"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.WRITE_LOCK_STALE_MS = exports.LOCK_DIR = exports.LOCK_MAX_AGE_MS = exports.LOCK_STALE_MS = void 0;
exports.setLockFallbackDir = setLockFallbackDir;
exports.resolveLockRoot = resolveLockRoot;
exports.lockFileIn = lockFileIn;
exports.lockPathFor = lockPathFor;
exports.isLocked = isLocked;
exports.acquireLock = acquireLock;
exports.acquireWriteLock = acquireWriteLock;
exports.waitForWriteLock = waitForWriteLock;
exports.acquireLockAt = acquireLockAt;
// プロセス（ウィンドウ）をまたぐロック（設計書 10 章「ロック」、ADR-0010）。
// ロックは pptx のフォルダではなく一時フォルダに置く（ADR-0014。同じ PC の別の VS Code とも共有する）。
const node_crypto_1 = require("node:crypto");
const fs = __importStar(require("node:fs/promises"));
const os = __importStar(require("node:os"));
const path = __importStar(require("node:path"));
const fsutil_1 = require("./fsutil.js");
exports.LOCK_STALE_MS = 10 * 60 * 1000;
/** 別のホストのロックでも、ファイルがこれより古ければ放棄とみなす（時刻を書き直し続けるロックで、いつまでも取れなくならないように）。 */
exports.LOCK_MAX_AGE_MS = 30 * 60 * 1000;
/** これより先の時刻のロックは壊れたものとみなす（時計のずれは許す）。 */
const LOCK_FUTURE_SKEW_MS = 5 * 60 * 1000;
const posix = process.platform !== 'win32';
/**
 * ロックを置く一時フォルダ。POSIX ではユーザーごと（`google-slides-sync-<uid>`、0700）にする。/tmp は同じ PC のほかの人と
 * 共有なので、共有の名前だとほかの人がシンボリックリンクやロックを先に置けるため（ADR-0019）。Windows の一時フォルダはユーザーごと。
 */
exports.LOCK_DIR = path.join(os.tmpdir(), posix ? `google-slides-sync-${process.getuid?.() ?? 'user'}` : 'google-slides-sync');
let lockRoot = exports.LOCK_DIR;
let lockFallback;
let lockRootReady;
/** 一時フォルダが安全でないときに代わりに使うフォルダ（拡張機能のグローバルストレージの下）。 */
function setLockFallbackDir(dir) {
    lockFallback = dir;
    lockRootReady = undefined;
}
/** 自分だけが使えるフォルダか確かめる（なければ 0700 で作る）。シンボリックリンク、持ち主が違う、ほかの人も使える、なら例外。 */
async function ensurePrivateDir(dir, uid) {
    try {
        await fs.mkdir(dir, { mode: 0o700 });
    }
    catch (e) {
        if ((0, fsutil_1.errCode)(e) === 'ENOENT')
            await fs.mkdir(dir, { recursive: true, mode: 0o700 });
        else if ((0, fsutil_1.errCode)(e) !== 'EEXIST')
            throw e;
    }
    if (uid === undefined)
        return; // Windows：一時フォルダはユーザーごと
    const st = await fs.lstat(dir);
    if (st.isSymbolicLink() || !st.isDirectory() || st.uid !== uid || (st.mode & 0o077) !== 0) {
        throw new Error(`ロックの一時フォルダ ${dir} が安全ではありません（シンボリックリンク、ほかのユーザーの持ち物、またはほかのユーザーも書ける）。` +
            'このフォルダを消してから、やり直してください');
    }
}
/**
 * ロックを置くフォルダを確かめて返す。dir が安全でなければ fallback を使い、なければ例外にする。
 */
async function resolveLockRoot(opts) {
    try {
        await ensurePrivateDir(opts.dir, opts.uid);
        return opts.dir;
    }
    catch (e) {
        if (!opts.fallback)
            throw e;
        await ensurePrivateDir(opts.fallback, opts.uid);
        return opts.fallback;
    }
}
/** 確かめたロックのフォルダの下の、サブフォルダ（locks・history-locks）の中のファイルのパス。 */
async function lockFileIn(sub, name) {
    lockRootReady ??= resolveLockRoot({ dir: exports.LOCK_DIR, uid: posix ? process.getuid?.() : undefined, fallback: lockFallback }).then((root) => (lockRoot = root), (e) => {
        lockRootReady = undefined;
        throw e;
    });
    const root = await lockRootReady;
    return path.join(root, sub, name);
}
function lockKey(pptxPath) {
    return `${(0, node_crypto_1.createHash)('sha256').update((0, fsutil_1.canonicalPath)(pptxPath)).digest('hex').slice(0, 32)}.lock`;
}
/**
 * `<ロックのフォルダ>/locks/<pptx のパスの SHA-256>.lock`。パスはシンボリックリンクや大文字小文字の違いを
 * そろえてから使う（別のウィンドウが別の書き方で同じファイルを開いても、同じロックになるように）。
 */
function lockPathFor(pptxPath) {
    return path.join(lockRoot, 'locks', lockKey(pptxPath));
}
function parseLock(text) {
    try {
        const v = JSON.parse(text);
        if (typeof v.pid === 'number' && typeof v.time === 'string' && typeof v.token === 'string') {
            return { pid: v.pid, time: v.time, token: v.token, hostname: typeof v.hostname === 'string' ? v.hostname : '' };
        }
    }
    catch {
        // 壊れている
    }
    return undefined;
}
function processAlive(pid) {
    try {
        process.kill(pid, 0);
        return true;
    }
    catch (e) {
        return (0, fsutil_1.errCode)(e) === 'EPERM';
    }
}
/**
 * 放棄されたロックか。時刻が 10 分以上前なら放棄とみなす。
 * 同じマシンで PID のプロセスがもういないときも放棄とみなす（落ちたウィンドウのロックで 10 分待たないため）。
 */
async function isStale(lockFile, content, opts) {
    const nowMs = opts.now().getTime();
    if (content && content.hostname !== opts.hostname) {
        // 別のホストのロック：時刻を書き直し続けても、ファイルが古ければ放棄とみなす
        const st = await fs.stat(lockFile).catch(() => undefined);
        if (!st || nowMs - st.mtimeMs >= exports.LOCK_MAX_AGE_MS)
            return true;
    }
    if (!content) {
        // 書きかけ・壊れたロック：ファイルの更新時刻で判断する
        try {
            const st = await fs.stat(lockFile);
            return nowMs - st.mtimeMs >= opts.staleMs;
        }
        catch {
            return true;
        }
    }
    const t = Date.parse(content.time);
    if (Number.isNaN(t) || nowMs - t >= opts.staleMs || t - nowMs > LOCK_FUTURE_SKEW_MS)
        return true;
    if (content.hostname === opts.hostname && content.pid !== opts.pid && !processAlive(content.pid))
        return true;
    return false;
}
function resolveOpts(opts) {
    return {
        now: opts.now ?? (() => new Date()),
        staleMs: opts.staleMs ?? exports.LOCK_STALE_MS,
        pid: opts.pid ?? process.pid,
        hostname: opts.hostname ?? os.hostname(),
    };
}
/** 有効なロックが別にあるか。 */
async function isLocked(pptxPath, options = {}) {
    const opts = resolveOpts(options);
    const p = await lockFileIn('locks', lockKey(pptxPath));
    let text;
    try {
        text = await fs.readFile(p, 'utf8');
    }
    catch {
        return false;
    }
    return !(await isStale(p, parseLock(text), opts));
}
/** pptx のロックを取る。他が持っていれば undefined。 */
async function acquireLock(pptxPath, options = {}) {
    return acquireLockAt(await lockFileIn('locks', lockKey(pptxPath)), options);
}
/**
 * pptx を書き込む間のロック（エディタの保存と、コメントの CLI の書き込みが重ならないように。ADR-0021）。
 * 同期のロックとは別のファイル。書き込みは短いので、2 分で放棄とみなす。他が持っていれば undefined。
 */
async function acquireWriteLock(pptxPath, options = {}) {
    return acquireLockAt(await lockFileIn('writes', lockKey(pptxPath)), { staleMs: exports.WRITE_LOCK_STALE_MS, ...options });
}
exports.WRITE_LOCK_STALE_MS = 2 * 60 * 1000;
/** 書き込みのロックを、待ちながら取る（取れなければ undefined）。 */
async function waitForWriteLock(pptxPath, waitMs, options = {}) {
    const until = Date.now() + waitMs;
    for (;;) {
        const lock = await acquireWriteLock(pptxPath, options);
        if (lock || Date.now() >= until)
            return lock;
        await new Promise((r) => setTimeout(r, 100));
    }
}
/** 指定したロックファイルでロックを取る。他が持っていれば undefined。 */
async function acquireLockAt(p, options = {}) {
    const opts = resolveOpts(options);
    await ensurePrivateDir(path.dirname(p), posix ? process.getuid?.() : undefined);
    const token = (0, node_crypto_1.randomUUID)();
    const content = { pid: opts.pid, hostname: opts.hostname, time: opts.now().toISOString(), token };
    for (let attempt = 0; attempt < 3; attempt++) {
        try {
            await fs.writeFile(p, JSON.stringify(content), { flag: 'wx', mode: 0o600 });
            return {
                path: p,
                release: async () => {
                    // 自分のロックのときだけ消す（失効して他に取られていたら触らない）。
                    try {
                        const cur = parseLock(await fs.readFile(p, 'utf8'));
                        if (cur?.token === token)
                            await fs.rm(p, { force: true });
                    }
                    catch {
                        // ない
                    }
                },
            };
        }
        catch (e) {
            if ((0, fsutil_1.errCode)(e) !== 'EEXIST')
                throw e;
        }
        let text;
        try {
            text = await fs.readFile(p, 'utf8');
        }
        catch (e) {
            if ((0, fsutil_1.errCode)(e) === 'ENOENT')
                continue; // 直前に解放された
            throw e;
        }
        const held = parseLock(text);
        if (!(await isStale(p, held, opts)))
            return undefined;
        // 放棄されたロックを取り直す。別プロセスと同時に取り直す場合に備え、
        // いったん固有の名前に移してから中身を確かめる。新しいロックを動かしてしまったら戻す。
        const aside = `${p}.${token}.stale`;
        try {
            await fs.rename(p, aside);
        }
        catch (e) {
            if ((0, fsutil_1.errCode)(e) === 'ENOENT')
                continue;
            throw e;
        }
        const movedText = await fs.readFile(aside, 'utf8').catch(() => '');
        const moved = parseLock(movedText);
        if (moved && held && moved.token !== held.token && !(await isStale(aside, moved, opts))) {
            // 他のプロセスが取り直したばかりのロックだった：戻して譲る
            await fs.link(aside, p).catch(() => undefined);
            await fs.rm(aside, { force: true });
            return undefined;
        }
        await fs.rm(aside, { force: true });
    }
    return undefined;
}
