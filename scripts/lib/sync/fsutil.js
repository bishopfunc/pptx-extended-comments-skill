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
exports.sha256 = sha256;
exports.canonicalPath = canonicalPath;
exports.tempPathFor = tempPathFor;
exports.writeFileAtomic = writeFileAtomic;
exports.createFileExclusiveAtomic = createFileExclusiveAtomic;
exports.createNumberedFile = createNumberedFile;
exports.readFileOrUndefined = readFileOrUndefined;
exports.pathExists = pathExists;
exports.moveFileNoClobber = moveFileNoClobber;
exports.isSameFile = isSameFile;
exports.errCode = errCode;
exports.hasZipSignature = hasZipSignature;
exports.isValidZip = isValidZip;
exports.formatTimestamp = formatTimestamp;
exports.parseTimestamp = parseTimestamp;
// sync で使うファイル操作の小道具（ADR-0010：一時ファイル + rename のアトミックな書き込み）。
const node_crypto_1 = require("node:crypto");
const node_fs_1 = require("node:fs");
const fs = __importStar(require("node:fs/promises"));
const path = __importStar(require("node:path"));
const zipLimits_1 = require("../pptx/zipLimits.js");
const saveTarget_1 = require("../shared/saveTarget.js");
function sha256(data) {
    return (0, node_crypto_1.createHash)('sha256').update(data).digest('hex');
}
/**
 * 同じファイルを指すパスを 1 つにそろえた鍵（ロック、履歴の索引、置き場所が同じかの判定に使う）。
 * まだないフォルダは、あるところまでさかのぼってシンボリックリンクを解決する（作る前と後で鍵が変わらないように）。
 * 大文字小文字を区別しないことが多い macOS と Windows では小文字にする（区別するボリュームでは、名前が大文字小文字だけ違う
 * 別のファイルが同じ鍵になるが、ロックが余分に効くだけで安全側）。
 */
function canonicalPath(p) {
    const abs = path.resolve(p);
    let dir = path.dirname(abs);
    const rest = [path.basename(abs)];
    for (;;) {
        try {
            dir = node_fs_1.realpathSync.native(dir);
            break;
        }
        catch {
            const parent = path.dirname(dir);
            if (parent === dir)
                break; // 根まで解決できない：そのまま使う
            rest.unshift(path.basename(dir));
            dir = parent;
        }
    }
    const out = path.join(dir, ...rest);
    return process.platform === 'darwin' || process.platform === 'win32' ? out.toLowerCase() : out;
}
/** 同じフォルダに置く一時ファイルの名前（退避一覧などの正規表現に当たらないよう、先頭を "." にする）。 */
function tempPathFor(target) {
    return path.join(path.dirname(target), `.${path.basename(target)}.${(0, node_crypto_1.randomUUID)().slice(0, 8)}.tmp`);
}
async function writeTemp(tmp, data, mode) {
    const fh = await fs.open(tmp, 'wx', mode);
    try {
        await fh.writeFile(data);
        await fh.sync();
    }
    finally {
        await fh.close();
    }
}
/** 同じフォルダの一時ファイルに書いてから rename で置き換える。既存ファイルのパーミッションは引き継ぐ。 */
async function writeFileAtomic(target, data) {
    let mode;
    try {
        mode = (await fs.stat(target)).mode & 0o777;
    }
    catch {
        mode = undefined;
    }
    const tmp = tempPathFor(target);
    try {
        await writeTemp(tmp, data, mode);
        await fs.rename(tmp, target);
    }
    catch (e) {
        await fs.rm(tmp, { force: true }).catch(() => undefined);
        throw e;
    }
}
/**
 * 同じ名前のファイルがないときだけ、アトミックに作る（一時ファイル → hard link）。
 * 既にあれば EEXIST を投げる。hard link が使えないファイルシステムでは、排他フラグで名前を確保してから rename する。
 */
async function createFileExclusiveAtomic(target, data) {
    const tmp = tempPathFor(target);
    try {
        await writeTemp(tmp, data);
        try {
            await fs.link(tmp, target);
        }
        catch (e) {
            const code = errCode(e);
            if (code === 'EEXIST')
                throw e;
            if (code !== 'EPERM' && code !== 'ENOTSUP' && code !== 'EXDEV' && code !== 'ENOSYS')
                throw e;
            // hard link が使えない：名前を排他フラグで確保してから置き換える。
            const fh = await fs.open(target, 'wx');
            await fh.close();
            await fs.rename(tmp, target);
        }
    }
    finally {
        await fs.rm(tmp, { force: true }).catch(() => undefined);
    }
}
/**
 * dir に「stem.ext」を作る。既にあれば「stem (1).ext」「stem (2).ext」…と番号を付け、空いている最初の名前で作る。
 * 既存のファイルは上書きしない（名前の確保と書き込みは createFileExclusiveAtomic で排他的に行う）。作ったファイルのパスを返す。
 */
async function createNumberedFile(dir, stem, ext, data) {
    const name = await (0, saveTarget_1.claimName)(stem, ext, async (n) => {
        try {
            await createFileExclusiveAtomic(path.join(dir, n), data);
            return true;
        }
        catch (e) {
            if (errCode(e) === 'EEXIST')
                return false;
            throw e;
        }
    });
    return path.join(dir, name);
}
async function readFileOrUndefined(p) {
    try {
        return await fs.readFile(p);
    }
    catch (e) {
        if (errCode(e) === 'ENOENT')
            return undefined;
        throw e;
    }
}
async function pathExists(p) {
    try {
        await fs.stat(p);
        return true;
    }
    catch {
        return false;
    }
}
/** 2 つのパスが同じファイルを指すか（大文字小文字を区別しないファイルシステムで、大文字小文字だけを変える名前変更など）。 */
/**
 * from を to へ移す。to が既にあれば（確かめたあとに別のウィンドウが作った場合も）上書きせず false。
 * ハードリンクを作ってから元を消す。ハードリンクを作れないファイルシステムでは、排他のコピーにする。
 */
async function moveFileNoClobber(from, to) {
    try {
        await fs.link(from, to);
    }
    catch (e) {
        if (errCode(e) === 'EEXIST')
            return false;
        if (errCode(e) === 'ENOENT')
            throw e;
        try {
            await fs.copyFile(from, to, node_fs_1.constants.COPYFILE_EXCL);
        }
        catch (e2) {
            if (errCode(e2) === 'EEXIST')
                return false;
            throw e2;
        }
    }
    await fs.rm(from, { force: true });
    return true;
}
async function isSameFile(a, b) {
    try {
        const [sa, sb] = await Promise.all([fs.stat(a), fs.stat(b)]);
        return sa.dev === sb.dev && sa.ino === sb.ino;
    }
    catch {
        return false;
    }
}
function errCode(e) {
    return typeof e === 'object' && e !== null && 'code' in e ? String(e.code) : undefined;
}
function hasZipSignature(bytes) {
    return bytes.length >= 4 && bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}
/** zip として開けるか（署名と、JSZip で目次を読めること。中身は展開しない。ファイルの数が多すぎれば false）。 */
async function isValidZip(bytes, limits) {
    if (!hasZipSignature(bytes))
        return false;
    try {
        await (0, zipLimits_1.openZip)(bytes, limits);
        return true;
    }
    catch {
        return false;
    }
}
/** ローカル時刻で YYYYMMDD_HHMMSS（ms を付けるなら _mmm も）。 */
function formatTimestamp(d, withMs) {
    const p = (n, w = 2) => String(n).padStart(w, '0');
    const base = `${d.getFullYear()}${p(d.getMonth() + 1)}${p(d.getDate())}_${p(d.getHours())}${p(d.getMinutes())}${p(d.getSeconds())}`;
    return withMs ? `${base}_${p(d.getMilliseconds(), 3)}` : base;
}
/** formatTimestamp(d, true) の逆。 */
function parseTimestamp(s) {
    const m = s.match(/^(\d{4})(\d{2})(\d{2})_(\d{2})(\d{2})(\d{2})(?:_(\d{3}))?$/);
    if (!m)
        return undefined;
    const n = m.slice(1).map((x) => (x === undefined ? 0 : Number(x)));
    return new Date(n[0], n[1] - 1, n[2], n[3], n[4], n[5], n[6]);
}
