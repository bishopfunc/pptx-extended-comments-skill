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
exports.GitShowError = void 0;
exports.findGitOnPath = findGitOnPath;
exports.resolveGit = resolveGit;
exports.readGitBlob = readGitBlob;
// git の版の pptx をバイト列のまま読む（「スライドの差分を表示（HEAD と比べる）」）。vscode に依存しない。
// git 拡張の API に頼らず `git show <ref>:./<path>` を encoding 'buffer' で実行する（バイナリを壊さない）。
// git は絶対パスで呼ぶ（Windows はコマンド名だけだと作業フォルダの git.exe を先に探すため、リポジトリに置かれた git.exe を実行しうる。ADR-0019）。
const node_child_process_1 = require("node:child_process");
const node_fs_1 = require("node:fs");
const path = __importStar(require("node:path"));
const gitUri_1 = require("../shared/gitUri.js");
/** pptx は数十 MB になることがある。 */
const MAX_BUFFER = 1024 * 1024 * 1024;
class GitShowError extends Error {
    reason;
    constructor(message, 
    /** notInRepo：git の管理下にない（git がない場合も含む）。notInRef：その版にファイルがない（新しいファイル）。 */
    reason) {
        super(message);
        this.reason = reason;
    }
}
exports.GitShowError = GitShowError;
/** いちばん近い、存在するフォルダ（削除したファイルでも、そのフォルダから git を実行できるように）。 */
async function existingDir(start) {
    let dir = start;
    for (;;) {
        try {
            if ((await node_fs_1.promises.stat(dir)).isDirectory())
                return dir;
        }
        catch {
            // 上へ
        }
        const up = path.dirname(dir);
        if (up === dir)
            return dir;
        dir = up;
    }
}
function defaultIsFile(p) {
    try {
        if (!(0, node_fs_1.statSync)(p).isFile())
            return false;
        if (process.platform !== 'win32')
            (0, node_fs_1.accessSync)(p, node_fs_1.constants.X_OK);
        return true;
    }
    catch {
        return false;
    }
}
/**
 * PATH から git の絶対パスを探す。相対・空の項目と、作業フォルダ（exclude）は飛ばす。Windows では git.exe だけを探す
 * （.bat・.cmd は cmd.exe を通るため使わない）。見つからなければ undefined。
 */
function findGitOnPath(opts = {}) {
    const platform = opts.platform ?? process.platform;
    const win = platform === 'win32';
    const p = win ? path.win32 : path.posix;
    const pathEnv = opts.pathEnv ?? (win ? (process.env.Path ?? process.env.PATH) : process.env.PATH);
    if (!pathEnv)
        return undefined;
    const isFile = opts.isFile ?? defaultIsFile;
    const norm = (d) => {
        const r = p.resolve(d);
        const t = r.length > 1 && (r.endsWith('/') || r.endsWith('\\')) && !/^[A-Za-z]:\\$/.test(r) ? r.slice(0, -1) : r;
        return win ? t.toLowerCase() : t;
    };
    const excluded = new Set((opts.exclude ?? []).map(norm));
    for (const raw of pathEnv.split(win ? ';' : ':')) {
        const dir = win ? raw.trim().replace(/^"(.*)"$/, '$1') : raw;
        if (!dir || !p.isAbsolute(dir) || (win && !/^([A-Za-z]:[\\/]|\\\\)/.test(dir)))
            continue;
        if (excluded.has(norm(dir)))
            continue;
        const candidate = p.join(dir, win ? 'git.exe' : 'git');
        if (isFile(candidate))
            return candidate;
    }
    return undefined;
}
/** 使う git の絶対パス。渡された（vscode.git の）パスが絶対パスならそれ、なければ PATH から探す。 */
function resolveGit(cwd, preferred) {
    if (preferred && path.isAbsolute(preferred) && (process.platform !== 'win32' || /\.exe$/i.test(preferred)))
        return preferred;
    return findGitOnPath({ exclude: [cwd] });
}
function run(git, args, cwd) {
    return new Promise((resolve, reject) => {
        (0, node_child_process_1.execFile)(git, args, { cwd, encoding: 'buffer', maxBuffer: MAX_BUFFER, windowsHide: true }, (err, stdout, stderr) => {
            if (err)
                reject(Object.assign(err, { stderrText: Buffer.from(stderr ?? '').toString('utf8') }));
            else
                resolve(stdout);
        });
    });
}
/** filePath（絶対パス）の ref の版の中身。リポジトリの外なら notInRepo、その版になければ notInRef の GitShowError。 */
async function readGitBlob(filePath, ref = 'HEAD', opts = {}) {
    const cwd = await existingDir(path.dirname(filePath));
    const git = resolveGit(cwd, opts.gitPath);
    if (!git)
        throw new GitShowError('git が見つかりません。git をインストールし、PATH に追加してください', 'notInRepo');
    try {
        await run(git, ['rev-parse', '--is-inside-work-tree'], cwd);
    }
    catch {
        throw new GitShowError(`${path.basename(filePath)} は git のリポジトリの中にありません`, 'notInRepo');
    }
    try {
        const out = await run(git, ['show', (0, gitUri_1.gitShowSpec)(ref, path.relative(cwd, filePath))], cwd);
        return new Uint8Array(out.buffer, out.byteOffset, out.byteLength);
    }
    catch {
        throw new GitShowError(`${ref} に ${path.basename(filePath)} はありません（新しいファイルか、git の管理外です）`, 'notInRef');
    }
}
