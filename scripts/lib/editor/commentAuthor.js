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
exports.cleanAuthorName = cleanAuthorName;
exports.gitUserName = gitUserName;
exports.osUserName = osUserName;
exports.resolveCommentAuthor = resolveCommentAuthor;
// コメントの作成者の名前（設定 slidesSync.commentAuthor）。vscode に依存しない。
// 未設定なら git の user.name、なければ OS のユーザー名。メールアドレスは使わない（pptx の authors.xml に名前として残るため）。
const node_child_process_1 = require("node:child_process");
const os = __importStar(require("node:os"));
const gitShow_1 = require("./gitShow.js");
/** 名前として使えるか（空でなく、メールアドレスに見えない）。前後の空白を除いて、長すぎれば切る。 */
function cleanAuthorName(v) {
    if (typeof v !== 'string')
        return undefined;
    const t = v.replace(/[\u0000-\u001f\u007f]/g, '').trim().slice(0, 100);
    if (!t || /\S+@\S+\.\S+/.test(t))
        return undefined;
    return t;
}
/** git の user.name（ホームフォルダで読む。リポジトリの設定は見ない）。読めなければ undefined。 */
function gitUserName(timeoutMs = 3000) {
    const home = os.homedir();
    const git = (0, gitShow_1.findGitOnPath)({ exclude: [process.cwd()] });
    if (!git)
        return Promise.resolve(undefined);
    return new Promise((resolve) => {
        (0, node_child_process_1.execFile)(git, ['config', '--global', '--get', 'user.name'], { cwd: home, timeout: timeoutMs, windowsHide: true, encoding: 'utf8' }, (err, out) => resolve(err ? undefined : cleanAuthorName(String(out))));
    });
}
function osUserName() {
    try {
        return cleanAuthorName(os.userInfo().username);
    }
    catch {
        return undefined;
    }
}
let gitNameCache;
/** 設定の名前、なければ git の user.name、なければ OS のユーザー名、どれもなければ「自分」。 */
async function resolveCommentAuthor(configured, deps = {}) {
    const c = cleanAuthorName(configured);
    if (c)
        return c;
    const git = deps.gitName ? await deps.gitName() : await (gitNameCache ??= gitUserName());
    return git ?? (deps.osName ?? osUserName)() ?? '自分';
}
