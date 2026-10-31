"use strict";
// git の版の pptx（git: の URI）の扱い（設計書 8 章「git の差分」）。vscode に依存しない（Extension Host と vitest で使う）。
//
// VS Code の git 拡張は、版の中身を git: の URI で開く。クエリは JSON で、{ path: 作業ツリーのファイルの絶対パス, ref } を持つ
// （extensions/git/src/uri.ts の toGitUri。サブモジュールでは submoduleOf も付く）。ref の意味は次のとおり。
//   '~'    インデックス（ソース管理の「変更」の差分の左側。右側は作業ツリーのファイル）
//   'HEAD' HEAD（「ステージされている変更」の差分の左側。右側はインデックス（ref が ''）。「ファイルを開く (HEAD)」でも単独で開く）
//   ''     インデックス（「ステージされている変更」の差分の右側）
//   '~1'〜'~3' マージの競合の各段、そのほかはコミットのハッシュなど（タイムラインの差分は左右とも git:）
Object.defineProperty(exports, "__esModule", { value: true });
exports.INDEX_NAME = exports.WORKING_TREE_NAME = void 0;
exports.documentSource = documentSource;
exports.documentCapabilities = documentCapabilities;
exports.parseGitUriQuery = parseGitUriQuery;
exports.gitUriQuery = gitUriQuery;
exports.gitOpenPlan = gitOpenPlan;
exports.gitRefName = gitRefName;
exports.isAbsentVersion = isAbsentVersion;
exports.gitShowSpec = gitShowSpec;
function documentSource(scheme) {
    if (scheme === 'file')
        return 'file';
    if (scheme === 'untitled')
        return 'untitled';
    return 'snapshot';
}
/** その出どころのドキュメントでできること。 */
function documentCapabilities(source) {
    switch (source) {
        case 'file':
            return { edit: true, sync: true, watch: true, save: true };
        case 'untitled':
            return { edit: true, sync: false, watch: false, save: true };
        case 'snapshot':
            return { edit: false, sync: false, watch: false, save: false };
    }
}
/** git: の URI のクエリを読む。git 拡張の形でなければ undefined。 */
function parseGitUriQuery(query) {
    if (!query)
        return undefined;
    let parsed;
    try {
        parsed = JSON.parse(query);
    }
    catch {
        return undefined;
    }
    if (!parsed || typeof parsed !== 'object')
        return undefined;
    const { path, ref, submoduleOf } = parsed;
    if (typeof path !== 'string' || !path || typeof ref !== 'string')
        return undefined;
    return { path, ref, ...(typeof submoduleOf === 'string' ? { submoduleOf } : {}) };
}
/** git 拡張の toGitUri と同じクエリ（ref の版を指す git: の URI を組み立てるときに使う）。 */
function gitUriQuery(fsPath, ref) {
    return JSON.stringify({ path: fsPath, ref });
}
function gitOpenPlan(params) {
    if (!params || params.submoduleOf)
        return 'viewer';
    if (params.ref === 'HEAD')
        return 'indexOrWorkingTree';
    return 'workingTree';
}
/** 差分の画面の列の名前。 */
exports.WORKING_TREE_NAME = '作業ツリー';
exports.INDEX_NAME = 'インデックス';
/** ref の版の表示名（列の見出し）。 */
function gitRefName(ref) {
    if (ref === '' || ref === '~' || ref === '~0')
        return exports.INDEX_NAME;
    if (/^~[1-3]$/.test(ref))
        return `競合の段 ${ref.slice(1)}`;
    if (/^[0-9a-f]{40}$/i.test(ref))
        return ref.slice(0, 7);
    return ref;
}
/**
 * git の版のバイト列が「その版にはファイルがない」を表すか。git 拡張のファイルシステムは、版にないファイルを
 * 読むと空のバイト列を返す（エラーにしない）。
 */
function isAbsentVersion(bytes) {
    return !bytes || bytes.length === 0;
}
/**
 * `git show` に渡す版とパス（cwd からの相対パス。`HEAD:./a/b.pptx` の形にすると、リポジトリのルートを調べずに済み、
 * macOS の /tmp と /private/tmp のようなシンボリックリンクの違いにも左右されない）。
 */
function gitShowSpec(ref, relativeFromCwd) {
    const rel = relativeFromCwd.split(/[\\/]+/).filter((s) => s && s !== '.').join('/');
    return `${ref}:./${rel}`;
}
