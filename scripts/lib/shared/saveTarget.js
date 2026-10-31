"use strict";
// 新しいプレゼンテーションと PDF の保存先（開いているフォルダに保存し、フォルダがないときだけ保存先を尋ねる）。
// vscode に依存しない。フォルダは URI でもパスでもよい（型引数）。
Object.defineProperty(exports, "__esModule", { value: true });
exports.UNTITLED_STEM = void 0;
exports.numberedName = numberedName;
exports.claimName = claimName;
exports.targetFolder = targetFolder;
exports.pdfStem = pdfStem;
/** 新しいプレゼンテーションの名前（Google スライドと同じ）。 */
exports.UNTITLED_STEM = '無題のプレゼンテーション';
/** 番号を付けた名前。0 なら「stem.ext」、1 以上なら「stem (n).ext」（Finder・エクスプローラーのコピーと同じ形）。 */
function numberedName(stem, ext, n) {
    return n === 0 ? `${stem}${ext}` : `${stem} (${n})${ext}`;
}
/** 番号を付ける上限（これを超えたら諦める。無限に回らないため）。 */
const MAX_NUMBER = 9999;
/**
 * 「stem.ext」「stem (1).ext」「stem (2).ext」… を順に claim に渡し、最初に取れた（true を返した）名前を返す。
 * claim は、名前が空いているかを見るだけでも、その名前で排他的に作る（既にあれば false）のでもよい。既存のファイルは上書きしない。
 */
async function claimName(stem, ext, claim) {
    for (let n = 0; n <= MAX_NUMBER; n++) {
        const name = numberedName(stem, ext, n);
        if (await claim(name))
            return name;
    }
    throw new Error(`「${stem}${ext}」と同じ名前のファイルが多すぎるため、名前を決められませんでした`);
}
/**
 * PDF・新しいプレゼンテーションを置くフォルダ。pptx の隣、なければワークスペースの最初のフォルダ。
 * どちらもない（フォルダを開いていないウィンドウの新しいプレゼンテーション）ときだけ undefined（保存先を尋ねる）。
 */
function targetFolder(choice) {
    return choice.fileFolder ?? choice.workspaceFolder;
}
/** PDF の名前の元（拡張子なし）。pptx のファイル名（新しいプレゼンテーションなら名前の欄の名前）から .pptx を除く。空なら UNTITLED_STEM。 */
function pdfStem(fileName) {
    return fileName.replace(/\.pptx$/i, '').trim() || exports.UNTITLED_STEM;
}
