"use strict";
// レビューコメントの表示用モデルと、編集操作のアンカーの指定（設計書 7 章「レビューコメント」、ADR-0021）。
// pptx の中の PowerPoint の新形式のコメント（ppt/comments/modernComment_*.xml、ppt/authors.xml）を Host が読み、Webview に送る。
// コメントは Google スライドとは同期しない（Push では取り除き、Pull では手元のコメントを引き継ぐ）。
// vscode にも DOM にも依存しない（Extension Host、Webview、CLI、vitest で使う）。
Object.defineProperty(exports, "__esModule", { value: true });
exports.COMMENT_TEXT_MAX = exports.COMMENT_IMAGES_MAX = exports.COMMENT_IMAGE_MAX_BYTES = void 0;
exports.initialsOf = initialsOf;
exports.formatCommentTime = formatCommentTime;
exports.parseCommentTime = parseCommentTime;
exports.shiftRange = shiftRange;
exports.COMMENT_IMAGE_MAX_BYTES = 2 * 1024 * 1024;
exports.COMMENT_IMAGES_MAX = 4;
/** コメントの本文の上限（文字数）。 */
exports.COMMENT_TEXT_MAX = 10_000;
/** 名前からイニシャルを作る（空白で区切った語の頭文字、最大 3 文字。和文の名前は先頭 2 文字）。 */
function initialsOf(name) {
    const words = name.trim().split(/\s+/).filter(Boolean);
    if (words.length === 0)
        return '?';
    if (words.length === 1) {
        const w = [...words[0]];
        return /^[A-Za-z]/.test(w[0] ?? '') ? w[0].toUpperCase() : w.slice(0, 2).join('');
    }
    return words
        .map((w) => [...w][0] ?? '')
        .join('')
        .toUpperCase()
        .slice(0, 3);
}
/**
 * PowerPoint の新形式と同じ形の作成日時（タイムゾーンを書かない UTC、ミリ秒まで）。
 * PowerPoint for the web の pptx で、変更の記録（changesInfo の dt）の最後と docProps/core.xml の modified（Z 付き）が
 * 秒まで一致したので、タイムゾーンのない値は UTC と判断した（設計書 7 章「レビューコメント」）。
 */
function formatCommentTime(d) {
    return d.toISOString().slice(0, -1);
}
/**
 * 作成日時を Date にする。タイムゾーンがなければ、zoneless が 'utc'（新形式。既定）なら UTC、'local'（旧形式の p:cm@dt）なら
 * ローカル時刻として読む。読めなければ undefined。
 */
function parseCommentTime(s, zoneless = 'utc') {
    const m = /^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})(?::(\d{2})(?:\.(\d{1,9}))?)?(Z|[+-]\d{2}:?\d{2})?$/.exec(s.trim());
    if (!m)
        return undefined;
    if (m[8]) {
        const d = new Date(s);
        return Number.isNaN(d.getTime()) ? undefined : d;
    }
    const ms = m[7] ? Number(m[7].slice(0, 3).padEnd(3, '0')) : 0;
    const parts = [Number(m[1]), Number(m[2]) - 1, Number(m[3]), Number(m[4]), Number(m[5]), Number(m[6] ?? 0), ms];
    const d = zoneless === 'utc' ? new Date(Date.UTC(...parts)) : new Date(...parts);
    return Number.isNaN(d.getTime()) ? undefined : d;
}
/**
 * テキストが before から after に変わったときの位置の付け替え。変わった部分は先頭と末尾の共通部分を除いた 1 つの区間とみなす。
 * 範囲の始まりは、変わった区間の中なら区間の始まりに、終わりは区間の新しい終わりに寄せる（範囲は置き換えた文字を含む）。
 * 範囲の直前に挿入した文字は範囲に含めず、直後に挿入した文字も含めない。
 */
function shiftRange(before, after, start, length) {
    let pre = 0;
    const max = Math.min(before.length, after.length);
    while (pre < max && before.charCodeAt(pre) === after.charCodeAt(pre))
        pre++;
    let suf = 0;
    while (suf < max - pre && before.charCodeAt(before.length - 1 - suf) === after.charCodeAt(after.length - 1 - suf))
        suf++;
    const oldEnd = before.length - suf;
    const newEnd = after.length - suf;
    const delta = newEnd - oldEnd;
    const a = Math.min(Math.max(0, start), before.length);
    const b = Math.min(Math.max(a, start + length), before.length);
    const mapStart = (p) => (p < pre ? p : p >= oldEnd ? p + delta : pre);
    const mapEnd = (p) => (p <= pre ? p : p >= oldEnd ? p + delta : newEnd);
    let na = mapStart(a);
    let nb = a === b ? na : mapEnd(b);
    na = Math.min(Math.max(0, na), after.length);
    nb = Math.min(Math.max(na, nb), after.length);
    return { start: na, length: nb - na };
}
