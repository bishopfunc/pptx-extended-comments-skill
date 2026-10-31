"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.isCommentXmlPath = isCommentXmlPath;
exports.commentSnapshot = commentSnapshot;
exports.countCommentChanges = countCommentChanges;
exports.summarizeCommentChanges = summarizeCommentChanges;
// 2 つの版のコメントの違いの要約（例「コメント +2・解決 1・返信 3」）。版の履歴の一覧の 2 行目と、外部の変更の取り込みの通知に使う。
// コメントのパーツ（ppt/comments/*.xml）の新形式のスレッドだけを見る。vscode に依存しない。
const comments_1 = require("./comments.js");
const xml_1 = require("./xml.js");
const commentImages_1 = require("./commentImages.js");
/** コメントのパーツのパス（新形式・旧形式とも ppt/comments/ の下）。 */
function isCommentXmlPath(path) {
    return /^ppt\/comments\/[^/]+\.xml$/i.test(path);
}
/** コメントのパーツの XML から、新形式のスレッドを集める。読めないパーツは飛ばす。 */
function commentSnapshot(texts) {
    const out = new Map();
    for (const text of texts) {
        let root;
        try {
            root = (0, xml_1.parseXml)(text).documentElement;
        }
        catch {
            continue;
        }
        for (const cm of (0, xml_1.kidsNamed)(root ?? undefined, 'p188:cm').slice(0, comments_1.COMMENT_READ_LIMITS.threadsPerPart)) {
            const id = (0, xml_1.attr)(cm, 'id')?.toLowerCase();
            if (!id || out.has(id))
                continue;
            const replies = new Map();
            const replyImages = new Map();
            for (const r of (0, comments_1.repliesOf)(cm, comments_1.ALL)) {
                const rid = (0, xml_1.attr)(r, 'id')?.toLowerCase();
                if (rid)
                    replies.set(rid, (0, comments_1.txBodyText)((0, xml_1.kid)(r, 'p188:txBody')));
                if (rid && (0, commentImages_1.imageKey)(r))
                    replyImages.set(rid, (0, commentImages_1.imageKey)(r));
            }
            out.set(id, { status: (0, comments_1.statusOf)(cm), text: (0, comments_1.txBodyText)((0, xml_1.kid)(cm, 'p188:txBody')), replies,
                ...((0, commentImages_1.imageKey)(cm) ? { images: (0, commentImages_1.imageKey)(cm) } : {}), ...(replyImages.size ? { replyImages } : {}) });
            if (out.size >= comments_1.COMMENT_READ_LIMITS.threadsTotal)
                return out;
        }
    }
    return out;
}
function countCommentChanges(before, after) {
    const c = { added: 0, removed: 0, resolved: 0, reopened: 0, replies: 0, repliesRemoved: 0, edited: 0 };
    for (const [id, a] of after) {
        const b = before.get(id);
        if (!b) {
            c.added++;
            c.replies += a.replies.size;
            continue;
        }
        if (b.status !== a.status) {
            if (a.status === 'resolved')
                c.resolved++;
            else
                c.reopened++;
        }
        if (b.text !== a.text || b.images !== a.images)
            c.edited++;
        for (const [rid, t] of a.replies) {
            const bt = b.replies.get(rid);
            if (bt === undefined)
                c.replies++;
            else if (bt !== t || b.replyImages?.get(rid) !== a.replyImages?.get(rid))
                c.edited++;
        }
        for (const rid of b.replies.keys())
            if (!a.replies.has(rid))
                c.repliesRemoved++;
    }
    for (const id of before.keys())
        if (!after.has(id))
            c.removed++;
    return c;
}
/** 要約（変化がなければ undefined）。 */
function summarizeCommentChanges(before, after) {
    const c = countCommentChanges(before, after);
    const parts = [];
    if (c.added)
        parts.push(`コメント +${c.added}`);
    if (c.removed)
        parts.push(`コメント −${c.removed}`);
    if (c.resolved)
        parts.push(`解決 ${c.resolved}`);
    if (c.reopened)
        parts.push(`再開 ${c.reopened}`);
    if (c.replies)
        parts.push(`返信 ${c.replies}`);
    if (c.repliesRemoved)
        parts.push(`返信 −${c.repliesRemoved}`);
    if (c.edited)
        parts.push(`編集 ${c.edited}`);
    return parts.length ? parts.join('・') : undefined;
}
