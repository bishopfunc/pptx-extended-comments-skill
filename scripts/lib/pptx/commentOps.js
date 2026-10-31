"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.shiftRange = void 0;
exports.isCommentOp = isCommentOp;
exports.applyCommentOp = applyCommentOp;
exports.captureTextAnchors = captureTextAnchors;
exports.shiftTextAnchors = shiftTextAnchors;
// コメントの編集操作（addComment など）と、テキストの編集に合わせた文字の範囲のアンカーの付け替え（設計書 7 章「レビューコメント」）。
// 自分の投稿（作成者の名前が設定の名前と同じ）だけを編集・削除できる（Google スライドと同じ）。解決・再開は誰でもできる。
// 旧形式のコメント（legacy:…）は読み取り専用。
const comments_1 = require("../shared/comments.js");
const commentImages_1 = require("../shared/commentImages.js");
const commentImages_2 = require("./commentImages.js");
const creationIds_1 = require("./creationIds.js");
const edit_1 = require("./edit.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const comments_2 = require("./comments.js");
const package_1 = require("./package.js");
const xml_1 = require("./xml.js");
const MSG = {
    noAuthor: 'コメントの作成者の名前が設定されていません（設定 slidesSync.commentAuthor）',
    empty: 'コメントの本文が空です',
    tooLong: `コメントの本文が長すぎます（${comments_1.COMMENT_TEXT_MAX} 文字まで）`,
    threadNotFound: 'コメントが見つかりません。削除された可能性があります',
    postNotFound: '返信が見つかりません。削除された可能性があります',
    legacy: 'このコメントは旧形式（Google スライドの書き出しなど）のため、表示だけできます。新しいコメントを追加してください',
    notMine: '自分のコメントだけを編集・削除できます',
    noText: 'この要素の文字にはコメントを付けられません',
    badRange: '文字の範囲が正しくありません',
    badElement: 'この要素にはコメントを付けられません',
};
function isCommentOp(op) {
    return (op.type === 'addComment' ||
        op.type === 'replyComment' ||
        op.type === 'editComment' ||
        op.type === 'attachCommentImage' ||
        op.type === 'removeCommentImage' ||
        op.type === 'deleteComment' ||
        op.type === 'resolveComment' ||
        op.type === 'reopenComment');
}
function authorOf(ctx) {
    const a = ctx?.comment?.author?.trim();
    if (!a)
        (0, edit_1.reject)(MSG.noAuthor);
    return a;
}
function nowOf(ctx) {
    return (0, comments_1.formatCommentTime)(ctx?.comment?.now ?? new Date());
}
function checkText(text) {
    if (typeof text !== 'string')
        (0, edit_1.reject)(MSG.empty);
    const t = text.replace(/\r\n?/g, '\n');
    if (!t.trim())
        (0, edit_1.reject)(MSG.empty);
    if (t.length > comments_1.COMMENT_TEXT_MAX)
        (0, edit_1.reject)(MSG.tooLong);
    return t;
}
function sameId(a, b) {
    return !!a && a.toLowerCase() === b.toLowerCase();
}
/** 書き換え用の DOM の上でスレッドを探す。 */
function findThread(txn, threadId) {
    if (typeof threadId !== 'string' || !threadId)
        (0, edit_1.reject)(MSG.threadNotFound);
    if (threadId.startsWith('legacy:'))
        (0, edit_1.reject)(MSG.legacy);
    for (const s of (0, inherit_1.readPresentation)(txn).slides) {
        const found = (0, comments_2.threadElsOf)(txn, s.path, comments_2.ALL);
        if (!found || !found.cms.some((cm) => sameId((0, xml_1.attr)(cm, 'id'), threadId)))
            continue;
        const root = (0, package_1.rootEl)(txn.edit(found.part));
        const cm = (0, xml_1.kidsNamed)(root, 'p188:cm').find((x) => sameId((0, xml_1.attr)(x, 'id'), threadId));
        if (cm)
            return { slide: s, part: found.part, cm };
    }
    (0, edit_1.reject)(MSG.threadNotFound);
}
function requireMine(txn, el, me) {
    if ((0, comments_2.authorNameOf)(txn, (0, xml_1.attr)(el, 'authorId')) !== me)
        (0, edit_1.reject)(MSG.notMine);
}
function postIn(cm, postId) {
    if (postId === undefined || sameId((0, xml_1.attr)(cm, 'id'), postId))
        return cm;
    const r = (0, comments_2.repliesOf)(cm, comments_2.ALL).find((x) => sameId((0, xml_1.attr)(x, 'id'), postId));
    if (!r)
        (0, edit_1.reject)(MSG.postNotFound);
    return r;
}
function replaceBody(post, text) {
    const doc = post.ownerDocument;
    const old = (0, xml_1.kid)(post, 'p188:txBody');
    const body = (0, comments_2.commentTxBody)(doc, text);
    if (old)
        post.replaceChild(body, old);
    else {
        const ext = (0, xml_1.kid)(post, 'p188:extLst');
        if (ext)
            post.insertBefore(body, ext);
        else
            post.appendChild(body);
    }
}
// ───────────── 追加 ─────────────
function buildAnchor(txn, slide, input, doc) {
    const d = doc;
    const sldRoot = (0, package_1.rootEl)(txn.edit(slide.path));
    const cId = (0, creationIds_1.ensureSlideCreationId)(sldRoot, (0, comments_2.usedSlideCreationIds)(txn, slide.path));
    const docMk = (0, xml_1.create)(d, 'pc:docMk');
    const sldMk = (0, xml_1.create)(d, 'pc:sldMk', { cId: String(cId), sldId: slide.id });
    if (!input || typeof input !== 'object' || input.kind === 'slide') {
        const a = (0, xml_1.create)(d, 'pc:sldMkLst');
        a.appendChild(docMk);
        a.appendChild(sldMk);
        return a;
    }
    if (input.kind !== 'element' && input.kind !== 'text')
        (0, edit_1.reject)(MSG.badElement);
    const parsed = (0, elements_1.parseElementId)(String(input.elementId));
    if (!parsed || parsed.slideId !== slide.id)
        (0, edit_1.reject)(MSG.badElement);
    const f = (0, edit_1.findElement)(txn, input.elementId);
    const mkName = (0, comments_2.monikerFor)(f.ref.el);
    const cNvPr = (0, inherit_1.cNvPrOf)(f.ref.el);
    if (!mkName || !cNvPr)
        (0, edit_1.reject)(MSG.badElement);
    const mk = (0, xml_1.create)(d, mkName, { id: (0, xml_1.attr)(cNvPr, 'id') ?? '0', creationId: uniqueShapeCreationId(txn, slide.path, sldRoot, cNvPr) });
    const a = (0, xml_1.create)(d, input.kind === 'text' ? 'ac:txMkLst' : 'ac:deMkLst');
    a.appendChild(docMk);
    a.appendChild(sldMk);
    a.appendChild(mk);
    if (input.kind === 'text') {
        if (!(0, xml_1.is)(f.ref.el, 'p:sp') || !(0, xml_1.kid)(f.ref.el, 'p:txBody'))
            (0, edit_1.reject)(MSG.noText);
        const len = (0, comments_2.shapeText)(f.ref.el).length;
        const { start, length } = input;
        if (!Number.isInteger(start) || !Number.isInteger(length) || start < 0 || length < 0 || start + length > len)
            (0, edit_1.reject)(MSG.badRange);
        a.appendChild((0, xml_1.create)(d, 'ac:txMk', { cp: String(start), len: String(length) }));
    }
    return a;
}
/**
 * 図形の a16:creationId（なければ付ける）。同じスライドのほかの図形にも同じ値があれば（古い版の複製や、ほかのツールの pptx）、
 * この図形に新しい値を付け、この図形（cNvPr@id）を指していた既存のアンカーも新しい値に書き換える。
 */
function uniqueShapeCreationId(txn, slidePath, sldRoot, cNvPr) {
    const cur = (0, creationIds_1.readShapeCreationId)(cNvPr);
    if (!cur)
        return (0, creationIds_1.writeShapeCreationId)(cNvPr);
    const lower = cur.toLowerCase();
    const tree = (0, elements_1.spTreeOf)(sldRoot);
    const dup = (tree ? (0, xml_1.descendants)(tree, 'p:cNvPr') : []).some((c) => c !== cNvPr && (0, creationIds_1.readShapeCreationId)(c)?.toLowerCase() === lower);
    if (!dup)
        return cur;
    const next = (0, creationIds_1.writeShapeCreationId)(cNvPr);
    const id = (0, xml_1.attr)(cNvPr, 'id');
    const part = (0, comments_2.commentPartOf)(txn, slidePath);
    const root = part ? (0, comments_2.safeRoot)(txn, part) : undefined;
    if (part && root) {
        for (const cm of (0, xml_1.kidsNamed)((0, package_1.rootEl)(txn.edit(part)), 'p188:cm')) {
            const mk = (0, comments_2.shapeMonikerOf)((0, comments_2.anchorElOf)(cm));
            if (mk && (0, xml_1.attr)(mk, 'id') === id && (0, xml_1.attr)(mk, 'creationId')?.toLowerCase() === lower)
                (0, xml_1.setAttr)(mk, 'creationId', next);
        }
    }
    return next;
}
function addComment(txn, op, ctx) {
    const me = authorOf(ctx);
    const text = checkText(op.text);
    const s = (0, edit_1.findSlide)(txn, op.slideId);
    const slide = { id: s.slideId, path: s.slidePath };
    const part = (0, comments_2.ensureCommentPart)(txn, slide);
    const root = (0, package_1.rootEl)(txn.edit(part));
    const doc = root.ownerDocument;
    const anchor = buildAnchor(txn, slide, op.anchor, doc);
    const id = (0, creationIds_1.newGuid)();
    const cm = (0, xml_1.create)(doc, 'p188:cm', { id, authorId: (0, comments_2.ensureAuthor)(txn, me), created: nowOf(ctx) });
    cm.appendChild(anchor);
    cm.appendChild((0, comments_2.commentTxBody)(doc, text));
    root.appendChild(cm);
    return { changes: [], comment: { threadId: id, postId: id } };
}
// ───────────── 返信・編集・削除・解決 ─────────────
function replyComment(txn, op, ctx) {
    const me = authorOf(ctx);
    const text = checkText(op.text);
    const { cm } = findThread(txn, op.threadId);
    const doc = cm.ownerDocument;
    // 子の順：アンカー、p188:pos、p188:replyLst、p188:txBody、p188:extLst
    let lst = (0, xml_1.kid)(cm, 'p188:replyLst');
    if (!lst) {
        lst = (0, xml_1.create)(doc, 'p188:replyLst');
        const before = (0, xml_1.kid)(cm, 'p188:txBody') ?? (0, xml_1.kid)(cm, 'p188:extLst');
        if (before)
            cm.insertBefore(lst, before);
        else
            cm.appendChild(lst);
    }
    const id = (0, creationIds_1.newGuid)();
    const reply = (0, xml_1.create)(doc, 'p188:reply', { id, authorId: (0, comments_2.ensureAuthor)(txn, me), created: nowOf(ctx) });
    reply.appendChild((0, comments_2.commentTxBody)(doc, text));
    lst.appendChild(reply);
    // 解決済みのスレッドに返信したら再開する（Google スライドと同じ）
    (0, xml_1.setAttr)(cm, 'status', undefined);
    return { changes: [], comment: { threadId: (0, xml_1.attr)(cm, 'id'), postId: id } };
}
function editComment(txn, op, ctx) {
    const me = authorOf(ctx);
    const text = checkText(op.text);
    const { cm } = findThread(txn, op.threadId);
    const post = postIn(cm, op.postId);
    requireMine(txn, post, me);
    replaceBody(post, text);
    return { changes: [], comment: { threadId: (0, xml_1.attr)(cm, 'id'), postId: (0, xml_1.attr)(post, 'id') } };
}
function deleteComment(txn, op, ctx) {
    const me = authorOf(ctx);
    const { cm, slide } = findThread(txn, op.threadId);
    const post = postIn(cm, op.postId);
    requireMine(txn, post, me);
    if (post === cm) {
        // 最初のコメントを消すとスレッドごと消える（Google スライドと同じ）
        (0, xml_1.remove)(cm);
        (0, comments_2.removeCommentPartIfEmpty)(txn, slide.path);
        return { changes: [] };
    }
    const lst = post.parentNode;
    (0, xml_1.remove)(post);
    if ((0, xml_1.kidsNamed)(lst, 'p188:reply').length === 0)
        (0, xml_1.remove)(lst);
    return { changes: [], comment: { threadId: (0, xml_1.attr)(cm, 'id'), postId: (0, xml_1.attr)(cm, 'id') } };
}
function imageComment(txn, op, ctx) {
    const { cm } = findThread(txn, op.threadId);
    const post = postIn(cm, op.postId);
    requireMine(txn, post, authorOf(ctx));
    let images = (0, commentImages_2.readCommentImages)(post);
    if (op.type === 'attachCommentImage') {
        if (!(0, commentImages_1.validCommentImage)(op.dataUrl))
            (0, edit_1.reject)('添付できる画像は2MB以下・1600万画素以下のPNGまたはJPEGです（1辺8192pxまで）');
        if (images.length >= comments_1.COMMENT_IMAGES_MAX)
            (0, edit_1.reject)(`画像は1投稿につき${comments_1.COMMENT_IMAGES_MAX}枚まで添付できます`);
        if (typeof op.name !== 'string')
            (0, edit_1.reject)('画像のファイル名が正しくありません');
        images.push({ id: (0, creationIds_1.newGuid)(), name: (0, commentImages_1.commentImageName)(op.name), dataUrl: op.dataUrl });
    }
    else {
        if (!images.some(i => i.id === op.imageId))
            (0, edit_1.reject)('添付画像が見つかりません');
        images = images.filter(i => i.id !== op.imageId);
    }
    (0, commentImages_2.writeCommentImages)(post, images);
    return { changes: [], comment: { threadId: (0, xml_1.attr)(cm, 'id'), postId: (0, xml_1.attr)(post, 'id') } };
}
function setStatus(txn, threadId, resolved) {
    const { cm } = findThread(txn, threadId);
    (0, xml_1.setAttr)(cm, 'status', resolved ? 'resolved' : undefined);
    return { changes: [], comment: { threadId: (0, xml_1.attr)(cm, 'id'), postId: (0, xml_1.attr)(cm, 'id') } };
}
function applyCommentOp(txn, op, ctx) {
    switch (op.type) {
        case 'addComment':
            return addComment(txn, op, ctx);
        case 'replyComment':
            return replyComment(txn, op, ctx);
        case 'editComment':
            return editComment(txn, op, ctx);
        case 'attachCommentImage':
        case 'removeCommentImage':
            return imageComment(txn, op, ctx);
        case 'deleteComment':
            return deleteComment(txn, op, ctx);
        case 'resolveComment':
            return setStatus(txn, op.threadId, true);
        case 'reopenComment':
            return setStatus(txn, op.threadId, false);
        default:
            return undefined;
    }
}
function captureTextAnchors(txn, op) {
    if (op.type !== 'setText' || op.cell)
        return undefined;
    const parsed = (0, elements_1.parseElementId)(op.elementId);
    if (!parsed)
        return undefined;
    const slide = (0, inherit_1.readPresentation)(txn).slides.find((s) => s.id === parsed.slideId);
    if (!slide || !(0, comments_2.commentPartOf)(txn, slide.path))
        return undefined;
    const root = (0, comments_2.safeRoot)(txn, slide.path);
    const ref = root ? (0, elements_1.indexElements)((0, elements_1.spTreeOf)(root)).find((x) => x.key === parsed.key) : undefined;
    if (!ref)
        return undefined;
    return { slidePath: slide.path, key: parsed.key, before: (0, comments_2.shapeText)(ref.el) };
}
var comments_3 = require("../shared/comments.js");
Object.defineProperty(exports, "shiftRange", { enumerable: true, get: function () { return comments_3.shiftRange; } });
/** setText の後：同じ図形への文字の範囲のアンカー（ac:txMk の cp・len）を付け替える。 */
function shiftTextAnchors(txn, cap) {
    if (!cap)
        return;
    const root = (0, comments_2.safeRoot)(txn, cap.slidePath);
    const refs = root ? (0, elements_1.indexElements)((0, elements_1.spTreeOf)(root)) : [];
    const ref = refs.find((x) => x.key === cap.key);
    if (!ref)
        return;
    const after = (0, comments_2.shapeText)(ref.el);
    if (after === cap.before)
        return;
    const found = (0, comments_2.threadElsOf)(txn, cap.slidePath, comments_2.ALL);
    if (!found)
        return;
    const hits = found.cms.filter((cm) => {
        const a = (0, comments_2.anchorElOf)(cm);
        const mk = (0, xml_1.is)(a, 'ac:txMkLst') ? (0, comments_2.shapeMonikerOf)(a) : undefined;
        return !!mk && (0, comments_2.resolveMoniker)(refs, mk)?.el === ref.el;
    });
    if (!hits.length)
        return;
    const ids = new Set(hits.map((cm) => (0, xml_1.attr)(cm, 'id')));
    const editRoot = (0, package_1.rootEl)(txn.edit(found.part));
    for (const cm of (0, xml_1.kidsNamed)(editRoot, 'p188:cm')) {
        if (!ids.has((0, xml_1.attr)(cm, 'id')))
            continue;
        const tx = (0, comments_2.textMonikerOf)((0, comments_2.anchorElOf)(cm));
        if (!tx)
            continue;
        const r = (0, comments_1.shiftRange)(cap.before, after, Number((0, xml_1.attr)(tx, 'cp') ?? 0) || 0, Number((0, xml_1.attr)(tx, 'len') ?? 0) || 0);
        if (String(r.start) !== (0, xml_1.attr)(tx, 'cp'))
            (0, xml_1.setAttr)(tx, 'cp', String(r.start));
        if (String(r.length) !== (0, xml_1.attr)(tx, 'len'))
            (0, xml_1.setAttr)(tx, 'len', String(r.length));
    }
}
