"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.CLIP_MSG = void 0;
exports.copySlides = copySlides;
exports.applySlideClipboardOp = applySlideClipboardOp;
const own_1 = require("../shared/own.js");
const edit_1 = require("./edit.js");
const comments_1 = require("./comments.js");
const creationIds_1 = require("./creationIds.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const layouts_1 = require("./layouts.js");
const package_1 = require("./package.js");
const parse_1 = require("./parse.js");
const slides_1 = require("./slides.js");
const xml_1 = require("./xml.js");
exports.CLIP_MSG = {
    forbiddenCopy: 'グラフ、SmartArt、OLE を含むスライドは、このバージョンではコピーできません',
    noSlides: '貼り付けるスライドがありません',
};
function layoutName(txn, layoutPath) {
    if (!layoutPath || !txn.has(layoutPath))
        return undefined;
    return (0, xml_1.attr)((0, xml_1.kid)((0, package_1.rootEl)(txn.dom(layoutPath)), 'p:cSld'), 'name');
}
/** spTree の直下のプレースホルダーのうち、a:xfrm を持たないもの。 */
function inheritingPlaceholders(root) {
    return (0, xml_1.kids)((0, elements_1.spTreeOf)(root)).filter((k) => (0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS) && !!(0, inherit_1.phOf)(k) && !(0, inherit_1.xfrmElOf)(k));
}
/**
 * スライドをコピーする（プレゼンテーションの順に並べる）。XML、XML が参照する関係、画像・メディアの中身、ノートの本文、
 * レイアウトのパスと名前を持つ。ノート・コメント・レイアウトの関係は持たない（貼り付けで作り直す）。
 */
function copySlides(txn, slideIds) {
    const pres = (0, inherit_1.readPresentation)(txn);
    for (const id of slideIds)
        (0, edit_1.findSlide)(txn, id);
    const wanted = new Set(slideIds);
    const payload = { slides: [], parts: {} };
    for (const s of pres.slides.filter((x) => wanted.has(x.id))) {
        const root = (0, package_1.rootEl)(txn.dom(s.path));
        const rels = (0, package_1.readRels)(txn, s.path);
        if ((0, elements_1.containsForbidden)(root) || rels.some((r) => slides_1.FORBIDDEN_REL_HINTS.some((h) => r.type.includes(h))))
            (0, edit_1.reject)(exports.CLIP_MSG.forbiddenCopy);
        const used = new Set((0, edit_1.rAttrs)(root).map((a) => a.value));
        const item = {
            xml: txn.text(s.path),
            rels: rels
                .filter((r) => used.has(r.id))
                .map((r) => ({ rId: r.id, type: r.type, target: r.target, external: r.external })),
            notes: (0, parse_1.readNotes)(txn, s.path),
            layout: {},
            placeholderXfrms: {},
        };
        const lp = (0, inherit_1.layoutPathOf)(txn, s.path);
        if (lp)
            item.layout.path = lp;
        const ln = layoutName(txn, lp);
        if (ln !== undefined)
            item.layout.name = ln;
        const ctx = (0, inherit_1.buildSlideCtx)(txn, s.path);
        for (const el of inheritingPlaceholders(root)) {
            const cid = (0, xml_1.attr)((0, inherit_1.cNvPrOf)(el), 'id');
            if (cid && !Object.hasOwn(item.placeholderXfrms, cid))
                item.placeholderXfrms[cid] = (0, parse_1.resolvedOwnXfrm)(ctx, el);
        }
        for (const r of item.rels) {
            if (r.external || !edit_1.COPY_PART_TYPES.includes(r.type) || !txn.has(r.target) || Object.hasOwn(payload.parts, r.target))
                continue;
            payload.parts[r.target] = txn.bytes(r.target);
        }
        payload.slides.push(item);
    }
    return payload;
}
/**
 * 貼り付け先のレイアウト：同じパスで同じ名前 → 同じ名前 → 同じパス → 最初のレイアウト。
 * exact は、同じパスで同じ名前のレイアウトが見つかったか（見つからなければプレースホルダーの位置を書き込む）。
 */
function pickLayout(txn, layouts, fallback, src) {
    const srcPath = typeof src.path === 'string' && layouts.includes(src.path) ? src.path : undefined;
    if (srcPath && (src.name === undefined || layoutName(txn, srcPath) === src.name))
        return { path: srcPath, exact: true };
    const byName = src.name !== undefined ? layouts.find((l) => layoutName(txn, l) === src.name) : undefined;
    return { path: byName ?? srcPath ?? fallback ?? layouts[0], exact: false };
}
/**
 * コピーしたスライドを afterSlideId の後ろ（null なら先頭）に、コピーした順で新しいパーツとして作る。
 * - レイアウト：pickLayout で選ぶ。
 * - 画像・メディア：同じ中身のパーツがあれば使い、なければ新しい名前で追加する（pasteElements と同じ）。
 * - 外部へのリンク：http(s)・mailto のハイパーリンクだけを持ち込む。スライド間リンク、その他の関係（タグなど）は外し、
 *   それを使っている r:id（a:hlinkClick などは要素ごと）を取り除く。
 * - ノート：本文だけを、新しいノートスライドに書く（setNotes と同じ）。コメントは持ち込まない。
 */
function pasteSlides(txn, op, clip) {
    if (!clip || !Array.isArray(clip.slides) || clip.slides.length === 0)
        (0, edit_1.reject)(exports.CLIP_MSG.noSlides);
    const pres = (0, inherit_1.readPresentation)(txn);
    let afterId;
    if (op.afterSlideId !== null)
        afterId = (0, edit_1.findSlide)(txn, op.afterSlideId).slideId;
    const layouts = (0, layouts_1.allLayoutPaths)(txn, pres);
    const fallbackLayout = (0, slides_1.firstLayout)(txn, pres);
    const resolvePart = (0, edit_1.partResolver)(txn, clip.parts);
    let first;
    for (const item of clip.slides) {
        const layout = pickLayout(txn, layouts, fallbackLayout, item.layout);
        if (!layout.path || !txn.has(layout.path))
            (0, edit_1.reject)(slides_1.SLIDE_MSG.noLayout);
        const slidePath = `ppt/slides/slide${(0, slides_1.nextPartNumber)(txn, 'ppt/slides', 'slide')}.xml`;
        const root = (0, package_1.rootEl)(txn.createXml(slidePath, item.xml));
        if (!(0, xml_1.is)(root, 'p:sld') || !(0, elements_1.spTreeOf)(root))
            (0, edit_1.reject)(edit_1.MSG.noClipboard);
        if ((0, elements_1.containsForbidden)(root))
            (0, edit_1.reject)(exports.CLIP_MSG.forbiddenCopy);
        // コメントは持ち込まない（p188:commentRel を拡張ごと外す）。スライドの作成 ID は、同じファイルに貼っても重ならないよう振り直す
        (0, creationIds_1.dropCommentRefs)(root);
        if ((0, creationIds_1.readSlideCreationId)(root) !== undefined)
            (0, creationIds_1.writeSlideCreationId)(root, undefined, (0, comments_1.usedSlideCreationIds)(txn, slidePath));
        const relMap = new Map();
        for (const a of (0, edit_1.rAttrs)(root)) {
            let nid = relMap.get(a.value);
            if (nid === undefined) {
                const rel = item.rels.find((x) => x.rId === a.value);
                if (!rel)
                    nid = null;
                else if (rel.external)
                    nid = (0, edit_1.isSafeExternalRel)(rel) ? txn.addRel(slidePath, rel.type, rel.target, true) : null;
                else if (edit_1.COPY_PART_TYPES.includes(rel.type)) {
                    const p = resolvePart(rel.target, rel.type);
                    nid = p ? txn.addRel(slidePath, rel.type, p) : null;
                }
                else
                    nid = null;
                relMap.set(a.value, nid);
            }
            if (nid === null)
                (0, edit_1.removeRAttr)(a.owner, a.name);
            else
                a.owner.setAttributeNS(xml_1.NS.r, a.name, nid);
        }
        txn.addRel(slidePath, package_1.REL.slideLayout, layout.path);
        // 別のレイアウトになったら、継承していた位置を書き込んでプレースホルダーが動かないようにする
        if (!layout.exact) {
            for (const el of inheritingPlaceholders(root)) {
                const x = (0, own_1.own)(item.placeholderXfrms ?? {}, (0, xml_1.attr)((0, inherit_1.cNvPrOf)(el), 'id'));
                if (x)
                    (0, edit_1.writeXfrm)(el, x, false);
            }
        }
        const newId = (0, slides_1.registerSlide)(txn, pres, slidePath, afterId);
        if (typeof item.notes === 'string' && item.notes !== '')
            (0, slides_1.writeNotes)(txn, slidePath, item.notes);
        first ??= newId;
        afterId = newId;
    }
    return { changes: [], selectSlide: first };
}
function applySlideClipboardOp(txn, op, clip) {
    if (op.type === 'pasteSlides')
        return pasteSlides(txn, op, clip);
    return undefined;
}
