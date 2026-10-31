"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.URI_SHAPE_CREATION_ID = exports.URI_COMMENT_REL = exports.URI_SLIDE_CREATION_ID = void 0;
exports.newGuid = newGuid;
exports.isGuid = isGuid;
exports.newSlideCreationId = newSlideCreationId;
exports.readSlideCreationId = readSlideCreationId;
exports.writeSlideCreationId = writeSlideCreationId;
exports.ensureSlideCreationId = ensureSlideCreationId;
exports.readShapeCreationId = readShapeCreationId;
exports.writeShapeCreationId = writeShapeCreationId;
exports.ensureShapeCreationId = ensureShapeCreationId;
exports.dropShapeCreationIds = dropShapeCreationIds;
exports.dropCommentRefs = dropCommentRefs;
// スライドと図形の作成 ID（p14:creationId、a16:creationId）。新形式のコメントのアンカーが使う（設計書 7 章「レビューコメント」）。
// どちらも PowerPoint が付ける識別子で、Google スライドと LibreOffice は書き出さない。必要になったとき（コメントを付けるとき）だけ付け、
// 複製した図形・スライドには同じ値を引き継がない。
const node_crypto_1 = require("node:crypto");
const xml_1 = require("./xml.js");
/** p:cSld/p:extLst/p:ext の uri。中に p14:creationId val（符号なし 32 ビット）。 */
exports.URI_SLIDE_CREATION_ID = '{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}';
/** p:sld/p:extLst/p:ext の uri。中に p188:commentRel r:id（新形式のコメントのパーツへの関係）。 */
exports.URI_COMMENT_REL = '{6950BFC3-D8DA-4A85-94F7-54DA5524770B}';
/** p:cNvPr/a:extLst/a:ext の uri。中に a16:creationId id（GUID）。 */
exports.URI_SHAPE_CREATION_ID = '{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}';
/** `{XXXXXXXX-XXXX-XXXX-XXXX-XXXXXXXXXXXX}`（大文字）。 */
function newGuid() {
    return `{${(0, node_crypto_1.randomUUID)().toUpperCase()}}`;
}
const GUID_RE = /^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/;
function isGuid(v) {
    return !!v && GUID_RE.test(v);
}
/** 0 でない符号なし 32 ビット整数。used と重ならない値。 */
function newSlideCreationId(used = []) {
    const taken = new Set(used);
    for (;;) {
        const n = (0, node_crypto_1.randomBytes)(4).readUInt32BE(0);
        if (n !== 0 && !taken.has(n))
            return n;
    }
}
function extIn(extLst, uri) {
    return (0, xml_1.kids)(extLst).find((e) => e.localName === 'ext' && (0, xml_1.attr)(e, 'uri') === uri);
}
/** p:cSld/p:extLst の中の p14:creationId 要素。 */
function slideCreationIdEl(sldRoot) {
    const ext = extIn((0, xml_1.kid)((0, xml_1.kid)(sldRoot, 'p:cSld'), 'p:extLst'), exports.URI_SLIDE_CREATION_ID);
    return (0, xml_1.kids)(ext).find((e) => (0, xml_1.is)(e, 'p14:creationId'));
}
/** スライドの p14:creationId（なければ undefined）。 */
function readSlideCreationId(sldRoot) {
    const v = Number((0, xml_1.attr)(slideCreationIdEl(sldRoot), 'val'));
    return Number.isInteger(v) && v > 0 && v <= 0xffffffff ? v : undefined;
}
/**
 * スライドの p14:creationId を書く（p:cSld/p:extLst の中。p:cSld の最後の子）。val を省略したら新しい値を作る。
 * 既にあれば書き換える。書いた値を返す。
 */
function writeSlideCreationId(sldRoot, val, used = []) {
    const doc = sldRoot.ownerDocument;
    const cSld = (0, xml_1.kid)(sldRoot, 'p:cSld');
    if (!cSld)
        throw new Error('p:cSld がありません');
    const v = val ?? newSlideCreationId(used);
    const existing = slideCreationIdEl(sldRoot);
    if (existing) {
        existing.setAttribute('val', String(v));
        return v;
    }
    let extLst = (0, xml_1.kid)(cSld, 'p:extLst');
    if (!extLst) {
        extLst = (0, xml_1.create)(doc, 'p:extLst');
        cSld.appendChild(extLst);
    }
    const ext = extIn(extLst, exports.URI_SLIDE_CREATION_ID) ?? extLst.appendChild((0, xml_1.create)(doc, 'p:ext', { uri: exports.URI_SLIDE_CREATION_ID }));
    const el = doc.createElementNS(xml_1.NS.p14, 'p14:creationId');
    el.setAttribute('val', String(v));
    ext.appendChild(el);
    return v;
}
/** スライドの p14:creationId（なければ付ける）。 */
function ensureSlideCreationId(sldRoot, used = []) {
    return readSlideCreationId(sldRoot) ?? writeSlideCreationId(sldRoot, undefined, used);
}
function shapeCreationIdEl(cNvPr) {
    const ext = extIn((0, xml_1.kid)(cNvPr, 'a:extLst'), exports.URI_SHAPE_CREATION_ID);
    return (0, xml_1.kids)(ext).find((e) => (0, xml_1.is)(e, 'a16:creationId'));
}
/** 図形（p:cNvPr）の a16:creationId（なければ undefined）。 */
function readShapeCreationId(cNvPr) {
    if (!cNvPr)
        return undefined;
    const v = (0, xml_1.attr)(shapeCreationIdEl(cNvPr), 'id');
    return isGuid(v) ? v : undefined;
}
/** 図形の a16:creationId を書く（id を省略したら新しい GUID）。既にあれば書き換える。 */
function writeShapeCreationId(cNvPr, id = newGuid()) {
    const doc = cNvPr.ownerDocument;
    const existing = shapeCreationIdEl(cNvPr);
    if (existing) {
        existing.setAttribute('id', id);
        return id;
    }
    let extLst = (0, xml_1.kid)(cNvPr, 'a:extLst');
    if (!extLst) {
        extLst = (0, xml_1.create)(doc, 'a:extLst');
        cNvPr.appendChild(extLst);
    }
    const ext = extIn(extLst, exports.URI_SHAPE_CREATION_ID) ?? extLst.appendChild((0, xml_1.create)(doc, 'a:ext', { uri: exports.URI_SHAPE_CREATION_ID }));
    const el = doc.createElementNS(xml_1.NS.a16, 'a16:creationId');
    el.setAttribute('id', id);
    ext.appendChild(el);
    return id;
}
/** 図形の a16:creationId（なければ付ける）。 */
function ensureShapeCreationId(cNvPr) {
    return readShapeCreationId(cNvPr) ?? writeShapeCreationId(cNvPr);
}
/** 部分木の中の a16:creationId を取り除く（複製・貼り付けた図形に、元の図形の ID を引き継がないため）。 */
function dropShapeCreationIds(el) {
    for (const c of [el, ...(0, xml_1.descendants)(el)].filter((e) => (0, xml_1.is)(e, 'p:cNvPr'))) {
        const extLst = (0, xml_1.kid)(c, 'a:extLst');
        for (const ext of (0, xml_1.kidsNamed)(extLst, 'a:ext'))
            if ((0, xml_1.attr)(ext, 'uri') === exports.URI_SHAPE_CREATION_ID)
                (0, xml_1.remove)(ext);
        if (extLst && (0, xml_1.kids)(extLst).length === 0)
            (0, xml_1.remove)(extLst);
    }
}
/**
 * スライドの p:extLst から新形式のコメントへの参照（p188:commentRel の拡張）を取り除く。空になった p:extLst も。
 * コピーしたスライドを貼り付けるとき（コメントは持ち込まない）に使う。
 */
function dropCommentRefs(sldRoot) {
    const extLst = (0, xml_1.kid)(sldRoot, 'p:extLst');
    for (const ext of (0, xml_1.kidsNamed)(extLst, 'p:ext'))
        if ((0, xml_1.attr)(ext, 'uri') === exports.URI_COMMENT_REL)
            (0, xml_1.remove)(ext);
    if (extLst && (0, xml_1.kids)(extLst).length === 0)
        (0, xml_1.remove)(extLst);
}
