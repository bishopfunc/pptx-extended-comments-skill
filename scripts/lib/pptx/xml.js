"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.EMU_MAX = exports.NS = void 0;
exports.parseXml = parseXml;
exports.serializeXml = serializeXml;
exports.isEl = isEl;
exports.is = is;
exports.isAny = isAny;
exports.kids = kids;
exports.kid = kid;
exports.kidsNamed = kidsNamed;
exports.path = path;
exports.descendants = descendants;
exports.attr = attr;
exports.numAttr = numAttr;
exports.boolAttr = boolAttr;
exports.setAttr = setAttr;
exports.create = create;
exports.remove = remove;
exports.removeKids = removeKids;
exports.textOf = textOf;
exports.fragment = fragment;
exports.importClean = importClean;
exports.insertChildInOrder = insertChildInOrder;
exports.ensureKid = ensureKid;
exports.checkChildOrder = checkChildOrder;
exports.orderedParents = orderedParents;
exports.clampCoord = clampCoord;
exports.clampSize = clampSize;
exports.emuCoord = emuCoord;
exports.emuSize = emuSize;
exports.normDeg = normDeg;
exports.rotUnits = rotUnits;
exports.emuInt = emuInt;
exports.stripInvalidXml = stripInvalidXml;
// 名前空間、DOM ヘルパー、子要素の順序表（設計書 7 章「書き込みの原則」）。
const xmldom_1 = require("../../vendor/xmldom.js");
exports.NS = {
    gssc: 'urn:google-slides-sync:comment-images:v1',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    p14: 'http://schemas.microsoft.com/office/powerpoint/2010/main',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    c: 'http://schemas.openxmlformats.org/drawingml/2006/chart',
    ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    /** 図形の作成 ID（a16:creationId）。 */
    a16: 'http://schemas.microsoft.com/office/drawing/2014/main',
    /** 新形式のコメント（PowerPoint 2018 以降）。 */
    p188: 'http://schemas.microsoft.com/office/powerpoint/2018/8/main',
    /** コメントのアンカー：スライド（pc:sldMk など）。 */
    pc: 'http://schemas.microsoft.com/office/powerpoint/2013/main/command',
    /** コメントのアンカー：図形・文字の範囲（ac:spMk、ac:txMk など）。 */
    ac: 'http://schemas.microsoft.com/office/drawing/2013/main/command',
    /** 旧形式のコメントのスレッド（p15:threadingInfo）。 */
    p15: 'http://schemas.microsoft.com/office/powerpoint/2012/main',
};
const ELEMENT_NODE = 1;
/** "a:xfrm" → [ns, local]。 */
function split(qname) {
    const i = qname.indexOf(':');
    if (i < 0)
        return ['', qname];
    const prefix = qname.slice(0, i);
    const ns = exports.NS[prefix];
    if (!ns)
        throw new Error(`unknown prefix ${prefix}`);
    return [ns, qname.slice(i + 1)];
}
const nameCache = new Map();
function q(qname) {
    let v = nameCache.get(qname);
    if (!v) {
        v = split(qname);
        nameCache.set(qname, v);
    }
    return v;
}
function parseXml(text) {
    const src = text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
    const doc = new xmldom_1.DOMParser({
        onError: (level, msg) => {
            if (level === 'fatalError' || level === 'error')
                throw new Error(`XML parse error: ${msg}`);
        },
    }).parseFromString(src, 'text/xml');
    return doc;
}
function serializeXml(doc) {
    return new xmldom_1.XMLSerializer().serializeToString(doc);
}
function isEl(n) {
    return !!n && n.nodeType === ELEMENT_NODE;
}
/** 要素が qname（"a:xfrm" など）か。 */
function is(el, qname) {
    if (!isEl(el))
        return false;
    const [ns, local] = q(qname);
    return el.localName === local && (el.namespaceURI ?? '') === ns;
}
function isAny(el, qnames) {
    return qnames.some((n) => is(el, n));
}
function kids(el) {
    const out = [];
    if (!el)
        return out;
    for (let n = el.firstChild; n; n = n.nextSibling)
        if (isEl(n))
            out.push(n);
    return out;
}
function kid(el, qname) {
    if (!el)
        return undefined;
    for (let n = el.firstChild; n; n = n.nextSibling)
        if (is(n, qname))
            return n;
    return undefined;
}
function kidsNamed(el, qname) {
    return kids(el).filter((k) => is(k, qname));
}
/** パスでたどる（"p:cSld/p:spTree"）。 */
function path(el, p) {
    let cur = el ?? undefined;
    for (const part of p.split('/')) {
        if (!cur)
            return undefined;
        cur = kid(cur, part);
    }
    return cur;
}
/** 子孫をすべて（文書順）。深い入れ子でもスタックを使い切らないよう、再帰しない。 */
function descendants(el, qname) {
    const out = [];
    // 次に見るノードの並び（後ろから取り出す）。子は逆順に積んで文書順にする
    const stack = [];
    const pushKids = (e) => {
        const ks = [];
        for (let n = e.firstChild; n; n = n.nextSibling)
            if (isEl(n))
                ks.push(n);
        for (let i = ks.length - 1; i >= 0; i--)
            stack.push(ks[i]);
    };
    pushKids(el);
    while (stack.length) {
        const n = stack.pop();
        if (!qname || is(n, qname))
            out.push(n);
        pushKids(n);
    }
    return out;
}
function attr(el, name) {
    if (!el)
        return undefined;
    if (name.includes(':')) {
        const [ns, local] = q(name);
        if (el.hasAttributeNS(ns, local))
            return el.getAttributeNS(ns, local) ?? undefined;
        return undefined;
    }
    return el.hasAttribute(name) ? (el.getAttribute(name) ?? undefined) : undefined;
}
function numAttr(el, name) {
    const v = attr(el, name);
    if (v === undefined || v === '')
        return undefined;
    const n = Number(v);
    return Number.isFinite(n) ? n : undefined;
}
function boolAttr(el, name) {
    const v = attr(el, name);
    if (v === undefined)
        return undefined;
    return v === '1' || v === 'true' || v === 'on';
}
function setAttr(el, name, value) {
    // XML に使えない文字は書かない（保存したパーツが XML として読めなくならないように。ADR-0019）
    if (typeof value === 'string')
        value = stripInvalidXml(value);
    if (name.includes(':')) {
        const [ns, local] = q(name);
        if (value === undefined)
            el.removeAttributeNS(ns, local);
        else
            el.setAttributeNS(ns, name, String(value));
        return;
    }
    if (value === undefined)
        el.removeAttribute(name);
    else
        el.setAttribute(name, String(value));
}
function create(doc, qname, attrs) {
    const [ns, local] = q(qname);
    const el = ns ? doc.createElementNS(ns, qname) : doc.createElement(local);
    if (attrs)
        for (const [k, v] of Object.entries(attrs))
            if (v !== undefined)
                setAttr(el, k, v);
    return el;
}
function remove(el) {
    if (el && el.parentNode)
        el.parentNode.removeChild(el);
}
function removeKids(el, qnames) {
    for (const k of kids(el))
        if (isAny(k, qnames))
            el.removeChild(k);
}
function textOf(el) {
    return el?.textContent ?? '';
}
/** XML 文字列の断片を要素にする（よく使う名前空間を宣言したうえで）。 */
function fragment(doc, xml) {
    const wrapped = `<w xmlns:a="${exports.NS.a}" xmlns:p="${exports.NS.p}" xmlns:r="${exports.NS.r}" xmlns:p14="${exports.NS.p14}" ` +
        `xmlns:mc="${exports.NS.mc}">${xml}</w>`;
    const tmp = parseXml(wrapped);
    const first = kids(tmp.documentElement)[0];
    if (!first)
        throw new Error('empty fragment');
    return importClean(doc, first);
}
/** 別の文書の要素を取り込み、親で宣言済みの冗長な xmlns 属性を取り除く。 */
function importClean(doc, el) {
    const imported = doc.importNode(el, true);
    stripRedundantNs(imported, doc.documentElement);
    return imported;
}
function stripRedundantNs(el, root) {
    if (!root)
        return;
    const declared = new Map();
    const attrs = root.attributes;
    for (let i = 0; i < attrs.length; i++) {
        const a = attrs.item(i);
        if (a && a.name.startsWith('xmlns:'))
            declared.set(a.name.slice(6), a.value);
    }
    const walk = (e) => {
        const at = e.attributes;
        const toRemove = [];
        for (let i = 0; i < at.length; i++) {
            const a = at.item(i);
            if (a && a.name.startsWith('xmlns:') && declared.get(a.name.slice(6)) === a.value)
                toRemove.push(a.name);
        }
        for (const n of toRemove)
            e.removeAttribute(n);
        for (const k of kids(e))
            walk(k);
    };
    walk(el);
}
// ---------------------------------------------------------------------------
// 子要素の順序表（ECMA-376）。同じ配列の要素は同じ順位（択一の要素）。
// ---------------------------------------------------------------------------
const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
const EFFECTS = ['a:effectLst', 'a:effectDag'];
const RPR_ORDER = [
    ['a:ln'],
    FILLS,
    EFFECTS,
    ['a:highlight'],
    ['a:uLnTx', 'a:uLn'],
    ['a:uFillTx', 'a:uFill'],
    ['a:latin'],
    ['a:ea'],
    ['a:cs'],
    ['a:sym'],
    ['a:hlinkClick'],
    ['a:hlinkMouseOver'],
    ['a:rtl'],
    ['a:extLst'],
];
const PPR_ORDER = [
    ['a:lnSpc'],
    ['a:spcBef'],
    ['a:spcAft'],
    ['a:buClrTx', 'a:buClr'],
    ['a:buSzTx', 'a:buSzPct', 'a:buSzPts'],
    ['a:buFontTx', 'a:buFont'],
    ['a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip'],
    ['a:tabLst'],
    ['a:defRPr'],
    ['a:extLst'],
];
const SPPR_ORDER = [
    ['a:xfrm'],
    ['a:custGeom', 'a:prstGeom'],
    FILLS,
    ['a:ln'],
    EFFECTS,
    ['a:scene3d'],
    ['a:sp3d'],
    ['a:extLst'],
];
const GRPSPPR_ORDER = [['a:xfrm'], FILLS, EFFECTS, ['a:scene3d'], ['a:extLst']];
const BODYPR_ORDER = [
    ['a:prstTxWarp'],
    ['a:noAutofit', 'a:normAutofit', 'a:spAutoFit'],
    ['a:scene3d'],
    ['a:sp3d'],
    ['a:flatTx'],
    ['a:extLst'],
];
const LN_ORDER = [
    ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill'],
    ['a:prstDash', 'a:custDash'],
    ['a:round', 'a:bevel', 'a:miter'],
    ['a:headEnd'],
    ['a:tailEnd'],
    ['a:extLst'],
];
const TXBODY_ORDER = [['a:bodyPr'], ['a:lstStyle'], ['a:p']];
const PARA_ORDER = [['a:pPr'], ['a:r', 'a:br', 'a:fld'], ['a:endParaRPr']];
const SHAPE_TREE_MEMBERS = ['p:sp', 'p:grpSp', 'p:graphicFrame', 'p:cxnSp', 'p:pic', 'p:contentPart', 'mc:AlternateContent'];
const LSTSTYLE_ORDER = [
    ['a:defPPr'],
    ['a:lvl1pPr'],
    ['a:lvl2pPr'],
    ['a:lvl3pPr'],
    ['a:lvl4pPr'],
    ['a:lvl5pPr'],
    ['a:lvl6pPr'],
    ['a:lvl7pPr'],
    ['a:lvl8pPr'],
    ['a:lvl9pPr'],
    ['a:extLst'],
];
const ORDER_TABLES = {
    'a:rPr': RPR_ORDER,
    'a:defRPr': RPR_ORDER,
    'a:endParaRPr': RPR_ORDER,
    'a:pPr': PPR_ORDER,
    'a:defPPr': PPR_ORDER,
    ...Object.fromEntries([1, 2, 3, 4, 5, 6, 7, 8, 9].map((i) => [`a:lvl${i}pPr`, PPR_ORDER])),
    'p:spPr': SPPR_ORDER,
    'p:grpSpPr': GRPSPPR_ORDER,
    'a:bodyPr': BODYPR_ORDER,
    'a:ln': LN_ORDER,
    'p:bgPr': [FILLS, EFFECTS, ['a:extLst']],
    'p:nvSpPr': [['p:cNvPr'], ['p:cNvSpPr'], ['p:nvPr']],
    'p:nvPicPr': [['p:cNvPr'], ['p:cNvPicPr'], ['p:nvPr']],
    'p:nvCxnSpPr': [['p:cNvPr'], ['p:cNvCxnSpPr'], ['p:nvPr']],
    'p:nvGrpSpPr': [['p:cNvPr'], ['p:cNvGrpSpPr'], ['p:nvPr']],
    'p:txBody': TXBODY_ORDER,
    'a:txBody': TXBODY_ORDER,
    'a:p': PARA_ORDER,
    'a:r': [['a:rPr'], ['a:t']],
    'a:fld': [['a:rPr'], ['a:pPr'], ['a:t']],
    'a:br': [['a:rPr']],
    'a:lstStyle': LSTSTYLE_ORDER,
    'a:tc': [['a:txBody'], ['a:tcPr'], ['a:extLst']],
    'p:sp': [['p:nvSpPr'], ['p:spPr'], ['p:style'], ['p:txBody'], ['p:extLst']],
    'p:pic': [['p:nvPicPr'], ['p:blipFill'], ['p:spPr'], ['p:style'], ['p:extLst']],
    'p:cxnSp': [['p:nvCxnSpPr'], ['p:spPr'], ['p:style'], ['p:extLst']],
    'p:grpSp': [['p:nvGrpSpPr'], ['p:grpSpPr'], SHAPE_TREE_MEMBERS, ['p:extLst']],
    'p:spTree': [['p:nvGrpSpPr'], ['p:grpSpPr'], SHAPE_TREE_MEMBERS, ['p:extLst']],
    'p:sld': [['p:cSld'], ['p:clrMapOvr'], ['p:transition'], ['p:timing'], ['p:extLst']],
    'p:notes': [['p:cSld'], ['p:clrMapOvr'], ['p:extLst']],
    'p:cSld': [['p:bg'], ['p:spTree'], ['p:custDataLst'], ['p:controls'], ['p:extLst']],
    'p:presentation': [
        ['p:sldMasterIdLst'],
        ['p:notesMasterIdLst'],
        ['p:handoutMasterIdLst'],
        ['p:sldIdLst'],
        ['p:sldSz'],
        ['p:notesSz'],
        ['p:smartTags'],
        ['p:embeddedFontLst'],
        ['p:custShowLst'],
        ['p:photoAlbum'],
        ['p:custDataLst'],
        ['p:kinsoku'],
        ['p:defaultTextStyle'],
        ['p:modifyVerifier'],
        ['p:extLst'],
    ],
};
function tableKey(el) {
    const ns = el.namespaceURI ?? '';
    const prefix = Object.keys(exports.NS).find((p) => exports.NS[p] === ns);
    if (!prefix)
        return undefined;
    return `${prefix}:${el.localName}`;
}
function rankIn(table, el) {
    for (let i = 0; i < table.length; i++)
        if (isAny(el, table[i]))
            return i;
    return -1;
}
/** 親の種類ごとの順序表に従って子要素を挿入する。表にない親なら末尾に追加。 */
function insertChildInOrder(parent, child) {
    const key = tableKey(parent);
    const table = key ? ORDER_TABLES[key] : undefined;
    if (!table) {
        parent.appendChild(child);
        return child;
    }
    const rank = rankIn(table, child);
    if (rank < 0) {
        parent.appendChild(child);
        return child;
    }
    for (const k of kids(parent)) {
        const r = rankIn(table, k);
        if (r > rank) {
            parent.insertBefore(child, k);
            return child;
        }
    }
    parent.appendChild(child);
    return child;
}
/** 子要素を取得し、なければ順序どおりに作る。 */
function ensureKid(parent, qname, attrs) {
    const existing = kid(parent, qname);
    if (existing)
        return existing;
    return insertChildInOrder(parent, create(parent.ownerDocument, qname, attrs));
}
/** 子要素の順序が表どおりかを検査する（integrity 用）。問題があれば説明を返す。 */
function checkChildOrder(el) {
    const key = tableKey(el);
    const table = key ? ORDER_TABLES[key] : undefined;
    if (!table)
        return undefined;
    let last = -1;
    let lastName = '';
    for (const k of kids(el)) {
        const r = rankIn(table, k);
        if (r < 0)
            continue;
        if (r < last)
            return `${key} の子要素 ${k.tagName} が ${lastName} の後にある`;
        last = r;
        lastName = k.tagName;
    }
    return undefined;
}
function orderedParents() {
    return Object.keys(ORDER_TABLES);
}
// ───────────── 数値の範囲（ADR-0019）─────────────
// pptx の座標と大きさは EMU の 32 ビット整数。細工した値（1e308 など）を読み書きしても、範囲の外の値や指数表記を書かない。
exports.EMU_MAX = 2 ** 31 - 1;
/** 読んだ座標を範囲に収める（丸めない）。NaN は 0。 */
function clampCoord(n) {
    if (Number.isNaN(n))
        return 0;
    return Math.max(-exports.EMU_MAX, Math.min(exports.EMU_MAX, n));
}
/** 読んだ大きさを 0 以上、範囲の内側に収める（丸めない）。 */
function clampSize(n) {
    return Math.max(0, clampCoord(n));
}
/** 書く座標（整数）。 */
function emuCoord(n) {
    return Math.round(clampCoord(n));
}
/** 書く大きさ（0 以上の整数）。 */
function emuSize(n) {
    return Math.round(clampSize(n));
}
/** 角度（度）を 0 以上 360 未満にする。有限でなければ 0。 */
function normDeg(r) {
    if (!Number.isFinite(r))
        return 0;
    const v = ((r % 360) + 360) % 360;
    return v >= 360 ? 0 : v;
}
/** 書く回転（60000 分の 1 度の整数）。 */
function rotUnits(deg) {
    return Math.round(normDeg(deg) * 60000) % 21600000;
}
function emuInt(n) {
    return String(emuCoord(n));
}
// ───────────── XML に使えない文字（ADR-0019）─────────────
const INVALID_XML_CHARS = /[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/g;
/** XML 1.0 に使えない文字（制御文字、U+FFFE・U+FFFF、対のないサロゲート）を除く。 */
function stripInvalidXml(s) {
    return s.replace(INVALID_XML_CHARS, '');
}
