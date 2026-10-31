"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MAX_GROUP_DEPTH = void 0;
exports.indexElements = indexElements;
exports.spTreeOf = spTreeOf;
exports.parseElementId = parseElementId;
exports.nextShapeId = nextShapeId;
exports.graphicKind = graphicKind;
exports.containsForbidden = containsForbidden;
exports.toSlideXfrm = toSlideXfrm;
exports.ownXfrm = ownXfrm;
const inherit_1 = require("./inherit.js");
const xml_1 = require("./xml.js");
/** グループの入れ子の深さの上限。これより深いグループは対応していない要素として扱う。 */
exports.MAX_GROUP_DEPTH = 64;
/** spTree（またはグループ）の下の図形を文書順に列挙する（グループの中も含む）。 */
function indexElements(spTree) {
    const out = [];
    const counts = new Map();
    const walk = (parent, groups) => {
        for (const k of (0, xml_1.kids)(parent)) {
            if (!(0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS))
                continue;
            const cnvId = (0, xml_1.attr)((0, inherit_1.cNvPrOf)(k), 'id') ?? '0';
            const n = (counts.get(cnvId) ?? 0) + 1;
            counts.set(cnvId, n);
            out.push({ el: k, cnvId, key: n === 1 ? cnvId : `${cnvId}#${n}`, groups });
            // 深すぎるグループの中は列挙しない（そのグループは対応していない要素になる。ADR-0019）
            if ((0, xml_1.is)(k, 'p:grpSp') && groups.length < exports.MAX_GROUP_DEPTH)
                walk(k, [k, ...groups]);
        }
    };
    if (spTree)
        walk(spTree, []);
    return out;
}
function spTreeOf(root) {
    return (0, xml_1.path)(root, 'p:cSld/p:spTree');
}
function parseElementId(id) {
    const i = id.indexOf(':');
    if (i <= 0)
        return undefined;
    return { slideId: id.slice(0, i), key: id.slice(i + 1) };
}
/** スライドの中のすべての cNvPr@id の最大値 + 1。 */
function nextShapeId(root) {
    const cSld = (0, xml_1.kid)(root, 'p:cSld');
    let max = 0;
    if (cSld) {
        for (const c of (0, xml_1.descendants)(cSld, 'p:cNvPr')) {
            const n = Number((0, xml_1.attr)(c, 'id'));
            if (Number.isFinite(n) && n > max)
                max = n;
        }
    }
    return max + 1;
}
/** 中身の種類（graphicFrame の graphicData の uri など）。 */
function graphicKind(el) {
    const target = (0, xml_1.is)(el, 'mc:AlternateContent') ? (0, inherit_1.altContentMain)(el) : el;
    if (!target || !(0, xml_1.is)(target, 'p:graphicFrame'))
        return undefined;
    const gd = (0, xml_1.path)(target, 'a:graphic/a:graphicData');
    const uri = (0, xml_1.attr)(gd, 'uri') ?? '';
    if (uri.endsWith('/table'))
        return 'table';
    if (uri.includes('/chart') || uri.includes('chartex') || uri.includes('/chartml'))
        return 'chart';
    if (uri.includes('/diagram'))
        return 'diagram';
    if (uri.includes('/ole'))
        return 'ole';
    return 'other';
}
/** 複製・貼り付けできない要素（グラフ、SmartArt、OLE）を含むか。 */
function containsForbidden(el) {
    const all = [el, ...(0, xml_1.descendants)(el)];
    for (const e of all) {
        if ((0, xml_1.is)(e, 'p:graphicFrame') || (0, xml_1.is)(e, 'mc:AlternateContent')) {
            const k = graphicKind(e);
            if (k === 'chart' || k === 'diagram' || k === 'ole')
                return true;
        }
        if ((0, xml_1.is)(e, 'a:graphicData')) {
            const uri = (0, xml_1.attr)(e, 'uri') ?? '';
            if (uri.includes('/chart') || uri.includes('/diagram') || uri.includes('/ole') || uri.includes('chartex'))
                return true;
        }
        if ((0, xml_1.is)(e, 'p:oleObj'))
            return true;
    }
    return false;
}
function groupFrame(g) {
    const x = (0, xml_1.path)(g, 'p:grpSpPr/a:xfrm');
    const base = (0, inherit_1.readXfrmEl)(x);
    const chOff = (0, xml_1.kid)(x, 'a:chOff');
    const chExt = (0, xml_1.kid)(x, 'a:chExt');
    const n = (e, a, d) => {
        const v = Number((0, xml_1.attr)(e, a));
        return Number.isFinite(v) && (0, xml_1.attr)(e, a) !== undefined ? v : d;
    };
    return {
        off: { x: base?.x ?? 0, y: base?.y ?? 0 },
        ext: { cx: base?.cx ?? 0, cy: base?.cy ?? 0 },
        chOff: { x: (0, xml_1.clampCoord)(n(chOff, 'x', base?.x ?? 0)), y: (0, xml_1.clampCoord)(n(chOff, 'y', base?.y ?? 0)) },
        chExt: { cx: (0, xml_1.clampSize)(n(chExt, 'cx', base?.cx ?? 0)), cy: (0, xml_1.clampSize)(n(chExt, 'cy', base?.cy ?? 0)) },
        rot: base?.rot ?? 0,
        flipH: base?.flipH ?? false,
        flipV: base?.flipV ?? false,
    };
}
/** グループの子の座標（子座標系）を、グループの親の座標系に変換する。 */
function mapThroughGroup(x, g) {
    // 子の座標系の大きさが 1 EMU 未満（0 や 1e-300 など）なら拡大しない（極端な倍率で値が壊れないように）
    const sx = g.chExt.cx >= 1 ? g.ext.cx / g.chExt.cx : 1;
    const sy = g.chExt.cy >= 1 ? g.ext.cy / g.chExt.cy : 1;
    let nx = g.off.x + (x.x - g.chOff.x) * sx;
    let ny = g.off.y + (x.y - g.chOff.y) * sy;
    const ncx = x.cx * sx;
    const ncy = x.cy * sy;
    let rot = x.rot;
    let flipH = x.flipH;
    let flipV = x.flipV;
    const gcx = g.off.x + g.ext.cx / 2;
    const gcy = g.off.y + g.ext.cy / 2;
    if (g.flipH) {
        const c = nx + ncx / 2;
        nx = 2 * gcx - c - ncx / 2;
        flipH = !flipH;
        rot = -rot;
    }
    if (g.flipV) {
        const c = ny + ncy / 2;
        ny = 2 * gcy - c - ncy / 2;
        flipV = !flipV;
        rot = -rot;
    }
    if (g.rot) {
        const rad = (g.rot * Math.PI) / 180;
        const cx = nx + ncx / 2 - gcx;
        const cy = ny + ncy / 2 - gcy;
        const rx = cx * Math.cos(rad) - cy * Math.sin(rad);
        const ry = cx * Math.sin(rad) + cy * Math.cos(rad);
        nx = gcx + rx - ncx / 2;
        ny = gcy + ry - ncy / 2;
        rot += g.rot;
    }
    return { x: (0, xml_1.emuCoord)(nx), y: (0, xml_1.emuCoord)(ny), cx: (0, xml_1.emuSize)(ncx), cy: (0, xml_1.emuSize)(ncy), rot: (0, xml_1.normDeg)(rot), flipH, flipV };
}
/** groups は内側から外側の順。 */
function toSlideXfrm(x, groups) {
    let cur = x;
    for (const g of groups)
        cur = mapThroughGroup(cur, groupFrame(g));
    return cur;
}
function ownXfrm(el) {
    return (0, inherit_1.readXfrmEl)((0, inherit_1.xfrmElOf)(el));
}
