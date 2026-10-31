"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ARRANGE_MSG = void 0;
exports.applyArrangeOp = applyArrangeOp;
const edit_1 = require("./edit.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const parse_1 = require("./parse.js");
const templates_1 = require("./templates.js");
const xml_1 = require("./xml.js");
exports.ARRANGE_MSG = {
    groupTooFew: 'グループ化するには 2 つ以上の要素を選んでください',
    groupSlides: '別々のスライドの要素はグループ化できません',
    groupPlaceholder: 'プレースホルダーはグループ化できません',
    notGroup: 'グループではない要素はグループ解除できません',
    badInset: 'テキストの余白の指定が正しくありません',
    badAnchor: 'テキストの配置の指定が正しくありません',
};
/** a:off / a:ext の上限（ST_Coordinate の範囲）。 */
const MAX_COORD = 51206400;
/** 回転を考慮した外接矩形。 */
function bounds(x) {
    if (!x.rot)
        return { x1: x.x, y1: x.y, x2: x.x + x.cx, y2: x.y + x.cy };
    const a = (x.rot * Math.PI) / 180;
    const cos = Math.abs(Math.cos(a));
    const sin = Math.abs(Math.sin(a));
    const w = x.cx * cos + x.cy * sin;
    const h = x.cx * sin + x.cy * cos;
    const cx = x.x + x.cx / 2;
    const cy = x.y + x.cy / 2;
    return { x1: cx - w / 2, y1: cy - h / 2, x2: cx + w / 2, y2: cy + h / 2 };
}
function selectIdsOf(root, slideId, els) {
    const set = new Set(els);
    return (0, elements_1.indexElements)((0, elements_1.spTreeOf)(root))
        .filter((r) => set.has(r.el))
        .map((r) => `${slideId}:${r.key}`);
}
// ---------------------------------------------------------------------------
// グループ化
// ---------------------------------------------------------------------------
function groupElements(txn, op) {
    const found = op.elementIds.map((id) => (0, edit_1.findElement)(txn, id));
    const first = found[0];
    if (!first)
        (0, edit_1.reject)(exports.ARRANGE_MSG.groupTooFew);
    for (const f of found) {
        if (f.slidePath !== first.slidePath)
            (0, edit_1.reject)(exports.ARRANGE_MSG.groupSlides);
        if (f.ref.groups.length)
            (0, edit_1.reject)(edit_1.MSG.groupChild);
        if ((0, inherit_1.phOf)(f.ref.el))
            (0, edit_1.reject)(exports.ARRANGE_MSG.groupPlaceholder);
    }
    const selected = new Set(found.map((f) => f.ref.el));
    if (selected.size < 2)
        (0, edit_1.reject)(exports.ARRANGE_MSG.groupTooFew);
    const root = first.root;
    const tree = (0, elements_1.spTreeOf)(root);
    // 文書順（重なり順）に並べる
    const members = (0, xml_1.kids)(tree).filter((k) => selected.has(k));
    const ctx = (0, inherit_1.buildSlideCtx)(txn, first.slidePath);
    let x1 = Infinity;
    let y1 = Infinity;
    let x2 = -Infinity;
    let y2 = -Infinity;
    for (const m of members) {
        const b = bounds((0, parse_1.resolvedOwnXfrm)(ctx, m));
        x1 = Math.min(x1, b.x1);
        y1 = Math.min(y1, b.y1);
        x2 = Math.max(x2, b.x2);
        y2 = Math.max(y2, b.y2);
    }
    const box = { x: Math.round(x1), y: Math.round(y1), cx: Math.round(x2 - x1), cy: Math.round(y2 - y1) };
    if (![box.x, box.y, box.cx, box.cy].every(Number.isFinite))
        (0, edit_1.reject)(edit_1.MSG.elementNotFound);
    const id = (0, elements_1.nextShapeId)(root);
    const group = (0, xml_1.fragment)(first.doc, (0, templates_1.groupXml)(id, `Group ${id - 1}`, box));
    // いちばん前面の要素の位置にグループを置き、要素を元の順で移す
    tree.insertBefore(group, members[members.length - 1]);
    for (const m of members)
        group.appendChild(m);
    // グループの中の図形はアニメーションの対象にできない（PowerPoint と同じく取り除く）
    const ids = new Set(members.map((m) => (0, xml_1.attr)((0, inherit_1.cNvPrOf)(m), 'id') ?? '').filter((x) => x));
    (0, edit_1.cleanTiming)(root, ids);
    return { changes: [], select: selectIdsOf(root, first.slideId, [group]) };
}
// ---------------------------------------------------------------------------
// グループ解除
// ---------------------------------------------------------------------------
const FILL_NAMES = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill'];
function sameXfrm(a, b) {
    const n = (r) => ((Math.round(r * 60000) % 21600000) + 21600000) % 21600000;
    return a.x === b.x && a.y === b.y && a.cx === b.cx && a.cy === b.cy && n(a.rot) === n(b.rot) && a.flipH === b.flipH && a.flipV === b.flipV;
}
/**
 * 子の a:grpFill（グループの塗りを使う）を、グループの塗りの複製に置き換える。グループに塗りがなければ塗りなし。
 * 入れ子のグループ（p:grpSpPr の a:grpFill）も置き換える（その子の a:grpFill が外側の塗りを引き継ぐように）。
 */
function replaceGroupFill(child, groupFill) {
    const pr = (0, xml_1.is)(child, 'p:grpSp') ? (0, xml_1.kid)(child, 'p:grpSpPr') : (0, xml_1.kid)(child, 'p:spPr');
    const gf = (0, xml_1.kid)(pr, 'a:grpFill');
    if (!pr || !gf)
        return;
    const doc = child.ownerDocument;
    const next = groupFill ? groupFill.cloneNode(true) : (0, xml_1.create)(doc, 'a:noFill');
    pr.insertBefore(next, gf);
    (0, xml_1.remove)(gf);
}
/** n 個の値を合計が total になるように比例で丸める（端数は最後に寄せる）。 */
function scaleParts(values, total) {
    const sum = values.reduce((a, b) => a + b, 0);
    if (!(sum > 0))
        return values;
    const out = values.map((v) => Math.max(0, Math.round((v * total) / sum)));
    const diff = Math.round(total) - out.reduce((a, b) => a + b, 0);
    out[out.length - 1] = Math.max(0, out[out.length - 1] + diff);
    return out;
}
/** 表の列の幅（a:gridCol@w）と行の高さ（a:tr@h）を、枠の大きさ cx × cy に合わせて比例で変える。 */
function scaleTable(frame, cx, cy) {
    const tbl = (0, xml_1.kid)((0, xml_1.kid)((0, xml_1.kid)(frame, 'a:graphic'), 'a:graphicData'), 'a:tbl');
    if (!tbl)
        return;
    const cols = (0, xml_1.kids)((0, xml_1.kid)(tbl, 'a:tblGrid')).filter((k) => (0, xml_1.is)(k, 'a:gridCol'));
    const ws = scaleParts(cols.map((c) => (0, xml_1.numAttr)(c, 'w') ?? 0), cx);
    cols.forEach((c, i) => (0, xml_1.setAttr)(c, 'w', ws[i]));
    const trs = (0, xml_1.kids)(tbl).filter((k) => (0, xml_1.is)(k, 'a:tr'));
    const hs = scaleParts(trs.map((r) => (0, xml_1.numAttr)(r, 'h') ?? 0), cy);
    trs.forEach((r, i) => (0, xml_1.setAttr)(r, 'h', hs[i]));
}
/**
 * 位置だけを書く要素（表、グラフ、OLE、mc:AlternateContent）を、拡大・縮小したグループから出すとき、
 * 画面に見えていた大きさ（グループの変換を当てた大きさ）に合わせる。表は列の幅と行の高さも比例で変える。
 */
function fitMoveOnly(m, own, onSlide) {
    if (Math.round(own.cx) === Math.round(onSlide.cx) && Math.round(own.cy) === Math.round(onSlide.cy))
        return;
    if ((0, xml_1.is)(m, 'p:graphicFrame'))
        scaleTable(m, onSlide.cx, onSlide.cy);
    (0, edit_1.writeExt)(m, onSlide.cx, onSlide.cy);
}
function ungroupElement(txn, op) {
    const select = [];
    const done = new Set();
    for (const id of op.elementIds) {
        const f = (0, edit_1.findElement)(txn, id);
        if (f.ref.groups.length)
            (0, edit_1.reject)(edit_1.MSG.groupChild);
        const g = f.ref.el;
        if (!(0, xml_1.is)(g, 'p:grpSp'))
            (0, edit_1.reject)(exports.ARRANGE_MSG.notGroup);
        if (done.has(g))
            continue;
        done.add(g);
        const tree = g.parentNode;
        const ctx = (0, inherit_1.buildSlideCtx)(txn, f.slidePath);
        const groupFill = (0, xml_1.kids)((0, xml_1.kid)(g, 'p:grpSpPr')).find((k) => (0, xml_1.isAny)(k, FILL_NAMES));
        const members = (0, xml_1.kids)(g).filter((k) => (0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS));
        for (const m of members) {
            // chOff/chExt が off/ext と同じで回転・反転もなければ座標はそのまま。そうでなければグループの変換を当てる
            const own = (0, parse_1.resolvedOwnXfrm)(ctx, m);
            const onSlide = (0, elements_1.toSlideXfrm)(own, [g]);
            if (!sameXfrm(own, onSlide) || !hasExplicitXfrm(m)) {
                const moveOnly = (0, edit_1.isUnsupportedEl)(m);
                (0, edit_1.writeXfrm)(m, onSlide, moveOnly);
                if (moveOnly)
                    fitMoveOnly(m, own, onSlide);
            }
            replaceGroupFill(m, groupFill);
            tree.insertBefore(m, g);
        }
        const gid = (0, xml_1.attr)((0, inherit_1.cNvPrOf)(g), 'id');
        (0, xml_1.remove)(g);
        if (gid)
            (0, edit_1.cleanTiming)(f.root, new Set([gid]));
        select.push(...selectIdsOf(f.root, f.slideId, members));
    }
    return { changes: [], select };
}
function hasExplicitXfrm(el) {
    const target = el;
    if ((0, xml_1.is)(target, 'p:graphicFrame'))
        return !!(0, xml_1.kid)(target, 'p:xfrm');
    if ((0, xml_1.is)(target, 'p:grpSp'))
        return !!(0, xml_1.kid)((0, xml_1.kid)(target, 'p:grpSpPr'), 'a:xfrm');
    if ((0, xml_1.is)(target, 'mc:AlternateContent'))
        return true;
    return !!(0, xml_1.kid)((0, xml_1.kid)(target, 'p:spPr'), 'a:xfrm');
}
// ---------------------------------------------------------------------------
// テキストの調整（余白と垂直方向の配置）
// ---------------------------------------------------------------------------
const INSET_ATTR = { l: 'lIns', t: 'tIns', r: 'rIns', b: 'bIns' };
/** 表のセルの余白（a:tcPr）。 */
const CELL_MARGIN_ATTR = { l: 'marL', t: 'marT', r: 'marR', b: 'marB' };
function setTextBody(txn, op) {
    const insets = op.insets ?? {};
    for (const k of Object.keys(insets)) {
        const v = insets[k];
        if (!(k in INSET_ATTR))
            (0, edit_1.reject)(exports.ARRANGE_MSG.badInset);
        if (v === undefined)
            continue;
        if (typeof v !== 'number' || !Number.isFinite(v) || v < 0 || v > MAX_COORD)
            (0, edit_1.reject)(exports.ARRANGE_MSG.badInset);
    }
    if (op.anchor !== undefined && !['t', 'ctr', 'b'].includes(op.anchor))
        (0, edit_1.reject)(exports.ARRANGE_MSG.badAnchor);
    const keys = Object.keys(INSET_ATTR).filter((k) => insets[k] !== undefined);
    if (!keys.length && op.anchor === undefined)
        return { changes: [] };
    let count = 0;
    for (const id of op.elementIds) {
        const f = (0, edit_1.findElement)(txn, id);
        for (const sp of (0, edit_1.textShapes)(f.ref.el)) {
            const tx = (0, xml_1.kid)(sp, 'p:txBody');
            const bodyPr = (0, xml_1.kid)(tx, 'a:bodyPr') ?? (0, xml_1.insertChildInOrder)(tx, (0, xml_1.create)(f.doc, 'a:bodyPr'));
            for (const k of keys)
                (0, xml_1.setAttr)(bodyPr, INSET_ATTR[k], Math.round(insets[k]));
            if (op.anchor !== undefined)
                (0, xml_1.setAttr)(bodyPr, 'anchor', op.anchor);
            count++;
        }
        // 表：セルごとの a:tcPr（a:tc の中で a:txBody の後、a:extLst の前に置く）
        for (const c of (0, edit_1.tableCells)(f.ref.el)) {
            const tcPr = (0, xml_1.kid)(c.tc, 'a:tcPr') ?? (0, xml_1.insertChildInOrder)(c.tc, (0, xml_1.create)(f.doc, 'a:tcPr'));
            for (const k of keys)
                (0, xml_1.setAttr)(tcPr, CELL_MARGIN_ATTR[k], Math.round(insets[k]));
            if (op.anchor !== undefined)
                (0, xml_1.setAttr)(tcPr, 'anchor', op.anchor);
            count++;
        }
    }
    if (count === 0)
        (0, edit_1.reject)(edit_1.MSG.noText);
    return { changes: [] };
}
// ---------------------------------------------------------------------------
// 振り分け
// ---------------------------------------------------------------------------
function applyArrangeOp(txn, op) {
    switch (op.type) {
        case 'groupElements':
            return groupElements(txn, op);
        case 'ungroupElement':
            return ungroupElement(txn, op);
        case 'setTextBody':
            return setTextBody(txn, op);
        default:
            return undefined;
    }
}
