"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.masterPaths = masterPaths;
exports.layoutPathsOf = layoutPathsOf;
exports.allLayoutPaths = allLayoutPaths;
exports.readLayouts = readLayouts;
const inherit_1 = require("./inherit.js");
const package_1 = require("./package.js");
const xml_1 = require("./xml.js");
/** サムネイルに描かないプレースホルダー（日付、フッター、スライド番号、ヘッダー）。 */
const SKIP_PH = new Set(['dt', 'ftr', 'sldNum', 'hdr']);
/** マスターのパスを p:sldMasterIdLst の順に返す。 */
function masterPaths(r, pres) {
    const out = [];
    for (const m of (0, xml_1.kids)((0, xml_1.kid)(pres.root, 'p:sldMasterIdLst'))) {
        if (!(0, xml_1.is)(m, 'p:sldMasterId'))
            continue;
        const rel = (0, package_1.relById)(r, pres.path, (0, xml_1.attr)(m, 'r:id') ?? '');
        if (rel && rel.type === package_1.REL.slideMaster && r.has(rel.target) && !out.includes(rel.target))
            out.push(rel.target);
    }
    // sldMasterIdLst にないマスター（壊れたファイル）も関係から拾う
    for (const rel of (0, package_1.relsOfType)(r, pres.path, package_1.REL.slideMaster))
        if (r.has(rel.target) && !out.includes(rel.target))
            out.push(rel.target);
    return out;
}
/** マスターごとに、p:sldLayoutIdLst の順でレイアウトのパスを返す。 */
function layoutPathsOf(r, masterPath) {
    const root = (0, package_1.rootEl)(r.dom(masterPath));
    const out = [];
    for (const l of (0, xml_1.kids)((0, xml_1.kid)(root, 'p:sldLayoutIdLst'))) {
        if (!(0, xml_1.is)(l, 'p:sldLayoutId'))
            continue;
        const rel = (0, package_1.relById)(r, masterPath, (0, xml_1.attr)(l, 'r:id') ?? '');
        if (rel && rel.type === package_1.REL.slideLayout && r.has(rel.target) && !out.includes(rel.target))
            out.push(rel.target);
    }
    for (const rel of (0, package_1.relsOfType)(r, masterPath, package_1.REL.slideLayout))
        if (r.has(rel.target) && !out.includes(rel.target))
            out.push(rel.target);
    return out;
}
/** プレゼンテーションで使えるすべてのレイアウトのパス。 */
function allLayoutPaths(r, pres) {
    return masterPaths(r, pres).flatMap((m) => layoutPathsOf(r, m));
}
function layoutInfo(r, layoutPath, masterRoot) {
    const root = (0, package_1.rootEl)(r.dom(layoutPath));
    const name = (0, xml_1.attr)((0, xml_1.kid)(root, 'p:cSld'), 'name') || (0, package_1.basename)(layoutPath).replace(/\.xml$/, '');
    const placeholders = [];
    for (const k of (0, xml_1.kids)((0, xml_1.path)(root, 'p:cSld/p:spTree'))) {
        if (!(0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS))
            continue;
        const ph = (0, inherit_1.phOf)(k);
        if (!ph || SKIP_PH.has(ph.type))
            continue;
        const own = (0, inherit_1.xfrmElOf)(k);
        const masterPh = (0, inherit_1.findMasterPh)(masterRoot, ph.type);
        const xfrm = (0, inherit_1.readXfrmEl)(own) ?? (masterPh ? (0, inherit_1.readXfrmEl)((0, inherit_1.xfrmElOf)(masterPh)) : undefined);
        if (!xfrm || xfrm.cx <= 0 || xfrm.cy <= 0)
            continue;
        placeholders.push({ type: ph.type, xfrm });
    }
    return { path: layoutPath, name, placeholders };
}
function readLayouts(r, pres) {
    const out = [];
    for (const m of masterPaths(r, pres)) {
        const masterRoot = (0, package_1.rootEl)(r.dom(m));
        for (const l of layoutPathsOf(r, m))
            out.push(layoutInfo(r, l, masterRoot));
    }
    return out;
}
