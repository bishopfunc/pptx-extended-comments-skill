"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.MEDIUM_STYLE_2_ACCENT_1 = void 0;
exports.findTableStyle = findTableStyle;
exports.tableFlags = tableFlags;
exports.resolveCellStyle = resolveCellStyle;
// 表のスタイル（ppt/tableStyles.xml の a:tblStyle と組み込みスタイル）をセルごとに解決する（設計書 7 章「表」）。
// ここではスタイルの要素を選ぶだけで、色や線の値への変換は parse.ts が行う。
const inherit_1 = require("./inherit.js");
const package_1 = require("./package.js");
const theme_1 = require("./theme.js");
const xml_1 = require("./xml.js");
const own_1 = require("../shared/own.js");
/** PowerPoint の既定の表スタイル「中間スタイル 2 - アクセント 1」。 */
exports.MEDIUM_STYLE_2_ACCENT_1 = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}';
const REL_TABLE_STYLES = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles';
const LN_LT1 = (w) => `<a:ln w="${w}" cmpd="sng"><a:solidFill><a:schemeClr val="lt1"/></a:solidFill></a:ln>`;
const TX_LT1_BOLD = '<a:tcTxStyle b="on"><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="lt1"/></a:tcTxStyle>';
const FILL_ACCENT1 = '<a:fill><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:fill>';
const FILL_TINT = (t) => `<a:fill><a:solidFill><a:schemeClr val="accent1"><a:tint val="${t}"/></a:schemeClr></a:solidFill></a:fill>`;
/** 組み込みスタイルは tableStyles.xml に定義が書かれず、ID だけで参照されることが多い。 */
const BUILTIN_XML = {
    [exports.MEDIUM_STYLE_2_ACCENT_1]: `<a:tblStyle xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" styleId="${exports.MEDIUM_STYLE_2_ACCENT_1}" styleName="Medium Style 2 - Accent 1">` +
        '<a:wholeTbl><a:tcTxStyle><a:fontRef idx="minor"><a:prstClr val="black"/></a:fontRef><a:schemeClr val="dk1"/></a:tcTxStyle>' +
        `<a:tcStyle><a:tcBdr><a:left>${LN_LT1(12700)}</a:left><a:right>${LN_LT1(12700)}</a:right><a:top>${LN_LT1(12700)}</a:top>` +
        `<a:bottom>${LN_LT1(12700)}</a:bottom><a:insideH>${LN_LT1(12700)}</a:insideH><a:insideV>${LN_LT1(12700)}</a:insideV></a:tcBdr>` +
        `${FILL_TINT(20000)}</a:tcStyle></a:wholeTbl>` +
        `<a:band1H><a:tcStyle><a:tcBdr/>${FILL_TINT(40000)}</a:tcStyle></a:band1H>` +
        '<a:band2H><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2H>' +
        `<a:band1V><a:tcStyle><a:tcBdr/>${FILL_TINT(40000)}</a:tcStyle></a:band1V>` +
        '<a:band2V><a:tcStyle><a:tcBdr/></a:tcStyle></a:band2V>' +
        `<a:lastCol>${TX_LT1_BOLD}<a:tcStyle><a:tcBdr/>${FILL_ACCENT1}</a:tcStyle></a:lastCol>` +
        `<a:firstCol>${TX_LT1_BOLD}<a:tcStyle><a:tcBdr/>${FILL_ACCENT1}</a:tcStyle></a:firstCol>` +
        `<a:lastRow>${TX_LT1_BOLD}<a:tcStyle><a:tcBdr><a:top>${LN_LT1(38100)}</a:top></a:tcBdr>${FILL_ACCENT1}</a:tcStyle></a:lastRow>` +
        `<a:firstRow>${TX_LT1_BOLD}<a:tcStyle><a:tcBdr><a:bottom>${LN_LT1(38100)}</a:bottom></a:tcBdr>${FILL_ACCENT1}</a:tcStyle></a:firstRow>` +
        '</a:tblStyle>',
};
const builtinCache = new Map();
function builtinStyle(id) {
    const xml = (0, own_1.own)(BUILTIN_XML, id);
    if (!xml)
        return undefined;
    let el = builtinCache.get(id);
    if (!el) {
        el = (0, package_1.rootEl)((0, xml_1.parseXml)(xml));
        builtinCache.set(id, el);
    }
    return el;
}
function tableStylesPath(r) {
    const pres = (0, inherit_1.readPresentation)(r);
    const rel = (0, package_1.readRels)(r, pres.path).find((x) => x.type === REL_TABLE_STYLES && !x.external);
    if (rel && r.has(rel.target))
        return rel.target;
    return r.has('ppt/tableStyles.xml') ? 'ppt/tableStyles.xml' : undefined;
}
/** a:tblPr が参照する表スタイル（a:tblStyle）。tableStyles.xml になければ組み込みスタイル、それもなければ undefined。 */
function findTableStyle(r, tblPr) {
    const inline = (0, xml_1.kid)(tblPr, 'a:tableStyle');
    if (inline)
        return inline;
    const id = (0, xml_1.kid)(tblPr, 'a:tableStyleId')?.textContent?.trim();
    if (!id)
        return undefined;
    const p = tableStylesPath(r);
    if (p) {
        const lst = (0, package_1.rootEl)(r.dom(p));
        const found = (0, xml_1.kids)(lst).find((k) => (0, xml_1.is)(k, 'a:tblStyle') && (0, xml_1.attr)(k, 'styleId')?.toUpperCase() === id.toUpperCase());
        if (found)
            return found;
    }
    return builtinStyle(id.toUpperCase());
}
function tableFlags(tblPr) {
    const on = (n) => {
        const v = (0, xml_1.attr)(tblPr, n);
        return v === '1' || v === 'true';
    };
    return {
        firstRow: on('firstRow'),
        lastRow: on('lastRow'),
        firstCol: on('firstCol'),
        lastCol: on('lastCol'),
        bandRow: on('bandRow'),
        bandCol: on('bandCol'),
    };
}
/** 適用する部分を優先度の低い順に並べる。 */
function applicableParts(f, p) {
    const all = { r0: 0, r1: p.nRows - 1, c0: 0, c1: p.nCols - 1 };
    const lastR = p.nRows - 1;
    const lastC = p.nCols - 1;
    const out = [{ name: 'a:wholeTbl', region: all }];
    const isHeaderRow = f.firstRow && p.row === 0;
    const isTotalRow = f.lastRow && p.rowEnd === lastR;
    const isFirstCol = f.firstCol && p.col === 0;
    const isLastCol = f.lastCol && p.colEnd === lastC;
    if (f.bandRow && !isHeaderRow && !isTotalRow) {
        const k = p.row - (f.firstRow ? 1 : 0);
        out.push({ name: k % 2 === 0 ? 'a:band1H' : 'a:band2H', region: { r0: p.row, r1: p.row, c0: 0, c1: lastC } });
    }
    if (f.bandCol && !isFirstCol && !isLastCol) {
        const k = p.col - (f.firstCol ? 1 : 0);
        out.push({ name: k % 2 === 0 ? 'a:band1V' : 'a:band2V', region: { r0: 0, r1: lastR, c0: p.col, c1: p.col } });
    }
    if (isLastCol)
        out.push({ name: 'a:lastCol', region: { r0: 0, r1: lastR, c0: lastC, c1: lastC } });
    if (isFirstCol)
        out.push({ name: 'a:firstCol', region: { r0: 0, r1: lastR, c0: 0, c1: 0 } });
    if (isTotalRow)
        out.push({ name: 'a:lastRow', region: { r0: lastR, r1: lastR, c0: 0, c1: lastC } });
    if (isHeaderRow)
        out.push({ name: 'a:firstRow', region: { r0: 0, r1: 0, c0: 0, c1: lastC } });
    const one = (r, c) => ({ r0: r, r1: r, c0: c, c1: c });
    if (isTotalRow && isLastCol)
        out.push({ name: 'a:seCell', region: one(lastR, lastC) });
    if (isTotalRow && isFirstCol)
        out.push({ name: 'a:swCell', region: one(lastR, 0) });
    if (isHeaderRow && isLastCol)
        out.push({ name: 'a:neCell', region: one(0, lastC) });
    if (isHeaderRow && isFirstCol)
        out.push({ name: 'a:nwCell', region: one(0, 0) });
    return out;
}
function onOff(el, name) {
    const v = (0, xml_1.attr)(el, name);
    if (v === 'on' || v === '1' || v === 'true')
        return true;
    if (v === 'off' || v === '0' || v === 'false')
        return false;
    return undefined; // def、または書かれていない
}
function refColor(ref, color) {
    return { ...color, phClr: (0, theme_1.colorOfParent)(ref, color) };
}
function fillOfTcStyle(tcStyle, color, theme) {
    const fill = (0, xml_1.kids)((0, xml_1.kid)(tcStyle, 'a:fill'))[0];
    if (fill)
        return { el: fill, color };
    const ref = (0, xml_1.kid)(tcStyle, 'a:fillRef');
    if (ref) {
        const idx = (0, xml_1.numAttr)(ref, 'idx') ?? 0;
        const entry = idx >= 1001 ? theme.bgFillStyles[idx - 1001] : idx > 0 ? theme.fillStyles[idx - 1] : undefined;
        if (idx === 0)
            return { el: noFillEl(), color };
        if (entry)
            return { el: entry, color: refColor(ref, color) };
    }
    return undefined;
}
let noFillCache;
function noFillEl() {
    noFillCache ??= (0, package_1.rootEl)((0, xml_1.parseXml)('<a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/>'));
    return noFillCache;
}
function lineOfEdge(edge, color, theme) {
    if (!edge)
        return undefined;
    const ln = (0, xml_1.kid)(edge, 'a:ln');
    if (ln)
        return { ln, color };
    const ref = (0, xml_1.kid)(edge, 'a:lnRef');
    if (ref) {
        const idx = (0, xml_1.numAttr)(ref, 'idx') ?? 0;
        const entry = idx > 0 ? theme.lineStyles[idx - 1] : undefined;
        if (entry)
            return { ln: entry, color: refColor(ref, color) };
    }
    return undefined;
}
/** セルに効く表スタイルの値（優先度の高い部分が勝つ）。style がなければ空。 */
function resolveCellStyle(style, flags, pos, color, theme) {
    const out = { borders: {}, text: {} };
    if (!style)
        return out;
    for (const { name, region } of applicableParts(flags, pos)) {
        const part = (0, xml_1.kid)(style, name);
        if (!part)
            continue;
        const tcStyle = (0, xml_1.kid)(part, 'a:tcStyle');
        const f = fillOfTcStyle(tcStyle, color, theme);
        if (f)
            out.fill = f;
        const bdr = (0, xml_1.kid)(tcStyle, 'a:tcBdr');
        if (bdr) {
            const edge = (outer, inner, isOuter) => lineOfEdge((0, xml_1.kid)(bdr, isOuter ? outer : inner), color, theme);
            const l = edge('a:left', 'a:insideV', pos.col <= region.c0);
            const r = edge('a:right', 'a:insideV', pos.colEnd >= region.c1);
            const t = edge('a:top', 'a:insideH', pos.row <= region.r0);
            const b = edge('a:bottom', 'a:insideH', pos.rowEnd >= region.r1);
            if (l)
                out.borders.l = l;
            if (r)
                out.borders.r = r;
            if (t)
                out.borders.t = t;
            if (b)
                out.borders.b = b;
        }
        const tx = (0, xml_1.kid)(part, 'a:tcTxStyle');
        if (tx) {
            const b = onOff(tx, 'b');
            if (b !== undefined)
                out.text.bold = b;
            const i = onOff(tx, 'i');
            if (i !== undefined)
                out.text.italic = i;
            const fontRef = (0, xml_1.kid)(tx, 'a:fontRef');
            const font = (0, xml_1.kid)(tx, 'a:font');
            if (fontRef) {
                const idx = (0, xml_1.attr)(fontRef, 'idx');
                if (idx === 'major' || idx === 'minor') {
                    out.text.latin = idx === 'major' ? '+mj-lt' : '+mn-lt';
                    out.text.ea = idx === 'major' ? '+mj-ea' : '+mn-ea';
                }
            }
            else if (font) {
                const lt = (0, xml_1.attr)((0, xml_1.kid)(font, 'a:latin'), 'typeface');
                const ea = (0, xml_1.attr)((0, xml_1.kid)(font, 'a:ea'), 'typeface');
                if (lt)
                    out.text.latin = lt;
                if (ea)
                    out.text.ea = ea;
            }
            const direct = (0, xml_1.kids)(tx).find((k) => (0, theme_1.isColorEl)(k));
            const c = direct ? (0, theme_1.resolveColor)(direct, color) : fontRef ? (0, theme_1.colorOfParent)(fontRef, color) : undefined;
            if (c)
                out.text.color = c;
        }
    }
    return out;
}
