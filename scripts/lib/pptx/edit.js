"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.COPY_PART_TYPES = exports.EXT_CONTENT_TYPE = exports.MSG = void 0;
exports.reject = reject;
exports.findSlide = findSlide;
exports.findElement = findElement;
exports.branches = branches;
exports.writeXfrm = writeXfrm;
exports.writeExt = writeExt;
exports.isUnsupportedEl = isUnsupportedEl;
exports.sameStyle = sameStyle;
exports.hexOf = hexOf;
exports.rPrToEndParaRPr = rPrToEndParaRPr;
exports.tableCells = tableCells;
exports.ownTables = ownTables;
exports.isHiddenCell = isHiddenCell;
exports.textShapes = textShapes;
exports.nextMediaPath = nextMediaPath;
exports.cleanTiming = cleanTiming;
exports.rAttrs = rAttrs;
exports.copyElements = copyElements;
exports.removeRAttr = removeRAttr;
exports.isSafeExternalRel = isSafeExternalRel;
exports.partResolver = partResolver;
exports.normalizeCrop = normalizeCrop;
exports.applyElementOp = applyElementOp;
const own_1 = require("../shared/own.js");
const creationIds_1 = require("./creationIds.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const paragraph_1 = require("./paragraph.js");
const package_1 = require("./package.js");
const parse_1 = require("./parse.js");
const templates_1 = require("./templates.js");
const types_1 = require("./types.js");
const xml_1 = require("./xml.js");
const FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
const LINE_FILLS = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill'];
exports.MSG = {
    elementNotFound: '対象の要素が見つかりません。スライドが変更された可能性があります',
    slideNotFound: '対象のスライドが見つかりません',
    groupChild: 'グループの中の要素は個別に操作できません',
    forbiddenCopy: 'グラフ、SmartArt、OLE を含む要素は、このバージョンでは複製・貼り付けできません',
    noText: 'この要素のテキストは編集できません',
    noFill: 'この要素の塗りや線は変更できません',
    imageFill: '画像による塗りは指定できません',
    badColor: '色の指定が正しくありません',
    noClipboard: '貼り付ける内容がありません',
    cellNotFound: '対象のセルが見つかりません。表が変更された可能性があります',
    hiddenCell: '結合で隠れているセルは編集できません',
    badImage: '画像の形式に対応していません（png、jpeg、gif のみ）',
    emptyImage: '画像のデータが空です',
    notPicture: 'この要素は画像ではないため、切り抜けません',
    badCrop: '切り抜きの範囲が正しくありません',
};
function reject(msg) {
    throw new types_1.OpRejectedError(msg);
}
function findSlide(txn, slideId) {
    const pres = (0, inherit_1.readPresentation)(txn);
    const s = pres.slides.find((x) => x.id === slideId);
    if (!s)
        reject(exports.MSG.slideNotFound);
    return { slideId, slidePath: s.path };
}
/** 書き換え用の DOM の上で要素を探す。 */
function findElement(txn, elementId) {
    const parsed = (0, elements_1.parseElementId)(elementId);
    if (!parsed)
        reject(exports.MSG.elementNotFound);
    const s = findSlide(txn, parsed.slideId);
    const doc = txn.edit(s.slidePath);
    const root = (0, package_1.rootEl)(doc);
    const ref = (0, elements_1.indexElements)((0, elements_1.spTreeOf)(root)).find((x) => x.key === parsed.key);
    if (!ref)
        reject(exports.MSG.elementNotFound);
    return { ...s, doc, root, ref };
}
/** mc:AlternateContent なら中の図形すべて、それ以外はその要素。 */
function branches(el) {
    if (!(0, xml_1.is)(el, 'mc:AlternateContent'))
        return [el];
    const out = [];
    for (const k of (0, xml_1.kids)(el))
        for (const c of (0, xml_1.kids)(k))
            if ((0, xml_1.isAny)(c, inherit_1.SHAPE_KINDS))
                out.push(c);
    return out;
}
// ---------------------------------------------------------------------------
// xfrm
// ---------------------------------------------------------------------------
function xfrmContainer(el) {
    const doc = el.ownerDocument;
    if ((0, xml_1.is)(el, 'p:graphicFrame')) {
        let xf = (0, xml_1.kid)(el, 'p:xfrm');
        if (!xf) {
            xf = (0, xml_1.create)(doc, 'p:xfrm');
            const graphic = (0, xml_1.kid)(el, 'a:graphic');
            if (graphic)
                el.insertBefore(xf, graphic);
            else
                el.appendChild(xf);
        }
        return { xf };
    }
    if ((0, xml_1.is)(el, 'p:grpSp')) {
        const grpSpPr = (0, xml_1.ensureKid)(el, 'p:grpSpPr');
        return { xf: (0, xml_1.ensureKid)(grpSpPr, 'a:xfrm') };
    }
    const spPr = (0, xml_1.ensureKid)(el, 'p:spPr');
    return { xf: (0, xml_1.ensureKid)(spPr, 'a:xfrm') };
}
function setOffExt(xf, x, moveOnly) {
    const doc = xf.ownerDocument;
    let off = (0, xml_1.kid)(xf, 'a:off');
    if (!off) {
        off = (0, xml_1.create)(doc, 'a:off');
        xf.insertBefore(off, xf.firstChild);
    }
    (0, xml_1.setAttr)(off, 'x', (0, xml_1.emuCoord)(x.x));
    (0, xml_1.setAttr)(off, 'y', (0, xml_1.emuCoord)(x.y));
    let ext = (0, xml_1.kid)(xf, 'a:ext');
    if (!ext) {
        ext = (0, xml_1.create)(doc, 'a:ext', { cx: (0, xml_1.emuSize)(x.cx), cy: (0, xml_1.emuSize)(x.cy) });
        if (off.nextSibling)
            xf.insertBefore(ext, off.nextSibling);
        else
            xf.appendChild(ext);
    }
    if (moveOnly)
        return;
    (0, xml_1.setAttr)(ext, 'cx', (0, xml_1.emuSize)(x.cx));
    (0, xml_1.setAttr)(ext, 'cy', (0, xml_1.emuSize)(x.cy));
    const rot = (0, xml_1.rotUnits)(x.rot);
    (0, xml_1.setAttr)(xf, 'rot', rot ? rot : undefined);
    (0, xml_1.setAttr)(xf, 'flipH', x.flipH ? '1' : undefined);
    (0, xml_1.setAttr)(xf, 'flipV', x.flipV ? '1' : undefined);
}
function writeXfrm(el, x, moveOnly) {
    for (const b of branches(el))
        setOffExt(xfrmContainer(b).xf, x, moveOnly);
}
/** 大きさ（a:ext）だけを書く。位置・回転・反転はそのまま。 */
function writeExt(el, cx, cy) {
    for (const b of branches(el)) {
        const xf = xfrmContainer(b).xf;
        const ext = (0, xml_1.kid)(xf, 'a:ext');
        if (!ext)
            continue;
        (0, xml_1.setAttr)(ext, 'cx', (0, xml_1.emuSize)(cx));
        (0, xml_1.setAttr)(ext, 'cy', (0, xml_1.emuSize)(cy));
    }
}
function isUnsupportedEl(el) {
    return !(0, xml_1.isAny)(el, ['p:sp', 'p:cxnSp', 'p:pic', 'p:grpSp']);
}
function setXfrm(txn, op) {
    for (const item of op.items) {
        const f = findElement(txn, item.elementId);
        if (f.ref.groups.length)
            reject(exports.MSG.groupChild);
        writeXfrm(f.ref.el, item.xfrm, isUnsupportedEl(f.ref.el));
    }
}
// ---------------------------------------------------------------------------
// テキストの書式
// ---------------------------------------------------------------------------
const STYLE_KEYS = ['bold', 'italic', 'underline', 'strike', 'size', 'color', 'latinFont', 'eaFont', 'baseline'];
function normVal(k, v) {
    if (v === undefined)
        return undefined;
    if (k === 'color' && typeof v === 'string')
        return v.replace('#', '').toUpperCase();
    if (k === 'size' && typeof v === 'number')
        return Math.round(v * 100);
    return v;
}
function sameProp(k, a, b) {
    return normVal(k, a) === normVal(k, b);
}
function sameStyle(a, b) {
    return STYLE_KEYS.every((k) => sameProp(k, a[k], b[k]));
}
function hexOf(color) {
    const h = color.replace('#', '').toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(h))
        reject(exports.MSG.badColor);
    return h;
}
function setFont(rPr, qname, typeface) {
    const cur = (0, xml_1.kid)(rPr, qname);
    if (typeface === undefined) {
        (0, xml_1.remove)(cur);
        return;
    }
    if (cur)
        (0, xml_1.setAttr)(cur, 'typeface', typeface);
    else
        (0, xml_1.insertChildInOrder)(rPr, (0, xml_1.create)(rPr.ownerDocument, qname, { typeface }));
}
/** rPr に 1 項目を書く（undefined なら明示的な値を取り除く）。 */
function writeProp(rPr, k, v, style) {
    const doc = rPr.ownerDocument;
    switch (k) {
        case 'bold':
            (0, xml_1.setAttr)(rPr, 'b', v === undefined ? undefined : v ? '1' : '0');
            break;
        case 'italic':
            (0, xml_1.setAttr)(rPr, 'i', v === undefined ? undefined : v ? '1' : '0');
            break;
        case 'underline': {
            const cur = (0, xml_1.attr)(rPr, 'u');
            (0, xml_1.setAttr)(rPr, 'u', v === undefined ? undefined : v ? (cur && cur !== 'none' ? cur : 'sng') : 'none');
            break;
        }
        case 'strike':
            (0, xml_1.setAttr)(rPr, 'strike', v === undefined ? undefined : v ? 'sngStrike' : 'noStrike');
            break;
        case 'size':
            (0, xml_1.setAttr)(rPr, 'sz', v === undefined ? undefined : Math.max(100, Math.min(400000, Math.round(v * 100))));
            break;
        case 'baseline':
            (0, xml_1.setAttr)(rPr, 'baseline', v === undefined ? undefined : Math.round(v));
            break;
        case 'color': {
            (0, xml_1.removeKids)(rPr, FILLS);
            if (v !== undefined) {
                const fill = (0, xml_1.create)(doc, 'a:solidFill');
                fill.appendChild((0, xml_1.create)(doc, 'a:srgbClr', { val: hexOf(v) }));
                (0, xml_1.insertChildInOrder)(rPr, fill);
            }
            break;
        }
        case 'latinFont':
            setFont(rPr, 'a:latin', v);
            // フォントを指定したら a:latin と a:ea の両方を書く
            if (v !== undefined && style.eaFont === undefined)
                setFont(rPr, 'a:ea', v);
            break;
        case 'eaFont':
            setFont(rPr, 'a:ea', v);
            if (v !== undefined && style.latinFont === undefined && !(0, xml_1.kid)(rPr, 'a:latin'))
                setFont(rPr, 'a:latin', v);
            break;
    }
}
/** base（元の explicit）と違う項目だけを書く。 */
function writeDiff(rPr, base, next) {
    for (const k of STYLE_KEYS) {
        if (sameProp(k, base[k], next[k]))
            continue;
        writeProp(rPr, k, next[k], next);
    }
}
/** 指定された項目だけを書く（setTextStyle）。 */
function writeStyle(rPr, style) {
    for (const k of STYLE_KEYS) {
        if (style[k] === undefined)
            continue;
        writeProp(rPr, k, style[k], style);
    }
}
function hasContent(el) {
    return el.attributes.length > 0 || (0, xml_1.kids)(el).length > 0;
}
function stripLinks(rPr) {
    if (!rPr)
        return;
    (0, xml_1.removeKids)(rPr, ['a:hlinkClick', 'a:hlinkMouseOver']);
}
// eslint-disable-next-line no-control-regex
function textEl(doc, text) {
    const t = (0, xml_1.create)(doc, 'a:t');
    t.appendChild(doc.createTextNode((0, xml_1.stripInvalidXml)(text)));
    return t;
}
function clone(el) {
    return (el ? el.cloneNode(true) : undefined);
}
function rPrToEndParaRPr(doc, rPr) {
    const end = (0, xml_1.create)(doc, 'a:endParaRPr');
    if (rPr) {
        for (let i = 0; i < rPr.attributes.length; i++) {
            const a = rPr.attributes.item(i);
            end.setAttributeNS(a.namespaceURI ?? null, a.name, a.value);
        }
        for (const k of (0, xml_1.kids)(rPr))
            if (!(0, xml_1.isAny)(k, ['a:hlinkClick', 'a:hlinkMouseOver']))
                end.appendChild(k.cloneNode(true));
    }
    return end;
}
function ensureTxBody(el) {
    let tx = (0, xml_1.kid)(el, 'p:txBody');
    if (tx)
        return tx;
    const doc = el.ownerDocument;
    tx = (0, xml_1.create)(doc, 'p:txBody');
    tx.appendChild((0, xml_1.create)(doc, 'a:bodyPr', { wrap: 'square', rtlCol: '0', anchor: 'ctr' }));
    tx.appendChild((0, xml_1.create)(doc, 'a:lstStyle'));
    return (0, xml_1.insertChildInOrder)(el, tx);
}
/** 要素（グループなら中も含む）の中の表のすべてのセル。mc:AlternateContent なら分岐ごとの表のセル。 */
function tableCells(el) {
    const out = [];
    for (const tbl of [el, ...(0, xml_1.descendants)(el)].filter((e) => (0, xml_1.is)(e, 'a:tbl'))) {
        (0, xml_1.kidsNamed)(tbl, 'a:tr').forEach((tr, row) => (0, xml_1.kidsNamed)(tr, 'a:tc').forEach((tc, col) => out.push({ tbl, tc, row, col })));
    }
    return out;
}
/** 表の要素そのもの（p:graphicFrame、または mc:AlternateContent の各分岐）の a:tbl。 */
function ownTables(el) {
    return branches(el)
        .map((b) => ((0, xml_1.is)(b, 'p:graphicFrame') ? (0, xml_1.path)(b, 'a:graphic/a:graphicData/a:tbl') : undefined))
        .filter((t) => !!t);
}
function isHiddenCell(tc) {
    return (0, xml_1.boolAttr)(tc, 'hMerge') === true || (0, xml_1.boolAttr)(tc, 'vMerge') === true;
}
function ensureCellTxBody(tc) {
    const cur = (0, xml_1.kid)(tc, 'a:txBody');
    if (cur)
        return cur;
    const doc = tc.ownerDocument;
    const tx = (0, xml_1.create)(doc, 'a:txBody');
    tx.appendChild((0, xml_1.create)(doc, 'a:bodyPr'));
    tx.appendChild((0, xml_1.create)(doc, 'a:lstStyle'));
    return (0, xml_1.insertChildInOrder)(tc, tx);
}
// ---------------------------------------------------------------------------
// テキストの置き換え（setText）
// ---------------------------------------------------------------------------
function setText(txn, op) {
    const f = findElement(txn, op.elementId);
    const el = f.ref.el;
    for (const pi of op.paragraphs)
        (0, paragraph_1.validateParagraphProps)(pi);
    if (op.cell) {
        setCellText(txn, f, op.cell, op.paragraphs);
        return;
    }
    if (!(0, xml_1.is)(el, 'p:sp'))
        reject(exports.MSG.noText);
    const ctx = (0, inherit_1.buildSlideCtx)(txn, f.slidePath);
    const txBody = ensureTxBody(el);
    const chain = (0, inherit_1.textChain)(ctx, el, (0, inherit_1.phSources)(ctx, el, 'slide'), 'slide');
    rewriteParagraphs(ctx, txBody, op.paragraphs, chain);
}
/** 表のセルのテキストを置き換える。規則は図形と同じ（a:bodyPr と a:lstStyle を残し、a:p だけを作り直す）。 */
function setCellText(txn, f, cell, paragraphs) {
    const tbls = ownTables(f.ref.el);
    if (!tbls.length)
        reject(exports.MSG.noText);
    const { row, col } = cell;
    if (!Number.isInteger(row) || !Number.isInteger(col) || row < 0 || col < 0)
        reject(exports.MSG.cellNotFound);
    const targets = tbls.map((tbl) => {
        const tc = (0, xml_1.kidsNamed)((0, xml_1.kidsNamed)(tbl, 'a:tr')[row], 'a:tc')[col];
        if (!tc)
            reject(exports.MSG.cellNotFound);
        if (isHiddenCell(tc))
            reject(exports.MSG.hiddenCell);
        return { tbl, tc };
    });
    const ctx = (0, inherit_1.buildSlideCtx)(txn, f.slidePath);
    for (const { tbl, tc } of targets) {
        rewriteParagraphs(ctx, ensureCellTxBody(tc), paragraphs, (0, parse_1.tableCellChain)(ctx, tbl, row, col));
    }
}
/**
 * txBody（図形の p:txBody、またはセルの a:txBody）の a:p を作り直す。設計書 7 章「テキストの置き換え」。
 * chain は段落の書式（レベルの既定）を引くための継承チェーン。
 */
function rewriteParagraphs(ctx, txBody, paragraphs, chain) {
    const doc = txBody.ownerDocument;
    const origPs = (0, xml_1.kidsNamed)(txBody, 'a:p');
    const origRuns = origPs.map(parse_1.runChildren);
    const used = new Set();
    const newPs = [];
    let prev;
    const buildRun = (ri) => {
        const origEl = ri.source ? origRuns[ri.source.p]?.[ri.source.r] : undefined;
        const key = ri.source ? `${ri.source.p}:${ri.source.r}` : '';
        const first = !!origEl && !used.has(key);
        if (origEl)
            used.add(key);
        const origRPr = (0, xml_1.kid)(origEl, 'a:rPr');
        const base = (0, parse_1.explicitOf)(origRPr, ctx.theme, ctx.color);
        const styleSame = sameStyle(base, ri.explicit);
        const makeRPr = () => {
            const rPr = origRPr ? clone(origRPr) : undefined;
            if (rPr && !first)
                stripLinks(rPr);
            if (styleSame)
                return rPr;
            const target = rPr ?? (0, xml_1.create)(doc, 'a:rPr');
            writeDiff(target, base, ri.explicit);
            return rPr || hasContent(target) ? target : undefined;
        };
        if (ri.kind === 'break') {
            if (first && (0, xml_1.is)(origEl, 'a:br') && styleSame)
                return clone(origEl);
            const br = (0, xml_1.create)(doc, 'a:br');
            const rPr = makeRPr();
            if (rPr)
                br.appendChild(rPr);
            return br;
        }
        if (ri.kind === 'field' && origEl && (0, xml_1.is)(origEl, 'a:fld')) {
            const fld = clone(origEl);
            if (!first)
                stripLinks((0, xml_1.kid)(fld, 'a:rPr'));
            if (!styleSame) {
                let rPr = (0, xml_1.kid)(fld, 'a:rPr');
                if (!rPr)
                    rPr = (0, xml_1.insertChildInOrder)(fld, (0, xml_1.create)(doc, 'a:rPr'));
                writeDiff(rPr, base, ri.explicit);
            }
            return fld;
        }
        if (first && (0, xml_1.is)(origEl, 'a:r') && styleSame && ((0, xml_1.kid)(origEl, 'a:t')?.textContent ?? '') === ri.text) {
            return clone(origEl);
        }
        const r = (0, xml_1.create)(doc, 'a:r');
        const rPr = makeRPr();
        if (rPr)
            r.appendChild(rPr);
        r.appendChild(textEl(doc, ri.text));
        return r;
    };
    paragraphs.forEach((pi) => {
        const src = pi.source !== undefined ? origPs[pi.source] : undefined;
        const p = (0, xml_1.create)(doc, 'a:p');
        let pPr;
        let endRPr;
        if (src) {
            pPr = clone((0, xml_1.kid)(src, 'a:pPr'));
            endRPr = clone((0, xml_1.kid)(src, 'a:endParaRPr'));
        }
        else if (prev) {
            pPr = clone((0, xml_1.kid)(prev, 'a:pPr'));
            endRPr = clone((0, xml_1.kid)(prev, 'a:endParaRPr'));
        }
        const props = { align: pi.align, level: pi.level, bullet: pi.bullet, lineSpacing: pi.lineSpacing };
        if ((0, paragraph_1.hasParagraphProps)(props)) {
            pPr ??= (0, xml_1.create)(doc, 'a:pPr');
            (0, paragraph_1.applyParagraphProps)(pPr, props, chain);
        }
        if (pPr)
            p.appendChild(pPr);
        for (const ri of pi.runs)
            p.appendChild(buildRun(ri));
        if (!endRPr && pi.runs.length === 0) {
            const srcRun = src ? (0, parse_1.runChildren)(src)[0] : prev ? (0, parse_1.runChildren)(prev)[0] : undefined;
            endRPr = rPrToEndParaRPr(doc, (0, xml_1.kid)(srcRun, 'a:rPr'));
        }
        if (endRPr)
            p.appendChild(endRPr);
        newPs.push(p);
        prev = p;
    });
    if (newPs.length === 0) {
        // 空のテキスト：a:endParaRPr だけを持つ a:p を 1 つ書く
        const p = (0, xml_1.create)(doc, 'a:p');
        const first = origPs[0];
        const pPr = clone((0, xml_1.kid)(first, 'a:pPr'));
        if (pPr)
            p.appendChild(pPr);
        p.appendChild(clone((0, xml_1.kid)(first, 'a:endParaRPr')) ?? rPrToEndParaRPr(doc, (0, xml_1.kid)((0, parse_1.runChildren)(first ?? p)[0], 'a:rPr')));
        newPs.push(p);
    }
    for (const p of origPs)
        (0, xml_1.remove)(p);
    const extLst = (0, xml_1.kid)(txBody, 'a:extLst');
    for (const p of newPs) {
        if (extLst)
            txBody.insertBefore(p, extLst);
        else
            txBody.appendChild(p);
    }
}
/** 要素（グループなら中の図形）のうち、テキストを持つ p:sp。 */
function textShapes(el) {
    const all = [el, ...(0, xml_1.descendants)(el)].filter((e) => (0, xml_1.is)(e, 'p:sp') && (0, xml_1.kid)(e, 'p:txBody'));
    return all;
}
/** txBody のすべてのランと、ランのない段落の a:endParaRPr に書式を書く。 */
function styleTxBody(txBody, style) {
    const doc = txBody.ownerDocument;
    for (const p of (0, xml_1.kidsNamed)(txBody, 'a:p')) {
        const runs = (0, parse_1.runChildren)(p);
        for (const r of runs) {
            const rPr = (0, xml_1.kid)(r, 'a:rPr') ?? (0, xml_1.insertChildInOrder)(r, (0, xml_1.create)(doc, 'a:rPr', { lang: undefined }));
            writeStyle(rPr, style);
        }
        let end = (0, xml_1.kid)(p, 'a:endParaRPr');
        if (!end && runs.length === 0)
            end = (0, xml_1.insertChildInOrder)(p, (0, xml_1.create)(doc, 'a:endParaRPr'));
        if (end)
            writeStyle(end, style);
    }
}
/** 図形（グループなら中の図形）と表（すべてのセル）のテキストに書式を当てる。 */
function setTextStyle(txn, op) {
    for (const id of op.elementIds) {
        const f = findElement(txn, id);
        for (const sp of textShapes(f.ref.el))
            styleTxBody((0, xml_1.kid)(sp, 'p:txBody'), op.style);
        for (const c of tableCells(f.ref.el)) {
            const tx = (0, xml_1.kid)(c.tc, 'a:txBody');
            if (tx)
                styleTxBody(tx, op.style);
        }
    }
}
function setParagraphStyle(txn, op) {
    const props = { align: op.align, bullet: op.bullet, lineSpacing: op.lineSpacing, indentDelta: op.indentDelta };
    (0, paragraph_1.validateParagraphProps)(props);
    if (!(0, paragraph_1.hasParagraphProps)(props))
        return;
    const ctxs = new Map();
    const apply = (doc, txBody, chain) => {
        for (const p of (0, xml_1.kidsNamed)(txBody, 'a:p')) {
            const pPr = (0, xml_1.kid)(p, 'a:pPr') ?? (0, xml_1.insertChildInOrder)(p, (0, xml_1.create)(doc, 'a:pPr'));
            (0, paragraph_1.applyParagraphProps)(pPr, props, chain);
        }
    };
    for (const id of op.elementIds) {
        const f = findElement(txn, id);
        let ctx = ctxs.get(f.slidePath);
        if (!ctx)
            ctxs.set(f.slidePath, (ctx = (0, inherit_1.buildSlideCtx)(txn, f.slidePath)));
        for (const sp of textShapes(f.ref.el)) {
            apply(f.doc, (0, xml_1.kid)(sp, 'p:txBody'), (0, inherit_1.textChain)(ctx, sp, (0, inherit_1.phSources)(ctx, sp, 'slide'), 'slide'));
        }
        for (const c of tableCells(f.ref.el)) {
            apply(f.doc, (0, xml_1.kid)(c.tc, 'a:txBody'), (0, parse_1.tableCellChain)(ctx, c.tbl, c.row, c.col));
        }
    }
}
// ---------------------------------------------------------------------------
// 塗りと線
// ---------------------------------------------------------------------------
function colorEl(doc, color, alpha) {
    const c = (0, xml_1.create)(doc, 'a:srgbClr', { val: hexOf(color) });
    if (alpha !== undefined && alpha < 1)
        c.appendChild((0, xml_1.create)(doc, 'a:alpha', { val: Math.round(Math.max(0, alpha) * 100000) }));
    return c;
}
function fillEl(doc, fill) {
    switch (fill.type) {
        case 'none':
            return (0, xml_1.create)(doc, 'a:noFill');
        case 'solid': {
            const f = (0, xml_1.create)(doc, 'a:solidFill');
            f.appendChild(colorEl(doc, fill.color, fill.alpha));
            return f;
        }
        case 'gradient': {
            const g = (0, xml_1.create)(doc, 'a:gradFill', { rotWithShape: '1' });
            const lst = (0, xml_1.create)(doc, 'a:gsLst');
            for (const s of fill.stops) {
                const gs = (0, xml_1.create)(doc, 'a:gs', { pos: Math.round(Math.min(1, Math.max(0, s.pos)) * 100000) });
                gs.appendChild(colorEl(doc, s.color, s.alpha));
                lst.appendChild(gs);
            }
            g.appendChild(lst);
            g.appendChild((0, xml_1.create)(doc, 'a:lin', { ang: Math.round((((fill.angle % 360) + 360) % 360) * 60000), scaled: '0' }));
            return g;
        }
        case 'image':
            return reject(exports.MSG.imageFill);
    }
}
/** 塗り・線を変えられる要素（グループなら中の図形）。 */
function paintTargets(el, forLine) {
    if ((0, xml_1.is)(el, 'p:grpSp'))
        return (0, xml_1.descendants)(el).filter((e) => (forLine ? (0, xml_1.isAny)(e, ['p:sp', 'p:cxnSp', 'p:pic']) : (0, xml_1.is)(e, 'p:sp')));
    if (forLine ? (0, xml_1.isAny)(el, ['p:sp', 'p:cxnSp', 'p:pic']) : (0, xml_1.isAny)(el, ['p:sp', 'p:pic']))
        return [el];
    return reject(exports.MSG.noFill);
}
function setFill(txn, op) {
    if (op.fill.type === 'image')
        reject(exports.MSG.imageFill);
    for (const id of op.elementIds) {
        const f = findElement(txn, id);
        for (const t of paintTargets(f.ref.el, false)) {
            const spPr = (0, xml_1.ensureKid)(t, 'p:spPr');
            (0, xml_1.removeKids)(spPr, FILLS);
            (0, xml_1.insertChildInOrder)(spPr, fillEl(f.doc, op.fill));
        }
    }
}
function writeLine(spPr, line) {
    const doc = spPr.ownerDocument;
    const ln = (0, xml_1.ensureKid)(spPr, 'a:ln');
    (0, xml_1.removeKids)(ln, LINE_FILLS);
    if (line.type === 'none') {
        (0, xml_1.insertChildInOrder)(ln, (0, xml_1.create)(doc, 'a:noFill'));
        return;
    }
    (0, xml_1.setAttr)(ln, 'w', Math.max(0, Math.round(line.width)));
    const sf = (0, xml_1.create)(doc, 'a:solidFill');
    sf.appendChild(colorEl(doc, line.color, line.alpha));
    (0, xml_1.insertChildInOrder)(ln, sf);
    (0, xml_1.removeKids)(ln, ['a:prstDash', 'a:custDash']);
    if (line.dash)
        (0, xml_1.insertChildInOrder)(ln, (0, xml_1.create)(doc, 'a:prstDash', { val: line.dash }));
    for (const [name, v] of [
        ['a:headEnd', line.headEnd],
        ['a:tailEnd', line.tailEnd],
    ]) {
        const cur = (0, xml_1.kid)(ln, name);
        if (v) {
            if (cur)
                (0, xml_1.setAttr)(cur, 'type', v);
            else
                (0, xml_1.insertChildInOrder)(ln, (0, xml_1.create)(doc, name, { type: v }));
        }
        else if (cur) {
            (0, xml_1.remove)(cur);
        }
    }
}
function setLine(txn, op) {
    for (const id of op.elementIds) {
        const f = findElement(txn, id);
        for (const t of paintTargets(f.ref.el, true))
            writeLine((0, xml_1.ensureKid)(t, 'p:spPr'), op.line);
    }
}
// ---------------------------------------------------------------------------
// 追加
// ---------------------------------------------------------------------------
function appendToTree(root, el) {
    const tree = (0, elements_1.spTreeOf)(root);
    if (!tree)
        throw new Error('spTree がありません');
    (0, xml_1.insertChildInOrder)(tree, el);
}
function addShape(txn, op) {
    const s = findSlide(txn, op.slideId);
    const doc = txn.edit(s.slidePath);
    const root = (0, package_1.rootEl)(doc);
    const id = (0, elements_1.nextShapeId)(root);
    const shape = op.shape;
    let xml;
    if (shape.type === 'textBox')
        xml = (0, templates_1.textBoxXml)(id, `TextBox ${id - 1}`, op.xfrm, op.text);
    else if (shape.type === 'line')
        xml = (0, templates_1.lineXml)(id, `Connector ${id - 1}`, shape.preset, op.xfrm);
    else
        xml = (0, templates_1.shapeXml)(id, `Shape ${id - 1}`, shape.preset, op.xfrm, op.text);
    appendToTree(root, (0, xml_1.fragment)(doc, xml));
    return { changes: [], select: [`${op.slideId}:${id}`] };
}
const IMAGE_EXT = { 'image/png': 'png', 'image/jpeg': 'jpeg', 'image/gif': 'gif' };
exports.EXT_CONTENT_TYPE = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    jpg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    tif: 'image/tiff',
    tiff: 'image/tiff',
    svg: 'image/svg+xml',
    webp: 'image/webp',
    emf: 'image/x-emf',
    wmf: 'image/x-wmf',
    wdp: 'image/vnd.ms-photo',
    mp4: 'video/mp4',
    m4v: 'video/mp4',
    mov: 'video/quicktime',
    mp3: 'audio/mpeg',
    m4a: 'audio/mp4',
    wav: 'audio/wav',
};
/** ppt/media/<prefix><既存の最大番号 + 1>.<ext>。 */
function nextMediaPath(txn, ext, prefix = 'image') {
    let max = 0;
    const re = new RegExp(`^ppt/media/${prefix}(\\d+)\\.`, 'i');
    for (const p of txn.paths()) {
        const m = re.exec(p);
        if (m)
            max = Math.max(max, Number(m[1]));
    }
    return `ppt/media/${prefix}${max + 1}.${ext}`;
}
function addImage(txn, op) {
    const ext = (0, own_1.own)(IMAGE_EXT, op.mime);
    if (!ext)
        reject(exports.MSG.badImage);
    if (!op.data || op.data.length === 0)
        reject(exports.MSG.emptyImage);
    const s = findSlide(txn, op.slideId);
    const media = nextMediaPath(txn, ext);
    txn.put(media, new Uint8Array(op.data));
    txn.ensureDefault(ext, op.mime);
    const rId = txn.addRel(s.slidePath, package_1.REL.image, media);
    const doc = txn.edit(s.slidePath);
    const root = (0, package_1.rootEl)(doc);
    const id = (0, elements_1.nextShapeId)(root);
    appendToTree(root, (0, xml_1.fragment)(doc, (0, templates_1.pictureXml)(id, `Picture ${id - 1}`, rId, op.xfrm)));
    return { changes: [], select: [`${op.slideId}:${id}`] };
}
// ---------------------------------------------------------------------------
// 削除
// ---------------------------------------------------------------------------
function deleteElements(txn, op) {
    const bySlide = new Map();
    for (const id of op.elementIds) {
        const f = findElement(txn, id);
        if (f.ref.groups.length)
            reject(exports.MSG.groupChild);
        const entry = bySlide.get(f.slidePath) ?? { slidePath: f.slidePath, els: [] };
        if (!entry.els.includes(f.ref.el))
            entry.els.push(f.ref.el);
        bySlide.set(f.slidePath, entry);
    }
    for (const { slidePath, els } of bySlide.values()) {
        const doc = txn.edit(slidePath);
        const root = (0, package_1.rootEl)(doc);
        const ids = new Set();
        for (const el of els) {
            for (const c of [el, ...(0, xml_1.descendants)(el)].filter((e) => (0, xml_1.is)(e, 'p:cNvPr')))
                ids.add((0, xml_1.attr)(c, 'id') ?? '');
            const c0 = (0, inherit_1.cNvPrOf)(el);
            if (c0)
                ids.add((0, xml_1.attr)(c0, 'id') ?? '');
            (0, xml_1.remove)(el);
        }
        // 残っているコネクタの接続先を外す
        for (const cx of [...(0, xml_1.descendants)(root, 'a:stCxn'), ...(0, xml_1.descendants)(root, 'a:endCxn')]) {
            if (ids.has((0, xml_1.attr)(cx, 'id') ?? ''))
                (0, xml_1.remove)(cx);
        }
        // 残っている同じ cNvPr@id の図形（重複 ID）を対象から外す
        const remaining = new Set((0, elements_1.indexElements)((0, elements_1.spTreeOf)(root)).map((r) => r.cnvId));
        for (const id of [...ids])
            if (remaining.has(id))
                ids.delete(id);
        cleanTiming(root, ids);
        pruneRels(txn, slidePath, root);
    }
}
/** p:timing の中で削除した要素を対象にしているものを取り除く。空になれば p:timing ごと取り除く。 */
function cleanTiming(root, ids) {
    const timing = (0, xml_1.kid)(root, 'p:timing');
    if (!timing || ids.size === 0)
        return;
    // アニメーションの効果（presetClass を持つ cTn の p:par、なければ最も内側の p:par）
    for (const tgt of (0, xml_1.descendants)(timing, 'p:spTgt')) {
        if (!ids.has((0, xml_1.attr)(tgt, 'spid') ?? ''))
            continue;
        let effect;
        let innermost;
        for (let n = tgt.parentNode; n && n !== timing; n = n.parentNode) {
            if ((0, xml_1.is)(n, 'p:par')) {
                innermost ??= n;
                if ((0, xml_1.attr)((0, xml_1.kid)(n, 'p:cTn'), 'presetClass') !== undefined) {
                    effect = n;
                    break;
                }
            }
        }
        const target = effect ?? innermost;
        if (target)
            (0, xml_1.remove)(target);
        else {
            // p:par の外（p:iterate など）なら、その親の要素ごと
            (0, xml_1.remove)(tgt.parentNode);
        }
    }
    for (const b of (0, xml_1.descendants)(timing).filter((e) => (0, xml_1.attr)(e, 'spid') !== undefined && ids.has((0, xml_1.attr)(e, 'spid') ?? ''))) {
        (0, xml_1.remove)(b);
    }
    // 空になった p:par を取り除く（子の p:childTnLst が空）
    let changed = true;
    while (changed) {
        changed = false;
        for (const lst of (0, xml_1.descendants)(timing, 'p:childTnLst')) {
            if ((0, xml_1.kids)(lst).length === 0) {
                const cTn = lst.parentNode;
                const par = cTn?.parentNode;
                if (par && (0, xml_1.isAny)(par, ['p:par', 'p:seq', 'p:excl']) && par.parentNode && (0, xml_1.is)(par.parentNode, 'p:childTnLst')) {
                    (0, xml_1.remove)(par);
                    changed = true;
                }
                else {
                    (0, xml_1.remove)(lst);
                    changed = true;
                }
            }
        }
    }
    for (const bl of (0, xml_1.descendants)(timing, 'p:bldLst'))
        if ((0, xml_1.kids)(bl).length === 0)
            (0, xml_1.remove)(bl);
    const hasTargets = (0, xml_1.descendants)(timing).some((e) => (0, xml_1.attr)(e, 'spid') !== undefined);
    if (!hasTargets)
        (0, xml_1.remove)(timing);
}
const PRUNABLE = new Set([
    package_1.REL.image,
    package_1.REL.hyperlink,
    package_1.REL.slide,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/media',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
    'http://schemas.microsoft.com/office/2007/relationships/media',
    'http://schemas.microsoft.com/office/2007/relationships/hdphoto',
]);
function referencedRIds(root) {
    const out = new Set();
    for (const e of [root, ...(0, xml_1.descendants)(root)]) {
        for (let i = 0; i < e.attributes.length; i++) {
            const a = e.attributes.item(i);
            if (a.namespaceURI === xml_1.NS.r)
                out.add(a.value);
        }
    }
    return out;
}
/** スライドの XML から参照されなくなった画像・リンクの関係を取り除く。 */
function pruneRels(txn, slidePath, root) {
    const used = referencedRIds(root);
    for (const rel of (0, package_1.readRels)(txn, slidePath)) {
        if (PRUNABLE.has(rel.type) && !used.has(rel.id))
            txn.removeRel(slidePath, rel.id);
    }
}
// ---------------------------------------------------------------------------
// 複製・コピー・貼り付け
// ---------------------------------------------------------------------------
/** プレースホルダーのように位置を継承している要素には、明示的な xfrm を書く。 */
function ensureExplicitXfrm(ctx, orig, copy) {
    for (const b of branches(copy)) {
        const has = (0, xml_1.is)(b, 'p:graphicFrame') ? (0, xml_1.kid)(b, 'p:xfrm') : (0, xml_1.is)(b, 'p:grpSp') ? (0, xml_1.kid)((0, xml_1.kid)(b, 'p:grpSpPr'), 'a:xfrm') : (0, xml_1.kid)((0, xml_1.kid)(b, 'p:spPr'), 'a:xfrm');
        if (!has)
            writeXfrm(b, (0, parse_1.resolvedOwnXfrm)(ctx, orig), false);
    }
}
function offsetEl(el, d) {
    if (!d)
        return;
    for (const b of branches(el)) {
        const xf = (0, xml_1.is)(b, 'p:graphicFrame') ? (0, xml_1.kid)(b, 'p:xfrm') : (0, xml_1.is)(b, 'p:grpSp') ? (0, xml_1.kid)((0, xml_1.kid)(b, 'p:grpSpPr'), 'a:xfrm') : (0, xml_1.kid)((0, xml_1.kid)(b, 'p:spPr'), 'a:xfrm');
        const off = (0, xml_1.kid)(xf, 'a:off');
        if (!off)
            continue;
        (0, xml_1.setAttr)(off, 'x', (0, xml_1.emuCoord)(Number((0, xml_1.attr)(off, 'x') ?? 0) + d));
        (0, xml_1.setAttr)(off, 'y', (0, xml_1.emuCoord)(Number((0, xml_1.attr)(off, 'y') ?? 0) + d));
    }
}
/**
 * コピーした部分木の cNvPr@id を振り直し、コネクタの接続先を付け替える。
 * 図形の作成 ID（a16:creationId）は引き継がない（同じ ID の図形が 2 つあると、コメントのアンカーがどちらを指すか分からないため）。
 */
function renumber(copies, start) {
    let next = start;
    const map = new Map();
    const tops = [];
    for (const c of copies) {
        (0, creationIds_1.dropShapeCreationIds)(c);
        const top = (0, inherit_1.cNvPrOf)(c);
        for (const n of [c, ...(0, xml_1.descendants)(c)].filter((e) => (0, xml_1.is)(e, 'p:cNvPr'))) {
            const old = (0, xml_1.attr)(n, 'id') ?? '';
            const nid = String(next++);
            if (!map.has(old))
                map.set(old, nid);
            (0, xml_1.setAttr)(n, 'id', nid);
            if (n === top)
                tops.push(Number(nid));
        }
        if (!top)
            tops.push(0);
    }
    for (const c of copies) {
        for (const cx of [...(0, xml_1.descendants)(c, 'a:stCxn'), ...(0, xml_1.descendants)(c, 'a:endCxn')]) {
            const nid = map.get((0, xml_1.attr)(cx, 'id') ?? '');
            if (nid)
                (0, xml_1.setAttr)(cx, 'id', nid);
            else
                (0, xml_1.remove)(cx);
        }
    }
    return tops;
}
function duplicateElements(txn, op) {
    const select = [];
    const bySlide = new Map();
    for (const id of op.elementIds) {
        const f = findElement(txn, id);
        if (f.ref.groups.length)
            reject(exports.MSG.groupChild);
        if ((0, elements_1.containsForbidden)(f.ref.el))
            reject(exports.MSG.forbiddenCopy);
        const e = bySlide.get(f.slidePath) ?? { slideId: f.slideId, slidePath: f.slidePath, els: [] };
        e.els.push(f.ref.el);
        bySlide.set(f.slidePath, e);
    }
    for (const { slideId, slidePath, els } of bySlide.values()) {
        const ctx = (0, inherit_1.buildSlideCtx)(txn, slidePath);
        const root = (0, package_1.rootEl)(txn.edit(slidePath));
        const copies = els.map((el) => {
            const c = el.cloneNode(true);
            ensureExplicitXfrm(ctx, el, c);
            offsetEl(c, op.offset);
            return c;
        });
        const ids = renumber(copies, (0, elements_1.nextShapeId)(root));
        copies.forEach((c, i) => {
            appendToTree(root, c);
            select.push(`${slideId}:${ids[i]}`);
        });
    }
    return { changes: [], select };
}
function rAttrs(el) {
    const out = [];
    for (const e of [el, ...(0, xml_1.descendants)(el)]) {
        for (let i = 0; i < e.attributes.length; i++) {
            const a = e.attributes.item(i);
            if (a.namespaceURI === xml_1.NS.r)
                out.push({ owner: e, name: a.name, value: a.value });
        }
    }
    return out;
}
exports.COPY_PART_TYPES = [
    package_1.REL.image,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/media',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
    'http://schemas.microsoft.com/office/2007/relationships/media',
    'http://schemas.microsoft.com/office/2007/relationships/hdphoto',
];
function copyElements(txn, elementIds) {
    const payload = { xml: [], rels: [], parts: {} };
    const seenRel = new Set();
    for (const id of elementIds) {
        const f = findElement(txn, id);
        if (f.ref.groups.length)
            reject(exports.MSG.groupChild);
        if ((0, elements_1.containsForbidden)(f.ref.el))
            reject(exports.MSG.forbiddenCopy);
        const ctx = (0, inherit_1.buildSlideCtx)(txn, f.slidePath);
        const c = f.ref.el.cloneNode(true);
        ensureExplicitXfrm(ctx, f.ref.el, c);
        payload.xml.push((0, xml_1.serializeXml)(c));
        const rels = (0, package_1.readRels)(txn, f.slidePath);
        for (const a of rAttrs(c)) {
            const rel = rels.find((x) => x.id === a.value);
            if (!rel)
                continue;
            // 同じ rId がスライドごとに違う先を指しうるので、スライドパスも鍵に含める
            const key = `${f.slidePath}#${rel.id}`;
            if (seenRel.has(key))
                continue;
            seenRel.add(key);
            if (payload.rels.some((x) => x.rId === rel.id && (x.target !== rel.target || x.type !== rel.type))) {
                // 別のスライドの同じ rId と衝突する場合は、XML 側の rId を書き換えて区別する
                const alt = `${rel.id}_${payload.xml.length}`;
                for (const b of rAttrs(c))
                    if (b.value === rel.id)
                        b.owner.setAttributeNS(xml_1.NS.r, b.name, alt);
                payload.xml[payload.xml.length - 1] = (0, xml_1.serializeXml)(c);
                payload.rels.push({ rId: alt, type: rel.type, target: rel.target, external: rel.external });
            }
            else if (!payload.rels.some((x) => x.rId === rel.id)) {
                payload.rels.push({ rId: rel.id, type: rel.type, target: rel.target, external: rel.external });
            }
            if (!rel.external && exports.COPY_PART_TYPES.includes(rel.type) && txn.has(rel.target) && !payload.parts[rel.target]) {
                payload.parts[rel.target] = txn.bytes(rel.target);
            }
            // スライド間リンク：貼り付け先が同じファイルかどうかを見分けるため、参照先スライドの中身も持つ
            // （貼り付けではスライドのパーツはコピーしない。中身が同じスライドがあるときだけリンクを保つ）
            if (!rel.external && rel.type === package_1.REL.slide && txn.has(rel.target) && !payload.parts[rel.target]) {
                payload.parts[rel.target] = txn.bytes(rel.target);
            }
        }
    }
    // トランザクションは読み取りだけ（edit は呼んだが書き換えていない）
    return payload;
}
function removeRAttr(owner, name) {
    if ((0, xml_1.isAny)(owner, ['a:hlinkClick', 'a:hlinkMouseOver', 'a:hlinkHover'])) {
        (0, xml_1.remove)(owner);
        return;
    }
    // p:tags（顧客データ）は r:id が必須なので要素ごと取り除き、空になった p:custDataLst も取り除く
    if ((0, xml_1.is)(owner, 'p:tags')) {
        const lst = owner.parentNode;
        (0, xml_1.remove)(owner);
        if (lst && (0, xml_1.is)(lst, 'p:custDataLst') && (0, xml_1.kids)(lst).length === 0)
            (0, xml_1.remove)(lst);
        return;
    }
    const local = name.includes(':') ? name.slice(name.indexOf(':') + 1) : name;
    owner.removeAttributeNS(xml_1.NS.r, local);
}
/** 貼り付けで持ち込んでよい外部への関係：http(s)・mailto のハイパーリンク。 */
function isSafeExternalRel(rel) {
    return rel.type === package_1.REL.hyperlink && /^(https?:\/\/|mailto:)/i.test(rel.target);
}
/**
 * クリップボードのパーツ（コピー元のパス → バイト列）を貼り付け先のパーツに対応づける関数を作る。
 * 同じパスに同じ中身があればそれを、ほかのメディアに同じ中身があればそれを使い、なければ新しい名前で追加する。
 * スライド（スライド間リンクの参照先）は、同じパスに同じ中身があるときだけ同じものとみなし、なければ null。
 */
function partResolver(txn, parts) {
    const partMap = new Map();
    return (target, type) => {
        if (partMap.has(target))
            return partMap.get(target);
        let out = null;
        const bytes = (0, own_1.own)(parts, target);
        if (type === package_1.REL.slide) {
            const cur = txn.has(target) ? txn.bytes(target) : undefined;
            out = bytes && cur && (0, package_1.sameData)(cur, bytes) ? target : null;
        }
        else if (bytes) {
            const cur = txn.has(target) ? txn.bytes(target) : undefined;
            if (cur && (0, package_1.sameData)(cur, bytes))
                out = target;
            else
                out = sameMedia(txn, bytes) ?? null;
            if (!out) {
                const ext = (0, package_1.extname)(target) || 'bin';
                const prefix = /^([a-zA-Z]+)\d*$/.exec((0, package_1.basename)(target).replace(/\.[^.]*$/, ''))?.[1] ?? 'image';
                const p = (0, package_1.dirname)(target) === 'ppt/media' ? nextMediaPath(txn, ext, prefix) : uniquePath(txn, target);
                txn.put(p, new Uint8Array(bytes));
                txn.ensureDefault(ext, (0, own_1.own)(exports.EXT_CONTENT_TYPE, ext) ?? 'application/octet-stream');
                out = p;
            }
        }
        partMap.set(target, out);
        return out;
    };
}
/** ppt/media の中で中身が bytes と同じパーツ（なければ undefined）。 */
function sameMedia(txn, bytes) {
    for (const p of txn.paths()) {
        if (!p.startsWith('ppt/media/'))
            continue;
        const cur = txn.bytes(p);
        if (cur && cur.length === bytes.length && (0, package_1.sameData)(cur, bytes))
            return p;
    }
    return undefined;
}
function pasteElements(txn, op, clip) {
    if (!clip || clip.xml.length === 0)
        reject(exports.MSG.noClipboard);
    const s = findSlide(txn, op.slideId);
    const doc = txn.edit(s.slidePath);
    const root = (0, package_1.rootEl)(doc);
    const relMap = new Map();
    const resolvePart = partResolver(txn, clip.parts);
    const copies = [];
    for (const xml of clip.xml) {
        const el = (0, xml_1.fragment)(doc, xml);
        if (!(0, xml_1.isAny)(el, inherit_1.SHAPE_KINDS))
            reject(exports.MSG.noClipboard);
        if ((0, elements_1.containsForbidden)(el))
            reject(exports.MSG.forbiddenCopy);
        for (const a of rAttrs(el)) {
            let nid = relMap.get(a.value);
            if (nid === undefined) {
                const rel = clip.rels.find((x) => x.rId === a.value);
                if (!rel)
                    nid = null;
                // 外部への関係は、http(s)・mailto のハイパーリンクだけを持ち込む（画像のリンクや file: などは外す。ADR-0019）
                else if (rel.external)
                    nid = isSafeExternalRel(rel) ? txn.addRel(s.slidePath, rel.type, rel.target, true) : null;
                else {
                    const p = resolvePart(rel.target, rel.type);
                    nid = p ? txn.addRel(s.slidePath, rel.type, p) : null;
                }
                relMap.set(a.value, nid);
            }
            if (nid === null)
                removeRAttr(a.owner, a.name);
            else
                a.owner.setAttributeNS(xml_1.NS.r, a.name, nid);
        }
        offsetEl(el, op.offset);
        copies.push(el);
    }
    const ids = renumber(copies, (0, elements_1.nextShapeId)(root));
    const select = [];
    copies.forEach((c, i) => {
        appendToTree(root, c);
        select.push(`${op.slideId}:${ids[i]}`);
    });
    return { changes: [], select };
}
function uniquePath(txn, p) {
    const dir = (0, package_1.dirname)(p);
    const ext = (0, package_1.extname)(p);
    const stem = (0, package_1.basename)(p).replace(/\.[^.]*$/, '').replace(/\d+$/, '');
    let n = 1;
    for (;;) {
        const cand = `${dir}/${stem}${n}${ext ? '.' + ext : ''}`;
        if (!txn.has(cand))
            return cand;
        n++;
    }
}
// ---------------------------------------------------------------------------
// 重なり順
// ---------------------------------------------------------------------------
function reorderElement(txn, op) {
    if (op.elementIds.length === 0)
        return;
    const found = op.elementIds.map((id) => findElement(txn, id));
    for (const f of found)
        if (f.ref.groups.length)
            reject(exports.MSG.groupChild);
    const bySlide = new Map();
    for (const f of found) {
        const arr = bySlide.get(f.slidePath) ?? [];
        if (!arr.includes(f.ref.el))
            arr.push(f.ref.el);
        bySlide.set(f.slidePath, arr);
    }
    for (const [slidePath, els] of bySlide) {
        const tree = (0, elements_1.spTreeOf)((0, package_1.rootEl)(txn.edit(slidePath)));
        reorderIn(tree, els, op.order);
    }
}
function reorderIn(tree, selected, order) {
    const shapes = () => (0, xml_1.kids)(tree).filter((k) => (0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS));
    const sel = new Set(selected);
    const list = shapes();
    let next;
    if (order === 'front')
        next = [...list.filter((x) => !sel.has(x)), ...list.filter((x) => sel.has(x))];
    else if (order === 'back')
        next = [...list.filter((x) => sel.has(x)), ...list.filter((x) => !sel.has(x))];
    else if (order === 'forward') {
        next = [...list];
        for (let i = next.length - 2; i >= 0; i--) {
            if (sel.has(next[i]) && !sel.has(next[i + 1]))
                [next[i], next[i + 1]] = [next[i + 1], next[i]];
        }
    }
    else {
        next = [...list];
        for (let i = 1; i < next.length; i++) {
            if (sel.has(next[i]) && !sel.has(next[i - 1]))
                [next[i], next[i - 1]] = [next[i - 1], next[i]];
        }
    }
    if (next.every((x, i) => x === list[i]))
        return;
    // 図形が並ぶ範囲（最初の図形の位置）に並べ直す
    const anchor = list[0];
    const marker = tree.ownerDocument.createComment('reorder');
    tree.insertBefore(marker, anchor);
    for (const x of list)
        (0, xml_1.remove)(x);
    for (const x of next)
        tree.insertBefore(x, marker);
    (0, xml_1.remove)(marker);
}
// ---------------------------------------------------------------------------
// 代替テキスト
// ---------------------------------------------------------------------------
/** p:cNvPr@descr を書く（空文字なら取り除く）。mc:AlternateContent なら分岐ごとの図形に書く。グループの中の要素も対象にする。 */
function setAltText(txn, op) {
    const f = findElement(txn, op.elementId);
    const descr = typeof op.descr === 'string' ? op.descr : '';
    for (const b of branches(f.ref.el)) {
        const c = (0, inherit_1.cNvPrOf)(b);
        if (c)
            (0, xml_1.setAttr)(c, 'descr', descr === '' ? undefined : descr);
    }
}
// ---------------------------------------------------------------------------
// 画像の切り抜き（設計書 7 章「画像の切り抜き」）
// ---------------------------------------------------------------------------
/** 残す部分の最小（元の画像の 1%）。0 にすると画像が見えなくなり、Google スライドでも PowerPoint でも戻しにくい。 */
const MIN_CROP_KEEP = 1000;
/**
 * 切り抜きの値を確かめて整数にする。各辺は、元の画像の外に広げる負の値（PowerPoint が書くことがある）も、
 * 元の画像の幅・高さの 10 倍まで受け付ける。左右（上下）の合計は、1% 以上が残る範囲にする。
 */
function normalizeCrop(c) {
    const out = { l: 0, t: 0, r: 0, b: 0 };
    for (const k of ['l', 't', 'r', 'b']) {
        const v = c[k];
        if (typeof v !== 'number' || !Number.isFinite(v) || v < -1_000_000 || v >= 100_000)
            reject(exports.MSG.badCrop);
        out[k] = Math.round(v);
    }
    if (out.l + out.r > 100_000 - MIN_CROP_KEEP || out.t + out.b > 100_000 - MIN_CROP_KEEP)
        reject(exports.MSG.badCrop);
    return out;
}
/** p:blipFill/a:srcRect を書く。すべて 0 なら取り除く。a:srcRect は a:blip の直後（a:stretch / a:tile の前）に置く。 */
function writeSrcRect(blipFill, c) {
    let sr = (0, xml_1.kid)(blipFill, 'a:srcRect');
    if (!c.l && !c.t && !c.r && !c.b) {
        (0, xml_1.remove)(sr);
        return;
    }
    if (!sr) {
        sr = (0, xml_1.create)(blipFill.ownerDocument, 'a:srcRect');
        const blip = (0, xml_1.kid)(blipFill, 'a:blip');
        blipFill.insertBefore(sr, blip ? blip.nextSibling : blipFill.firstChild);
    }
    for (const k of ['l', 't', 'r', 'b'])
        (0, xml_1.setAttr)(sr, k, c[k] ? c[k] : undefined);
}
function cropImage(txn, op) {
    const f = findElement(txn, op.elementId);
    if (f.ref.groups.length)
        reject(exports.MSG.groupChild);
    const crop = normalizeCrop(op.crop);
    const pics = branches(f.ref.el).filter((b) => (0, xml_1.is)(b, 'p:pic'));
    if (!pics.length)
        reject(exports.MSG.notPicture);
    for (const pic of pics) {
        const bf = (0, xml_1.kid)(pic, 'p:blipFill');
        if (!bf)
            reject(exports.MSG.notPicture);
        writeSrcRect(bf, crop);
    }
    writeXfrm(f.ref.el, op.xfrm, false);
}
// ---------------------------------------------------------------------------
// 振り分け
// ---------------------------------------------------------------------------
function applyElementOp(txn, op, clip) {
    switch (op.type) {
        case 'setXfrm':
            setXfrm(txn, op);
            return { changes: [] };
        case 'setText':
            setText(txn, op);
            return { changes: [] };
        case 'setTextStyle':
            setTextStyle(txn, op);
            return { changes: [] };
        case 'setParagraphStyle':
            setParagraphStyle(txn, op);
            return { changes: [] };
        case 'setFill':
            setFill(txn, op);
            return { changes: [] };
        case 'setLine':
            setLine(txn, op);
            return { changes: [] };
        case 'addShape':
            return addShape(txn, op);
        case 'addImage':
            return addImage(txn, op);
        case 'deleteElements':
            deleteElements(txn, op);
            return { changes: [] };
        case 'duplicateElements':
            return duplicateElements(txn, op);
        case 'pasteElements':
            return pasteElements(txn, op, clip);
        case 'reorderElement':
            reorderElement(txn, op);
            return { changes: [] };
        case 'setAltText':
            setAltText(txn, op);
            return { changes: [] };
        case 'cropImage':
            cropImage(txn, op);
            return { changes: [] };
        default:
            return undefined;
    }
}
