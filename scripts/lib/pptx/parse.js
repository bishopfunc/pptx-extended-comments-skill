"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_DECK_LIMITS = exports.DISPLAYABLE_MIME = void 0;
exports.displayableMime = displayableMime;
exports.buildDeck = buildDeck;
exports.buildSlide = buildSlide;
exports.resolveFill = resolveFill;
exports.explicitOf = explicitOf;
exports.runChildren = runChildren;
exports.tableCellChain = tableCellChain;
exports.notesPathOf = notesPathOf;
exports.notesBodyShape = notesBodyShape;
exports.readNotes = readNotes;
exports.resolvedOwnXfrm = resolvedOwnXfrm;
const elements_1 = require("./elements.js");
const zipLimits_1 = require("./zipLimits.js");
const inherit_1 = require("./inherit.js");
const layouts_1 = require("./layouts.js");
const package_1 = require("./package.js");
const tableStyle_1 = require("./tableStyle.js");
const theme_1 = require("./theme.js");
const xml_1 = require("./xml.js");
const own_1 = require("../shared/own.js");
exports.DISPLAYABLE_MIME = {
    png: 'image/png',
    jpg: 'image/jpeg',
    jpeg: 'image/jpeg',
    jpe: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    svg: 'image/svg+xml',
    webp: 'image/webp',
};
function displayableMime(p) {
    return (0, own_1.own)(exports.DISPLAYABLE_MIME, (0, package_1.extname)(p));
}
const FILL_NAMES = ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:blipFill', 'a:pattFill', 'a:grpFill'];
exports.DEFAULT_DECK_LIMITS = { maxSlides: 2000, maxElements: 200_000 };
function countElements(els) {
    let n = 0;
    const stack = [...els];
    while (stack.length) {
        const e = stack.pop();
        n++;
        if (e.kind === 'group')
            stack.push(...e.children);
    }
    return n;
}
function buildDeck(r, limits = exports.DEFAULT_DECK_LIMITS) {
    const pres = (0, inherit_1.readPresentation)(r);
    if (pres.slides.length > limits.maxSlides)
        throw new zipLimits_1.ZipTooLargeError(`スライドが ${pres.slides.length} 枚あります`);
    let total = 0;
    const slides = pres.slides.map((s) => {
        const slide = buildSlide(r, s.id, s.path, pres.defaultTextStyle);
        total += countElements(slide.elements) + countElements(slide.inherited);
        if (total > limits.maxElements)
            throw new zipLimits_1.ZipTooLargeError('図形が多すぎます');
        return slide;
    });
    return { slideSize: pres.slideSize, slides, revision: 0, layouts: (0, layouts_1.readLayouts)(r, pres) };
}
function buildSlide(r, slideId, slidePath, defaultTextStyle) {
    const ctx = (0, inherit_1.buildSlideCtx)(r, slidePath, defaultTextStyle);
    const show = (0, inherit_1.showMasterShapes)(ctx);
    const inherited = [];
    if (show.master && ctx.masterRoot && ctx.masterPath) {
        inherited.push(...parseTree(ctx, ctx.masterRoot, ctx.masterPath, 'master', `${slideId}:master:`, true));
    }
    if (show.layout && ctx.layoutRoot && ctx.layoutPath) {
        inherited.push(...parseTree(ctx, ctx.layoutRoot, ctx.layoutPath, 'layout', `${slideId}:layout:`, true));
    }
    const elements = parseTree(ctx, ctx.slideRoot, slidePath, 'slide', `${slideId}:`, false);
    // スキップ（p:sld@show="0"）と、スライド自身の背景（p:cSld/p:bg）。どちらもなければ項目を省く
    const showAttr = (0, xml_1.attr)(ctx.slideRoot, 'show');
    const hidden = showAttr === '0' || showAttr === 'false';
    const ownBackground = !!(0, xml_1.path)(ctx.slideRoot, 'p:cSld/p:bg');
    return {
        id: slideId,
        partPath: slidePath,
        background: resolveBackground(ctx),
        inherited,
        elements,
        notes: readNotes(r, slidePath),
        layoutName: (0, xml_1.attr)((0, xml_1.kid)(ctx.layoutRoot, 'p:cSld'), 'name'),
        layoutPath: ctx.layoutPath,
        ...(hidden ? { hidden: true } : {}),
        ...(ownBackground ? { ownBackground: true } : {}),
    };
}
function parseTree(ctx, root, part, owner, idPrefix, skipPh) {
    const tree = (0, elements_1.spTreeOf)(root);
    const refs = (0, elements_1.indexElements)(tree);
    const byEl = new Map(refs.map((x) => [x.el, x]));
    const out = [];
    for (const k of (0, xml_1.kids)(tree)) {
        const ref = byEl.get(k);
        if (!ref)
            continue;
        if (skipPh && (0, inherit_1.phOf)(k))
            continue;
        out.push(parseElement(ctx, ref, byEl, part, owner, idPrefix, skipPh));
    }
    return out;
}
const NO_CAPS = { move: false, resize: false, rotate: false, editText: false };
function parseElement(ctx, ref, byEl, part, owner, idPrefix, inherited, groupFill) {
    const el = ref.el;
    const cNvPr = (0, inherit_1.cNvPrOf)(el);
    const id = idPrefix + ref.key;
    const name = (0, xml_1.attr)(cNvPr, 'name') ?? '';
    const src = owner === 'slide' || owner === 'layout' ? (0, inherit_1.phSources)(ctx, el, owner) : {};
    const own = (0, inherit_1.resolveXfrm)(el, src);
    const xfrm = (0, elements_1.toSlideXfrm)(own, ref.groups);
    const ph = (0, inherit_1.phOf)(el);
    const placeholder = ph ? { type: ph.type, idx: ph.idx } : undefined;
    const inGroup = ref.groups.length > 0;
    const caps = (editText, full = true) => inherited
        ? NO_CAPS
        : inGroup
            ? { move: false, resize: false, rotate: false, editText }
            : full
                ? { move: true, resize: true, rotate: true, editText }
                : { move: true, resize: false, rotate: false, editText: false };
    const descr = (0, xml_1.attr)(cNvPr, 'descr');
    const base = { id, name, xfrm, ...(placeholder ? { placeholder } : {}), ...(descr ? { altText: descr } : {}) };
    if ((0, xml_1.is)(el, 'p:grpSp')) {
        // 深すぎるグループの中は読まない（indexElements も列挙しない。ADR-0019）
        if (ref.groups.length >= elements_1.MAX_GROUP_DEPTH) {
            return { ...base, kind: 'unsupported', label: 'グループ（入れ子が深すぎます）', capabilities: caps(false, false) };
        }
        const fill = resolveFill(ctx, (0, xml_1.kids)((0, xml_1.kid)(el, 'p:grpSpPr')).find((k) => (0, xml_1.isAny)(k, FILL_NAMES)), part, ctx.color, groupFill);
        const children = [];
        for (const k of (0, xml_1.kids)(el)) {
            const cref = byEl.get(k);
            if (!cref)
                continue;
            children.push(parseElement(ctx, cref, byEl, part, owner, idPrefix, inherited, fill ?? groupFill));
        }
        const c = caps(false);
        return { ...base, kind: 'group', children, capabilities: c };
    }
    if ((0, xml_1.is)(el, 'p:sp') || (0, xml_1.is)(el, 'p:cxnSp')) {
        const isConnector = (0, xml_1.is)(el, 'p:cxnSp');
        const txBody = (0, xml_1.kid)(el, 'p:txBody');
        const text = txBody ? parseTextBody(ctx, el, txBody, src, owner) : undefined;
        return {
            ...base,
            kind: 'shape',
            geometry: resolveGeometry(el, src, xfrm),
            fill: isConnector ? { type: 'none' } : shapeFill(ctx, el, src, part, groupFill),
            line: shapeLine(ctx, el, src),
            ...(text ? { text } : {}),
            isTextBox: (0, xml_1.boolAttr)((0, xml_1.kid)((0, inherit_1.nvOf)(el), 'p:cNvSpPr'), 'txBox') ?? false,
            isConnector,
            capabilities: isConnector ? { ...caps(false), editText: false } : caps(inGroup ? !!txBody : true),
        };
    }
    if ((0, xml_1.is)(el, 'p:pic')) {
        const blip = (0, xml_1.path)(el, 'p:blipFill/a:blip');
        const embed = (0, xml_1.attr)(blip, 'r:embed');
        const link = (0, xml_1.attr)(blip, 'r:link');
        const rel = embed ? (0, package_1.relById)(ctx.r, part, embed) : undefined;
        const media = rel && !rel.external ? rel.target : undefined;
        if (!media || !displayableMime(media) || !ctx.r.has(media)) {
            const ext = media ? (0, package_1.extname)(media).toUpperCase() : link ? 'リンク' : '不明';
            return { ...base, kind: 'unsupported', label: `画像（${ext}）`, capabilities: caps(false, false) };
        }
        const sr = (0, xml_1.path)(el, 'p:blipFill/a:srcRect');
        const crop = sr
            ? { l: (0, xml_1.numAttr)(sr, 'l') ?? 0, t: (0, xml_1.numAttr)(sr, 't') ?? 0, r: (0, xml_1.numAttr)(sr, 'r') ?? 0, b: (0, xml_1.numAttr)(sr, 'b') ?? 0 }
            : undefined;
        return {
            ...base,
            kind: 'picture',
            mediaPath: media,
            ...(crop && (crop.l || crop.t || crop.r || crop.b) ? { crop } : {}),
            line: shapeLine(ctx, el, src),
            capabilities: caps(false),
        };
    }
    // graphicFrame、AlternateContent、contentPart
    const gk = (0, elements_1.graphicKind)(el);
    if (gk === 'table') {
        const frame = (0, xml_1.is)(el, 'mc:AlternateContent') ? (0, inherit_1.altContentMain)(el) : el;
        const tbl = (0, xml_1.path)(frame, 'a:graphic/a:graphicData/a:tbl');
        if (tbl) {
            const t = parseTable(ctx, tbl, part, owner);
            // 拡大・縮小したグループの中の表は、グループの変換に合わせて列の幅と行の高さを見せる（グループ解除の結果と同じ）
            if (inGroup && own.cx > 0 && own.cy > 0) {
                const sx = xfrm.cx / own.cx;
                const sy = xfrm.cy / own.cy;
                if (sx !== 1)
                    t.columns = t.columns.map((w) => w * sx);
                if (sy !== 1)
                    for (const r of t.rows)
                        r.height *= sy;
            }
            // セルのテキストは、スライドのトップレベルの表だけ編集できる（大きさと回転は変えない）
            const capabilities = inherited || inGroup ? caps(false, false) : { move: true, resize: false, rotate: false, editText: true };
            // 表の大きさは列の幅と行の高さの合計で決まる（PowerPoint と同じ）。Google スライドが書き出す pptx は
            // p:xfrm の ext にダミーの値（3000000 など）を入れるので、枠の値は使わない。
            const sumW = t.columns.reduce((a, w) => a + w, 0);
            const sumH = t.rows.reduce((a, r) => a + r.height, 0);
            const tableXfrm = !inGroup && sumW > 0 && sumH > 0 ? { ...base.xfrm, cx: sumW, cy: sumH } : base.xfrm;
            return { ...base, xfrm: tableXfrm, kind: 'table', ...t, capabilities };
        }
    }
    let label = 'オブジェクト';
    if (gk === 'table')
        label = '表';
    else if (gk === 'chart')
        label = 'グラフ';
    else if (gk === 'diagram')
        label = 'SmartArt';
    else if (gk === 'ole')
        label = 'OLE';
    else if ((0, xml_1.is)(el, 'p:contentPart'))
        label = 'インク';
    else if ((0, xml_1.is)(el, 'mc:AlternateContent')) {
        const main = (0, inherit_1.altContentMain)(el);
        if (main && (0, xml_1.is)(main, 'p:sp'))
            label = '図形';
    }
    return { ...base, kind: 'unsupported', label, capabilities: caps(false, false) };
}
// ---------------------------------------------------------------------------
// 形状
// ---------------------------------------------------------------------------
function resolveGeometry(el, src, xfrm) {
    const spPrs = [(0, xml_1.kid)(el, 'p:spPr'), (0, xml_1.kid)(src.layoutPh, 'p:spPr'), (0, xml_1.kid)(src.masterPh, 'p:spPr')];
    for (const spPr of spPrs) {
        const prst = (0, xml_1.kid)(spPr, 'a:prstGeom');
        if (prst) {
            const adjust = {};
            for (const gd of (0, xml_1.kids)((0, xml_1.kid)(prst, 'a:avLst'))) {
                const m = /^val\s+(-?\d+(\.\d+)?)$/.exec((0, xml_1.attr)(gd, 'fmla') ?? '');
                if (m)
                    adjust[(0, xml_1.attr)(gd, 'name') ?? ''] = Number(m[1]);
            }
            return Object.keys(adjust).length
                ? { type: 'preset', preset: (0, xml_1.attr)(prst, 'prst') ?? 'rect', adjust }
                : { type: 'preset', preset: (0, xml_1.attr)(prst, 'prst') ?? 'rect' };
        }
        const cust = (0, xml_1.kid)(spPr, 'a:custGeom');
        if (cust)
            return { type: 'custom', paths: customPaths(cust, xfrm.cx, xfrm.cy) };
    }
    return { type: 'preset', preset: 'rect' };
}
function customPaths(cust, shapeW, shapeH) {
    const out = [];
    for (const p of (0, xml_1.kids)((0, xml_1.kid)(cust, 'a:pathLst'))) {
        // 細工した値（1e308、1e-300 など）でも、座標は EMU の範囲に収める（ADR-0019）
        const w = (0, xml_1.clampSize)((0, xml_1.numAttr)(p, 'w') ?? shapeW);
        const h = (0, xml_1.clampSize)((0, xml_1.numAttr)(p, 'h') ?? shapeH);
        const guides = evalGuides(cust, w >= 1 ? w : shapeW, h >= 1 ? h : shapeH);
        const val = (s) => {
            if (s === undefined)
                return 0;
            const n = Number(s);
            if (Number.isFinite(n))
                return (0, xml_1.clampCoord)(n);
            return (0, xml_1.clampCoord)(guides.get(s) ?? 0);
        };
        const pt = (e) => [val((0, xml_1.attr)(e, 'x')), val((0, xml_1.attr)(e, 'y'))];
        const d = [];
        let cur = [0, 0];
        let start = [0, 0];
        for (const c of (0, xml_1.kids)(p)) {
            const pts = (0, xml_1.kids)(c).filter((k) => (0, xml_1.is)(k, 'a:pt')).map(pt);
            switch (c.localName) {
                case 'moveTo':
                    if (pts[0]) {
                        cur = [pts[0][0], pts[0][1]];
                        start = cur;
                        d.push(`M${cur[0]} ${cur[1]}`);
                    }
                    break;
                case 'lnTo':
                    if (pts[0]) {
                        cur = [pts[0][0], pts[0][1]];
                        d.push(`L${cur[0]} ${cur[1]}`);
                    }
                    break;
                case 'cubicBezTo':
                    if (pts.length >= 3) {
                        d.push(`C${pts.map((q) => `${q[0]} ${q[1]}`).join(' ')}`);
                        cur = [pts[2][0], pts[2][1]];
                    }
                    break;
                case 'quadBezTo':
                    if (pts.length >= 2) {
                        d.push(`Q${pts.map((q) => `${q[0]} ${q[1]}`).join(' ')}`);
                        cur = [pts[1][0], pts[1][1]];
                    }
                    break;
                case 'arcTo': {
                    const wR = val((0, xml_1.attr)(c, 'wR'));
                    const hR = val((0, xml_1.attr)(c, 'hR'));
                    const st = (val((0, xml_1.attr)(c, 'stAng')) / 60000) * (Math.PI / 180);
                    const sw = (val((0, xml_1.attr)(c, 'swAng')) / 60000) * (Math.PI / 180);
                    const cx0 = cur[0] - wR * Math.cos(st);
                    const cy0 = cur[1] - hR * Math.sin(st);
                    const arcTo = (a) => {
                        const ex = cx0 + wR * Math.cos(a);
                        const ey = cy0 + hR * Math.sin(a);
                        return [round3(ex), round3(ey)];
                    };
                    const segs = Math.abs(sw) >= 2 * Math.PI - 1e-6 ? 2 : 1;
                    for (let i = 1; i <= segs; i++) {
                        const part = sw / segs;
                        const end = arcTo(st + part * i);
                        d.push(`A${wR} ${hR} 0 ${Math.abs(part) > Math.PI ? 1 : 0} ${part > 0 ? 1 : 0} ${end[0]} ${end[1]}`);
                        cur = end;
                    }
                    break;
                }
                case 'close':
                    d.push('Z');
                    cur = start;
                    break;
                default:
                    break;
            }
        }
        out.push({
            // 1 EMU 未満の幅・高さで割らない（描画で極端な倍率にならないように）
            w: w >= 1 ? w : shapeW,
            h: h >= 1 ? h : shapeH,
            fill: (0, xml_1.attr)(p, 'fill') !== 'none',
            stroke: (0, xml_1.attr)(p, 'stroke') !== '0' && (0, xml_1.attr)(p, 'stroke') !== 'false',
            d: d.join(' '),
        });
    }
    return out;
}
function round3(n) {
    return Math.round(n * 1000) / 1000;
}
function evalGuides(cust, w, h) {
    const g = new Map();
    const ss = Math.min(w, h);
    const base = {
        w,
        h,
        l: 0,
        t: 0,
        r: w,
        b: h,
        hc: w / 2,
        vc: h / 2,
        wd2: w / 2,
        wd3: w / 3,
        wd4: w / 4,
        wd5: w / 5,
        wd6: w / 6,
        wd8: w / 8,
        wd10: w / 10,
        wd32: w / 32,
        hd2: h / 2,
        hd3: h / 3,
        hd4: h / 4,
        hd5: h / 5,
        hd6: h / 6,
        hd8: h / 8,
        hd10: h / 10,
        hd32: h / 32,
        ss,
        ls: Math.max(w, h),
        ssd2: ss / 2,
        ssd4: ss / 4,
        ssd6: ss / 6,
        ssd8: ss / 8,
        ssd16: ss / 16,
        ssd32: ss / 32,
        cd2: 10800000,
        cd4: 5400000,
        cd8: 2700000,
        '3cd4': 16200000,
        '3cd8': 8100000,
        '5cd8': 13500000,
        '7cd8': 18900000,
    };
    for (const [k, v] of Object.entries(base))
        g.set(k, v);
    const v = (s) => {
        if (s === undefined)
            return 0;
        const n = Number(s);
        if (Number.isFinite(n))
            return n;
        return g.get(s) ?? 0;
    };
    const deg = (x) => (x / 60000) * (Math.PI / 180);
    for (const gd of [...(0, xml_1.kids)((0, xml_1.kid)(cust, 'a:avLst')), ...(0, xml_1.kids)((0, xml_1.kid)(cust, 'a:gdLst'))]) {
        const name = (0, xml_1.attr)(gd, 'name');
        const parts = ((0, xml_1.attr)(gd, 'fmla') ?? '').trim().split(/\s+/);
        const [op, a1, a2, a3] = parts;
        const x = v(a1);
        const y = v(a2);
        const z = v(a3);
        let r = 0;
        switch (op) {
            case 'val':
                r = x;
                break;
            case '*/':
                r = z ? (x * y) / z : 0;
                break;
            case '+-':
                r = x + y - z;
                break;
            case '+/':
                r = z ? (x + y) / z : 0;
                break;
            case '?:':
                r = x > 0 ? y : z;
                break;
            case 'abs':
                r = Math.abs(x);
                break;
            case 'sqrt':
                r = Math.sqrt(Math.max(0, x));
                break;
            case 'max':
                r = Math.max(x, y);
                break;
            case 'min':
                r = Math.min(x, y);
                break;
            case 'pin':
                r = y < x ? x : y > z ? z : y;
                break;
            case 'mod':
                r = Math.sqrt(x * x + y * y + z * z);
                break;
            case 'sin':
                r = x * Math.sin(deg(y));
                break;
            case 'cos':
                r = x * Math.cos(deg(y));
                break;
            case 'tan':
                r = x * Math.tan(deg(y));
                break;
            case 'at2':
                r = (Math.atan2(y, x) * 180 * 60000) / Math.PI;
                break;
            case 'cat2':
                r = x * Math.cos(Math.atan2(z, y));
                break;
            case 'sat2':
                r = x * Math.sin(Math.atan2(z, y));
                break;
            default:
                r = 0;
        }
        if (name)
            g.set(name, Number.isFinite(r) ? r : 0);
    }
    return g;
}
// ---------------------------------------------------------------------------
// 塗りと線
// ---------------------------------------------------------------------------
function fillElOf(container) {
    return (0, xml_1.kids)(container).find((k) => (0, xml_1.isAny)(k, FILL_NAMES));
}
function toFillColor(c) {
    if (!c)
        return { type: 'none' };
    return c.alpha < 1 ? { type: 'solid', color: (0, theme_1.cssColor)(c), alpha: c.alpha } : { type: 'solid', color: (0, theme_1.cssColor)(c) };
}
/** 塗りの要素を Fill にする。part は blip の r:embed を解決するパーツ。 */
function resolveFill(ctx, fillEl, part, color, groupFill) {
    if (!fillEl)
        return undefined;
    switch (fillEl.localName) {
        case 'noFill':
            return { type: 'none' };
        case 'solidFill':
            return toFillColor((0, theme_1.colorOfParent)(fillEl, color));
        case 'gradFill': {
            const stops = (0, xml_1.kids)((0, xml_1.kid)(fillEl, 'a:gsLst')).map((gs) => {
                const c = (0, theme_1.colorOfParent)(gs, color) ?? { hex: '000000', alpha: 1 };
                const s = { pos: ((0, xml_1.numAttr)(gs, 'pos') ?? 0) / 100000, color: (0, theme_1.cssColor)(c) };
                if (c.alpha < 1)
                    s.alpha = c.alpha;
                return s;
            });
            stops.sort((a, b) => a.pos - b.pos);
            const lin = (0, xml_1.kid)(fillEl, 'a:lin');
            if (!stops.length)
                return { type: 'none' };
            return { type: 'gradient', stops, angle: ((0, xml_1.numAttr)(lin, 'ang') ?? 0) / 60000 };
        }
        case 'blipFill': {
            const embed = (0, xml_1.attr)((0, xml_1.kid)(fillEl, 'a:blip'), 'r:embed');
            const rel = embed ? (0, package_1.relById)(ctx.r, part, embed) : undefined;
            if (rel && !rel.external && displayableMime(rel.target) && ctx.r.has(rel.target)) {
                return { type: 'image', mediaPath: rel.target };
            }
            return { type: 'solid', color: '#D9D9D9' };
        }
        case 'pattFill':
            return toFillColor((0, theme_1.colorOfParent)((0, xml_1.kid)(fillEl, 'a:fgClr'), color));
        case 'grpFill':
            return groupFill ?? { type: 'none' };
        default:
            return undefined;
    }
}
function styleRef(ctx, el, name) {
    const ref = (0, xml_1.path)(el, `p:style/${name}`);
    if (!ref)
        return undefined;
    const phClr = (0, theme_1.colorOfParent)(ref, ctx.color);
    return { idx: (0, xml_1.numAttr)(ref, 'idx') ?? 0, color: { ...ctx.color, phClr } };
}
function shapeFill(ctx, el, src, part, groupFill) {
    const own = resolveFill(ctx, fillElOf((0, xml_1.kid)(el, 'p:spPr')), part, ctx.color, groupFill);
    if (own)
        return own;
    // プレースホルダーはレイアウト・マスターの spPr の塗りも継承する
    if (src.layoutPh) {
        const f = resolveFill(ctx, fillElOf((0, xml_1.kid)(src.layoutPh, 'p:spPr')), ctx.layoutPath ?? part, ctx.color, groupFill);
        if (f)
            return f;
    }
    if (src.masterPh) {
        const f = resolveFill(ctx, fillElOf((0, xml_1.kid)(src.masterPh, 'p:spPr')), ctx.masterPath ?? part, ctx.color, groupFill);
        if (f)
            return f;
    }
    const ref = styleRef(ctx, el, 'a:fillRef');
    if (ref && ref.idx > 0) {
        const list = ref.idx >= 1001 ? ctx.theme.bgFillStyles : ctx.theme.fillStyles;
        const entry = list[(ref.idx >= 1001 ? ref.idx - 1001 : ref.idx - 1)];
        const f = resolveFill(ctx, entry, part, ref.color, groupFill);
        if (f)
            return f;
    }
    return { type: 'none' };
}
function lineFillColor(ln, color) {
    const f = (0, xml_1.kids)(ln).find((k) => (0, xml_1.isAny)(k, ['a:noFill', 'a:solidFill', 'a:gradFill', 'a:pattFill']));
    if (!f)
        return undefined;
    if ((0, xml_1.is)(f, 'a:noFill'))
        return { none: true };
    if ((0, xml_1.is)(f, 'a:solidFill'))
        return { none: false, c: (0, theme_1.colorOfParent)(f, color) };
    if ((0, xml_1.is)(f, 'a:gradFill'))
        return { none: false, c: (0, theme_1.colorOfParent)((0, xml_1.kids)((0, xml_1.kid)(f, 'a:gsLst'))[0], color) };
    return { none: false, c: (0, theme_1.colorOfParent)((0, xml_1.kid)(f, 'a:fgClr'), color) };
}
const DASHES = new Set(['solid', 'dash', 'dot', 'dashDot', 'lgDash', 'sysDash', 'sysDot']);
function shapeLine(ctx, el, src) {
    const lns = [];
    const own = (0, xml_1.path)(el, 'p:spPr/a:ln');
    if (own)
        lns.push({ ln: own, color: ctx.color });
    for (const ph of [src.layoutPh, src.masterPh]) {
        const l = (0, xml_1.path)(ph, 'p:spPr/a:ln');
        if (l)
            lns.push({ ln: l, color: ctx.color });
    }
    const ref = styleRef(ctx, el, 'a:lnRef');
    if (ref && ref.idx > 0) {
        const entry = ctx.theme.lineStyles[ref.idx - 1];
        if (entry)
            lns.push({ ln: entry, color: ref.color });
    }
    return lineOf(lns);
}
/** 線の要素の並び（優先度の高い順）から Line を作る。塗りは最初に塗りを持つ要素、幅・破線は最初に書かれている要素。 */
function lineOf(lns) {
    if (!lns.length)
        return { type: 'none' };
    let fill;
    for (const l of lns) {
        fill = lineFillColor(l.ln, l.color);
        if (fill)
            break;
    }
    if (!fill || fill.none || !fill.c)
        return { type: 'none' };
    const first = (get) => {
        for (const l of lns) {
            const v = get(l.ln);
            if (v !== undefined)
                return v;
        }
        return undefined;
    };
    const width = first((ln) => (0, xml_1.numAttr)(ln, 'w')) ?? 9525;
    const dash = first((ln) => (0, xml_1.attr)((0, xml_1.kid)(ln, 'a:prstDash'), 'val'));
    const head = first((ln) => (0, xml_1.attr)((0, xml_1.kid)(ln, 'a:headEnd'), 'type'));
    const tail = first((ln) => (0, xml_1.attr)((0, xml_1.kid)(ln, 'a:tailEnd'), 'type'));
    const line = { type: 'solid', color: (0, theme_1.cssColor)(fill.c), width };
    if (fill.c.alpha < 1)
        line.alpha = fill.c.alpha;
    if (dash && DASHES.has(dash))
        line.dash = dash;
    if (head && head !== 'none')
        line.headEnd = head;
    if (tail && tail !== 'none')
        line.tailEnd = tail;
    return line;
}
// ---------------------------------------------------------------------------
// 背景
// ---------------------------------------------------------------------------
function resolveBackground(ctx) {
    const found = (0, inherit_1.findBackground)(ctx);
    if (!found)
        return { type: 'solid', color: '#FFFFFF' };
    const { bg, part } = found;
    if ((0, xml_1.is)(bg, 'p:bgPr'))
        return resolveFill(ctx, fillElOf(bg), part, ctx.color) ?? { type: 'solid', color: '#FFFFFF' };
    // p:bgRef
    const idx = (0, xml_1.numAttr)(bg, 'idx') ?? 0;
    const phClr = (0, theme_1.colorOfParent)(bg, ctx.color);
    const color = { ...ctx.color, phClr };
    if (idx === 0)
        return { type: 'none' };
    const entry = idx >= 1001 ? ctx.theme.bgFillStyles[idx - 1001] : ctx.theme.fillStyles[idx - 1];
    const f = resolveFill(ctx, entry, part, color);
    if (f)
        return f;
    return toFillColor(phClr);
}
function isFontRef(e) {
    return e.fontRef !== undefined;
}
function lookupP(pPr, entries, get) {
    if (pPr) {
        const v = get(pPr);
        if (v !== undefined)
            return v;
    }
    for (const e of entries) {
        if (isFontRef(e))
            continue;
        const v = get(e);
        if (v !== undefined)
            return v;
    }
    return undefined;
}
function lookupR(rPr, entries, get, fromFontRef) {
    if (rPr) {
        const v = get(rPr);
        if (v !== undefined)
            return v;
    }
    for (const e of entries) {
        if (isFontRef(e)) {
            const v = fromFontRef?.(e.fontRef);
            if (v !== undefined)
                return v;
            continue;
        }
        const d = (0, xml_1.kid)(e, 'a:defRPr');
        if (!d)
            continue;
        const v = get(d);
        if (v !== undefined)
            return v;
    }
    return undefined;
}
const getters = {
    bold: (e) => (0, xml_1.boolAttr)(e, 'b'),
    italic: (e) => (0, xml_1.boolAttr)(e, 'i'),
    underline: (e) => {
        const u = (0, xml_1.attr)(e, 'u');
        return u === undefined ? undefined : u !== 'none';
    },
    strike: (e) => {
        const s = (0, xml_1.attr)(e, 'strike');
        return s === undefined ? undefined : s !== 'noStrike';
    },
    size: (e) => {
        const s = (0, xml_1.numAttr)(e, 'sz');
        return s === undefined ? undefined : s / 100;
    },
    baseline: (e) => (0, xml_1.numAttr)(e, 'baseline'),
    latinRaw: (e) => {
        const t = (0, xml_1.attr)((0, xml_1.kid)(e, 'a:latin'), 'typeface');
        return t ? t : undefined;
    },
    eaRaw: (e) => {
        const t = (0, xml_1.attr)((0, xml_1.kid)(e, 'a:ea'), 'typeface');
        return t ? t : undefined;
    },
};
function colorGetter(color) {
    return (e) => {
        const f = (0, xml_1.kid)(e, 'a:solidFill');
        if (f)
            return (0, theme_1.colorOfParent)(f, color);
        const g = (0, xml_1.kid)(e, 'a:gradFill');
        if (g)
            return (0, theme_1.colorOfParent)((0, xml_1.kids)((0, xml_1.kid)(g, 'a:gsLst'))[0], color);
        return undefined;
    };
}
/** a:rPr に書かれている値だけを RunStyle にする（テーマのフォントは解決する）。 */
function explicitOf(rPr, theme, color) {
    const s = {};
    if (!rPr)
        return s;
    const b = getters.bold(rPr);
    if (b !== undefined)
        s.bold = b;
    const i = getters.italic(rPr);
    if (i !== undefined)
        s.italic = i;
    const u = getters.underline(rPr);
    if (u !== undefined)
        s.underline = u;
    const st = getters.strike(rPr);
    if (st !== undefined)
        s.strike = st;
    const sz = getters.size(rPr);
    if (sz !== undefined)
        s.size = sz;
    const c = colorGetter(color)(rPr);
    if (c)
        s.color = (0, theme_1.cssColor)(c);
    const lt = getters.latinRaw(rPr);
    if (lt)
        s.latinFont = (0, theme_1.resolveThemeFont)(theme, lt, 'latin');
    const ea = getters.eaRaw(rPr);
    if (ea)
        s.eaFont = (0, theme_1.resolveThemeFont)(theme, ea, 'ea');
    const bl = getters.baseline(rPr);
    if (bl !== undefined)
        s.baseline = bl;
    return s;
}
function resolveRun(ctx, rPr, entries, hasLink) {
    const theme = ctx.theme;
    const colorGet = colorGetter(ctx.color);
    let color = lookupR(rPr, entries, colorGet, (f) => f.color);
    if (hasLink && !(rPr && colorGet(rPr)))
        color = { hex: (0, own_1.own)(ctx.theme.colors, (0, own_1.own)(ctx.clrMap, 'hlink') ?? 'hlink') ?? '0000FF', alpha: 1 };
    if (!color) {
        const tx1 = ctx.clrMap.tx1 ?? 'dk1';
        color = { hex: (0, own_1.own)(ctx.theme.colors, tx1) ?? '000000', alpha: 1 };
    }
    const latinRaw = lookupR(rPr, entries, getters.latinRaw, (f) => f.latin) ?? '+mn-lt';
    const eaRaw = lookupR(rPr, entries, getters.eaRaw, (f) => f.ea) ?? '+mn-ea';
    let ea = (0, theme_1.resolveThemeFont)(theme, eaRaw, 'ea') ?? '';
    if (!ea)
        ea = theme.minor.jpan || theme.minor.ea;
    let underline = lookupR(rPr, entries, getters.underline) ?? false;
    if (hasLink && rPr && getters.underline(rPr) === undefined)
        underline = true;
    return {
        bold: lookupR(rPr, entries, getters.bold, (f) => f.bold) ?? false,
        italic: lookupR(rPr, entries, getters.italic, (f) => f.italic) ?? false,
        underline,
        strike: lookupR(rPr, entries, getters.strike) ?? false,
        size: lookupR(rPr, entries, getters.size) ?? 18,
        color: (0, theme_1.cssColor)(color),
        latinFont: (0, theme_1.resolveThemeFont)(theme, latinRaw, 'latin') ?? theme.minor.latin,
        eaFont: ea,
        baseline: lookupR(rPr, entries, getters.baseline) ?? 0,
    };
}
const RUN_KINDS = ['a:r', 'a:br', 'a:fld'];
function runChildren(p) {
    return (0, xml_1.kids)(p).filter((k) => (0, xml_1.isAny)(k, RUN_KINDS));
}
function spacingPt(el, sizePt) {
    if (!el)
        return undefined;
    const pts = (0, xml_1.kid)(el, 'a:spcPts');
    if (pts)
        return ((0, xml_1.numAttr)(pts, 'val') ?? 0) / 100;
    const pct = (0, xml_1.kid)(el, 'a:spcPct');
    if (pct)
        return (((0, xml_1.numAttr)(pct, 'val') ?? 0) / 100000) * sizePt * 1.2;
    return undefined;
}
function parseTextBody(ctx, el, txBody, src, owner) {
    const chain = (0, inherit_1.textChain)(ctx, el, src, owner);
    const bodyAttr = (get) => {
        for (const b of chain.bodyPrs) {
            const v = get(b);
            if (v !== undefined)
                return v;
        }
        return undefined;
    };
    const anchorRaw = bodyAttr((b) => (0, xml_1.attr)(b, 'anchor')) ?? 't';
    const anchor = anchorRaw === 'ctr' ? 'ctr' : anchorRaw === 'b' ? 'b' : 't';
    const vertRaw = bodyAttr((b) => (0, xml_1.attr)(b, 'vert')) ?? 'horz';
    const vertical = ['horz', 'vert', 'vert270', 'eaVert'].includes(vertRaw)
        ? vertRaw
        : 'other';
    const autofit = bodyAttr((b) => (0, xml_1.kids)(b).find((k) => (0, xml_1.isAny)(k, ['a:noAutofit', 'a:normAutofit', 'a:spAutoFit'])));
    const norm = autofit && (0, xml_1.is)(autofit, 'a:normAutofit') ? autofit : undefined;
    const paragraphs = (0, xml_1.kids)(txBody)
        .filter((k) => (0, xml_1.is)(k, 'a:p'))
        .map((p, pi) => parseParagraph(ctx, chain, p, pi));
    return {
        paragraphs,
        anchor,
        insets: {
            l: bodyAttr((b) => (0, xml_1.numAttr)(b, 'lIns')) ?? 91440,
            t: bodyAttr((b) => (0, xml_1.numAttr)(b, 'tIns')) ?? 45720,
            r: bodyAttr((b) => (0, xml_1.numAttr)(b, 'rIns')) ?? 91440,
            b: bodyAttr((b) => (0, xml_1.numAttr)(b, 'bIns')) ?? 45720,
        },
        wrap: (bodyAttr((b) => (0, xml_1.attr)(b, 'wrap')) ?? 'square') !== 'none',
        fontScale: norm ? ((0, xml_1.numAttr)(norm, 'fontScale') ?? 100000) / 100000 : 1,
        lineSpaceReduction: norm ? ((0, xml_1.numAttr)(norm, 'lnSpcReduction') ?? 0) / 100000 : 0,
        vertical,
    };
}
function parseParagraph(ctx, chain, p, pi) {
    const pPr = (0, xml_1.kid)(p, 'a:pPr');
    const level = Math.min(8, Math.max(0, (0, xml_1.numAttr)(pPr, 'lvl') ?? 0));
    const entries = (0, inherit_1.levelPPrs)(chain, level);
    const algnRaw = lookupP(pPr, entries, (e) => (0, xml_1.attr)(e, 'algn')) ?? 'l';
    const align = algnRaw === 'ctr' ? 'ctr' : algnRaw === 'r' ? 'r' : algnRaw === 'just' || algnRaw === 'dist' || algnRaw === 'justLow' ? 'just' : 'l';
    // ランの書式は、段落自身の a:pPr/a:defRPr も継承元に含める（Google スライドはこれを使う）。
    const runEntries = pPr ? [pPr, ...entries] : entries;
    const runs = [];
    runChildren(p).forEach((r, ri) => {
        const rPr = (0, xml_1.kid)(r, 'a:rPr');
        const hasLink = !!(0, xml_1.kid)(rPr, 'a:hlinkClick');
        const kind = (0, xml_1.is)(r, 'a:br') ? 'break' : (0, xml_1.is)(r, 'a:fld') ? 'field' : 'text';
        const run = {
            kind,
            text: kind === 'break' ? '\n' : (0, xml_1.textOf)((0, xml_1.kid)(r, 'a:t')),
            resolved: resolveRun(ctx, rPr, runEntries, hasLink),
            explicit: explicitOf(rPr, ctx.theme, ctx.color),
            source: { p: pi, r: ri },
        };
        if (hasLink)
            run.hasLink = true;
        runs.push(run);
    });
    const endRPr = (0, xml_1.kid)(p, 'a:endParaRPr');
    const endSize = (endRPr && getters.size(endRPr)) ?? lookupR(undefined, runEntries, getters.size) ?? 18;
    const firstSize = runs[0]?.resolved.size ?? endSize;
    // 箇条書き
    const buType = lookupP(pPr, entries, (e) => (0, xml_1.kids)(e).find((k) => (0, xml_1.isAny)(k, ['a:buNone', 'a:buChar', 'a:buAutoNum', 'a:buBlip'])));
    let bullet;
    if (buType && !(0, xml_1.is)(buType, 'a:buNone')) {
        bullet = {};
        if ((0, xml_1.is)(buType, 'a:buChar'))
            bullet.char = (0, xml_1.attr)(buType, 'char') ?? '•';
        else if ((0, xml_1.is)(buType, 'a:buAutoNum'))
            bullet.autoNumber = (0, xml_1.attr)(buType, 'type') ?? 'arabicPeriod';
        else
            bullet.char = '•';
        const buClr = lookupP(pPr, entries, (e) => (0, xml_1.kid)(e, 'a:buClr'));
        const bc = buClr ? (0, theme_1.colorOfParent)(buClr, ctx.color) : undefined;
        if (bc)
            bullet.color = (0, theme_1.cssColor)(bc);
        const buFont = lookupP(pPr, entries, (e) => (0, xml_1.attr)((0, xml_1.kid)(e, 'a:buFont'), 'typeface'));
        if (buFont)
            bullet.font = (0, theme_1.resolveThemeFont)(ctx.theme, buFont, 'latin');
    }
    const para = {
        runs,
        align,
        level,
        indent: lookupP(pPr, entries, (e) => (0, xml_1.numAttr)(e, 'indent')) ?? 0,
        marginLeft: lookupP(pPr, entries, (e) => (0, xml_1.numAttr)(e, 'marL')) ?? 0,
        endSize,
        source: pi,
        // ランの a:rPr を除いたときの書式（「書式をクリア」で Webview が使う）
        baseStyle: resolveRun(ctx, undefined, runEntries, false),
    };
    if (bullet)
        para.bullet = bullet;
    const lnSpc = lookupP(pPr, entries, (e) => (0, xml_1.kid)(e, 'a:lnSpc'));
    if (lnSpc) {
        const pct = (0, xml_1.kid)(lnSpc, 'a:spcPct');
        const pts = (0, xml_1.kid)(lnSpc, 'a:spcPts');
        if (pct)
            para.lineSpacing = ((0, xml_1.numAttr)(pct, 'val') ?? 100000) / 100000;
        else if (pts)
            para.lineSpacingPt = ((0, xml_1.numAttr)(pts, 'val') ?? 0) / 100;
    }
    const before = spacingPt(lookupP(pPr, entries, (e) => (0, xml_1.kid)(e, 'a:spcBef')), firstSize);
    if (before !== undefined)
        para.spaceBefore = before;
    const after = spacingPt(lookupP(pPr, entries, (e) => (0, xml_1.kid)(e, 'a:spcAft')), firstSize);
    if (after !== undefined)
        para.spaceAfter = after;
    return para;
}
// ---------------------------------------------------------------------------
// 表（設計書 7 章「表」）
// ---------------------------------------------------------------------------
const DEFAULT_CELL_INSETS = { l: 91440, r: 91440, t: 45720, b: 45720 };
function parseTable(ctx, tbl, part, owner) {
    const tblPr = (0, xml_1.kid)(tbl, 'a:tblPr');
    const columns = (0, xml_1.kids)((0, xml_1.kid)(tbl, 'a:tblGrid'))
        .filter((k) => (0, xml_1.is)(k, 'a:gridCol'))
        .map((g) => (0, xml_1.numAttr)(g, 'w') ?? 0);
    const trs = (0, xml_1.kids)(tbl).filter((k) => (0, xml_1.is)(k, 'a:tr'));
    const style = (0, tableStyle_1.findTableStyle)(ctx.r, tblPr);
    const flags = (0, tableStyle_1.tableFlags)(tblPr);
    const nRows = trs.length;
    const nCols = Math.max(columns.length, ...trs.map((tr) => (0, xml_1.kidsNamed)(tr, 'a:tc').length));
    const rows = trs.map((tr, ri) => ({
        height: (0, xml_1.numAttr)(tr, 'h') ?? 0,
        cells: (0, xml_1.kidsNamed)(tr, 'a:tc').map((tc, ci) => parseCell(ctx, tc, part, owner, style, flags, ri, ci, nRows, nCols)),
    }));
    return { columns, rows };
}
/** セルのテキストの継承チェーン：セルの lstStyle → 表スタイルの tcTxStyle → defaultTextStyle（マスター上なら otherStyle も）。 */
function cellChain(ctx, tc, t, owner) {
    const items = [];
    const ls = (0, xml_1.kid)((0, xml_1.kid)(tc, 'a:txBody'), 'a:lstStyle');
    if (ls)
        items.push({ kind: 'lst', el: ls });
    if (t.bold !== undefined || t.italic !== undefined || t.latin || t.ea || t.color) {
        items.push({ kind: 'fontRef', latin: t.latin, ea: t.ea, color: t.color, bold: t.bold, italic: t.italic });
    }
    if (owner === 'master') {
        const os = (0, xml_1.kid)((0, xml_1.kid)(ctx.masterRoot, 'p:txStyles'), 'p:otherStyle');
        if (os)
            items.push({ kind: 'lst', el: os });
    }
    if (ctx.defaultTextStyle)
        items.push({ kind: 'lst', el: ctx.defaultTextStyle });
    return { items, bodyPrs: [] };
}
/** a:tbl の (row, col) のセルの継承チェーン（編集で段落の書式を書くときに使う。スライド上の表）。 */
function tableCellChain(ctx, tbl, row, col) {
    const tblPr = (0, xml_1.kid)(tbl, 'a:tblPr');
    const trs = (0, xml_1.kidsNamed)(tbl, 'a:tr');
    const tc = (0, xml_1.kidsNamed)(trs[row], 'a:tc')[col];
    if (!tc)
        return { items: [], bodyPrs: [] };
    const nRows = trs.length;
    const nCols = Math.max((0, xml_1.kidsNamed)((0, xml_1.kid)(tbl, 'a:tblGrid'), 'a:gridCol').length, ...trs.map((tr) => (0, xml_1.kidsNamed)(tr, 'a:tc').length));
    const gridSpan = Math.max(1, (0, xml_1.numAttr)(tc, 'gridSpan') ?? 1);
    const rowSpan = Math.max(1, (0, xml_1.numAttr)(tc, 'rowSpan') ?? 1);
    const pos = {
        row,
        col,
        rowEnd: Math.min(nRows - 1, row + rowSpan - 1),
        colEnd: Math.min(nCols - 1, col + gridSpan - 1),
        nRows,
        nCols,
    };
    const st = (0, tableStyle_1.resolveCellStyle)((0, tableStyle_1.findTableStyle)(ctx.r, tblPr), (0, tableStyle_1.tableFlags)(tblPr), pos, ctx.color, ctx.theme);
    return cellChain(ctx, tc, st.text, 'slide');
}
function parseCell(ctx, tc, part, owner, style, flags, row, col, nRows, nCols) {
    const gridSpan = Math.max(1, (0, xml_1.numAttr)(tc, 'gridSpan') ?? 1);
    const rowSpan = Math.max(1, (0, xml_1.numAttr)(tc, 'rowSpan') ?? 1);
    const hidden = ((0, xml_1.boolAttr)(tc, 'hMerge') ?? false) || ((0, xml_1.boolAttr)(tc, 'vMerge') ?? false);
    const pos = {
        row,
        col,
        rowEnd: Math.min(nRows - 1, row + rowSpan - 1),
        colEnd: Math.min(nCols - 1, col + gridSpan - 1),
        nRows,
        nCols,
    };
    const st = (0, tableStyle_1.resolveCellStyle)(style, flags, pos, ctx.color, ctx.theme);
    const tcPr = (0, xml_1.kid)(tc, 'a:tcPr');
    const fill = resolveFill(ctx, fillElOf(tcPr), part, ctx.color) ??
        (st.fill ? resolveFill(ctx, st.fill.el, part, st.fill.color) : undefined) ??
        { type: 'none' };
    const border = (name, fromStyle) => {
        const own = (0, xml_1.kid)(tcPr, name);
        const lns = [];
        if (own)
            lns.push({ ln: own, color: ctx.color });
        if (fromStyle)
            lns.push(fromStyle);
        return lineOf(lns);
    };
    const borders = {
        l: border('a:lnL', st.borders.l),
        t: border('a:lnT', st.borders.t),
        r: border('a:lnR', st.borders.r),
        b: border('a:lnB', st.borders.b),
    };
    const txBody = (0, xml_1.kid)(tc, 'a:txBody');
    const chain = cellChain(ctx, tc, st.text, owner);
    const paragraphs = (0, xml_1.kids)(txBody)
        .filter((k) => (0, xml_1.is)(k, 'a:p'))
        .map((p, pi) => parseParagraph(ctx, chain, p, pi));
    const anchorRaw = (0, xml_1.attr)(tcPr, 'anchor') ?? 't';
    const vertRaw = (0, xml_1.attr)(tcPr, 'vert') ?? 'horz';
    const text = {
        paragraphs,
        anchor: anchorRaw === 'ctr' ? 'ctr' : anchorRaw === 'b' ? 'b' : 't',
        insets: {
            l: (0, xml_1.numAttr)(tcPr, 'marL') ?? DEFAULT_CELL_INSETS.l,
            t: (0, xml_1.numAttr)(tcPr, 'marT') ?? DEFAULT_CELL_INSETS.t,
            r: (0, xml_1.numAttr)(tcPr, 'marR') ?? DEFAULT_CELL_INSETS.r,
            b: (0, xml_1.numAttr)(tcPr, 'marB') ?? DEFAULT_CELL_INSETS.b,
        },
        wrap: true,
        fontScale: 1,
        lineSpaceReduction: 0,
        vertical: ['horz', 'vert', 'vert270', 'eaVert'].includes(vertRaw) ? vertRaw : 'other',
    };
    return { hidden, gridSpan, rowSpan, fill, borders, text };
}
// ---------------------------------------------------------------------------
// ノート
// ---------------------------------------------------------------------------
function notesPathOf(r, slidePath) {
    const rel = (0, package_1.relsOfType)(r, slidePath, package_1.REL.notesSlide)[0];
    return rel && r.has(rel.target) ? rel.target : undefined;
}
function notesBodyShape(notesRoot) {
    for (const k of (0, xml_1.kids)((0, elements_1.spTreeOf)(notesRoot))) {
        if (!(0, xml_1.is)(k, 'p:sp'))
            continue;
        const ph = (0, inherit_1.phOf)(k);
        if (ph && ph.type === 'body')
            return k;
    }
    return undefined;
}
function readNotes(r, slidePath) {
    const np = notesPathOf(r, slidePath);
    if (!np)
        return '';
    const body = notesBodyShape((0, package_1.rootEl)(r.dom(np)));
    const tx = (0, xml_1.kid)(body, 'p:txBody');
    if (!tx)
        return '';
    return (0, xml_1.kids)(tx)
        .filter((k) => (0, xml_1.is)(k, 'a:p'))
        .map((p) => runChildren(p)
        .map((r2) => ((0, xml_1.is)(r2, 'a:br') ? '\n' : (0, xml_1.textOf)((0, xml_1.kid)(r2, 'a:t'))))
        .join(''))
        .join('\n');
}
// ---------------------------------------------------------------------------
// 編集で使う：要素の解決済みの位置
// ---------------------------------------------------------------------------
function resolvedOwnXfrm(ctx, el) {
    return (0, elements_1.ownXfrm)(el) ?? (0, inherit_1.resolveXfrm)(el, (0, inherit_1.phSources)(ctx, el, 'slide'));
}
