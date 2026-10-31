"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SHAPE_KINDS = void 0;
exports.presentationPath = presentationPath;
exports.readPresentation = readPresentation;
exports.nvOf = nvOf;
exports.cNvPrOf = cNvPrOf;
exports.altContentMain = altContentMain;
exports.phOf = phOf;
exports.normalizePhType = normalizePhType;
exports.txStyleKey = txStyleKey;
exports.findLayoutPh = findLayoutPh;
exports.findMasterPh = findMasterPh;
exports.themeFor = themeFor;
exports.layoutPathOf = layoutPathOf;
exports.masterPathOf = masterPathOf;
exports.buildSlideCtx = buildSlideCtx;
exports.phSources = phSources;
exports.readXfrmEl = readXfrmEl;
exports.xfrmElOf = xfrmElOf;
exports.resolveXfrm = resolveXfrm;
exports.textChain = textChain;
exports.levelPPrs = levelPPrs;
exports.findBackground = findBackground;
exports.showMasterShapes = showMasterShapes;
const package_1 = require("./package.js");
const theme_1 = require("./theme.js");
const xml_1 = require("./xml.js");
exports.SHAPE_KINDS = ['p:sp', 'p:pic', 'p:cxnSp', 'p:grpSp', 'p:graphicFrame', 'p:contentPart', 'mc:AlternateContent'];
function presentationPath(r) {
    const rel = (0, package_1.readRels)(r, '').find((x) => x.type === package_1.REL.officeDocument);
    return rel?.target ?? 'ppt/presentation.xml';
}
function readPresentation(r) {
    const p = presentationPath(r);
    const root = (0, package_1.rootEl)(r.dom(p));
    const rels = (0, package_1.readRels)(r, p);
    const sz = (0, xml_1.kid)(root, 'p:sldSz');
    const slides = (0, xml_1.kids)((0, xml_1.kid)(root, 'p:sldIdLst'))
        .filter((e) => (0, xml_1.is)(e, 'p:sldId'))
        .map((e) => {
        const rId = (0, xml_1.attr)(e, 'r:id') ?? '';
        return { id: (0, xml_1.attr)(e, 'id') ?? '', rId, path: rels.find((x) => x.id === rId)?.target ?? '' };
    })
        .filter((s) => s.path && r.has(s.path))
        // 同じスライドを何度も指していれば、最初の 1 つだけ（細工した pptx でスライドを水増しさせない。ADR-0019）
        .filter((s, i, all) => all.findIndex((x) => x.path === s.path) === i);
    return {
        path: p,
        root,
        slideSize: { cx: (0, xml_1.numAttr)(sz, 'cx') ?? 9144000, cy: (0, xml_1.numAttr)(sz, 'cy') ?? 6858000 },
        defaultTextStyle: (0, xml_1.kid)(root, 'p:defaultTextStyle'),
        slides,
    };
}
/** 図形要素の nv*Pr 要素。 */
function nvOf(el) {
    const target = (0, xml_1.is)(el, 'mc:AlternateContent') ? altContentMain(el) : el;
    if (!target)
        return undefined;
    return (0, xml_1.kids)(target).find((k) => (k.localName ?? '').startsWith('nv'));
}
function cNvPrOf(el) {
    return (0, xml_1.kid)(nvOf(el), 'p:cNvPr');
}
/** mc:AlternateContent の中の図形要素（Choice、なければ Fallback）。 */
function altContentMain(el) {
    const choice = (0, xml_1.kid)(el, 'mc:Choice');
    const fromChoice = (0, xml_1.kids)(choice)[0];
    if (fromChoice)
        return fromChoice;
    return (0, xml_1.kids)((0, xml_1.kid)(el, 'mc:Fallback'))[0];
}
function phOf(el) {
    const ph = (0, xml_1.kid)((0, xml_1.kid)(nvOf(el), 'p:nvPr'), 'p:ph');
    if (!ph)
        return undefined;
    return { type: (0, xml_1.attr)(ph, 'type') ?? 'obj', idx: (0, xml_1.numAttr)(ph, 'idx') ?? 0 };
}
function normalizePhType(t) {
    if (t === 'ctrTitle')
        return 'title';
    if (['subTitle', 'obj', 'pic', 'tbl', 'chart', 'dgm', 'media', 'clipArt'].includes(t))
        return 'body';
    return t;
}
function txStyleKey(t) {
    const n = normalizePhType(t);
    if (n === 'title')
        return 'p:titleStyle';
    if (n === 'body')
        return 'p:bodyStyle';
    return 'p:otherStyle';
}
function spTreeOf(root) {
    return (0, xml_1.path)(root, 'p:cSld/p:spTree');
}
function placeholdersIn(root) {
    const out = [];
    for (const k of (0, xml_1.kids)(spTreeOf(root))) {
        if (!(0, xml_1.isAny)(k, exports.SHAPE_KINDS))
            continue;
        const ph = phOf(k);
        if (ph)
            out.push({ el: k, ph });
    }
    return out;
}
/** スライド → レイアウト：idx、なければ正規化した type。 */
function findLayoutPh(layoutRoot, ph) {
    const all = placeholdersIn(layoutRoot);
    const byIdx = all.find((x) => x.ph.idx === ph.idx);
    if (byIdx)
        return byIdx.el;
    const t = normalizePhType(ph.type);
    return all.find((x) => normalizePhType(x.ph.type) === t)?.el;
}
/** レイアウト → マスター：正規化した type だけ。 */
function findMasterPh(masterRoot, type) {
    const t = normalizePhType(type);
    return placeholdersIn(masterRoot).find((x) => normalizePhType(x.ph.type) === t)?.el;
}
const themeCache = new WeakMap();
function themeFor(r, themePath) {
    if (!themePath || !r.has(themePath))
        return (0, theme_1.parseTheme)(undefined);
    const doc = r.dom(themePath);
    let t = themeCache.get(doc);
    if (!t) {
        t = (0, theme_1.parseTheme)((0, package_1.rootEl)(doc));
        themeCache.set(doc, t);
    }
    return t;
}
function layoutPathOf(r, slidePath) {
    return (0, package_1.relsOfType)(r, slidePath, package_1.REL.slideLayout)[0]?.target;
}
function masterPathOf(r, layoutPath) {
    if (!layoutPath)
        return undefined;
    return (0, package_1.relsOfType)(r, layoutPath, package_1.REL.slideMaster)[0]?.target;
}
function buildSlideCtx(r, slidePath, defaultTextStyle) {
    const slideRoot = (0, package_1.rootEl)(r.dom(slidePath));
    const layoutPath = layoutPathOf(r, slidePath);
    const layoutRoot = layoutPath && r.has(layoutPath) ? (0, package_1.rootEl)(r.dom(layoutPath)) : undefined;
    const masterPath = masterPathOf(r, layoutPath);
    const masterRoot = masterPath && r.has(masterPath) ? (0, package_1.rootEl)(r.dom(masterPath)) : undefined;
    const themePath = masterPath ? (0, package_1.relsOfType)(r, masterPath, package_1.REL.theme)[0]?.target : undefined;
    const theme = themeFor(r, themePath);
    const clrMap = (0, theme_1.parseClrMap)((0, xml_1.path)(slideRoot, 'p:clrMapOvr/a:overrideClrMapping')) ??
        (0, theme_1.parseClrMap)((0, xml_1.path)(layoutRoot, 'p:clrMapOvr/a:overrideClrMapping')) ??
        (0, theme_1.parseClrMap)((0, xml_1.kid)(masterRoot, 'p:clrMap')) ??
        theme_1.DEFAULT_CLRMAP;
    return {
        r,
        slidePath,
        slideRoot,
        layoutPath,
        layoutRoot,
        masterPath,
        masterRoot,
        theme,
        clrMap,
        color: { colors: theme.colors, clrMap },
        defaultTextStyle,
    };
}
/** owner は図形が置かれているパーツの種類。 */
function phSources(ctx, el, owner) {
    const ph = phOf(el);
    if (!ph || owner === 'master')
        return {};
    if (owner === 'layout')
        return { masterPh: findMasterPh(ctx.masterRoot, ph.type) };
    const layoutPh = findLayoutPh(ctx.layoutRoot, ph);
    const masterPh = findMasterPh(ctx.masterRoot, layoutPh ? (phOf(layoutPh)?.type ?? ph.type) : ph.type);
    return { layoutPh, masterPh };
}
// ---------------------------------------------------------------------------
// 位置と大きさ
// ---------------------------------------------------------------------------
function readXfrmEl(x) {
    if (!x)
        return undefined;
    const off = (0, xml_1.kid)(x, 'a:off');
    const ext = (0, xml_1.kid)(x, 'a:ext');
    if (!off && !ext)
        return undefined;
    // 細工した値でも、座標は EMU の範囲、大きさは 0 以上、回転は 0〜360 に収める（ADR-0019）
    return {
        x: (0, xml_1.clampCoord)((0, xml_1.numAttr)(off, 'x') ?? 0),
        y: (0, xml_1.clampCoord)((0, xml_1.numAttr)(off, 'y') ?? 0),
        cx: (0, xml_1.clampSize)((0, xml_1.numAttr)(ext, 'cx') ?? 0),
        cy: (0, xml_1.clampSize)((0, xml_1.numAttr)(ext, 'cy') ?? 0),
        rot: (0, xml_1.normDeg)(((0, xml_1.numAttr)(x, 'rot') ?? 0) / 60000),
        flipH: (0, xml_1.boolAttr)(x, 'flipH') ?? false,
        flipV: (0, xml_1.boolAttr)(x, 'flipV') ?? false,
    };
}
/** 図形要素自身の xfrm 要素（graphicFrame は p:xfrm、グループは grpSpPr/a:xfrm）。 */
function xfrmElOf(el) {
    const target = (0, xml_1.is)(el, 'mc:AlternateContent') ? altContentMain(el) : el;
    if (!target)
        return undefined;
    if ((0, xml_1.is)(target, 'p:graphicFrame'))
        return (0, xml_1.kid)(target, 'p:xfrm');
    if ((0, xml_1.is)(target, 'p:grpSp'))
        return (0, xml_1.path)(target, 'p:grpSpPr/a:xfrm');
    return (0, xml_1.path)(target, 'p:spPr/a:xfrm');
}
function resolveXfrm(el, src) {
    return (readXfrmEl(xfrmElOf(el)) ??
        (src.layoutPh ? readXfrmEl(xfrmElOf(src.layoutPh)) : undefined) ??
        (src.masterPh ? readXfrmEl(xfrmElOf(src.masterPh)) : undefined) ?? {
        x: 0,
        y: 0,
        cx: 0,
        cy: 0,
        rot: 0,
        flipH: false,
        flipV: false,
    });
}
function textChain(ctx, el, src, owner) {
    const items = [];
    const bodyPrs = [];
    const txOf = (e) => (0, xml_1.kid)(e, 'p:txBody');
    const add = (tx) => {
        const ls = (0, xml_1.kid)(tx, 'a:lstStyle');
        if (ls)
            items.push({ kind: 'lst', el: ls });
        const bp = (0, xml_1.kid)(tx, 'a:bodyPr');
        if (bp)
            bodyPrs.push(bp);
    };
    add(txOf(el));
    add(txOf(src.layoutPh));
    add(txOf(src.masterPh));
    const ph = phOf(el);
    const fontRef = (0, xml_1.path)(el, 'p:style/a:fontRef');
    if (fontRef) {
        const idx = (0, xml_1.attr)(fontRef, 'idx');
        const color = (0, theme_1.colorOfParent)(fontRef, ctx.color);
        items.push({
            kind: 'fontRef',
            latin: idx === 'major' ? '+mj-lt' : idx === 'minor' ? '+mn-lt' : undefined,
            ea: idx === 'major' ? '+mj-ea' : idx === 'minor' ? '+mn-ea' : undefined,
            color,
        });
    }
    if (ph) {
        const ts = (0, xml_1.kid)((0, xml_1.kid)(ctx.masterRoot, 'p:txStyles'), txStyleKey(ph.type));
        if (ts)
            items.push({ kind: 'lst', el: ts });
    }
    else if (owner === 'slide' || owner === 'layout') {
        if (ctx.defaultTextStyle)
            items.push({ kind: 'lst', el: ctx.defaultTextStyle });
    }
    else {
        // マスター上のプレースホルダー以外の図形は otherStyle、それから defaultTextStyle
        const ts = (0, xml_1.kid)((0, xml_1.kid)(ctx.masterRoot, 'p:txStyles'), 'p:otherStyle');
        if (ts)
            items.push({ kind: 'lst', el: ts });
        if (ctx.defaultTextStyle)
            items.push({ kind: 'lst', el: ctx.defaultTextStyle });
    }
    return { items, bodyPrs };
}
/** レベル（0 始まり）の段落書式要素を、チェーンの順に並べる（lvlNpPr、なければ defPPr）。 */
function levelPPrs(chain, level) {
    const out = [];
    const name = `a:lvl${Math.min(9, Math.max(1, level + 1))}pPr`;
    for (const it of chain.items) {
        if (it.kind === 'fontRef') {
            out.push({ fontRef: it });
            continue;
        }
        const lvl = (0, xml_1.kid)(it.el, name);
        if (lvl)
            out.push(lvl);
        const def = (0, xml_1.kid)(it.el, 'a:defPPr');
        if (def)
            out.push(def);
    }
    return out;
}
/** 背景の解決結果：要素（bgPr か bgRef）と、それが置かれているパーツ。 */
function findBackground(ctx) {
    const cands = [
        [ctx.slideRoot, ctx.slidePath],
        [ctx.layoutRoot, ctx.layoutPath],
        [ctx.masterRoot, ctx.masterPath],
    ];
    for (const [root, part] of cands) {
        const bg = (0, xml_1.path)(root, 'p:cSld/p:bg');
        if (bg && part) {
            const inner = (0, xml_1.kid)(bg, 'p:bgPr') ?? (0, xml_1.kid)(bg, 'p:bgRef');
            if (inner)
                return { bg: inner, part };
        }
    }
    return undefined;
}
/** showMasterSp の判定（設計書 9 章「スライドの描画順」）。 */
function showMasterShapes(ctx) {
    const slideShow = (0, xml_1.attr)(ctx.slideRoot, 'showMasterSp') !== '0' && (0, xml_1.attr)(ctx.slideRoot, 'showMasterSp') !== 'false';
    const layoutShow = (0, xml_1.attr)(ctx.layoutRoot, 'showMasterSp') !== '0' && (0, xml_1.attr)(ctx.layoutRoot, 'showMasterSp') !== 'false';
    return { master: slideShow && layoutShow, layout: slideShow };
}
