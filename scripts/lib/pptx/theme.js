"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.DEFAULT_CLRMAP = void 0;
exports.parseTheme = parseTheme;
exports.parseClrMap = parseClrMap;
exports.resolveThemeFont = resolveThemeFont;
exports.isColorEl = isColorEl;
exports.colorOfParent = colorOfParent;
exports.resolveColor = resolveColor;
exports.fromHex = fromHex;
exports.toHex = toHex;
exports.cssColor = cssColor;
// テーマ色、カラーマップ、色変換、フォント（設計書 7 章「色」）。
const xml_1 = require("./xml.js");
const own_1 = require("../shared/own.js");
exports.DEFAULT_CLRMAP = {
    bg1: 'lt1',
    tx1: 'dk1',
    bg2: 'lt2',
    tx2: 'dk2',
    accent1: 'accent1',
    accent2: 'accent2',
    accent3: 'accent3',
    accent4: 'accent4',
    accent5: 'accent5',
    accent6: 'accent6',
    hlink: 'hlink',
    folHlink: 'folHlink',
};
const DEFAULT_COLORS = {
    dk1: '000000',
    lt1: 'FFFFFF',
    dk2: '44546A',
    lt2: 'E7E6E6',
    accent1: '4472C4',
    accent2: 'ED7D31',
    accent3: 'A5A5A5',
    accent4: 'FFC000',
    accent5: '5B9BD5',
    accent6: '70AD47',
    hlink: '0563C1',
    folHlink: '954F72',
};
function parseTheme(root) {
    const colors = { ...DEFAULT_COLORS };
    const scheme = (0, xml_1.path)(root, 'a:themeElements/a:clrScheme');
    for (const c of (0, xml_1.kids)(scheme)) {
        const choice = (0, xml_1.kids)(c)[0];
        if (!choice)
            continue;
        const rc = baseColor(choice, { colors: DEFAULT_COLORS, clrMap: exports.DEFAULT_CLRMAP });
        if (rc && c.localName && c.localName !== '__proto__')
            colors[c.localName] = rc.hex;
    }
    const fontScheme = (0, xml_1.path)(root, 'a:themeElements/a:fontScheme');
    const fonts = (name) => {
        const f = (0, xml_1.kid)(fontScheme, name);
        const jpan = (0, xml_1.kids)(f).find((k) => k.localName === 'font' && (0, xml_1.attr)(k, 'script') === 'Jpan');
        return {
            latin: (0, xml_1.attr)((0, xml_1.kid)(f, 'a:latin'), 'typeface') ?? 'Calibri',
            ea: (0, xml_1.attr)((0, xml_1.kid)(f, 'a:ea'), 'typeface') ?? '',
            jpan: (0, xml_1.attr)(jpan, 'typeface') ?? '',
        };
    };
    const fmt = (0, xml_1.path)(root, 'a:themeElements/a:fmtScheme');
    return {
        colors,
        major: fonts('a:majorFont'),
        minor: fonts('a:minorFont'),
        fillStyles: (0, xml_1.kids)((0, xml_1.kid)(fmt, 'a:fillStyleLst')),
        lineStyles: (0, xml_1.kids)((0, xml_1.kid)(fmt, 'a:lnStyleLst')),
        bgFillStyles: (0, xml_1.kids)((0, xml_1.kid)(fmt, 'a:bgFillStyleLst')),
    };
}
function parseClrMap(el) {
    if (!el)
        return undefined;
    const m = {};
    for (const k of Object.keys(exports.DEFAULT_CLRMAP)) {
        const v = (0, xml_1.attr)(el, k);
        if (v)
            m[k] = v;
    }
    return { ...exports.DEFAULT_CLRMAP, ...m };
}
/** テーマのフォント指定（+mj-lt など）を実際の名前にする。 */
function resolveThemeFont(theme, typeface, kind) {
    if (typeface === undefined)
        return undefined;
    const t = typeface;
    const pick = (f, k) => {
        if (k === 'latin')
            return f.latin;
        return f.ea !== '' ? f.ea : f.jpan;
    };
    switch (t) {
        case '+mj-lt':
            return theme.major.latin;
        case '+mn-lt':
            return theme.minor.latin;
        case '+mj-ea':
            return pick(theme.major, 'ea');
        case '+mn-ea':
            return pick(theme.minor, 'ea');
        case '+mj-cs':
            return pick(theme.major, kind);
        case '+mn-cs':
            return pick(theme.minor, kind);
        default:
            return t;
    }
}
const PRESET_COLORS = {
    black: '000000',
    white: 'FFFFFF',
    red: 'FF0000',
    green: '008000',
    blue: '0000FF',
    yellow: 'FFFF00',
    cyan: '00FFFF',
    magenta: 'FF00FF',
    gray: '808080',
    grey: '808080',
    darkGray: 'A9A9A9',
    lightGray: 'D3D3D3',
    orange: 'FFA500',
    purple: '800080',
    navy: '000080',
    maroon: '800000',
    olive: '808000',
    teal: '008080',
    silver: 'C0C0C0',
    lime: '00FF00',
    pink: 'FFC0CB',
    brown: 'A52A2A',
    gold: 'FFD700',
    darkBlue: '00008B',
    darkRed: '8B0000',
    darkGreen: '006400',
    lightBlue: 'ADD8E6',
};
const COLOR_ELEMENTS = ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'];
function isColorEl(el) {
    return !!el && COLOR_ELEMENTS.includes(el.localName ?? '');
}
/** 色を含む親（a:solidFill など）から最初の色要素を探して解決する。 */
function colorOfParent(parent, ctx) {
    const c = (0, xml_1.kids)(parent).find((k) => isColorEl(k));
    return c ? resolveColor(c, ctx) : undefined;
}
/** 色要素（a:srgbClr など）を解決する。例外は投げない。 */
function resolveColor(el, ctx) {
    try {
        const base = baseColor(el, ctx);
        if (!base)
            return undefined;
        return applyTransforms(base, el);
    }
    catch {
        return undefined;
    }
}
function baseColor(el, ctx) {
    switch (el.localName) {
        case 'srgbClr':
            return hexOk((0, xml_1.attr)(el, 'val'));
        case 'sysClr':
            return hexOk((0, xml_1.attr)(el, 'lastClr')) ?? ((0, xml_1.attr)(el, 'val') === 'window' ? { hex: 'FFFFFF', alpha: 1 } : { hex: '000000', alpha: 1 });
        case 'prstClr': {
            const v = (0, own_1.own)(PRESET_COLORS, (0, xml_1.attr)(el, 'val'));
            return v ? { hex: v, alpha: 1 } : { hex: '000000', alpha: 1 };
        }
        case 'scrgbClr': {
            const r = linToSrgb(((0, xml_1.numAttr)(el, 'r') ?? 0) / 100000);
            const g = linToSrgb(((0, xml_1.numAttr)(el, 'g') ?? 0) / 100000);
            const b = linToSrgb(((0, xml_1.numAttr)(el, 'b') ?? 0) / 100000);
            return { hex: toHex(r * 255, g * 255, b * 255), alpha: 1 };
        }
        case 'hslClr': {
            const h = ((0, xml_1.numAttr)(el, 'hue') ?? 0) / 60000;
            const s = ((0, xml_1.numAttr)(el, 'sat') ?? 0) / 100000;
            const l = ((0, xml_1.numAttr)(el, 'lum') ?? 0) / 100000;
            const [r, g, b] = hslToRgb(h, s, l);
            return { hex: toHex(r, g, b), alpha: 1 };
        }
        case 'schemeClr': {
            const val = (0, xml_1.attr)(el, 'val') ?? '';
            if (val === 'phClr')
                return ctx.phClr ? { ...ctx.phClr } : { hex: '000000', alpha: 1 };
            const mapped = (0, own_1.own)(ctx.clrMap, val) ?? val;
            const hex = (0, own_1.own)(ctx.colors, mapped);
            return hex ? { hex, alpha: 1 } : { hex: '000000', alpha: 1 };
        }
        default:
            return undefined;
    }
}
function hexOk(v) {
    if (!v || !/^[0-9a-fA-F]{6}$/.test(v))
        return undefined;
    return { hex: v.toUpperCase(), alpha: 1 };
}
function applyTransforms(c, el) {
    let [r, g, b] = fromHex(c.hex);
    let alpha = c.alpha;
    for (const t of (0, xml_1.kids)(el)) {
        const v = ((0, xml_1.numAttr)(t, 'val') ?? 0) / 100000;
        switch (t.localName) {
            case 'alpha':
                alpha = v;
                break;
            case 'lumMod':
            case 'lumOff': {
                const [h, s, l] = rgbToHsl(r, g, b);
                const nl = t.localName === 'lumMod' ? l * v : l + v;
                [r, g, b] = hslToRgb(h, s, clamp01(nl));
                break;
            }
            case 'satMod': {
                const [h, s, l] = rgbToHsl(r, g, b);
                [r, g, b] = hslToRgb(h, clamp01(s * v), l);
                break;
            }
            case 'shade': {
                // 線形 RGB で黒に近づける
                [r, g, b] = [r, g, b].map((x) => linToSrgb(srgbToLin(x / 255) * v) * 255);
                break;
            }
            case 'tint': {
                [r, g, b] = [r, g, b].map((x) => {
                    const lin = srgbToLin(x / 255);
                    return linToSrgb(lin * v + (1 - v)) * 255;
                });
                break;
            }
            default:
                break; // それ以外の変換は無視
        }
    }
    return { hex: toHex(r, g, b), alpha: clamp01(alpha) };
}
function clamp01(x) {
    return Math.min(1, Math.max(0, x));
}
function srgbToLin(c) {
    return c <= 0.04045 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
}
function linToSrgb(c) {
    const x = clamp01(c);
    return x <= 0.0031308 ? x * 12.92 : 1.055 * Math.pow(x, 1 / 2.4) - 0.055;
}
function fromHex(hex) {
    const n = parseInt(hex.replace('#', ''), 16) || 0;
    return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}
function toHex(r, g, b) {
    const f = (x) => Math.round(Math.min(255, Math.max(0, x)))
        .toString(16)
        .padStart(2, '0');
    return (f(r) + f(g) + f(b)).toUpperCase();
}
function rgbToHsl(r, g, b) {
    const rn = r / 255;
    const gn = g / 255;
    const bn = b / 255;
    const max = Math.max(rn, gn, bn);
    const min = Math.min(rn, gn, bn);
    const l = (max + min) / 2;
    if (max === min)
        return [0, 0, l];
    const d = max - min;
    const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    let h;
    if (max === rn)
        h = (gn - bn) / d + (gn < bn ? 6 : 0);
    else if (max === gn)
        h = (bn - rn) / d + 2;
    else
        h = (rn - gn) / d + 4;
    return [h * 60, s, l];
}
function hslToRgb(h, s, l) {
    if (s === 0)
        return [l * 255, l * 255, l * 255];
    const hue = (((h % 360) + 360) % 360) / 360;
    const q2 = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p2 = 2 * l - q2;
    const f = (t0) => {
        let t = t0;
        if (t < 0)
            t += 1;
        if (t > 1)
            t -= 1;
        if (t < 1 / 6)
            return p2 + (q2 - p2) * 6 * t;
        if (t < 1 / 2)
            return q2;
        if (t < 2 / 3)
            return p2 + (q2 - p2) * (2 / 3 - t) * 6;
        return p2;
    };
    return [f(hue + 1 / 3) * 255, f(hue) * 255, f(hue - 1 / 3) * 255];
}
function cssColor(c) {
    return `#${c.hex}`;
}
