"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BULLET_CHAR = exports.BULLET_HANG = exports.INDENT_STEP = void 0;
exports.hasParagraphProps = hasParagraphProps;
exports.validateParagraphProps = validateParagraphProps;
exports.applyParagraphProps = applyParagraphProps;
const inherit_1 = require("./inherit.js");
const types_1 = require("./types.js");
const xml_1 = require("./xml.js");
/** レベル 1 段のインデント（0.5 インチ）。lstStyle にレベルごとの既定がないときに使う。 */
exports.INDENT_STEP = 457200;
/** 箇条書きの記号の幅（ぶら下げ）。Google スライドの既定（marL=457200、indent=-342900）に合わせる。 */
exports.BULLET_HANG = 342900;
/** Google スライドの箇条書きの記号。 */
exports.BULLET_CHAR = '●';
const MAX_LEVEL = 8;
const BU_TYPES = ['a:buNone', 'a:buAutoNum', 'a:buChar', 'a:buBlip'];
function hasParagraphProps(p) {
    return (p.align !== undefined ||
        p.bullet !== undefined ||
        p.lineSpacing !== undefined ||
        p.level !== undefined ||
        (p.indentDelta !== undefined && p.indentDelta !== 0));
}
const ALIGNS = ['l', 'ctr', 'r', 'just'];
const BULLETS = ['none', 'bullet', 'number'];
/** 値の検査。おかしな値は op ごと拒否する（パーツは変わらない）。 */
function validateParagraphProps(p) {
    if (p.align !== undefined && !ALIGNS.includes(p.align))
        throw new types_1.OpRejectedError('配置の指定が正しくありません');
    if (p.bullet !== undefined && !BULLETS.includes(p.bullet))
        throw new types_1.OpRejectedError('箇条書きの指定が正しくありません');
    if (p.lineSpacing !== undefined && !(Number.isFinite(p.lineSpacing) && p.lineSpacing > 0 && p.lineSpacing <= 132)) {
        throw new types_1.OpRejectedError('行間の指定が正しくありません');
    }
    if (p.level !== undefined && !Number.isFinite(p.level))
        throw new types_1.OpRejectedError('レベルの指定が正しくありません');
    if (p.indentDelta !== undefined && !Number.isFinite(p.indentDelta))
        throw new types_1.OpRejectedError('インデントの指定が正しくありません');
}
function clampLevel(n) {
    return Math.min(MAX_LEVEL, Math.max(0, Math.round(n)));
}
/** レベルの継承元の段落書式（lstStyle の lvlNpPr / defPPr、チェーンの順）。 */
function inherited(chain, level) {
    return (0, inherit_1.levelPPrs)(chain, level).filter((e) => e.fontRef === undefined);
}
function lookup(pPr, entries, name) {
    const own = (0, xml_1.numAttr)(pPr, name);
    if (own !== undefined)
        return own;
    for (const e of entries) {
        const v = (0, xml_1.numAttr)(e, name);
        if (v !== undefined)
            return v;
    }
    return undefined;
}
/**
 * レベルを変える。marL を明示している段落は 1 段ごとに INDENT_STEP ずらす。
 * 明示していなければ lstStyle のレベルごとの既定に任せる（既定が段ごとに変わらないときだけ marL を書く）。
 */
function setLevel(pPr, from, to, chain) {
    (0, xml_1.setAttr)(pPr, 'lvl', to > 0 ? to : undefined);
    const delta = (to - from) * exports.INDENT_STEP;
    const own = (0, xml_1.numAttr)(pPr, 'marL');
    if (own !== undefined) {
        (0, xml_1.setAttr)(pPr, 'marL', Math.max(0, own + delta));
        return;
    }
    const fromM = lookup(undefined, inherited(chain, from), 'marL') ?? 0;
    const toM = lookup(undefined, inherited(chain, to), 'marL') ?? 0;
    if (toM === fromM)
        (0, xml_1.setAttr)(pPr, 'marL', Math.max(0, fromM + delta));
}
function setLineSpacing(pPr, v) {
    const doc = pPr.ownerDocument;
    (0, xml_1.removeKids)(pPr, ['a:lnSpc']);
    const ln = (0, xml_1.create)(doc, 'a:lnSpc');
    ln.appendChild((0, xml_1.create)(doc, 'a:spcPct', { val: Math.round(v * 100000) }));
    (0, xml_1.insertChildInOrder)(pPr, ln);
}
/**
 * 記号を外したときのインデント（ぶら下げを解く）。付けたときの逆にする：
 * ぶら下げが BULLET_HANG で marL が 1 段以上なら marL を 1 段戻す。それ以外は 1 行目の位置（marL + indent）に本文をそろえる。
 */
function unhang(marL, indent) {
    if (indent === -exports.BULLET_HANG && marL >= exports.INDENT_STEP)
        return [marL - exports.INDENT_STEP, 0];
    return [Math.max(0, marL + indent), 0];
}
/** marL / indent を書く。継承元と同じ値になる属性は消して継承に戻す（付けて外したとき元の pPr に戻るように）。 */
function writeIndent(pPr, marL, indent, inhMarL, inhIndent) {
    (0, xml_1.setAttr)(pPr, 'marL', marL === inhMarL ? undefined : marL);
    (0, xml_1.setAttr)(pPr, 'indent', indent === inhIndent ? undefined : indent);
}
/**
 * 箇条書き。bullet は a:buFont（Arial）と a:buChar（●）、number は a:buFont（+mj-lt）と a:buAutoNum（arabicPeriod）、
 * none は a:buNone。記号を付けるとき、ぶら下げ（負の indent）がなければ marL を 1 段増やして indent=-342900 にする。
 * 外すときは unhang でぶら下げを解く。付ける・外すは互いに逆になるようにする（繰り返しても本文の位置がずれない）。
 * 継承元（lstStyle）にぶら下げがあり、外したときの位置にいるなら、付けるときは属性を消して継承のぶら下げに戻す。
 */
function setBullet(pPr, kind, chain) {
    const doc = pPr.ownerDocument;
    const entries = inherited(chain, (0, xml_1.numAttr)(pPr, 'lvl') ?? 0);
    const marL = lookup(pPr, entries, 'marL') ?? 0;
    const indent = lookup(pPr, entries, 'indent') ?? 0;
    const inhMarL = lookup(undefined, entries, 'marL') ?? 0;
    const inhIndent = lookup(undefined, entries, 'indent') ?? 0;
    (0, xml_1.removeKids)(pPr, [...BU_TYPES, 'a:buFontTx', 'a:buFont']);
    if (kind === 'none') {
        (0, xml_1.insertChildInOrder)(pPr, (0, xml_1.create)(doc, 'a:buNone'));
        if (indent < 0)
            writeIndent(pPr, ...unhang(marL, indent), inhMarL, inhIndent);
        return;
    }
    if (kind === 'bullet') {
        (0, xml_1.insertChildInOrder)(pPr, (0, xml_1.create)(doc, 'a:buFont', { typeface: 'Arial' }));
        (0, xml_1.insertChildInOrder)(pPr, (0, xml_1.create)(doc, 'a:buChar', { char: exports.BULLET_CHAR }));
    }
    else {
        (0, xml_1.insertChildInOrder)(pPr, (0, xml_1.create)(doc, 'a:buFont', { typeface: '+mj-lt' }));
        (0, xml_1.insertChildInOrder)(pPr, (0, xml_1.create)(doc, 'a:buAutoNum', { type: 'arabicPeriod' }));
    }
    if (indent >= 0) {
        const [offMarL, offIndent] = inhIndent < 0 ? unhang(inhMarL, inhIndent) : [NaN, NaN];
        if (marL === offMarL && indent === offIndent)
            writeIndent(pPr, inhMarL, inhIndent, inhMarL, inhIndent);
        else
            writeIndent(pPr, marL + exports.INDENT_STEP, -exports.BULLET_HANG, inhMarL, inhIndent);
    }
}
/** a:pPr に段落の書式を書く。chain は段落を持つ図形の継承チェーン（レベルの既定を引くため）。 */
function applyParagraphProps(pPr, props, chain) {
    if (props.align !== undefined)
        (0, xml_1.setAttr)(pPr, 'algn', props.align);
    const cur = clampLevel((0, xml_1.numAttr)(pPr, 'lvl') ?? 0);
    let target = props.level !== undefined ? clampLevel(props.level) : cur;
    if (props.indentDelta)
        target = clampLevel(target + props.indentDelta);
    if (target !== cur)
        setLevel(pPr, cur, target, chain);
    if (props.lineSpacing !== undefined)
        setLineSpacing(pPr, props.lineSpacing);
    if (props.bullet !== undefined)
        setBullet(pPr, props.bullet, chain);
}
