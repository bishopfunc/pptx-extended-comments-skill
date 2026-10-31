"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.esc = esc;
exports.xfrmXml = xfrmXml;
exports.shapeXml = shapeXml;
exports.textBoxXml = textBoxXml;
exports.lineXml = lineXml;
exports.pictureXml = pictureXml;
exports.emptySlideXml = emptySlideXml;
exports.placeholderSpXml = placeholderSpXml;
exports.notesBodyXml = notesBodyXml;
exports.notesSlideXml = notesSlideXml;
exports.notesMasterXml = notesMasterXml;
exports.groupXml = groupXml;
const xml_1 = require("./xml.js");
const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NSDECL = `xmlns:a="${xml_1.NS.a}" xmlns:r="${xml_1.NS.r}" xmlns:p="${xml_1.NS.p}"`;
function esc(s) {
    return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function xfrmXml(x, tag = 'a:xfrm') {
    const attrs = [];
    const rot = (0, xml_1.rotUnits)(x.rot);
    if (rot)
        attrs.push(`rot="${rot}"`);
    if (x.flipH)
        attrs.push('flipH="1"');
    if (x.flipV)
        attrs.push('flipV="1"');
    return (`<${tag}${attrs.length ? ' ' + attrs.join(' ') : ''}><a:off x="${(0, xml_1.emuCoord)(x.x)}" y="${(0, xml_1.emuCoord)(x.y)}"/>` +
        `<a:ext cx="${(0, xml_1.emuSize)(x.cx)}" cy="${(0, xml_1.emuSize)(x.cy)}"/></${tag}>`);
}
function paragraphsXml(text, algn) {
    const pPr = algn ? `<a:pPr algn="${algn}"/>` : '';
    if (!text)
        return `<a:p>${pPr}<a:endParaRPr lang="ja-JP" altLang="en-US" dirty="0"/></a:p>`;
    return text
        .split('\n')
        .map((line) => line
        ? `<a:p>${pPr}<a:r><a:rPr lang="ja-JP" altLang="en-US" dirty="0"/><a:t>${esc(line)}</a:t></a:r></a:p>`
        : `<a:p>${pPr}<a:endParaRPr lang="ja-JP" altLang="en-US" dirty="0"/></a:p>`)
        .join('');
}
const ACCENT_FILL = '<a:solidFill><a:schemeClr val="accent1"/></a:solidFill>';
const ACCENT_LINE = '<a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:solidFill></a:ln>';
function shapeXml(id, name, prst, x, text) {
    return (`<p:sp ${NSDECL}><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr>${xfrmXml(x)}<a:prstGeom prst="${esc(prst)}"><a:avLst/></a:prstGeom>${ACCENT_FILL}${ACCENT_LINE}</p:spPr>` +
        `<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/>${paragraphsXml(text, 'ctr')}</p:txBody></p:sp>`);
}
function textBoxXml(id, name, x, text) {
    return (`<p:sp ${NSDECL}><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
        `<p:spPr>${xfrmXml(x)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>` +
        `<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/>${paragraphsXml(text)}</p:txBody></p:sp>`);
}
function lineXml(id, name, prst, x) {
    return (`<p:cxnSp ${NSDECL}><p:nvCxnSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
        `<p:spPr>${xfrmXml(x)}<a:prstGeom prst="${esc(prst)}"><a:avLst/></a:prstGeom>` +
        `<a:ln w="12700"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`);
}
function pictureXml(id, name, rId, x) {
    return (`<p:pic ${NSDECL}><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
        `<p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
        `<p:spPr>${xfrmXml(x)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
}
function emptySlideXml() {
    return (`${XML_DECL}<p:sld ${NSDECL}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
        `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`);
}
function placeholderSpXml(id, name, phAttrs) {
    return (`<p:sp ${NSDECL}><p:nvSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
        `<p:nvPr><p:ph${phAttrs}/></p:nvPr></p:nvSpPr><p:spPr/>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" altLang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`);
}
function notesBodyXml(id) {
    return placeholderSpXml(id, `Notes Placeholder ${id - 1}`, ' type="body" idx="1"');
}
function notesSlideXml() {
    return (`${XML_DECL}<p:notes ${NSDECL}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>` +
        `<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>` +
        `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>` +
        `<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>` +
        `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="ja-JP" altLang="en-US" dirty="0"/></a:p></p:txBody></p:sp>` +
        `</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`);
}
function notesLevel(i) {
    return (`<a:lvl${i}pPr marL="${(i - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
        `<a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill>` +
        `<a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i}pPr>`);
}
/** ノートマスターの雛形。notesSz（cx, cy）に合わせて配置する。 */
function notesMasterXml(cx, cy) {
    const ph = (id, name, attrs, x, y, w, h, body) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${Math.round(x)}" y="${Math.round(y)}"/><a:ext cx="${Math.round(w)}" cy="${Math.round(h)}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${attrs.includes('sldImg') ? '<a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln>' : ''}</p:spPr>${body}</p:sp>`;
    const txt = (s) => `<p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:r><a:rPr lang="ja-JP"/><a:t>${s}</a:t></a:r></a:p></p:txBody>`;
    const m = cx * 0.1;
    return (`${XML_DECL}<p:notesMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>` +
        `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>` +
        ph(2, 'Header Placeholder 1', 'type="hdr" sz="quarter"', 0, 0, cx / 2, cy * 0.05, txt('')) +
        ph(3, 'Date Placeholder 2', 'type="dt" idx="1"', cx / 2, 0, cx / 2, cy * 0.05, txt('')) +
        ph(4, 'Slide Image Placeholder 3', 'type="sldImg" idx="2"', m, cy * 0.08, cx - 2 * m, cy * 0.37, '') +
        ph(5, 'Notes Placeholder 4', 'type="body" sz="quarter" idx="3"', m, cy * 0.48, cx - 2 * m, cy * 0.4, txt('テキストを入力')) +
        ph(6, 'Footer Placeholder 5', 'type="ftr" sz="quarter" idx="4"', 0, cy * 0.95, cx / 2, cy * 0.05, txt('')) +
        ph(7, 'Slide Number Placeholder 6', 'type="sldNum" sz="quarter" idx="5"', cx / 2, cy * 0.95, cx / 2, cy * 0.05, txt('')) +
        `</p:spTree></p:cSld>` +
        `<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>` +
        `<p:notesStyle>${[1, 2, 3, 4, 5, 6, 7, 8, 9].map(notesLevel).join('')}</p:notesStyle></p:notesMaster>`);
}
/** 新しいグループ（中身は空。子は呼び出し側で移す）。chOff/chExt は off/ext と同じにして、子の座標を変えずに済ませる。 */
function groupXml(id, name, box) {
    const x = (0, xml_1.emuCoord)(box.x);
    const y = (0, xml_1.emuCoord)(box.y);
    const cx = (0, xml_1.emuSize)(box.cx);
    const cy = (0, xml_1.emuSize)(box.cy);
    return (`<p:grpSp ${NSDECL}><p:nvGrpSpPr><p:cNvPr id="${id}" name="${esc(name)}"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
        `<p:grpSpPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/>` +
        `<a:chOff x="${x}" y="${y}"/><a:chExt cx="${cx}" cy="${cy}"/></a:xfrm></p:grpSpPr></p:grpSp>`);
}
