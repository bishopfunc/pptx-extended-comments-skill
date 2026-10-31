"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.checkIntegrity = checkIntegrity;
// パッケージの整合性検査（設計書 14 章）。テストと保存前の自己検査で使う。
const inherit_1 = require("./inherit.js");
const package_1 = require("./package.js");
const comments_1 = require("./comments.js");
const xml_1 = require("./xml.js");
const ORDER_CHECKED = new Set((0, xml_1.orderedParents)());
function checkIntegrity(r) {
    const issues = [];
    const add = (part, message) => issues.push({ part, message });
    // XML として読めるか
    const xmlParts = r.paths().filter((p) => /\.(xml|rels)$/i.test(p));
    const broken = new Set();
    for (const p of xmlParts) {
        try {
            r.dom(p);
        }
        catch (e) {
            broken.add(p);
            add(p, `XML として読めない: ${e.message}`);
        }
    }
    if (broken.has(package_1.CT_PATH))
        return issues;
    // すべての関係の参照先が存在する
    for (const rp of r.paths().filter((p) => p.endsWith('.rels'))) {
        if (broken.has(rp))
            continue;
        const src = (0, package_1.sourceOfRels)(rp);
        if (src !== '' && !r.has(src))
            add(rp, `rels の元のパーツ ${src} が存在しない`);
        const ids = new Set();
        for (const rel of (0, package_1.readRels)(r, src)) {
            if (ids.has(rel.id))
                add(rp, `関係 ID ${rel.id} が重複している`);
            ids.add(rel.id);
            if (rel.external)
                continue;
            if (!r.has(rel.target))
                add(rp, `関係 ${rel.id} の参照先 ${rel.target} が存在しない`);
        }
    }
    // [Content_Types].xml がすべてのパーツを覆う
    for (const p of r.paths()) {
        if (p === package_1.CT_PATH)
            continue;
        if (!(0, package_1.contentTypeOf)(r, p))
            add(package_1.CT_PATH, `${p} の Content Type がない`);
    }
    const ctRoot = (0, package_1.rootEl)(r.dom(package_1.CT_PATH));
    for (const o of (0, xml_1.descendants)(ctRoot).filter((e) => e.localName === 'Override')) {
        const name = ((0, xml_1.attr)(o, 'PartName') ?? '').replace(/^\//, '');
        if (!r.has(name))
            add(package_1.CT_PATH, `Override の ${name} が存在しない`);
    }
    let pres;
    try {
        pres = (0, inherit_1.readPresentation)(r);
    }
    catch (e) {
        add('ppt/presentation.xml', `presentation.xml を読めない: ${e.message}`);
        return issues;
    }
    const presRels = (0, package_1.readRels)(r, pres.path);
    const slideIds = new Set();
    for (const s of (0, xml_1.kidsNamed)((0, xml_1.kid)(pres.root, 'p:sldIdLst'), 'p:sldId')) {
        const id = (0, xml_1.attr)(s, 'id') ?? '';
        if (slideIds.has(id))
            add(pres.path, `スライド ID ${id} が重複している`);
        slideIds.add(id);
        const n = Number(id);
        if (!(n >= 256 && n <= 2147483647))
            add(pres.path, `スライド ID ${id} が範囲外`);
        const rel = presRels.find((x) => x.id === (0, xml_1.attr)(s, 'r:id'));
        if (!rel || rel.type !== package_1.REL.slide)
            add(pres.path, `スライド ID ${id} の r:id が slide の関係を指していない`);
    }
    // p14:sectionLst と p:custShowLst の ID が存在する
    for (const sid of (0, xml_1.descendants)(pres.root, 'p14:sldId')) {
        if (!slideIds.has((0, xml_1.attr)(sid, 'id') ?? ''))
            add(pres.path, `セクションのスライド ID ${(0, xml_1.attr)(sid, 'id')} が存在しない`);
    }
    const seenInSections = new Set();
    for (const sid of (0, xml_1.descendants)(pres.root, 'p14:sldId')) {
        const id = (0, xml_1.attr)(sid, 'id') ?? '';
        if (seenInSections.has(id))
            add(pres.path, `スライド ID ${id} が複数のセクションにある`);
        seenInSections.add(id);
    }
    if ((0, xml_1.descendants)(pres.root, 'p14:sectionLst').length) {
        const secOrder = (0, xml_1.descendants)(pres.root, 'p14:sldId').map((x) => (0, xml_1.attr)(x, 'id') ?? '');
        const order = pres.slides.map((s) => s.id);
        if (secOrder.join(',') !== order.join(','))
            add(pres.path, 'セクションの順とスライドの順が一致しない');
    }
    for (const sld of (0, xml_1.descendants)((0, xml_1.kid)(pres.root, 'p:custShowLst') ?? pres.root, 'p:sld')) {
        if (!(0, xml_1.is)(sld.parentNode, 'p:sldLst'))
            continue;
        const rel = presRels.find((x) => x.id === (0, xml_1.attr)(sld, 'r:id'));
        if (!rel || rel.type !== package_1.REL.slide || !r.has(rel.target))
            add(pres.path, `目的別スライドショーの r:id ${(0, xml_1.attr)(sld, 'r:id')} が存在しない`);
    }
    // スライドごとに cNvPr@id が一意、ノートとスライドが 1 対 1
    const notesOwner = new Map();
    for (const s of pres.slides) {
        if (broken.has(s.path))
            continue;
        const root = (0, package_1.rootEl)(r.dom(s.path));
        const ids = new Map();
        for (const c of (0, xml_1.descendants)((0, xml_1.kid)(root, 'p:cSld') ?? root, 'p:cNvPr')) {
            const id = (0, xml_1.attr)(c, 'id') ?? '';
            ids.set(id, (ids.get(id) ?? 0) + 1);
        }
        for (const [id, n] of ids)
            if (n > 1)
                add(s.path, `cNvPr@id ${id} が ${n} 個ある`);
        const notes = (0, package_1.relsOfType)(r, s.path, package_1.REL.notesSlide);
        if (notes.length > 1)
            add(s.path, 'ノートスライドの関係が複数ある');
        for (const n of notes) {
            if (notesOwner.has(n.target))
                add(n.target, `ノートスライドが複数のスライドから参照されている`);
            notesOwner.set(n.target, s.path);
            // Google スライドが書き出す pptx は、ノートからスライドへの関係を持たない（PowerPoint も開ける）。
            // 関係がないのは許し、あるのに別のスライドを指している場合だけを問題にする。
            const back = (0, package_1.relsOfType)(r, n.target, package_1.REL.slide);
            if (back.length > 1 || (back.length === 1 && back[0].target !== s.path)) {
                add(n.target, `ノートスライドの slide の関係が ${s.path} を指していない`);
            }
        }
        // 子要素の順序
        checkOrderIn(root, s.path, add);
        // スライドの XML の r:id などが、スライドの関係を指している（複製で残った p188:commentRel など）
        checkRIdRefs(r, root, s.path, add);
    }
    for (const p of r.paths()) {
        if (/^ppt\/notesSlides\/[^/]+\.xml$/.test(p) && !notesOwner.has(p))
            add(p, 'どのスライドからも参照されていないノートスライド');
    }
    for (const p of r.paths()) {
        if (/^ppt\/notesSlides\/[^/]+\.xml$/.test(p) && !broken.has(p))
            checkOrderIn((0, package_1.rootEl)(r.dom(p)), p, add);
    }
    checkOrderIn(pres.root, pres.path, add);
    // ノートマスター
    // （python-pptx の出力のように p:notesMasterIdLst がないファイルもあり、PowerPoint は開けるので問題にしない）
    const nmIds = (0, xml_1.descendants)(pres.root, 'p:notesMasterId');
    for (const id of nmIds) {
        const rel = presRels.find((x) => x.id === (0, xml_1.attr)(id, 'r:id'));
        if (!rel || rel.type !== package_1.REL.notesMaster)
            add(pres.path, 'p:notesMasterId の r:id が notesMaster を指していない');
    }
    return issues;
}
/** パーツの XML の中の、関係 ID の属性（r:id、r:embed など）がすべて、そのパーツの関係にあるか。 */
function checkRIdRefs(r, root, part, add) {
    const rels = (0, package_1.readRels)(r, part);
    const ids = new Set(rels.map((x) => x.id));
    // p188:commentRel は、新形式のコメントの関係を指す（古い版の複製で、別の関係の ID を指したまま残ることがある）
    for (const cr of (0, xml_1.descendants)(root, 'p188:commentRel')) {
        const rel = rels.find((x) => x.id === (0, xml_1.attr)(cr, 'r:id'));
        if (rel && rel.type !== comments_1.REL_MODERN_COMMENTS)
            add(part, `p188:commentRel の r:id="${rel.id}" がコメントの関係ではない`);
    }
    const seen = new Set();
    for (const e of [root, ...(0, xml_1.descendants)(root)]) {
        for (let i = 0; i < e.attributes.length; i++) {
            const a = e.attributes.item(i);
            if (a.namespaceURI !== xml_1.NS.r || ids.has(a.value))
                continue;
            const key = `${e.tagName} ${a.name}=${a.value}`;
            if (seen.has(key))
                continue;
            seen.add(key);
            add(part, `${e.tagName} の ${a.name}="${a.value}" の関係がない`);
        }
    }
}
function checkOrderIn(root, part, add) {
    for (const e of [root, ...(0, xml_1.descendants)(root)]) {
        const tag = e.tagName;
        if (!ORDER_CHECKED.has(tag))
            continue;
        const msg = (0, xml_1.checkChildOrder)(e);
        if (msg)
            add(part, msg);
    }
}
