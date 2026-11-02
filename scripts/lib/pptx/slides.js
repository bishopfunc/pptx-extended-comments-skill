"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.FORBIDDEN_REL_HINTS = exports.SLIDE_MSG = void 0;
exports.isCommentRel = isCommentRel;
exports.nextSlideId = nextSlideId;
exports.nextPartNumber = nextPartNumber;
exports.firstLayout = firstLayout;
exports.registerSlide = registerSlide;
exports.writeNotes = writeNotes;
exports.setLayoutRel = setLayoutRel;
exports.applySlideOp = applySlideOp;
const creationIds_1 = require("./creationIds.js");
const elements_1 = require("./elements.js");
const edit_1 = require("./edit.js");
const inherit_1 = require("./inherit.js");
const package_1 = require("./package.js");
const parse_1 = require("./parse.js");
const layouts_1 = require("./layouts.js");
const templates_1 = require("./templates.js");
const xml_1 = require("./xml.js");
const SECTION_EXT_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';
const MODERN_COMMENTS = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';
exports.SLIDE_MSG = {
    lastSlide: '最後の 1 枚のスライドは削除できません',
    forbiddenDup: 'グラフ、SmartArt、OLE を含むスライドは、このバージョンでは複製できません',
    badIndex: '移動先の位置が正しくありません',
    noLayout: 'スライドのレイアウトが見つかりません',
};
function isCommentRel(type) {
    return type === package_1.REL.comments || type === MODERN_COMMENTS || type.endsWith('/comments');
}
// ---------------------------------------------------------------------------
// 採番
// ---------------------------------------------------------------------------
/** max(256, 既存の最大値 + 1)。2147483648 以上なら [256, 2147483647] の未使用の最小値。 */
function nextSlideId(existing) {
    const max = existing.length ? Math.max(...existing) : 0;
    const cand = Math.max(256, max + 1);
    if (cand < 2147483648)
        return cand;
    const used = new Set(existing);
    for (let i = 256; i <= 2147483647; i++)
        if (!used.has(i))
            return i;
    throw new Error('スライド ID を採番できません');
}
function nextPartNumber(txn, dir, stem) {
    let max = 0;
    const re = new RegExp(`^${dir}/${stem}(\\d+)\\.xml$`);
    for (const p of txn.paths()) {
        const m = re.exec(p);
        if (m)
            max = Math.max(max, Number(m[1]));
    }
    return max + 1;
}
function presRoot(txn, pres) {
    return (0, package_1.rootEl)(txn.edit(pres.path));
}
function sldIdLst(root) {
    let lst = (0, xml_1.kid)(root, 'p:sldIdLst');
    if (!lst)
        lst = (0, xml_1.insertChildInOrder)(root, (0, xml_1.create)(root.ownerDocument, 'p:sldIdLst'));
    return lst;
}
// ---------------------------------------------------------------------------
// セクション（p14:sectionLst）
// ---------------------------------------------------------------------------
function sectionLst(root) {
    const extLst = (0, xml_1.kid)(root, 'p:extLst');
    for (const ext of (0, xml_1.kidsNamed)(extLst, 'p:ext')) {
        if ((0, xml_1.attr)(ext, 'uri') === SECTION_EXT_URI) {
            const s = (0, xml_1.kid)(ext, 'p14:sectionLst');
            if (s)
                return s;
        }
    }
    return (0, xml_1.descendants)(root, 'p14:sectionLst')[0];
}
function sectionIdLists(root) {
    const lst = sectionLst(root);
    if (!lst)
        return [];
    return (0, xml_1.kidsNamed)(lst, 'p14:section').map((s) => {
        let ids = (0, xml_1.kid)(s, 'p14:sldIdLst');
        if (!ids) {
            ids = (0, xml_1.create)(root.ownerDocument, 'p14:sldIdLst');
            s.insertBefore(ids, s.firstChild);
        }
        return ids;
    });
}
/** 新しい ID を、afterId と同じセクションのすぐ後ろに入れる（afterId がなければ最初のセクションの先頭）。 */
function insertIntoSection(root, newId, afterId) {
    const lists = sectionIdLists(root);
    if (!lists.length)
        return;
    const el = (0, xml_1.create)(root.ownerDocument, 'p14:sldId', { id: newId });
    if (afterId !== undefined) {
        for (const l of lists) {
            const prev = (0, xml_1.kidsNamed)(l, 'p14:sldId').find((x) => (0, xml_1.attr)(x, 'id') === afterId);
            if (prev) {
                if (prev.nextSibling)
                    l.insertBefore(el, prev.nextSibling);
                else
                    l.appendChild(el);
                return;
            }
        }
        lists[lists.length - 1].appendChild(el);
        return;
    }
    const first = lists[0];
    first.insertBefore(el, first.firstChild);
}
function removeFromSections(root, id) {
    for (const l of sectionIdLists(root))
        for (const x of (0, xml_1.kidsNamed)(l, 'p14:sldId'))
            if ((0, xml_1.attr)(x, 'id') === id)
                (0, xml_1.remove)(x);
}
// ---------------------------------------------------------------------------
// 追加
// ---------------------------------------------------------------------------
const NEW_SLIDE_PH = new Set(['title', 'ctrTitle', 'subTitle', 'body', 'obj']);
function firstLayout(txn, pres) {
    const masterRel = (0, package_1.relsOfType)(txn, pres.path, package_1.REL.slideMaster)[0];
    if (!masterRel)
        return undefined;
    return (0, package_1.relsOfType)(txn, masterRel.target, package_1.REL.slideLayout)[0]?.target;
}
function registerSlide(txn, pres, slidePath, afterId) {
    const root = presRoot(txn, pres);
    const lst = sldIdLst(root);
    const existing = (0, xml_1.kidsNamed)(lst, 'p:sldId').map((x) => Number((0, xml_1.attr)(x, 'id')));
    const newId = String(nextSlideId(existing.filter((n) => Number.isFinite(n))));
    const rId = txn.addRel(pres.path, package_1.REL.slide, slidePath);
    txn.addOverride(slidePath, package_1.CT.slide);
    const el = (0, xml_1.create)(root.ownerDocument, 'p:sldId', { id: newId, 'r:id': rId });
    const prev = afterId !== undefined ? (0, xml_1.kidsNamed)(lst, 'p:sldId').find((x) => (0, xml_1.attr)(x, 'id') === afterId) : undefined;
    if (prev) {
        if (prev.nextSibling)
            lst.insertBefore(el, prev.nextSibling);
        else
            lst.appendChild(el);
    }
    else {
        lst.insertBefore(el, lst.firstChild);
    }
    insertIntoSection(root, newId, prev ? afterId : undefined);
    return newId;
}
/** レイアウトのプレースホルダー（spTree の直下の図形）。 */
function layoutPlaceholders(layoutRoot) {
    return (0, xml_1.kids)((0, elements_1.spTreeOf)(layoutRoot)).filter((k) => (0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS) && !!(0, inherit_1.phOf)(k));
}
/** レイアウトのプレースホルダー layoutPh と同じ type / idx の空のプレースホルダーをスライドの spTree に入れる。 */
function addEmptyPlaceholder(tree, layoutPh, id) {
    const phEl = (0, xml_1.kid)((0, xml_1.kid)((0, inherit_1.nvOf)(layoutPh), 'p:nvPr'), 'p:ph');
    let attrs = '';
    const t = (0, xml_1.attr)(phEl, 'type');
    if (t)
        attrs += ` type="${t}"`;
    const idx = (0, xml_1.attr)(phEl, 'idx');
    if (idx !== undefined)
        attrs += ` idx="${idx}"`;
    const label = (0, inherit_1.normalizePhType)((0, inherit_1.phOf)(layoutPh).type) === 'title' ? 'Title' : 'Content Placeholder';
    (0, xml_1.insertChildInOrder)(tree, (0, xml_1.fragment)(tree.ownerDocument, (0, templates_1.placeholderSpXml)(id, `${label} ${id - 1}`, attrs)));
}
function addSlide(txn, op) {
    const pres = (0, inherit_1.readPresentation)(txn);
    const after = op.afterSlideId !== null ? pres.slides.find((s) => s.id === op.afterSlideId) : undefined;
    if (op.afterSlideId !== null && !after)
        (0, edit_1.reject)('対象のスライドが見つかりません');
    let layoutPath;
    if (op.layoutPath !== undefined) {
        // 指定されたレイアウト：このプレゼンテーションのマスターに属するものだけを受け付ける
        if (!(0, layouts_1.allLayoutPaths)(txn, pres).includes(op.layoutPath))
            (0, edit_1.reject)(exports.SLIDE_MSG.noLayout);
        layoutPath = op.layoutPath;
    }
    else {
        const ref = after ?? pres.slides[0];
        layoutPath = (ref ? (0, inherit_1.layoutPathOf)(txn, ref.path) : undefined) ?? firstLayout(txn, pres);
    }
    if (!layoutPath || !txn.has(layoutPath))
        (0, edit_1.reject)(exports.SLIDE_MSG.noLayout);
    const num = nextPartNumber(txn, 'ppt/slides', 'slide');
    const slidePath = `ppt/slides/slide${num}.xml`;
    const doc = txn.createXml(slidePath, (0, templates_1.emptySlideXml)());
    const root = (0, package_1.rootEl)(doc);
    const tree = (0, elements_1.spTreeOf)(root);
    const layoutRoot = (0, package_1.rootEl)(txn.dom(layoutPath));
    let id = (0, elements_1.nextShapeId)(root);
    for (const k of layoutPlaceholders(layoutRoot)) {
        if (!NEW_SLIDE_PH.has((0, inherit_1.phOf)(k).type))
            continue;
        addEmptyPlaceholder(tree, k, id++);
    }
    txn.addRel(slidePath, package_1.REL.slideLayout, layoutPath);
    const newId = registerSlide(txn, pres, slidePath, after?.id);
    return { changes: [], selectSlide: newId };
}
// ---------------------------------------------------------------------------
// 複製
// ---------------------------------------------------------------------------
exports.FORBIDDEN_REL_HINTS = ['/chart', '/diagram', '/oleObject', '/package'];
function duplicateSlide(txn, op) {
    const pres = (0, inherit_1.readPresentation)(txn);
    const src = pres.slides.find((s) => s.id === op.slideId);
    if (!src)
        (0, edit_1.reject)('対象のスライドが見つかりません');
    const srcRoot = (0, package_1.rootEl)(txn.dom(src.path));
    const rels = (0, package_1.readRels)(txn, src.path);
    if ((0, elements_1.containsForbidden)(srcRoot) || rels.some((r) => exports.FORBIDDEN_REL_HINTS.some((h) => r.type.includes(h)))) {
        (0, edit_1.reject)(exports.SLIDE_MSG.forbiddenDup);
    }
    const num = nextPartNumber(txn, 'ppt/slides', 'slide');
    const slidePath = `ppt/slides/slide${num}.xml`;
    txn.put(slidePath, txn.text(src.path));
    // 関係：ノートとコメント以外をそのまま（同じ rId で）コピーする。参照先のパーツは共有する。
    // 外した関係を r:id で参照している要素（新形式のコメントの p188:commentRel など）も取り除く
    const srcRels = (0, package_1.relsPathOf)(src.path);
    if (txn.has(srcRels)) {
        txn.put((0, package_1.relsPathOf)(slidePath), txn.text(srcRels));
        for (const r of rels) {
            if (r.type === package_1.REL.notesSlide || isCommentRel(r.type)) {
                removeRIdUses(txn, slidePath, r.id);
                txn.removeRel(slidePath, r.id);
            }
        }
    }
    // スライドの作成 ID（p14:creationId）は新しく振る（コメントのアンカーが、元のスライドと取り違えないように）
    const copyRoot = (0, package_1.rootEl)(txn.edit(slidePath));
    if ((0, creationIds_1.readSlideCreationId)(copyRoot) !== undefined) {
        const used = pres.slides.map((s) => (0, creationIds_1.readSlideCreationId)((0, package_1.rootEl)(txn.dom(s.path)))).filter((v) => v !== undefined);
        (0, creationIds_1.writeSlideCreationId)(copyRoot, undefined, used);
    }
    const newId = registerSlide(txn, pres, slidePath, src.id);
    // ノート
    const srcNotes = (0, parse_1.notesPathOf)(txn, src.path);
    if (srcNotes) {
        const n = nextPartNumber(txn, 'ppt/notesSlides', 'notesSlide');
        const notesPath = `ppt/notesSlides/notesSlide${n}.xml`;
        txn.put(notesPath, txn.text(srcNotes));
        txn.addOverride(notesPath, package_1.CT.notesSlide);
        const srcNotesRels = (0, package_1.relsPathOf)(srcNotes);
        if (txn.has(srcNotesRels)) {
            txn.put((0, package_1.relsPathOf)(notesPath), txn.text(srcNotesRels));
            const relsDoc = txn.edit((0, package_1.relsPathOf)(notesPath));
            for (const r of (0, xml_1.kids)((0, package_1.rootEl)(relsDoc))) {
                if ((0, xml_1.attr)(r, 'Type') === package_1.REL.slide)
                    r.setAttribute('Target', `../slides/slide${num}.xml`);
            }
        }
        else {
            const master = (0, package_1.relsOfType)(txn, pres.path, package_1.REL.notesMaster)[0]?.target;
            if (master)
                txn.addRel(notesPath, package_1.REL.notesMaster, master);
            txn.addRel(notesPath, package_1.REL.slide, slidePath);
        }
        txn.addRel(slidePath, package_1.REL.notesSlide, notesPath);
    }
    return { changes: [], selectSlide: newId };
}
// ---------------------------------------------------------------------------
// 削除
// ---------------------------------------------------------------------------
/** パーツ partPath の XML から、rId を使っている属性（a:hlinkClick などは要素ごと）を取り除く。 */
function removeRIdUses(txn, partPath, rId) {
    if (!txn.has(partPath) || !/\.xml$/i.test(partPath))
        return;
    const root = (0, package_1.rootEl)(txn.dom(partPath));
    const uses = [];
    for (const e of [root, ...(0, xml_1.descendants)(root)]) {
        for (let i = 0; i < e.attributes.length; i++) {
            const a = e.attributes.item(i);
            if (a.namespaceURI === xml_1.NS.r && a.value === rId) {
                uses.push(e);
                break;
            }
        }
    }
    if (!uses.length)
        return;
    const editRoot = (0, package_1.rootEl)(txn.edit(partPath));
    for (const e of [editRoot, ...(0, xml_1.descendants)(editRoot)]) {
        const names = [];
        for (let i = 0; i < e.attributes.length; i++) {
            const a = e.attributes.item(i);
            if (a.namespaceURI === xml_1.NS.r && a.value === rId)
                names.push(a.localName ?? a.name);
        }
        if (!names.length)
            continue;
        const ext = extensionOf(e);
        if (ext) {
            // 拡張（p:ext / a:ext）の中の参照（p188:commentRel など）は、拡張ごと取り除く。空になった extLst も
            const lst = ext.parentNode;
            (0, xml_1.remove)(ext);
            if (lst && (0, xml_1.is)(lst, 'p:extLst') && (0, xml_1.kids)(lst).length === 0)
                (0, xml_1.remove)(lst);
            else if (lst && (0, xml_1.is)(lst, 'a:extLst') && (0, xml_1.kids)(lst).length === 0)
                (0, xml_1.remove)(lst);
        }
        else if ((0, xml_1.isAny)(e, ['a:hlinkClick', 'a:hlinkMouseOver', 'a:hlinkHover', 'p:sld', 'p:sldId'])) {
            (0, xml_1.remove)(e);
        }
        else {
            for (const n of names)
                e.removeAttributeNS(xml_1.NS.r, n);
        }
    }
}
/** e が拡張（p:ext / a:ext）の中にあれば、その拡張。 */
function extensionOf(e) {
    for (let n = e.parentNode; n && n.nodeType === 1; n = n.parentNode) {
        if ((0, xml_1.isAny)(n, ['p:ext', 'a:ext']))
            return n;
    }
    return undefined;
}
function allRelsSources(txn) {
    return txn
        .paths()
        .filter((p) => p.endsWith('.rels'))
        .map(package_1.sourceOfRels);
}
function deletePart(txn, p) {
    txn.delete(p);
    txn.delete((0, package_1.relsPathOf)(p));
    txn.removeOverride(p);
}
function deleteSlide(txn, op) {
    const pres = (0, inherit_1.readPresentation)(txn);
    const s = pres.slides.find((x) => x.id === op.slideId);
    if (!s)
        (0, edit_1.reject)('対象のスライドが見つかりません');
    if (pres.slides.length <= 1)
        (0, edit_1.reject)(exports.SLIDE_MSG.lastSlide);
    const root = presRoot(txn, pres);
    for (const x of (0, xml_1.kidsNamed)((0, xml_1.kid)(root, 'p:sldIdLst'), 'p:sldId'))
        if ((0, xml_1.attr)(x, 'id') === s.id)
            (0, xml_1.remove)(x);
    removeFromSections(root, s.id);
    const slideRels = (0, package_1.readRels)(txn, s.path);
    const toDelete = new Set([s.path]);
    for (const r of slideRels) {
        if (r.external)
            continue;
        if (r.type === package_1.REL.notesSlide || isCommentRel(r.type))
            toDelete.add(r.target);
    }
    // 削除するパーツが参照しているパーツ（共有パーツの候補）
    const candidates = new Set();
    for (const p of toDelete)
        for (const r of (0, package_1.readRels)(txn, p))
            if (!r.external && !toDelete.has(r.target))
                candidates.add(r.target);
    // すべての rels を走査し、削除するパーツを指す関係とその r:id の利用を取り除く
    for (const src of allRelsSources(txn)) {
        if (toDelete.has(src))
            continue;
        for (const r of (0, package_1.readRels)(txn, src)) {
            if (r.external || !toDelete.has(r.target))
                continue;
            removeRIdUses(txn, src, r.id);
            txn.removeRel(src, r.id);
        }
    }
    for (const p of toDelete)
        deletePart(txn, p);
    // 他から参照されなくなった共有パーツ（画像など）を削除する
    const referenced = new Set();
    for (const src of allRelsSources(txn))
        for (const r of (0, package_1.readRels)(txn, src))
            if (!r.external)
                referenced.add(r.target);
    for (const c of candidates) {
        if (referenced.has(c))
            continue;
        if (/^ppt\/(slides|slideLayouts|slideMasters|notesMasters|theme|handoutMasters)\//.test(c))
            continue;
        if (c === pres.path)
            continue;
        deletePart(txn, c);
    }
}
/** 選んだスライドをまとめて削除する（同じトランザクションで 1 枚ずつ。すべてのスライドは削除できない）。 */
function deleteSlides(txn, op) {
    const ids = [...new Set(op.slideIds)];
    const pres = (0, inherit_1.readPresentation)(txn);
    if (!ids.length)
        (0, edit_1.reject)('削除するスライドがありません');
    if (ids.some((id) => !pres.slides.some((x) => x.id === id)))
        (0, edit_1.reject)('対象のスライドが見つかりません');
    if (ids.length >= pres.slides.length)
        (0, edit_1.reject)(exports.SLIDE_MSG.lastSlide);
    for (const slideId of ids)
        deleteSlide(txn, { type: 'deleteSlide', slideId });
}
// ---------------------------------------------------------------------------
// 移動
// ---------------------------------------------------------------------------
/**
 * 選んだスライドを、今の順のまま、選んでいないスライドの並びの toIndex の位置にまとめて移す。
 * 目的の並びを先に決め、前から順に 1 枚ずつ moveSlide で置く（セクションの組み直しも moveSlide に任せる）。
 */
function moveSlides(txn, op) {
    const pres = (0, inherit_1.readPresentation)(txn);
    const order = pres.slides.map((x) => x.id);
    const ids = new Set(op.slideIds);
    if (!ids.size)
        (0, edit_1.reject)('移動するスライドがありません');
    if ([...ids].some((id) => !order.includes(id)))
        (0, edit_1.reject)('対象のスライドが見つかりません');
    const moving = order.filter((id) => ids.has(id));
    const rest = order.filter((id) => !ids.has(id));
    if (!Number.isInteger(op.toIndex) || op.toIndex < 0 || op.toIndex > rest.length)
        (0, edit_1.reject)(exports.SLIDE_MSG.badIndex);
    const target = [...rest.slice(0, op.toIndex), ...moving, ...rest.slice(op.toIndex)];
    const cur = [...order];
    target.forEach((id, i) => {
        if (cur[i] === id)
            return;
        moveSlide(txn, { type: 'moveSlide', slideId: id, toIndex: i });
        cur.splice(cur.indexOf(id), 1);
        cur.splice(i, 0, id);
    });
}
function moveSlide(txn, op) {
    const pres = (0, inherit_1.readPresentation)(txn);
    const from = pres.slides.findIndex((x) => x.id === op.slideId);
    if (from < 0)
        (0, edit_1.reject)('対象のスライドが見つかりません');
    const n = pres.slides.length;
    if (!Number.isInteger(op.toIndex) || op.toIndex < 0 || op.toIndex >= n)
        (0, edit_1.reject)(exports.SLIDE_MSG.badIndex);
    if (from === op.toIndex)
        return;
    const root = presRoot(txn, pres);
    const lst = sldIdLst(root);
    const items = (0, xml_1.kidsNamed)(lst, 'p:sldId');
    const el = items.find((x) => (0, xml_1.attr)(x, 'id') === op.slideId);
    (0, xml_1.remove)(el);
    const rest = (0, xml_1.kidsNamed)(lst, 'p:sldId');
    const before = rest[op.toIndex];
    if (before)
        lst.insertBefore(el, before);
    else
        lst.appendChild(el);
    // セクション：セクションの順に並べたときにスライドの順と一致するよう組み直す
    const lists = sectionIdLists(root);
    if (!lists.length)
        return;
    const order = (0, xml_1.kidsNamed)(lst, 'p:sldId').map((x) => (0, xml_1.attr)(x, 'id') ?? '');
    removeFromSections(root, op.slideId);
    const secEl = (0, xml_1.create)(root.ownerDocument, 'p14:sldId', { id: op.slideId });
    const idx = order.indexOf(op.slideId);
    if (idx > 0) {
        const predId = order[idx - 1];
        for (const l of lists) {
            const pred = (0, xml_1.kidsNamed)(l, 'p14:sldId').find((x) => (0, xml_1.attr)(x, 'id') === predId);
            if (pred) {
                if (pred.nextSibling)
                    l.insertBefore(secEl, pred.nextSibling);
                else
                    l.appendChild(secEl);
                return;
            }
        }
        lists[lists.length - 1].appendChild(secEl);
    }
    else {
        const succId = order[1];
        for (const l of lists) {
            const succ = (0, xml_1.kidsNamed)(l, 'p14:sldId').find((x) => (0, xml_1.attr)(x, 'id') === succId);
            if (succ) {
                l.insertBefore(secEl, succ);
                return;
            }
        }
        lists[0].insertBefore(secEl, lists[0].firstChild);
    }
}
// ---------------------------------------------------------------------------
// ノート
// ---------------------------------------------------------------------------
function ensureNotesMaster(txn, pres) {
    const existing = (0, package_1.relsOfType)(txn, pres.path, package_1.REL.notesMaster)[0];
    if (existing && txn.has(existing.target))
        return existing.target;
    const root = presRoot(txn, pres);
    const sz = (0, xml_1.kid)(root, 'p:notesSz');
    const n = nextPartNumber(txn, 'ppt/notesMasters', 'notesMaster');
    const nmPath = `ppt/notesMasters/notesMaster${n}.xml`;
    txn.put(nmPath, (0, templates_1.notesMasterXml)((0, xml_1.numAttr)(sz, 'cx') ?? 6858000, (0, xml_1.numAttr)(sz, 'cy') ?? 9144000));
    txn.addOverride(nmPath, package_1.CT.notesMaster);
    // テーマ：スライドマスターのテーマをコピーする
    const firstSlide = pres.slides[0];
    const master = (0, inherit_1.masterPathOf)(txn, firstSlide ? (0, inherit_1.layoutPathOf)(txn, firstSlide.path) : firstLayout(txn, pres));
    const srcTheme = master ? (0, package_1.relsOfType)(txn, master, package_1.REL.theme)[0]?.target : undefined;
    const t = nextPartNumber(txn, 'ppt/theme', 'theme');
    const themePath = `ppt/theme/theme${t}.xml`;
    const themeText = (srcTheme && txn.text(srcTheme)) ??
        `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<a:theme xmlns:a="${xml_1.NS.a}" name="Office Theme"><a:themeElements/></a:theme>`;
    txn.put(themePath, themeText);
    txn.addOverride(themePath, package_1.CT.theme);
    txn.addRel(nmPath, package_1.REL.theme, themePath);
    const rId = txn.addRel(pres.path, package_1.REL.notesMaster, nmPath);
    let lst = (0, xml_1.kid)(root, 'p:notesMasterIdLst');
    if (!lst)
        lst = (0, xml_1.insertChildInOrder)(root, (0, xml_1.create)(root.ownerDocument, 'p:notesMasterIdLst'));
    for (const k of (0, xml_1.kids)(lst))
        (0, xml_1.remove)(k);
    lst.appendChild((0, xml_1.create)(root.ownerDocument, 'p:notesMasterId', { 'r:id': rId }));
    return nmPath;
}
function createNotesSlide(txn, pres, slidePath) {
    const nm = ensureNotesMaster(txn, pres);
    const n = nextPartNumber(txn, 'ppt/notesSlides', 'notesSlide');
    const notesPath = `ppt/notesSlides/notesSlide${n}.xml`;
    txn.put(notesPath, (0, templates_1.notesSlideXml)());
    txn.addOverride(notesPath, package_1.CT.notesSlide);
    txn.addRel(notesPath, package_1.REL.notesMaster, nm);
    txn.addRel(notesPath, package_1.REL.slide, slidePath);
    txn.addRel(slidePath, package_1.REL.notesSlide, notesPath);
    return notesPath;
}
function setNotes(txn, op) {
    const s = (0, edit_1.findSlide)(txn, op.slideId);
    writeNotes(txn, s.slidePath, op.text);
}
/** スライドのノートの本文を text にする（ノートスライドがなければ作る）。同じなら何もしない。 */
function writeNotes(txn, slidePath, rawText) {
    const text = rawText.replace(/\r\n?/g, '\n');
    if ((0, parse_1.readNotes)(txn, slidePath) === text)
        return;
    const pres = (0, inherit_1.readPresentation)(txn);
    const notesPath = (0, parse_1.notesPathOf)(txn, slidePath) ?? createNotesSlide(txn, pres, slidePath);
    const doc = txn.edit(notesPath);
    const root = (0, package_1.rootEl)(doc);
    let body = (0, parse_1.notesBodyShape)(root);
    if (!body) {
        body = (0, xml_1.fragment)(doc, (0, templates_1.notesBodyXml)((0, elements_1.nextShapeId)(root)));
        (0, xml_1.insertChildInOrder)((0, elements_1.spTreeOf)(root), body);
    }
    let tx = (0, xml_1.kid)(body, 'p:txBody');
    if (!tx) {
        tx = (0, xml_1.create)(doc, 'p:txBody');
        tx.appendChild((0, xml_1.create)(doc, 'a:bodyPr'));
        tx.appendChild((0, xml_1.create)(doc, 'a:lstStyle'));
        (0, xml_1.insertChildInOrder)(body, tx);
    }
    const origPs = (0, xml_1.kidsNamed)(tx, 'a:p');
    const first = origPs[0];
    const pPr = (0, xml_1.kid)(first, 'a:pPr');
    const firstRun = first ? (0, parse_1.runChildren)(first).find((r) => (0, xml_1.is)(r, 'a:r')) : undefined;
    const rPr = (0, xml_1.kid)(firstRun, 'a:rPr');
    for (const p of origPs)
        (0, xml_1.remove)(p);
    const lines = text === '' ? [''] : text.split('\n');
    for (const line of lines) {
        const p = (0, xml_1.create)(doc, 'a:p');
        if (pPr)
            p.appendChild(pPr.cloneNode(true));
        if (line) {
            const r = (0, xml_1.create)(doc, 'a:r');
            if (rPr) {
                const c = rPr.cloneNode(true);
                for (const k of (0, xml_1.kids)(c))
                    if ((0, xml_1.isAny)(k, ['a:hlinkClick', 'a:hlinkMouseOver']))
                        (0, xml_1.remove)(k);
                r.appendChild(c);
            }
            const t = (0, xml_1.create)(doc, 'a:t');
            t.appendChild(doc.createTextNode((0, xml_1.stripInvalidXml)(line)));
            r.appendChild(t);
            p.appendChild(r);
        }
        else {
            const end = (0, xml_1.create)(doc, 'a:endParaRPr', { lang: (0, xml_1.attr)(rPr, 'lang') ?? 'ja-JP' });
            p.appendChild(end);
        }
        tx.appendChild(p);
    }
}
// ---------------------------------------------------------------------------
// スキップ（p:sld@show）
// ---------------------------------------------------------------------------
function setSlideHidden(txn, op) {
    const want = op.hidden === true;
    for (const id of new Set(op.slideIds)) {
        const s = (0, edit_1.findSlide)(txn, id);
        const show = (0, xml_1.attr)((0, package_1.rootEl)(txn.dom(s.slidePath)), 'show');
        if (want ? show === '0' : show === undefined)
            continue;
        (0, xml_1.setAttr)((0, package_1.rootEl)(txn.edit(s.slidePath)), 'show', want ? '0' : undefined);
    }
}
// ---------------------------------------------------------------------------
// 背景（p:cSld/p:bg）
// ---------------------------------------------------------------------------
/**
 * 単色なら p:bg/p:bgPr（a:solidFill と a:effectLst。p:bgPr は塗りと効果が必須）を書き、既存の p:bg（p:bgRef も）を置き換える。
 * null なら p:bg を取り除き、レイアウト・マスターの背景に戻す。p:bg は p:cSld の最初の子。
 */
function setSlideBackground(txn, op) {
    let color;
    if (op.fill !== null) {
        if (!op.fill || op.fill.type !== 'solid' || typeof op.fill.color !== 'string')
            (0, edit_1.reject)(edit_1.MSG.badColor);
        color = (0, edit_1.hexOf)(op.fill.color);
    }
    for (const id of new Set(op.slideIds)) {
        const s = (0, edit_1.findSlide)(txn, id);
        const cur = (0, xml_1.kid)((0, package_1.rootEl)(txn.dom(s.slidePath)), 'p:cSld');
        if (!cur)
            (0, edit_1.reject)(edit_1.MSG.slideNotFound);
        if (color === undefined && !(0, xml_1.kid)(cur, 'p:bg'))
            continue;
        const doc = txn.edit(s.slidePath);
        const cSld = (0, xml_1.kid)((0, package_1.rootEl)(doc), 'p:cSld');
        for (const bg of (0, xml_1.kidsNamed)(cSld, 'p:bg'))
            (0, xml_1.remove)(bg);
        if (color === undefined)
            continue;
        const bg = (0, xml_1.create)(doc, 'p:bg');
        const bgPr = (0, xml_1.create)(doc, 'p:bgPr');
        const fill = (0, xml_1.create)(doc, 'a:solidFill');
        fill.appendChild((0, xml_1.create)(doc, 'a:srgbClr', { val: color }));
        bgPr.appendChild(fill);
        bgPr.appendChild((0, xml_1.create)(doc, 'a:effectLst'));
        bg.appendChild(bgPr);
        (0, xml_1.insertChildInOrder)(cSld, bg);
    }
}
// ---------------------------------------------------------------------------
// レイアウトの適用
// ---------------------------------------------------------------------------
const PH_ATTRS = ['type', 'orient', 'sz', 'idx'];
/** スライドのプレースホルダーの p:ph を、レイアウトのプレースホルダーと同じ type / orient / sz / idx にする。 */
function retargetPlaceholder(el, layoutPh) {
    const src = (0, xml_1.kid)((0, xml_1.kid)((0, inherit_1.nvOf)(layoutPh), 'p:nvPr'), 'p:ph');
    for (const b of (0, edit_1.branches)(el)) {
        const ph = (0, xml_1.kid)((0, xml_1.kid)((0, inherit_1.nvOf)(b), 'p:nvPr'), 'p:ph');
        if (!ph)
            continue;
        for (const a of PH_ATTRS)
            (0, xml_1.setAttr)(ph, a, (0, xml_1.attr)(src, a));
    }
}
/** 文字のないプレースホルダー（p:sp で、a:t がすべて空白）。 */
function isEmptyPlaceholder(el) {
    if (!(0, xml_1.is)(el, 'p:sp'))
        return false;
    if ((0, xml_1.descendants)(el, 'a:blipFill').length)
        return false;
    return (0, xml_1.descendants)(el, 'a:t').every((t) => (t.textContent ?? '').trim() === '');
}
/**
 * スライドのレイアウトを変える（Google スライドの「レイアウトを適用」に近い動き）。
 * - スライドのプレースホルダーは、新しいレイアウトのプレースホルダーと正規化した type で照合し、同じ type が複数あれば
 *   idx が同じものを優先する。照合できたら p:ph の type / idx をレイアウトのものに書き換える（位置・書式はレイアウトから継承）。
 *   スライドに a:xfrm があればそのまま（PowerPoint と同じく、動かした位置は保つ）。
 * - 照合できず a:xfrm も持たないプレースホルダーは、文字がなければ取り除き、文字があれば古いレイアウト・マスターから
 *   継承していた位置と大きさを a:xfrm に書く（その場から動かないように）。
 * - 新しいレイアウトのタイトル・本文のプレースホルダー（addSlide と同じ種類）で、スライドにないものは空で追加する。
 * - 最後に slideLayout の関係の参照先を変える。今と同じレイアウトなら何もしない。
 */
function applyLayout(txn, op) {
    const pres = (0, inherit_1.readPresentation)(txn);
    if (!(0, layouts_1.allLayoutPaths)(txn, pres).includes(op.layoutPath) || !txn.has(op.layoutPath))
        (0, edit_1.reject)(exports.SLIDE_MSG.noLayout);
    const layoutPhs = layoutPlaceholders((0, package_1.rootEl)(txn.dom(op.layoutPath)));
    for (const slideId of new Set(op.slideIds)) {
        const s = (0, edit_1.findSlide)(txn, slideId);
        const oldLayout = (0, inherit_1.layoutPathOf)(txn, s.slidePath);
        if (oldLayout === op.layoutPath)
            continue;
        const root = (0, package_1.rootEl)(txn.edit(s.slidePath));
        const tree = (0, elements_1.spTreeOf)(root);
        if (!tree)
            (0, edit_1.reject)(edit_1.MSG.slideNotFound);
        // 古いレイアウトでの継承（関係を書き換える前に作る）
        const ctx = (0, inherit_1.buildSlideCtx)(txn, s.slidePath);
        const used = new Set();
        const removedIds = new Set();
        for (const el of (0, xml_1.kids)(tree).filter((k) => (0, xml_1.isAny)(k, inherit_1.SHAPE_KINDS))) {
            const ph = (0, inherit_1.phOf)(el);
            if (!ph)
                continue;
            const t = (0, inherit_1.normalizePhType)(ph.type);
            const cands = layoutPhs.filter((l) => !used.has(l) && (0, inherit_1.normalizePhType)((0, inherit_1.phOf)(l).type) === t);
            const match = cands.find((l) => (0, inherit_1.phOf)(l).idx === ph.idx) ?? cands[0];
            if (match) {
                used.add(match);
                retargetPlaceholder(el, match);
                continue;
            }
            if ((0, inherit_1.xfrmElOf)(el))
                continue;
            if (isEmptyPlaceholder(el)) {
                const cid = (0, xml_1.attr)((0, inherit_1.cNvPrOf)(el), 'id');
                if (cid)
                    removedIds.add(cid);
                (0, xml_1.remove)(el);
                continue;
            }
            (0, edit_1.writeXfrm)(el, (0, parse_1.resolvedOwnXfrm)(ctx, el), false);
        }
        let id = (0, elements_1.nextShapeId)(root);
        for (const l of layoutPhs) {
            if (used.has(l) || !NEW_SLIDE_PH.has((0, inherit_1.phOf)(l).type))
                continue;
            addEmptyPlaceholder(tree, l, id++);
        }
        if (removedIds.size)
            (0, edit_1.cleanTiming)(root, removedIds);
        setLayoutRel(txn, s.slidePath, op.layoutPath);
    }
}
/** スライドの slideLayout の関係を layoutPath に向ける（なければ追加する）。 */
function setLayoutRel(txn, slidePath, layoutPath) {
    const rel = (0, package_1.relsOfType)(txn, slidePath, package_1.REL.slideLayout)[0];
    if (!rel) {
        txn.addRel(slidePath, package_1.REL.slideLayout, layoutPath);
        return;
    }
    const relsRoot = (0, package_1.rootEl)(txn.editRels(slidePath));
    for (const r of (0, xml_1.kids)(relsRoot))
        if ((0, xml_1.attr)(r, 'Id') === rel.id)
            r.setAttribute('Target', (0, package_1.relativeTarget)(slidePath, layoutPath));
}
function applySlideOp(txn, op) {
    switch (op.type) {
        case 'addSlide':
            return addSlide(txn, op);
        case 'duplicateSlide':
            return duplicateSlide(txn, op);
        case 'deleteSlide':
            deleteSlide(txn, op);
            return { changes: [] };
        case 'moveSlide':
            moveSlide(txn, op);
            return { changes: [] };
        case 'deleteSlides':
            deleteSlides(txn, op);
            return { changes: [] };
        case 'moveSlides':
            moveSlides(txn, op);
            return { changes: [] };
        case 'setNotes':
            setNotes(txn, op);
            return { changes: [] };
        case 'setSlideHidden':
            setSlideHidden(txn, op);
            return { changes: [] };
        case 'setSlideBackground':
            setSlideBackground(txn, op);
            return { changes: [] };
        case 'applyLayout':
            applyLayout(txn, op);
            return { changes: [] };
        default:
            return undefined;
    }
}
