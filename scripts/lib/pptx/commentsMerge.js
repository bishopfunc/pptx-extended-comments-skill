"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.mergeComments = mergeComments;
// ディスク上で書き換えられたコメント（エージェントの CLI など）を、未保存の編集を持つエディタの内容に取り込む
// 3 者マージ（設計書 8 章「外部変更とコメント」、ADR-0021）。base はエディタが最後に読み込んだ・保存した内容、disk は今のディスクの内容、
// local はエディタのメモリ上の内容（Txn）。スライドの中身は base と disk で同じこと（contentSha256 が同じ）を呼び出し側が確かめる。
//
// スレッドと返信は ID で突き合わせる。
// - 片方だけで足したものは足す。片方だけで消し、もう片方で変えていなければ消す（変えていれば残す）。
// - 本文・状態（解決済みか）・アンカーは、片方だけで変えたものを採る。両方で変えていればエディタ側を採る
//   （作成日時のほかに更新日時が記録されないため。どちらも自分の投稿にしか起きない）。
const comments_1 = require("./comments.js");
const commentsSummary_1 = require("./commentsSummary.js");
const commentOps_1 = require("./commentOps.js");
const commentImages_1 = require("./commentImages.js");
const creationIds_1 = require("./creationIds.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const package_1 = require("./package.js");
const xml_1 = require("./xml.js");
function threadsOf(r) {
    const out = new Map();
    for (const s of (0, inherit_1.readPresentation)(r).slides) {
        for (const cm of (0, comments_1.threadElsOf)(r, s.path, comments_1.ALL)?.cms ?? []) {
            const id = (0, xml_1.attr)(cm, 'id')?.toLowerCase();
            if (id && !out.has(id))
                out.set(id, { id, slideId: s.id, slidePath: s.path, cm });
        }
    }
    return out;
}
function commentTexts(r) {
    const out = [];
    for (const s of (0, inherit_1.readPresentation)(r).slides) {
        const found = (0, comments_1.threadElsOf)(r, s.path);
        if (found)
            out.push(r.text(found.part) ?? '');
    }
    return out;
}
const bodyOf = (post) => (0, comments_1.txBodyText)((0, xml_1.kid)(post, 'p188:txBody'));
const anchorKey = (cm) => {
    const a = (0, comments_1.anchorElOf)(cm);
    return a ? (0, xml_1.serializeXml)(a) : '';
};
/** base から変えたのはどちらか。両方なら local。 */
function pick(b, l, d) {
    if (l === d)
        return 'local';
    if (b !== undefined && l === b)
        return 'disk';
    return 'local';
}
function postChanged(b, x) {
    return !b || bodyOf(b) !== bodyOf(x) || (0, commentImages_1.imageKey)(b) !== (0, commentImages_1.imageKey)(x);
}
function threadChanged(b, x) {
    if (postChanged(b, x) || (0, comments_1.statusOf)(b) !== (0, comments_1.statusOf)(x) || anchorKey(b) !== anchorKey(x))
        return true;
    const br = new Map((0, comments_1.repliesOf)(b, comments_1.ALL).map((r) => [(0, xml_1.attr)(r, 'id')?.toLowerCase(), r]));
    const xr = (0, comments_1.repliesOf)(x, comments_1.ALL);
    if (br.size !== xr.length)
        return true;
    return xr.some((r) => postChanged(br.get((0, xml_1.attr)(r, 'id')?.toLowerCase()), r));
}
/**
 * base → disk のコメントの変化を、local（エディタの内容）に 3 者マージする。
 */
function mergeComments(txn, base, disk) {
    const before = (0, commentsSummary_1.commentSnapshot)(commentTexts(txn));
    const B = threadsOf(base);
    const L = threadsOf(txn);
    const D = threadsOf(disk);
    const localSlides = new Map((0, inherit_1.readPresentation)(txn).slides.map((s) => [s.id, s]));
    const report = { droppedForDeletedSlides: 0 };
    const touchedSlides = new Set();
    // 作成者：ディスク側の作成者を、同じ ID で足す
    const da = (0, comments_1.authorsPartOf)(disk);
    const dAuthors = da ? (0, xml_1.kidsNamed)((0, comments_1.safeRoot)(disk, da), 'p188:author') : [];
    const lAuthorIds = new Set((0, xml_1.kidsNamed)((0, comments_1.safeRoot)(txn, (0, comments_1.authorsPartOf)(txn) ?? '') ?? undefined, 'p188:author').map((a) => (0, xml_1.attr)(a, 'id')));
    const needAuthors = dAuthors.filter((a) => !lAuthorIds.has((0, xml_1.attr)(a, 'id')));
    const editCm = (t) => {
        const part = (0, comments_1.threadElsOf)(txn, t.slidePath)?.part;
        if (!part)
            return undefined;
        return (0, xml_1.kidsNamed)((0, package_1.rootEl)(txn.edit(part)), 'p188:cm').find((x) => (0, xml_1.attr)(x, 'id')?.toLowerCase() === t.id);
    };
    const ids = new Set([...B.keys(), ...L.keys(), ...D.keys()]);
    for (const id of ids) {
        const b = B.get(id);
        const l = L.get(id);
        const d = D.get(id);
        if (l && d) {
            if (b && !threadChanged(b.cm, d.cm))
                continue; // ディスク側は変えていない
            const target = editCm(l);
            if (target)
                mergeInto(target, b?.cm, l.cm, d.cm);
            continue;
        }
        if (l && !d) {
            // ディスク側で消した：エディタ側で変えていなければ消す
            if (b && !threadChanged(b.cm, l.cm)) {
                const target = editCm(l);
                if (target) {
                    (0, xml_1.remove)(target);
                    touchedSlides.add(l.slidePath);
                }
            }
            continue;
        }
        if (!l && d) {
            // ディスク側で足した、またはエディタ側で消したがディスク側で変えた
            if (b && !threadChanged(b.cm, d.cm))
                continue;
            const slide = localSlides.get(d.slideId);
            if (!slide) {
                report.droppedForDeletedSlides++;
                continue;
            }
            importThread(txn, disk, d, slide);
        }
    }
    if (needAuthors.length) {
        const root = (0, package_1.rootEl)(txn.edit((0, comments_1.ensureAuthorsPart)(txn)));
        for (const a of needAuthors)
            root.appendChild((0, xml_1.importClean)(root.ownerDocument, a));
    }
    for (const p of touchedSlides)
        (0, comments_1.removeCommentPartIfEmpty)(txn, p);
    report.summary = (0, commentsSummary_1.summarizeCommentChanges)(before, (0, commentsSummary_1.commentSnapshot)(commentTexts(txn)));
    return report;
}
/** target（エディタ側のスレッドの要素）に、base → disk の変化を書く。 */
function mergeInto(target, b, l, d) {
    const doc = target.ownerDocument;
    mergeImages(target, b, l, d);
    if (pick(b && bodyOf(b), bodyOf(l), bodyOf(d)) === 'disk') {
        const body = (0, xml_1.kid)(d, 'p188:txBody');
        const old = (0, xml_1.kid)(target, 'p188:txBody');
        if (body && old)
            target.replaceChild((0, xml_1.importClean)(doc, body), old);
    }
    if (pick(b && (0, comments_1.statusOf)(b), (0, comments_1.statusOf)(l), (0, comments_1.statusOf)(d)) === 'disk')
        (0, xml_1.setAttr)(target, 'status', (0, xml_1.attr)(d, 'status'));
    if (pick(b && anchorKey(b), anchorKey(l), anchorKey(d)) === 'disk') {
        const a = (0, comments_1.anchorElOf)(d);
        const old = (0, comments_1.anchorElOf)(target);
        if (a && old)
            target.replaceChild((0, xml_1.importClean)(doc, a), old);
    }
    // 返信：ID の和。片方で消して、もう片方で変えていなければ消す
    const bR = new Map((b ? (0, comments_1.repliesOf)(b, comments_1.ALL) : []).map((r) => [(0, xml_1.attr)(r, 'id')?.toLowerCase(), r]));
    const lR = new Map((0, comments_1.repliesOf)(target, comments_1.ALL).map((r) => [(0, xml_1.attr)(r, 'id')?.toLowerCase(), r]));
    const dR = new Map((0, comments_1.repliesOf)(d, comments_1.ALL).map((r) => [(0, xml_1.attr)(r, 'id')?.toLowerCase(), r]));
    const keep = [];
    for (const rid of new Set([...lR.keys(), ...dR.keys()])) {
        const br = bR.get(rid);
        const lr = lR.get(rid);
        const dr = dR.get(rid);
        if (lr && dr) {
            mergeImages(lr, br, lr, dr);
            if (pick(br && bodyOf(br), bodyOf(lr), bodyOf(dr)) === 'disk') {
                const nb = (0, xml_1.kid)(dr, 'p188:txBody');
                const ob = (0, xml_1.kid)(lr, 'p188:txBody');
                if (nb && ob)
                    lr.replaceChild((0, xml_1.importClean)(doc, nb), ob);
            }
            keep.push(lr);
        }
        else if (lr) {
            if (!br || postChanged(br, lr))
                keep.push(lr);
        }
        else if (dr) {
            if (!br || postChanged(br, dr))
                keep.push((0, xml_1.importClean)(doc, dr));
        }
    }
    keep.sort((x, y) => ((0, xml_1.attr)(x, 'created') ?? '').localeCompare((0, xml_1.attr)(y, 'created') ?? ''));
    let lst = (0, xml_1.kid)(target, 'p188:replyLst');
    if (lst)
        for (const r of (0, xml_1.kidsNamed)(lst, 'p188:reply'))
            (0, xml_1.remove)(r);
    if (keep.length) {
        if (!lst) {
            lst = (0, xml_1.create)(doc, 'p188:replyLst');
            const before = (0, xml_1.kid)(target, 'p188:txBody') ?? (0, xml_1.kid)(target, 'p188:extLst');
            if (before)
                target.insertBefore(lst, before);
            else
                target.appendChild(lst);
        }
        for (const r of keep)
            lst.appendChild(r);
    }
    else if (lst)
        (0, xml_1.remove)(lst);
}
/** 画像と本文を別々にマージし、本文の編集中にも外部の添付を取り込む。 */
function mergeImages(target, b, l, d) {
    if (pick(b && (0, commentImages_1.imageKey)(b), (0, commentImages_1.imageKey)(l), (0, commentImages_1.imageKey)(d)) !== 'disk')
        return;
    const old = (0, commentImages_1.imageExtension)(target);
    if (old)
        (0, xml_1.remove)(old);
    const incoming = (0, commentImages_1.imageExtension)(d);
    if (!incoming)
        return;
    let list = (0, xml_1.kid)(target, 'p188:extLst');
    if (!list) {
        list = (0, xml_1.create)(target.ownerDocument, 'p188:extLst');
        target.appendChild(list);
    }
    list.appendChild((0, xml_1.importClean)(target.ownerDocument, incoming));
}
/**
 * ディスク側だけにあるスレッドを、エディタの同じスライド（sldId）に足す。図形のアンカーは、ディスク側の図形と同じ cNvPr@id で
 * 同じ種類のエディタ側の図形に付ける（スライドの中身は base と同じなので、ID は同じ図形を指す）。エディタ側の図形に作成 ID が
 * なければディスク側の値を写し、あればアンカーをエディタ側の値に書き換える。エディタで消して別の図形が同じ ID を使っていれば
 * （文字列も位置・大きさも違う）、付けずに元の内容が削除されたアンカーとして残す。
 * 文字の範囲は、エディタでの未保存のテキストの編集に合わせて付け替える（ディスク側の文字列は base と同じ）。
 */
function importThread(txn, disk, d, slide) {
    const part = (0, comments_1.ensureCommentPart)(txn, slide);
    const root = (0, package_1.rootEl)(txn.edit(part));
    const cm = (0, xml_1.importClean)(root.ownerDocument, d.cm);
    const sldRoot = (0, package_1.rootEl)(txn.edit(slide.path));
    const cId = (0, creationIds_1.ensureSlideCreationId)(sldRoot, (0, comments_1.usedSlideCreationIds)(txn, slide.path));
    for (const mk of (0, xml_1.descendants)(cm, 'pc:sldMk')) {
        (0, xml_1.setAttr)(mk, 'sldId', slide.id);
        (0, xml_1.setAttr)(mk, 'cId', String(cId));
    }
    const anchor = (0, comments_1.anchorElOf)(cm);
    const mk = anchor && ((0, xml_1.is)(anchor, 'ac:deMkLst') || (0, xml_1.is)(anchor, 'ac:txMkLst')) ? (0, comments_1.shapeMonikerOf)(anchor) : undefined;
    if (mk) {
        const dRoot = (0, comments_1.safeRoot)(disk, d.slidePath);
        const dRef = dRoot ? (0, comments_1.resolveMoniker)((0, elements_1.indexElements)((0, elements_1.spTreeOf)(dRoot)), mk) : undefined;
        const lRefs = (0, elements_1.indexElements)((0, elements_1.spTreeOf)(sldRoot));
        const guid = (0, xml_1.attr)(mk, 'creationId');
        let lRef = guid ? lRefs.find((x) => (0, creationIds_1.readShapeCreationId)((0, inherit_1.cNvPrOf)(x.el))?.toLowerCase() === guid.toLowerCase()) : undefined;
        if (!lRef && dRef) {
            const cand = lRefs.find((x) => x.cnvId === dRef.cnvId && x.el.localName === dRef.el.localName);
            if (cand && sameElement(dRef.el, cand.el))
                lRef = cand;
            const cNvPr = lRef && (0, inherit_1.cNvPrOf)(lRef.el);
            if (lRef && cNvPr) {
                const local = (0, creationIds_1.readShapeCreationId)(cNvPr);
                if (!local && guid)
                    (0, creationIds_1.writeShapeCreationId)(cNvPr, guid);
                (0, xml_1.setAttr)(mk, 'id', lRef.cnvId);
                if (local)
                    (0, xml_1.setAttr)(mk, 'creationId', local);
            }
            else if (!guid) {
                // ID だけのアンカーが、別の図形を指さないように
                (0, xml_1.setAttr)(mk, 'creationId', (0, creationIds_1.newGuid)());
            }
        }
        const tx = (0, comments_1.textMonikerOf)(anchor);
        if (tx && dRef && lRef) {
            const r = (0, commentOps_1.shiftRange)((0, comments_1.shapeText)(dRef.el), (0, comments_1.shapeText)(lRef.el), Number((0, xml_1.attr)(tx, 'cp') ?? 0) || 0, Number((0, xml_1.attr)(tx, 'len') ?? 0) || 0);
            (0, xml_1.setAttr)(tx, 'cp', String(r.start));
            (0, xml_1.setAttr)(tx, 'len', String(r.length));
        }
    }
    root.appendChild(cm);
}
/** 同じ cNvPr@id の 2 つの図形が、同じ図形か（文字列か、位置・大きさが同じ。どちらも変えていれば別の図形とみなす）。 */
function sameElement(a, b) {
    return (0, comments_1.shapeText)(a) === (0, comments_1.shapeText)(b) || xfrmKey(a) === xfrmKey(b);
}
function xfrmKey(el) {
    const x = (0, xml_1.descendants)(el, 'a:xfrm')[0] ?? (0, xml_1.descendants)(el, 'p:xfrm')[0];
    return x ? (0, xml_1.serializeXml)(x) : '';
}
