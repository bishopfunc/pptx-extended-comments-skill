"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.BROKEN_COMMENT_PART_MSG = exports.ALL = exports.COMMENT_READ_LIMITS = exports.CT_AUTHORS = exports.CT_MODERN_COMMENTS = exports.REL_LEGACY_AUTHORS = exports.REL_AUTHORS = exports.REL_MODERN_COMMENTS = void 0;
exports.monikerFor = monikerFor;
exports.txBodyText = txBodyText;
exports.shapeText = shapeText;
exports.presentationPartOf = presentationPartOf;
exports.authorsPartOf = authorsPartOf;
exports.commentPartOf = commentPartOf;
exports.brokenCommentPartsOf = brokenCommentPartsOf;
exports.safeRoot = safeRoot;
exports.readAuthors = readAuthors;
exports.anchorElOf = anchorElOf;
exports.repliesOf = repliesOf;
exports.statusOf = statusOf;
exports.shapeMonikerOf = shapeMonikerOf;
exports.textMonikerOf = textMonikerOf;
exports.resolveMoniker = resolveMoniker;
exports.threadElsOf = threadElsOf;
exports.readComments = readComments;
exports.anchorModel = anchorModel;
exports.ensureAuthorsPart = ensureAuthorsPart;
exports.ensureAuthor = ensureAuthor;
exports.authorNameOf = authorNameOf;
exports.ensureCommentPart = ensureCommentPart;
exports.usedSlideCreationIds = usedSlideCreationIds;
exports.removeCommentPartIfEmpty = removeCommentPartIfEmpty;
exports.removeCommentPart = removeCommentPart;
exports.commentTxBody = commentTxBody;
const comments_1 = require("../shared/comments.js");
const creationIds_1 = require("./creationIds.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const package_1 = require("./package.js");
const types_1 = require("./types.js");
const xml_1 = require("./xml.js");
const creationIds_2 = require("./creationIds.js");
const commentImages_1 = require("./commentImages.js");
exports.REL_MODERN_COMMENTS = 'http://schemas.microsoft.com/office/2018/10/relationships/comments';
exports.REL_AUTHORS = 'http://schemas.microsoft.com/office/2018/10/relationships/authors';
exports.REL_LEGACY_AUTHORS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors';
exports.CT_MODERN_COMMENTS = 'application/vnd.ms-powerpoint.comments+xml';
exports.CT_AUTHORS = 'application/vnd.ms-powerpoint.authors+xml';
/** 外から来たコメントのパーツを読むときの上限。超えた分は読み飛ばす（モデルと Webview への送信を大きくしない）。 */
exports.COMMENT_READ_LIMITS = {
    /** 1 つのパーツのスレッドの数。 */
    threadsPerPart: 2000,
    /** 1 つのスレッドの返信の数。 */
    repliesPerThread: 1000,
    /** 全体のスレッドの数。 */
    threadsTotal: 20000,
    /** 1 件の本文の文字数（超えた分は表示しない。XML はそのまま残す）。 */
    textChars: 20000,
    imageCharsTotal: 24 * 1024 * 1024,
    /** 文字の範囲のアンカーの引用の文字数（パネルには 1 行だけ出す）。 */
    quoteChars: 500,
    /** 作成者の名前の文字数（エディタの設定の名前と同じ上限）。イニシャルは 3 文字まで。 */
    authorChars: 100,
    /** 全体の投稿（返信を含む）の数。 */
    postsTotal: 20000,
    /**
     * 全体の文字数（本文・引用・作成者の名前。投稿ごとに数える）。作成者の名前や引用は投稿ごとに写るので、
     * 小さな pptx からでも大きなモデルを作れる（長い名前の作成者 1 人と、多くの短いコメント）。
     */
    charsTotal: 5_000_000,
};
const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
/** 図形の種類 → アンカーのモニカー（ac:*Mk）の名前。 */
const MONIKER = {
    sp: 'ac:spMk',
    pic: 'ac:picMk',
    graphicFrame: 'ac:graphicFrameMk',
    cxnSp: 'ac:cxnSpMk',
    grpSp: 'ac:grpSpMk',
};
const MONIKER_NAMES = new Set(['spMk', 'picMk', 'graphicFrameMk', 'cxnSpMk', 'grpSpMk']);
function monikerFor(el) {
    return el.namespaceURI === xml_1.NS.p ? MONIKER[el.localName ?? ''] : undefined;
}
// ───────────── 文字列 ─────────────
/** a:p の並び（図形の p:txBody、コメントの p188:txBody）の文字列。段落は \n、改行（a:br）も \n（どちらも 1 文字）。 */
function txBodyText(txBody) {
    return (0, xml_1.kidsNamed)(txBody, 'a:p')
        .map((p) => (0, xml_1.kids)(p)
        .map((k) => ((0, xml_1.is)(k, 'a:br') ? '\n' : (0, xml_1.isAny)(k, ['a:r', 'a:fld']) ? ((0, xml_1.kid)(k, 'a:t')?.textContent ?? '') : ''))
        .join(''))
        .join('\n');
}
/** 図形（p:sp）の文字列。テキストのない図形は ''。文字の範囲のアンカーはこの文字列の UTF-16 の位置で数える。 */
function shapeText(el) {
    return (0, xml_1.is)(el, 'p:sp') ? txBodyText((0, xml_1.kid)(el, 'p:txBody')) : '';
}
function clip(s, max = exports.COMMENT_READ_LIMITS.textChars) {
    return s.length > max ? `${s.slice(0, max)}…` : s;
}
/** 作成者の名前とイニシャル（外から来るので長さを抑える。XML はそのまま）。 */
function authorInfo(id, rawName, rawInitials) {
    const name = clip(rawName, exports.COMMENT_READ_LIMITS.authorChars);
    const initials = rawInitials ? [...rawInitials].slice(0, 3).join('') : (0, comments_1.initialsOf)(name);
    return { id, name, initials };
}
// ───────────── パーツの場所 ─────────────
function presentationPartOf(r) {
    return (0, inherit_1.readPresentation)(r).path;
}
/** 新形式の作成者のパーツ（なければ undefined）。 */
function authorsPartOf(r) {
    const pres = presentationPartOf(r);
    const rel = (0, package_1.relsOfType)(r, pres, exports.REL_AUTHORS).find((x) => !x.external && r.has(x.target));
    return rel?.target;
}
/** スライドの新形式のコメントのパーツ（XML として読めるもの。なければ undefined）。 */
function commentPartOf(r, slidePath) {
    return (0, package_1.relsOfType)(r, slidePath, exports.REL_MODERN_COMMENTS).find((x) => !x.external && r.has(x.target) && !!safeRoot(r, x.target))?.target;
}
/** スライドの新形式のコメントのパーツのうち、あるが XML として読めないもの（壊れている）。 */
function brokenCommentPartsOf(r, slidePath) {
    return (0, package_1.relsOfType)(r, slidePath, exports.REL_MODERN_COMMENTS)
        .filter((x) => !x.external && r.has(x.target) && !safeRoot(r, x.target))
        .map((x) => x.target);
}
/** XML として読めればルート要素。壊れていれば undefined（コメントのパーツは外から来るので、読めなくても止めない）。 */
function safeRoot(r, path) {
    try {
        return (0, package_1.rootEl)(r.dom(path)) ?? undefined;
    }
    catch {
        return undefined;
    }
}
function readAuthors(r) {
    const out = new Map();
    const p = authorsPartOf(r);
    const root = p ? safeRoot(r, p) : undefined;
    for (const a of (0, xml_1.kidsNamed)(root, 'p188:author')) {
        const id = (0, xml_1.attr)(a, 'id');
        if (!id || out.has(id))
            continue;
        out.set(id, authorInfo(id, (0, xml_1.attr)(a, 'name') ?? '', (0, xml_1.attr)(a, 'initials')));
    }
    return out;
}
// ───────────── スレッドの XML ─────────────
/** p188:cm の子のうち、アンカー（*MkLst、p188:unknownAnchor）。 */
function anchorElOf(cm) {
    return (0, xml_1.kids)(cm).find((k) => (k.localName ?? '').endsWith('MkLst') || (0, xml_1.is)(k, 'p188:unknownAnchor'));
}
/** スレッドの返信の要素。表示用には上限まで。書き換える処理は ALL を渡す（上限より後ろの返信を消さないように）。 */
function repliesOf(cm, limit = exports.COMMENT_READ_LIMITS.repliesPerThread) {
    return (0, xml_1.kidsNamed)((0, xml_1.kid)(cm, 'p188:replyLst'), 'p188:reply').slice(0, limit);
}
/** repliesOf・threadElsOf に渡す「上限なし」（XML を書き換える処理で使う）。 */
exports.ALL = Number.POSITIVE_INFINITY;
function statusOf(cm) {
    const s = (0, xml_1.attr)(cm, 'status');
    return s === 'resolved' || s === 'closed' ? 'resolved' : 'active';
}
/** アンカーの中の図形のモニカー（最後のもの。グループの中の図形ならその図形）。 */
function shapeMonikerOf(anchor) {
    const mks = (0, xml_1.kids)(anchor).filter((k) => k.namespaceURI === xml_1.NS.ac && MONIKER_NAMES.has(k.localName ?? ''));
    return mks[mks.length - 1];
}
function textMonikerOf(anchor) {
    return (0, xml_1.kids)(anchor).find((k) => (0, xml_1.is)(k, 'ac:txMk'));
}
/**
 * スライドの中で、モニカーが指す図形。creationId があれば creationId で、なければ cNvPr@id で探す。
 * 同じ creationId の図形が複数あれば（古い版の複製や、ほかのツールの pptx）、cNvPr@id も同じものを選ぶ。
 */
function resolveMoniker(refs, mk) {
    const cid = (0, xml_1.attr)(mk, 'creationId');
    const id = (0, xml_1.attr)(mk, 'id');
    if (cid) {
        const lower = cid.toLowerCase();
        const hits = refs.filter((x) => (0, creationIds_1.readShapeCreationId)((0, inherit_1.cNvPrOf)(x.el))?.toLowerCase() === lower);
        if (hits.length <= 1)
            return hits[0];
        return hits.find((x) => x.cnvId === id) ?? hits[0];
    }
    return id === undefined ? undefined : refs.find((x) => x.cnvId === id);
}
/** スライドの新形式のスレッドの要素。表示用には上限まで。書き換える処理は ALL を渡す。 */
function threadElsOf(r, slidePath, limit = exports.COMMENT_READ_LIMITS.threadsPerPart) {
    const part = commentPartOf(r, slidePath);
    if (!part)
        return undefined;
    const root = safeRoot(r, part);
    if (!root)
        return undefined;
    return { part, cms: (0, xml_1.kidsNamed)(root, 'p188:cm').slice(0, limit) };
}
// ───────────── 表示用モデル ─────────────
/** パッケージのすべてのコメント（新形式と旧形式）。 */
function readComments(r) {
    const pres = (0, inherit_1.readPresentation)(r);
    const authors = readAuthors(r);
    const legacyAuthors = readLegacyAuthors(r, pres.path);
    const threads = [];
    const counts = {};
    let posts = 0;
    let chars = 0;
    let truncated = false;
    let imageChars = 0;
    let imagesTruncated = false;
    for (const s of pres.slides) {
        if (threads.length >= exports.COMMENT_READ_LIMITS.threadsTotal || truncated) {
            truncated = true;
            break;
        }
        const found = threadElsOf(r, s.path);
        const legacy = readLegacy(r, s.path, s.id, legacyAuthors);
        if (!found && !legacy.length)
            continue;
        const slideRoot = safeRoot(r, s.path);
        const refs = slideRoot ? (0, elements_1.indexElements)((0, elements_1.spTreeOf)(slideRoot)) : [];
        const list = [];
        for (const cm of found?.cms ?? []) {
            const t = threadModel(cm, s.id, refs, authors);
            if (t)
                list.push(t);
        }
        list.push(...legacy);
        list.sort((a, b) => (a.root.created < b.root.created ? -1 : a.root.created > b.root.created ? 1 : 0));
        for (const t of list) {
            for (const p of [t.root, ...t.replies]) {
                if (!p.images)
                    continue;
                p.images = p.images.filter(image => {
                    if (imageChars + image.dataUrl.length > exports.COMMENT_READ_LIMITS.imageCharsTotal) {
                        imagesTruncated = true;
                        return false;
                    }
                    imageChars += image.dataUrl.length;
                    return true;
                });
            }
            const c = (counts[s.id] ??= { open: 0, resolved: 0 });
            if (t.status === 'resolved')
                c.resolved++;
            else
                c.open++;
            if (truncated)
                continue;
            // 全体の上限：超えたら、それ以降のスレッドはモデルに入れない（数は counts に数える）
            const size = threadChars(t);
            if (posts + 1 + t.replies.length > exports.COMMENT_READ_LIMITS.postsTotal || chars + size > exports.COMMENT_READ_LIMITS.charsTotal) {
                truncated = true;
                continue;
            }
            posts += 1 + t.replies.length;
            chars += size;
            threads.push(t);
        }
    }
    if (threads.length > exports.COMMENT_READ_LIMITS.threadsTotal)
        truncated = true;
    return { threads: threads.slice(0, exports.COMMENT_READ_LIMITS.threadsTotal), counts, ...(truncated || imagesTruncated ? { truncated: true } : {}) };
}
function threadChars(t) {
    const post = (p) => p.text.length + p.author.length + p.initials.length + p.id.length + p.created.length;
    const quote = t.anchor.kind === 'text' ? (t.anchor.quote?.length ?? 0) : 0;
    return post(t.root) + t.replies.reduce((n, r) => n + post(r), 0) + quote;
}
function postModel(el, authors) {
    const id = (0, xml_1.attr)(el, 'id');
    if (!id)
        return undefined;
    const a = authors.get((0, xml_1.attr)(el, 'authorId') ?? '');
    const name = a?.name ?? '不明な作成者';
    const images = (0, commentImages_1.readCommentImages)(el);
    return {
        id,
        author: name,
        initials: a?.initials ?? (0, comments_1.initialsOf)(name),
        created: (0, xml_1.attr)(el, 'created') ?? '',
        text: clip(txBodyText((0, xml_1.kid)(el, 'p188:txBody'))),
        ...(images.length ? { images } : {}),
    };
}
function threadModel(cm, slideId, refs, authors) {
    const root = postModel(cm, authors);
    if (!root)
        return undefined;
    const replies = repliesOf(cm)
        .map((x) => postModel(x, authors))
        .filter((x) => !!x);
    return { id: root.id, slideId, format: 'modern', status: statusOf(cm), anchor: anchorModel(anchorElOf(cm), slideId, refs), root, replies };
}
function anchorModel(anchor, slideId, refs) {
    if (!anchor || (0, xml_1.is)(anchor, 'pc:sldMkLst'))
        return { kind: 'slide' };
    const isDe = (0, xml_1.is)(anchor, 'ac:deMkLst');
    const isTx = (0, xml_1.is)(anchor, 'ac:txMkLst');
    if (!isDe && !isTx)
        return { kind: 'unknown', tag: anchor.tagName };
    const mk = shapeMonikerOf(anchor);
    if (!mk)
        return { kind: 'unknown', tag: anchor.tagName };
    const ref = resolveMoniker(refs, mk);
    const elementId = ref ? `${slideId}:${ref.key}` : undefined;
    const orphaned = !ref;
    if (isTx) {
        const tx = textMonikerOf(anchor);
        const start = Math.max(0, Number((0, xml_1.attr)(tx, 'cp') ?? 0) || 0);
        const length = Math.max(0, Number((0, xml_1.attr)(tx, 'len') ?? 0) || 0);
        const text = ref ? shapeText(ref.el) : undefined;
        return {
            kind: 'text',
            elementId,
            orphaned,
            start,
            length,
            ...(text !== undefined ? { quote: clip(text.slice(start, start + length), exports.COMMENT_READ_LIMITS.quoteChars) } : {}),
        };
    }
    const isTable = mk.localName === 'graphicFrameMk' && (!ref || (0, xml_1.descendants)(ref.el, 'a:tbl').length > 0);
    return isTable ? { kind: 'table', elementId, orphaned } : { kind: 'element', elementId, orphaned };
}
// ───────────── 旧形式（読み取り専用） ─────────────
function readLegacyAuthors(r, presPath) {
    const out = new Map();
    const rel = (0, package_1.relsOfType)(r, presPath, exports.REL_LEGACY_AUTHORS).find((x) => !x.external && r.has(x.target));
    const root = rel ? safeRoot(r, rel.target) : undefined;
    for (const a of (0, xml_1.kidsNamed)(root, 'p:cmAuthor')) {
        const id = (0, xml_1.attr)(a, 'id') ?? '';
        out.set(id, authorInfo(id, (0, xml_1.attr)(a, 'name') ?? '', (0, xml_1.attr)(a, 'initials')));
    }
    return out;
}
/** 旧形式の p:cm@dt（タイムゾーンなしはローカル時刻）を、表示用モデルの形（タイムゾーンを書かない UTC）にそろえる。 */
function legacyTime(dt) {
    const d = (0, comments_1.parseCommentTime)(dt, 'local');
    return d ? (0, comments_1.formatCommentTime)(d) : dt;
}
function readLegacy(r, slidePath, slideId, authors) {
    const rel = (0, package_1.relsOfType)(r, slidePath, package_1.REL.comments).find((x) => !x.external && r.has(x.target));
    const root = rel ? safeRoot(r, rel.target) : undefined;
    if (!root)
        return [];
    const threads = [];
    const byKey = new Map();
    for (const cm of (0, xml_1.kidsNamed)(root, 'p:cm').slice(0, exports.COMMENT_READ_LIMITS.threadsPerPart)) {
        const authorId = (0, xml_1.attr)(cm, 'authorId') ?? '';
        const key = `${authorId}:${(0, xml_1.attr)(cm, 'idx') ?? ''}`;
        const a = authors.get(authorId);
        const name = a?.name || '不明な作成者';
        const post = {
            id: `legacy:${key}`,
            author: name,
            initials: a?.initials ?? (0, comments_1.initialsOf)(name),
            created: legacyTime((0, xml_1.attr)(cm, 'dt') ?? ''),
            text: clip((0, xml_1.kid)(cm, 'p:text')?.textContent ?? ''),
        };
        // 返信（p15:threadingInfo/p15:parentCm）なら親のスレッドに足す
        const parent = (0, xml_1.descendants)(cm, 'p15:parentCm')[0];
        const parentThread = parent ? byKey.get(`${(0, xml_1.attr)(parent, 'authorId') ?? ''}:${(0, xml_1.attr)(parent, 'idx') ?? ''}`) : undefined;
        if (parentThread) {
            if (parentThread.replies.length < exports.COMMENT_READ_LIMITS.repliesPerThread)
                parentThread.replies.push(post);
            continue;
        }
        const t = { id: post.id, slideId, format: 'legacy', status: 'active', anchor: { kind: 'slide' }, root: post, replies: [] };
        byKey.set(key, t);
        threads.push(t);
    }
    return threads;
}
// ───────────── 書き込み：作成者 ─────────────
function newXmlPart(txn, path, rootQName) {
    return txn.createXml(path, `${DECL}<${rootQName} xmlns:a="${xml_1.NS.a}" xmlns:r="${xml_1.NS.r}" xmlns:p188="${xml_1.NS.p188}"/>`);
}
function uniquePartPath(txn, path) {
    if (!txn.has(path))
        return path;
    const stem = path.replace(/\.xml$/i, '');
    for (let n = 1;; n++)
        if (!txn.has(`${stem}_${n}.xml`))
            return `${stem}_${n}.xml`;
}
/**
 * 作成者のパーツ（なければ作る）。先の関係が無いパーツを指していれば、その関係を取り除く。
 * 関係のない ppt/authors.xml（p188:authorLst として読める）があれば、それに関係を付けて使う（そこにある作成者を不明にしない）。
 */
function ensureAuthorsPart(txn) {
    const existing = authorsPartOf(txn);
    if (existing && safeRoot(txn, existing))
        return existing;
    const pres = presentationPartOf(txn);
    for (const rel of (0, package_1.relsOfType)(txn, pres, exports.REL_AUTHORS))
        if (!rel.external && !txn.has(rel.target))
            txn.removeRel(pres, rel.id);
    const orphan = 'ppt/authors.xml';
    if (!existing && txn.has(orphan) && (0, xml_1.is)(safeRoot(txn, orphan), 'p188:authorLst')) {
        txn.addRel(pres, exports.REL_AUTHORS, orphan);
        txn.addOverride(orphan, exports.CT_AUTHORS);
        return orphan;
    }
    const path = uniquePartPath(txn, 'ppt/authors.xml');
    newXmlPart(txn, path, 'p188:authorLst');
    txn.addRel(pres, exports.REL_AUTHORS, path);
    txn.addOverride(path, exports.CT_AUTHORS);
    return path;
}
/**
 * 名前の作成者の ID（なければ足す）。userId には名前を入れ、providerId は "None"（メールアドレスなどは書かない）。
 */
function ensureAuthor(txn, name) {
    const path = ensureAuthorsPart(txn);
    const root = (0, package_1.rootEl)(txn.edit(path));
    const all = (0, xml_1.kidsNamed)(root, 'p188:author');
    const found = all.find((a) => (0, xml_1.attr)(a, 'name') === name && (0, xml_1.attr)(a, 'providerId') === 'None') ?? all.find((a) => (0, xml_1.attr)(a, 'name') === name);
    const id = (0, xml_1.attr)(found, 'id');
    if (found && id)
        return id;
    const a = (0, xml_1.create)(root.ownerDocument, 'p188:author', {
        id: (0, creationIds_1.newGuid)(),
        name,
        initials: (0, comments_1.initialsOf)(name),
        userId: name,
        providerId: 'None',
    });
    root.appendChild(a);
    return (0, xml_1.attr)(a, 'id');
}
/** 作成者の名前（authorId から）。 */
function authorNameOf(r, authorId) {
    return authorId ? readAuthors(r).get(authorId)?.name : undefined;
}
// ───────────── 書き込み：スライドのコメントのパーツ ─────────────
/** スライドのコメントのパーツが壊れている（XML として読めない）ときの拒否の文。 */
exports.BROKEN_COMMENT_PART_MSG = 'このスライドのコメントのパーツが壊れているため、コメントを書き込めません（pptx を修復してください）';
/**
 * スライドのコメントのパーツ（なければ作り、関係・p188:commentRel・Content Type を足す）。
 * 古い版の複製などで残った、無いパーツへの関係と、コメントのパーツを指さない p188:commentRel は取り除き、
 * スライドの p188:commentRel が 1 つになるようにする。パーツはあるが XML として読めなければ拒否する
 * （別のパーツを作ると、読む側は壊れた方を見て、書いたコメントが見えなくなる）。
 */
function ensureCommentPart(txn, slide) {
    const existing = commentPartOf(txn, slide.path);
    if (!existing && brokenCommentPartsOf(txn, slide.path).length)
        throw new types_1.OpRejectedError(exports.BROKEN_COMMENT_PART_MSG);
    // 無いパーツへの関係を取り除く
    for (const rel of (0, package_1.relsOfType)(txn, slide.path, exports.REL_MODERN_COMMENTS)) {
        if (rel.external || txn.has(rel.target))
            continue;
        txn.removeRel(slide.path, rel.id);
        txn.removeOverride(rel.target);
    }
    let path = existing;
    let rId = path ? (0, package_1.relsOfType)(txn, slide.path, exports.REL_MODERN_COMMENTS).find((x) => !x.external && x.target === path)?.id : undefined;
    if (!path || !rId) {
        const sldRoot = (0, package_1.rootEl)(txn.edit(slide.path));
        const cId = (0, creationIds_1.ensureSlideCreationId)(sldRoot, usedSlideCreationIds(txn, slide.path));
        const hex = (n) => n.toString(16).toUpperCase();
        path = uniquePartPath(txn, `ppt/comments/modernComment_${hex(Number(slide.id) || 0)}_${hex(cId)}.xml`);
        newXmlPart(txn, path, 'p188:cmLst');
        rId = txn.addRel(slide.path, exports.REL_MODERN_COMMENTS, path);
        txn.addOverride(path, exports.CT_MODERN_COMMENTS);
    }
    syncCommentRel(txn, slide.path, rId);
    return path;
}
/** p:sld/p:extLst の p188:commentRel の拡張を、rId を指す 1 つだけにする（変える必要がなければスライドを書き換えない）。 */
function syncCommentRel(txn, slidePath, rId) {
    const exts = (root) => (0, xml_1.kidsNamed)((0, xml_1.kid)(root, 'p:extLst'), 'p:ext').filter((e) => (0, xml_1.attr)(e, 'uri') === creationIds_2.URI_COMMENT_REL);
    const ok = (e) => {
        const rels = (0, xml_1.kidsNamed)(e, 'p188:commentRel');
        return rels.length === 1 && (0, xml_1.attr)(rels[0], 'r:id') === rId;
    };
    const cur = safeRoot(txn, slidePath);
    if (cur) {
        const found = exts(cur);
        if (found.length === 1 && ok(found[0]))
            return;
    }
    const sldRoot = (0, package_1.rootEl)(txn.edit(slidePath));
    const doc = sldRoot.ownerDocument;
    const found = exts(sldRoot);
    const keep = found[0];
    for (const e of found.slice(1))
        (0, xml_1.remove)(e);
    if (keep) {
        for (const k of (0, xml_1.kids)(keep))
            (0, xml_1.remove)(k);
        const rel = doc.createElementNS(xml_1.NS.p188, 'p188:commentRel');
        rel.setAttributeNS(xml_1.NS.r, 'r:id', rId);
        keep.appendChild(rel);
        return;
    }
    // p:sld/p:extLst（p:sld の最後の子）に p188:commentRel
    const extLst = (0, xml_1.ensureKid)(sldRoot, 'p:extLst');
    const ext = (0, xml_1.create)(doc, 'p:ext', { uri: creationIds_2.URI_COMMENT_REL });
    const rel = doc.createElementNS(xml_1.NS.p188, 'p188:commentRel');
    rel.setAttributeNS(xml_1.NS.r, 'r:id', rId);
    ext.appendChild(rel);
    extLst.appendChild(ext);
}
/** ほかのスライドの p14:creationId（重ならない値を振るため）。 */
function usedSlideCreationIds(r, except) {
    const out = [];
    for (const s of (0, inherit_1.readPresentation)(r).slides) {
        if (s.path === except)
            continue;
        const root = safeRoot(r, s.path);
        const v = root ? (0, creationIds_1.readSlideCreationId)(root) : undefined;
        if (v !== undefined)
            out.push(v);
    }
    return out;
}
/** スレッドがなくなったコメントのパーツを、関係・p188:commentRel・Content Type ごと取り除く。 */
function removeCommentPartIfEmpty(txn, slidePath) {
    const part = commentPartOf(txn, slidePath);
    if (!part)
        return;
    const root = safeRoot(txn, part);
    if (!root || (0, xml_1.kidsNamed)(root, 'p188:cm').length > 0)
        return;
    removeCommentPart(txn, slidePath);
}
/** スライドの新形式のコメントのパーツを、関係・p188:commentRel・Content Type ごと取り除く。 */
function removeCommentPart(txn, slidePath) {
    for (const rel of (0, package_1.readRels)(txn, slidePath).filter((x) => x.type === exports.REL_MODERN_COMMENTS)) {
        const sldRoot = (0, package_1.rootEl)(txn.edit(slidePath));
        for (const cr of (0, xml_1.descendants)(sldRoot, 'p188:commentRel')) {
            if ((0, xml_1.attr)(cr, 'r:id') !== rel.id)
                continue;
            const ext = cr.parentNode;
            const lst = ext?.parentNode;
            (0, xml_1.remove)(ext && (0, xml_1.is)(ext, 'p:ext') ? ext : cr);
            if (lst && (0, xml_1.is)(lst, 'p:extLst') && (0, xml_1.kids)(lst).length === 0)
                (0, xml_1.remove)(lst);
        }
        txn.removeRel(slidePath, rel.id);
        if (!rel.external && txn.has(rel.target)) {
            txn.delete(rel.target);
            txn.delete((0, package_1.relsPathOf)(rel.target));
            txn.removeOverride(rel.target);
        }
    }
}
/** p188:txBody を作る（段落は \n で分ける）。 */
function commentTxBody(doc, text) {
    const body = (0, xml_1.create)(doc, 'p188:txBody');
    body.appendChild((0, xml_1.create)(doc, 'a:bodyPr'));
    body.appendChild((0, xml_1.create)(doc, 'a:lstStyle'));
    for (const line of (0, xml_1.stripInvalidXml)(text).replace(/\r\n?/g, '\n').split('\n')) {
        const p = (0, xml_1.create)(doc, 'a:p');
        if (line) {
            const run = (0, xml_1.create)(doc, 'a:r');
            run.appendChild((0, xml_1.create)(doc, 'a:rPr', { lang: /[぀-ヿ㐀-鿿]/.test(line) ? 'ja-JP' : 'en-US' }));
            const t = (0, xml_1.create)(doc, 'a:t');
            t.appendChild(doc.createTextNode(line));
            run.appendChild(t);
            p.appendChild(run);
        }
        body.appendChild(p);
    }
    return body;
}
