"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.isCommentRelType = isCommentRelType;
exports.stripComments = stripComments;
exports.strippedCommentsMessage = strippedCommentsMessage;
const package_1 = require("./package.js");
const xml_1 = require("./xml.js");
const zipLimits_1 = require("./zipLimits.js");
/** コメントの関係の Type（旧形式と、PowerPoint 2018 以降の新形式）。 */
const COMMENT_REL_TYPES = new Set([
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors',
    'http://schemas.microsoft.com/office/2018/10/relationships/comments',
    'http://schemas.microsoft.com/office/2018/10/relationships/authors',
]);
/** コメントのパーツの Content Type。 */
const COMMENT_CONTENT_TYPES = new Set([
    'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
    'application/vnd.ms-powerpoint.comments+xml',
    'application/vnd.ms-powerpoint.authors+xml',
]);
/** 関係でも Content Type でも見つからなかったときのための、コメントのパーツの置き場所。 */
const COMMENT_PATH = /^ppt\/(comments\/[^/]+\.xml|commentAuthors\.xml|authors\.xml)$/i;
function isCommentRelType(type) {
    return COMMENT_REL_TYPES.has(type) || type.endsWith('/comments') || type.endsWith('/commentAuthors');
}
/**
 * pptx からコメントのパーツ、それを指す関係、[Content_Types].xml の Override、
 * スライドの p:extLst にある新形式のコメントへの参照を取り除く。
 * 展開は zipLimits の上限を確かめながら行い、読むのは rels・[Content_Types].xml・コメント・参照しているパーツだけ。
 */
async function stripComments(bytes, limits = zipLimits_1.DEFAULT_ZIP_LIMITS) {
    const zip = await (0, zipLimits_1.openZip)(bytes, limits);
    const budget = { total: 0 };
    const entries = new Map();
    for (const f of Object.values(zip.files))
        if (!f.dir)
            entries.set(f.name.replace(/^\/+/, ''), f);
    const unchanged = { bytes, comments: 0, removedParts: [] };
    const ctEntry = entries.get(package_1.CT_PATH);
    if (!ctEntry)
        return unchanged;
    const readText = async (name) => decodeXml(await (0, zipLimits_1.readZipEntry)(entries.get(name), budget, limits));
    // コメントのパーツ：Content Type、関係の参照先、置き場所のどれかで見つける
    const ctDoc = (0, xml_1.parseXml)(await readText(package_1.CT_PATH));
    const ctRoot = ctDoc.documentElement;
    const commentParts = new Set();
    for (const o of (0, xml_1.kids)(ctRoot)) {
        if (o.localName !== 'Override' || !COMMENT_CONTENT_TYPES.has((0, xml_1.attr)(o, 'ContentType') ?? ''))
            continue;
        const part = findEntry(entries, ((0, xml_1.attr)(o, 'PartName') ?? '').replace(/^\/+/, ''));
        if (part)
            commentParts.add(part);
    }
    for (const name of entries.keys())
        if (COMMENT_PATH.test(name))
            commentParts.add(name);
    const relsDocs = new Map();
    for (const name of entries.keys()) {
        if (!name.endsWith('.rels'))
            continue;
        const doc = (0, xml_1.parseXml)(await readText(name));
        const root = doc.documentElement;
        if (!root)
            continue;
        const src = (0, package_1.sourceOfRels)(name);
        const drop = [];
        for (const rel of (0, xml_1.kids)(root)) {
            if (rel.localName !== 'Relationship' || !isCommentRelType((0, xml_1.attr)(rel, 'Type') ?? ''))
                continue;
            drop.push(rel);
            if ((0, xml_1.attr)(rel, 'TargetMode') === 'External')
                continue;
            const target = findEntry(entries, (0, package_1.resolveTarget)(src, (0, xml_1.attr)(rel, 'Target') ?? ''));
            if (target)
                commentParts.add(target);
        }
        if (drop.length)
            relsDocs.set(name, { doc, drop });
    }
    if (commentParts.size === 0 && relsDocs.size === 0)
        return unchanged;
    // 取り除くパーツと、その rels
    const removed = new Set();
    let comments = 0;
    for (const part of commentParts) {
        comments += countComments(await readText(part));
        removed.add(part);
        const rp = findEntry(entries, (0, package_1.relsPathOf)(part));
        if (rp)
            removed.add(rp);
    }
    const changed = new Map();
    for (const [name, { doc, drop }] of relsDocs) {
        if (removed.has(name))
            continue;
        const ids = new Set(drop.map((r) => (0, xml_1.attr)(r, 'Id') ?? ''));
        for (const r of drop)
            (0, xml_1.remove)(r);
        changed.set(name, (0, xml_1.serializeXml)(doc));
        // 元のパーツが関係 ID で参照していれば（新形式のコメントの p188:commentRel など）、その参照も取り除く
        const src = (0, package_1.sourceOfRels)(name);
        if (src === '' || removed.has(src) || !entries.has(src) || !/\.xml$/i.test(src))
            continue;
        const text = await readText(src);
        if (![...ids].some((id) => id && text.includes(`"${id}"`)))
            continue;
        const srcDoc = (0, xml_1.parseXml)(text);
        if (removeRelRefs(srcDoc.documentElement, ids))
            changed.set(src, (0, xml_1.serializeXml)(srcDoc));
    }
    let ctChanged = false;
    for (const o of (0, xml_1.kids)(ctRoot)) {
        if (o.localName !== 'Override')
            continue;
        const part = findEntry(entries, ((0, xml_1.attr)(o, 'PartName') ?? '').replace(/^\/+/, ''));
        if (part && removed.has(part)) {
            (0, xml_1.remove)(o);
            ctChanged = true;
        }
    }
    if (ctChanged)
        changed.set(package_1.CT_PATH, (0, xml_1.serializeXml)(ctDoc));
    // 書き出す。手を入れないパーツは、元の圧縮方法のまま、圧縮済みのデータを使う（展開も再圧縮もしない）
    for (const [name, f] of entries) {
        if (removed.has(name)) {
            zip.remove(f.name);
            continue;
        }
        const text = changed.get(name);
        if (text !== undefined) {
            zip.file(f.name, text, { date: f.date, compression: 'DEFLATE', createFolders: false });
            continue;
        }
        f.options.compression = isStored(f) ? 'STORE' : 'DEFLATE';
    }
    removeEmptyDirs(zip, removed);
    const out = await zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
    return { bytes: out, comments, removedParts: [...removed].sort() };
}
function decodeXml(data) {
    return new TextDecoder('utf-8', { ignoreBOM: true }).decode(data);
}
/** zip の中のパーツ名を、大文字小文字を区別せずに探す（パーツ名は大文字小文字を区別しない。OPC）。 */
function findEntry(entries, name) {
    if (!name)
        return undefined;
    if (entries.has(name))
        return name;
    const lower = name.toLowerCase();
    for (const k of entries.keys())
        if (k.toLowerCase() === lower)
            return k;
    return undefined;
}
/** コメントの数（旧形式の p:cm、新形式の p188:cm と p188:reply）。読めなければ 0。 */
function countComments(text) {
    try {
        const root = (0, xml_1.parseXml)(text).documentElement;
        if (!root)
            return 0;
        return [root, ...(0, xml_1.descendants)(root)].filter((e) => e.localName === 'cm' || e.localName === 'reply').length;
    }
    catch {
        return 0;
    }
}
/** r:id などの関係 ID の属性が ids を指す要素を取り除く。拡張（p:ext）の中なら p:ext ごと、空になった p:extLst も。 */
function removeRelRefs(root, ids) {
    let removedAny = false;
    for (const el of [root, ...(0, xml_1.descendants)(root)]) {
        if (!el.parentNode)
            continue;
        const refs = Array.from({ length: el.attributes.length }, (_, i) => el.attributes.item(i)).some((a) => a.namespaceURI === xml_1.NS.r && ids.has(a.value));
        if (!refs || el === root)
            continue;
        let target = el;
        for (let n = el; n && (0, xml_1.isEl)(n); n = n.parentNode) {
            if (n.localName === 'ext') {
                target = n;
                break;
            }
        }
        const parent = target.parentNode;
        (0, xml_1.remove)(target);
        removedAny = true;
        if (parent && (0, xml_1.isEl)(parent) && parent.localName === 'extLst' && (0, xml_1.kids)(parent).length === 0)
            (0, xml_1.remove)(parent);
    }
    return removedAny;
}
/** 元が無圧縮（STORE）か。JSZip の内部の値を見る（分からなければ DEFLATE とみなす）。 */
function isStored(f) {
    const magic = f._data?.compression?.magic;
    return magic === '\x00\x00';
}
/** 取り除いたパーツのフォルダのエントリが残っていて、中が空になれば消す（ppt/comments/ など）。 */
function removeEmptyDirs(zip, removed) {
    const dirs = new Set([...removed].map((p) => (0, package_1.dirname)(p)).filter(Boolean));
    for (const d of dirs) {
        const prefix = `${d}/`;
        const dirEntry = Object.values(zip.files).find((f) => f.dir && f.name.replace(/^\/+/, '') === prefix);
        if (!dirEntry)
            continue;
        const others = Object.values(zip.files).some((f) => !f.dir && f.name.replace(/^\/+/, '').startsWith(prefix));
        if (!others)
            zip.remove(dirEntry.name);
    }
}
/** ログに出す件数の説明。 */
function strippedCommentsMessage(r) {
    if (!r.removedParts.length)
        return undefined;
    return `コメント ${r.comments} 件を取り除いてアップロードしました（Google のコメントはそのまま残ります）`;
}
