"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.isCommentPartPath = isCommentPartPath;
exports.normalizeForContent = normalizeForContent;
exports.sameContentPart = sameContentPart;
exports.contentSha256 = contentSha256;
// コメントを除いた pptx の中身のハッシュ（設計書 10 章「ローカルの変更とコメント」、ADR-0021）。
// コメントは Google スライドと同期しないので、コメントだけの変更は「ローカルの変更」に数えない。
// 除くもの：コメントと作成者のパーツ（新形式・旧形式）、それを指す関係と Content Type、スライドの p188:commentRel、
// コメントのアンカーのために付けた作成 ID（スライドの p14:creationId、図形の a16:creationId）。
// rels・[Content_Types].xml・スライドの XML は、どちらの側でも同じ形に書き直してから比べる（この拡張機能が書き直した XML と、
// 元のままの XML を同じ中身として扱うため）。vscode に依存しない。
const node_crypto_1 = require("node:crypto");
const creationIds_1 = require("./creationIds.js");
const package_1 = require("./package.js");
const stripComments_1 = require("./stripComments.js");
const xml_1 = require("./xml.js");
const zipLimits_1 = require("./zipLimits.js");
const COMMENT_CONTENT_TYPES = new Set([
    'application/vnd.openxmlformats-officedocument.presentationml.comments+xml',
    'application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml',
    'application/vnd.ms-powerpoint.comments+xml',
    'application/vnd.ms-powerpoint.authors+xml',
]);
const COMMENT_PATH = /^ppt\/(comments\/[^/]+|commentAuthors\.xml|authors\.xml)$/i;
const COMMENT_RELS_PATH = /^ppt\/comments\/_rels\/[^/]+\.rels$/i;
const SLIDE_PATH = /^ppt\/slides\/[^/]+\.xml$/i;
const DROP_EXT_URIS = new Set([creationIds_1.URI_COMMENT_REL, creationIds_1.URI_SLIDE_CREATION_ID, creationIds_1.URI_SHAPE_CREATION_ID]);
/** コメントのパーツ（中身の比較から除く）か。 */
function isCommentPartPath(path) {
    return COMMENT_PATH.test(path) || COMMENT_RELS_PATH.test(path);
}
function rewrite(text, fn) {
    try {
        const doc = (0, xml_1.parseXml)(text);
        const root = doc.documentElement;
        if (!root)
            return text;
        fn(root);
        return (0, xml_1.serializeXml)(doc);
    }
    catch {
        return text; // 読めない XML はそのまま比べる
    }
}
/**
 * 中身の比較に使う形。コメントのパーツなら null（比べない）。
 */
function normalizeForContent(path, data) {
    if (data === null || isCommentPartPath(path))
        return null;
    if (typeof data !== 'string')
        return data;
    if (path.endsWith('.rels')) {
        return rewrite(data, (root) => {
            for (const r of (0, xml_1.kids)(root))
                if (r.localName === 'Relationship' && (0, stripComments_1.isCommentRelType)((0, xml_1.attr)(r, 'Type') ?? ''))
                    (0, xml_1.remove)(r);
        });
    }
    if (path === package_1.CT_PATH) {
        return rewrite(data, (root) => {
            for (const o of (0, xml_1.kids)(root)) {
                if (o.localName !== 'Override')
                    continue;
                const name = ((0, xml_1.attr)(o, 'PartName') ?? '').replace(/^\/+/, '');
                if (COMMENT_CONTENT_TYPES.has((0, xml_1.attr)(o, 'ContentType') ?? '') || isCommentPartPath(name))
                    (0, xml_1.remove)(o);
            }
        });
    }
    if (SLIDE_PATH.test(path)) {
        return rewrite(data, (root) => {
            for (const ext of [root, ...(0, xml_1.descendants)(root)].filter((e) => e.localName === 'ext' && DROP_EXT_URIS.has((0, xml_1.attr)(e, 'uri') ?? ''))) {
                const lst = ext.parentNode;
                (0, xml_1.remove)(ext);
                if (lst && (0, xml_1.isEl)(lst) && lst.localName === 'extLst' && (0, xml_1.kids)(lst).length === 0)
                    (0, xml_1.remove)(lst);
            }
        });
    }
    return data;
}
/** 2 つのパーツが、コメントを除いて同じ中身か。 */
function sameContentPart(path, a, b) {
    const x = normalizeForContent(path, a);
    const y = normalizeForContent(path, b);
    if (x === null || y === null)
        return x === y;
    if (typeof x === 'string' || typeof y === 'string')
        return x === y;
    return x.length === y.length && x.every((v, i) => v === y[i]);
}
const utf8 = new TextDecoder('utf-8', { ignoreBOM: true });
/**
 * コメントを除いた中身の SHA-256（16 進）。zip として開けなければ undefined（中身では比べられない）。
 * 画像などのバイナリのパーツは展開せず、zip に書かれた CRC-32 と大きさで比べる（同じ中身なら圧縮し直しても変わらない）。
 */
async function contentSha256(bytes, limits = zipLimits_1.DEFAULT_ZIP_LIMITS) {
    let zip;
    try {
        zip = await (0, zipLimits_1.openZip)(bytes, limits);
    }
    catch {
        return undefined;
    }
    const h = (0, node_crypto_1.createHash)('sha256');
    const budget = { total: 0 };
    const entries = (0, zipLimits_1.zipEntries)(zip)
        .map((f) => ({ f, name: f.name.replace(/^\/+/, '') }))
        .sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));
    try {
        for (const { f, name } of entries) {
            if (isCommentPartPath(name))
                continue;
            if (!(0, package_1.isXmlPath)(name)) {
                const d = f._data;
                if (d && typeof d.crc32 === 'number' && typeof d.uncompressedSize === 'number') {
                    h.update(`${name}\0bin:${d.crc32}:${d.uncompressedSize}\0`);
                    continue;
                }
                h.update(`${name}\0`).update(await (0, zipLimits_1.readZipEntry)(f, budget, limits)).update('\0');
                continue;
            }
            const norm = normalizeForContent(name, utf8.decode(await (0, zipLimits_1.readZipEntry)(f, budget, limits)));
            if (norm === null)
                continue;
            h.update(`${name}\0`).update(norm).update('\0');
        }
    }
    catch {
        return undefined;
    }
    return h.digest('hex');
}
