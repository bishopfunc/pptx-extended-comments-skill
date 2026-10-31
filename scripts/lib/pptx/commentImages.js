"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.COMMENT_IMAGES_URI = void 0;
exports.imageExtension = imageExtension;
exports.imageKey = imageKey;
exports.readCommentImages = readCommentImages;
exports.writeCommentImages = writeCommentImages;
// コメントのextLstに自己完結した添付画像を保存する。この拡張機能独自の拡張。
// 別パーツへの参照を持たないため、コメントの取り込み・Pull・除去にも追従する。
const comments_1 = require("../shared/comments.js");
const commentImages_1 = require("../shared/commentImages.js");
const xml_1 = require("./xml.js");
exports.COMMENT_IMAGES_URI = 'urn:google-slides-sync:comment-images:v1';
function imageExtension(post) {
    return (0, xml_1.kidsNamed)((0, xml_1.kid)(post, 'p188:extLst'), 'p188:ext').find(e => (0, xml_1.attr)(e, 'uri') === exports.COMMENT_IMAGES_URI);
}
function imageKey(post) {
    const ext = imageExtension(post);
    return ext ? (0, xml_1.serializeXml)(ext) : '';
}
function readCommentImages(post) {
    const out = [];
    const seen = new Set();
    for (const el of (0, xml_1.kidsNamed)((0, xml_1.kid)(imageExtension(post), 'gssc:images'), 'gssc:image').slice(0, comments_1.COMMENT_IMAGES_MAX)) {
        const id = (0, xml_1.attr)(el, 'id');
        const dataUrl = el.textContent ?? '';
        if (!id || seen.has(id) || !(0, commentImages_1.validCommentImage)(dataUrl))
            continue;
        seen.add(id);
        out.push({ id, name: (0, commentImages_1.commentImageName)((0, xml_1.attr)(el, 'name') ?? ''), dataUrl });
    }
    return out;
}
function writeCommentImages(post, images) {
    const old = imageExtension(post);
    if (old)
        (0, xml_1.remove)(old);
    if (!images.length) {
        const list = (0, xml_1.kid)(post, 'p188:extLst');
        if (list && !list.firstChild)
            (0, xml_1.remove)(list);
        return;
    }
    const doc = post.ownerDocument;
    const list = (0, xml_1.ensureKid)(post, 'p188:extLst');
    const ext = (0, xml_1.create)(doc, 'p188:ext', { uri: exports.COMMENT_IMAGES_URI });
    const content = (0, xml_1.create)(doc, 'gssc:images');
    for (const image of images) {
        const el = (0, xml_1.create)(doc, 'gssc:image', { id: image.id, name: image.name });
        el.appendChild(doc.createTextNode(image.dataUrl));
        content.appendChild(el);
    }
    ext.appendChild(content);
    list.appendChild(ext);
}
