"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.validCommentImage = validCommentImage;
exports.commentImageName = commentImageName;
// コメントの添付画像の検査（PNG/JPEG の data URL の形式、署名、サイズ、画素数）と、画像の名前の整え方。拡張機能のホストと Webview で共有する。
const comments_1 = require("./comments.js");
/** PNG/JPEGだけを許可する。MIME、署名、サイズ、画素数をHostとWebviewで同じように検査。 */
function validCommentImage(dataUrl) {
    if (typeof dataUrl !== 'string' || dataUrl.length > Math.ceil(comments_1.COMMENT_IMAGE_MAX_BYTES / 3) * 4 + 32)
        return false;
    const m = /^data:image\/(png|jpeg);base64,([A-Za-z0-9+/]+={0,2})$/.exec(dataUrl);
    if (!m || m[2].length % 4 !== 0)
        return false;
    let bytes;
    try {
        bytes = atob(m[2]);
    }
    catch {
        return false;
    }
    if (!bytes.length || bytes.length > comments_1.COMMENT_IMAGE_MAX_BYTES)
        return false;
    const u8 = (i) => bytes.charCodeAt(i);
    const u16 = (i) => u8(i) * 256 + u8(i + 1);
    const u32 = (i) => u16(i) * 65536 + u16(i + 2);
    const dimensions = (w, h) => w > 0 && h > 0 && w <= 8192 && h <= 8192 && w * h <= 16_000_000;
    if (m[1] === 'png') {
        return bytes.length >= 33 && bytes.slice(0, 8) === '\x89PNG\r\n\x1a\n' &&
            u32(8) === 13 && bytes.slice(12, 16) === 'IHDR' && dimensions(u32(16), u32(20));
    }
    if (u16(0) !== 0xffd8)
        return false;
    // JPEGのSOFから幅・高さを読む。圧縮データを展開しない。
    let i = 2;
    while (i + 4 < bytes.length) {
        if (u8(i++) !== 0xff)
            return false;
        while (i < bytes.length && u8(i) === 0xff)
            i++;
        const marker = u8(i++);
        if (marker === 0xda || marker === 0xd9)
            return false;
        if (marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7))
            continue;
        const size = u16(i);
        if (size < 2 || i + size > bytes.length)
            return false;
        if ([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf].includes(marker)) {
            return size >= 8 && dimensions(u16(i + 5), u16(i + 3));
        }
        i += size;
    }
    return false;
}
function commentImageName(name) {
    return name.replace(/[\u0000-\u001f\u007f\ufffe\uffff]|\p{Surrogate}/gu, '').split(/[\\/]/).pop().trim().slice(0, 160) || '画像';
}
