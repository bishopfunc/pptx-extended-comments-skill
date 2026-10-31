"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.PptxDoc = exports.OpRejectedError = void 0;
const arrange_1 = require("./arrange.js");
const commentOps_1 = require("./commentOps.js");
const comments_1 = require("./comments.js");
const elements_1 = require("./elements.js");
const inherit_1 = require("./inherit.js");
const commentsMerge_1 = require("./commentsMerge.js");
const edit_1 = require("./edit.js");
const contentHash_1 = require("./contentHash.js");
const integrity_1 = require("./integrity.js");
const package_1 = require("./package.js");
const parse_1 = require("./parse.js");
const slides_1 = require("./slides.js");
const slideClipboard_1 = require("./slideClipboard.js");
const tableOps_1 = require("./tableOps.js");
const types_1 = require("./types.js");
var types_2 = require("./types.js");
Object.defineProperty(exports, "OpRejectedError", { enumerable: true, get: function () { return types_2.OpRejectedError; } });
class PptxDoc {
    pkg;
    modelCache;
    constructor(pkg) {
        this.pkg = pkg;
    }
    static async load(bytes, opts = {}) {
        const pkg = await package_1.Package.load(bytes, opts);
        return new PptxDoc(pkg);
    }
    toModel(limits) {
        if (!limits && this.modelCache && this.modelCache.version === this.pkg.version)
            return structuredClone(this.modelCache.deck);
        const deck = (0, parse_1.buildDeck)(this.pkg, limits);
        deck.comments = (0, comments_1.readComments)(this.pkg);
        this.modelCache = { version: this.pkg.version, deck };
        return structuredClone(deck);
    }
    mediaMap() {
        return this.mediaFor(this.pkg.paths().filter((p) => p.startsWith('ppt/media/') || p.includes('/media/')));
    }
    mediaFor(paths) {
        const out = {};
        for (const p of paths) {
            const mime = (0, parse_1.displayableMime)(p);
            if (!mime)
                continue;
            const data = this.pkg.bytes(p);
            if (!data)
                continue;
            out[p] = { mime, data };
        }
        return out;
    }
    apply(op, ctx) {
        const txn = new package_1.Txn(this.pkg);
        try {
            // テキストの編集では、同じ図形への文字の範囲のコメントの位置を付け替える（同じ undo の単位に入れる）
            const anchors = (0, commentOps_1.captureTextAnchors)(txn, op);
            const res = (0, commentOps_1.applyCommentOp)(txn, op, ctx) ??
                (0, edit_1.applyElementOp)(txn, op, ctx?.clipboard) ??
                (0, arrange_1.applyArrangeOp)(txn, op) ??
                (0, slides_1.applySlideOp)(txn, op) ??
                (0, tableOps_1.applyTableOp)(txn, op) ??
                (0, slideClipboard_1.applySlideClipboardOp)(txn, op, ctx?.slideClipboard);
            if (!res)
                throw new types_1.OpRejectedError('未対応の編集操作です');
            (0, commentOps_1.shiftTextAnchors)(txn, anchors);
            const changes = txn.commit();
            const out = { changes };
            if (res.select)
                out.select = res.select;
            if (res.selectSlide)
                out.selectSlide = res.selectSlide;
            if (res.comment)
                out.comment = res.comment;
            return out;
        }
        catch (e) {
            txn.rollback();
            throw e;
        }
    }
    restore(changes, direction) {
        const list = direction === 'before' ? [...changes].reverse() : changes;
        for (const c of list)
            this.pkg.set(c.path, direction === 'before' ? c.before : c.after);
    }
    copyElements(elementIds) {
        const txn = new package_1.Txn(this.pkg);
        try {
            return (0, edit_1.copyElements)(txn, elementIds);
        }
        finally {
            txn.rollback();
        }
    }
    copySlides(slideIds) {
        const txn = new package_1.Txn(this.pkg);
        try {
            return (0, slideClipboard_1.copySlides)(txn, slideIds);
        }
        finally {
            txn.rollback();
        }
    }
    checkIntegrity() {
        return (0, integrity_1.checkIntegrity)(this.pkg);
    }
    hasChanges() {
        return this.pkg.hasChanges();
    }
    /**
     * ディスク上で変わったコメントを取り込む（base はエディタが最後に読み込んだ・保存した内容、disk は今のディスクの内容）。
     * 変わったパーツを返す（undo の単位にする）。スライドの中身が base と disk で同じことは、呼び出し側が確かめる。
     */
    async mergeExternalComments(base, disk) {
        const [b, d] = await Promise.all([package_1.Package.load(base), package_1.Package.load(disk)]);
        const txn = new package_1.Txn(this.pkg);
        try {
            const report = (0, commentsMerge_1.mergeComments)(txn, b, d);
            return { changes: txn.commit(), ...report };
        }
        catch (e) {
            txn.rollback();
            throw e;
        }
    }
    hasContentChanges() {
        return this.pkg.hasChangesBy(contentHash_1.sameContentPart);
    }
    markSaved() {
        this.pkg.markSaved();
    }
    save(options) {
        return this.pkg.save(options?.compression ?? 'DEFLATE');
    }
    /**
     * 図形の文字列（コメントの文字の範囲と同じ数え方。段落と改行は \n）。テキストのない要素は ''、見つからなければ undefined。
     */
    elementText(elementId) {
        const parsed = (0, elements_1.parseElementId)(elementId);
        if (!parsed)
            return undefined;
        const slide = (0, inherit_1.readPresentation)(this.pkg).slides.find((s) => s.id === parsed.slideId);
        const root = slide ? (0, comments_1.safeRoot)(this.pkg, slide.path) : undefined;
        const ref = root ? (0, elements_1.indexElements)((0, elements_1.spTreeOf)(root)).find((x) => x.key === parsed.key) : undefined;
        return ref ? (0, comments_1.shapeText)(ref.el) : undefined;
    }
    /** テスト・デバッグ用：パーツの中身。 */
    part(path) {
        return this.pkg.get(path);
    }
    /** テスト・デバッグ用：パーツのパス一覧。 */
    partPaths() {
        return this.pkg.paths();
    }
}
exports.PptxDoc = PptxDoc;
