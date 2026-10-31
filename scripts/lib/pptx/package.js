"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.create = exports.Txn = exports.Package = exports.CT = exports.REL = exports.CT_PATH = void 0;
exports.isXmlPath = isXmlPath;
exports.dirname = dirname;
exports.basename = basename;
exports.extname = extname;
exports.resolveTarget = resolveTarget;
exports.relativeTarget = relativeTarget;
exports.relsPathOf = relsPathOf;
exports.sourceOfRels = sourceOfRels;
exports.sameData = sameData;
exports.readRels = readRels;
exports.relById = relById;
exports.relsOfType = relsOfType;
exports.contentTypeOf = contentTypeOf;
exports.rootEl = rootEl;
// zip の読み書き、パーツ管理、[Content_Types].xml、rels（設計書 7 章「読み込み」）。
const jszip_1 = __importDefault(require("../../vendor/jszip.js"));
const zipLimits_1 = require("./zipLimits.js");
const xml_1 = require("./xml.js");
Object.defineProperty(exports, "create", { enumerable: true, get: function () { return xml_1.create; } });
exports.CT_PATH = '[Content_Types].xml';
exports.REL = {
    officeDocument: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
    slide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide',
    slideLayout: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
    slideMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
    theme: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme',
    notesSlide: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
    notesMaster: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster',
    image: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
    hyperlink: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink',
    comments: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments',
    chart: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
    diagramData: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData',
    oleObject: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject',
    package: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/package',
};
exports.CT = {
    slide: 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml',
    notesSlide: 'application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml',
    notesMaster: 'application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml',
    theme: 'application/vnd.openxmlformats-officedocument.theme+xml',
    rels: 'application/vnd.openxmlformats-package.relationships+xml',
};
function isXmlPath(p) {
    const lower = p.toLowerCase();
    return lower.endsWith('.xml') || lower.endsWith('.rels') || lower.endsWith('.vml');
}
function dirname(p) {
    const i = p.lastIndexOf('/');
    return i < 0 ? '' : p.slice(0, i);
}
function basename(p) {
    const i = p.lastIndexOf('/');
    return i < 0 ? p : p.slice(i + 1);
}
function extname(p) {
    const b = basename(p);
    const i = b.lastIndexOf('.');
    return i < 0 ? '' : b.slice(i + 1).toLowerCase();
}
/** base パーツから見た相対ターゲットを絶対パスにする。 */
function resolveTarget(basePart, target) {
    let t = target;
    try {
        t = decodeURI(target);
    }
    catch {
        // そのまま
    }
    if (t.startsWith('/'))
        return normalize(t.slice(1));
    const dir = dirname(basePart);
    return normalize(dir ? `${dir}/${t}` : t);
}
function normalize(p) {
    const out = [];
    for (const seg of p.split('/')) {
        if (seg === '' || seg === '.')
            continue;
        if (seg === '..')
            out.pop();
        else
            out.push(seg);
    }
    return out.join('/');
}
/** from パーツから to パーツへの相対パス。 */
function relativeTarget(fromPart, toPart) {
    const from = dirname(fromPart).split('/').filter(Boolean);
    const to = toPart.split('/').filter(Boolean);
    let i = 0;
    while (i < from.length && i < to.length - 1 && from[i] === to[i])
        i++;
    const ups = from.length - i;
    return [...Array(ups).fill('..'), ...to.slice(i)].join('/');
}
/** パーツ p の rels のパス。 */
function relsPathOf(p) {
    if (p === '')
        return '_rels/.rels';
    return `${dirname(p) ? dirname(p) + '/' : ''}_rels/${basename(p)}.rels`;
}
/** rels のパスから元のパーツのパス。 */
function sourceOfRels(relsPath) {
    const dir = dirname(dirname(relsPath));
    const name = basename(relsPath).replace(/\.rels$/, '');
    if (name === '')
        return '';
    return dir ? `${dir}/${name}` : name;
}
const encoder = new TextEncoder();
const utf8 = new TextDecoder('utf-8', { ignoreBOM: true });
class Package {
    parts = new Map();
    order = [];
    domCache = new Map();
    saved = new Map();
    /** パーツを書き戻すたびに増える（モデルのキャッシュ用）。 */
    version = 0;
    /** 展開は上限を確かめながら 1 つずつ行う（大きすぎる・壊れた zip で落ちないように。ADR-0019）。 */
    static async load(bytes, opts = {}) {
        const limits = opts.zipLimits ?? zipLimits_1.DEFAULT_ZIP_LIMITS;
        const zip = await (0, zipLimits_1.openZip)(bytes, limits);
        const pkg = new Package();
        const budget = { total: 0 };
        for (const f of (0, zipLimits_1.zipEntries)(zip)) {
            const name = f.name.replace(/^\/+/, '');
            const data = await (0, zipLimits_1.readZipEntry)(f, budget, limits);
            // BOM は残す（JSZip の async('string') と同じ）
            if (isXmlPath(name))
                pkg.parts.set(name, utf8.decode(data));
            else
                pkg.parts.set(name, data);
            pkg.order.push(name);
        }
        if (!pkg.parts.has(exports.CT_PATH))
            throw new Error('[Content_Types].xml がありません。pptx ではないファイルです');
        pkg.markSaved();
        return pkg;
    }
    has(path) {
        return this.parts.has(path);
    }
    paths() {
        return [...this.parts.keys()];
    }
    get(path) {
        return this.parts.get(path) ?? null;
    }
    text(path) {
        const v = this.parts.get(path);
        if (v === undefined)
            return undefined;
        if (typeof v === 'string')
            return v;
        return new TextDecoder().decode(v);
    }
    bytes(path) {
        const v = this.parts.get(path);
        if (v === undefined)
            return undefined;
        return typeof v === 'string' ? encoder.encode(v) : v;
    }
    dom(path) {
        let d = this.domCache.get(path);
        if (!d) {
            const t = this.text(path);
            if (t === undefined)
                throw new Error(`part not found: ${path}`);
            d = (0, xml_1.parseXml)(t);
            this.domCache.set(path, d);
        }
        return d;
    }
    /** パーツを書き換える（null で削除）。DOM のキャッシュを捨てる。 */
    set(path, data) {
        this.domCache.delete(path);
        this.version++;
        if (data === null) {
            this.parts.delete(path);
            return;
        }
        if (!this.parts.has(path))
            this.order.push(path);
        this.parts.set(path, data);
    }
    markSaved() {
        this.saved = new Map(this.parts);
        this.sameCache = new WeakMap();
    }
    /** hasChangesBy の比べ方ごとの、パーツごとの前回の結果（保存した内容と今の内容が前回と同じなら比べ直さない）。 */
    sameCache = new WeakMap();
    hasChanges() {
        if (this.saved.size !== this.parts.size)
            return true;
        for (const [k, v] of this.parts) {
            const s = this.saved.get(k);
            if (s === undefined)
                return true;
            if (s === v)
                continue;
            if (!sameData(s, v))
                return true;
        }
        return false;
    }
    /** 前回の load / markSaved から、same で比べて違うパーツがあるか（コメントだけの変更を除くときに使う）。 */
    hasChangesBy(same) {
        // 同期の状態は編集のたびに確かめるので、変わっていないパーツの比べ直し（XML の読み直し）はしない
        let cache = this.sameCache.get(same);
        if (!cache)
            this.sameCache.set(same, (cache = new Map()));
        const check = (k, s, v) => {
            const c = cache.get(k);
            if (c && c.s === s && c.v === v)
                return c.same;
            const r = same(k, s, v);
            cache.set(k, { s, v, same: r });
            return r;
        };
        for (const [k, v] of this.parts) {
            const s = this.saved.get(k);
            if (s === v)
                continue;
            if (!check(k, s ?? null, v))
                return true;
        }
        for (const [k, s] of this.saved)
            if (!this.parts.has(k) && !check(k, s, null))
                return true;
        return false;
    }
    async save(compression = 'DEFLATE') {
        const zip = new jszip_1.default();
        const names = [exports.CT_PATH, ...this.order.filter((p) => p !== exports.CT_PATH && this.parts.has(p))];
        const seen = new Set();
        for (const name of names) {
            if (seen.has(name))
                continue;
            seen.add(name);
            zip.file(name, this.parts.get(name), { createFolders: false });
        }
        return zip.generateAsync({
            type: 'uint8array',
            compression,
            compressionOptions: compression === 'DEFLATE' ? { level: 6 } : undefined,
        });
    }
}
exports.Package = Package;
function sameData(a, b) {
    if (a === b)
        return true;
    if (a === null || b === null)
        return false;
    if (typeof a === 'string' || typeof b === 'string')
        return typeof a === typeof b && a === b;
    if (a.length !== b.length)
        return false;
    for (let i = 0; i < a.length; i++)
        if (a[i] !== b[i])
            return false;
    return true;
}
/**
 * 1 回の apply の間の変更をまとめる。DOM は Package のキャッシュとは別に作り直し、
 * commit で文字列に戻して書き込む。失敗したら rollback で元に戻す。
 */
class Txn {
    pkg;
    originals = new Map();
    doms = new Map();
    constructor(pkg) {
        this.pkg = pkg;
    }
    has(path) {
        return this.doms.has(path) || this.pkg.has(path);
    }
    paths() {
        const s = new Set(this.pkg.paths());
        for (const k of this.doms.keys())
            s.add(k);
        return [...s];
    }
    text(path) {
        const d = this.doms.get(path);
        if (d)
            return (0, xml_1.serializeXml)(d);
        return this.pkg.text(path);
    }
    bytes(path) {
        const d = this.doms.get(path);
        if (d)
            return encoder.encode((0, xml_1.serializeXml)(d));
        return this.pkg.bytes(path);
    }
    /** 読み取り用の DOM。編集中のパーツなら編集中の DOM を返す。 */
    dom(path) {
        return this.doms.get(path) ?? this.pkg.dom(path);
    }
    /** 書き換え用の DOM（このトランザクション専用のコピー）。 */
    edit(path) {
        let d = this.doms.get(path);
        if (!d) {
            const t = this.pkg.text(path);
            if (t === undefined)
                throw new Error(`part not found: ${path}`);
            d = (0, xml_1.parseXml)(t);
            this.doms.set(path, d);
        }
        return d;
    }
    /** 新しい XML パーツを DOM で作る。 */
    createXml(path, xml) {
        this.put(path, xml);
        return this.edit(path);
    }
    put(path, data) {
        if (!this.originals.has(path))
            this.originals.set(path, this.pkg.get(path));
        this.doms.delete(path);
        this.pkg.set(path, data);
    }
    delete(path) {
        if (this.has(path))
            this.put(path, null);
    }
    commit() {
        for (const [path, d] of this.doms) {
            const s = (0, xml_1.serializeXml)(d);
            const cur = this.pkg.get(path);
            if (cur !== s) {
                if (!this.originals.has(path))
                    this.originals.set(path, cur);
                this.pkg.set(path, s);
            }
        }
        this.doms.clear();
        const changes = [];
        for (const [path, before] of this.originals) {
            const after = this.pkg.get(path);
            if (sameData(before, after)) {
                // 変わっていなければ元の値（同じ文字列）に戻しておく
                if (before !== after)
                    this.pkg.set(path, before);
                continue;
            }
            changes.push({ path, before, after });
        }
        return changes;
    }
    rollback() {
        this.doms.clear();
        for (const [path, before] of this.originals)
            this.pkg.set(path, before);
        this.originals.clear();
    }
    // ---- rels ----
    /** パーツ p の rels を DOM で編集する（なければ作る）。 */
    editRels(part) {
        const rp = relsPathOf(part);
        if (!this.has(rp)) {
            this.put(rp, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="${xml_1.NS.rel}"></Relationships>`);
        }
        return this.edit(rp);
    }
    addRel(part, type, target, external = false) {
        const doc = this.editRels(part);
        const root = doc.documentElement;
        const used = new Set((0, xml_1.kids)(root).map((r) => (0, xml_1.attr)(r, 'Id')));
        let n = 1;
        while (used.has(`rId${n}`))
            n++;
        const id = `rId${n}`;
        const el = doc.createElementNS(xml_1.NS.rel, 'Relationship');
        el.setAttribute('Id', id);
        el.setAttribute('Type', type);
        el.setAttribute('Target', external ? target : relativeTarget(part, target));
        if (external)
            el.setAttribute('TargetMode', 'External');
        root.appendChild(el);
        return id;
    }
    removeRel(part, id) {
        const rp = relsPathOf(part);
        if (!this.has(rp))
            return;
        const doc = this.edit(rp);
        for (const r of (0, xml_1.kids)(doc.documentElement))
            if ((0, xml_1.attr)(r, 'Id') === id)
                r.parentNode.removeChild(r);
    }
    // ---- content types ----
    addOverride(path, contentType) {
        const doc = this.edit(exports.CT_PATH);
        const root = doc.documentElement;
        const name = `/${path}`;
        for (const o of (0, xml_1.kids)(root))
            if (o.localName === 'Override' && (0, xml_1.attr)(o, 'PartName') === name)
                o.parentNode.removeChild(o);
        const el = doc.createElementNS(xml_1.NS.ct, 'Override');
        el.setAttribute('PartName', name);
        el.setAttribute('ContentType', contentType);
        root.appendChild(el);
    }
    removeOverride(path) {
        const doc = this.edit(exports.CT_PATH);
        const root = doc.documentElement;
        const name = `/${path}`.toLowerCase();
        for (const o of (0, xml_1.kids)(root)) {
            if (o.localName === 'Override' && ((0, xml_1.attr)(o, 'PartName') ?? '').toLowerCase() === name)
                o.parentNode.removeChild(o);
        }
    }
    /** 大文字小文字を区別せずに同じ拡張子の Default がなければ追加する。 */
    ensureDefault(ext, contentType) {
        const cur = this.dom(exports.CT_PATH).documentElement;
        const exists = (0, xml_1.kids)(cur).some((d) => d.localName === 'Default' && ((0, xml_1.attr)(d, 'Extension') ?? '').toLowerCase() === ext.toLowerCase());
        if (exists)
            return;
        const doc = this.edit(exports.CT_PATH);
        const root = doc.documentElement;
        const el = doc.createElementNS(xml_1.NS.ct, 'Default');
        el.setAttribute('Extension', ext);
        el.setAttribute('ContentType', contentType);
        const firstOverride = (0, xml_1.kids)(root).find((k) => k.localName === 'Override');
        if (firstOverride)
            root.insertBefore(el, firstOverride);
        else
            root.appendChild(el);
    }
}
exports.Txn = Txn;
// ---- 読み取り用ヘルパー ----
function readRels(r, part) {
    const rp = relsPathOf(part);
    if (!r.has(rp))
        return [];
    const root = r.dom(rp).documentElement;
    if (!root)
        return [];
    return (0, xml_1.kids)(root)
        .filter((k) => k.localName === 'Relationship')
        .map((k) => {
        const external = (0, xml_1.attr)(k, 'TargetMode') === 'External';
        const raw = (0, xml_1.attr)(k, 'Target') ?? '';
        return {
            id: (0, xml_1.attr)(k, 'Id') ?? '',
            type: (0, xml_1.attr)(k, 'Type') ?? '',
            rawTarget: raw,
            target: external ? raw : resolveTarget(part, raw),
            external,
        };
    });
}
function relById(r, part, id) {
    return readRels(r, part).find((x) => x.id === id);
}
function relsOfType(r, part, type) {
    return readRels(r, part).filter((x) => x.type === type);
}
function contentTypeOf(r, path) {
    const root = r.dom(exports.CT_PATH).documentElement;
    const name = `/${path}`.toLowerCase();
    for (const o of (0, xml_1.kids)(root)) {
        if (o.localName === 'Override' && ((0, xml_1.attr)(o, 'PartName') ?? '').toLowerCase() === name)
            return (0, xml_1.attr)(o, 'ContentType');
    }
    const ext = extname(path);
    for (const d of (0, xml_1.kids)(root)) {
        if (d.localName === 'Default' && ((0, xml_1.attr)(d, 'Extension') ?? '').toLowerCase() === ext)
            return (0, xml_1.attr)(d, 'ContentType');
    }
    return undefined;
}
function rootEl(doc) {
    return doc.documentElement;
}
