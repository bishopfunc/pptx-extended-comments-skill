"use strict";
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.USAGE = void 0;
exports.parseArgs = parseArgs;
exports.authorOf = authorOf;
exports.normalizeId = normalizeId;
exports.addOp = addOp;
exports.runCli = runCli;
// slides-comments：pptx のレビューコメントをターミナルから読み書きする CLI（エージェント向け。設計書 7 章「レビューコメント」、ADR-0021）。
// 拡張機能と同じモジュール（PptxDoc とコメントの編集操作）で読み書きする。vscode に依存しない。
// 書き込みは一時ファイル + rename。同期（Push / Pull）の途中なら書かない。エディタの保存とは書き込みのロックで重ならないようにし、
// 読んでから書くまでにファイルが変わっていたら、読み直してやり直す。
const fs = __importStar(require("node:fs/promises"));
const path = __importStar(require("node:path"));
const index_1 = require("../pptx/index.js");
const types_1 = require("../pptx/types.js");
const commentAuthor_1 = require("../editor/commentAuthor.js");
const format_1 = require("./format.js");
const comments_1 = require("../shared/comments.js");
const fsutil_1 = require("../sync/fsutil.js");
const lock_1 = require("../sync/lock.js");
exports.USAGE = `使い方: slides-comments <コマンド> <deck.pptx> [オプション]

  outline <deck.pptx> [--slide N] [--json]   スライドごとの図形（cNvPr の ID、種類、名前、位置、文字列）、表、ノート
  list    <deck.pptx> [--slide N] [--open | --resolved] [--json]
                                    コメントの一覧（git の textconv にも使える）
  add     <deck.pptx> --slide N [--shape ID|NAME] [--text-range START:LEN | --quote "…"] --text "…" [--author 名前]
  reply   <deck.pptx> --thread ID --text "…" [--author 名前]
  edit    <deck.pptx> --thread ID [--reply ID] --text "…" [--author 名前]
  attach  <deck.pptx> --thread ID [--reply ID] --image file.png [--author 名前]
  detach  <deck.pptx> --thread ID [--reply ID] --image-id ID [--author 名前]
  resolve <deck.pptx> --thread ID
  reopen  <deck.pptx> --thread ID
  delete  <deck.pptx> --thread ID [--reply ID] [--author 名前] [--force]

  --slide は 1 始まりのスライドの番号（outline と list では「2,5」「2-4」も）。--shape は outline の #ID（cNvPr の ID）か名前。
  --text-range は図形の文字列の位置（0 始まり、段落と改行は 1 文字）と長さ。--quote は図形の文字列の中の引用（最初に見つかった所）。
  --thread・--reply の ID は list に出る GUID（{ } はなくてもよい）。
  作成者は --author、なければ環境変数 SLIDES_COMMENTS_AUTHOR、なければ「Claude」（メールアドレスは使えない）。
  自分の投稿だけを編集・削除できる。最初の投稿を消すとスレッドごと消えるので、ほかの人の返信があるスレッドは
  --force を付けないと消さない（resolve を使ってください）。
  add と reply は、スレッド（投稿）の ID を出力する。edit・resolve・reopen・delete はスレッドの ID を出力する。`;
class UsageError extends Error {
}
const VALUE_OPTS = new Set(['slide', 'shape', 'text-range', 'quote', 'text', 'author', 'thread', 'reply', 'image', 'image-id']);
const FLAG_OPTS = new Set(['json', 'help', 'force', 'open', 'resolved']);
function parseArgs(argv) {
    const [cmd, ...rest] = argv;
    if (!cmd || cmd === '--help' || cmd === '-h' || cmd === 'help')
        throw new UsageError('');
    const opts = new Map();
    let file;
    for (let i = 0; i < rest.length; i++) {
        const a = rest[i];
        if (a.startsWith('--')) {
            const name = a.slice(2);
            if (FLAG_OPTS.has(name))
                opts.set(name, true);
            else if (VALUE_OPTS.has(name)) {
                const v = rest[++i];
                if (v === undefined)
                    throw new UsageError(`--${name} に値がありません`);
                opts.set(name, v);
            }
            else
                throw new UsageError(`知らないオプションです：${a}`);
        }
        else if (file === undefined)
            file = a;
        else
            throw new UsageError(`余分な引数です：${a}`);
    }
    if (!file)
        throw new UsageError('pptx のパスがありません');
    return { cmd, file, opts };
}
function str(a, name, required = false) {
    const v = a.opts.get(name);
    if (typeof v === 'string')
        return v;
    if (required)
        throw new UsageError(`--${name} を指定してください`);
    return undefined;
}
/** --text の \n（2 文字）は改行にする（シェルで改行を渡しにくいため）。 */
function bodyText(a) {
    return str(a, 'text', true).replace(/\\n/g, '\n');
}
/** 作成者の名前（エディタの設定と同じ決まりで整える：制御文字を除き、100 文字まで、メールアドレスは使えない）。 */
function authorOf(a, io) {
    const raw = str(a, 'author') ?? io.env.SLIDES_COMMENTS_AUTHOR;
    if (raw === undefined || !raw.trim())
        return 'Claude';
    const name = (0, commentAuthor_1.cleanAuthorName)(raw.replace(/[\ud800-\udfff\ufffe\uffff]/g, ''));
    if (!name)
        throw new UsageError('作成者の名前が使えません（空、またはメールアドレスに見えます）。--author か SLIDES_COMMENTS_AUTHOR に名前を指定してください');
    return name;
}
/** スレッド・投稿の ID：{ } のない GUID も受け付ける（list の出力と同じ形にする）。旧形式の ID（legacy:…）はそのまま。 */
function normalizeId(v) {
    const t = v.trim().replace(/^\{|\}$/g, '');
    return /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(t) ? `{${t.toUpperCase()}}` : v.trim();
}
/** --slide・--shape・--text-range・--quote から、addComment の op を作る。 */
function addOp(doc, a, io) {
    const deck = doc.toModel();
    const n = Number(str(a, 'slide', true));
    const slide = Number.isInteger(n) ? deck.slides[n - 1] : undefined;
    if (!slide)
        throw new UsageError(`スライド ${str(a, 'slide')} がありません（1〜${deck.slides.length}）`);
    const shape = str(a, 'shape');
    const range = str(a, 'text-range');
    const quote = str(a, 'quote');
    if ((range || quote) && !shape)
        throw new UsageError('--text-range と --quote には --shape も指定してください');
    if (range && quote)
        throw new UsageError('--text-range と --quote は片方だけを指定してください');
    let anchor = { kind: 'slide' };
    if (shape) {
        const el = (0, format_1.outlineOf)(doc).slides[n - 1].elements.find((e) => e.cNvPrId === shape.replace(/^#/, '')) ??
            (0, format_1.outlineOf)(doc).slides[n - 1].elements.find((e) => e.name === shape);
        if (!el)
            throw new UsageError(`スライド ${n} に図形 ${shape} がありません（outline で ID と名前を確かめてください）`);
        anchor = { kind: 'element', elementId: el.elementId };
        if (range || quote) {
            const text = doc.elementText(el.elementId) ?? '';
            let start;
            let length;
            if (range) {
                const m = /^(\d+):(\d+)$/.exec(range);
                if (!m)
                    throw new UsageError('--text-range は START:LEN の形で指定してください（例 0:5）');
                start = Number(m[1]);
                length = Number(m[2]);
            }
            else {
                const q = quote.replace(/\\n/g, '\n');
                start = text.indexOf(q);
                if (!q || start < 0)
                    throw new UsageError(`図形 ${shape} の文字列に「${quote}」がありません`);
                if (text.indexOf(q, start + 1) >= 0)
                    io.err(`注意：「${quote}」は複数あります。最初の位置（${start}）に付けます`);
                length = q.length;
            }
            anchor = { kind: 'text', elementId: el.elementId, start, length };
        }
    }
    return { type: 'addComment', slideId: slide.id, anchor, text: bodyText(a) };
}
async function opFor(doc, a, io, author) {
    const thread = () => normalizeId(str(a, 'thread', true));
    const replyId = () => {
        const r = str(a, 'reply');
        return r === undefined ? undefined : normalizeId(r);
    };
    switch (a.cmd) {
        case 'add':
            return addOp(doc, a, io);
        case 'reply':
            return { type: 'replyComment', threadId: thread(), text: bodyText(a) };
        case 'edit':
            return { type: 'editComment', threadId: thread(), postId: replyId() ?? thread(), text: bodyText(a) };
        case 'attach': {
            const image = path.resolve(str(a, 'image', true));
            if ((await fs.stat(image)).size > comments_1.COMMENT_IMAGE_MAX_BYTES)
                throw new UsageError('添付画像は2MB以下にしてください');
            const bytes = await fs.readFile(image);
            const mime = bytes[0] === 0x89 ? 'image/png' : 'image/jpeg';
            return { type: 'attachCommentImage', threadId: thread(), postId: replyId() ?? thread(),
                name: path.basename(image), dataUrl: `data:${mime};base64,${bytes.toString('base64')}` };
        }
        case 'detach':
            return { type: 'removeCommentImage', threadId: thread(), postId: replyId() ?? thread(), imageId: normalizeId(str(a, 'image-id', true)) };
        case 'resolve':
            return { type: 'resolveComment', threadId: thread() };
        case 'reopen':
            return { type: 'reopenComment', threadId: thread() };
        case 'delete': {
            const reply = replyId();
            if (!reply && !a.opts.has('force')) {
                // 最初の投稿を消すとスレッドごと（返信も）消える。ほかの人の返信があれば、--force がなければ消さない
                const t = doc.toModel().comments?.threads.find((x) => x.id.toLowerCase() === thread().toLowerCase());
                const others = t?.replies.filter((r) => r.author !== author) ?? [];
                if (others.length) {
                    throw new types_1.OpRejectedError(`ほかの人の返信（${others.length} 件）があるため、スレッドを削除しませんでした。resolve で解決済みにしてください（返信ごと消すには --force）`);
                }
            }
            return { type: 'deleteComment', threadId: thread(), ...(reply ? { postId: reply } : {}) };
        }
        default:
            throw new UsageError(`知らないコマンドです：${a.cmd}`);
    }
}
/** 書き込み：読む → 適用 → 一時ファイル + rename。読んでから書くまでに変わっていれば、やり直す。 */
async function writeOp(file, a, io) {
    if (await (0, lock_1.isLocked)(file).catch(() => false)) {
        io.err('このファイルは同期（Push / Pull）の途中です。終わってからやり直してください');
        return 3;
    }
    const lock = await (0, lock_1.waitForWriteLock)(file, 5000).catch(() => undefined);
    if (!lock) {
        io.err('エディタ（または同期）がこのファイルを書き込んでいます。少し待ってからやり直してください');
        return 3;
    }
    const syncing = async () => {
        if (!(await (0, lock_1.isLocked)(file).catch(() => false)))
            return false;
        io.err('このファイルは同期（Push / Pull）の途中です。終わってからやり直してください');
        return true;
    };
    try {
        // 書き込みのロックを待つ間に、同期が始まったかもしれない
        if (await syncing())
            return 3;
        for (let attempt = 0; attempt < 3; attempt++) {
            const bytes = new Uint8Array(await fs.readFile(file));
            const doc = await index_1.PptxDoc.load(bytes);
            const author = authorOf(a, io);
            const op = await opFor(doc, a, io, author);
            const res = doc.apply(op, { comment: { author } });
            const issues = doc.checkIntegrity();
            if (issues.length)
                io.err(`注意：整合性の検査で ${issues.length} 件の問題：${issues.map((i) => `${i.part}: ${i.message}`).join(' / ')}`);
            const out = await doc.save();
            const now = await fs.readFile(file).catch(() => undefined);
            if (!now || (0, fsutil_1.sha256)(now) !== (0, fsutil_1.sha256)(bytes))
                continue; // 読んでいる間に変わった
            if (await syncing())
                return 3;
            await (0, fsutil_1.writeFileAtomic)(file, out);
            if (a.cmd === 'add')
                io.out(res.comment.threadId);
            else if (a.cmd === 'reply')
                io.out(res.comment.postId);
            else if (res.comment)
                io.out(res.comment.threadId);
            else if (op.type === 'deleteComment')
                io.out(op.threadId);
            return 0;
        }
        io.err('書き込む間にファイルが何度も変わったため、中止しました');
        return 3;
    }
    finally {
        await lock.release().catch(() => undefined);
    }
}
async function runCli(argv, io) {
    let a;
    try {
        a = parseArgs(argv);
    }
    catch (e) {
        if (e instanceof UsageError && e.message)
            io.err(e.message);
        io.err(exports.USAGE);
        return e instanceof UsageError && !e.message ? 0 : 1;
    }
    // シンボリックリンクはたどる（リンクそのものを rename で置き換えると、元の pptx に書かれない。ロックも元の pptx で取る）
    const given = path.resolve(a.file);
    const file = await fs.realpath(given).catch(() => given);
    try {
        if (a.cmd === 'outline' || a.cmd === 'list') {
            const doc = await index_1.PptxDoc.load(new Uint8Array(await fs.readFile(file)));
            const json = a.opts.has('json');
            let slides;
            try {
                slides = (0, format_1.parseSlideFilter)(str(a, 'slide'), doc.toModel().slides.length);
            }
            catch (e) {
                throw new UsageError(e.message);
            }
            if (a.cmd === 'outline') {
                const o = (0, format_1.outlineOf)(doc);
                if (slides)
                    o.slides = o.slides.filter((x) => slides.has(x.number));
                io.out(json ? JSON.stringify(o, null, 2) : (0, format_1.formatOutline)(o));
            }
            else {
                if (a.opts.has('open') && a.opts.has('resolved'))
                    throw new UsageError('--open と --resolved は片方だけを指定してください');
                const filter = { ...(slides ? { slides } : {}), ...(a.opts.has('open') ? { status: 'active' } : a.opts.has('resolved') ? { status: 'resolved' } : {}) };
                if (json) {
                    const c = doc.toModel().comments;
                    const num = new Map(doc.toModel().slides.map((x, i) => [x.id, i + 1]));
                    const threads = (0, format_1.filterThreadList)(doc, c?.threads ?? [], filter).map((t) => ({ ...t, slideNumber: num.get(t.slideId) }));
                    io.out(JSON.stringify({ ...c, threads }, null, 2));
                }
                else
                    io.out((0, format_1.formatThreads)(doc, filter));
            }
            return 0;
        }
        const st = await fs.stat(file);
        if (st.nlink > 1) {
            // rename で書くと、ほかのハードリンクとは別のファイルになる（エディタが開いている方に書かれない）
            throw new UsageError('このファイルはハードリンクされているため、書き込めません（コピーしてから使ってください）');
        }
        return await writeOp(file, a, io);
    }
    catch (e) {
        if (e instanceof UsageError) {
            io.err(e.message);
            return 1;
        }
        if (e instanceof types_1.OpRejectedError) {
            io.err(e.message);
            if (/見つかりません/.test(e.message))
                io.err('list で今の ID を確かめてください（{ } を含む GUID）');
            return 2;
        }
        io.err(`失敗しました：${e instanceof Error ? e.message : String(e)}`);
        return 1;
    }
}
