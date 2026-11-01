"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.outlineOf = outlineOf;
exports.terminalSafe = terminalSafe;
exports.parseSlideFilter = parseSlideFilter;
exports.formatOutline = formatOutline;
exports.filterThreadList = filterThreadList;
exports.formatThreads = formatThreads;
const comments_1 = require("../shared/comments.js");
const model_1 = require("../shared/model.js");
function kindOf(e) {
    switch (e.kind) {
        case 'shape':
            return e.isConnector ? 'connector' : e.isTextBox ? 'textBox' : 'shape';
        case 'picture':
        case 'group':
        case 'table':
            return e.kind;
        default:
            return 'other';
    }
}
/**
 * 文字の大きさと、枠からはみ出すおそれ（エージェントはスライドを見られないので、手がかりを出す）。
 * はみ出しは、文字の幅（和文は 1 文字が大きさと同じ、欧文は 0.55 倍）で行数を数え、行の高さを大きさの 1.2 倍として見積もる。
 */
function textHints(body, x) {
    const sizes = [];
    const scale = body.fontScale || 1;
    const width = Math.max(0, x.cx - body.insets.l - body.insets.r) / model_1.EMU_PER_PT;
    const height = Math.max(0, x.cy - body.insets.t - body.insets.b) / model_1.EMU_PER_PT;
    let used = 0;
    for (const p of body.paragraphs) {
        const lineSpacing = (p.lineSpacing ?? 1) * (1 - (body.lineSpaceReduction || 0));
        let lineMax = p.endSize * scale;
        let lineWidth = 0;
        let lines = 1;
        const indent = Math.max(0, (p.marginLeft + p.indent) / model_1.EMU_PER_PT);
        for (const r of p.runs) {
            const size = r.resolved.size * scale;
            if (r.kind === 'break') {
                lines++;
                lineWidth = 0;
                continue;
            }
            if (r.text.trim())
                sizes.push(size);
            lineMax = Math.max(lineMax, size);
            for (const ch of r.text) {
                const w = /[\u3000-\u9fff\uff00-\uffef]/.test(ch) ? size : size * 0.55;
                if (body.wrap && width > 0 && lineWidth + w > width - indent && lineWidth > 0) {
                    lines++;
                    lineWidth = 0;
                }
                lineWidth += w;
            }
        }
        used += (p.spaceBefore ?? 0) + (p.spaceAfter ?? 0) + (p.lineSpacingPt ?? lineMax * 1.2 * lineSpacing) * lines;
    }
    if (!sizes.length)
        return {};
    const round = (v) => Math.round(v * 10) / 10;
    return {
        fontPt: { min: round(Math.min(...sizes)), max: round(Math.max(...sizes)) },
        ...(height > 0 && used > height * 1.25 && used - height > 8 ? { mayOverflow: true } : {}),
    };
}
function outlineOf(doc) {
    const deck = doc.toModel();
    const { cx, cy } = deck.slideSize;
    const slides = deck.slides.map((s, i) => {
        const elements = [];
        const walk = (els, group) => {
            for (const e of els) {
                const cNvPrId = e.id.slice(e.id.indexOf(':') + 1);
                const x = e.xfrm;
                const o = {
                    cNvPrId,
                    elementId: e.id,
                    kind: kindOf(e),
                    name: e.name,
                    ...(e.placeholder ? { placeholder: e.placeholder.type } : {}),
                    ...(group ? { group } : {}),
                    box: { x: x.x, y: x.y, w: x.cx, h: x.cy },
                    bboxPt: (0, model_1.bboxPt)(x),
                };
                if (e.kind === 'shape' && e.text)
                    Object.assign(o, textHints(e.text, x));
                if (e.kind === 'shape') {
                    const text = doc.elementText(e.id) ?? '';
                    if (text) {
                        o.text = text;
                        let start = 0;
                        o.lines = text.split('\n').map((t) => {
                            const line = { start, text: t };
                            start += t.length + 1;
                            return line;
                        });
                    }
                }
                if (e.kind === 'table') {
                    o.table = e.rows.map((r) => r.cells.map((c) => (c.hidden ? '' : c.text.paragraphs.map((p) => p.runs.map((x) => x.text).join('')).join('\n'))));
                }
                elements.push(o);
                if (e.kind === 'group')
                    walk(e.children, cNvPrId);
            }
        };
        walk(s.elements);
        const titleEl = elements.find((e) => e.placeholder === 'title' || e.placeholder === 'ctrTitle') ?? elements.find((e) => e.text);
        return {
            number: i + 1,
            id: s.id,
            title: (titleEl?.text ?? '').split('\n')[0] ?? '',
            ...(s.layoutName ? { layout: s.layoutName } : {}),
            ...(s.hidden ? { hidden: true } : {}),
            elements,
            notes: s.notes,
            comments: deck.comments?.counts[s.id] ?? { open: 0, resolved: 0 },
        };
    });
    return { slideSize: { cx, cy }, slides };
}
const KIND_LABEL = {
    shape: '図形',
    textBox: 'テキストボックス',
    connector: '線',
    picture: '画像',
    group: 'グループ',
    table: '表',
    other: 'その他',
};
/**
 * 端末にそのまま出せない文字（\n と \t 以外の制御文字、文字の向きを変える制御文字）を \uXXXX に置き換える。
 * コメントや図形の文字列は外から来る（他人の pptx）。git の textconv で差分に出ると、エスケープシーケンスで表示を隠したり、
 * クリップボードを書き換えたりできるため。--json は JSON の書き方でエスケープされるので使わない。
 */
function terminalSafe(s) {
    return s.replace(/[\u0000-\u0008\u000b-\u001f\u007f-\u009f\u061c\u200e\u200f\u202a-\u202e\u2066-\u2069]/g, (c) => `\\u${c.charCodeAt(0).toString(16).padStart(4, '0')}`);
}
/** outline と list の --slide（「2」「2,5」「2-4」）。 */
function parseSlideFilter(v, count) {
    if (v === undefined)
        return undefined;
    const out = new Set();
    for (const part of v.split(',')) {
        const m = /^\s*(\d+)\s*(?:-\s*(\d+)\s*)?$/.exec(part);
        if (!m)
            throw new Error(`--slide は番号で指定してください（例 2、2,5、2-4）：${v}`);
        const a = Number(m[1]);
        const b = m[2] ? Number(m[2]) : a;
        for (let n = Math.min(a, b); n <= Math.max(a, b) && n <= count; n++)
            if (n >= 1)
                out.add(n);
    }
    return out;
}
function formatOutline(o) {
    // 単位はここで 1 回だけ書き、図形の行は bbox [左, 上, 右, 下] の数だけにする
    const out = [`スライドの大きさ [${(0, model_1.ptOf)(o.slideSize.cx)}, ${(0, model_1.ptOf)(o.slideSize.cy)}]。図形の bbox は [左, 上, 右, 下]。どちらも pt（1pt = 12700 EMU）`];
    for (const s of o.slides) {
        const c = s.comments.open || s.comments.resolved ? `  コメント 未解決 ${s.comments.open}・解決済み ${s.comments.resolved}` : '';
        out.push(`スライド ${s.number}（ID ${s.id}）「${s.title}」${s.layout ? `  レイアウト：${s.layout}` : ''}${s.hidden ? '  スキップ' : ''}${c}`);
        for (const e of s.elements) {
            const indent = e.group ? '    ' : '  ';
            const ph = e.placeholder ? `（プレースホルダー ${e.placeholder}）` : '';
            const grp = e.group ? `（グループ #${e.group} の中）` : '';
            const f = e.fontPt ? `  文字 ${e.fontPt.min === e.fontPt.max ? e.fontPt.min : `${e.fontPt.min}〜${e.fontPt.max}`}pt` : '';
            const over = e.mayOverflow ? '  ⚠ 枠からはみ出すおそれ（見積もり）' : '';
            out.push(`${indent}#${e.cNvPrId} ${KIND_LABEL[e.kind]}${ph}${grp}「${e.name}」 bbox [${e.bboxPt.join(', ')}]${f}${over}`);
            for (const l of e.lines ?? [])
                out.push(`${indent}    ${String(l.start).padStart(4)}│ ${l.text}`);
            for (const row of e.table ?? [])
                out.push(`${indent}    | ${row.map((x) => x.replace(/\n/g, ' / ')).join(' | ')} |`);
        }
        if (s.notes) {
            out.push('  ノート：');
            for (const l of s.notes.split('\n'))
                out.push(`    ${l}`);
        }
        out.push('');
    }
    return terminalSafe(out.join('\n').trimEnd());
}
function anchorLabel(a, names) {
    const el = (id) => (id ? `#${id.slice(id.indexOf(':') + 1)}「${names.get(id) ?? ''}」` : '');
    switch (a.kind) {
        case 'slide':
            return 'スライド';
        case 'element':
            return a.orphaned ? '図形（元の内容は削除されました）' : `図形 ${el(a.elementId)}`;
        case 'table':
            return a.orphaned ? '表（元の内容は削除されました）' : `表 ${el(a.elementId)}`;
        case 'text':
            return a.orphaned
                ? `文字の範囲 ${a.start}:${a.length}（元の内容は削除されました）`
                : `図形 ${el(a.elementId)}の文字「${(a.quote ?? '').replace(/\n/g, '\\n')}」（${a.start}:${a.length}）`;
        case 'unknown':
            return `スライド（${a.tag}）`;
    }
}
/** 作成日時を、この PC のタイムゾーンの「2026-10-30 21:15」にする（読めなければそのまま）。 */
const time = (s) => {
    const d = (0, comments_1.parseCommentTime)(s);
    if (!d)
        return s;
    const p = (n) => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())} ${p(d.getHours())}:${p(d.getMinutes())}`;
};
/** list の絞り込み。 */
function filterThreadList(doc, threads, f = {}) {
    const num = new Map(doc.toModel().slides.map((s, i) => [s.id, i + 1]));
    return threads.filter((t) => (!f.slides || f.slides.has(num.get(t.slideId) ?? 0)) && (!f.status || t.status === f.status));
}
function formatThreads(doc, f = {}) {
    const deck = doc.toModel();
    const num = new Map(deck.slides.map((s, i) => [s.id, i + 1]));
    const names = new Map();
    const walk = (els) => {
        for (const e of els) {
            names.set(e.id, e.name);
            if (e.kind === 'group')
                walk(e.children);
        }
    };
    for (const s of deck.slides)
        walk(s.elements);
    const all = deck.comments?.threads ?? [];
    const threads = filterThreadList(doc, all, f);
    if (!threads.length) {
        if (deck.comments?.truncated)
            return 'コメントが多すぎるため、表示できません';
        return all.length ? '該当するコメントはありません' : 'コメントはありません';
    }
    const out = [];
    if (deck.comments?.truncated)
        out.push('注意：コメントが多すぎるため、一部だけを表示します', '');
    const open = threads.filter((t) => t.status !== 'resolved').length;
    out.push(`未解決 ${open}・解決済み ${threads.length - open}`, '');
    for (const t of threads) {
        const state = t.status === 'resolved' ? '解決済み' : '未解決';
        out.push(`スライド ${num.get(t.slideId) ?? '?'} · ${state}${t.format === 'legacy' ? ' · 旧形式（読み取り専用）' : ''} · ${t.id}`);
        out.push(`  場所：${anchorLabel(t.anchor, names)}`);
        const body = (s, indent) => s.split('\n').join(`\n${indent}`);
        out.push(`  ${t.root.author} · ${time(t.root.created)}：${body(t.root.text, '    ')}`);
        for (const image of t.root.images ?? [])
            out.push(`    画像：${image.name}  [${image.id}]`);
        for (const r of t.replies) {
            out.push(`    ↳ ${r.author} · ${time(r.created)}：${body(r.text, '      ')}  [${r.id}]`);
            for (const image of r.images ?? [])
                out.push(`      画像：${image.name}  [${image.id}]`);
        }
        out.push('');
    }
    return terminalSafe(out.join('\n').trimEnd());
}
