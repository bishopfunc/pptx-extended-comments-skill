"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.TABLE_MSG = void 0;
exports.applyTableOp = applyTableOp;
const edit_1 = require("./edit.js");
const parse_1 = require("./parse.js");
const xml_1 = require("./xml.js");
exports.TABLE_MSG = {
    notTable: 'この要素は表ではありません',
    insertRowMerge: '縦に結合したセルの途中には行を挿入できません。先にセルの結合を解除してください',
    insertColMerge: '横に結合したセルの途中には列を挿入できません。先にセルの結合を解除してください',
    deleteRowMerge: '縦に結合したセルを含む行は削除できません。先にセルの結合を解除してください',
    deleteColMerge: '横に結合したセルを含む列は削除できません。先にセルの結合を解除してください',
    lastRow: '表の最後の行は削除できません。表を削除してください',
    lastCol: '表の最後の列は削除できません。表を削除してください',
    notMerged: '結合したセルではありません',
};
/** 行・列の識別子（PowerPoint 2016 以降が a:tr / a:gridCol の a:extLst に書く a16:rowId / a16:colId）の名前空間。 */
const A16 = 'http://schemas.microsoft.com/office/drawing/2014/main';
const MERGE_ATTRS = ['gridSpan', 'rowSpan', 'hMerge', 'vMerge'];
/** 操作する表（分岐ごと）。表でなければ・グループの中なら拒否する。 */
function targets(txn, elementId) {
    const f = (0, edit_1.findElement)(txn, elementId);
    if (f.ref.groups.length)
        (0, edit_1.reject)(edit_1.MSG.groupChild);
    const out = [];
    for (const b of (0, edit_1.branches)(f.ref.el)) {
        if (!(0, xml_1.is)(b, 'p:graphicFrame'))
            continue;
        const tbl = (0, xml_1.path)(b, 'a:graphic/a:graphicData/a:tbl');
        if (tbl)
            out.push({ frame: b, tbl });
    }
    if (!out.length)
        (0, edit_1.reject)(exports.TABLE_MSG.notTable);
    return out;
}
function rowsOf(tbl) {
    return (0, xml_1.kidsNamed)(tbl, 'a:tr');
}
function gridColsOf(tbl) {
    return (0, xml_1.kidsNamed)((0, xml_1.kid)(tbl, 'a:tblGrid'), 'a:gridCol');
}
function span(tc, name) {
    return Math.max(1, (0, xml_1.numAttr)(tc, name) ?? 1);
}
function checkIndex(n, size) {
    if (!Number.isInteger(n) || n < 0 || n >= size)
        (0, edit_1.reject)(edit_1.MSG.cellNotFound);
}
/** 結合の起点のセル（gridSpan / rowSpan が 2 以上）と、その範囲。 */
function mergeOrigins(tbl) {
    const out = [];
    rowsOf(tbl).forEach((tr, row) => (0, xml_1.kidsNamed)(tr, 'a:tc').forEach((tc, col) => {
        if ((0, xml_1.boolAttr)(tc, 'hMerge') || (0, xml_1.boolAttr)(tc, 'vMerge'))
            return;
        const rows = span(tc, 'rowSpan');
        const cols = span(tc, 'gridSpan');
        if (rows > 1 || cols > 1)
            out.push({ row, col, rows, cols, tc });
    }));
    return out;
}
/** 行の境界 b（b 行目の上）が縦の結合を切るか。 */
function cutsRowMerge(tbl, b) {
    if (mergeOrigins(tbl).some((m) => m.row < b && b < m.row + m.rows))
        return true;
    return (0, xml_1.kidsNamed)(rowsOf(tbl)[b], 'a:tc').some((tc) => (0, xml_1.boolAttr)(tc, 'vMerge') === true);
}
/** 列の境界 b（b 列目の左）が横の結合を切るか。 */
function cutsColMerge(tbl, b) {
    if (mergeOrigins(tbl).some((m) => m.col < b && b < m.col + m.cols))
        return true;
    return rowsOf(tbl).some((tr) => (0, xml_1.boolAttr)((0, xml_1.kidsNamed)(tr, 'a:tc')[b], 'hMerge') === true);
}
/** 枠（p:graphicFrame の p:xfrm/a:ext）の大きさを変える。 */
function growFrame(frame, dcx, dcy) {
    const ext = (0, xml_1.path)(frame, 'p:xfrm/a:ext');
    if (!ext)
        return;
    if (dcx)
        (0, xml_1.setAttr)(ext, 'cx', (0, xml_1.emuSize)(((0, xml_1.numAttr)(ext, 'cx') ?? 0) + dcx));
    if (dcy)
        (0, xml_1.setAttr)(ext, 'cy', (0, xml_1.emuSize)(((0, xml_1.numAttr)(ext, 'cy') ?? 0) + dcy));
}
/** 複製した a:tr / a:gridCol の a16:rowId / a16:colId に、表の中で使われていない値を振る。 */
function freshIds(tbl, copy, local) {
    const isId = (e) => e.namespaceURI === A16 && e.localName === local;
    const targetsInCopy = (0, xml_1.descendants)(copy).filter(isId);
    if (!targetsInCopy.length)
        return;
    const used = new Set((0, xml_1.descendants)(tbl).filter(isId).map((e) => Number(e.getAttribute('val'))));
    const max = Math.max(0, ...[...used].filter(Number.isFinite));
    let next = max + 1;
    for (const e of targetsInCopy) {
        // 値は 32 ビットの符号なし整数。上限を超えたら、使われていない値を小さいほうから探す
        if (next > 4294967295) {
            next = 1;
            while (used.has(next))
                next++;
        }
        e.setAttribute('val', String(next));
        used.add(next);
        next++;
    }
}
/** セルを複製して空にする（結合を外し、a:bodyPr / a:lstStyle / a:tcPr は残す）。 */
function emptyCellFrom(tc) {
    const c = tc.cloneNode(true);
    for (const a of MERGE_ATTRS)
        c.removeAttribute(a);
    const tx = (0, xml_1.kid)(c, 'a:txBody');
    if (!tx)
        return c;
    const doc = c.ownerDocument;
    const ps = (0, xml_1.kidsNamed)(tx, 'a:p');
    const first = ps[0];
    // 空の段落：最初の段落の a:pPr と、最初のランの a:rPr から作った a:endParaRPr（ランがなければ元の a:endParaRPr）
    const p = (0, xml_1.create)(doc, 'a:p');
    const pPr = (0, xml_1.kid)(first, 'a:pPr');
    if (pPr)
        p.appendChild(pPr.cloneNode(true));
    const firstRun = first ? (0, parse_1.runChildren)(first)[0] : undefined;
    const runRPr = (0, xml_1.kid)(firstRun, 'a:rPr');
    const ownEnd = (0, xml_1.kid)(first, 'a:endParaRPr');
    p.appendChild(runRPr || !ownEnd ? (0, edit_1.rPrToEndParaRPr)(doc, runRPr) : ownEnd.cloneNode(true));
    for (const x of ps)
        (0, xml_1.remove)(x);
    const extLst = (0, xml_1.kid)(tx, 'a:extLst');
    if (extLst)
        tx.insertBefore(p, extLst);
    else
        tx.appendChild(p);
    return c;
}
/** 親の子 ref の前に入れる（ref がなければ a:extLst の前、それもなければ末尾）。 */
function insertAt(parent, el, ref) {
    const before = ref ?? (0, xml_1.kid)(parent, 'a:extLst');
    if (before)
        parent.insertBefore(el, before);
    else
        parent.appendChild(el);
}
// ---------------------------------------------------------------------------
// 行・列の挿入
// ---------------------------------------------------------------------------
function insertTableRow(txn, op) {
    const ts = targets(txn, op.elementId);
    for (const { tbl } of ts) {
        const rows = rowsOf(tbl);
        checkIndex(op.row, rows.length);
        const b = op.below ? op.row + 1 : op.row;
        if (b < rows.length && cutsRowMerge(tbl, b))
            (0, edit_1.reject)(exports.TABLE_MSG.insertRowMerge);
    }
    for (const { frame, tbl } of ts) {
        const rows = rowsOf(tbl);
        const ref = rows[op.row];
        const copy = ref.cloneNode(false);
        for (const k of (0, xml_1.kids)(ref))
            copy.appendChild((0, xml_1.is)(k, 'a:tc') ? emptyCellFrom(k) : k.cloneNode(true));
        freshIds(tbl, copy, 'rowId');
        insertAt(tbl, copy, rows[op.below ? op.row + 1 : op.row]);
        growFrame(frame, 0, (0, xml_1.numAttr)(ref, 'h') ?? 0);
    }
}
function insertTableColumn(txn, op) {
    const ts = targets(txn, op.elementId);
    for (const { tbl } of ts) {
        const cols = gridColsOf(tbl);
        checkIndex(op.col, cols.length);
        const b = op.right ? op.col + 1 : op.col;
        if (b < cols.length && cutsColMerge(tbl, b))
            (0, edit_1.reject)(exports.TABLE_MSG.insertColMerge);
    }
    for (const { frame, tbl } of ts) {
        const cols = gridColsOf(tbl);
        const b = op.right ? op.col + 1 : op.col;
        const refCol = cols[op.col];
        const gc = refCol.cloneNode(true);
        freshIds(tbl, gc, 'colId');
        insertAt((0, xml_1.kid)(tbl, 'a:tblGrid'), gc, cols[b]);
        for (const tr of rowsOf(tbl)) {
            const tcs = (0, xml_1.kidsNamed)(tr, 'a:tc');
            const ref = tcs[Math.min(op.col, tcs.length - 1)];
            if (!ref)
                continue;
            // a:tc の後ろには a:extLst しか来ないので、b 番目の a:tc の前（なければ a:extLst の前）に入れる
            insertAt(tr, emptyCellFrom(ref), tcs[b]);
        }
        growFrame(frame, (0, xml_1.numAttr)(refCol, 'w') ?? 0, 0);
    }
}
// ---------------------------------------------------------------------------
// 行・列の削除
// ---------------------------------------------------------------------------
function deleteTableRow(txn, op) {
    const ts = targets(txn, op.elementId);
    for (const { tbl } of ts) {
        const rows = rowsOf(tbl);
        checkIndex(op.row, rows.length);
        if (rows.length <= 1)
            (0, edit_1.reject)(exports.TABLE_MSG.lastRow);
        // 縦の結合（rowSpan の起点か vMerge）にかかる行は消せない。横の結合だけなら行ごと消えるので構わない
        const inMerge = (0, xml_1.kidsNamed)(rows[op.row], 'a:tc').some((tc) => span(tc, 'rowSpan') > 1 || (0, xml_1.boolAttr)(tc, 'vMerge') === true);
        if (inMerge || mergeOrigins(tbl).some((m) => m.row < op.row && op.row < m.row + m.rows))
            (0, edit_1.reject)(exports.TABLE_MSG.deleteRowMerge);
    }
    for (const { frame, tbl } of ts) {
        const tr = rowsOf(tbl)[op.row];
        const h = (0, xml_1.numAttr)(tr, 'h') ?? 0;
        (0, xml_1.remove)(tr);
        growFrame(frame, 0, -h);
    }
}
function deleteTableColumn(txn, op) {
    const ts = targets(txn, op.elementId);
    for (const { tbl } of ts) {
        const cols = gridColsOf(tbl);
        checkIndex(op.col, cols.length);
        if (cols.length <= 1)
            (0, edit_1.reject)(exports.TABLE_MSG.lastCol);
        const inMerge = rowsOf(tbl).some((tr) => {
            const tc = (0, xml_1.kidsNamed)(tr, 'a:tc')[op.col];
            return !!tc && (span(tc, 'gridSpan') > 1 || (0, xml_1.boolAttr)(tc, 'hMerge') === true);
        });
        if (inMerge || mergeOrigins(tbl).some((m) => m.col < op.col && op.col < m.col + m.cols))
            (0, edit_1.reject)(exports.TABLE_MSG.deleteColMerge);
    }
    for (const { frame, tbl } of ts) {
        const gc = gridColsOf(tbl)[op.col];
        const w = (0, xml_1.numAttr)(gc, 'w') ?? 0;
        (0, xml_1.remove)(gc);
        for (const tr of rowsOf(tbl))
            (0, xml_1.remove)((0, xml_1.kidsNamed)(tr, 'a:tc')[op.col]);
        growFrame(frame, -w, 0);
    }
}
// ---------------------------------------------------------------------------
// 結合の解除
// ---------------------------------------------------------------------------
function unmergeTableCell(txn, op) {
    const ts = targets(txn, op.elementId);
    for (const { tbl } of ts) {
        const rows = rowsOf(tbl);
        checkIndex(op.row, rows.length);
        const tc = (0, xml_1.kidsNamed)(rows[op.row], 'a:tc')[op.col];
        if (!tc)
            (0, edit_1.reject)(edit_1.MSG.cellNotFound);
        if ((0, xml_1.boolAttr)(tc, 'hMerge') || (0, xml_1.boolAttr)(tc, 'vMerge') || (span(tc, 'gridSpan') === 1 && span(tc, 'rowSpan') === 1)) {
            (0, edit_1.reject)(exports.TABLE_MSG.notMerged);
        }
    }
    for (const { tbl } of ts) {
        const rows = rowsOf(tbl);
        const tc = (0, xml_1.kidsNamed)(rows[op.row], 'a:tc')[op.col];
        const nRows = span(tc, 'rowSpan');
        const nCols = span(tc, 'gridSpan');
        tc.removeAttribute('gridSpan');
        tc.removeAttribute('rowSpan');
        // 隠れていたセルを見せる（中身はそのまま。ふつうは空）
        for (let r = op.row; r < Math.min(rows.length, op.row + nRows); r++) {
            const tcs = (0, xml_1.kidsNamed)(rows[r], 'a:tc');
            for (let c = op.col; c < Math.min(tcs.length, op.col + nCols); c++) {
                if (r === op.row && c === op.col)
                    continue;
                tcs[c].removeAttribute('hMerge');
                tcs[c].removeAttribute('vMerge');
            }
        }
    }
}
function applyTableOp(txn, op) {
    switch (op.type) {
        case 'insertTableRow':
            insertTableRow(txn, op);
            return { changes: [] };
        case 'insertTableColumn':
            insertTableColumn(txn, op);
            return { changes: [] };
        case 'deleteTableRow':
            deleteTableRow(txn, op);
            return { changes: [] };
        case 'deleteTableColumn':
            deleteTableColumn(txn, op);
            return { changes: [] };
        case 'unmergeTableCell':
            unmergeTableCell(txn, op);
            return { changes: [] };
        default:
            return undefined;
    }
}
