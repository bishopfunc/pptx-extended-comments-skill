"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ptOf = exports.EMU_PER_PX = exports.EMU_PER_PT = exports.EMU_PER_INCH = void 0;
exports.bboxPt = bboxPt;
exports.EMU_PER_INCH = 914400;
exports.EMU_PER_PT = 12700;
/** Webview で 100% 表示のときの換算（96dpi）。 */
exports.EMU_PER_PX = 9525;
/**
 * EMU を pt にする（小数 1 桁）。AI エージェントに渡す位置と大きさ（選択の共有、コメントの CLI の outline）は pt で書く。
 * % と違って縦と横が同じ単位になり、×12700 でちょうど EMU に戻る。
 */
const ptOf = (emu) => Math.round((emu / exports.EMU_PER_PT) * 10) / 10;
exports.ptOf = ptOf;
/** 位置と大きさを、pt の bbox [左, 上, 右, 下] にする。 */
function bboxPt(x) {
    return [(0, exports.ptOf)(x.x), (0, exports.ptOf)(x.y), (0, exports.ptOf)(x.x + x.cx), (0, exports.ptOf)(x.y + x.cy)];
}
