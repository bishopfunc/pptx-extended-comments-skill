"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.OpRejectedError = void 0;
/** 編集操作が拒否された（存在しない ID、グラフを含むスライドの複製など）。メッセージは利用者向けの日本語。 */
class OpRejectedError extends Error {
    constructor(message) {
        super(message);
        this.name = 'OpRejectedError';
    }
}
exports.OpRejectedError = OpRejectedError;
// 実装側は次の静的メソッドを持つ：
//   PptxDoc.load(bytes: Uint8Array): Promise<PptxDoc>
