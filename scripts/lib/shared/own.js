"use strict";
// 表（プレーンなオブジェクト）を、外から来たキー（pptx の中身、Webview のメッセージ）で引くとき、
// __proto__・constructor・toString などの組み込みの名前を値とみなさないための関数（ADR-0019）。
Object.defineProperty(exports, "__esModule", { value: true });
exports.own = own;
/** rec が key を自分のキーとして持つときだけ、その値を返す。 */
function own(rec, key) {
    return typeof key === 'string' && Object.hasOwn(rec, key) ? rec[key] : undefined;
}
