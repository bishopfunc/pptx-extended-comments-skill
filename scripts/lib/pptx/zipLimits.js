"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
exports.ZipTooLargeError = exports.ZIP_TOO_LARGE_MESSAGE = exports.DEFAULT_ZIP_LIMITS = void 0;
exports.openZip = openZip;
exports.zipEntries = zipEntries;
exports.readZipEntry = readZipEntry;
// zip（pptx）を展開するときの上限（ADR-0019）。展開すると巨大になる zip（zip bomb）でプロセスが落ちないよう、
// 1 つずつ少しずつ展開しながらバイト数を数え、上限を超えたらそこで止める。vscode に依存しない。
const jszip_1 = __importDefault(require("../../vendor/jszip.js"));
exports.DEFAULT_ZIP_LIMITS = {
    maxEntries: 10_000,
    maxEntryBytes: 100 * 1024 * 1024,
    maxTotalBytes: 1024 * 1024 * 1024,
    maxRatio: 500,
    minRatioCheckBytes: 10 * 1024 * 1024,
};
exports.ZIP_TOO_LARGE_MESSAGE = 'pptx が大きすぎるか、壊れています';
class ZipTooLargeError extends Error {
    constructor(detail) {
        super(`${exports.ZIP_TOO_LARGE_MESSAGE}（${detail}）`);
        this.name = 'ZipTooLargeError';
    }
}
exports.ZipTooLargeError = ZipTooLargeError;
/** zip を開く（中身はまだ展開しない）。ファイルの数が上限を超えれば ZipTooLargeError。 */
async function openZip(bytes, limits = exports.DEFAULT_ZIP_LIMITS) {
    const zip = await jszip_1.default.loadAsync(bytes);
    const n = Object.keys(zip.files).length;
    if (n > limits.maxEntries)
        throw new ZipTooLargeError(`ファイルが ${n} 個あります`);
    return zip;
}
/** zip の中のディレクトリでないファイル。 */
function zipEntries(zip) {
    return Object.values(zip.files).filter((f) => !f.dir);
}
/** 1 つのファイルを、上限を確かめながら展開する。 */
function readZipEntry(f, budget, limits = exports.DEFAULT_ZIP_LIMITS) {
    const compressed = f._data?.compressedSize;
    return new Promise((resolve, reject) => {
        const chunks = [];
        let size = 0;
        let done = false;
        const stream = f.internalStream('uint8array');
        const fail = (e) => {
            if (done)
                return;
            done = true;
            stream.pause();
            chunks.length = 0;
            reject(e);
        };
        stream
            .on('data', (chunk) => {
            if (done)
                return;
            size += chunk.length;
            budget.total += chunk.length;
            if (size > limits.maxEntryBytes)
                return fail(new ZipTooLargeError(`${f.name} が大きすぎます`));
            if (budget.total > limits.maxTotalBytes)
                return fail(new ZipTooLargeError('展開した合計が大きすぎます'));
            if (compressed !== undefined && size > limits.minRatioCheckBytes && size > compressed * limits.maxRatio) {
                return fail(new ZipTooLargeError(`${f.name} の圧縮率が高すぎます`));
            }
            chunks.push(chunk);
        })
            .on('error', fail)
            .on('end', () => {
            if (done)
                return;
            done = true;
            if (chunks.length === 1)
                return resolve(chunks[0]);
            const out = new Uint8Array(size);
            let off = 0;
            for (const c of chunks) {
                out.set(c, off);
                off += c.length;
            }
            resolve(out);
        })
            .resume();
    });
}
