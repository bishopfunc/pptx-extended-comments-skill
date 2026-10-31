#!/usr/bin/env node
"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
// slides-comments の入り口（スキルの scripts/slides-comments.js になる）。
// 引数を lib/cli/slidesComments.js の runCli に渡し、その終了コードで終わる。処理の中身は runCli にある。
const slidesComments_1 = require("./lib/cli/slidesComments.js");
void (0, slidesComments_1.runCli)(process.argv.slice(2), {
    out: (s) => process.stdout.write(`${s}\n`),
    err: (s) => process.stderr.write(`${s}\n`),
    env: process.env,
}).then((code) => {
    process.exitCode = code;
});
