---
name: pptx-extended-comments
description: Attach images to PPTX review comments to show revision proposals, before-and-after comparisons, and alternatives made in parallel. Use when reviewing a .pptx and the feedback is clearer as a picture. The images are shown in the Google Slides Sync editor for VS Code. Not for authoring or editing the slides themselves.
---

# PPTX の画像付きコメント

pptx のレビューコメントに画像を添付して、スライドの修正案を見た目ごと伝えるためのスキル。
スライドの編集や画像の描画はこのスキルでは行わないので、その環境で使えるスライドのスキルやツールを使う。

コメントの読み書きには、このフォルダの `scripts/slides-comments.js`（以下、CLI）を使う。
CLI は Node.js 20 以上で動き、VS Code や追加の npm パッケージは要らない。
実行するときは、この `SKILL.md` の場所から CLI の絶対パスを求める。

## 画像付きコメントを付ける手順

1. `outline` でスライドの図形と文字列を、`list --json` で既存のコメントを読み、対象のスライドと要素を確かめる。
2. 比較に使う画像を PNG か JPEG で用意する。変更前と修正案は同じ寸法に揃え、文字が読める大きさにする。
3. 既存のスレッドに付けるなら、そのスレッド ID を使う。新しく付けるなら `add` でコメントを作る。`add` はスレッド ID を出力する。
4. `attach` で画像を一枚ずつ添付する。コメントの本文には、画像の順番と変更の理由を書く。
5. `list --json` で読み戻し、対象のスライド、本文、画像がそろっているかを確かめる。失敗した添付だけをやり直す。

```sh
CLI=/absolute/path/to/pptx-extended-comments/scripts/slides-comments.js
node "$CLI" add deck.pptx --slide 3 --text '見出しを結論に変える案。画像は変更前、修正案の順。' --author Reviewer
# THREAD には add が出力したスレッド ID を入れる
node "$CLI" attach deck.pptx --thread THREAD --image before.png --author Reviewer
node "$CLI" attach deck.pptx --thread THREAD --image proposal.png --author Reviewer
node "$CLI" list deck.pptx --slide 3 --json
```

作成者（`--author`）には、自分の名前かエージェントの名前を使う。
CLI で編集や削除ができるのは自分の投稿だけなので、ほかの人のコメントと返信はそのまま残る。
画像を作れなかったときや添付に失敗したときは、そのことを依頼者に報告する。

## 使いどころ

- **修正案**：見出し、図表、レイアウトの修正案を画像で示す。レビューだけを頼まれたときは、元のスライドを変えない。
- **変更前後の比較**：同じコメントに変更前と修正後を添付し、何が変わったかを見比べられるようにする。
- **並列で作った案の比較**：複数の作業者が別々に作った案を、一つのコメントに並べる。手順は次の節にある。

## 並列で作った案を比べる手順

並列で作業できて、それが許されている場合に限る。

1. 基準にする版を決め、作業者ごとに別のコピーと、担当するスライドや案を渡す。
2. 各作業者が自分のコピーを編集し、結果を画像に描き出す。
3. 一人が画像を集め、案 A、案 B の順にコメントへ添付する。

同じ pptx に複数の作業者が同時に保存しないようにする。
CLI のロックは同時の書き込みを防ぐだけで、変更を自動でマージはしないので、同時に保存すると片方の変更が失われる。
採用された案を反映するときは、元の版がその後に変わっていないかを確かめ、依頼の範囲で変更をまとめてから画像を描き直す。

## コマンド

| コマンド | 内容 |
| --- | --- |
| `outline <deck.pptx> [--slide N] [--json]` | スライドごとの図形（ID、種類、名前、位置、文字列）、表、ノート |
| `list <deck.pptx> [--slide N] [--open \| --resolved] [--json]` | コメントの一覧（画像を含む） |
| `add <deck.pptx> --slide N [--shape ID] [--text-range START:LEN \| --quote "…"] --text "…"` | コメントを作る |
| `reply <deck.pptx> --thread ID --text "…"` | スレッドに返信する |
| `edit <deck.pptx> --thread ID [--reply ID] --text "…"` | 自分の投稿の本文を書き換える |
| `attach <deck.pptx> --thread ID [--reply ID] --image file.png` | 投稿に画像を添付する |
| `detach <deck.pptx> --thread ID [--reply ID] --image-id ID` | 投稿から画像を外す |
| `resolve <deck.pptx> --thread ID` | スレッドを解決済みにする |
| `reopen <deck.pptx> --thread ID` | 解決済みのスレッドを開き直す |
| `delete <deck.pptx> --thread ID [--reply ID] [--force]` | 自分の投稿を消す |

オプションの意味は次のとおり。

- `--slide`：1 から数えたスライドの番号。`outline` と `list` では `2,5` や `2-4` とも書ける。
- `--shape`：`outline` に出る図形の ID か名前。コメントをその図形に付ける。
- `--text-range`：図形の文字列の位置（0 から数え、段落の区切りと改行は 1 文字）と長さ。`--quote` を使うと、図形の文字列で最初に見つかった引用の位置に付ける。
- `--thread` と `--reply`：`list` に出る ID（前後の `{ }` は省いてよい）。返信に添付するときは `--reply` も指定する。
- `--author`：作成者の名前。省くと環境変数 `SLIDES_COMMENTS_AUTHOR`、それもなければ「Claude」になる。メールアドレスは使えない。
- `--force`：ほかの人の返信があるスレッドの最初の投稿を消すときに必要。最初の投稿を消すとスレッドごと消えるので、ふつうは `resolve` を使う。

引数を付けずに実行すると、使い方を表示する。

## 画像の制限

- 一つの投稿に添付できるのは 4 枚まで。それより多いときは、同じスレッドに自分の返信を作り、そこへ添付する。
- 形式は PNG か JPEG で、一枚 2 MB まで。各辺は 8192 px 以下、画素数は 1600 万以下にする。

## 画像の保存先と表示

画像は外部のファイルを参照するのではなく、pptx の中のコメントのデータに埋め込まれる。
そのため pptx を一つ渡せば画像も一緒に渡り、画像が別のフォルダに書き出されることもない。

画像を表示できるのは、VS Code の拡張機能 Google Slides Sync のエディタである。
画像の保存形式はこの拡張機能の独自のものなので、PowerPoint などほかのアプリで画像が表示されることは想定していない。
Google Slides Sync で Google スライドと同期しても、コメントと画像は Google へは送られず、手元の pptx に残る。
