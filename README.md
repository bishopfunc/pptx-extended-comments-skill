# pptx-extended-comments-skill

PPTX のレビューコメントに画像を添付して、スライドの修正案や変更前後の比較を伝えるためのエージェント用スキルです。
VS Code の拡張機能 [Google Slides Sync](https://github.com/bishopfunc/google_slides_sync_extension) のエディタで、コメントと画像を表示できます。

- `SKILL.md`：スキルの説明と手順（エージェントが読む）
- `scripts/slides-comments.js`：コメントを読み書きする CLI の入り口（Node.js 20 以上。追加の npm パッケージは不要）
- `scripts/lib/`：CLI の処理。ソースのファイルごとの JS で、関数の名前とコメントはソースと同じ
- `scripts/vendor/`：依存パッケージ（jszip、@xmldom/xmldom）
- `scripts/README.md`：scripts のファイルの一覧と、CLI が読み書きするもの

## 導入

このリポジトリを、エージェントのスキルのフォルダに `pptx-extended-comments` という名前でクローンします。

```sh
# Claude Code（全プロジェクトで使う）
git clone https://github.com/bishopfunc/pptx-extended-comments-skill.git ~/.claude/skills/pptx-extended-comments

# Claude Code（このプロジェクトだけで使う）
git clone https://github.com/bishopfunc/pptx-extended-comments-skill.git .claude/skills/pptx-extended-comments
```

更新するときは、クローンしたフォルダで `git pull` します。

VS Code の拡張機能を入れていれば、コマンド「Google Slides Sync: エージェント用のレビュー手順を書き出す」でも、開いているフォルダの `.google-slides-sync/skills/pptx-extended-comments/` にこのスキルを書き出せます。

## CLI

```sh
node scripts/slides-comments.js outline deck.pptx            # スライドごとの図形と文字列
node scripts/slides-comments.js list deck.pptx --json        # 今のコメント（画像を含む）
node scripts/slides-comments.js add deck.pptx --slide 3 --text '見出しを結論に変える案' --author Reviewer
node scripts/slides-comments.js attach deck.pptx --thread '{スレッドID}' --image proposal.png --author Reviewer
```

引数なしで実行すると、すべてのコマンドとオプションを表示します。

## ソース

`scripts/` のファイルはビルドで作ります。直接は編集しません。
ソースは Google Slides Sync のリポジトリの `src/cli/` と `src/pptx/` などにあり、拡張機能のエディタと同じ処理で pptx を読み書きします。
このリポジトリは、拡張機能のリポジトリにサブモジュール（`skills/pptx-extended-comments-skill/`）として入っていて、拡張機能のビルド（`npm run build`）が `scripts/` を書き出します。

処理を確かめられるように、一つの JS にまとめず、TypeScript のソースから型だけを取り除いて、ファイルごとに書き出しています。
どのファイルが何をするかは `scripts/README.md` の一覧にあります。コマンドの振り分けと、pptx の読み書きの手順は `scripts/lib/cli/slidesComments.js` です。
CLI はネットワークに接続せず、ほかのプログラムも実行しません。
