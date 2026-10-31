# pptx-extended-comments-skill

pptx のレビューコメントに画像を添付して、スライドの修正案や変更前後の比較を伝えるための、エージェント（Claude Code など）用のスキルです。
添付した画像は、VS Code の拡張機能 [Google Slides Sync](https://github.com/bishopfunc/google_slides_sync_extension) のエディタで表示できます。

## 導入

このリポジトリを、スキルのフォルダに `pptx-extended-comments` という名前でクローンします。

```sh
# すべてのプロジェクトで使う
git clone https://github.com/bishopfunc/pptx-extended-comments-skill.git ~/.claude/skills/pptx-extended-comments

# このプロジェクトだけで使う
git clone https://github.com/bishopfunc/pptx-extended-comments-skill.git .claude/skills/pptx-extended-comments
```

更新するときは、クローンしたフォルダで `git pull` します。
動かすには Node.js 20 以上が要ります（追加の npm パッケージは要りません）。

## 使い方

エージェントに「このスライドの修正案を画像付きのコメントで出して」のように頼むと、エージェントが `SKILL.md` の手順でコメントを付けます。
コメントを読み書きする CLI は、直接実行することもできます。

```sh
node scripts/slides-comments.js outline deck.pptx         # スライドごとの図形と文字列
node scripts/slides-comments.js list deck.pptx --json     # 今のコメント（画像を含む）
node scripts/slides-comments.js add deck.pptx --slide 3 --text '見出しを結論に変える案' --author Reviewer
node scripts/slides-comments.js attach deck.pptx --thread '{スレッドID}' --image proposal.png --author Reviewer
```

引数を付けずに実行すると、すべてのコマンドとオプションを表示します。

## 中身

- `SKILL.md`：エージェントが読む手順
- `scripts/`：コメントを読み書きする CLI。ネットワークに接続せず、ほかのプログラムも実行しません。ファイルの一覧と、CLI が読み書きするものは `scripts/README.md` にあります。
