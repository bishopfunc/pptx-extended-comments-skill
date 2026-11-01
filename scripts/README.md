# scripts の中身

コメントを読み書きする CLI です。
このフォルダのファイルは Google Slides Sync（https://github.com/bishopfunc/google_slides_sync_extension）のソースからビルドで作るので、直接は編集しません。
JS は TypeScript のソースから型だけを取り除いたもので、関数の名前とコメントはソースと同じです。

- `slides-comments.js`：入り口。コマンドの振り分けと pptx を読み書きする手順は `lib/cli/slidesComments.js` にあります。
- `lib/`：CLI が使う処理。
- `vendor/`：依存パッケージ（jszip、@xmldom/xmldom）を、中身を変えずにパッケージごとに一つにまとめたもの。ライセンスは `vendor/THIRD_PARTY_LICENSES.md` にあります。

CLI はネットワークに接続せず、ほかのプログラムも実行しません（`lib/editor/gitShow.js` などは読み込まれるだけで、CLI からは呼ばれません）。
読み書きするのは次のものだけです。

- 引数で指定した pptx：書き込むときは、同じフォルダの一時ファイルに書いてから置き換える
- `attach` の `--image` で指定した画像：読むだけ
- OS の一時フォルダの `google-slides-sync-<ユーザー ID>/`：エディタと同時に書き込まないためのロック

| ファイル | ソース | 内容 |
| --- | --- | --- |
| `slides-comments.js` | `src/cli/main.ts` | slides-comments の入り口（スキルの scripts/slides-comments.js になる）。 |
| `lib/cli/format.js` | `src/cli/format.ts` | slides-comments の outline と list の中身と書式（エージェントが読む。list は git の textconv にも使う）。 |
| `lib/cli/slidesComments.js` | `src/cli/slidesComments.ts` | slides-comments：pptx のレビューコメントをターミナルから読み書きする CLI。 |
| `lib/editor/commentAuthor.js` | `src/editor/commentAuthor.ts` | コメントの作成者の名前（設定 slidesSync.commentAuthor）。 |
| `lib/editor/gitShow.js` | `src/editor/gitShow.ts` | git の版の pptx をバイト列のまま読む（「スライドの差分を表示（HEAD と比べる）」）。 |
| `lib/pptx/arrange.js` | `src/pptx/arrange.ts` | 配置の操作：グループ化・グループ解除と、テキスト枠の余白・垂直方向の配置。 |
| `lib/pptx/commentImages.js` | `src/pptx/commentImages.ts` | コメントのextLstに自己完結した添付画像を保存する。 |
| `lib/pptx/commentOps.js` | `src/pptx/commentOps.ts` | コメントの編集操作（addComment など）と、テキストの編集に合わせた文字の範囲のアンカーの付け替え。 |
| `lib/pptx/comments.js` | `src/pptx/comments.ts` | レビューコメントの読み取りと、パッケージの組み立て。 |
| `lib/pptx/commentsMerge.js` | `src/pptx/commentsMerge.ts` | ディスク上で書き換えられたコメント（エージェントの CLI など）を、未保存の編集を持つエディタの内容に取り込む 3 者マージ。 |
| `lib/pptx/commentsSummary.js` | `src/pptx/commentsSummary.ts` | 2 つの版のコメントの違いの要約（例「コメント +2・解決 1・返信 3」）。 |
| `lib/pptx/contentHash.js` | `src/pptx/contentHash.ts` | コメントを除いた pptx の中身のハッシュ。 |
| `lib/pptx/creationIds.js` | `src/pptx/creationIds.ts` | スライドと図形の作成 ID（p14:creationId、a16:creationId）。 |
| `lib/pptx/edit.js` | `src/pptx/edit.ts` | EditOp の適用（要素の操作）。 |
| `lib/pptx/elements.js` | `src/pptx/elements.ts` | スライドの図形の列挙と ID（`<スライドID>:<cNvPr@id>[#n]`）、グループの座標変換。 |
| `lib/pptx/index.js` | `src/pptx/index.ts` | PptxDoc：pptx パッケージの読み込み、モデル生成、編集、undo 用の書き戻し、保存。 |
| `lib/pptx/inherit.js` | `src/pptx/inherit.ts` | 継承チェーンの解決（プレースホルダー、lstStyle、p:style、背景）。 |
| `lib/pptx/integrity.js` | `src/pptx/integrity.ts` | パッケージの整合性検査。 |
| `lib/pptx/layouts.js` | `src/pptx/layouts.ts` | レイアウトの一覧。 |
| `lib/pptx/package.js` | `src/pptx/package.ts` | zip の読み書き、パーツ管理、[Content_Types].xml、rels。 |
| `lib/pptx/paragraph.js` | `src/pptx/paragraph.ts` | 段落の書式（配置、箇条書き、行間、レベルとインデント）の書き込み。 |
| `lib/pptx/parse.js` | `src/pptx/parse.ts` | スライド → DeckModel。 |
| `lib/pptx/slideClipboard.js` | `src/pptx/slideClipboard.ts` | スライドのコピーと貼り付け。 |
| `lib/pptx/slides.js` | `src/pptx/slides.ts` | スライドの追加・複製・削除・移動、ノート、スキップ、背景、レイアウトの適用。 |
| `lib/pptx/stripComments.js` | `src/pptx/stripComments.ts` | アップロードする pptx からコメントを取り除く。 |
| `lib/pptx/tableOps.js` | `src/pptx/tableOps.ts` | 表の行・列の挿入と削除、セルの結合の解除。 |
| `lib/pptx/tableStyle.js` | `src/pptx/tableStyle.ts` | 表のスタイル（ppt/tableStyles.xml の a:tblStyle と組み込みスタイル）をセルごとに解決する。 |
| `lib/pptx/templates.js` | `src/pptx/templates.ts` | 新規図形・スライド・ノートマスターの XML 雛形。 |
| `lib/pptx/theme.js` | `src/pptx/theme.ts` | テーマ色、カラーマップ、色変換、フォント。 |
| `lib/pptx/types.js` | `src/pptx/types.ts` | pptx モジュールの公開契約。 |
| `lib/pptx/xml.js` | `src/pptx/xml.ts` | 名前空間、DOM ヘルパー、子要素の順序表。 |
| `lib/pptx/zipLimits.js` | `src/pptx/zipLimits.ts` | zip（pptx）を展開するときの上限。 |
| `lib/shared/commentImages.js` | `src/shared/commentImages.ts` | コメントの添付画像の検査（PNG/JPEG の data URL の形式、署名、サイズ、画素数）と、画像の名前の整え方。 |
| `lib/shared/comments.js` | `src/shared/comments.ts` | レビューコメントの表示用モデルと、編集操作のアンカーの指定。 |
| `lib/shared/gitUri.js` | `src/shared/gitUri.ts` | git の版の pptx（git: の URI）の扱い。 |
| `lib/shared/model.js` | `src/shared/model.ts` | 表示用モデル。 |
| `lib/shared/own.js` | `src/shared/own.ts` | 表（プレーンなオブジェクト）を、外から来たキー（pptx の中身、Webview のメッセージ）で引くとき、 __proto__・constructor・toString などの組み込みの名前を値とみなさないための関数。 |
| `lib/shared/saveTarget.js` | `src/shared/saveTarget.ts` | 新しいプレゼンテーションと PDF の保存先（開いているフォルダに保存し、フォルダがないときだけ保存先を尋ねる）。 |
| `lib/sync/fsutil.js` | `src/sync/fsutil.ts` | sync で使うファイル操作の小道具。 |
| `lib/sync/lock.js` | `src/sync/lock.ts` | プロセス（ウィンドウ）をまたぐロック。 |
