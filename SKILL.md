---
name: pptx-extended-comments
description: Use image attachments in PPTX review comments to present revision proposals, before-and-after comparisons, and results from parallel slide editing. Applies to the extended image comments supported by Google Slides Sync, not general slide authoring.
---

# PPTX拡張コメント

画像付きコメントを使い、スライドの修正提案を見た目と一緒に伝える。スライドの編集・画像への描画には、その環境で利用できるスライドスキルやツールを使う。

このフォルダの `SKILL.md` と `scripts/` 一式を一緒に配置する。Node.js 20以上で動作し、CLIの実行にVS Codeや追加のnpmパッケージは不要。`scripts/` はGoogle Slides Syncのソースからビルドしたもので、直接は編集しない。各ファイルの役割は `scripts/README.md` にある。

## 画像付きコメントの使用

1. 対象ページと提案内容を確認し、比較に使う画像をPNG/JPEGで用意する。変更前後は同じ寸法に揃え、文字が読める大きさにする。
2. このスキルの `scripts/slides-comments.js` を使う。実行時は、この `SKILL.md` の場所からCLIの絶対パスを求める。`outline` と `list --json` で対象ページ・要素・既存コメントを確認し、既存のスレッドを使うか、`add` で提案用のコメントを作る。
3. `attach` で画像を付ける。本文には画像の順番と変更理由を書く。投稿後は `list --json` で対象ページ・本文・画像を読み戻す。失敗した添付だけを再実行する。

```sh
node /absolute/path/pptx-extended-comments/scripts/slides-comments.js add deck.pptx --slide 3 --text '見出しを結論に変える案。画像は変更前、修正案の順。' --author Reviewer
# CLIのパスと作成者は実際の値にし、THREADにはaddが返したスレッドIDを入れる。
node /absolute/path/pptx-extended-comments/scripts/slides-comments.js attach deck.pptx --thread 'THREAD' --image before.png --author Reviewer
node /absolute/path/pptx-extended-comments/scripts/slides-comments.js attach deck.pptx --thread 'THREAD' --image proposal.png --author Reviewer
```

既存の返信への添付は `attach --thread THREAD --reply REPLY --image file.png --author Reviewer`。作成者には自分の名前またはエージェント名を使い、他の人のコメント・返信を保持する。画像を作成・添付できていない場合は、その旨を報告する。

## ユースケース

- **修正提案**：見出し、図表、レイアウトの修正案を画像で添付し、提案内容を具体的に示す。レビューだけの依頼では元のスライド内容を変更しない。
- **変更前後の比較**：同じコメントに変更前と修正後を添付し、何が変わったか確認できるようにする。
- **並列で作った案の比較**：基準版を固定し、並列実行が利用可能で許可されていれば、作業者ごとに別のコピーと担当ページ・案を渡す。各担当が編集・描画した結果を集め、一人の担当が案A・案Bの画像を順番にコメントへ添付する。同じpptxへ同時保存しない。採用案を反映する際は元の版を再確認し、依頼の範囲で変更を統合して描画し直す。ファイルロックは変更内容の自動マージではない。

## 画像付きコメントの仕様

- 一投稿4枚まで。PNG/JPEGで各2MBまで、各辺8192px以下・1600万画素以下。さらに必要なら同じスレッドに自分の返信を作り、そこへ添付する。
- 画像のdata URL（Base64）と名前・IDを、pptx内のコメントXMLの `p188:extLst` に独自URI `urn:google-slides-sync:comment-images:v1` の拡張として保存する。外部画像ファイルの参照ではなく、外部フォルダへの自動書き出しも行わない。
- 画像表示はGoogle Slides Syncのエディタで対応する。PowerPoint標準の画像コメントとの互換は主張しない。GoogleへのPush・PDF出力用の送信データからは添付画像を含むローカルコメントを除去し、Pullでは手元のコメントと画像を引き継ぐ。
