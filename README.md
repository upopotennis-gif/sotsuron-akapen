# 卒論 赤ペンチェック

卒業論文のWordファイル（.docx）を読み、「レイアウトフォーマット（2026）」と「卒業論文及びアブストラクトを書くにあたって」に合っているかを判定して、赤ペンの講評とPDFを返すページ。ファイルはブラウザの中だけで読み、外部に送らない。

- `engine.js` … docx を読んで判定（基準は先頭の `RULE`、各チェックは `analyze()` の中）
- `ui.js` … 画面・赤ペンシート・PDF（html2canvas＋jsPDF）・一括ZIP・CSV
- `index.html` … 見た目
- 動作確認用の架空の論文は `python3 tools/make_samples.py` で `samples/` に作る（公開しない）
