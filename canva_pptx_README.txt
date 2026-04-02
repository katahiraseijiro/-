================================================================================
CSV → PowerPoint（pptx）→ Canva：手順
================================================================================

■ 用意するもの
  - Python 3.10 以降（インストール時に「Add Python to PATH」にチェック推奨）
  - 本フォルダの csv_to_pptx.py / requirements-pptx.txt

■ 1回だけ：ライブラリ導入
  コマンドプロンプトまたは PowerShell で本フォルダへ移動し：

    python -m pip install -r requirements-pptx.txt

■ pptx を作る
  （Canva用に作った bulk_only CSV を指定）

    python csv_to_pptx.py -i canva_part16_最幸ママ_bulk_only.csv -o output_part16.pptx --title "最幸ママ・アカデミー"

  表紙が不要なら --title を省略。

  PowerShell ラッパー例：

    .\run_csv_to_pptx.ps1 -InputCsv ".\canva_part19_最幸ママ_bulk_only.csv" -Output ".\output_part19.pptx" -Title "最幸ママ・パート19"

■ Canva へ載せる
  1. Canva で「作成」→「インポート」または「ファイルをアップロード」で pptx を読み込む
  2. 読み込み後、フォント・色・ロゴをブランドに合わせて一括調整
  3. 画像・図は VisualIdea 列を見ながら手動で差し込み（または別素材）

■ CSV の列名
  Headline, Subheadline, Body, CTA, Hook, VisualIdea
  （canva_*_bulk_only.csv と同じ）

■ トラブル
  - python が認識されない → PATH を確認するか、Microsoft Store の Python ではなく python.org 版を推奨
  - 文字化け → CSV を UTF-8（BOM 可）で保存し直す

================================================================================
