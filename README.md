# Canva 用スライド素材（CSV → PowerPoint）

ウェビナー用のストーリーライン・スライド原稿を **CSV** で管理し、`python-pptx` で **PowerPoint（.pptx）** を生成します。生成したファイルは **Canva にインポート**してデザイン調整する想定です（Canva API は使用しません）。

同じリポジトリの **`web/`** に **Next.js**（App Router・TypeScript・Tailwind CSS・ESLint）を用意しています。ルート直下は既存の CSV / Python と共存させるため、npm の命名規則（フォルダ名に大文字を含むとプロジェクト名に使えない）にも合わせ **`web/` サブフォルダ**に初期化しています。

## Next.js（`web/`）

```powershell
cd web
npm install
npm run dev
```

- 開発サーバー: 通常は [http://localhost:3000](http://localhost:3000)  
- 本番ビルド: `npm run build` → `npm run start`

ソースは `web/app/` 以下（App Router）です。

## 必要な環境

- **Python 3.10 以降**（インストール時に「Add Python to PATH」にチェック推奨）
- 本リポジトリの `csv_to_pptx.py` と `requirements-pptx.txt`

## セットアップ（初回のみ）

プロジェクトフォルダで:

```powershell
python -m pip install -r requirements-pptx.txt
```

依存パッケージは Python の `site-packages` に入ります（このフォルダ直下には置かれません）。

## pptx の生成

### Python から直接

```powershell
python csv_to_pptx.py -i canva_part16_最幸ママ_bulk_only.csv -o output_part16.pptx --title "最幸ママ・アカデミー"
```

表紙スライドが不要な場合は `--title` を省略します。

```powershell
python csv_to_pptx.py -i canva_part19_最幸ママ_bulk_only.csv -o output_part19.pptx --title "最幸ママ・パート19"
```

### PowerShell ラッパー

依存関係のインストールと実行をまとめて行う場合:

```powershell
.\run_csv_to_pptx.ps1 -InputCsv ".\canva_part19_最幸ママ_bulk_only.csv" -Output ".\output_part19.pptx" -Title "最幸ママ・パート19"
```

## CSV の列

次の列名を想定しています（`canva_*_bulk_only.csv` と同じ形式）。

| 列名 | 用途の目安 |
|------|------------|
| `Headline` | スライドタイトル |
| `Subheadline` | 本文上部 |
| `Body` | 本文 |
| `CTA` | 行動喚起（本文に追記） |
| `Hook` | フック（本文に追記） |
| `VisualIdea` | ビジュアル案（本文に追記） |

CSV は **UTF-8**（BOM 可）で保存してください。

## Canva への載せ方

1. Canva で「作成」→「インポート」またはファイルをアップロードし、**.pptx** を読み込む  
2. 読み込み後、フォント・色・ロゴをブランドに合わせて調整  
3. 画像・図は `VisualIdea` を参照しながら差し替え

詳しい運用メモは `canva_pptx_README.txt` および `canva_300slides_workflow.txt` を参照してください。

## リポジトリ内の主なファイル

| 種別 | ファイル例 |
|------|------------|
| Next.js アプリ | `web/`（`package.json`, `web/app/`） |
| 生成スクリプト | `csv_to_pptx.py`, `run_csv_to_pptx.ps1` |
| 依存定義 | `requirements-pptx.txt` |
| マスター・マッピング | `webinar_slide_map_22rows.csv`, `webinar_22slides_master_spec.txt` |
| スライドモデル（テキスト） | `webinar_slide_model_*.txt` |
| 最幸ママ用 CSV | `canva_part16_最幸ママ_bulk_only.csv`, `canva_part19_最幸ママ_bulk_only.csv` など |

生成した **`.pptx` は `.gitignore` で除外**しています（バイナリの肥大化を避けるため）。

## トラブルシューティング

- **`python` が認識されない** — PATH を確認するか、[python.org](https://www.python.org/downloads/) の公式インストーラで入れ直す  
- **文字化け** — CSV を UTF-8 で保存し直す  
- **`pip` が使えない** — `python -m pip install -r requirements-pptx.txt` を試す

## ライセンス

リポジトリ利用者の判断に委ねます（未設定の場合はリポジトリ所有者の方針に従ってください）。
