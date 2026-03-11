# TALO IFC 積算システム

TALOログハウスのIFCファイル（ArchiCAD 22 JPN / IFC2X3）から部材を自動抽出し、積算用Excel・3Dビューア・キット積算Excelを生成するシステム。

GitHub Pages: https://adachitalo.github.io/talosekisan/

## 機能一覧

| 機能 | スクリプト | 出力 |
|------|-----------|------|
| 部材抽出・積算Excel | `extract_ifc.py` | 部材一覧Excel（分類別集計付き） |
| キット積算Excel | `extract_ifc.py` | キット価格計算Excel（HI/IE/LO/カスタム自動判別） |
| 額縁拾い3Dビューア | `build_gakubuchi_viewer.py` | 額縁部材一覧 + 3Dビューア + 寸法チェック |
| 廻り縁拾い3Dビューア | `build_mawari_buchi_viewer.py` | 廻り縁部材一覧 + 3Dビューア |
| 巾木拾い3Dビューア | `build_habaki_viewer.py` | 巾木部材一覧 + 3Dビューア |
| 集計シート統合 | `add_sheets_to_excel.py` | 上記JSONを部材一覧Excelに統合 |
| 額縁拾い・木取り計算 | `pages/gakubuchi.html` | PDF取込→額縁計算→木取り最適化→寸法チェック |

## 仕組み

1. IFCファイルをリポジトリにpush（またはWebからアップロード）
2. GitHub Actionsが自動実行（変更IFCのみ差分処理、キャッシュ対応）
3. 部材一覧Excel・キット積算Excel・3DビューアHTMLがGitHub Pagesにデプロイ

## セットアップ

### GitHub Pages を有効化

Settings → Pages → Source: **GitHub Actions** を選択

### IFCファイルを追加

方法1: Webアップロード（https://adachitalo.github.io/talosekisan/upload.html）

方法2: gitでpush
```bash
cp /path/to/model.ifc ifc/
git add ifc/model.ifc
git commit -m "Add IFC file"
git push
```

### 結果を確認

- Actionsタブでビルド状況を確認
- 完了後 https://adachitalo.github.io/talosekisan/ で各ビューア・Excelにアクセス

## 手動実行

Actions タブ → 「IFC → 積算ツール」→ Run workflow

## ローカル実行

```bash
pip install -r requirements.txt
python scripts/extract_ifc.py model.ifc output/buzai.xlsx output/kit.xlsx
```

## ファイル構成

```
├── .github/workflows/process-ifc.yml   # GitHub Actions（差分処理+キャッシュ対応）
├── scripts/
│   ├── extract_ifc.py                  # 部材抽出・積算Excel・キット積算Excel
│   ├── build_gakubuchi_viewer.py       # 額縁拾い3Dビューア（寸法チェック付き）
│   ├── build_mawari_buchi_viewer.py    # 廻り縁拾い3Dビューア
│   ├── build_habaki_viewer.py          # 巾木拾い3Dビューア
│   └── add_sheets_to_excel.py          # 集計シート統合
├── pages/
│   ├── gakubuchi.html                  # 額縁拾い・木取り計算ツール（寸法チェック付き）
│   └── upload.html                     # IFCファイルWebアップロード
├── ifc/                                # IFCファイル格納フォルダ
├── requirements.txt                     # Python依存パッケージ
└── README.md
```

## 額縁拾い・木取り計算ツール（gakubuchi.html）

ブラウザ単体で動作する額縁拾い・木取り計算ツール（サーバー不要）。

- PDFインポート（ARCHICAD建具表PDF対応）または建具マスタ（94件）から手動選択
- IFCビューア準拠の表裏分離ルール（表面=額縁+補助部材、裏面=額縁のみ）
- 額縁3方/4方の切替、間仕切壁対応（T-bar自動連動）
- 建具の分割機能（同一建具を個別設定用に分割可能）
- VMW/VMSD規格寸法チェック（CAD入力ミスを自動検出・推定正解を提示）
- First Fit Decreasing法による木取り最適化
- 印刷機能付き（木取り根拠の表示/非表示切替可）

## キット積算Excel

IFCファイルから抽出した部材数量をキット積算テンプレートに自動入力。

- テンプレートExcelはスクリプトに内蔵（base64埋め込み、外部ファイル不要）
- モデル名からHI/IE/LO系を自動判別し、対応シートに入力（それ以外はカスタムシート）
- 代理店定価・CP定価・エンドユーザー定価の3段階価格を自動計算

## VMW/VMSD 規格寸法チェック

PDF取り込み（gakubuchi.html）とIFC処理（build_gakubuchi_viewer.py）の両方で、
マーカー名と実寸法を規格値と照合し、不一致があれば警告を表示する。

対応品番: VMW-1〜VMW-13b（15品番）、VMSD（3サイズ）

- gakubuchi.html: 赤枠の警告パネル + テーブル行ハイライト
- 3Dビューア: 画面中央の警告ダイアログ
- JSON集計: dim_errors フィールドで部材一覧Excel統合時にも参照可能

## 廻り縁タイプ

| タイプ | 使用箇所 |
|--------|----------|
| 廻り縁１ | 1F天井×ログ壁、1F天井×梁(L345)、2F勾配天井×ログ壁(斜め) |
| 廻り縁２ | 1F天井×間仕切壁、2F水平天井×壁/梁(勾配上側)、2F天井×棟木 |
| 廻り縁３ | 2F水平天井×壁/梁(勾配下側) |
