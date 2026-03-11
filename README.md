# TALO IFC 積算システム

TALOログハウスのIFCファイル（ArchiCAD 22 JPN / IFC2X3）から部材を自動抽出し、積算用Excel・3Dビューアを生成するシステム。

## 機能一覧

| 機能 | スクリプト | 出力 |
|------|-----------|------|
| 部材抽出・積算Excel | `extract_ifc.py` | 部材一覧Excel + 3Dビューア |
| 額縁拾い3Dビューア | `build_gakubuchi_viewer.py` | 額縁部材一覧 + 3Dビューア |
| 廻り縁拾い3Dビューア | `build_mawari_buchi_viewer.py` | 廻り縁部材一覧 + 3Dビューア |
| 巾木拾い3Dビューア | `build_habaki_viewer.py` | 巾木部材一覧 + 3Dビューア |
| 集計シート統合 | `add_sheets_to_excel.py` | 上記JSONを部材一覧Excelに統合 |
| 額縁拾いHTMLツール | `pages/額縁拾い.html` | PDF取込→額縁計算→木取り最適化 |

## 仕組み

1. IFCファイルをリポジトリにpush
2. GitHub Actionsが自動実行（変更IFCのみ差分処理）
3. 積算Excel・3DビューアHTMLがGitHub Pagesにデプロイ

## セットアップ

### GitHub Pages を有効化

Settings → Pages → Source: **GitHub Actions** を選択

### IFCファイルを追加してpush

```bash
cp /path/to/model.ifc .
git add model.ifc
git commit -m "Add IFC file"
git push
```

### 結果を確認

- Actionsタブでビルド状況を確認
- 完了後 `https://adachitalo.github.io/talosekisan/` で各ビューアにアクセス

## 手動実行

Actions タブ → 「IFC 積算処理」→ Run workflow

## ローカル実行

```bash
pip install -r requirements.txt
python scripts/extract_ifc.py model.ifc output/
```

## ファイル構成

```
├── .github/workflows/process-ifc.yml   # GitHub Actions（差分処理対応）
├── scripts/
│   ├── extract_ifc.py                  # 部材抽出・積算Excel
│   ├── build_gakubuchi_viewer.py       # 額縁拾い3Dビューア
│   ├── build_mawari_buchi_viewer.py    # 廻り縁拾い3Dビューア
│   ├── build_habaki_viewer.py          # 巾木拾い3Dビューア
│   └── add_sheets_to_excel.py          # 集計シート統合
├── pages/
│   └── 額縁拾い.html                    # 額縁拾い・木取り計算ツール
├── requirements.txt                     # Python依存パッケージ
├── *.ifc                               # IFCファイル（ユーザーが追加）
└── README.md
```

## 額縁拾いHTMLツール

ブラウザ単体で動作する額縁拾い・木取り計算ツール（サーバー不要）。

- PDFインポート（ARCHICAD建具表PDF対応）または建具マスタから手動選択
- IFCビューア準拠の表裏ルール（表面=フルセット、裏面=額縁のみ）
- 額縁3方/4方の切替、間仕切壁対応
- First Fit Decreasing法による木取り最適化
- 印刷機能付き

## 廻り縁タイプ

| タイプ | 使用箇所 |
|--------|----------|
| 廻り縁１ | 1F天井×ログ壁、1F天井×梁(L345)、2F勾配天井×ログ壁(斜め) |
| 廻り縁２ | 1F天井×間仕切壁、2F水平天井×壁/梁(勾配上側)、2F天井×棟木 |
| 廻り縁３ | 2F水平天井×壁/梁(勾配下側) |
