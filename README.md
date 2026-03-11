# TALO 廻り縁拾い 3Dビューア

IFCファイルから廻り縁（crown molding）の必要数量を自動算出し、3Dビューア付きHTMLをGitHub Pagesに公開します。

## 仕組み

1. IFCファイルをリポジトリにpush
2. GitHub Actionsが自動実行（Python + ifcopenshell）
3. 3DビューアHTMLがGitHub Pagesにデプロイ

## セットアップ

### 1. リポジトリ作成

```bash
git clone <このリポジトリ>
cd mawari-buchi-viewer
```

### 2. GitHub Pages を有効化

Settings → Pages → Source: **GitHub Actions** を選択

### 3. IFCファイルを追加してpush

```bash
cp /path/to/your-model.ifc .
git add your-model.ifc
git commit -m "Add IFC file"
git push
```

### 4. 結果を確認

- Actions タブでビルド状況を確認
- 完了後、`https://<user>.github.io/<repo>/` でビューアが表示されます

## 手動実行

Actions タブ → 「IFC → 廻り縁拾い3Dビューア」→ Run workflow

特定のファイルを指定する場合は `ifc_file` パラメータにファイル名を入力。

## ローカル実行

```bash
pip install -r requirements.txt
python scripts/build_mawari_buchi_viewer.py input.ifc output/index.html
open output/index.html
```

## ファイル構成

```
├── .github/workflows/process-ifc.yml  # GitHub Actions ワークフロー
├── scripts/
│   └── build_mawari_buchi_viewer.py   # メイン処理スクリプト
├── requirements.txt                    # Python依存パッケージ
├── *.ifc                              # IFCファイル（ユーザーが追加）
└── README.md
```

## 廻り縁タイプ

| タイプ | 使用箇所 | 3Dビューア色 |
|--------|----------|-------------|
| 廻り縁１ | 1F天井×ログ壁、1F天井×梁(L345)、2F勾配天井×ログ壁(斜め) | オレンジレッド |
| 廻り縁２ | 1F天井×間仕切壁、2F水平天井×壁/梁(勾配上側)、2F天井×棟木 | 水色 |
| 廻り縁３ | 2F水平天井×壁/梁(勾配下側) | ライムグリーン |
