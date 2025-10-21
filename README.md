# STREAMED → freee 変換ツール

STREAMED（経費精算システム）からダウンロードしたCSVを、freee会計ソフト用のExcel形式に変換するWebアプリケーションです。

## 🌟 特徴

- ✅ **複数ファイル一括処理** - 一度に複数のSTREAMED CSVを処理可能
- ✅ **部署名の自動正規化** - 表記ゆれを統一
- ✅ **取引先名の自動照合** - freeeの取引先マスタと照合
- ✅ **設定ファイルアップロード対応** - 固定設定ファイルなしでも動作
- ✅ **リアルタイム進捗表示** - 処理状況を視覚的に確認
- ✅ **Webアプリ** - インストール不要、ブラウザだけで利用可能

## 🚀 使い方（ユーザー向け）

### 1. アプリにアクセス
デプロイされたStreamlit CloudのURLにアクセス

### 2. 入力形式を選択
サイドバーで以下から選択：
- **test** - テスト用Excel
- **freee** - freee形式Excel
- **streamed** - STREAMED CSV（推奨）

### 3. 設定ファイルをアップロード（STREAMEDの場合・任意）
- 部署マッピングExcel
- 取引先一覧Excel

※アップロードしない場合は、configフォルダの固定ファイルを使用

### 4. データファイルをアップロード
- STREAMED CSVファイル（複数可）
- freee取引先CSV（STREAMEDの場合は必須）

### 5. 出力形式を選択
- **test** - テスト用Excel
- **freee** - freee用Excel（推奨）

### 6. 実行
「実行」ボタンをクリック

### 7. ダウンロード
- 個別ダウンロード
- 一括ZIPダウンロード（複数ファイルの場合）

## 💻 開発者向け

### ローカル環境でのセットアップ

#### 1. リポジトリをクローン
```bash
git clone https://github.com/Nebo-git/ricoSTREAMED_project01.git
cd ricoSTREAMED_project01
```

#### 2. 必要なパッケージをインストール
```bash
pip install -r requirements.txt
```

#### 3. アプリを起動
```bash
streamlit run streamlit_app.py
```

ブラウザが自動で開きます（開かない場合は `http://localhost:8501` にアクセス）

### ファイル構成

```
ricoSTREAMED_project01/
├── streamlit_app.py          # メインアプリケーション
├── requirements.txt          # 必要なPythonパッケージ
├── README.md                 # このファイル
├── ARCHITECTURE.md           # システム構成の詳細
├── config/                   # 設定ファイル（デフォルト用）
│   ├── dept_mapping.xlsx     # 部署名マッピング
│   └── partner_list.xlsx     # 取引先一覧
├── reader/                   # 入力ファイル読み込みモジュール
│   ├── __init__.py
│   ├── test_reader01.py
│   ├── freee_reader.py
│   └── rico_streamed_csvreader.py
├── processor/                # データ処理・変換モジュール
│   ├── __init__.py
│   ├── dept_normalizer.py
│   ├── partner_resolver.py
│   └── voucher_formatter.py
└── exporter/                 # 出力ファイル生成モジュール
    ├── __init__.py
    └── freee_exporter.py
```

詳細は [ARCHITECTURE.md](ARCHITECTURE.md) を参照してください。

## 🌐 デプロイ（Streamlit Cloud）

### 前提条件
- GitHubアカウント
- Streamlit Cloudアカウント（GitHubで認証可能）

### 手順

#### 1. GitHubリポジトリを作成（初回のみ）
1. GitHub（https://github.com）にログイン
2. 「New repository」をクリック
3. リポジトリ名を入力
4. **Public** を選択（Streamlit Cloud無料プランの場合）
5. 「Create repository」をクリック

#### 2. コードをpush
```bash
git init
git add .
git commit -m "初回コミット"
git branch -M main
git remote add origin https://github.com/ユーザー名/リポジトリ名.git
git push -u origin main
```

#### 3. Streamlit Cloudでデプロイ
1. https://streamlit.io/cloud にアクセス
2. GitHubアカウントでログイン
3. 「New app」をクリック
4. リポジトリを選択
5. Main file path: `streamlit_app.py`
6. 「Deploy!」をクリック

数分後、URLが発行されます！

## 🔧 技術スタック

- **Python 3.x**
- **Streamlit** - Webアプリケーションフレームワーク
- **pandas** - データ処理
- **openpyxl** - Excel読み書き

## 📋 システム要件

### Streamlit Cloud（推奨）
- メモリ: 1GB RAM
- ストレージ: 無制限
- 実行時間: 無制限

### ローカル環境
- Python 3.8以上
- メモリ: 512MB以上推奨

## ⚠️ 制限事項

- **ファイルサイズ**: 1ファイル200MB未満推奨
- **同時処理**: 無料プランでは同時アクセス数に制限なし（ただしリソース共有）
- **設定ファイル**: configフォルダの固定ファイルは機密情報を含むため、GitHubに含めないことを推奨

## 🐛 トラブルシューティング

### エラーが出る場合
- Streamlit Cloud管理画面のログを確認
- ローカルで `streamlit run streamlit_app.py` を実行して動作確認

### ファイルが大きすぎる
- 1ファイル200MB未満に分割
- 不要な列を削除してからアップロード

### 処理が遅い
- 無料プランのため、大量ファイルは時間がかかります
- ファイル数を減らして複数回に分けて実行

### 設定ファイルが読み込めない
- アップロードした設定ファイルの形式を確認
- configフォルダのサンプルファイルと同じ形式か確認

## 📞 サポート

質問や問題があれば、GitHubのIssuesで報告してください。

## 📝 ライセンス

このプロジェクトは私的利用を目的としています。

## 🙏 謝辞

- Streamlit Community
- freee API ドキュメント
- STREAMED サポートチーム