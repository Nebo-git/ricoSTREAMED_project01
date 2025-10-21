# Excel to CSV Converter - Streamlit版

## 🚀 ローカルでテスト

### 1. 必要なパッケージをインストール
```bash
pip install -r requirements.txt
```

### 2. アプリを起動
```bash
streamlit run streamlit_app.py
```

ブラウザが自動で開きます（開かない場合は `http://localhost:8501` にアクセス）

---

## 📦 デプロイ手順（Streamlit Cloud）

### Phase 2: GitHubにアップロード

#### 1. GitHubリポジトリを作成
1. GitHub（https://github.com）にログイン
2. 「New repository」をクリック
3. リポジトリ名: `excel-csv-converter`（任意）
4. Public を選択
5. 「Create repository」をクリック

#### 2. ファイルをアップロード

**アップロードするファイル：**
```
excel-csv-converter/
├── streamlit_app.py          ← メインアプリ
├── requirements.txt          ← パッケージリスト
├── config/
│   ├── dept_mapping.xlsx
│   └── partner_list.xlsx
├── reader/
│   ├── __init__.py
│   ├── test_reader01.py
│   ├── freee_reader.py
│   └── rico_streamed_csvreader.py
├── processor/
│   ├── __init__.py
│   ├── config.py
│   ├── dept_normalizer.py
│   ├── partner_resolver.py
│   └── voucher_formatter.py
└── exporter/
    ├── __init__.py
    └── freee_exporter.py
```

**アップロード方法（GitHub Web UI）：**
1. リポジトリページで「Add file」→「Upload files」
2. 上記のファイル・フォルダをドラッグ&ドロップ
3. 「Commit changes」をクリック

---

### Phase 3: Streamlit Cloudでデプロイ

#### 1. Streamlit Cloudにアクセス
https://streamlit.io/cloud にアクセス

#### 2. サインアップ/ログイン
- GitHubアカウントで認証

#### 3. 新しいアプリをデプロイ
1. 「New app」をクリック
2. リポジトリを選択: `excel-csv-converter`
3. Main file path: `streamlit_app.py`
4. 「Deploy!」をクリック

#### 4. デプロイ完了！
数分後、URLが発行されます（例: `https://your-app.streamlit.app`）

---

## 🎨 使い方

### 基本操作
1. **サイドバー**で入力形式・出力形式を選択
2. **ファイルをアップロード**（複数可）
3. **STREAMEDの場合**: freee取引先CSVもアップロード
4. **「実行」ボタン**をクリック
5. **ダウンロード**ボタンで結果を取得

### 特徴
- ✅ 複数ファイル一括処理
- ✅ リアルタイム進捗表示
- ✅ エラー詳細表示
- ✅ 個別ファイルダウンロード
- ✅ モダンで見やすいUI

---

## 💡 Tips

### 無料プランの制限
- リソース: 1GB RAM
- 同時アクセス: 制限なし
- 実行時間: 無制限

### トラブルシューティング
- **エラーが出る場合**: ログを確認（Streamlit Cloud管理画面）
- **ファイルが大きすぎる**: 1ファイル200MB未満推奨
- **処理が遅い**: 無料プランのため、大量ファイルは時間がかかる

---

## 📞 サポート

質問や問題があれば、GitHubのIssuesで報告してください。