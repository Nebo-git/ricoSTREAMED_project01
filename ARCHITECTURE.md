# システム構成（ARCHITECTURE）

このドキュメントでは、STREAMED → freee 変換ツールのシステム構成、データフロー、各モジュールの役割を説明します。

## 📊 全体アーキテクチャ

```
┌─────────────────┐
│  Streamlit UI   │ ← ユーザーインターフェース
└────────┬────────┘
         │
         ↓
┌─────────────────┐
│ streamlit_app.py│ ← メインコントローラー
└────────┬────────┘
         │
    ┌────┴────┬───────────┬──────────┐
    │         │           │          │
    ↓         ↓           ↓          ↓
┌────────┐ ┌─────────┐ ┌────────┐ ┌────────┐
│ Reader │ │Processor│ │Exporter│ │ Config │
└────────┘ └─────────┘ └────────┘ └────────┘
```

## 🔄 データフロー

### 全体の流れ

```
1. ファイルアップロード
   ↓
2. Reader: ファイル読み込み・検証
   ↓
3. Processor: データ変換処理
   ├─ DeptNormalizer: 部署名正規化
   ├─ PartnerResolver: 取引先照合
   └─ VoucherFormatter: freee形式整形
   ↓
4. Exporter: Excel出力
   ↓
5. ダウンロード
```

### 詳細フロー（STREAMED → freee）

```
[STREAMED CSV] + [freee取引先CSV] + [設定ファイル（任意）]
        ↓
┌──────────────────────────────────┐
│  RicoStreamedCSVReader           │
│  - CSV読み込み                    │
│  - データ検証                     │
└──────────┬───────────────────────┘
           ↓
    [data_list: List[Dict]]
           ↓
┌──────────────────────────────────┐
│  DeptNormalizer                  │
│  - 部署名の表記ゆれを統一         │
│  - dept_mapping.xlsx 参照        │
└──────────┬───────────────────────┘
           ↓
    [正規化されたdata_list]
           ↓
┌──────────────────────────────────┐
│  PartnerResolver                 │
│  - 取引先名をfreee IDに変換       │
│  - partner_list.xlsx 参照        │
│  - freee取引先CSVで補完          │
└──────────┬───────────────────────┘
           ↓
    [取引先解決済みdata_list]
           ↓
┌──────────────────────────────────┐
│  VoucherFormatter                │
│  - freee仕訳形式に整形            │
│  - 必須項目の補完                 │
└──────────┬───────────────────────┘
           ↓
    [freee形式data_list]
           ↓
┌──────────────────────────────────┐
│  FreeeExcelExporter              │
│  - Excelファイル生成              │
└──────────┬───────────────────────┘
           ↓
    [freee用Excel]
```

## 📁 モジュール詳細

### 1. streamlit_app.py（メインアプリケーション）

**役割**: UIとビジネスロジックの統合

**主な機能**:
- ユーザーインターフェースの構築
- ファイルアップロード処理
- 各モジュールの呼び出しと連携
- エラーハンドリング
- 結果の表示とダウンロード提供

**主な関数**:
- `main()` - アプリのエントリーポイント
- `process_files()` - ファイル処理のメインループ
- `show_results()` - 処理結果の表示

---

### 2. reader/ （入力モジュール）

各種形式のファイルを読み込み、統一されたデータ構造に変換します。

#### 2-1. test_reader01.py
**クラス**: `TestExcelReader`

**役割**: テスト用Excel（B2:日付, C3:金額形式）を読み込み

**入力**: 
- Excelファイル（.xlsx, .xls）

**出力**: 
- `data_list: List[Dict]` - 統一された辞書のリスト
- `errors: List[str]` - エラーメッセージのリスト

**処理内容**:
1. B2セルから日付を読み取り
2. C3セルから金額を読み取り
3. データ検証（日付形式、金額の数値チェック）
4. 辞書形式に変換

---

#### 2-2. freee_reader.py
**クラス**: `FreeeExcelReader`

**役割**: freee形式Excelを読み込み

**入力**: 
- freee形式Excelファイル

**出力**: 
- `data_list: List[Dict]`
- `errors: List[str]`

**処理内容**:
1. freee項目名を検索
2. データ行を抽出
3. 各列を辞書のキーにマッピング

---

#### 2-3. rico_streamed_csvreader.py
**クラス**: `RicoStreamedCSVReader`

**役割**: STREAMED CSVを読み込み

**入力**: 
- STREAMED CSVファイル

**出力**: 
- `data_list: List[Dict]`
- `errors: List[str]`

**処理内容**:
1. CSVをpandasで読み込み
2. STREAMED特有の列名を標準化
3. 日付・金額のフォーマット検証
4. 必須項目のチェック

**主な項目**:
- 日付
- 取引先名
- 部署名
- 金額
- 摘要
- など

---

### 3. processor/ （処理モジュール）

データの変換・正規化・整形を行います。

#### 3-1. dept_normalizer.py
**クラス**: `DeptNormalizer`

**役割**: 部署名の表記ゆれを統一

**設定ファイル**: `config/dept_mapping.xlsx`

**フォーマット**:
| 入力部署名 | 出力部署名 |
|-----------|-----------|
| 営業部 | 営業部 |
| 営業 | 営業部 |
| Sales | 営業部 |

**処理内容**:
1. data_listの各レコードから部署名を取得
2. dept_mapping.xlsxで定義されたマッピングを適用
3. マッピングに存在しない場合は元の名前をそのまま使用

**メソッド**:
- `__init__(dept_mapping_path)` - マッピングファイルを読み込み
- `normalize(data_list)` - 部署名を正規化

---

#### 3-2. partner_resolver.py
**クラス**: `PartnerResolver`

**役割**: 取引先名をfreeeの取引先IDに変換

**設定ファイル**: 
- `config/partner_list.xlsx` - 優先
- freee取引先CSV - 補助

**処理ロジック**:
1. partner_list.xlsxで取引先名を検索
2. 見つからない場合、freee取引先CSVで検索
3. 見つからない場合、エラーまたは空白

**メソッド**:
- `__init__(partner_list_path, freee_csv_path)` - 設定ファイル読み込み
- `resolve(data_list)` - 取引先名を解決

---

#### 3-3. voucher_formatter.py
**クラス**: `VoucherFormatter`

**役割**: freee会計の仕訳形式に整形

**処理内容**:
1. freee必須項目の追加
   - 収支区分
   - 決済期日
   - 口座区分
   - 税区分
   など
2. 日付フォーマットの統一（YYYY/MM/DD）
3. 金額のフォーマット（数値型）
4. 空白項目の補完

**入力**: STREAMED形式のdata_list

**出力**: freee形式のdata_list

**主な変換**:
- STREAMED列名 → freee列名
- 部署名 → 部門
- 取引先名 → 取引先コード

**メソッド**:
- `__init__(format_type)` - フォーマットタイプを指定
- `format(data_list)` - データを整形

---

### 4. exporter/ （出力モジュール）

処理済みデータをExcelファイルとして出力します。

#### 4-1. freee_exporter.py

**クラス**: 
- `TestExcelExporter` - テスト用Excel出力
- `FreeeExcelExporter` - freee用Excel出力

**役割**: データをExcelファイルに書き出し

**入力**: 
- `data_list: List[Dict]`
- `filename: str` - 元のファイル名

**出力**: 
- Excelファイルのパス

**処理内容**:
1. pandasでDataFrameを作成
2. openpyxlでExcel書き込み
3. 列幅の自動調整
4. ヘッダー行のスタイリング（任意）

**freee用の列順**:
- 取引日
- 決済期日
- 収支区分
- 取引先コード
- 取引先名
- 部門
- 勘定科目
- 税区分
- 金額
- 摘要
- など

**メソッド**:
- `__init__(output_dir)` - 出力ディレクトリを指定
- `export(data_list, filename)` - Excelファイルを生成

---

### 5. config/ （設定ファイル）

#### 5-1. dept_mapping.xlsx
**形式**: Excel（2列）

**内容**:
- 列A: 入力部署名（STREAMED上の表記）
- 列B: 出力部署名（freeeでの統一表記）

**例**:
```
入力部署名 | 出力部署名
---------|----------
営業部    | 営業部
営業      | 営業部
Sales    | 営業部
経理部    | 管理部
```

**用途**: 部署名の表記ゆれを吸収

---

#### 5-2. partner_list.xlsx
**形式**: Excel（複数列）

**必須列**:
- 取引先名
- 取引先コード（freee ID）

**任意列**:
- 正式名称
- 略称
- その他のメタデータ

**例**:
```
取引先名       | 取引先コード | 正式名称
-------------|------------|------------
株式会社ABC   | 12345      | 株式会社ABC商事
ABC          | 12345      | 株式会社ABC商事
```

**用途**: 取引先名の名寄せとfreee IDの解決

---

## 🔧 データ構造

### data_list の形式

各モジュール間で受け渡される `data_list` は以下の形式です：

```python
data_list = [
    {
        "日付": "2025-10-21",
        "取引先名": "株式会社ABC",
        "部署名": "営業部",
        "金額": 10000,
        "摘要": "交通費",
        # ... その他の項目
    },
    {
        "日付": "2025-10-22",
        # ...
    }
]
```

### エラーハンドリング

各Readerは `(data_list, errors)` のタプルを返します：

```python
data_list, errors = reader.read_and_validate()

# errors の例
errors = [
    "行3: 日付が不正です",
    "行5: 金額が数値ではありません",
]
```

---

## 🚦 処理フロー制御

### エラー発生時
- エラーは `errors` リストに蓄積
- 処理は継続（可能な限り）
- 最終的にUIでエラー一覧を表示

### 設定ファイル不在時
- アップロードされた設定ファイルを優先使用
- なければ `config/` フォルダのデフォルトファイルを使用
- デフォルトファイルもなければ空の辞書で動作

---

## 🔐 セキュリティ考慮事項

### 機密情報の扱い
- `config/dept_mapping.xlsx` - 部署情報（機密性: 低）
- `config/partner_list.xlsx` - 取引先情報（機密性: 高）

**推奨**:
- GitHubにはconfigフォルダを含めない（`.gitignore`で除外）
- 設定ファイルはアップロード方式で毎回提供

### 一時ファイル
- アップロードされたファイルは一時ディレクトリに保存
- 処理完了後は自動削除

---

## 📈 拡張性

### 新しい入力形式の追加

1. `reader/` に新しいReaderクラスを作成
2. `read_and_validate()` メソッドを実装
3. `streamlit_app.py` でReaderを選択する条件分岐を追加

### 新しい出力形式の追加

1. `exporter/` に新しいExporterクラスを作成
2. `export()` メソッドを実装
3. `streamlit_app.py` でExporterを選択する条件分岐を追加

### 新しい処理の追加

1. `processor/` に新しいProcessorクラスを作成
2. `process()` メソッドを実装（入力: data_list, 出力: data_list）
3. `streamlit_app.py` の処理フローに追加

---

## 🧪 テスト

### 単体テスト（推奨）
各モジュールごとにテストを作成：

```
tests/
├── test_reader.py
├── test_processor.py
└── test_exporter.py
```

### 統合テスト
実際のサンプルファイルを使用して、全体フローをテスト

---

## 📊 パフォーマンス

### ボトルネック
- Excelファイルの読み書き（openpyxl）
- 大量データの処理（pandas）

### 最適化案
- pandas の `read_csv` で chunksizeを指定
- 並列処理（複数ファイルを同時処理）

---

## 🔄 今後の改善案

1. **エラー処理の強化**
   - より詳細なエラーメッセージ
   - 修正提案の表示

2. **プレビュー機能**
   - 実行前にデータをプレビュー
   - 変換結果を確認してからダウンロード

3. **ログ機能**
   - 処理履歴の保存
   - エラーログの蓄積

4. **設定のカスタマイズ**
   - UIから設定を編集
   - 設定プロファイルの保存

5. **バリデーションの強化**
   - freee APIとの連携で取引先を検証
   - 勘定科目の妥当性チェック

---

## 📞 問い合わせ

システム構成に関する質問は、GitHubのIssuesまでお願いします。