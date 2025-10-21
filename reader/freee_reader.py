import pandas as pd
from datetime import datetime


class FreeeExcelReader:
    """freee形式Excelファイルを読み込み、データを検証するクラス"""
    
    # freee用の全列名
    FREEE_COLUMNS = [
        '[表題行]', '日付', '伝票番号', '決算整理仕訳',
        '借方勘定科目', '借方科目コード', '借方補助科目', '借方取引先', '借方取引先コード',
        '借方部門', '借方品目', '借方メモタグ', '借方セグメント1', '借方セグメント2', '借方セグメント3',
        '借方金額', '借方税区分', '借方税額',
        '貸方勘定科目', '貸方科目コード', '貸方補助科目', '貸方取引先', '貸方取引先コード',
        '貸方部門', '貸方品目', '貸方メモタグ', '貸方セグメント1', '貸方セグメント2', '貸方セグメント3',
        '貸方金額', '貸方税区分', '貸方税額', '摘要'
    ]
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.data_list = []
        self.errors = []
    
    def read_and_validate(self):
        """freee形式のExcelを読み込み、データを検証する
        
        Returns:
            tuple: (data_list, errors)
                - data_list: list[dict] - 各行の検証済みデータ
                - errors: list[str] - エラーメッセージのリスト
        """
        try:
            # 1行目をヘッダーとして読み込み
            df = pd.read_excel(self.file_path, engine='openpyxl', header=0)
            
            # 列名を取得
            columns = df.columns.tolist()
            
            # 必須列のチェック
            if '日付' not in columns:
                raise Exception("「日付」列が見つかりません")
            if '金額' not in columns:
                raise Exception("「金額」列が見つかりません")
            
            # 各行を処理
            for idx, row in df.iterrows():
                self._process_row(row, idx + 2, columns)
            
            return self.data_list, self.errors
        
        except Exception as e:
            raise Exception(f"ファイル読み込みエラー: {str(e)}")
    
    def _process_row(self, row, row_number, columns):
        """各行を処理する"""
        try:
            # データ辞書を作成（全列を含む）
            data = {}
            
            # 全列をループして辞書に格納
            for col in columns:
                value = row[col]
                data[col] = '' if pd.isna(value) else value
            
            # 「日付」列の検証（共通キー）
            date_value = row['日付']
            validated_date = self._validate_date(date_value, row_number)
            data['日付'] = validated_date
            
            # 「金額」列の検証（共通キー）
            amount_value = row['金額']
            validated_amount = self._validate_amount(amount_value, row_number)
            data['金額'] = validated_amount
            
            # 「借方金額」「貸方金額」があれば、「金額」の値で上書き
            if '借方金額' in columns:
                data['借方金額'] = validated_amount
            if '貸方金額' in columns:
                data['貸方金額'] = validated_amount
            
            # エラーフラグ用キーを追加
            data['_errors'] = []
            
            self.data_list.append(data)
        
        except Exception as e:
            self.errors.append(f"{row_number}行目: 行の処理エラー ({str(e)})")
    
    def _validate_date(self, value, row_number):
        """日付を検証し、yyyy-mm-dd形式に変換する"""
        if pd.isna(value) or value == '':
            error_msg = f"{row_number}行目: 日付が空白です"
            self.errors.append(error_msg)
            return '形式エラー'
        
        # 数値の場合（yyyymmdd形式）
        if isinstance(value, (int, float)):
            try:
                date_str = str(int(value))
                if len(date_str) == 8:
                    year = date_str[:4]
                    month = date_str[4:6]
                    day = date_str[6:8]
                    datetime(int(year), int(month), int(day))
                    return f"{year}-{month}-{day}"
                else:
                    raise ValueError("8桁ではありません")
            except:
                error_msg = f"{row_number}行目: 日付が形式エラー（値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        # 文字列の場合
        elif isinstance(value, str):
            try:
                # yyyymmdd形式
                if len(value) == 8 and value.isdigit():
                    year = value[:4]
                    month = value[4:6]
                    day = value[6:8]
                    datetime(int(year), int(month), int(day))
                    return f"{year}-{month}-{day}"
                # yyyy/mm/dd形式
                elif '/' in value:
                    parts = value.split('/')
                    if len(parts) == 3:
                        year, month, day = parts
                        datetime(int(year), int(month), int(day))
                        return f"{year.zfill(4)}-{month.zfill(2)}-{day.zfill(2)}"
                    else:
                        raise ValueError("yyyy/mm/dd形式ではありません")
                else:
                    raise ValueError("対応していない日付形式です")
            except:
                error_msg = f"{row_number}行目: 日付が形式エラー（値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        # datetime型の場合
        elif isinstance(value, datetime):
            return value.strftime('%Y-%m-%d')
        
        else:
            error_msg = f"{row_number}行目: 日付が形式エラー（値: {value}）"
            self.errors.append(error_msg)
            return '形式エラー'
    
    def _validate_amount(self, value, row_number):
        """金額を検証する"""
        if pd.isna(value) or value == '':
            error_msg = f"{row_number}行目: 金額が空白です"
            self.errors.append(error_msg)
            return '形式エラー'
        
        if isinstance(value, (int, float)):
            return value
        
        elif isinstance(value, str):
            try:
                return float(value)
            except:
                error_msg = f"{row_number}行目: 金額が形式エラー（値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        else:
            error_msg = f"{row_number}行目: 金額が形式エラー（値: {value}）"
            self.errors.append(error_msg)
            return '形式エラー'
        