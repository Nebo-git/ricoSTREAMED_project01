import pandas as pd
from datetime import datetime


class TestExcelReader:
    """テスト用Excelファイルを読み込み、データを検証するクラス"""
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.data_list = []
        self.errors = []
    
    def read_and_validate(self):
        """全シートを読み込み、データを検証する
        
        Returns:
            tuple: (data_list, errors)
                - data_list: list[dict] - 各行の検証済みデータ
                - errors: list[str] - エラーメッセージのリスト
        """
        try:
            excel_file = pd.ExcelFile(self.file_path, engine='openpyxl')
            
            for sheet_name in excel_file.sheet_names:
                self._process_sheet(excel_file, sheet_name)
            
            return self.data_list, self.errors
        
        except Exception as e:
            raise Exception(f"ファイル読み込みエラー: {str(e)}")
    
    def _process_sheet(self, excel_file, sheet_name):
        """各シートを処理する"""
        try:
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)
            
            # B2 (行1, 列1) の日付を取得
            date_value = self._get_cell_value(df, 1, 1)
            validated_date = self._validate_date(date_value, sheet_name)
            
            # C3 (行2, 列2) の金額を取得
            amount_value = self._get_cell_value(df, 2, 2)
            validated_amount = self._validate_amount(amount_value, sheet_name)
            
            # データを追加（freee項目名に準拠）
            self.data_list.append({
                '日付': validated_date,
                '金額': validated_amount,
                'シート名': sheet_name,
                '_errors': []
            })
        
        except Exception as e:
            self.errors.append(f"{sheet_name}: シート読み込みエラー ({str(e)})")
            self.data_list.append({
                '日付': '形式エラー',
                '金額': '形式エラー',
                'シート名': sheet_name,
                '_errors': [f"シート読み込みエラー: {str(e)}"]
            })
    
    def _get_cell_value(self, df, row, col):
        """セルの値を安全に取得する"""
        try:
            if row < len(df) and col < len(df.columns):
                value = df.iloc[row, col]
                return value if pd.notna(value) else None
            return None
        except:
            return None
    
    def _validate_date(self, value, sheet_name):
        """日付を検証し、yyyy-mm-dd形式に変換する"""
        if value is None:
            error_msg = f"{sheet_name}: 日付が空白です（セルB2）"
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
                error_msg = f"{sheet_name}: 日付が形式エラー（セルB2の値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        # 文字列の場合
        elif isinstance(value, str):
            try:
                if len(value) == 8 and value.isdigit():
                    year = value[:4]
                    month = value[4:6]
                    day = value[6:8]
                    datetime(int(year), int(month), int(day))
                    return f"{year}-{month}-{day}"
                else:
                    raise ValueError("yyyymmdd形式ではありません")
            except:
                error_msg = f"{sheet_name}: 日付が形式エラー（セルB2の値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        # datetime型の場合
        elif isinstance(value, datetime):
            return value.strftime('%Y-%m-%d')
        
        else:
            error_msg = f"{sheet_name}: 日付が形式エラー（セルB2の値: {value}）"
            self.errors.append(error_msg)
            return '形式エラー'
    
    def _validate_amount(self, value, sheet_name):
        """金額を検証する"""
        if value is None:
            error_msg = f"{sheet_name}: 金額が空白です（セルC3）"
            self.errors.append(error_msg)
            return '形式エラー'
        
        if isinstance(value, (int, float)):
            return value
        
        elif isinstance(value, str):
            try:
                return float(value)
            except:
                error_msg = f"{sheet_name}: 金額が形式エラー（セルC3の値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        else:
            error_msg = f"{sheet_name}: 金額が形式エラー（セルC3の値: {value}）"
            self.errors.append(error_msg)
            return '形式エラー'
        