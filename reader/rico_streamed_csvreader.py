"""
reader/rico_streamed_csvreader.py - リコホテルズSTREAMED形式CSVを読み込み、freee形式に変換
"""

import pandas as pd
from datetime import datetime


class RicoStreamedCSVReader:
    """リコホテルズSTREAMED形式CSVを読み込み、データを検証するクラス"""
    
    # STREAMED形式のカラムマッピング
    REQUIRED_COLUMNS = [
        '日付', '伝票番号', '借方勘定科目', '借方補助科目', '借方部門',
        '借方金額', '借方税区分', '貸方勘定科目', '貸方補助科目',
        '貸方部門', '貸方金額', '貸方税区分', '摘要'
    ]
    
    def __init__(self, file_path):
        self.file_path = file_path
        self.data_list = []
        self.errors = []
    
    def read_and_validate(self):
        """STREAMED形式のCSVを読み込み、データを検証する
        
        Returns:
            tuple: (data_list, errors)
                - data_list: list[dict] - 各行の検証済みデータ
                - errors: list[str] - エラーメッセージのリスト
        """
        try:
            # UTF-8でCSVを読み込み（エンコーディング自動判定も試行）
            try:
                df = pd.read_csv(self.file_path, encoding='utf-8', header=0)
            except UnicodeDecodeError:
                # UTF-8で失敗した場合はCP932（Windows Shift-JIS）を試行
                try:
                    df = pd.read_csv(self.file_path, encoding='cp932', header=0)
                except UnicodeDecodeError:
                    # それでも失敗したらShift-JISを試行
                    df = pd.read_csv(self.file_path, encoding='shift_jis', header=0)
            
            # 列名を取得
            columns = df.columns.tolist()
            
            # 必須列のチェック
            missing_columns = [col for col in self.REQUIRED_COLUMNS if col not in columns]
            if missing_columns:
                raise Exception(f"必須列が見つかりません: {', '.join(missing_columns)}")
            
            # 各行を処理
            for idx, row in df.iterrows():
                self._process_row(row, idx + 2, columns)  # idx+2 (ヘッダーが1行目なので)
            
            return self.data_list, self.errors
        
        except Exception as e:
            raise Exception(f"ファイル読み込みエラー: {str(e)}")
    
    def _process_row(self, row, row_number, columns):
        """各行を処理する"""
        try:
            # データ辞書を作成
            data = {}
            
            # 全列をループして辞書に格納
            for col in columns:
                value = row[col]
                data[col] = '' if pd.isna(value) else str(value).strip()
            
            # 「日付」列の検証（共通キー）
            date_value = row['日付']
            validated_date = self._validate_date(date_value, row_number)
            data['日付'] = validated_date
            
            # 「借方金額」と「貸方金額」の検証（共通キー）
            borrow_amount = row['借方金額']
            lend_amount = row['貸方金額']
            validated_borrow = self._validate_amount(borrow_amount, row_number, '借方金額')
            validated_lend = self._validate_amount(lend_amount, row_number, '貸方金額')
            data['借方金額'] = validated_borrow
            data['貸方金額'] = validated_lend
            
            # 共通キー「金額」を設定（借方金額を使用）
            data['金額'] = validated_borrow
            
            # 補助科目を取引先に変換
            data['借方取引先'] = data.pop('借方補助科目', '')
            data['貸方取引先'] = data.pop('貸方補助科目', '')
            
            # エラーフラグ用キーを追加
            data['_errors'] = []
            data['候補'] = ''  # マッチング結果用
            
            self.data_list.append(data)
        
        except Exception as e:
            self.errors.append(f"{row_number}行目: 行の処理エラー ({str(e)})")
    
    def _validate_date(self, value, row_number):
        """日付を検証し、yyyy-mm-dd形式に変換する"""
        if pd.isna(value) or value == '':
            error_msg = f"{row_number}行目: 日付が空白です"
            self.errors.append(error_msg)
            return '形式エラー'
        
        value_str = str(value).strip()
        
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
                if len(value_str) == 8 and value_str.isdigit():
                    year = value_str[:4]
                    month = value_str[4:6]
                    day = value_str[6:8]
                    datetime(int(year), int(month), int(day))
                    return f"{year}-{month}-{day}"
                # yyyy/mm/dd形式
                elif '/' in value_str:
                    parts = value_str.split('/')
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
    
    def _validate_amount(self, value, row_number, field_name):
        """金額を検証する"""
        if pd.isna(value) or value == '':
            error_msg = f"{row_number}行目: {field_name}が空白です"
            self.errors.append(error_msg)
            return '形式エラー'
        
        if isinstance(value, (int, float)):
            return value
        
        elif isinstance(value, str):
            try:
                return float(value)
            except:
                error_msg = f"{row_number}行目: {field_name}が形式エラー（値: {value}）"
                self.errors.append(error_msg)
                return '形式エラー'
        
        else:
            error_msg = f"{row_number}行目: {field_name}が形式エラー（値: {value}）"
            self.errors.append(error_msg)
            return '形式エラー'
        