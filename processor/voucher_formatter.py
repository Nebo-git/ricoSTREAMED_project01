"""
processor/voucher_formatter.py - 伝票番号を整形するクラス
"""

from datetime import datetime
from processor.config import DEPT_CODES, IMPORT_FORMAT_CODES


class VoucherFormatter:
    """伝票番号を再生成・整形するクラス
    
    フォーマット: [インポート形式][部門][月][元の伝票番号]
    例: 2110001
      - 2: STREAMED
      - 1: 本部
      - 10: 10月（実行日基準）
      - 001: 元の伝票番号（3桁ゼロパディング）
    
    7桁形式: DDMMVVV
    """
    
    def __init__(self, import_format: str):
        """
        Args:
            import_format: str - インポート形式（"CR", "STREAMED", "総振"）
        """
        self.import_format = import_format
        self.import_code = IMPORT_FORMAT_CODES.get(import_format)
        
        if not self.import_code:
            raise ValueError(f"不正なインポート形式: {import_format}")
        
        # 実行日の月を取得
        self.current_month = datetime.now().month
    
    def format(self, data_list: list[dict]) -> list[dict]:
        """伝票番号を整形する
        
        Args:
            data_list: list[dict] - 検証済みデータリスト
        
        Returns:
            list[dict] - 伝票番号整形済みデータリスト
        """
        for data in data_list:
            try:
                # 生成用の情報を取得
                voucher_num = data.get('伝票番号', '')
                borrow_dept = data.get('借方部門', '本部')
                
                # 伝票番号を生成（実行日の月を使用）
                formatted_voucher = self._generate_voucher(
                    voucher_num, borrow_dept, self.current_month
                )
                
                data['伝票番号'] = formatted_voucher
            
            except Exception as e:
                # エラーが発生した場合
                self._add_error(data, str(e))
                data['伝票番号'] = f"ERR_{data.get('伝票番号', 'UNKNOWN')}"
        
        return data_list
    
    def _generate_voucher(self, voucher_num: str, dept_name: str, month: int) -> str:
        """伝票番号を生成する
        
        Args:
            voucher_num: str - 元の伝票番号
            dept_name: str - 部門名
            month: int - 月（実行日基準）
        
        Returns:
            str - 生成済み伝票番号（7桁）
        """
        # 部門コードを取得
        dept_code = DEPT_CODES.get(dept_name)
        if not dept_code:
            raise ValueError(f"部門コードが見つかりません: {dept_name}")
        
        # 元の伝票番号をパース
        try:
            voucher_base = int(voucher_num)
            if voucher_base < 0 or voucher_base > 999:
                raise ValueError(f"伝票番号が範囲外（0-999）: {voucher_base}")
        except ValueError as e:
            raise ValueError(f"伝票番号が数値ではありません: {voucher_num}") from e
        
        # フォーマット: [インポート形式][部門][月（2桁）][伝票番号（3桁）]
        formatted = f"{self.import_code}{dept_code}{month:02d}{voucher_base:03d}"
        
        return formatted
    
    def _add_error(self, data: dict, error_msg: str):
        """エラーメッセージを_errorsに追記
        
        Args:
            data: dict - データ行
            error_msg: str - エラーメッセージ
        """
        if '_errors' not in data:
            data['_errors'] = []
        
        if isinstance(data['_errors'], str):
            data['_errors'] = [data['_errors']] if data['_errors'] else []
        
        data['_errors'].append(f"伝票番号生成エラー: {error_msg}")
"""
processor/voucher_formatter.py - 伝票番号を整形するクラス
"""

from datetime import datetime
from processor.config import DEPT_CODES, IMPORT_FORMAT_CODES


class VoucherFormatter:
    """伝票番号を再生成・整形するクラス
    
    フォーマット: [インポート形式][部門][月][元の伝票番号]
    例: 2112001
      - 2: STREAMED
      - 1: 本部
      - 12: 12月
      - 001: 元の伝票番号（3桁ゼロパディング）
    
    7桁形式: DDMMVVV
    """
    
    def __init__(self, import_format: str):
        """
        Args:
            import_format: str - インポート形式（"CR", "STREAMED", "総振"）
        """
        self.import_format = import_format
        self.import_code = IMPORT_FORMAT_CODES.get(import_format)
        
        if not self.import_code:
            raise ValueError(f"不正なインポート形式: {import_format}")
    
    def format(self, data_list: list[dict]) -> list[dict]:
        """伝票番号を整形する
        
        Args:
            data_list: list[dict] - 検証済みデータリスト
        
        Returns:
            list[dict] - 伝票番号整形済みデータリスト
        """
        for data in data_list:
            try:
                # 生成用の情報を取得
                date_str = data.get('日付', '')
                voucher_num = data.get('伝票番号', '')
                borrow_dept = data.get('借方部門', '本部')
                
                # 伝票番号を生成
                formatted_voucher = self._generate_voucher(
                    date_str, voucher_num, borrow_dept
                )
                
                data['伝票番号'] = formatted_voucher
            
            except Exception as e:
                # エラーが発生した場合
                self._add_error(data, str(e))
                data['伝票番号'] = f"ERR_{data.get('伝票番号', 'UNKNOWN')}"
        
        return data_list
    
    def _generate_voucher(self, date_str: str, voucher_num: str, dept_name: str) -> str:
        """伝票番号を生成する
        
        Args:
            date_str: str - 日付（yyyy-mm-dd形式）
            voucher_num: str - 元の伝票番号
            dept_name: str - 部門名
        
        Returns:
            str - 生成済み伝票番号（7桁）
        """
        # 部門コードを取得
        dept_code = DEPT_CODES.get(dept_name)
        if not dept_code:
            raise ValueError(f"部門コードが見つかりません: {dept_name}")
        
        # 日付から月を抽出
        try:
            month = int(date_str.split('-')[1])
            if not (1 <= month <= 12):
                raise ValueError(f"月が範囲外: {month}")
        except (IndexError, ValueError) as e:
            raise ValueError(f"日付形式エラー: {date_str}") from e
        
        # 元の伝票番号をパース
        try:
            voucher_base = int(voucher_num)
            if voucher_base < 0 or voucher_base > 999:
                raise ValueError(f"伝票番号が範囲外（0-999）: {voucher_base}")
        except ValueError as e:
            raise ValueError(f"伝票番号が数値ではありません: {voucher_num}") from e
        
        # フォーマット: [インポート形式][部門][月（2桁）][伝票番号（3桁）]
        formatted = f"{self.import_code}{dept_code}{month:02d}{voucher_base:03d}"
        
        return formatted
    
    def _add_error(self, data: dict, error_msg: str):
        """エラーメッセージを_errorsに追記
        
        Args:
            data: dict - データ行
            error_msg: str - エラーメッセージ
        """
        if '_errors' not in data:
            data['_errors'] = []
        
        if isinstance(data['_errors'], str):
            data['_errors'] = [data['_errors']] if data['_errors'] else []
        
        data['_errors'].append(f"伝票番号生成エラー: {error_msg}")