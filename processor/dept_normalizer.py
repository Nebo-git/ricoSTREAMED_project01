"""
processor/dept_normalizer.py - 部門名を正規化するクラス
"""

import pandas as pd
from pathlib import Path
from processor.config import DEPT_CODES


class DeptNormalizer:
    """部門名を正規化するクラス"""
    
    def __init__(self, dept_mapping_path):
        """
        Args:
            dept_mapping_path: str - dept_mapping.xlsx のパス
        """
        self.dept_mapping_path = Path(dept_mapping_path)
        self.dept_map = {}
        self._load_dept_mapping()
    
    def _load_dept_mapping(self):
        """部門マッピング辞書を読み込む
        
        dept_mapping.xlsx:
        A: 元の名称, B: 正式名称
        """
        try:
            df = pd.read_excel(self.dept_mapping_path, header=0)
            
            # 列名を取得（最初の2列を使用）
            col1, col2 = df.columns[0], df.columns[1]
            
            # マッピング辞書を作成
            for _, row in df.iterrows():
                original = str(row[col1]).strip()
                formal = str(row[col2]).strip()
                self.dept_map[original] = formal
        
        except Exception as e:
            raise Exception(f"部門マッピングファイル読み込みエラー: {str(e)}")
    
    def normalize(self, data_list: list[dict]) -> list[dict]:
        """部門名を正規化する
        
        Args:
            data_list: list[dict] - 検証済みデータリスト
        
        Returns:
            list[dict] - 部門名正規化済みデータリスト
        """
        # ファイル全体を1部門と仮定して、最初に見つかった部門を取得
        default_dept = self._find_first_dept(data_list)
        
        # 各行を処理
        for data in data_list:
            # 借方部門の正規化
            borrow_dept = data.get('借方部門', '').strip()
            if not borrow_dept:
                # 空欄の場合はデフォルト部門を使用
                data['借方部門'] = default_dept
            else:
                normalized_borrow = self.dept_map.get(borrow_dept, f"未登録_{borrow_dept}")
                data['借方部門'] = normalized_borrow
                
                # マッピング内に存在しない場合はエラーフラグ
                if borrow_dept not in self.dept_map:
                    self._add_error(data, f"借方部門が未登録: {borrow_dept}")
            
            # 貸方部門の正規化
            lend_dept = data.get('貸方部門', '').strip()
            if not lend_dept:
                # 空欄の場合はデフォルト部門を使用
                data['貸方部門'] = default_dept
            else:
                normalized_lend = self.dept_map.get(lend_dept, f"未登録_{lend_dept}")
                data['貸方部門'] = normalized_lend
                
                # マッピング内に存在しない場合はエラーフラグ
                if lend_dept not in self.dept_map:
                    self._add_error(data, f"貸方部門が未登録: {lend_dept}")
        
        return data_list
    
    def _find_first_dept(self, data_list: list[dict]) -> str:
        """最初に見つかった部門を取得（空欄でない）
        
        Args:
            data_list: list[dict]
        
        Returns:
            str - デフォルト部門（見つからない場合は"本部"）
        """
        for data in data_list:
            borrow_dept = data.get('借方部門', '').strip()
            if borrow_dept:
                return self.dept_map.get(borrow_dept, borrow_dept)
            
            lend_dept = data.get('貸方部門', '').strip()
            if lend_dept:
                return self.dept_map.get(lend_dept, lend_dept)
        
        # デフォルトは「本部」
        return "本部"
    
    def _add_error(self, data: dict, error_msg: str):
        """エラーメッセージを_errorsに追記
        
        Args:
            data: dict - データ行
            error_msg: str - エラーメッセージ
        """
        if '_errors' not in data:
            data['_errors'] = []
        
        if isinstance(data['_errors'], str):
            # 文字列の場合はリストに変換
            data['_errors'] = [data['_errors']] if data['_errors'] else []
        
        data['_errors'].append(error_msg)