"""
processor/partner_resolver.py - 取引先名を正規化・解決するクラス
"""

import pandas as pd
from pathlib import Path
from difflib import SequenceMatcher


class PartnerResolver:
    """取引先名を正規化・解決するクラス
    
    優先順位:
    1. partner_list.xlsx（固定リスト・最優先）
    2. freee取引先CSV（ユーザー提供）
    3. 類似度マッチング（複数候補提示）
    """
    
    def __init__(self, partner_list_path, freee_csv_path):
        """
        Args:
            partner_list_path: str - partner_list.xlsx のパス
            freee_csv_path: str - freee取引先CSV（UTF-8/CP932/Shift-JIS）のパス
        """
        self.partner_map = {}        # 固定リスト（最優先）
        self.freee_partners = []     # freee取引先リスト
        self.freee_partner_map = {}  # freee取引先マップ
        
        self._load_partner_list(partner_list_path)
        self._load_freee_csv(freee_csv_path)
    
    def _load_partner_list(self, partner_list_path):
        """partner_list.xlsx を読み込む
        
        partner_list.xlsx:
        A: 元の名称, B: 正式名称
        """
        try:
            df = pd.read_excel(partner_list_path, header=0)
            
            # 列名を取得
            col1, col2 = df.columns[0], df.columns[1]
            
            # マッピング辞書を作成
            for _, row in df.iterrows():
                original = str(row[col1]).strip()
                formal = str(row[col2]).strip()
                self.partner_map[original] = formal
        
        except Exception as e:
            raise Exception(f"取引先リストファイル読み込みエラー: {str(e)}")
    
    def _load_freee_csv(self, freee_csv_path):
        """freee取引先CSVを読み込む
        
        freee_csv_path:
        A列: 取引先名, Q列: ステータス
        ステータスが「使用しない」のものは除外
        """
        try:
            # UTF-8でCSVを読み込み（エンコーディング自動判定）
            try:
                df = pd.read_csv(freee_csv_path, encoding='utf-8', header=0)
            except UnicodeDecodeError:
                # UTF-8で失敗した場合はCP932を試行
                try:
                    df = pd.read_csv(freee_csv_path, encoding='cp932', header=0)
                except UnicodeDecodeError:
                    # それでも失敗したらShift-JISを試行
                    df = pd.read_csv(freee_csv_path, encoding='shift_jis', header=0)
            
            # A列とQ列を取得
            partner_col = df.columns[0]  # A列
            status_col = df.columns[16] if len(df.columns) > 16 else None  # Q列（0から16=Q）
            
            # データを処理
            for _, row in df.iterrows():
                partner_name = str(row[partner_col]).strip()
                status = str(row[status_col]).strip() if status_col and status_col in df.columns else "使用"
                
                # ステータスが「使用しない」でないものを追加
                if status != "使用しない":
                    self.freee_partners.append(partner_name)
                    self.freee_partner_map[partner_name] = partner_name
        
        except Exception as e:
            raise Exception(f"freee取引先CSV読み込みエラー: {str(e)}")
    
    def resolve(self, data_list: list[dict]) -> list[dict]:
        """取引先名を解決する
        
        Args:
            data_list: list[dict] - 検証済みデータリスト
        
        Returns:
            list[dict] - 取引先解決済みデータリスト
        """
        for data in data_list:
            # 1. 空欄コピー処理（最初に実行）
            borrow_partner = data.get('借方取引先', '').strip()
            lend_partner = data.get('貸方取引先', '').strip()
            
            if not borrow_partner and lend_partner:
                # 借方が空欄 → 貸方をコピー
                data['借方取引先'] = lend_partner
                borrow_partner = lend_partner
            elif not lend_partner and borrow_partner:
                # 貸方が空欄 → 借方をコピー
                data['貸方取引先'] = borrow_partner
                lend_partner = borrow_partner
            
            # 候補列を初期化
            if '候補' not in data:
                data['候補'] = ''
            
            # 2. 借方取引先の解決（元の名称は変更しない）
            borrow_match_type = 'none'
            borrow_candidate = ''
            if borrow_partner:
                borrow_match_type, borrow_candidate = self._resolve_partner(borrow_partner)
                data['借方取引先_match_type'] = borrow_match_type  # 色付け用
            else:
                data['借方取引先_match_type'] = 'none'
            
            # 3. 貸方取引先の解決（元の名称は変更しない）
            lend_match_type = 'none'
            lend_candidate = ''
            if lend_partner:
                lend_match_type, lend_candidate = self._resolve_partner(lend_partner)
                data['貸方取引先_match_type'] = lend_match_type  # 色付け用
            else:
                data['貸方取引先_match_type'] = 'none'
            
            # 4. 候補列は1つだけ（重複排除）
            if borrow_candidate and lend_candidate:
                # 両方に候補がある場合、同じなら1つだけ、異なれば両方
                if borrow_candidate == lend_candidate:
                    data['候補'] = borrow_candidate
                else:
                    data['候補'] = f"{borrow_candidate} / {lend_candidate}"
            elif borrow_candidate:
                data['候補'] = borrow_candidate
            elif lend_candidate:
                data['候補'] = lend_candidate
        
        return data_list
    
    def _resolve_partner(self, partner_name: str) -> tuple[str, str]:
        """取引先名を解決する
        
        Args:
            partner_name: str - 解決対象の取引先名
        
        Returns:
            tuple: (match_type, candidate_name)
                - match_type: 'partner_list', 'freee_exact', 'fuzzy', 'none'
                - candidate_name: 候補名（fuzzyの場合のみ）
        """
        # 1. 固定リストで完全一致（半角全角・スペース区別）
        if partner_name in self.partner_map:
            return 'partner_list', ''
        
        # 2. freee取引先で完全一致（半角全角・スペース区別）
        if partner_name in self.freee_partner_map:
            return 'freee_exact', ''
        
        # 3. 類似度マッチング
        candidates = self._fuzzy_match(partner_name)
        
        if candidates:
            # 最高スコア候補の名前のみを返す
            return 'fuzzy', candidates[0]['name']
        else:
            return 'none', ''
    
    def _fuzzy_match(self, partner_name: str, threshold=0.6, max_candidates=1) -> list[dict]:
        """類似度マッチングで取引先候補を検索
        
        Args:
            partner_name: str - 検索対象の取引先名
            threshold: float - スコア閾値（0-1）
            max_candidates: int - 最大候補数（デフォルト1）
        
        Returns:
            list[dict] - 候補リスト（スコア降順）
                [{'name': '...',  'score': 0.95}]
        """
        candidates = []
        
        # freee取引先すべてとの類似度を計算
        for freee_partner in self.freee_partners:
            score = self._calculate_similarity(partner_name, freee_partner)
            
            if score >= threshold:
                candidates.append({
                    'name': freee_partner,
                    'score': score
                })
        
        # スコア降順でソート
        candidates.sort(key=lambda x: x['score'], reverse=True)
        
        # 最大候補数までカット
        return candidates[:max_candidates]
    
    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """2つの文字列の類似度を計算（0-1）
        
        Args:
            str1: str
            str2: str
        
        Returns:
            float - 類似度スコア（0-1）
        """
        return SequenceMatcher(None, str1, str2).ratio()