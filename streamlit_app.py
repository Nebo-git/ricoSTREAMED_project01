"""
streamlit_app.py - Streamlit版メインアプリ
"""

import streamlit as st
from pathlib import Path
import tempfile
import os

from reader.test_reader01 import TestExcelReader
from reader.freee_reader import FreeeExcelReader
from reader.rico_streamed_csvreader import RicoStreamedCSVReader
from processor.dept_normalizer import DeptNormalizer
from processor.partner_resolver import PartnerResolver
from processor.voucher_formatter import VoucherFormatter
from exporter.freee_exporter import TestExcelExporter, FreeeExcelExporter


# プロジェクトのルートディレクトリ
PROJECT_ROOT = Path(__file__).parent

# 一時ディレクトリを作成
TEMP_DIR = Path(tempfile.gettempdir()) / "streamlit_converter"
TEMP_DIR.mkdir(exist_ok=True)

# ページ設定
st.set_page_config(
    page_title="Excel to CSV Converter",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# カスタムCSS
st.markdown("""
<style>
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        color: #1976D2;
        text-align: center;
        padding: 1rem 0;
    }
    .step-header {
        font-size: 1.5rem;
        font-weight: bold;
        color: #424242;
        margin-top: 2rem;
        padding: 0.5rem;
        border-left: 4px solid #2196F3;
        background-color: #E3F2FD;
    }
    .success-box {
        padding: 1rem;
        background-color: #C8E6C9;
        border-radius: 0.5rem;
        border-left: 4px solid #4CAF50;
    }
    .error-box {
        padding: 1rem;
        background-color: #FFCDD2;
        border-radius: 0.5rem;
        border-left: 4px solid #F44336;
    }
    .info-box {
        padding: 1rem;
        background-color: #FFF9C4;
        border-radius: 0.5rem;
        border-left: 4px solid #FFC107;
    }
</style>
""", unsafe_allow_html=True)


def main():
    # ヘッダー
    st.markdown('<div class="main-header">📊 Excel to CSV Converter</div>', unsafe_allow_html=True)
    
    # サイドバー: 設定
    with st.sidebar:
        st.header("⚙️ 設定")
        
        # 入力形式選択
        input_type = st.radio(
            "📥 インプット形式",
            options=["test", "freee", "streamed"],
            format_func=lambda x: {
                "test": "テスト用Excel (B2:日付, C3:金額)",
                "freee": "freee形式Excel (freee項目名からデータ照合)",
                "streamed": "リコホテルズ_STREAMED_csv"
            }[x],
            help="処理するファイルの形式を選択してください"
        )
        
        st.divider()
        
        # 出力形式選択
        output_type = st.radio(
            "📤 アウトプット形式",
            options=["test", "freee"],
            format_func=lambda x: "テスト用Excel" if x == "test" else "freee用Excel",
            help="出力するファイルの形式を選択してください"
        )
        
        st.divider()
        
        # 設定ファイルアップロード（STREAMED用のみ表示）
        dept_mapping_file = None
        partner_list_file = None
        
        if input_type == "streamed":
            st.subheader("📋 設定ファイル（任意）")
            st.caption("アップロードしない場合は、configフォルダの固定ファイルを使用します")
            
            dept_mapping_file = st.file_uploader(
                "部署マッピングExcel",
                type=["xlsx"],
                help="部署名の正規化に使用",
                key="dept_mapping"
            )
            
            partner_list_file = st.file_uploader(
                "取引先一覧Excel",
                type=["xlsx"],
                help="取引先名の照合に使用",
                key="partner_list"
            )
            
            # アップロード状態の表示
            if dept_mapping_file or partner_list_file:
                st.divider()
                st.markdown("#### ✅ アップロード済み")
                if dept_mapping_file:
                    st.success(f"📋 {dept_mapping_file.name}")
                if partner_list_file:
                    st.success(f"📋 {partner_list_file.name}")
            
            st.divider()
        
        # 使い方
        with st.expander("📖 使い方"):
            st.markdown("""
            **基本的な流れ：**
            1. インプット形式を選択
            2. （STREAMEDの場合）設定ファイルをアップロード（任意）
            3. ファイルをアップロード
            4. （STREAMEDの場合）freee取引先CSVをアップロード
            5. アウトプット形式を選択
            6. 「実行」ボタンをクリック
            7. 生成されたファイルをダウンロード
            """)
    
    # メインエリア
    st.markdown('<div class="step-header">📂 Step 1: ファイルをアップロード</div>', unsafe_allow_html=True)
    
    # ファイルアップロード
    if input_type == "streamed":
        uploaded_files = st.file_uploader(
            "STREAMED CSVファイルを選択（複数可）",
            type=["csv"],
            accept_multiple_files=True,
            help="複数ファイルを一度にアップロードできます"
        )
    else:
        uploaded_files = st.file_uploader(
            f"{'テスト用' if input_type == 'test' else 'freee形式'}Excelファイルを選択（複数可）",
            type=["xlsx", "xls"],
            accept_multiple_files=True,
            help="複数ファイルを一度にアップロードできます"
        )
    
    # STREAMED用のfreee取引先CSV
    freee_partner_file = None
    if input_type == "streamed" and uploaded_files:
        st.markdown('<div class="step-header">📂 Step 2: freee取引先CSVをアップロード</div>', unsafe_allow_html=True)
        st.markdown('<div class="info-box">💡 configフォルダの固定リストを最優先で参照します。このCSVは補助的に参照されます。</div>', unsafe_allow_html=True)
        
        freee_partner_file = st.file_uploader(
            "freee取引先CSVを選択",
            type=["csv"],
            help="取引先名の照合に使用します"
        )
    
    # 実行ボタン
    if uploaded_files:
        step_num = 3 if input_type == "streamed" else 2
        st.markdown(f'<div class="step-header">▶️ Step {step_num}: 実行</div>', unsafe_allow_html=True)
        
        col1, col2, col3 = st.columns([1, 2, 1])
        with col2:
            if input_type == "streamed" and not freee_partner_file:
                st.warning("⚠️ freee取引先CSVをアップロードしてください")
                execute_button = st.button("実行", disabled=True, use_container_width=True)
            else:
                button_color = "#2196F3" if output_type == "test" else "#4CAF50"
                execute_button = st.button(
                    f"🚀 {'テスト用' if output_type == 'test' else 'freee用'}で実行",
                    type="primary",
                    use_container_width=True
                )
        
        if execute_button:
            # STREAMEDの場合は設定ファイルも渡す
            if input_type == "streamed":
                process_files(
                    uploaded_files, 
                    input_type, 
                    output_type, 
                    freee_partner_file,
                    dept_mapping_file,
                    partner_list_file
                )
            else:
                process_files(uploaded_files, input_type, output_type)


def process_files(uploaded_files, input_type, output_type, freee_partner_file=None, dept_mapping_file=None, partner_list_file=None):
    """ファイルを処理する"""
    
    progress_bar = st.progress(0)
    status_text = st.empty()
    
    all_errors = []
    output_files = []
    
    total_files = len(uploaded_files)
    
    for idx, uploaded_file in enumerate(uploaded_files):
        try:
            status_text.text(f"処理中... ({idx + 1}/{total_files}) {uploaded_file.name}")
            progress_bar.progress((idx + 1) / total_files)
            
            # 一時ファイルとして保存
            temp_path = TEMP_DIR / uploaded_file.name
            temp_path.write_bytes(uploaded_file.getvalue())
            
            # Reader選択
            if input_type == "test":
                reader = TestExcelReader(str(temp_path))
            elif input_type == "freee":
                reader = FreeeExcelReader(str(temp_path))
            elif input_type == "streamed":
                reader = RicoStreamedCSVReader(str(temp_path))
            
            # データ読み込み
            data_list, errors = reader.read_and_validate()
            
            # STREAMED処理
            if input_type == "streamed":
                # freee取引先CSVを一時保存
                freee_csv_path = TEMP_DIR / "freee_partners.csv"
                freee_csv_path.write_bytes(freee_partner_file.getvalue())
                
                # 設定ファイルのパスを決定（アップロードされていればそちらを使用）
                if dept_mapping_file:
                    dept_mapping_path = TEMP_DIR / "temp_dept_mapping.xlsx"
                    dept_mapping_path.write_bytes(dept_mapping_file.getvalue())
                    st.sidebar.info("✅ アップロードされた部署マッピングを使用")
                else:
                    dept_mapping_path = PROJECT_ROOT / "config" / "dept_mapping.xlsx"
                    st.sidebar.info("📁 configフォルダの部署マッピングを使用")
                
                if partner_list_file:
                    partner_list_path = TEMP_DIR / "temp_partner_list.xlsx"
                    partner_list_path.write_bytes(partner_list_file.getvalue())
                    st.sidebar.info("✅ アップロードされた取引先一覧を使用")
                else:
                    partner_list_path = PROJECT_ROOT / "config" / "partner_list.xlsx"
                    st.sidebar.info("📁 configフォルダの取引先一覧を使用")
                
                dept_normalizer = DeptNormalizer(str(dept_mapping_path))
                data_list = dept_normalizer.normalize(data_list)
                
                partner_resolver = PartnerResolver(
                    str(partner_list_path),
                    str(freee_csv_path)
                )
                data_list = partner_resolver.resolve(data_list)
                
                voucher_formatter = VoucherFormatter("STREAMED")
                data_list = voucher_formatter.format(data_list)
            
            # Exporter選択
            if output_type == "test":
                exporter = TestExcelExporter(output_dir=str(TEMP_DIR))
            elif output_type == "freee":
                exporter = FreeeExcelExporter(output_dir=str(TEMP_DIR))
            
            # Excel出力
            output_path = exporter.export(data_list, uploaded_file.name)
            output_files.append((uploaded_file.name, output_path))
            
            if errors:
                all_errors.extend([f"{uploaded_file.name}: {e}" for e in errors])
        
        except Exception as e:
            all_errors.append(f"{uploaded_file.name}: {str(e)}")
    
    progress_bar.empty()
    status_text.empty()
    
    # 結果表示
    show_results(output_files, all_errors, output_type)


def show_results(output_files, all_errors, output_type):
    """処理結果を表示する"""
    
    st.markdown('<div class="step-header">✅ 処理完了</div>', unsafe_allow_html=True)
    
    if all_errors:
        st.markdown('<div class="error-box">', unsafe_allow_html=True)
        st.warning(f"⚠️ {len(all_errors)}件のエラーがありました")
        with st.expander("エラー詳細を表示"):
            for error in all_errors[:20]:
                st.text(f"- {error}")
            if len(all_errors) > 20:
                st.text(f"... 他{len(all_errors)-20}件")
        st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="success-box">', unsafe_allow_html=True)
        st.success(f"✅ {len(output_files)}個のファイルが正常に処理されました！")
        st.markdown('</div>', unsafe_allow_html=True)
    
    # ダウンロードボタン
    st.markdown("### 📥 ダウンロード")
    
    # 一括ダウンロードボタン
    if len(output_files) > 1:
        st.markdown("#### 🎁 一括ダウンロード")
        
        import zipfile
        import io
        
        # ZIPファイルを作成
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for original_name, output_path in output_files:
                output_path = Path(output_path)
                with open(output_path, "rb") as f:
                    zip_file.writestr(output_path.name, f.read())
        
        zip_buffer.seek(0)
        
        st.download_button(
            label="📦 すべてのファイルをZIPでダウンロード",
            data=zip_buffer,
            file_name=f"output_files_{output_type}.zip",
            mime="application/zip",
            key="download_all_zip",
            use_container_width=True
        )
        
        st.divider()
    
    # 個別ダウンロード
    st.markdown("#### 📄 個別ダウンロード")
    
    for idx, (original_name, output_path) in enumerate(output_files):
        output_path = Path(output_path)
        
        with open(output_path, "rb") as f:
            file_data = f.read()
        
        col1, col2 = st.columns([3, 1])
        with col1:
            st.text(f"📄 {output_path.name}")
        with col2:
            st.download_button(
                label="⬇️ DL",
                data=file_data,
                file_name=output_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key=f"download_{idx}_{output_path.name}",
                use_container_width=True
            )
    
    # リセットボタン
    st.divider()
    if st.button("🔄 新しいファイルを処理する", type="primary", use_container_width=True):
        st.rerun()


if __name__ == "__main__":
    main()