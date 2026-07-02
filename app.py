# --- EXTERNAL LIBRARIES ---
import streamlit as st # type: ignore
import pandas as pd # type: ignore
import numpy as np
import cv2
from PIL import Image
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
from docx import Document
from docx.shared import Inches

# --- PYTHON STANDARD LIBRARY ---
import os
import io
import re
import time
import zipfile
import shutil
from io import BytesIO
from datetime import datetime, date

# --- PAGE CONFIGURATION MUST BE AT THE VERY TOP ---
st.set_page_config(page_title="Workspace Tools", page_icon="🧰", layout="wide")

st.markdown(
    """
    <style>
    /* 1. Base styling for smooth transitions on all standard and download buttons */
    .stButton > button, .stDownloadButton > button {
        transition: all 0.3s cubic-bezier(0.25, 0.8, 0.25, 1);
        border-radius: 8px !important;
    }

    /* 2. Hover State: Lifts the button, adds a shadow, and triggers the animation */
    .stButton > button:hover, .stDownloadButton > button:hover {
        transform: translateY(-4px);
        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.2) !important; 
        border-color: #2e7d32 !important;
        /* Subtle green border match */
        animation: gentle-pulse 1.5s infinite ease-in-out; 
        z-index: 1;
        /* Ensures the glowing button pops above other elements */
    }

    /* 3. Click/Active State: Pushes the button down when clicked */
    .stButton > button:active, .stDownloadButton > button:active {
        transform: translateY(2px) scale(0.98) !important;
        box-shadow: 0 2px 4px rgba(0, 0, 0, 0.1) !important;
        animation: none;
        /* Stops the pulse while clicking */
    }

    /* 4. The Keyframes for the continuous pulsing effect while hovering */
    @keyframes gentle-pulse {
        0% { transform: translateY(-4px) scale(1); }
        50% { transform: translateY(-4px) scale(1.03); }
        100% { transform: translateY(-4px) scale(1); }
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- SESSION STATE ---
if 'uploader_key' not in st.session_state:
    st.session_state.uploader_key = 0
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'release_data' not in st.session_state:
    st.session_state.release_data = None
if 'generated_filename' not in st.session_state:
    st.session_state.generated_filename = ""
if 'release_filename' not in st.session_state:
    st.session_state.release_filename = ""
if 'selected_type' not in st.session_state:
    st.session_state.selected_type = None
if 'show_account_type_modal' not in st.session_state:
    st.session_state.show_account_type_modal = False
if 'process_confirm' not in st.session_state:
    st.session_state.process_confirm = False
if 'cms_id_warning' not in st.session_state:
    st.session_state.cms_id_warning = None

def reset_app():
    st.session_state.uploader_key += 1
    st.session_state.processed_data = None
    st.session_state.release_data = None
    st.session_state.selected_type = None
    st.session_state.show_account_type_modal = False
    st.session_state.process_confirm = False
    st.session_state.cms_id_warning = None
    st.rerun()

# --- SIDEBAR NAVIGATION ---
st.sidebar.title("🛠️ Workspace")
selected_tool = st.sidebar.radio(
    "Select Tool:",
    ["VRP Mapper", "Field Result", "E-SIGN FIXER"]
)

st.sidebar.divider()
st.sidebar.info("Select a tool from the menu above to get started.")


# ==========================================
# TOOL 1: VRP ACCOUNT TEMPLATE MAPPER
# ==========================================
if selected_tool == "VRP Mapper":
    st.title("🚀 VRP Mapper")

    # --- 1. SETTINGS & TEMPLATE ---
    template_filename = 'demand_letter_template.csv'
    if not os.path.exists(template_filename):
        st.error(f"❌ Error: '{template_filename}' not found.")
        st.stop()
    df_tmp = pd.read_csv(template_filename, dtype=str)

    # --- CMS ID REFERENCE LOADING ---
    @st.cache_data
    def load_cms_data():
        ref_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "cmd_id.xlsx")
        if os.path.exists(ref_path):
            try:
                df_cms = pd.read_excel(ref_path, dtype=str)
                df_cms.columns = df_cms.columns.str.strip().str.upper()
                
                # Find the CMS ID column
                cms_col = next((c for c in df_cms.columns if c in ['CMS ID', 'CMS_ID']), None)
                ch_col = next((c for c in df_cms.columns if c in ['CH CODE', 'CH_CODE']), None)
                
                if ch_col and cms_col:
                    return dict(zip(df_cms[ch_col].str.strip().str.upper(), df_cms[cms_col].str.strip()))
            except Exception:
                pass
        return {}

    cms_mapping = load_cms_data()

    # --- AMOUNTS REFERENCE LOADING ---
    @st.cache_data
    def load_amounts_data():
        ref_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "amounts.csv")
        if os.path.exists(ref_path):
            try:
                df_amt = pd.read_csv(ref_path, dtype=str)
                df_amt.columns = df_amt.columns.str.strip()
                
                # Find ChCode, Amount OB, and Principal Amount Due columns by exact name
                ch_col = next((c for c in df_amt.columns if c.upper().replace(" ", "") in ['CHCODE', 'CH_CODE']), None)
                ob_col = next((c for c in df_amt.columns if c.strip().upper() == 'AMOUNT OB'), None)
                due_col = next((c for c in df_amt.columns if c.strip().upper() == 'PRINCIPAL AMOUNT DUE'), None)
                
                if ch_col and ob_col and due_col:
                    df_amt[ch_col] = df_amt[ch_col].str.strip().str.upper()
                    
                    # Create dictionary mappings for easy lookup
                    ob_dict = dict(zip(df_amt[ch_col], df_amt[ob_col].str.strip()))
                    due_dict = dict(zip(df_amt[ch_col], df_amt[due_col].str.strip()))
                    return ob_dict, due_dict
            except Exception:
                pass
        return {}, {}

    amounts_ob_mapping, amounts_due_mapping = load_amounts_data()

    # --- FULL AREA CLUSTER REFERENCE MAPPING ---
    AREA_MAPPING = {
        'ABRA': 'NORTH LUZON', 'AURORA': 'NORTH LUZON', 'BATAAN': 'NORTH LUZON', 
        'BENGUET': 'NORTH LUZON', 'BULACAN': 'NORTH LUZON', 'CAGAYAN': 'NORTH LUZON', 
        'ILOCOS NORTE': 'NORTH LUZON', 'ILOCOS SUR': 'NORTH LUZON', 'ISABELA': 'NORTH LUZON', 
        'LA UNION': 'NORTH LUZON', 'NUEVA ECIJA': 'NORTH LUZON', 'PAMPANGA': 'NORTH LUZON', 
        'PANGASINAN': 'NORTH LUZON', 'TARLAC': 'NORTH LUZON', 'ZAMBALES': 'NORTH LUZON',
        'NCR': 'NCR', 'METRO MANILA': 'NCR',
        'BATANGAS': 'SOUTH LUZON', 'CAMARINES SUR': 'SOUTH LUZON', 'CAVITE': 'SOUTH LUZON', 
        'LAGUNA': 'SOUTH LUZON', 'OCCIDENTAL MINDORO': 'SOUTH LUZON', 'ORIENTAL MINDORO': 'SOUTH LUZON', 
        'QUEZON PROVINCE': 'SOUTH LUZON', 'RIZAL': 'SOUTH LUZON',
        'AKLAN': 'VISAYAS', 'ANTIQUE': 'VISAYAS', 'BOHOL': 'VISAYAS', 'CAPIZ': 'VISAYAS', 
        'CEBU': 'VISAYAS', 'GUIMARAS': 'VISAYAS', 'ILOILO': 'VISAYAS', 'LEYTE': 'VISAYAS', 
        'NEGROS OCCIDENTAL': 'VISAYAS', 'NEGROS ORIENTAL': 'VISAYAS',
        'AGUSAN DEL SUR': 'MINDANAO', 'BUKIDNON': 'MINDANAO', 'COTABATO': 'MINDANAO', 
        'DAVAO DE ORO': 'MINDANAO', 'DAVAO DEL NORTE': 'MINDANAO', 'DAVAO DEL SUR': 'MINDANAO', 
        'DAVAO ORIENTAL': 'MINDANAO', 'LANAO DEL NORTE': 'MINDANAO', 'LANAO DEL SUR': 'MINDANAO', 
        'MAGUINDANAO': 'MINDANAO', 'MISAMIS OCCIDENTAL': 'MINDANAO', 'MISAMIS ORIENTAL': 'MINDANAO', 
        'SARANGANI': 'MINDANAO', 'SOUTH COTABATO': 'MINDANAO', 'SULTAN KUDARAT': 'MINDANAO', 
        'SULU': 'MINDANAO', 'SURIGAO DEL SUR': 'MINDANAO', 'TARLAC': 'MINDANAO', 'TAWI-TAWI': 'MINDANAO', 
        'ZAMBOANGA DEL NORTE': 'MINDANAO', 'ZAMBOANGA DEL SUR': 'MINDANAO', 'ZAMBOANGA SIBUGAY': 'MINDANAO'
    }

    st.divider()

    # --- INPUT SECTION ---
    src_file = st.file_uploader("Upload VRP ACCOUNTS CSV", type=['csv'], key=f"uploader_{st.session_state.uploader_key}")

    if src_file:
        df_master = pd.read_csv(src_file, dtype=str, keep_default_na=False)
        df_master.columns = df_master.columns.str.strip()
        
        # --- Auto-detect the Trans/Details-only mapping file by name ---
        TRANS_DETAILS_ONLY_FILENAME = "NEW VRP ACCOUNTS FOR TAG_PIF REVISIT- NO DL_JOREM"
        uploaded_name_clean = re.sub(r'[\s_]+', ' ', os.path.splitext(src_file.name)[0].strip().upper())
        target_name_clean = re.sub(r'[\s_]+', ' ', TRANS_DETAILS_ONLY_FILENAME.strip().upper())
    
        force_trans_details = uploaded_name_clean.startswith(target_name_clean)
        
        pasted_codes = st.text_area("Paste Reference Codes (Optional filter - one per line):", height=150)
        
        col1, col2 = st.columns([1, 1])
        with col1:
            if st.button("✨ Process and Map Data", use_container_width=True):
                st.session_state.show_account_type_modal = True
                st.session_state.process_confirm = False
        with col2:
            if st.button("🔄 Reset / Clear All", use_container_width=True):
                reset_app()

        if st.session_state.show_account_type_modal:
            left, middle, right = st.columns([1, 2, 1])
            with middle:
                st.markdown("<h3 style='text-align:center'>Select Type of Account</h3>", unsafe_allow_html=True)
                if force_trans_details:
                    st.caption("ℹ️ Recognized Trans/Details-only file — pre-selected below.")
                selected_type = st.radio(
                    "", ["DL", "Trans/Details"],
                    index=1 if force_trans_details else 0,
                    key=f"account_type_radio_{uploaded_name_clean}"
                )
                if st.button("Process Now", type="primary", use_container_width=True):
                    st.session_state.selected_type = selected_type
                    st.session_state.process_confirm = True
                    st.session_state.show_account_type_modal = False

        if st.session_state.process_confirm:
            progress_bar = st.progress(0, text="Initializing processing...")
            time.sleep(0.2)
            
            selected_type = st.session_state.selected_type
            
            progress_bar.progress(15, text="Filtering target codes...")
            if pasted_codes.strip():
                codes_list = [c.strip() for c in pasted_codes.replace(',', '\n').split('\n') if c.strip()]
                df_src = df_master[df_master['REF CODE'].isin(codes_list)].copy()
            else:
                df_src = df_master.copy()

            df_src = df_src.reset_index(drop=True)
            df_out = pd.DataFrame(columns=df_tmp.columns, index=range(len(df_src)))
            time.sleep(0.2)
            
            progress_bar.progress(35, text="Mapping source columns to template...")
            mapping = {
                'ACCOUNT NUMBER': 'account_no', 'BANK': 'bank', 'PLACEMENT': 'placement',
                'CH CODE': 'ch_code', 'CH NAME': 'ch_name', 'ADD TYPE': 'address_type',
                'ADDRESS': 'address', 'MUNICIPALITY': 'municipality', 'DL TYPE': 'dl_type',
                'REF CODE': 'ref_code', 'AUTOFIELD DATE': 'autofield_date',
                'PULLOUT DATE': 'pullout_date', 'ENDO DATE': 'endorsement_date', 'AREA': 'area'
            }
            
            for src_col, target_col in mapping.items():
                if src_col in df_src.columns:
                    if target_col == 'account_no':
                        df_out[target_col] = df_src[src_col].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
                        df_out[target_col] = df_out[target_col].replace(['nan', 'None', ''], '0')
                    elif target_col == 'address':
                        df_out[target_col] = (
                            df_src[src_col]
                            .fillna('')
                            .str.replace('Blk', 'Block', case=False, regex=False)
                            .str.replace('BRGY', 'BARANGAY', case=False, regex=False)
                            .str.replace(r'\bSTA\b\.?', 'SANTA', case=False, regex=True)
                            .str.replace(r'\bSTO\b\.?', 'SANTO', case=False, regex=True)
                            .str.upper()
                        )
                    else:
                        df_out[target_col] = df_src[src_col]
            
            time.sleep(0.2)

            progress_bar.progress(60, text="Calculating constants, amounts, and CMS IDs...")
            
            # --- MAP OUTSTANDING BALANCE AND AMOUNT DUE VIA CH CODE ---
            if 'CH CODE' in df_src.columns:
                ch_codes_clean = df_src['CH CODE'].astype(str).str.strip().str.upper()
                
                # Fetch amounts from amounts.csv via mapping dictionary
                mapped_ob = ch_codes_clean.map(amounts_ob_mapping).fillna('')
                mapped_due = ch_codes_clean.map(amounts_due_mapping).fillna('')
                
                df_out['outstanding_balance'] = mapped_ob
                if 'amount_due' in df_out.columns:
                    df_out['amount_due'] = mapped_due
                
                # --- CMS ID Mapping ---
                df_out['cms_id'] = ch_codes_clean.map(cms_mapping).fillna('')
                
                # Halt process if ANY CMS ID is blank and show a warning
                blank_cms_mask = df_out['cms_id'] == ''
                if blank_cms_mask.any():
                    # Get the missing CH CODES to tell the user what to fix
                    missing_ch_codes = sorted(set(df_src.loc[blank_cms_mask, 'CH CODE'].astype(str).str.strip()))
                    
                    # Clear the progress bar and reset the processing state
                    progress_bar.empty()
                    st.session_state.process_confirm = False
                    
                    # Display the pop-up error and completely halt the script
                    st.error(
                        f"🚨 **PROCESS HALTED:** Found {blank_cms_mask.sum()} row(s) with a blank CMS ID.\n\n"
                        f"**Missing CH CODE(s):** {', '.join(missing_ch_codes)}\n\n"
                        f"Please update your 'cmd_id.xlsx' reference file and try again."
                    )
                    st.stop()

            df_out['shared_or_exclusive'] = "SHARED"
            df_out['type_of_account'] = selected_type
            df_out['month'] = datetime.now().strftime('%B').upper()
            df_out['account_type'] = "HOUSING"
            df_out['form_code'] = "vid04qNT"

            if 'FINAL AREA' in df_src.columns:
                df_out['final_area'] = df_src['FINAL AREA'].replace('', pd.NA).fillna(df_src['AREA'])
            else:
                df_out['final_area'] = df_src['AREA']
            
            df_out['area_cluster'] = df_src['AREA'].str.strip().str.upper().map(AREA_MAPPING)

            def determine_visit_type(row):
                if selected_type == "Trans/Details":
                    return "REGULAR"
                
                remark = str(row.get('FIELD REMARKS', '')).strip().upper()
                bank = str(row.get('BANK', '')).strip().upper()
                
                if "PIF FORECLOSURE" in bank: 
                    return "OTS"
                
                regular_banks = ["CBS HOUSING LOAN", "PIF PROVIDENT", "SBF MORTGAGE", "UBP HOME MORTGAGE", "SBC HOME LOAN", "SBF HOMELOAN"]
                if any(target in bank for target in regular_banks): 
                    return "REGULAR"
                
                if remark in ["NEW ENDO", "REVISIT", "NEW ENDO-PRI ADD"]: 
                    return "REGULAR"
                    
                if remark == "CARAVAN": 
                    return "CARAVAN"
                
                return ""
                
            df_out['visit_type'] = df_src.apply(determine_visit_type, axis=1)

            progress_bar.progress(80, text="Formatting dates...")
            date_fields = ['autofield_date', 'endorsement_date', 'pullout_date']
            for field in date_fields:
                st.session_state.processed_data = None  
                df_out[field] = pd.to_datetime(df_out[field], errors='coerce').dt.strftime('%d-%m-%Y')
                df_out[field] = df_out[field].fillna('')
            time.sleep(0.2)

            progress_bar.progress(95, text="Generating final files...")
            total_accounts = len(df_out)
            if not df_src.empty:
                unique_banks = df_src['BANK'].astype(str).str.strip().str.upper().replace({'NAN': ''}).unique()
                unique_banks = [b for b in unique_banks if b]
                if len(unique_banks) == 1 and unique_banks[0] == "PIF HOMELOAN":
                    bank_val = "PIF HOME LOAN"
                elif len(unique_banks) == 1:
                    bank_val = unique_banks[0]
                else:
                    bank_val = "MC2 OTHERS"
            else:
                bank_val = "FILTERED"
            
            st.session_state.generated_filename = f"{bank_val}_{total_accounts}.csv"
            st.session_state.release_filename = f"Released_to_{bank_val}_{total_accounts}.csv"
            st.session_state.processed_data = df_out
            
            df_rel = pd.DataFrame()
            df_rel['REF CODE'] = df_src['REF CODE']
            df_rel['RELEASED TO'] = df_src['RELEASED TO']
            st.session_state.release_data = df_rel
            
            progress_bar.progress(100, text="Done!")
            time.sleep(0.3)
            progress_bar.empty()

    if st.session_state.processed_data is not None:
        df_out = st.session_state.processed_data
        df_release = st.session_state.release_data
        main_fn = st.session_state.generated_filename
        rel_fn = st.session_state.release_filename
        total_accounts = len(df_out)

        _, center_box, _ = st.columns([1, 1, 1])
        with center_box:
            st.success(f"✅ Processed {total_accounts} accounts!")

        st.write("### 📥 Download Processed Files")
        col_dl1, col_dl2 = st.columns(2)
        
        with col_dl1:
            csv_main = df_out.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📥 Main CSV ({main_fn})", 
                data=csv_main, 
                file_name=main_fn, 
                mime="text/csv", 
                use_container_width=True
            )
            
        with col_dl2:
            csv_rel = df_release.to_csv(index=False).encode('utf-8')
            st.download_button(
                label=f"📊 Released to ({rel_fn})", 
                data=csv_rel, 
                file_name=rel_fn, 
                mime="text/csv",
                use_container_width=True
            )

        st.divider()
        st.subheader("📋 Main Data Preview")
        st.dataframe(df_out.head(10))
        st.subheader("📑 Release File Preview")
        st.dataframe(df_release.head(10))


# ==========================================
# TOOL 2: FIELD RESULT (REPLICATED VBA MACRO)
# ==========================================
elif selected_tool == "Field Result":
    st.title("📊 Field Result Column Extractor")
    st.write("Upload an Excel workbook containing a sheet named **'RESULT'** to run the macro extraction.")
    
    excel_file = st.file_uploader("Upload Workbook (.xlsx)", type=['xlsx'])
    
    if excel_file:
        try:
            progress_bar = st.progress(0, text="Reading Excel file...")
            
            xl_file = pd.ExcelFile(excel_file)
            sheet_names = xl_file.sheet_names
            target_sheet = next((sheet for sheet in sheet_names if sheet.upper() == "RESULT"), None)
            
            if not target_sheet:
                progress_bar.empty()
                st.error("❌ Error: A sheet named 'RESULT' was not found in this workbook.")
            else:
                progress_bar.progress(25, text="Extracting relevant columns...")
                time.sleep(0.2)
                
                df_source = pd.read_excel(excel_file, sheet_name=target_sheet, header=None, dtype=str)
                
                max_required_col_idx = 41
                if df_source.shape[1] <= max_required_col_idx:
                    for i in range(df_source.shape[1], max_required_col_idx + 1):
                        df_source[i] = ""
                
                df_target = pd.DataFrame()
                df_target[0] = df_source[41]
                df_target[1] = df_source[2]
                df_target[2] = df_source[3]
                df_target[3] = df_source[4]
                df_target[4] = df_source[27]
                df_target[5] = df_source[32]
                df_target[6] = df_source[36]
                
                progress_bar.progress(50, text="Converting dates to native objects...")
                time.sleep(0.2)
                
                def format_field_result_date(val):
                    if pd.isna(val): 
                        return val
                    val_str = str(val).strip()
                    if re.match(r'^\d{4}-\d{2}-\d{2}( \d{2}:\d{2}:\d{2}(\.\d+)?)?$', val_str) or \
                       re.match(r'^\d{1,2}/\d{1,2}/\d{4}( \d{2}:\d{2}:\d{2}(\.\d+)?)?$', val_str):
                        try:
                            return pd.to_datetime(val_str).date()
                        except:
                            return val
                    return val

                for col in df_target.columns:
                    df_target[col] = df_target[col].apply(format_field_result_date)
                
                progress_bar.progress(75, text="Generating stylized Excel with mm/dd/yyyy format...")
                time.sleep(0.2)
                
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = "RESULT"
                calibri_font = Font(name="Calibri", size=9)
                
                for r_idx, row_values in enumerate(df_target.values):
                    for c_idx, cell_value in enumerate(row_values):
                        
                        # --- UPDATED: FORMAT SET TO mm/dd/yyyy ---
                        if isinstance(cell_value, (date, datetime)):
                            formatted_text_value = cell_value.strftime('%m/%d/%Y')
                        else:
                            formatted_text_value = cell_value
                            
                        cell = ws_out.cell(row=r_idx + 1, column=c_idx + 1, value=formatted_text_value)
                        cell.font = calibri_font
                
                for col in ws_out.columns:
                    max_len = 0
                    col_letter = get_column_letter(col[0].column)
                    for cell in col:
                        if cell.value is not None:
                            max_len = max(max_len, len(str(cell.value)))
                    ws_out.column_dimensions[col_letter].width = max(max_len + 3, 10)
                
                # --- Filename uses dashes to prevent OS errors ---
                current_date_str = datetime.now().strftime('%m-%d-%Y')
                generated_fn = f"FIELD RESULT {current_date_str}.xlsx"
                
                output_buffer = io.BytesIO()
                wb_out.save(output_buffer)
                output_buffer.seek(0)
                
                progress_bar.progress(100, text="Process Complete!")
                time.sleep(0.3)
                progress_bar.empty()
                
                st.success("✅ Extraction and formatting (mm/dd/yyyy) successful!")
                
                st.download_button(
                    label=f"📥 Download Extracted File ({generated_fn})",
                    data=output_buffer,
                    file_name=generated_fn,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                
                st.divider()
                st.subheader("📋 Extracted Data Preview")
                
                preview_df = df_target.copy()
                preview_df.columns = preview_df.iloc[0]
                preview_df = preview_df.drop(preview_df.index[0]).reset_index(drop=True)
                st.dataframe(preview_df.head(15))
                
        except Exception as e:
            st.error(f"❌ Error processing file: {str(e)}")


# ==========================================
# TOOL 3: E-SIGN FIXER
# ==========================================
elif selected_tool == "E-SIGN FIXER":
    st.title("✒️ E-SIGN FIXER")
    st.write("Upload an image page containing multiple sheet signatures to crop them out into isolated transparent files.")
    
    TARGET_SIZE = 250 

    uploaded_file = st.file_uploader("Upload Image", type=["png", "jpg", "jpeg"])
    
    if uploaded_file:
        try:
            progress_bar = st.progress(0, text="Loading image...")
            
            image = Image.open(uploaded_file).convert("RGB")
            img = np.array(image)
            
            with st.expander("🖼️ Show Original Image", expanded=False):
                st.image(img, width=350)
            
            progress_bar.progress(25, text="Processing...")
            time.sleep(0.25)
            
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)
            _, thresh = cv2.threshold(gray, 220, 255, cv2.THRESH_BINARY_INV)
            
            kernel = np.ones((3, 3), np.uint8)
            thresh = cv2.dilate(thresh, kernel, iterations=1)
            
            progress_bar.progress(50, text="Extracting signatures...")
            time.sleep(0.25)
            
            contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
            contours = sorted(contours, key=lambda c: (cv2.boundingRect(c)[1], cv2.boundingRect(c)[0]))
            
            extracted_signatures = []
            
            progress_bar.progress(75, text="Cropping...")
            time.sleep(0.25)
            
            for cnt in contours:
                area = cv2.contourArea(cnt)
                if area < 1000:
                    continue
                    
                x, y, w, h = cv2.boundingRect(cnt)
                padding = 15
                
                x1 = max(0, x - padding)
                y1 = max(0, y - padding)
                x2 = min(img.shape[1], x + w + padding)
                y2 = min(img.shape[0], y + h + padding)
                
                crop = img[y1:y2, x1:x2]
                mask = np.any(crop < 240, axis=2).astype(np.uint8) * 255
                ys, xs = np.where(mask > 0)
                
                if len(xs) == 0 or len(ys) == 0:
                    continue
                    
                crop = crop[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
                mask = mask[ys.min():ys.max() + 1, xs.min():xs.max() + 1]
                
                gray_crop = cv2.cvtColor(crop, cv2.COLOR_RGB2GRAY)
                rgba = np.zeros((gray_crop.shape[0], gray_crop.shape[1], 4), dtype=np.uint8)
                rgba[:, :, 0:3] = gray_crop[:, :, np.newaxis]
                rgba[:, :, 3] = mask
                
                h_sig, w_sig = rgba.shape[:2]
                max_dim = max(h_sig, w_sig)
                square_canvas = np.zeros((max_dim, max_dim, 4), dtype=np.uint8)
                y_offset = (max_dim - h_sig) // 2
                x_offset = (max_dim - w_sig) // 2
                square_canvas[y_offset:y_offset+h_sig, x_offset:x_offset+w_sig] = rgba
                
                final_sig = cv2.resize(square_canvas, (TARGET_SIZE, TARGET_SIZE), interpolation=cv2.INTER_AREA)
                extracted_signatures.append(final_sig)
            
            progress_bar.progress(90, text="Generating files...")
            
            if extracted_signatures:
                zip_buffer = BytesIO()
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
                    for idx, sig in enumerate(extracted_signatures, start=1):
                        filename = f"signature_{idx}.PNG"
                        img_byte_arr = io.BytesIO()
                        Image.fromarray(sig).save(img_byte_arr, format='PNG')
                        zipf.writestr(filename, img_byte_arr.getvalue())
                
                doc = Document()
                doc.add_heading('Extracted E-Signatures', 0)
                for idx, sig in enumerate(extracted_signatures, start=1):
                    img_byte_arr = io.BytesIO()
                    Image.fromarray(sig).save(img_byte_arr, format='PNG')
                    img_byte_arr.seek(0)
                    doc.add_paragraph(f'Signature {idx}:')
                    doc.add_picture(img_byte_arr, width=Inches(2.5))
                
                word_buffer = io.BytesIO()
                doc.save(word_buffer)
                
                progress_bar.progress(100, text="Complete!")
                st.success(f"✅ Extracted {len(extracted_signatures)} signatures.")
                
                c1, c2 = st.columns(2)
                with c1:
                    st.download_button("📦 ZIP Images", zip_buffer.getvalue(), "signatures.zip", "application/zip", use_container_width=True)
                with c2:
                    st.download_button("📝 Word Document", word_buffer.getvalue(), "signatures.docx", use_container_width=True)
                
                st.divider()
                cols = st.columns(4)
                for idx, sig in enumerate(extracted_signatures):
                    cols[idx % 4].image(sig, caption=f"Sig {idx+1}")
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
