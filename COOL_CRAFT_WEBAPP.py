# COOL_CRAFT_WEBAPP.py
import streamlit as st
import pandas as pd
import io
import re
from math import ceil
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import MergedCell

# ================ PAGE SETUP ================
st.set_page_config(page_title="CoolCraft VRF TDS Generator", layout="wide")

# ================ CSS ================
st.markdown("""
<style>
.header {font-size:28px; font-weight:700; color:#fff; text-align:center; padding:14px;
 background:linear-gradient(90deg,#0047ab,#e63946);}
.card {background:#fff; padding:14px; border-radius:10px; box-shadow:0 1px 6px rgba(0,0,0,0.06);}
</style>
""", unsafe_allow_html=True)
st.markdown('<div class="header">❄️ CoolCraft TDS – HVAC Technical Data Sheet Generator</div>', unsafe_allow_html=True)

# Conversion constants
KW_TO_HP = 1.0 / 0.745699872
TON_TO_HP = 3.517 / 0.745699872

# ================ HELPERS ================
@st.cache_data(ttl=600)
def load_excel_all_sheets(uploaded_file):
    return pd.read_excel(uploaded_file, sheet_name=None)

def normalize_name(s: str) -> str:
    if s is None:
        return ""
    s = str(s).strip().lower()
    s = re.sub(r'\s+', ' ', s)
    s = s.replace('\u200b', '')
    return s

def build_normalized_map(df: pd.DataFrame) -> dict:
    return {orig: normalize_name(orig) for orig in df.columns}

def find_capacity_column_by_type(df: pd.DataFrame, unit_type: str) -> str:
    norm_map = build_normalized_map(df)
    if unit_type.lower() == "indoor":
        for orig, norm in norm_map.items():
            if "cooling capacity" in norm and "kw" in norm: return orig
        for orig, norm in norm_map.items():
            if "capacity" in norm and "kw" in norm: return orig
        for orig, norm in norm_map.items():
            if "kw" in norm: return orig
    else:
        for orig, norm in norm_map.items():
            if "hp" in norm and ("capacity" in norm or "hp" in norm): return orig
        for orig, norm in norm_map.items():
            if "horsepower" in norm or re.search(r'\bhp\b', norm): return orig
        for orig, norm in norm_map.items():
            if "capacity" in norm and ("hp" in norm or "horsepower" in norm): return orig
    return None

def expand_combo_instances(combo):
    inst = []
    for hp, cnt in sorted(combo.items(), reverse=True):
        inst.extend([hp] * cnt)
    return inst

def greedy_combo(target_cap, sizes):
    rem = int(round(target_cap))
    combo = {}
    for s in sizes:
        cnt = rem // s
        if cnt > 0:
            combo[s] = int(cnt)
            rem -= s * cnt
    if rem > 0 and sizes:
        smallest = min(sizes)
        combo[smallest] = combo.get(smallest, 0) + 1
    return combo

def generate_candidate_combos(target_cap, sizes, max_suggestions=12):
    # enhanced algorithm: try exact match first
    exact_match = {s: int(target_cap // s) for s in sizes if target_cap % s == 0 and target_cap // s > 0}
    raw = []
    if exact_match: raw.append(exact_match)
    raw.append(greedy_combo(target_cap, sizes))
    for s in sizes[:6]:
        raw.append({s: ceil(target_cap / s)})
    uniq = {tuple(sorted(c.items(), reverse=True)): c for c in raw}
    combos = list(uniq.values())[:max_suggestions]
    scored = []
    for c in combos:
        total = sum(expand_combo_instances(c))
        units = sum(c.values())
        closeness = 1.0 / (1 + abs(total - target_cap) / max(1, target_cap))
        unit_score = 1.0 / (1 + units)
        score = 0.6 * closeness + 0.4 * unit_score
        scored.append((score, c))
    scored.sort(key=lambda x: x[0], reverse=True)
    return [c for _, c in scored]

def find_nearest_row(df, target_val, cap_col):
    diffs = (pd.to_numeric(df[cap_col], errors='coerce') - float(target_val)).abs()
    idx = diffs.idxmin()
    return df.loc[idx] if pd.notna(idx) else None

def export_excel(df: pd.DataFrame, sheet_name='VRF_TDS_Report'):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    thin = Side(border_style="thin", color="000000")
    border_all = Border(left=thin, right=thin, top=thin, bottom=thin)
    center_align = Alignment(horizontal='center', vertical='center')
    header_font = Font(bold=True, color='FFFFFF')

    col_count = df.shape[1] if df.shape[1] > 0 else 1
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=col_count)
    title_cell = ws.cell(row=1, column=1, value="❄️ CoolCraft Technical Data Sheet (TDS)")
    title_cell.font = Font(size=14, bold=True, color="FFFFFF")
    title_cell.alignment = center_align
    title_cell.fill = PatternFill(start_color='4B8BBE', end_color='4B8BBE', fill_type='solid')

    for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), start=2):
        for c_idx, value in enumerate(row, start=1):
            cell = ws.cell(row=r_idx, column=c_idx, value=value)
            if r_idx == 2:
                cell.font = header_font
                cell.fill = PatternFill(start_color='2F5597', end_color='2F5597', fill_type='solid')
                cell.alignment = center_align
            else:
                fill_color = "EAF1FB" if r_idx % 2 == 0 else "FFFFFF"
                if isinstance(value, (int, float)):
                    if value >= 90: fill_color = "1e824c"
                    elif value >= 70: fill_color = "f39c12"
                    else: fill_color = "c0392b"
                cell.fill = PatternFill(start_color=fill_color, end_color=fill_color, fill_type='solid')
            cell.border = border_all

    for col_cells in ws.columns:
        length = max(len(str(cell.value or "")) for cell in col_cells)
        first_cell = next((cell for cell in col_cells if not isinstance(cell, MergedCell)), None)
        if first_cell:
            try: ws.column_dimensions[first_cell.column_letter].width = min(length + 3, 50)
            except: pass

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio

# ================ SIDEBAR / WIZARD ================
st.sidebar.header("Start New TDS")
brand = st.sidebar.selectbox("Brand", ["Toshiba", "Carrier"])
system_type = st.sidebar.selectbox("System Type", ["VRF", "Non-VRF"])
product_type = st.sidebar.selectbox("Product Type", ["Cassette", "High Wall", "Ductable", "Other"])
unit_type = st.sidebar.selectbox("Unit Type", ["Indoor", "Outdoor"])
combination_mode = st.sidebar.selectbox("Combination Mode", ["Automatic", "Manual"])
combination_type = None
if unit_type == "Outdoor" and product_type == "Other":
    combination_type = st.sidebar.selectbox("Outdoor selection", ["Combination", "High Efficiency", "Single Unit"])

uploaded_file = st.file_uploader("Upload Excel TDS", type=["xlsx"])
if uploaded_file is None:
    st.warning("Please upload an Excel file to proceed.")
    st.stop()

sheets = load_excel_all_sheets(uploaded_file)
sheet_choice = st.selectbox("Select sheet", list(sheets)) if len(sheets) > 1 else list(sheets)[0]
df = sheets[sheet_choice].copy()

# Detect capacity column
cap_label_type = "HP" if unit_type=="Outdoor" else "kW"
cap_col = find_capacity_column_by_type(df, unit_type)
if cap_col is None:
    st.warning(f"No capacity column ({cap_label_type}) found. Automatic combo disabled.")
    sizes_available = []
else:
    df[cap_col] = pd.to_numeric(df[cap_col], errors='coerce')
    sizes_available = sorted(list({float(x) for x in df[cap_col].dropna().unique()}))

st.subheader("Loaded Dataset Preview")
st.dataframe(df.head())

# ================ COMBINATION GENERATION ================
if combination_mode == "Automatic" and sizes_available:
    unit_input = st.radio("Provide load in:", [cap_label_type, "kW" if cap_label_type=="HP" else "HP", "Ton"], horizontal=True)
    default_value = 100.0 if cap_label_type=="HP" else 10.0
    load_val = st.number_input(f"Enter load value ({unit_input})", min_value=0.1, value=default_value, step=0.1)
    if st.button("Generate Combos"):
        # convert to target in cap units
        if cap_label_type == "HP":
            target_cap = load_val if unit_input=="HP" else (load_val * KW_TO_HP if unit_input=="kW" else load_val * TON_TO_HP)
        else:
            target_cap = load_val if unit_input=="kW" else (load_val * 0.745699872 if unit_input=="HP" else load_val * 3.517)
        enriched = []
        sizes_desc = sorted(sizes_available, reverse=True)
        normalized_sizes = [int(s) if float(s).is_integer() else float(s) for s in sizes_desc]
        for combo in generate_candidate_combos(target_cap, normalized_sizes):
            rows = []
            for idx, cap in enumerate(expand_combo_instances(combo)):
                exact = df[pd.to_numeric(df[cap_col], errors='coerce') == float(cap)]
                if not exact.empty:
                    chosen = exact.iloc[0].to_dict()
                else:
                    nearest = find_nearest_row(df, cap, cap_col)
                    chosen = nearest.to_dict() if nearest is not None else {}
                chosen['_instance'] = idx + 1
                rows.append(chosen)
            total_cap = sum(pd.to_numeric([r.get(cap_col, 0) for r in rows], errors='coerce'))
            enriched.append({'combo': combo, 'rows': rows, 'total_cap': total_cap, 'units': len(rows)})
        st.session_state['enriched'] = enriched
else:
    st.info("Manual combination mode or no numeric capacity column detected.")

# ================ SHOW & EXPORT ================
if 'enriched' in st.session_state:
    enriched = st.session_state['enriched']
    for i, e in enumerate(enriched, 1):
        desc = " + ".join(f"{v}×{k}{cap_label_type}" for k, v in e['combo'].items())
        st.markdown(f"**Option {i}**: {desc} — Units: {e.get('units',0)} — Total: {round(e.get('total_cap',0), 3)} {cap_label_type}")
        st.dataframe(pd.DataFrame(e['rows']).head(10))

    choice = st.selectbox("Choose option", range(1, len(enriched) + 1)) - 1
    chosen = enriched[choice]

    mod_rows = []
    for r in chosen['rows']:
        inst = int(r.get('_instance', 0))
        sel_row = r.copy()
        if 'model' in df.columns and cap_col in r and pd.notna(r.get(cap_col, None)):
            model_list = df[pd.to_numeric(df[cap_col], errors='coerce') == float(r[cap_col])]['model'].dropna().unique().tolist()
            if not model_list: model_list = df['model'].dropna().unique().tolist()
            model_list = sorted(model_list)
            sel = st.selectbox(f"Instance {inst}", model_list, key=f"ov_{inst}")
            matches = df[df['model'] == sel]
            if not matches.empty:
                sel_row = matches.iloc[0].to_dict()
                sel_row['_instance'] = inst
        else:
            cap_display = r.get(cap_col, r.get('_cap_input', 'N/A'))
            st.markdown(f"Instance {inst} — {cap_label_type}: {cap_display}")
            sel_row['_instance'] = inst
        mod_rows.append(sel_row)

    out_df = pd.DataFrame(mod_rows)

    # metadata
    client = st.text_input("Client Name", "Client")
    manuf = st.text_input("Manufacturer", brand)
    billing = st.text_input("Billing/Sales", "")
    rdate = st.date_input("Report Date", datetime.now().date())
    out_df['Client'] = client
    out_df['Manufacturer'] = manuf
    out_df['BillingSales'] = billing
    out_df['ReportDate'] = rdate.strftime('%Y-%m-%d')
    out_df['SelectedCombo'] = " + ".join(f"{v}×{k}{cap_label_type}" for k, v in chosen['combo'].items())
    out_df[f'ComboTotal{cap_label_type}'] = chosen.get('total_cap', 0)
    out_df['ComboUnits'] = chosen.get('units', 0)

    meta_cols = ['Client', 'Manufacturer', 'BillingSales', 'ReportDate', 'SelectedCombo', f'ComboTotal{cap_label_type}', 'ComboUnits']
    extra_cols = [c for c in df.columns if c in out_df.columns and c not in meta_cols]
    final_cols = meta_cols + extra_cols if extra_cols else meta_cols
    out_df = out_df[final_cols]

    st.subheader("TDS Preview with Ratings")
    num_cols = list(out_df.select_dtypes(include=['number']).columns)
    def style_func(x):
        if isinstance(x, (int, float)):
            if x >= 90: return 'background-color: #1e824c; color:white'
            elif x >= 70: return 'background-color: #f39c12; color:white'
            else: return 'background-color: #c0392b; color:white'
        return ''
    st.dataframe(out_df.style.applymap(style_func, subset=num_cols) if num_cols else out_df)

    if st.button("Download Excel"):
        bytes_xl = export_excel(out_df)
        fname = f"TDS_{client.replace(' ','_')}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        st.download_button("📥 Download", data=bytes_xl, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
