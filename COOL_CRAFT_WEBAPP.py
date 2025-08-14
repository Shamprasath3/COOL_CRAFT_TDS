import streamlit as st
import pandas as pd
import os
import io
import re
from math import ceil
from datetime import datetime
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.cell.cell import MergedCell

# ------------------ PAGE SETUP ------------------
st.set_page_config(page_title="CoolCraft VRF TDS Generator", layout="wide")

st.markdown("""
<style>
.header {font-size:28px; font-weight:700; color:#fff; text-align:center; padding:14px;
 background:linear-gradient(90deg,#0047ab,#e63946);}
.card {background:#fff; padding:14px; border-radius:10px; box-shadow:0 1px 6px rgba(0,0,0,0.06); margin-bottom:10px;}
</style>
""", unsafe_allow_html=True)

st.markdown('<div class="header">❄️ CoolCraft TDS – HVAC Technical Data Sheet Generator</div>', unsafe_allow_html=True)

# ------------------ DATA SOURCES ------------------
DATA_SOURCES = {
    # Replace with relative paths for deployment
    ("Toshiba", "VRF", "Other", "Outdoor Unit", "Single Unit"): "data/TOS_VRF_SINGLE.xlsx",
    ("Toshiba", "VRF", "Other", "Outdoor Unit", "High Efficiency"): "data/TOS_VRF_HI_EFM.xlsx",
    ("Toshiba", "VRF", "Other", "Outdoor Unit", "Combination"): "data/TOS_VRF_COMB.xlsx",
    ("Toshiba", "VRF", "Cassette", "Indoor Unit", None): "data/4-way Cassette U series.xlsx",
    ("Toshiba", "VRF", "High Wall", "Indoor Unit", None): "data/Highwall_USeries.xlsx",
    ("Toshiba", "VRF", "Ductable", "Indoor Unit", None): "data/HS Ductable TDS.xlsx",
}

KW_TO_HP = 1.0 / 0.745699872
TON_TO_HP = 3.517 / 0.745699872

# ------------------ HELPERS ------------------
@st.cache_data(ttl=600)
def load_excel_all_sheets(path: str):
    if os.path.exists(path):
        xls = pd.ExcelFile(path)
    else:
        fallback = os.path.join("/mnt/data", os.path.basename(path))
        if os.path.exists(fallback):
            xls = pd.ExcelFile(fallback)
        else:
            raise FileNotFoundError(f"File not found: {path}")
    return {sheet: pd.read_excel(xls, sheet_name=sheet) for sheet in xls.sheet_names}

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
        inst.extend([hp]*cnt)
    return inst

def greedy_combo(target_cap, sizes):
    rem = int(round(target_cap))
    combo = {}
    for s in sizes:
        cnt = rem // s
        if cnt>0:
            combo[s]=int(cnt)
            rem-=s*cnt
    if rem>0 and sizes:
        smallest=min(sizes)
        combo[smallest]=combo.get(smallest,0)+1
    return combo

def generate_candidate_combos(target_cap, sizes, max_suggestions=12):
    raw=[greedy_combo(target_cap,sizes)]
    for s in sizes[:6]: raw.append({s:ceil(target_cap/s)})
    uniq={tuple(sorted(c.items(), reverse=True)): c for c in raw}
    combos=list(uniq.values())[:max_suggestions]
    scored=[]
    for c in combos:
        total=sum(expand_combo_instances(c))
        units=sum(c.values())
        closeness=1.0/(1+abs(total-target_cap)/max(1,target_cap))
        unit_score=1.0/(1+units)
        score=0.6*closeness+0.4*unit_score
        scored.append((score,c))
    scored.sort(key=lambda x:x[0],reverse=True)
    return [c for _,c in scored]

def find_nearest_row(df,target_val,cap_col):
    diffs=(pd.to_numeric(df[cap_col],errors='coerce')-float(target_val)).abs()
    idx=diffs.idxmin()
    return df.loc[idx] if pd.notna(idx) else None

# ------------------ METADATA PANEL ------------------
st.sidebar.header("TDS Metadata")
client = st.sidebar.text_input("Client Name", "Client")
manuf = st.sidebar.text_input("Manufacturer", "Toshiba")
billing = st.sidebar.text_input("Billing/Sales", "")
rdate = st.sidebar.date_input("Report Date", datetime.now().date())

# ------------------ WIZARD ------------------
st.sidebar.header("Start New TDS")
brand = st.sidebar.selectbox("Brand", ["Toshiba","Carrier"])
system_type = st.sidebar.selectbox("System Type", ["VRF","Non-VRF"])
product_type = st.sidebar.selectbox("Product Type", ["Cassette","High Wall","Ductable","Other"])
unit_type = st.sidebar.selectbox("Unit Type", ["Indoor","Outdoor"])
combination_mode = st.sidebar.selectbox("Combination Mode", ["Automatic","Manual"])
combination_type=None
if unit_type=="Outdoor" and product_type=="Other":
    combination_type=st.sidebar.selectbox("Outdoor selection",["Combination","High Efficiency","Single Unit"])

if st.sidebar.button("Proceed"):
    st.session_state['wizard']={'brand':brand,'system_type':system_type,'product_type':product_type,
                                'unit_type':unit_type,'combination_mode':combination_mode,'combination_type':combination_type}

wizard=st.session_state.get('wizard')
if not wizard: st.stop()

def map_key(w):
    if w['unit_type']=="Outdoor" and w['product_type']=="Other":
        return (w['brand'],w['system_type'],"Other","Outdoor Unit",w.get('combination_type'))
    return (w['brand'],w['system_type'],w['product_type'],"Indoor Unit",None)

path=DATA_SOURCES.get(map_key(wizard))
if not path: st.error("No mapping found for this selection."); st.stop()

sheets=load_excel_all_sheets(path)
sheet_choice=st.selectbox("Select sheet", list(sheets)) if len(sheets)>1 else list(sheets)[0]
df=sheets[sheet_choice].copy()

cap_label_type="HP" if wizard['unit_type']=="Outdoor" else "kW"
cap_col=find_capacity_column_by_type(df,wizard['unit_type'])
sizes_available=[]
if cap_col: 
    df[cap_col]=pd.to_numeric(df[cap_col],errors='coerce')
    sizes_available=sorted(list({float(x) for x in df[cap_col].dropna().unique()}))

# ------------------ TABS ------------------
tab1, tab2 = st.tabs(["Automatic Mode","Manual Mode"])

# --- AUTOMATIC MODE ---
with tab1:
    st.subheader("Automatic Combination Generation")
    if not cap_col or not sizes_available:
        st.info("Automatic mode requires a numeric capacity column.")
    else:
        unit_input=st.radio("Provide load in:",[cap_label_type,"kW" if cap_label_type=="HP" else "HP","Ton"],horizontal=True)
        default_val=100.0 if cap_label_type=="HP" else 10.0
        load_val=st.number_input(f"Enter load value ({unit_input})", min_value=0.1,value=default_val,step=0.1)
        if st.button("Generate Combos (Automatic)"):
            target_cap = load_val * KW_TO_HP if (cap_label_type=="HP" and unit_input=="kW") else load_val
            combos = generate_candidate_combos(target_cap,sizes_available)
            st.session_state['auto_combos']=combos
            for idx,c in enumerate(combos,1):
                with st.container():
                    st.markdown(f"**Option {idx}:** {' + '.join(f'{v}×{k}{cap_label_type}' for k,v in c.items())}")
                    st.dataframe(pd.DataFrame(expand_combo_instances(c),columns=[cap_label_type]))

# --- MANUAL MODE ---
with tab2:
    st.subheader("Manual Combination")
    man_in=st.text_input(f"Enter {cap_label_type} sizes (use +, e.g., 3.5+3.5+2)","")
    if st.button("Create Combo (Manual)"):
        if man_in.strip():
            sizes=[float(x) for x in man_in.split("+") if x.strip()]
            combo_dict={}
            for s in sizes:
                combo_dict[s]=combo_dict.get(s,0)+1
            st.session_state['manual_combo']={'sizes':sizes,'combo_dict':combo_dict}
            st.markdown(f"**Combo:** {' + '.join(f'{v}×{k}{cap_label_type}' for k,v in combo_dict.items())}")
            st.dataframe(pd.DataFrame({'Instance':range(1,len(sizes)+1),cap_label_type:sizes}))

