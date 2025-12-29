import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import shutil
import os
import io

# Page Config
st.set_page_config(page_title="Noon Price Comp", layout="centered")
st.title("📊 Partner Price Comp Tool")

# 1. Initialize Session State for Data Persistence
if 'summary_df' not in st.session_state:
    st.session_state.summary_df = None
if 'zip_buffer' not in st.session_state:
    st.session_state.zip_buffer = None
if 'zip_name' not in st.session_state:
    st.session_state.zip_name = None
if 'partner_count' not in st.session_state:
    st.session_state.partner_count = 0

SELECTED_COLS = [
    "Psku", "SKU", "Title En", "Comp Link", 
    "Latest Comp Price All", "Offer Price", 
    "Adjustment needed", "Comp Bb Seller Name", "noon link"
]

# 2. File Upload (Supports CSV & Excel)
uploaded_file = st.file_uploader("Upload Master File (Excel or CSV)", type=["xlsx", "csv"])

if uploaded_file:
    if st.button("🚀 Process & Generate ZIP"):
        with st.spinner('Processing files...'):
            # Detect file type
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
            else:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                
            df.columns = [c.strip() for c in df.columns]
            
            # Filter for NC/NCO
            mask = df["Price Comp Bucket"].str.upper().isin({"NC", "NCO"})
            filtered = df.loc[mask].copy()
            
            if filtered.empty:
                st.error("No 'NC' or 'NCO' rows found.")
            else:
                # Calculations
                filtered["Latest Comp Price All"] = pd.to_numeric(filtered["Latest Comp Price All"], errors='coerce')
                filtered["Offer Price"] = pd.to_numeric(filtered["Offer Price"], errors='coerce')
                filtered["Adjustment needed"] = filtered["Offer Price"] - filtered["Latest Comp Price All"]
                
                # Hyperlinks
                sku_cfg = filtered["SKU Config"].astype(str)
                filtered["noon link"] = '=HYPERLINK("http://noon.com/egypt-en/' + sku_cfg + '/p/", "View on Noon")'
                filtered["Comp Link"] = '=HYPERLINK("' + filtered["Comp Link"].astype(str) + '", "View Competitor")'

                # Setup Temporary Directory
                folder_name = f"Comp_{datetime.now():%d-%m}"
                out_dir = Path(folder_name)
                if out_dir.exists(): shutil.rmtree(out_dir)
                out_dir.mkdir()

                # 3. Create Master Summary (Standalone Excel)
                st.session_state.summary_df = (
                    filtered.groupby(["ID Partner", "Partner Name"])["SKU"]
                    .nunique()
                    .reset_index(name="NC_NCO_SKU_Count")
                    .sort_values(by="NC_NCO_SKU_Count", ascending=False)
                )
                st.session_state.summary_df.to_excel(out_dir / "Partner_PriceComp_Counts.xlsx", index=False)

                # 4. Partner Export Loop
                grouped = filtered.groupby("ID Partner", sort=False)
                st.session_state.partner_count = len(grouped)
                
                for pid, grp in grouped:
                    p_name = grp["Partner Name"].iloc[0] if "Partner Name" in grp else f"Partner_{pid}"
                    safe_name = "".join(c for c in str(p_name) if c.isalnum() or c in " -_").strip()[:50]
                    fname = out_dir / f"{safe_name}_PriceComp.xlsx"
                    
                    with pd.ExcelWriter(fname, engine='openpyxl') as writer:
                        cols_to_write =
