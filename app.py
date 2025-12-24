import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import shutil
import os

# Page Config
st.set_page_config(page_title="Noon Price Comp", layout="centered")
st.title("📊 Partner Price Comp Tool")
st.info("Upload the master Excel file to generate partner-specific exports.")

SELECTED_COLS = [
    "Psku", "SKU", "Title En", "Comp Link", 
    "Latest Comp Price All", "Offer Price", 
    "Adjustment needed", "Comp Bb Seller Name", "noon link"
]

uploaded_file = st.file_uploader("Upload Master Excel File", type=["xlsx"])

if uploaded_file:
    if st.button("🚀 Process & Generate ZIP"):
        with st.spinner('Processing...'):
            # 1. Read Data
            df = pd.read_excel(uploaded_file, engine='openpyxl')
            df.columns = [c.strip() for c in df.columns]
            
            # 2. Filter
            mask = df["Price Comp Bucket"].str.upper().isin({"NC", "NCO"})
            filtered = df.loc[mask].copy()
            
            if filtered.empty:
                st.error("No 'NC' or 'NCO' rows found.")
            else:
                # 3. Calculations
                filtered["Latest Comp Price All"] = pd.to_numeric(filtered["Latest Comp Price All"], errors='coerce')
                filtered["Offer Price"] = pd.to_numeric(filtered["Offer Price"], errors='coerce')
                filtered["Adjustment needed"] = filtered["Offer Price"] - filtered["Latest Comp Price All"]
                
                sku_cfg = filtered["SKU Config"].astype(str)
                filtered["noon link"] = '=HYPERLINK("http://noon.com/egypt-en/' + sku_cfg + '/p/", "View on Noon")'
                filtered["Comp Link"] = '=HYPERLINK("' + filtered["Comp Link"].astype(str) + '", "View Competitor")'

                # 4. Create Output Folder
                folder_name = f"Comp_{datetime.now():%d-%m}"
                out_dir = Path(folder_name)
                if out_dir.exists(): shutil.rmtree(out_dir)
                out_dir.mkdir()

                # 5. Group & Save
                grouped = filtered.groupby("ID Partner", sort=False)
                for pid, grp in grouped:
                    p_name = grp["Partner Name"].iloc[0] if "Partner Name" in grp else f"Partner_{pid}"
                    safe_name = "".join(c for c in str(p_name) if c.isalnum() or c in " -_").strip()[:50]
                    fname = out_dir / f"{safe_name}_PriceComp.xlsx"
                    grp.to_excel(fname, index=False, columns=[c for c in SELECTED_COLS if c in grp.columns])

                # 6. Create ZIP
                shutil.make_archive(folder_name, 'zip', out_dir)
                
                with open(f"{folder_name}.zip", "rb") as f:
                    st.success("✅ Done!")
                    st.download_button(
                        label="⬇️ Download Partner Files (ZIP)",
                        data=f,
                        file_name=f"{folder_name}.zip",
                        mime="application/zip"
                    )
                
                # Cleanup
                shutil.rmtree(out_dir)
                os.remove(f"{folder_name}.zip")