import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
import shutil
import os
import io

# 1. Page Configuration
st.set_page_config(page_title="Noon Price Comp", layout="centered")
st.title("📊 Partner Price Comp Tool")

# 2. Initialize Session State for Data Persistence
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

# 3. File Upload (Supports CSV & Excel)
uploaded_file = st.file_uploader("Upload Master File (Excel or CSV)", type=["xlsx", "csv"])

if uploaded_file:
    if st.button("🚀 Process & Generate ZIP"):
        with st.spinner('Processing files...'):
            try:
                # Detect and read file type
                if uploaded_file.name.endswith('.csv'):
                    df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
                else:
                    df = pd.read_excel(uploaded_file, engine='openpyxl')
                    
                df.columns = [c.strip() for c in df.columns]
                
                # Filter for NC/NCO
                mask = df["Price Comp Bucket"].str.upper().isin({"NC", "NCO"})
                filtered = df.loc[mask].copy()
                
                if filtered.empty:
                    st.error("No 'NC' or 'NCO' rows found in the data.")
                else:
                    # Calculations
                    filtered["Latest Comp Price All"] = pd.to_numeric(filtered["Latest Comp Price All"], errors='coerce')
                    filtered["Offer Price"] = pd.to_numeric(filtered["Offer Price"], errors='coerce')
                    filtered["Adjustment needed"] = filtered["Offer Price"] - filtered["Latest Comp Price All"]
                    
                    # Generate Hyperlinks
                    sku_cfg = filtered["SKU Config"].astype(str)
                    filtered["noon link"] = '=HYPERLINK("http://noon.com/egypt-en/' + sku_cfg + '/p/", "View on Noon")'
                    filtered["Comp Link"] = '=HYPERLINK("' + filtered["Comp Link"].astype(str) + '", "View Competitor")'

                    # Setup Temporary Directory
                    folder_name = f"Comp_{datetime.now():%d-%m}"
                    out_dir = Path(folder_name)
                    if out_dir.exists(): 
                        shutil.rmtree(out_dir)
                    out_dir.mkdir()

                    # 4. Create Master Summary (Pivot-style Excel)
                    st.session_state.summary_df = (
                        filtered.groupby(["ID Partner", "Partner Name"])["SKU"]
                        .nunique()
                        .reset_index(name="NC_NCO_SKU_Count")
                        .sort_values(by="NC_NCO_SKU_Count", ascending=False)
                    )
                    st.session_state.summary_df.to_excel(out_dir / "Partner_PriceComp_Counts.xlsx", index=False)

                    # 5. Partner Export Loop
                    grouped = filtered.groupby("ID Partner", sort=False)
                    st.session_state.partner_count = len(grouped)
                    
                    for pid, grp in grouped:
                        p_name = grp["Partner Name"].iloc[0] if "Partner Name" in grp else f"Partner_{pid}"
                        safe_name = "".join(c for c in str(p_name) if c.isalnum() or c in " -_").strip()[:50]
                        fname = out_dir / f"{safe_name}_PriceComp.xlsx"
                        
                        with pd.ExcelWriter(fname, engine='openpyxl') as writer:
                            cols_to_write = [c for c in SELECTED_COLS if c in grp.columns]
                            grp.to_excel(writer, index=False, columns=cols_to_write, sheet_name='PriceComp')
                            
                            # Individual partner summary tab
                            partner_count = grp['SKU'].nunique()
                            pd.DataFrame({
                                "Partner ID": [pid],
                                "Partner Name": [p_name],
                                "NC_NCO_SKU_Count": [partner_count]
                            }).to_excel(writer, index=False, sheet_name='Summary')

                    # 6. Package into ZIP
                    zip_path = shutil.make_archive(folder_name, 'zip', out_dir)
                    with open(zip_path, "rb") as f:
                        st.session_state.zip_buffer = f.read()
                    st.session_state.zip_name = f"{folder_name}.zip"

                    # Final cleanup
                    shutil.rmtree(out_dir)
                    if os.path.exists(zip_path):
                        os.remove(zip_path)

            except Exception as e:
                st.error(f"Error processing file: {str(e)}")

# --- 7. PERSISTENT DISPLAY (Outside the button logic) ---
if st.session_state.summary_df is not None:
    # Display Chart
    st.subheader("🔝 Top 10 Sellers by NC/NCO Count")
    st.bar_chart(data=st.session_state.summary_df.head(10), x="Partner Name", y="NC_NCO_SKU_Count")
    
    # Persistent Download Button
    st.success(f"✅ Generated files for {st.session_state.partner_count} partners.")
    st.download_button(
        label="⬇️ Download ZIP",
        data=st.session_state.zip_buffer,
        file_name=st.session_state.zip_name,
        mime="application/zip"
    )
