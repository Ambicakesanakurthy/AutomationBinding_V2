# Import required libraries
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import difflib
 
# Set Streamlit page config
st.set_page_config(page_title="Automatic Binding Tool", layout="centered")
 
# Custom CSS styling
st.markdown("""
<style>
.stApp {
    background-color: #0070AD;
}
.form-box {
    background-color: white;
    padding: 40px 30px;
    border-radius: 15px;
    max-width: 600px;
    margin: auto;
    color: #333;
    box-shadow: 0 4px 12px rgba(0,0,0,0.2);
}
h1 {
    text-align: center;
    color: white;
    margin-bottom: 10px;
}
.sub {
    text-align: center;
    font-size: 16px;
    color: #FF0090;
    margin-bottom: 30px;
}
.stButton>button {
    background-color: #1abc9c;
    color: white;
    font-weight: bold;
    border-radius: 8px;
    padding: 10px 24px;
}
.stButton>button:hover {
    background-color: #16a085;
}
</style>
""", unsafe_allow_html=True)
 
# Add title and description
st.markdown('<div class="form-box"><h1>TGML Binding Tool</h1>', unsafe_allow_html=True)
st.markdown('', unsafe_allow_html=True)
st.markdown('<p class="sub">Upload TGML & Excel File to Update Bindings</p>', unsafe_allow_html=True)
 
# File uploaders
tgml_file = st.file_uploader("TGML File", type=["tgml", "xml"])
excel_file = st.file_uploader("Excel File", type="xlsx")
sheet_name = None
 
# Get sheet names
if excel_file:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        sheet_name = st.selectbox("Select a sheet from the Excel file", sheet_names)
    except Exception as e:
        st.error(f"Error reading Excel: {e}")
 
# Button logic
if st.button("Submit and Download") and tgml_file and excel_file and sheet_name:
    try:
        tree = ET.parse(tgml_file)
        root = tree.getroot()
 
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
        required_columns = ["First Label", "Second Label", "Third Label"]
 
        # Check columns
        for col in required_columns:
            if col not in df.columns:
                st.error(f"Column '{col}' missing in Excel!")
                st.stop()
 
        label_to_bind = {}
        all_labels = set()
        seen_labels = set()
 
        # Count populated labels
        total_labels = 0
        for idx, row in df.iterrows():
            bind = str(row.get("Nomenclature", "")).strip()
            for col in required_columns:
                label = row.get(col)
                if pd.isna(label) or not str(label).strip():
                    continue
                label = str(label).strip()
                key = label.lower()
                if key in seen_labels:
                    st.error(f"Duplicate label '{label}' at row {idx+2}")
                    st.stop()
                seen_labels.add(key)
                label_to_bind[key] = bind
                all_labels.add(key)
                total_labels += 1
 
        # Replacing logic
        replaced_labels = set()
        non_replaced_labels = set(all_labels)
        in_group = False
        current_label_key = None
        inside_target_text = False
 
        for elem in root.iter():
            if elem.tag == "Group":
                in_group = True
            elif elem.tag == "Text" and in_group:
                text_name = elem.attrib.get("Name", "").strip()
                key = text_name.lower()
                matches = difflib.get_close_matches(key, all_labels, n=1, cutoff=0.85)
                if matches:
                    current_label_key = matches[0]
                    inside_target_text = True
                else:
                    current_label_key = None
                    inside_target_text = False
            elif elem.tag == "Bind" and in_group and inside_target_text and current_label_key:
                new_bind = label_to_bind[current_label_key]
                if new_bind:
                    elem.set("Name", new_bind)
                    replaced_labels.add(current_label_key)
                    if current_label_key in non_replaced_labels:
                        non_replaced_labels.remove(current_label_key)
            elif elem.tag == "Text" and inside_target_text:
                inside_target_text = False
                current_label_key = None
 
        # Save updated TGML
        output = BytesIO()
        tree.write(output, encoding="utf-8", xml_declaration=True)
        output.seek(0)
 
        # Show counts
        st.success("✅ Binding completed successfully!")
        st.info(f"🔹 Total unique labels in Excel: {total_labels}")
        st.info(f"✅ Replaced labels: {len(replaced_labels)}")
        st.warning(f"❌ Labels not replaced: {len(non_replaced_labels)}")
 
        # Write non-replaced labels Excel
        if non_replaced_labels:
            out_df = pd.DataFrame(sorted(non_replaced_labels), columns=["Unmatched Labels"])
            out_excel = BytesIO()
            with pd.ExcelWriter(out_excel, engine="xlsxwriter") as writer:
                out_df.to_excel(writer, index=False)
            out_excel.seek(0)
            st.download_button(
                "Download Unmatched Labels Excel",
                out_excel,
                file_name="unmatched_labels.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
 
        st.download_button("Download Updated TGML", output, file_name=f"updated_{tgml_file.name}", mime="application/xml")
 
    except Exception as e:
        st.error(f"Error: {e}")
 
st.markdown('</div>', unsafe_allow_html=True)
 
