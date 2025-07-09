# Import required libraries
import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
from io import BytesIO
import difflib
 
# Set page configuration
st.set_page_config(page_title="Automatic Binding Tool", layout="centered")
 
# Add custom CSS styling
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
    box-shadow: 0 4px 12px rgba(0, 0, 0, 0.2);
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
st.markdown('<div class="form-box">', unsafe_allow_html=True)
st.markdown('<h1>TGML Binding Tool</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub">Upload TGML & Excel File to Update Bindings</p>', unsafe_allow_html=True)
 
# File uploaders
tgml_file = st.file_uploader("TGML File", type=["tgml", "xml"])
excel_file = st.file_uploader("Excel File", type="xlsx")
sheet_name = None
 
# Read sheet names if Excel is uploaded
if excel_file:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        sheet_name = st.selectbox("Select a sheet from the Excel file", sheet_names)
    except Exception as e:
        st.error(f"Error in Reading Excel Sheet Names: {e}")
 
# Main processing logic
if st.button("Submit and Download") and tgml_file and excel_file and sheet_name:
    try:
        # Parse XML
        tree = ET.parse(tgml_file)
        root = tree.getroot()
 
        # Read Excel
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
 
        # Initialize label mapping
        label_to_bind = {}
        excel_labels = set()   # labels present in Excel
        used_labels = set()    # labels matched and used in TGML
 
        required_columns = ["First Label", "Second Label", "Third Label"]
 
        # Validate columns
        for column in required_columns:
            if column not in df.columns:
                st.error(f"'{column}' column is not available in the Excel sheet. Please check!")
                st.stop()
 
        # Collect labels only from Excel
        for idx, row in df.iterrows():
            bind = str(row.get("Nomenclature", "")).strip()
            for col in required_columns:
                label = row.get(col)
                if pd.isna(label) or not str(label).strip():
                    continue
                label_str = str(label).strip()
                label_key = label_str.lower()
 
                if label_key in excel_labels:
                    st.error(f"Duplicate label found in Excel: '{label_str}' at row {idx+2}, column '{col}'")
                    st.stop()
 
                excel_labels.add(label_key)
                label_to_bind[label_key] = bind
 
        # Replace in TGML
        in_group = False
        inside_target_text = False
        current_label_key = None
 
        for elem in root.iter():
            if elem.tag == "Group":
                in_group = True
            elif elem.tag == "Text" and in_group:
                text_name = elem.attrib.get("Name", "").strip()
                text_key = text_name.lower()
                # Fuzzy match ONLY with Excel labels
                matches = difflib.get_close_matches(text_key, excel_labels, n=1, cutoff=0.85)
                if matches:
                    current_label_key = matches[0]
                    inside_target_text = True
                    used_labels.add(current_label_key)
                else:
                    current_label_key = None
                    inside_target_text = False
            elif elem.tag == "Bind" and in_group and inside_target_text and current_label_key:
                new_bind = label_to_bind.get(current_label_key)
                if new_bind:
                    elem.set("Name", new_bind)
            elif elem.tag == "Text" and inside_target_text:
                inside_target_text = False
                current_label_key = None
 
        # Save updated TGML
        output = BytesIO()
        tree.write(output, encoding="utf-8", xml_declaration=True)
output.seek(0)
 
        # Compute unmatched labels
        not_replaced_labels = sorted(excel_labels - used_labels)
 
        # Show counts
        st.success("✅ Binding completed successfully!")
st.info(f"📝 Total Labels in Excel: {len(excel_labels)}")
st.info(f"✅ Replaced Labels: {len(used_labels)}")
        st.warning(f"❌ Not Replaced Labels: {len(not_replaced_labels)}")
 
        # Download updated TGML
st.download_button("Download Updated TGML", output, file_name=f"updated_{tgml_file.name}", mime="application/xml")
 
        # Save unmatched labels to Excel
        if not_replaced_labels:
            unmatched_df = pd.DataFrame(not_replaced_labels, columns=["Unmatched Labels"])
            unmatched_output = BytesIO()
            with pd.ExcelWriter(unmatched_output, engine="xlsxwriter") as writer:
                unmatched_df.to_excel(writer, index=False)
unmatched_output.seek(0)
            st.download_button("Download Unmatched Labels", unmatched_output, file_name="unmatched_labels.xlsx")
 
    except Exception as e:
        st.error(f"Error: {e}")
 
st.markdown('</div>', unsafe_allow_html=True)
