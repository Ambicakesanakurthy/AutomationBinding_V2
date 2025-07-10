# Import required libraries
import streamlit as st  # For creating web UI
import pandas as pd  # For reading Excel file
import xml.etree.ElementTree as ET  # For parsing and editing TGML
from io import BytesIO
import difflib  # For fuzzy matching of label names
 
# Set title on browser tab and center-align the layout
st.set_page_config(page_title="Automatic Binding Tool", layout="centered")
 
# Add custom CSS for styling the background and form
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
    color: pink;
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
 
# File uploaders and input
tgml_file = st.file_uploader("TGML File", type=["tgml", "xml"])
excel_file = st.file_uploader("Excel File", type="xlsx")
sheet_name = None
 
if excel_file:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_names = xls.sheet_names
        sheet_name = st.selectbox("Select a sheet from the Excel file", sheet_names)
    except Exception as e:
        st.error(f"Error reading Excel sheet names: {e}")
 
if st.button("Submit and Download") and tgml_file and excel_file and sheet_name:
    try:
        # Parse XML
        tree = ET.parse(tgml_file)
        root = tree.getroot()
 
        # Read Excel
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
 
        label_to_bind = {}
        all_labels = []
        seen_labels = set()
 
        required_columns = ["First Label", "Second Label", "Third Label"]
 
        for column in required_columns:
            if column not in df.columns:
                st.error(f"'{column}' column is missing in Excel sheet!")
                st.stop()
 
        # Collect labels from Excel
        for idx, row in df.iterrows():
            bind = str(row.get("Nomenclature", "")).strip()
            for col in required_columns:
                label = row.get(col)
                if pd.isna(label) or not str(label).strip():
                    continue
                label = str(label).strip()
                key = label.lower()
                if key in seen_labels:
                    st.error(f"Duplicate label in Excel: '{label}' Row {idx+2}, column '{col}'")
                    st.stop()
                seen_labels.add(key)
                label_to_bind[key] = bind
                all_labels.append(key)
 
        # Collect all Text labels found in TGML
        tgml_labels = set()
        for elem in root.iter():
            if elem.tag == "Text":
                text_name = elem.attrib.get("Name", "").strip().lower()
                tgml_labels.add(text_name)
 
        # Identify labels from Excel NOT found in TGML
        unmatched_labels = []
        for lbl in all_labels:
            matches = difflib.get_close_matches(lbl, tgml_labels, n=1, cutoff=0.85)
            if not matches:
                unmatched_labels.append(lbl)
 
        # Show warning but continue
        if unmatched_labels:
            st.warning(f"⚠️ These labels were NOT found in the TGML file and were not replaced:\n\n{', '.join(unmatched_labels)}")
 
        # Replace in TGML
        in_group = False
        inside_target_text = False
        current_label_key = None
 
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
            elif elem.tag == "Text" and inside_target_text:
                inside_target_text = False
                current_label_key = None
 
        # Save new file
        output = BytesIO()
        tree.write(output, encoding="utf-8", xml_declaration=True)
        output.seek(0)
 
        st.download_button("Download Updated TGML", output, file_name=f"updated_{tgml_file.name}", mime="application/xml")
        st.success("Binding completed successfully!")
 
    except Exception as e:
        st.error(f"Error: {e}")
 
st.markdown('</div>', unsafe_allow_html=True)
