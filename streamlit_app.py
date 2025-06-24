import streamlit as st
import pandas as pd
import xml.etree.ElementTree as ET
 
# Page configuration
st.set_page_config(page_title="Automatic Binding Tool", layout="centered")
 
# Custom CSS styling
st.markdown("""
    <style>
    body {
        background-color: #2980b9;
    }
    h1 {
        text-align: center;
        color: #2c3e50;
    }
    .sub {
        text-align: center;
        color: #7f8c8d;
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
 
# App header
st.markdown('<div class="main">', unsafe_allow_html=True)
st.markdown('<h1>Automatic Binding Tool</h1>', unsafe_allow_html=True)
st.markdown('<p class="sub">Upload TGML & Excel File to Update Bindings</p>', unsafe_allow_html=True)
 
# File uploaders
tgml_file = st.file_uploader("TGML File", type="tgml")
excel_file = st.file_uploader("Excel File", type="xlsx")
sheet_name = None
 
# Extract sheet names from Excel
if excel_file:
    try:
        xls = pd.ExcelFile(excel_file)
        sheet_name = st.selectbox("Select Sheet Name", xls.sheet_names)
    except Exception as e:
        st.error(f"Error reading Excel file: {e}")
 
# Run logic
if st.button("Submit and Download") and tgml_file and excel_file and sheet_name:
    try:
        # Load XML
        tree = ET.parse(tgml_file)
        root = tree.getroot()
 
        # Load Excel
        df = pd.read_excel(excel_file, sheet_name=sheet_name)
 
        label_to_bind = {}
        seen_labels = {}
 
        # Check for duplicates (case-insensitive)
        for idx, row in df.iterrows():
            nomenclature = str(row.get("Nomenclature", "")).strip()
            for col in ["First Label", "Second Label", "Third Label"]:
                label = str(row.get(col, "")).strip()
                if label:
                    label_key = label.lower()
                    if label_key in seen_labels:
                        prev_row = seen_labels[label_key] + 2
                        curr_row = idx + 2
                        raise ValueError(
                            f"Duplicate label '{label}' found at row {curr_row} (already exists at row {prev_row})."
                        )
                    label_to_bind[label_key] = nomenclature
                    seen_labels[label_key] = idx
 
        # Replace Bind in XML
        in_group = False
        current_text = None
        inside_target_text = False
 
        for elem in root.iter():
            if elem.tag == "Group":
                in_group = True
            elif elem.tag == "Text" and in_group:
                current_text = elem.attrib.get("Name", "").strip()
                inside_target_text = current_text.lower() in label_to_bind
            elif elem.tag == "Bind" and in_group and inside_target_text:
                new_bind = label_to_bind.get(current_text.lower())
                if new_bind:
                    elem.set("Name", new_bind)
            elif elem.tag == "Text" and inside_target_text:
                inside_target_text = False
 
        # Save modified XML
        output_file = "updated_" + tgml_file.name
        tree.write(output_file, encoding="utf-8", xml_declaration=True)
 
        with open(output_file, "rb") as f:
            st.download_button("Download Updated TGML", f, file_name=output_file)
 
        st.success("Binding completed successfully!")
 
    except Exception as e:
        st.error(f"Error: {e}")
 
st.markdown('</div>', unsafe_allow_html=True)
