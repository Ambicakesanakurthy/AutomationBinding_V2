#Import required libraries
import streamlit as st     # For creating web UI
import pandas as pd        # For reading excel file
import xml.etree.ElementTree as ET   # For parsing and editing TGML
from io import BytesIO
import difflib    # For Fuzzy Matching of label names

# Set title on browser tab and center-align the layout
st.set_page_config(page_title="Automatic Binding Tool", layout="centered")
 
# Add custom CSS for styling the background and form
st.markdown("""
    <style>
   body {
        background-color: #0070AD;
    }
 
    .form-box {
        background-color:#0070AD;
        padding: 40px 30px;
        border-radius: 15px;
        max-width: 600px;
        margin: 50px auto;
        color: #0070AD;
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
st.markdown('<div class="main">', unsafe_allow_html=True)
#title of the app
st.markdown('<h1>TGML Binding Tool</h1>', unsafe_allow_html=True)
# sub text with instreuctions
st.markdown('<p class="sub">Upload TGML & Excel File to Update Bindings</p>', unsafe_allow_html=True)
 
# File uploaders and input
tgml_file = st.file_uploader("TGML File", type=["tgml", "xml"])
excel_file = st.file_uploader("Excel File", type="xlsx")
sheet_name = None   #holds seclected sheet name

if excel_file:
    try:
     # Reads sheet names from excel file
     xls = pd.ExcelFile(excel_file)
     # List of sheet names
     sheet_names = xls.sheet_names     # get all sheet names
     # Creates the Dropdown menu
     sheet_name = st.selectbox("select a sheet from the excel file", sheet_names)
    except Exception as e:
     st.error(f"Error in Reading Excel Sheet Names: {e}")
     
 
# Button and logic
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
                st.error(f"'{column}' Column is not available in the Excel sheet, please check!")
                st.stop()
        
        for idx, row in df.iterrows():
            bind = str(row.get("Nomenclature", "")).strip()
            for col in required_columns:
                label = row.get(col)
                if pd.isna(label) or not str(label).strip():
                    continue
                label = str(label).strip()
                key = label.lower()
                if key in seen_labels:
                    st.error(f"Duplicate label found in Excel: '{label}' Row {idx+2}, column '{col}'")
                    st.stop()
                seen_labels.add(key)
                label_to_bind[key] = bind
                all_labels.append(key)
    
 
        # Replace in TGML
        in_group = False
        inside_target_text = False
        current_text_key = None
 
        for elem in root.iter():
            if elem.tag == "Group":
                in_group = True   #entering a group block
            elif elem.tag == "Text" and in_group:
                #get text name
                text_name = elem.attrib.get("Name", "").strip()  
                key = text_name.lower()
                # perform fuzzy match tp find closest label
                matches = difflib.get_close_matches(key, all_labels, n=1, cutoff = 0.85)
             
                if matches:
                    current_label_key = matches[0]
                    # update to matched label
                    inside_target_text = True
                else:
                    current_label_key = None
                    inside_target_text = False
                 
            elif elem.tag == "Bind" and in_group and inside_target_text and current_label_key:
                new_bind = label_to_bind[current_label_key]
                if new_bind:
                     elem.set("Name", new_bind)
                 
            elif elem.tag == "Text" and inside_target_text:
                 # Reset after target text block ends
                inside_target_text = False
                current_label_key = None
 
        # Save new file
        output = BytesIO()
        tree.write(output, encoding="utf-8", xml_declaration=True)
        output.seek(0)
 
        st.download_button("Download Updated TGML", output, file_name=f"updated_{tgml_file.name}", mime = "application/xml")
        st.success("Binding completed successfully!")
 
    except Exception as e:
        # shows error if something goes wrong
        st.error(f"Error: {e}")
 
st.markdown('</div>', unsafe_allow_html=True)
