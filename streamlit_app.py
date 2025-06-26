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
    /* set the full page background color */
    body {
        background-color: #0070AD;
    }
    /* style the content box */
    main {
         color: pink;
     }

     /* title styling */
    h1 {
        text-align: center;
        color: #114488;
        font-size: 32px;
    }
    p {
       text-align: center;
       margin-bottom: 20px;
       font-size: 14px;
       color: black;
     }
    /* sub title styling*/
    .sub {
        text-align: center;
        color: pink;
        margin-bottom: 30px;
        font-size: 16px;
    }
    /* button styling */
    .stButton>button {
        background-color: #1abc9c;
        color: white;
        font-weight: bold;
        border-radius: 8px;
        padding: 10px 24px;
    }
    /* Hover effect for buttn=on */
    .stButton>button:hover {
        background-color: #16a085;
    }
    .block-container {
        background-color: #0070AD !important;
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
     
        for idx, row in df.iterrows():
            nomenclature = str(row.get("Nomenclature", "")).strip()
            for col in ["First Label", "Second Label", "Third Label"]:
                label = str(row.get(col, "")).strip()
                if label:
                    lower_label = label.lower()
                    if lower_label in seen_labels:
                        st.error(f"Duplicate label found in Excel : '{label}' Row {idx+2}, column '{col}')")
                        st.stop()
                    seen_labels.add(lower_label)
                    label_to_bind[lower_label] = nomenclature
                    all_labels.append(lower_label)
 
        # Replace in TGML
        in_group = False
        current_text = None
        inside_target_text = False
 
        for elem in root.iter():
            if elem.tag == "Group":
                in_group = True   #entering a group block
            elif elem.tag == "Text" and in_group:
                #get text name
                text_name = elem.attrib.get("Name", "").strip()  
                text_name_lower = text_name.lower()
                # perform fuzzy match tp find closest label
                matches = difflib.get_close_matches(text_name_lower, all_labels, n=1, cutoff = 0.85)
                if matches:
                    current_label_key = matches[0]
                    # update to matched label
                    inside_target_text = True
                else:
                    inside_target_text = False
                 
            elif elem.tag == "Bind" and in_group and inside_target_text:
                if current_label_key and current_label_key in label_to_bind:
                     elem.set("Name", label_to_bind[current_label_key])
                 
            elif elem.tag == "Text" and inside_target_text:
                 # Reset after target text block ends
                inside_target_text = False
                current_label_key = None
 
        # Save new file
        output_file = BytesIO()
        tree.write(output_file, encoding="utf-8", xml_declaration=True)
        output.seek(0)
 
        st.download_button("Download Updated TGML", output, file_name=f"updated_{tgml_file.name}", mime = "application/xml")
        st.success("Binding completed successfully!")
 
    except Exception as e:
        # shows error if something goes wrong
        st.error(f"Error: {e}")
 
st.markdown('</div>', unsafe_allow_html=True)
