import streamlit as st
import pandas as pd
import os
import tempfile
import zipfile
import io
import base64
from datetime import datetime
import re
import requests
from pypdf import PdfReader, PdfWriter
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# Set page configuration
st.set_page_config(
    page_title="EPI Generator",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="expanded",
)

# Custom styling
st.markdown("""
<style>
    .main-header {
        font-size: 2em;
        font-weight: bold;
        color: #005670;
        margin-bottom: 0.5em;
    }
    .subheader {
        font-size: 1.5em;
        font-weight: bold;
        color: #3F4443;
        margin-bottom: 0.5em;
    }
    .stButton>button {
        background-color: #3F4443;
        color: #CDD325;
        border: none;
        padding: 0.5em 1em;
        font-weight: bold;
    }
    .stButton>button:hover {
        background-color: #2D3231;
        color: #CDD325;
    }
    .download-btn {
        display: inline-block;
        padding: 0.5em 1em;
        color: white;
        background-color: #4CAF50;
        text-decoration: none;
        font-weight: bold;
        border-radius: 4px;
        margin-top: 1em;
    }
    .download-btn:hover {
        background-color: #45a049;
    }
</style>
""", unsafe_allow_html=True)

# Function to create a download link for the zip file
def get_download_link(buffer, filename):
    b64 = base64.b64encode(buffer.getvalue()).decode()
    return f'<a href="data:application/zip;base64,{b64}" download="{filename}" class="download-btn">Download {filename}</a>'

# Load Arial font if available, otherwise use a default font
try:
    pdfmetrics.registerFont(TTFont('Arial', 'arial.ttf'))
    font_name = 'Arial'
except:
    # Use a default font that's available in the environment
    font_name = 'Helvetica'

def fill_epi_template(input_pdf_path, output_pdf_path, field_data_by_page):
    """
    Fill a PDF template with data
    """
    try:
        reader = PdfReader(input_pdf_path)
        writer = PdfWriter()

        for page_num, fields in field_data_by_page.items():
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=letter)
            
            # Use the font
            can.setFont(font_name, 7.5)

            for field, (text, x, y) in fields.items():
                can.drawString(x, y, str(text))

            can.save()
            packet.seek(0)
            new_pdf = PdfReader(packet)

            page = reader.pages[page_num]
            page.merge_page(new_pdf.pages[0])
            writer.add_page(page)

        with open(output_pdf_path, "wb") as output_file:
            writer.write(output_file)
        return True
    except Exception as e:
        st.error(f"Error creating PDF: {e}")
        return False

def match_lenders_to_sov(sheets_df, df_lenders):
    """
    Matches lenders to each row in the SOV based on Location Name.
    """
    if "Location Name" not in sheets_df.columns or "Location Name" not in df_lenders.columns:
        st.error("Missing 'Location Name' column in either SOV or Lenders sheet.")
        return None

    # Strip extra spaces and normalize column values
    sheets_df["Location Name"] = sheets_df["Location Name"].str.strip()
    df_lenders["Location Name"] = df_lenders["Location Name"].str.strip()

    # Merge based on Location Name
    matched_df = sheets_df.merge(df_lenders, on="Location Name", how="left", suffixes=("", "_lender"))
    return matched_df

def get_value(df, row, col, default=""):
    """Safely get a value from a DataFrame"""
    try:
        if row < len(df) and col < len(df.columns):
            value = df.iloc[row, col]
            return value if pd.notna(value) else default
        return default
    except Exception:
        return default

def format_money(value, default=""):
    """Format a number as currency"""
    try:
        if str(value).strip().lower() == "excluded":
            return "Excluded"
        elif str(value).strip().lower() == "included":
            return "Included"
        elif pd.notna(value) and value != "":
            return "{:,}".format(int(float(value)))
        else:
            return default
    except (ValueError, TypeError):
        return str(value) if pd.notna(value) else default

def process_sov(file, template_path, producer):
    """Process the SOV file and generate EPI PDFs"""
    
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            # Check if the template file exists
            if not os.path.exists(template_path):
                st.error(f"Template file not found: {template_path}")
                return None
            
            # Read the Excel file
            required_sheets = ["SOV", "Property Coverage Information", "Additional Remarks Page 1", "Additional Remarks Page 2", "Lenders"]
            
            try:
                excel_data = pd.ExcelFile(file)
                
                # Check for required sheets
                missing_sheets = [sheet for sheet in required_sheets if sheet not in excel_data.sheet_names]
                if missing_sheets:
                    st.error(f"Missing required sheets: {', '.join(missing_sheets)}")
                    return None
                
                # Parse data from sheets
                data = {
                    sheet: excel_data.parse(sheet, header=11).dropna(how="all") if sheet == "SOV" else excel_data.parse(sheet).dropna(how="all")
                    for sheet in required_sheets
                }
            except Exception as e:
                st.error(f"Error reading Excel file: {e}")
                return None
            
            sheets_df = data["SOV"]
            df_lenders = data["Lenders"]
            df_carriers = data["Additional Remarks Page 1"]
            df_carriers2 = data["Additional Remarks Page 2"]
            
            # Normalize column names
            sheets_df.columns = sheets_df.columns.str.strip()
            df_lenders.columns = df_lenders.columns.str.strip()
            
            # Check for missing columns
            if "Location Name" not in sheets_df.columns:
                st.error("The 'Location Name' column is missing in the SOV sheet. Please check your input file.")
                return None
            
            if "Location Name" not in df_lenders.columns:
                st.error("The 'Location Name' column is missing in the Lenders sheet. Please check your input file.")
                return None
            
            # Match lenders to SOV
            matched_sov = match_lenders_to_sov(sheets_df, df_lenders)
            if matched_sov is None:
                return None
            
            # Output directory for PDFs
            output_dir = os.path.join(temp_dir, "output")
            os.makedirs(output_dir, exist_ok=True)
            
            # Create a zip file to store all generated PDFs
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w") as zip_file:
                # Create progress bar
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                # Process each row
                successful_pdfs = 0
                total_rows = len(matched_sov)
                
                for index, row in matched_sov.iterrows():
                    try:
                        status_text.text(f"Processing {index + 1} of {total_rows}...")
                        
                        # Extract data from row
                        entity_name = row.get("Entity Name", "")
                        certnum = row.get('Cert Number', "")
                        location_name = row.get("Location Name", "")
                        street_address = row.get("*Street Address", "")
                        
                        # Skip if mandatory fields are missing
                        if pd.isna(location_name) or pd.isna(street_address):
                            continue
                        
                        # Extract building and business income values
                        buildingvalue = row.get('*Real Property Value ($)', 0)
                        businessincomesov = row.get('BI/Rental Income ($)', 0)
                        
                        # Format address
                        try:
                            town_address = f"{row.get('*City', '')}, {row.get('*State Code', '')}, {str(int(float(row.get('*Zip', 0)))).zfill(5)}"
                        except:
                            town_address = f"{row.get('*City', '')}, {row.get('*State Code', '')}, {row.get('*Zip', '')}"
                        
                        # Lender information
                        lender_name = row.get("Lender Name", "N/A")
                        lender_address =  "" if pd.isna(row.get('Street Address', '')) else row.get('Street Address', '')
                       
                        lender_address_additional = "" if pd.isna(row.get('Street Address 2', '')) else row.get('Street Address 2', '')
                        
                        try:
                            city = "" if pd.isna(row.get('City', '')) else row.get('City', '')
                            state = "" if pd.isna(row.get('State', '')) else row.get('State', '')
                            zipcode = row.get('Zipcode', 0)
                            
                            if pd.isna(zipcode):
                                zipcode_str = ""
                            else:
                                zipcode_str = str(int(float(zipcode))).zfill(5)
                            
                            lender_address2 = f"{city}, {state} {zipcode_str}".strip()
                        except:
                            city = "" if pd.isna(row.get('City', '')) else row.get('City', '')
                            state = "" if pd.isna(row.get('State', '')) else row.get('State', '')
                            zipcode = "" if pd.isna(row.get('Zipcode', '')) else row.get('Zipcode', '')
                            lender_address2 = f"{city}, {state} {zipcode}".strip()
                                                
                        # Lender checkboxes
                        lender_cos = row.get("Contract of Sale", 0)
                        lender_llp = row.get("Lenders Loss Payable", 0)
                        lender_lp = row.get("Loss Payee", 0)
                        lender_m = row.get("Mortgagee", 0)
                        other_lender = "" if pd.isna(row.get('Other (Type in below)', '')) else row.get('Other (Type in below)', '')
                        
                        # Lender additional information
                        lender_info = []
                        for i in range(1, 13):
                            field_name = f'Lender Additional Information {i}'
                            value = row.get(field_name, "")
                            lender_info.append("" if pd.isna(value) else str(value))
                        
                        # Date information
                        current_date = datetime.now().strftime("%m/%d/%Y")
                        current_date_year = int(datetime.now().strftime("%Y")) - 2000
                        next_year = current_date_year + 1
                        
                        # Get coverage information
                        coverage_info = data["Property Coverage Information"]
                        
                        # Format dates
                        try:
                            effective_date_obj = coverage_info.iloc[0, 1]
                            expiration_date_obj = coverage_info.iloc[1, 1]
                            effective_date = effective_date_obj.strftime("%m/%d/%Y") if pd.notna(effective_date_obj) else ""
                            expiration_date = expiration_date_obj.strftime("%m/%d/%Y") if pd.notna(expiration_date_obj) else ""
                        except:
                            effective_date = ""
                            expiration_date = ""
                        
                        # Client information
                        csa_name = str(get_value(coverage_info, 3, 1, ""))
                        client_name = str(get_value(coverage_info, 4, 1, ""))
                        client_line_address = str(get_value(coverage_info, 5, 1, ""))
                        client_2nd_line_address = f"{get_value(coverage_info, 6, 1, '')}, {get_value(coverage_info, 7, 1, '')} {get_value(coverage_info, 8, 1, '')}"
                        
                        # Additional information
                        value = coverage_info.iloc[3, 3] if 3 < len(coverage_info.index) and 3 < len(coverage_info.columns) else ""
                        csaemail = "" if pd.isnull(value) else str(value).strip()
                        csanumb = "(813) 839-7330"
                        
                        # Agency customer ID
                        try:
                            value = coverage_info.iloc[5, 2] if 5 < len(coverage_info.index) and 2 < len(coverage_info.columns) else ""
                            if str(value).strip().lower() == "excluded":
                                agencycustomerid = "Excluded"
                            elif str(value).strip().lower() == "included":
                                agencycustomerid = "Included"
                            elif pd.notnull(value):
                                agencycustomerid = str(int(float(value)))
                            else:
                                agencycustomerid = ""
                        except (ValueError, TypeError):
                            agencycustomerid = str(value).strip() if pd.notnull(value) else ""
                        
                        # Loan numbers
                        try:
                            value = coverage_info.iloc[7, 2] if 7 < len(coverage_info.index) and 2 < len(coverage_info.columns) else ""
                            if str(value).strip().lower() == "excluded":
                                loannumbers = "Excluded"
                            elif str(value).strip().lower() == "included":
                                loannumbers = "Included"
                            elif pd.notnull(value):
                                loannumbers = str(int(float(value)))
                            else:
                                loannumbers = ""
                        except (ValueError, TypeError):
                            loannumbers = str(value).strip() if pd.notnull(value) else ""
                        
                        # Lender information
                        lender_cos_check = "X" if lender_cos == 1 else ""
                        lender_llp_check = "X" if lender_llp == 1 else ""
                        lender_lp_check = "X" if lender_lp == 1 else ""
                        lender_m_check = "X" if lender_m == 1 else ""
                        
                        if other_lender == "":
                            other_lender_check = ""
                            other_lender_text = ""
                        else:
                            other_lender_check = "X"
                            other_lender_text = other_lender
                        
                        # Building coverage
                        a = 13
                        buildingcheck = "X" if get_value(coverage_info, a-2, 1, 0) == 1 else ""
                        buildingpersonalcheck = "X" if get_value(coverage_info, a-2, 3, 0) == 1 else ""
                        
                        # Coverage types
                        basic = "X" if get_value(coverage_info, a+1, 1, 0) == 1 else ""
                        broad = "X" if get_value(coverage_info, a+2, 1, 0) == 1 else ""
                        special = "X" if get_value(coverage_info, a+3, 1, 0) == 1 else ""
                        other = "X" if get_value(coverage_info, a+4, 1, 0) == 1 else ""
                        
                        # Extract various coverage details and limits
                        # For brevity, I'm showing a few examples - the real implementation would include all the fields
                        
                        # Commercial Property Coverage Amount
                        try:
                            value = get_value(coverage_info, a+5, 1, "")
                            if str(value).strip().lower() == "excluded":
                                coverageamount = "Excluded"
                            elif str(value).strip().lower() == "included":
                                coverageamount = "Included"
                            elif pd.notnull(value) and value != "":
                                coverageamount = "${:,}".format(int(float(value)))
                            else:
                                coverageamount = ""
                        except (ValueError, TypeError):
                            coverageamount = "0"

                        # Deductible
                        try:
                            value = get_value(coverage_info, a+6, 1, "")
                            val_str = str(value).strip().lower()
                            if val_str == "excluded":
                                deductible = "Excluded"
                            elif val_str == "included":
                                deductible = "Included"
                            elif val_str == "see attached":
                                deductible = "See Attached"
                            elif pd.notnull(value) and value != "":
                                deductible = "${:,}".format(int(float(value)))
                            else:
                                deductible = ""
                        except (ValueError, TypeError):
                            deductible = ""




                
                        
                        # Business Income
                        Businessincome = "X" if get_value(coverage_info, a+8, 1, 0) == 1 else ""
                        Businessincomeno = "X" if get_value(coverage_info, a+9, 1, 0) == 1 else ""
                        Businessincomena = "X" if get_value(coverage_info, a+10, 1, 0) == 1 else ""
                        Bicheck = "X" if get_value(coverage_info, a+12, 1, 0) == 1 else ""
                        Rentalvalue = "X" if get_value(coverage_info, a+14, 1, 0) == 1 else ""
                        
                        # BI Limit
                        try:
                            value = get_value(coverage_info, a+15, 1, "")
                            if str(value).strip().lower() == "excluded":
                                bilim = "Excluded"
                            elif str(value).strip().lower() == "included":
                                bilim = "Included"
                            elif pd.notnull(value) and value != "":
                                bilim = "${:,}".format(int(float(value)))
                            else:
                                bilim = ""
                        except (ValueError, TypeError):
                            bilim = "0"
                        
                        # Months
                        try:
                            value = get_value(coverage_info, a+16, 1, "")
                            if str(value).strip().lower() == "excluded":
                                months = "Excluded"
                            elif str(value).strip().lower() == "included":
                                months = "Included"
                            elif pd.notnull(value) and value != "":
                                months = "{:,}".format(int(float(value)))
                            else:
                                months = ""
                        except (ValueError, TypeError):
                            months = "0"
                        
                        losscheck = "X" if get_value(coverage_info, a+16, 1, "") != "" else ""
                        
                        # Continue with all other fields...
                        # For brevity, we'll skip ahead to just a few more examples
                        
                        # Blanket Coverage
                        blanket_yes = "X" if get_value(coverage_info, a+19, 1, 0) == 1 else ""
                        blanket_no = "X" if get_value(coverage_info, a+20, 1, 0) == 1 else ""
                        blanket_na = "X" if get_value(coverage_info, a+21, 1, 0) == 1 else ""
                        
                        try:
                            value = get_value(coverage_info, a+18, 1, "")
                            if str(value).strip().lower() == "excluded":
                                blanket_lim = "Excluded"
                            elif str(value).strip().lower() == "included":
                                blanket_lim = "Included"
                            elif pd.notnull(value) and value != "":
                                blanket_lim = "{:,}".format(int(float(value)))
                            else:
                                blanket_lim = ""
                        except (ValueError, TypeError):
                            blanket_lim = "0"
                        
                        # Terrorism
                        terror_yes = "X" if get_value(coverage_info, a+23, 1, 0) == 1 else ""
                        terror_no = "X" if get_value(coverage_info, a+24, 1, 0) == 1 else ""
                        terror_na = "X" if get_value(coverage_info, a+25, 1, 0) == 1 else ""
                        
                        # Other checkbox fields
                        terror_exclusions_yes = "X" if get_value(coverage_info, a+27, 1, 0) == 1 else ""
                        terror_exclusions_no = "X" if get_value(coverage_info, a+28, 1, 0) == 1 else ""
                        terror_exclusions_na = "X" if get_value(coverage_info, a+29, 1, 0) == 1 else ""
                        
                        domestic_terror_exclusions_yes = "X" if get_value(coverage_info, a+31, 1, 0) == 1 else ""
                        domestic_terror_exclusions_no = "X" if get_value(coverage_info, a+32, 1, 0) == 1 else ""
                        domestic_terror_exclusions_na = "X" if get_value(coverage_info, a+33, 1, 0) == 1 else ""
                        
                        # Fungus
                        limited_fungus_coverage_yes = "X" if get_value(coverage_info, a+35, 1, 0) == 1 else ""
                        limited_fungus_coverage_no = "X" if get_value(coverage_info, a+36, 1, 0) == 1 else ""
                        limited_fungus_coverage_na = "X" if get_value(coverage_info, a+37, 1, 0) == 1 else ""
                        
                        # Process the rest of the checkbox fields following the same pattern...
                                    
                        value = coverage_info.iloc[a+38, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                fungus_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() in ["excluded", "included"]:
                                    fungus_lim = value_str.capitalize()
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    fungus_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    fungus_lim = value_str  # Preserve any text as-is
                        
                        except (ValueError, TypeError):
                            fungus_lim = "0"
            
            
                        value = coverage_info.iloc[a+39, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                fungus_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    fungus_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    fungus_ded = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    fungus_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    fungus_ded = value_str  # Preserve any text as-is
                        
                        except (ValueError, TypeError):
                            fungus_ded = "0"
            
            
            
                        # Fungus Exclusion
                        value = coverage_info.iloc[a+41, 1]  # Starting at 41
                        if value == 1:
                            fungus_exclusion_yes = "X"
                        else:
                            fungus_exclusion_yes = ""
                        
                        value = coverage_info.iloc[a+42, 1]  # Incremented to 42
                        if value == 1:
                            fungus_exclusion_no = "X"
                        else:
                            fungus_exclusion_no = ""
                        
                        value = coverage_info.iloc[a+43, 1]  # Incremented to 43
                        if value == 1:
                            fungus_exclusion_na = "X"
                        else:
                            fungus_exclusion_na = ""
            
            
                        # Replacement Cost
                        value = coverage_info.iloc[a+46, 1]  # Starting at 45
                        if value == 1:
                            replacement_cost_yes = "X"
                        else:
                            replacement_cost_yes = ""
                        
                        value = coverage_info.iloc[a+47, 1]  # Incremented to 46
                        if value == 1:
                            replacement_cost_no = "X"
                        else:
                            replacement_cost_no = ""
                        
                        value = coverage_info.iloc[a+48, 1]  # Incremented to 47
                        if value == 1:
                            replacement_cost_na = "X"
                        else:
                            replacement_cost_na = ""
            
                        # Agreed Value
                        value = coverage_info.iloc[a+50, 1]  # Starting at 49
                        if value == 1:
                            agreed_value_yes = "X"
                        else:
                            agreed_value_yes = ""
                        
                        value = coverage_info.iloc[a+51, 1]  # Incremented to 50
                        if value == 1:
                            agreed_value_no = "X"
                        else:
                            agreed_value_no = ""
                        
                        value = coverage_info.iloc[a+52, 1]  # Incremented to 51
                        if value == 1:
                            agreed_value_na = "X"
                        else:
                            agreed_value_na = ""
            
                        # Co-Insurance
                        value = coverage_info.iloc[a+54, 1]  # Starting at 54
                        if value == 1:
                            co_insurance_yes = "X"
                        else:
                            co_insurance_yes = ""
                        
                        value = coverage_info.iloc[a+55, 1]  # Incremented to 55
                        if value == 1:
                            co_insurance_no = "X"
                        else:
                            co_insurance_no = ""
                        
                        value = coverage_info.iloc[a+56, 1]  # Incremented to 56
                        if value == 1:
                            co_insurance_na = "X"
                        else:
                            co_insurance_na = ""
            
                        value = coverage_info.iloc[a+57, 1]
                        
                        try:
                            if str(value).strip().lower() == "excluded":
                                co_ins_percent = "Excluded"
                            elif str(value).strip().lower() == "included":
                                co_ins_percent = "Included"
                            elif pd.notnull(value):
                                co_ins_percent = "{:.0f}".format(float(value) * 100)  # Multiply by 100 and format as a percentage
                            else:
                                co_ins_percent = ""
                        except (ValueError, TypeError):
                            co_ins_percent = "0"
            
                        # Equipment Breakdown
                        value = coverage_info.iloc[a+59, 1]  # Starting at 59
                        if value == 1:
                            equipment_breakdown_yes = "X"
                        else:
                            equipment_breakdown_yes = ""
                        
                        value = coverage_info.iloc[a+60, 1]  # Incremented to 60
                        if value == 1:
                            equipment_breakdown_no = "X"
                        else:
                            equipment_breakdown_no = ""
                        
                        value = coverage_info.iloc[a+61, 1]  # Incremented to 61
                        if value == 1:
                            equipment_breakdown_na = "X"
                        else:
                            equipment_breakdown_na = ""
                        
                        value = coverage_info.iloc[a+62, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                equipment_breakdown_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    equipment_breakdown_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    equipment_breakdown_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    equipment_breakdown_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    equipment_breakdown_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            equipment_breakdown_lim = "0"
            
            
                        value = coverage_info.iloc[a+63, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                equipment_breakdown_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    equipment_breakdown_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    equipment_breakdown_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    equipment_breakdown_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    equipment_breakdown_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    equipment_breakdown_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            equipment_breakdown_ded = "0"
            
            
            
            #Ordinance or Law
                        # Coverage A
                        value = coverage_info.iloc[a+66, 1]  # Starting at 66
                        if value == 1:
                            coverage_a_yes = "X"
                        else:
                            coverage_a_yes = ""
                        
                        value = coverage_info.iloc[a+67, 1]  # Incremented to 67
                        if value == 1:
                            coverage_a_no = "X"
                        else:
                            coverage_a_no = ""
                        
                        value = coverage_info.iloc[a+68, 1]  # Incremented to 68
                        if value == 1:
                            coverage_a_na = "X"
                        else:
                            coverage_a_na = ""
                        
                        value = coverage_info.iloc[a+69, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                coverage_a_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    coverage_a_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    coverage_a_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    coverage_a_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    coverage_a_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            coverage_a_lim = "0"
                                    
                        value = coverage_info.iloc[a+70, 1]
                        
                        try:
                            if str(value).strip().lower() == "excluded":
                                coverage_a_ded = "Excluded"
                            elif str(value).strip().lower() == "included":
                                coverage_a_ded = "Included"
                            elif str(value).strip().lower() == "see attached":
                                coverage_a_ded = "See Attached"  
                            elif pd.notnull(value):
                                coverage_a_ded = "${:,}".format(int(float(value)))
                            else:
                                coverage_a_ded = ""
                        except (ValueError, TypeError):
                            coverage_a_ded = "0"
            
            
                        # Coverage B
                        value = coverage_info.iloc[a+72, 1]  # Starting at 72
                        if value == 1:
                            coverage_b_yes = "X"
                        else:
                            coverage_b_yes = ""
                        
                        value = coverage_info.iloc[a+73, 1]  # Incremented to 73
                        if value == 1:
                            coverage_b_no = "X"
                        else:
                            coverage_b_no = ""
                        
                        value = coverage_info.iloc[a+74, 1]  # Incremented to 74
                        if value == 1:
                            coverage_b_na = "X"
                        else:
                            coverage_b_na = ""
                        
                        value = coverage_info.iloc[a+75, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                coverage_b_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    coverage_b_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    coverage_b_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    coverage_b_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    coverage_b_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            coverage_b_lim = "0"
                        
                        value = coverage_info.iloc[a+76, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                coverage_b_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    coverage_b_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    coverage_b_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    coverage_b_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    coverage_b_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    coverage_b_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            coverage_b_ded = "0"
            
                
                        # Coverage C
                        value = coverage_info.iloc[a+78, 1]  # Starting at 78
                        if value == 1:
                            coverage_c_yes = "X"
                        else:
                            coverage_c_yes = ""
                        
                        value = coverage_info.iloc[a+79, 1]  # Incremented to 79
                        if value == 1:
                            coverage_c_no = "X"
                        else:
                            coverage_c_no = ""
                        
                        value = coverage_info.iloc[a+80, 1]  # Incremented to 80
                        if value == 1:
                            coverage_c_na = "X"
                        else:
                            coverage_c_na = ""
                        
                        value = coverage_info.iloc[a+81, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                coverage_c_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    coverage_c_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    coverage_c_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    coverage_c_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    coverage_c_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            coverage_c_lim = "0"
                        
                        value = coverage_info.iloc[a+82, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                coverage_c_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    coverage_c_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    coverage_c_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    coverage_c_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    coverage_c_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    coverage_c_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            coverage_c_ded = "0"
            
                        # Earth Movement
                        value = coverage_info.iloc[a+84, 1]  # Starting at 84
                        if value == 1:
                            earth_movement_yes = "X"
                        else:
                            earth_movement_yes = ""
                        
                        value = coverage_info.iloc[a+85, 1]  # Incremented to 85
                        if value == 1:
                            earth_movement_no = "X"
                        else:
                            earth_movement_no = ""
                        
                        value = coverage_info.iloc[a+86, 1]  # Incremented to 86
                        if value == 1:
                            earth_movement_na = "X"
                        else:
                            earth_movement_na = ""
                        
                        value = coverage_info.iloc[a+87, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                earth_movement_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    earth_movement_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    earth_movement_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    earth_movement_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    earth_movement_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            earth_movement_lim = "0"
                        
                        value = coverage_info.iloc[a+88, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                earth_movement_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    earth_movement_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    earth_movement_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    earth_movement_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    earth_movement_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    earth_movement_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            earth_movement_ded = "0"
            
            
                        # Flood
                        value = coverage_info.iloc[a+90, 1]  # Starting at 90
                        if value == 1:
                            flood_yes = "X"
                        else:
                            flood_yes = ""
                        
                        value = coverage_info.iloc[a+91, 1]  # Incremented to 91
                        if value == 1:
                            flood_no = "X"
                        else:
                            flood_no = ""
                        
                        value = coverage_info.iloc[a+92, 1]  # Incremented to 92
                        if value == 1:
                            flood_na = "X"
                        else:
                            flood_na = ""
                        
                        value = coverage_info.iloc[a+93, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                flood_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    flood_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    flood_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    flood_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    flood_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            flood_lim = "0"
                        
                        value = coverage_info.iloc[a+94, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                flood_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    flood_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    flood_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    flood_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    flood_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    flood_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            flood_ded = "0"
                                    # Here we would continue processing all the remaining fields from the original code
                                    # For brevity, we're showing a streamlined version
                                    
                                    # Many more fields would be processed here...
                        
                        # Wind/Hail fields
                        wind_hail_yes = "X" if get_value(coverage_info, a+96, 1, 0) == 1 else ""
                        wind_hail_no = "X" if get_value(coverage_info, a+97, 1, 0) == 1 else ""
                        wind_hail_na = "X" if get_value(coverage_info, a+98, 1, 0) == 1 else ""
                        wind_hail_yes2 = "X" if get_value(coverage_info, a+101, 1, 0) == 1 else ""
                        wind_hail_no2 = "X" if get_value(coverage_info, a+102, 1, 0) == 1 else ""
                        
                        value = coverage_info.iloc[a+99, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                wind_hail_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    wind_hail_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    wind_hail_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    wind_hail_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    wind_hail_lim = value_str  # Preserve any other text as-is
                        except (ValueError, TypeError):
                            wind_hail_lim = "0"

                        value = coverage_info.iloc[a+100, 1]
            
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                wind_hail_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    wind_hail_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    wind_hail_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    wind_hail_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    wind_hail_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    wind_hail_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            wind_hail_ded = "0"


                    
                        # Process values for building and business income if box is checked
                        value = get_value(coverage_info, a+115, 1, 0)
                        if value == 1:
                            bilim = "${:,}".format(int(float(businessincomesov))) if pd.notna(businessincomesov) and businessincomesov != 0 else ""
                            coverageamount = "${:,}".format(int(float(buildingvalue))) if pd.notna(buildingvalue) and buildingvalue != 0 else ""
                                    # Wind/Hail
                        value = coverage_info.iloc[a+96, 1]  # Starting at 96
                        if value == 1:
                            wind_hail_yes = "X"
                        else:
                            wind_hail_yes = ""
                        
                        value = coverage_info.iloc[a+97, 1]  # Incremented to 97
                        if value == 1:
                            wind_hail_no = "X"
                        else:
                            wind_hail_no = ""
                        
                        value = coverage_info.iloc[a+98, 1]  # Incremented to 98
                        if value == 1:
                            wind_hail_na = "X"
                        else:
                            wind_hail_na = ""
                        
                        value = coverage_info.iloc[a+101, 1]  # Incremented to 98
                        if value == 1:
                            wind_hail_yes2= "X"
                        else:
                            wind_hail_yes2 = ""
                        value = coverage_info.iloc[a+102, 1]  # Incremented to 98
                        if value == 1:
                            wind_hail_no2= "X"
                        else:
                            wind_hail_no2 = ""
                            
                        value = coverage_info.iloc[a+99, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                wind_hail_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    wind_hail_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    wind_hail_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    wind_hail_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    wind_hail_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            wind_hail_lim = "0"
                        
                        value = coverage_info.iloc[a+100, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                wind_hail_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    wind_hail_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    wind_hail_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    wind_hail_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    wind_hail_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    wind_hail_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            wind_hail_ded = "0"
                            
                        # Named Windstorm
                        value = coverage_info.iloc[a+104, 1]  # Starting at 104
                        if value == 1:
                            named_windstorm_yes = "X"
                        else:
                            named_windstorm_yes = ""
                        
                        value = coverage_info.iloc[a+105, 1]  # Incremented to 105
                        if value == 1:
                            named_windstorm_no = "X"
                        else:
                            named_windstorm_no = ""
                        
                        value = coverage_info.iloc[a+106, 1]  # Incremented to 106
                        if value == 1:
                            named_windstorm_na = "X"
                        else:
                            named_windstorm_na = ""
                        
                        value = coverage_info.iloc[a+109, 1]  # Incremented for secondary yes
                        if value == 1:
                            named_windstorm_yes2 = "X"
                        else:
                            named_windstorm_yes2 = ""
                        
                        value = coverage_info.iloc[a+110, 1]  # Incremented for secondary no
                        if value == 1:
                            named_windstorm_no2 = "X"
                        else:
                            named_windstorm_no2 = ""
                        
                        value = coverage_info.iloc[a+107, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                named_windstorm_lim = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    named_windstorm_lim = "Excluded"
                                elif value_str.lower() == "included":
                                    named_windstorm_lim = "Included"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    named_windstorm_lim = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    named_windstorm_lim = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            named_windstorm_lim = "0"
                        
                        value = coverage_info.iloc[a+108, 1]
                        
                        try:
                            if pd.isna(value) or str(value).strip() == "":
                                named_windstorm_ded = ""
                            else:
                                value_str = str(value).strip()
                                
                                if value_str.lower() == "excluded":
                                    named_windstorm_ded = "Excluded"
                                elif value_str.lower() == "included":
                                    named_windstorm_ded = "Included"
                                elif value_str.lower() == "see attached":
                                    named_windstorm_ded = "See Attached"
                                elif value_str.replace(',', '').replace('$', '').replace('.', '').isdigit():
                                    named_windstorm_ded = "${:,}".format(int(float(value_str.replace(',', '').replace('$', ''))))
                                else:
                                    named_windstorm_ded = value_str  # Preserve any other text as-is
                        
                        except (ValueError, TypeError):
                            named_windstorm_ded = "0"
            
            
                        # Subrogation
                        value = coverage_info.iloc[a+112, 1]  # Starting at 112
                        if value == 1:
                            subrogation_yes = "X"
                        else:
                            subrogation_yes = ""
                        
                        value = coverage_info.iloc[a+113, 1]  # Incremented to 113
                        if value == 1:
                            subrogation_no = "X"
                        else:
                            subrogation_no = ""
                        
                        value = coverage_info.iloc[a+114, 1]  # Incremented to 114
                        if value == 1:
                            subrogation_na = "X"
                        else:
                            subrogation_na = ""


                        
                     
                        # Set up the dynamic position for lender info on page 2
                        c = 580
                        for i, info in enumerate(lender_info):
                            if not info:  # If this lender info is empty
                                c = 570 - (i * 10)
                                break
                            if i == 11:  # If we've gone through all 12 and none are empty
                                c = 440
                        
                        # Prepare field data by page for PDF template
                        field_data_by_page = {
                            0: {  # Page 1
                                "Current_Date": (current_date, 530, 750),
                                "effective date": (effective_date, 325, 565),
                                "expiration date": (expiration_date, 405, 565),
                                "location_name": (location_name, 150, 510),
                                "Street_Address": (street_address, 300, 510),
                                "Town_Address": (town_address, 300, 500),
                                "csaemail": (csaemail, 160, 632),
                                "csa_Name": (csa_name, 25, 670),
                                "csa_number": (csanumb, 200, 692),
                                "csa_fax": (csanumb, 50, 632),
                                "EntitynamePlacer": ("Additional Named Insured Below", 25, 590),
                                "client_Name": (client_name, 25, 580),
                                "client_Line_Address": (client_line_address, 25, 570),
                                "client_2nd_Line_Address": (client_2nd_line_address, 25, 560),
                                "agency customer id": (agencycustomerid, 80, 608),
                                "loan number": (loannumbers, 325, 587),
                                "policy number":("See Attached", 490, 587),
                                
                                # Building Check
                                "building": (buildingcheck, 362, 525),
                                "business personal property": (buildingpersonalcheck, 438, 525),
                                
                                # Coverage Information
                                "basic": (basic, 250, 446.5),
                                "broad": (broad, 300, 446.5),
                                "special": (special, 355, 446.5),
                                "other": (other, 414, 446.5),
                                
                                # Commercial Property Coverage Amount 
                                "Coverage Amount": (coverageamount, 255, 435.5),
                                "deductible amount": (deductible, 500, 435),
                                
                                # Business Income
                                "Businessincome": (Businessincome, 264, 411),
                                "BI No check": (Businessincomeno, 278, 411),
                                "BI NA check": (Businessincomena, 293, 411),
                                "BI Check": (Bicheck, 24, 411),
                                "Rental Value": (Rentalvalue, 110, 411),
                                "bilim": (bilim, 355, 411),
                                "months": (months, 570, 411),
                                "losscheck": (losscheck, 450, 411),
                                
                                # Blanket Coverage
                                "Blanket Yes": (blanket_yes, 264, 398),
                                "Blanket No": (blanket_no, 278, 398),
                                "Blanket NA": (blanket_na, 293, 398),
                                "blanket lim": (blanket_lim, 500, 399),
                                
                                # Terrorism
                                "Terrorism Yes": (terror_yes, 264, 386),
                                "Terrorism No": (terror_no, 278, 386),
                                "Terrorism NA": (terror_na, 293, 386),
                                
                                # Terrorism Specific Exclusions
                                "Terrorism Exclusions Yes": (terror_exclusions_yes, 264, 374),
                                "Terrorism Exclusions No": (terror_exclusions_no, 278, 374),
                                "Terrorism Exclusions NA": (terror_exclusions_na, 293, 374),
                                
                                # Domestic Terrorism Exclusions
                                "Domestic Terrorism Exclusions Yes": (domestic_terror_exclusions_yes, 264, 362),
                                "Domestic Terrorism Exclusions No": (domestic_terror_exclusions_no, 278, 362),
                                "Domestic Terrorism Exclusions NA": (domestic_terror_exclusions_na, 293, 362),
                                
                                # Limited Fungus Coverage
                                "Limited Fungus Coverage Yes": (limited_fungus_coverage_yes, 264, 350),
                                "Limited Fungus Coverage No": (limited_fungus_coverage_no, 278, 350),
                                "Limited Fungus Coverage NA": (limited_fungus_coverage_na, 293, 350),
                                "Fungus lim":(fungus_lim,355,350),
                                "Fungus DED":(fungus_ded,520,351.5),
                                # Here you would add all the remaining field mappings...
                                # For brevity, we're only showing a subset
                                 # Fungus Exclusion
                                "Fungus Exclusion Yes": (fungus_exclusion_yes, 264, 338),  # Subtracted 12 from the y-axis
                                "Fungus Exclusion No": (fungus_exclusion_no, 278, 338),    # Subtracted 12 from the y-axis
                                "Fungus Exclusion NA": (fungus_exclusion_na, 293, 338),    # Subtracted 12 from the y-axis
                                # Replacement Cost
                                "Replacement Cost Yes": (replacement_cost_yes, 264, 326),  # Subtracted 12 from the y-axis
                                "Replacement Cost No": (replacement_cost_no, 278, 326),    # Subtracted 12 from the y-axis
                                "Replacement Cost NA": (replacement_cost_na, 293, 326),    # Subtracted 12 from the y-axis
            
                                # Agreed Value
                                "Agreed Value Yes": (agreed_value_yes, 264, 314),  # Subtracted 12 from the y-axis
                                "Agreed Value No": (agreed_value_no, 278, 314),    # Subtracted 12 from the y-axis
                                "Agreed Value NA": (agreed_value_na, 293, 314),    # Subtracted 12 from the y-axis
            
                                # Co-Insurance
                                "Co-Insurance Yes": (co_insurance_yes, 264, 302),  # Subtracted 12 from the y-axis
                                "Co-Insurance No": (co_insurance_no, 278, 302),    # Subtracted 12 from the y-axis
                                "Co-Insurance NA": (co_insurance_na, 293, 302),    # Subtracted 12 from the y-axis
                                "Coinsurance %":(co_ins_percent,350,303),
                                # Equipment Breakdown
                                "Equipment Breakdown Yes": (equipment_breakdown_yes, 264, 290),  # Adjusted y-axis to 302 - 12
                                "Equipment Breakdown No": (equipment_breakdown_no, 278, 290),    # Adjusted y-axis to 302 - 12
                                "Equipment Breakdown NA": (equipment_breakdown_na, 293, 290),    # Adjusted y-axis to 302 - 12
                                "Equipment Breakdown Limit": (equipment_breakdown_lim, 355, 290),  # Adjusted y-axis to 302 - 12
                                "Equipment Breakdown Deductible": (equipment_breakdown_ded, 520, 291.5),  # Adjusted y-axis to 302 - 12
            
                                # Coverage A
                                "Coverage A Yes": (coverage_a_yes, 264, 278),  # Adjusted y-axis to 290 - 12
                                "Coverage A No": (coverage_a_no, 278, 278),    # Adjusted y-axis to 290 - 12
                                "Coverage A NA": (coverage_a_na, 293, 278),    # Adjusted y-axis to 290 - 12
                                "Coverage A Limit": (coverage_a_lim, 355, 278),  # Adjusted y-axis to 290 - 12
                                "Coverage A Deductible": (coverage_a_ded, 520, 279.5),  # Adjusted y-axis to 290 - 12
            
                                # Coverage B
                                "Coverage B Yes": (coverage_b_yes, 264, 266),  # Adjusted y-axis to 278 - 12
                                "Coverage B No": (coverage_b_no, 278, 266),    # Adjusted y-axis to 278 - 12
                                "Coverage B NA": (coverage_b_na, 293, 266),    # Adjusted y-axis to 278 - 12
                                "Coverage B Limit": (coverage_b_lim, 355, 266),  # Adjusted y-axis to 278 - 12
                                "Coverage B Deductible": (coverage_b_ded, 520, 267.5),  # Adjusted y-axis to 278 - 12
            
                                # Coverage C
                                "Coverage C Yes": (coverage_c_yes, 264, 254),  # Adjusted y-axis to 266 - 12
                                "Coverage C No": (coverage_c_no, 278, 254),    # Adjusted y-axis to 266 - 12
                                "Coverage C NA": (coverage_c_na, 293, 254),    # Adjusted y-axis to 266 - 12
                                "Coverage C Limit": (coverage_c_lim, 355, 254),  # Adjusted y-axis to 266 - 12
                                "Coverage C Deductible": (coverage_c_ded, 520, 255.5),  # Adjusted y-axis to 266 - 12
            
                                # Earth Movement
                                "Earth Movement Yes": (earth_movement_yes, 264, 242),  # Adjusted y-axis to 254 - 12
                                "Earth Movement No": (earth_movement_no, 278, 242),    # Adjusted y-axis to 254 - 12
                                "Earth Movement NA": (earth_movement_na, 293, 242),    # Adjusted y-axis to 254 - 12
                                "Earth Movement Limit": (earth_movement_lim, 355, 242),  # Adjusted y-axis to 254 - 12
                                "Earth Movement Deductible": (earth_movement_ded, 520, 243.5),  # Adjusted y-axis to 254 - 12
            
                                # Flood
                                "Flood Yes": (flood_yes, 264, 230),  # Adjusted y-axis to 242 - 12
                                "Flood No": (flood_no, 278, 230),    # Adjusted y-axis to 242 - 12
                                "Flood NA": (flood_na, 293, 230),    # Adjusted y-axis to 242 - 12
                                "Flood Limit": (flood_lim, 355, 230),  # Adjusted y-axis to 242 - 12
                                "Flood Deductible": (flood_ded, 520, 231.5),  # Adjusted y-axis to 242 - 12
                                
                                # Wind/Hail
                                "Wind/Hail Yes": (wind_hail_yes, 264, 218),  # Adjusted y-axis to 230 - 12
                                "Wind/Hail No": (wind_hail_no, 278, 218),    # Adjusted y-axis to 230 - 12
                                "Wind/Hail NA": (wind_hail_na, 293, 218),    # Adjusted y-axis to 230 - 12
                                "Wind/Hail Limit": (coverageamount, 355, 218),  # Adjusted y-axis to 230 - 12
                                "Wind/Hail Deductible": (wind_hail_ded, 520, 219.5),  # Adjusted y-axis to 230 - 12
                                "Wind/Hail y": (wind_hail_yes2, 103, 219),
                                "Wind/Hail n": (wind_hail_no2, 135, 219),
            
            
                                # Named Windstorm
                                "Named Windstorm Yes": (named_windstorm_yes, 264, 206),  # Adjusted y-axis to 218 - 12
                                "Named Windstorm No": (named_windstorm_no, 278, 206),    # Adjusted y-axis to 218 - 12
                                "Named Windstorm NA": (named_windstorm_na, 293, 206),    # Adjusted y-axis to 218 - 12
                                "Named Windstorm Limit": (coverageamount, 355, 206),  # Adjusted y-axis to 218 - 12
                                "Named Windstorm Deductible": (named_windstorm_ded, 520, 207.5),  # Adjusted y-axis to 218 - 12
                                "Named Windstorm y": (named_windstorm_yes2, 103, 207),  # Adjusted y-axis to 218 - 12
                                "Named Windstorm n": (named_windstorm_no2, 135, 207),   # Adjusted y-axis to 218 - 12
                                # Certificate Holder Specification
                                "Lender Contract of Sale": (lender_cos_check, 23, 128.5),
                                "Lender Loss Payable": (lender_llp_check, 120, 128.5),
                                "Lender Loss Payee": (lender_lp_check, 233.5, 128.5),
                                "Lender Mortgagee": (lender_m_check, 23, 117),
                                "Other Lender": (other_lender_check, 118.5, 117),
                                "Other Lender Text": (other_lender_text, 135, 117),

                                # Permission to waive
                                "PermYes": (subrogation_yes, 264, 194),  # Adjusted y-axis to 242 - 12
                                "PermNo": (subrogation_no, 278,194),    # Adjusted y-axis to 242 - 12
                                "PermNA": (subrogation_na, 293, 194),    # Adjusted y-axis to 242 - 12
                
                                # Lender 
                                "Lender Name": (lender_name, 50, 90),
                                "Lender Address": (lender_address, 50, 80),
                                "Addtional Lender Address": (lender_address_additional, 50, 70),
                                "Lender Address 2": (lender_address2, 50, 60),
                            },
                            1: {  # Page 2
                                "Client Name": (client_name, 310, 690),
                                "Additional Information": ("Additional Named Insured", 22, c),
                                "Statement": ("-30-day notice of cancellation applies except 10-day notice for non-payment of premium to Additional Interest / Mortgagee / Loss Payee", 22, 590),
                                "Entitynamer": (entity_name, 22, c-10),
                            }
                        }
                        
                        # Add lender information to page 2
                        for i, info in enumerate(lender_info):
                            if info:  # Only add if not empty
                                field_data_by_page[1][f"Lender {i+1}"] = (info, 22, 570 - (i * 10))
                        
                        # Add page 2 carriers from Additional Remarks Page 1
                        page_2_carriers = {}
                        y_position = c-30  # Starting Y-coordinate for placement
                        previous_index = None  # To track the last processed index
                        
                        for index, row in df_carriers.iterrows():  # Loop through the DataFrame using index
                            try:
                                value = row.iloc[8] if 8 < len(row) else ""  # Access the column value safely
                                
                                # Include the row if the value is not NaN and contains meaningful text
                                if not pd.isnull(value) and str(value).strip() != "":
                                    if previous_index is not None:
                                        # Calculate the number of skipped lines based on the index difference
                                        skipped_rows = index - previous_index - 1
                                        y_position -= skipped_rows * 10  # Adjust Y-position for skipped rows
                                    
                                    # Add the current value to the carriers
                                    page_2_carriers[str(index)] = (value, 22, y_position)
                                    y_position -= 10  # Decrease Y-coordinate for the current row
                                    previous_index = index  # Update the last processed index
                            except Exception as e:
                                continue
                        
                        # Add carriers to Page 2
                        field_data_by_page[1].update(page_2_carriers)
                        
                        # Add page 3 for additional remarks
                        field_data_by_page.update({
                            2: {  # Page 3
                                "Client Name": (client_name, 310, 690),
                            }
                        })
                        
                        # Add remarks from Additional Remarks Page 2
                        page_3_carriers = {}
                        y_position = 580  # Starting Y-coordinate for placement
                        previous_index = None  # To track the last processed index
                        
                        for index, row in df_carriers2.iterrows():  # Loop through the DataFrame using index
                            try:
                                value = row.iloc[1] if 1 < len(row) else ""  # Access the column value safely
                                
                                # Include the row if the value is not NaN and contains meaningful text
                                if not pd.isnull(value) and str(value).strip() != "":
                                    if previous_index is not None:
                                        # Calculate the number of skipped lines based on the index difference
                                        skipped_rows = index - previous_index - 1
                                        y_position -= skipped_rows * 10  # Adjust Y-position for skipped rows
                                    
                                    # Add the current value to the carriers
                                    page_3_carriers[str(index)] = (value, 22, y_position)
                                    y_position -= 10  # Decrease Y-coordinate for the current row
                                    previous_index = index  # Update the last processed index
                            except Exception as e:
                                continue
                        
                        # Add carriers to Page 3
                        field_data_by_page[2].update(page_3_carriers)
                        
                        # Create a safe filename
                        clean_lender_name = re.sub(r'[^A-Za-z\s;:]', '', lender_name)
                        lender_first_word = lender_name.split()[0] if lender_name and len(lender_name.split()) > 0 else 'NA'
                        safe_location_name = ''.join(c for c in location_name if c.isalnum() or c in ' _-')[:30]
                        safe_location_name = safe_location_name.replace(' ', '_')
                        
                        filename = f"{current_date_year}-{next_year} EPI {safe_location_name}_{lender_first_word}_{certnum}.pdf"
                        output_pdf = os.path.join(output_dir, filename)
                        
                        # Fill the template
                        if fill_epi_template(template_path, output_pdf, field_data_by_page):
                            # Add to zip file
                            zip_file.write(output_pdf, filename)
                            successful_pdfs += 1
                        
                        # Update progress
                        progress_bar.progress((index + 1) / total_rows)
                    
                    except Exception as e:
                        status_text.text(f"Error processing row {index + 1}: {e}")
                        st.warning(f"Error processing row {index + 1}: {e}. Check to make sure location name matches from the lender tab to SOV tab.")
                
                # Update final status
                progress_bar.progress(1.0)
                status_text.text(f"Completed: Generated {successful_pdfs} PDFs out of {total_rows} rows.")
            
            # Return the zip file for download
            zip_buffer.seek(0)
            return zip_buffer
            
        except Exception as e:
            st.error(f"Error processing file: {e}")
            return None

# Main function
def main():
    st.markdown('<p class="main-header">EPI Generator</p>', unsafe_allow_html=True)
    st.markdown("Generate Evidence of Property Insurance from Excel data")
    
    # File uploader - placing it earlier in the interface
    st.markdown('<p class="subheader">Upload Excel File</p>', unsafe_allow_html=True)
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xlsx", "xls"])
    
    # Producer selection
    st.markdown('<p class="subheader">Select Producer</p>', unsafe_allow_html=True)
    producers = sorted(["Garet Marr", "Jared Kahhan", "Matt Armstrong", "Matt Harrell", "Michael Shadeed", "Brad Young", "Michael Conlon", "Evan Seacat", "Patrick Hall", "Pete Moore","Custom Format"])
    producer = st.selectbox("Producer", producers)
    
    # Output folder name
    st.markdown('<p class="subheader">Enter Output ZIP Name</p>', unsafe_allow_html=True)
    folder_name = st.text_input("ZIP Filename (without .zip extension)")
    
    # Template selection based on producer
    template_options = {
        "Garet Marr": "EPI Template - Garet Marr.pdf",
        "Jared Kahhan": "EPI Template - Jared Kahhan.pdf",
        "Matt Armstrong": "EPI Template - Matt Armstrong.pdf",
        "Matt Harrell": "EPI Template - Matt Harrell.pdf",
        "Michael Shadeed": "EPI Template - Michael Shadeed.pdf",
        "Brad Young": "EPI Template - Brad Young.pdf",
        "Michael Conlon": "EPI Template - Michael Conlon.pdf",
        "Evan Seacat": "EPI Template - Evan Seacat.pdf",
        "Patrick Hall": "EPI Template - Patrick Hall.pdf",
        "Pete Moore": "EPI Template - Pete Moore.pdf",
        "Custom Format": "EPI Template - Custom Format.pdf"
    }
    
    # Get template path based on producer selection
    template_path = template_options.get(producer)
    
    # Allow custom template path
    use_custom_path = st.checkbox("Use custom template path")
    if use_custom_path:
        template_path = st.text_input("Enter custom template path")
    
    # Process button - only one instance with a unique key
    if st.button("Generate EPIs", key="generate_button"):
        if not folder_name:
            st.error("Please enter a ZIP filename")
        elif not uploaded_file:
            st.error("Please upload an Excel file")
        elif not template_path:
            st.error("Please select a template or provide a custom path")
        elif not os.path.exists(template_path):
            st.error(f"Template file not found: {template_path}")
        else:
            with st.spinner("Processing... This may take a few minutes for large files"):
                # Process the file with the local template
                zip_buffer = process_sov(uploaded_file, template_path, producer)
                

            
                if zip_buffer:
                # Create download link
                    download_filename = f"{folder_name}.zip"
                    st.success(f"✅ {download_filename} is ready for download!")
                    st.download_button(
                        label=f"Download {download_filename}",
                        data=zip_buffer,
                        file_name=download_filename,
                        mime="application/zip"
                    )


    
    # Instructions
    with st.expander("How to use this app"):
        st.markdown("""
        1. Upload your Excel file with the SOV data
        2. Select the producer from the dropdown
        3. Enter a name for the output ZIP file (without .zip extension)
        4. The app will automatically select the PDF template based on the producer
        5. Click 'Generate EPIs' to process the file
        6. Download the ZIP file containing all generated PDFs
        
        **Required Excel Sheets:**
        - SOV (with header at row 11)
        - Property Coverage Information
        - Additional Remarks Page 1
        - Additional Remarks Page 2
        - Lenders
        
        **Required Columns in SOV Sheet:**
        - Location Name
        - Entity Name
        - *Street Address
        - *City
        - *State Code
        - *Zip
        - *Real Property Value ($)
        - BI/Rental Income ($)
        
        **Required Columns in Lenders Sheet:**
        - Location Name (must match the SOV sheet)
        - Lender Name
        - Street Address
        - Street Address 2 (optional)
        - City
        - State
        - Zipcode
        - Contract of Sale
        - Lenders Loss Payable
        - Loss Payee
        - Mortgagee
        - Other (Type in below)
        - Lender Additional Information 1-12 (optional)
        """)
    
    # About section
    with st.expander("About"):
        st.markdown("""
        **EPI Generator**
        
        This application generates Evidence of Property Insurance (EPI) from an Excel file containing insurance data.
        It processes the SOV (Schedule of Values) data, matches it with lender information, and creates 
        individual PDF certificates for each location.
        
        The app is built with Streamlit and deployed on Render.
        
        **Need help?**
        
        If you encounter any issues or have questions, please contact support.
        """)

# Run the main function
if __name__ == "__main__":
    main()
