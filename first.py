import streamlit as st
import pandas as pd
from fpdf import FPDF
import tempfile
import os
from datetime import datetime
import PyPDF2
from io import BytesIO
import fitz
import re
from typing import Optional, List, Dict, Tuple

# Initialize session state
if 'pdf_generated' not in st.session_state:
    st.session_state.pdf_generated = False
if 'generated_pdf_bytes' not in st.session_state:
    st.session_state.generated_pdf_bytes = None
if 'company_code' not in st.session_state:
    st.session_state.company_code = "PRF"
if 'show_advanced' not in st.session_state:
    st.session_state.show_advanced = False
if 'extracted_data' not in st.session_state:
    st.session_state.extracted_data = {}
if 'cusdec_file' not in st.session_state:
    st.session_state.cusdec_file = None

class PaymentRequestPDF(FPDF):
    def __init__(self):
        super().__init__()
        self.set_auto_page_break(auto=True, margin=15)
    
    def header(self):
        self.set_font('Arial', 'B', 16)
        self.cell(0, 10, 'Payment Request Form', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

def find_label_rect_on_page(page: fitz.Page, keywords: List[str]) -> Optional[Tuple[fitz.Rect, str]]:
    """Find the label (e.g., 'Exporter', 'Consignee', 'Vessel/Flight') on the page and return its bounding box."""
    lower_keywords = [k.lower() for k in keywords]

    # --- Search in blocks first (usually accurate for printed forms)
    for b in page.get_text("blocks"):
        text = (b[4] or "").strip()
        lower_text = text.lower()
        if any(k in lower_text for k in lower_keywords):
            return fitz.Rect(b[0], b[1], b[2], b[3]), text

    # --- Fallback: search word by word
    for w in page.get_text("words"):
        wt = (w[4] or "").strip()
        lower_wt = wt.lower()
        if any(k in lower_wt for k in lower_keywords):
            return fitz.Rect(w[0], w[1], w[2], w[3]), wt

    return None

def expand_rect(rect: fitz.Rect, page: fitz.Page,
                right_px: float = 200, down_px: float = 80) -> fitz.Rect:
    """Expand the found label area to include the full cage box region."""
    x0 = max(0, rect.x0 - 4)
    y0 = max(0, rect.y0 - 2)
    x1 = min(page.rect.x1 - 5, rect.x1 + right_px)
    y1 = min(page.rect.y1 - 5, rect.y1 + down_px)
    return fitz.Rect(x0, y0, x1, y1)

def extract_text_in_rect(page: fitz.Page, rect: fitz.Rect) -> List[str]:
    """Extract and preserve text line order inside a given rectangular region."""
    words = page.get_text("words")
    selected = [w for w in words if fitz.Rect(w[:4]).intersects(rect)]
    if not selected:
        return []

    # Sort top-to-bottom then left-to-right
    selected.sort(key=lambda w: (round(w[1], 2), w[0]))

    # Group into lines
    lines_map = {}
    for w in selected:
        y = round(w[1], 1)
        lines_map.setdefault(y, []).append(w)

    lines = []
    for y in sorted(lines_map.keys()):
        row_words = sorted(lines_map[y], key=lambda x: x[0])
        line_text = " ".join(t[4] for t in row_words).strip()
        if line_text:
            lines.append(line_text)

    return lines

def extract_label_cage_text(
    pdf_path: str,
    keywords: List[str],
    page_no: Optional[int] = None,
    right_expand: float = 200,
    down_expand: float = 80
) -> Optional[Dict]:
    """
    Extracts all text within the cage of the given label (e.g., 'Vessel/Flight').
    Keeps text order exactly as shown in CUSDEC.
    """
    doc = fitz.open(pdf_path)
    pages_to_check = [page_no] if page_no is not None else range(len(doc))

    for pno in pages_to_check:
        page = doc[pno]
        found = find_label_rect_on_page(page, keywords)
        if not found:
            continue

        label_rect, label_text = found
        cage_rect = expand_rect(label_rect, page, right_expand, down_expand)
        cage_lines = extract_text_in_rect(page, cage_rect)

        return {
            "page": pno,
            "label": label_text,
            "rect": cage_rect,
            "lines": cage_lines
        }

    return None

def extract_company_name_from_consignee(lines: List[str]) -> str:
    """
    Extract company name from consignee lines.
    Looks for known company names in the extracted text.
    """
    # Known company mappings
    company_mappings = {
        "MAS CAPITAL PVT LTD": "MAS CAPITAL PVT LTD",
        "BODYLINE PVT LTD": "BODYLINE PVT LTD", 
        "UNICHELA PVT LTD": "UNICHELA PVT LTD",
        "MAS CAPITAL": "MAS CAPITAL PVT LTD",
        "BODYLINE": "BODYLINE PVT LTD",
        "UNICHELA": "UNICHELA PVT LTD"
    }
    
    # Combine all lines for searching
    full_text = " ".join(lines).upper()
    
    # Look for exact company names
    for company_key, company_full in company_mappings.items():
        if company_key in full_text:
            return company_full
    
    # If no exact match found, return the first meaningful line after "Consignee"
    for i, line in enumerate(lines):
        line_upper = line.upper()
        if "CONSIGNEE" in line_upper and i + 1 < len(lines):
            # Return the next line which should be the company name
            return lines[i + 1].strip()
    
    return ""


def extract_invoice_number_from_customs_ref(lines: List[str]) -> Dict[str, str]:
    """
    Extract invoice number components from Customs Reference Number lines.
    Expected format: "S 52194 30/09/2025"
    Returns: {'prefix': 'S', 'number': '52194', 'year': '2025'}
    """
    # Combine all lines for searching
    full_text = " ".join(lines)
    
    # Look for the pattern: letter + space + digits + space + date
    pattern = r'([A-Z])\s+(\d+)\s+(\d{2}/\d{2}/(\d{4}))'
    match = re.search(pattern, full_text)
    
    if match:
        prefix = match.group(1)  # S
        number = match.group(2)  # 52194
        year = match.group(4)    # 2025 (from 30/09/2025)
        
        return {
            'prefix': prefix,
            'number': number,
            'year': year
        }
    
    # Alternative pattern if the first one doesn't match
    pattern2 = r'([A-Z])\s*(\d{4,6})\s*.*?(\d{4})'
    match2 = re.search(pattern2, full_text)
    
    if match2:
        return {
            'prefix': match2.group(1),
            'number': match2.group(2),
            'year': match2.group(3)
        }
    
    return {'prefix': '', 'number': '', 'year': ''}


def extract_invoice_date_from_customs_ref(lines: List[str]) -> str:
    """
    Extract invoice date from Customs Reference Number lines.
    Expected format: "S 52194 30/09/2025"
    Returns: '30/09/2025'
    """
    # Combine all lines for searching
    full_text = " ".join(lines)
    
    # Look for the date pattern: dd/mm/yyyy
    date_pattern = r'(\d{2}/\d{2}/\d{4})'
    matches = re.findall(date_pattern, full_text)
    
    if matches:
        # Return the first date found (which should be the invoice date)
        return matches[0]
    
    return ""


def extract_gross_value_from_total_declaration(lines: List[str]) -> float:
    """
    Extract gross value from Total Declaration lines.
    Expected format: "499,180" before "Total Declaration"
    Returns: 499180.00
    """
    # Combine all lines for searching
    full_text = " ".join(lines)
    
    # Look for number patterns before "Total Declaration"
    # Pattern: digits with commas followed by "Total Declaration"
    pattern = r'([\d,]+)\s*Total Declaration'
    match = re.search(pattern, full_text)
    
    if match:
        # Extract the number and remove commas
        number_str = match.group(1).replace(',', '')
        try:
            return float(number_str)
        except ValueError:
            pass
    
    # Alternative: look for any number in the lines near "Total Declaration"
    for i, line in enumerate(lines):
        if "total declaration" in line.lower():
            # Check previous line for number
            if i > 0:
                prev_line = lines[i-1]
                # Extract numbers with commas
                numbers = re.findall(r'[\d,]+', prev_line)
                if numbers:
                    try:
                        return float(numbers[0].replace(',', ''))
                    except ValueError:
                        pass
    
    return 0.0

def get_cost_center_by_company(company_name: str) -> str:
    """
    Get cost center based on company name.
    """
    cost_center_map = {
        "BODYLINE PVT LTD": "B051ADMN01",
        "UNICHELA PVT LTD": "A051ADMN01", 
        "MAS CAPITAL PVT LTD": "MCAP051ADMN01"
    }
    
    return cost_center_map.get(company_name, "B051ADMN01")  # Default to Bodyline

def extract_vat_amount_from_tax_table(lines: List[str]) -> float:
    """
    Extract VAT amount from tax table lines.
    Looks for "VAT" or "VTD" row and extracts the amount.
    Format: "VAT 1,749,003 18.00 314,821 1"
    Returns: 314821.00 only if trailing flag is 1, otherwise 0.0
    """
    # Combine all lines for searching
    full_text = " ".join(lines)
    
    # Look for VAT or VTD row pattern with trailing flag
    # Pattern: VAT/VTD followed by numbers and then amount and trailing 0/1
    pattern = r'(?:VAT|VTD)\s+[\d,.]+\s+[\d,.]+\s+([\d,]+)\s*([01])'
    match = re.search(pattern, full_text)
    
    if match:
        amount_str = match.group(1).replace(',', '')
        trailing_flag = match.group(2)
        
        # Only include amount if trailing flag is 1
        if trailing_flag == '1':
            try:
                return float(amount_str)
            except ValueError:
                pass
        else:
            # Trailing flag is 0, so VAT amount should be 0
            return 0.0
    
    # Alternative: line by line search with flag check
    for line in lines:
        line_upper = line.upper()
        if any(tax_type in line_upper for tax_type in ['VAT', 'VTD']):
            # Extract all numbers from the line
            numbers = re.findall(r'[\d,]+', line)
            if len(numbers) >= 4:
                # The amount is typically the 3rd number, and flag is the 4th
                amount_str = numbers[2].replace(',', '')
                trailing_flag = numbers[3]
                
                if trailing_flag == '1':
                    try:
                        return float(amount_str)
                    except ValueError:
                        pass
                else:
                    return 0.0
            elif len(numbers) >= 3:
                # If no flag, assume we should include the amount
                try:
                    return float(numbers[2].replace(',', ''))
                except ValueError:
                    pass
    
    return 0.0


def extract_vat_from_summary_of_taxes(pdf_path: str) -> float:
    """
    Extract VAT amount from Summary of Taxes section across all pages.
    Looks for the pattern and extracts the VAT amount.
    """
    doc = fitz.open(pdf_path)
    vat_amount = 0.0
    
    for page_num in range(len(doc)):
        page = doc[page_num]
        
        # Look for Summary of Taxes section
        summary_result = extract_label_cage_text(
            pdf_path=pdf_path,
            keywords=["Summary of Taxes"],
            page_no=page_num,
            right_expand=150,
            down_expand=150
        )
        
        if summary_result:
            st.success(f"✅ Found Summary of Taxes on page: {page_num + 1}")
            st.write(f"Label: {summary_result['label']}")
            st.write("Extracted text (preserved order):")
            
            for line in summary_result["lines"]:
                st.write(line)
            
            # Look for VAT line in the extracted text
            for line in summary_result["lines"]:
                if "VAT" in line.upper():
                    # Extract numbers from the VAT line
                    numbers = re.findall(r'[\d,]+', line)
                    if numbers:
                        # Usually the first number after "VAT" is the amount
                        vat_str = numbers[0].replace(',', '')
                        try:
                            vat_amount = float(vat_str)
                            st.success(f"✅ Extracted VAT Amount: {vat_amount:,.2f}")
                            doc.close()
                            return vat_amount
                        except ValueError:
                            continue
    
    doc.close()
    return vat_amount


def extract_office_code_from_customs_ref(lines: List[str]) -> str:
    """
    Extract office code from Customs Reference Number lines.
    Expected format: "CBBI2 Colombo Boi Imports(Air"
    Returns: 'CBBI2'
    """
    # Combine all lines for searching
    full_text = " ".join(lines)
    
    # Look for office code pattern: 4-5 uppercase letters + digits
    pattern = r'([A-Z]{3,4}\d{1,2})\s+[A-Za-z]'
    match = re.search(pattern, full_text)
    
    if match:
        return match.group(1)
    
    return "CBBI1"  # Default fallback


def process_cusdec_pdf(uploaded_file) -> Dict:
    """
    Process CUSDEC PDF and extract relevant data
    """
    extracted_data = {}
    
    # Save uploaded file to temporary location
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        tmp_file.write(uploaded_file.getvalue())
        tmp_path = tmp_file.name
    
    try:
        # Extract Consignee information
        consignee_result = extract_label_cage_text(
            pdf_path=tmp_path,
            keywords=["Consignee"],
            right_expand=80,
            down_expand=50
        )
        
        if consignee_result:
            company_name = extract_company_name_from_consignee(consignee_result["lines"])
            if company_name:
                extracted_data['company_name'] = company_name
                st.success(f"✅ Automatically detected Company: {company_name}")

        # Extract Customs Reference Number for Invoice Number and Date
        customs_ref_result = extract_label_cage_text(
            pdf_path=tmp_path,
            keywords=["Customs Reference Number"],
            right_expand=150,
            down_expand=50
        )
        
        if customs_ref_result:
            office_code = extract_office_code_from_customs_ref(customs_ref_result["lines"])
            if office_code:
                extracted_data['office_code'] = office_code
                st.success(f"✅ Automatically detected Office Code: {office_code}")

            # Extract invoice number components
            invoice_data = extract_invoice_number_from_customs_ref(customs_ref_result["lines"])
            if invoice_data['prefix'] and invoice_data['number'] and invoice_data['year']:
                extracted_data['invoice_prefix'] = invoice_data['prefix']
                extracted_data['invoice_number'] = invoice_data['number']
                extracted_data['invoice_year'] = invoice_data['year']
                st.success(f"✅ Automatically detected Invoice: {invoice_data['prefix']} {invoice_data['number']} {invoice_data['year']}")
            
            # Extract invoice date
            invoice_date_str = extract_invoice_date_from_customs_ref(customs_ref_result["lines"])
            if invoice_date_str:
                extracted_data['invoice_date'] = invoice_date_str
                st.success(f"✅ Automatically detected Invoice Date: {invoice_date_str}")
        
        # Extract Gross Value from Total Declaration
        total_declaration_result = extract_label_cage_text(
            pdf_path=tmp_path,
            keywords=["Total Declaration"],
            right_expand=100,
            down_expand=50
        )
        
        if total_declaration_result:
            gross_value = extract_gross_value_from_total_declaration(total_declaration_result["lines"])
            if gross_value > 0:
                extracted_data['gross_value'] = gross_value
                st.success(f"✅ Automatically detected Gross Value: {gross_value:,.2f}")
        
        # Try to extract VAT from Summary of Taxes first (for multi-page PDFs)
        vat_amount = extract_vat_from_summary_of_taxes(tmp_path)
        
        if vat_amount > 0:
            extracted_data['vat_amount'] = vat_amount
            st.success(f"✅ Automatically detected VAT Amount from Summary: {vat_amount:,.2f}")
        else:
            # Fallback to tax table extraction if Summary of Taxes not found
            tax_table_result = extract_label_cage_text(
                pdf_path=tmp_path,
                keywords=["Amount", "Rate", "Tax Base"],
                right_expand=200,
                down_expand=100
            )
            
            if tax_table_result:
                vat_amount = extract_vat_amount_from_tax_table(tax_table_result["lines"])
                if vat_amount > 0:
                    extracted_data['vat_amount'] = vat_amount
                    st.success(f"✅ Automatically detected VAT Amount from Tax Table: {vat_amount:,.2f}")
                
                # Display extraction details for debugging
                with st.expander("View Tax Table Extraction Details"):
                    st.write("Extracted Lines:")
                    for i, line in enumerate(tax_table_result["lines"]):
                        st.write(f"{i}: {line}")
                    st.write(f"Parsed VAT Amount: {vat_amount}")
        
    except Exception as e:
        st.error(f"Error processing PDF: {str(e)}")
    finally:
        # Clean up temporary file
        if os.path.exists(tmp_path):
            os.unlink(tmp_path)
    
    return extracted_data


def create_payment_request_pdf(data):
    pdf = PaymentRequestPDF()
    pdf.add_page()
    
      # Add border around the entire page
    pdf.set_draw_color(0, 0, 0)  # Black color
    pdf.set_line_width(0.5)  # Line width
    pdf.rect(8, 8, 190, 150)   

    # Set font
    pdf.set_font('Arial', '', 12)
    
    # Company Name and Currency
    pdf.cell(45, 10, 'Company Name', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['company_name'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, 'Currency :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['currency'], 0, 1)
    
    # Gross Value and Invoice Number
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'Gross Value :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, f"{data['gross_value']:,.2f}", 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, 'Invoice Number :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['invoice_number'], 0, 1)
    
    # Vendor Code and Invoice Date
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'Vendor Code :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['vendor_code'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, 'Invoice Date :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['invoice_date'], 0, 1)
    
    # Vendor and SOURCEEMAIL
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'Vendor :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['vendor'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, 'SOURCEEMAIL :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['source_email'], 0, 1)
    
    # PO Number (centered)
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, '', 0, 0)
    pdf.cell(45, 10, '', 0, 0)
    pdf.cell(35, 10, 'PO :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['po_number'], 0, 1)
    
    # Description and Cost Center
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'Description', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['description'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, 'Cost Center', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['cost_center'], 0, 1)
    
    # Functional Area and Payment Method
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'Functional Area :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['functional_area'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, 'Payment Method :', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, data['payment_method'], 0, 1)
    
    # GL Account and Amount
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'GL Account', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['gl_account'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, '', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, f"{data['gl_amount']:,.2f}", 0, 1)
    
    # VAT GL Account and Amount
    pdf.set_font('Arial', '', 12)
    pdf.cell(45, 10, 'VAT GL Account', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(45, 10, data['vat_gl_account'], 0, 0)
    pdf.set_font('Arial', '', 12)
    pdf.cell(35, 10, '', 0, 0)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, f"{data['vat_amount']:,.2f}", 0, 1)
    
    return pdf

def merge_pdfs(pdf1_bytes, pdf2_bytes):
    """Merge two PDF files"""
    merger = PyPDF2.PdfMerger()
    
    # Add first PDF (the one we generated)
    merger.append(BytesIO(pdf1_bytes))
    
    # Add second PDF (the uploaded CUSDEC)
    merger.append(BytesIO(pdf2_bytes))
    
    # Create merged PDF in memory
    merged_pdf = BytesIO()
    merger.write(merged_pdf)
    merger.close()
    
    return merged_pdf.getvalue()

def main():
    st.set_page_config(page_title="Payment Request Form Generator", layout="wide")
    
    # Initialize session state for extracted data if not exists
    if 'extracted_data' not in st.session_state:
        st.session_state.extracted_data = {}
    if 'cusdec_file' not in st.session_state:
        st.session_state.cusdec_file = None
    
    
    st.title("📄 Payment Request Form Generator")
    
    # Create tabs
    tab1, tab2 = st.tabs(["WC Creation", "PayReq Creation"])

    with tab1:
        st.write("Fill in the details below to generate a Payment Request Form PDF")
        
        # CUSDEC PDF Upload
        st.subheader("📤 Upload CUSDEC PDF")
        cusdec_file = st.file_uploader("Upload CUSDEC PDF for auto-fill and merging", type="pdf", key="cusdec_uploader")
        
        # Auto-process checkbox
        auto_process = st.checkbox("Automatically extract data after upload", value=True)
        
        if cusdec_file is not None:
            # DEBUG: Show what's in extracted_data
            st.write("🔍 DEBUG - Current extracted_data:", st.session_state.extracted_data)
            
            # Track if we've processed this file
            current_file_id = f"{cusdec_file.name}_{cusdec_file.size}"
            
            if ('last_processed_file' not in st.session_state or 
                st.session_state.last_processed_file != current_file_id):
                
                if auto_process:
                    with st.spinner("Extracting data from CUSDEC PDF..."):
                        extracted_data = process_cusdec_pdf(cusdec_file)
                        st.session_state.extracted_data = extracted_data
                        st.session_state.cusdec_file = cusdec_file
                        st.session_state.last_processed_file = current_file_id
                        st.success(f"✅ Data extracted automatically!")
                        # DEBUG: Show what was just extracted
                        st.write("🔍 DEBUG - Freshly extracted data:", extracted_data)
                        st.rerun()
        
        # 🔥 CRITICAL FIX FOR STREAMLIT CLOUD 🔥
        # Initialize form fields with extracted data BEFORE they render
        if st.session_state.extracted_data:
            # Force update the form field values in session state
            for key in ['invoice_prefix', 'invoice_number', 'invoice_year']:
                if key in st.session_state.extracted_data:
                    # Set the form field values in session state
                    field_value = str(st.session_state.extracted_data[key])
                    if f"{key}_input" not in st.session_state:
                        st.session_state[f"{key}_input"] = field_value  
                        
        # Toggle for advanced fields
        st.session_state.show_advanced = st.checkbox("Show All Fields", value=False)
        
        # Create two columns for better layout
        col1, col2 = st.columns(2)
        
        with col1:

            st.subheader("Essential Information")
            company_options = ["", "BODYLINE PVT LTD", "UNICHELA PVT LTD", "MAS CAPITAL PVT LTD"]

            # Cloud-safe company selection
            extracted_company = st.session_state.extracted_data.get('company_name', '')
            # Handle both cases: empty string or None
            if not extracted_company:
                company_index = 0
            else:
                company_index = company_options.index(extracted_company) if extracted_company in company_options else 0

            company_name = st.selectbox("Company Name *", company_options, index=company_index, key="company_select")

            # Invoice Number components - Cloud optimized
            st.write("Invoice Number *")
            col_inv1, col_inv2, col_inv3 = st.columns(3)
            with col_inv1:
                invoice_prefix = st.text_input("Prefix *", 
                                            value=str(st.session_state.extracted_data.get('invoice_prefix', '')),
                                            key="prefix_input")

            with col_inv2:
                invoice_number = st.text_input("Number *", 
                                            value=str(st.session_state.extracted_data.get('invoice_number', '')),
                                            key="number_input")

            with col_inv3:
                invoice_year = st.text_input("Year *", 
                                            value=str(st.session_state.extracted_data.get('invoice_year', '')),
                                            key="year_input")
            
            # Auto-fill invoice date if extracted, otherwise today's date
            default_invoice_date = datetime.now()
            if 'invoice_date' in st.session_state.extracted_data:
                try:
                    # Parse the date string "30/09/2025" to datetime object
                    date_str = st.session_state.extracted_data['invoice_date']
                    day, month, year = map(int, date_str.split('/'))
                    default_invoice_date = datetime(year, month, day)
                    # st.success(f"✅ Auto-filled Invoice Date: {date_str}")
                except:
                    pass

            invoice_date = st.date_input("Invoice Date *", value=default_invoice_date, key="wc_invoice_date")

            # Auto-fill PO Number if extracted, otherwise blank
            default_po_number = st.session_state.extracted_data.get('po_number', '')
            po_number = st.text_input("PO Number", value=default_po_number)
            
            # GL Account dropdown with manual option
            st.write("GL Account *")
            gl_account_options = ["72022181", "10016000", "Other"]
            selected_gl_account = st.selectbox("Select GL Account", gl_account_options, key="gl_account_select")
            
            if selected_gl_account == "Other":
                gl_account = st.text_input("Enter GL Account manually", value="")
            else:
                gl_account = selected_gl_account
            
            # Cost Center - removed auto-mapping, just text input
            cost_center = st.text_input("Cost Center", value="")

        with col2:
            # Essential fields - always visible
            st.subheader("Financial Information")
            # Auto-fill gross value if extracted, otherwise 0.0
            default_gross_value = st.session_state.extracted_data.get('gross_value', 0.0)
            gross_value = st.number_input("Gross Value *", min_value=0.0, value=float(default_gross_value), step=1000.0)

            # Auto-fill VAT amount if extracted, otherwise 0.0
            default_vat_amount = st.session_state.extracted_data.get('vat_amount', 0.0)
            vat_amount = st.number_input("VAT Amount", min_value=0.0, value=float(default_vat_amount), step=100.0)
            
            # Calculate GL Amount
            gl_amount = gross_value - vat_amount
            
            # Display calculated values
            st.info(f"**Calculated GL Amount:** {gl_amount:,.2f}")
            
            # Advanced fields - conditionally visible
            if st.session_state.show_advanced:
                st.subheader("Advanced Settings")
                # Auto-fill currency if extracted, otherwise use default values
                default_currency = st.session_state.extracted_data.get('currency', 'LKR')
                currency = st.text_input("Currency", value=default_currency)
                
                # Auto-fill vendor code if extracted, otherwise use default values
                default_vendor_code = st.session_state.extracted_data.get('vendor_code', '400554')
                vendor_code = st.text_input("Vendor Code", value=default_vendor_code)
                
                # Auto-fill vendor if extracted, otherwise use default values
                default_vendor = st.session_state.extracted_data.get('vendor', 'DGC')
                vendor = st.text_input("Vendor", value=default_vendor)
                
                # Auto-fill source email if extracted, otherwise use default values
                default_source_email = st.session_state.extracted_data.get('source_email', 'chamithwi@masholdings.com')
                source_email = st.text_input("SOURCEEMAIL", value=default_source_email)
                
                # Auto-fill description if extracted, otherwise use default values
                default_description = st.session_state.extracted_data.get('description', 'CUSTOM DUTY')
                description = st.text_input("Description", value=default_description)
                
                # Auto-fill functional area if extracted, otherwise use default values
                default_functional_area = st.session_state.extracted_data.get('functional_area', 'Z019')
                functional_area = st.text_input("Functional Area", value=default_functional_area)
                
                # Auto-fill payment method if extracted, otherwise use default values
                default_payment_method = st.session_state.extracted_data.get('payment_method', 'P')
                payment_method = st.text_input("Payment Method", value=default_payment_method)
                
                # Auto-fill VAT GL account if extracted, otherwise use default values
                default_vat_gl_account = st.session_state.extracted_data.get('vat_gl_account', '17003030')
                vat_gl_account = st.text_input("VAT GL Account", value=default_vat_gl_account)
            else:
                # Set default values for hidden fields
                currency = st.session_state.extracted_data.get('currency', 'LKR')
                vendor_code = st.session_state.extracted_data.get('vendor_code', '400554')
                vendor = st.session_state.extracted_data.get('vendor', 'DGC')
                source_email = st.session_state.extracted_data.get('source_email', 'chamithwi@masholdings.com')
                description = st.session_state.extracted_data.get('description', 'CUSTOM DUTY')
                functional_area = st.session_state.extracted_data.get('functional_area', 'Z019')
                payment_method = st.session_state.extracted_data.get('payment_method', 'P')
                vat_gl_account = st.session_state.extracted_data.get('vat_gl_account', '17003030')
        
        # Format invoice number
        formatted_invoice_number = f"{invoice_prefix} {invoice_number} {invoice_year}"
        
        # Prepare data for PDF
        data = {
            'company_name': company_name,
            'currency': currency,
            'gross_value': gross_value,
            'invoice_number': formatted_invoice_number,
            'vendor_code': vendor_code,
            'invoice_date': invoice_date.strftime('%d/%m/%Y'),
            'vendor': vendor,
            'source_email': source_email,
            'po_number': po_number,
            'description': description,
            'cost_center': cost_center,
            'functional_area': functional_area,
            'payment_method': payment_method,
            'gl_account': gl_account,
            'vat_gl_account': vat_gl_account,
            'vat_amount': vat_amount,
            'gl_amount': gl_amount
        }

        # Generate filename based on company name
        company_code_map = {
            "BODYLINE PVT LTD": "BPL",
            "UNICHELA PVT LTD": "UPL", 
            "MAS CAPITAL PVT LTD": "MCPL"
        }

        company_code = company_code_map.get(company_name, "PRF")

        # Single download button that generates and downloads the PDF
        col_download1, col_download2 = st.columns(2)
        
        with col_download1:
            # Validate required fields
            missing_fields = []
            
            if not company_name:
                missing_fields.append("Company Name")
            if not invoice_prefix:
                missing_fields.append("Invoice Prefix")
            if not invoice_number:
                missing_fields.append("Invoice Number")
            if not invoice_year:
                missing_fields.append("Invoice Year")
            if gross_value <= 0:
                missing_fields.append("Gross Value (must be greater than 0)")
            if not gl_account:
                missing_fields.append("GL Account")
            
            if missing_fields:
                st.error(f"❌ Please fill in all required fields: {', '.join(missing_fields)}")
            
            # Generate filename
            filename_single = f"PRF-{invoice_number}-{company_code}.pdf" if invoice_number else "payment_request.pdf"
            
            # Create and download PDF directly
            try:
                pdf = create_payment_request_pdf(data)
                
                # Save to bytes
                pdf_bytes = pdf.output(dest='S').encode('latin-1')
                
                st.download_button(
                    label="Download Payment Request Only",
                    data=pdf_bytes,
                    file_name=filename_single,
                    mime="application/pdf",
                    type="primary",
                    disabled=bool(missing_fields)
                )
            except Exception as e:
                st.error(f"❌ Error generating PDF: {str(e)}")
        
        with col_download2:
            if st.session_state.get('cusdec_file') is not None:
                # Validate required fields
                missing_fields_merged = []
                
                if not company_name:
                    missing_fields_merged.append("Company Name")
                if not invoice_prefix:
                    missing_fields_merged.append("Invoice Prefix")
                if not invoice_number:
                    missing_fields_merged.append("Invoice Number")
                if not invoice_year:
                    missing_fields_merged.append("Invoice Year")
                if gross_value <= 0:
                    missing_fields_merged.append("Gross Value (must be greater than 0)")
                if not gl_account:
                    missing_fields_merged.append("GL Account")
                
                if missing_fields_merged:
                    st.error(f"❌ Please fill in all required fields: {', '.join(missing_fields_merged)}")
                
                # Generate filename
                filename_merged = f"PRF-{invoice_number}-{company_code}-merged.pdf" if invoice_number else "payment_request_merged.pdf"
                
                # Create and download merged PDF directly
                try:
                    # Create Payment Request PDF
                    pdf = create_payment_request_pdf(data)
                    pdf_bytes = pdf.output(dest='S').encode('latin-1')
                    
                    # Merge with CUSDEC PDF
                    merged_pdf_bytes = merge_pdfs(pdf_bytes, st.session_state.cusdec_file.getvalue())
                    
                    st.download_button(
                        label="Download Merged PDF with CUSDEC",
                        data=merged_pdf_bytes,
                        file_name=filename_merged,
                        mime="application/pdf",
                        type="secondary",
                        disabled=bool(missing_fields_merged)
                    )
                except Exception as e:
                    st.error(f"❌ Error generating merged PDF: {str(e)}")
            else:
                st.button("Download Merged PDF with CUSDEC", disabled=True, help="Please upload a CUSDEC PDF first")

        # Display preview of data
        st.subheader("Form Preview")
        
        # Show ALL fields in preview regardless of advanced toggle
        preview_fields = [
            'Company Name', 'Currency', 'Gross Value', 'Invoice Number', 
            'Vendor Code', 'Invoice Date', 'Vendor', 'SOURCEEMAIL',
            'PO Number', 'Description', 'Cost Center', 'Functional Area',
            'Payment Method', 'GL Account', 'VAT GL Account', 'VAT Amount',
            'GL Amount'
        ]
        
        preview_data = {
            'Field': preview_fields,
            'Value': [data.get(field.lower().replace(' ', '_'), '') for field in preview_fields]
        }
        st.table(preview_data)

    with tab2:
        st.subheader("PayReq Creation")
        
        # Upload template Excel file
        st.write("### Step 1: Upload PayReq Template")
        template_file = st.file_uploader("Upload PayReq Excel Template", type=["xlsx", "xls"])
        
        if template_file:
            st.success("Template uploaded successfully!")
        
        st.write("### Step 2: Enter Payment Details")
        
        # Basic information
        col1, col2 = st.columns(2)
        
        with col1:
            company_code_map = {
                "BODYLINE PVT LTD": "B050",
                "UNICHELA PVT LTD": "A050", 
                "MAS CAPITAL PVT LTD": "MCAP"
            }
            selected_company = st.selectbox("COMPANY", list(company_code_map.keys()), index=1)
            company_code = company_code_map[selected_company]
            
            vendor_code = st.text_input("VENDOR CODE", "0000400554")
            vendor_name = st.text_input("VENDOR NAME", "Director General of Customs")

        with col2:
            payment_mode = st.selectbox("PAYMENT MODE", ["S", "C", "T"], index=0)
            currency = st.selectbox("CURRENCY", ["LKR", "USD"], index=0)
            description = st.text_input("DESCRIPTION", "VAT Claimable")
        
        # Initialize session state for invoice data
        if 'payreq_invoice_data' not in st.session_state:
            st.session_state.payreq_invoice_data = pd.DataFrame({
                'INV. DATE': [],
                'Office Code': [],
                'Year': [],
                'Serial': [],
                'CUSDEC': [],
                'AMOUNT': [],
                'GL A/C': [],
                'F A': [],
                'CUSDEC_FILE': []  # Store the uploaded file for reference
            })
        
        st.write("### Step 3: Upload CUSDEC PDFs and Auto-fill Data")
        
        # Combined CUSDEC PDF uploader for both auto-fill and merging
        uploaded_cusdecs = st.file_uploader(
            "Upload CUSDEC PDFs (for auto-fill and merging)", 
            type="pdf", 
            accept_multiple_files=True,
            key="payreq_cusdec_uploader"
        )
        
        # Store uploaded files in session state for merging
        if uploaded_cusdecs:
            st.session_state.payreq_cusdec_files = uploaded_cusdecs
 
        # Process each uploaded CUSDEC
        if uploaded_cusdecs:
            for i, cusdec_file in enumerate(uploaded_cusdecs):
                if f"processed_{cusdec_file.name}" not in st.session_state:
                    st.write(f"**Processing {cusdec_file.name}...**")
                    
                    # Extract data from CUSDEC PDF
                    with st.spinner(f"Extracting data from {cusdec_file.name}..."):
                        extracted_data = process_cusdec_pdf(cusdec_file)
                        
                        # Add to table if extraction successful
                        if any(key in extracted_data for key in ['company_name', 'invoice_number', 'gross_value', 'vat_amount']):
                            # Get invoice components
                            inv_prefix = extracted_data.get('invoice_prefix', 'S')
                            inv_number = extracted_data.get('invoice_number', '')
                            inv_year = extracted_data.get('invoice_year', '2025')
                            inv_date = extracted_data.get('invoice_date', datetime.now().strftime('%d/%m/%Y'))
                            gross_value = extracted_data.get('gross_value', 0.0)
                            vat_amount = extracted_data.get('vat_amount', 0.0)  # Get VAT amount
                            
                            # Validate that Gross Value equals VAT Amount (only VAT should be included)
                            if gross_value == vat_amount:
                                new_row = {
                                    'INV. DATE': inv_date,
                                    'Office Code': extracted_data.get('office_code', 'CBBI1'),
                                    'Year': inv_year,
                                    'Serial': inv_prefix,
                                    'CUSDEC': inv_number,
                                    'AMOUNT': vat_amount,  # Use VAT amount only
                                    'GL A/C': '17003030',
                                    'F A': 'Z016',
                                    'CUSDEC_FILE': cusdec_file.name
                                }
                                
                                # Add to dataframe
                                new_df = pd.DataFrame([new_row])
                                st.session_state.payreq_invoice_data = pd.concat(
                                    [st.session_state.payreq_invoice_data, new_df], 
                                    ignore_index=True
                                )
                                
                                # Mark as processed
                                st.session_state[f"processed_{cusdec_file.name}"] = True
                                st.success(f"✅ Added data from {cusdec_file.name}")
                            else:
                                st.error(f"❌ Rejected {cusdec_file.name}: Gross Value ({gross_value:,.2f}) ≠ VAT Amount ({vat_amount:,.2f}) - Only VAT claims allowed")
                        else:
                            st.warning(f"⚠️ Could not extract sufficient data from {cusdec_file.name}")



        st.write("### Step 4: Review and Edit Invoice Data")
        
        if not st.session_state.payreq_invoice_data.empty:
            # Editable dataframe
            edited_df = st.data_editor(
                st.session_state.payreq_invoice_data,
                num_rows="dynamic",
                use_container_width=True,
                key="payreq_data_editor"
            )
            
            # Update session state
            st.session_state.payreq_invoice_data = edited_df
            
            # Calculate totals
            total_amount = edited_df['AMOUNT'].sum()
            st.info(f"**Total Amount: {total_amount:,.2f}**")
            
            # Manual add button
            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.button("Add Manual Entry"):
                    new_manual_row = pd.DataFrame({
                        'INV. DATE': [datetime.now().strftime('%d/%m/%Y')],
                        'Office Code': ['CBBI1'],
                        'Year': ['2025'],
                        'Serial': ['S'],
                        'CUSDEC': [''],
                        'AMOUNT': [0.00],
                        'GL A/C': ['17003030'],
                        'F A': ['Z016'],
                        'CUSDEC_FILE': ['']
                    })
                    st.session_state.payreq_invoice_data = pd.concat(
                        [st.session_state.payreq_invoice_data, new_manual_row], 
                        ignore_index=True
                    )
                    st.rerun()
            
            with col_btn2:
                if st.button("Clear All Data"):
                    st.session_state.payreq_invoice_data = pd.DataFrame({
                        'INV. DATE': [],
                        'Office Code': [],
                        'Year': [],
                        'Serial': [],
                        'CUSDEC': [],
                        'AMOUNT': [],
                        'GL A/C': [],
                        'F A': [],
                        'CUSDEC_FILE': []
                    })
                    # Clear processed flags
                    for key in list(st.session_state.keys()):
                        if key.startswith('processed_'):
                            del st.session_state[key]
                    st.rerun()
        else:
            st.info("No invoice data yet. Upload CUSDEC PDFs to auto-fill or add manual entries.")
            
            # Add first manual entry if empty
            if st.button("Start with Manual Entry"):
                st.session_state.payreq_invoice_data = pd.DataFrame({
                    'INV. DATE': [datetime.now().strftime('%d/%m/%Y')],
                    'Office Code': ['CBBI1'],
                    'Year': ['2025'],
                    'Serial': ['S'],
                    'CUSDEC': [''],
                    'AMOUNT': [0.00],
                    'GL A/C': ['17003030'],
                    'F A': ['Z016'],
                    'CUSDEC_FILE': ['']
                })
                st.rerun()
        
        # Generate final files
        # Generate final files
        if not st.session_state.payreq_invoice_data.empty:
            st.write("### Step 5: Generate Files")
            
            col_gen1, col_gen2 = st.columns(2)
            
            def prepare_output_data():
                output_data = []
                for _, row in st.session_state.payreq_invoice_data.iterrows():
                    invoice_no = f"{row['Serial']}{row['CUSDEC']}{row['Year']}" if row['CUSDEC'] else f"{row['Serial']}{row['Year']}"
                    cost_center_map = {
                        "BODYLINE PVT LTD": "B050COMN01",
                        "UNICHELA PVT LTD": "A050COMN01", 
                        "MAS CAPITAL PVT LTD": "MCAPCOMN01"
                    }
                    cost_center = cost_center_map[selected_company]
                    assignment = selected_company.upper().replace(' PVT LTD', '')

                    output_data.append({
                        'INV. DATE': row['INV. DATE'],
                        'INVOICE NO': invoice_no,
                        'VENDOR': vendor_code,
                        'GL A/C': row['GL A/C'],
                        'TEXT': description,
                        'COST CENTER': cost_center,
                        'ASSIGNEMENT': assignment,
                        'F A': row['F A'],
                        'INT.ORDER': '',
                        'N O F': '',
                        'Plant': company_code,
                        'AMOUNT': f"{row['AMOUNT']:,.2f}",
                        'VAT (11%)': '0.00',
                        'VAT in LKR': '0.00',
                        'NBT in LKR': '0.00',
                        'TOTAL': f"{row['AMOUNT']:,.2f}",
                        'Office Code': row['Office Code'],
                        'Year': row['Year'],
                        'Serial': row['Serial'],
                        'CUSDEC': row['CUSDEC']
                    })

                total_amount = st.session_state.payreq_invoice_data['AMOUNT'].sum()
                if output_data:
                    output_data.append({
                        'INV. DATE': '', 'INVOICE NO': '', 'VENDOR': '', 'GL A/C': '',
                        'TEXT': '', 'COST CENTER': '', 'ASSIGNEMENT': '', 'F A': '',
                        'INT.ORDER': '', 'N O F': '', 'Plant': '', 'AMOUNT': f"{total_amount:,.2f}",
                        'VAT (11%)': '0.00', 'VAT in LKR': '0.00', 'NBT in LKR': '0.00',
                        'TOTAL': f"{total_amount:,.2f}", 'Office Code': '', 'Year': '',
                        'Serial': '', 'CUSDEC': ''
                    })

                df = pd.DataFrame(output_data)

                # Ensure the correct column order
                column_order = [
                    'INV. DATE', 'INVOICE NO', 'VENDOR', 'GL A/C', 'TEXT',
                    'COST CENTER', 'ASSIGNEMENT', 'F A', 'INT.ORDER', 'N O F', 'Plant',
                    'AMOUNT', 'VAT (11%)', 'VAT in LKR', 'NBT in LKR', 'TOTAL',
                    'Office Code', 'Year', 'Serial', 'CUSDEC'
                ]
                df = df[column_order]
                return df

            # ---- Excel Only ----
            with col_gen1:
                if st.button("Generate PayReq Excel Only", type="primary"):
                    try:
                        output_df = prepare_output_data()
                        output = BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            output_df.to_excel(writer, index=False, sheet_name='Payment Requisition')

                        excel_data = output.getvalue()
                        filename = f"PayReq_{company_code}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"

                        st.download_button(
                            label="Download PayReq Excel",
                            data=excel_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                        st.success("PayReq Excel generated successfully!")
                    except Exception as e:
                        st.error(f"Error generating Excel: {str(e)}")

            # ---- Excel + PDF ----
            with col_gen2:
                if st.button("Generate with Merged CUSDEC PDFs", type="secondary"):
                    if 'payreq_cusdec_files' in st.session_state and st.session_state.payreq_cusdec_files:
                        try:
                            # Generate Excel first
                            output_df = prepare_output_data()
                            excel_output = BytesIO()
                            with pd.ExcelWriter(excel_output, engine='openpyxl') as writer:
                                output_df.to_excel(writer, index=False, sheet_name='Payment Requisition')
                            excel_data = excel_output.getvalue()

                            # Merge CUSDEC PDFs
                            pdf_merger = PyPDF2.PdfMerger()
                            for cusdec_file in st.session_state.payreq_cusdec_files:
                                pdf_merger.append(BytesIO(cusdec_file.getvalue()))

                            combined_pdf = BytesIO()
                            pdf_merger.write(combined_pdf)
                            pdf_merger.close()
                            combined_pdf_data = combined_pdf.getvalue()

                            # Download buttons
                            col_dl1, col_dl2 = st.columns(2)
                            with col_dl1:
                                excel_filename = f"PayReq_{company_code}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
                                st.download_button(
                                    label="Download Excel File",
                                    data=excel_data,
                                    file_name=excel_filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
                            with col_dl2:
                                pdf_filename = f"Combined_CUSDEC_{company_code}_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf"
                                st.download_button(
                                    label="Download Combined CUSDEC PDF",
                                    data=combined_pdf_data,
                                    file_name=pdf_filename,
                                    mime="application/pdf"
                                )

                            st.success("Both Excel and combined PDF generated successfully!")
                        except Exception as e:
                            st.error(f"Error generating files: {str(e)}")
                    else:
                        st.warning("No CUSDEC PDFs uploaded for merging")



if __name__ == "__main__":
    main()



