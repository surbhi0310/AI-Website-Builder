"""
fix_excel_and_mentor.py - Fixes Excel formatting and replaces any remaining old mentor names
1. Replaces any remaining instances of 'Kalpana Dwivedi' or 'PJ-IVD-2026' with 'Mr. Dinesh Sereno' and 'P053'
2. Regenerates Step 1 Document and Complete Code Document
3. Professional Excel styling: auto column width, number formatting, no ###, visible gridlines
4. Updates all files in LMS_Submission_Package and converts to PDF
"""

import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import win32com.client
import shutil

BASE_DIR = Path(r"C:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")
DOCS_DIR = BASE_DIR / "docs"
PDF_DIR = DOCS_DIR / "pdf_submissions"
DATA_DIR = BASE_DIR / "data"
PACKAGE_DIR = BASE_DIR / "LMS_Submission_Package"

# 1. Update test files & generator files
def update_source_comments():
    print("--- Updating code comments and Step 1 generator ---")
    files_to_check = [
        BASE_DIR / "generate_step1_doc.py",
        BASE_DIR / "tests" / "test_data.py",
        BASE_DIR / "tests" / "test_analysis.py",
        BASE_DIR / "tests" / "test_app.py",
        BASE_DIR / "BUILD_PLAN.md",
        BASE_DIR / "README.md"
    ]
    for fp in files_to_check:
        if fp.exists():
            text = fp.read_text(encoding="utf-8")
            text = text.replace("Kalpana Dwivedi", "Mr. Dinesh Sereno")
            text = text.replace("PJ-IVD-2026", "P053")
            text = text.replace("ABES_OND26_MCA Batch 02", "ABES_OND26_MCA2ndYear_Data Analytics_Batch 01")
            fp.write_text(text, encoding="utf-8")
            print(f"Updated: {fp.name}")

    # Re-run generate_step1_doc.py
    os.system(f'cd /d "{BASE_DIR}" && python generate_step1_doc.py')
    # Copy generated doc into docs/01_Project_Understanding.docx
    shutil.copy2(BASE_DIR / "STEP_01_Project_Understanding.docx", DOCS_DIR / "01_Project_Understanding.docx")
    print("Regenerated docs/01_Project_Understanding.docx with Mr. Dinesh Sereno and P053")

# 2. Format Excel files to perfection
def format_excel_file(file_path):
    print(f"Formatting Excel: {file_path.name}...")
    wb = openpyxl.load_workbook(file_path)
    
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")
    data_font = Font(name="Calibri", size=10, color="1F2937")
    
    thin_border_side = Side(style="thin", color="E2E8F0")
    thin_border = Border(left=thin_border_side, right=thin_border_side, top=thin_border_side, bottom=thin_border_side)
    
    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        # Enable grid lines
        ws.views.sheetView[0].showGridLines = True
        
        # Check max row/col
        max_r = ws.max_row
        max_c = ws.max_column
        if max_r < 1 or max_c < 1:
            continue
            
        # Style Header Row (row 1)
        for col_idx in range(1, max_c + 1):
            cell = ws.cell(row=1, column=col_idx)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = thin_border
        ws.row_dimensions[1].height = 28
        
        # Style Data Rows
        for r_idx in range(2, max_r + 1):
            ws.row_dimensions[r_idx].height = 20
            is_zebra = (r_idx % 2 == 1)
            for c_idx in range(1, max_c + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                cell.font = data_font
                cell.border = thin_border
                if is_zebra and cell.fill.fill_type is None:
                    cell.fill = zebra_fill
                
                # Number formatting
                val = cell.value
                if isinstance(val, (int, float)):
                    # Format as currency or standard decimal
                    col_header = str(ws.cell(row=1, column=c_idx).value).lower()
                    if any(w in col_header for w in ['amount', 'price', 'total', 'subtotal', 'variance', 'leakage', 'spend', 'tax']):
                        cell.number_format = '#,##0.00'
                    elif isinstance(val, int) or (isinstance(val, float) and val.is_integer()):
                        cell.number_format = '#,##0'
                elif val is not None and ('date' in str(ws.cell(row=1, column=c_idx).value).lower()):
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    
        # Auto-fit column widths with ample padding (eliminates ### completely)
        for col in ws.columns:
            max_len = 0
            col_letter = get_column_letter(col[0].column)
            for cell in col:
                val_str = ""
                if cell.value is not None:
                    if isinstance(cell.value, float):
                        val_str = f"{cell.value:,.2f}"
                    elif isinstance(cell.value, int):
                        val_str = f"{cell.value:,}"
                    else:
                        val_str = str(cell.value)
                if len(val_str) > max_len:
                    max_len = len(val_str)
            # Add padding
            ws.column_dimensions[col_letter].width = max(max_len + 4, 14)
            
    wb.save(file_path)
    print(f"  [OK] Successfully formatted and saved: {file_path.name}")

# Execute Excel Formatter
def format_all_excels():
    excel_files = [
        DATA_DIR / "invoice_data.xlsx",
        DOCS_DIR / "04_Traceability_Matrix.xlsx",
        DOCS_DIR / "08_Database_Design.xlsx",
        PACKAGE_DIR / "2500320140152_SurbhiSharma_Project_Database.xlsx",
        PACKAGE_DIR / "2500320140152_SurbhiSharma_04_Traceability_Matrix.xlsx",
        PACKAGE_DIR / "2500320140152_SurbhiSharma_08_Database_Design.xlsx"
    ]
    for ef in excel_files:
        if ef.exists():
            format_excel_file(ef)

# 3. Regenerate Complete Code docx
def regenerate_code_and_reports():
    print("--- Regenerating Complete Code Document ---")
    os.system(f'cd /d "{BASE_DIR}" && python generate_lms_special_reports.py')

# 4. Convert DOCX to PDF via Word COM
def convert_updated_docs_to_pdf():
    print("--- Converting updated DOCX to PDF via Word COM ---")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    docs_to_convert = [
        (DOCS_DIR / "01_Project_Understanding.docx", PDF_DIR / "01_Project_Understanding.pdf", PACKAGE_DIR / "2500320140152_SurbhiSharma_01_Project_Understanding.pdf"),
        (DOCS_DIR / "Complete_Code.docx", PDF_DIR / "Complete_Code.pdf", PACKAGE_DIR / "2500320140152_SurbhiSharma_Source_Code.pdf"),
        (DOCS_DIR / "GitHub_Report.docx", PDF_DIR / "GitHub_Report.pdf", PACKAGE_DIR / "2500320140152_SurbhiSharma_GitHub_and_Streamlit_Details.pdf")
    ]
    
    try:
        for doc_in, pdf_out, pkg_out in docs_to_convert:
            if doc_in.exists():
                print(f"Converting: {doc_in.name} -> {pdf_out.name}...")
                doc = word.Documents.Open(str(doc_in.resolve()))
                doc.SaveAs(str(pdf_out.resolve()), FileFormat=17)
                doc.Close()
                print(f"  [OK] Saved: {pdf_out.name}")
                shutil.copy2(pdf_out, pkg_out)
                print(f"  [OK] Copied to LMS package: {pkg_out.name}")
    finally:
        word.Quit()

if __name__ == "__main__":
    update_source_comments()
    regenerate_code_and_reports()
    format_all_excels()
    convert_updated_docs_to_pdf()
    print("\nAll Excel formatting and mentor updates complete!")
