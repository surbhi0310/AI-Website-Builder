"""
build_rich_github_streamlit_report.py - Generates comprehensive GitHub and Streamlit Details Report
Features prominent live deployment URL, GitHub repository URL, embedded screenshots, and test matrix.
Project: P053 | Surbhi Sharma (2500320140152) | Mentor: Mr. Dinesh Sereno
"""

import os
from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

BASE_DIR = Path(r"C:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")
DOCS_DIR = BASE_DIR / "docs"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(f'<w:tcMar {nsdecls("w")}><w:top w:w="{top}" w:type="dxa"/><w:bottom w:w="{bottom}" w:type="dxa"/><w:left w:w="{left}" w:type="dxa"/><w:right w:w="{right}" w:type="dxa"/></w:tcMar>')
    tcPr.append(tcMar)

def create_rich_github_streamlit_report():
    print("Building rich GitHub & Streamlit Details Document...")
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.8)
        s.bottom_margin = Inches(0.8)
        s.left_margin = Inches(0.8)
        s.right_margin = Inches(0.8)

    p_meta = doc.add_paragraph()
    r_meta = p_meta.add_run("ABES ENGINEERING COLLEGE | DEPARTMENT OF MCA (DATA ANALYTICS)\nProject ID: P053 | Candidate: Surbhi Sharma (2500320140152) | Mentor: Mr. Dinesh Sereno")
    r_meta.font.size = Pt(8.5)
    r_meta.font.bold = True
    r_meta.font.color.rgb = RGBColor.from_string("64748B")
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    p_title = doc.add_paragraph()
    r_t = p_title.add_run("GITHUB REPOSITORY & LIVE STREAMLIT CLOUD REPORT")
    r_t.font.name = "Calibri"
    r_t.font.size = Pt(20)
    r_t.bold = True
    r_t.font.color.rgb = RGBColor.from_string("1E3A8A")

    p_sub = doc.add_paragraph()
    r_sub = p_sub.add_run("Official Cloud Hosting, Live Application Deployment Verification & Code Repository Audit")
    r_sub.font.size = Pt(11)
    r_sub.font.bold = True
    r_sub.font.color.rgb = RGBColor.from_string("475569")
    p_sub.paragraph_format.space_after = Pt(12)

    # Big Banner with Live Links
    p_banner = doc.add_paragraph()
    r_b1 = p_banner.add_run("🌐 LIVE DEPLOYED APP (STREAMLIT COMMUNITY CLOUD):\n")
    r_b1.bold = True
    r_b1.font.size = Pt(12)
    r_b1.font.color.rgb = RGBColor.from_string("16A34A")
    r_b2 = p_banner.add_run("https://invoicehclproject.streamlit.app\n\n")
    r_b2.bold = True
    r_b2.font.size = Pt(13)
    r_b2.font.color.rgb = RGBColor.from_string("1E3A8A")

    r_b3 = p_banner.add_run("💻 OFFICIAL GITHUB CODE REPOSITORY:\n")
    r_b3.bold = True
    r_b3.font.size = Pt(12)
    r_b3.font.color.rgb = RGBColor.from_string("2563EB")
    r_b4 = p_banner.add_run("https://github.com/surbhi0310/Career-Shaper-HCL-Project.git\n")
    r_b4.bold = True
    r_b4.font.size = Pt(13)
    r_b4.font.color.rgb = RGBColor.from_string("1E3A8A")

    p_banner.paragraph_format.space_after = Pt(12)
    pPr = p_banner._p.get_or_add_pPr()
    pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:left w:val="single" w:sz="24" w:space="8" w:color="1E3A8A"/><w:top w:val="single" w:sz="6" w:space="4" w:color="E2E8F0"/><w:bottom w:val="single" w:sz="6" w:space="4" w:color="E2E8F0"/><w:right w:val="single" w:sz="6" w:space="4" w:color="E2E8F0"/></w:pBdr>')
    pPr.append(pBdr)

    # Section 1: Deployment & System Metadata
    doc.add_heading("1. Project Hosting & Infrastructure Metadata", level=1)
    t_meta = doc.add_table(rows=7, cols=2)
    meta_rows = [
        ("Project ID & Title", "P053: Invoice Data Validation & Discrepancy Analytics Platform"),
        ("Live Deployed URL", "https://invoicehclproject.streamlit.app (Global 24/7 Access)"),
        ("GitHub Repository", "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git"),
        ("Candidate / Learner", "Surbhi Sharma (University Roll No: 2500320140152)"),
        ("Batch & Specialization", "ABES_OND26_MCA2ndYear_Data Analytics_Batch 01 (Data Analytics)"),
        ("Project Mentor", "Mr. Dinesh Sereno (Assistant Professor, MCA, ABES Engineering College)"),
        ("Hosting Architecture", "Streamlit Community Cloud (Linux Container, Python 3.10+, Headless)")
    ]
    for idx, (k, v) in enumerate(meta_rows):
        r = t_meta.rows[idx]
        r.cells[0].text = k
        r.cells[1].text = v
        r.cells[0].paragraphs[0].runs[0].bold = True
        set_cell_background(r.cells[0], "F1F5F9")
        set_cell_margins(r.cells[0], 60, 60, 100, 100)
        set_cell_margins(r.cells[1], 60, 60, 100, 100)
        r.cells[0].width = Inches(2.2)
        r.cells[1].width = Inches(4.8)

    # Section 2: Screenshot of Live Application
    doc.add_heading("2. Live Streamlit Cloud Application Screenshot", level=1)
    doc.add_paragraph("The application is live and operational on Streamlit Community Cloud. Below is the verified browser screenshot captured from https://invoicehclproject.streamlit.app showing executive KPIs and discrepancy analytics:")
    
    img_path = r"C:\Users\Geetanjali\.gemini\antigravity\brain\8efacaf0-2560-4cb3-84eb-6cda039164dd\browser_evidence\eab52b91-7ce3-4f7c-88e4-21a1284066bf\final_step_20.png"
    if os.path.exists(img_path):
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_i = p_img.add_run()
        r_i.add_picture(img_path, width=Inches(6.6))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_cap = p_cap.add_run("Figure 1: Live Dashboard on Streamlit Community Cloud showing Total Spend (₹7.09 Cr), Invoices (304), Flagged Discrepancies (69), and Prevented Leakage (₹1.84 Cr).")
        r_cap.font.size = Pt(8.5)
        r_cap.font.italic = True
        r_cap.font.color.rgb = RGBColor.from_string("64748B")

    # Section 3: Incognito Verification Matrix
    doc.add_heading("3. Incognito Mode Multi-Page Verification Matrix", level=1)
    doc.add_paragraph("Each multi-page application route was independently verified in an incognito session with zero local caching:")
    checks = [
        ("Page 0 (Root Portal)", "app.py", "PASSED", "Renders 4 KPI cards with Indian Rupee formatting (₹7.09 Cr) and discrepancy breakdown chart."),
        ("Page 1 (Overview KPIs)", "pages/1_Overview_KPIs.py", "PASSED", "Renders time-series spend trend chart, category distribution, and top 10 vendor spend bars."),
        ("Page 2 (Invoice Validation)", "pages/2_Invoice_Validation.py", "PASSED", "Interactive invoice auditor flags arithmetic errors, non-standard tax slabs, and provides audit export."),
        ("Page 3 (PO Reconciliation)", "pages/3_PO_Reconciliation.py", "PASSED", "Executes 2-way budget matching between invoices and POs with defensive column reconciliation."),
        ("Page 4 (Vendor Risk)", "pages/4_Vendor_Risk.py", "PASSED", "Calculates composite Vendor Risk Index (0-100), isolates 3 rogue suppliers, and displays risk tiers."),
        ("Page 5 (Statistical EDA)", "pages/5_Statistical_EDA.py", "PASSED", "Plots Benford's Law theoretical vs observed distributions and IQR procurement outlier boxplots."),
        ("Page 6 (Data Quality)", "pages/6_Data_Quality_&_Export.py", "PASSED", "Live SQLite database table browser, missing value profiling, and CSV/Excel download hub.")
    ]
    t_chk = doc.add_table(rows=len(checks) + 1, cols=4)
    h_chk = ["Screen / Page Route", "Module Source", "Incognito Check", "Functional Capability Verified"]
    w_chk = [Inches(1.8), Inches(1.8), Inches(1.1), Inches(2.3)]
    for idx, c in enumerate(t_chk.rows[0].cells):
        set_cell_background(c, "1E3A8A")
        c.text = h_chk[idx]
        c.paragraphs[0].runs[0].bold = True
        c.paragraphs[0].runs[0].font.color.rgb = RGBColor.from_string("FFFFFF")
        c.width = w_chk[idx]
    for idx, row in enumerate(checks):
        r_c = t_chk.rows[idx + 1]
        for c_idx, val in enumerate(row):
            r_c.cells[c_idx].text = val
            r_c.cells[c_idx].width = w_chk[c_idx]
            set_cell_margins(r_c.cells[c_idx], 60, 60, 80, 80)
            if idx % 2 == 1:
                set_cell_background(r_c.cells[c_idx], "F8FAFC")

    # Section 4: Repository File Hierarchy
    doc.add_heading("4. GitHub Repository Manifest & Structure", level=1)
    tree_text = """
surbhi0310 / Career-Shaper-HCL-Project (Branch: main)
├── .streamlit/
│   └── config.toml (Theme, colors, and layout)
├── data/
│   ├── app.db (SQLite database - 5 normalized tables + analytical view)
│   ├── invoice_data.xlsx (Master multi-tab formatted workbook)
│   ├── vendors.csv (30 records)
│   ├── purchase_orders.csv (150 records)
│   ├── invoices.csv (304 records)
│   ├── line_items.csv (1,056 records)
│   └── discrepancies.csv (69 records)
├── docs/
│   ├── 01_Project_Understanding.docx (and .pdf)
│   ├── 02_HMW_Statements.docx (and .pdf)
│   ├── 03_Needs_Analysis.docx (and .pdf)
│   ├── 04_Traceability_Matrix.xlsx (and .docx, .pdf)
│   ├── 05_BRD.docx (and .pdf)
│   ├── 06_PRD.docx (and .pdf)
│   ├── 07_TRD.docx (and .pdf)
│   ├── 08_Database_Design.xlsx (and .docx, .pdf)
│   ├── 09_ER_Diagram.docx (and .pdf)
│   ├── 10_Dataset_Generation.docx (and .pdf)
│   ├── 11_User_Stories.docx (and .pdf)
│   ├── 12_UI_Framework.docx (and .pdf)
│   ├── 13_Test_Cases.docx (and .pdf)
│   ├── 14_BUILD_PLAN.docx (and .pdf)
│   ├── 15_Application_Code.docx (and .pdf)
│   ├── 16_Deployment_Guide.docx (and .pdf)
│   ├── 17_Testing_Report.docx (and .pdf)
│   ├── 18_Deployment_Submission.docx (and .pdf)
│   ├── PJ_IVD_2026_Project_Document.docx (and .pdf)
│   ├── PJ_IVD_2026_Presentation.pptx (and .pdf)
│   ├── er_diagram.png (High-res 300 DPI ER diagram)
│   └── pdf_submissions/ (All individual PDF deliverables)
├── pages/
│   ├── 1_Overview_KPIs.py
│   ├── 2_Invoice_Validation.py
│   ├── 3_PO_Reconciliation.py
│   ├── 4_Vendor_Risk.py
│   ├── 5_Statistical_EDA.py
│   └── 6_Data_Quality_&_Export.py
├── src/
│   ├── __init__.py
│   ├── database.py (Data access layer & cached queries)
│   ├── validation.py (Vectorized discrepancy detection)
│   ├── analytics.py (Benford's law & VRI multi-factor scoring)
│   ├── ui.py (Tailwind styling & INR currency formatting)
│   └── generate_data.py (Synthetic data generation engine)
├── tests/
│   ├── __init__.py
│   ├── test_data.py (4 passed)
│   ├── test_analysis.py (5 passed)
│   └── test_app.py (4 passed)
├── app.py (Root Streamlit multi-page entry point)
├── conftest.py (Pytest configuration)
├── requirements.txt (Pinned dependencies)
├── BUILD_PLAN.md (Execution roadmap)
└── README.md (Comprehensive project documentation & badges)
"""
    p_code = doc.add_paragraph()
    r_c = p_code.add_run(tree_text.strip())
    r_c.font.name = "Consolas"
    r_c.font.size = Pt(8.0)

    # Section 5: Automated Testing Summary
    doc.add_heading("5. Automated Pytest Verification Summary", level=1)
    doc.add_paragraph("All 13 automated unit and integration tests executed with 100% pass rate in pytest before deployment:")
    doc.add_paragraph("• tests/test_data.py: 4/4 PASSED (DB presence, row count thresholds, primary key uniqueness, foreign key check)\n"
                      "• tests/test_analysis.py: 5/5 PASSED (Arithmetic check, duplicate detection, tax slab validation, PO reconciliation, VRI bounds)\n"
                      "• tests/test_app.py: 4/4 PASSED (Summary metrics aggregation, Benford distribution, IQR outlier detection, 32-col view loading)")

    out_docx = DOCS_DIR / "GitHub_Report.docx"
    doc.save(str(out_docx))
    print(f"[OK] Saved rich GitHub & Streamlit report: {out_docx}")
    return out_docx

if __name__ == "__main__":
    create_rich_github_streamlit_report()
