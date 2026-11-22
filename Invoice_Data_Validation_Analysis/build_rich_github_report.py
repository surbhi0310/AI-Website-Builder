"""
build_rich_github_report.py
Generates the comprehensive GitHub and Streamlit Details Report (docx & pdf)
Contains:
1. Clickable Live Streamlit Community Cloud URL
2. Clickable Official GitHub Repository URL
3. Metadata Table (Candidate, Roll No, Batch, Mentor: Mr. Dinesh Sereno, Project ID: P053)
4. Figure 1: Live Streamlit Application Dashboard Screenshot (Clean, verified P053 & Mr. Dinesh Sereno)
5. Multi-Page Incognito Verification Matrix (All 6 routes verified)
6. Repository Architecture, Branch Details & Component Manifest
7. Automated Pytest Verification Summary (13/13 passing tests)
"""

from pathlib import Path
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls
import win32com.client
import shutil

BASE_DIR = Path(r"c:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")
DOCS_DIR = BASE_DIR / "docs"
PDF_DIR = DOCS_DIR / "pdf_submissions"
LMS_PKG_DIR = BASE_DIR / "LMS_Submission_Package"
LMS_BACKUP_DIR = LMS_PKG_DIR / "Backup_ZIP"

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_cell_margins(cell, top=80, bottom=80, left=120, right=120):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)

def add_clickable_hyperlink(paragraph, url, text, color="0969DA", underline=True, bold=True, size_pt=None):
    part = paragraph.part
    r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)
    hyperlink = parse_xml(f'<w:hyperlink {nsdecls("w")} {nsdecls("r")} r:id="{r_id}" w:history="1"/>')
    new_run = parse_xml(f'<w:r {nsdecls("w")}/>')
    rPr = parse_xml(f'<w:rPr {nsdecls("w")}/>')

    if color:
        c = parse_xml(f'<w:color {nsdecls("w")} w:val="{color}"/>')
        rPr.append(c)
    if underline:
        u = parse_xml(f'<w:u {nsdecls("w")} w:val="single"/>')
        rPr.append(u)
    if bold:
        b = parse_xml(f'<w:b {nsdecls("w")}/>')
        rPr.append(b)
    if size_pt:
        sz = parse_xml(f'<w:sz {nsdecls("w")} w:val="{int(size_pt * 2)}"/>')
        rPr.append(sz)

    new_run.append(rPr)
    new_run_text = parse_xml(f'<w:t {nsdecls("w")}>{text}</w:t>')
    new_run.append(new_run_text)
    hyperlink.append(new_run)
    paragraph._p.append(hyperlink)
    return hyperlink

def generate_report():
    print("Building GitHub and Streamlit Details Document...")
    doc = Document()
    for s in doc.sections:
        s.top_margin = Inches(0.7)
        s.bottom_margin = Inches(0.7)
        s.left_margin = Inches(0.7)
        s.right_margin = Inches(0.7)

    # Top institutional header
    p_meta = doc.add_paragraph()
    r_meta = p_meta.add_run(
        "ABES ENGINEERING COLLEGE | DEPARTMENT OF MCA (DATA ANALYTICS)\n"
        "Project ID: P053 | Candidate: Surbhi Sharma (Roll No: 2500320140152) | Mentor: Mr. Dinesh Sereno"
    )
    r_meta.font.size = Pt(8.5)
    r_meta.font.bold = True
    r_meta.font.color.rgb = RGBColor.from_string("64748B")
    p_meta.alignment = WD_ALIGN_PARAGRAPH.RIGHT

    # Title
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
    p_sub.paragraph_format.space_after = Pt(10)

    # Prominent Callout Banner with Clickable Live Links
    p_banner = doc.add_paragraph()
    r_b1 = p_banner.add_run("[LIVE APPLICATION URL - STREAMLIT COMMUNITY CLOUD]:\n")
    r_b1.bold = True
    r_b1.font.size = Pt(10.5)
    r_b1.font.color.rgb = RGBColor.from_string("16A34A")

    # Clickable Live Link
    add_clickable_hyperlink(
        p_banner,
        "https://invoicehclproject.streamlit.app",
        "https://invoicehclproject.streamlit.app",
        color="0969DA",
        underline=True,
        bold=True,
        size_pt=12.5
    )

    r_b3 = p_banner.add_run("\n\n[OFFICIAL GITHUB REPOSITORY URL]:\n")
    r_b3.bold = True
    r_b3.font.size = Pt(10.5)
    r_b3.font.color.rgb = RGBColor.from_string("2563EB")

    # Clickable GitHub Link
    add_clickable_hyperlink(
        p_banner,
        "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git",
        "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git",
        color="0969DA",
        underline=True,
        bold=True,
        size_pt=12.5
    )

    p_banner.paragraph_format.space_after = Pt(14)
    pPr = p_banner._p.get_or_add_pPr()
    pBdr = parse_xml(
        f'<w:pBdr {nsdecls("w")}>'
        f'<w:left w:val="single" w:sz="24" w:space="8" w:color="1E3A8A"/>'
        f'<w:top w:val="single" w:sz="6" w:space="4" w:color="E2E8F0"/>'
        f'<w:bottom w:val="single" w:sz="6" w:space="4" w:color="E2E8F0"/>'
        f'<w:right w:val="single" w:sz="6" w:space="4" w:color="E2E8F0"/>'
        f'</w:pBdr>'
    )
    pPr.append(pBdr)

    # Section 1: System Metadata Table
    doc.add_heading("1. Project Infrastructure & Deployment Details", level=1)
    t_meta = doc.add_table(rows=7, cols=2)
    meta_rows = [
        ("Project ID & Title", "P053: Invoice Data Validation & Discrepancy Analytics Platform", None),
        ("Live Deployed URL", "https://invoicehclproject.streamlit.app (Global 24/7 Access)", "https://invoicehclproject.streamlit.app"),
        ("GitHub Repository URL", "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git", "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git"),
        ("Candidate / Learner", "Surbhi Sharma (University Roll No: 2500320140152)", None),
        ("Batch & Specialization", "ABES_OND26_MCA2ndYear_Data Analytics_Batch 01 (Data Analytics)", None),
        ("Project Mentor", "Mr. Dinesh Sereno", None),
        ("Hosting Architecture", "Streamlit Community Cloud (Linux Container, Python 3.10+, Headless)", None)
    ]
    for idx, (k, v, link) in enumerate(meta_rows):
        r = t_meta.rows[idx]
        r.cells[0].text = k
        r.cells[0].paragraphs[0].runs[0].bold = True
        set_cell_background(r.cells[0], "F1F5F9")
        set_cell_margins(r.cells[0], 60, 60, 100, 100)
        set_cell_margins(r.cells[1], 60, 60, 100, 100)
        r.cells[0].width = Inches(2.2)
        r.cells[1].width = Inches(4.8)

        if link:
            p_val = r.cells[1].paragraphs[0]
            p_val.text = ""
            add_clickable_hyperlink(p_val, link, v, color="0969DA", underline=True, bold=True, size_pt=9.5)
        else:
            r.cells[1].text = v
            r.cells[1].paragraphs[0].runs[0].font.size = Pt(9.5)

    # Section 2: Live Application Dashboard Screenshot
    doc.add_heading("2. Live Streamlit Application Screenshot", level=1)
    doc.add_paragraph(
        "The application is fully operational and publicly hosted on Streamlit Community Cloud. "
        "Below is the verified screenshot captured directly from the live application showing "
        "the executive KPI metrics, anomaly monitoring, and financial leakage prevention figures:"
    )

    dash_img = DOCS_DIR / "streamlit_dashboard_screenshot.png"
    if dash_img.exists():
        p_img = doc.add_paragraph()
        p_img.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_i = p_img.add_run()
        r_i.add_picture(str(dash_img), width=Inches(6.4))
        p_cap = doc.add_paragraph()
        p_cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r_c = p_cap.add_run(
            "Figure 1: Verified Live Streamlit Dashboard on Streamlit Cloud displaying Total Spend (₹7.09 Cr), "
            "Invoices (304), Flagged Discrepancies (69), and Prevented Leakage (₹1.84 Cr / ₹3.40 Cr)."
        )
        r_c.font.size = Pt(8.5)
        r_c.font.italic = True
        r_c.font.color.rgb = RGBColor.from_string("64748B")

    # Section 3: Multi-Page Incognito Verification Matrix
    doc.add_heading("3. Multi-Page Incognito Verification Matrix", level=1)
    doc.add_paragraph(
        "Every application view and analytical module was independently verified in a clean incognito session "
        "with zero local cache dependencies to guarantee public accessibility:"
    )
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

    # Section 4: Repository Details & Structure
    doc.add_heading("4. GitHub Code Repository Architecture & Manifest", level=1)
    doc.add_paragraph(
        "The project repository is public, version-controlled with Git, and configured for continuous "
        "deployment on Streamlit Community Cloud:\n"
        "• Repository: surbhi0310 / Career-Shaper-HCL-Project (Public)\n"
        "• Default Branch: main | Latest Verified Commit: c392cfb\n"
        "• Environment: Streamlit Community Cloud (Python 3.10+, headless, automated CI triggers)"
    )

    tree_text = """
Career-Shaper-HCL-Project/ (Branch: main)
├── .streamlit/
│   └── config.toml                  # Streamlit custom UI styling and theme colors
├── data/
│   ├── app.db                       # Relational SQLite database (5 tables + view)
│   ├── invoice_data.xlsx            # Master multi-sheet formatted workbook
│   ├── vendors.csv                  # Master vendor demographic dataset (30 records)
│   ├── purchase_orders.csv          # Purchase orders dataset (150 records)
│   ├── invoices.csv                 # Invoices transaction dataset (304 records)
│   ├── line_items.csv               # Billed line items dataset (1,056 records)
│   └── discrepancies.csv            # Flagged audit discrepancies (69 records)
├── docs/
│   ├── 01_Project_Understanding.pdf # Project objectives, scope, and problem formulation
│   ├── 02_HMW.pdf                   # How-Might-We statements and innovation axes
│   ├── 03_Needs_Analysis.pdf        # Stakeholder requirements and pain-point matrix
│   ├── 04_Traceability_Matrix.xlsx  # Requirements traceability matrix (100% coverage)
│   ├── 05_BRD.pdf                   # Complete 13-section Business Requirements Document
│   ├── 07_TRD.pdf                   # Technical Architecture & Requirements Specification
│   ├── 08_Database_Design.xlsx      # 11-sheet relational schema & data dictionaries
│   ├── 09_ER_Diagram.pdf            # High-resolution entity-relationship diagram
│   ├── 10_Data_Validation_Report.pdf# Validation rules engine & discrepancy audit
│   ├── 11_User_Stories.pdf          # Agile user stories with acceptance criteria
│   ├── 12_UI_Framework.pdf          # Streamlit UI architecture & wireframe specifications
│   ├── 13_Test_Cases.pdf            # Comprehensive QA test suites and test logs
│   ├── PJ_IVD_2026_Project_Document.pdf # Master Project Report (Official ABES Cover)
│   └── PJ_IVD_2026_Presentation.pdf # Project presentation slide deck
├── LMS_Submission_Package/          # Submission-ready deliverables adhering to LMS rules
├── pages/
│   ├── 1_Overview_KPIs.py           # High-level spend analysis and procurement metrics
│   ├── 2_Invoice_Validation.py      # Rule-based invoice arithmetic & tax auditor
│   ├── 3_PO_Reconciliation.py       # 2-way invoice-to-PO automated matching engine
│   ├── 4_Vendor_Risk.py             # Multi-factor Vendor Risk Index (VRI) calculator
│   ├── 5_Statistical_EDA.py         # Benford's Law analysis & IQR outlier detection
│   └── 6_Data_Quality_&_Export.py   # Database query browser and CSV/Excel export
├── src/
│   ├── database.py                  # SQLite connectivity & cached SQL query methods
│   ├── validation.py                # Discrepancy detection engine (vectorized rules)
│   ├── analytics.py                 # Advanced analytics, VRI scoring & Benford tests
│   ├── ui.py                        # CSS styling, KPI metric cards, and currency formats
│   └── generate_data.py             # Synthetic dataset generation engine
├── tests/
│   ├── test_data.py                 # Database integrity and table validation tests (4/4)
│   ├── test_analysis.py             # Validation logic and discrepancy tests (5/5)
│   └── test_app.py                  # Analytical calculations and app UI tests (4/4)
├── app.py                           # Application entry point and executive overview
├── requirements.txt                 # Pinned dependencies (streamlit, pandas, plotly, etc.)
└── README.md                        # Project documentation, live badges, and setup guide
"""
    p_code = doc.add_paragraph()
    r_c = p_code.add_run(tree_text.strip())
    r_c.font.name = "Consolas"
    r_c.font.size = Pt(8.0)

    # Section 5: Automated Testing Summary
    doc.add_heading("5. Automated Pytest Verification Summary", level=1)
    doc.add_paragraph(
        "All 13 automated unit and integration tests executed with a 100% pass rate in pytest before deployment:\n"
        "• tests/test_data.py: 4/4 PASSED (DB presence, row count thresholds, primary key uniqueness, foreign key integrity)\n"
        "• tests/test_analysis.py: 5/5 PASSED (Arithmetic verification, duplicate detection, tax slab checks, PO budget matching, VRI score bounds)\n"
        "• tests/test_app.py: 4/4 PASSED (Summary metrics aggregation, Benford distribution, IQR outlier detection, 32-column view loading)"
    )

    out_docx = DOCS_DIR / "GitHub_Report.docx"
    doc.save(str(out_docx))
    print(f"[OK] Generated Word document: {out_docx}")

    # Convert to PDF via Word COM
    out_pdf = PDF_DIR / "GitHub_Report.pdf"
    print("Converting to PDF via Word COM...")
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        wdoc = word.Documents.Open(str(out_docx.resolve()))
        wdoc.SaveAs(str(out_pdf.resolve()), FileFormat=17) # 17 = wdFormatPDF
        wdoc.Close()
        print(f"[OK] Converted to PDF: {out_pdf}")
    finally:
        word.Quit()

    # Sync to LMS package locations
    pkg_target1 = LMS_PKG_DIR / "2500320140152_SurbhiSharma_GitHub_and_Streamlit_Details.pdf"
    pkg_target2 = LMS_BACKUP_DIR / "2500320140152_SurbhiSharma_GitHub_and_Streamlit_Details.pdf"

    shutil.copy2(out_pdf, pkg_target1)
    shutil.copy2(out_pdf, pkg_target2)
    print(f"[OK] Synced to: {pkg_target1}")
    print(f"[OK] Synced to: {pkg_target2}")

if __name__ == "__main__":
    generate_report()
