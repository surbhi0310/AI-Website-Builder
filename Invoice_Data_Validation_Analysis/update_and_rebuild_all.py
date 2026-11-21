"""
update_and_rebuild_all.py - Complete metadata update, document regeneration, and LMS package build
New Metadata:
- Roll No: 2500320140152
- Batch ID: ABES_OND26_MCA2ndYear_Data Analytics_Batch 01
- Project ID: P053
- Project Mentor: Mr. Dinesh Sereno
- Domain: Data Analytics
- Candidate: Surbhi Sharma
"""

import os
import re
import sys
import shutil
from pathlib import Path

BASE_DIR = Path(r"C:\Users\Geetanjali\Downloads\AI-website-Builder-final-project-main\Invoice_Data_Validation_Analysis")

# 1. Update text files / generator scripts
files_to_update = [
    BASE_DIR / "build_abes_doc.py",
    BASE_DIR / "build_abes_presentation.py",
    BASE_DIR / "build_step_docs.py",
    BASE_DIR / "build_missing_steps.py",
    BASE_DIR / "generate_lms_special_reports.py",
    BASE_DIR / "prepare_lms_package.py",
    BASE_DIR / "README.md",
    BASE_DIR / "BUILD_PLAN.md",
    BASE_DIR / "conftest.py"
]

replacements = [
    ("PJ-IVD-2026", "P053"),
    ("Kalpana Dwivedi", "Mr. Dinesh Sereno"),
    ("ABES_OND26_MCA Batch 02", "ABES_OND26_MCA2ndYear_Data Analytics_Batch 01"),
    ("ABES_OND26_MCA2ndYear_Data_Analytics_Batch_01", "ABES_OND26_MCA2ndYear_Data Analytics_Batch 01"),
    ("https://github.com/2500320140152/invoice-validation-analytics", "https://github.com/surbhi0310/Career-Shaper-HCL-Project.git")
]

for fp in files_to_update:
    if fp.exists():
        content = fp.read_text(encoding="utf-8")
        for old_str, new_str in replacements:
            content = content.replace(old_str, new_str)
        fp.write_text(content, encoding="utf-8")
        print(f"[OK] Updated metadata in {fp.name}")

# Also update pages/ metadata comments
for page_file in (BASE_DIR / "pages").glob("*.py"):
    content = page_file.read_text(encoding="utf-8")
    for old_str, new_str in replacements:
        content = content.replace(old_str, new_str)
    page_file.write_text(content, encoding="utf-8")
print("[OK] Updated pages/ scripts")

# Also update src/ metadata comments
for src_file in (BASE_DIR / "src").glob("*.py"):
    content = src_file.read_text(encoding="utf-8")
    for old_str, new_str in replacements:
        content = content.replace(old_str, new_str)
    src_file.write_text(content, encoding="utf-8")
print("[OK] Updated src/ scripts")

print("\nExecuting generators with updated metadata...")

# Run build_abes_doc.py
os.system(f'cd /d "{BASE_DIR}" && python build_abes_doc.py')

# Run build_abes_presentation.py
os.system(f'cd /d "{BASE_DIR}" && python build_abes_presentation.py')

# Run build_step_docs.py
os.system(f'cd /d "{BASE_DIR}" && python build_step_docs.py')

# Run build_missing_steps.py
os.system(f'cd /d "{BASE_DIR}" && python build_missing_steps.py')

# Run generate_lms_special_reports.py
os.system(f'cd /d "{BASE_DIR}" && python generate_lms_special_reports.py')

print("\nConverting all DOCX and PPTX to PDF...")
# Run convert_all_to_pdf.py
os.system(f'cd /d "{BASE_DIR}" && python convert_all_to_pdf.py')

print("\nAssembling LMS Submission Package...")
# Run prepare_lms_package.py
os.system(f'cd /d "{BASE_DIR}" && python prepare_lms_package.py')

print("\nAll tasks completed successfully!")
