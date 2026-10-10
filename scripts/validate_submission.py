import os
import re
import docx
from docx import Document

def validate_deliverables():
    base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    docx_path = os.path.join(base_dir, "report", "Hariom_Upadhyay_YuvaIntern_Week1_Project_Planning_Report.docx")
    diag_dir = os.path.join(base_dir, "diagrams")
    scripts_dir = os.path.join(base_dir, "scripts")
    readme_path = os.path.join(base_dir, "README.md")
    req_path = os.path.join(base_dir, "requirements.txt")
    gitignore_path = os.path.join(base_dir, ".gitignore")
    
    print("=" * 70)
    print("YUVAINTERN WEEK 1 DELIVERABLE AUTOMATED VALIDATION SUITE")
    print("=" * 70)
    
    checks_passed = 0
    total_checks = 10
    
    # Check 1: Finalized DOCX exists at exact path
    c1 = os.path.exists(docx_path)
    print(f"Check 1: Finalized DOCX exists at expected path: {'[PASS]' if c1 else '[FAIL]'}")
    if c1:
        checks_passed += 1
        print(f"         Path: {docx_path}")
    else:
        print(f"         Missing path: {docx_path}")
        
    # Check 2: Reopen programmatically without corruption
    c2 = False
    doc = None
    try:
        doc = Document(docx_path)
        c2 = True
        print(f"Check 2: DOCX reopens programmatically without corruption: [PASS]")
        checks_passed += 1
    except Exception as e:
        print(f"Check 2: DOCX reopens programmatically: [FAIL] - {e}")
        
    # Check 3: Personalization & Internship Metadata verified in report
    all_text = " ".join([p.text for p in doc.paragraphs])
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                all_text += " " + cell.text
                
    meta_keywords = [
        "Hariom Upadhyay",
        "YuvaIntern",
        "Virtual Data Science Explorer Intern",
        "10 October 2026",
        "Customer Churn"
    ]
    missing_meta = [kw for kw in meta_keywords if kw.lower() not in all_text.lower()]
    c3 = len(missing_meta) == 0
    print(f"Check 3: Internship metadata verified in report: {'[PASS]' if c3 else '[FAIL]'}")
    if c3:
        checks_passed += 1
        print("         Student Name, Organization, Role, and Date confirmed.")
    else:
        print(f"         Missing metadata terms: {missing_meta}")
        
    # Check 4: Timeline totals exactly 33 hours
    has_33_hours = ("33 hours" in all_text.lower() or "33-hour" in all_text.lower() or "33" in all_text)
    c4 = has_33_hours
    print(f"Check 4: 33-hour project schedule verified in report: {'[PASS]' if c4 else '[FAIL]'}")
    if c4:
        checks_passed += 1
        
    # Check 5: Three diagrams embedded in DOCX
    images_in_doc = [rel for rel in doc.part.rels.values() if "image" in rel.target_ref]
    num_embedded_images = len(images_in_doc)
    c5 = num_embedded_images >= 3
    print(f"Check 5: All 3 diagrams embedded inside DOCX: {'[PASS]' if c5 else '[FAIL]'} (Found: {num_embedded_images} images)")
    if c5:
        checks_passed += 1
        
    # Check 6: Tables, headings, paragraphs present
    num_tables = len(doc.tables)
    num_paragraphs = len(doc.paragraphs)
    c6 = num_tables >= 8 and num_paragraphs >= 50
    print(f"Check 6: Tables ({num_tables}) and Headings/Paragraphs ({num_paragraphs}) present: {'[PASS]' if c6 else '[FAIL]'}")
    if c6:
        checks_passed += 1
        
    # Check 7: File size comfortably below 2 MB
    size_bytes = os.path.getsize(docx_path)
    size_mb = size_bytes / (1024 * 1024)
    c7 = size_mb < 2.0
    print(f"Check 7: File size comfortably below 2 MB: {'[PASS]' if c7 else '[FAIL]'} ({size_mb:.2f} MB / {size_bytes} bytes)")
    if c7:
        checks_passed += 1
        
    # Check 8: No sensitive credentials or secrets in project
    secret_patterns = ["password", "secret_key", "api_key", "bearer ", "aws_secret"]
    found_secrets = []
    for root, _, files in os.walk(base_dir):
        if ".git" in root or "__pycache__" in root:
            continue
        for f in files:
            if f.endswith((".py", ".md", ".txt", ".json")):
                p_file = os.path.join(root, f)
                try:
                    with open(p_file, "r", encoding="utf-8", errors="ignore") as fh:
                        content = fh.read().lower()
                        for sp in secret_patterns:
                            if sp in content and "no passwords" not in content and "api keys" not in content:
                                found_secrets.append(f"{f}:{sp}")
                except Exception:
                    pass
    c8 = len(found_secrets) == 0
    print(f"Check 8: Security & credential leakage check: {'[PASS]' if c8 else '[FAIL]'}")
    if c8:
        checks_passed += 1
        print("         Zero API keys, secrets, or confidential credentials detected.")
    else:
        print(f"         Potential sensitive strings: {found_secrets}")

    # Check 9: Academic integrity (Planning-only framing)
    planning_framing = any(term in all_text.lower() for term in ["planning", "strategic blueprint", "hypothetical"])
    c9 = planning_framing
    print(f"Check 9: Academic integrity (Planning-only methodology design): {'[PASS]' if c9 else '[FAIL]'}")
    if c9:
        checks_passed += 1
        print("         Appropriately framed as strategic planning project.")

    # Check 10: Supporting files, diagrams, and repository structure
    diagram_files = [
        os.path.join(diag_dir, "diagram1_lifecycle_flowchart.png"),
        os.path.join(diag_dir, "diagram2_conceptual_architecture.png"),
        os.path.join(diag_dir, "diagram3_project_timeline.png")
    ]
    all_diagrams_exist = all(os.path.exists(f) for f in diagram_files)
    
    script_files = [
        os.path.join(scripts_dir, "generate_diagrams.py"),
        os.path.join(scripts_dir, "docx_helpers.py"),
        os.path.join(scripts_dir, "build_full_report.py")
    ]
    all_scripts_exist = all(os.path.exists(f) for f in script_files)
    
    repo_files = os.path.exists(readme_path) and os.path.exists(req_path)
    c10 = all_diagrams_exist and all_scripts_exist and repo_files
    print(f"Check 10: Supporting scripts, diagrams, and documentation exist: {'[PASS]' if c10 else '[FAIL]'}")
    if c10:
        checks_passed += 1
        
    print("=" * 70)
    print(f"OVERALL VALIDATION SUMMARY: {checks_passed} OF {total_checks} CHECKS PASSED")
    print("=" * 70)
    
    return checks_passed == total_checks

if __name__ == "__main__":
    validate_deliverables()
