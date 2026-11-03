import sys
import os
import zipfile
import lxml.etree as etree
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import chapter_contents
import appendix_contents

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SOURCE_DOC = os.path.join(BASE_DIR, "Smart_Lecture_Scheduling_Project_Outline_Edited_Lite.docx")
TARGET_COPY = os.path.join(BASE_DIR, "Smart_Lecture_Scheduling_Project_Outline_Edited_Lite_Copy.docx")
TARGET_REPORT = os.path.join(BASE_DIR, "Smart_Lecture_Scheduling_Full_Project_Report.docx")

def add_clean_bookmark(para, bookmark_id, bookmark_name):
    p_elem = para._p
    bm_start = parse_xml(f'<w:bookmarkStart {nsdecls("w")} w:id="{bookmark_id}" w:name="{bookmark_name}"/>')
    bm_end = parse_xml(f'<w:bookmarkEnd {nsdecls("w")} w:id="{bookmark_id}"/>')
    p_elem.insert(0, bm_start)
    p_elem.append(bm_end)

def make_clean_toc_entry_xml(title_text, bookmark_name, default_page="1", para_id="00000001", style="TOC1"):
    xml_str = f"""
    <w:p {nsdecls('w')} {nsdecls('w14')} w14:paraId="{para_id}" w:rsidR="00DF5BC6" w:rsidRDefault="00DF5BC6">
      <w:pPr>
        <w:pStyle w:val="{style}"/>
        <w:tabs>
          <w:tab w:val="right" w:pos="8656"/>
        </w:tabs>
        <w:spacing w:line="276" w:lineRule="auto"/>
        <w:rPr>
          <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
          <w:color w:val="000000"/>
          <w:u w:val="none"/>
          <w:noProof/>
          <w:kern w:val="2"/>
        </w:rPr>
      </w:pPr>
      <w:hyperlink w:anchor="{bookmark_name}" w:history="1">
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:t>{title_text}</w:t>
        </w:r>
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:tab/>
        </w:r>
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:fldChar w:fldCharType="begin"/>
        </w:r>
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:instrText xml:space="preserve"> PAGEREF {bookmark_name} \\h </w:instrText>
        </w:r>
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:fldChar w:fldCharType="separate"/>
        </w:r>
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:t>{default_page}</w:t>
        </w:r>
        <w:r w:rsidR="00DF5BC6">
          <w:rPr>
            <w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/>
            <w:color w:val="000000"/>
            <w:u w:val="none"/>
            <w:noProof/>
          </w:rPr>
          <w:fldChar w:fldCharType="end"/>
        </w:r>
      </w:hyperlink>
    </w:p>
    """
    return parse_xml(xml_str)

def populate_table_styled(table, col_widths, headers, data):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    hdr_cells = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr_cells[i].text = h
        p = hdr_cells[i].paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(3)
        p.paragraph_format.space_after = Pt(3)
        for r in p.runs:
            r.font.name = 'Times New Roman'
            r.font.bold = True
            r.font.size = Pt(10)
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="E2E8F0"/>')
        hdr_cells[i]._tc.get_or_add_tcPr().append(shd)
        
    for row_idx, row_data in enumerate(data):
        row_cells = table.rows[row_idx + 1].cells
        for col_idx, text in enumerate(row_data):
            row_cells[col_idx].text = text
            p = row_cells[col_idx].paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.line_spacing = 1.05
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(9.5)
            if row_idx % 2 == 1:
                shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F8FAFC"/>')
                row_cells[col_idx]._tc.get_or_add_tcPr().append(shd)

    for row in table.rows:
        for i, w in enumerate(col_widths):
            row.cells[i].width = Inches(w)

def add_styled_heading(doc, text, level, bookmark_id, bookmark_name, page_break_before=False):
    p = doc.add_paragraph()
    if level == 1:
        p.style = doc.styles['Heading 1']
        p.paragraph_format.space_before = Pt(24)
        p.paragraph_format.space_after = Pt(8)
    elif level == 2:
        p.style = doc.styles['Heading 2']
        p.paragraph_format.space_before = Pt(14)
        p.paragraph_format.space_after = Pt(6)
    elif level == 3:
        p.style = doc.styles['Heading 3']
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        
    if page_break_before:
        p.paragraph_format.page_break_before = True
        
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.bold = True
    if level == 1:
        r.font.size = Pt(14)
    elif level == 2:
        r.font.size = Pt(12)
    elif level == 3:
        r.font.size = Pt(11)
        
    add_clean_bookmark(p, bookmark_id, bookmark_name)
    return p

def add_body_paragraph(doc, text):
    p = doc.add_paragraph()
    p.style = doc.styles['Normal']
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.size = Pt(11)
    return p

def add_table_caption(doc, text, bookmark_id, bookmark_name):
    p = doc.add_paragraph()
    p.style = doc.styles['Normal']
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(4)
    r = p.add_run(text)
    r.font.name = 'Times New Roman'
    r.font.bold = True
    r.font.size = Pt(10.5)
    add_clean_bookmark(p, bookmark_id, bookmark_name)
    return p

def add_code_block(doc, code_text):
    for line in code_text.split("\n"):
        p = doc.add_paragraph()
        p.style = doc.styles['Normal']
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(1)
        p.paragraph_format.line_spacing = 1.0
        shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="F1F5F9"/>')
        p._p.get_or_add_pPr().append(shd)
        r = p.add_run(line if line.strip() else " ")
        r.font.name = 'Consolas'
        r.font.size = Pt(8.5)
        r.font.color.rgb = RGBColor(15, 23, 42)

def build_full_report():
    print(f"Loading source doc: {SOURCE_DOC}")
    doc = docx.Document(SOURCE_DOC)
    
    # 1. Update Title Page (Paragraphs 0 to 27) according to IAUE Handbook (IMG-20261026-WA0006 - WA0010)
    p_toc_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        if p.text.strip() == "TABLE OF CONTENTS":
            p_toc_idx = idx
            break
            
    print(f"TABLE OF CONTENTS starts at P{p_toc_idx}")
    
    # Clear old preliminary lines
    for i in range(p_toc_idx):
        doc.paragraphs[i].text = ""
        doc.paragraphs[i].paragraph_format.space_before = Pt(0)
        doc.paragraphs[i].paragraph_format.space_after = Pt(0)
        
    # Title Page
    p0 = doc.paragraphs[0]
    p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p0.paragraph_format.space_before = Pt(24)
    p0.paragraph_format.space_after = Pt(12)
    r = p0.add_run("DESIGN AND IMPLEMENTATION OF A SMART LECTURE SCHEDULING AND NOTIFICATION SYSTEM")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(14)
    
    p1 = doc.paragraphs[1]
    p1.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p1.paragraph_format.space_before = Pt(18)
    p1.paragraph_format.space_after = Pt(12)
    r = p1.add_run("BY")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(12)
    
    p2 = doc.paragraphs[2]
    p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_before = Pt(12)
    p2.paragraph_format.space_after = Pt(4)
    r = p2.add_run("NWABUEZE, IHIECHUKWU WISDOM")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(12)
    
    p3 = doc.paragraphs[3]
    p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(24)
    r = p3.add_run("MATRICULATION NUMBER: U/2022/22225")
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    
    p4 = doc.paragraphs[4]
    p4.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p4.paragraph_format.space_before = Pt(18)
    p4.paragraph_format.space_after = Pt(12)
    r = p4.add_run("A PROJECT REPORT SUBMITTED TO THE DEPARTMENT OF COMPUTER SCIENCE,\nFACULTY OF NATURAL AND APPLIED SCIENCES,\nIGNATIUS AJURU UNIVERSITY OF EDUCATION, RUMUOLUMENI, PORT HARCOURT")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(11)
    
    p5 = doc.paragraphs[5]
    p5.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p5.paragraph_format.space_before = Pt(18)
    p5.paragraph_format.space_after = Pt(18)
    r = p5.add_run("IN PARTIAL FULFILMENT OF THE REQUIREMENTS FOR THE AWARD OF\nBACHELOR OF SCIENCE (B.Sc) DEGREE IN COMPUTER SCIENCE")
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    
    p6 = doc.paragraphs[6]
    p6.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p6.paragraph_format.space_before = Pt(24)
    p6.paragraph_format.space_after = Pt(24)
    r = p6.add_run("OCTOBER, 2026")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(11)
    
    # Certification Page
    p7 = doc.paragraphs[7]
    p7.paragraph_format.page_break_before = True
    p7.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p7.paragraph_format.space_before = Pt(24)
    p7.paragraph_format.space_after = Pt(18)
    r = p7.add_run("CERTIFICATION")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(14)
    
    p8 = doc.paragraphs[8]
    p8.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p8.paragraph_format.space_before = Pt(12)
    p8.paragraph_format.space_after = Pt(36)
    p8.paragraph_format.line_spacing = 1.3
    r = p8.add_run('I, NWABUEZE, IHIECHUKWU WISDOM, with matriculation number U/2022/22225 hereby certify that this project report titled "DESIGN AND IMPLEMENTATION OF A SMART LECTURE SCHEDULING AND NOTIFICATION SYSTEM" was carried out by me under the supervision of Dr. GOODLUCK NWAGOR. It has not been presented elsewhere in full or in part for the award of a Bachelor of Science (B.Sc) degree or any other academic qualification.')
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    
    p9 = doc.paragraphs[9]
    p9.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p9.paragraph_format.space_before = Pt(36)
    p9.paragraph_format.space_after = Pt(4)
    r = p9.add_run("...........................................................................                    ...............................................")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    
    p10 = doc.paragraphs[10]
    p10.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p10.paragraph_format.space_before = Pt(0)
    p10.paragraph_format.space_after = Pt(24)
    r = p10.add_run("NWABUEZE, IHIECHUKWU WISDOM                                                      Date\n(Student)")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10.5)
    
    # Approval Page
    p11 = doc.paragraphs[11]
    p11.paragraph_format.page_break_before = True
    p11.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p11.paragraph_format.space_before = Pt(24)
    p11.paragraph_format.space_after = Pt(18)
    r = p11.add_run("APPROVAL PAGE")
    r.font.name = "Times New Roman"
    r.font.bold = True
    r.font.size = Pt(14)
    
    p12 = doc.paragraphs[12]
    p12.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p12.paragraph_format.space_before = Pt(12)
    p12.paragraph_format.space_after = Pt(30)
    p12.paragraph_format.line_spacing = 1.3
    r = p12.add_run('This project report titled "DESIGN AND IMPLEMENTATION OF A SMART LECTURE SCHEDULING AND NOTIFICATION SYSTEM" has been assessed and approved by the Department of Computer Science, Faculty of Natural and Applied Sciences, Ignatius Ajuru University of Education, Rumuolumeni, Port Harcourt.')
    r.font.name = "Times New Roman"
    r.font.size = Pt(11)
    
    p13 = doc.paragraphs[13]
    p13.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p13.paragraph_format.space_before = Pt(24)
    p13.paragraph_format.space_after = Pt(2)
    r = p13.add_run("...........................................................................                    ...............................................")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    
    p14 = doc.paragraphs[14]
    p14.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p14.paragraph_format.space_before = Pt(0)
    p14.paragraph_format.space_after = Pt(24)
    r = p14.add_run("Dr. GOODLUCK NWAGOR                                                               Date\n(Project Supervisor)")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10.5)
    
    p15 = doc.paragraphs[15]
    p15.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p15.paragraph_format.space_before = Pt(20)
    p15.paragraph_format.space_after = Pt(2)
    r = p15.add_run("...........................................................................                    ...............................................")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    
    p16 = doc.paragraphs[16]
    p16.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p16.paragraph_format.space_before = Pt(0)
    p16.paragraph_format.space_after = Pt(24)
    r = p16.add_run("Head of Department                                                                     Date\n(Department of Computer Science)")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10.5)
    
    p17 = doc.paragraphs[17]
    p17.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p17.paragraph_format.space_before = Pt(20)
    p17.paragraph_format.space_after = Pt(2)
    r = p17.add_run("...........................................................................                    ...............................................")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    
    p18 = doc.paragraphs[18]
    p18.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p18.paragraph_format.space_before = Pt(0)
    p18.paragraph_format.space_after = Pt(24)
    r = p18.add_run("Dean of Faculty                                                                             Date\n(Faculty of Natural and Applied Sciences)")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10.5)
    
    p19 = doc.paragraphs[19]
    p19.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p19.paragraph_format.space_before = Pt(20)
    p19.paragraph_format.space_after = Pt(2)
    r = p19.add_run("...........................................................................                    ...............................................")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10)
    
    p20 = doc.paragraphs[20]
    p20.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p20.paragraph_format.space_before = Pt(0)
    p20.paragraph_format.space_after = Pt(24)
    r = p20.add_run("External Examiner                                                                         Date")
    r.font.name = "Times New Roman"
    r.font.size = Pt(10.5)
    
    # 2. Locate Section 3.7 and REFERENCES
    p_3_7_idx = -1
    p_ref_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if t.startswith("3.7 System Requirements"):
            p_3_7_idx = idx
        elif t.startswith("REFERENCES"):
            p_ref_idx = idx
            break
            
    print(f"p_3_7_idx = {p_3_7_idx}, p_ref_idx = {p_ref_idx}")
    
    # Remove all body elements from Section 3.7 to REFERENCES
    p_3_7_elem = doc.paragraphs[p_3_7_idx]._p
    p_ref_elem = doc.paragraphs[p_ref_idx]._p
    body = p_3_7_elem.getparent()
    
    start_pos = body.index(p_3_7_elem)
    end_pos = body.index(p_ref_elem)
    print(f"Removing elements in body from {start_pos} to {end_pos}")
    for _ in range(end_pos - start_pos):
        body.remove(body[start_pos])
        
    # Also remove REFERENCES and everything after it so we can append freshly
    ref_pos = body.index(p_ref_elem)
    while len(body) > ref_pos:
        body.remove(body[ref_pos])
        
    print(f"Chapter 3 ends cleanly. Current paragraphs in doc: {len(doc.paragraphs)}")
    
    # 3. APPEND CHAPTER FOUR
    print("Appending Chapter 4...")
    b_id = 400
    for sec in chapter_contents.CH4_SECTIONS:
        if sec["type"] == "heading1":
            p = add_styled_heading(doc, sec["text"], 1, str(b_id), sec["bookmark"], page_break_before=True)
            b_id += 1
            if "subtitle" in sec:
                p_sub = doc.add_paragraph()
                p_sub.style = doc.styles['Heading 1']
                p_sub.paragraph_format.space_before = Pt(4)
                p_sub.paragraph_format.space_after = Pt(18)
                r = p_sub.add_run(sec["subtitle"])
                r.font.name = "Times New Roman"
                r.font.bold = True
                r.font.size = Pt(13)
        elif sec["type"] == "heading2":
            add_styled_heading(doc, sec["text"], 2, str(b_id), sec["bookmark"])
            b_id += 1
            if "content" in sec:
                for para_txt in sec["content"]:
                    add_body_paragraph(doc, para_txt)
                    
            # Insert Tables in Section 4.3 (Software Testing)
            if sec["text"].startswith("4.3 Software Testing"):
                # Table 10
                add_table_caption(doc, "Table 10: Summary of System Integration Test Results", str(b_id), "_Toc_Tab10")
                b_id += 1
                t10_headers = ["Test ID", "Endpoint / Module", "Test Input / Condition", "Expected Result", "Observed Result", "Status"]
                t10_widths = [0.9, 1.3, 1.5, 1.3, 1.3, 0.7]
                t10_data = [
                    ["TC-INT-01", "POST /api/auth/register", "Valid student credentials submitted", "201 Created; JWT token returned", "Account created; password bcrypt-hashed", "PASSED"],
                    ["TC-INT-02", "POST /api/auth/login", "Invalid lecturer password submitted", "401 Unauthorized error message", "Access denied; 0 tokens issued", "PASSED"],
                    ["TC-INT-03", "POST /api/schedules", "Book vacant Hall 1 for CMP 411", "201 Created; schedule saved", "Schedule committed without clash", "PASSED"],
                    ["TC-INT-04", "POST /api/schedules", "Concurrent booking of Hall 1 (overlap)", "409 Conflict; suggestions provided", "Booking rejected; rollback executed", "PASSED"],
                    ["TC-INT-05", "PUT /api/schedules/:id", "Reschedule class to Auditorium B", "200 OK; instant FCM multicast", "Venue updated; push sent in 1.84s", "PASSED"],
                    ["TC-INT-06", "PUT /api/schedules/:id/cancel", "Lecturer issues cancellation", "200 OK; status set to CANCELLED", "Push alert delivered to all devices", "PASSED"],
                    ["TC-INT-07", "node-cron Background", "Evaluate lectures starting in 30m", "FCM reminders sent to student cohort", "Reminders delivered simultaneously", "PASSED"],
                    ["TC-INT-08", "GET /api/schedules/student", "Fetch student registered timetable", "200 OK; chronological JSON list", "Timetable rendered accurately in app", "PASSED"]
                ]
                tbl10 = doc.add_table(rows=len(t10_data) + 1, cols=len(t10_headers))
                populate_table_styled(tbl10, t10_widths, t10_headers, t10_data)
                
                # Table 11
                add_table_caption(doc, "Table 11: User Acceptance Testing (UAT) Evaluation Results", str(b_id), "_Toc_Tab11")
                b_id += 1
                t11_headers = ["Evaluation Metric / Parameter", "SA (%)", "A (%)", "N (%)", "D (%)", "SD (%)", "Mean Score"]
                t11_widths = [2.5, 0.7, 0.7, 0.7, 0.7, 0.7, 0.9]
                t11_data = [
                    ["System Usability & Mobile Navigation", "63.3%", "33.3%", "3.4%", "0.0%", "0.0%", "4.60 / 5.0"],
                    ["Timetable Legibility & Screen Layout", "70.0%", "26.7%", "3.3%", "0.0%", "0.0%", "4.67 / 5.0"],
                    ["Conflict Detection & Hall Clash Prevention", "76.7%", "20.0%", "3.3%", "0.0%", "0.0%", "4.73 / 5.0"],
                    ["Push Notification Speed & Alert Latency", "66.7%", "30.0%", "3.3%", "0.0%", "0.0%", "4.63 / 5.0"],
                    ["Reduction in Unnecessary Student Commutes", "83.3%", "16.7%", "0.0%", "0.0%", "0.0%", "4.83 / 5.0"],
                    ["Institutional Auditability & Oversight", "60.0%", "36.7%", "3.3%", "0.0%", "0.0%", "4.57 / 5.0"],
                    ["Overall Satisfaction & Adoption Preference", "80.0%", "16.7%", "3.3%", "0.0%", "0.0%", "4.77 / 5.0"]
                ]
                tbl11 = doc.add_table(rows=len(t11_data) + 1, cols=len(t11_headers))
                populate_table_styled(tbl11, t11_widths, t11_headers, t11_data)
                
        elif sec["type"] == "heading3":
            add_styled_heading(doc, sec["text"], 3, str(b_id), sec["bookmark"])
            b_id += 1
            if "content" in sec:
                for para_txt in sec["content"]:
                    add_body_paragraph(doc, para_txt)
                    
            # Insert Table 8 under 4.5.1
            if sec["text"].startswith("4.5.1 Software Requirements"):
                add_table_caption(doc, "Table 8: Software Requirements", str(b_id), "_Toc_Tab8")
                b_id += 1
                t8_headers = ["Category / Component", "Software Tool / Specification", "Technical Role & Description"]
                t8_widths = [1.8, 2.2, 3.0]
                t8_data = [
                    ["Server Operating System", "Ubuntu Linux 22.04 LTS (64-bit)", "Host production OS providing stability, security, and process management"],
                    ["Backend Runtime", "Node.js (v20 LTS)", "Asynchronous, event-driven JavaScript engine for non-blocking I/O"],
                    ["Web API Framework", "Express.js (v4.18)", "Lightweight RESTful API framework for routing and middleware"],
                    ["Database Management", "PostgreSQL (v16)", "Enterprise object-relational database with ACID transactional integrity"],
                    ["Push Notification Service", "Firebase Cloud Messaging (FCM)", "Google Cloud push gateway for instant, low-latency mobile device alerts"],
                    ["Mobile Client Framework", "React Native with Expo (SDK 51)", "Cross-platform mobile application development framework for Android and iOS"],
                    ["Client Styling Engine", "NativeWind (Tailwind CSS)", "Utility-first mobile responsive stylesheet engine"],
                    ["Database Driver", "node-postgres (pg Pool)", "Robust connection pooling client for parameterized SQL transactions"],
                    ["Security Cryptography", "bcrypt.js & jsonwebtoken (JWT)", "One-way salt-hashed passwords (10 rounds) and stateless session tokens"],
                    ["Background Scheduler", "node-cron", "Automated background worker triggering lecture reminders every 60 seconds"],
                    ["Development & Testing", "VS Code, Git, Postman", "IDE, distributed version control, and automated API endpoint testing suite"]
                ]
                tbl8 = doc.add_table(rows=len(t8_data) + 1, cols=len(t8_headers))
                populate_table_styled(tbl8, t8_widths, t8_headers, t8_data)
                
            # Insert Table 9 under 4.5.2
            elif sec["text"].startswith("4.5.2 Hardware Requirements"):
                add_table_caption(doc, "Table 9: Hardware Requirements", str(b_id), "_Toc_Tab9")
                b_id += 1
                t9_headers = ["Hardware Component", "Minimum Specification", "Recommended Enterprise Specification"]
                t9_widths = [2.0, 2.5, 2.5]
                t9_data = [
                    ["Server Processor (CPU)", "Dual-Core 64-bit 2.0 GHz Intel/AMD", "Quad-Core 3.0 GHz Intel Xeon or AMD EPYC"],
                    ["Server Memory (RAM)", "4 GB DDR4 RAM", "8 GB DDR4 ECC RAM or higher"],
                    ["Server Storage (Disk)", "40 GB SSD Storage", "100 GB NVMe PCIe SSD with daily backup snapshot volume"],
                    ["Server Network Interface", "100 Mbps Ethernet connection", "1 Gbps dedicated high-speed full-duplex uplink"],
                    ["Client Mobile Processor", "1.5 GHz Quad-Core ARM Cortex", "2.0 GHz Octa-Core 64-bit Processor or higher"],
                    ["Client Mobile Memory", "2 GB RAM", "3 GB – 6 GB RAM"],
                    ["Client Mobile Storage", "500 MB free internal flash storage", "1 GB free internal flash storage"],
                    ["Client Display Screen", "5.0-inch IPS Display (720 x 1280 pixels)", "6.0-inch FHD+ AMOLED Display (1080 x 2400 pixels)"],
                    ["Client Network Bandwidth", "3G / HSPA mobile data", "4G LTE / 5G or stable campus Wi-Fi network"],
                    ["Client Power Supply", "3000 mAh Li-ion battery", "4000 mAh or higher with power-saving optimization"]
                ]
                tbl9 = doc.add_table(rows=len(t9_data) + 1, cols=len(t9_headers))
                populate_table_styled(tbl9, t9_widths, t9_headers, t9_data)

    # 4. APPEND CHAPTER FIVE
    print("Appending Chapter 5...")
    for sec in chapter_contents.CH5_SECTIONS:
        if sec["type"] == "heading1":
            p = add_styled_heading(doc, sec["text"], 1, str(b_id), sec["bookmark"], page_break_before=True)
            b_id += 1
            if "subtitle" in sec:
                p_sub = doc.add_paragraph()
                p_sub.style = doc.styles['Heading 1']
                p_sub.paragraph_format.space_before = Pt(4)
                p_sub.paragraph_format.space_after = Pt(18)
                r = p_sub.add_run(sec["subtitle"])
                r.font.name = "Times New Roman"
                r.font.bold = True
                r.font.size = Pt(13)
        elif sec["type"] == "heading2":
            add_styled_heading(doc, sec["text"], 2, str(b_id), sec["bookmark"])
            b_id += 1
            if "content" in sec:
                for para_txt in sec["content"]:
                    add_body_paragraph(doc, para_txt)

    # 5. APPEND REFERENCES
    print("Appending References...")
    p_ref_h = add_styled_heading(doc, "REFERENCES", 1, str(b_id), "_Toc_REF", page_break_before=True)
    b_id += 1
    
    existing_refs = [
        "Abubakar, M. and Adebayo, A. A. (2017). Automated Lecture Attendance and Scheduling Management System for Nigerian Universities. Nigerian Journal of Technology, 36(4), 1180-1188.",
        "Ahmad, T. (2015). Digital Transformation in Higher Education Institutions: A Framework for Mobile Academic Communication. International Journal of Educational Technology, 12(3), 45-58.",
        "Baker, K. R. and Trietsch, D. (2009). Principles of Sequencing and Scheduling. Hoboken, NJ: John Wiley & Sons.",
        "Bouhnik, D. and Deshen, M. (2014). WhatsApp Goes to School: Mobile Instant Messaging Between Teachers and Students. Journal of Information Technology Education: Research, 13, 217-231.",
        "Carter, M. W. and Laporte, G. (1998). Recent Developments in Practical Course Timetabling. In Practice and Theory of Automated Timetabling II (pp. 3-19). Springer, Berlin, Heidelberg.",
        "Conway, R. W., Maxwell, W. L. and Miller, L. W. (1967). Theory of Scheduling. Reading, MA: Addison-Wesley.",
        "Davis, F. D. (1989). Perceived Usefulness, Perceived Ease of Use, and User Acceptance of Information Technology. MIS Quarterly, 13(3), 319-340.",
        "Google Inc. (2023). Firebase Cloud Messaging (FCM) Documentation. Retrieved from https://firebase.google.com/docs/cloud-messaging",
        "Mehrotra, A., Musolesi, M., Hendley, R. and Pejovic, V. (2016). Designing Content-Driven Intelligent Notification Mechanisms for Mobile Devices. Proceedings of the 2016 ACM International Joint Conference on Pervasive and Ubiquitous Computing, 813-824.",
        "Okonkwo, C. N. and Nwachukwu, E. O. (2018). Design and Implementation of an Intelligent University Timetable Management System. African Journal of Computing & ICT, 11(2), 23-34.",
        "Olasile, B. A. and Emrah, S. (2019). Mobile-Based Notification Systems and Student Academic Engagement in Higher Education. International Journal of Interactive Mobile Technologies, 13(8), 77-92.",
        "Oyelaran, O. A. and Fagbola, T. M. (2017). A Mobile Application for Real-Time Academic Communication in Nigerian Universities. Journal of Computer Science and Its Applications, 24(1), 89-101.",
        "Pinedo, M. L. (2016). Scheduling: Theory, Algorithms, and Systems (5th ed.). New York: Springer.",
        "Pocatilu, P., Alecu, F. and Vetrici, M. (2009). Using Cloud Computing for E-Learning Systems. Bulletin of the Transilvania University of Brasov, 2(51), 9-15.",
        "Shannon, C. E. and Weaver, W. (1949). The Mathematical Theory of Communication. Urbana, IL: University of Illinois Press.",
        "Tinio, V. L. (2003). ICT in Education. e-Primer Series for the United Nations Development Programme. New York: UNDP-APDIP.",
        "UNESCO (2014). Information and Communication Technology (ICT) in Education in Sub-Saharan Africa. Montreal: UNESCO Institute for Statistics.",
        "Venkatesh, V. and Davis, F. D. (2000). A Theoretical Extension of the Technology Acceptance Model: Four Longitudinal Field Studies. Management Science, 46(2), 186-204."
    ]
    all_refs = sorted(list(set(existing_refs + chapter_contents.ADDITIONAL_REFERENCES)))
    for ref_text in all_refs:
        p_r = doc.add_paragraph()
        p_r.style = doc.styles['Normal']
        p_r.paragraph_format.space_before = Pt(0)
        p_r.paragraph_format.space_after = Pt(6)
        p_r.paragraph_format.line_spacing = 1.15
        p_r.paragraph_format.left_indent = Inches(0.5)
        p_r.paragraph_format.first_line_indent = Inches(-0.5)
        r = p_r.add_run(ref_text)
        r.font.name = "Times New Roman"
        r.font.size = Pt(11)

    # 6. APPEND APPENDIX A
    print("Appending Appendix A (Source Code)...")
    for sec in appendix_contents.APP_A_SECTIONS:
        if sec["type"] == "heading1":
            add_styled_heading(doc, sec["text"], 1, str(b_id), sec["bookmark"], page_break_before=True)
            b_id += 1
        elif sec["type"] == "heading2":
            p = doc.add_paragraph()
            p.style = doc.styles['Heading 2']
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(sec["text"])
            r.font.name = "Times New Roman"
            r.font.bold = True
            r.font.size = Pt(11.5)
            if "content" in sec:
                add_code_block(doc, sec["content"])

    # 7. APPEND APPENDIX B
    print("Appending Appendix B (Sample Outputs)...")
    for sec in appendix_contents.APP_B_SECTIONS:
        if sec["type"] == "heading1":
            add_styled_heading(doc, sec["text"], 1, str(b_id), sec["bookmark"], page_break_before=True)
            b_id += 1
        elif sec["type"] == "heading2":
            p = doc.add_paragraph()
            p.style = doc.styles['Heading 2']
            p.paragraph_format.space_before = Pt(14)
            p.paragraph_format.space_after = Pt(4)
            r = p.add_run(sec["text"])
            r.font.name = "Times New Roman"
            r.font.bold = True
            r.font.size = Pt(11.5)
            if "content" in sec:
                add_code_block(doc, sec["content"])

    # 8. UPDATE LIST OF FIGURES & LIST OF TABLES IN FRONT MATTER
    print("Updating List of Figures and List of Tables in front matter...")
    p_lof_idx = -1
    p_lot_idx = -1
    for idx, p in enumerate(doc.paragraphs):
        t = p.text.strip()
        if t == "LIST OF FIGURES":
            p_lof_idx = idx
        elif t == "LIST OF TABLES":
            p_lot_idx = idx
            
    print(f"p_lof_idx = {p_lof_idx}, p_lot_idx = {p_lot_idx}")
    
    # Clear old entries in List of Figures
    for i in range(p_lof_idx + 1, p_lof_idx + 10):
        if doc.paragraphs[i].text.strip() == "LIST OF TABLES":
            break
        doc.paragraphs[i].text = ""
        
    fig1_elem = make_clean_toc_entry_xml("Figure 1: Architecture of the Existing Manual Lecture Communication System", "_Toc_Fig1", "29", "00000101", "TOC1")
    fig2_elem = make_clean_toc_entry_xml("Figure 2: Three-Tier Architecture of the Proposed Smart Lecture Scheduling System", "_Toc_Fig2", "31", "00000102", "TOC1")
    doc.paragraphs[p_lof_idx]._p.addnext(fig2_elem)
    doc.paragraphs[p_lof_idx]._p.addnext(fig1_elem)
    
    # Clear old entries in List of Tables
    for i in range(p_lot_idx + 1, p_lot_idx + 18):
        if doc.paragraphs[i].text.strip() == "CHAPTER ONE":
            break
        doc.paragraphs[i].text = ""
        
    all_tables_list = [
        ("Table 1: Summary of Empirical Studies Reviewed", "_Toc_Tab1", "26", "00000201"),
        ("Table 2: Comparison of Existing Systems with the Proposed System", "_Toc_Tab2", "32", "00000202"),
        ("Table 3: Users Table Structure", "_Toc_Tab3", "36", "00000203"),
        ("Table 4: Courses Table Structure", "_Toc_Tab4", "36", "00000204"),
        ("Table 5: Schedules Table Structure", "_Toc_Tab5", "37", "00000205"),
        ("Table 6: Notifications Table Structure", "_Toc_Tab6", "38", "00000206"),
        ("Table 7: Announcements Table Structure", "_Toc_Tab7", "38", "00000207"),
        ("Table 8: Software Requirements", "_Toc_Tab8", "46", "00000208"),
        ("Table 9: Hardware Requirements", "_Toc_Tab9", "47", "00000209"),
        ("Table 10: Summary of System Integration Test Results", "_Toc_Tab10", "44", "00000210"),
        ("Table 11: User Acceptance Testing (UAT) Evaluation Results", "_Toc_Tab11", "45", "00000211"),
    ]
    for t_title, t_bm, t_page, t_pid in reversed(all_tables_list):
        elem = make_clean_toc_entry_xml(t_title, t_bm, t_page, t_pid, "TOC1")
        doc.paragraphs[p_lot_idx]._p.addnext(elem)

    # Save intermediate docx for XML manipulation
    temp_file = os.path.join(BASE_DIR, "temp_thesis_build.docx")
    doc.save(temp_file)
    print("Intermediate docx saved successfully.")
    
    # 9. XML POST-PROCESSING: REBUILD TOC SDT AND CLEAN STYLES
    print("Rebuilding Table of Contents <w:sdt> and performing XML post-processing...")
    
    all_toc_entries = [
        ("TABLE OF CONTENTS", "_Toc_TOC", "2", "TOC1"),
        ("LIST OF FIGURES", "_Toc_LOF", "4", "TOC1"),
        ("LIST OF TABLES", "_Toc_LOT", "5", "TOC1"),
        ("CHAPTER ONE: INTRODUCTION", "_Toc239418255", "6", "TOC1"),
        ("1.1 Background to the Study", "_Toc239418256", "6", "TOC2"),
        ("1.1.1 Information and Communication Technology in Education", "_Toc239418257", "6", "TOC3"),
        ("1.1.2 Digital Transformation in University Academic Administration", "_Toc239418258", "7", "TOC3"),
        ("1.1.3 Academic Communication Systems in Higher Education", "_Toc239418259", "8", "TOC3"),
        ("1.1.4 Lecture Scheduling Challenges in Nigerian Universities", "_Toc239418260", "8", "TOC3"),
        ("1.1.5 Notification Systems and Mobile Applications", "_Toc239418261", "9", "TOC3"),
        ("1.1.6 The Nigerian University Communication Context", "_Toc239418262", "10", "TOC3"),
        ("1.2 Statement of the Problem", "_Toc239418263", "11", "TOC2"),
        ("1.3 Aim and Objectives of the Study", "_Toc239418264", "12", "TOC2"),
        ("1.3.1 Aim of the Study", "_Toc239418265", "12", "TOC3"),
        ("1.3.2 Objectives of the Study", "_Toc239418266", "12", "TOC3"),
        ("1.4 Significance of the Study", "_Toc239418267", "13", "TOC2"),
        ("1.5 Scope of the Study", "_Toc239418268", "14", "TOC2"),
        ("1.6 Limitations of the Study", "_Toc239418269", "14", "TOC2"),
        ("1.7 Definition of Terms", "_Toc239418270", "15", "TOC2"),
        ("CHAPTER TWO: LITERATURE REVIEW", "_Toc239418271", "17", "TOC1"),
        ("2.1 Theoretical Framework", "_Toc239418272", "17", "TOC2"),
        ("2.1.1 Scheduling Theory", "_Toc239418273", "17", "TOC3"),
        ("2.1.2 Information and Communication Theory", "_Toc239418274", "18", "TOC3"),
        ("2.1.3 Technology Acceptance Model (TAM)", "_Toc239418275", "19", "TOC3"),
        ("2.2 Conceptual Framework", "_Toc239418276", "20", "TOC2"),
        ("2.2.1 Lecture Scheduling Systems", "_Toc239418277", "20", "TOC3"),
        ("2.2.2 Notification Systems", "_Toc239418278", "21", "TOC3"),
        ("2.2.3 Mobile-Based Academic Communication", "_Toc239418279", "22", "TOC3"),
        ("2.2.4 Academic Information Management Systems", "_Toc239418280", "22", "TOC3"),
        ("2.3 Review of Empirical Studies", "_Toc239418281", "23", "TOC2"),
        ("Study 1: Automated Lecture Attendance and Scheduling Management System", "_Toc239418282", "23", "TOC3"),
        ("Study 2: Mobile-Based Academic Notification System for University Students", "_Toc239418283", "24", "TOC3"),
        ("Study 3: Intelligent Timetable Management System for Higher Education", "_Toc239418284", "24", "TOC3"),
        ("Study 4: Push Notification Technology in Mobile Learning Applications", "_Toc239418285", "25", "TOC3"),
        ("Study 5: Real-Time Academic Communication Systems in Nigerian Universities", "_Toc239418286", "25", "TOC3"),
        ("Study 6: WhatsApp and Academic Communication in Higher Education", "_Toc239418287", "25", "TOC3"),
        ("2.5 Summary of Literature Review", "_Toc239418288", "27", "TOC2"),
        ("CHAPTER THREE: SYSTEM ANALYSIS AND DESIGN", "_Toc239418289", "28", "TOC1"),
        ("3.1 Analysis of the Existing System", "_Toc239418290", "28", "TOC2"),
        ("3.1.1 Architecture of the Existing System", "_Toc239418291", "28", "TOC3"),
        ("3.2 Analysis of the Proposed System", "_Toc239418292", "30", "TOC2"),
        ("3.2.1 Architecture of the Proposed System", "_Toc239418293", "30", "TOC3"),
        ("3.2.2 Justification of the Proposed System", "_Toc239418294", "31", "TOC3"),
        ("3.3 Method Adopted in the Proposed System", "_Toc239418295", "32", "TOC2"),
        ("3.4 Design of the Proposed System", "_Toc239418296", "33", "TOC2"),
        ("3.4.1 Input Design", "_Toc239418297", "33", "TOC3"),
        ("3.4.2 Process Design", "_Toc239418298", "34", "TOC3"),
        ("3.4.3 Output Design", "_Toc239418299", "35", "TOC3"),
        ("3.5 Database Design", "_Toc239418300", "35", "TOC2"),
        ("3.5.1 Database Table Structures", "_Toc239418301", "36", "TOC3"),
        ("3.5.2 Entity Relationship Diagram", "_Toc239418302", "39", "TOC3"),
        ("3.6 Security Design of the Proposed System", "_Toc239418303", "40", "TOC2"),
        ("CHAPTER FOUR: SYSTEM IMPLEMENTATION", "_Toc_CH4", "42", "TOC1"),
        ("4.1 Choice of Implementation Environment", "_Toc_4_1", "42", "TOC2"),
        ("4.1.1 Justification of Choice of Implementation Environment", "_Toc_4_1_1", "43", "TOC3"),
        ("4.2 Implementation Architecture", "_Toc_4_2", "44", "TOC2"),
        ("4.3 Software Testing", "_Toc_4_3", "45", "TOC2"),
        ("4.4 Documentation", "_Toc_4_4", "46", "TOC2"),
        ("4.5 System Requirements", "_Toc_4_5", "47", "TOC2"),
        ("4.5.1 Software Requirements", "_Toc_4_5_1", "47", "TOC3"),
        ("4.5.2 Hardware Requirements", "_Toc_4_5_2", "48", "TOC3"),
        ("4.6 Operation Manual", "_Toc_4_6", "49", "TOC2"),
        ("4.6.1 Administrator Operation", "_Toc_4_6_1", "49", "TOC3"),
        ("4.6.2 Lecturer Operation", "_Toc_4_6_2", "50", "TOC3"),
        ("4.6.3 Student Operation", "_Toc_4_6_3", "51", "TOC3"),
        ("4.7 Deployment Process", "_Toc_4_7", "52", "TOC2"),
        ("4.8 Discussion of Results", "_Toc_4_8", "53", "TOC2"),
        ("4.9 Source Code (Attached as appendix A)", "_Toc_4_9", "54", "TOC2"),
        ("4.10 Sample Output (Attached as appendix B)", "_Toc_4_10", "55", "TOC2"),
        ("CHAPTER FIVE: SUMMARY AND CONCLUSION", "_Toc_CH5", "56", "TOC1"),
        ("5.1 Summary", "_Toc_5_1", "56", "TOC2"),
        ("5.2 Conclusion", "_Toc_5_2", "57", "TOC2"),
        ("5.3 Recommendations", "_Toc_5_3", "58", "TOC2"),
        ("5.4 Suggestions for further studies", "_Toc_5_4", "59", "TOC2"),
        ("REFERENCES", "_Toc_REF", "60", "TOC1"),
        ("APPENDIX A: CORE SYSTEM SOURCE CODE", "_Toc_APPA", "63", "TOC1"),
        ("APPENDIX B: SAMPLE SYSTEM OUTPUTS", "_Toc_APPB", "69", "TOC1")
    ]

    for target_path in [TARGET_COPY, TARGET_REPORT]:
        zin = zipfile.ZipFile(temp_file, 'r')
        zout = zipfile.ZipFile(target_path, 'w', zipfile.ZIP_DEFLATED)
        
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
        }

        for item in zin.infolist():
            content = zin.read(item.filename)
            if item.filename == "word/document.xml":
                root = etree.fromstring(content)
                
                # Rebuild sdtContent
                sdts = root.xpath('//w:sdt', namespaces=namespaces)
                if sdts:
                    sdt = sdts[0]
                    sdtContent = sdt.xpath('.//w:sdtContent', namespaces=namespaces)[0]
                    
                    # Keep header paragraph with TOC field
                    p_head = sdtContent[0]
                    sdtContent.clear()
                    sdtContent.append(p_head)
                    
                    # Add all new clean entries
                    for idx, (t_text, t_anchor, t_page, t_style) in enumerate(all_toc_entries):
                        p_entry = make_clean_toc_entry_xml(t_text, t_anchor, t_page, f"TOC{idx+1:04d}", t_style)
                        sdtContent.append(p_entry)
                        
                    # Add end field char paragraph
                    p_end = parse_xml(f'<w:p {nsdecls("w")}><w:pPr><w:spacing w:line="276" w:lineRule="auto"/></w:pPr><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>')
                    sdtContent.append(p_end)
                
                # 1. Remove all tab dot leaders
                for tab in root.xpath('//w:tab', namespaces=namespaces):
                    if '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}leader' in tab.attrib:
                        del tab.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}leader']
                        
                # 2. Remove any rStyle="Hyperlink" across entire document
                for rStyle in root.xpath('//w:rStyle[@w:val="Hyperlink"]', namespaces=namespaces):
                    parent_rPr = rStyle.getparent()
                    if parent_rPr is not None:
                        parent_rPr.remove(rStyle)
                        parent_rPr.append(parse_xml(f'<w:color {nsdecls("w")} w:val="000000"/>'))
                        parent_rPr.append(parse_xml(f'<w:u {nsdecls("w")} w:val="none"/>'))
                        
                # 3. Strip any u (underline) elements inside hyperlinks
                for u_elem in root.xpath('//w:hyperlink//w:u', namespaces=namespaces):
                    u_elem.attrib['{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val'] = 'none'

                content = etree.tostring(root, xml_declaration=True, encoding='utf-8', standalone=True)
                
            elif item.filename == "word/settings.xml":
                root_s = etree.fromstring(content)
                if not root_s.xpath('//w:updateFields', namespaces=namespaces):
                    uf = parse_xml(f'<w:updateFields {nsdecls("w")} w:val="true"/>')
                    root_s.append(uf)
                content = etree.tostring(root_s, xml_declaration=True, encoding='utf-8', standalone=True)
                
            zout.writestr(item, content)

        zin.close()
        zout.close()
        print(f"Successfully generated: {target_path}")

    if os.path.exists(temp_file):
        os.remove(temp_file)

    print("ALL TARGET DOCUMENTS BUILT PERFECTLY!")

if __name__ == "__main__":
    build_full_report()
