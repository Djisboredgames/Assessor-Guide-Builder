"""
docx_generator.py - Generates CDU Assessor Guide .docx files using python-docx

Based on CDU Assessor Guide v7 template (October 2025).
"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os
from datetime import datetime


# ── Styling constants ──────────────────────────────────────────────
FONT_BODY = "Arial"
FONT_HEADING = "Arial"
COLOR_CDU_GOLD = RGBColor(0xFF, 0xD7, 0x00)
COLOR_CDU_DARK = RGBColor(0x1A, 0x1A, 0x2E)
COLOR_BLACK = RGBColor(0, 0, 0)
COLOR_GREY = RGBColor(0x66, 0x66, 0x66)
COLOR_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
COLOR_HEADER_BG = "1A1A2E"
COLOR_HIGHLIGHT = "FFFF00"
COLOR_LIGHT_GREY = "F2F2F2"
COLOR_TABLE_HEADER = "D5D5D5"

BORDER_COLOR = "999999"

PT_BODY = Pt(10)
PT_SMALL = Pt(9)
PT_HEADING1 = Pt(14)
PT_HEADING2 = Pt(12)
PT_TITLE = Pt(20)


def set_cell_shading(cell, color_hex):
    """Apply background shading to a table cell."""
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color_hex}" w:val="clear"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def set_cell_border(cell, **kwargs):
    """Set cell borders. kwargs: top, bottom, left, right with values like '1' for thin."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(f'<w:tcBorders {nsdecls("w")}/>')
    for edge, val in kwargs.items():
        element = parse_xml(
            f'<w:{edge} {nsdecls("w")} w:val="single" w:sz="4" w:space="0" w:color="{BORDER_COLOR}"/>'
        )
        tcBorders.append(element)
    tcPr.append(tcBorders)


def styled_paragraph(doc_or_cell, text, bold=False, size=PT_BODY, font=FONT_BODY,
                      alignment=None, space_before=0, space_after=0, color=COLOR_BLACK,
                      italic=False):
    """Add a styled paragraph to a document or table cell."""
    if hasattr(doc_or_cell, 'add_paragraph'):
        p = doc_or_cell.add_paragraph()
    else:
        p = doc_or_cell.paragraphs[0] if doc_or_cell.paragraphs else doc_or_cell.add_paragraph()
    
    run = p.add_run(text)
    run.font.name = font
    run.font.size = size
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = color
    
    if alignment:
        p.alignment = alignment
    
    pf = p.paragraph_format
    pf.space_before = Pt(space_before)
    pf.space_after = Pt(space_after)
    
    return p


def add_heading_styled(doc, text, level=1):
    """Add a numbered heading matching CDU template style."""
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.font.name = FONT_HEADING
    run.font.bold = True
    run.font.size = PT_HEADING1 if level == 1 else PT_HEADING2
    run.font.color.rgb = COLOR_BLACK
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    return p


def add_bullet(doc, text, level=0, bold=False):
    """Add a bullet point paragraph."""
    p = doc.add_paragraph(style='List Bullet')
    p.clear()
    run = p.add_run(text)
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    run.font.bold = bold
    if level > 0:
        p.paragraph_format.left_indent = Inches(0.5 * (level + 1))
    return p


def add_checkbox_row(table, question_text, yes_checked=False, no_checked=False):
    """Add a row with a Yes/No checkbox question to a table."""
    row = table.add_row()
    row.cells[0].text = ""
    p = row.cells[0].paragraphs[0]
    run = p.add_run(question_text)
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    
    check_yes = "☒" if yes_checked else "☐"
    check_no = "☒" if no_checked else "☐"
    
    p2 = row.cells[1].paragraphs[0]
    run2 = p2.add_run(f"{check_yes} Y  {check_no} N")
    run2.font.name = FONT_BODY
    run2.font.size = PT_BODY
    
    return row


def build_assessor_guide(unit_data, assessment_config, output_path):
    """
    Build the complete CDU Assessor Guide .docx document.
    
    Args:
        unit_data: UnitData object with all unit information
        assessment_config: dict with assessment task configuration
        output_path: path to save the .docx file
    """
    doc = Document()
    
    # ── Page setup ──
    section = doc.sections[0]
    section.page_width = Inches(8.27)   # A4
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(1)
    section.right_margin = Inches(1)
    
    # Set default font
    style = doc.styles['Normal']
    font = style.font
    font.name = FONT_BODY
    font.size = PT_BODY
    
    # ══════════════════════════════════════════════════════════════
    # COVER / TITLE AREA
    # ══════════════════════════════════════════════════════════════
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("Assessor Guide")
    run.font.name = FONT_HEADING
    run.font.size = PT_TITLE
    run.font.bold = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{unit_data.code} {unit_data.title}")
    run.font.name = FONT_HEADING
    run.font.size = Pt(14)
    run.font.bold = True
    run.font.color.rgb = COLOR_CDU_DARK
    p.paragraph_format.space_after = Pt(24)
    
    # ══════════════════════════════════════════════════════════════
    # PURPOSE
    # ══════════════════════════════════════════════════════════════
    styled_paragraph(doc, "Purpose", bold=True, size=PT_HEADING2, space_before=12, space_after=6)
    styled_paragraph(
        doc,
        "The Assessor Guide provides the Assessor (VET Lecturer) with instructions and "
        "assessment instruments to assess the unit of competency or cluster of units. "
        "The guide should be read in conjunction with the Student Unit Guide and all "
        "developed individual assessment instruments for this unit of competency.",
        space_after=12
    )
    
    # ══════════════════════════════════════════════════════════════
    # 1. UNIT DETAILS
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "1. Unit Details")
    
    # Unit details table
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    table.autofit = True
    
    # Unit Code and Title
    row = table.add_row()
    row.cells[0].text = ""
    p = row.cells[0].paragraphs[0]
    run = p.add_run("Unit Code and Title")
    run.font.bold = True
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    row.cells[1].text = f"{unit_data.code} {unit_data.title}"
    
    # Application
    row = table.add_row()
    p = row.cells[0].paragraphs[0]
    run = p.add_run("Application of the Unit or Descriptor")
    run.font.bold = True
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    row.cells[1].text = unit_data.application or "[Insert from training.gov.au]"
    
    # Prerequisites
    check_yes = "☒" if unit_data.has_prerequisites else "☐"
    check_no = "☐" if unit_data.has_prerequisites else "☒"
    row = table.add_row()
    p = row.cells[0].paragraphs[0]
    run = p.add_run("Are there pre-requisites for this unit?")
    run.font.bold = True
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    p2 = row.cells[1].paragraphs[0]
    run2 = p2.add_run(f"{check_yes} Yes  {check_no} No")
    run2.font.name = FONT_BODY
    run2.font.size = PT_BODY
    if unit_data.has_prerequisites and unit_data.prerequisites:
        p3 = row.cells[1].add_paragraph()
        run3 = p3.add_run(", ".join(unit_data.prerequisites))
        run3.font.name = FONT_BODY
        run3.font.size = PT_BODY
    
    # Licensing
    check_yes = "☒" if unit_data.has_licensing else "☐"
    check_no = "☐" if unit_data.has_licensing else "☒"
    row = table.add_row()
    p = row.cells[0].paragraphs[0]
    run = p.add_run("Are there licensing or regulatory requirements?")
    run.font.bold = True
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    p2 = row.cells[1].paragraphs[0]
    run2 = p2.add_run(f"{check_yes} Yes  {check_no} No")
    run2.font.name = FONT_BODY
    run2.font.size = PT_BODY
    
    # Format table widths
    for row in table.rows:
        row.cells[0].width = Inches(2.5)
        row.cells[1].width = Inches(4)
    
    # ══════════════════════════════════════════════════════════════
    # 2. UNIT ASSESSMENT SUMMARY
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "2. Unit Assessment Summary")
    
    tasks = assessment_config.get("tasks", [])
    table = doc.add_table(rows=1, cols=4)
    table.style = 'Table Grid'
    
    # Header row
    headers = ["Assessment\ntask number", "Assessment Method and Description", "No. of\nAttempts", "Due Date"]
    for i, header in enumerate(headers):
        cell = table.rows[0].cells[i]
        p = cell.paragraphs[0]
        run = p.add_run(header)
        run.font.bold = True
        run.font.name = FONT_BODY
        run.font.size = PT_SMALL
        set_cell_shading(cell, COLOR_TABLE_HEADER)
    
    for task in tasks:
        row = table.add_row()
        row.cells[0].text = str(task.get("number", ""))
        row.cells[1].text = task.get("method", "")
        row.cells[2].text = str(task.get("attempts", "3"))
        row.cells[3].text = task.get("due_date", "")
    
    # ══════════════════════════════════════════════════════════════
    # 3. ASSESSMENT CONDITIONS / CONTEXT
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "3. Unit Assessment Conditions / Context and Specific Resources")
    
    styled_paragraph(
        doc,
        "Access www.training.gov.au and review the Assessment Conditions or the Context "
        "and Specific Resource Requirements, and answer the following questions:",
        space_after=6
    )
    
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    
    # Q1 - Additional assessor requirements
    add_checkbox_row(
        table,
        "1. Are there any additional Assessor requirements other than meeting the 2025 Standards for RTOs?",
        yes_checked=unit_data.has_additional_assessor_reqs,
        no_checked=not unit_data.has_additional_assessor_reqs
    )
    # Detail row
    row = table.add_row()
    row.cells[0].merge(row.cells[1])
    row.cells[0].text = assessment_config.get("assessor_req_detail", "")
    
    # Q2 - Simulated environment
    add_checkbox_row(
        table,
        "2. Is a simulated assessment environment included in the conditions/context?",
        yes_checked=unit_data.has_simulated_environment,
        no_checked=not unit_data.has_simulated_environment
    )
    row = table.add_row()
    row.cells[0].merge(row.cells[1])
    row.cells[0].text = ""
    
    # Q3 - Resources
    add_checkbox_row(
        table,
        "3. Does CDU have all the required listed resources?",
        yes_checked=unit_data.has_required_resources,
        no_checked=not unit_data.has_required_resources
    )
    row = table.add_row()
    row.cells[0].merge(row.cells[1])
    row.cells[0].text = ""
    
    # ══════════════════════════════════════════════════════════════
    # 4. PERFORMANCE EVIDENCE REQUIREMENTS
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "4. Performance Evidence Requirements")
    
    styled_paragraph(
        doc,
        "Access www.training.gov.au and review the Performance Evidence or Critical "
        "Aspects of Evidence and answer the following questions:",
        space_after=6
    )
    
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    
    add_checkbox_row(
        table,
        "1. Is there a requirement for volume and frequency?",
        yes_checked=unit_data.has_volume_frequency,
        no_checked=not unit_data.has_volume_frequency
    )
    row = table.add_row()
    row.cells[0].merge(row.cells[1])
    if unit_data.volume_frequency_detail:
        p = row.cells[0].paragraphs[0]
        run = p.add_run(unit_data.volume_frequency_detail)
        run.font.name = FONT_BODY
        run.font.size = PT_BODY
        run.font.italic = True
    
    add_checkbox_row(
        table,
        "2. Is there a requirement for actual work placement hours?",
        yes_checked=unit_data.has_work_placement,
        no_checked=not unit_data.has_work_placement
    )
    row = table.add_row()
    row.cells[0].merge(row.cells[1])
    row.cells[0].text = unit_data.work_placement_detail or ""
    
    # ══════════════════════════════════════════════════════════════
    # 5. REASONABLE ADJUSTMENT (boilerplate)
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "5. Reasonable Adjustment")
    
    styled_paragraph(
        doc,
        "Should a student have specific needs that require reasonable adjustment, "
        "then the Assessor must:",
        space_after=4
    )
    
    adjustments = [
        "review the unit requirements and determine that any adjustments will not compromise the outcome.",
        "determine the adjustments to be made in consultation with the student and a specialist if necessary.",
        "document the adjustments made in the appropriate section on the Assessment Agreement in the Student Unit Guide.",
        "ensure the student's privacy and confidentiality is protected in relation to any personal information, such as a medical condition."
    ]
    for adj in adjustments:
        add_bullet(doc, adj)
    
    styled_paragraph(doc, "", space_after=4)
    styled_paragraph(
        doc,
        "The following are examples of wording that could be used to describe reasonable "
        "adjustment that was provided for an assessment task:",
        space_after=4
    )
    
    examples = [
        "Resources were customised to suit student needs.",
        "Assistive/adaptive technology was used.",
        "A sign language Interpreter was used.",
        "The time was extended for this task to suit student needs.",
        "The assessment was delivered verbally to suit student needs.",
        "Rest breaks were provided during this assessment to suit student needs.",
        "An alternative venue was provided to suit student needs.",
        "The student assessment responses were recorded by video/audio.",
        "Plain English was used to describe these tasks.",
        "The assessment task was completed online to suit student needs."
    ]
    for ex in examples:
        add_bullet(doc, ex)
    
    # ══════════════════════════════════════════════════════════════
    # 6. RPL (boilerplate)
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "6. Recognition of Prior Learning (RPL)")
    
    rpl_items = [
        "To assess a candidate's skills and knowledge, use practical methods like on-site questioning and observation.",
        "Ensure you understand the relevant competencies and qualifications appropriate to the candidate's goals.",
        "Encourage a competency conversation, which allows candidates to share their experiences.",
        "It's essential to verify the information gathered from interviews and observations with someone who has had the opportunity to observe the candidate's skills over time.",
        "Keep detailed records of conversations, skills demonstrations, and any documents reviewed to support the claim of prior learning.",
        "If a candidate wants to RPL a full qualification, skill set, or accredited course, but has skill gaps, work with them to create a plan to complete training for any outstanding units."
    ]
    for item in rpl_items:
        add_bullet(doc, item)
    
    # ══════════════════════════════════════════════════════════════
    # 7. FEEDBACK (boilerplate)
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "7. Feedback")
    styled_paragraph(
        doc,
        "The Assessor must provide written feedback at the end of each assessment task "
        "in the provided feedback area of each assessment instrument.",
        space_after=4
    )
    styled_paragraph(
        doc,
        "The Student will sign that they received the feedback on the 'Assessment Summary' "
        "upon completing all assessment tasks.",
        space_after=8
    )
    
    # ══════════════════════════════════════════════════════════════
    # 8. RECORD KEEPING (boilerplate)
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "8. Record Keeping")
    styled_paragraph(
        doc,
        "The following documentation must be stored per the team's repository process, "
        "i.e., TAFE SharePoint Team Document Libraries, ShareStream, Learnline or a hard "
        "copy storage filing system and per the Records and Information Management Policy and Procedure.",
        space_after=4
    )
    styled_paragraph(doc, "The VET Team or Assessor must retain for each VET student all records of assessment, including:", space_after=4)
    add_bullet(doc, "completed student assessment tasks containing Assessor feedback.")
    add_bullet(doc, "completed and signed (by the student and Assessor) Assessment Summary or Learnline Gradebook.")
    add_bullet(doc, "RPL evidence (if applicable).")
    
    styled_paragraph(doc, "", space_after=4)
    styled_paragraph(doc, "The VET Team or Assessor must retain for each VET unit documents, including:", space_after=4)
    add_bullet(doc, "Student Unit Guide")
    add_bullet(doc, "Assessor Guide")
    add_bullet(doc, "RPL Guides")
    add_bullet(doc, "Assessment tools")
    
    # ══════════════════════════════════════════════════════════════
    # 9. RESULTING ASSESSMENTS (boilerplate)
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "9. Resulting Assessments")
    styled_paragraph(
        doc,
        "After the assessment task has been completed and a judgement has been made, "
        "the Assessor records the result and provides written feedback in the appropriate "
        "space on the assessment task instrument.",
        space_after=4
    )
    styled_paragraph(doc, "The result for an assessment task can be:", bold=True, space_after=4)
    add_bullet(doc, "Satisfactory")
    add_bullet(doc, "Unsatisfactory")
    
    styled_paragraph(doc, "", space_after=2)
    styled_paragraph(doc, "The final result for the unit can be:", bold=True, space_after=4)
    add_bullet(doc, "CA – Competency Achieved")
    add_bullet(doc, "NYC – Not Yet Competent")
    add_bullet(doc, "IP – Insufficient Participation.")
    
    # ══════════════════════════════════════════════════════════════
    # 10. AI STATEMENT (boilerplate)
    # ══════════════════════════════════════════════════════════════
    add_heading_styled(doc, "10. Artificial Intelligence Software")
    styled_paragraph(doc, "CDU encourages CDU staff and students to:", space_after=4)
    ai_items = [
        "Explore the benefits of AI in a range of AI tools, such as ChatGPT.",
        "Share knowledge and insights regarding AI through a community of practice.",
        "Discuss AI and the use of LLMs, like ChatGPT, in the context of the unit, professional ethics and integrity.",
        "Articulate acceptable use of AI in any assigned learning activity or Unit assessment.",
        "Adhere to the CDU Student Academic Integrity Policy.",
    ]
    for item in ai_items:
        add_bullet(doc, item)
    
    styled_paragraph(doc, "", space_after=4)
    styled_paragraph(doc, "CDU will not tolerate the following:", bold=True, space_after=4)
    ai_no = [
        "Misrepresenting AI generated content as your own in teaching, learning or research.",
        "Failure to disclose the use of AI.",
        "Use of third-party resources or AI in any other form that compromises academic integrity.",
        "Contract cheating in any form.",
    ]
    for item in ai_no:
        add_bullet(doc, item)
    
    styled_paragraph(doc, "", space_after=4)
    styled_paragraph(doc, "Team/unit specific AI statement:", bold=True, space_after=4)
    styled_paragraph(
        doc,
        assessment_config.get("ai_statement", "[Insert team/unit specific AI statement here]"),
        italic=True, space_after=8
    )
    
    # ══════════════════════════════════════════════════════════════
    # ASSESSMENT TASK SECTIONS
    # ══════════════════════════════════════════════════════════════
    for task in tasks:
        doc.add_page_break()
        _build_assessment_task_section(doc, unit_data, task, assessment_config)
    
    # ══════════════════════════════════════════════════════════════
    # ASSESSMENT MAPPING MATRIX
    # ══════════════════════════════════════════════════════════════
    doc.add_page_break()
    _build_mapping_matrix(doc, unit_data, tasks, assessment_config)
    
    # ══════════════════════════════════════════════════════════════
    # DOCUMENT APPROVAL RECORD
    # ══════════════════════════════════════════════════════════════
    doc.add_page_break()
    add_heading_styled(doc, "Document Approval and Validation Record")
    
    table = doc.add_table(rows=4, cols=2)
    table.style = 'Table Grid'
    labels = ["Document title", "Date of pre-validation", "Approval by Team Leader", "Date of next review"]
    values = [
        f"Assessor Guide - {unit_data.code}",
        datetime.now().strftime("%d/%m/%Y"),
        "",
        ""
    ]
    for i, (label, value) in enumerate(zip(labels, values)):
        table.rows[i].cells[0].text = label
        table.rows[i].cells[1].text = value
        for cell in table.rows[i].cells:
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.name = FONT_BODY
                    run.font.size = PT_BODY
    
    # ── Footer info ──
    styled_paragraph(doc, "", space_after=12)
    styled_paragraph(
        doc,
        f"Assessor Guide | Academic Quality & Integrity (AQ&I) | {datetime.now().strftime('%B %Y')}",
        size=Pt(8), color=COLOR_GREY, alignment=WD_ALIGN_PARAGRAPH.CENTER
    )
    
    # ── Save ──
    doc.save(output_path)
    return output_path


def _build_assessment_task_section(doc, unit_data, task, config):
    """Build an individual assessment task section (AT#)."""
    task_num = task.get("number", 1)
    task_type = task.get("type", "questioning")
    task_method = task.get("method", "Questioning (written or verbal)")
    task_title = task.get("title", "")
    
    # Section header
    type_labels = {
        "questioning": "Questioning",
        "observation": "Direct Observation",
        "project": "Project",
        "portfolio": "Portfolio",
        "third_party": "Third Party Report"
    }
    type_label = type_labels.get(task_type, task_type.title())
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"AT{task_num} {type_label}")
    run.font.name = FONT_HEADING
    run.font.size = Pt(18)
    run.font.bold = True
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"{unit_data.code} {unit_data.title}")
    run.font.name = FONT_HEADING
    run.font.size = Pt(12)
    run.font.bold = True
    run.font.color.rgb = COLOR_CDU_DARK
    p.paragraph_format.space_after = Pt(12)
    
    # Assessment Details table
    styled_paragraph(doc, "Assessment Details", bold=True, size=PT_HEADING2, space_after=6)
    
    table = doc.add_table(rows=0, cols=2)
    table.style = 'Table Grid'
    
    # Task Description
    _add_detail_row(table, "Task Description", task.get("description", _get_default_description(task_type)))
    _add_detail_row(table, "Assessment Method", task_method)
    _add_detail_row(table, "Assessment Instrument", f"{unit_data.code} AT{task_num} {type_label} v1")
    
    # Context
    context = task.get("context", "")
    if not context:
        context = "☐ Simulated workplace  ☐ Active workplace  ☐ Training Room  ☐ Own Environment"
    _add_detail_row(table, "Context of Assessment", context)
    
    # Evidence
    evidence = task.get("evidence", _get_default_evidence(task_type))
    _add_detail_row(table, "Evidence to be submitted", evidence)
    
    for row in table.rows:
        row.cells[0].width = Inches(1.8)
        row.cells[1].width = Inches(4.7)
    
    # Additional sections based on task type
    styled_paragraph(doc, "", space_after=6)
    
    if task_type == "questioning":
        _build_questioning_benchmarks(doc, unit_data, task, config)
    elif task_type == "observation":
        _build_observation_section(doc, unit_data, task)
    elif task_type == "third_party":
        _build_third_party_section(doc, unit_data, task)
    else:
        styled_paragraph(doc, "Benchmark Guide (answer guide)", bold=True, size=PT_HEADING2, space_after=6)
        styled_paragraph(doc, "[Develop benchmark criteria for this assessment task]", italic=True)


def _add_detail_row(table, label, value):
    """Add a label-value row to a details table."""
    row = table.add_row()
    p = row.cells[0].paragraphs[0]
    run = p.add_run(label)
    run.font.bold = True
    run.font.name = FONT_BODY
    run.font.size = PT_BODY
    
    p2 = row.cells[1].paragraphs[0]
    run2 = p2.add_run(value)
    run2.font.name = FONT_BODY
    run2.font.size = PT_BODY


def _get_default_description(task_type):
    descriptions = {
        "questioning": "This assessment task asks you to answer a series of written questions. These questions are designed to show your knowledge related to the unit of competency while following workplace policies, procedures, and processes.",
        "observation": "For this assessment, you will be observed performing workplace tasks to demonstrate your practical skills and competency.",
        "project": "This assessment requires you to complete a project demonstrating your practical application of skills and knowledge.",
        "portfolio": "Collate the listed workplace documents and submit them to the Assessor.",
        "third_party": "Gather workplace evidence from the direct Supervisor of actual workplace performance."
    }
    return descriptions.get(task_type, "")


def _get_default_evidence(task_type):
    evidence = {
        "questioning": "The completion and submission of the identified questionnaire assessment, along with any additional information or documents outlined in the questions.",
        "observation": "Completed observation checklist signed by the Assessor.",
        "project": "Complete project documentation as specified.",
        "portfolio": "Collated workplace documents in PDF or Word format.",
        "third_party": "Signed Third Party Report."
    }
    return evidence.get(task_type, "")


def _build_questioning_benchmarks(doc, unit_data, task, config):
    """Build the questioning section with benchmark answers from knowledge evidence."""
    styled_paragraph(doc, "Benchmark Guide (answer guide)", bold=True, size=PT_HEADING2, space_after=6)
    
    questions = task.get("questions", [])
    
    if not questions and unit_data.knowledge_evidence:
        # Auto-generate question structure from knowledge evidence
        styled_paragraph(
            doc,
            "The following questions are mapped to the Knowledge Evidence requirements. "
            "Benchmark answers are provided as a guide for the Assessor.",
            italic=True, space_after=8
        )
        
        q_num = 1
        for ke in unit_data.knowledge_evidence:
            text = ke.text if hasattr(ke, 'text') else str(ke)
            if text and len(text) > 10:
                table = doc.add_table(rows=2, cols=2)
                table.style = 'Table Grid'
                
                # Question row
                q_cell = table.rows[0].cells[0]
                p = q_cell.paragraphs[0]
                run = p.add_run(f"Q{q_num}")
                run.font.bold = True
                run.font.name = FONT_BODY
                run.font.size = PT_BODY
                
                q_text_cell = table.rows[0].cells[1]
                p = q_text_cell.paragraphs[0]
                run = p.add_run(f"Describe/explain: {text}")
                run.font.name = FONT_BODY
                run.font.size = PT_BODY
                
                # Benchmark row
                b_cell = table.rows[1].cells[0]
                b_cell.merge(table.rows[1].cells[1])
                p = b_cell.paragraphs[0]
                run = p.add_run("[Assessor benchmark answer - to be completed]")
                run.font.name = FONT_BODY
                run.font.size = PT_BODY
                run.font.italic = True
                run.font.color.rgb = COLOR_GREY
                
                # Add S/US checkboxes
                p2 = b_cell.add_paragraph()
                run2 = p2.add_run("☐ S   ☐ US")
                run2.font.name = FONT_BODY
                run2.font.size = PT_BODY
                
                styled_paragraph(doc, "", space_after=4)
                q_num += 1
    elif questions:
        styled_paragraph(
            doc,
            "The following questions are mapped to the Knowledge Evidence requirements. "
            "Benchmark answers are provided as a guide for the Assessor.",
            italic=True, space_after=8
        )
        
        for q in questions:
            table = doc.add_table(rows=2, cols=2)
            table.style = 'Table Grid'
            
            # Question number cell
            q_cell = table.rows[0].cells[0]
            p = q_cell.paragraphs[0]
            run = p.add_run(f"Q{q.get('number', '')}")
            run.font.bold = True
            run.font.name = FONT_BODY
            run.font.size = PT_BODY
            
            # Question text cell — support both "text" and "question" keys
            q_text_cell = table.rows[0].cells[1]
            q_text = q.get("question", "") or q.get("text", "")
            p = q_text_cell.paragraphs[0]
            run = p.add_run(q_text)
            run.font.name = FONT_BODY
            run.font.size = PT_BODY
            
            # Benchmark answer
            b_cell = table.rows[1].cells[0]
            b_cell.merge(table.rows[1].cells[1])
            benchmark = q.get("benchmark", "[Benchmark answer - to be completed]")
            p = b_cell.paragraphs[0]
            run = p.add_run(benchmark)
            run.font.name = FONT_BODY
            run.font.size = PT_BODY
            run.font.italic = True
            if benchmark.startswith("["):
                run.font.color.rgb = COLOR_GREY
            
            # S/US checkboxes
            p2 = b_cell.add_paragraph()
            run2 = p2.add_run("☐ S   ☐ US")
            run2.font.name = FONT_BODY
            run2.font.size = PT_BODY
            
            styled_paragraph(doc, "", space_after=4)
    else:
        styled_paragraph(doc, "[Develop questions and benchmark answers here]", italic=True)


def _build_observation_section(doc, unit_data, task):
    """Build the direct observation section."""
    styled_paragraph(doc, "Performance Evidence to be collected during this task", bold=True, space_after=6)
    
    if unit_data.performance_evidence:
        for pe in unit_data.performance_evidence:
            add_bullet(doc, pe)
    else:
        styled_paragraph(doc, "[List performance evidence from training.gov.au]", italic=True)
    
    styled_paragraph(doc, "", space_after=8)
    styled_paragraph(doc, "Benchmark Guide (answer guide)", bold=True, size=PT_HEADING2, space_after=6)
    styled_paragraph(
        doc,
        f"The {unit_data.code} AT{task.get('number', '')} Direct Observation is the Benchmark "
        "document for this assessment.",
        space_after=4
    )
    styled_paragraph(
        doc,
        "NOTE: The observable tasks and their benchmarks will be included in the AT document. "
        "Do not provide the full observation checklist to the student.",
        italic=True
    )


def _build_third_party_section(doc, unit_data, task):
    """Build the third party report section."""
    styled_paragraph(doc, "Additional Assessor Instructions", bold=True, space_after=6)
    styled_paragraph(doc, "Assessor to ensure that students have all workplace requirements.", space_after=4)
    styled_paragraph(doc, "Assessor to ensure the following:", space_after=4)
    add_bullet(doc, "The workplace is safe.")
    add_bullet(doc, "The Student has had a suitable workplace induction (if it is a work placement).")
    
    styled_paragraph(doc, "", space_after=8)
    styled_paragraph(doc, "Benchmark Guide (answer guide)", bold=True, size=PT_HEADING2, space_after=6)
    styled_paragraph(doc, "N/A", space_after=4)


def _build_mapping_matrix(doc, unit_data, tasks, config=None):
    """Build the assessment mapping matrix with auto-mapped question references."""
    add_heading_styled(doc, "Assessment Mapping Matrix")
    
    styled_paragraph(
        doc,
        "This assessment mapping matrix outlines the assessment strategy for the unit "
        "of competence in line with Training Package requirements.",
        space_after=8
    )
    
    num_tasks = len(tasks)
    num_cols = 2 + num_tasks  # Requirements + Student Resource + one per AT
    
    # ── Build question → KE mapping from AI questions ──
    # For each KE text, find which Q numbers cover it
    ai_questions = []
    questioning_task_idx = None  # Which AT column is the questioning task
    observation_task_idx = None
    third_party_task_idx = None
    
    for i, task in enumerate(tasks):
        if task["type"] == "questioning":
            questioning_task_idx = i
            ai_questions = task.get("questions", [])
        elif task["type"] == "observation":
            observation_task_idx = i
        elif task["type"] == "third_party":
            third_party_task_idx = i
    
    def _match_ke_to_questions(ke_text: str) -> str:
        """Find which question numbers cover a given KE item."""
        if not ai_questions:
            return ""
        ke_lower = ke_text.lower().strip()
        matched = []
        for q in ai_questions:
            # Check ke_items field
            q_ke_items = q.get("ke_items", [])
            for item in q_ke_items:
                if isinstance(item, str):
                    # Fuzzy match: check if KE text is contained in the item or vice versa
                    item_lower = item.lower().strip()
                    if (ke_lower in item_lower or item_lower in ke_lower or
                        _fuzzy_overlap(ke_lower, item_lower)):
                        matched.append(f"Q{q['number']}")
                        break
            else:
                # Also check if the question text references this KE topic
                q_text = q.get("question", "").lower()
                q_bench = q.get("benchmark", "").lower()
                # Extract key words from KE (3+ letter words)
                ke_words = [w for w in ke_lower.split() if len(w) > 3]
                if ke_words:
                    hits = sum(1 for w in ke_words if w in q_text or w in q_bench)
                    if hits >= len(ke_words) * 0.6:  # 60% of key words match
                        matched.append(f"Q{q['number']}")
        
        # Deduplicate and sort
        seen = set()
        unique = []
        for m in matched:
            if m not in seen:
                seen.add(m)
                unique.append(m)
        return ", ".join(unique[:4])  # Cap at 4 references per row
    
    def _fuzzy_overlap(a: str, b: str) -> bool:
        """Check if two strings share enough significant words."""
        words_a = set(w for w in a.split() if len(w) > 3)
        words_b = set(w for w in b.split() if len(w) > 3)
        if not words_a or not words_b:
            return False
        overlap = words_a & words_b
        return len(overlap) >= min(len(words_a), len(words_b)) * 0.5
    
    def _set_cell_text(cell, text, bold=False, size=Pt(8)):
        """Set cell text with consistent formatting."""
        p = cell.paragraphs[0]
        p.clear()
        run = p.add_run(text)
        run.font.name = FONT_BODY
        run.font.size = size
        run.font.bold = bold
    
    # ══════════════════════════════════════════════════════════════
    # Elements & Performance Criteria
    # ══════════════════════════════════════════════════════════════
    styled_paragraph(doc, "Elements and Performance Criteria", bold=True, space_after=6)
    
    if unit_data.elements:
        table = doc.add_table(rows=1, cols=num_cols)
        table.style = 'Table Grid'
        
        # Header
        _set_cell_text(table.rows[0].cells[0], "Mandatory Unit Requirements", bold=True)
        set_cell_shading(table.rows[0].cells[0], COLOR_TABLE_HEADER)
        _set_cell_text(table.rows[0].cells[1], "Student Resource Reference", bold=True)
        set_cell_shading(table.rows[0].cells[1], COLOR_TABLE_HEADER)
        for i, task in enumerate(tasks):
            cell = table.rows[0].cells[2 + i]
            label = f"Assessment {task.get('number', i+1)}\n{task.get('method', '')[:30]}"
            _set_cell_text(cell, label, bold=True)
            set_cell_shading(cell, COLOR_TABLE_HEADER)
        
        for element in unit_data.elements:
            row = table.add_row()
            _set_cell_text(row.cells[0], f"Element {element.number}: {element.title}", bold=True)
            set_cell_shading(row.cells[0], COLOR_LIGHT_GREY)
            
            for pc in element.performance_criteria:
                row = table.add_row()
                _set_cell_text(row.cells[0], f"{pc.number} {pc.text}")
                
                # Auto-map: observation tasks cover all PCs
                if observation_task_idx is not None:
                    _set_cell_text(row.cells[2 + observation_task_idx], "R")
                
                # Third party can cover all PCs too
                if third_party_task_idx is not None:
                    _set_cell_text(row.cells[2 + third_party_task_idx], "R")
    else:
        styled_paragraph(doc, "[Elements and Performance Criteria will be populated from training.gov.au data]", italic=True)
    
    # ══════════════════════════════════════════════════════════════
    # Foundation Skills
    # ══════════════════════════════════════════════════════════════
    styled_paragraph(doc, "", space_after=8)
    styled_paragraph(doc, "Foundation Skills", bold=True, space_after=6)
    
    if unit_data.foundation_skills:
        table = doc.add_table(rows=1, cols=num_cols + 1)  # +1 for Description
        table.style = 'Table Grid'
        
        headers = ["Skill", "Description", "Embedded into PC"]
        for i, task in enumerate(tasks):
            headers.append(f"AT{task.get('number', i+1)}")
        
        for i, h in enumerate(headers):
            if i < len(table.rows[0].cells):
                _set_cell_text(table.rows[0].cells[i], h, bold=True)
                set_cell_shading(table.rows[0].cells[i], COLOR_TABLE_HEADER)
        
        for fs in unit_data.foundation_skills:
            row = table.add_row()
            _set_cell_text(row.cells[0], fs.skill)
            _set_cell_text(row.cells[1], fs.description)
            _set_cell_text(row.cells[2], fs.embedded_pc if hasattr(fs, 'embedded_pc') else "")
    else:
        styled_paragraph(doc, "[Foundation Skills will be populated from training.gov.au data]", italic=True)
    
    # ══════════════════════════════════════════════════════════════
    # Knowledge Evidence — with auto-mapped Q numbers
    # ══════════════════════════════════════════════════════════════
    styled_paragraph(doc, "", space_after=8)
    styled_paragraph(doc, "Knowledge Evidence", bold=True, space_after=6)
    
    if unit_data.knowledge_evidence:
        table = doc.add_table(rows=1, cols=num_cols)
        table.style = 'Table Grid'
        
        _set_cell_text(table.rows[0].cells[0], "Knowledge Evidence", bold=True)
        set_cell_shading(table.rows[0].cells[0], COLOR_TABLE_HEADER)
        _set_cell_text(table.rows[0].cells[1], "Student Resource", bold=True)
        set_cell_shading(table.rows[0].cells[1], COLOR_TABLE_HEADER)
        for i, task in enumerate(tasks):
            cell = table.rows[0].cells[2 + i]
            _set_cell_text(cell, f"AT{task.get('number', i+1)}", bold=True)
            set_cell_shading(cell, COLOR_TABLE_HEADER)
        
        for ke in unit_data.knowledge_evidence:
            text = ke.text if hasattr(ke, 'text') else str(ke)
            row = table.add_row()
            _set_cell_text(row.cells[0], text)
            
            # Auto-map questions to KE
            if questioning_task_idx is not None and ai_questions:
                q_refs = _match_ke_to_questions(text)
                if q_refs:
                    _set_cell_text(row.cells[2 + questioning_task_idx], q_refs)
    
    # ══════════════════════════════════════════════════════════════
    # Performance Evidence — with observation references
    # ══════════════════════════════════════════════════════════════
    styled_paragraph(doc, "", space_after=8)
    styled_paragraph(doc, "Performance Evidence", bold=True, space_after=6)
    
    if unit_data.performance_evidence:
        table = doc.add_table(rows=1, cols=num_cols)
        table.style = 'Table Grid'
        
        _set_cell_text(table.rows[0].cells[0], "Performance Evidence", bold=True)
        set_cell_shading(table.rows[0].cells[0], COLOR_TABLE_HEADER)
        _set_cell_text(table.rows[0].cells[1], "Student Resource", bold=True)
        set_cell_shading(table.rows[0].cells[1], COLOR_TABLE_HEADER)
        for i, task in enumerate(tasks):
            cell = table.rows[0].cells[2 + i]
            _set_cell_text(cell, f"AT{task.get('number', i+1)}", bold=True)
            set_cell_shading(cell, COLOR_TABLE_HEADER)
        
        for pi, pe in enumerate(unit_data.performance_evidence):
            row = table.add_row()
            _set_cell_text(row.cells[0], pe)
            
            # Observation covers all PE
            if observation_task_idx is not None:
                _set_cell_text(row.cells[2 + observation_task_idx], "R")
            
            # Third party covers all PE  
            if third_party_task_idx is not None:
                _set_cell_text(row.cells[2 + third_party_task_idx], "R")
    
    # ══════════════════════════════════════════════════════════════
    # Assessment Conditions
    # ══════════════════════════════════════════════════════════════
    styled_paragraph(doc, "", space_after=8)
    styled_paragraph(doc, "Assessment Conditions", bold=True, space_after=6)
    
    if unit_data.assessment_conditions:
        table = doc.add_table(rows=2, cols=2)
        table.style = 'Table Grid'
        _set_cell_text(table.rows[0].cells[0], "Assessment Conditions", bold=True)
        set_cell_shading(table.rows[0].cells[0], COLOR_TABLE_HEADER)
        _set_cell_text(table.rows[0].cells[1], "How addressed", bold=True)
        set_cell_shading(table.rows[0].cells[1], COLOR_TABLE_HEADER)
        
        _set_cell_text(table.rows[1].cells[0], unit_data.assessment_conditions)
        _set_cell_text(table.rows[1].cells[1], "[Explain how the Assessment Conditions have been addressed]")
