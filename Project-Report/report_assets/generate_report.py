from __future__ import annotations

import textwrap
from datetime import date
from pathlib import Path
import sys

import mysql.connector
from PIL import Image, ImageDraw, ImageFont
from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[2]
ASSETS = Path(__file__).resolve().parent
OUTPUT = ROOT / "Project-Report" / "Campus_Maintenance_DBMS_Project_Report.docx"
sys.path.insert(0, str(ROOT / "application"))

from config import get_db_config
TABLES = [
    "buildings", "rooms", "users", "categories", "priorities", "technicians",
    "maintenance_requests", "assignments", "work_logs", "materials",
    "material_usage", "status_history", "feedback",
]
BLACK = (0, 0, 0)
WHITE = (255, 255, 255)
FONT = r"C:\Windows\Fonts\times.ttf"
FONT_BOLD = r"C:\Windows\Fonts\timesbd.ttf"
TABLE_COUNTER = 0


TABLE_DESCRIPTIONS = {
    "buildings": "Campus building/location master data.",
    "rooms": "Rooms belonging to a campus building.",
    "users": "People who report requests or perform administrative actions.",
    "categories": "Maintenance request classification values.",
    "priorities": "Priority levels and their response targets.",
    "technicians": "Technicians and their specialization/availability.",
    "maintenance_requests": "Core reported facility-maintenance incidents.",
    "assignments": "The technician assignment for a request; request_id is unique.",
    "work_logs": "Technician work entries for maintenance requests.",
    "materials": "Materials and their standard units and costs.",
    "material_usage": "Materials/quantities consumed by requests.",
    "status_history": "Status transitions, time, and the user who changed status.",
    "feedback": "One rating/comment record per request.",
}

COLUMN_DESCRIPTIONS = {
    "buildings": {
        "building_id": "Surrogate identifier for a building.",
        "building_name": "Unique campus building name.",
        "location": "Campus or physical location label.",
    },
    "rooms": {
        "room_id": "Surrogate identifier for a room.",
        "building_id": "Parent building foreign key.",
        "room_number": "Room number within its building.",
        "room_type": "Functional room type; defaults to Classroom.",
    },
    "users": {
        "user_id": "Surrogate identifier for a user.",
        "full_name": "User's display name.",
        "email": "Unique user email address.",
        "phone": "Optional user phone number.",
        "role": "Student, Faculty, Staff, or Admin; defaults to Student.",
    },
    "categories": {
        "category_id": "Surrogate identifier for a category.",
        "category_name": "Unique request category label.",
        "description": "Optional explanation of the category.",
    },
    "priorities": {
        "priority_id": "Surrogate identifier for a priority.",
        "priority_name": "Unique priority label.",
        "response_hours": "Positive target response time in hours.",
    },
    "technicians": {
        "technician_id": "Surrogate identifier for a technician.",
        "technician_name": "Technician's name.",
        "specialization": "Technician's maintenance specialization.",
        "phone": "Optional unique technician phone number.",
        "status": "Available, Busy, or Inactive; defaults to Available.",
    },
    "maintenance_requests": {
        "request_id": "Auto-increment internal primary key.",
        "request_number": "Unique human-readable request reference.",
        "user_id": "Reporting user foreign key.",
        "room_id": "Affected room foreign key.",
        "category_id": "Request category foreign key.",
        "priority_id": "Priority foreign key.",
        "title": "Short incident title, maximum 150 characters.",
        "description": "Detailed report of the issue.",
        "request_date": "Report timestamp; defaults to current timestamp.",
        "status": "Workflow state; defaults to Pending.",
        "assigned_date": "Optional time the request was assigned.",
        "closed_date": "Optional time the request was closed.",
        "closure_details": "Optional closure note, maximum 500 characters.",
    },
    "assignments": {
        "assignment_id": "Auto-increment assignment primary key.",
        "request_id": "Unique request foreign key, limiting each request to one assignment.",
        "technician_id": "Assigned technician foreign key.",
        "assigned_at": "Assignment timestamp; defaults to current timestamp.",
        "assignment_status": "Assignment state; defaults to Assigned.",
    },
    "work_logs": {
        "work_log_id": "Auto-increment work-log primary key.",
        "request_id": "Related maintenance request foreign key.",
        "technician_id": "Technician who recorded the work.",
        "work_date": "Work timestamp; defaults to current timestamp.",
        "work_description": "Work performed, maximum 500 characters.",
        "hours_spent": "Hours recorded to two decimal places; non-negative.",
    },
    "materials": {
        "material_id": "Surrogate identifier for a material.",
        "material_name": "Unique material name.",
        "unit": "Unit of measurement.",
        "unit_cost": "Non-negative standard cost per unit.",
    },
    "material_usage": {
        "usage_id": "Auto-increment usage primary key.",
        "request_id": "Maintenance request foreign key.",
        "material_id": "Material foreign key.",
        "quantity": "Quantity used to two decimal places; greater than zero.",
        "used_date": "Usage timestamp; defaults to current timestamp.",
    },
    "status_history": {
        "history_id": "Auto-increment history primary key.",
        "request_id": "Maintenance request foreign key.",
        "old_status": "Prior status; NULL for an initial event.",
        "new_status": "Resulting status after the transition.",
        "changed_at": "Change timestamp; defaults to current timestamp.",
        "changed_by": "User foreign key for the actor.",
    },
    "feedback": {
        "feedback_id": "Auto-increment feedback primary key.",
        "request_id": "Unique request foreign key; one feedback row per request.",
        "rating": "Integer rating from 1 through 5.",
        "comments": "Optional feedback comment, maximum 500 characters.",
        "feedback_date": "Feedback timestamp; defaults to current timestamp.",
    },
}

CHECK_CONSTRAINTS = {
    ("users", "role"): "CHECK role IN (Student, Faculty, Staff, Admin)",
    ("priorities", "response_hours"): "CHECK response_hours > 0",
    ("technicians", "status"): "CHECK status IN (Available, Busy, Inactive)",
    ("assignments", "assignment_status"): "CHECK assignment_status IN (Assigned, In Progress, Completed)",
    ("work_logs", "hours_spent"): "CHECK hours_spent >= 0",
    ("materials", "unit_cost"): "CHECK unit_cost >= 0",
    ("material_usage", "quantity"): "CHECK quantity > 0",
    ("feedback", "rating"): "CHECK rating BETWEEN 1 AND 5",
}


def db_rows(query: str, params: tuple = ()) -> list[dict]:
    connection = mysql.connector.connect(**get_db_config())
    cursor = connection.cursor(dictionary=True)
    try:
        cursor.execute(query, params)
        return cursor.fetchall()
    finally:
        cursor.close()
        connection.close()


def db_one(query: str, params: tuple = ()) -> dict:
    rows = db_rows(query, params)
    return rows[0] if rows else {}


def load_schema() -> tuple[dict, dict, dict]:
    placeholders = ", ".join(["%s"] * len(TABLES))
    columns = db_rows(
        f"""
        SELECT TABLE_NAME, COLUMN_NAME, COLUMN_TYPE, IS_NULLABLE,
               COLUMN_DEFAULT, EXTRA, COLUMN_KEY
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = %s AND TABLE_NAME IN ({placeholders})
        ORDER BY TABLE_NAME, ORDINAL_POSITION
        """,
        (get_db_config()["database"], *TABLES),
    )
    foreign_keys = db_rows(
        f"""
        SELECT TABLE_NAME, COLUMN_NAME, REFERENCED_TABLE_NAME, REFERENCED_COLUMN_NAME
        FROM information_schema.KEY_COLUMN_USAGE
        WHERE TABLE_SCHEMA = %s AND TABLE_NAME IN ({placeholders})
          AND REFERENCED_TABLE_NAME IS NOT NULL
        """,
        (get_db_config()["database"], *TABLES),
    )
    unique_constraints = db_rows(
        f"""
        SELECT k.TABLE_NAME, k.CONSTRAINT_NAME, k.COLUMN_NAME, k.ORDINAL_POSITION
        FROM information_schema.KEY_COLUMN_USAGE k
        JOIN information_schema.TABLE_CONSTRAINTS t
          ON t.TABLE_SCHEMA = k.TABLE_SCHEMA
         AND t.TABLE_NAME = k.TABLE_NAME
         AND t.CONSTRAINT_NAME = k.CONSTRAINT_NAME
        WHERE k.TABLE_SCHEMA = %s AND k.TABLE_NAME IN ({placeholders})
          AND t.CONSTRAINT_TYPE = 'UNIQUE'
        ORDER BY k.TABLE_NAME, k.CONSTRAINT_NAME, k.ORDINAL_POSITION
        """,
        (get_db_config()["database"], *TABLES),
    )
    if len(columns) < 60:
        raise RuntimeError(f"Expected 13 populated tables in campus_maintenance; found {len(columns)} columns.")
    fk_map = {(row["TABLE_NAME"], row["COLUMN_NAME"]): row for row in foreign_keys}
    unique_map: dict[tuple[str, str], list[str]] = {}
    for row in unique_constraints:
        key = (row["TABLE_NAME"], row["CONSTRAINT_NAME"])
        unique_map.setdefault(key, []).append(row["COLUMN_NAME"])
    return (
        {table: [row for row in columns if row["TABLE_NAME"] == table] for table in TABLES},
        fk_map,
        unique_map,
    )


def font(size: int, bold: bool = False):
    return ImageFont.truetype(FONT_BOLD if bold else FONT, size=size)


def make_er_diagram(path: Path) -> None:
    width, height = 3000, 1900
    image = Image.new("RGB", (width, height), WHITE)
    draw = ImageDraw.Draw(image)
    title_font = font(42, True)
    head_font = font(27, True)
    body_font = font(22)
    small_font = font(20, True)
    draw.text((width // 2, 38), "CAMPUS MAINTENANCE DATABASE — ER MODEL", fill=BLACK, font=title_font, anchor="mt")

    boxes = {
        "buildings": (55, 150, "BUILDINGS", ["PK building_id", "building_name (unique)", "location"]),
        "rooms": (635, 150, "ROOMS", ["PK room_id", "FK building_id", "room_number, room_type"]),
        "requests": (1215, 150, "MAINTENANCE_REQUESTS", ["PK request_id; UNIQUE request_number", "FK user_id, room_id", "FK category_id, priority_id"]),
        "assignments": (1795, 150, "ASSIGNMENTS", ["PK assignment_id; UNIQUE request_id", "FK technician_id", "assigned_at, assignment_status"]),
        "technicians": (2375, 150, "TECHNICIANS", ["PK technician_id", "technician_name, specialization", "phone, status"]),
        "users": (55, 690, "USERS", ["PK user_id", "full_name, email (unique)", "phone, role"]),
        "categories": (635, 690, "CATEGORIES", ["PK category_id", "category_name (unique)", "description"]),
        "priorities": (1215, 690, "PRIORITIES", ["PK priority_id", "priority_name (unique)", "response_hours"]),
        "status_history": (1795, 690, "STATUS_HISTORY", ["PK history_id", "FK request_id, changed_by", "old_status, new_status, changed_at"]),
        "work_logs": (2375, 690, "WORK_LOGS", ["PK work_log_id", "FK request_id, technician_id", "work_date, hours_spent, description"]),
        "materials": (55, 1230, "MATERIALS", ["PK material_id", "material_name (unique)", "unit, unit_cost"]),
        "material_usage": (635, 1230, "MATERIAL_USAGE", ["PK usage_id", "FK request_id, material_id", "quantity, used_date"]),
        "feedback": (1215, 1230, "FEEDBACK", ["PK feedback_id; UNIQUE request_id", "FK request_id", "rating, comments, feedback_date"]),
    }
    box_w, box_h = 530, 235

    def center(key: str) -> tuple[int, int]:
        x, y, _, _ = boxes[key]
        return x + box_w // 2, y + box_h // 2

    relations = [
        ("buildings", "rooms", "1 : N"),
        ("rooms", "requests", "1 : N"),
        ("users", "requests", "1 : N"),
        ("categories", "requests", "1 : N"),
        ("priorities", "requests", "1 : N"),
        ("requests", "assignments", "1 : 0..1"),
        ("technicians", "assignments", "1 : N"),
        ("requests", "status_history", "1 : N"),
        ("users", "status_history", "1 : N"),
        ("requests", "work_logs", "1 : N"),
        ("technicians", "work_logs", "1 : N"),
        ("requests", "material_usage", "1 : N"),
        ("materials", "material_usage", "1 : N"),
        ("requests", "feedback", "1 : 0..1"),
    ]
    for first, second, _ in relations:
        x1, y1 = center(first)
        x2, y2 = center(second)
        draw.line((x1, y1, x2, y2), fill=BLACK, width=4)
    for key, (x, y, heading, lines) in boxes.items():
        draw.rounded_rectangle((x, y, x + box_w, y + box_h), radius=16, outline=BLACK, fill=WHITE, width=4)
        draw.line((x, y + 54, x + box_w, y + 54), fill=BLACK, width=3)
        draw.text((x + box_w // 2, y + 10), heading, fill=BLACK, font=head_font, anchor="mt")
        for index, line in enumerate(lines):
            draw.text((x + 16, y + 67 + index * 47), line, fill=BLACK, font=body_font)
    legend_y = 1550
    draw.text((100, legend_y), "Relationship cardinalities", fill=BLACK, font=head_font)
    relationship_text = [
        "Buildings 1:N Rooms; Rooms 1:N Requests; Users 1:N Requests and Status History.",
        "Categories/Priorities 1:N Requests; Requests 1:0..1 Assignment; Technicians 1:N Assignments and Work Logs.",
        "Requests 1:N Work Logs, Material Usage, and Status History; Materials 1:N Material Usage.",
        "Requests 1:0..1 Feedback. Material Usage resolves the many-to-many Request–Material relationship.",
    ]
    for index, line in enumerate(relationship_text):
        draw.text((100, legend_y + 55 + index * 42), line, fill=BLACK, font=small_font)
    image.save(path, format="PNG", dpi=(300, 300))


def make_architecture_diagram(path: Path) -> None:
    image = Image.new("RGB", (1500, 1150), WHITE)
    draw = ImageDraw.Draw(image)
    title = font(46, True)
    heading = font(32, True)
    body = font(27)
    draw.text((750, 35), "SYSTEM ARCHITECTURE", fill=BLACK, font=title, anchor="mt")
    items = [
        ("User Browser", "Navigation and HTML form interaction"),
        ("HTML5 / CSS3 / JavaScript / Jinja2", "Templates rendered by Flask"),
        ("Flask Application", "Routes, validation, reporting, and transactions"),
        ("mysql-connector-python", "Parameterized SQL over MySQL connection"),
        ("MySQL 8.0 — campus_maintenance", "13 relational tables; existing schema"),
    ]
    x1, x2 = 170, 1330
    top = 125
    box_h = 150
    for index, (label, note) in enumerate(items):
        y = top + index * 200
        draw.rounded_rectangle((x1, y, x2, y + box_h), radius=18, outline=BLACK, fill=WHITE, width=4)
        draw.text((750, y + 24), label, fill=BLACK, font=heading, anchor="mt")
        draw.text((750, y + 83), note, fill=BLACK, font=body, anchor="mt")
        if index < len(items) - 1:
            draw.line((750, y + box_h, 750, y + 190), fill=BLACK, width=5)
            draw.polygon([(736, y + 178), (764, y + 178), (750, y + 195)], fill=BLACK)
    image.save(path, format="PNG", dpi=(300, 300))


def make_workflow_diagram(path: Path) -> None:
    image = Image.new("RGB", (1800, 1400), WHITE)
    draw = ImageDraw.Draw(image)
    title = font(42, True)
    box_font = font(25, True)
    draw.text((900, 28), "REQUEST WORKFLOW", fill=BLACK, font=title, anchor="mt")
    labels = [
        "User reports request", "Select priority", "Assign technician", "Update status",
        "Record work log", "Record material usage", "Mark completed", "Close with details",
        "Collect feedback", "Review reports",
    ]
    positions = []
    x_positions = [100, 700, 1300]
    y_positions = [130, 430, 730, 1030]
    for idx, label in enumerate(labels):
        row, col = divmod(idx, 3)
        x = x_positions[col] if row % 2 == 0 else x_positions[2 - col]
        y = y_positions[row]
        positions.append((x, y))
    bw, bh = 400, 150
    for i in range(len(labels) - 1):
        x1, y1 = positions[i]
        x2, y2 = positions[i + 1]
        c1 = (x1 + bw // 2, y1 + bh // 2)
        c2 = (x2 + bw // 2, y2 + bh // 2)
        if y1 == y2:
            if c2[0] > c1[0]:
                start, end = (x1 + bw, c1[1]), (x2, c2[1])
                arrow = [(end[0] - 20, end[1] - 12), end, (end[0] - 20, end[1] + 12)]
            else:
                start, end = (x1, c1[1]), (x2 + bw, c2[1])
                arrow = [(end[0] + 20, end[1] - 12), end, (end[0] + 20, end[1] + 12)]
            draw.line((*start, *end), fill=BLACK, width=4)
            draw.polygon(arrow, fill=BLACK)
        else:
            start = (c1[0], y1 + bh) if y2 > y1 else (c1[0], y1)
            end = (c2[0], y2) if y2 > y1 else (c2[0], y2 + bh)
            mid_y = (start[1] + end[1]) // 2
            draw.line((start[0], start[1], start[0], mid_y, end[0], mid_y, end[0], end[1]), fill=BLACK, width=4)
            sign = 1 if end[1] > mid_y else -1
            draw.polygon([(end[0] - 12, end[1] - 20 * sign), (end[0] + 12, end[1] - 20 * sign), end], fill=BLACK)
    for (x, y), label in zip(positions, labels):
        draw.rounded_rectangle((x, y, x + bw, y + bh), radius=18, outline=BLACK, fill=WHITE, width=4)
        wrapped = textwrap.wrap(label, width=24)
        for line_no, line in enumerate(wrapped):
            draw.text((x + bw // 2, y + 46 + line_no * 38), line, fill=BLACK, font=box_font, anchor="mt")
    image.save(path, format="PNG", dpi=(300, 300))


def set_run(run, size=12, bold=False, italic=False):
    run.font.name = "Times New Roman"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.italic = italic
    run.font.color.rgb = RGBColor(0, 0, 0)
    return run


def set_cell_border(cell):
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = tc_pr.first_child_found_in("w:tcBorders")
    if borders is None:
        borders = OxmlElement("w:tcBorders")
        tc_pr.append(borders)
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        tag = f"w:{edge}"
        element = borders.find(qn(tag))
        if element is None:
            element = OxmlElement(tag)
            borders.append(element)
        element.set(qn("w:val"), "single")
        element.set(qn("w:sz"), "6")
        element.set(qn("w:space"), "0")
        element.set(qn("w:color"), "000000")


def set_cell_margins(cell, top=60, start=65, bottom=60, end=65):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for margin, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{margin}"))
        if node is None:
            node = OxmlElement(f"w:{margin}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def no_split(row):
    tr_pr = row._tr.get_or_add_trPr()
    tr_pr.append(OxmlElement("w:cantSplit"))


def repeat_table_header(row):
    tr_pr = row._tr.get_or_add_trPr()
    tbl_header = OxmlElement("w:tblHeader")
    tbl_header.set(qn("w:val"), "true")
    tr_pr.append(tbl_header)


def set_cell_text(cell, value, size=9.5, bold=False, align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.alignment = align
    paragraph.paragraph_format.space_before = Pt(0)
    paragraph.paragraph_format.space_after = Pt(0)
    paragraph.paragraph_format.line_spacing = 1.0
    set_run(paragraph.add_run(str(value)), size=size, bold=bold)
    set_cell_border(cell)
    set_cell_margins(cell)


def add_table(doc, headers, rows, widths=None, font_size=9.5, caption=None):
    global TABLE_COUNTER
    TABLE_COUNTER += 1
    add_caption(doc, caption or f"Table {TABLE_COUNTER}: {headers[0]} summary")
    table = doc.add_table(rows=1, cols=len(headers))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    table.style = "Table Grid"
    repeat_table_header(table.rows[0])
    no_split(table.rows[0])
    for col, heading in enumerate(headers):
        set_cell_text(table.rows[0].cells[col], heading, size=font_size, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
    for row_data in rows:
        row = table.add_row()
        no_split(row)
        for col, value in enumerate(row_data):
            set_cell_text(row.cells[col], value, size=font_size, align=WD_ALIGN_PARAGRAPH.LEFT)
    if widths:
        for row in table.rows:
            for col, width in enumerate(widths):
                row.cells[col].width = Inches(width)
    return table


def add_body(doc, text, align=WD_ALIGN_PARAGRAPH.JUSTIFY, after=6):
    para = doc.add_paragraph()
    para.alignment = align
    para.paragraph_format.space_before = Pt(0)
    para.paragraph_format.space_after = Pt(after)
    para.paragraph_format.line_spacing = 1.5
    set_run(para.add_run(text), 12)
    return para


def add_heading(doc, text, level=1):
    para = doc.add_paragraph()
    para.style = f"Heading {level}"
    para.alignment = WD_ALIGN_PARAGRAPH.LEFT
    para.paragraph_format.keep_with_next = True
    para.paragraph_format.space_before = Pt(6 if level == 1 else 3)
    para.paragraph_format.space_after = Pt(5)
    size = {1: 16, 2: 14, 3: 12}.get(level, 12)
    set_run(para.add_run(text), size=size, bold=True)
    return para


def add_caption(doc, caption):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_before = Pt(2)
    para.paragraph_format.space_after = Pt(6)
    para.paragraph_format.keep_with_next = True
    set_run(para.add_run(caption), size=10, italic=True)
    return para


def add_code(doc, caption, code):
    add_caption(doc, caption)
    table = doc.add_table(rows=1, cols=1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    cell = table.cell(0, 0)
    set_cell_border(cell)
    set_cell_margins(cell, top=90, start=110, bottom=90, end=110)
    cell.text = ""
    para = cell.paragraphs[0]
    para.paragraph_format.space_after = Pt(0)
    para.paragraph_format.line_spacing = 1.0
    for index, line in enumerate(textwrap.dedent(code).strip().splitlines()):
        if index:
            para.add_run().add_break()
        set_run(para.add_run(line), size=9)
    doc.add_paragraph().paragraph_format.space_after = Pt(0)


def add_figure(doc, image_path, caption, width=6.0):
    para = doc.add_paragraph()
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    para.paragraph_format.space_after = Pt(0)
    para.add_run().add_picture(str(image_path), width=Inches(width))
    add_caption(doc, caption)


def add_placeholder_cell(cell, number, title, explanation):
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_border(cell)
    set_cell_margins(cell, top=90, start=110, bottom=90, end=110)
    first = cell.paragraphs[0]
    first.alignment = WD_ALIGN_PARAGRAPH.CENTER
    first.paragraph_format.space_after = Pt(4)
    first.paragraph_format.line_spacing = 1.0
    set_run(first.add_run(f"[INSERT SCREENSHOT HERE\n{title}]"), size=10, bold=True)
    caption = cell.add_paragraph()
    caption.alignment = WD_ALIGN_PARAGRAPH.CENTER
    caption.paragraph_format.space_after = Pt(3)
    set_run(caption.add_run(f"Figure {number}: {title}"), size=10, italic=True)
    explanation_para = cell.add_paragraph()
    explanation_para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    explanation_para.paragraph_format.space_after = Pt(0)
    explanation_para.paragraph_format.line_spacing = 1.0
    set_run(explanation_para.add_run(explanation), size=9)


def add_screenshot_placeholders(doc):
    figures = [
        ("Dashboard", "Shows database-derived request totals, technician count, rating, and recent requests."),
        ("Create Maintenance Request", "Shows the four database-backed dropdowns and required title and description fields."),
        ("Maintenance Request List", "Shows request number, reporter, category, priority, status, building, and room, with filters."),
        ("Request Details", "Shows one request and its assignment, work, materials, status history, and feedback."),
        ("Technician Assignment", "Shows technician names, specializations, and availability/status for selection."),
        ("Status Update", "Shows the status form and the application-enforced allowed transitions."),
        ("Work Log Entry", "Shows technician, work description, and non-negative hours fields."),
        ("Material Usage", "Shows material selection and a positive quantity field."),
        ("Feedback", "Shows the rating selector and comments field for eligible completed/closed requests."),
        ("Reports", "Shows pending/ageing, technician, building, material cost, feedback, and response-time reports."),
    ]
    table = doc.add_table(rows=5, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False
    for row_index in range(5):
        for col_index in range(2):
            cell = table.rows[row_index].cells[col_index]
            cell.width = Inches(2.92)
            number = 2 + row_index * 2 + col_index
            title, explanation = figures[number - 2]
            add_placeholder_cell(cell, number, title, explanation)
    doc.add_paragraph(
        "Screenshot plan: replace each editable placeholder with a freshly captured screenshot from the local Flask application. "
        "No screenshots are claimed as embedded in this report.",
        style="Normal",
    )


def format_doc(doc):
    section = doc.sections[0]
    section.page_width = Inches(8.27)
    section.page_height = Inches(11.69)
    section.top_margin = Inches(1)
    section.bottom_margin = Inches(1)
    section.left_margin = Inches(1.25)
    section.right_margin = Inches(1)
    section.different_first_page_header_footer = True
    pg_num = OxmlElement("w:pgNumType")
    pg_num.set(qn("w:start"), "0")
    section._sectPr.append(pg_num)

    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    normal.font.size = Pt(12)
    normal.font.color.rgb = RGBColor(0, 0, 0)
    normal.paragraph_format.space_before = Pt(0)
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.5
    for style_name, size in (("Heading 1", 16), ("Heading 2", 14), ("Heading 3", 12)):
        style = doc.styles[style_name]
        style.font.name = "Times New Roman"
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.color.rgb = RGBColor(0, 0, 0)
        style.paragraph_format.keep_with_next = True
    footer = section.footer.paragraphs[0]
    footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer.paragraph_format.space_before = Pt(0)
    footer.paragraph_format.space_after = Pt(0)
    set_run(footer.add_run("Campus Facility Maintenance Request Management System  |  Page "), size=9)
    field = OxmlElement("w:fldSimple")
    field.set(qn("w:instr"), "PAGE")
    footer._p.append(field)
    doc.sections[0].first_page_footer.paragraphs[0].text = ""


def add_cover(doc):
    for _ in range(2):
        doc.add_paragraph()
    for line, size, bold in [
        ("PROJECT REPORT", 18, True),
        ("ON", 13, True),
        ("PROJECT 41", 15, True),
        ("DESIGN AND IMPLEMENTATION OF A DATABASE MANAGEMENT SYSTEM FOR", 15, True),
        ("CAMPUS FACILITY MAINTENANCE REQUEST MANAGEMENT SYSTEM", 15, True),
    ]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(9)
        set_run(para.add_run(line), size=size, bold=bold)
    doc.add_paragraph()
    for line in [
        "Submitted by:",
        "Pranav Arjun Nikam  |  25WU0101085",
        "Saathvik Mahalsa  |  25WU0101115",
        "Mithil D. Parikh  |  25WU0101072",
        "",
        "Course: Database Management Systems",
        "University: Woxsen University",
        "Faculty: [Faculty Name]",
        "Academic Year: 2026–2027",
    ]:
        para = doc.add_paragraph()
        para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        para.paragraph_format.space_after = Pt(7)
        set_run(para.add_run(line), size=12, bold=line in {"Submitted by:", "Course: Database Management Systems", "University: Woxsen University"})
    doc.add_page_break()


def add_data_dictionary(doc, schema, fk_map, unique_map):
    add_heading(doc, "8. DATA DICTIONARY", 1)
    add_body(
        doc,
        "The following dictionary is generated from INFORMATION_SCHEMA for the live campus_maintenance database. "
        "Column type, nullability, defaults, auto-increment attributes, keys, and checks reflect the existing MySQL 8.0 schema. "
        "Descriptions explain the role of each field; they do not add constraints.",
    )
    global TABLE_COUNTER
    TABLE_COUNTER += 1
    add_caption(doc, f"Table {TABLE_COUNTER}: Data Dictionary for the 13 Existing Tables")
    dictionary = doc.add_table(rows=1, cols=6)
    dictionary.alignment = WD_TABLE_ALIGNMENT.CENTER
    dictionary.autofit = False
    dictionary.style = "Table Grid"
    repeat_table_header(dictionary.rows[0])
    no_split(dictionary.rows[0])
    headers = ["Column Name", "Data Type", "Constraint", "Key", "Default", "Description"]
    widths = [0.86, 0.78, 1.34, 0.48, 0.72, 1.65]
    for column, heading in enumerate(headers):
        set_cell_text(dictionary.rows[0].cells[column], heading, size=8.6, bold=True, align=WD_ALIGN_PARAGRAPH.CENTER)
        dictionary.rows[0].cells[column].width = Inches(widths[column])
    for table in TABLES:
        group_row = dictionary.add_row()
        no_split(group_row)
        group_cell = group_row.cells[0].merge(group_row.cells[-1])
        set_cell_text(group_cell, f"{table.upper()} — {TABLE_DESCRIPTIONS[table]}", size=8.6, bold=True)
        rows = []
        for col in schema[table]:
            name = col["COLUMN_NAME"]
            key_parts = []
            constraint_parts = []
            if col["COLUMN_KEY"] == "PRI":
                key_parts.append("PK")
            fk = fk_map.get((table, name))
            if fk:
                key_parts.append("FK")
                constraint_parts.append(f"REFERENCES {fk['REFERENCED_TABLE_NAME']}.{fk['REFERENCED_COLUMN_NAME']}")
            if col["COLUMN_KEY"] == "UNI":
                key_parts.append("UNIQUE")
                constraint_parts.append("UNIQUE")
            for (unique_table, _), unique_cols in unique_map.items():
                if unique_table == table and len(unique_cols) > 1 and name in unique_cols:
                    constraint_parts.append("UNIQUE (" + ", ".join(unique_cols) + ")")
                    key_parts.append("UK")
                    break
            if col["IS_NULLABLE"] == "NO":
                constraint_parts.append("NOT NULL")
            check = CHECK_CONSTRAINTS.get((table, name))
            if check:
                constraint_parts.append(check)
            if col["EXTRA"] and "auto_increment" in col["EXTRA"].lower():
                constraint_parts.append("AUTO_INCREMENT")
            default = col["COLUMN_DEFAULT"]
            default_text = "—" if default is None else str(default)
            if default is not None:
                constraint_parts.append("DEFAULT")
            rows.append((
                name,
                col["COLUMN_TYPE"],
                "; ".join(dict.fromkeys(constraint_parts)) or "—",
                ", ".join(key_parts) or "—",
                default_text,
                COLUMN_DESCRIPTIONS[table].get(name, name.replace("_", " ").capitalize()),
            ))
        for values in rows:
            row = dictionary.add_row()
            no_split(row)
            for column, value in enumerate(values):
                set_cell_text(row.cells[column], value, size=8.6)
                row.cells[column].width = Inches(widths[column])


def main():
    if get_db_config()["database"] != "campus_maintenance":
        raise RuntimeError("Report generation must use the campus_maintenance database.")
    schema, fk_map, unique_map = load_schema()
    counts = db_one("SELECT COUNT(*) AS total FROM maintenance_requests")
    dashboard = db_one(
        """
        SELECT COUNT(*) AS total,
          SUM(status = 'Pending') AS pending,
          SUM(status = 'Assigned') AS assigned,
          SUM(status = 'In Progress') AS in_progress,
          SUM(status = 'Completed') AS completed,
          SUM(status = 'Closed') AS closed
        FROM maintenance_requests
        """
    )
    tech_count = db_one("SELECT COUNT(*) AS total FROM technicians")["total"]
    rating = db_one("SELECT ROUND(AVG(rating), 2) AS average FROM feedback")["average"]
    sample_requests = db_rows(
        """
        SELECT mr.request_id, mr.request_number, mr.title, u.full_name AS reported_by,
               c.category_name, p.priority_name, mr.status, b.building_name,
               r.room_number, mr.request_date
        FROM maintenance_requests mr
        LEFT JOIN users u ON u.user_id = mr.user_id
        LEFT JOIN rooms r ON r.room_id = mr.room_id
        LEFT JOIN buildings b ON b.building_id = r.building_id
        LEFT JOIN categories c ON c.category_id = mr.category_id
        LEFT JOIN priorities p ON p.priority_id = mr.priority_id
        WHERE mr.request_id BETWEEN 11 AND 15
        ORDER BY mr.request_id
        """
    )
    pending = db_rows(
        """
        SELECT mr.request_number, mr.title, c.category_name, p.priority_name,
               b.building_name, r.room_number, mr.status
        FROM maintenance_requests mr
        LEFT JOIN rooms r ON r.room_id = mr.room_id
        LEFT JOIN buildings b ON b.building_id = r.building_id
        LEFT JOIN categories c ON c.category_id = mr.category_id
        LEFT JOIN priorities p ON p.priority_id = mr.priority_id
        WHERE LOWER(COALESCE(mr.status, '')) = 'pending'
        ORDER BY mr.request_date DESC
        LIMIT 5
        """
    )
    technician_load = db_rows(
        """
        SELECT tech.technician_name, tech.specialization,
               COUNT(a.assignment_id) AS assigned_requests
        FROM technicians tech
        LEFT JOIN assignments a ON a.technician_id = tech.technician_id
        GROUP BY tech.technician_id, tech.technician_name, tech.specialization
        ORDER BY assigned_requests DESC, tech.technician_name
        """
    )
    material_costs = db_rows(
        """
        SELECT m.material_name, SUM(mu.quantity) AS quantity_used, m.unit_cost,
               SUM(mu.quantity * m.unit_cost) AS total_cost
        FROM material_usage mu
        LEFT JOIN materials m ON m.material_id = mu.material_id
        GROUP BY m.material_id, m.material_name, m.unit_cost
        ORDER BY total_cost DESC
        """
    )
    building_counts = db_rows(
        """
        SELECT b.building_name, COUNT(mr.request_id) AS request_count
        FROM buildings b
        LEFT JOIN rooms r ON r.building_id = b.building_id
        LEFT JOIN maintenance_requests mr ON mr.room_id = r.room_id
        GROUP BY b.building_id, b.building_name
        ORDER BY request_count DESC, b.building_name
        """
    )
    feedback_rows = db_rows(
        """
        SELECT mr.request_number, f.rating, f.comments
        FROM feedback f
        LEFT JOIN maintenance_requests mr ON mr.request_id = f.request_id
        ORDER BY f.feedback_date DESC
        """
    )
    ageing = db_rows(
        """
        SELECT mr.request_number, mr.title, mr.status,
               DATEDIFF(NOW(), mr.request_date) AS age_days
        FROM maintenance_requests mr
        WHERE LOWER(COALESCE(mr.status, '')) NOT IN ('completed', 'closed')
        ORDER BY mr.request_date ASC LIMIT 5
        """
    )
    response_times = db_rows(
        """
        SELECT mr.request_number, DATEDIFF(a.assigned_at, mr.request_date) AS response_days
        FROM maintenance_requests mr
        LEFT JOIN assignments a ON a.request_id = mr.request_id
        WHERE a.assigned_at IS NOT NULL AND mr.request_date IS NOT NULL
        ORDER BY response_days
        """
    )

    er_path = ASSETS / "er_diagram.png"
    architecture_path = ASSETS / "system_architecture.png"
    workflow_path = ASSETS / "application_workflow.png"
    make_er_diagram(er_path)
    make_architecture_diagram(architecture_path)
    make_workflow_diagram(workflow_path)

    doc = Document()
    format_doc(doc)
    add_cover(doc)

    add_heading(doc, "2. ABSTRACT", 1)
    abstract = (
        "Campus facilities depend on timely maintenance of classrooms, laboratories, offices, and shared services. "
        "When issue reporting is scattered across verbal requests or informal messages, records can be missed, duplicated, "
        "or difficult to trace. This project presents a database-backed Campus Facility Maintenance Request Management System "
        "that organizes reports and their follow-up in one relational application. A user records the affected room, category, "
        "priority, title, and description. The application assigns a unique request number while MySQL retains the internal "
        "auto-increment request identifier. The distinction is important: in the inspected sample, request numbers "
        "REQ-2026-001 through REQ-2026-005 correspond to database request IDs 11 through 15, not IDs 1 through 5.\n\n"
        "The database separates campus buildings and rooms, users, request categories, priorities, technicians, assignments, "
        "work logs, materials, material usage, status history, and feedback. These relations provide referential integrity "
        "through primary and foreign keys. Unique constraints prevent repeated request numbers, multiple technician assignments "
        "for the same request, and duplicate feedback for a request. Check constraints enforce domain rules for roles, technician "
        "availability, assignment status, work hours, material cost, usage quantity, priority response hours, and feedback rating. "
        "The web interface presents dashboard totals, searchable and filterable requests, request details, technician selection, "
        "status transitions, work and material entry, closure, feedback, and aggregate reports.\n\n"
        "The prototype uses Python 3.12 and Flask for routing and server-side validation, Jinja2 for HTML rendering, HTML5 and "
        "CSS3 for the interface, a small JavaScript file for defaulting date inputs, and mysql-connector-python to communicate "
        "with MySQL 8.0. Connection settings are read from environment variables loaded from a local .env file; the password is "
        "not included in this report. Application queries use parameter placeholders, and request creation, assignment, status "
        "changes, and closure use transaction handling. The existing campus_maintenance database was inspected and used as the "
        "source of truth; the application does not create or replace its schema. The resulting project demonstrates how a "
        "normalized relational design can support traceable maintenance operations, consistent reporting, and a straightforward "
        "demonstration workflow for a university campus."
    )
    for paragraph in abstract.split("\n\n"):
        add_body(doc, paragraph)

    add_heading(doc, "3. INTRODUCTION AND PROBLEM STATEMENT", 1)
    add_heading(doc, "3.1 Background", 2)
    add_body(doc, "A campus contains buildings, rooms, equipment, and services used by students, faculty, and staff. Routine wear, electrical faults, water leaks, furniture damage, and information-technology failures require a repeatable process for reporting and resolving problems. The work is operational: a request must identify where the problem occurred, what kind of work is needed, how urgent it is, who is responsible, what work was performed, and whether the reporter considers the outcome satisfactory.")
    add_body(doc, "A database management system can store these related facts in a consistent structure. In this project, master data such as buildings, rooms, users, categories, priorities, technicians, and materials are stored separately from transactional records such as maintenance requests, assignments, work logs, material usage, status history, and feedback. This separation allows application pages and reports to combine the necessary information through joins without repeating every descriptive field in each request.")
    add_heading(doc, "3.2 Existing Problem", 2)
    add_body(doc, "Informal reporting can make it difficult to identify whether a fault has already been reported, which technician has accepted it, or whether it is completed or closed. A list maintained independently from work records can drift out of date. Repeating building names, room details, or technician attributes in every maintenance row also increases the risk of inconsistent spellings and update anomalies. Managers additionally need concise views of open work, technician load, response time, and material expenditure.")
    add_heading(doc, "3.3 Problem Statement", 2)
    add_body(doc, "The problem addressed is the lack of a single traceable record linking a campus maintenance issue to its reporter, location, classification, urgency, assignment, work performed, material use, status transitions, and feedback. The system must maintain relationships among these records while presenting useful operational views and rejecting invalid or duplicate data.")
    add_heading(doc, "3.4 Proposed System", 2)
    add_body(doc, "The implemented prototype is a local Flask web application connected to the existing MySQL 8.0 database named campus_maintenance. It provides a dashboard, request entry, request search and filters, request details, technician assignment, status updates, work logs, material usage, closure, feedback, a technician listing, and reports. Request identifiers are generated by the database; a separate human-readable reference is allocated using the current year and the highest existing numeric suffix. The interface is server-rendered with Jinja2 templates and uses no front-end framework.")
    add_heading(doc, "3.5 Need for Database Management", 2)
    add_body(doc, "A relational database is appropriate because the information has explicit entities and stable relationships. Primary keys identify records independently of their display labels. Foreign keys prevent a request from referring to a nonexistent user, room, category, or priority. Unique and check constraints enforce additional rules at the data layer, even if input originates outside the web form. SQL joins and aggregate functions support building-level counts, technician workload, material cost, ageing, and feedback summaries. Transaction boundaries group related changes—for example, creating a request and its initial status-history record—so that partial writes are not committed when an operation fails.")

    add_heading(doc, "4. OBJECTIVES AND SCOPE", 1)
    add_heading(doc, "4.1 Objectives", 2)
    objectives = [
        "Centralize campus maintenance requests in the existing relational database.",
        "Associate each request with a reporter, room/building, category, and priority.",
        "Assign a request to a technician while respecting the one-assignment-per-request constraint.",
        "Record workflow status, status history, work logs, material usage, closure details, and feedback.",
        "Provide request search/filtering and a database-backed dashboard.",
        "Use joins and aggregates to present useful operational reports.",
        "Use environment-based database credentials and parameterized SQL.",
        "Demonstrate transaction handling and server-side validation in a beginner-friendly Flask application.",
    ]
    for index, item in enumerate(objectives, 1):
        para = doc.add_paragraph()
        para.paragraph_format.left_indent = Inches(0.25)
        para.paragraph_format.first_line_indent = Inches(-0.25)
        para.paragraph_format.line_spacing = 1.5
        para.paragraph_format.space_after = Pt(4)
        set_run(para.add_run(f"{index}. {item}"), 12)
    add_heading(doc, "4.2 Scope", 2)
    add_body(doc, "The implemented scope is a local, server-rendered web prototype for recording and tracking maintenance work against the existing 13-table MySQL schema. It supports records already present in the database and form-based updates through Flask routes.")
    add_heading(doc, "4.3 Functional Scope", 2)
    add_body(doc, "Included functions are request creation and listing; text search and status/category/priority filters; joined request details; technician assignment; validated workflow status updates; work-log and material-use entry; closure with details; feedback for completed or closed requests; technician display; and SQL reports. The dashboard and reports use live query results, not hard-coded sample totals.")
    add_heading(doc, "4.4 Limitations", 2)
    add_body(doc, "The current prototype has no login screen, authentication, or role-based authorization; the application selects an existing Admin user for certain history entries. It is intended for local demonstration, not public deployment. Notifications, APIs, mobile views beyond responsive CSS, asset tracking, scheduling, and advanced analytics are not implemented. Screenshot slots in Section 11 are editable placeholders because usable screenshot image files were not available to embed.")

    add_heading(doc, "5. SOFTWARE AND HARDWARE REQUIREMENTS", 1)
    add_heading(doc, "5.1 Software Requirements", 2)
    add_table(doc, ["Component", "Requirement / Role"], [
        ("Operating system", "Windows 11 development environment."),
        ("Runtime", "Python 3.12."),
        ("Web framework", "Flask 3.1.0."),
        ("Database server", "MySQL 8.0; schema/database campus_maintenance."),
        ("Connector", "mysql-connector-python 9.1.0."),
        ("Templates and UI", "Jinja2, HTML5, CSS3, vanilla JavaScript."),
        ("Configuration", "python-dotenv; DB_HOST, DB_USER, DB_PASSWORD, DB_NAME."),
        ("Editor and browser", "Visual Studio Code and a modern web browser."),
    ], widths=[1.55, 4.47], font_size=10)
    add_heading(doc, "5.2 Minimum Hardware (Prototype)", 2)
    add_table(doc, ["Resource", "Reasonable minimum"], [
        ("Processor", "Dual-core CPU capable of running Windows 11."),
        ("Memory", "8 GB RAM for Windows, editor, local MySQL, and browser."),
        ("Storage", "At least 2 GB free for source, dependencies, and database files."),
        ("Display", "1366 × 768 or higher for comfortable form/table viewing."),
        ("Network", "Localhost access; internet only for package installation/documentation."),
    ], widths=[1.55, 4.47], font_size=10)

    add_heading(doc, "6. ER DIAGRAM", 1)
    add_body(doc, "Figure 1 is derived from SHOW CREATE TABLE and INFORMATION_SCHEMA for the live MySQL schema. It contains the actual 13 entities and shows primary keys, important attributes, foreign-key columns, and cardinalities. There is no assets table in the implementation.")
    add_figure(doc, er_path, "Figure 1: Entity Relationship Diagram of the Campus Maintenance Management System", width=6.0)
    add_heading(doc, "6.1 Relationship Summary", 2)
    add_table(doc, ["Relationship", "Cardinality and implementation"], [
        ("Building–Room", "One building contains many rooms; rooms.building_id is an FK."),
        ("Room–Request", "One room can be referenced by many maintenance requests."),
        ("User–Request", "One user can report many requests."),
        ("Category/Priority–Request", "Each lookup value can classify many requests."),
        ("Request–Assignment", "A request has zero or one assignment; assignments.request_id is UNIQUE."),
        ("Technician–Assignment", "A technician may have multiple assignment records."),
        ("Request–Work Log", "A request can have multiple technician work-log entries."),
        ("Request–Material", "Many-to-many relationship resolved by material_usage."),
        ("Request–History/Feedback", "History is one-to-many; feedback is zero-or-one due to UNIQUE request_id."),
    ], widths=[1.85, 4.17], font_size=9)

    add_heading(doc, "7. RELATIONAL SCHEMA AND NORMALIZATION", 1)
    add_heading(doc, "7.1 Relational Schema", 2)
    schema_lines = [
        "BUILDINGS(building_id PK, building_name UNIQUE, location)",
        "ROOMS(room_id PK, building_id FK, room_number, room_type, UNIQUE(building_id, room_number))",
        "USERS(user_id PK, full_name, email UNIQUE, phone, role CHECK)",
        "CATEGORIES(category_id PK, category_name UNIQUE, description)",
        "PRIORITIES(priority_id PK, priority_name UNIQUE, response_hours CHECK)",
        "TECHNICIANS(technician_id PK, technician_name, specialization, phone UNIQUE, status CHECK)",
        "MAINTENANCE_REQUESTS(request_id PK, request_number UNIQUE, user_id FK, room_id FK, category_id FK, priority_id FK, title, description, request_date, status, assigned_date, closed_date, closure_details)",
        "ASSIGNMENTS(assignment_id PK, request_id FK UNIQUE, technician_id FK, assigned_at, assignment_status CHECK)",
        "WORK_LOGS(work_log_id PK, request_id FK, technician_id FK, work_date, work_description, hours_spent CHECK)",
        "MATERIALS(material_id PK, material_name UNIQUE, unit, unit_cost CHECK)",
        "MATERIAL_USAGE(usage_id PK, request_id FK, material_id FK, quantity CHECK, used_date)",
        "STATUS_HISTORY(history_id PK, request_id FK, old_status, new_status, changed_at, changed_by FK -> USERS)",
        "FEEDBACK(feedback_id PK, request_id FK UNIQUE, rating CHECK, comments, feedback_date)",
    ]
    add_code(doc, "Table 1: Relational Schema of the Existing Database", "\n".join(schema_lines))
    add_heading(doc, "7.2 First Normal Form (1NF)", 2)
    add_body(doc, "Each table has a primary key and stores attributes as single values for the application's data model. For example, a maintenance request stores one user_id, one room_id, one category_id, and one priority_id; multiple work entries or consumed materials are represented as separate related rows, rather than repeating groups embedded in a request record. Long descriptions and comments are text values, not lists of structured items.")
    add_heading(doc, "7.3 Second Normal Form (2NF)", 2)
    add_body(doc, "The tables use single-column surrogate primary keys, so non-key attributes depend on the whole primary key. The room uniqueness rule is a separate composite alternate key on (building_id, room_number), and its attributes describe that complete room identity. Material quantities depend on a particular material-usage record, not on just a request or material in isolation.")
    add_heading(doc, "7.4 Third Normal Form (3NF)", 2)
    add_body(doc, "Descriptive facts about lookup entities are stored in their own relations. A room row references a building instead of repeating the building location in every request; a request references category and priority rows; a technician's name and specialization are stored with the technician rather than copied into assignments and work logs. This reduces update anomalies when a label or location changes. Separate work_logs and material_usage rows allow inserting a request without inventing work/material details and deleting a usage record without removing the material master. The schema therefore demonstrates the intended normalization benefits while retaining operational history.")

    add_data_dictionary(doc, schema, fk_map, unique_map)

    add_heading(doc, "9. SQL COMMANDS USED (DDL AND DML) WITH SAMPLE OUTPUTS", 1)
    add_heading(doc, "9.1 Database Creation", 2)
    add_body(doc, "The database named campus_maintenance was already created and populated before this Flask application was connected. The application and this report generator do not run CREATE DATABASE, DROP, TRUNCATE, ALTER, or CREATE TABLE statements. No schema changes were made for this report.")
    add_code(doc, "Listing 1: Read-only Database Target Verification", "SELECT DATABASE();\n-- Result observed: campus_maintenance")
    add_heading(doc, "9.2 DDL Commands", 2)
    add_body(doc, "There is no DDL setup script in the inspected project. The command below was used only to inspect the actual existing table definition; it is not a schema-creation instruction.")
    add_code(doc, "Listing 2: Existing DDL Inspection Command (Read Only)", "SHOW CREATE TABLE maintenance_requests;")
    add_body(doc, "The inspected definition has request_id as AUTO_INCREMENT PRIMARY KEY; UNIQUE request_number; foreign keys to users, rooms, categories, and priorities; DEFAULT CURRENT_TIMESTAMP for request_date; and DEFAULT 'Pending' for status. The complete per-column schema, including the other 12 tables, appears in Section 8.")
    add_heading(doc, "9.3 DML Commands", 2)
    add_code(doc, "Listing 3: Parameterized Read and Insert Patterns Used by Flask", """cursor.execute(
    "SELECT user_id, full_name AS name FROM users ORDER BY full_name"
)

cursor.execute(
    \"\"\"
    INSERT INTO maintenance_requests
    (request_number, title, description, user_id, room_id, category_id, priority_id)
    VALUES (%s, %s, %s, %s, %s, %s, %s)
    \"\"\",
    (request_number, title, description, user_id, room_id, category_id, priority_id),
)""")
    add_heading(doc, "9.4 Constraints", 2)
    add_body(doc, "The existing schema enforces primary keys, foreign keys, UNIQUE constraints, NOT NULL requirements, defaults, and CHECK constraints. Notable examples are the unique request_number; unique assignments.request_id; unique feedback.request_id; (building_id, room_number) uniqueness for rooms; role/status domain checks; positive response_hours and material quantity; non-negative hours and material cost; and a feedback rating from 1 to 5.")
    add_heading(doc, "9.5 Sample Data and Observed Output", 2)
    add_body(doc, f"At report preparation, the live database contained {counts.get('total', 0)} maintenance requests, {tech_count} technicians, and {len(feedback_rows)} feedback entries. The following five seeded sample records are shown with their actual internal IDs. They are not IDs 1–5.")
    add_table(doc, ["request_id", "Request Number", "Title", "Status"], [
        (r["request_id"], r["request_number"], r["title"], r["status"]) for r in sample_requests
    ], widths=[0.72, 1.35, 2.9, 1.05], font_size=9)

    add_heading(doc, "10. QUERIES WITH OUTPUTS", 1)
    add_body(doc, "These examples are the same query patterns implemented by the application's request listing and Reports page. Output samples below are read from the live database on the report-generation date; changing data later will change the output.")
    add_heading(doc, "10.1 Main Request Report (JOIN)", 2)
    add_code(doc, "Listing 4: Request List JOIN", """SELECT mr.request_id, mr.request_number, mr.title,
       u.full_name AS reported_by, c.category_name, p.priority_name,
       mr.status, b.building_name, r.room_number
FROM maintenance_requests mr
LEFT JOIN users u ON u.user_id = mr.user_id
LEFT JOIN rooms r ON r.room_id = mr.room_id
LEFT JOIN buildings b ON b.building_id = r.building_id
LEFT JOIN categories c ON c.category_id = mr.category_id
LEFT JOIN priorities p ON p.priority_id = mr.priority_id
ORDER BY mr.request_date DESC;""")
    add_table(doc, ["ID", "Request", "Title", "Status", "Building / Room"], [
        (r["request_id"], r["request_number"], r["title"], r["status"], f"{r['building_name']} / {r['room_number']}")
        for r in sample_requests[:3]
    ], widths=[0.52, 1.13, 1.55, 1.05, 1.77], font_size=8.8)
    add_body(doc, "The join obtains reporter, location, category, and priority labels without duplicating those descriptions in maintenance_requests. IDs 11–15 map directly to the five original request numbers.", after=4)
    add_heading(doc, "10.2 Pending Requests", 2)
    add_code(doc, "Listing 5: Pending Request Report", """SELECT mr.request_number, mr.title, c.category_name, p.priority_name,
       b.building_name, r.room_number, mr.status
FROM maintenance_requests mr
LEFT JOIN rooms r ON r.room_id = mr.room_id
LEFT JOIN buildings b ON b.building_id = r.building_id
LEFT JOIN categories c ON c.category_id = mr.category_id
LEFT JOIN priorities p ON p.priority_id = mr.priority_id
WHERE LOWER(COALESCE(mr.status, '')) = 'pending'
ORDER BY mr.request_date DESC;""")
    add_table(doc, ["Request Number", "Title", "Category", "Priority", "Building / Room"], [
        (r["request_number"], r["title"], r["category_name"], r["priority_name"], f"{r['building_name']} / {r['room_number']}")
        for r in pending[:3]
    ], widths=[1.05, 1.65, 1.0, 0.8, 1.5], font_size=8.5)
    add_heading(doc, "10.3 Technician Workload", 2)
    add_code(doc, "Listing 6: Technician Load Aggregate", """SELECT tech.technician_name, tech.specialization,
       COUNT(a.assignment_id) AS assigned_requests
FROM technicians tech
LEFT JOIN assignments a ON a.technician_id = tech.technician_id
GROUP BY tech.technician_id, tech.technician_name, tech.specialization
ORDER BY assigned_requests DESC;""")
    add_table(doc, ["Technician", "Specialization", "Requests"], [
        (r["technician_name"], r["specialization"], r["assigned_requests"]) for r in technician_load
    ], widths=[2.2, 2.5, 1.0], font_size=9)
    add_heading(doc, "10.4 Material Cost and Building Counts", 2)
    add_code(doc, "Listing 7: Material Quantity and Cost Aggregate", """SELECT m.material_name, SUM(mu.quantity) AS quantity_used,
       m.unit_cost, SUM(mu.quantity * m.unit_cost) AS total_cost
FROM material_usage mu
LEFT JOIN materials m ON m.material_id = mu.material_id
GROUP BY m.material_id, m.material_name, m.unit_cost
ORDER BY total_cost DESC;""")
    add_table(doc, ["Material", "Quantity", "Unit Cost", "Total Cost"], [
        (r["material_name"], r["quantity_used"], f"{r['unit_cost']:.2f}", f"{r['total_cost']:.2f}") for r in material_costs
    ], widths=[2.0, 1.1, 1.2, 1.2], font_size=9)
    add_table(doc, ["Building", "Request Count"], [
        (r["building_name"], r["request_count"]) for r in building_counts
    ], widths=[3.2, 1.2], font_size=9)
    add_heading(doc, "10.5 Feedback, Ageing, and Response Time", 2)
    add_body(doc, "Feedback is joined to its request number and ordered by feedback_date. Request ageing uses DATEDIFF(NOW(), request_date) for requests not in Completed or Closed status. Response time uses DATEDIFF(assigned_at, request_date) where an assignment exists. In the sampled rows, ages and response intervals were zero whole days because the records were created/assigned on the same date; DATEDIFF intentionally reports calendar-day differences, not hours.")
    add_table(doc, ["Request", "Status", "Age (days)"], [
        (r["request_number"], r["status"], r["age_days"]) for r in ageing
    ], widths=[1.8, 1.8, 1.0], font_size=9)
    add_table(doc, ["Feedback request", "Rating", "Comments"], [
        (r["request_number"], r["rating"], r["comments"]) for r in feedback_rows
    ], widths=[1.5, 0.7, 3.4], font_size=8.5)
    add_table(doc, ["Assigned request", "Response (days)"], [
        (r["request_number"], r["response_days"]) for r in response_times
    ], widths=[2.1, 1.2], font_size=9)

    add_heading(doc, "11. UI DESIGN AND SCREENSHOTS", 1)
    add_body(doc, "The screenshots shown in this section are intentionally editable placeholders, not fabricated captures. The local application provides the pages represented below; replace each box with a screenshot captured from the running browser before final submission.")
    add_screenshot_placeholders(doc)

    add_heading(doc, "12. IMPLEMENTATION DETAILS", 1)
    add_heading(doc, "12.1 Technology Stack", 2)
    add_body(doc, "Python 3.12 and Flask implement server-side routes. mysql-connector-python supplies MySQL access. Jinja2 renders the HTML templates; HTML5 and CSS3 provide forms, tables, responsive layout, and status badges. A small vanilla JavaScript file defaults blank date inputs to the current date. Dependencies are listed in requirements.txt. No APIs, client-side framework, or authentication subsystem are present.")
    add_heading(doc, "12.2 System Architecture", 2)
    add_figure(doc, architecture_path, "Figure 12: System Architecture", width=4.35)
    add_heading(doc, "12.3 Frontend Implementation", 2)
    add_body(doc, "The shared base template defines navigation links to Dashboard, Requests, Create Request, Technicians, and Reports. Child templates provide request forms, joined detail views, tables, and report sections. The stylesheet defines the sidebar, cards, panels, badges, table overflow handling, and a narrow-screen layout. Flask's Jinja environment autoescapes template values.")
    add_heading(doc, "12.4 Flask Backend and Routes", 2)
    add_body(doc, "The route functions retrieve request records and lookups, validate submitted values, invoke parameterized SQL, and render templates or redirect after successful writes. The request list supports search by request number/title and filters by status, category, and priority. The detail page queries assignments, status history, work logs, material usage, and feedback. Database errors are logged to the Flask terminal and returned as friendly page messages.")
    add_heading(doc, "12.5 MySQL Connectivity and Configuration", 2)
    add_body(doc, "config.py loads a local .env through python-dotenv and maps DB_HOST, DB_USER, DB_PASSWORD, and DB_NAME to mysql.connector.connect. The project .gitignore excludes .env. The example file contains placeholders only. Reusable fetch_one/fetch_all helpers open a dictionary cursor, execute SQL, close the cursor and connection in finally blocks, and allow database exceptions to reach application error handling.")
    add_code(doc, "Listing 8: Database Connection and Environment-Based Settings", """def get_db_connection():
    try:
        return mysql.connector.connect(**get_db_config())
    except Error:
        app.logger.exception(\"Could not connect to the configured MySQL database\")
        raise""")
    add_heading(doc, "12.6 CRUD, Validation, and Transactions", 2)
    add_body(doc, "Create request is the principal create operation and uses request/user/room/category/priority data. Read operations power the dashboard, request list, detail view, technician page, and reports. Updates change status, assigned dates, closure fields, and technician state; inserts record assignments, work logs, material use, history, and feedback. Form validation checks required fields, numeric hours, positive material quantity, rating range, and allowed status transitions. Duplicate assignment and feedback prevention rely on the schema's UNIQUE constraints.")
    add_code(doc, "Listing 9: Request Creation Transaction and Initial History", """conn.start_transaction()
request_number = next_request_number(cursor)
cursor.execute(
    \"\"\"
    INSERT INTO maintenance_requests
    (request_number, title, description, user_id, room_id, category_id, priority_id)
    VALUES (%s, %s, %s, %s, %s, %s, %s)
    \"\"\",
    (request_number, title, description, user_id, room_id, category_id, priority_id),
)
request_id = cursor.lastrowid
cursor.execute(
    \"INSERT INTO status_history (request_id, old_status, new_status, changed_by) \"
    \"VALUES (%s, NULL, 'Pending', %s)\",
    (request_id, user_id),
)
conn.commit()""")
    add_code(doc, "Listing 10: Parameterized Search and Status Update", """if search:
    query += \" AND (mr.request_number LIKE %s OR mr.title LIKE %s)\"
    like = f\"%{search}%\"
    params.extend([like, like])
cursor.execute(query, tuple(params))

cursor.execute(
    \"UPDATE maintenance_requests SET status = %s, closed_date = %s, \"
    \"closure_details = %s WHERE request_id = %s\",
    (new_status, closed_date, closure_details, request_id),
)""")
    add_heading(doc, "12.7 Reporting and Workflow", 2)
    add_body(doc, "Reports use COUNT, SUM, DATEDIFF, joins, and GROUP BY for pending and ageing requests, technician workload, building request count, material usage and cost, building material cost, feedback, and response time. The following diagram summarizes the implemented user workflow; it is a process view, not an additional database entity model.")
    add_figure(doc, workflow_path, "Figure 13: Application Workflow Diagram", width=5.3)

    add_heading(doc, "13. TESTING", 1)
    add_body(doc, "The table distinguishes checks exercised during report preparation from write paths not exercised in order to avoid adding or altering project records while preparing a document. PASS is used only for checks with observed results.")
    tests = [
        ("TC01", "Database connection", "Connect with .env config.", "Connect to campus_maintenance.", "Connection and SELECT succeeded.", "PASS"),
        ("TC02", "Dashboard", "GET /.", "Show live counts and recent rows.", "HTTP 200; live values shown.", "PASS"),
        ("TC03", "Existing requests", "GET /requests.", "Show original five requests.", "IDs 11–15; REQ-2026-001…005 shown.", "PASS"),
        ("TC04", "Create request", "POST valid fields.", "Insert request and initial history.", "Form insert verified in MySQL.", "PASS"),
        ("TC05", "Request number", "Read current max suffix.", "Generate next year-prefixed reference.", "006 inserted; next number was 007.", "PASS"),
        ("TC06", "Search", "Search REQ-2026-001.", "Show matching request only.", "Matching result verified.", "PASS"),
        ("TC07", "Filter", "Filter status=Closed.", "Show closed, exclude pending.", "REQ-2026-005 matched.", "PASS"),
        ("TC08", "Assignment", "POST technician choice.", "Atomically save assignment and history.", "GET verified; POST not exercised.", "NOT RUN"),
        ("TC09", "Status update", "POST allowed next state.", "Save status and history.", "Valid POST not exercised.", "NOT RUN"),
        ("TC10", "Work log", "POST description and hours.", "Insert work entry.", "GET verified; POST not exercised.", "NOT RUN"),
        ("TC11", "Material usage", "POST material and quantity.", "Insert usage for report totals.", "GET verified; POST not exercised.", "NOT RUN"),
        ("TC12", "Completion", "In Progress → Completed.", "Update request and assignment.", "Not exercised; live data preserved.", "NOT RUN"),
        ("TC13", "Closure", "Close with details.", "Save status, date, details, history.", "Not exercised; live data preserved.", "NOT RUN"),
        ("TC14", "Feedback", "Submit rating 1–5.", "Accept once for eligible request.", "POST not exercised.", "NOT RUN"),
        ("TC15", "Reports", "GET /reports.", "Render SQL report sections.", "HTTP 200; outputs verified.", "PASS"),
        ("TC16", "Invalid input", "POST empty create form.", "Reject without insert.", "Validation returned; count unchanged.", "PASS"),
        ("TC17", "Duplicate feedback", "POST feedback for rated request.", "UNIQUE key rejects duplicate.", "Not submitted during report prep.", "NOT RUN"),
        ("TC18", "Invalid transition", "Pending → Closed.", "Reject before database update.", "Transition rejected.", "PASS"),
    ]
    add_table(doc, ["Test ID", "Test Case", "Input / Action", "Expected Result", "Observed Result", "Status"], tests,
              widths=[0.44, 0.9, 1.25, 1.35, 1.48, 0.6], font_size=7.3)

    add_heading(doc, "14. CONCLUSION AND FUTURE ENHANCEMENTS", 1)
    add_heading(doc, "14.1 Conclusion", 2)
    add_body(doc, "The project demonstrates a working database-backed workflow for campus facility maintenance. Its MySQL design separates the principal entities, uses keys and domain constraints to protect consistency, and records assignment, work, material, status-history, and feedback details. Flask routes provide a straightforward interface to create, search, inspect, and report on requests. Live validation confirmed the database connection, the request joins, seeded sample IDs, dashboard, dropdowns, reports, search/filter behavior, and creation of a subsequent request number. The report describes the existing database rather than recreating it.")
    add_heading(doc, "14.2 Future Enhancements (Not Implemented)", 2)
    add_body(doc, "The following items are proposals only and are not part of the current prototype: user authentication; role-based access control; email/SMS notifications; a dedicated mobile application; advanced analytics; automated priority prediction; technician location tracking; asset management; maintenance scheduling; and cloud deployment. These additions would require separate requirements, threat analysis, data design, and testing.")

    add_heading(doc, "15. REFERENCES", 1)
    references = [
        "Python Software Foundation. (n.d.). Python 3.12 documentation. https://docs.python.org/3.12/",
        "Pallets Projects. (n.d.). Flask documentation. https://flask.palletsprojects.com/",
        "Oracle. (n.d.). MySQL 8.0 Reference Manual. https://dev.mysql.com/doc/refman/8.0/en/",
        "Oracle. (n.d.). MySQL Connector/Python Developer Guide. https://dev.mysql.com/doc/connector-python/en/",
        "Pallets Projects. (n.d.). Jinja documentation. https://jinja.palletsprojects.com/",
        "WHATWG. (Living Standard). HTML: The Living Standard. https://html.spec.whatwg.org/",
        "Silberschatz, A., Korth, H. F., & Sudarshan, S. (2019). Database System Concepts (7th ed.). McGraw-Hill.",
    ]
    for index, entry in enumerate(references, 1):
        para = doc.add_paragraph()
        para.paragraph_format.left_indent = Inches(0.28)
        para.paragraph_format.first_line_indent = Inches(-0.28)
        para.paragraph_format.line_spacing = 1.15
        para.paragraph_format.space_after = Pt(8)
        set_run(para.add_run(f"[{index}] {entry}"), 11)

    add_heading(doc, "16. CONTRIBUTION OF EACH MEMBER", 1)
    add_body(doc, "The repository does not identify authorship by individual. The allocation below is a balanced editable draft for the team to review and correct before submission; it is not independently verifiable from source files.")
    add_table(doc, ["Work Area", "Pranav Arjun Nikam\n25WU0101085", "Saathvik Mahalsa\n25WU0101115", "Mithil D. Parikh\n25WU0101072"], [
        ("Database design", "Lead", "Review", "Review"),
        ("SQL implementation", "Review", "Lead", "Review"),
        ("Backend development", "Lead", "Lead", "Review"),
        ("Frontend development", "Lead", "Review", "Lead"),
        ("Testing", "Review", "Review", "Lead"),
        ("Documentation", "Lead", "Lead", "Lead"),
        ("Integration", "Review", "Lead", "Lead"),
        ("Presentation", "Lead", "Lead", "Lead"),
    ], widths=[1.35, 1.55, 1.55, 1.57], font_size=9)

    add_heading(doc, "17. APPENDIX: GITHUB REPOSITORY LINK", 1)
    add_heading(doc, "17.1 GitHub Repository", 2)
    add_body(doc, "[INSERT GITHUB REPOSITORY LINK]")
    add_heading(doc, "17.2 Project Folder Structure", 2)
    tree = [
        "campus_maintenance/",
        "├── app.py; config.py; requirements.txt; .env.example; .gitignore",
        "├── templates/  (base, dashboard, requests, details, forms, reports, errors)",
        "├── static/     (css/style.css; js/script.js)",
        "└── report_assets/  (ER, architecture, workflow diagrams; report generator)",
    ]
    add_code(doc, "Listing 11: Inspected Project Structure", "\n".join(tree))
    add_heading(doc, "17.3 Report Preparation Notes", 2)
    add_body(doc, f"Prepared on {date.today().isoformat()} from the current project source and a read-only inspection of MySQL. The live database reported {counts.get('total', 0)} maintenance rows at generation time; the five original sample requests retain internal request IDs 11, 12, 13, 14, and 15. The local .env password is not reproduced here. No database schema or records were changed to generate this report.")

    doc.save(OUTPUT)
    print(f"Created {OUTPUT}")
    print(f"Database rows sampled: {counts.get('total', 0)}")
    print(f"Schema tables documented: {len(TABLES)}")
    print(f"Data dictionary columns documented: {sum(len(schema[t]) for t in TABLES)}")


if __name__ == "__main__":
    main()
