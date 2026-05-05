import customtkinter as ctk
from tkinter import filedialog, messagebox
from PIL import Image as PILImage
import openpyxl
from openpyxl.styles import PatternFill, Font
import urllib.request
import threading
import os
import sys
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Spacer, Image as RLImage, Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT

# ── Version ───────────────────────────────────────────────
CURRENT_VERSION = "1.0.0"
VERSION_URL = "https://raw.githubusercontent.com/MrSamuelB/schedule-report-payroll/refs/heads/main/version.txt"
UPDATE_URL = "https://github.com/MrSamuelB/schedule-report-payroll/releases/latest"

# ── Rules ─────────────────────────────────────────────────
COLUMNS_TO_DELETE = [
    "medical record number", "branch name", "phone", "task type",
    "address", "city", "state", "zip code", "employee id"
]

TASK_NAMES_TO_DELETE = [
    "30-day summary", "additional hospital records", "admit summary",
    "case conference and 60 day summary", "cms 485", "consent", "cota sup",
    "pta sup", "lpn sup", "discharge summary", "dnr", "dpa", "emergency plan",
    "follow up vpoc", "f2f encounter", "f2f notes", "frequency order",
    "hospitalization", "insurance payment", "insurance verification",
    "lab results", "medication profile", "msw communication", "nomnc",
    "otcomm", "ovn", "patient communication", "physician order", "ptcomm",
    "referral approval", "referral", "release of information form", "orc",
    "rocdocs", "soc email", "transfer summary", "wound company ovn"
]

STATUS_DELETE = "submitted with signature (mv)"
STATUS_HIGHLIGHT = ["not started", "saved"]

RED_FONT = Font(color="FF0000")
RED_FILL = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")

# ── Logo path ─────────────────────────────────────────────
def resource_path(filename):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, filename)
    return os.path.join(os.path.dirname(os.path.abspath(__file__)), filename)

LOGO_PATH = resource_path("Logo.png")

# ── Auto-update ───────────────────────────────────────────
def check_for_updates():
    try:
        with urllib.request.urlopen(VERSION_URL, timeout=5) as response:
            latest_version = response.read().decode("utf-8").strip()
        if latest_version != CURRENT_VERSION:
            answer = messagebox.askyesno(
                "Update Available",
                f"A new version is available ({latest_version}).\n\n"
                f"You are on version {CURRENT_VERSION}.\n\n"
                "Would you like to go to the download page?"
            )
            if answer:
                import webbrowser
                webbrowser.open(UPDATE_URL)
    except Exception:
        pass

# ── Helpers ───────────────────────────────────────────────
def get_col_index(ws, name):
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value and str(cell.value).strip().lower() == name.lower():
            return col_idx
    return None

def move_column_to(ws, from_idx, to_idx):
    if from_idx == to_idx:
        return
    col_data = [row[from_idx - 1].value for row in ws.iter_rows()]
    ws.delete_cols(from_idx)
    ws.insert_cols(to_idx)
    for row_idx, value in enumerate(col_data, start=1):
        ws.cell(row=row_idx, column=to_idx, value=value)

# ── Process data ──────────────────────────────────────────
def process_data(ws):
    # Step 1: Delete specified columns
    cols_to_remove = []
    for col_idx, cell in enumerate(ws[1], start=1):
        if cell.value and str(cell.value).strip().lower() in COLUMNS_TO_DELETE:
            cols_to_remove.append(col_idx)
    for col_idx in sorted(cols_to_remove, reverse=True):
        ws.delete_cols(col_idx)

    # Step 2: Move Provider Last → col B, Provider First → col C
    provider_last_idx = get_col_index(ws, "Provider Last")
    if provider_last_idx:
        move_column_to(ws, provider_last_idx, 2)
    provider_first_idx = get_col_index(ws, "Provider First")
    if provider_first_idx:
        move_column_to(ws, provider_first_idx, 3)

    # Step 3: Sort by Task Name alphabetically
    task_name_idx = get_col_index(ws, "Task Name")
    if task_name_idx:
        data_rows = [list(row) for row in ws.iter_rows(min_row=2, values_only=True)]
        data_rows.sort(key=lambda r: str(r[task_name_idx - 1] or "").lower())
        for row_idx, row_data in enumerate(data_rows, start=2):
            for col_idx, value in enumerate(row_data, start=1):
                ws.cell(row=row_idx, column=col_idx, value=value)

    # Step 4: Delete rows where Task Name matches list
    task_name_idx = get_col_index(ws, "Task Name")
    if task_name_idx:
        rows_to_delete = []
        for row in ws.iter_rows(min_row=2):
            cell_value = str(row[task_name_idx - 1].value or "").strip().lower()
            if cell_value in TASK_NAMES_TO_DELETE:
                rows_to_delete.append(row[0].row)
        for row_idx in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row_idx)

    # Step 5: Handle Status column
    status_idx = get_col_index(ws, "Status")
    if status_idx:
        rows_to_delete = []
        for row in ws.iter_rows(min_row=2):
            cell_value = str(row[status_idx - 1].value or "").strip().lower()
            if cell_value == STATUS_DELETE:
                rows_to_delete.append(row[0].row)
            elif cell_value in STATUS_HIGHLIGHT:
                for cell in row:
                    cell.font = RED_FONT
                    cell.fill = RED_FILL
        for row_idx in sorted(rows_to_delete, reverse=True):
            ws.delete_rows(row_idx)

    return ws

# ── PDF generation ────────────────────────────────────────
def generate_pdf(ws, save_path, is_debug=False):
    pagesize = letter
    page_width, page_height = pagesize
    margin = 0.5 * inch
    usable_width = page_width - 2 * margin

    all_rows = list(ws.iter_rows(values_only=True))
    if not all_rows:
        return

    headers = [str(c) if c is not None else "" for c in all_rows[0]]
    data_rows = all_rows[1:]

    # Find status column
    status_col_idx = None
    for i, h in enumerate(headers):
        if h.strip().lower() == "status":
            status_col_idx = i
            break

    col_count = len(headers)
    col_widths = [usable_width / col_count] * col_count

    # Styles
    normal_style = ParagraphStyle(
        "cell", fontName="Helvetica", fontSize=8, leading=11, wordWrap="LTR"
    )
    red_style = ParagraphStyle(
        "red_cell", fontName="Helvetica", fontSize=8, leading=11,
        wordWrap="LTR", textColor=colors.red
    )
    header_style = ParagraphStyle(
        "header", fontName="Helvetica-Bold", fontSize=8, leading=11,
        textColor=colors.white, wordWrap="LTR"
    )
    title_style = ParagraphStyle(
        "title", fontName="Helvetica-Bold", fontSize=14,
        leading=18, alignment=TA_CENTER
    )
    summary_label_style = ParagraphStyle(
        "sum_label", fontName="Helvetica-Bold", fontSize=10,
        leading=14, alignment=TA_LEFT
    )
    summary_value_style = ParagraphStyle(
        "sum_value", fontName="Helvetica", fontSize=10,
        leading=14, alignment=TA_LEFT
    )
    summary_red_style = ParagraphStyle(
        "sum_red", fontName="Helvetica-Bold", fontSize=10,
        leading=14, alignment=TA_LEFT, textColor=colors.red
    )
    summary_billable_style = ParagraphStyle(
        "sum_billable", fontName="Helvetica-Bold", fontSize=10,
        leading=14, alignment=TA_LEFT, textColor=colors.HexColor("#1a5276")
    )

    # Count visits
    total_visits = len(data_rows)
    red_visits = 0
    if status_col_idx is not None:
        for row in data_rows:
            val = str(row[status_col_idx] or "").strip().lower()
            if val in STATUS_HIGHLIGHT:
                red_visits += 1
    billable_visits = total_visits - red_visits

    # Build table rows
    table_data = [[Paragraph(h, header_style) for h in headers]]
    for row_idx, row in enumerate(data_rows, start=1):
        is_red = False
        if status_col_idx is not None:
            val = str(row[status_col_idx] or "").strip().lower()
            if val in STATUS_HIGHLIGHT:
                is_red = True
        style = red_style if is_red else normal_style
        cells = [Paragraph(str(c) if c is not None else "", style) for c in row]
        table_data.append(cells)

    # Table styling
    table_style_cmds = [
        ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#1a5276")),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.HexColor("#cccccc")),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1),
         [colors.white, colors.HexColor("#eaf0fb")]),
    ]

    table = Table(table_data, colWidths=col_widths, repeatRows=1)
    table.setStyle(TableStyle(table_style_cmds))

    # Summary table
    summary_data = [
        [Paragraph("Total Visits:", summary_label_style),
         Paragraph(str(total_visits), summary_value_style)],
        [Paragraph("Red Visits (Not Started / Saved):", summary_red_style),
         Paragraph(str(red_visits), summary_red_style)],
        [Paragraph("Billable Visits:", summary_billable_style),
         Paragraph(str(billable_visits), summary_billable_style)],
    ]
    summary_table = Table(summary_data, colWidths=[3*inch, 1*inch])
    summary_table.setStyle(TableStyle([
        ("LINEABOVE", (0, 0), (-1, 0), 1, colors.HexColor("#cccccc")),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
    ]))

    # Build elements
    elements = []

    if os.path.exists(LOGO_PATH):
        logo = RLImage(LOGO_PATH, width=0.8*inch, height=0.8*inch)
        elements.append(logo)

    elements.append(Spacer(1, 0.05*inch))
    elements.append(Paragraph("Schedule Report for Payroll", title_style))

    if is_debug:
        debug_style = ParagraphStyle(
            "debug", fontName="Helvetica-Bold", fontSize=10,
            leading=14, textColor=colors.red, alignment=TA_CENTER
        )
        elements.append(Spacer(1, 0.05*inch))
        elements.append(Paragraph("DEBUG VERSION — includes deleted rows", debug_style))

    elements.append(Spacer(1, 0.15*inch))
    elements.append(table)
    elements.append(Spacer(1, 0.2*inch))
    elements.append(summary_table)

    doc = SimpleDocTemplate(
        save_path,
        pagesize=pagesize,
        leftMargin=margin,
        rightMargin=margin,
        topMargin=margin,
        bottomMargin=margin
    )
    doc.build(elements)

# ── Upload & process ──────────────────────────────────────
def upload_file():
    file_path = filedialog.askopenfilename(
        title="Select your schedule file",
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if not file_path:
        return

    status_label.configure(text="Loading file...")
    app.update()

    workbook = openpyxl.load_workbook(file_path)
    ws = workbook.active

    row_count = ws.max_row - 1
    col_count = ws.max_column
    status_label.configure(
        text=f"Loaded: {row_count} rows, {col_count} columns. Processing..."
    )
    app.update()

    # Debug PDF (before processing)
    debug_pdf_path = filedialog.asksaveasfilename(
        title="Save DEBUG PDF as",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile="debug_report.pdf"
    )
    if debug_pdf_path:
        status_label.configure(text="Generating debug PDF...")
        app.update()
        generate_pdf(ws, debug_pdf_path, is_debug=True)

    # Apply all rules
    status_label.configure(text="Applying rules...")
    app.update()
    ws = process_data(ws)

    # Clean PDF
    clean_pdf_path = filedialog.asksaveasfilename(
        title="Save clean PDF as",
        defaultextension=".pdf",
        filetypes=[("PDF files", "*.pdf")],
        initialfile="schedule_report.pdf"
    )
    if clean_pdf_path:
        status_label.configure(text="Generating clean PDF...")
        app.update()
        generate_pdf(ws, clean_pdf_path, is_debug=False)

    # Excel file
    excel_path = filedialog.asksaveasfilename(
        title="Save Excel file as",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="schedule_report.xlsx"
    )
    if excel_path:
        status_label.configure(text="Saving Excel file...")
        app.update()
        workbook.save(excel_path)

    status_label.configure(
        text=f"All done! Rows in final report: {ws.max_row - 1}"
    )

# ── Build UI ──────────────────────────────────────────────
app = ctk.CTk()
app.title("Schedule Report For Payroll")
app.geometry("800x600")

# Logo on main window
if os.path.exists(LOGO_PATH):
    pil_image = PILImage.open(LOGO_PATH)
    logo_image = ctk.CTkImage(light_image=pil_image, dark_image=pil_image, size=(120, 120))
    logo_label = ctk.CTkLabel(app, image=logo_image, text="")
    logo_label.pack(pady=(20, 5))

title_label = ctk.CTkLabel(
    app,
    text="Schedule Report For Payroll",
    font=ctk.CTkFont(size=24, weight="bold")
)
title_label.pack(pady=(5, 0))

subtitle_label = ctk.CTkLabel(
    app,
    text="Upload a schedule, apply rules, and export a PDF and Excel file.",
    font=ctk.CTkFont(size=14)
)
subtitle_label.pack(pady=5)

version_label = ctk.CTkLabel(
    app,
    text=f"Version {CURRENT_VERSION}",
    font=ctk.CTkFont(size=11),
    text_color="gray"
)
version_label.pack(pady=0)

upload_btn = ctk.CTkButton(
    app,
    text="Upload Schedule",
    font=ctk.CTkFont(size=16),
    height=50,
    width=250,
    command=upload_file
)
upload_btn.pack(pady=20)

status_label = ctk.CTkLabel(
    app,
    text="No file loaded yet.",
    font=ctk.CTkFont(size=13),
    text_color="gray"
)
status_label.pack(pady=5)

columns_frame = ctk.CTkScrollableFrame(
    app,
    width=600,
    height=150,
    label_text=""
)
columns_frame.pack(pady=10, padx=20, fill="both", expand=True)

threading.Thread(target=check_for_updates, daemon=True).start()

app.mainloop()