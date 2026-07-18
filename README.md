import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Create workbook and setup sheets
wb = openpyxl.Workbook()

# Setup sheet 1: Dashboard / Summary
ws_dash = wb.active
ws_dash.title = "Inspection Summary"
ws_dash.views.sheetView[0].showGridLines = True

# Setup sheet 2: Detailed Inspection Checklist
ws_check = wb.create_sheet(title="QC Checklist Template")
ws_check.views.sheetView[0].showGridLines = True

# Color Palette: Deep Slate & Warm Amber Accents
COLOR_HEADER_FILL = "2C3E50"  # Deep Slate Blue
COLOR_HEADER_TEXT = "FFFFFF"  # White
COLOR_SECTION_FILL = "EAEDED" # Light Neutral Grey
COLOR_PASS_FILL = "D4EFDF"    # Soft Green
COLOR_FAIL_FILL = "FADBD8"    # Soft Red
COLOR_BORDER = "BDC3C7"       # Muted Grey Border

# Styles
font_title = Font(name="Segoe UI", size=16, bold=True, color="2C3E50")
font_section = Font(name="Segoe UI", size=11, bold=True, color="2C3E50")
font_header = Font(name="Segoe UI", size=11, bold=True, color=COLOR_HEADER_TEXT)
font_body = Font(name="Segoe UI", size=10)
font_body_bold = Font(name="Segoe UI", size=10, bold=True)
font_pass = Font(name="Segoe UI", size=10, bold=True, color="196F3D")
font_fail = Font(name="Segoe UI", size=10, bold=True, color="943126")

fill_header = PatternFill(start_color=COLOR_HEADER_FILL, end_color=COLOR_HEADER_FILL, fill_type="solid")
fill_section = PatternFill(start_color=COLOR_SECTION_FILL, end_color=COLOR_SECTION_FILL, fill_type="solid")
fill_pass = PatternFill(start_color=COLOR_PASS_FILL, end_color=COLOR_PASS_FILL, fill_type="solid")
fill_fail = PatternFill(start_color=COLOR_FAIL_FILL, end_color=COLOR_FAIL_FILL, fill_type="solid")

thin_border = Border(
    left=Side(style='thin', color=COLOR_BORDER),
    right=Side(style='thin', color=COLOR_BORDER),
    top=Side(style='thin', color=COLOR_BORDER),
    bottom=Side(style='thin', color=COLOR_BORDER)
)

# -------------------------------------------------------------
# POPULATE SHEET 1: INSPECTION SUMMARY / HEADER INFO
# -------------------------------------------------------------
ws_dash["A1"] = "WOODEN DOOR & CABINET QUALITY CONTROL REPORT"
ws_dash["A1"].font = font_title
ws_dash.row_dimensions[1].height = 25

# Meta Metadata
metadata = [
    ("Report ID:", "QC-2026-0001", "Date:", "2026-07-18"),
    ("Inspector Name:", "Fauzan Khalid", "Factory Section:", "Assembly & Finishing"),
    ("Project / Order No:", "PO-8892-KITCHEN", "Client Name:", "Premium Residential Units"),
    ("Material Category:", "Solid Wood & Melamine Combi", "Total Batch Qty:", 120)
]

for idx, row_data in enumerate(metadata, start=3):
    ws_dash.cell(row=idx, column=1, value=row_data[0]).font = font_body_bold
    ws_dash.cell(row=idx, column=2, value=row_data[1]).font = font_body
    ws_dash.cell(row=idx, column=4, value=row_data[2]).font = font_body_bold
    ws_dash.cell(row=idx, column=5, value=row_data[3]).font = font_body
    ws_dash.row_dimensions[idx].height = 18

# Key Performance Box (Formulas linking to Sheet 2)
ws_dash.cell(row=8, column=1, value="BATCH QUALITY METRICS").font = font_section
ws_dash.row_dimensions[8].height = 20

metrics_headers = ["Total Checkpoints Evaluated", "Passed Items", "Failed Items", "Acceptance Rate (%)"]
for idx, h in enumerate(metrics_headers, start=1):
    cell = ws_dash.cell(row=9, column=idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
ws_dash.row_dimensions[9].height = 28

# Formulas linking to checklist sheet
ws_dash.cell(row=10, column=1, value="=COUNTA('QC Checklist Template'!C5:C24)").font = font_body_bold
ws_dash.cell(row=10, column=2, value="=COUNTIF('QC Checklist Template'!C5:C24, \"Pass\")").font = font_body_bold
ws_dash.cell(row=10, column=3, value="=COUNTIF('QC Checklist Template'!C5:C24, \"Fail\")").font = font_body_bold
ws_dash.cell(row=10, column=4, value="=B10/A10").font = font_body_bold
ws_dash.cell(row=10, column=4).number_format = "0.0%"

for col in range(1, 5):
    c = ws_dash.cell(row=10, column=col)
    c.alignment = Alignment(horizontal="center", vertical="center")
    c.border = thin_border
ws_dash.row_dimensions[10].height = 22

# Notes / Action item box below
ws_dash.cell(row=13, column=1, value="NON-CONFORMANCE REPORT & CORRECTIVE ACTIONS (NCR)").font = font_section
ncr_headers = ["Item/Process Code", "Defect Description", "Root Cause Analysis", "Corrective Action Required", "Disposition Status"]
for idx, h in enumerate(ncr_headers, start=1):
    cell = ws_dash.cell(row=14, column=idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="center", vertical="center")
ws_dash.row_dimensions[14].height = 24

sample_ncr = [
    ("QC-CHECK-02", "Moisture Content at 13.5% on Solid Oak planks.", "Raw materials stored near an open bay window during rainfall.", "Move batch back to kiln drying; seal off bay window.", "Pending Re-inspection"),
    ("QC-CHECK-07", "Melamine door edges showing micro-chipping.", "Main panel saw scoring blade is dull and misaligned.", "Replace cutting edge & recalibrate alignment before resuming.", "Resolved & Verified")
]

for idx, ncr_data in enumerate(sample_ncr, start=15):
    for col_idx, value in enumerate(ncr_data, start=1):
        cell = ws_dash.cell(row=idx, column=col_idx, value=value)
        cell.font = font_body
        cell.border = thin_border
        if col_idx == 5:
            cell.fill = fill_fail if "Pending" in value else fill_pass
            cell.font = font_fail if "Pending" in value else font_pass
    ws_dash.row_dimensions[idx].height = 22

# -------------------------------------------------------------
# POPULATE SHEET 2: DETAILED QUALITY CHECKLIST TEMPLATE
# -------------------------------------------------------------
ws_check["A1"] = "DETAILED STEP-BY-STEP QC PRODUCTION CHECKLIST"
ws_check["A1"].font = font_title
ws_check.row_dimensions[1].height = 25

headers_check = ["QC Code", "Inspection Checkpoint / Parameter", "Status (Pass/Fail)", "Engineering Specification / Tolerance Limit", "Inspector Notes & Remarks"]
for idx, h in enumerate(headers_check, start=1):
    cell = ws_check.cell(row=4, column=idx, value=h)
    cell.font = font_header
    cell.fill = fill_header
    cell.alignment = Alignment(horizontal="left" if idx!=3 else "center", vertical="center")
ws_check.row_dimensions[4].height = 26

checklist_data = [
    # INCOMING
    ("SECTION 1: INCOMING RAW MATERIAL QUALITY", "", "", "", ""),
    ("QC-CHECK-01", "Solid Wood Species Verification & Grain Pattern Match", "Pass", "Must match master control samples visually.", "Approved."),
    ("QC-CHECK-02", "Solid Wood Moisture Content Verification", "Fail", "Strictly between 6.0% and 10.0% (Pin Meter checked).", "Failed at 13.5% on Batch #3."),
    ("QC-CHECK-03", "Melamine Board Core Density & Thickness Check", "Pass", "Thickness variance within +/- 0.2mm (Caliper checked).", "18.0mm nominal verified."),
    ("QC-CHECK-04", "Melamine Surface Lamination Visual Inspection", "Pass", "Zero surface scratches, telegraphing glue, or color fade.", "Perfect surface condition."),
    
    # IN-PROCESS
    ("SECTION 2: IN-PROCESS MACHINING & PROCESSING", "", "", "", ""),
    ("QC-CHECK-05", "Panel Cutting Dimensions & Sizing Precision", "Pass", "Length & width variance within +/- 0.5mm.", "All cuts square."),
    ("QC-CHECK-06", "Panel Squareness (Diagonal Cross-Measurement Check)", "Pass", "Diagonal difference must be <= 1.0mm.", "Perfect 90-degree corners."),
    ("QC-CHECK-07", "Melamine Edge Banding Adhesion & Clean Trim", "Fail", "No visible glue lines, zero peeling, zero edge chipping.", "Micro-chipping found on back edge."),
    ("QC-CHECK-08", "Hinge Boring & Shelf-Pin Hole Line Alignment", "Pass", "Position variance within +/- 0.5mm.", "CNC drilling accurately calibrated."),
    ("QC-CHECK-09", "Solid Wood Profile Sanding & Smoothness Check", "Pass", "No cross-grain sanding scratches, no rough end-grains.", "Ready for spray booth."),
    
    # FINAL
    ("SECTION 3: FINAL ASSEMBLY & STRUCTURAL INTEGRITY", "", "", "", ""),
    ("QC-CHECK-10", "Solid Wood Joint Gaps & Miter Joint Fitment", "Pass", "Zero visible gaps at joints; no glue squeeze-out.", "Shaker profile joints flush."),
    ("QC-CHECK-11", "Hardware Installation & Mechanical Functionality", "Pass", "Soft-close mechanisms, sliders, and hinges glide smoothly.", "Tested manually."),
    ("QC-CHECK-12", "Cabinet Box Door Reveal Gaps & Alignment Parallelism", "Pass", "Reveal gaps uniform between 2.0mm and 3.0mm.", "Aligned flush."),
    ("QC-CHECK-13", "Final Coat Polish/Lacquer Uniformity", "Pass", "Zero orange peel, paint runs, or debris contamination.", "Gloss/Matte level uniform.")
]

current_row = 5
for item in checklist_data:
    if item[1] == "" and item[2] == "":  # Section Header
        ws_check.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=5)
        cell = ws_check.cell(row=current_row, column=1, value=item[0])
        cell.font = font_section
        cell.fill = fill_section
        ws_check.row_dimensions[current_row].height = 22
        for col in range(1, 6):
            ws_check.cell(row=current_row, column=col).border = thin_border
    else:
        ws_check.cell(row=current_row, column=1, value=item[0]).font = font_body_bold
        ws_check.cell(row=current_row, column=2, value=item[1]).font = font_body
        
        status_cell = ws_check.cell(row=current_row, column=3, value=item[2])
        status_cell.alignment = Alignment(horizontal="center")
        if item[2] == "Pass":
            status_cell.fill = fill_pass
            status_cell.font = font_pass
        elif item[2] == "Fail":
            status_cell.fill = fill_fail
            status_cell.font = font_fail
            
        ws_check.cell(row=current_row, column=4, value=item[3]).font = font_body
        ws_check.cell(row=current_row, column=5, value=item[4]).font = font_body
        
        for col in range(1, 6):
            ws_check.cell(row=current_row, column=col).border = thin_border
        ws_check.row_dimensions[current_row].height = 20
    current_row += 1

# Auto-adjust column widths
for ws in [ws_dash, ws_check]:
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            if cell.value and not str(cell.value).startswith("="):
                max_len = max(max_len, len(str(cell.value)))
        ws.column_dimensions[col_letter].width = max(max_len + 4, 13)

# Specific custom adjustments to columns
ws_dash.column_dimensions['A'].width = 18
ws_dash.column_dimensions['B'].width = 35
ws_dash.column_dimensions['C'].width = 30
ws_dash.column_dimensions['D'].width = 35
ws_dash.column_dimensions['E'].width = 25

ws_check.column_dimensions['A'].width = 15
ws_check.column_dimensions['B'].width = 45
ws_check.column_dimensions['C'].width = 18
ws_check.column_dimensions['D'].width = 45
ws_check.column_dimensions['E'].width = 35

# Save workbook
file_path = "Cabinet_Factory_QC_Report_Template.xlsx"
wb.save(file_path)
print(f"File successfully created: {file_path}")
