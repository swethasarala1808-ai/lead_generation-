import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
from openpyxl.utils import get_column_letter
import os, time

class ExcelExporter:
    def export(self, leads, job_id):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Leads"

        # ── Colors ──
        HEADER_BG   = "1A1A2E"   # deep navy
        HEADER_FG   = "E0E0E0"   # light silver
        ALT_ROW     = "F0F4FF"   # soft blue-white
        ACCENT      = "0F3460"   # navy accent
        GREEN_SCORE = "00B050"   # quality score green
        GOLD        = "FFD700"

        thin = Side(style="thin", color="CCCCCC")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        headers = list(leads[0].keys()) if leads else []

        # Title row
        ws.merge_cells(f"A1:{get_column_letter(len(headers))}1")
        title_cell = ws["A1"]
        title_cell.value = "🎯  LeadGen Pro — Quality Lead Report"
        title_cell.font = Font(name="Calibri", bold=True, size=16, color=GOLD)
        title_cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 36

        # Sub-title row
        ws.merge_cells(f"A2:{get_column_letter(len(headers))}2")
        sub = ws["A2"]
        sub.value = f"Total Leads: {len(leads)}  |  Generated: {time.strftime('%d %b %Y %H:%M')}"
        sub.font = Font(name="Calibri", italic=True, size=10, color="AAAAAA")
        sub.fill = PatternFill("solid", fgColor=HEADER_BG)
        sub.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[2].height = 20

        # Header row
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = Font(name="Calibri", bold=True, size=10, color=HEADER_FG)
            cell.fill = PatternFill("solid", fgColor=ACCENT)
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = border
        ws.row_dimensions[3].height = 30

        # Data rows
        for row_idx, lead in enumerate(leads, 4):
            is_alt = (row_idx % 2 == 0)
            for col_idx, header in enumerate(headers, 1):
                value = lead.get(header, "")
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.font = Font(name="Calibri", size=9)
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
                cell.border = border
                if is_alt:
                    cell.fill = PatternFill("solid", fgColor=ALT_ROW)

                # Special formatting
                if header == "Lead Quality Score" and value:
                    score = int(value.replace("%",""))
                    if score >= 85:
                        cell.font = Font(name="Calibri", size=9, bold=True, color=GREEN_SCORE)
                    elif score >= 70:
                        cell.font = Font(name="Calibri", size=9, color="FFA500")
                if header == "GSTIN" and value:
                    cell.font = Font(name="Calibri", size=9, color="0070C0")
                if header == "Registration Type":
                    if "MSME" in str(value):
                        cell.font = Font(name="Calibri", size=9, bold=True, color="7030A0")
                    elif "Startup" in str(value):
                        cell.font = Font(name="Calibri", size=9, bold=True, color="C00000")

        # Column widths
        col_widths = {
            "Sr No": 6, "Company Name": 28, "Business Type": 22, "Registration Type": 22,
            "Owner / Contact Person": 22, "Designation": 16, "Mobile Number": 15,
            "Alternate Mobile": 15, "Email ID": 30, "Website": 28,
            "Address": 30, "City": 14, "State": 16, "Pincode": 10,
            "Industry": 14, "Annual Turnover": 16, "No. of Employees": 14,
            "Year Established": 14, "GSTIN": 22, "WhatsApp Available": 14,
            "Lead Quality Score": 14, "Source": 14, "Remarks": 24,
        }
        for col_idx, header in enumerate(headers, 1):
            ws.column_dimensions[get_column_letter(col_idx)].width = col_widths.get(header, 15)

        # Freeze panes
        ws.freeze_panes = "A4"

        # Auto-filter
        ws.auto_filter.ref = f"A3:{get_column_letter(len(headers))}{len(leads)+3}"

        # Summary sheet
        ws2 = wb.create_sheet("Summary")
        ws2["A1"] = "Lead Generation Summary"
        ws2["A1"].font = Font(name="Calibri", bold=True, size=14, color=HEADER_FG)
        ws2["A1"].fill = PatternFill("solid", fgColor=HEADER_BG)
        ws2.merge_cells("A1:C1")
        ws2["A1"].alignment = Alignment(horizontal="center")

        summary_data = [
            ("Total Leads Generated", len(leads)),
            ("Industry", leads[0].get("Industry","") if leads else ""),
            ("City", leads[0].get("City","") if leads else ""),
            ("MSME Registered", sum(1 for l in leads if "MSME" in l.get("Registration Type",""))),
            ("Startup Registered", sum(1 for l in leads if "Startup" in l.get("Registration Type",""))),
            ("Leads with GSTIN", sum(1 for l in leads if l.get("GSTIN",""))),
            ("Leads with Website", sum(1 for l in leads if l.get("Website",""))),
            ("Leads with WhatsApp", sum(1 for l in leads if l.get("WhatsApp Available","") == "Yes")),
            ("High Quality Leads (>85%)", sum(1 for l in leads if l.get("Lead Quality Score","").replace("%","").isdigit() and int(l.get("Lead Quality Score","0%").replace("%","")) >= 85)),
        ]
        for r, (label, val) in enumerate(summary_data, 3):
            ws2.cell(row=r, column=1, value=label).font = Font(name="Calibri", bold=True, size=10)
            ws2.cell(row=r, column=2, value=val).font = Font(name="Calibri", size=10)
            ws2.cell(row=r, column=1).fill = PatternFill("solid", fgColor="E8EAF6")
        ws2.column_dimensions["A"].width = 30
        ws2.column_dimensions["B"].width = 20

        os.makedirs("/home/claude/lead_generation-/exports", exist_ok=True)
        path = f"/home/claude/lead_generation-/exports/leads_{job_id}.xlsx"
        wb.save(path)
        return path
