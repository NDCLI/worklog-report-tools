from collections import Counter
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from data_engine import ReportResult


DARK_HEADER_FILL = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="444444", end_color="444444", fill_type="solid")
WARN_FILL = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
INFO_FILL = PatternFill(start_color="DDEBF7", end_color="DDEBF7", fill_type="solid")
OK_FILL = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
ZEBRA_FILL = PatternFill(start_color="F8F9FA", end_color="F8F9FA", fill_type="solid")

FONT_TITLE = Font(name="Times New Roman", size=12, bold=True, color="FFFFFF")
FONT_HEADER = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_TASK = Font(name="Times New Roman", size=10, bold=True, color="FFFFFF")
FONT_REGULAR = Font(name="Times New Roman", size=10, bold=False, color="000000")
FONT_BOLD = Font(name="Times New Roman", size=10, bold=True, color="000000")
FONT_RED = Font(name="Times New Roman", size=10, bold=False, color="C00000")
FONT_RED_BOLD = Font(name="Times New Roman", size=10, bold=True, color="C00000")
FONT_WARN = Font(name="Segoe UI", size=10, bold=True, color="9C6500")
FONT_INFO = Font(name="Segoe UI", size=10, bold=True, color="1F4E78")
FONT_OK = Font(name="Segoe UI", size=10, bold=True, color="375623")
FONT_AUDIT_HDR = Font(name="Segoe UI", size=10, bold=True, color="FFFFFF")
FONT_AUDIT_REG = Font(name="Segoe UI", size=10, bold=False, color="000000")

THIN_GRAY = Side(style="thin", color="BFBFBF")
THIN_BLACK = Side(style="thin", color="000000")
MEDIUM_BLACK = Side(style="medium", color="000000")

ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")
ALIGN_LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True)
ALIGN_BOTTOM_RIGHT = Alignment(horizontal="right", vertical="bottom")


def _clean_num(val: float):
    if abs(val - round(val)) < 1e-6:
        return int(round(val))
    return round(val, 2)


def _make_border(left=THIN_GRAY, right=THIN_GRAY, top=THIN_GRAY, bottom=THIN_GRAY):
    return Border(left=left, right=right, top=top, bottom=bottom)


def _write_template1_by_task(wb: openpyxl.Workbook, report: ReportResult, highlight_mismatch: bool = True):
    ws = wb.active
    ws.title = "Task"
    ws.views.sheetView[0].showGridLines = True

    title_text = f"{report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')}"

    # Row 1: Title banner
    ws.merge_cells("A1:P1")
    ws["A1"] = title_text
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = DARK_HEADER_FILL
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28
    for c in range(1, 17):
        ws.cell(1, c).fill = DARK_HEADER_FILL
        ws.cell(1, c).border = _make_border(
            left=MEDIUM_BLACK if c == 1 else THIN_BLACK,
            right=MEDIUM_BLACK if c == 16 else THIN_BLACK,
            top=MEDIUM_BLACK,
            bottom=THIN_BLACK,
        )

    # Row 2: Column Headers
    headers = [
        "TASK",
        "Annotator",
        "Work Load",
        "Work BBOX",
        "Work Time",
        "Err\nFrame",
        "Err\nBBOX",
        "FRAME Per\nHour",
        "BBOX Per\nHour",
        "Err Rate\nFrame (%)",
        "Err Rate\nBBOX (%)",
        "TASK",
        "Avg\nFrame/h",
        "Avg\nBBOX/h",
        "Err Rate\nFrame (%)",
        "Err Rate\nBBOX (%)",
    ]
    ws.row_dimensions[2].height = 54
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(2, c, h)
        cell.font = FONT_HEADER
        cell.fill = DARK_HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(
            left=MEDIUM_BLACK if c in (1, 8, 12, 13) else THIN_BLACK,
            right=MEDIUM_BLACK if c in (7, 11, 12, 16) else THIN_BLACK,
            top=THIN_BLACK,
            bottom=MEDIUM_BLACK,
        )

    cur_row = 3
    for grp in report.task_groups:
        start_r = cur_row
        end_r = cur_row + len(grp.rows) - 1

        for idx, rdata in enumerate(grp.rows):
            r = cur_row + idx
            ws.row_dimensions[r].height = 20
            is_first = (idx == 0)
            is_last = (idx == len(grp.rows) - 1)

            ws.cell(r, 1, grp.task if is_first else None)
            ws.cell(r, 2, rdata.annotator)
            ws.cell(r, 3, _clean_num(rdata.work_load))
            ws.cell(r, 4, _clean_num(rdata.work_bbox))
            ws.cell(r, 5, _clean_num(rdata.work_time))
            ws.cell(r, 6, _clean_num(rdata.total_err_frame))
            ws.cell(r, 7, _clean_num(rdata.total_err_bbox))

            # Live Excel formulas for KPIs
            ws.cell(r, 8, f"=IF(E{r}>0,C{r}/E{r},0)")
            ws.cell(r, 9, f"=IF(E{r}>0,D{r}/E{r},0)")
            ws.cell(r, 10, f"=IF(C{r}>0,F{r}/C{r}*100,0)")
            ws.cell(r, 11, f"=IF(D{r}>0,G{r}/D{r}*100,0)")
            ws.cell(r, 12, grp.task if is_first else None)

            for c in range(1, 17):
                cell = ws.cell(r, c)
                top_b = MEDIUM_BLACK if is_first else THIN_GRAY
                bot_b = MEDIUM_BLACK if is_last else THIN_GRAY
                if c in (1, 12, 13, 14, 15, 16):
                    # Merged cells keep outer borders cleanly
                    left_b = MEDIUM_BLACK if c in (1, 12, 13) else THIN_BLACK
                    right_b = MEDIUM_BLACK if c in (1, 12, 16) else THIN_BLACK
                    cell.border = _make_border(left=left_b, right=right_b, top=top_b, bottom=bot_b)
                    if c in (13, 14, 15, 16):
                        cell.fill = ZEBRA_FILL
                else:
                    left_b = MEDIUM_BLACK if c == 8 else THIN_GRAY
                    right_b = MEDIUM_BLACK if c in (7, 11) else THIN_GRAY
                    cell.border = _make_border(left=left_b, right=right_b, top=top_b, bottom=bot_b)

                if c in (1, 12):
                    cell.fill = DARK_HEADER_FILL
                    cell.font = FONT_TASK
                    cell.alignment = ALIGN_CENTER
                elif c == 2:
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_CENTER
                    if highlight_mismatch and rdata.has_mismatch:
                        cell.fill = WARN_FILL
                elif c in (3, 4, 5, 6, 7):
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.##" if c == 5 else "0"
                    if highlight_mismatch and rdata.has_mismatch and c in (3, 4):
                        cell.fill = WARN_FILL
                elif c in (8, 9):
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"
                elif c in (10, 11):
                    cell.font = FONT_RED
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"

        # Merge TASK cols (1 and 12) and 4 Average cols (13, 14, 15, 16) for this task group
        if end_r > start_r:
            ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=1)
            ws.merge_cells(start_row=start_r, start_column=12, end_row=end_r, end_column=12)
            ws.merge_cells(start_row=start_r, start_column=13, end_row=end_r, end_column=13)
            ws.merge_cells(start_row=start_r, start_column=14, end_row=end_r, end_column=14)
            ws.merge_cells(start_row=start_r, start_column=15, end_row=end_r, end_column=15)
            ws.merge_cells(start_row=start_r, start_column=16, end_row=end_r, end_column=16)

        c13 = ws.cell(start_r, 13, f"=IF(SUM(E{start_r}:E{end_r})>0,SUM(C{start_r}:C{end_r})/SUM(E{start_r}:E{end_r}),0)")
        c14 = ws.cell(start_r, 14, f"=IF(SUM(E{start_r}:E{end_r})>0,SUM(D{start_r}:D{end_r})/SUM(E{start_r}:E{end_r}),0)")
        c15 = ws.cell(start_r, 15, f"=IF(SUM(C{start_r}:C{end_r})>0,SUM(F{start_r}:F{end_r})/SUM(C{start_r}:C{end_r})*100,0)")
        c16 = ws.cell(start_r, 16, f"=IF(SUM(D{start_r}:D{end_r})>0,SUM(G{start_r}:G{end_r})/SUM(D{start_r}:D{end_r})*100,0)")
        for avg_cell in (c13, c14, c15, c16):
            avg_cell.font = FONT_RED_BOLD
            avg_cell.alignment = ALIGN_CENTER
            avg_cell.number_format = "0.00"

        cur_row = end_r + 1

    # Column widths matching visual balance
    widths = [11, 13, 12, 12, 11, 11, 11, 13, 13, 13, 13, 11, 13, 13, 13, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C3"


def _write_template2_by_annotator(wb: openpyxl.Workbook, report: ReportResult, highlight_mismatch: bool = True):
    ws = wb.create_sheet(title="Annotator")
    ws.views.sheetView[0].showGridLines = True

    title_text = f"{report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')}"

    # Row 1: Title banner
    ws.merge_cells("A1:J1")
    ws["A1"] = title_text
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = DARK_HEADER_FILL
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28
    for c in range(1, 11):
        ws.cell(1, c).fill = DARK_HEADER_FILL
        ws.cell(1, c).border = _make_border(
            left=MEDIUM_BLACK if c == 1 else THIN_BLACK,
            right=MEDIUM_BLACK if c == 10 else THIN_BLACK,
            top=MEDIUM_BLACK,
            bottom=THIN_BLACK,
        )

    # Row 2: Column Headers
    headers = [
        "Annotator",
        "Work\nLoad",
        "Work\nBBOX",
        "Work Time",
        "Err\nFrame",
        "Err\nBBOX",
        "FRAME Per\nHour",
        "BBOX Per\nHour",
        "Err Rate\nFrame (%)",
        "Err Rate\nBBOX (%)",
    ]
    ws.row_dimensions[2].height = 54
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(2, c, h)
        cell.font = FONT_HEADER
        cell.fill = DARK_HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(
            left=MEDIUM_BLACK if c in (1, 7) else THIN_BLACK,
            right=MEDIUM_BLACK if c in (6, 10) else THIN_BLACK,
            top=THIN_BLACK,
            bottom=MEDIUM_BLACK,
        )

    for idx, rdata in enumerate(report.annotator_rows):
        r = 3 + idx
        ws.row_dimensions[r].height = 20
        is_last = (idx == len(report.annotator_rows) - 1)

        ws.cell(r, 1, rdata.annotator)
        ws.cell(r, 2, _clean_num(rdata.work_load))
        ws.cell(r, 3, _clean_num(rdata.work_bbox))
        ws.cell(r, 4, _clean_num(rdata.work_time))
        ws.cell(r, 5, _clean_num(rdata.total_err_frame))
        ws.cell(r, 6, _clean_num(rdata.total_err_bbox))

        ws.cell(r, 7, f"=IF(D{r}>0,B{r}/D{r},0)")
        ws.cell(r, 8, f"=IF(D{r}>0,C{r}/D{r},0)")
        ws.cell(r, 9, f"=IF(B{r}>0,E{r}/B{r}*100,0)")
        ws.cell(r, 10, f"=IF(C{r}>0,F{r}/C{r}*100,0)")

        for c in range(1, 11):
            cell = ws.cell(r, c)
            left_b = MEDIUM_BLACK if c in (1, 7) else THIN_GRAY
            right_b = MEDIUM_BLACK if c in (6, 10) else THIN_GRAY
            bot_b = MEDIUM_BLACK if is_last else THIN_GRAY
            cell.border = _make_border(left=left_b, right=right_b, top=THIN_GRAY, bottom=bot_b)

            if c == 1:
                cell.font = FONT_BOLD
                cell.alignment = ALIGN_CENTER
                if highlight_mismatch and rdata.has_mismatch:
                    cell.fill = WARN_FILL
            elif c in (2, 3, 4, 5, 6):
                cell.font = FONT_REGULAR
                cell.alignment = ALIGN_RIGHT
                cell.number_format = "0.##" if c == 4 else "0"
                if highlight_mismatch and rdata.has_mismatch and c in (2, 3):
                    cell.fill = WARN_FILL
            elif c in (7, 8):
                cell.font = FONT_BOLD
                cell.alignment = ALIGN_RIGHT
                cell.number_format = "0.00"
            elif c in (9, 10):
                cell.font = FONT_RED
                cell.alignment = ALIGN_RIGHT
                cell.number_format = "0.00"

    widths = [15, 11, 12, 11, 11, 11, 13, 13, 13, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B3"


def _write_reviewer_sheet(wb: openpyxl.Workbook, report: ReportResult):
    ws = wb.create_sheet(title="Reviewer")
    ws.views.sheetView[0].showGridLines = True

    title_text = f"{report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')}"
    ws.merge_cells("A1:L1")
    ws["A1"] = title_text
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = DARK_HEADER_FILL
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28
    for c in range(1, 13):
        ws.cell(1, c).fill = DARK_HEADER_FILL
        ws.cell(1, c).border = _make_border(
            left=MEDIUM_BLACK if c == 1 else THIN_BLACK,
            right=MEDIUM_BLACK if c == 12 else THIN_BLACK,
            top=MEDIUM_BLACK,
            bottom=THIN_BLACK,
        )

    headers = [
        "INSPECTOR",
        "Annotator",
        "Work Load",
        "Work BBOX",
        "Work Time",
        "First\nInspection\nError (By\nFrame)",
        "Second\nInspection\nError (By\nFrame)",
        "FRAME Per\nHour",
        "BBOX Per\nHour",
        "Error Rate\nFrame (%)",
        "Error Rate\nBBOX (%)",
        "First\nInspector\nRecall (By\nFrame)",
    ]
    ws.row_dimensions[2].height = 56
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(2, c, h)
        cell.font = FONT_HEADER
        cell.fill = DARK_HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(
            left=MEDIUM_BLACK if c in (1, 8) else THIN_BLACK,
            right=MEDIUM_BLACK if c in (7, 12) else THIN_BLACK,
            top=THIN_BLACK,
            bottom=MEDIUM_BLACK,
        )

    cur_row = 3
    for grp in report.inspector_groups:
        start_r = cur_row
        end_r = cur_row + len(grp.rows) - 1
        for idx, rdata in enumerate(grp.rows):
            r = cur_row + idx
            ws.row_dimensions[r].height = 20
            is_first = (idx == 0)
            is_last = (idx == len(grp.rows) - 1)

            ws.cell(r, 1, grp.inspector if is_first else None)
            ws.cell(r, 2, rdata.annotator)
            ws.cell(r, 3, _clean_num(rdata.work_load))
            ws.cell(r, 4, _clean_num(rdata.work_bbox))
            ws.cell(r, 5, _clean_num(rdata.work_time))
            ws.cell(r, 6, _clean_num(rdata.err_first))
            ws.cell(r, 7, _clean_num(rdata.err_second))

            for c in range(1, 13):
                cell = ws.cell(r, c)
                top_b = MEDIUM_BLACK if is_first else THIN_GRAY
                bot_b = MEDIUM_BLACK if is_last else THIN_GRAY
                if c in (1, 8, 9, 10, 11, 12):
                    left_b = MEDIUM_BLACK if c in (1, 8) else THIN_BLACK
                    right_b = MEDIUM_BLACK if c in (1, 12) else THIN_BLACK
                    cell.border = _make_border(left=left_b, right=right_b, top=top_b, bottom=bot_b)
                    if c >= 8:
                        cell.fill = ZEBRA_FILL
                else:
                    left_b = THIN_GRAY
                    right_b = MEDIUM_BLACK if c == 7 else THIN_GRAY
                    cell.border = _make_border(left=left_b, right=right_b, top=top_b, bottom=bot_b)

                if c == 1:
                    cell.fill = DARK_HEADER_FILL
                    cell.font = FONT_TASK
                    cell.alignment = ALIGN_CENTER
                elif c == 2:
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_CENTER
                elif c in (3, 4, 5, 6, 7):
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.##" if c == 5 else "#,##0"

        if end_r > start_r:
            ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=1)
            for c in (8, 9, 10, 11, 12):
                ws.merge_cells(start_row=start_r, start_column=c, end_row=end_r, end_column=c)

        c8 = ws.cell(start_r, 8, f"=IF(SUM(E{start_r}:E{end_r})>0,SUM(C{start_r}:C{end_r})/SUM(E{start_r}:E{end_r}),0)")
        c9 = ws.cell(start_r, 9, f"=IF(SUM(E{start_r}:E{end_r})>0,SUM(D{start_r}:D{end_r})/SUM(E{start_r}:E{end_r}),0)")
        c10 = ws.cell(start_r, 10, f"=IF(SUM(C{start_r}:C{end_r})>0,SUM(G{start_r}:G{end_r})/SUM(C{start_r}:C{end_r})*100,0)")
        c11 = ws.cell(start_r, 11, round(grp.err_rate_bbox, 2))
        c12 = ws.cell(start_r, 12, f"=IF((SUM(F{start_r}:F{end_r})+SUM(G{start_r}:G{end_r}))>0,SUM(F{start_r}:F{end_r})/(SUM(F{start_r}:F{end_r})+SUM(G{start_r}:G{end_r})),1)")
        for cell_m, fmt, fnt in (
            (c8, "0.00", FONT_RED_BOLD),
            (c9, "0.00", FONT_RED_BOLD),
            (c10, "0.00", FONT_BOLD),
            (c11, "0.00", FONT_BOLD),
            (c12, "0%", FONT_BOLD),
        ):
            cell_m.font = fnt
            cell_m.alignment = ALIGN_CENTER
            cell_m.number_format = fmt

        cur_row = end_r + 1

    # Grand Total row
    if cur_row > 3:
        last_data = cur_row - 1
        ws.row_dimensions[cur_row].height = 22
        ws.cell(cur_row, 1, "TỔNG").font = FONT_BOLD
        ws.cell(cur_row, 1).alignment = ALIGN_CENTER
        ws.cell(cur_row, 3, f"=SUM(C3:C{last_data})").number_format = "#,##0"
        ws.cell(cur_row, 4, f"=SUM(D3:D{last_data})").number_format = "#,##0"
        ws.cell(cur_row, 5, f"=SUM(E3:E{last_data})").number_format = "0.00"
        ws.cell(cur_row, 6, f"=SUM(F3:F{last_data})").number_format = "#,##0"
        ws.cell(cur_row, 7, f"=SUM(G3:G{last_data})").number_format = "#,##0"
        for c in range(3, 8):
            ws.cell(cur_row, c).font = FONT_BOLD
            ws.cell(cur_row, c).alignment = ALIGN_RIGHT

    widths = [13, 13, 12, 12, 11, 13, 13, 13, 13, 13, 13, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C3"


def _write_ranking_sheet(wb: openpyxl.Workbook, report: ReportResult):
    ws = wb.create_sheet(title="Xep_Hang")
    ws.views.sheetView[0].showGridLines = True

    workers = sorted(
        [r for r in report.annotator_rows if not r.is_reviewer and (r.work_load > 0 or r.work_bbox > 0)],
        key=lambda r: (r.total_err_rate, r.total_err_bbox_rate, -r.work_load, r.annotator.upper()),
    )
    reviewers = sorted(
        [g for g in report.inspector_groups if g.total_work_load > 0 or g.total_work_bbox > 0],
        key=lambda g: (g.err_rate_frame, g.err_rate_bbox, -g.total_work_load, g.inspector.upper()),
    )

    def write_rank_block(start_row: int, title: str, name_hdr: str, rows_list, is_rev: bool = False) -> int:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=11)
        ws.cell(start_row, 1, title)
        ws.cell(start_row, 1).font = FONT_AUDIT_HDR
        ws.row_dimensions[start_row].height = 24
        for c in range(1, 12):
            ws.cell(start_row, c).fill = DARK_HEADER_FILL

        hdr_row = start_row + 1
        hdrs = [
            "Hạng",
            name_hdr,
            "Work Load (Frame)",
            "Work BBOX",
            "Work Time",
            "Pangyo Frame" if is_rev else "Err Frame",
            "Pangyo BBOX" if is_rev else "Err BBOX",
            "Err Rate Frame (%)",
            "Err Rate BBOX (%)",
            "1st Insp. Recall (%)" if is_rev else "FRAME / Hour",
            "FRAME / Hour" if is_rev else "BBOX / Hour",
        ]
        ws.row_dimensions[hdr_row].height = 28
        for c, h in enumerate(hdrs, start=1):
            cell = ws.cell(hdr_row, c, h)
            cell.font = FONT_HEADER
            cell.fill = SUB_HEADER_FILL
            cell.alignment = ALIGN_CENTER
            cell.border = _make_border(THIN_BLACK, THIN_BLACK, THIN_BLACK, MEDIUM_BLACK)

        cur = hdr_row + 1
        for rank_idx, rdata in enumerate(rows_list, start=1):
            ws.row_dimensions[cur].height = 20
            ws.cell(cur, 1, rank_idx)
            if is_rev:
                ws.cell(cur, 2, rdata.inspector)
                ws.cell(cur, 3, _clean_num(rdata.total_work_load))
                ws.cell(cur, 4, _clean_num(rdata.total_work_bbox))
                ws.cell(cur, 5, _clean_num(rdata.total_work_time))
                ws.cell(cur, 6, _clean_num(rdata.total_err_second))
                ws.cell(cur, 7, _clean_num(rdata.total_err_second_bbox))
                ws.cell(cur, 8, f"=IF(C{cur}>0,F{cur}/C{cur}*100,0)")
                ws.cell(cur, 9, f"=IF(D{cur}>0,G{cur}/D{cur}*100,0)")
                ws.cell(cur, 10, round(rdata.first_inspector_recall, 0) / 100.0)
                ws.cell(cur, 11, f"=IF(E{cur}>0,C{cur}/E{cur},0)")
            else:
                ws.cell(cur, 2, rdata.annotator)
                ws.cell(cur, 3, _clean_num(rdata.work_load))
                ws.cell(cur, 4, _clean_num(rdata.work_bbox))
                ws.cell(cur, 5, _clean_num(rdata.work_time))
                ws.cell(cur, 6, _clean_num(rdata.total_err_frame))
                ws.cell(cur, 7, _clean_num(rdata.total_err_bbox))
                ws.cell(cur, 8, f"=IF(C{cur}>0,F{cur}/C{cur}*100,0)")
                ws.cell(cur, 9, f"=IF(D{cur}>0,G{cur}/D{cur}*100,0)")
                ws.cell(cur, 10, f"=IF(E{cur}>0,C{cur}/E{cur},0)")
                ws.cell(cur, 11, f"=IF(E{cur}>0,D{cur}/E{cur},0)")

            for c in range(1, 12):
                cell = ws.cell(cur, c)
                cell.border = _make_border(THIN_GRAY, THIN_GRAY, THIN_GRAY, THIN_GRAY)
                if rank_idx <= 3:
                    cell.fill = OK_FILL
                if c in (1, 2):
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_CENTER
                elif c in (3, 4, 5, 6, 7):
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.##" if c == 5 else "#,##0"
                elif c in (8, 9):
                    cell.font = FONT_RED_BOLD if rank_idx > 3 else FONT_GREEN_BOLD
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"
                elif is_rev and c == 10:
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_CENTER
                    cell.number_format = "0%"
                else:
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"
            cur += 1
        return cur + 2

    next_r = write_rank_block(
        1,
        f"BẢNG XẾP HẠNG ANNOTATOR (WORKER) — TỈ LỆ LỖI THẤP NHẤT ({report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')})",
        "Annotator (Worker)",
        workers,
        is_rev=False,
    )
    write_rank_block(
        next_r,
        f"BẢNG XẾP HẠNG REVIEWER (LẦN 1) — TỈ LỆ LỖI THẤP NHẤT ({report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')})",
        "Reviewer (Lần 1)",
        reviewers,
        is_rev=True,
    )

    widths = [8, 16, 16, 14, 12, 12, 12, 16, 16, 15, 14]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def _write_audit_and_notes_sheet(wb: openpyxl.Workbook, report: ReportResult):
    ws = wb.create_sheet(title="Audit_&_Notes")
    ws.views.sheetView[0].showGridLines = True

    # Part 1: Comparison table
    ws.merge_cells("A1:K1")
    ws["A1"] = (
        f"ĐỐI CHIẾU WORK LOG VS REVIEW ({report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')}) | "
        f"Nguồn chọn: {'Work Log + Note' if report.source_mode == 'worklog' else 'Review (#IMAGE)'} | "
        f"Lọc Review theo: {'Ngày làm việc (Cột V)' if report.rework_date_mode == 'col_v' else 'DATE (Cột A)'}"
    )
    ws["A1"].font = FONT_AUDIT_HDR
    ws["A1"].fill = DARK_HEADER_FILL
    ws["A1"].alignment = ALIGN_LEFT
    ws.row_dimensions[1].height = 26

    audit_headers = [
        "TASK", "Annotator",
        "Work Log (Frame)", "Review (Frame)", "Lệch Frame",
        "Work Log (BBOX)", "Review (BBOX)", "Lệch BBOX",
        "Số Note", "Trạng thái", "Chi tiết ngày"
    ]
    ws.row_dimensions[2].height = 26
    for c, h in enumerate(audit_headers, start=1):
        cell = ws.cell(2, c, h)
        cell.font = FONT_AUDIT_HDR
        cell.fill = SUBHEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)

    cur_row = 3
    for rdata in report.all_rows:
        ws.row_dimensions[cur_row].height = 20
        sk = rdata.status_kind
        if sk == "mismatch":
            status = "⚠ LỆCH DATA"
        elif sk == "pending":
            status = "⏳ Chưa có data"
        else:
            status = "✓ Khớp 100%"
        m_dates = rdata.status_dates_summary
        vals = [
            rdata.task,
            rdata.annotator,
            _clean_num(rdata.wl_frames),
            _clean_num(rdata.rw_frames),
            _clean_num(rdata.diff_frame),
            _clean_num(rdata.wl_bbox),
            _clean_num(rdata.rw_bbox),
            _clean_num(rdata.diff_bbox),
            rdata.note_count,
            status,
            m_dates,
        ]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(cur_row, c, v)
            cell.border = _make_border()
            cell.font = FONT_AUDIT_REG
            if c in (1, 2, 9, 10):
                cell.alignment = ALIGN_CENTER
            elif c == 11:
                cell.alignment = ALIGN_LEFT
            else:
                cell.alignment = ALIGN_RIGHT
            if sk == "mismatch":
                cell.fill = WARN_FILL
                if c in (5, 8, 10, 11):
                    cell.font = FONT_WARN
            elif sk == "pending":
                cell.fill = INFO_FILL
                if c in (10, 11):
                    cell.font = FONT_INFO
            elif c == 10:
                cell.fill = OK_FILL
                cell.font = FONT_OK
        cur_row += 1

    # Part 1B: Block-level Mismatch & Pending Review Breakdown (matched by Frame Start/End)
    if report.daily_diffs:
        real_cnt = sum(1 for d in report.daily_diffs if d.kind == "mismatch")
        pend_cnt = sum(1 for d in report.daily_diffs if d.kind == "pending_rework")
        cur_row += 2
        ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=13)
        ws.cell(
            cur_row, 1,
            f"CHI TIẾT ĐỐI CHIẾU THEO FRAME START/END: {real_cnt} BLOCK LỆCH DATA · {pend_cnt} BLOCK CHƯA CÓ DATA REVIEW"
        ).font = FONT_AUDIT_HDR
        for c in range(1, 14):
            ws.cell(cur_row, c).fill = DARK_HEADER_FILL
        ws.row_dimensions[cur_row].height = 24
        cur_row += 1

        diff_headers = [
            "Loại", "Annotator", "Ngày Work Log", "Ngày Review", "TASK", "Frame Start ~ End",
            "Work Log (Frame)", "Review (Frame)", "Lệch Frame",
            "Work Log (BBOX)", "Review (BBOX)", "Lệch BBOX",
            "Ghi chú / Nguyên nhân"
        ]
        ws.row_dimensions[cur_row].height = 24
        for c, h in enumerate(diff_headers, start=1):
            cell = ws.cell(cur_row, c, h)
            cell.font = FONT_AUDIT_HDR
            cell.fill = SUBHEADER_FILL
            cell.alignment = ALIGN_CENTER
            cell.border = _make_border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
        cur_row += 1

        for d in report.daily_diffs:
            ws.row_dimensions[cur_row].height = 20
            is_pend = d.is_pending
            kind_lbl = "⏳ Chưa có data" if is_pend else "⚠ LỆCH DATA"
            row_fill = INFO_FILL if is_pend else WARN_FILL
            accent_font = FONT_INFO if is_pend else FONT_WARN
            vals = [
                kind_lbl,
                d.annotator,
                d.date.strftime("%Y-%m-%d"),
                d.rw_date.strftime("%Y-%m-%d") if d.rw_date else "—",
                d.task,
                d.frame_range_display,
                _clean_num(d.wl_frames),
                _clean_num(d.rw_frames),
                _clean_num(d.diff_frame),
                _clean_num(d.wl_bbox),
                _clean_num(d.rw_bbox),
                _clean_num(d.diff_bbox),
                d.reason,
            ]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(cur_row, c, v)
                cell.border = _make_border()
                cell.fill = row_fill
                cell.font = accent_font if c in (1, 3, 6, 9, 12, 13) else FONT_AUDIT_REG
                cell.alignment = ALIGN_LEFT if c == 13 else (ALIGN_CENTER if c in (1, 2, 3, 4, 5, 6) else ALIGN_RIGHT)
            cur_row += 1

    # Part 2: Parsed Work_log Notes detail (removed Sheet & Cell columns for clean view)
    cur_row += 2
    ws.merge_cells(start_row=cur_row, start_column=1, end_row=cur_row, end_column=6)
    ws.cell(cur_row, 1, f"CHI TIẾT EXCEL NOTES ĐỌC ĐƯỢC TỪ WORK LOG ({len(report.filtered_notes)} notes)").font = FONT_AUDIT_HDR
    for c in range(1, 7):
        ws.cell(cur_row, c).fill = DARK_HEADER_FILL
    ws.row_dimensions[cur_row].height = 24
    cur_row += 1

    note_headers = ["Annotator", "Ngày", "TASK", "Work Log (Frame)", "Work Log (BBOX)", "Nội dung Note gốc"]
    ws.row_dimensions[cur_row].height = 24
    for c, h in enumerate(note_headers, start=1):
        cell = ws.cell(cur_row, c, h)
        cell.font = FONT_AUDIT_HDR
        cell.fill = SUBHEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(left=THIN_BLACK, right=THIN_BLACK, top=THIN_BLACK, bottom=THIN_BLACK)
    cur_row += 1

    for n in report.filtered_notes:
        ws.row_dimensions[cur_row].height = 20
        vals = [
            n.annotator_display,
            n.date.strftime("%Y-%m-%d"),
            n.task,
            _clean_num(n.frames_in_cell),
            n.parsed_bbox if n.parsed_bbox is not None else "",
            n.raw_note.replace("\n", " | "),
        ]
        for c, v in enumerate(vals, start=1):
            cell = ws.cell(cur_row, c, v)
            cell.border = _make_border()
            cell.font = FONT_AUDIT_REG
            cell.alignment = ALIGN_LEFT if c == 6 else (ALIGN_RIGHT if c in (4, 5) else ALIGN_CENTER)
        cur_row += 1

    widths = [16, 14, 16, 16, 13, 28, 16, 16, 13, 16, 26, 14, 36]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def export_report_to_excel(
    report: ReportResult,
    output_path: str,
    include_template1: bool = True,
    include_template2: bool = True,
    include_audit: bool = True,
    highlight_mismatch: bool = True,
):
    wb = openpyxl.Workbook()
    first_used = False

    if include_template1:
        _write_template1_by_task(wb, report, highlight_mismatch=highlight_mismatch)
        first_used = True

    if include_template2:
        if not first_used:
            # Remove default empty sheet after creating Template 2
            default_ws = wb.active
            _write_template2_by_annotator(wb, report, highlight_mismatch=highlight_mismatch)
            wb.remove(default_ws)
            first_used = True
        else:
            _write_template2_by_annotator(wb, report, highlight_mismatch=highlight_mismatch)

    # Always include the Reviewer sheet and Ranking sheet
    if first_used:
        _write_reviewer_sheet(wb, report)
        _write_ranking_sheet(wb, report)

    if include_audit:
        if not first_used:
            default_ws = wb.active
            _write_audit_and_notes_sheet(wb, report)
            wb.remove(default_ws)
        else:
            _write_audit_and_notes_sheet(wb, report)

    wb.save(output_path)
