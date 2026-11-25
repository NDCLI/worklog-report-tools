from collections import Counter
from typing import List, Optional

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from data_engine import InspectorSummary, ReportResult


DARK_HEADER_FILL = PatternFill(start_color="333333", end_color="333333", fill_type="solid")
SUBHEADER_FILL = PatternFill(start_color="444444", end_color="444444", fill_type="solid")
SUB_HEADER_FILL = SUBHEADER_FILL
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
FONT_GREEN_BOLD = Font(name="Times New Roman", size=10, bold=True, color="375623")
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
    ws.merge_cells("A1:R1")
    ws["A1"] = title_text
    ws["A1"].font = FONT_TITLE
    ws["A1"].fill = DARK_HEADER_FILL
    ws["A1"].alignment = ALIGN_CENTER
    ws.row_dimensions[1].height = 28
    for c in range(1, 19):
        ws.cell(1, c).fill = DARK_HEADER_FILL
        ws.cell(1, c).border = _make_border(
            left=MEDIUM_BLACK if c == 1 else THIN_BLACK,
            right=MEDIUM_BLACK if c == 18 else THIN_BLACK,
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
        "1st Err\n(BBOX)",
        "2nd Err\n(BBOX)",
        "1st Err\n(Frame)",
        "2nd Err\n(Frame)",
        "FRAME Per\nHour",
        "BBOX Per\nHour",
        "Err Rate\nBBOX (%)",
        "Err Rate\nFrame (%)",
        "TASK",
        "Avg\nFrame/h",
        "Avg\nBBOX/h",
        "Err Rate\nBBOX (%)",
        "Err Rate\nFrame (%)",
    ]
    ws.row_dimensions[2].height = 54
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(2, c, h)
        cell.font = FONT_HEADER
        cell.fill = DARK_HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(
            left=MEDIUM_BLACK if c in (1, 10, 14, 15) else THIN_BLACK,
            right=MEDIUM_BLACK if c in (9, 13, 14, 18) else THIN_BLACK,
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
            ws.cell(r, 6, _clean_num(rdata.err_first_bbox))
            ws.cell(r, 7, _clean_num(rdata.err_second_bbox))
            ws.cell(r, 8, _clean_num(rdata.err_first))
            ws.cell(r, 9, _clean_num(rdata.err_second))

            # Live Excel formulas for KPIs
            ws.cell(r, 10, f"=IF(E{r}>0,C{r}/E{r},0)")
            ws.cell(r, 11, f"=IF(E{r}>0,D{r}/E{r},0)")
            ws.cell(r, 12, f"=IF(D{r}>0,(F{r}+G{r})/D{r}*100,0)")
            ws.cell(r, 13, f"=IF(C{r}>0,(H{r}+I{r})/C{r}*100,0)")
            ws.cell(r, 14, grp.task if is_first else None)

            for c in range(1, 19):
                cell = ws.cell(r, c)
                top_b = MEDIUM_BLACK if is_first else THIN_GRAY
                bot_b = MEDIUM_BLACK if is_last else THIN_GRAY
                if c in (1, 14, 15, 16, 17, 18):
                    left_b = MEDIUM_BLACK if c in (1, 14, 15) else THIN_BLACK
                    right_b = MEDIUM_BLACK if c in (1, 14, 18) else THIN_BLACK
                    cell.border = _make_border(left=left_b, right=right_b, top=top_b, bottom=bot_b)
                    if c >= 15:
                        cell.fill = ZEBRA_FILL
                else:
                    left_b = MEDIUM_BLACK if c == 10 else THIN_GRAY
                    right_b = MEDIUM_BLACK if c in (9, 13) else THIN_GRAY
                    cell.border = _make_border(left=left_b, right=right_b, top=top_b, bottom=bot_b)

                if c in (1, 14):
                    cell.fill = DARK_HEADER_FILL
                    cell.font = FONT_TASK
                    cell.alignment = ALIGN_CENTER
                elif c == 2:
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_CENTER
                    if highlight_mismatch and rdata.has_mismatch:
                        cell.fill = WARN_FILL
                elif c in (3, 4, 5, 6, 7, 8, 9):
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.##" if c == 5 else "0"
                    if highlight_mismatch and rdata.has_mismatch and c in (3, 4):
                        cell.fill = WARN_FILL
                elif c in (10, 11):
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"
                elif c in (12, 13):
                    cell.font = FONT_RED
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"

        # Merge TASK cols (1 and 14) and 4 Average cols (15, 16, 17, 18)
        if end_r > start_r:
            ws.merge_cells(start_row=start_r, start_column=1, end_row=end_r, end_column=1)
            for c in (14, 15, 16, 17, 18):
                ws.merge_cells(start_row=start_r, start_column=c, end_row=end_r, end_column=c)

        c15 = ws.cell(start_r, 15, f"=IF(SUM(E{start_r}:E{end_r})>0,SUM(C{start_r}:C{end_r})/SUM(E{start_r}:E{end_r}),0)")
        c16 = ws.cell(start_r, 16, f"=IF(SUM(E{start_r}:E{end_r})>0,SUM(D{start_r}:D{end_r})/SUM(E{start_r}:E{end_r}),0)")
        c17 = ws.cell(start_r, 17, f"=IF(SUM(D{start_r}:D{end_r})>0,(SUM(F{start_r}:F{end_r})+SUM(G{start_r}:G{end_r}))/SUM(D{start_r}:D{end_r})*100,0)")
        c18 = ws.cell(start_r, 18, f"=IF(SUM(C{start_r}:C{end_r})>0,(SUM(H{start_r}:H{end_r})+SUM(I{start_r}:I{end_r}))/SUM(C{start_r}:C{end_r})*100,0)")
        for avg_cell in (c15, c16, c17, c18):
            avg_cell.font = FONT_RED_BOLD
            avg_cell.alignment = ALIGN_CENTER
            avg_cell.number_format = "0.00"

        cur_row = end_r + 1

    widths = [11, 13, 12, 12, 11, 11, 11, 11, 11, 13, 13, 13, 13, 11, 13, 13, 13, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "C3"


def _write_template2_by_annotator(wb: openpyxl.Workbook, report: ReportResult, highlight_mismatch: bool = True):
    ws = wb.create_sheet(title="Annotator")
    ws.views.sheetView[0].showGridLines = True

    title_text = f"{report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')}"

    # Row 1: Title banner
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

    # Row 2: Column Headers
    headers = [
        "Annotator",
        "Work\nLoad",
        "Work\nBBOX",
        "Work Time",
        "1st Err\n(BBOX)",
        "2nd Err\n(BBOX)",
        "1st Err\n(Frame)",
        "2nd Err\n(Frame)",
        "FRAME Per\nHour",
        "BBOX Per\nHour",
        "Err Rate\nBBOX (%)",
        "Err Rate\nFrame (%)",
    ]
    ws.row_dimensions[2].height = 54
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(2, c, h)
        cell.font = FONT_HEADER
        cell.fill = DARK_HEADER_FILL
        cell.alignment = ALIGN_CENTER
        cell.border = _make_border(
            left=MEDIUM_BLACK if c in (1, 9) else THIN_BLACK,
            right=MEDIUM_BLACK if c in (8, 12) else THIN_BLACK,
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
        ws.cell(r, 5, _clean_num(rdata.err_first_bbox))
        ws.cell(r, 6, _clean_num(rdata.err_second_bbox))
        ws.cell(r, 7, _clean_num(rdata.err_first))
        ws.cell(r, 8, _clean_num(rdata.err_second))

        ws.cell(r, 9, f"=IF(D{r}>0,B{r}/D{r},0)")
        ws.cell(r, 10, f"=IF(D{r}>0,C{r}/D{r},0)")
        ws.cell(r, 11, f"=IF(C{r}>0,(E{r}+F{r})/C{r}*100,0)")
        ws.cell(r, 12, f"=IF(B{r}>0,(G{r}+H{r})/B{r}*100,0)")

        for c in range(1, 13):
            cell = ws.cell(r, c)
            left_b = MEDIUM_BLACK if c in (1, 9) else THIN_GRAY
            right_b = MEDIUM_BLACK if c in (8, 12) else THIN_GRAY
            bot_b = MEDIUM_BLACK if is_last else THIN_GRAY
            cell.border = _make_border(left=left_b, right=right_b, top=THIN_GRAY, bottom=bot_b)

            if c == 1:
                cell.font = FONT_BOLD
                cell.alignment = ALIGN_CENTER
                if highlight_mismatch and rdata.has_mismatch:
                    cell.fill = WARN_FILL
            elif c in (2, 3, 4, 5, 6, 7, 8):
                cell.font = FONT_REGULAR
                cell.alignment = ALIGN_RIGHT
                cell.number_format = "0.##" if c == 4 else "0"
                if highlight_mismatch and rdata.has_mismatch and c in (2, 3):
                    cell.fill = WARN_FILL
            elif c in (9, 10):
                cell.font = FONT_BOLD
                cell.alignment = ALIGN_RIGHT
                cell.number_format = "0.00"
            elif c in (11, 12):
                cell.font = FONT_RED
                cell.alignment = ALIGN_RIGHT
                cell.number_format = "0.00"

    widths = [15, 11, 12, 11, 11, 11, 11, 11, 13, 13, 13, 13]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "B3"


def _write_inspector_sheet(wb: openpyxl.Workbook, report: ReportResult, groups: List[InspectorSummary]):
    ws = wb.create_sheet(title="Inspector")
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
    for grp in groups:
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
        [r for r in report.annotator_rows if not r.is_reviewer and r.work_load > 0],
        key=lambda r: (r.total_err_rate, -r.work_load, r.annotator.upper()),
    )
    reviewers = sorted(
        [g for g in report.inspector_groups if g.total_work_load > 0],
        key=lambda g: (g.err_rate_frame, -g.total_work_load, g.inspector.upper()),
    )

    def write_rank_block(start_row: int, title: str, name_hdr: str, rows_list, is_rev: bool = False) -> int:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=3)
        title_cell = ws.cell(start_row, 1, title)
        title_cell.font = FONT_AUDIT_HDR
        if is_rev:
            title_cell.alignment = ALIGN_LEFT
        ws.row_dimensions[start_row].height = 36 if is_rev else 24
        for c in range(1, 4):
            ws.cell(start_row, c).fill = DARK_HEADER_FILL

        hdr_row = start_row + 1
        hdrs = [
            "Hạng",
            name_hdr,
            "Err Rate Frame (%)",
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
                ws.cell(cur, 3, round(rdata.err_rate_frame, 2))
            else:
                ws.cell(cur, 2, rdata.annotator)
                ws.cell(cur, 3, round(rdata.total_err_rate, 2))

            for c in range(1, 4):
                cell = ws.cell(cur, c)
                cell.border = _make_border(THIN_GRAY, THIN_GRAY, THIN_GRAY, THIN_GRAY)
                if rank_idx <= 3:
                    cell.fill = OK_FILL
                if c in (1, 2):
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_CENTER
                else:
                    cell.font = FONT_RED_BOLD if rank_idx > 3 else FONT_GREEN_BOLD
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"
            cur += 1
        return cur + 2

    next_r = write_rank_block(
        1,
        f"BẢNG XẾP HẠNG ANNOTATOR (WORKER) — TỈ LỆ LỖI FRAME THẤP NHẤT ({report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')})",
        "Annotator (Worker)",
        workers,
        is_rev=False,
    )
    next_r = write_rank_block(
        next_r,
        f"BẢNG XẾP HẠNG INSPECTOR (LẦN 1) — {report.start_date:%Y.%m.%d} ~ {report.end_date:%Y.%m.%d}\n"
        "Tỉ lệ lỗi Frame thấp nhất theo Annotator đang chọn",
        "Inspector (Lần 1)",
        reviewers,
        is_rev=True,
    )

    # Also write Performance Ranking (Low Frame Error Rate + High Volume, 50/50)
    def _score_perf(items, get_load, get_bbox, get_err):
        if not items:
            return []
        errs = [get_err(x) for x in items]
        min_e, max_e = min(errs), max(errs)
        max_l = max((get_load(x) for x in items), default=0.0)
        max_b = max((get_bbox(x) for x in items), default=0.0)
        out = []
        for x in items:
            q = ((max_e - get_err(x)) / (max_e - min_e) * 100.0) if max_e > min_e else 100.0
            vf = (get_load(x) / max_l * 100.0) if max_l > 0 else 0.0
            vb = (get_bbox(x) / max_b * 100.0) if max_b > 0 else vf
            tot = 0.5 * q + 0.5 * (0.5 * vf + 0.5 * vb)
            out.append((x, tot))
        out.sort(key=lambda t: (-t[1], get_err(t[0]), -get_load(t[0])))
        return out

    def write_perf_block(start_row: int, title: str, name_hdr: str, scored_list, is_rev: bool = False) -> int:
        ws.merge_cells(start_row=start_row, start_column=1, end_row=start_row, end_column=6)
        tcell = ws.cell(start_row, 1, title)
        tcell.font = FONT_AUDIT_HDR
        ws.row_dimensions[start_row].height = 24
        for c in range(1, 7):
            ws.cell(start_row, c).fill = DARK_HEADER_FILL
        hdr_row = start_row + 1
        hdrs = ["Hạng", name_hdr, "Work Load (Frame)", "Work BBOX", "Err Rate Frame (%)", "Điểm Tổng"]
        ws.row_dimensions[hdr_row].height = 28
        for c, h in enumerate(hdrs, start=1):
            cell = ws.cell(hdr_row, c, h)
            cell.font = FONT_HEADER
            cell.fill = SUB_HEADER_FILL
            cell.alignment = ALIGN_CENTER
            cell.border = _make_border(THIN_BLACK, THIN_BLACK, THIN_BLACK, MEDIUM_BLACK)
        cur = hdr_row + 1
        for rank_idx, (rdata, score) in enumerate(scored_list, start=1):
            ws.row_dimensions[cur].height = 20
            ws.cell(cur, 1, rank_idx)
            if is_rev:
                ws.cell(cur, 2, rdata.inspector)
                ws.cell(cur, 3, _clean_num(rdata.total_work_load))
                ws.cell(cur, 4, _clean_num(rdata.total_work_bbox))
                ws.cell(cur, 5, round(rdata.err_rate_frame, 2))
            else:
                ws.cell(cur, 2, rdata.annotator)
                ws.cell(cur, 3, _clean_num(rdata.work_load))
                ws.cell(cur, 4, _clean_num(rdata.work_bbox))
                ws.cell(cur, 5, round(rdata.total_err_rate, 2))
            ws.cell(cur, 6, round(score, 1))
            for c in range(1, 7):
                cell = ws.cell(cur, c)
                cell.border = _make_border(THIN_GRAY, THIN_GRAY, THIN_GRAY, THIN_GRAY)
                if rank_idx <= 3:
                    cell.fill = OK_FILL
                if c in (1, 2):
                    cell.font = FONT_BOLD
                    cell.alignment = ALIGN_CENTER
                elif c in (3, 4):
                    cell.font = FONT_REGULAR
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "#,##0"
                elif c == 5:
                    cell.font = FONT_RED_BOLD if rank_idx > 3 else FONT_GREEN_BOLD
                    cell.alignment = ALIGN_RIGHT
                    cell.number_format = "0.00"
                else:
                    cell.font = FONT_GREEN_BOLD if rank_idx <= 3 else FONT_BOLD
                    cell.alignment = ALIGN_CENTER
                    cell.number_format = "0.0"
            cur += 1
        return cur + 2

    w_perf = _score_perf(workers, lambda r: r.work_load, lambda r: r.work_bbox, lambda r: r.total_err_rate)
    r_perf = _score_perf(reviewers, lambda g: g.total_work_load, lambda g: g.total_work_bbox, lambda g: g.err_rate_frame)
    next_r = write_perf_block(
        next_r,
        f"HIỆU SUẤT ANNOTATOR — LỖI THẤP & SẢN LƯỢNG NHIỀU ({report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')})",
        "Annotator (Worker)",
        w_perf,
        is_rev=False,
    )
    write_perf_block(
        next_r,
        f"HIỆU SUẤT INSPECTOR — LỖI THẤP & SẢN LƯỢNG NHIỀU ({report.start_date.strftime('%Y.%m.%d')} ~ {report.end_date.strftime('%Y.%m.%d')})",
        "Inspector (Lần 1)",
        r_perf,
        is_rev=True,
    )

    widths = [10, 22, 18, 16, 18, 14]
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
            status = "⚠ KHÔNG KHỚP"
        elif sk == "pending":
            only_miss_bb = bool(rdata.pending_diffs) and all(d.is_missing_bbox for d in rdata.pending_diffs)
            status = "⏳ Chưa nhập BBOX" if only_miss_bb else "⏳ Chưa có data"
        else:
            status = "✓ Khớp"
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
            cell.alignment = ALIGN_CENTER
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
            f"CHI TIẾT ĐỐI CHIẾU THEO FRAME START/END: {real_cnt} BLOCK KHÔNG KHỚP · {pend_cnt} BLOCK CHƯA CÓ DATA / BBOX"
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
            if d.is_missing_bbox:
                kind_lbl = "⏳ Chưa nhập BBOX"
            elif is_pend:
                kind_lbl = "⏳ Chưa có data"
            elif d.is_dist_frame_mismatch:
                kind_lbl = "⚠ SAI FRAME"
            elif d.is_job_id_mismatch:
                kind_lbl = "⚠ SAI ID JOB"
            else:
                kind_lbl = "⚠ KHÔNG KHỚP"
            row_fill = INFO_FILL if is_pend else WARN_FILL
            accent_font = FONT_INFO if is_pend else FONT_WARN
            miss_rw_bb = (d.rw_frames > 0 and d.wl_bbox > 0 and d.rw_bbox <= 0)
            miss_wl_bb = (d.wl_frames > 0 and d.rw_bbox > 0 and d.wl_bbox <= 0)
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
                "—" if miss_wl_bb else _clean_num(d.wl_bbox),
                "—" if miss_rw_bb else _clean_num(d.rw_bbox),
                "—" if (miss_rw_bb or miss_wl_bb) else _clean_num(d.diff_bbox),
                d.reason,
            ]
            for c, v in enumerate(vals, start=1):
                cell = ws.cell(cur_row, c, v)
                cell.border = _make_border()
                cell.fill = row_fill
                cell.font = accent_font if c in (1, 3, 6, 9, 12, 13) else FONT_AUDIT_REG
                cell.alignment = ALIGN_CENTER
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
    include_reviewer: bool = True,
    include_ranking: bool = True,
    include_audit: bool = True,
    highlight_mismatch: bool = True,
    inspector_groups: Optional[List[InspectorSummary]] = None,
):
    wb = openpyxl.Workbook()
    default_ws = wb.active
    used_default = False

    if include_template1:
        _write_template1_by_task(wb, report, highlight_mismatch=highlight_mismatch)
        used_default = True

    if include_template2:
        _write_template2_by_annotator(wb, report, highlight_mismatch=highlight_mismatch)

    if include_reviewer:
        groups = report.inspector_groups if inspector_groups is None else inspector_groups
        _write_inspector_sheet(wb, report, groups)

    if include_ranking:
        _write_ranking_sheet(wb, report)

    if include_audit:
        _write_audit_and_notes_sheet(wb, report)

    if not used_default and len(wb.sheetnames) > 1 and default_ws in wb.worksheets:
        wb.remove(default_ws)

    wb.save(output_path)
