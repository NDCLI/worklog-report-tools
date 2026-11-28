import datetime
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from typing import Dict, Iterable, List, Optional, Tuple

import openpyxl
from openpyxl.utils import get_column_letter


MONTH_MAP = {
    "january": 1, "jan": 1,
    "february": 2, "feb": 2,
    "march": 3, "mar": 3,
    "april": 4, "apr": 4,
    "may": 5,
    "june": 6, "jun": 6,
    "july": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "october": 10, "oct": 10,
    "november": 11, "nov": 11,
    "december": 12, "dec": 12,
}

# Preferred display names matching team report conventions
PREFERRED_DISPLAY_NAMES = {
    "CLINH": "CLINH",
    "DHIEN": "DHIEN",
    "DUNG": "DUNG",
    "HOA": "HOA",
    "HOAHT": "HOA.ht",
    "HUONG": "HUONG",
    "HUYEN": "HUYEN",
    "LANH": "LANH",
    "MAI": "MAI",
    "NMAI": "NMAI",
    "OANH": "OANH",
    "TIEN": "TIEN",
    "TOAN": "TOAN",
    "TONG": "TONG",
    "TU": "TU",
    "TUYET": "TUYET",
    "CHINH": "CHINH",
    "TAM": "TAM",
    # 7 First Reviewers (INSPECTOR Lần 1)
    "MLINH": "MLINH",
    "THUY": "THUY",
    "NGAN": "NGAN",
    "XUAN": "XUAN",
    "HIEN03": "HIEN_03",
    "KNGAN": "KNGAN",
    "NGUYEN": "NGUYEN",
}

FIRST_REVIEWER_KEYS = {
    "MLINH", "THUY", "NGAN", "XUAN", "HIEN03", "KNGAN", "NGUYEN"
}

# Skip sample & legacy sheets in Work_log
SKIP_NON_DATA_WORKLOG_SHEETS = {
    "SAMPLE", "HIEN", "CUC"
}
# Legacy alias kept for compatibility
SKIP_WORKLOG_SHEETS = SKIP_NON_DATA_WORKLOG_SHEETS | {
    "MLINH", "THUY", "NGAN", "XUAN", "HIEN_03", "KNGAN", "NGUYEN"
}


def strip_rew_suffix(task: str) -> str:
    t = normalize_task(task)
    return t[:-4] if t.endswith("_REW") else t

# Skip example & legacy (~1 year old, 2024 / early 2025) sheets in Inspection_And_Rework
SKIP_REWORK_SHEETS = {
    "EXAMPLE(ANPR_ENG_230704)",
    "THUY(",
    "PANKYO (HIEN)",
    "HIEN(HOA,XAY,CLINH,CUC)",
    "CUC (HUE,TAM,HIEN)",
}

# Inactive / legacy people (~1 year old or no longer active)
INACTIVE_ANNOTATOR_KEYS = {
    "CUC", "HIEN", "HUE", "XAY", "PANKYO", "MJ", "SJ", "HJ"
}

SUBROW_KEYWORDS = ("_PASS", "_START", "_END", "START_FRAME", "END_FRAME", "_SKIP")


def normalize_annotator_key(name: str) -> str:
    """Normalize annotator/reviewer name into an alphanumeric uppercase key for reliable matching."""
    if not name:
        return ""
    cleaned = re.sub(r"[^A-Za-z0-9]", "", str(name).strip().upper())
    if cleaned == "MYLINH":
        return "MLINH"
    return cleaned


def display_annotator_name(raw_name: str) -> str:
    key = normalize_annotator_key(raw_name)
    if key in PREFERRED_DISPLAY_NAMES:
        return PREFERRED_DISPLAY_NAMES[key]
    return str(raw_name).strip().upper()


def unique_inspector_names(raw_names: Iterable[str]) -> Tuple[str, ...]:
    names_by_key = {}
    for name in raw_names:
        display = display_annotator_name(name)
        if display:
            key = normalize_annotator_key(name) or display.casefold()
            names_by_key.setdefault(key, display)
    return tuple(sorted(names_by_key.values()))


def normalize_task(task: str) -> str:
    if not task:
        return ""
    return str(task).strip().upper()


def normalize_job_name(name: str) -> str:
    """Normalize Task/Job Name (e.g. 'HOPJT_#2', 'OD_HOPJT_vid_1') for matching against Distribution Sheet."""
    if not name:
        return ""
    return re.sub(r"\s+", "", str(name).strip().upper())


def to_date(val) -> Optional[datetime.date]:
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    if isinstance(val, str):
        s = val.strip()
        for fmt in ("%Y-%m-%d", "%Y.%m.%d", "%Y/%m/%d", "%d/%m/%Y", "%m/%d/%Y"):
            try:
                return datetime.datetime.strptime(s[:10], fmt).date()
            except ValueError:
                pass
    return None


def parse_box_from_note(note: str) -> Tuple[Optional[int], List[int]]:
    """
    Parse BBOX counts from an Excel cell Note.
    Handles formats like:
      - Box: 887 / BBox:179 / box; 230 / box 169 / lBox:727
      - Multiple blocks inside a single note (sums all occurrences)
    Returns (total_or_None, list_of_numbers).
    """
    if not note:
        return None, []
    matches = re.findall(r"(?:b?box|lbox)\s*[:;=\-\s]\s*(\d+)", note, flags=re.IGNORECASE)
    if not matches:
        return None, []
    nums = [int(m) for m in matches]
    return sum(nums), nums


def load_xlsx_comments(xlsx_path: str) -> Dict[str, Dict[str, str]]:
    """
    Fast XML reader for ALL Excel Notes (legacy comments) and Threaded Comments.
    Returns: {sheet_name: {cell_ref (e.g. 'GW32'): comment_text}}
    """
    result: Dict[str, Dict[str, str]] = {}
    ns_main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    ns_tc = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"

    with zipfile.ZipFile(xlsx_path, "r") as z:
        namelist = set(z.namelist())
        wb_xml = ET.fromstring(z.read("xl/workbook.xml"))
        wb_rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        rid_to_target = {rel.attrib["Id"]: rel.attrib["Target"] for rel in wb_rels}

        for sheet_el in wb_xml.findall(f".//{{{ns_main}}}sheet"):
            sname = sheet_el.attrib["name"]
            rid = sheet_el.attrib.get(f"{{{ns_rel}}}id")
            if not rid or rid not in rid_to_target:
                continue
            target = rid_to_target[rid]
            if not target.startswith("xl/"):
                target = "xl/" + target.lstrip("/")
            sheet_dir, sheet_file = target.rsplit("/", 1)
            rels_path = f"{sheet_dir}/_rels/{sheet_file}.rels"
            c_map: Dict[str, str] = {}
            tc_map: Dict[str, List[str]] = defaultdict(list)

            if rels_path in namelist:
                s_rels = ET.fromstring(z.read(rels_path))
                for rel in s_rels:
                    rtype = rel.attrib.get("Type", "").lower()
                    if "comment" not in rtype:
                        continue
                    ctarget = rel.attrib["Target"]
                    parts = (sheet_dir + "/" + ctarget).split("/")
                    norm: List[str] = []
                    for p in parts:
                        if p == "..":
                            if norm:
                                norm.pop()
                        elif p and p != ".":
                            norm.append(p)
                    cpath = "/".join(norm)
                    if cpath not in namelist:
                        continue
                    croot = ET.fromstring(z.read(cpath))
                    if "threadedcomments" in rtype:
                        for tc in croot.findall(f".//{{{ns_tc}}}threadedComment"):
                            ref = tc.attrib.get("ref")
                            text_el = tc.find(f"{{{ns_tc}}}text")
                            txt = "".join(text_el.itertext()).strip() if text_el is not None else "".join(tc.itertext()).strip()
                            if ref and txt:
                                tc_map[ref].append(txt)
                    else:
                        for c in croot.findall(f".//{{{ns_main}}}comment"):
                            ref = c.attrib.get("ref")
                            txt = "".join(c.itertext()).strip()
                            if ref and txt:
                                c_map[ref] = txt

            # Prefer real threaded comment text over the Excel placeholder text
            for ref, msgs in tc_map.items():
                c_map[ref] = "\n".join(msgs)

            result[sname] = c_map
    return result


@dataclass
class WorklogNoteRecord:
    annotator_key: str
    annotator_display: str
    sheet_name: str
    date: datetime.date
    task: str
    cell_ref: str
    frames_in_cell: float
    parsed_bbox: Optional[int]
    raw_note: str


def _parse_int(val) -> Optional[int]:
    if isinstance(val, (int, float)):
        return int(round(val))
    if isinstance(val, str):
        s = val.strip().replace(",", "")
        m = re.search(r"-?\d+", s)
        if m:
            return int(m.group(0))
    return None


def _extract_job_id(url: str) -> str:
    if not url:
        return ""
    m = re.search(r"/jobs/(\d+)", str(url))
    return m.group(1) if m else ""


def _extract_cvat_task_id(url: str) -> str:
    if not url:
        return ""
    m = re.search(r"/tasks/(\d+)", str(url))
    return m.group(1) if m else ""


def _fmt_range(sf: Optional[int], ef: Optional[int]) -> str:
    if sf is not None and ef is not None:
        return f"{sf}~{ef}"
    if sf is not None:
        return f"{sf}~?"
    if ef is not None:
        return f"?~{ef}"
    return ""


@dataclass
class WorklogBlockRecord:
    annotator_key: str
    annotator_display: str
    sheet_name: str
    date: datetime.date
    task: str
    job_id: str
    start_frame: Optional[int]
    end_frame: Optional[int]
    frames: float
    bbox: int
    cell_ref: str
    job_name: str = ""                         # Col E (Job / Task Name, e.g. 'HOPJT_#2', 'OD_HOPJT_vid_1')


@dataclass
class DistributionJobRecord:
    sheet_name: str
    row_num: int
    task_name: str                             # Raw TASK NAME from distribution sheet (e.g. 'HOPJT_#2')
    task_name_key: str                         # Normalized key (e.g. 'HOPJT_#2')
    job_url: str                               # Full JOB URL (e.g. 'http://.../jobs/4410')
    job_id: str                                # Trailing ID after /jobs/ (e.g. '4410')
    start_frame: int
    end_frame: int
    assignee_raw: str = ""                     # Raw NAME column text (e.g. 'C.Linh (33050-)')
    assignee_keys: Tuple[str, ...] = ()        # Normalized annotator keys (e.g. ('CLINH',))
    cvat_task_id: str = ""                     # CVAT Task ID after /tasks/ (e.g. '735')
    reviewer_raw: str = ""                     # Raw INSPECTION column text (e.g. 'Hien.03(32500-32559)')
    reviewer_keys: Tuple[str, ...] = ()        # Normalized reviewer keys (e.g. ('HIEN03',))


@dataclass
class WorklogDailyRecord:
    annotator_key: str
    annotator_display: str
    sheet_name: str
    date: datetime.date
    task: str
    frames: float = 0.0
    bbox: int = 0
    work_time: float = 0.0
    cell_refs: List[str] = field(default_factory=list)


@dataclass
class ReworkRecord:
    sheet_name: str
    row_num: int
    date_inspect: Optional[datetime.date]      # Col A (DATE)
    date_work: Optional[datetime.date]         # Col V (Date of Secondary Review / actual work date)
    task: str
    worker_key: str
    worker_display: str
    inspector: str
    job_url: str
    frames: float                              # Col O row 1
    bbox: float                                # Col O row 2
    err_first: float                           # Col Q (ISSUE FRAME)
    err_first_bbox: float                      # Col R (ISSUE BBOX)
    err_second: float                          # Col S (PANGYO FEEDBACK FRAME)
    err_second_bbox: float                     # Col T (PANGYO FEEDBACK BBOX)
    work_time: float = 0.0                     # Col P (WORKING TIME for 1st Inspector)
    notes: str = ""
    job_id: str = ""
    start_frame: Optional[int] = None          # Col E (WORKING FRAME START)
    end_frame: Optional[int] = None            # Col F (WORKING FRAME END)


@dataclass
class LoadedDataset:
    worklog_path: str = ""
    rework_path: str = ""
    distribution_path: str = ""
    worklog_daily: List[WorklogDailyRecord] = field(default_factory=list)
    worklog_blocks: List[WorklogBlockRecord] = field(default_factory=list)
    worklog_notes: List[WorklogNoteRecord] = field(default_factory=list)
    reviewer_worklog_daily: List[WorklogDailyRecord] = field(default_factory=list)
    reviewer_worklog_blocks: List[WorklogBlockRecord] = field(default_factory=list)
    reviewer_worklog_notes: List[WorklogNoteRecord] = field(default_factory=list)
    rework_records: List[ReworkRecord] = field(default_factory=list)
    distribution_jobs: List[DistributionJobRecord] = field(default_factory=list)
    annotators: Dict[str, str] = field(default_factory=dict)  # key -> display name
    tasks: List[str] = field(default_factory=list)
    min_date: Optional[datetime.date] = None
    max_date: Optional[datetime.date] = None
    warnings: List[str] = field(default_factory=list)


def _resolve_col_dates(rows: list, month_row_idx: Optional[int], date_row_idx: int) -> Dict[int, datetime.date]:
    """
    Resolve column index -> true datetime.date.
    Fixes Excel sheets where the Month header says 'November' but hidden date serials
    were copied from September/October, or where early cells have a typo year.
    """
    m_row = rows[month_row_idx] if month_row_idx is not None else []
    d_row = rows[date_row_idx]

    # Segment columns by month header
    segments: List[Tuple[int, int, Optional[int]]] = []  # (start_col, end_col, month_num)
    month_starts: List[Tuple[int, int]] = []
    for c_i in range(8, len(d_row)):
        if c_i < len(m_row) and m_row[c_i] is not None:
            m_str = str(m_row[c_i]).strip().lower()
            if m_str in MONTH_MAP:
                month_starts.append((c_i, MONTH_MAP[m_str]))

    if not month_starts:
        # Fallback: trust date cells directly
        col_dates: Dict[int, datetime.date] = {}
        for c_i in range(8, len(d_row)):
            d = to_date(d_row[c_i])
            if d:
                col_dates[c_i] = d
        return col_dates

    # Prefix before first month header (if any)
    if month_starts[0][0] > 8:
        segments.append((8, month_starts[0][0], None))
    for idx, (sc, m_num) in enumerate(month_starts):
        ec = month_starts[idx + 1][0] if idx + 1 < len(month_starts) else len(d_row)
        segments.append((sc, ec, m_num))

    # Determine majority year across the whole date row as fallback
    all_years = [to_date(v).year for v in d_row[8:] if to_date(v) is not None]
    default_year = Counter(all_years).most_common(1)[0][0] if all_years else datetime.date.today().year

    col_dates = {}
    for sc, ec, header_month in segments:
        seg_years = [to_date(d_row[c]).year for c in range(sc, ec) if c < len(d_row) and to_date(d_row[c]) is not None]
        seg_year = Counter(seg_years).most_common(1)[0][0] if seg_years else default_year
        for c in range(sc, ec):
            if c >= len(d_row):
                continue
            v = d_row[c]
            d = to_date(v)
            day = None
            cell_month = None
            cell_year = None
            if d:
                day = d.day
                cell_month = d.month
                cell_year = d.year
            elif isinstance(v, (int, float)) and 1 <= int(v) <= 31:
                day = int(v)

            if day is None:
                continue
            m = header_month if header_month is not None else cell_month
            y = seg_year if header_month is not None else (cell_year or default_year)
            if m and y:
                try:
                    col_dates[c] = datetime.date(y, m, day)
                except ValueError:
                    pass
    return col_dates


def parse_worklog(
    xlsx_path: str, progress_cb=None
) -> Tuple[
    List[WorklogDailyRecord],
    List[WorklogBlockRecord],
    List[WorklogNoteRecord],
    List[str],
    List[WorklogDailyRecord],
    List[WorklogBlockRecord],
    List[WorklogNoteRecord],
]:
    daily_records: List[WorklogDailyRecord] = []
    block_records: List[WorklogBlockRecord] = []
    note_records: List[WorklogNoteRecord] = []
    rev_daily: List[WorklogDailyRecord] = []
    rev_blocks: List[WorklogBlockRecord] = []
    rev_notes: List[WorklogNoteRecord] = []
    warnings: List[str] = []

    if progress_cb:
        progress_cb(5, "Đang đọc toàn bộ Excel Notes trong Work_log...")
    comments_by_sheet = load_xlsx_comments(xlsx_path)

    if progress_cb:
        progress_cb(15, "Đang mở file Work_log...")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    total_sheets = max(len(sheet_names), 1)

    for s_idx, sname in enumerate(sheet_names):
        if progress_cb:
            pct = 15 + int((s_idx / total_sheets) * 45)
            progress_cb(pct, f"Đang đọc Work_log sheet: {sname} ({s_idx+1}/{total_sheets})")

        if sname.strip().upper() in SKIP_NON_DATA_WORKLOG_SHEETS:
            continue

        ws = wb[sname]
        rows = list(ws.iter_rows(values_only=True))
        if len(rows) < 5:
            continue

        month_r = None
        date_r = None
        for idx, r in enumerate(rows[:45]):
            if month_r is None and any(str(v).strip().lower() in MONTH_MAP for v in r if v is not None):
                month_r = idx
            if date_r is None and len(r) > 4 and str(r[4]).strip().lower() == "job":
                date_r = idx

        # Fallback date row search: row with >= 10 date values in cols >= 8
        if date_r is None:
            for idx, r in enumerate(rows[:45]):
                cnt = sum(1 for v in r[8:] if to_date(v) is not None)
                if cnt >= 10:
                    date_r = idx
                    break

        if date_r is None:
            continue

        col_dates = _resolve_col_dates(rows, month_r, date_r)
        if not col_dates:
            continue

        ann_key = normalize_annotator_key(sname)
        ann_disp = display_annotator_name(sname)
        c_map = comments_by_sheet.get(sname, {})

        # Accumulate per (date, task)
        agg: Dict[Tuple[datetime.date, str], WorklogDailyRecord] = {}

        def get_rec(dt: datetime.date, task: str) -> WorklogDailyRecord:
            k = (dt, task)
            if k not in agg:
                agg[k] = WorklogDailyRecord(
                    annotator_key=ann_key,
                    annotator_display=ann_disp,
                    sheet_name=sname,
                    date=dt,
                    task=task,
                )
            return agg[k]

        # 1. Read MH (Work Time) section above date_r
        in_mh = False
        for r_i in range(date_r):
            r = rows[r_i]
            e_val = str(r[4]).strip().upper() if len(r) > 4 and r[4] is not None else ""
            g_val = str(r[6]).strip().upper() if len(r) > 6 and r[6] is not None else ""
            if e_val == "MH":
                in_mh = True
            elif "FRAME/HOUR" in e_val or g_val == "TOTAL_TIME":
                in_mh = False
            if in_mh and g_val:
                task_name = normalize_task(g_val)
                for c_i, dt in col_dates.items():
                    if c_i < len(r):
                        v = r[c_i]
                        if isinstance(v, (int, float)) and v > 0:
                            get_rec(dt, task_name).work_time += float(v)

        # 2. Read Task rows below date_r (Frames + Notes for BBOX + Start/End Frame blocks)
        for r_i in range(date_r + 1, len(rows)):
            r = rows[r_i]
            if len(r) <= 1 or r[1] is None:
                continue
            mode = normalize_task(r[1])
            if not mode or any(kw in mode for kw in SUBROW_KEYWORDS):
                continue
            # Keep _REW tasks ONLY for the 7 First Reviewers; skip any _REW summary row on a Worker sheet
            if mode.endswith("_REW") and ann_key not in FIRST_REVIEWER_KEYS:
                continue

            # Job Name (col E), Job ID (col F) and paired start_frame / end_frame sub-rows directly below r_i
            job_name_str = str(r[4]).strip() if (len(r) > 4 and r[4] is not None) else ""
            job_id_val = _parse_int(r[5]) if len(r) > 5 else None
            job_id_str = str(job_id_val) if job_id_val is not None else ""
            start_row, end_row = None, None
            for off in (1, 2, 3):
                if r_i + off < len(rows):
                    sub_m = str(rows[r_i + off][1] or "").strip().upper()
                    if "START" in sub_m:
                        start_row = rows[r_i + off]
                    elif "END" in sub_m:
                        end_row = rows[r_i + off]

            for c_i, dt in col_dates.items():
                v = r[c_i] if c_i < len(r) else None
                frames_val = float(v) if isinstance(v, (int, float)) and v > 0 else 0.0
                ref = f"{get_column_letter(c_i + 1)}{r_i + 1}"
                note_txt = c_map.get(ref, "")
                bbox_val = 0

                if frames_val > 0:
                    rec = get_rec(dt, mode)
                    rec.frames += frames_val
                    if ref not in rec.cell_refs:
                        rec.cell_refs.append(ref)

                # Only count Note BBOX on cells that actually have frames > 0 (ignores stale notes on cleared cells)
                if note_txt and frames_val > 0:
                    parsed_bb, _ = parse_box_from_note(note_txt)
                    if parsed_bb is not None and parsed_bb > 0:
                        bbox_val = parsed_bb
                        rec = get_rec(dt, mode)
                        rec.bbox += bbox_val
                        if ref not in rec.cell_refs:
                            rec.cell_refs.append(ref)
                    target_notes = rev_notes if ann_key in FIRST_REVIEWER_KEYS else note_records
                    target_notes.append(WorklogNoteRecord(
                        annotator_key=ann_key,
                        annotator_display=ann_disp,
                        sheet_name=sname,
                        date=dt,
                        task=mode,
                        cell_ref=ref,
                        frames_in_cell=frames_val,
                        parsed_bbox=parsed_bb,
                        raw_note=note_txt,
                    ))

                if frames_val > 0:
                    sf = _parse_int(start_row[c_i]) if (start_row is not None and c_i < len(start_row)) else None
                    ef = _parse_int(end_row[c_i]) if (end_row is not None and c_i < len(end_row)) else None
                    target_blocks = rev_blocks if ann_key in FIRST_REVIEWER_KEYS else block_records
                    target_blocks.append(WorklogBlockRecord(
                        annotator_key=ann_key,
                        annotator_display=ann_disp,
                        sheet_name=sname,
                        date=dt,
                        task=mode,
                        job_id=job_id_str,
                        start_frame=sf,
                        end_frame=ef,
                        frames=frames_val,
                        bbox=bbox_val,
                        cell_ref=ref,
                        job_name=job_name_str,
                    ))

        if ann_key in FIRST_REVIEWER_KEYS:
            rev_daily.extend(agg.values())
        else:
            daily_records.extend(agg.values())

    wb.close()
    return daily_records, block_records, note_records, warnings, rev_daily, rev_blocks, rev_notes


def parse_rework(xlsx_path: str, progress_cb=None) -> Tuple[List[ReworkRecord], List[str]]:
    records: List[ReworkRecord] = []
    warnings: List[str] = []

    if progress_cb:
        progress_cb(62, "Đang đọc Notes trong file Inspection & Rework...")
    comments_by_sheet = load_xlsx_comments(xlsx_path)

    if progress_cb:
        progress_cb(68, "Đang mở file Inspection & Rework...")
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    sheet_names = wb.sheetnames
    total_sheets = max(len(sheet_names), 1)

    for s_idx, sname in enumerate(sheet_names):
        if progress_cb:
            pct = 68 + int((s_idx / total_sheets) * 30)
            progress_cb(pct, f"Đang đọc Review sheet: {sname} ({s_idx+1}/{total_sheets})")

        s_up = sname.strip().upper()
        if s_up.startswith("EXAMPLE") or s_up in SKIP_REWORK_SHEETS:
            continue

        ws = wb[sname]
        rows = list(ws.iter_rows(min_row=1, max_col=26, values_only=True))
        if len(rows) < 2:
            continue

        # Detect header columns dynamically (exact match first, then substring with exclusions)
        header = [re.sub(r"\s+", " ", str(v).strip().upper()) if v is not None else "" for v in rows[0]]

        def find_col(exact_names: List[str], contains_names: List[str], default_idx: int, exclude: List[str] = None) -> int:
            exclude = exclude or []
            for name in exact_names:
                for i, h in enumerate(header):
                    if h == name:
                        return i
            for pat in contains_names:
                for i, h in enumerate(header):
                    if pat in h and not any(ex in h for ex in exclude):
                        return i
            return default_idx

        col_date_a = find_col(["DATE"], ["DATE"], 0, exclude=["REWORK", "SECONDARY"])
        col_task = find_col(["TASK"], ["TASK"], 1)
        col_url = find_col(["JOB(URL)", "JOB (URL)"], ["JOB", "URL"], 2)
        col_tot_fr = find_col(["TOTAL FRAME"], ["TOTAL FRAME"], 3)
        col_start_fr = find_col(["WORKING FRAME START"], ["FRAME START"], 4)
        col_end_fr = find_col(["WORKING FRAME END"], ["FRAME END"], 5)
        col_insp = find_col(["INSPECTOR"], ["INSPECTOR"], 7)
        col_worker = find_col(["WORKER"], ["WORKER"], 8)
        col_img = find_col(["#IMAGE", "IMAGE"], ["#IMAGE", "IMAGE"], 14)
        col_work_time = find_col(["WORKING TIME"], ["WORKING TIME"], 15)
        col_err1 = find_col(["ISSUE FRAME"], ["ISSUE FRAME"], 16, exclude=["NUMBER"])
        col_err1_b = find_col(["ISSUE BBOX"], ["ISSUE BBOX"], 17)
        col_err2 = find_col(["PANGYO FEEDBACK FRAME"], ["FEEDBACK FRAME"], 18)
        col_err2_b = find_col(["PANGYO FEEDBACK BBOX"], ["FEEDBACK BBOX"], 19)
        col_date_v = find_col(["DATE OF SECONDARY REVIEW"], ["SECONDARY REVIEW"], 21, exclude=["REVIEWER"])

        c_map = comments_by_sheet.get(sname, {})
        data_rows = rows[1:]
        n_rows = len(data_rows)

        def cell(row, idx):
            return row[idx] if idx < len(row) else None

        def num(v) -> float:
            return float(v) if isinstance(v, (int, float)) else 0.0

        for i, r in enumerate(data_rows):
            # ONLY read the first (main) row of each block where WORKING FRAME START and END are present.
            # Detail sub-rows below the first row only list individual issue frame numbers and must NEVER be summed!
            sf_val = _parse_int(cell(r, col_start_fr))
            ef_val = _parse_int(cell(r, col_end_fr))
            if sf_val is None or ef_val is None:
                continue

            worker_raw = cell(r, col_worker)
            task_raw = cell(r, col_task)
            if not worker_raw:
                continue

            w_key = normalize_annotator_key(worker_raw)
            if w_key in INACTIVE_ANNOTATOR_KEYS:
                continue

            dt_a = to_date(cell(r, col_date_a))
            dt_v = to_date(cell(r, col_date_v))
            frames = num(cell(r, col_img))

            # Second row of the block holds BBOX in col O (#IMAGE)
            bbox = 0.0
            if i + 1 < n_rows:
                nr = data_rows[i + 1]
                if _parse_int(cell(nr, col_start_fr)) is None and _parse_int(cell(nr, col_end_fr)) is None:
                    bbox = num(cell(nr, col_img))

            # Collect any notes (and fallback TASK if missing on main row) attached to rows belonging to this block
            excel_row_num = i + 2
            block_notes = []
            j = i
            while j < n_rows:
                if j > i and (_parse_int(cell(data_rows[j], col_start_fr)) is not None or _parse_int(cell(data_rows[j], col_end_fr)) is not None):
                    break
                if j > i and not task_raw:
                    sub_task = cell(data_rows[j], col_task)
                    if sub_task and str(sub_task).strip():
                        task_raw = sub_task
                rnum = j + 2
                prefix_matches = [f"{ref}: {txt}" for ref, txt in c_map.items() if re.fullmatch(rf"[A-Z]+{rnum}", ref)]
                block_notes.extend(prefix_matches)
                j += 1

            url_str = str(cell(r, col_url) or "").strip()
            job_id_str = _extract_job_id(url_str)
            base_task = normalize_task(task_raw)
            insp_str = str(cell(r, col_insp) or "").strip()
            insp_key = normalize_annotator_key(insp_str)
            e1_fr = num(cell(r, col_err1))
            e1_bb = num(cell(r, col_err1_b))
            e2_fr = num(cell(r, col_err2))
            e2_bb = num(cell(r, col_err2_b))
            notes_joined = " | ".join(block_notes)

            # Worker record (also carries inspector + Col P WORKING TIME for the Reviewer tab)
            records.append(ReworkRecord(
                sheet_name=sname,
                row_num=excel_row_num,
                date_inspect=dt_a,
                date_work=dt_v,
                task=base_task,
                worker_key=w_key,
                worker_display=display_annotator_name(worker_raw),
                inspector=insp_str,
                job_url=url_str,
                frames=frames,
                bbox=bbox,
                err_first=e1_fr,
                err_first_bbox=e1_bb,
                err_second=e2_fr,
                err_second_bbox=e2_bb,
                work_time=num(cell(r, col_work_time)),
                notes=notes_joined,
                job_id=job_id_str,
                start_frame=sf_val,
                end_frame=ef_val,
            ))

    wb.close()
    return records, warnings


def _resolve_missing_rework_tasks(
    rework_records: List[ReworkRecord],
    worklog_blocks: List[WorklogBlockRecord],
) -> None:
    """Infer missing TASK on Rework records by matching Start/End Frame, Frame/BBOX count, and Job ID against Work_log blocks."""
    missing = [r for r in rework_records if not r.task]
    if not missing:
        return

    blocks_by_worker: Dict[str, List[WorklogBlockRecord]] = defaultdict(list)
    for b in worklog_blocks:
        if b.task:
            blocks_by_worker[b.annotator_key].append(b)

    task_by_job_id: Dict[str, Counter] = defaultdict(Counter)
    for r in rework_records:
        if r.task and r.job_id:
            task_by_job_id[r.job_id][r.task] += 1
    for b in worklog_blocks:
        if b.task and b.job_id:
            task_by_job_id[b.job_id][b.task] += 1

    for r in missing:
        ref_dt = r.date_work or r.date_inspect
        cands = blocks_by_worker.get(r.worker_key, [])
        best_task = ""
        best_score = None

        for b in cands:
            same_range = (
                r.start_frame is not None and r.end_frame is not None
                and b.start_frame == r.start_frame and b.end_frame == r.end_frame
            )
            same_fr = abs(b.frames - r.frames) < 1e-6 and r.frames > 0
            same_bb = abs(b.bbox - r.bbox) < 1e-6 and r.bbox > 0
            same_job = bool(r.job_id and b.job_id and r.job_id == b.job_id)
            day_gap = abs((ref_dt - b.date).days) if (ref_dt and b.date) else 999

            tier = None
            if same_range and same_fr and (same_bb or b.bbox == 0 or r.bbox == 0):
                tier = 0
            elif same_range:
                tier = 1
            elif same_job and same_fr:
                tier = 2
            elif same_fr and same_bb and day_gap <= 14:
                tier = 3

            if tier is not None:
                score = (tier, 0 if same_job else 1, 0 if same_bb else 1, 0 if same_fr else 1, day_gap)
                if best_score is None or score < best_score:
                    best_score = score
                    best_task = b.task

        if not best_task and r.job_id and r.job_id in task_by_job_id:
            best_task = task_by_job_id[r.job_id].most_common(1)[0][0]

        r.task = best_task or "UNKNOWN"


def _parse_dist_assignees(raw_val) -> Tuple[str, Tuple[str, ...]]:
    """
    Parse annotator name(s) from Distribution Sheet NAME column.
    Examples:
      'C.Linh (33050-)'              -> ('C.Linh (33050-)', ('CLINH',))
      'Nmai(32500-32744)'            -> ('Nmai(32500-32744)', ('NMAI',))
      'Nmai(0-100) / C.Linh(101-200)' -> (..., ('NMAI', 'CLINH'))
    """
    if raw_val is None:
        return "", ()
    raw_str = str(raw_val).strip()
    if not raw_str:
        return "", ()
    # Remove parenthetical frame hints like '(33050-)' or '(32500-32744)'
    cleaned = re.sub(r"\([^)]*\)", " ", raw_str)
    # Split multiple assignees separated by '/', ',', ';', '+', '&', newline
    parts = re.split(r"[/,;+&\n]+", cleaned)
    keys: List[str] = []
    for p in parts:
        # Remove standalone trailing numbers/ranges if any (keep digits glued to a name, e.g. 'Hien.03' -> HIEN03)
        p_txt = re.sub(r"(?<![A-Za-z0-9._])\d+\s*[-~–—]?\s*\d*\b", " ", p).strip()
        k = normalize_annotator_key(p_txt)
        if k and k not in keys:
            keys.append(k)
    return raw_str, tuple(keys)


def parse_distribution(xlsx_path: str, progress_cb=None) -> Tuple[List[DistributionJobRecord], List[str]]:
    """
    Parse distribution_sheet_<year>.xlsx across all sheets.
    Uses actual Excel merged_cells ranges so vertically merged TASK NAME, JOB URL,
    and Frame cells are resolved accurately even when intermediate NAME rows are blank.
    """
    records: List[DistributionJobRecord] = []
    warnings: List[str] = []

    if progress_cb:
        progress_cb(92, "Đang đọc Distribution Sheet (check ID Job & Name)...")

    wb = openpyxl.load_workbook(xlsx_path, read_only=False, data_only=True)
    for sname in wb.sheetnames:
        ws = wb[sname]
        max_r = ws.max_row or 0
        max_c = min(ws.max_column or 0, 20)
        if max_r < 2 or max_c < 2:
            continue

        # Build exact value grid with merged cells expanded
        grid: List[List[object]] = [
            [ws.cell(r, c).value for c in range(1, max_c + 1)]
            for r in range(1, max_r + 1)
        ]
        for rng in ws.merged_cells.ranges:
            if rng.min_col > max_c:
                continue
            top_val = grid[rng.min_row - 1][rng.min_col - 1]
            if top_val is None:
                continue
            end_r = min(rng.max_row, max_r)
            end_c = min(rng.max_col, max_c)
            for rr in range(rng.min_row - 1, end_r):
                for cc in range(rng.min_col - 1, end_c):
                    if grid[rr][cc] is None:
                        grid[rr][cc] = top_val

        # Dynamically locate header row containing JOB URL / URL and FRAME
        hdr_idx = None
        col_task = 0
        col_url = 1
        col_frame = 2
        col_name = 3
        col_insp: Optional[int] = None
        for r_i, r in enumerate(grid[:20]):
            norm = [re.sub(r"\s+", " ", str(v).strip().upper()) if v is not None else "" for v in r]
            url_i = next((i for i, h in enumerate(norm) if "JOB URL" in h or "JOB_URL" in h or h == "URL" or "JOB(URL)" in h), None)
            fr_i = next((i for i, h in enumerate(norm) if h == "FRAME" or h == "FRAMES" or "FRAME RANGE" in h), None)
            if url_i is not None and fr_i is not None:
                hdr_idx = r_i
                col_url = url_i
                col_frame = fr_i
                tk_i = next((i for i, h in enumerate(norm) if "TASK NAME" in h or "TASK_NAME" in h or h == "TASK" or "JOB NAME" in h), None)
                col_task = tk_i if tk_i is not None else max(0, url_i - 1)
                nm_i = next(
                    (
                        i for i, h in enumerate(norm)
                        if i not in (col_task, col_url, col_frame)
                        and (
                            h == "NAME"
                            or "ANNOTATOR" in h
                            or "WORKER" in h
                            or "ASSIGNEE" in h
                            or ("NAME" in h and "TASK" not in h and "JOB" not in h and "FILE" not in h)
                        )
                    ),
                    None,
                )
                col_name = nm_i if nm_i is not None else (col_frame + 1)
                ins_i = next(
                    (
                        i for i, h in enumerate(norm)
                        if i > col_name and ("INSP" in h or "INPECT" in h or "REVIEW" in h)
                    ),
                    None,
                )
                col_insp = ins_i
                break

        if hdr_idx is None:
            continue

        # Aggregate rows by (task_name_key, job_id, sf, ef) so all sub-row assignees merge cleanly
        seg_map: Dict[Tuple[str, str, int, int], DistributionJobRecord] = {}
        for idx, r in enumerate(grid[hdr_idx + 1:], start=hdr_idx + 2):
            raw_tk = r[col_task] if col_task < len(r) else None
            raw_url = r[col_url] if col_url < len(r) else None
            raw_fr = r[col_frame] if col_frame < len(r) else None
            raw_nm = r[col_name] if col_name < len(r) else None

            url_s = str(raw_url).strip() if raw_url is not None else ""
            fr_s = str(raw_fr).strip() if raw_fr is not None else ""
            if not url_s or not fr_s:
                continue

            jid = _extract_job_id(url_s)
            if not jid:
                m_tail = re.search(r"/(\d+)\s*$", url_s)
                if m_tail:
                    jid = m_tail.group(1)
            if not jid:
                continue

            m_fr = re.search(r"(\d+)\s*[-~–—]\s*(\d+)", fr_s)
            if not m_fr:
                continue
            sf, ef = int(m_fr.group(1)), int(m_fr.group(2))
            if sf > ef:
                sf, ef = ef, sf

            tk_s = str(raw_tk).strip() if raw_tk is not None else ""
            tk_key = normalize_job_name(tk_s)
            assignee_raw, assignee_keys = _parse_dist_assignees(raw_nm)
            raw_ins = r[col_insp] if (col_insp is not None and col_insp < len(r)) else None
            reviewer_raw, reviewer_keys = _parse_dist_assignees(raw_ins)

            seg_key = (tk_key, jid, sf, ef)
            if seg_key not in seg_map:
                rec = DistributionJobRecord(
                    sheet_name=sname,
                    row_num=idx,
                    task_name=tk_s,
                    task_name_key=tk_key,
                    job_url=url_s,
                    job_id=jid,
                    start_frame=sf,
                    end_frame=ef,
                    assignee_raw=assignee_raw,
                    assignee_keys=assignee_keys,
                    cvat_task_id=_extract_cvat_task_id(url_s),
                    reviewer_raw=reviewer_raw,
                    reviewer_keys=reviewer_keys,
                )
                seg_map[seg_key] = rec
                records.append(rec)
            else:
                rec = seg_map[seg_key]
                if assignee_keys:
                    merged_keys = list(rec.assignee_keys)
                    for k in assignee_keys:
                        if k not in merged_keys:
                            merged_keys.append(k)
                    rec.assignee_keys = tuple(merged_keys)
                if assignee_raw and assignee_raw not in rec.assignee_raw:
                    rec.assignee_raw = (
                        f"{rec.assignee_raw}; {assignee_raw}" if rec.assignee_raw else assignee_raw
                    )
                if reviewer_keys:
                    merged_rk = list(rec.reviewer_keys)
                    for k in reviewer_keys:
                        if k not in merged_rk:
                            merged_rk.append(k)
                    rec.reviewer_keys = tuple(merged_rk)
                if reviewer_raw and reviewer_raw not in rec.reviewer_raw:
                    rec.reviewer_raw = (
                        f"{rec.reviewer_raw}; {reviewer_raw}" if rec.reviewer_raw else reviewer_raw
                    )

    wb.close()
    return records, warnings


def load_dataset(
    worklog_path: str,
    rework_path: str,
    progress_cb=None,
    distribution_path: str = "",
) -> LoadedDataset:
    ds = LoadedDataset(
        worklog_path=worklog_path,
        rework_path=rework_path,
        distribution_path=distribution_path,
    )

    if worklog_path:
        wl_daily, wl_blocks, wl_notes, wl_warn, rev_daily, rev_blocks, rev_notes = parse_worklog(worklog_path, progress_cb=progress_cb)
        ds.worklog_daily = wl_daily
        ds.worklog_blocks = wl_blocks
        ds.worklog_notes = wl_notes
        ds.reviewer_worklog_daily = rev_daily
        ds.reviewer_worklog_blocks = rev_blocks
        ds.reviewer_worklog_notes = rev_notes
        ds.warnings.extend(wl_warn)

    if rework_path:
        rw_recs, rw_warn = parse_rework(rework_path, progress_cb=progress_cb)
        ds.rework_records = rw_recs
        ds.warnings.extend(rw_warn)

    if distribution_path and os.path.isfile(distribution_path):
        try:
            dist_jobs, dist_warn = parse_distribution(distribution_path, progress_cb=progress_cb)
            ds.distribution_jobs = dist_jobs
            ds.warnings.extend(dist_warn)
        except Exception as exc:
            ds.warnings.append(f"Không đọc được Distribution Sheet: {exc}")

    _resolve_missing_rework_tasks(ds.rework_records, ds.worklog_blocks)

    # Determine global max date and last active date per person to exclude inactive/old (~1 year) accounts
    all_dates: List[datetime.date] = []
    last_active_by_ann: Dict[str, datetime.date] = {}

    for r in ds.worklog_daily:
        if r.frames > 0 or r.work_time > 0 or r.bbox > 0:
            all_dates.append(r.date)
            prev = last_active_by_ann.get(r.annotator_key)
            if prev is None or r.date > prev:
                last_active_by_ann[r.annotator_key] = r.date

    for r in ds.rework_records:
        dt = r.date_work or r.date_inspect
        if dt and (r.frames > 0 or r.bbox > 0):
            all_dates.append(dt)
            prev = last_active_by_ann.get(r.worker_key)
            if prev is None or dt > prev:
                last_active_by_ann[r.worker_key] = dt

    if all_dates:
        ds.min_date = min(all_dates)
        ds.max_date = max(all_dates)

    # Filter out people who have no data or whose last activity was > 60 days before max_date (or ~1 year old)
    cutoff_date = (ds.max_date - datetime.timedelta(days=60)) if ds.max_date else datetime.date(2026, 9, 1)
    active_keys = {
        k for k, last_dt in last_active_by_ann.items()
        if k not in INACTIVE_ANNOTATOR_KEYS and last_dt >= cutoff_date
    }

    ds.worklog_daily = [r for r in ds.worklog_daily if r.annotator_key in active_keys]
    ds.worklog_blocks = [b for b in ds.worklog_blocks if b.annotator_key in active_keys]
    ds.worklog_notes = [n for n in ds.worklog_notes if n.annotator_key in active_keys]
    ds.rework_records = [r for r in ds.rework_records if r.worker_key in active_keys]

    ann_map: Dict[str, str] = {}
    task_set = set()
    for r in ds.worklog_daily:
        ann_map[r.annotator_key] = r.annotator_display
        if r.task and (r.frames > 0 or r.work_time > 0 or r.bbox > 0):
            task_set.add(r.task)
    for r in ds.rework_records:
        if r.worker_key not in ann_map:
            ann_map[r.worker_key] = r.worker_display
        if r.task and (r.frames > 0 or r.bbox > 0):
            task_set.add(r.task)

    ds.annotators = dict(sorted(ann_map.items(), key=lambda kv: kv[1]))
    ds.tasks = sorted(task_set)

    if progress_cb:
        progress_cb(100, "Hoàn tất tải dữ liệu!")
    return ds


@dataclass
class DailyDiffRecord:
    date: datetime.date                        # Work_log date (or Rework date if missing in WL)
    task: str
    annotator_key: str
    annotator: str
    wl_frames: float
    rw_frames: float
    wl_bbox: float
    rw_bbox: float
    wl_cells: str = ""
    rw_rows: str = ""
    reason: str = ""
    kind: str = "mismatch"                     # "mismatch" or "pending_rework"
    rw_date: Optional[datetime.date] = None    # Review/Inspect date in Rework (may differ from WL date)
    job_id: str = ""
    wl_range: str = ""                         # e.g. "5899~5999"
    rw_range: str = ""                         # e.g. "5900~5999"
    inspectors: Tuple[str, ...] = ()

    @property
    def diff_frame(self) -> float:
        return self.wl_frames - self.rw_frames

    @property
    def diff_bbox(self) -> float:
        return self.wl_bbox - self.rw_bbox

    @property
    def is_pending(self) -> bool:
        return self.kind == "pending_rework"

    @property
    def is_missing_bbox(self) -> bool:
        return self.is_pending and (
            (self.wl_bbox > 0 and self.rw_bbox <= 0 and self.rw_frames > 0)
            or (self.rw_bbox > 0 and self.wl_bbox <= 0 and self.wl_frames > 0)
        )

    @property
    def is_job_id_mismatch(self) -> bool:
        return (not self.is_pending) and ("Sai ID Job" in self.reason or "Lệch ID Job" in self.reason)

    @property
    def is_dist_frame_mismatch(self) -> bool:
        return (not self.is_pending) and ("Sai Frame " in self.reason)

    @property
    def frame_range_display(self) -> str:
        if self.wl_range and self.rw_range:
            if self.wl_range == self.rw_range:
                return self.wl_range
            return f"Work Log:{self.wl_range} ≠ Review:{self.rw_range}"
        return self.wl_range or self.rw_range or "—"


@dataclass
class AggregatedRow:
    task: str
    annotator_key: str
    annotator: str
    work_load: float
    work_bbox: float
    work_time: float
    err_first: float
    err_second: float
    err_first_bbox: float = 0.0
    err_second_bbox: float = 0.0
    # Audit fields (always computed from BOTH files so we can highlight mismatches)
    wl_frames: float = 0.0
    wl_bbox: float = 0.0
    rw_frames: float = 0.0
    rw_bbox: float = 0.0
    note_count: int = 0
    daily_diffs: List[DailyDiffRecord] = field(default_factory=list)

    @property
    def is_reviewer(self) -> bool:
        return self.annotator_key in FIRST_REVIEWER_KEYS or self.task.endswith("_REW")

    @property
    def role_label(self) -> str:
        return "Inspector Lần 1" if self.is_reviewer else "Worker"

    @property
    def frame_per_hour(self) -> float:
        return self.work_load / self.work_time if self.work_time > 0 else 0.0

    @property
    def bbox_per_hour(self) -> float:
        return self.work_bbox / self.work_time if self.work_time > 0 else 0.0

    @property
    def total_err_frame(self) -> float:
        """
        Reviewer Lần 1: PANGYO FEEDBACK FRAME only (err_second).
        Worker: 1st Review ISSUE FRAME + PANGYO FEEDBACK FRAME (err_first + err_second).
        """
        return self.err_second if self.is_reviewer else (self.err_first + self.err_second)

    @property
    def total_err_bbox(self) -> float:
        """
        Reviewer Lần 1: PANGYO FEEDBACK BBOX only (err_second_bbox).
        Worker: 1st Review ISSUE BBOX + PANGYO FEEDBACK BBOX (err_first_bbox + err_second_bbox).
        """
        return self.err_second_bbox if self.is_reviewer else (self.err_first_bbox + self.err_second_bbox)

    @property
    def first_err_rate(self) -> float:
        return (self.err_first / self.work_load * 100.0) if self.work_load > 0 else 0.0

    @property
    def second_err_rate(self) -> float:
        return (self.err_second / self.work_load * 100.0) if self.work_load > 0 else 0.0

    @property
    def total_err_rate(self) -> float:
        return (self.total_err_frame / self.work_load * 100.0) if self.work_load > 0 else 0.0

    @property
    def total_err_bbox_rate(self) -> float:
        return (self.total_err_bbox / self.work_bbox * 100.0) if self.work_bbox > 0 else 0.0

    @property
    def diff_frame(self) -> float:
        return self.wl_frames - self.rw_frames

    @property
    def diff_bbox(self) -> float:
        return self.wl_bbox - self.rw_bbox

    @property
    def real_diffs(self) -> List[DailyDiffRecord]:
        return [d for d in self.daily_diffs if d.kind == "mismatch"]

    @property
    def pending_diffs(self) -> List[DailyDiffRecord]:
        return [d for d in self.daily_diffs if d.kind == "pending_rework"]

    @property
    def mismatch_dates(self) -> List[datetime.date]:
        return sorted({d.date for d in self.real_diffs})

    @property
    def pending_dates(self) -> List[datetime.date]:
        return sorted({d.date for d in self.pending_diffs})

    @property
    def mismatch_dates_str(self) -> str:
        dates = self.mismatch_dates
        return ", ".join(d.strftime("%m/%d") for d in dates) if dates else ""

    @property
    def pending_dates_str(self) -> str:
        dates = self.pending_dates
        return ", ".join(d.strftime("%m/%d") for d in dates) if dates else ""

    @property
    def status_dates_summary(self) -> str:
        parts = []
        if self.mismatch_dates:
            parts.append(f"Không khớp: {self.mismatch_dates_str}")
        no_rw_dates = sorted({d.date for d in self.pending_diffs if not d.is_missing_bbox})
        miss_bb_dates = sorted({d.date for d in self.pending_diffs if d.is_missing_bbox})
        if no_rw_dates:
            parts.append("Chưa có Review: " + ", ".join(d.strftime("%m/%d") for d in no_rw_dates))
        if miss_bb_dates:
            parts.append("Chưa nhập BBOX: " + ", ".join(d.strftime("%m/%d") for d in miss_bb_dates))
        return " | ".join(parts)

    @property
    def has_mismatch(self) -> bool:
        """True ONLY when both files have data (or Rework has extra data) and values disagree."""
        return len(self.real_diffs) > 0

    @property
    def has_pending(self) -> bool:
        """True when Work_log has data on one or more days not yet entered in Rework."""
        return len(self.pending_diffs) > 0

    @property
    def status_kind(self) -> str:
        if self.has_mismatch:
            return "mismatch"
        if self.has_pending:
            return "pending"
        return "ok"


@dataclass
class TaskSummary:
    task: str
    rows: List[AggregatedRow]
    total_work_load: float
    total_work_bbox: float
    total_work_time: float
    total_err_first: float
    total_err_second: float
    total_err_first_bbox: float = 0.0
    total_err_second_bbox: float = 0.0

    @property
    def total_err_frame(self) -> float:
        return sum(r.total_err_frame for r in self.rows)

    @property
    def total_err_bbox(self) -> float:
        return sum(r.total_err_bbox for r in self.rows)

    @property
    def avg_frame_per_hour(self) -> float:
        return self.total_work_load / self.total_work_time if self.total_work_time > 0 else 0.0

    @property
    def avg_bbox_per_hour(self) -> float:
        return self.total_work_bbox / self.total_work_time if self.total_work_time > 0 else 0.0

    @property
    def total_task_err_rate(self) -> float:
        return (self.total_err_frame / self.total_work_load * 100.0) if self.total_work_load > 0 else 0.0

    @property
    def total_task_err_bbox_rate(self) -> float:
        return (self.total_err_bbox / self.total_work_bbox * 100.0) if self.total_work_bbox > 0 else 0.0


@dataclass
class InspectorWorkerRow:
    inspector_key: str
    inspector: str
    annotator_key: str
    annotator: str
    work_load: float
    work_bbox: float
    work_time: float
    err_first: float
    err_second: float
    err_first_bbox: float = 0.0
    err_second_bbox: float = 0.0


@dataclass
class InspectorSummary:
    inspector_key: str
    inspector: str
    rows: List[InspectorWorkerRow]
    total_work_load: float
    total_work_bbox: float
    total_work_time: float
    total_err_first: float
    total_err_second: float
    total_err_first_bbox: float = 0.0
    total_err_second_bbox: float = 0.0

    @property
    def frame_per_hour(self) -> float:
        return self.total_work_load / self.total_work_time if self.total_work_time > 0 else 0.0

    @property
    def bbox_per_hour(self) -> float:
        return self.total_work_bbox / self.total_work_time if self.total_work_time > 0 else 0.0

    @property
    def err_rate_frame(self) -> float:
        """Reviewer Error Rate (By Frame): PANGYO FEEDBACK FRAME / Work Load * 100."""
        return (self.total_err_second / self.total_work_load * 100.0) if self.total_work_load > 0 else 0.0

    @property
    def err_rate_bbox(self) -> float:
        """Reviewer Error Rate (By BBOX): PANGYO FEEDBACK BBOX / Work BBOX * 100."""
        return (self.total_err_second_bbox / self.total_work_bbox * 100.0) if self.total_work_bbox > 0 else 0.0

    @property
    def first_inspector_recall(self) -> float:
        """First Inspector Recall (By Frame): First Insp Err / (First Insp Err + Second Insp Err) * 100."""
        denom = self.total_err_first + self.total_err_second
        return (self.total_err_first / denom * 100.0) if denom > 0 else 100.0


@dataclass
class ReviewerDiffRecord:
    date: datetime.date
    rw_date: Optional[datetime.date]
    inspector_key: str
    inspector: str
    annotator: str
    task: str
    job_id: str = ""
    wl_range: str = ""
    rw_range: str = ""
    ann_range: str = ""
    wl_frames: float = 0.0
    rw_frames: float = 0.0
    ann_frames: float = 0.0
    wl_bbox: float = 0.0
    rw_bbox: float = 0.0
    ann_bbox: float = 0.0
    wl_cells: str = ""
    rw_rows: str = ""
    ann_cells: str = ""
    reason: str = ""
    kind: str = "mismatch"  # "mismatch" or "pending_rework"

    @property
    def is_pending(self) -> bool:
        return self.kind == "pending_rework"

    @property
    def frame_range_display(self) -> str:
        parts = []
        if self.wl_range:
            parts.append(f"WL Inspector: {self.wl_range}")
        if self.rw_range and self.rw_range != self.wl_range:
            parts.append(f"RW:{self.rw_range}")
        if self.ann_range and self.ann_range != self.wl_range:
            parts.append(f"Ann:{self.ann_range}")
        return " | ".join(parts) if parts else "—"


@dataclass
class ReviewerAuditSummaryRow:
    task: str
    inspector_key: str
    inspector: str
    wl_frames: float = 0.0
    rw_frames: float = 0.0
    wl_bbox: float = 0.0
    rw_bbox: float = 0.0
    wl_time: float = 0.0
    rw_time: float = 0.0
    ann_frames: float = 0.0
    ann_bbox: float = 0.0
    note_count: int = 0
    daily_diffs: List[ReviewerDiffRecord] = field(default_factory=list)

    @property
    def diff_frame(self) -> float:
        return self.wl_frames - self.rw_frames

    @property
    def diff_bbox(self) -> float:
        return self.wl_bbox - self.rw_bbox

    @property
    def diff_time(self) -> float:
        return self.wl_time - self.rw_time

    @property
    def real_diffs(self) -> List[ReviewerDiffRecord]:
        return [d for d in self.daily_diffs if d.kind == "mismatch"]

    @property
    def pending_diffs(self) -> List[ReviewerDiffRecord]:
        return [d for d in self.daily_diffs if d.kind == "pending_rework"]

    @property
    def status_kind(self) -> str:
        if self.real_diffs:
            return "mismatch"
        if self.pending_diffs:
            return "pending"
        return "ok"

    @property
    def status_dates_summary(self) -> str:
        parts = []
        m_dates = sorted({d.date for d in self.real_diffs})
        p_dates = sorted({d.date for d in self.pending_diffs})
        if m_dates:
            parts.append("Không khớp: " + ", ".join(d.strftime("%m/%d") for d in m_dates))
        if p_dates:
            parts.append("Chưa có Review: " + ", ".join(d.strftime("%m/%d") for d in p_dates))
        return " | ".join(parts)


@dataclass
class ReportResult:
    start_date: datetime.date
    end_date: datetime.date
    source_mode: str                  # "rework" or "worklog"
    rework_date_mode: str             # "col_v" or "col_a"
    task_groups: List[TaskSummary]    # Template 1 (Grouped by TASK)
    annotator_rows: List[AggregatedRow]  # Template 2 (Sorted by Annotator A-Z)
    all_rows: List[AggregatedRow]     # Raw (Annotator, Task) rows for Audit tab
    filtered_notes: List[WorklogNoteRecord]
    daily_diffs: List[DailyDiffRecord] = field(default_factory=list)
    inspector_groups: List[InspectorSummary] = field(default_factory=list)
    reviewer_audit_rows: List[ReviewerAuditSummaryRow] = field(default_factory=list)
    reviewer_daily_diffs: List[ReviewerDiffRecord] = field(default_factory=list)
    reviewer_notes: List[WorklogNoteRecord] = field(default_factory=list)


def _classify_daily_diff(wl_fr: float, rw_fr: float, wl_bb: float, rw_bb: float) -> Tuple[str, str]:
    """
    Return (kind, reason) where kind is 'pending_rework', 'mismatch', or 'ok'.
    BBOX is only considered mismatched when BOTH Work Log and Review have BBOX data (> 0)
    and the values disagree. If either side has no BBOX entered (0 / blank), it is NOT a
    BBOX error.
    """
    has_wl = (wl_fr > 0 or wl_bb > 0)
    has_rw = (rw_fr > 0 or rw_bb > 0)
    if has_wl and not has_rw:
        return "pending_rework", "Chưa có data bên Review"
    if has_rw and not has_wl:
        return "mismatch", "Có ở Review, thiếu bên Work Log"

    diff_f = abs(wl_fr - rw_fr) > 1e-6
    has_wl_bb = wl_bb > 0
    has_rw_bb = rw_bb > 0
    diff_b = has_wl_bb and has_rw_bb and (abs(wl_bb - rw_bb) > 1e-6)
    missing_rw_bb = has_wl_bb and not has_rw_bb
    missing_wl_bb = has_rw_bb and not has_wl_bb

    if diff_f and diff_b:
        return "mismatch", "Lệch cả Frame & BBOX"
    if diff_f:
        if missing_rw_bb:
            return "mismatch", "Lệch Frame (chưa nhập BBOX bên Review)"
        if missing_wl_bb:
            return "mismatch", "Lệch Frame (chưa nhập BBOX bên Work Log)"
        return "mismatch", "Khớp BBOX, lệch Frame"
    if diff_b:
        return "mismatch", "Khớp Frame, lệch BBOX (kiểm tra Note)"
    if missing_rw_bb:
        return "pending_rework", "Chưa nhập BBOX bên Review"
    if missing_wl_bb:
        return "pending_rework", "Chưa nhập BBOX bên Work Log"
    return "ok", ""


def _format_rw_locs(recs: List[ReworkRecord]) -> str:
    by_sheet = defaultdict(list)
    for r in recs:
        short_s = r.sheet_name.split("(")[0].strip() or r.sheet_name
        by_sheet[short_s].append(r.row_num)
    parts = []
    for s, nums in by_sheet.items():
        rnums = ",".join(f"R{n}" for n in sorted(set(nums))[:5])
        if len(nums) > 5:
            rnums += "..."
        parts.append(f"{s}!{rnums}")
    return "; ".join(parts)


def _match_blocks_with_frame_ranges(
    ds: LoadedDataset,
    start_date: datetime.date,
    end_date: datetime.date,
    selected_annotators: Optional[set],
) -> Tuple[Dict[Tuple[str, str], List[DailyDiffRecord]], Dict[Tuple[str, str], List[float]], List[ReworkRecord]]:
    """
    Match Work_log blocks in [start_date, end_date] against ALL Review records by
    (Annotator, Task, Job ID, Start Frame ~ End Frame) + Frames/BBOX without filtering
    by any date column in Inspection_And_Rework.
    Returns (diffs_by_ann_task, rw_agg_all, matched_rw_records).
    """
    diffs_by_ann_task: Dict[Tuple[str, str], List[DailyDiffRecord]] = defaultdict(list)
    rw_agg_all: Dict[Tuple[str, str], List[float]] = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0, 0.0, 0.0])
    matched_rw_ids: set = set()
    matched_rw_records: List[ReworkRecord] = []

    def _consume_rw(key: Tuple[str, str], recs: List[ReworkRecord]):
        for rec in recs:
            if id(rec) not in matched_rw_ids:
                matched_rw_ids.add(id(rec))
                matched_rw_records.append(rec)
                rw_agg_all[key][0] += rec.frames
                rw_agg_all[key][1] += rec.bbox
                rw_agg_all[key][2] += rec.err_first
                rw_agg_all[key][3] += rec.err_second
                rw_agg_all[key][4] += rec.err_first_bbox
                rw_agg_all[key][5] += rec.err_second_bbox

    # Index Distribution Sheet jobs by Job ID (merging multiple frame segments under the same Job ID),
    # by CVAT Task ID (/tasks/<id>), and by normalized Task Name
    dist_by_id: Dict[str, DistributionJobRecord] = {}
    dist_by_cvat_task: Dict[str, List[DistributionJobRecord]] = defaultdict(list)
    dist_by_task: Dict[str, List[DistributionJobRecord]] = defaultdict(list)
    for dj in ds.distribution_jobs:
        if dj.job_id in dist_by_id:
            prev = dist_by_id[dj.job_id]
            merged_keys = list(prev.assignee_keys)
            for k in dj.assignee_keys:
                if k not in merged_keys:
                    merged_keys.append(k)
            dist_by_id[dj.job_id] = DistributionJobRecord(
                sheet_name=prev.sheet_name,
                row_num=prev.row_num,
                task_name=prev.task_name or dj.task_name,
                task_name_key=prev.task_name_key or dj.task_name_key,
                job_url=prev.job_url,
                job_id=prev.job_id,
                start_frame=min(prev.start_frame, dj.start_frame),
                end_frame=max(prev.end_frame, dj.end_frame),
                assignee_raw=prev.assignee_raw,
                assignee_keys=tuple(merged_keys),
                cvat_task_id=prev.cvat_task_id or dj.cvat_task_id,
            )
        else:
            dist_by_id[dj.job_id] = dj
        if dj.cvat_task_id:
            dist_by_cvat_task[dj.cvat_task_id].append(dj)
        if dj.task_name_key:
            dist_by_task[dj.task_name_key].append(dj)

    def _same_task_jobs_for(recorded_id: str, fallback_tn_key: str) -> List[DistributionJobRecord]:
        """Return ONLY Distribution jobs belonging to the EXACT SAME Task (never cross-task!)."""
        rec_dist = dist_by_id.get(recorded_id) if recorded_id else None
        if rec_dist is not None:
            if rec_dist.cvat_task_id and rec_dist.cvat_task_id in dist_by_cvat_task:
                return dist_by_cvat_task[rec_dist.cvat_task_id]
            if rec_dist.task_name_key and rec_dist.task_name_key in dist_by_task:
                return dist_by_task[rec_dist.task_name_key]
        if fallback_tn_key and fallback_tn_key in dist_by_task:
            return dist_by_task[fallback_tn_key]
        return []

    def _find_frame_dist(
        task_jobs: List[DistributionJobRecord],
        recorded_id: str,
        annotator_key: str,
        sf: int,
        ef: int,
    ) -> Optional[DistributionJobRecord]:
        if not task_jobs:
            return None
        hits = [d for d in task_jobs if d.start_frame <= sf and ef <= d.end_frame]
        if not hits:
            hits = [d for d in task_jobs if d.start_frame <= sf <= d.end_frame]
        if not hits:
            return None
        return (
            next((d for d in hits if d.job_id == recorded_id), None)
            or next((d for d in hits if annotator_key and annotator_key in d.assignee_keys), None)
            or hits[0]
        )

    def _resolve_dist_task_key(wb: WorklogBlockRecord, matched_rws: List[ReworkRecord]) -> str:
        if not dist_by_task:
            return ""
        raw_k = normalize_job_name(wb.job_name)
        if raw_k and raw_k in dist_by_task:
            return raw_k
        # Try stripping leading mode prefix (e.g. "OD_HOPJT_#2" -> "HOPJT_#2") or unique suffix match
        if raw_k:
            for pfx in ("OD_", "PFE_", "SSEG_", "ANPR_"):
                if raw_k.startswith(pfx) and raw_k[len(pfx):] in dist_by_task:
                    return raw_k[len(pfx):]
            sub_hits = [k for k in dist_by_task if k in raw_k or raw_k in k]
            if len(sub_hits) == 1:
                return sub_hits[0]
        # Fallback: infer Task Name from recorded Job ID in Work Log or Review
        if wb.job_id and wb.job_id in dist_by_id:
            return dist_by_id[wb.job_id].task_name_key
        for rw in matched_rws:
            if rw.job_id and rw.job_id in dist_by_id:
                return dist_by_id[rw.job_id].task_name_key
        return ""

    def _diagnose_dist_mismatch(
        side_label: str,
        recorded_id: str,
        annotator_key: str,
        sf: int,
        ef: int,
        frame_dist: Optional[DistributionJobRecord],
    ) -> str:
        rec_dist = dist_by_id.get(recorded_id) if recorded_id else None

        # If [sf, ef] is already inside the recorded Job ID's own Distribution frame range -> 100% VALID!
        if rec_dist is not None and rec_dist.start_frame <= sf and ef <= rec_dist.end_frame:
            return ""

        # Case 1: Both recorded ID and frame range resolve to Distribution jobs (guaranteed SAME task!)
        if rec_dist is not None and frame_dist is not None:
            if rec_dist.job_id == frame_dist.job_id:
                return ""
            rec_is_owner = bool(annotator_key and annotator_key in rec_dist.assignee_keys)
            frm_is_owner = bool(annotator_key and annotator_key in frame_dist.assignee_keys)
            if rec_is_owner and not frm_is_owner:
                # Recorded Job ID IS assigned to this annotator -> ID is RIGHT, Frame is WRONG!
                return (
                    f"Sai Frame {side_label}: ID {rec_dist.job_id} ({rec_dist.start_frame}~{rec_dist.end_frame}) đúng, "
                    f"ghi nhầm sang frame Job {frame_dist.job_id} ({frame_dist.start_frame}~{frame_dist.end_frame}) [{rec_dist.task_name}]"
                )
            if frm_is_owner and not rec_is_owner:
                # Frame's Job IS assigned to this annotator -> Frame is RIGHT, Job ID is WRONG!
                return (
                    f"Sai ID Job {side_label}: ghi {rec_dist.job_id} ({rec_dist.start_frame}~{rec_dist.end_frame}), "
                    f"đúng là {frame_dist.job_id} ({frame_dist.start_frame}~{frame_dist.end_frame}) [{frame_dist.task_name}]"
                )
            if rec_is_owner and frm_is_owner:
                # Same annotator owns both jobs -> flag clearly that ID and Frame disagree
                return (
                    f"Lệch ID Job / Frame {side_label}: ghi ID {rec_dist.job_id} ({rec_dist.start_frame}~{rec_dist.end_frame}), "
                    f"frame thuộc Job {frame_dist.job_id} ({frame_dist.start_frame}~{frame_dist.end_frame}) [{frame_dist.task_name}]"
                )
            # Neither has assignee match -> fallback to Job ID mismatch warning
            return (
                f"Sai ID Job {side_label}: ghi {rec_dist.job_id} ({rec_dist.start_frame}~{rec_dist.end_frame}), "
                f"đúng là {frame_dist.job_id} ({frame_dist.start_frame}~{frame_dist.end_frame}) [{frame_dist.task_name}]"
            )

        # Case 2: Recorded ID is valid in Distribution, but frame [sf, ef] falls outside ALL jobs in that task
        if rec_dist is not None and frame_dist is None:
            return (
                f"Sai Frame {side_label}: ID {rec_dist.job_id} có frame {rec_dist.start_frame}~{rec_dist.end_frame}, "
                f"đang ghi {sf}~{ef} [{rec_dist.task_name}]"
            )

        # Case 3: Recorded ID missing/unknown, but frame matches a known Distribution job in the same task
        if rec_dist is None and frame_dist is not None:
            return (
                f"Sai ID Job {side_label}: ghi {recorded_id or '(trống)'}, "
                f"đúng là {frame_dist.job_id} ({frame_dist.start_frame}~{frame_dist.end_frame}) [{frame_dist.task_name}]"
            )
        return ""

    def _check_dist_job_ids(wb: WorklogBlockRecord, matched_rws: List[ReworkRecord]) -> str:
        if not ds.distribution_jobs or wb.start_frame is None or wb.end_frame is None:
            return ""
        fallback_tn = _resolve_dist_task_key(wb, matched_rws)

        sf, ef = wb.start_frame, wb.end_frame
        if sf > ef:
            sf, ef = ef, sf

        issues: List[str] = []
        wl_task_jobs = _same_task_jobs_for(wb.job_id, fallback_tn)
        if wl_task_jobs:
            wl_frame_dist = _find_frame_dist(wl_task_jobs, wb.job_id, wb.annotator_key, sf, ef)
            wl_issue = _diagnose_dist_mismatch("Work Log", wb.job_id, wb.annotator_key, sf, ef, wl_frame_dist)
            if wl_issue:
                issues.append(wl_issue)

        seen_rw_jids = []
        for rw in matched_rws:
            if rw.job_id and rw.job_id not in seen_rw_jids:
                seen_rw_jids.append(rw.job_id)
        for rw_jid in seen_rw_jids:
            rw_sf = matched_rws[0].start_frame if matched_rws and matched_rws[0].start_frame is not None else sf
            rw_ef = matched_rws[0].end_frame if matched_rws and matched_rws[0].end_frame is not None else ef
            if rw_sf > rw_ef:
                rw_sf, rw_ef = rw_ef, rw_sf
            rw_task_jobs = _same_task_jobs_for(rw_jid, fallback_tn)
            if rw_task_jobs:
                rw_frame_dist = _find_frame_dist(rw_task_jobs, rw_jid, wb.annotator_key, rw_sf, rw_ef)
                rw_issue = _diagnose_dist_mismatch("Review", rw_jid, wb.annotator_key, rw_sf, rw_ef, rw_frame_dist)
                if rw_issue:
                    issues.append(rw_issue)

        return " | ".join(issues)

    # Index ALL Rework records by (worker_key, task) across ALL dates
    rw_by_wt: Dict[Tuple[str, str], List[ReworkRecord]] = defaultdict(list)
    for r in ds.rework_records:
        if selected_annotators is not None and r.worker_key not in selected_annotators:
            continue
        if r.frames > 0 or r.bbox > 0 or (r.start_frame is not None and r.end_frame is not None):
            rw_by_wt[(r.worker_key, r.task)].append(r)

    # Index ALL Work_log blocks by (annotator_key, task) (for cross-date reservation & N-to-1 merge checks)
    wl_all_by_wt: Dict[Tuple[str, str], List[WorklogBlockRecord]] = defaultdict(list)
    wl_in_range: List[WorklogBlockRecord] = []
    wl_outside: List[WorklogBlockRecord] = []
    for wb in ds.worklog_blocks:
        if selected_annotators is not None and wb.annotator_key not in selected_annotators:
            continue
        wl_all_by_wt[(wb.annotator_key, wb.task)].append(wb)
        if start_date <= wb.date <= end_date:
            wl_in_range.append(wb)
        else:
            wl_outside.append(wb)

    def _exact_match_score(wb: WorklogBlockRecord, c: ReworkRecord, reserved_ids: set) -> Tuple[int, int, int, int, int]:
        """Rank exact-range Review candidates strictly by content (unused, Job ID, Frames/BBOX from Note); never require matching dates."""
        if id(c) in matched_rw_ids:
            used_pen = 2
        elif id(c) in reserved_ids:
            used_pen = 1
        else:
            used_pen = 0
        job_pen = 0 if (wb.job_id and c.job_id and c.job_id == wb.job_id) else (1 if (not wb.job_id or not c.job_id) else 2)
        same_fr = abs(c.frames - wb.frames) <= 1e-6
        same_bb = (wb.bbox > 0 and c.bbox > 0 and abs(c.bbox - wb.bbox) <= 1e-6)
        if same_fr and (same_bb or wb.bbox == 0 or c.bbox == 0):
            val_pen = 0 if same_bb else 1
        elif same_fr:
            val_pen = 2
        else:
            val_pen = 3
        # Final tie-breaker only for completely identical duplicates: prefer review on/after worklog date
        d_ref = c.date_inspect or c.date_work
        order_pen = 0 if (d_ref is not None and d_ref >= wb.date) else 1
        return (used_pen, job_pen, val_pen, order_pen, c.row_num)

    # Pass 0: Pre-reserve Review records that ALREADY exactly match Work_log blocks OUTSIDE [start_date, end_date].
    # Prevents a block in the current period (e.g. 11/25 184~190) from stealing a Review row written on 11/25
    # that actually reviewed an earlier day's Work_log block (e.g. 11/21 177~183)!
    reserved_rw_ids: set = set()
    for wb_out in wl_outside:
        if wb_out.start_frame is None or wb_out.end_frame is None:
            continue
        wt_key = (wb_out.annotator_key, wb_out.task)
        cands = rw_by_wt.get(wt_key, [])
        exact = [
            c for c in cands
            if id(c) not in reserved_rw_ids
            and c.start_frame == wb_out.start_frame and c.end_frame == wb_out.end_frame
            and (not wb_out.job_id or not c.job_id or c.job_id == wb_out.job_id)
        ]
        if exact:
            best = min(exact, key=lambda c: _exact_match_score(wb_out, c, reserved_rw_ids))
            reserved_rw_ids.add(id(best))
            continue
        # Also reserve exact 1-to-N splits belonging strictly to wb_out
        if wb_out.job_id:
            splits = [
                c for c in cands
                if id(c) not in reserved_rw_ids
                and c.job_id == wb_out.job_id
                and c.start_frame is not None and c.end_frame is not None
                and c.start_frame >= wb_out.start_frame and c.end_frame <= wb_out.end_frame
            ]
            if splits:
                s_min = min(c.start_frame for c in splits)
                s_max = max(c.end_frame for c in splits)
                if s_min == wb_out.start_frame and s_max == wb_out.end_frame:
                    for c in splits:
                        reserved_rw_ids.add(id(c))

    resolved_wb_indices: set = set()

    # Pass 1: Exact Start/End Frame matches across ALL in-range Work_log blocks first
    # (so an unreviewed block earlier in the list can NEVER steal another block's exact match).
    for idx_wb, wb in enumerate(wl_in_range):
        if wb.start_frame is None or wb.end_frame is None:
            continue
        wt_key = (wb.annotator_key, wb.task)
        cands = rw_by_wt.get(wt_key, [])
        exact = [
            c for c in cands
            if c.start_frame == wb.start_frame and c.end_frame == wb.end_frame
            and (not wb.job_id or not c.job_id or c.job_id == wb.job_id)
        ]
        if not exact:
            exact = [c for c in cands if c.start_frame == wb.start_frame and c.end_frame == wb.end_frame]
        if not exact:
            continue

        matched_single = min(exact, key=lambda c: _exact_match_score(wb, c, reserved_rw_ids))
        resolved_wb_indices.add(idx_wb)
        _consume_rw(wt_key, [matched_single])

        wl_rng = _fmt_range(wb.start_frame, wb.end_frame)
        wl_loc = f"{wb.sheet_name}!{wb.cell_ref}"
        job_err = _check_dist_job_ids(wb, [matched_single])
        exact_vals = abs(matched_single.frames - wb.frames) <= 1e-6 and abs(matched_single.bbox - wb.bbox) <= 1e-6
        if exact_vals and not job_err:
            continue  # 100% exact match regardless of review date!
        if exact_vals:
            diff_kind, diff_reason = "mismatch", job_err
        else:
            diff_kind, diff_reason = _classify_daily_diff(wb.frames, matched_single.frames, wb.bbox, matched_single.bbox)
            if job_err:
                diff_kind = "mismatch"
                diff_reason = f"{diff_reason} | {job_err}"
        diffs_by_ann_task[wt_key].append(DailyDiffRecord(
            date=wb.date,
            rw_date=matched_single.date_inspect or matched_single.date_work,
            task=wb.task,
            annotator_key=wb.annotator_key,
            annotator=wb.annotator_display,
            wl_frames=wb.frames,
            rw_frames=matched_single.frames,
            wl_bbox=wb.bbox,
            rw_bbox=matched_single.bbox,
            wl_cells=wl_loc,
            rw_rows=_format_rw_locs([matched_single]),
            inspectors=unique_inspector_names([matched_single.inspector]),
            reason=diff_reason,
            kind=diff_kind,
            job_id=wb.job_id or matched_single.job_id,
            wl_range=wl_rng,
            rw_range=_fmt_range(matched_single.start_frame, matched_single.end_frame),
        ))

    # Pass 2: Overlapping Start/End Frame matches for remaining unresolved Work_log blocks (NO date cutoff)
    for idx_wb, wb in enumerate(wl_in_range):
        if idx_wb in resolved_wb_indices or wb.start_frame is None or wb.end_frame is None:
            continue
        wt_key = (wb.annotator_key, wb.task)
        cands = rw_by_wt.get(wt_key, [])
        overlaps: List[ReworkRecord] = []
        for c in cands:
            if c.start_frame is None or c.end_frame is None:
                continue
            if max(wb.start_frame, c.start_frame) > min(wb.end_frame, c.end_frame):
                continue
            if wb.job_id and c.job_id and wb.job_id != c.job_id:
                continue
            # Skip Review rows already claimed as an exact match by another block,
            # unless this Review row strictly encloses wb (N-to-1 merge across multiple WL blocks).
            is_enclosing_merge = (c.start_frame <= wb.start_frame and c.end_frame >= wb.end_frame and (c.start_frame < wb.start_frame or c.end_frame > wb.end_frame))
            if (id(c) in matched_rw_ids or id(c) in reserved_rw_ids) and not is_enclosing_merge:
                continue
            overlaps.append(c)

        if not overlaps:
            continue

        resolved_wb_indices.add(idx_wb)
        _consume_rw(wt_key, overlaps)
        wl_rng = _fmt_range(wb.start_frame, wb.end_frame)
        wl_loc = f"{wb.sheet_name}!{wb.cell_ref}"
        sum_fr = sum(c.frames for c in overlaps)
        sum_bb = sum(c.bbox for c in overlaps)
        min_sf = min(c.start_frame for c in overlaps if c.start_frame is not None)
        max_ef = max(c.end_frame for c in overlaps if c.end_frame is not None)
        job_err = _check_dist_job_ids(wb, overlaps)

        # Check 1-to-N split exact match (1 WL block split into multiple RW rows)
        is_exact_overlap = (
            min_sf == wb.start_frame and max_ef == wb.end_frame
            and abs(sum_fr - wb.frames) <= 1e-6 and abs(sum_bb - wb.bbox) <= 1e-6
        )
        if is_exact_overlap and not job_err:
            continue

        # Check N-to-1 merge exact match (multiple WL blocks merged into 1 RW row)
        if not is_exact_overlap and len(overlaps) == 1:
            rw_one = overlaps[0]
            if rw_one.start_frame is not None and rw_one.end_frame is not None:
                sibs = [
                    s for s in wl_all_by_wt.get(wt_key, [])
                    if s.start_frame is not None and s.end_frame is not None
                    and (not s.job_id or not rw_one.job_id or s.job_id == rw_one.job_id)
                    and max(s.start_frame, rw_one.start_frame) <= min(s.end_frame, rw_one.end_frame)
                ]
                if len(sibs) > 1:
                    s_min = min(s.start_frame for s in sibs)
                    s_max = max(s.end_frame for s in sibs)
                    s_fr = sum(s.frames for s in sibs)
                    s_bb = sum(s.bbox for s in sibs)
                    if s_min == rw_one.start_frame and s_max == rw_one.end_frame and abs(s_fr - rw_one.frames) <= 1e-6 and abs(s_bb - rw_one.bbox) <= 1e-6:
                        if not job_err:
                            continue
                        is_exact_overlap = True

        rw_rng = _fmt_range(min_sf, max_ef)
        if is_exact_overlap and job_err:
            diff_kind = "mismatch"
            reason = job_err
        elif min_sf != wb.start_frame or max_ef != wb.end_frame:
            diff_kind = "mismatch"
            reason = f"Lệch Frame Start/End (Work Log {wl_rng} ≠ Review {rw_rng})"
            if job_err:
                reason = f"{reason} | {job_err}"
        else:
            diff_kind, reason = _classify_daily_diff(wb.frames, sum_fr, wb.bbox, sum_bb)
            if job_err:
                diff_kind = "mismatch"
                reason = f"{reason} | {job_err}"
        rw_dt = min((c.date_inspect or c.date_work for c in overlaps if (c.date_inspect or c.date_work)), default=None)
        diffs_by_ann_task[wt_key].append(DailyDiffRecord(
            date=wb.date,
            rw_date=rw_dt,
            task=wb.task,
            annotator_key=wb.annotator_key,
            annotator=wb.annotator_display,
            wl_frames=wb.frames,
            rw_frames=sum_fr,
            wl_bbox=wb.bbox,
            rw_bbox=sum_bb,
            wl_cells=wl_loc,
            rw_rows=_format_rw_locs(overlaps),
            inspectors=unique_inspector_names(record.inspector for record in overlaps),
            reason=reason,
            kind=diff_kind,
            job_id=wb.job_id or overlaps[0].job_id,
            wl_range=wl_rng,
            rw_range=rw_rng,
        ))

    # Pass 2.5 & 3 & 4: Remaining unresolved Work_log blocks
    for idx_wb, wb in enumerate(wl_in_range):
        if idx_wb in resolved_wb_indices:
            continue
        wt_key = (wb.annotator_key, wb.task)
        cands = rw_by_wt.get(wt_key, [])
        wl_rng = _fmt_range(wb.start_frame, wb.end_frame)
        wl_loc = f"{wb.sheet_name}!{wb.cell_ref}"

        # 2.5. Strict typo-recovery when Start/End Frame has a typo (e.g. Review typed 23000~23049 instead of 33000~33049):
        # MUST match unused & unreserved Review row on SAME (Annotator, Task, Job ID, Frames) AND exact BBOX from Work Log Note!
        # Never steals a Review row belonging to another day's Work Log block, and never uses date as a match filter.
        if wb.start_frame is not None and wb.end_frame is not None and wb.job_id and wb.frames > 0 and wb.bbox > 0:
            same_job_cands = [
                c for c in cands
                if id(c) not in matched_rw_ids
                and id(c) not in reserved_rw_ids
                and c.job_id == wb.job_id
                and abs(c.frames - wb.frames) <= 1e-6
                and c.bbox > 0 and abs(c.bbox - wb.bbox) <= 1e-6
            ]
            if same_job_cands:
                sj_match = min(
                    same_job_cands,
                    key=lambda c: (
                        0 if ((c.date_inspect or c.date_work) is not None and (c.date_inspect or c.date_work) >= wb.date) else 1,
                        c.row_num,
                    ),
                )
                _consume_rw(wt_key, [sj_match])
                rw_rng = _fmt_range(sj_match.start_frame, sj_match.end_frame)
                job_err = _check_dist_job_ids(wb, [sj_match])
                reason = f"Lệch Frame Start/End (Work Log {wl_rng} ≠ Review {rw_rng})"
                if job_err:
                    reason = f"{reason} | {job_err}"
                diffs_by_ann_task[wt_key].append(DailyDiffRecord(
                    date=wb.date,
                    rw_date=sj_match.date_inspect or sj_match.date_work,
                    task=wb.task,
                    annotator_key=wb.annotator_key,
                    annotator=wb.annotator_display,
                    wl_frames=wb.frames,
                    rw_frames=sj_match.frames,
                    wl_bbox=wb.bbox,
                    rw_bbox=sj_match.bbox,
                    wl_cells=wl_loc,
                    rw_rows=_format_rw_locs([sj_match]),
                    inspectors=unique_inspector_names([sj_match.inspector]),
                    reason=reason,
                    kind="mismatch",
                    job_id=wb.job_id,
                    wl_range=wl_rng,
                    rw_range=rw_rng,
                ))
                continue

        # 3. Fallback when WL block lacks Start/End Frame: match unused & unreserved Review by Job ID + Frames/BBOX
        if wb.start_frame is None or wb.end_frame is None:
            job_matches = [
                c for c in cands
                if id(c) not in matched_rw_ids
                and id(c) not in reserved_rw_ids
                and wb.job_id and c.job_id == wb.job_id
            ]
            if job_matches:
                # Prefer exact (frames, bbox) single match if one exists
                exact_fb = [
                    c for c in job_matches
                    if abs(c.frames - wb.frames) <= 1e-6
                    and (wb.bbox == 0 or c.bbox == 0 or abs(c.bbox - wb.bbox) <= 1e-6)
                ]
                if exact_fb:
                    job_matches = [exact_fb[0]]
                _consume_rw(wt_key, job_matches)
                sum_fr = sum(c.frames for c in job_matches)
                sum_bb = sum(c.bbox for c in job_matches)
                if abs(sum_fr - wb.frames) <= 1e-6 and abs(sum_bb - wb.bbox) <= 1e-6:
                    continue
                diff_kind, diff_reason = _classify_daily_diff(wb.frames, sum_fr, wb.bbox, sum_bb)
                rw_dt = job_matches[0].date_inspect or job_matches[0].date_work
                diffs_by_ann_task[wt_key].append(DailyDiffRecord(
                    date=wb.date,
                    rw_date=rw_dt,
                    task=wb.task,
                    annotator_key=wb.annotator_key,
                    annotator=wb.annotator_display,
                    wl_frames=wb.frames,
                    rw_frames=sum_fr,
                    wl_bbox=wb.bbox,
                    rw_bbox=sum_bb,
                    wl_cells=wl_loc,
                    rw_rows=_format_rw_locs(job_matches),
                    inspectors=unique_inspector_names(record.inspector for record in job_matches),
                    reason=diff_reason,
                    kind=diff_kind,
                    job_id=wb.job_id,
                    wl_range=wl_rng,
                    rw_range="",
                ))
                continue

        # 4. Not found anywhere in Rework -> pending_rework (or mismatch if Work Log Job ID is wrong vs Distribution!)
        job_err = _check_dist_job_ids(wb, [])
        diffs_by_ann_task[wt_key].append(DailyDiffRecord(
            date=wb.date,
            rw_date=None,
            task=wb.task,
            annotator_key=wb.annotator_key,
            annotator=wb.annotator_display,
            wl_frames=wb.frames,
            rw_frames=0.0,
            wl_bbox=wb.bbox,
            rw_bbox=0.0,
            wl_cells=wl_loc,
            rw_rows="",
            reason=f"Chưa có data bên Review | {job_err}" if job_err else "Chưa có data bên Review",
            kind="mismatch" if job_err else "pending_rework",
            job_id=wb.job_id,
            wl_range=wl_rng,
            rw_range="",
        ))

    # 5. Audit unmatched Review (Inspection) records in [start_date, end_date] against Distribution Sheet
    if ds.distribution_jobs:
        for r in ds.rework_records:
            if id(r) in matched_rw_ids:
                continue
            if selected_annotators is not None and r.worker_key not in selected_annotators:
                continue
            if r.start_frame is None or r.end_frame is None or not r.job_id:
                continue
            r_dt = r.date_work or r.date_inspect
            if r_dt is None or not (start_date <= r_dt <= end_date):
                continue
            task_jobs = _same_task_jobs_for(r.job_id, "")
            if not task_jobs:
                continue
            rw_sf, rw_ef = r.start_frame, r.end_frame
            if rw_sf > rw_ef:
                rw_sf, rw_ef = rw_ef, rw_sf
            rw_frame_dist = _find_frame_dist(task_jobs, r.job_id, r.worker_key, rw_sf, rw_ef)
            rw_issue = _diagnose_dist_mismatch("Review", r.job_id, r.worker_key, rw_sf, rw_ef, rw_frame_dist)
            if rw_issue:
                wt_key = (r.worker_key, r.task)
                diffs_by_ann_task[wt_key].append(DailyDiffRecord(
                    date=r_dt,
                    rw_date=r.date_inspect or r.date_work,
                    task=r.task,
                    annotator_key=r.worker_key,
                    annotator=r.worker_display,
                    wl_frames=0.0,
                    rw_frames=r.frames,
                    wl_bbox=0.0,
                    rw_bbox=r.bbox,
                    wl_cells="",
                    rw_rows=_format_rw_locs([r]),
                    inspectors=unique_inspector_names([r.inspector]),
                    reason=rw_issue,
                    kind="mismatch",
                    job_id=r.job_id,
                    wl_range="",
                    rw_range=_fmt_range(r.start_frame, r.end_frame),
                ))

    return diffs_by_ann_task, rw_agg_all, matched_rw_records


def _build_reviewer_audit(
    ds: LoadedDataset,
    start_date: datetime.date,
    end_date: datetime.date,
    selected_tasks: Optional[set] = None,
) -> Tuple[List[ReviewerAuditSummaryRow], List[ReviewerDiffRecord], List[WorklogNoteRecord]]:
    """Cross-check Reviewer Work Log vs Inspection & Rework, Annotator Work Log, and Distribution Sheet."""
    wl_agg = defaultdict(lambda: [0.0, 0.0, 0.0])  # (insp_key, base_task) -> [frames, bbox, time]
    for r in ds.reviewer_worklog_daily:
        if not (start_date <= r.date <= end_date):
            continue
        base_task = strip_rew_suffix(r.task)
        if selected_tasks is not None and base_task not in selected_tasks and r.task not in selected_tasks:
            continue
        if r.frames == 0 and r.bbox == 0 and r.work_time == 0:
            continue
        k = (r.annotator_key, base_task)
        wl_agg[k][0] += r.frames
        wl_agg[k][1] += r.bbox
        wl_agg[k][2] += r.work_time

    rw_agg = defaultdict(lambda: [0.0, 0.0, 0.0])  # (insp_key, base_task) -> [frames, bbox, time]
    rw_by_it = defaultdict(list)
    for rec in ds.rework_records:
        insp_key = normalize_annotator_key(rec.inspector)
        if insp_key not in FIRST_REVIEWER_KEYS:
            continue
        dt = rec.date_inspect or rec.date_work
        if dt is None or not (start_date <= dt <= end_date):
            continue
        base_task = strip_rew_suffix(rec.task)
        if selected_tasks is not None and base_task not in selected_tasks and rec.task not in selected_tasks:
            continue
        if rec.frames == 0 and rec.bbox == 0 and rec.work_time == 0:
            continue
        rw_by_it[(insp_key, base_task)].append(rec)

    # Index all Rework rows & Reviewer WL blocks (across ALL dates for cross-date matching) & Annotator WL blocks
    rw_all_by_it = defaultdict(list)
    for rec in ds.rework_records:
        insp_key = normalize_annotator_key(rec.inspector)
        if insp_key not in FIRST_REVIEWER_KEYS:
            continue
        rw_all_by_it[(insp_key, strip_rew_suffix(rec.task))].append(rec)

    rev_all_by_it = defaultdict(list)
    for wb in ds.reviewer_worklog_blocks:
        rev_all_by_it[(wb.annotator_key, strip_rew_suffix(wb.task))].append(wb)

    ann_blocks_by_task = defaultdict(list)
    for ab in ds.worklog_blocks:
        ann_blocks_by_task[strip_rew_suffix(ab.task)].append(ab)

    # Distribution segments indexed per TASK (each sheet = 1 mode; TASK NAME = 1 CVAT task).
    # Job IDs can repeat across sheets/servers, so ALWAYS scope lookups to the same task first.
    dist_by_tn: Dict[str, List[DistributionJobRecord]] = defaultdict(list)
    dist_by_sheet: Dict[str, List[DistributionJobRecord]] = defaultdict(list)
    for dj in ds.distribution_jobs:
        if dj.task_name_key:
            dist_by_tn[dj.task_name_key].append(dj)
        dist_by_sheet[str(dj.sheet_name).strip().upper()].append(dj)

    def _dist_task_segments(wb_blk: WorklogBlockRecord, base_task: str) -> List[DistributionJobRecord]:
        # 1. Prefer resolving by Job/Task Name first so we always check against the right task
        tn = normalize_job_name(wb_blk.job_name)
        if tn and tn in dist_by_tn:
            return dist_by_tn[tn]
        if tn:
            for pfx in ("OD_", "PFE_", "SSEG_", "ANPR_", "PPE_", "DPM_"):
                if tn.startswith(pfx) and tn[len(pfx):] in dist_by_tn:
                    return dist_by_tn[tn[len(pfx):]]
            hits = [k for k in dist_by_tn if k in tn or tn in k]
            if len(hits) == 1:
                return dist_by_tn[hits[0]]
        # 2. Fallback: resolve by recorded Job ID inside the same mode sheet (when col E Job Name is empty/unrecognized)
        sheet_segs = dist_by_sheet.get(base_task, [])
        if wb_blk.job_id and sheet_segs:
            same_id = [s for s in sheet_segs if s.job_id == wb_blk.job_id]
            if same_id:
                ctid = same_id[0].cvat_task_id
                tkey = same_id[0].task_name_key
                return [s for s in sheet_segs if (ctid and s.cvat_task_id == ctid) or (not ctid and s.task_name_key == tkey)]
        return []

    def _seg_range(segs: List[DistributionJobRecord]) -> str:
        return f"{min(s.start_frame for s in segs)}~{max(s.end_frame for s in segs)}"

    def _check_reviewer_dist(
        wb_blk: WorklogBlockRecord, base_task: str, ann_keys: set
    ) -> List[str]:
        if not ds.distribution_jobs or wb_blk.start_frame is None or wb_blk.end_frame is None:
            return []
        segs = _dist_task_segments(wb_blk, base_task)
        if not segs:
            return []
        sf, ef = min(wb_blk.start_frame, wb_blk.end_frame), max(wb_blk.start_frame, wb_blk.end_frame)
        task_lbl = segs[0].task_name or base_task
        rec_segs = [s for s in segs if wb_blk.job_id and s.job_id == wb_blk.job_id]
        # Segments (rows) in this task that physically contain the frames
        owner_segs = [s for s in segs if s.start_frame <= sf <= s.end_frame]
        owner_segs.sort(key=lambda s: (
            0 if s.job_id == wb_blk.job_id else 1,
            0 if (ann_keys and set(s.assignee_keys) & ann_keys) else 1,
        ))
        owner = owner_segs[0] if owner_segs else None

        issues: List[str] = []
        id_ok = bool(rec_segs) and any(s.start_frame <= sf <= s.end_frame for s in rec_segs) and any(
            s.start_frame <= ef <= s.end_frame for s in rec_segs
        )
        frame_wrong = False
        if not id_ok:
            if not rec_segs and wb_blk.job_id:
                sheet_segs = dist_by_sheet.get(base_task, [])
                other_segs = [s for s in sheet_segs if s.job_id == wb_blk.job_id]
                if other_segs and any(s.start_frame <= sf <= s.end_frame for s in other_segs) and any(
                    s.start_frame <= ef <= s.end_frame for s in other_segs
                ):
                    actual_tn = other_segs[0].task_name or other_segs[0].task_name_key
                    raw_first = str(wb_blk.job_name or "").splitlines()[0].strip() if wb_blk.job_name else "(trống)"
                    issues.append(
                        f"Sai Task Name WL Inspector: ghi '{raw_first}', đúng là '{actual_tn}' (Job {wb_blk.job_id})"
                    )
                    return issues
            rec_is_ann = bool(ann_keys) and any(set(s.assignee_keys) & ann_keys for s in rec_segs)
            frm_is_ann = owner is not None and bool(ann_keys) and bool(set(owner.assignee_keys) & ann_keys)
            if rec_segs and rec_is_ann and not frm_is_ann:
                # Job ID belongs to the annotator -> ID right, Start/End frame typed wrong
                frame_wrong = True
                ann_lbl = next((s.assignee_raw for s in rec_segs if set(s.assignee_keys) & ann_keys), "")
                issues.append(
                    f"Sai Frame WL Inspector: Job {wb_blk.job_id} ({_seg_range(rec_segs)}) của {ann_lbl} đúng, "
                    f"nhưng ghi frame {sf}~{ef}"
                    + (f" (thuộc Job {owner.job_id})" if owner is not None else "")
                    + f" [{task_lbl}]"
                )
            elif owner is not None and owner.job_id != wb_blk.job_id:
                rec_txt = f"{wb_blk.job_id} ({_seg_range(rec_segs)})" if rec_segs else f"{wb_blk.job_id or '(trống)'} (không có trong task)"
                issues.append(
                    f"Sai ID Job WL Inspector: ghi {rec_txt}, frame {sf}~{ef} thuộc Job {owner.job_id} "
                    f"({owner.start_frame}~{owner.end_frame}) [{task_lbl}]"
                )
            elif rec_segs:
                frame_wrong = True
                issues.append(
                    f"Sai Frame WL Inspector: Job {wb_blk.job_id} có frame {_seg_range(rec_segs)}, đang ghi {sf}~{ef} [{task_lbl}]"
                )
            elif owner is None:
                issues.append(f"Frame {sf}~{ef} không thuộc Job nào trong task [{task_lbl}]")
        elif wb_blk.job_name and task_lbl:
            # Job ID & Frame match Distribution, verify Task Name matches too
            wb_tn = _clean_tn(wb_blk.job_name)
            dist_tn = _clean_tn(task_lbl)
            wb_tn_nopfx = wb_tn
            for pfx in ("OD_", "PFE_", "SSEG_", "ANPR_", "PPE_", "DPM_"):
                if wb_tn_nopfx.startswith(pfx):
                    wb_tn_nopfx = wb_tn_nopfx[len(pfx):]
                    break
            if wb_tn and dist_tn and wb_tn != dist_tn and wb_tn_nopfx != dist_tn and dist_tn not in wb_tn and wb_tn not in dist_tn:
                raw_first = str(wb_blk.job_name).splitlines()[0].strip()
                issues.append(
                    f"Sai Task Name WL Inspector: ghi '{raw_first}', đúng là '{task_lbl}' (Job {wb_blk.job_id})"
                )

        return issues

    def _clean_tn(s: str) -> str:
        first = str(s or "").split("\n")[0]
        first = re.sub(r"\([^)]*\)", "", first)
        return normalize_job_name(first)

    def _tiered_match(
        wb_blk: WorklogBlockRecord,
        cands: list,
        get_tn,
        used_ids: Optional[set] = None,
        reserved_ids: Optional[set] = None,
        single: bool = False,
        max_tier: int = 5,
    ):
        """
        Match a Reviewer WL block against candidates (Rework rows or Annotator WL blocks) strictly by
        Task, Job ID, Frame Start/End, Frames & BBOX from Note (never using date as a match condition):
          1. Exact Start/End + same Job ID
          2. Exact Start/End + same task name (or Job ID missing / task name unknown)
          3. Overlapping frames + same Job ID (unclaimed/unreserved)
          4. Strict Start/End typo fallback: same Job ID + exact Frames + exact BBOX (>0) (unclaimed/unreserved)
          5. Overlapping frames + same task name (unclaimed/unreserved)
        Returns (matched_list, tier).
        """
        sf, ef = wb_blk.start_frame, wb_blk.end_frame
        wid = wb_blk.job_id
        wtn = _clean_tn(wb_blk.job_name)
        res_set = reserved_ids or set()
        used_set = used_ids or set()

        def same_id(c) -> bool:
            return bool(wid and c.job_id and c.job_id == wid)

        def same_tn(c) -> bool:
            t = get_tn(c)
            return bool(wtn and t and t == wtn)

        def tn_unknown(c) -> bool:
            t = get_tn(c)
            return not wtn or not t or not wid or not c.job_id

        def has_rng(c) -> bool:
            return c.start_frame is not None and c.end_frame is not None

        def is_free(c) -> bool:
            return id(c) not in used_set and id(c) not in res_set

        def pick(lst):
            if not single or len(lst) <= 1:
                return lst
            def _key(c):
                u_pen = 2 if id(c) in used_set else (1 if id(c) in res_set else 0)
                same_fr = abs(c.frames - wb_blk.frames) <= 1e-6
                same_bb = (wb_blk.bbox > 0 and c.bbox > 0 and abs(c.bbox - wb_blk.bbox) <= 1e-6)
                v_pen = 0 if (same_fr and same_bb) else (1 if same_fr else 2)
                c_dt = getattr(c, "date_inspect", None) or getattr(c, "date_work", None) or getattr(c, "date", None)
                ord_pen = 0 if (c_dt is not None and c_dt >= wb_blk.date) else 1
                return (u_pen, v_pen, ord_pen, getattr(c, "row_num", 0))
            return [min(lst, key=_key)]

        if sf is None or ef is None:
            if wid and max_tier >= 4:
                free_id = [c for c in cands if same_id(c) and is_free(c)]
                return pick(free_id), 4
            return [], 0
        exact = [c for c in cands if c.start_frame == sf and c.end_frame == ef]
        t1 = [c for c in exact if same_id(c)]
        if t1:
            return pick(t1), 1
        if max_tier < 2:
            return [], 0
        t2 = [c for c in exact if same_tn(c) or tn_unknown(c)]
        if t2:
            return pick(t2), 2
        if max_tier < 3:
            return [], 0
        t3 = [c for c in cands if has_rng(c) and is_free(c) and same_id(c) and max(sf, c.start_frame) <= min(ef, c.end_frame)]
        if t3:
            return t3, 3
        if max_tier < 4:
            return [], 0
        t4 = [
            c for c in cands
            if same_id(c) and is_free(c)
            and wb_blk.frames > 0 and abs(c.frames - wb_blk.frames) <= 1e-6
            and wb_blk.bbox > 0 and c.bbox > 0 and abs(c.bbox - wb_blk.bbox) <= 1e-6
        ]
        if t4:
            return pick(t4), 4
        if max_tier < 5:
            return [], 0
        t5 = [c for c in cands if has_rng(c) and is_free(c) and same_tn(c) and max(sf, c.start_frame) <= min(ef, c.end_frame)]
        if t5:
            return t5, 5
        return [], 0

    note_counts = defaultdict(int)
    rev_notes: List[WorklogNoteRecord] = []
    for n in ds.reviewer_worklog_notes:
        if not (start_date <= n.date <= end_date):
            continue
        base_task = strip_rew_suffix(n.task)
        if selected_tasks is not None and base_task not in selected_tasks and n.task not in selected_tasks:
            continue
        note_counts[(n.annotator_key, base_task)] += 1
        rev_notes.append(n)
    rev_notes.sort(key=lambda x: (x.annotator_display, x.date, x.task))

    diffs_by_it = defaultdict(list)
    matched_rw_ids = set()

    # Pre-reserve Rework rows that already exactly match outside-period Reviewer Work Log blocks
    reserved_rev_rw_ids: set = set()
    rev_in_range: List[WorklogBlockRecord] = []
    for wb in ds.reviewer_worklog_blocks:
        base_task = strip_rew_suffix(wb.task)
        if start_date <= wb.date <= end_date:
            if selected_tasks is not None and base_task not in selected_tasks and wb.task not in selected_tasks:
                continue
            rev_in_range.append(wb)
        elif wb.start_frame is not None and wb.end_frame is not None:
            it_key = (wb.annotator_key, base_task)
            m_out, _ = _tiered_match(
                wb, rw_all_by_it.get(it_key, []), lambda c: "",
                used_ids=reserved_rev_rw_ids, single=True, max_tier=2,
            )
            for m in m_out:
                reserved_rev_rw_ids.add(id(m))

    # Multi-pass matching for in-range Reviewer WL blocks (Exact first -> Overlaps & Strict Typo fallback)
    rev_rw_matches: Dict[int, Tuple[List[ReworkRecord], int]] = {}
    for idx_wb, wb in enumerate(rev_in_range):
        it_key = (wb.annotator_key, strip_rew_suffix(wb.task))
        m_rws, m_tier = _tiered_match(
            wb, rw_all_by_it.get(it_key, []), lambda c: "",
            used_ids=matched_rw_ids, reserved_ids=reserved_rev_rw_ids, single=True, max_tier=2,
        )
        if m_rws:
            rev_rw_matches[idx_wb] = (m_rws, m_tier)
            for m in m_rws:
                if id(m) not in matched_rw_ids:
                    matched_rw_ids.add(id(m))
                    rw_agg[it_key][0] += m.frames
                    rw_agg[it_key][1] += m.bbox
                    rw_agg[it_key][2] += m.work_time

    for idx_wb, wb in enumerate(rev_in_range):
        if idx_wb in rev_rw_matches:
            continue
        it_key = (wb.annotator_key, strip_rew_suffix(wb.task))
        m_rws, m_tier = _tiered_match(
            wb, rw_all_by_it.get(it_key, []), lambda c: "",
            used_ids=matched_rw_ids, reserved_ids=reserved_rev_rw_ids, single=True, max_tier=5,
        )
        rev_rw_matches[idx_wb] = (m_rws, m_tier)
        for m in m_rws:
            if id(m) not in matched_rw_ids:
                matched_rw_ids.add(id(m))
                rw_agg[it_key][0] += m.frames
                rw_agg[it_key][1] += m.bbox
                rw_agg[it_key][2] += m.work_time

    for idx_wb, wb in enumerate(rev_in_range):
        base_task = strip_rew_suffix(wb.task)
        it_key = (wb.annotator_key, base_task)
        insp_disp = wb.annotator_display
        wl_rng = _fmt_range(wb.start_frame, wb.end_frame)
        wl_loc = f"{wb.sheet_name}!{wb.cell_ref}"

        matched_rws, rw_tier = rev_rw_matches.get(idx_wb, ([], 0))

        # Match against Annotator Work Log (same mode, across ALL dates; scoped by Job ID / task name)
        ann_cands = ann_blocks_by_task.get(base_task, [])
        matched_anns, ann_tier = _tiered_match(wb, ann_cands, lambda a: _clean_tn(a.job_name))

        rw_fr = sum(c.frames for c in matched_rws)
        rw_bb = sum(c.bbox for c in matched_rws)
        rw_dt = min((c.date_inspect or c.date_work for c in matched_rws if (c.date_inspect or c.date_work)), default=None)
        rw_rng = _fmt_range(
            min((c.start_frame for c in matched_rws if c.start_frame is not None), default=None),
            max((c.end_frame for c in matched_rws if c.end_frame is not None), default=None),
        ) if matched_rws else ""
        ann_fr = sum(a.frames for a in matched_anns)
        ann_bb = sum(a.bbox for a in matched_anns)
        ann_rng = _fmt_range(
            min((a.start_frame for a in matched_anns if a.start_frame is not None), default=None),
            max((a.end_frame for a in matched_anns if a.end_frame is not None), default=None),
        ) if matched_anns else ""
        ann_name = ", ".join(sorted({a.annotator_display for a in matched_anns})) if matched_anns else (
            ", ".join(sorted({c.worker_display for c in matched_rws})) if matched_rws else "—"
        )

        reasons: List[str] = []
        kind = "ok"
        if not matched_rws:
            kind = "pending_rework"
            reasons.append("Chưa có data bên Review")
        else:
            if wl_rng and rw_rng and wl_rng != rw_rng:
                kind = "mismatch"
                reasons.append(f"Lệch Start/End vs Review ({wl_rng} ≠ {rw_rng})")
            if abs(wb.frames - rw_fr) > 1e-6:
                kind = "mismatch"
                reasons.append(f"Lệch Frame vs Review ({int(wb.frames)} ≠ {int(rw_fr)})")
            if wb.bbox > 0 and rw_bb > 0 and abs(wb.bbox - rw_bb) > 1e-6:
                kind = "mismatch"
                reasons.append(f"Lệch BBOX vs Review ({int(wb.bbox)} ≠ {int(rw_bb)})")

        if matched_anns:
            if wl_rng and ann_rng and wl_rng != ann_rng and ann_tier == 4:
                kind = "mismatch"
                reasons.append(f"Lệch Start/End vs WL Annotator ({wl_rng} ≠ {ann_rng})")
            if abs(wb.frames - ann_fr) > 1e-6:
                kind = "mismatch"
                reasons.append(f"Lệch Frame vs WL Annotator ({int(wb.frames)} ≠ {int(ann_fr)})")
            if wb.bbox > 0 and ann_bb > 0 and abs(wb.bbox - ann_bb) > 1e-6:
                kind = "mismatch"
                reasons.append(f"Lệch BBOX vs WL Annotator ({int(wb.bbox)} ≠ {int(ann_bb)})")
        elif wb.start_frame is not None and wb.end_frame is not None:
            if kind == "ok":
                kind = "mismatch"
            reasons.append("Không tìm thấy block tương ứng ở WL Annotator")

        ann_keys = {a.annotator_key for a in matched_anns} | {c.worker_key for c in matched_rws}
        dist_issues = _check_reviewer_dist(wb, base_task, ann_keys)
        if dist_issues:
            kind = "mismatch"
            reasons.extend(dist_issues)

        if kind != "ok":
            diffs_by_it[it_key].append(ReviewerDiffRecord(
                date=wb.date,
                rw_date=rw_dt,
                inspector_key=wb.annotator_key,
                inspector=insp_disp,
                annotator=ann_name,
                task=base_task,
                job_id=wb.job_id or (matched_rws[0].job_id if matched_rws else ""),
                wl_range=wl_rng,
                rw_range=rw_rng,
                ann_range=ann_rng,
                wl_frames=wb.frames,
                rw_frames=rw_fr,
                ann_frames=ann_fr,
                wl_bbox=wb.bbox,
                rw_bbox=rw_bb,
                ann_bbox=ann_bb,
                wl_cells=wl_loc,
                rw_rows=_format_rw_locs(matched_rws),
                ann_cells=", ".join(f"{a.sheet_name}!{a.cell_ref}" for a in matched_anns),
                reason=" | ".join(reasons),
                kind=kind,
            ))

    # Rework rows in period NOT matched by any Reviewer Work Log block (check across ALL dates before flagging!)
    for it_key, rlist in rw_by_it.items():
        insp_key, base_task = it_key
        rev_cands = rev_all_by_it.get(it_key, [])
        for rec in rlist:
            if id(rec) in matched_rw_ids:
                continue
            dt = rec.date_inspect or rec.date_work
            if dt is None:
                continue
            # Check if a matching Reviewer Work Log block exists on ANY date (handles date lag)
            has_any_wl_match = False
            if rec.start_frame is not None and rec.end_frame is not None:
                for wb_any in rev_cands:
                    if wb_any.start_frame is not None and wb_any.end_frame is not None:
                        if max(rec.start_frame, wb_any.start_frame) <= min(rec.end_frame, wb_any.end_frame):
                            if not rec.job_id or not wb_any.job_id or rec.job_id == wb_any.job_id:
                                has_any_wl_match = True
                                break
            elif rec.job_id:
                has_any_wl_match = any(wb_any.job_id == rec.job_id for wb_any in rev_cands)

            if has_any_wl_match:
                continue

            diffs_by_it[it_key].append(ReviewerDiffRecord(
                date=dt,
                rw_date=dt,
                inspector_key=insp_key,
                inspector=display_annotator_name(insp_key),
                annotator=rec.worker_display,
                task=base_task,
                job_id=rec.job_id,
                wl_range="",
                rw_range=_fmt_range(rec.start_frame, rec.end_frame),
                ann_range="",
                wl_frames=0.0,
                rw_frames=rec.frames,
                ann_frames=0.0,
                wl_bbox=0.0,
                rw_bbox=rec.bbox,
                ann_bbox=0.0,
                wl_cells="",
                rw_rows=_format_rw_locs([rec]),
                ann_cells="",
                reason="Có bên Review nhưng chưa ghi Work Log Inspector",
                kind="pending_rework",
            ))

    summary_rows: List[ReviewerAuditSummaryRow] = []
    all_keys = set(wl_agg.keys()) | set(rw_agg.keys())
    for insp_key, base_task in sorted(all_keys, key=lambda kv: (kv[1], display_annotator_name(kv[0]))):
        wl_fr, wl_bb, wl_tm = wl_agg.get((insp_key, base_task), (0.0, 0.0, 0.0))
        rw_fr, rw_bb, rw_tm = rw_agg.get((insp_key, base_task), (0.0, 0.0, 0.0))
        if wl_fr == 0 and wl_bb == 0 and wl_tm == 0 and rw_fr == 0 and rw_bb == 0 and rw_tm == 0:
            continue
        it_diffs = sorted(diffs_by_it.get((insp_key, base_task), []), key=lambda d: d.date)
        summary_rows.append(ReviewerAuditSummaryRow(
            task=base_task,
            inspector_key=insp_key,
            inspector=display_annotator_name(insp_key),
            wl_frames=wl_fr,
            rw_frames=rw_fr,
            wl_bbox=wl_bb,
            rw_bbox=rw_bb,
            wl_time=wl_tm,
            rw_time=rw_tm,
            note_count=note_counts.get((insp_key, base_task), 0),
            daily_diffs=it_diffs,
        ))

    all_rev_diffs: List[ReviewerDiffRecord] = []
    for r in summary_rows:
        all_rev_diffs.extend(r.daily_diffs)
    all_rev_diffs.sort(key=lambda d: (0 if d.kind == "mismatch" else 1, d.inspector.upper(), d.date, d.task))

    return summary_rows, all_rev_diffs, rev_notes


def build_report(
    ds: LoadedDataset,
    start_date: datetime.date,
    end_date: datetime.date,
    selected_annotators: Optional[set] = None,  # set of annotator_keys, None = all
    selected_tasks: Optional[set] = None,       # set of task strings for all report views (Task/Annotator/Reviewer/Xep_Hang/Audit), None = all
    source_mode: str = "rework",                # "rework" or "worklog"
    rework_date_mode: str = "col_v",            # legacy parameter (ignored; matched by Job ID + Frame Start/End)
) -> ReportResult:
    # Aggregate Work_log by (annotator_key, task)
    wl_agg_all = defaultdict(lambda: [0.0, 0.0, 0.0])  # [frames, bbox, work_time]
    for r in ds.worklog_daily:
        if not (start_date <= r.date <= end_date):
            continue
        if selected_annotators is not None and r.annotator_key not in selected_annotators:
            continue
        if r.frames == 0 and r.bbox == 0 and r.work_time == 0:
            continue
        k = (r.annotator_key, r.task)
        wl_agg_all[k][0] += r.frames
        wl_agg_all[k][1] += r.bbox
        wl_agg_all[k][2] += r.work_time

    # Count notes in range
    note_counts_all = defaultdict(int)
    filtered_notes: List[WorklogNoteRecord] = []
    for n in ds.worklog_notes:
        if not (start_date <= n.date <= end_date):
            continue
        if selected_annotators is not None and n.annotator_key not in selected_annotators:
            continue
        note_counts_all[(n.annotator_key, n.task)] += 1
        if selected_tasks is None or n.task in selected_tasks:
            filtered_notes.append(n)
    filtered_notes.sort(key=lambda x: (x.annotator_display, x.date, x.task))

    # Match Review rows directly to Work_log blocks in [start_date, end_date] by Job ID & Start/End Frame
    diffs_by_ann_task, rw_agg_all, matched_rw_records = _match_blocks_with_frame_ranges(ds, start_date, end_date, selected_annotators)

    # Combine ALL (annotator, task) rows
    all_keys = set(wl_agg_all.keys()) | set(rw_agg_all.keys())
    unfiltered_task_rows: List[AggregatedRow] = []
    for ann_key, task in all_keys:
        wl_fr, wl_bb, wl_tm = wl_agg_all.get((ann_key, task), (0.0, 0.0, 0.0))
        rw_fr, rw_bb, err1, err2, err1_bb, err2_bb = rw_agg_all.get((ann_key, task), (0.0, 0.0, 0.0, 0.0, 0.0, 0.0))

        if source_mode == "worklog":
            wl_load, wl_box = wl_fr, wl_bb
        else:
            wl_load, wl_box = rw_fr, rw_bb

        if wl_load == 0 and wl_box == 0 and wl_tm == 0 and err1 == 0 and err2 == 0 and wl_fr == 0 and rw_fr == 0:
            continue

        disp = ds.annotators.get(ann_key, display_annotator_name(ann_key))
        unfiltered_task_rows.append(AggregatedRow(
            task=task,
            annotator_key=ann_key,
            annotator=disp,
            work_load=wl_load,
            work_bbox=wl_box,
            work_time=wl_tm,
            err_first=err1,
            err_second=err2,
            err_first_bbox=err1_bb,
            err_second_bbox=err2_bb,
            wl_frames=wl_fr,
            wl_bbox=wl_bb,
            rw_frames=rw_fr,
            rw_bbox=rw_bb,
            note_count=note_counts_all.get((ann_key, task), 0),
            daily_diffs=sorted(diffs_by_ann_task.get((ann_key, task), []), key=lambda d: d.date),
        ))

    # Filter rows by selected_tasks for Template 1 (By TASK) and Audit tab
    all_rows: List[AggregatedRow] = [
        r for r in unfiltered_task_rows
        if selected_tasks is None or r.task in selected_tasks
    ]

    # Collect filtered daily diffs for Audit tab (real mismatches first, then pending rework)
    filtered_daily_diffs: List[DailyDiffRecord] = []
    for r in all_rows:
        filtered_daily_diffs.extend(r.daily_diffs)
    filtered_daily_diffs.sort(key=lambda d: (0 if d.kind == "mismatch" else 1, d.annotator.upper(), d.date, d.task))

    # Build Template 1: Group by TASK (Workers first: OD, SSEG, PFE, ANPR; then 1st Reviewers: OD_REW, SSEG_REW, PFE_REW, ANPR_REW)
    task_priority = {
        "OD": 0, "SSEG": 1, "PFE": 2, "ANPR": 3,
        "OD_REW": 10, "SSEG_REW": 11, "PFE_REW": 12, "ANPR_REW": 13,
    }
    grouped = defaultdict(list)
    for row in all_rows:
        grouped[row.task].append(row)

    task_groups: List[TaskSummary] = []
    for task in sorted(grouped.keys(), key=lambda t: (1 if t.endswith("_REW") else 0, task_priority.get(t, 99), t)):
        t_rows = sorted(grouped[task], key=lambda r: r.annotator.upper())
        task_groups.append(TaskSummary(
            task=task,
            rows=t_rows,
            total_work_load=sum(r.work_load for r in t_rows),
            total_work_bbox=sum(r.work_bbox for r in t_rows),
            total_work_time=sum(r.work_time for r in t_rows),
            total_err_first=sum(r.err_first for r in t_rows),
            total_err_second=sum(r.err_second for r in t_rows),
            total_err_first_bbox=sum(r.err_first_bbox for r in t_rows),
            total_err_second_bbox=sum(r.err_second_bbox for r in t_rows),
        ))

    # Build Template 2: Grand total across selected tasks per Annotator/Reviewer (Workers A-Z, then 1st Reviewers A-Z)
    by_ann = defaultdict(list)
    for row in all_rows:
        by_ann[row.annotator_key].append(row)

    annotator_rows: List[AggregatedRow] = []
    for ann_key, rlist in by_ann.items():
        tasks_joined = ", ".join(sorted(set(r.task for r in rlist)))
        combined_diffs: List[DailyDiffRecord] = []
        for r in rlist:
            combined_diffs.extend(r.daily_diffs)
        combined_diffs.sort(key=lambda d: (d.date, d.task))
        annotator_rows.append(AggregatedRow(
            task=tasks_joined,
            annotator_key=ann_key,
            annotator=rlist[0].annotator,
            work_load=sum(r.work_load for r in rlist),
            work_bbox=sum(r.work_bbox for r in rlist),
            work_time=sum(r.work_time for r in rlist),
            err_first=sum(r.err_first for r in rlist),
            err_second=sum(r.err_second for r in rlist),
            err_first_bbox=sum(r.err_first_bbox for r in rlist),
            err_second_bbox=sum(r.err_second_bbox for r in rlist),
            wl_frames=sum(r.wl_frames for r in rlist),
            wl_bbox=sum(r.wl_bbox for r in rlist),
            rw_frames=sum(r.rw_frames for r in rlist),
            rw_bbox=sum(r.rw_bbox for r in rlist),
            note_count=sum(r.note_count for r in rlist),
            daily_diffs=combined_diffs,
        ))
    annotator_rows.sort(key=lambda r: (1 if r.is_reviewer else 0, r.annotator.upper()))

    all_rows.sort(key=lambda r: (1 if r.task.endswith("_REW") else 0, task_priority.get(r.task, 99), r.task, r.annotator.upper()))

    # Build Inspector Groups (1st Reviewers: HIEN_03, KNGAN, MLINH, NGAN, NGUYEN, THUY, XUAN)
    # Grouped by INSPECTOR -> Annotator (exact match to sample template)
    insp_worker_agg = defaultdict(lambda: [0.0, 0.0, 0.0, 0.0, 0.0, 0.0, 0.0])  # [fr, bb, time, e1, e2, e1_bb, e2_bb]
    for rec in matched_rw_records:
        if selected_tasks is not None and rec.task not in selected_tasks:
            continue
        insp_key = normalize_annotator_key(rec.inspector)
        if insp_key not in FIRST_REVIEWER_KEYS:
            continue
        k = (insp_key, rec.worker_key)
        insp_worker_agg[k][0] += rec.frames
        insp_worker_agg[k][1] += rec.bbox
        insp_worker_agg[k][2] += rec.work_time
        insp_worker_agg[k][3] += rec.err_first
        insp_worker_agg[k][4] += rec.err_second
        insp_worker_agg[k][5] += rec.err_first_bbox
        insp_worker_agg[k][6] += rec.err_second_bbox

    by_insp: Dict[str, List[InspectorWorkerRow]] = defaultdict(list)
    for (insp_key, w_key), vals in insp_worker_agg.items():
        if vals[0] == 0 and vals[1] == 0 and vals[2] == 0:
            continue
        by_insp[insp_key].append(InspectorWorkerRow(
            inspector_key=insp_key,
            inspector=display_annotator_name(insp_key),
            annotator_key=w_key,
            annotator=ds.annotators.get(w_key, display_annotator_name(w_key)),
            work_load=vals[0],
            work_bbox=vals[1],
            work_time=vals[2],
            err_first=vals[3],
            err_second=vals[4],
            err_first_bbox=vals[5],
            err_second_bbox=vals[6],
        ))

    inspector_groups: List[InspectorSummary] = []
    for insp_key in sorted(by_insp.keys(), key=lambda k: display_annotator_name(k)):
        i_rows = sorted(by_insp[insp_key], key=lambda r: r.annotator.upper())
        inspector_groups.append(InspectorSummary(
            inspector_key=insp_key,
            inspector=display_annotator_name(insp_key),
            rows=i_rows,
            total_work_load=sum(r.work_load for r in i_rows),
            total_work_bbox=sum(r.work_bbox for r in i_rows),
            total_work_time=sum(r.work_time for r in i_rows),
            total_err_first=sum(r.err_first for r in i_rows),
            total_err_second=sum(r.err_second for r in i_rows),
            total_err_first_bbox=sum(r.err_first_bbox for r in i_rows),
            total_err_second_bbox=sum(r.err_second_bbox for r in i_rows),
        ))

    rev_audit_rows, rev_daily_diffs, rev_notes = _build_reviewer_audit(
        ds, start_date, end_date, selected_tasks=selected_tasks
    )

    return ReportResult(
        start_date=start_date,
        end_date=end_date,
        source_mode=source_mode,
        rework_date_mode=rework_date_mode,
        task_groups=task_groups,
        annotator_rows=annotator_rows,
        all_rows=all_rows,
        filtered_notes=filtered_notes,
        daily_diffs=filtered_daily_diffs,
        inspector_groups=inspector_groups,
        reviewer_audit_rows=rev_audit_rows,
        reviewer_daily_diffs=rev_daily_diffs,
        reviewer_notes=rev_notes,
    )
