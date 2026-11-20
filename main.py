import datetime
import json
import os
import sys
import unicodedata
from collections import Counter
from typing import Dict, Optional, Set

from PyQt6.QtCore import QDate, Qt, QThread, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QFont, QIcon, QKeySequence, QPainter, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QScrollArea,
    QSplitter,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from data_engine import LoadedDataset, ReportResult, build_report, load_dataset
from excel_exporter import export_report_to_excel
from ui_theme import (
    DARK_THEME,
    LIGHT_THEME,
    ReportTableDelegate,
    build_stylesheet,
    svg_icon,
    svg_pixmap,
)
from ui_widgets import (
    FileDropCard,
    FlowLayout,
    MetricCard,
    SegmentedControl,
    ToastBanner,
)


APP_NAME = "Report Tools"
APP_VERSION = "2.1"


class CalendarDateEdit(QDateEdit):
    """QDateEdit with an always-visible crisp calendar button on the right side."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setCalendarPopup(True)
        self.setDisplayFormat("yyyy.MM.dd")
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setToolTip("Bấm vào biểu tượng lịch bên phải để chọn ngày")

    def paintEvent(self, event):
        super().paintEvent(event)
        painter = QPainter(self)
        btn_w = 28
        icon_sz = 15
        x = self.width() - btn_w + (btn_w - icon_sz) // 2
        y = (self.height() - icon_sz) // 2
        pm = svg_pixmap("calendar", "#2563EB", icon_sz)
        painter.drawPixmap(x, y, icon_sz, icon_sz, pm)
        painter.end()


def _resource_path(relative: str) -> str:
    base = getattr(sys, "_MEIPASS", os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative)


def _app_base_dir() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


def _settings_path() -> Optional[str]:
    candidates = [
        os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "WorklogReportPro"),
        _app_base_dir(),
    ]
    for folder in candidates:
        try:
            os.makedirs(folder, exist_ok=True)
            return os.path.join(folder, "worklog_report_settings.json")
        except Exception:
            continue
    return None


def _load_settings() -> dict:
    p = _settings_path()
    if p and os.path.isfile(p):
        try:
            with open(p, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def _save_settings(data: dict):
    p = _settings_path()
    if not p:
        return
    try:
        with open(p, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


ONEDRIVE_PRIMARY_PATH = r"C:\Users\HVV-AI33\OneDrive - hanwhaworkplace\HVV - AI Lab Annotation - Tài liệu"


def _get_onedrive_dirs() -> list:
    """
    Return candidate OneDrive / SharePoint synced directories (checking both NFC and NFD Unicode forms,
    plus dynamic sub-directory discovery under OneDrive - hanwhaworkplace).
    """
    cands = [
        ONEDRIVE_PRIMARY_PATH,
        unicodedata.normalize("NFC", ONEDRIVE_PRIMARY_PATH),
        unicodedata.normalize("NFD", ONEDRIVE_PRIMARY_PATH),
    ]
    home = os.path.expanduser("~")
    od_roots = [
        r"C:\Users\HVV-AI33\OneDrive - hanwhaworkplace",
        os.environ.get("OneDriveCommercial", ""),
        os.environ.get("OneDrive", ""),
    ]
    try:
        if os.path.isdir(home):
            for entry in os.listdir(home):
                if entry.lower().startswith("onedrive"):
                    od_roots.append(os.path.join(home, entry))
    except Exception:
        pass

    for root in od_roots:
        if not root or not os.path.isdir(root):
            continue
        try:
            for sub in os.listdir(root):
                if "ai lab annotation" in sub.lower():
                    cands.append(os.path.join(root, sub))
        except Exception:
            pass
        cands.append(root)

    seen = set()
    unique = []
    for p in cands:
        if p and p not in seen:
            seen.add(p)
            unique.append(p)
    return unique


def _find_default_files(search_dirs):
    """Auto-detect original Work_log and Inspection_And_Rework files in candidate dirs."""
    wl_found, rw_found = "", ""
    for d in search_dirs:
        if not d or not os.path.isdir(d):
            continue
        wl_cands, rw_cands = [], []
        try:
            entries = os.listdir(d)
        except Exception:
            continue
        for fn in entries:
            low = fn.lower()
            if not low.endswith(".xlsx") or low.startswith("~$") or "copy" in low:
                continue
            full = os.path.join(d, fn)
            if "work_log" in low and "re_id" not in low:
                wl_cands.append(full)
            elif "inspection" in low and "rework" in low:
                rw_cands.append(full)
        wl_cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        rw_cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        if not wl_found and wl_cands:
            wl_found = wl_cands[0]
        if not rw_found and rw_cands:
            rw_found = rw_cands[0]
        if wl_found and rw_found:
            break
    return wl_found, rw_found


def _fmt_int(v: float) -> str:
    if abs(v - round(v)) < 1e-6:
        return f"{int(round(v)):,}"
    return f"{v:,.2f}"


def _fmt_time(v: float) -> str:
    if abs(v - round(v)) < 1e-6:
        return str(int(round(v)))
    return f"{v:.2f}".rstrip("0").rstrip(".")


def _fmt_f2(v: float) -> str:
    return f"{v:.2f}"


class LoaderThread(QThread):
    progress = pyqtSignal(int, str)
    finished_ok = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, wl_path: str, rw_path: str):
        super().__init__()
        self.wl_path = wl_path
        self.rw_path = rw_path

    def run(self):
        try:
            ds = load_dataset(
                self.wl_path,
                self.rw_path,
                progress_cb=lambda p, msg: self.progress.emit(p, msg),
            )
            self.finished_ok.emit(ds)
        except Exception as exc:
            self.failed.emit(str(exc))


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — Tổng hợp Work Log & Inspection/Review")
        self.resize(1480, 880)
        self.setMinimumSize(1120, 680)
        self.setAcceptDrops(True)

        icon_file = _resource_path("app_icon.ico")
        if os.path.isfile(icon_file):
            self.setWindowIcon(QIcon(icon_file))

        self.settings = _load_settings()
        self.theme = DARK_THEME if self.settings.get("theme") == "dark" else LIGHT_THEME
        self.dataset: Optional[LoadedDataset] = None
        self.current_report: Optional[ReportResult] = None
        self.loader_thread: Optional[LoaderThread] = None
        self._updating_filters = False
        self._isolated_annotator: Optional[str] = None
        self._task_chips: Dict[str, QPushButton] = {}

        self._build_ui()
        self._apply_theme(self.theme)
        self._setup_shortcuts()
        self._restore_files_and_load()

    # ------------------------------------------------------------------
    # UI CONSTRUCTION
    # ------------------------------------------------------------------
    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(0, 0, 0, 0)
        root.setSpacing(0)

        # ---- 1. TOP COMMAND BAR ----
        top_bar = QFrame()
        top_bar.setObjectName("TopBar")
        top_bar.setFixedHeight(58)
        top_lay = QHBoxLayout(top_bar)
        top_lay.setContentsMargins(16, 8, 16, 8)
        top_lay.setSpacing(12)

        logo = QLabel()
        logo.setFixedSize(34, 34)
        logo.setAlignment(Qt.AlignmentFlag.AlignCenter)
        logo.setStyleSheet("background-color: #2563EB; border-radius: 9px;")
        logo.setPixmap(svg_pixmap("chart", "#FFFFFF", 19))
        top_lay.addWidget(logo)

        title_col = QVBoxLayout()
        title_col.setSpacing(0)
        lbl_t = QLabel(APP_NAME)
        lbl_t.setObjectName("AppTitle")
        lbl_s = QLabel("Tổng hợp Work Log (+ Excel Notes) & Inspection/Review")
        lbl_s.setObjectName("AppSubtitle")
        title_col.addWidget(lbl_t)
        title_col.addWidget(lbl_s)
        top_lay.addLayout(title_col)

        top_lay.addStretch(1)

        self.btn_onedrive = QPushButton(" ☁ Nguồn OneDrive")
        self.btn_onedrive.setToolTip(f"Nạp trực tiếp từ thư mục OneDrive Sync:\n{ONEDRIVE_PRIMARY_PATH}")
        self.btn_onedrive.clicked.connect(self.switch_to_onedrive_source)
        top_lay.addWidget(self.btn_onedrive)

        self.btn_reload = QPushButton(" Nạp lại (F5)")
        self.btn_reload.setToolTip("Đọc lại 2 file Excel mới nhất từ ổ đĩa / OneDrive (F5)")
        self.btn_reload.clicked.connect(self.start_load_data)
        top_lay.addWidget(self.btn_reload)

        self.btn_theme = QPushButton()
        self.btn_theme.setToolTip("Chuyển giao diện Sáng / Tối")
        self.btn_theme.setFixedHeight(32)
        self.btn_theme.clicked.connect(self.toggle_theme)
        top_lay.addWidget(self.btn_theme)

        self.btn_copy = QPushButton(" Copy bảng")
        self.btn_copy.setToolTip("Copy bảng đang xem vào Clipboard (Ctrl+C)")
        self.btn_copy.clicked.connect(self.copy_current_table)
        top_lay.addWidget(self.btn_copy)

        self.btn_export = QPushButton(" Xuất Excel (.xlsx)")
        self.btn_export.setObjectName("SuccessBtn")
        self.btn_export.setToolTip("Xuất báo cáo Excel chuẩn mẫu kèm công thức (Ctrl+E)")
        self.btn_export.clicked.connect(self.export_excel)
        top_lay.addWidget(self.btn_export)

        root.addWidget(top_bar)

        # Thin animated progress bar directly under top bar
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setFixedHeight(4)
        self.progress.setVisible(False)
        root.addWidget(self.progress)

        # ---- 2. MAIN WORKSPACE SPLITTER ----
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        root.addWidget(splitter, 1)

        # ---- LEFT SIDEBAR ----
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setMinimumWidth(320)
        sidebar.setMaximumWidth(400)
        sb_lay = QVBoxLayout(sidebar)
        sb_lay.setContentsMargins(14, 12, 14, 12)
        sb_lay.setSpacing(12)

        # Section A: Date Range
        sb_lay.addWidget(self._section_label("KHOẢNG THỜI GIAN"))
        d_row = QHBoxLayout()
        d_row.setSpacing(6)
        self.dt_start = CalendarDateEdit()
        self.dt_start.setDate(QDate(2026, 11, 9))
        self.dt_start.dateChanged.connect(self._on_date_changed)
        lbl_arrow = QLabel("→")
        lbl_arrow.setStyleSheet("font-weight: 700; color: #64748B;")
        self.dt_end = CalendarDateEdit()
        self.dt_end.setDate(QDate(2026, 11, 15))
        self.dt_end.dateChanged.connect(self._on_date_changed)
        d_row.addWidget(self.dt_start, 1)
        d_row.addWidget(lbl_arrow)
        d_row.addWidget(self.dt_end, 1)
        sb_lay.addLayout(d_row)

        p_row = QHBoxLayout()
        p_row.setSpacing(5)
        for lbl, mode in [
            ("Tuần trước", "last_week"),
            ("Tuần này", "this_week"),
            ("Tháng này", "this_month"),
            ("Tất cả", "all"),
        ]:
            b = QPushButton(lbl)
            b.setObjectName("PresetBtn")
            b.clicked.connect(lambda _c=False, m=mode: self._apply_date_preset(m))
            p_row.addWidget(b, 1)
        sb_lay.addLayout(p_row)

        # Section B: Data Source & Mode (Segmented Controls)
        sb_lay.addWidget(self._section_label("NGUỒN DỮ LIỆU & ĐỐI CHIẾU"))
        lbl_src = QLabel("Sản lượng Frame / BBOX:")
        lbl_src.setStyleSheet("font-size: 11.5px; color: #64748B;")
        sb_lay.addWidget(lbl_src)
        self.seg_source = SegmentedControl([
            ("Review #IMAGE", "rework"),
            ("Work Log + Note", "worklog"),
        ])
        if self.settings.get("source_mode") in ("rework", "worklog"):
            self.seg_source.set_value(self.settings["source_mode"])
        self.seg_source.changed.connect(self._on_option_changed)
        sb_lay.addWidget(self.seg_source)

        self.chk_highlight = QCheckBox("Cảnh báo highlight nếu 2 file lệch nhau")
        self.chk_highlight.setChecked(True)
        self.chk_highlight.toggled.connect(self.refresh_report)
        sb_lay.addWidget(self.chk_highlight)

        # Section C: Task Filter Chips
        t_hdr = QHBoxLayout()
        t_hdr.addWidget(self._section_label("LỌC TASK"))
        t_hdr.addStretch(1)
        b_tall = QPushButton("Tất cả")
        b_tall.setObjectName("GhostBtn")
        b_tall.clicked.connect(lambda: self._set_all_task_chips(True))
        b_tact = QPushButton("Có data")
        b_tact.setObjectName("GhostBtn")
        b_tact.clicked.connect(self._select_active_tasks)
        b_tnone = QPushButton("Bỏ")
        b_tnone.setObjectName("GhostBtn")
        b_tnone.clicked.connect(lambda: self._set_all_task_chips(False))
        t_hdr.addWidget(b_tall)
        t_hdr.addWidget(b_tact)
        t_hdr.addWidget(b_tnone)
        sb_lay.addLayout(t_hdr)

        self.task_scroll = QScrollArea()
        self.task_scroll.setWidgetResizable(True)
        self.task_scroll.setFrameShape(QFrame.Shape.NoFrame)
        self.task_scroll.viewport().setObjectName("TransparentHost")
        self.task_scroll.setMaximumHeight(96)
        self.task_chip_host = QWidget()
        self.task_chip_host.setObjectName("TransparentHost")
        self.task_flow = FlowLayout(self.task_chip_host, margin=2, h_spacing=5, v_spacing=5)
        self.task_scroll.setWidget(self.task_chip_host)
        sb_lay.addWidget(self.task_scroll)

        # Section D: Annotator Filter (1 person or many)
        a_hdr = QHBoxLayout()
        a_hdr.addWidget(self._section_label("ANNOTATOR (1 HOẶC NHIỀU NGƯỜI)"))
        a_hdr.addStretch(1)
        b_aall = QPushButton("Tất cả")
        b_aall.setObjectName("GhostBtn")
        b_aall.clicked.connect(lambda: self._set_all_anns(True))
        b_aact = QPushButton("Có data")
        b_aact.setObjectName("GhostBtn")
        b_aact.clicked.connect(self._select_active_annotators)
        b_anone = QPushButton("Bỏ")
        b_anone.setObjectName("GhostBtn")
        b_anone.clicked.connect(lambda: self._set_all_anns(False))
        a_hdr.addWidget(b_aall)
        a_hdr.addWidget(b_aact)
        a_hdr.addWidget(b_anone)
        sb_lay.addLayout(a_hdr)

        self.ed_search = QLineEdit()
        self.ed_search.setPlaceholderText("Tìm nhanh annotator... (Ctrl+F)")
        self.ed_search.setClearButtonEnabled(True)
        self.ed_search.textChanged.connect(self._filter_ann_search)
        sb_lay.addWidget(self.ed_search)

        self.list_anns = QListWidget()
        self.list_anns.setToolTip("Tick để chọn nhiều người • Nhấp đúp vào tên để xem riêng 1 người")
        self.list_anns.itemChanged.connect(self._on_ann_item_changed)
        self.list_anns.itemDoubleClicked.connect(self._isolate_single_annotator)
        sb_lay.addWidget(self.list_anns, 1)

        lbl_hint = QLabel("Mẹo: Nhấp đúp 1 tên để xem riêng người đó.")
        lbl_hint.setStyleSheet("font-size: 11px; color: #64748B; font-style: italic;")
        sb_lay.addWidget(lbl_hint)

        splitter.addWidget(sidebar)

        # ---- RIGHT CONTENT WORKSPACE ----
        main_area = QWidget()
        main_lay = QVBoxLayout(main_area)
        main_lay.setContentsMargins(16, 12, 16, 10)
        main_lay.setSpacing(12)

        # File Drop Cards Row
        files_row = QHBoxLayout()
        files_row.setSpacing(10)
        self.card_wl = FileDropCard("FILE WORK LOG (CẢ EXCEL NOTES)", "Kéo thả hoặc chọn Work_log_*.xlsx", "#2563EB")
        self.card_wl.browse_clicked.connect(lambda: self._browse_file("wl"))
        self.card_wl.file_dropped.connect(lambda p: self._set_file_and_reload("wl", p))

        self.card_rw = FileDropCard("FILE INSPECTION & REVIEW", "Kéo thả hoặc chọn Inspection_And_Rework_*.xlsx", "#059669")
        self.card_rw.browse_clicked.connect(lambda: self._browse_file("rw"))
        self.card_rw.file_dropped.connect(lambda p: self._set_file_and_reload("rw", p))

        files_row.addWidget(self.card_wl, 1)
        files_row.addWidget(self.card_rw, 1)
        main_lay.addLayout(files_row)

        # 5 Executive KPI Cards
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(10)
        self.kpi_people = MetricCard("Nhân sự trong kỳ", "users", "#2563EB")
        self.kpi_frames = MetricCard("Work Load (Frame)", "layers", "#0F172A")
        self.kpi_bbox = MetricCard("Work BBOX", "box", "#7C3AED")
        self.kpi_time = MetricCard("Work Time (MH)", "clock", "#0284C7")
        self.kpi_quality = MetricCard("Lỗi & Đối chiếu 2 file", "shield", "#059669", clickable=True)
        self.kpi_quality.setToolTip("Bấm để xem ngay danh sách ngày bị lệch giữa 2 file")
        self.kpi_quality.clicked.connect(lambda: (self.tabs.setCurrentIndex(4), self.audit_subtabs.setCurrentIndex(0)))
        for card in (self.kpi_people, self.kpi_frames, self.kpi_bbox, self.kpi_time, self.kpi_quality):
            kpi_row.addWidget(card, 1)
        main_lay.addLayout(kpi_row)

        # Single-person isolation banner (hidden unless active)
        self.isolate_banner = QFrame()
        self.isolate_banner.setObjectName("IsolateBanner")
        self.isolate_banner.setVisible(False)
        ib_lay = QHBoxLayout(self.isolate_banner)
        ib_lay.setContentsMargins(12, 6, 10, 6)
        self.lbl_isolate = QLabel("")
        self.lbl_isolate.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        btn_unisolate = QPushButton("× Hiện lại tất cả")
        btn_unisolate.setObjectName("PrimaryBtn")
        btn_unisolate.setFixedHeight(26)
        btn_unisolate.clicked.connect(self._clear_isolation)
        ib_lay.addWidget(self.lbl_isolate)
        ib_lay.addStretch(1)
        ib_lay.addWidget(btn_unisolate)
        main_lay.addWidget(self.isolate_banner)

        # Tabs (Task, Annotator, Reviewer, Bảng Xếp Hạng, Đối chiếu & Notes)
        self.tabs = QTabWidget()

        # Tab 1: Task (Workers only)
        self.tbl_t1 = QTableWidget()
        self._setup_table(self.tbl_t1)
        self.del_t1 = ReportTableDelegate(divider_col=[9, 13, 14], parent=self.tbl_t1)
        self.tbl_t1.setItemDelegate(self.del_t1)
        self.tbl_t1.cellDoubleClicked.connect(self._on_t1_double_click)
        self.tabs.addTab(self.tbl_t1, "Task")

        # Tab 2: Annotator (Workers only)
        self.tbl_t2 = QTableWidget()
        self._setup_table(self.tbl_t2)
        self.del_t2 = ReportTableDelegate(divider_col=8, parent=self.tbl_t2)
        self.tbl_t2.setItemDelegate(self.del_t2)
        self.tbl_t2.cellDoubleClicked.connect(self._on_t2_double_click)
        self.tabs.addTab(self.tbl_t2, "Annotator")

        # Tab 3: Reviewer (Dedicated tab for 1st Reviewers matching exact INSPECTOR -> Annotator template)
        self.tbl_rev = QTableWidget()
        self._setup_table(self.tbl_rev)
        self.del_rev = ReportTableDelegate(divider_col=7, parent=self.tbl_rev)
        self.tbl_rev.setItemDelegate(self.del_rev)
        self.tbl_rev.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_rev, r))
        self.tabs.addTab(self.tbl_rev, "Reviewer")

        # Tab 4: Bảng Xếp Hạng (Lowest Error Rate Ranking for Workers & 1st Reviewers)
        tab_rank = QWidget()
        rank_lay = QVBoxLayout(tab_rank)
        rank_lay.setContentsMargins(10, 10, 10, 10)
        rank_lay.setSpacing(8)

        rank_top = QHBoxLayout()
        lbl_rk_info = QLabel("🏆 Bảng xếp hạng tỉ lệ lỗi thấp nhất (Hạng 1 = Tỉ lệ lỗi thấp nhất)")
        lbl_rk_info.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        rank_top.addWidget(lbl_rk_info)
        rank_top.addStretch(1)
        rank_top.addWidget(QLabel("Xếp hạng ưu tiên theo:"))
        self.seg_rank_by = SegmentedControl([
            ("Tỉ lệ lỗi Frame (%)", "frame"),
            ("Tỉ lệ lỗi BBOX (%)", "bbox"),
        ])
        self.seg_rank_by.setFixedWidth(280)
        self.seg_rank_by.changed.connect(lambda _v: self._render_ranking_tables(self.current_report) if self.current_report else None)
        rank_top.addWidget(self.seg_rank_by)
        rank_lay.addLayout(rank_top)

        rank_split = QSplitter(Qt.Orientation.Horizontal)

        w_rank_box = QWidget()
        w_rlay = QVBoxLayout(w_rank_box)
        w_rlay.setContentsMargins(0, 0, 4, 0)
        w_rlay.setSpacing(5)
        self.lbl_rank_worker = QLabel("🥇 Xếp hạng Annotator (Worker — Lỗi Lần 1 + Pangyo)")
        self.lbl_rank_worker.setStyleSheet("font-weight: 700; font-size: 12.5px; color: #2563EB;")
        w_rlay.addWidget(self.lbl_rank_worker)
        self.tbl_rank_worker = QTableWidget()
        self._setup_table(self.tbl_rank_worker)
        self.tbl_rank_worker.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_rank_worker, r))
        w_rlay.addWidget(self.tbl_rank_worker)
        rank_split.addWidget(w_rank_box)

        rev_rank_box = QWidget()
        rev_rlay = QVBoxLayout(rev_rank_box)
        rev_rlay.setContentsMargins(4, 0, 0, 0)
        rev_rlay.setSpacing(5)
        self.lbl_rank_reviewer = QLabel("🥇 Xếp hạng Người Review Lần 1 (Inspector — Lỗi Pangyo)")
        self.lbl_rank_reviewer.setStyleSheet("font-weight: 700; font-size: 12.5px; color: #059669;")
        rev_rlay.addWidget(self.lbl_rank_reviewer)
        self.tbl_rank_reviewer = QTableWidget()
        self._setup_table(self.tbl_rank_reviewer)
        self.tbl_rank_reviewer.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_rank_reviewer, r))
        rev_rlay.addWidget(self.tbl_rank_reviewer)
        rank_split.addWidget(rev_rank_box)

        rank_split.setSizes([600, 600])
        rank_lay.addWidget(rank_split, 1)
        self.tabs.addTab(tab_rank, "Bảng Xếp Hạng")

        # Tab 4: Audit, Daily Mismatches & Excel Notes Inspector
        tab_audit = QWidget()
        audit_lay = QVBoxLayout(tab_audit)
        audit_lay.setContentsMargins(10, 10, 10, 10)
        audit_lay.setSpacing(8)

        audit_split = QSplitter(Qt.Orientation.Vertical)
        top_a = QWidget()
        top_alay = QVBoxLayout(top_a)
        top_alay.setContentsMargins(0, 0, 0, 0)
        top_alay.setSpacing(5)
        self.lbl_audit_hdr = QLabel("Đối chiếu Work Log (+ Note) vs Review (#IMAGE)")
        self.lbl_audit_hdr.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        top_alay.addWidget(self.lbl_audit_hdr)
        self.tbl_audit = QTableWidget()
        self._setup_table(self.tbl_audit)
        self.tbl_audit.cellClicked.connect(self._on_audit_row_clicked)
        top_alay.addWidget(self.tbl_audit)
        audit_split.addWidget(top_a)

        bot_wrap = QWidget()
        bot_lay = QVBoxLayout(bot_wrap)
        bot_lay.setContentsMargins(0, 6, 0, 0)
        bot_lay.setSpacing(5)

        filter_row = QHBoxLayout()
        self.lbl_detail_hdr = QLabel("Chi tiết đối chiếu theo Frame Start/End & Excel Notes")
        self.lbl_detail_hdr.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        filter_row.addWidget(self.lbl_detail_hdr)
        filter_row.addStretch(1)
        filter_row.addWidget(QLabel("Lọc theo người:"))
        self.cb_note_ann = QComboBox()
        self.cb_note_ann.setMinimumWidth(155)
        self.cb_note_ann.currentIndexChanged.connect(self._on_detail_ann_changed)
        filter_row.addWidget(self.cb_note_ann)
        bot_lay.addLayout(filter_row)

        self.audit_subtabs = QTabWidget()

        # Sub-tab 0: Daily Mismatch Dates
        self.tbl_daily_diffs = QTableWidget()
        self._setup_table(self.tbl_daily_diffs)
        self.tbl_daily_diffs.cellDoubleClicked.connect(self._on_daily_diff_double_click)
        self.audit_subtabs.addTab(self.tbl_daily_diffs, "Đối chiếu Start/End (0)")

        # Sub-tab 1: Excel Notes
        self.tbl_notes = QTableWidget()
        self._setup_table(self.tbl_notes)
        self.audit_subtabs.addTab(self.tbl_notes, "Excel Notes Work Log (0)")

        bot_lay.addWidget(self.audit_subtabs, 1)
        audit_split.addWidget(bot_wrap)
        audit_split.setSizes([290, 290])

        audit_lay.addWidget(audit_split)
        self.tabs.addTab(tab_audit, "Đối chiếu && Notes")

        main_lay.addWidget(self.tabs, 1)

        # Bottom Export Sheet Selector Bar
        foot = QHBoxLayout()
        foot.setSpacing(14)
        lbl_exp = QLabel("Sheet khi xuất Excel:")
        lbl_exp.setStyleSheet("font-size: 12px; font-weight: 600; color: #64748B;")
        foot.addWidget(lbl_exp)
        self.chk_exp_t1 = QCheckBox("Task")
        self.chk_exp_t1.setChecked(True)
        self.chk_exp_t2 = QCheckBox("Annotator")
        self.chk_exp_t2.setChecked(True)
        self.chk_exp_audit = QCheckBox("Đối chiếu & Notes")
        self.chk_exp_audit.setChecked(True)
        foot.addWidget(self.chk_exp_t1)
        foot.addWidget(self.chk_exp_t2)
        foot.addWidget(self.chk_exp_audit)
        foot.addStretch(1)

        self.lbl_period_badge = QLabel("")
        self.lbl_period_badge.setStyleSheet("font-size: 12px; font-weight: 700; color: #2563EB;")
        foot.addWidget(self.lbl_period_badge)
        main_lay.addLayout(foot)

        splitter.addWidget(main_area)
        splitter.setSizes([335, 1145])

        # Status Bar + Floating Toast
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Kéo thả 2 file Excel vào cửa sổ hoặc bấm chọn file để bắt đầu.")

        self.toast = ToastBanner(self)

    def _section_label(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("SectionOverline")
        return lbl

    def _setup_table(self, tbl: QTableWidget):
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        tbl.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        tbl.verticalHeader().setVisible(False)
        tbl.verticalHeader().setDefaultSectionSize(28)
        tbl.horizontalHeader().setStretchLastSection(False)
        tbl.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+F"), self, activated=lambda: (self.ed_search.setFocus(), self.ed_search.selectAll()))
        QShortcut(QKeySequence("Ctrl+E"), self, activated=self.export_excel)
        QShortcut(QKeySequence("Ctrl+C"), self, activated=self.copy_current_table)
        QShortcut(QKeySequence("F5"), self, activated=self.start_load_data)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "toast"):
            self.toast.reposition()

    # ------------------------------------------------------------------
    # GLOBAL DRAG & DROP (.xlsx anywhere on window)
    # ------------------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for u in event.mimeData().urls():
                if u.toLocalFile().lower().endswith(".xlsx"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        changed = False
        for u in event.mimeData().urls():
            p = u.toLocalFile()
            if not p.lower().endswith(".xlsx"):
                continue
            low = os.path.basename(p).lower()
            if "inspection" in low or "rework" in low:
                self.card_rw.set_path(p, self.theme["name"] == "dark")
                changed = True
            else:
                self.card_wl.set_path(p, self.theme["name"] == "dark")
                changed = True
        if changed:
            self._save_current_settings()
            self.start_load_data()
            event.acceptProposedAction()

    # ------------------------------------------------------------------
    # THEME MANAGEMENT
    # ------------------------------------------------------------------
    def _apply_theme(self, theme: dict):
        self.theme = theme
        is_dark = theme["name"] == "dark"
        self.setStyleSheet(build_stylesheet(theme))
        self.del_t1.set_theme(theme)
        self.del_t2.set_theme(theme)
        self.del_rev.set_theme(theme)

        ic_color = theme["text_secondary"]
        self.btn_reload.setIcon(svg_icon("refresh", ic_color, 15))
        self.btn_copy.setIcon(svg_icon("copy", ic_color, 15))
        self.btn_export.setIcon(svg_icon("download", "#FFFFFF", 15))
        if is_dark:
            self.btn_theme.setIcon(svg_icon("sun", "#FBBF24", 16))
            self.btn_theme.setText(" Sáng")
        else:
            self.btn_theme.setIcon(svg_icon("moon", "#475569", 16))
            self.btn_theme.setText(" Tối")

        self.card_wl.set_path(self.card_wl.path(), is_dark)
        self.card_rw.set_path(self.card_rw.path(), is_dark)

        if self.current_report:
            self._render_kpis(self.current_report)
            self._render_template1(self.current_report)
            self._render_template2(self.current_report)
            self._render_reviewer_tab(self.current_report)
            self._render_ranking_tables(self.current_report)
            self._render_audit_table(self.current_report)
            self._render_daily_diffs_table()
            self._render_notes_table()

    def toggle_theme(self):
        next_t = DARK_THEME if self.theme["name"] == "light" else LIGHT_THEME
        self._apply_theme(next_t)
        self._save_current_settings()

    def _save_current_settings(self):
        self.settings["theme"] = self.theme["name"]
        self.settings["worklog_path"] = self.card_wl.path()
        self.settings["rework_path"] = self.card_rw.path()
        self.settings["source_mode"] = self.seg_source.current_value()
        _save_settings(self.settings)

    # ------------------------------------------------------------------
    # FILE DISCOVERY & LOADING
    # ------------------------------------------------------------------
    def _restore_files_and_load(self):
        base = _app_base_dir()
        parent = os.path.dirname(base)
        od_dirs = _get_onedrive_dirs()

        # 1. Highest priority: OneDrive / SharePoint synced folder
        od_wl, od_rw = _find_default_files(od_dirs)
        # 2. Fallback: local folder next to .exe
        local_wl, local_rw = _find_default_files([base, parent, os.getcwd()])

        saved_wl = self.settings.get("worklog_path", "")
        saved_rw = self.settings.get("rework_path", "")

        # Prefer OneDrive whenever available; otherwise use saved or local
        if od_wl:
            wl = saved_wl if (saved_wl and os.path.isfile(saved_wl) and "onedrive" in saved_wl.lower()) else od_wl
        else:
            wl = saved_wl if (saved_wl and os.path.isfile(saved_wl)) else local_wl

        if od_rw:
            rw = saved_rw if (saved_rw and os.path.isfile(saved_rw) and "onedrive" in saved_rw.lower()) else od_rw
        else:
            rw = saved_rw if (saved_rw and os.path.isfile(saved_rw)) else local_rw

        is_dark = self.theme["name"] == "dark"
        self.card_wl.set_path(wl, is_dark)
        self.card_rw.set_path(rw, is_dark)
        self._save_current_settings()

        if wl or rw:
            self.start_load_data()

    def switch_to_onedrive_source(self):
        od_dirs = _get_onedrive_dirs()
        od_wl, od_rw = _find_default_files(od_dirs)
        if not od_wl and not od_rw:
            # Open file dialog starting at OneDrive folder if found
            start_d = next((d for d in od_dirs if os.path.isdir(d)), _app_base_dir())
            QMessageBox.information(
                self,
                "Chưa tìm thấy file trong OneDrive",
                f"Không tự tìm thấy Work_log / Rework trong:\n{ONEDRIVE_PRIMARY_PATH}\n\n"
                f"Bạn hãy chọn file trực tiếp trong hộp thoại tiếp theo.",
            )
            self._browse_file("wl", start_override=start_d)
            return

        is_dark = self.theme["name"] == "dark"
        if od_wl:
            self.card_wl.set_path(od_wl, is_dark)
        if od_rw:
            self.card_rw.set_path(od_rw, is_dark)
        self._save_current_settings()
        self.toast.show_toast("☁ Đã chuyển sang nguồn OneDrive Sync!", "info")
        self.start_load_data()

    def _browse_file(self, kind: str, start_override: str = ""):
        cur = self.card_wl.path() if kind == "wl" else self.card_rw.path()
        if start_override and os.path.isdir(start_override):
            start_dir = start_override
        elif cur and os.path.isdir(os.path.dirname(cur)):
            start_dir = os.path.dirname(cur)
        else:
            od_dirs = _get_onedrive_dirs()
            start_dir = next((d for d in od_dirs if os.path.isdir(d)), _app_base_dir())
        title = "Chọn file Work Log (.xlsx)" if kind == "wl" else "Chọn file Inspection & Review (.xlsx)"
        path, _ = QFileDialog.getOpenFileName(self, title, start_dir, "Excel Workbook (*.xlsx)")
        if path:
            self._set_file_and_reload(kind, path)

    def _set_file_and_reload(self, kind: str, path: str):
        is_dark = self.theme["name"] == "dark"
        if kind == "wl":
            self.card_wl.set_path(path, is_dark)
        else:
            self.card_rw.set_path(path, is_dark)
        self._save_current_settings()
        self.start_load_data()

    def start_load_data(self):
        wl = self.card_wl.path()
        rw = self.card_rw.path()
        if not wl and not rw:
            self.toast.show_toast("Vui lòng chọn hoặc kéo thả file Excel trước!", "warn")
            return
        if wl and not os.path.isfile(wl):
            QMessageBox.warning(self, "Không tìm thấy file", f"File Work Log không tồn tại:\n{wl}")
            return
        if rw and not os.path.isfile(rw):
            QMessageBox.warning(self, "Không tìm thấy file", f"File Review không tồn tại:\n{rw}")
            return

        if self.loader_thread and self.loader_thread.isRunning():
            return

        self.btn_reload.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.status_bar.showMessage("Đang đọc dữ liệu & Excel Notes...")

        self.loader_thread = LoaderThread(wl, rw)
        self.loader_thread.progress.connect(self._on_load_progress)
        self.loader_thread.finished_ok.connect(self._on_load_done)
        self.loader_thread.failed.connect(self._on_load_fail)
        self.loader_thread.start()

    def _on_load_progress(self, pct: int, msg: str):
        self.progress.setValue(pct)
        self.status_bar.showMessage(msg)

    def _on_load_done(self, ds: LoadedDataset):
        self.dataset = ds
        self.btn_reload.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.progress.setVisible(False)
        self._isolated_annotator = None
        self.isolate_banner.setVisible(False)

        # Build Task chips
        self._updating_filters = True
        self._clear_task_chips()
        priority = {"OD": 0, "SSEG": 1, "PFE": 2}
        ordered_tasks = sorted(ds.tasks, key=lambda t: (priority.get(t, 99), t))
        for t in ordered_tasks:
            btn = QPushButton(t)
            btn.setObjectName("ChipBtn")
            btn.setCheckable(True)
            btn.setChecked(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.toggled.connect(self._on_task_chip_toggled)
            self.task_flow.addWidget(btn)
            self._task_chips[t] = btn

        # Build Annotator list
        self.list_anns.clear()
        for key, disp in ds.annotators.items():
            it = QListWidgetItem(disp)
            it.setData(Qt.ItemDataRole.UserRole, key)
            it.setData(Qt.ItemDataRole.UserRole + 1, disp)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked)
            self.list_anns.addItem(it)
        self._updating_filters = False

        # Auto-select latest complete Mon-Sun week
        if ds.max_date:
            ref = ds.max_date
            mon = ref - datetime.timedelta(days=ref.weekday())
            if ref.weekday() < 6:
                s_d = mon - datetime.timedelta(days=7)
                e_d = s_d + datetime.timedelta(days=6)
            else:
                s_d, e_d = mon, ref
            self.dt_start.blockSignals(True)
            self.dt_end.blockSignals(True)
            self.dt_start.setDate(QDate(s_d.year, s_d.month, s_d.day))
            self.dt_end.setDate(QDate(e_d.year, e_d.month, e_d.day))
            self.dt_start.blockSignals(False)
            self.dt_end.blockSignals(False)

        msg = (
            f"✓ Đã nạp {len(ds.annotators)} annotators · {len(ds.worklog_notes):,} Notes · "
            f"{len(ds.rework_records):,} dòng Review"
        )
        self.status_bar.showMessage(msg)
        self.toast.show_toast(msg, "success")
        self.refresh_report()

    def _on_load_fail(self, err: str):
        self.btn_reload.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.progress.setVisible(False)
        self.status_bar.showMessage("Lỗi khi nạp file Excel!")
        QMessageBox.critical(self, "Lỗi đọc file Excel", f"Không thể đọc file:\n{err}")

    # ------------------------------------------------------------------
    # FILTER INTERACTIONS
    # ------------------------------------------------------------------
    def _clear_task_chips(self):
        while self.task_flow.count():
            item = self.task_flow.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._task_chips.clear()

    def _on_option_changed(self, _val):
        self._save_current_settings()
        self.refresh_report()

    def _on_date_changed(self):
        self.refresh_report()

    def _apply_date_preset(self, mode: str):
        ref = self.dataset.max_date if (self.dataset and self.dataset.max_date) else datetime.date.today()
        mon = ref - datetime.timedelta(days=ref.weekday())
        if mode == "this_week":
            s, e = mon, mon + datetime.timedelta(days=6)
        elif mode == "last_week":
            s = mon - datetime.timedelta(days=7)
            e = s + datetime.timedelta(days=6)
        elif mode == "this_month":
            s = ref.replace(day=1)
            next_m = (s.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
            e = next_m - datetime.timedelta(days=1)
        else:
            s = self.dataset.min_date if (self.dataset and self.dataset.min_date) else ref
            e = ref

        self.dt_start.blockSignals(True)
        self.dt_end.blockSignals(True)
        self.dt_start.setDate(QDate(s.year, s.month, s.day))
        self.dt_end.setDate(QDate(e.year, e.month, e.day))
        self.dt_start.blockSignals(False)
        self.dt_end.blockSignals(False)
        self.refresh_report()

    def _on_task_chip_toggled(self, _checked):
        if not self._updating_filters:
            self.refresh_report()

    def _set_all_task_chips(self, checked: bool):
        self._updating_filters = True
        for btn in self._task_chips.values():
            btn.setChecked(checked)
        self._updating_filters = False
        self.refresh_report()

    def _get_period_bounds(self):
        qd_s, qd_e = self.dt_start.date(), self.dt_end.date()
        s_d = datetime.date(qd_s.year(), qd_s.month(), qd_s.day())
        e_d = datetime.date(qd_e.year(), qd_e.month(), qd_e.day())
        return (e_d, s_d) if s_d > e_d else (s_d, e_d)

    def _select_active_tasks(self):
        if not self.dataset:
            return
        s_d, e_d = self._get_period_bounds()
        rep_all = build_report(
            self.dataset, s_d, e_d, None, None,
            source_mode=self.seg_source.current_value(),
        )
        active = {r.task for r in rep_all.all_rows}
        self._updating_filters = True
        for t, btn in self._task_chips.items():
            btn.setChecked(t in active)
        self._updating_filters = False
        self.refresh_report()

    def _set_all_anns(self, checked: bool):
        self._isolated_annotator = None
        self.isolate_banner.setVisible(False)
        self._updating_filters = True
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.list_anns.count()):
            it = self.list_anns.item(i)
            if not it.isHidden():
                it.setCheckState(state)
        self._updating_filters = False
        self.refresh_report()

    def _select_active_annotators(self):
        if not self.dataset:
            return
        self._isolated_annotator = None
        self.isolate_banner.setVisible(False)
        s_d, e_d = self._get_period_bounds()
        rep_all = build_report(
            self.dataset, s_d, e_d, None, None,
            source_mode=self.seg_source.current_value(),
        )
        active = {r.annotator_key for r in rep_all.all_rows}
        self._updating_filters = True
        for i in range(self.list_anns.count()):
            it = self.list_anns.item(i)
            it.setCheckState(Qt.CheckState.Checked if it.data(Qt.ItemDataRole.UserRole) in active else Qt.CheckState.Unchecked)
        self._updating_filters = False
        self.refresh_report()

    def _isolate_single_annotator(self, item: QListWidgetItem):
        key = item.data(Qt.ItemDataRole.UserRole)
        disp = item.data(Qt.ItemDataRole.UserRole + 1)
        self._isolate_by_key(key, disp)

    def _isolate_by_key(self, key: str, disp: str):
        self._isolated_annotator = key
        self._updating_filters = True
        for i in range(self.list_anns.count()):
            it = self.list_anns.item(i)
            it.setCheckState(Qt.CheckState.Checked if it.data(Qt.ItemDataRole.UserRole) == key else Qt.CheckState.Unchecked)
        self._updating_filters = False
        self.lbl_isolate.setText(f"Đang xem riêng 1 người: {disp}")
        self.isolate_banner.setVisible(True)
        self.refresh_report()

    def _clear_isolation(self):
        self._set_all_anns(True)

    def _filter_ann_search(self, text: str):
        q = text.strip().lower()
        for i in range(self.list_anns.count()):
            it = self.list_anns.item(i)
            disp = str(it.data(Qt.ItemDataRole.UserRole + 1) or it.text()).lower()
            it.setHidden(bool(q) and q not in disp)

    def _on_ann_item_changed(self, _item):
        if not self._updating_filters:
            checked = self._get_checked_annotators()
            if len(checked) != 1:
                self._isolated_annotator = None
                self.isolate_banner.setVisible(False)
            self.refresh_report()

    def _get_checked_annotators(self) -> Set[str]:
        res = set()
        for i in range(self.list_anns.count()):
            it = self.list_anns.item(i)
            if it.checkState() == Qt.CheckState.Checked:
                res.add(it.data(Qt.ItemDataRole.UserRole))
        return res

    def _get_checked_tasks(self) -> Set[str]:
        return {t for t, btn in self._task_chips.items() if btn.isChecked()}

    # ------------------------------------------------------------------
    # REPORT COMPUTATION & RENDERING
    # ------------------------------------------------------------------
    def refresh_report(self):
        if not self.dataset:
            return
        s_d, e_d = self._get_period_bounds()
        src = self.seg_source.current_value()

        # Update sidebar live frame counts per annotator for this date range
        rep_period_all = build_report(self.dataset, s_d, e_d, None, None, source_mode=src)
        fr_by_ann = {}
        for r in rep_period_all.all_rows:
            fr_by_ann[r.annotator_key] = fr_by_ann.get(r.annotator_key, 0.0) + r.work_load
        self._updating_filters = True
        for i in range(self.list_anns.count()):
            it = self.list_anns.item(i)
            k = it.data(Qt.ItemDataRole.UserRole)
            disp = it.data(Qt.ItemDataRole.UserRole + 1)
            fr = fr_by_ann.get(k, 0.0)
            it.setText(f"{disp}   ({ _fmt_int(fr) } fr)" if fr > 0 else disp)
        self._updating_filters = False

        # Compute filtered report
        self.current_report = build_report(
            self.dataset,
            start_date=s_d,
            end_date=e_d,
            selected_annotators=self._get_checked_annotators(),
            selected_tasks=self._get_checked_tasks(),
            source_mode=src,
        )
        self.lbl_period_badge.setText(f"Kỳ báo cáo: {s_d.strftime('%Y.%m.%d')} ~ {e_d.strftime('%Y.%m.%d')}")

        self._render_kpis(self.current_report)
        self._render_template1(self.current_report)
        self._render_template2(self.current_report)
        self._render_reviewer_tab(self.current_report)
        self._render_ranking_tables(self.current_report)
        self._render_audit_table(self.current_report)
        self._populate_note_filter(self.current_report)
        self._render_daily_diffs_table()
        self._render_notes_table()

    def _render_kpis(self, rep: ReportResult):
        t = self.theme
        rows = rep.all_rows
        uniq_anns = len({r.annotator_key for r in rows})
        uniq_tasks = len({r.task for r in rows})
        tot_fr = sum(r.work_load for r in rows)
        tot_bb = sum(r.work_bbox for r in rows)
        tot_tm = sum(r.work_time for r in rows)
        tot_err = sum(r.err_first + r.err_second for r in rows)
        avg_fph = (tot_fr / tot_tm) if tot_tm > 0 else 0.0
        avg_bph = (tot_bb / tot_tm) if tot_tm > 0 else 0.0
        err_rate = (tot_err / tot_fr * 100.0) if tot_fr > 0 else 0.0
        real_diff_days = sum(1 for d in rep.daily_diffs if d.kind == "mismatch")
        pend_diff_days = sum(1 for d in rep.daily_diffs if d.kind == "pending_rework")
        days = (rep.end_date - rep.start_date).days + 1

        self.kpi_people.set_data(f"{uniq_anns} người", f"{len(rows)} dòng · {uniq_tasks} Task", t["accent"])
        self.kpi_frames.set_data(_fmt_int(tot_fr), f"TB {avg_fph:.2f} frame/h", t["text_primary"])
        self.kpi_bbox.set_data(_fmt_int(tot_bb), f"TB {avg_bph:.2f} bbox/h", "#8B5CF6" if t["name"] == "dark" else "#7C3AED")
        self.kpi_time.set_data(f"{_fmt_time(tot_tm)} h", f"Trong {days} ngày", "#38BDF8" if t["name"] == "dark" else "#0284C7")
        if real_diff_days > 0:
            extra = f" · {pend_diff_days} chưa có Review" if pend_diff_days else ""
            self.kpi_quality.set_data(
                f"{err_rate:.2f}%",
                f"⚠ {real_diff_days} ngày lệch{extra} (bấm xem)",
                t["warn"],
                "alert",
            )
        elif pend_diff_days > 0:
            self.kpi_quality.set_data(
                f"{err_rate:.2f}%",
                f"⏳ {pend_diff_days} ngày chưa có data Review",
                t["accent"],
                "clock",
            )
        else:
            self.kpi_quality.set_data(f"{err_rate:.2f}%", f"✓ Khớp 100% ({len(rep.filtered_notes)} notes)", t["success"], "shield")

    def _make_item(
        self,
        text: str,
        align=Qt.AlignmentFlag.AlignRight | Qt.AlignmentFlag.AlignVCenter,
        bold: bool = False,
        fg: Optional[str] = None,
        bg: Optional[str] = None,
        tooltip: str = "",
        user_data=None,
    ) -> QTableWidgetItem:
        it = QTableWidgetItem(text)
        it.setTextAlignment(int(align))
        if bold:
            f = QFont("Segoe UI", 9)
            f.setBold(True)
            it.setFont(f)
        if fg:
            it.setForeground(QBrush(QColor(fg)))
        if bg:
            it.setBackground(QBrush(QColor(bg)))
        if tooltip:
            it.setToolTip(tooltip)
        if user_data is not None:
            it.setData(Qt.ItemDataRole.UserRole, user_data)
        return it

    def _build_mismatch_tooltip(self, r, prefix: str = "") -> str:
        lines = []
        if prefix:
            lines.append(prefix)
        if r.has_mismatch or r.has_pending:
            lines.append(
                f"Đối chiếu tổng kỳ:\n"
                f"  • Frame: Work Log = {_fmt_int(r.wl_frames)} | Review = {_fmt_int(r.rw_frames)} (chênh {_fmt_int(r.diff_frame)})\n"
                f"  • BBOX:  Work Log = {_fmt_int(r.wl_bbox)} | Review = {_fmt_int(r.rw_bbox)} (chênh {_fmt_int(r.diff_bbox)})"
            )
            if r.real_diffs:
                lines.append("⚠ Block BỊ LỆCH data (so khớp theo Frame Start/End):")
                for d in r.real_diffs[:5]:
                    rw_dt_s = f" | Review {d.rw_date.strftime('%m/%d')}" if d.rw_date else ""
                    rng_s = f" | {d.frame_range_display}" if d.frame_range_display != "—" else ""
                    lines.append(
                        f"  • Work Log {d.date.strftime('%Y-%m-%d')}{rw_dt_s} [{d.task}{rng_s}]: "
                        f"Frame {_fmt_int(d.wl_frames)}/{_fmt_int(d.rw_frames)} | "
                        f"BBOX {_fmt_int(d.wl_bbox)}/{_fmt_int(d.rw_bbox)} — {d.reason}"
                    )
                if len(r.real_diffs) > 5:
                    lines.append(f"  ... và {len(r.real_diffs) - 5} block lệch khác")
            if r.pending_diffs:
                lines.append("⏳ Block CHƯA CÓ DATA bên Review:")
                for d in r.pending_diffs[:5]:
                    rng_s = f" | {d.frame_range_display}" if d.frame_range_display != "—" else ""
                    lines.append(
                        f"  • Work Log {d.date.strftime('%Y-%m-%d')} [{d.task}{rng_s}]: "
                        f"Work Log có {_fmt_int(d.wl_frames)} frame / {_fmt_int(d.wl_bbox)} bbox"
                    )
                if len(r.pending_diffs) > 5:
                    lines.append(f"  ... và {len(r.pending_diffs) - 5} block khác")
            lines.append(f"(Nhấp đúp dòng để xem chi tiết Frame Start/End của {r.annotator})")
        else:
            lines.append(f"✓ Khớp 100% giữa Work Log và Review (nhấp đúp để xem Excel Notes của {r.annotator})")
        return "\n".join(lines)

    def _row_status_style(self, r, hl: bool):
        """Return (bg_color, prefix_icon) based on 3-state status."""
        if not hl:
            return None, ""
        t = self.theme
        sk = r.status_kind
        if sk == "mismatch":
            return t["warn_soft"], "⚠ "
        if sk == "pending":
            return t["accent_soft"], "⏳ "
        return None, ""

    def _render_task_grouped_table(self, tbl: QTableWidget, delegate: ReportTableDelegate, groups, name_col_hdr: str = "Annotator", err_prefix: str = "Err", strip_rew_suffix: bool = False):
        tbl.clearSpans()
        t = self.theme
        headers = [
            "TASK",
            name_col_hdr,
            "Work\nLoad",
            "Work\nBBOX",
            "Work\nTime",
            "1st Err\n(BBOX)",
            "2nd Err\n(BBOX)",
            "1st Err\n(Frame)",
            "2nd Err\n(Frame)",
            "FRAME\n/ Hour",
            "BBOX\n/ Hour",
            "Err Rate\nBBOX (%)",
            "Err Rate\nFrame (%)",
            "TASK",
            "Avg\nFrame/h",
            "Avg\nBBOX/h",
            "Err Rate\nBBOX (%)",
            "Err Rate\nFrame (%)",
        ]
        total_rows = sum(len(g.rows) for g in groups)
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(total_rows)

        hl = self.chk_highlight.isChecked()
        task_avg_bg = "#F1F5F9" if t["name"] == "light" else "#162032"
        group_bottoms = set()
        cur = 0
        for grp in groups:
            start_r = cur
            n = len(grp.rows)
            group_bottoms.add(start_r + n - 1)
            disp_task = grp.task[:-4] if (strip_rew_suffix and grp.task.endswith("_REW")) else grp.task
            for idx, r in enumerate(grp.rows):
                row_i = cur + idx
                row_bg, prefix_ic = self._row_status_style(r, hl)
                kpi_bg = t["tbl_kpi_bg"]
                err_info = (
                    f"Worker (Lần 1 + Pangyo):\n"
                    f"  • Lỗi BBOX:  {_fmt_int(r.total_err_bbox)} (Lần 1: {_fmt_int(r.err_first_bbox)} + Pangyo: {_fmt_int(r.err_second_bbox)})\n"
                    f"  • Lỗi Frame: {_fmt_int(r.total_err_frame)} (Lần 1: {_fmt_int(r.err_first)} + Pangyo: {_fmt_int(r.err_second)})"
                )
                tip = self._build_mismatch_tooltip(r, prefix=err_info)
                has_any_issue = (r.has_mismatch or r.has_pending)

                if idx == 0:
                    tbl.setItem(row_i, 0, self._make_item(disp_task, Qt.AlignmentFlag.AlignCenter, True, t["task_badge_fg"], t["task_badge_bg"]))

                ann_lbl = f"{prefix_ic}{r.annotator}"
                tbl.setItem(row_i, 1, self._make_item(ann_lbl, Qt.AlignmentFlag.AlignCenter, False, None, row_bg, tip, user_data=(r.annotator, has_any_issue)))
                tbl.setItem(row_i, 2, self._make_item(_fmt_int(r.work_load), bg=row_bg, tooltip=tip))
                tbl.setItem(row_i, 3, self._make_item(_fmt_int(r.work_bbox), bg=row_bg, tooltip=tip))
                tbl.setItem(row_i, 4, self._make_item(_fmt_time(r.work_time), tooltip=tip))
                tbl.setItem(row_i, 5, self._make_item(_fmt_int(r.err_first_bbox), tooltip=tip))
                tbl.setItem(row_i, 6, self._make_item(_fmt_int(r.err_second_bbox), tooltip=tip))
                tbl.setItem(row_i, 7, self._make_item(_fmt_int(r.err_first), tooltip=tip))
                tbl.setItem(row_i, 8, self._make_item(_fmt_int(r.err_second), tooltip=tip))
                tbl.setItem(row_i, 9, self._make_item(_fmt_f2(r.frame_per_hour), bold=True, bg=kpi_bg))
                tbl.setItem(row_i, 10, self._make_item(_fmt_f2(r.bbox_per_hour), bold=True, bg=kpi_bg))
                tbl.setItem(row_i, 11, self._make_item(_fmt_f2(r.total_err_bbox_rate), bold=True, fg=t["danger"], tooltip=tip))
                tbl.setItem(row_i, 12, self._make_item(_fmt_f2(r.total_err_rate), bold=True, fg=t["danger"], tooltip=tip))

            avg_tip = (
                f"Chỉ số trung bình & tỉ lệ lỗi của toàn bộ nhóm TASK {disp_task} ({n} người):\n"
                f"  • Tổng Frame: {_fmt_int(grp.total_work_load)} | Tổng BBOX: {_fmt_int(grp.total_work_bbox)} | Tổng Time: {_fmt_time(grp.total_work_time)}h\n"
                f"  • Tổng lỗi BBOX: {_fmt_int(grp.total_err_bbox)} ({_fmt_f2(grp.total_task_err_bbox_rate)}%) | Tổng lỗi Frame: {_fmt_int(grp.total_err_frame)} ({_fmt_f2(grp.total_task_err_rate)}%)"
            )
            avg_align = Qt.AlignmentFlag.AlignCenter
            tbl.setItem(start_r, 13, self._make_item(disp_task, avg_align, True, t["task_badge_fg"], t["task_badge_bg"], tooltip=avg_tip))
            tbl.setItem(start_r, 14, self._make_item(_fmt_f2(grp.avg_frame_per_hour), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
            tbl.setItem(start_r, 15, self._make_item(_fmt_f2(grp.avg_bbox_per_hour), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
            tbl.setItem(start_r, 16, self._make_item(_fmt_f2(grp.total_task_err_bbox_rate), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
            tbl.setItem(start_r, 17, self._make_item(_fmt_f2(grp.total_task_err_rate), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))

            if n > 1:
                tbl.setSpan(start_r, 0, n, 1)
                for c in (13, 14, 15, 16, 17):
                    tbl.setSpan(start_r, c, n, 1)
            cur += n

        delegate.set_group_bottom_rows(group_bottoms)
        hdr = tbl.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(13, QHeaderView.ResizeMode.Fixed)
        tbl.setColumnWidth(0, 62)
        tbl.setColumnWidth(1, 90)
        tbl.setColumnWidth(13, 62)
        hdr.setFixedHeight(44)

    def _render_person_summary_table(self, tbl: QTableWidget, delegate: ReportTableDelegate, rows, name_col_hdr: str = "Annotator", err_prefix: str = "Err"):
        t = self.theme
        headers = [
            name_col_hdr,
            "Work\nLoad",
            "Work\nBBOX",
            "Work\nTime",
            "1st Err\n(BBOX)",
            "2nd Err\n(BBOX)",
            "1st Err\n(Frame)",
            "2nd Err\n(Frame)",
            "FRAME Per\nHour",
            "BBOX Per\nHour",
            "Err Rate\nBBOX (%)",
            "Err Rate\nFrame (%)",
        ]
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(len(rows))

        hl = self.chk_highlight.isChecked()

        for row_i, r in enumerate(rows):
            row_bg, prefix_ic = self._row_status_style(r, hl)
            kpi_bg = t["tbl_kpi_bg"]
            role_s = f"[Worker — {r.task}] Lỗi BBOX={_fmt_int(r.total_err_bbox)} ({_fmt_int(r.err_first_bbox)}+{_fmt_int(r.err_second_bbox)}) | Lỗi Frame={_fmt_int(r.total_err_frame)} ({_fmt_int(r.err_first)}+{_fmt_int(r.err_second)})"
            tip = self._build_mismatch_tooltip(r, prefix=role_s)
            lbl = f"{prefix_ic}{r.annotator}"
            has_any_issue = (r.has_mismatch or r.has_pending)

            tbl.setItem(row_i, 0, self._make_item(lbl, Qt.AlignmentFlag.AlignCenter, True, None, row_bg, tip, user_data=(r.annotator, has_any_issue)))
            tbl.setItem(row_i, 1, self._make_item(_fmt_int(r.work_load), bg=row_bg, tooltip=tip))
            tbl.setItem(row_i, 2, self._make_item(_fmt_int(r.work_bbox), bg=row_bg, tooltip=tip))
            tbl.setItem(row_i, 3, self._make_item(_fmt_time(r.work_time), tooltip=tip))
            tbl.setItem(row_i, 4, self._make_item(_fmt_int(r.err_first_bbox), tooltip=tip))
            tbl.setItem(row_i, 5, self._make_item(_fmt_int(r.err_second_bbox), tooltip=tip))
            tbl.setItem(row_i, 6, self._make_item(_fmt_int(r.err_first), tooltip=tip))
            tbl.setItem(row_i, 7, self._make_item(_fmt_int(r.err_second), tooltip=tip))
            tbl.setItem(row_i, 8, self._make_item(_fmt_f2(r.frame_per_hour), bold=True, bg=kpi_bg))
            tbl.setItem(row_i, 9, self._make_item(_fmt_f2(r.bbox_per_hour), bold=True, bg=kpi_bg))
            tbl.setItem(row_i, 10, self._make_item(_fmt_f2(r.total_err_bbox_rate), bold=True, fg=t["danger"], tooltip=tip))
            tbl.setItem(row_i, 11, self._make_item(_fmt_f2(r.total_err_rate), bold=True, fg=t["danger"], tooltip=tip))

        delegate.set_group_bottom_rows(set())
        hdr = tbl.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        tbl.setColumnWidth(0, 110)
        hdr.setFixedHeight(44)

    def _render_template1(self, rep: ReportResult):
        # Tab 1 (Task): Workers ONLY
        worker_groups = [g for g in rep.task_groups if not g.task.endswith("_REW")]
        self._render_task_grouped_table(self.tbl_t1, self.del_t1, worker_groups, name_col_hdr="Annotator", err_prefix="Err", strip_rew_suffix=False)

    def _render_template2(self, rep: ReportResult):
        # Tab 2 (Annotator): Workers ONLY
        worker_rows = [r for r in rep.annotator_rows if not r.is_reviewer]
        self._render_person_summary_table(self.tbl_t2, self.del_t2, worker_rows, name_col_hdr="Annotator", err_prefix="Err")

    def _render_reviewer_tab(self, rep: ReportResult):
        # Tab 3 (Reviewer): Exact sample template grouped by INSPECTOR -> Annotator
        tbl = self.tbl_rev
        tbl.clearSpans()
        t = self.theme
        headers = [
            "INSPECTOR",
            "Annotator",
            "Work\nLoad",
            "Work\nBBOX",
            "Work\nTime",
            "1st Insp.\nError (Fr)",
            "2nd Insp.\nError (Fr)",
            "FRAME\n/ Hour",
            "BBOX\n/ Hour",
            "Err Rate\nBBOX (%)",
            "Err Rate\nFrame (%)",
            "1st Inspector\nRecall (%)",
        ]
        groups = rep.inspector_groups
        data_rows_cnt = sum(len(g.rows) for g in groups)
        total_rows = data_rows_cnt + (1 if data_rows_cnt > 0 else 0)
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(total_rows)

        kpi_bg = "#F1F5F9" if t["name"] == "light" else "#162032"
        tot_bg = "#E2E8F0" if t["name"] == "light" else "#1E293B"
        group_bottoms = set()
        cur = 0
        for grp in groups:
            start_r = cur
            n = len(grp.rows)
            group_bottoms.add(start_r + n - 1)
            for idx, r in enumerate(grp.rows):
                row_i = cur + idx
                tip = (
                    f"Reviewer {grp.inspector} chấm bài của {r.annotator}:\n"
                    f"  • Lỗi Lần 1 (Review): {_fmt_int(r.err_first)} Frame | {_fmt_int(r.err_first_bbox)} BBOX\n"
                    f"  • Lỗi Lần 2 (Pangyo): {_fmt_int(r.err_second)} Frame | {_fmt_int(r.err_second_bbox)} BBOX"
                )
                if idx == 0:
                    tbl.setItem(row_i, 0, self._make_item(grp.inspector, Qt.AlignmentFlag.AlignCenter, True, t["task_badge_fg"], t["task_badge_bg"]))
                tbl.setItem(row_i, 1, self._make_item(r.annotator, Qt.AlignmentFlag.AlignCenter, False, None, None, tip, user_data=(r.annotator, False)))
                tbl.setItem(row_i, 2, self._make_item(_fmt_int(r.work_load), tooltip=tip))
                tbl.setItem(row_i, 3, self._make_item(_fmt_int(r.work_bbox), tooltip=tip))
                tbl.setItem(row_i, 4, self._make_item(_fmt_time(r.work_time), tooltip=tip))
                tbl.setItem(row_i, 5, self._make_item(_fmt_int(r.err_first), tooltip=tip))
                tbl.setItem(row_i, 6, self._make_item(_fmt_int(r.err_second), tooltip=tip))

            sum_tip = (
                f"Tổng hợp Reviewer {grp.inspector} ({n} Annotators):\n"
                f"  • Tổng Load: {_fmt_int(grp.total_work_load)} Frame | {_fmt_int(grp.total_work_bbox)} BBOX | {_fmt_time(grp.total_work_time)}h\n"
                f"  • Lỗi Pangyo: {_fmt_int(grp.total_err_second)} Frame ({_fmt_f2(grp.err_rate_frame)}%) | {_fmt_int(grp.total_err_second_bbox)} BBOX ({_fmt_f2(grp.err_rate_bbox)}%)\n"
                f"  • Recall Lần 1: {_fmt_int(grp.total_err_first)} / ({_fmt_int(grp.total_err_first)} + {_fmt_int(grp.total_err_second)}) = {grp.first_inspector_recall:.0f}%"
            )
            c_align = Qt.AlignmentFlag.AlignCenter
            tbl.setItem(start_r, 7, self._make_item(_fmt_f2(grp.frame_per_hour), c_align, True, t["danger"], bg=kpi_bg, tooltip=sum_tip))
            tbl.setItem(start_r, 8, self._make_item(_fmt_f2(grp.bbox_per_hour), c_align, True, t["danger"], bg=kpi_bg, tooltip=sum_tip))
            tbl.setItem(start_r, 9, self._make_item(_fmt_f2(grp.err_rate_bbox), c_align, True, None, bg=kpi_bg, tooltip=sum_tip))
            tbl.setItem(start_r, 10, self._make_item(_fmt_f2(grp.err_rate_frame), c_align, True, None, bg=kpi_bg, tooltip=sum_tip))
            tbl.setItem(start_r, 11, self._make_item(f"{grp.first_inspector_recall:.0f}%", c_align, True, t["accent_text"], bg=kpi_bg, tooltip=sum_tip))

            if n > 1:
                tbl.setSpan(start_r, 0, n, 1)
                for c in (7, 8, 9, 10, 11):
                    tbl.setSpan(start_r, c, n, 1)
            cur += n

        # Grand Total footer row
        if data_rows_cnt > 0:
            tr = data_rows_cnt
            tot_load = sum(g.total_work_load for g in groups)
            tot_bb = sum(g.total_work_bbox for g in groups)
            tot_tm = sum(g.total_work_time for g in groups)
            tot_e1 = sum(g.total_err_first for g in groups)
            tot_e2 = sum(g.total_err_second for g in groups)
            tot_e2_bb = sum(g.total_err_second_bbox for g in groups)
            tot_fph = (tot_load / tot_tm) if tot_tm > 0 else 0.0
            tot_bph = (tot_bb / tot_tm) if tot_tm > 0 else 0.0
            tot_er_fr = (tot_e2 / tot_load * 100.0) if tot_load > 0 else 0.0
            tot_er_bb = (tot_e2_bb / tot_bb * 100.0) if tot_bb > 0 else 0.0
            tot_rec = (tot_e1 / (tot_e1 + tot_e2) * 100.0) if (tot_e1 + tot_e2) > 0 else 100.0

            tbl.setItem(tr, 0, self._make_item("TỔNG", Qt.AlignmentFlag.AlignCenter, True, bg=tot_bg))
            tbl.setItem(tr, 1, self._make_item(f"{len(groups)} Reviewer", Qt.AlignmentFlag.AlignCenter, True, bg=tot_bg))
            tbl.setItem(tr, 2, self._make_item(_fmt_int(tot_load), bold=True, bg=tot_bg))
            tbl.setItem(tr, 3, self._make_item(_fmt_int(tot_bb), bold=True, bg=tot_bg))
            tbl.setItem(tr, 4, self._make_item(f"{tot_tm:.2f}", bold=True, bg=tot_bg))
            tbl.setItem(tr, 5, self._make_item(_fmt_int(tot_e1), bold=True, bg=tot_bg))
            tbl.setItem(tr, 6, self._make_item(_fmt_int(tot_e2), bold=True, bg=tot_bg))
            tbl.setItem(tr, 7, self._make_item(_fmt_f2(tot_fph), Qt.AlignmentFlag.AlignCenter, True, t["danger"], bg=tot_bg))
            tbl.setItem(tr, 8, self._make_item(_fmt_f2(tot_bph), Qt.AlignmentFlag.AlignCenter, True, t["danger"], bg=tot_bg))
            tbl.setItem(tr, 9, self._make_item(_fmt_f2(tot_er_bb), Qt.AlignmentFlag.AlignCenter, True, bg=tot_bg))
            tbl.setItem(tr, 10, self._make_item(_fmt_f2(tot_er_fr), Qt.AlignmentFlag.AlignCenter, True, bg=tot_bg))
            tbl.setItem(tr, 11, self._make_item(f"{tot_rec:.0f}%", Qt.AlignmentFlag.AlignCenter, True, t["accent_text"], bg=tot_bg))

        self.del_rev.set_group_bottom_rows(group_bottoms)
        hdr = tbl.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        tbl.setColumnWidth(0, 96)
        tbl.setColumnWidth(1, 96)
        hdr.setFixedHeight(44)

    def _render_ranking_tables(self, rep: ReportResult):
        t = self.theme
        by_bbox = (self.seg_rank_by.current_value() == "bbox")

        def worker_sort_key(r):
            primary = r.total_err_bbox_rate if by_bbox else r.total_err_rate
            secondary = r.total_err_rate if by_bbox else r.total_err_bbox_rate
            return (primary, secondary, -r.work_load, r.annotator.upper())

        def rev_sort_key(g):
            primary = g.err_rate_bbox if by_bbox else g.err_rate_frame
            secondary = g.err_rate_frame if by_bbox else g.err_rate_bbox
            return (primary, secondary, -g.total_work_load, g.inspector.upper())

        workers = sorted(
            [r for r in rep.annotator_rows if not r.is_reviewer and (r.work_load > 0 or r.work_bbox > 0)],
            key=worker_sort_key,
        )
        reviewers = sorted(
            [g for g in rep.inspector_groups if g.total_work_load > 0 or g.total_work_bbox > 0],
            key=rev_sort_key,
        )

        self.lbl_rank_worker.setText(f"🥇 Xếp hạng Annotator / Worker ({len(workers)} người — Tỉ lệ lỗi thấp nhất)")
        self.lbl_rank_reviewer.setText(f"🥇 Xếp hạng Người Review Lần 1 ({len(reviewers)} người — Tỉ lệ lỗi Pangyo thấp nhất)")

        # Populate Worker Ranking
        w_tbl = self.tbl_rank_worker
        w_headers = [
            "Hạng", "Annotator", "Work\nLoad", "Work\nBBOX", "Work\nTime",
            "Err\nFrame", "Err\nBBOX", "Err Rate\nFrame (%)", "Err Rate\nBBOX (%)",
            "FRAME\n/ Hour", "BBOX\n/ Hour",
        ]
        w_tbl.setColumnCount(len(w_headers))
        w_tbl.setHorizontalHeaderLabels(w_headers)
        w_tbl.setRowCount(len(workers))
        medals = {1: "🥇 1", 2: "🥈 2", 3: "🥉 3"}
        for idx, r in enumerate(workers):
            rank_num = idx + 1
            rank_lbl = medals.get(rank_num, str(rank_num))
            is_top3 = (rank_num <= 3)
            bg = t["success_soft"] if is_top3 else None
            rate_fg = t["success"] if is_top3 else t["danger"]
            has_any_issue = (r.has_mismatch or r.has_pending)
            tip = (
                f"Hạng #{rank_num} — {r.annotator} ({r.task}):\n"
                f"  • Sản lượng: {_fmt_int(r.work_load)} Frame | {_fmt_int(r.work_bbox)} BBOX | {_fmt_time(r.work_time)}h\n"
                f"  • Số lỗi: {_fmt_int(r.total_err_frame)} Frame ({_fmt_f2(r.total_err_rate)}%) | "
                f"{_fmt_int(r.total_err_bbox)} BBOX ({_fmt_f2(r.total_err_bbox_rate)}%)"
            )
            w_tbl.setItem(idx, 0, self._make_item(rank_lbl, Qt.AlignmentFlag.AlignCenter, True, rate_fg if is_top3 else None, bg, tip))
            w_tbl.setItem(idx, 1, self._make_item(r.annotator, Qt.AlignmentFlag.AlignCenter, True, None, bg, tip, user_data=(r.annotator, has_any_issue)))
            w_tbl.setItem(idx, 2, self._make_item(_fmt_int(r.work_load), bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 3, self._make_item(_fmt_int(r.work_bbox), bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 4, self._make_item(_fmt_time(r.work_time), bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 5, self._make_item(_fmt_int(r.total_err_frame), bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 6, self._make_item(_fmt_int(r.total_err_bbox), bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 7, self._make_item(_fmt_f2(r.total_err_rate), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 8, self._make_item(_fmt_f2(r.total_err_bbox_rate), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 9, self._make_item(_fmt_f2(r.frame_per_hour), bold=True, bg=bg or t["tbl_kpi_bg"], tooltip=tip))
            w_tbl.setItem(idx, 10, self._make_item(_fmt_f2(r.bbox_per_hour), bold=True, bg=bg or t["tbl_kpi_bg"], tooltip=tip))
        wh = w_tbl.horizontalHeader()
        wh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        wh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        wh.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        w_tbl.setColumnWidth(0, 56)
        w_tbl.setColumnWidth(1, 86)
        wh.setFixedHeight(42)

        # Populate Reviewer Lần 1 Ranking
        r_tbl = self.tbl_rank_reviewer
        r_headers = [
            "Hạng", "Reviewer", "Work\nLoad", "Work\nBBOX", "Work\nTime",
            "Pangyo\nFrame", "Pangyo\nBBOX", "Err Rate\nFrame (%)", "Err Rate\nBBOX (%)",
            "1st Insp.\nRecall (%)", "FRAME\n/ Hour",
        ]
        r_tbl.setColumnCount(len(r_headers))
        r_tbl.setHorizontalHeaderLabels(r_headers)
        r_tbl.setRowCount(len(reviewers))
        for idx, g in enumerate(reviewers):
            rank_num = idx + 1
            rank_lbl = medals.get(rank_num, str(rank_num))
            is_top3 = (rank_num <= 3)
            bg = t["success_soft"] if is_top3 else None
            rate_fg = t["success"] if is_top3 else t["danger"]
            tip = (
                f"Hạng #{rank_num} — Reviewer {g.inspector}:\n"
                f"  • Sản lượng: {_fmt_int(g.total_work_load)} Frame | {_fmt_int(g.total_work_bbox)} BBOX | {_fmt_time(g.total_work_time)}h\n"
                f"  • Lỗi Pangyo: {_fmt_int(g.total_err_second)} Frame ({_fmt_f2(g.err_rate_frame)}%) | {_fmt_int(g.total_err_second_bbox)} BBOX ({_fmt_f2(g.err_rate_bbox)}%)\n"
                f"  • Recall Lần 1: {g.first_inspector_recall:.0f}%"
            )
            r_tbl.setItem(idx, 0, self._make_item(rank_lbl, Qt.AlignmentFlag.AlignCenter, True, rate_fg if is_top3 else None, bg, tip))
            r_tbl.setItem(idx, 1, self._make_item(g.inspector, Qt.AlignmentFlag.AlignCenter, True, None, bg, tip))
            r_tbl.setItem(idx, 2, self._make_item(_fmt_int(g.total_work_load), bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 3, self._make_item(_fmt_int(g.total_work_bbox), bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 4, self._make_item(_fmt_time(g.total_work_time), bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 5, self._make_item(_fmt_int(g.total_err_second), bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 6, self._make_item(_fmt_int(g.total_err_second_bbox), bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 7, self._make_item(_fmt_f2(g.err_rate_frame), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 8, self._make_item(_fmt_f2(g.err_rate_bbox), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 9, self._make_item(f"{g.first_inspector_recall:.0f}%", Qt.AlignmentFlag.AlignCenter, True, t["accent_text"], bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 10, self._make_item(_fmt_f2(g.frame_per_hour), bold=True, bg=bg or t["tbl_kpi_bg"], tooltip=tip))
        rh = r_tbl.horizontalHeader()
        rh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        rh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        rh.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        r_tbl.setColumnWidth(0, 56)
        r_tbl.setColumnWidth(1, 86)
        rh.setFixedHeight(42)

    def _render_audit_table(self, rep: ReportResult):
        tbl = self.tbl_audit
        t = self.theme
        headers = [
            "TASK", "Annotator",
            "Work Log (Frame)", "Review (Frame)", "Lệch Frame",
            "Work Log (BBOX)", "Review (BBOX)", "Lệch BBOX",
            "Trạng thái", "Chi tiết ngày"
        ]
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(len(rep.all_rows))

        real_cnt = 0
        pend_cnt = 0
        for row_i, r in enumerate(rep.all_rows):
            sk = r.status_kind
            if sk == "mismatch":
                real_cnt += 1
                bg = t["warn_soft"]
                st_txt = "⚠ LỆCH DATA"
                st_fg = t["warn"]
                st_bg = t["warn_soft"]
            elif sk == "pending":
                pend_cnt += 1
                bg = t["accent_soft"]
                st_txt = "⏳ Chưa có data"
                st_fg = t["accent_text"]
                st_bg = t["accent_soft"]
            else:
                bg = None
                st_txt = "✓ Khớp 100%"
                st_fg = t["success"]
                st_bg = t["success_soft"]

            dates_str = r.status_dates_summary or "—"
            tip = self._build_mismatch_tooltip(r)
            has_any_issue = (sk != "ok")

            tbl.setItem(row_i, 0, self._make_item(r.task, Qt.AlignmentFlag.AlignCenter, True, bg=bg, tooltip=tip))
            tbl.setItem(row_i, 1, self._make_item(r.annotator, Qt.AlignmentFlag.AlignCenter, True, bg=bg, tooltip=tip, user_data=(r.annotator, has_any_issue)))
            tbl.setItem(row_i, 2, self._make_item(_fmt_int(r.wl_frames), bg=bg))
            tbl.setItem(row_i, 3, self._make_item(_fmt_int(r.rw_frames), bg=bg))
            tbl.setItem(row_i, 4, self._make_item(_fmt_int(r.diff_frame), bold=has_any_issue, fg=st_fg if has_any_issue else None, bg=bg))
            tbl.setItem(row_i, 5, self._make_item(_fmt_int(r.wl_bbox), bg=bg))
            tbl.setItem(row_i, 6, self._make_item(_fmt_int(r.rw_bbox), bg=bg))
            tbl.setItem(row_i, 7, self._make_item(_fmt_int(r.diff_bbox), bold=has_any_issue, fg=st_fg if has_any_issue else None, bg=bg))
            tbl.setItem(row_i, 8, self._make_item(st_txt, Qt.AlignmentFlag.AlignCenter, True, fg=st_fg, bg=st_bg))
            tbl.setItem(row_i, 9, self._make_item(dates_str, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, bold=has_any_issue, fg=st_fg if has_any_issue else t["text_muted"], bg=bg, tooltip=tip))

        tbl.resizeColumnsToContents()
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.horizontalHeader().setFixedHeight(36)
        summary_parts = []
        if real_cnt:
            summary_parts.append(f"⚠ {real_cnt} dòng lệch data")
        if pend_cnt:
            summary_parts.append(f"⏳ {pend_cnt} dòng chưa có data Review")
        summary_str = " · ".join(summary_parts) if summary_parts else "✓ Khớp 100%"
        self.lbl_audit_hdr.setText(
            f"Đối chiếu Work Log (+ Note) vs Review (#IMAGE): {len(rep.all_rows)} dòng ({summary_str}) "
            f"— bấm 1 dòng để xem chi tiết bên dưới"
        )

    def _populate_note_filter(self, rep: ReportResult):
        prev = self.cb_note_ann.currentData()
        self.cb_note_ann.blockSignals(True)
        self.cb_note_ann.clear()
        self.cb_note_ann.addItem("Tất cả annotator", "")
        names = sorted({n.annotator_display for n in rep.filtered_notes} | {d.annotator for d in rep.daily_diffs})
        for nm in names:
            self.cb_note_ann.addItem(nm, nm)
        if prev and prev in names:
            self.cb_note_ann.setCurrentIndex(self.cb_note_ann.findData(prev))
        self.cb_note_ann.blockSignals(False)

    def _on_detail_ann_changed(self):
        self._render_daily_diffs_table()
        self._render_notes_table()

    def _render_daily_diffs_table(self):
        if not self.current_report:
            return
        t = self.theme
        sel_ann = self.cb_note_ann.currentData() or ""
        diffs = [d for d in self.current_report.daily_diffs if not sel_ann or d.annotator == sel_ann]
        real_c = sum(1 for d in diffs if d.kind == "mismatch")
        pend_c = sum(1 for d in diffs if d.kind == "pending_rework")

        tbl = self.tbl_daily_diffs
        headers = [
            "Loại", "Annotator", "Ngày Work Log", "Ngày Review", "TASK", "Frame Start ~ End",
            "Work Log (Frame)", "Review (Frame)", "Lệch Frame",
            "Work Log (BBOX)", "Review (BBOX)", "Lệch BBOX",
            "Ghi chú / Nguyên nhân"
        ]
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(len(diffs))

        for row_i, d in enumerate(diffs):
            is_pend = d.is_pending
            bg = t["accent_soft"] if is_pend else t["warn_soft"]
            accent_fg = t["accent_text"] if is_pend else t["warn"]
            date_fg = t["accent_text"] if is_pend else t["danger"]
            kind_lbl = "⏳ Chưa có data" if is_pend else "⚠ LỆCH DATA"
            rw_dt_str = d.rw_date.strftime("%Y-%m-%d") if d.rw_date else "—"
            tip = "Nhấp đúp dòng này để xem Excel Notes của ngày này"

            tbl.setItem(row_i, 0, self._make_item(kind_lbl, Qt.AlignmentFlag.AlignCenter, True, fg=accent_fg, bg=bg, tooltip=tip, user_data=(d.annotator, d.date)))
            tbl.setItem(row_i, 1, self._make_item(d.annotator, Qt.AlignmentFlag.AlignCenter, True, bg=bg, tooltip=tip))
            tbl.setItem(row_i, 2, self._make_item(d.date.strftime("%Y-%m-%d"), Qt.AlignmentFlag.AlignCenter, True, fg=date_fg, bg=bg, tooltip=tip))
            tbl.setItem(row_i, 3, self._make_item(rw_dt_str, Qt.AlignmentFlag.AlignCenter, bg=bg, tooltip=tip))
            tbl.setItem(row_i, 4, self._make_item(d.task, Qt.AlignmentFlag.AlignCenter, True, bg=bg))
            tbl.setItem(row_i, 5, self._make_item(d.frame_range_display, Qt.AlignmentFlag.AlignCenter, bold=True, fg=date_fg if "≠" in d.frame_range_display else None, bg=bg, tooltip=tip))
            tbl.setItem(row_i, 6, self._make_item(_fmt_int(d.wl_frames), bg=bg))
            tbl.setItem(row_i, 7, self._make_item(_fmt_int(d.rw_frames), bg=bg))
            tbl.setItem(row_i, 8, self._make_item(_fmt_int(d.diff_frame), bold=True, fg=accent_fg, bg=bg))
            tbl.setItem(row_i, 9, self._make_item(_fmt_int(d.wl_bbox), bg=bg))
            tbl.setItem(row_i, 10, self._make_item(_fmt_int(d.rw_bbox), bg=bg))
            tbl.setItem(row_i, 11, self._make_item(_fmt_int(d.diff_bbox), bold=True, fg=accent_fg, bg=bg))
            tbl.setItem(row_i, 12, self._make_item(d.reason, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, bold=True, fg=accent_fg, bg=bg))

        tbl.resizeColumnsToContents()
        tbl.horizontalHeader().setStretchLastSection(True)
        tbl.horizontalHeader().setFixedHeight(34)
        self.audit_subtabs.setTabText(0, f"Đối chiếu Start/End ({real_c} lệch · {pend_c} chưa có Review)")
        suffix = f" của {sel_ann}" if sel_ann else ""
        self.lbl_detail_hdr.setText(f"Chi tiết đối chiếu theo Frame Start/End & Excel Notes{suffix}")

    def _render_notes_table(self, highlight_date: Optional[datetime.date] = None):
        if not self.current_report:
            return
        t = self.theme
        sel_ann = self.cb_note_ann.currentData() or ""
        notes = [n for n in self.current_report.filtered_notes if not sel_ann or n.annotator_display == sel_ann]
        real_pairs = {(d.annotator, d.date) for d in self.current_report.daily_diffs if d.kind == "mismatch"}
        pend_pairs = {(d.annotator, d.date) for d in self.current_report.daily_diffs if d.kind == "pending_rework"}

        ntbl = self.tbl_notes
        headers = ["Annotator", "Ngày", "TASK", "Work Log (Frame)", "Work Log (BBOX)", "Nội dung Note gốc"]
        ntbl.setColumnCount(len(headers))
        ntbl.setHorizontalHeaderLabels(headers)
        ntbl.setRowCount(len(notes))

        first_hl_row = None
        for row_i, n in enumerate(notes):
            pair = (n.annotator_display, n.date)
            is_real = pair in real_pairs
            is_pend = pair in pend_pairs and not is_real
            bg = t["warn_soft"] if is_real else (t["accent_soft"] if is_pend else None)
            dt_fg = t["danger"] if is_real else (t["accent_text"] if is_pend else None)
            if highlight_date and n.date == highlight_date and first_hl_row is None:
                first_hl_row = row_i
            ntbl.setItem(row_i, 0, self._make_item(n.annotator_display, Qt.AlignmentFlag.AlignCenter, True, bg=bg))
            ntbl.setItem(row_i, 1, self._make_item(n.date.strftime("%Y-%m-%d"), Qt.AlignmentFlag.AlignCenter, bold=(is_real or is_pend), fg=dt_fg, bg=bg))
            ntbl.setItem(row_i, 2, self._make_item(n.task, Qt.AlignmentFlag.AlignCenter, bg=bg))
            ntbl.setItem(row_i, 3, self._make_item(_fmt_int(n.frames_in_cell), bg=bg))
            bb_str = _fmt_int(n.parsed_bbox) if n.parsed_bbox is not None else "(trống)"
            bb_fg = t["accent_text"] if n.parsed_bbox is not None else t["danger"]
            ntbl.setItem(row_i, 4, self._make_item(bb_str, bold=True, fg=bb_fg, bg=bg))
            ntbl.setItem(row_i, 5, self._make_item(n.raw_note.replace("\n", " | "), Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, bg=bg))

        ntbl.resizeColumnsToContents()
        ntbl.horizontalHeader().setStretchLastSection(True)
        ntbl.horizontalHeader().setFixedHeight(34)
        self.audit_subtabs.setTabText(1, f"Excel Notes Work Log ({len(notes)})")
        if first_hl_row is not None:
            ntbl.selectRow(first_hl_row)
            ntbl.scrollToItem(ntbl.item(first_hl_row, 0))

    # ------------------------------------------------------------------
    # DRILL-DOWN CLICK HANDLERS
    # ------------------------------------------------------------------
    def _jump_to_audit_for(self, ann_name: str, has_mismatch: bool):
        if not ann_name:
            return
        idx = self.cb_note_ann.findData(ann_name)
        if idx >= 0:
            self.cb_note_ann.setCurrentIndex(idx)
        self.tabs.setCurrentIndex(4)
        self.audit_subtabs.setCurrentIndex(0 if has_mismatch else 1)

    def _on_t1_double_click(self, row: int, _col: int):
        it = self.tbl_t1.item(row, 1)
        if it and it.data(Qt.ItemDataRole.UserRole):
            ann, has_m = it.data(Qt.ItemDataRole.UserRole)
            self._jump_to_audit_for(ann, has_m)

    def _on_t2_double_click(self, row: int, _col: int):
        it = self.tbl_t2.item(row, 0)
        if it and it.data(Qt.ItemDataRole.UserRole):
            ann, has_m = it.data(Qt.ItemDataRole.UserRole)
            self._jump_to_audit_for(ann, has_m)

    def _on_rev_tot_double_click(self, row: int, _col: int):
        it = self.tbl_rev_tot.item(row, 0)
        if it and it.data(Qt.ItemDataRole.UserRole):
            ann, has_m = it.data(Qt.ItemDataRole.UserRole)
            self._jump_to_audit_for(ann, has_m)

    def _on_rank_double_click(self, tbl: QTableWidget, row: int):
        it = tbl.item(row, 1)
        if it and it.data(Qt.ItemDataRole.UserRole):
            ann, has_m = it.data(Qt.ItemDataRole.UserRole)
            self._jump_to_audit_for(ann, has_m)

    def _on_audit_row_clicked(self, row: int, _col: int):
        it = self.tbl_audit.item(row, 1)
        if it and it.data(Qt.ItemDataRole.UserRole):
            ann, has_m = it.data(Qt.ItemDataRole.UserRole)
            idx = self.cb_note_ann.findData(ann)
            if idx >= 0:
                self.cb_note_ann.setCurrentIndex(idx)
            self.audit_subtabs.setCurrentIndex(0 if has_m else 1)

    def _on_daily_diff_double_click(self, row: int, _col: int):
        it = self.tbl_daily_diffs.item(row, 0)
        if it and it.data(Qt.ItemDataRole.UserRole):
            ann, dt = it.data(Qt.ItemDataRole.UserRole)
            idx = self.cb_note_ann.findData(ann)
            if idx >= 0:
                self.cb_note_ann.setCurrentIndex(idx)
            self.audit_subtabs.setCurrentIndex(1)
            self._render_notes_table(highlight_date=dt)

    # ------------------------------------------------------------------
    # COPY & EXPORT
    # ------------------------------------------------------------------
    def copy_current_table(self):
        idx = self.tabs.currentIndex()
        if idx == 0:
            tbl = self.tbl_t1
        elif idx == 1:
            tbl = self.tbl_t2
        elif idx == 2:
            tbl = self.tbl_rev
        elif idx == 3:
            tbl = self.tbl_rank_reviewer if self.tbl_rank_reviewer.hasFocus() else self.tbl_rank_worker
        else:
            if self.audit_subtabs.currentIndex() == 0 and self.tbl_daily_diffs.rowCount() > 0:
                tbl = self.tbl_daily_diffs
            else:
                tbl = self.tbl_audit
        if tbl.rowCount() == 0:
            self.toast.show_toast("Không có dữ liệu để copy.", "warn")
            return
        keep_cols = [c for c in range(tbl.columnCount()) if tbl.horizontalHeaderItem(c) and tbl.horizontalHeaderItem(c).text().strip()]
        lines = ["\t".join(tbl.horizontalHeaderItem(c).text().replace("\n", " ") for c in keep_cols)]
        for r in range(tbl.rowCount()):
            vals = []
            for c in keep_cols:
                it = tbl.item(r, c)
                txt = it.text().replace("⚠ ", "").replace("⏳ ", "").replace("\n", " ") if it else ""
                vals.append(txt)
            lines.append("\t".join(vals))
        QApplication.clipboard().setText("\n".join(lines))
        self.toast.show_toast("✓ Đã copy bảng vào Clipboard (dán vào Excel bằng Ctrl+V)", "success")

    def export_excel(self):
        if not self.current_report or not self.current_report.all_rows:
            self.toast.show_toast("Chưa có dữ liệu trong kỳ / bộ lọc đang chọn!", "warn")
            return
        inc_t1 = self.chk_exp_t1.isChecked()
        inc_t2 = self.chk_exp_t2.isChecked()
        inc_au = self.chk_exp_audit.isChecked()
        if not (inc_t1 or inc_t2 or inc_au):
            self.toast.show_toast("Hãy tick ít nhất 1 sheet để xuất Excel!", "warn")
            return

        rep = self.current_report
        wl_p = self.card_wl.path() or ""
        default_dir = _app_base_dir() if ("onedrive" in wl_p.lower() or not wl_p) else os.path.dirname(wl_p)
        default_name = f"Report_{rep.start_date.strftime('%Y.%m.%d')}_to_{rep.end_date.strftime('%Y.%m.%d')}.xlsx"
        out_path, _ = QFileDialog.getSaveFileName(
            self,
            "Lưu báo cáo Excel",
            os.path.join(default_dir, default_name),
            "Excel Workbook (*.xlsx)",
        )
        if not out_path:
            return
        if not out_path.lower().endswith(".xlsx"):
            out_path += ".xlsx"

        try:
            export_report_to_excel(
                rep,
                out_path,
                include_template1=inc_t1,
                include_template2=inc_t2,
                include_audit=inc_au,
                highlight_mismatch=self.chk_highlight.isChecked(),
            )
            self.toast.show_toast(f"✓ Đã xuất Excel: {os.path.basename(out_path)}", "success", 4500)
            self.status_bar.showMessage(f"Đã xuất thành công: {out_path}")
            ans = QMessageBox.question(
                self,
                "Xuất Excel thành công",
                f"Đã lưu báo cáo tại:\n{out_path}\n\nMở file Excel ngay bây giờ?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ans == QMessageBox.StandardButton.Yes:
                os.startfile(out_path)
        except Exception as exc:
            QMessageBox.critical(self, "Lỗi xuất Excel", f"Không thể lưu file:\n{exc}")


def main():
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    win = MainWindow()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
