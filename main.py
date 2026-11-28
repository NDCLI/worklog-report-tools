import datetime
import json
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import replace
from typing import Dict, List, Optional, Set, Tuple

from PyQt6.QtCore import QDate, QSignalBlocker, QSize, Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QBrush, QColor, QFont, QFontMetrics, QIcon, QKeySequence, QPainter, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QCheckBox,
    QComboBox,
    QDateEdit,
    QDialog,
    QFileDialog,
    QFormLayout,
    QFrame,
    QGridLayout,
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
    QStackedWidget,
    QStatusBar,
    QTableWidget,
    QTableWidgetItem,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from data_engine import (
    FIRST_REVIEWER_KEYS,
    DailyDiffRecord,
    InspectorSummary,
    LoadedDataset,
    ReportResult,
    build_report,
    display_annotator_name,
    load_dataset,
    normalize_annotator_key,
    unique_inspector_names,
)
from excel_exporter import export_report_to_excel
from sharepoint_client import (
    SharePointDownloadThread,
    SharePointLoginDialog,
    clear_sp_session,
    extract_sp_site_url,
    has_saved_sp_session,
)
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
from updater import (
    ReleaseInfo,
    UpdateBanner,
    UpdateCheckThread,
    UpdateDialog,
    cleanup_old_update_files,
)


APP_NAME = "Annotator & Inspector Report"
APP_VERSION = "2.12"

DEFAULT_SP_WL_URL = (
    "https://hanwhaworkplace.sharepoint.com/:x:/r/sites/HVV-AILabAnnotation/"
    "Shared%20Documents/Work_log_2026_11.xlsx.xlsx"
    "?d=w0f7b0e57e89249a1938c009da82fefa4&csf=1&web=1&e=eqHrz1"
)
DEFAULT_SP_RW_URL = (
    "https://hanwhaworkplace.sharepoint.com/:x:/r/sites/HVV-AILabAnnotation/"
    "Shared%20Documents/Inspection_And_Rework_2026.11.xlsx"
    "?d=w3f3285489dac4f98addefbbd88d29790&csf=1&web=1&e=kUHBnR"
)
DEFAULT_SP_DIST_URL = (
    "https://hanwhaworkplace.sharepoint.com/:x:/r/sites/HVV-AILabAnnotation/"
    "_layouts/15/Doc.aspx?sourcedoc=%7B62933D2F-977C-4136-9CFC-D0485D980413%7D"
    "&file=distribution_sheet_2026.xlsx&action=default&mobileredirect=true&wdwpf=doclib-c"
)


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
        btn_w = 22
        icon_sz = 13
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
    """Auto-detect original Work_log, Inspection_And_Rework, and Distribution Sheet files in candidate dirs."""
    wl_found, rw_found, dist_found = "", "", ""
    for d in search_dirs:
        if not d or not os.path.isdir(d):
            continue
        wl_cands, rw_cands, dist_cands = [], [], []
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
            elif "distribution" in low:
                dist_cands.append(full)
        wl_cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        rw_cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        dist_cands.sort(key=lambda p: os.path.getmtime(p), reverse=True)
        if not wl_found and wl_cands:
            wl_found = wl_cands[0]
        if not rw_found and rw_cands:
            rw_found = rw_cands[0]
        if not dist_found and dist_cands:
            dist_found = dist_cands[0]
        if wl_found and rw_found and dist_found:
            break
    return wl_found, rw_found, dist_found


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


def _merge_job_id_diff_rows(diffs: List[DailyDiffRecord]) -> List[DailyDiffRecord]:
    groups = defaultdict(list)
    for i, diff in enumerate(diffs):
        if (
            diff.is_pending
            or not diff.reason.startswith("Sai ID Job Work Log:")
            or " | " in diff.reason
        ):
            continue
        wl_match = re.fullmatch(r"(-?\d+)~(-?\d+)", diff.wl_range)
        rw_match = re.fullmatch(r"(-?\d+)~(-?\d+)", diff.rw_range)
        if not wl_match or not rw_match or not re.fullmatch(r"[^;]+!R\d+", diff.rw_rows):
            continue
        start, end = sorted(map(int, wl_match.groups()))
        if [start, end] != sorted(map(int, rw_match.groups())):
            continue
        key = (diff.annotator_key, diff.task, diff.job_id, diff.reason)
        groups[key].append((start, end, i))

    display_rows = dict(enumerate(diffs))
    for blocks in groups.values():
        blocks.sort()
        if any(end >= next_start for (_, end, _), (next_start, _, _) in zip(blocks, blocks[1:])):
            continue
        runs = []
        for block in blocks:
            if not runs or runs[-1][-1][1] + 1 != block[0]:
                runs.append([])
            runs[-1].append(block)

        for run in runs:
            if len(run) < 2:
                continue
            indices = sorted(i for _, _, i in run)
            rows = [diffs[i] for i in indices]
            if len({row.rw_rows for row in rows}) != len(rows):
                continue
            first = min(rows, key=lambda row: row.date)
            frame_range = f"{run[0][0]}~{run[-1][1]}"
            display_rows[indices[0]] = replace(
                first,
                rw_date=min((row.rw_date for row in rows if row.rw_date is not None), default=None),
                wl_frames=sum(row.wl_frames for row in rows),
                rw_frames=sum(row.rw_frames for row in rows),
                wl_bbox=sum(row.wl_bbox for row in rows),
                rw_bbox=sum(row.rw_bbox for row in rows),
                wl_range=frame_range,
                rw_range=frame_range,
                inspectors=unique_inspector_names(name for row in rows for name in row.inspectors),
            )
            for i in indices[1:]:
                del display_rows[i]

    return list(display_rows.values())


class LoaderThread(QThread):
    progress = pyqtSignal(int, str)
    finished_ok = pyqtSignal(object)
    failed = pyqtSignal(str)

    def __init__(self, wl_path: str, rw_path: str, dist_path: str = ""):
        super().__init__()
        self.wl_path = wl_path
        self.rw_path = rw_path
        self.dist_path = dist_path

    def run(self):
        try:
            ds = load_dataset(
                self.wl_path,
                self.rw_path,
                progress_cb=lambda p, msg: self.progress.emit(p, msg),
                distribution_path=self.dist_path,
            )
            self.finished_ok.emit(ds)
        except Exception as exc:
            self.failed.emit(str(exc))


def _sp_cache_dir() -> str:
    base = os.path.join(os.environ.get("APPDATA") or os.path.expanduser("~"), "WorklogReportPro", "sp_cache")
    os.makedirs(base, exist_ok=True)
    return base


class SharePointDialog(QDialog):
    """Dialog for embedded SharePoint browser login and 2 Excel link inputs (zero Azure IT approval)."""

    def __init__(self, settings: dict, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Kết nối SharePoint (Trình duyệt nhúng)")
        self.setMinimumWidth(560)
        self.settings = settings
        self.sync_requested = False

        lay = QVBoxLayout(self)
        lay.setContentsMargins(18, 16, 18, 16)
        lay.setSpacing(12)

        hdr = QLabel("☁ Lấy dữ liệu trực tiếp từ SharePoint")
        hdr.setStyleSheet("font-size: 15px; font-weight: 700;")
        lay.addWidget(hdr)

        sub = QLabel(
            "Đăng nhập tài khoản Microsoft công ty (không cần IT phê duyệt) và bấm tải dữ liệu. "
            "Các link SharePoint mặc định (Work_log, Review, Distribution) đã được cấu hình sẵn."
        )
        sub.setWordWrap(True)
        sub.setStyleSheet("font-size: 12px; color: #64748B;")
        lay.addWidget(sub)

        self.links_box = QWidget()
        form = QFormLayout(self.links_box)
        form.setContentsMargins(0, 0, 0, 0)
        form.setSpacing(8)
        self.ed_wl_url = QLineEdit(settings.get("sp_wl_url") or DEFAULT_SP_WL_URL)
        self.ed_wl_url.setPlaceholderText("https://<tenant>.sharepoint.com/... (Copy link file Work_log)")
        self.ed_rw_url = QLineEdit(settings.get("sp_rw_url") or DEFAULT_SP_RW_URL)
        self.ed_rw_url.setPlaceholderText("https://<tenant>.sharepoint.com/... (Copy link file Inspection & Rework)")
        self.ed_dist_url = QLineEdit(settings.get("sp_dist_url") or DEFAULT_SP_DIST_URL)
        self.ed_dist_url.setPlaceholderText("https://<tenant>.sharepoint.com/... (Distribution Sheet — check ID Job)")

        form.addRow("Link Work_log:", self.ed_wl_url)
        form.addRow("Link Review:", self.ed_rw_url)
        form.addRow("Link Distribution:", self.ed_dist_url)
        self.links_box.setVisible(False)
        lay.addWidget(self.links_box)

        self.btn_toggle_links = QPushButton("⚙ Đổi link SharePoint (nâng cao)")
        self.btn_toggle_links.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_toggle_links.setFlat(True)
        self.btn_toggle_links.setStyleSheet("text-align: left; font-size: 11.5px; color: #64748B; padding: 2px 0;")
        self.btn_toggle_links.clicked.connect(lambda: self.links_box.setVisible(not self.links_box.isVisible()))
        lay.addWidget(self.btn_toggle_links)

        # Auth status row
        auth_row = QHBoxLayout()
        auth_row.setSpacing(8)
        self.lbl_auth = QLabel("Chưa đăng nhập")
        self.lbl_auth.setStyleSheet("font-weight: 600;")
        auth_row.addWidget(self.lbl_auth, 1)

        self.btn_login = QPushButton("Đăng nhập SharePoint")
        self.btn_login.setObjectName("PrimaryBtn")
        self.btn_login.clicked.connect(self._do_login)
        auth_row.addWidget(self.btn_login)

        self.btn_logout = QPushButton("Đăng xuất")
        self.btn_logout.clicked.connect(self._do_logout)
        auth_row.addWidget(self.btn_logout)
        lay.addLayout(auth_row)

        # Action buttons
        btn_row = QHBoxLayout()
        btn_row.addStretch(1)
        self.btn_cancel = QPushButton("Đóng")
        self.btn_cancel.clicked.connect(self.reject)
        self.btn_sync = QPushButton("☁ Tải dữ liệu từ SharePoint")
        self.btn_sync.setObjectName("SuccessBtn")
        self.btn_sync.clicked.connect(self._do_sync)
        btn_row.addWidget(self.btn_cancel)
        btn_row.addWidget(self.btn_sync)
        lay.addLayout(btn_row)

        self._refresh_auth_status()

    def _refresh_auth_status(self):
        host = has_saved_sp_session(_sp_cache_dir())
        if host:
            self.lbl_auth.setText(f"✓ Đã lưu phiên đăng nhập ({host})")
            self.lbl_auth.setStyleSheet("font-weight: 700; color: #059669;")
            self.btn_logout.setEnabled(True)
        else:
            self.lbl_auth.setText("Chưa đăng nhập SharePoint")
            self.lbl_auth.setStyleSheet("font-weight: 600; color: #64748B;")
            self.btn_logout.setEnabled(False)

    def _do_login(self):
        start_url = extract_sp_site_url(self.ed_wl_url.text(), self.ed_rw_url.text(), self.ed_dist_url.text())
        login_dlg = SharePointLoginDialog(start_url, _sp_cache_dir(), self)
        login_dlg.exec()
        self._refresh_auth_status()

    def _do_logout(self):
        clear_sp_session(_sp_cache_dir())
        self._refresh_auth_status()

    def _do_sync(self):
        if not (self.ed_wl_url.text().strip() or self.ed_rw_url.text().strip()):
            QMessageBox.warning(self, "Thiếu link SharePoint", "Vui lòng dán ít nhất 1 link SharePoint (.xlsx).")
            return
        self.sync_requested = True
        self.accept()

    def values(self) -> dict:
        return {
            "sp_wl_url": self.ed_wl_url.text().strip(),
            "sp_rw_url": self.ed_rw_url.text().strip(),
            "sp_dist_url": self.ed_dist_url.text().strip(),
        }


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — Báo cáo & Check Data")
        self.setAcceptDrops(True)

        icon_file = _resource_path("app_icon.ico")
        if os.path.isfile(icon_file):
            self.setWindowIcon(QIcon(icon_file))

        self.settings = _load_settings()
        self.theme = DARK_THEME if self.settings.get("theme") == "dark" else LIGHT_THEME
        self.dataset: Optional[LoadedDataset] = None
        self.current_report: Optional[ReportResult] = None
        self._period_report: Optional[ReportResult] = None
        self._linked_audit_key: Optional[Tuple[str, str]] = None
        self.loader_thread: Optional[LoaderThread] = None
        self.update_check_thread: Optional[UpdateCheckThread] = None
        self.latest_release: Optional[ReleaseInfo] = None
        self._update_check_silent = True
        self._updating_filters = False
        self._isolated_annotator: Optional[str] = None
        self._task_chips: Dict[str, QPushButton] = {}

        cleanup_old_update_files()

        self._build_ui()
        self._apply_theme(self.theme)
        self._setup_shortcuts()
        self._restore_files_only()
        self._show_welcome_page()

        self.winId()
        window = self.windowHandle()
        work_area = window.screen().availableGeometry()
        client_area = work_area.marginsRemoved(window.frameMargins())
        startup_size = QSize(1800, 880).boundedTo(client_area.size())
        self.setMinimumSize(startup_size)
        self.resize(startup_size)
        frame = self.frameGeometry()
        frame.moveCenter(work_area.center())
        self.move(frame.topLeft())

        QTimer.singleShot(1500, lambda: self.check_for_updates(silent=True))

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
        title_top = QHBoxLayout()
        title_top.setSpacing(6)
        title_top.setContentsMargins(0, 0, 0, 0)
        lbl_t = QLabel(APP_NAME)
        lbl_t.setObjectName("AppTitle")
        title_top.addWidget(lbl_t)

        self.btn_version = QPushButton(f"v{APP_VERSION}")
        self.btn_version.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_version.setFixedHeight(20)
        self.btn_version.setToolTip("Kiểm tra cập nhật phiên bản mới")
        self.btn_version.clicked.connect(self._on_version_badge_clicked)
        title_top.addWidget(self.btn_version)
        title_top.addStretch(1)

        lbl_s = QLabel("Báo cáo Năng suất & Đối chiếu Dữ liệu (Check Data)")
        lbl_s.setObjectName("AppSubtitle")
        title_col.addLayout(title_top)
        title_col.addWidget(lbl_s)
        top_lay.addLayout(title_col)

        top_lay.addStretch(1)

        self.btn_back_home = QPushButton(" ← Đổi nguồn")
        self.btn_back_home.setToolTip("Quay lại màn hình đăng nhập SharePoint / chọn file đầu vào")
        self.btn_back_home.clicked.connect(self._show_welcome_page)
        top_lay.addWidget(self.btn_back_home)

        self.btn_onedrive = QPushButton(" ☁ Nguồn OneDrive")
        self.btn_onedrive.setToolTip(f"Nạp trực tiếp từ thư mục OneDrive Sync:\n{ONEDRIVE_PRIMARY_PATH}")
        self.btn_onedrive.clicked.connect(self.switch_to_onedrive_source)
        top_lay.addWidget(self.btn_onedrive)

        self.btn_sharepoint = QPushButton(" SharePoint")
        self.btn_sharepoint.setToolTip("Đăng nhập tài khoản Microsoft để tải trực tiếp 2 file Excel từ SharePoint")
        self.btn_sharepoint.clicked.connect(self.open_sharepoint_dialog)
        top_lay.addWidget(self.btn_sharepoint)

        self.btn_reload = QPushButton(" Nạp lại (F5)")
        self.btn_reload.setToolTip("Tải lại dữ liệu mới nhất (tự đồng bộ lại SharePoint nếu đang dùng nguồn SharePoint) (F5)")
        self.btn_reload.clicked.connect(self.reload_data)
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

        # In-app update notification banner (visible across Welcome & Workspace)
        self.update_banner = UpdateBanner(self)
        self.update_banner.update_clicked.connect(lambda: self.open_update_dialog(auto_start=True))
        self.update_banner.details_clicked.connect(lambda: self.open_update_dialog(auto_start=False))
        root.addWidget(self.update_banner)

        # ---- 2. PAGE STACK: Page 0 = Welcome (Login / Load), Page 1 = Main Workspace ----
        self.page_stack = QStackedWidget()
        root.addWidget(self.page_stack, 1)

        self.welcome_page = self._build_welcome_page()
        self.page_stack.addWidget(self.welcome_page)

        # ---- MAIN WORKSPACE SPLITTER (Page 1) ----
        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setHandleWidth(1)
        self.page_stack.addWidget(splitter)

        # ---- LEFT SIDEBAR (compact to maximize table workspace) ----
        sidebar = QFrame()
        sidebar.setObjectName("Sidebar")
        sidebar.setMinimumWidth(225)
        sidebar.setMaximumWidth(275)
        sb_lay = QVBoxLayout(sidebar)
        sb_lay.setContentsMargins(9, 8, 9, 8)
        sb_lay.setSpacing(7)

        # Section A: Date Range — Month/Year first, then Weeks + Full Month, then manual Start -> End
        sb_lay.addWidget(self._section_label("KHOẢNG THỜI GIAN"))

        wm_row = QHBoxLayout()
        wm_row.setSpacing(4)
        self.cb_week_month = QComboBox()
        self.cb_week_month.setToolTip("Chọn Tháng/Năm để xem cả tháng hoặc từng tuần (Thứ 2 → Chủ nhật)")
        self.cb_week_month.currentIndexChanged.connect(self._rebuild_week_buttons)
        wm_row.addWidget(self.cb_week_month, 1)

        self.btn_full_month = QPushButton("Cả tháng")
        self.btn_full_month.setObjectName("PresetBtn")
        self.btn_full_month.setCheckable(True)
        self.btn_full_month.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_full_month.setToolTip("Chọn toàn bộ các ngày trong tháng đang chọn")
        self.btn_full_month.clicked.connect(self._apply_selected_month)
        wm_row.addWidget(self.btn_full_month)
        sb_lay.addLayout(wm_row)

        self.week_btns_layout = QGridLayout()
        self.week_btns_layout.setSpacing(4)
        self.week_btns_layout.setContentsMargins(0, 0, 0, 0)
        sb_lay.addLayout(self.week_btns_layout)
        self._week_buttons = []

        d_row = QHBoxLayout()
        d_row.setSpacing(3)
        self.dt_start = CalendarDateEdit()
        self.dt_start.setDate(QDate(2026, 11, 9))
        self.dt_start.dateChanged.connect(self._on_date_changed)
        lbl_arrow = QLabel("→")
        lbl_arrow.setStyleSheet("font-weight: 700; color: #64748B; font-size: 11px;")
        self.dt_end = CalendarDateEdit()
        self.dt_end.setDate(QDate(2026, 11, 15))
        self.dt_end.dateChanged.connect(self._on_date_changed)
        d_row.addWidget(self.dt_start, 1)
        d_row.addWidget(lbl_arrow)
        d_row.addWidget(self.dt_end, 1)
        sb_lay.addLayout(d_row)

        self._populate_week_months()

        # Section B: Data Source & Mode (Segmented Controls)
        sb_lay.addWidget(self._section_label("NGUỒN DỮ LIỆU"))
        self.seg_source = SegmentedControl([
            ("Review", "rework"),
            ("Work Log", "worklog"),
        ])
        self.seg_source.setToolTip("Chọn nguồn sản lượng Frame / BBOX: Review (#IMAGE) hoặc Work Log (+ Note)")
        if self.settings.get("source_mode") in ("rework", "worklog"):
            self.seg_source.set_value(self.settings["source_mode"])
        self.seg_source.changed.connect(self._on_option_changed)
        sb_lay.addWidget(self.seg_source)

        self.chk_highlight = QCheckBox("Tô màu dòng không khớp")
        self.chk_highlight.setToolTip("Tô màu cảnh báo khi dữ liệu giữa hai file không khớp")
        self.chk_highlight.setChecked(True)
        self.chk_highlight.toggled.connect(self.refresh_report)
        sb_lay.addWidget(self.chk_highlight)

        # Section C2: TASK Filter
        t_hdr = QHBoxLayout()
        t_hdr.setSpacing(2)
        t_hdr.addWidget(self._section_label("LỌC TASK"))
        t_hdr.addStretch(1)
        b_tall = QPushButton("Tất cả")
        b_tall.setObjectName("GhostBtn")
        b_tall.clicked.connect(lambda: self._set_all_tasks(True))
        b_tnone = QPushButton("Bỏ")
        b_tnone.setObjectName("GhostBtn")
        b_tnone.clicked.connect(lambda: self._set_all_tasks(False))
        t_hdr.addWidget(b_tall)
        t_hdr.addWidget(b_tnone)
        sb_lay.addLayout(t_hdr)

        self.list_tasks = QListWidget()
        self.list_tasks.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.list_tasks.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.list_tasks.setMaximumHeight(130)
        self.list_tasks.setToolTip("Tick để chọn/bỏ chọn TASK đưa vào báo cáo tổng hợp & xếp hạng")
        self.list_tasks.itemChanged.connect(self._on_task_item_changed)
        sb_lay.addWidget(self.list_tasks)

        self.people_filters = QStackedWidget()
        self.annotator_filter, self.ed_search, self.list_anns = self._build_people_filter(
            "Annotator", self._set_all_anns, self._select_active_annotators,
            self._on_ann_item_changed, self._isolate_single_annotator,
        )
        self.inspector_filter, self.ed_inspector_search, self.list_inspectors = self._build_people_filter(
            "Inspector", self._set_all_inspectors, self._select_active_inspectors,
            self._on_inspector_item_changed, self._isolate_single_inspector,
        )
        self.people_filters.addWidget(self.annotator_filter)
        self.people_filters.addWidget(self.inspector_filter)
        sb_lay.addWidget(self.people_filters, 1)

        splitter.addWidget(sidebar)

        # ---- RIGHT CONTENT WORKSPACE ----
        main_area = QWidget()
        main_lay = QVBoxLayout(main_area)
        main_lay.setContentsMargins(16, 12, 16, 10)
        main_lay.setSpacing(12)

        # File Drop Cards Row
        files_container = QWidget()
        files_row = QHBoxLayout(files_container)
        files_row.setContentsMargins(0, 0, 0, 0)
        files_row.setSpacing(10)
        self.card_wl = FileDropCard("FILE WORK LOG (CẢ EXCEL NOTES)", "Kéo thả hoặc chọn Work_log_*.xlsx", "#2563EB")
        self.card_wl.browse_clicked.connect(lambda: self._browse_file("wl"))
        self.card_wl.file_dropped.connect(lambda p: self._set_file_and_reload("wl", p))

        self.card_rw = FileDropCard("FILE INSPECTION & REVIEW", "Kéo thả hoặc chọn Inspection_And_Rework_*.xlsx", "#059669")
        self.card_rw.browse_clicked.connect(lambda: self._browse_file("rw"))
        self.card_rw.file_dropped.connect(lambda p: self._set_file_and_reload("rw", p))

        self.card_dist = FileDropCard("FILE DISTRIBUTION (CHECK ID JOB)", "Tùy chọn: distribution_sheet_*.xlsx", "#7C3AED")
        self.card_dist.browse_clicked.connect(lambda: self._browse_file("dist"))
        self.card_dist.file_dropped.connect(lambda p: self._set_file_and_reload("dist", p))

        files_row.addWidget(self.card_wl, 1)
        files_row.addWidget(self.card_rw, 1)
        files_row.addWidget(self.card_dist, 1)
        main_lay.addWidget(files_container)
        files_container.hide()

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

        # Tabs (Task, Annotator, Inspector, Bảng Xếp Hạng, Đối chiếu & Notes)
        self.tabs = QTabWidget()

        # Tab 1: Task (Workers only) — Main table on top, per-Task summary table below
        tab_t1 = QWidget()
        t1_lay = QVBoxLayout(tab_t1)
        t1_lay.setContentsMargins(0, 0, 0, 0)
        t1_lay.setSpacing(0)

        t1_split = QSplitter(Qt.Orientation.Vertical)
        t1_split.setChildrenCollapsible(False)

        self.tbl_t1 = QTableWidget()
        self._setup_table(self.tbl_t1)
        self.del_t1 = ReportTableDelegate(divider_col=9, parent=self.tbl_t1)
        self.tbl_t1.setItemDelegate(self.del_t1)
        self.tbl_t1.cellDoubleClicked.connect(self._on_t1_double_click)
        t1_split.addWidget(self.tbl_t1)

        t1_sum_wrap = QWidget()
        t1_sum_lay = QVBoxLayout(t1_sum_wrap)
        t1_sum_lay.setContentsMargins(8, 6, 8, 6)
        t1_sum_lay.setSpacing(5)
        lbl_t1_sum = QLabel("📊 Tổng hợp năng suất & tỉ lệ lỗi theo từng TASK")
        lbl_t1_sum.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        t1_sum_lay.addWidget(lbl_t1_sum)

        self.tbl_t1_summary = QTableWidget()
        self._setup_table(self.tbl_t1_summary)
        self.tbl_t1_summary.setMinimumHeight(110)
        t1_sum_lay.addWidget(self.tbl_t1_summary)
        t1_split.addWidget(t1_sum_wrap)

        t1_split.setStretchFactor(0, 3)
        t1_split.setStretchFactor(1, 1)
        t1_lay.addWidget(t1_split)
        self.tabs.addTab(tab_t1, "Task")

        # Tab 2: Annotator (Workers only)
        self.tbl_t2 = QTableWidget()
        self._setup_table(self.tbl_t2)
        self.del_t2 = ReportTableDelegate(divider_col=8, parent=self.tbl_t2)
        self.tbl_t2.setItemDelegate(self.del_t2)
        self.tbl_t2.cellDoubleClicked.connect(self._on_t2_double_click)
        self.tabs.addTab(self.tbl_t2, "Annotator")

        # Tab 3: Inspector (Dedicated tab for 1st Inspectors matching exact INSPECTOR -> Annotator template)
        self.tbl_rev = QTableWidget()
        self._setup_table(self.tbl_rev)
        self.del_rev = ReportTableDelegate(divider_col=7, parent=self.tbl_rev)
        self.tbl_rev.setItemDelegate(self.del_rev)
        self.tbl_rev.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_rev, r))
        self.tabs.addTab(self.tbl_rev, "Inspector")

        # Tab 4: Bảng Xếp Hạng (2 sub-tabs: Tỉ lệ lỗi Frame (%) & Hiệu suất lỗi thấp + sản lượng nhiều)
        self.tab_rank = QWidget()
        rank_root_lay = QVBoxLayout(self.tab_rank)
        rank_root_lay.setContentsMargins(10, 10, 10, 10)
        rank_root_lay.setSpacing(8)

        self.rank_mode_tabs = QTabWidget()

        # Sub-tab 0: Tỉ lệ lỗi Frame (%)
        page_rank_err = QWidget()
        rank_lay = QVBoxLayout(page_rank_err)
        rank_lay.setContentsMargins(10, 10, 10, 10)
        rank_lay.setSpacing(8)

        rank_top = QHBoxLayout()
        lbl_rk_info = QLabel("🏆 Bảng xếp hạng tỉ lệ lỗi Frame thấp nhất (Hạng 1 = Tỉ lệ lỗi Frame thấp nhất)")
        lbl_rk_info.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        rank_top.addWidget(lbl_rk_info)
        rank_top.addStretch(1)
        rank_lay.addLayout(rank_top)

        rank_split = QSplitter(Qt.Orientation.Horizontal)

        w_rank_box = QWidget()
        w_rlay = QVBoxLayout(w_rank_box)
        w_rlay.setContentsMargins(0, 0, 4, 0)
        w_rlay.setSpacing(5)
        self.lbl_rank_worker = QLabel("🥇 Xếp hạng Annotator (Worker — Lỗi Lần 1 + Lần 2)")
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
        self.lbl_rank_reviewer = QLabel("🥇 Xếp hạng Inspector Lần 1")
        self.lbl_rank_reviewer.setStyleSheet("font-weight: 700; font-size: 12.5px; color: #059669;")
        rev_rlay.addWidget(self.lbl_rank_reviewer)
        self.tbl_rank_reviewer = QTableWidget()
        self._setup_table(self.tbl_rank_reviewer)
        self.tbl_rank_reviewer.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_rank_reviewer, r))
        rev_rlay.addWidget(self.tbl_rank_reviewer)
        rank_split.addWidget(rev_rank_box)

        rank_split.setSizes([600, 600])
        rank_lay.addWidget(rank_split, 1)
        self.rank_mode_tabs.addTab(page_rank_err, "Tỉ lệ lỗi Frame (%)")

        # Sub-tab 1: Hiệu suất (Tỉ lệ lỗi Frame thấp & Sản lượng nhiều)
        page_rank_perf = QWidget()
        perf_lay = QVBoxLayout(page_rank_perf)
        perf_lay.setContentsMargins(10, 10, 10, 10)
        perf_lay.setSpacing(8)

        perf_top = QHBoxLayout()
        lbl_pf_info = QLabel("⭐ Xếp hạng Hiệu suất: Tỉ lệ lỗi thấp & Sản lượng cao (Thang 100 điểm)")
        lbl_pf_info.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        lbl_pf_info.setToolTip(
            "Điểm Tổng = (Trọng số Chất lượng × Điểm lỗi Frame thấp) + (Trọng số Sản lượng × Điểm sản lượng Frame & BBOX)"
        )
        perf_top.addWidget(lbl_pf_info)
        perf_top.addStretch(1)
        perf_top.addWidget(QLabel("Trọng số:"))
        self.seg_perf_weight = SegmentedControl([
            ("Cân bằng 50/50", "50_50"),
            ("Lỗi thấp 60/40", "60_40"),
            ("Sản lượng 40/60", "40_60"),
        ])
        self.seg_perf_weight.setFixedWidth(360)
        self.seg_perf_weight.changed.connect(lambda _v: self._render_performance_tables(self.current_report) if self.current_report else None)
        perf_top.addWidget(self.seg_perf_weight)
        perf_lay.addLayout(perf_top)

        perf_split = QSplitter(Qt.Orientation.Horizontal)

        w_perf_box = QWidget()
        w_play = QVBoxLayout(w_perf_box)
        w_play.setContentsMargins(0, 0, 4, 0)
        w_play.setSpacing(5)
        self.lbl_perf_worker = QLabel("⭐ Hiệu suất Annotator / Worker")
        self.lbl_perf_worker.setStyleSheet("font-weight: 700; font-size: 12.5px; color: #2563EB;")
        w_play.addWidget(self.lbl_perf_worker)
        self.tbl_perf_worker = QTableWidget()
        self._setup_table(self.tbl_perf_worker)
        self.tbl_perf_worker.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_perf_worker, r))
        w_play.addWidget(self.tbl_perf_worker)
        perf_split.addWidget(w_perf_box)

        rev_perf_box = QWidget()
        rev_play = QVBoxLayout(rev_perf_box)
        rev_play.setContentsMargins(4, 0, 0, 0)
        rev_play.setSpacing(5)
        self.lbl_perf_reviewer = QLabel("⭐ Hiệu suất Inspector Lần 1")
        self.lbl_perf_reviewer.setStyleSheet("font-weight: 700; font-size: 12.5px; color: #059669;")
        rev_play.addWidget(self.lbl_perf_reviewer)
        self.tbl_perf_reviewer = QTableWidget()
        self._setup_table(self.tbl_perf_reviewer)
        self.tbl_perf_reviewer.cellDoubleClicked.connect(lambda r, _c: self._on_rank_double_click(self.tbl_perf_reviewer, r))
        rev_play.addWidget(self.tbl_perf_reviewer)
        perf_split.addWidget(rev_perf_box)

        perf_split.setSizes([600, 600])
        perf_lay.addWidget(perf_split, 1)
        self.rank_mode_tabs.addTab(page_rank_perf, "Hiệu suất")

        rank_root_lay.addWidget(self.rank_mode_tabs)
        self.tabs.addTab(self.tab_rank, "Bảng Xếp Hạng")

        # Tab 6: Kiểm tra dữ liệu (2 sub-tabs: Check Annotator & Check Inspector)
        self.tab_audit = QWidget()
        audit_lay = QVBoxLayout(self.tab_audit)
        audit_lay.setContentsMargins(10, 10, 10, 10)
        audit_lay.setSpacing(8)

        self.audit_mode_tabs = QTabWidget()

        page_ann, self.cb_audit_status, self.tbl_audit, self.lbl_detail_hdr, self.tbl_daily_diffs = self._build_audit_page()
        self.cb_audit_status.currentIndexChanged.connect(self._on_audit_status_changed)
        self.tbl_audit.currentCellChanged.connect(self._render_daily_diffs_table)
        self.audit_mode_tabs.addTab(page_ann, "Annotator")

        page_rev, self.cb_rev_audit_status, self.tbl_rev_audit, self.lbl_rev_detail_hdr, self.tbl_rev_diffs = self._build_audit_page()
        self.cb_rev_audit_status.currentIndexChanged.connect(
            lambda: self._render_reviewer_audit_table(self.current_report) if self.current_report else None
        )
        self.tbl_rev_audit.currentCellChanged.connect(self._render_rev_diffs_table)
        self.audit_mode_tabs.addTab(page_rev, "Inspector")

        audit_lay.addWidget(self.audit_mode_tabs)
        self.tabs.addTab(self.tab_audit, "Kiểm tra dữ liệu")
        self.tabs.currentChanged.connect(self._sync_people_filter_panel)

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
        self.chk_exp_rev = QCheckBox("Inspector")
        self.chk_exp_rev.setChecked(True)
        self.chk_exp_rank = QCheckBox("Bảng Xếp Hạng")
        self.chk_exp_rank.setChecked(True)
        self.chk_exp_audit = QCheckBox("Kiểm tra dữ liệu")
        self.chk_exp_audit.setChecked(True)
        foot.addWidget(self.chk_exp_t1)
        foot.addWidget(self.chk_exp_t2)
        foot.addWidget(self.chk_exp_rev)
        foot.addWidget(self.chk_exp_rank)
        foot.addWidget(self.chk_exp_audit)
        foot.addStretch(1)

        self.lbl_period_badge = QLabel("")
        self.lbl_period_badge.setStyleSheet("font-size: 12px; font-weight: 700; color: #2563EB;")
        foot.addWidget(self.lbl_period_badge)
        main_lay.addLayout(foot)

        splitter.addWidget(main_area)
        splitter.setSizes([240, 1240])

        # Status Bar + Floating Toast
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_bar.showMessage("Kéo thả 2 file Excel vào cửa sổ hoặc bấm chọn file để bắt đầu.")

        self.btn_status_update = QPushButton(f"v{APP_VERSION} • Kiểm tra cập nhật")
        self.btn_status_update.setCursor(Qt.CursorShape.PointingHandCursor)
        self.btn_status_update.setFlat(True)
        self.btn_status_update.setToolTip("Kiểm tra phiên bản mới từ GitHub Releases")
        self.btn_status_update.clicked.connect(lambda: self.check_for_updates(silent=False))
        self.status_bar.addPermanentWidget(self.btn_status_update)

        self.toast = ToastBanner(self)

    def _build_welcome_page(self) -> QWidget:
        page = QWidget()
        outer = QVBoxLayout(page)
        outer.setContentsMargins(36, 28, 36, 28)
        outer.setSpacing(18)
        outer.addStretch(1)

        # Hero title
        hero = QVBoxLayout()
        hero.setSpacing(4)
        lbl_h = QLabel(f"Chào mừng đến với {APP_NAME}")
        lbl_h.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_h.setStyleSheet("font-size: 22px; font-weight: 700;")
        lbl_sub = QLabel("Chọn đăng nhập SharePoint để lấy dữ liệu trực tuyến hoặc nạp 2 file Excel từ máy / OneDrive")
        lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_sub.setStyleSheet("font-size: 13px; color: #64748B;")
        hero.addWidget(lbl_h)
        hero.addWidget(lbl_sub)
        outer.addLayout(hero)

        # Two side-by-side cards
        cards_wrap = QHBoxLayout()
        cards_wrap.setSpacing(18)
        cards_wrap.addStretch(1)

        # Card 1: SharePoint Login & Sync
        sp_card = QFrame()
        sp_card.setObjectName("SectionCard")
        sp_card.setMinimumWidth(440)
        sp_card.setMaximumWidth(500)
        sp_lay = QVBoxLayout(sp_card)
        sp_lay.setContentsMargins(20, 18, 20, 18)
        sp_lay.setSpacing(12)

        sp_title = QLabel("☁ Đăng nhập SharePoint (Online)")
        sp_title.setStyleSheet("font-size: 15px; font-weight: 700;")
        sp_lay.addWidget(sp_title)

        self.w_sp_links_box = QWidget()
        sp_form = QFormLayout(self.w_sp_links_box)
        sp_form.setContentsMargins(0, 0, 0, 0)
        sp_form.setSpacing(8)
        self.w_ed_wl_url = QLineEdit(self.settings.get("sp_wl_url") or DEFAULT_SP_WL_URL)
        self.w_ed_wl_url.setPlaceholderText("Dán 'Copy link' file Work_log (.xlsx)")
        self.w_ed_rw_url = QLineEdit(self.settings.get("sp_rw_url") or DEFAULT_SP_RW_URL)
        self.w_ed_rw_url.setPlaceholderText("Dán 'Copy link' file Inspection & Rework (.xlsx)")
        self.w_ed_dist_url = QLineEdit(self.settings.get("sp_dist_url") or DEFAULT_SP_DIST_URL)
        self.w_ed_dist_url.setPlaceholderText("Link distribution_sheet_*.xlsx (tùy chọn — check ID Job)")
        sp_form.addRow("Link Work_log:", self.w_ed_wl_url)
        sp_form.addRow("Link Review:", self.w_ed_rw_url)
        sp_form.addRow("Link Distribution:", self.w_ed_dist_url)
        self.w_sp_links_box.setVisible(False)
        sp_lay.addWidget(self.w_sp_links_box)

        self.w_btn_toggle_links = QPushButton("⚙ Đổi link SharePoint (nâng cao)")
        self.w_btn_toggle_links.setCursor(Qt.CursorShape.PointingHandCursor)
        self.w_btn_toggle_links.setFlat(True)
        self.w_btn_toggle_links.setStyleSheet("text-align: left; font-size: 11.5px; color: #64748B; padding: 2px 0;")
        self.w_btn_toggle_links.clicked.connect(lambda: self.w_sp_links_box.setVisible(not self.w_sp_links_box.isVisible()))
        sp_lay.addWidget(self.w_btn_toggle_links)

        sp_auth_row = QHBoxLayout()
        sp_auth_row.setSpacing(8)
        self.w_lbl_sp_auth = QLabel("Chưa đăng nhập")
        self.w_lbl_sp_auth.setStyleSheet("font-size: 12px; font-weight: 600; color: #64748B;")
        sp_auth_row.addWidget(self.w_lbl_sp_auth, 1)

        self.w_btn_sp_login = QPushButton("Đăng nhập SharePoint")
        self.w_btn_sp_login.setObjectName("PrimaryBtn")
        self.w_btn_sp_login.setCursor(Qt.CursorShape.PointingHandCursor)
        self.w_btn_sp_login.clicked.connect(self._welcome_sp_login)
        sp_auth_row.addWidget(self.w_btn_sp_login)

        self.w_btn_sp_logout = QPushButton("Đăng xuất")
        self.w_btn_sp_logout.setCursor(Qt.CursorShape.PointingHandCursor)
        self.w_btn_sp_logout.clicked.connect(self._welcome_sp_logout)
        sp_auth_row.addWidget(self.w_btn_sp_logout)
        sp_lay.addLayout(sp_auth_row)

        sp_lay.addStretch(1)
        self.w_btn_sp_sync = QPushButton("☁ Tải từ SharePoint && Vào báo cáo")
        self.w_btn_sp_sync.setObjectName("SuccessBtn")
        self.w_btn_sp_sync.setCursor(Qt.CursorShape.PointingHandCursor)
        self.w_btn_sp_sync.clicked.connect(self._welcome_sp_sync)
        sp_lay.addWidget(self.w_btn_sp_sync)
        cards_wrap.addWidget(sp_card)

        # Card 2: Local / OneDrive Excel Files
        loc_card = QFrame()
        loc_card.setObjectName("SectionCard")
        loc_card.setMinimumWidth(440)
        loc_card.setMaximumWidth(500)
        loc_lay = QVBoxLayout(loc_card)
        loc_lay.setContentsMargins(20, 18, 20, 18)
        loc_lay.setSpacing(12)

        loc_title = QLabel("📂 Nạp file Excel từ máy / OneDrive")
        loc_title.setStyleSheet("font-size: 15px; font-weight: 700;")
        loc_lay.addWidget(loc_title)

        self.w_card_wl = FileDropCard("FILE WORK LOG (CẢ EXCEL NOTES)", "Kéo thả hoặc chọn Work_log_*.xlsx", "#2563EB")
        self.w_card_wl.browse_clicked.connect(lambda: self._browse_file("wl", auto_load=False))
        self.w_card_wl.file_dropped.connect(lambda p: self._set_file_only("wl", p))
        loc_lay.addWidget(self.w_card_wl)

        self.w_card_rw = FileDropCard("FILE INSPECTION & REVIEW", "Kéo thả hoặc chọn Inspection_And_Rework_*.xlsx", "#059669")
        self.w_card_rw.browse_clicked.connect(lambda: self._browse_file("rw", auto_load=False))
        self.w_card_rw.file_dropped.connect(lambda p: self._set_file_only("rw", p))
        loc_lay.addWidget(self.w_card_rw)

        self.w_card_dist = FileDropCard("FILE DISTRIBUTION (CHECK ID JOB)", "Tùy chọn: distribution_sheet_*.xlsx", "#7C3AED")
        self.w_card_dist.browse_clicked.connect(lambda: self._browse_file("dist", auto_load=False))
        self.w_card_dist.file_dropped.connect(lambda p: self._set_file_only("dist", p))
        loc_lay.addWidget(self.w_card_dist)

        loc_lay.addStretch(1)
        loc_btns = QHBoxLayout()
        loc_btns.setSpacing(8)
        self.w_btn_od = QPushButton("☁ Tìm OneDrive Sync")
        self.w_btn_od.setCursor(Qt.CursorShape.PointingHandCursor)
        self.w_btn_od.clicked.connect(self._welcome_find_onedrive)
        loc_btns.addWidget(self.w_btn_od)

        self.w_btn_load = QPushButton("▶ Nạp dữ liệu && Vào báo cáo")
        self.w_btn_load.setObjectName("PrimaryBtn")
        self.w_btn_load.setCursor(Qt.CursorShape.PointingHandCursor)
        self.w_btn_load.clicked.connect(self.start_load_data)
        loc_btns.addWidget(self.w_btn_load, 1)
        loc_lay.addLayout(loc_btns)
        cards_wrap.addWidget(loc_card)

        cards_wrap.addStretch(1)
        outer.addLayout(cards_wrap)

        self.w_lbl_status = QLabel("")
        self.w_lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.w_lbl_status.setStyleSheet("font-size: 12.5px; font-weight: 600; color: #2563EB;")
        outer.addWidget(self.w_lbl_status)
        outer.addStretch(2)
        return page

    def _show_welcome_page(self):
        self.page_stack.setCurrentIndex(0)
        for b in (self.btn_back_home, self.btn_onedrive, self.btn_sharepoint, self.btn_reload, self.btn_copy, self.btn_export):
            b.setVisible(False)
        self._refresh_welcome_sp_auth()

    def _show_main_workspace(self):
        self.page_stack.setCurrentIndex(1)
        for b in (self.btn_back_home, self.btn_onedrive, self.btn_sharepoint, self.btn_reload, self.btn_copy, self.btn_export):
            b.setVisible(True)

    def _refresh_welcome_sp_auth(self):
        host = has_saved_sp_session(_sp_cache_dir())
        if host:
            self.w_lbl_sp_auth.setText(f"✓ Đã đăng nhập ({host})")
            self.w_lbl_sp_auth.setStyleSheet("font-size: 12px; font-weight: 700; color: #059669;")
            self.w_btn_sp_logout.setEnabled(True)
        else:
            self.w_lbl_sp_auth.setText("Chưa đăng nhập SharePoint")
            self.w_lbl_sp_auth.setStyleSheet("font-size: 12px; font-weight: 600; color: #64748B;")
            self.w_btn_sp_logout.setEnabled(False)

    def _save_welcome_sp_fields(self):
        self.settings["sp_wl_url"] = self.w_ed_wl_url.text().strip()
        self.settings["sp_rw_url"] = self.w_ed_rw_url.text().strip()
        self.settings["sp_dist_url"] = self.w_ed_dist_url.text().strip()
        self._save_current_settings()

    def _welcome_sp_login(self):
        self._save_welcome_sp_fields()
        start_url = extract_sp_site_url(self.w_ed_wl_url.text(), self.w_ed_rw_url.text(), self.w_ed_dist_url.text())
        dlg = SharePointLoginDialog(start_url, _sp_cache_dir(), self)
        dlg.exec()
        self._refresh_welcome_sp_auth()

    def _welcome_sp_logout(self):
        clear_sp_session(_sp_cache_dir())
        self._refresh_welcome_sp_auth()

    def _welcome_sp_sync(self):
        self._save_welcome_sp_fields()
        wl_u = self.w_ed_wl_url.text().strip()
        rw_u = self.w_ed_rw_url.text().strip()
        dist_u = self.w_ed_dist_url.text().strip()
        if not (wl_u or rw_u):
            QMessageBox.warning(self, "Thiếu link SharePoint", "Vui lòng dán ít nhất 1 link SharePoint (.xlsx).")
            return
        self._start_sharepoint_sync(wl_u, rw_u, dist_u)

    def _welcome_find_onedrive(self):
        od_dirs = _get_onedrive_dirs()
        od_wl, od_rw, od_dist = _find_default_files(od_dirs)
        if not od_wl and not od_rw:
            QMessageBox.warning(self, "Không tìm thấy file OneDrive", f"Không tìm thấy file Excel trong:\n{ONEDRIVE_PRIMARY_PATH}")
            return
        is_dark = self.theme["name"] == "dark"
        if od_wl:
            self.card_wl.set_path(od_wl, is_dark)
            self.w_card_wl.set_path(od_wl, is_dark)
        if od_rw:
            self.card_rw.set_path(od_rw, is_dark)
            self.w_card_rw.set_path(od_rw, is_dark)
        if od_dist:
            self.card_dist.set_path(od_dist, is_dark)
            self.w_card_dist.set_path(od_dist, is_dark)
        self._save_current_settings()
        self.toast.show_toast("✓ Đã tìm thấy file trong OneDrive Sync. Bấm 'Nạp dữ liệu' để mở.", "success")

    def _section_label(self, text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setObjectName("SectionOverline")
        return lbl

    def _build_people_filter(self, role: str, set_all, select_active, on_change, isolate):
        panel = QWidget()
        layout = QVBoxLayout(panel)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(7)

        header = QHBoxLayout()
        header.setSpacing(2)
        header.addWidget(self._section_label(role))
        header.addStretch(1)
        for label, tooltip, action in (
            ("Tất cả", "Chọn tất cả tên đang hiển thị", lambda: set_all(True)),
            ("Có data", "Chọn người có dữ liệu trong kỳ và Task hiện tại", select_active),
            ("Bỏ", "Bỏ chọn các tên đang hiển thị", lambda: set_all(False)),
        ):
            button = QPushButton(label)
            button.setObjectName("GhostBtn")
            button.setToolTip(tooltip)
            button.clicked.connect(action)
            header.addWidget(button)
        layout.addLayout(header)

        search = QLineEdit()
        search.setAccessibleName(f"Tìm {role}")
        search.setPlaceholderText(f"Tìm {role}... (Ctrl+F)")
        search.setClearButtonEnabled(True)
        layout.addWidget(search)

        names = QListWidget()
        names.setAccessibleName(role)
        names.setToolTip("Đánh dấu để chọn nhiều người; nhấp đúp tên để xem riêng.")
        names.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        names.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        names.itemChanged.connect(on_change)
        names.itemDoubleClicked.connect(isolate)
        search.textChanged.connect(lambda text: self._filter_people_search(names, text))
        layout.addWidget(names, 1)

        hint = QLabel("Mẹo: Nhấp đúp tên để xem riêng.")
        hint.setStyleSheet("font-size: 11px; color: #64748B; font-style: italic;")
        layout.addWidget(hint)
        return panel, search, names

    def _build_audit_page(self):
        page = QWidget()
        layout = QVBoxLayout(page)
        layout.setContentsMargins(0, 6, 0, 0)
        split = QSplitter(Qt.Orientation.Vertical)
        split.setChildrenCollapsible(False)

        top = QWidget()
        top_layout = QVBoxLayout(top)
        top_layout.setContentsMargins(0, 0, 0, 0)
        top_layout.setSpacing(5)
        filters = QHBoxLayout()
        filters.addWidget(QLabel("Lọc trạng thái:"))
        status = QComboBox()
        status.setMinimumWidth(190)
        for label, key in (
            ("Tất cả", ""), ("Không khớp", "mismatch"),
            ("Chưa có data / BBOX", "pending"), ("Khớp", "ok"),
        ):
            status.addItem(label, key)
        filters.addWidget(status)
        filters.addStretch(1)
        top_layout.addLayout(filters)
        summary = QTableWidget()
        self._setup_table(summary)
        summary.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        summary.horizontalHeader().setMinimumHeight(36)
        top_layout.addWidget(summary)
        split.addWidget(top)

        bottom = QWidget()
        bottom_layout = QVBoxLayout(bottom)
        bottom_layout.setContentsMargins(0, 6, 0, 0)
        bottom_layout.setSpacing(5)
        heading = QLabel("Chọn một dòng ở bảng trên để xem chi tiết.")
        heading.setTextFormat(Qt.TextFormat.PlainText)
        heading.setStyleSheet("font-weight: 700; font-size: 12.5px;")
        bottom_layout.addWidget(heading)
        details = QTableWidget()
        self._setup_table(details)
        details.setWordWrap(True)
        details.verticalHeader().setMinimumSectionSize(30)
        bottom_layout.addWidget(details)
        split.addWidget(bottom)
        split.setSizes([240, 360])
        layout.addWidget(split)
        return page, status, summary, heading, details

    def _setup_table(self, tbl: QTableWidget):
        tbl.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        tbl.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        tbl.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tbl.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        tbl.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        tbl.verticalHeader().setVisible(False)
        tbl.verticalHeader().setDefaultSectionSize(28)
        tbl.horizontalHeader().setStretchLastSection(False)
        tbl.horizontalHeader().setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)

    def _sync_people_filter_panel(self):
        show_inspector = self.tabs.currentWidget() is self.tbl_rev
        self.people_filters.setCurrentWidget(self.inspector_filter if show_inspector else self.annotator_filter)
        self.isolate_banner.setVisible(not show_inspector and self._isolated_annotator is not None)

    def _focus_people_search(self):
        edit = self.ed_inspector_search if self.people_filters.currentWidget() is self.inspector_filter else self.ed_search
        edit.setFocus()
        edit.selectAll()

    def _setup_shortcuts(self):
        QShortcut(QKeySequence("Ctrl+F"), self, activated=self._focus_people_search)
        QShortcut(QKeySequence("Ctrl+E"), self, activated=self.export_excel)
        QShortcut(QKeySequence("Ctrl+C"), self, activated=self.copy_current_table)
        QShortcut(QKeySequence("F5"), self, activated=self.reload_data)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, "toast"):
            self.toast.reposition()
        if hasattr(self, "tbl_daily_diffs") and self.tbl_daily_diffs.rowCount() > 0:
            self.tbl_daily_diffs.resizeRowsToContents()
        if hasattr(self, "tbl_rev_diffs") and self.tbl_rev_diffs.rowCount() > 0:
            self.tbl_rev_diffs.resizeRowsToContents()

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
        is_dark = self.theme["name"] == "dark"
        for u in event.mimeData().urls():
            p = u.toLocalFile()
            if not p.lower().endswith(".xlsx"):
                continue
            low = os.path.basename(p).lower()
            if "distribution" in low:
                self.card_dist.set_path(p, is_dark)
                self.w_card_dist.set_path(p, is_dark)
                changed = True
            elif "inspection" in low or "rework" in low:
                self.card_rw.set_path(p, is_dark)
                self.w_card_rw.set_path(p, is_dark)
                changed = True
            else:
                self.card_wl.set_path(p, is_dark)
                self.w_card_wl.set_path(p, is_dark)
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
        self.btn_sharepoint.setIcon(svg_icon("cloud", theme["accent"], 15))
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
        self.card_dist.set_path(self.card_dist.path(), is_dark)
        if hasattr(self, "w_card_wl"):
            self.w_card_wl.set_path(self.card_wl.path(), is_dark)
            self.w_card_rw.set_path(self.card_rw.path(), is_dark)
            self.w_card_dist.set_path(self.card_dist.path(), is_dark)

        if hasattr(self, "update_banner"):
            self.update_banner.apply_theme(theme)
        self._style_update_controls()

        if self.current_report:
            self._render_template1(self.current_report)
            self._render_template2(self.current_report)
            self._render_inspector_tab()
            self._render_ranking_tables(self.current_report)
            self._render_performance_tables(self.current_report)
            self._render_audit_table(self.current_report)
            self._render_reviewer_audit_table(self.current_report)

    def toggle_theme(self):
        next_t = DARK_THEME if self.theme["name"] == "light" else LIGHT_THEME
        self._apply_theme(next_t)
        self._save_current_settings()

    def _save_current_settings(self):
        self.settings["theme"] = self.theme["name"]
        self.settings["worklog_path"] = self.card_wl.path()
        self.settings["rework_path"] = self.card_rw.path()
        self.settings["dist_path"] = self.card_dist.path()
        self.settings["source_mode"] = self.seg_source.current_value()
        _save_settings(self.settings)

    # ------------------------------------------------------------------
    # FILE DISCOVERY & LOADING
    # ------------------------------------------------------------------
    def _restore_files_only(self):
        base = _app_base_dir()
        parent = os.path.dirname(base)
        od_dirs = _get_onedrive_dirs()

        od_wl, od_rw, od_dist = _find_default_files(od_dirs)
        local_wl, local_rw, local_dist = _find_default_files([base, parent, os.getcwd()])

        saved_wl = self.settings.get("worklog_path", "")
        saved_rw = self.settings.get("rework_path", "")
        saved_dist = self.settings.get("dist_path", "")

        if od_wl:
            wl = saved_wl if (saved_wl and os.path.isfile(saved_wl) and "onedrive" in saved_wl.lower()) else od_wl
        else:
            wl = saved_wl if (saved_wl and os.path.isfile(saved_wl)) else local_wl

        if od_rw:
            rw = saved_rw if (saved_rw and os.path.isfile(saved_rw) and "onedrive" in saved_rw.lower()) else od_rw
        else:
            rw = saved_rw if (saved_rw and os.path.isfile(saved_rw)) else local_rw

        if od_dist:
            dist_p = saved_dist if (saved_dist and os.path.isfile(saved_dist) and "onedrive" in saved_dist.lower()) else od_dist
        else:
            dist_p = saved_dist if (saved_dist and os.path.isfile(saved_dist)) else local_dist

        is_dark = self.theme["name"] == "dark"
        self.card_wl.set_path(wl, is_dark)
        self.card_rw.set_path(rw, is_dark)
        self.card_dist.set_path(dist_p, is_dark)
        self.w_card_wl.set_path(wl, is_dark)
        self.w_card_rw.set_path(rw, is_dark)
        self.w_card_dist.set_path(dist_p, is_dark)
        self._save_current_settings()

    def open_sharepoint_dialog(self):
        dlg = SharePointDialog(self.settings, self)
        res = dlg.exec()
        vals = dlg.values()
        self.settings.update(vals)
        self.w_ed_wl_url.setText(vals["sp_wl_url"])
        self.w_ed_rw_url.setText(vals["sp_rw_url"])
        self.w_ed_dist_url.setText(vals["sp_dist_url"])
        self._save_current_settings()
        self._refresh_welcome_sp_auth()
        if res == QDialog.DialogCode.Accepted and dlg.sync_requested:
            self._start_sharepoint_sync(vals["sp_wl_url"], vals["sp_rw_url"], vals["sp_dist_url"])

    def _start_sharepoint_sync(self, wl_url: str, rw_url: str, dist_url: str = "", _retried: bool = False):
        self._pending_sp_urls = (wl_url, rw_url, dist_url, _retried)
        self.btn_sharepoint.setEnabled(False)
        self.btn_reload.setEnabled(False)
        self.w_btn_sp_sync.setEnabled(False)
        self.w_btn_load.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(8)
        self.w_lbl_status.setText("Đang kết nối SharePoint...")
        self.status_bar.showMessage("Đang kết nối SharePoint...")

        self._sp_dl_thread = SharePointDownloadThread(wl_url, rw_url, _sp_cache_dir(), dist_url=dist_url)
        self._sp_dl_thread.progress.connect(self._on_load_progress)
        self._sp_dl_thread.finished_ok.connect(lambda wl_p, rw_p, dist_p: self._on_sp_sync_done(wl_p, rw_p, dist_p, "SharePoint"))
        self._sp_dl_thread.need_login.connect(self._on_sp_need_login)
        self._sp_dl_thread.failed.connect(self._on_sp_sync_fail)
        self._sp_dl_thread.start()

    def _on_sp_need_login(self):
        self.btn_sharepoint.setEnabled(True)
        self.btn_reload.setEnabled(True)
        self.w_btn_sp_sync.setEnabled(True)
        self.w_btn_load.setEnabled(True)
        self.progress.setVisible(False)
        self.w_lbl_status.setText("")
        wl_url, rw_url, dist_url, retried = getattr(self, "_pending_sp_urls", ("", "", "", False))
        if retried:
            QMessageBox.warning(self, "Chưa đăng nhập SharePoint", "Không xác thực được phiên SharePoint. Vui lòng thử đăng nhập lại.")
            return
        dlg = SharePointLoginDialog(extract_sp_site_url(wl_url, rw_url, dist_url), _sp_cache_dir(), self)
        res = dlg.exec()
        self._refresh_welcome_sp_auth()
        if res == QDialog.DialogCode.Accepted:
            self._start_sharepoint_sync(wl_url, rw_url, dist_url, _retried=True)

    def _on_sp_sync_done(self, wl_path: str, rw_path: str, dist_path: str, user: str):
        self.btn_sharepoint.setEnabled(True)
        self.btn_reload.setEnabled(True)
        self.w_btn_sp_sync.setEnabled(True)
        self.w_btn_load.setEnabled(True)
        is_dark = (self.theme["name"] == "dark")
        if wl_path:
            self.card_wl.set_path(wl_path, is_dark)
            self.w_card_wl.set_path(wl_path, is_dark)
        if rw_path:
            self.card_rw.set_path(rw_path, is_dark)
            self.w_card_rw.set_path(rw_path, is_dark)
        if dist_path:
            self.card_dist.set_path(dist_path, is_dark)
            self.w_card_dist.set_path(dist_path, is_dark)
        self._save_current_settings()
        self.toast.show_toast(f"✓ Đã tải từ {user}, đang phân tích dữ liệu...", "success", 3500)
        self.start_load_data()

    def _on_sp_sync_fail(self, err: str):
        self.btn_sharepoint.setEnabled(True)
        self.btn_reload.setEnabled(True)
        self.w_btn_sp_sync.setEnabled(True)
        self.w_btn_load.setEnabled(True)
        self.progress.setVisible(False)
        self.w_lbl_status.setText("")
        self.status_bar.showMessage("Lỗi tải dữ liệu từ SharePoint!")
        QMessageBox.critical(self, "Lỗi SharePoint", err)

    def switch_to_onedrive_source(self):
        od_dirs = _get_onedrive_dirs()
        od_wl, od_rw, od_dist = _find_default_files(od_dirs)
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
            self.w_card_wl.set_path(od_wl, is_dark)
        if od_rw:
            self.card_rw.set_path(od_rw, is_dark)
            self.w_card_rw.set_path(od_rw, is_dark)
        if od_dist:
            self.card_dist.set_path(od_dist, is_dark)
            self.w_card_dist.set_path(od_dist, is_dark)
        self._save_current_settings()
        self.toast.show_toast("☁ Đã chuyển sang nguồn OneDrive Sync!", "info")
        self.start_load_data()

    def _browse_file(self, kind: str, start_override: str = "", auto_load: bool = True):
        if kind == "wl":
            cur = self.card_wl.path()
            title = "Chọn file Work Log (.xlsx)"
        elif kind == "dist":
            cur = self.card_dist.path()
            title = "Chọn file Distribution Sheet (.xlsx)"
        else:
            cur = self.card_rw.path()
            title = "Chọn file Inspection & Review (.xlsx)"
        if start_override and os.path.isdir(start_override):
            start_dir = start_override
        elif cur and os.path.isdir(os.path.dirname(cur)):
            start_dir = os.path.dirname(cur)
        else:
            od_dirs = _get_onedrive_dirs()
            start_dir = next((d for d in od_dirs if os.path.isdir(d)), _app_base_dir())
        path, _ = QFileDialog.getOpenFileName(self, title, start_dir, "Excel Workbook (*.xlsx)")
        if path:
            if auto_load:
                self._set_file_and_reload(kind, path)
            else:
                self._set_file_only(kind, path)

    def _set_file_only(self, kind: str, path: str):
        is_dark = self.theme["name"] == "dark"
        if kind == "wl":
            self.card_wl.set_path(path, is_dark)
            self.w_card_wl.set_path(path, is_dark)
        elif kind == "dist":
            self.card_dist.set_path(path, is_dark)
            self.w_card_dist.set_path(path, is_dark)
        else:
            self.card_rw.set_path(path, is_dark)
            self.w_card_rw.set_path(path, is_dark)
        self._save_current_settings()

    def _set_file_and_reload(self, kind: str, path: str):
        self._set_file_only(kind, path)
        self.start_load_data()

    def reload_data(self):
        """Reload data: if currently using SharePoint-synced files, re-download fresh copies from SharePoint first."""
        wl = self.card_wl.path()
        rw = self.card_rw.path()
        dist_p = self.card_dist.path()
        sp_dir = os.path.normcase(os.path.abspath(_sp_cache_dir()))
        using_sp = any(
            p and os.path.normcase(os.path.abspath(p)).startswith(sp_dir)
            for p in (wl, rw, dist_p)
        )
        sp_wl_url = (self.settings.get("sp_wl_url") or DEFAULT_SP_WL_URL).strip()
        sp_rw_url = (self.settings.get("sp_rw_url") or DEFAULT_SP_RW_URL).strip()
        sp_dist_url = (self.settings.get("sp_dist_url") or DEFAULT_SP_DIST_URL).strip()
        if using_sp and (sp_wl_url or sp_rw_url):
            self._start_sharepoint_sync(sp_wl_url, sp_rw_url, sp_dist_url)
            return
        self.start_load_data()

    def start_load_data(self):
        wl = self.card_wl.path()
        rw = self.card_rw.path()
        dist_p = self.card_dist.path()
        if not wl and not rw:
            self.toast.show_toast("Vui lòng chọn hoặc kéo thả file Excel trước!", "warn")
            return
        if wl and not os.path.isfile(wl):
            QMessageBox.warning(self, "Không tìm thấy file", f"File Work Log không tồn tại:\n{wl}")
            return
        if rw and not os.path.isfile(rw):
            QMessageBox.warning(self, "Không tìm thấy file", f"File Review không tồn tại:\n{rw}")
            return
        if dist_p and not os.path.isfile(dist_p):
            dist_p = ""

        if self.loader_thread and self.loader_thread.isRunning():
            return

        self.btn_reload.setEnabled(False)
        self.btn_export.setEnabled(False)
        self.w_btn_load.setEnabled(False)
        self.w_btn_sp_sync.setEnabled(False)
        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.w_lbl_status.setText("Đang đọc dữ liệu & Excel Notes...")
        self.status_bar.showMessage("Đang đọc dữ liệu & Excel Notes...")

        self.loader_thread = LoaderThread(wl, rw, dist_p)
        self.loader_thread.progress.connect(self._on_load_progress)
        self.loader_thread.finished_ok.connect(self._on_load_done)
        self.loader_thread.failed.connect(self._on_load_fail)
        self.loader_thread.start()

    def _on_load_progress(self, pct: int, msg: str):
        self.progress.setValue(pct)
        self.w_lbl_status.setText(msg)
        self.status_bar.showMessage(msg)

    def _on_load_done(self, ds: LoadedDataset):
        self.dataset = ds
        self.btn_reload.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.w_btn_load.setEnabled(True)
        self.w_btn_sp_sync.setEnabled(True)
        self.progress.setVisible(False)
        self.w_lbl_status.setText("")
        self._show_main_workspace()
        self._isolated_annotator = None
        self.isolate_banner.setVisible(False)

        # Build TASK & Annotator lists
        self._updating_filters = True
        self.list_tasks.clear()
        task_priority = {
            "OD": 0, "SSEG": 1, "PFE": 2, "ANPR": 3,
            "OD_REW": 10, "SSEG_REW": 11, "PFE_REW": 12, "ANPR_REW": 13,
        }
        for t in sorted(ds.tasks, key=lambda x: (1 if x.endswith("_REW") else 0, task_priority.get(x, 99), x)):
            it = QListWidgetItem(t)
            it.setData(Qt.ItemDataRole.UserRole, t)
            it.setFlags(it.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            it.setCheckState(Qt.CheckState.Checked)
            self.list_tasks.addItem(it)

        inspector_keys = {normalize_annotator_key(rec.inspector) for rec in ds.rework_records} & FIRST_REVIEWER_KEYS
        inspector_names = {
            key: display_annotator_name(key) for key in sorted(inspector_keys, key=display_annotator_name)
        }
        self._populate_people_list(self.list_anns, ds.annotators, self.ed_search.text())
        self._populate_people_list(self.list_inspectors, inspector_names, self.ed_inspector_search.text())
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
            self._populate_week_months(select_ym=(s_d.year, s_d.month))

        msg = f"✓ Đã nạp {len(ds.annotators)} annotators · {len(ds.rework_records):,} dòng Review"
        self.status_bar.showMessage(msg)
        self.toast.show_toast(msg, "success")
        self.refresh_report()

    def _on_load_fail(self, err: str):
        self.btn_reload.setEnabled(True)
        self.btn_export.setEnabled(True)
        self.w_btn_load.setEnabled(True)
        self.w_btn_sp_sync.setEnabled(True)
        self.progress.setVisible(False)
        self.w_lbl_status.setText("")
        self.status_bar.showMessage("Lỗi khi nạp file Excel!")
        QMessageBox.critical(self, "Lỗi đọc file Excel", f"Không thể đọc file:\n{err}")

    # ------------------------------------------------------------------
    # FILTER INTERACTIONS
    # ------------------------------------------------------------------
    def _on_option_changed(self, _val):
        self._save_current_settings()
        self.refresh_report()

    def _on_date_changed(self):
        self._sync_period_button_states()
        self.refresh_report()

    @staticmethod
    def _weeks_in_month(year: int, month: int):
        """Return list of full (Mon, Sun) weeks whose Monday falls in (year, month).
        If the last week extends past the end of the month, those remaining days of the week are included."""
        first_day = datetime.date(year, month, 1)
        next_m = (first_day.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
        last_day = next_m - datetime.timedelta(days=1)
        # First Monday in the month
        days_to_mon = (7 - first_day.weekday()) % 7
        cur = first_day + datetime.timedelta(days=days_to_mon)
        weeks = []
        while cur <= last_day:
            weeks.append((cur, cur + datetime.timedelta(days=6)))
            cur += datetime.timedelta(days=7)
        return weeks

    @staticmethod
    def _month_bounds(year: int, month: int):
        first_day = datetime.date(year, month, 1)
        next_m = (first_day.replace(day=28) + datetime.timedelta(days=4)).replace(day=1)
        return first_day, next_m - datetime.timedelta(days=1)

    def _populate_week_months(self, select_ym=None):
        """Populate month dropdown for week-of-month selection."""
        months = []
        if self.dataset and self.dataset.min_date and self.dataset.max_date:
            y, m = self.dataset.max_date.year, self.dataset.max_date.month
            min_ym = (self.dataset.min_date.year, self.dataset.min_date.month)
            while (y, m) >= min_ym:
                months.append((y, m))
                m -= 1
                if m == 0:
                    m = 12
                    y -= 1
        else:
            qd = self.dt_start.date()
            ref_y, ref_m = qd.year(), qd.month()
            for offset in range(12):
                m = ref_m - offset
                y = ref_y
                while m <= 0:
                    m += 12
                    y -= 1
                months.append((y, m))

        self.cb_week_month.blockSignals(True)
        self.cb_week_month.clear()
        for y, m in months:
            self.cb_week_month.addItem(f"Tháng {m:02d}/{y}", (y, m))
        target_idx = 0
        if select_ym:
            for i in range(self.cb_week_month.count()):
                if self.cb_week_month.itemData(i) == select_ym:
                    target_idx = i
                    break
        if self.cb_week_month.count() > 0:
            self.cb_week_month.setCurrentIndex(target_idx)
        self.cb_week_month.blockSignals(False)
        self._rebuild_week_buttons()

    def _rebuild_week_buttons(self):
        while self.week_btns_layout.count():
            item = self.week_btns_layout.takeAt(0)
            w = item.widget()
            if w:
                w.deleteLater()
        self._week_buttons = []

        ym = self.cb_week_month.currentData()
        if not ym:
            return
        year, month = ym
        ms, me = self._month_bounds(year, month)
        self.btn_full_month.setToolTip(
            f"Cả tháng {month:02d}/{year}: {ms.strftime('%d/%m/%Y')} → {me.strftime('%d/%m/%Y')}"
        )
        weeks = self._weeks_in_month(year, month)
        cols = 3
        for idx, (s_d, e_d) in enumerate(weeks, start=1):
            btn = QPushButton(f"Tuần {idx}")
            btn.setObjectName("PresetBtn")
            btn.setCheckable(True)
            btn.setCursor(Qt.CursorShape.PointingHandCursor)
            btn.setToolTip(
                f"Tuần {idx} tháng {month:02d}/{year}: "
                f"{s_d.strftime('%d/%m/%Y')} → {e_d.strftime('%d/%m/%Y')}"
            )
            btn.clicked.connect(lambda _c=False, s=s_d, e=e_d: self._apply_date_range(s, e))
            r_i, c_i = divmod(idx - 1, cols)
            self.week_btns_layout.addWidget(btn, r_i, c_i)
            self._week_buttons.append((btn, s_d, e_d))
        self._sync_period_button_states()

    def _apply_selected_month(self):
        ym = self.cb_week_month.currentData()
        if not ym:
            return
        s_d, e_d = self._month_bounds(ym[0], ym[1])
        self._apply_date_range(s_d, e_d)

    def _sync_period_button_states(self):
        cur_s, cur_e = self._get_period_bounds()
        ym = self.cb_week_month.currentData()
        is_full_month = False
        if ym:
            ms, me = self._month_bounds(ym[0], ym[1])
            is_full_month = (cur_s == ms and cur_e == me)
        self.btn_full_month.blockSignals(True)
        self.btn_full_month.setChecked(is_full_month)
        self.btn_full_month.blockSignals(False)
        for btn, s_d, e_d in self._week_buttons:
            btn.blockSignals(True)
            btn.setChecked(cur_s == s_d and cur_e == e_d)
            btn.blockSignals(False)

    def _apply_date_range(self, s: datetime.date, e: datetime.date):
        self.dt_start.blockSignals(True)
        self.dt_end.blockSignals(True)
        self.dt_start.setDate(QDate(s.year, s.month, s.day))
        self.dt_end.setDate(QDate(e.year, e.month, e.day))
        self.dt_start.blockSignals(False)
        self.dt_end.blockSignals(False)
        self._sync_period_button_states()
        self.refresh_report()

    def _get_period_bounds(self):
        qd_s, qd_e = self.dt_start.date(), self.dt_end.date()
        s_d = datetime.date(qd_s.year(), qd_s.month(), qd_s.day())
        e_d = datetime.date(qd_e.year(), qd_e.month(), qd_e.day())
        return (e_d, s_d) if s_d > e_d else (s_d, e_d)

    def _set_all_tasks(self, checked: bool):
        self._updating_filters = True
        state = Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked
        for i in range(self.list_tasks.count()):
            self.list_tasks.item(i).setCheckState(state)
        self._updating_filters = False
        self.refresh_report()

    def _on_task_item_changed(self, _item):
        if not self._updating_filters:
            self.refresh_report()

    def _get_checked_tasks(self) -> Optional[Set[str]]:
        if self.list_tasks.count() == 0:
            return None
        res = set()
        for i in range(self.list_tasks.count()):
            it = self.list_tasks.item(i)
            if it.checkState() == Qt.CheckState.Checked:
                res.add(it.data(Qt.ItemDataRole.UserRole))
        return res

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
            self.dataset, s_d, e_d, None, self._get_checked_tasks(),
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

    def _populate_people_list(self, names: QListWidget, display_names: Dict[str, str], search: str):
        names.clear()
        for key, display in display_names.items():
            item = QListWidgetItem(display)
            item.setData(Qt.ItemDataRole.UserRole, key)
            item.setData(Qt.ItemDataRole.UserRole + 1, display)
            item.setFlags(item.flags() | Qt.ItemFlag.ItemIsUserCheckable)
            item.setCheckState(Qt.CheckState.Checked)
            names.addItem(item)
        self._filter_people_search(names, search)

    def _filter_people_search(self, names: QListWidget, text: str):
        query = text.strip().casefold()
        key_query = normalize_annotator_key(text)
        for i in range(names.count()):
            item = names.item(i)
            display = item.data(Qt.ItemDataRole.UserRole + 1).casefold()
            key = item.data(Qt.ItemDataRole.UserRole)
            item.setHidden(query not in display and (not key_query or key_query not in key))

    def _set_all_inspectors(self, checked: bool):
        self._set_inspector_checks(lambda _key: checked, visible_only=True)

    def _select_active_inspectors(self):
        groups = self._period_report.inspector_groups if self._period_report else []
        active = {group.inspector_key for group in groups}
        self._set_inspector_checks(lambda key: key in active)

    def _isolate_single_inspector(self, item: QListWidgetItem):
        key = item.data(Qt.ItemDataRole.UserRole)
        self._set_inspector_checks(lambda other: other == key)

    def _set_inspector_checks(self, checked_for_key, *, visible_only: bool = False):
        self._updating_filters = True
        for i in range(self.list_inspectors.count()):
            item = self.list_inspectors.item(i)
            if not visible_only or not item.isHidden():
                checked = checked_for_key(item.data(Qt.ItemDataRole.UserRole))
                item.setCheckState(Qt.CheckState.Checked if checked else Qt.CheckState.Unchecked)
        self._updating_filters = False
        self._render_inspector_tab()

    def _on_inspector_item_changed(self, _item):
        if not self._updating_filters:
            self._render_inspector_tab()

    def _selected_inspector_groups(self) -> List[InspectorSummary]:
        if not self._period_report:
            return []
        checked = self._get_checked_people(self.list_inspectors)
        return [group for group in self._period_report.inspector_groups if group.inspector_key in checked]

    def _on_ann_item_changed(self, _item):
        if not self._updating_filters:
            checked = self._get_checked_annotators()
            if len(checked) != 1:
                self._isolated_annotator = None
                self.isolate_banner.setVisible(False)
            self.refresh_report()

    def _get_checked_annotators(self) -> Set[str]:
        return self._get_checked_people(self.list_anns)

    def _get_checked_people(self, names: QListWidget) -> Set[str]:
        checked = set()
        for i in range(names.count()):
            item = names.item(i)
            if item.checkState() == Qt.CheckState.Checked:
                checked.add(item.data(Qt.ItemDataRole.UserRole))
        return checked

    # ------------------------------------------------------------------
    # REPORT COMPUTATION & RENDERING
    # ------------------------------------------------------------------
    def refresh_report(self):
        if not self.dataset:
            return
        s_d, e_d = self._get_period_bounds()
        src = self.seg_source.current_value()

        selected_tasks = self._get_checked_tasks()
        rep_period_all = build_report(self.dataset, s_d, e_d, None, selected_tasks, source_mode=src)
        self._period_report = rep_period_all
        fr_by_ann = {}
        for r in rep_period_all.all_rows:
            fr_by_ann[r.annotator_key] = fr_by_ann.get(r.annotator_key, 0.0) + r.work_load
        fr_by_insp = {group.inspector_key: group.total_work_load for group in rep_period_all.inspector_groups}
        self._updating_filters = True
        for names, frames_by_key in ((self.list_anns, fr_by_ann), (self.list_inspectors, fr_by_insp)):
            for i in range(names.count()):
                item = names.item(i)
                key = item.data(Qt.ItemDataRole.UserRole)
                display = item.data(Qt.ItemDataRole.UserRole + 1)
                frames = frames_by_key.get(key, 0.0)
                item.setText(f"{display}   ({_fmt_int(frames)} fr)" if frames > 0 else display)
        self._updating_filters = False

        # Compute filtered report
        self.current_report = build_report(
            self.dataset,
            start_date=s_d,
            end_date=e_d,
            selected_annotators=self._get_checked_annotators(),
            selected_tasks=selected_tasks,
            source_mode=src,
        )
        self.lbl_period_badge.setText(f"Kỳ báo cáo: {s_d.strftime('%Y.%m.%d')} ~ {e_d.strftime('%Y.%m.%d')}")

        self._render_template1(self.current_report)
        self._render_template2(self.current_report)
        self._render_inspector_tab()
        self._render_ranking_tables(self.current_report)
        self._render_performance_tables(self.current_report)
        self._linked_audit_key = None
        self._render_audit_table(self.current_report)
        self._render_reviewer_audit_table(self.current_report)

    def _make_item(
        self,
        text: str,
        align=Qt.AlignmentFlag.AlignCenter,
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

    def _auto_fit_columns(
        self,
        tbl: QTableWidget,
        headers: list,
        header_height: int = 44,
        stretch_last: bool = True,
        min_widths: Optional[Dict[int, int]] = None,
    ):
        """
        Size every non-last column using Qt's exact ResizeToContents (with wordWrap=False so
        content expands the column instead of wrapping or eliding with '...'), then lock those
        columns and Stretch the last column flush to the window edge with wordWrap=True.
        """
        hdr = tbl.horizontalHeader()
        hdr.setDefaultAlignment(Qt.AlignmentFlag.AlignCenter)
        hdr.setFixedHeight(header_height)
        hdr.setStretchLastSection(False)

        # Temporarily disable word-wrap so Qt measures exact single-line width (including CSS padding & bold/emoji)
        tbl.setWordWrap(False)
        n_cols = len(headers)
        for c in range(n_cols):
            if stretch_last and c == n_cols - 1:
                continue
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.ResizeToContents)

        # Force Qt to compute ResizeToContents widths using live stylesheet font & item padding
        tbl.resizeColumnsToContents()

        h_fm = QFontMetrics(hdr.font())
        for c, h_text in enumerate(headers):
            if stretch_last and c == n_cols - 1:
                if min_widths and c in min_widths:
                    hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
                    tbl.setColumnWidth(c, min_widths[c])
                    hdr.setStretchLastSection(True)
                else:
                    hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Stretch)
                continue

            lines = str(h_text).split("\n")
            hdr_w = max(h_fm.horizontalAdvance(ln) for ln in lines) + 22
            qt_w = tbl.columnWidth(c) + 8
            w = max(qt_w, hdr_w)
            if min_widths and c in min_widths:
                w = max(w, min_widths[c])
            hdr.setSectionResizeMode(c, QHeaderView.ResizeMode.Fixed)
            tbl.setColumnWidth(c, w)

        tbl.setWordWrap(True)
        tbl.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

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
                lines.append("⚠ Block không khớp (đối chiếu theo Frame Start/End):")
                for d in r.real_diffs[:5]:
                    rw_dt_s = f" | Review {d.rw_date.strftime('%m/%d')}" if d.rw_date else ""
                    rng_s = f" | {d.frame_range_display}" if d.frame_range_display != "—" else ""
                    lines.append(
                        f"  • Work Log {d.date.strftime('%Y-%m-%d')}{rw_dt_s} [{d.task}{rng_s}]: "
                        f"Frame {_fmt_int(d.wl_frames)}/{_fmt_int(d.rw_frames)} | "
                        f"BBOX {_fmt_int(d.wl_bbox)}/{_fmt_int(d.rw_bbox)} — {d.reason}"
                    )
                if len(r.real_diffs) > 5:
                    lines.append(f"  ... và {len(r.real_diffs) - 5} block không khớp khác")
            if r.pending_diffs:
                lines.append("⏳ Block CHƯA CÓ DATA / CHƯA NHẬP BBOX:")
                for d in r.pending_diffs[:5]:
                    rng_s = f" | {d.frame_range_display}" if d.frame_range_display != "—" else ""
                    lines.append(
                        f"  • Work Log {d.date.strftime('%Y-%m-%d')} [{d.task}{rng_s}]: "
                        f"Frame {_fmt_int(d.wl_frames)}/{_fmt_int(d.rw_frames)} | "
                        f"BBOX {_fmt_int(d.wl_bbox)}/{_fmt_int(d.rw_bbox)} — {d.reason}"
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
        ]
        total_rows = sum(len(g.rows) for g in groups)
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.setRowCount(total_rows)

        hl = self.chk_highlight.isChecked()
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
                    f"Worker (Lần 1 + Lần 2):\n"
                    f"  • Lỗi BBOX:  {_fmt_int(r.total_err_bbox)} (Lần 1: {_fmt_int(r.err_first_bbox)} + Lần 2: {_fmt_int(r.err_second_bbox)})\n"
                    f"  • Lỗi Frame: {_fmt_int(r.total_err_frame)} (Lần 1: {_fmt_int(r.err_first)} + Lần 2: {_fmt_int(r.err_second)})"
                )
                tip = self._build_mismatch_tooltip(r, prefix=err_info)

                if idx == 0:
                    tbl.setItem(row_i, 0, self._make_item(disp_task, Qt.AlignmentFlag.AlignCenter, True, t["task_badge_fg"], t["task_badge_bg"]))

                ann_lbl = f"{prefix_ic}{r.annotator}"
                tbl.setItem(row_i, 1, self._make_item(ann_lbl, Qt.AlignmentFlag.AlignCenter, False, None, row_bg, tip, user_data=(r.annotator_key, r.task)))
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

            if n > 1:
                tbl.setSpan(start_r, 0, n, 1)
            cur += n

        delegate.set_group_bottom_rows(group_bottoms)
        hdr = tbl.horizontalHeader()
        hdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        tbl.setColumnWidth(0, 70)
        tbl.setColumnWidth(1, 96)
        hdr.setFixedHeight(44)

        # Render per-TASK summary table below
        stbl = self.tbl_t1_summary
        s_headers = [
            "TASK",
            "Số người",
            "Tổng Frame",
            "Tổng BBOX",
            "Avg\nFrame/h",
            "Avg\nBBOX/h",
            "Err Rate\nBBOX (%)",
            "Err Rate\nFrame (%)",
        ]
        stbl.setColumnCount(len(s_headers))
        stbl.setHorizontalHeaderLabels(s_headers)
        stbl.setRowCount(len(groups))
        task_avg_bg = t["tbl_kpi_bg"]
        avg_align = Qt.AlignmentFlag.AlignCenter
        for idx, grp in enumerate(groups):
            disp_task = grp.task[:-4] if (strip_rew_suffix and grp.task.endswith("_REW")) else grp.task
            n = len(grp.rows)
            avg_tip = (
                f"Chỉ số trung bình & tỉ lệ lỗi của toàn bộ nhóm TASK {disp_task} ({n} người):\n"
                f"  • Tổng Frame: {_fmt_int(grp.total_work_load)} | Tổng BBOX: {_fmt_int(grp.total_work_bbox)} | Tổng Time: {_fmt_time(grp.total_work_time)}h\n"
                f"  • Tổng lỗi BBOX: {_fmt_int(grp.total_err_bbox)} ({_fmt_f2(grp.total_task_err_bbox_rate)}%) | Tổng lỗi Frame: {_fmt_int(grp.total_err_frame)} ({_fmt_f2(grp.total_task_err_rate)}%)"
            )
            stbl.setItem(idx, 0, self._make_item(disp_task, avg_align, True, t["task_badge_fg"], t["task_badge_bg"], tooltip=avg_tip))
            stbl.setItem(idx, 1, self._make_item(str(n), avg_align, False, None, task_avg_bg, tooltip=avg_tip))
            stbl.setItem(idx, 2, self._make_item(_fmt_int(grp.total_work_load), bg=task_avg_bg, tooltip=avg_tip))
            stbl.setItem(idx, 3, self._make_item(_fmt_int(grp.total_work_bbox), bg=task_avg_bg, tooltip=avg_tip))
            stbl.setItem(idx, 4, self._make_item(_fmt_f2(grp.avg_frame_per_hour), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
            stbl.setItem(idx, 5, self._make_item(_fmt_f2(grp.avg_bbox_per_hour), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
            stbl.setItem(idx, 6, self._make_item(_fmt_f2(grp.total_task_err_bbox_rate), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
            stbl.setItem(idx, 7, self._make_item(_fmt_f2(grp.total_task_err_rate), avg_align, True, t["danger"], bg=task_avg_bg, tooltip=avg_tip))
        shdr = stbl.horizontalHeader()
        shdr.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        shdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        stbl.setColumnWidth(0, 96)
        shdr.setMinimumHeight(46)

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

            tbl.setItem(row_i, 0, self._make_item(lbl, Qt.AlignmentFlag.AlignCenter, True, None, row_bg, tip, user_data=(r.annotator_key, None)))
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

    def _render_inspector_tab(self):
        tbl = self.tbl_rev
        tbl.clearSpans()
        tbl.clearContents()
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
        groups = self._selected_inspector_groups()
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
                    f"Inspector {grp.inspector} chấm bài của {r.annotator}:\n"
                    f"  • Lỗi Lần 1 (Review): {_fmt_int(r.err_first)} Frame | {_fmt_int(r.err_first_bbox)} BBOX\n"
                    f"  • Lỗi Lần 2: {_fmt_int(r.err_second)} Frame | {_fmt_int(r.err_second_bbox)} BBOX"
                )
                if idx == 0:
                    tbl.setItem(row_i, 0, self._make_item(grp.inspector, Qt.AlignmentFlag.AlignCenter, True, t["task_badge_fg"], t["task_badge_bg"]))
                tbl.setItem(row_i, 1, self._make_item(r.annotator, Qt.AlignmentFlag.AlignCenter, False, None, None, tip, user_data=(r.annotator_key, None)))
                tbl.setItem(row_i, 2, self._make_item(_fmt_int(r.work_load), tooltip=tip))
                tbl.setItem(row_i, 3, self._make_item(_fmt_int(r.work_bbox), tooltip=tip))
                tbl.setItem(row_i, 4, self._make_item(_fmt_time(r.work_time), tooltip=tip))
                tbl.setItem(row_i, 5, self._make_item(_fmt_int(r.err_first), tooltip=tip))
                tbl.setItem(row_i, 6, self._make_item(_fmt_int(r.err_second), tooltip=tip))

            sum_tip = (
                f"Tổng hợp Inspector {grp.inspector} ({n} Annotators):\n"
                f"  • Tổng Load: {_fmt_int(grp.total_work_load)} Frame | {_fmt_int(grp.total_work_bbox)} BBOX | {_fmt_time(grp.total_work_time)}h\n"
                f"  • Lỗi Lần 2: {_fmt_int(grp.total_err_second)} Frame ({_fmt_f2(grp.err_rate_frame)}%) | {_fmt_int(grp.total_err_second_bbox)} BBOX ({_fmt_f2(grp.err_rate_bbox)}%)\n"
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
            tbl.setItem(tr, 1, self._make_item(f"{len(groups)} Inspector", Qt.AlignmentFlag.AlignCenter, True, bg=tot_bg))
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

        workers = sorted(
            [r for r in rep.annotator_rows if not r.is_reviewer and r.work_load > 0],
            key=lambda r: (r.total_err_rate, -r.work_load, r.annotator.upper()),
        )
        reviewers = sorted(
            [g for g in rep.inspector_groups if g.total_work_load > 0],
            key=lambda g: (g.err_rate_frame, -g.total_work_load, g.inspector.upper()),
        )

        self.lbl_rank_worker.setText(f"🥇 Xếp hạng Annotator / Worker ({len(workers)} người — Tỉ lệ lỗi Frame thấp nhất)")
        self.lbl_rank_reviewer.setText(f"🥇 Inspector Lần 1 ({len(reviewers)} người) · Theo Annotator đang chọn")

        # Populate Worker Ranking (Frame Error Rate only)
        w_tbl = self.tbl_rank_worker
        w_headers = [
            "Hạng", "Annotator", "Err Rate\nFrame (%)",
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
            tip = (
                f"Hạng #{rank_num} — {r.annotator} ({r.task}):\n"
                f"  • Sản lượng: {_fmt_int(r.work_load)} Frame | {_fmt_int(r.work_bbox)} BBOX | {_fmt_time(r.work_time)}h\n"
                f"  • Lỗi Frame: {_fmt_int(r.total_err_frame)} ({_fmt_f2(r.total_err_rate)}%)"
            )
            w_tbl.setItem(idx, 0, self._make_item(rank_lbl, Qt.AlignmentFlag.AlignCenter, True, rate_fg if is_top3 else None, bg, tip))
            w_tbl.setItem(idx, 1, self._make_item(r.annotator, Qt.AlignmentFlag.AlignCenter, True, None, bg, tip, user_data=(r.annotator_key, None)))
            w_tbl.setItem(idx, 2, self._make_item(_fmt_f2(r.total_err_rate), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
        wh = w_tbl.horizontalHeader()
        wh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        wh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        w_tbl.setColumnWidth(0, 64)
        wh.setFixedHeight(42)

        # Populate Inspector Lần 1 Ranking (Frame Error Rate only)
        r_tbl = self.tbl_rank_reviewer
        r_headers = [
            "Hạng", "Inspector", "Err Rate\nFrame (%)",
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
                f"Hạng #{rank_num} — Inspector {g.inspector}:\n"
                f"  • Sản lượng: {_fmt_int(g.total_work_load)} Frame | {_fmt_int(g.total_work_bbox)} BBOX | {_fmt_time(g.total_work_time)}h\n"
                f"  • Lỗi Frame: {_fmt_int(g.total_err_second)} ({_fmt_f2(g.err_rate_frame)}%)\n"
                f"  • Recall Lần 1: {g.first_inspector_recall:.0f}%"
            )
            r_tbl.setItem(idx, 0, self._make_item(rank_lbl, Qt.AlignmentFlag.AlignCenter, True, rate_fg if is_top3 else None, bg, tip))
            r_tbl.setItem(idx, 1, self._make_item(g.inspector, Qt.AlignmentFlag.AlignCenter, True, None, bg, tip))
            r_tbl.setItem(idx, 2, self._make_item(_fmt_f2(g.err_rate_frame), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
        rh = r_tbl.horizontalHeader()
        rh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        rh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        r_tbl.setColumnWidth(0, 64)
        rh.setFixedHeight(42)

    @staticmethod
    def _score_performance(items, get_load, get_bbox, get_err, w_err: float, w_vol: float):
        """Return list of (item, total_score, quality_score, volume_score) sorted by highest score."""
        if not items:
            return []
        errs = [get_err(x) for x in items]
        min_e, max_e = min(errs), max(errs)
        max_load = max((get_load(x) for x in items), default=0.0)
        max_bbox = max((get_bbox(x) for x in items), default=0.0)
        scored = []
        for x in items:
            err = get_err(x)
            q_score = ((max_e - err) / (max_e - min_e) * 100.0) if max_e > min_e else 100.0
            vf = (get_load(x) / max_load * 100.0) if max_load > 0 else 0.0
            vb = (get_bbox(x) / max_bbox * 100.0) if max_bbox > 0 else vf
            v_score = 0.5 * vf + 0.5 * vb
            total = w_err * q_score + w_vol * v_score
            scored.append((x, total, q_score, v_score))
        scored.sort(key=lambda tup: (-tup[1], get_err(tup[0]), -get_load(tup[0])))
        return scored

    def _render_performance_tables(self, rep: ReportResult):
        t = self.theme
        mode = self.seg_perf_weight.current_value()
        w_err, w_vol = {"50_50": (0.5, 0.5), "60_40": (0.6, 0.4), "40_60": (0.4, 0.6)}.get(mode, (0.5, 0.5))

        workers_raw = [r for r in rep.annotator_rows if not r.is_reviewer and r.work_load > 0]
        reviewers_raw = [g for g in rep.inspector_groups if g.total_work_load > 0]

        w_scored = self._score_performance(
            workers_raw, lambda r: r.work_load, lambda r: r.work_bbox, lambda r: r.total_err_rate, w_err, w_vol
        )
        r_scored = self._score_performance(
            reviewers_raw, lambda g: g.total_work_load, lambda g: g.total_work_bbox, lambda g: g.err_rate_frame, w_err, w_vol
        )

        self.lbl_perf_worker.setText(f"⭐ Hiệu suất Annotator / Worker ({len(w_scored)} người — Tỉ lệ lỗi thấp & Sản lượng cao)")
        self.lbl_perf_reviewer.setText(f"⭐ Hiệu suất Inspector Lần 1 ({len(r_scored)} người — Tỉ lệ lỗi thấp & Sản lượng cao)")

        medals = {1: "🥇 1", 2: "🥈 2", 3: "🥉 3"}

        # Populate Worker Performance Table
        w_tbl = self.tbl_perf_worker
        w_headers = ["Hạng", "Annotator", "Work Load\n(Frame)", "Work\nBBOX", "Err Rate\nFrame (%)", "Điểm\nTổng"]
        w_tbl.setColumnCount(len(w_headers))
        w_tbl.setHorizontalHeaderLabels(w_headers)
        w_tbl.setRowCount(len(w_scored))
        for idx, (r, tot, q_sc, v_sc) in enumerate(w_scored):
            rank_num = idx + 1
            rank_lbl = medals.get(rank_num, str(rank_num))
            is_top3 = (rank_num <= 3)
            bg = t["success_soft"] if is_top3 else None
            score_fg = t["success"] if is_top3 else t["accent_text"]
            rate_fg = t["success"] if r.total_err_rate <= 1.0 else t["danger"]
            tip = (
                f"Hạng #{rank_num} — {r.annotator} ({r.task}):\n"
                f"  • Điểm Tổng: {tot:.1f} / 100 (Lỗi thấp: {q_sc:.1f} × {int(w_err*100)}% + Sản lượng: {v_sc:.1f} × {int(w_vol*100)}%)\n"
                f"  • Sản lượng: {_fmt_int(r.work_load)} Frame | {_fmt_int(r.work_bbox)} BBOX | {_fmt_time(r.work_time)}h\n"
                f"  • Tốc độ: {_fmt_f2(r.frame_per_hour)} fr/h | {_fmt_f2(r.bbox_per_hour)} bbox/h\n"
                f"  • Lỗi Frame: {_fmt_int(r.total_err_frame)} ({_fmt_f2(r.total_err_rate)}%)"
            )
            w_tbl.setItem(idx, 0, self._make_item(rank_lbl, Qt.AlignmentFlag.AlignCenter, True, score_fg if is_top3 else None, bg, tip))
            w_tbl.setItem(idx, 1, self._make_item(r.annotator, Qt.AlignmentFlag.AlignCenter, True, None, bg, tip, user_data=(r.annotator_key, None)))
            w_tbl.setItem(idx, 2, self._make_item(_fmt_int(r.work_load), bold=is_top3, bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 3, self._make_item(_fmt_int(r.work_bbox), bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 4, self._make_item(_fmt_f2(r.total_err_rate), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
            w_tbl.setItem(idx, 5, self._make_item(f"{tot:.1f}", Qt.AlignmentFlag.AlignCenter, True, fg=score_fg, bg=bg, tooltip=tip))
        wh = w_tbl.horizontalHeader()
        wh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        wh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        w_tbl.setColumnWidth(0, 62)
        wh.setFixedHeight(42)

        # Populate Inspector Performance Table
        r_tbl = self.tbl_perf_reviewer
        r_headers = ["Hạng", "Inspector", "Work Load\n(Frame)", "Work\nBBOX", "Err Rate\nFrame (%)", "Điểm\nTổng"]
        r_tbl.setColumnCount(len(r_headers))
        r_tbl.setHorizontalHeaderLabels(r_headers)
        r_tbl.setRowCount(len(r_scored))
        for idx, (g, tot, q_sc, v_sc) in enumerate(r_scored):
            rank_num = idx + 1
            rank_lbl = medals.get(rank_num, str(rank_num))
            is_top3 = (rank_num <= 3)
            bg = t["success_soft"] if is_top3 else None
            score_fg = t["success"] if is_top3 else t["accent_text"]
            rate_fg = t["success"] if g.err_rate_frame <= 0.5 else t["danger"]
            tip = (
                f"Hạng #{rank_num} — Inspector {g.inspector}:\n"
                f"  • Điểm Tổng: {tot:.1f} / 100 (Lỗi thấp: {q_sc:.1f} × {int(w_err*100)}% + Sản lượng: {v_sc:.1f} × {int(w_vol*100)}%)\n"
                f"  • Sản lượng: {_fmt_int(g.total_work_load)} Frame | {_fmt_int(g.total_work_bbox)} BBOX | {_fmt_time(g.total_work_time)}h\n"
                f"  • Lỗi Frame: {_fmt_int(g.total_err_second)} ({_fmt_f2(g.err_rate_frame)}%) | Recall: {g.first_inspector_recall:.0f}%"
            )
            r_tbl.setItem(idx, 0, self._make_item(rank_lbl, Qt.AlignmentFlag.AlignCenter, True, score_fg if is_top3 else None, bg, tip))
            r_tbl.setItem(idx, 1, self._make_item(g.inspector, Qt.AlignmentFlag.AlignCenter, True, None, bg, tip))
            r_tbl.setItem(idx, 2, self._make_item(_fmt_int(g.total_work_load), bold=is_top3, bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 3, self._make_item(_fmt_int(g.total_work_bbox), bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 4, self._make_item(_fmt_f2(g.err_rate_frame), bold=True, fg=rate_fg, bg=bg, tooltip=tip))
            r_tbl.setItem(idx, 5, self._make_item(f"{tot:.1f}", Qt.AlignmentFlag.AlignCenter, True, fg=score_fg, bg=bg, tooltip=tip))
        rh = r_tbl.horizontalHeader()
        rh.setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        rh.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        r_tbl.setColumnWidth(0, 62)
        rh.setFixedHeight(42)

    def _selected_audit_summary(self, table: QTableWidget):
        item = table.item(table.currentRow(), 0)
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    def _render_audit_summary(self, table: QTableWidget, rows, status: str, *, inspector: bool, select_key=None):
        previous = self._selected_audit_summary(table)
        key_to_restore = select_key if select_key is not None else (previous[0] if previous else None)
        headers = ["Task", "Inspector" if inspector else "Annotator", "Trạng thái", "Dòng cần xử lý"]
        visible_rows = [row for row in rows if not status or row.status_kind == status]
        t = self.theme
        restore_index = None

        with QSignalBlocker(table):
            table.clearSelection()
            table.setCurrentItem(None)
            table.clearContents()
            table.setColumnCount(len(headers))
            table.setHorizontalHeaderLabels(headers)
            table.setRowCount(len(visible_rows))
            for i, row in enumerate(visible_rows):
                person, person_key = (row.inspector, row.inspector_key) if inspector else (row.annotator, row.annotator_key)
                key = (person_key, row.task)
                issues = [d for d in row.daily_diffs if d.kind in ("mismatch", "pending_rework")]
                issues.sort(key=lambda d: d.is_pending)
                if not inspector:
                    issues = _merge_job_id_diff_rows(issues)
                if row.status_kind == "mismatch":
                    label, fg, tint = "⚠ KHÔNG KHỚP", t["warn"], t["warn_soft"]
                elif row.status_kind == "pending":
                    only_bbox = not inspector and all(d.is_missing_bbox for d in issues)
                    label = "⏳ Chưa nhập BBOX" if only_bbox else "⏳ Chưa có data"
                    fg, tint = t["accent_text"], t["accent_soft"]
                else:
                    label, fg, tint = "✓ Khớp", t["success"], t["success_soft"]
                tip = "Mở từ báo cáo Inspector; không đổi bộ lọc Annotator." if not inspector and key == self._linked_audit_key else ""
                table.setItem(i, 0, self._make_item(row.task, bold=True, tooltip=tip, user_data=(key, person, issues)))
                table.setItem(i, 1, self._make_item(person, bold=True, tooltip=tip))
                table.setItem(i, 2, self._make_item(label, bold=True, fg=fg, bg=tint))
                table.setItem(i, 3, self._make_item(_fmt_int(len(issues)), bold=bool(issues), fg=fg if issues else t["text_muted"]))
                if key == key_to_restore:
                    restore_index = i
            if restore_index is not None:
                table.setCurrentCell(restore_index, 0)

    def _on_audit_status_changed(self):
        self._linked_audit_key = None
        if self.current_report:
            self._render_audit_table(self.current_report)

    def _render_audit_table(self, rep: ReportResult, select_key=None):
        rows = rep.all_rows
        if self._linked_audit_key and self._period_report:
            linked = next((row for row in self._period_report.all_rows if (row.annotator_key, row.task) == self._linked_audit_key), None)
            if linked is not None:
                rows = [*rows, linked]
        self._render_audit_summary(self.tbl_audit, rows, self.cb_audit_status.currentData(), inspector=False, select_key=select_key)
        self._render_daily_diffs_table()

    def _audit_details(self, summary: QTableWidget, heading: QLabel):
        selected = self._selected_audit_summary(summary)
        if not selected:
            heading.setText("Chọn một dòng ở bảng trên để xem chi tiết." if summary.rowCount() else "Không có dòng phù hợp với bộ lọc hiện tại.")
            return []
        (_, task), person, issues = selected
        count = f"{_fmt_int(len(issues))} dòng cần xử lý" if issues else "Không có dòng cần xử lý"
        heading.setText(f"Chi tiết: {person} · {task} — {count}")
        return issues

    def _render_daily_diffs_table(self):
        t = self.theme
        diffs = self._audit_details(self.tbl_audit, self.lbl_detail_hdr)

        tbl = self.tbl_daily_diffs
        headers = [
            "Loại", "Annotator", "Inspector", "Ngày\nWork Log", "Ngày\nReview", "Task", "Job ID", "Frame\nStart ~ End",
            "Work Log\n(Frame / BBOX)", "Review\n(Frame / BBOX)", "Chênh\n(Frame / BBOX)", "Nguyên nhân",
        ]
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.horizontalHeaderItem(2).setToolTip("Tên Inspector trên các dòng Review được đối chiếu.")
        tbl.horizontalHeaderItem(3).setToolTip("Ngày ghi bên Work Log (chỉ để tham khảo vị trí, không dùng so sánh lệch).")
        tbl.horizontalHeaderItem(4).setToolTip("Ngày ghi bên Review (thường sau Work Log 1 hoặc vài ngày — không tính là lệch).")
        tbl.setRowCount(len(diffs))
        tbl.setWordWrap(True)

        for row_i, d in enumerate(diffs):
            is_pend = d.is_pending
            bg = t["accent_soft"] if is_pend else t["warn_soft"]
            accent_fg = t["accent_text"] if is_pend else t["warn"]
            rng_err_fg = t["accent_text"] if is_pend else t["danger"]
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
            rw_dt_str = d.rw_date.strftime("%Y-%m-%d") if d.rw_date else "—"

            if d.wl_range and d.rw_range and d.wl_range != d.rw_range:
                rng_txt = f"WL: {d.wl_range}\nRW: {d.rw_range}"
                rng_fg = rng_err_fg
            else:
                rng_txt = d.wl_range or d.rw_range or "—"
                rng_fg = None

            miss_rw_bb = (d.rw_frames > 0 and d.wl_bbox > 0 and d.rw_bbox <= 0)
            miss_wl_bb = (d.wl_frames > 0 and d.rw_bbox > 0 and d.wl_bbox <= 0)
            wl_bb_str = "—" if miss_wl_bb else _fmt_int(d.wl_bbox)
            rw_bb_str = "—" if miss_rw_bb else _fmt_int(d.rw_bbox)
            diff_bb_str = "—" if (miss_rw_bb or miss_wl_bb) else _fmt_int(d.diff_bbox)

            tbl.setItem(row_i, 0, self._make_item(kind_lbl, bold=True, fg=accent_fg, bg=bg))
            tbl.setItem(row_i, 1, self._make_item(d.annotator, bold=True, bg=bg))
            tbl.setItem(row_i, 2, self._make_item(", ".join(d.inspectors) or "—", bold=True, bg=bg))
            tbl.setItem(row_i, 3, self._make_item(d.date.strftime("%Y-%m-%d"), bg=bg))
            tbl.setItem(row_i, 4, self._make_item(rw_dt_str, bg=bg))
            tbl.setItem(row_i, 5, self._make_item(d.task, bold=True, bg=bg))
            tbl.setItem(row_i, 6, self._make_item(d.job_id or "—", bold=True, bg=bg))
            tbl.setItem(row_i, 7, self._make_item(rng_txt, bold=True, fg=rng_fg, bg=bg, tooltip=d.frame_range_display))
            tbl.setItem(row_i, 8, self._make_item(f"{_fmt_int(d.wl_frames)} / {wl_bb_str}", bg=bg))
            tbl.setItem(row_i, 9, self._make_item(f"{_fmt_int(d.rw_frames)} / {rw_bb_str}", bg=bg))
            tbl.setItem(row_i, 10, self._make_item(f"{_fmt_int(d.diff_frame)} / {diff_bb_str}", bold=True, fg=accent_fg, bg=bg))
            tbl.setItem(row_i, 11, self._make_item(d.reason, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, bold=True, fg=accent_fg, bg=bg, tooltip=d.reason))

        self._auto_fit_columns(tbl, headers, min_widths={len(headers) - 1: 280})
        tbl.resizeRowsToContents()

    def _render_reviewer_audit_table(self, rep: ReportResult):
        self._render_audit_summary(self.tbl_rev_audit, rep.reviewer_audit_rows, self.cb_rev_audit_status.currentData(), inspector=True)
        self._render_rev_diffs_table()

    def _render_rev_diffs_table(self):
        t = self.theme
        diffs = self._audit_details(self.tbl_rev_audit, self.lbl_rev_detail_hdr)

        tbl = self.tbl_rev_diffs
        headers = [
            "Loại", "Inspector", "Annotator", "Ngày\nWork Log", "Ngày\nReview", "Task", "Job ID",
            "Frame\nStart ~ End", "WL Inspector\n(Frame / BBOX)", "Review\n(Frame / BBOX)",
            "WL Annotator\n(Frame / BBOX)", "Nguyên nhân",
        ]
        tbl.setColumnCount(len(headers))
        tbl.setHorizontalHeaderLabels(headers)
        tbl.horizontalHeaderItem(3).setToolTip("Ngày ghi bên Work Log Inspector (chỉ để tham khảo, không dùng so sánh lệch).")
        tbl.horizontalHeaderItem(4).setToolTip("Ngày ghi bên Review (có thể lệch ngày so với Work Log — không tính là lệch).")
        tbl.setRowCount(len(diffs))
        tbl.setWordWrap(True)

        for row_i, d in enumerate(diffs):
            is_pend = d.is_pending
            bg = t["accent_soft"] if is_pend else t["warn_soft"]
            accent_fg = t["accent_text"] if is_pend else t["warn"]
            kind_lbl = "⏳ Chưa có data" if is_pend else "⚠ KHÔNG KHỚP"
            rw_dt_str = d.rw_date.strftime("%Y-%m-%d") if d.rw_date else "—"
            wl_s = f"{_fmt_int(d.wl_frames)} / {_fmt_int(d.wl_bbox)}" if (d.wl_frames or d.wl_bbox) else "—"
            rw_s = f"{_fmt_int(d.rw_frames)} / {_fmt_int(d.rw_bbox)}" if (d.rw_frames or d.rw_bbox) else "—"
            ann_s = f"{_fmt_int(d.ann_frames)} / {_fmt_int(d.ann_bbox)}" if (d.ann_frames or d.ann_bbox) else "—"

            tbl.setItem(row_i, 0, self._make_item(kind_lbl, Qt.AlignmentFlag.AlignCenter, True, fg=accent_fg, bg=bg))
            tbl.setItem(row_i, 1, self._make_item(d.inspector, Qt.AlignmentFlag.AlignCenter, True, bg=bg))
            tbl.setItem(row_i, 2, self._make_item(d.annotator, Qt.AlignmentFlag.AlignCenter, True, bg=bg))
            tbl.setItem(row_i, 3, self._make_item(d.date.strftime("%Y-%m-%d"), Qt.AlignmentFlag.AlignCenter, bg=bg))
            tbl.setItem(row_i, 4, self._make_item(rw_dt_str, Qt.AlignmentFlag.AlignCenter, bg=bg))
            tbl.setItem(row_i, 5, self._make_item(d.task, bold=True, bg=bg))
            tbl.setItem(row_i, 6, self._make_item(d.job_id or "—", bold=True, bg=bg))
            frame_range = d.frame_range_display
            tbl.setItem(row_i, 7, self._make_item(frame_range.replace(" | ", "\n"), bold=True, bg=bg, tooltip=frame_range))
            tbl.setItem(row_i, 8, self._make_item(wl_s, bg=bg))
            tbl.setItem(row_i, 9, self._make_item(rw_s, bg=bg))
            tbl.setItem(row_i, 10, self._make_item(ann_s, bg=bg))
            tbl.setItem(row_i, 11, self._make_item(d.reason, Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignVCenter, True, fg=accent_fg, bg=bg, tooltip=d.reason))

        self._auto_fit_columns(tbl, headers, min_widths={len(headers) - 1: 280})
        tbl.resizeRowsToContents()

    # ------------------------------------------------------------------
    # DRILL-DOWN CLICK HANDLERS
    # ------------------------------------------------------------------
    def _jump_to_audit_for(self, ann_key: str, task: Optional[str], *, from_inspector: bool = False):
        if not self.current_report:
            return

        def matching_rows(report):
            return [row for row in report.all_rows if row.annotator_key == ann_key and (task is None or row.task == task)]

        candidates = matching_rows(self.current_report)
        linked = not candidates and from_inspector and self._period_report is not None
        if linked:
            candidates = matching_rows(self._period_report)
        if not candidates:
            self.toast.show_toast("Không có dòng đối chiếu phù hợp trong kỳ và bộ lọc hiện tại.", "warn")
            return

        priority = {"mismatch": 0, "pending": 1, "ok": 2}
        target = min(candidates, key=lambda row: priority[row.status_kind])
        key = (target.annotator_key, target.task)
        self._linked_audit_key = key if linked else None
        status = self.cb_audit_status.currentData()
        if status and status != target.status_kind:
            with QSignalBlocker(self.cb_audit_status):
                self.cb_audit_status.setCurrentIndex(0)
        self.tabs.setCurrentWidget(self.tab_audit)
        self.audit_mode_tabs.setCurrentIndex(0)
        self._render_audit_table(self.current_report, select_key=key)
        self.tbl_audit.setFocus()
        self.tbl_audit.scrollToItem(self.tbl_audit.currentItem())

    def _on_t1_double_click(self, row: int, _col: int):
        it = self.tbl_t1.item(row, 1)
        if it and it.data(Qt.ItemDataRole.UserRole):
            self._jump_to_audit_for(*it.data(Qt.ItemDataRole.UserRole))

    def _on_t2_double_click(self, row: int, _col: int):
        it = self.tbl_t2.item(row, 0)
        if it and it.data(Qt.ItemDataRole.UserRole):
            self._jump_to_audit_for(*it.data(Qt.ItemDataRole.UserRole))

    def _on_rank_double_click(self, tbl: QTableWidget, row: int):
        it = tbl.item(row, 1)
        if it and it.data(Qt.ItemDataRole.UserRole):
            self._jump_to_audit_for(*it.data(Qt.ItemDataRole.UserRole), from_inspector=tbl is self.tbl_rev)

    # ------------------------------------------------------------------
    # COPY & EXPORT
    # ------------------------------------------------------------------
    def copy_current_table(self):
        cur_w = self.tabs.currentWidget()
        if self.tabs.currentIndex() == 0:
            tbl = self.tbl_t1_summary if self.tbl_t1_summary.hasFocus() else self.tbl_t1
        elif cur_w is self.tbl_t2:
            tbl = self.tbl_t2
        elif cur_w is self.tbl_rev:
            tbl = self.tbl_rev
        elif cur_w is self.tab_rank:
            if self.rank_mode_tabs.currentIndex() == 1:
                tbl = self.tbl_perf_reviewer if self.tbl_perf_reviewer.hasFocus() else self.tbl_perf_worker
            else:
                tbl = self.tbl_rank_reviewer if self.tbl_rank_reviewer.hasFocus() else self.tbl_rank_worker
        else:
            if self.audit_mode_tabs.currentIndex() == 1:
                summary, details = self.tbl_rev_audit, self.tbl_rev_diffs
            else:
                summary, details = self.tbl_audit, self.tbl_daily_diffs
            tbl = summary if summary.hasFocus() or not details.rowCount() else details
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
        inc_t1 = self.chk_exp_t1.isChecked()
        inc_t2 = self.chk_exp_t2.isChecked()
        inc_rev = self.chk_exp_rev.isChecked()
        inc_rank = self.chk_exp_rank.isChecked()
        inc_au = self.chk_exp_audit.isChecked()
        include_other_sheets = inc_t1 or inc_t2 or inc_rank or inc_au
        if not (inc_rev or include_other_sheets):
            self.toast.show_toast("Hãy tick ít nhất 1 sheet để xuất Excel!", "warn")
            return

        rep = self.current_report
        inspector_groups = self._selected_inspector_groups()
        if not rep or not ((inc_rev and inspector_groups) or (include_other_sheets and rep.all_rows)):
            self.toast.show_toast("Chưa có dữ liệu trong kỳ / bộ lọc đang chọn!", "warn")
            return

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
                include_reviewer=inc_rev,
                include_ranking=inc_rank,
                include_audit=inc_au,
                highlight_mismatch=self.chk_highlight.isChecked(),
                inspector_groups=inspector_groups,
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

    # ------------------------------------------------------------------
    # AUTO-UPDATE CHECK & IN-PLACE PORTABLE UPDATER
    # ------------------------------------------------------------------
    def _style_update_controls(self):
        if not hasattr(self, "btn_version"):
            return
        t = self.theme
        if self.latest_release is not None:
            self.btn_version.setText(f"v{APP_VERSION} • Có v{self.latest_release.version}")
            self.btn_version.setStyleSheet(
                f"QPushButton {{ background-color: {t['accent']}; color: #FFFFFF; border: none; "
                f"border-radius: 9px; padding: 0 8px; font-size: 10.5px; font-weight: 700; }}"
                f"QPushButton:hover {{ background-color: {t['accent_hover']}; }}"
            )
            self.btn_status_update.setText(f"⬆ Có bản mới v{self.latest_release.version} — Nhấn để cập nhật")
            self.btn_status_update.setStyleSheet(
                f"QPushButton {{ color: {t['accent_text']}; font-size: 11.5px; font-weight: 700; "
                f"border: none; background: transparent; padding: 0 8px; }}"
                f"QPushButton:hover {{ text-decoration: underline; }}"
            )
        else:
            self.btn_version.setText(f"v{APP_VERSION}")
            self.btn_version.setStyleSheet(
                f"QPushButton {{ background-color: {t['bg_subtle']}; color: {t['text_muted']}; "
                f"border: 1px solid {t['border']}; border-radius: 9px; padding: 0 7px; "
                f"font-size: 10.5px; font-weight: 600; }}"
                f"QPushButton:hover {{ border-color: {t['accent']}; color: {t['accent_text']}; }}"
            )
            self.btn_status_update.setText(f"v{APP_VERSION} • Kiểm tra cập nhật")
            self.btn_status_update.setStyleSheet(
                f"QPushButton {{ color: {t['text_muted']}; font-size: 11.5px; font-weight: 600; "
                f"border: none; background: transparent; padding: 0 8px; }}"
                f"QPushButton:hover {{ color: {t['accent_text']}; }}"
            )

    def _on_version_badge_clicked(self):
        if self.latest_release is not None:
            self.open_update_dialog(auto_start=False)
        else:
            self.check_for_updates(silent=False)

    def check_for_updates(self, silent: bool = False):
        if self.update_check_thread and self.update_check_thread.isRunning():
            if not silent:
                self.toast.show_toast("Đang kiểm tra cập nhật...", "info", 2000)
            return

        self._update_check_silent = silent
        if not silent:
            self.btn_status_update.setEnabled(False)
            self.btn_status_update.setText("Đang kiểm tra cập nhật...")
            self.status_bar.showMessage("Đang kiểm tra phiên bản mới từ GitHub...")

        self.update_check_thread = UpdateCheckThread(APP_VERSION, parent=self)
        self.update_check_thread.update_available.connect(self._on_update_available)
        self.update_check_thread.up_to_date.connect(self._on_update_up_to_date)
        self.update_check_thread.failed.connect(self._on_update_check_failed)
        self.update_check_thread.finished.connect(self._on_update_check_done)
        self.update_check_thread.start()

    def _on_update_check_done(self):
        self.btn_status_update.setEnabled(True)
        self._style_update_controls()

    def _on_update_available(self, release: ReleaseInfo):
        self.latest_release = release
        self._style_update_controls()
        self.update_banner.show_release(release, APP_VERSION)
        if not self._update_check_silent:
            self.status_bar.showMessage(f"Tìm thấy phiên bản mới v{release.version}.")
            self.open_update_dialog(auto_start=False)

    def _on_update_up_to_date(self, latest_ver: str):
        self.latest_release = None
        self.update_banner.setVisible(False)
        self._style_update_controls()
        if not self._update_check_silent:
            msg = f"✓ Bạn đang dùng phiên bản mới nhất (v{APP_VERSION})."
            self.toast.show_toast(msg, "success", 3500)
            self.status_bar.showMessage(msg)

    def _on_update_check_failed(self, err_msg: str):
        if not self._update_check_silent:
            self.toast.show_toast("Không thể kiểm tra cập nhật", "error", 3500)
            self.status_bar.showMessage("Kiểm tra cập nhật thất bại.")
            QMessageBox.warning(self, "Không thể kiểm tra cập nhật", err_msg)

    def open_update_dialog(self, auto_start: bool = False):
        if not self.latest_release:
            self.check_for_updates(silent=False)
            return
        dlg = UpdateDialog(
            release=self.latest_release,
            current_version=APP_VERSION,
            theme=self.theme,
            auto_start=auto_start,
            parent=self,
        )
        dlg.exec()


def _apply_windows_taskbar_icon(win: QMainWindow, icon_path: str):
    """Explicitly set Win32 ICON_BIG / ICON_SMALL so Windows taskbar & Alt+Tab always show app_icon.ico."""
    if sys.platform != "win32" or not os.path.isfile(icon_path):
        return
    try:
        import ctypes
        user32 = ctypes.windll.user32
        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040
        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1
        hwnd = int(win.winId())
        hicon_big = user32.LoadImageW(None, icon_path, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
        hicon_small = user32.LoadImageW(None, icon_path, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)
        if hicon_big:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)
        if hicon_small:
            user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small)
    except Exception:
        pass


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("ReportTools.WorklogReport.App")
        except Exception:
            pass
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    icon_file = _resource_path("app_icon.ico")
    if os.path.isfile(icon_file):
        app.setWindowIcon(QIcon(icon_file))
    win = MainWindow()
    win.show()
    _apply_windows_taskbar_icon(win, icon_file)
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
