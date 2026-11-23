from PyQt6.QtCore import QByteArray, QRect, QSize, Qt
from PyQt6.QtGui import QColor, QIcon, QPainter, QPen, QPixmap
from PyQt6.QtSvg import QSvgRenderer
from PyQt6.QtWidgets import QStyledItemDelegate


LIGHT_THEME = {
    "name": "light",
    "bg_canvas": "#D8E0EA",
    "bg_surface": "#E6ECF3",
    "bg_subtle": "#DCE4EE",
    "bg_hover": "#CDD9E8",
    "bg_selected": "#BFD6F6",
    "border": "#BCC8D8",
    "border_strong": "#94A3B8",
    "text_primary": "#0F172A",
    "text_secondary": "#1E293B",
    "text_muted": "#475569",
    "accent": "#2563EB",
    "accent_hover": "#1D4ED8",
    "accent_soft": "#DBEAFE",
    "accent_text": "#1D4ED8",
    "success": "#059669",
    "success_hover": "#047857",
    "success_soft": "#D1FAE5",
    "warn": "#B45309",
    "warn_soft": "#FEF3C7",
    "warn_border": "#F59E0B",
    "danger": "#DC2626",
    "danger_soft": "#FEE2E2",
    "tbl_bg": "#FFFFFF",
    "tbl_header_bg": "#1E293B",
    "tbl_header_fg": "#F8FAFC",
    "tbl_header_border": "#334155",
    "tbl_grid": "#CBD5E1",
    "tbl_divider": "#0F172A",
    "tbl_kpi_bg": "#F1F5F9",
    "task_badge_bg": "#1E293B",
    "task_badge_fg": "#FFFFFF",
}

DARK_THEME = {
    "name": "dark",
    "bg_canvas": "#0B1120",
    "bg_surface": "#111827",
    "bg_subtle": "#162032",
    "bg_hover": "#1E293B",
    "bg_selected": "#1E3A5F",
    "border": "#1F2937",
    "border_strong": "#334155",
    "text_primary": "#F8FAFC",
    "text_secondary": "#CBD5E1",
    "text_muted": "#94A3B8",
    "accent": "#3B82F6",
    "accent_hover": "#2563EB",
    "accent_soft": "#172554",
    "accent_text": "#93C5FD",
    "success": "#10B981",
    "success_hover": "#059669",
    "success_soft": "#064E3B",
    "warn": "#F59E0B",
    "warn_soft": "#451A03",
    "warn_border": "#92400E",
    "danger": "#F87171",
    "danger_soft": "#450A0A",
    "tbl_bg": "#111827",
    "tbl_header_bg": "#0F172A",
    "tbl_header_fg": "#F8FAFC",
    "tbl_header_border": "#334155",
    "tbl_grid": "#1F2937",
    "tbl_divider": "#64748B",
    "tbl_kpi_bg": "#162032",
    "task_badge_bg": "#0F172A",
    "task_badge_fg": "#F8FAFC",
}


SVG_ICONS = {
    "chart": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M3 3v18h18"/><path d="M18 17V9"/><path d="M13 17V5"/><path d="M8 17v-3"/></svg>""",
    "file": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/><polyline points="14 2 14 8 20 8"/><line x1="8" y1="13" x2="16" y2="13"/><line x1="8" y1="17" x2="16" y2="17"/></svg>""",
    "refresh": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polyline points="23 4 23 10 17 10"/><path d="M20.49 15a9 9 0 1 1-2.12-9.36L23 10"/></svg>""",
    "download": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/></svg>""",
    "copy": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="9" y="9" width="13" height="13" rx="2" ry="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/></svg>""",
    "sun": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="5"/><line x1="12" y1="1" x2="12" y2="3"/><line x1="12" y1="21" x2="12" y2="23"/><line x1="4.22" y1="4.22" x2="5.64" y2="5.64"/><line x1="18.36" y1="18.36" x2="19.78" y2="19.78"/><line x1="1" y1="12" x2="3" y2="12"/><line x1="21" y1="12" x2="23" y2="12"/><line x1="4.22" y1="19.78" x2="5.64" y2="18.36"/><line x1="18.36" y1="5.64" x2="19.78" y2="4.22"/></svg>""",
    "moon": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 12.79A9 9 0 1 1 11.21 3 7 7 0 0 0 21 12.79z"/></svg>""",
    "users": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17 21v-2a4 4 0 0 0-4-4H5a4 4 0 0 0-4 4v2"/><circle cx="9" cy="7" r="4"/><path d="M23 21v-2a4 4 0 0 0-3-3.87"/><path d="M16 3.13a4 4 0 0 1 0 7.75"/></svg>""",
    "layers": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><polygon points="12 2 2 7 12 12 22 7 12 2"/><polyline points="2 17 12 22 22 17"/><polyline points="2 12 12 17 22 12"/></svg>""",
    "box": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M21 16V8a2 2 0 0 0-1-1.73l-7-4a2 2 0 0 0-2 0l-7 4A2 2 0 0 0 3 8v8a2 2 0 0 0 1 1.73l7 4a2 2 0 0 0 2 0l7-4A2 2 0 0 0 21 16z"/><polyline points="3.27 6.96 12 12.01 20.73 6.96"/><line x1="12" y1="22.08" x2="12" y2="12"/></svg>""",
    "clock": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="10"/><polyline points="12 6 12 12 16 14"/></svg>""",
    "shield": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/><polyline points="9 12 11 14 15 10"/></svg>""",
    "alert": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/></svg>""",
    "search": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><circle cx="11" cy="11" r="8"/><line x1="21" y1="21" x2="16.65" y2="16.65"/></svg>""",
    "folder": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/></svg>""",
    "calendar": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>""",
    "cloud": """<svg viewBox="0 0 24 24" fill="none" stroke="{color}" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M18 10h-1.26A8 8 0 1 0 9 20h9a5 5 0 0 0 0-10z"/></svg>""",
}


def svg_pixmap(name: str, color: str = "#475569", size: int = 18) -> QPixmap:
    svg_str = SVG_ICONS.get(name, SVG_ICONS["chart"]).format(color=color)
    renderer = QSvgRenderer(QByteArray(svg_str.encode("utf-8")))
    pm = QPixmap(QSize(size * 2, size * 2))
    pm.fill(Qt.GlobalColor.transparent)
    painter = QPainter(pm)
    renderer.render(painter)
    painter.end()
    pm.setDevicePixelRatio(2.0)
    return pm


def svg_icon(name: str, color: str = "#475569", size: int = 18) -> QIcon:
    return QIcon(svg_pixmap(name, color=color, size=size))


def build_stylesheet(t: dict) -> str:
    return f"""
QMainWindow {{
    background-color: {t["bg_canvas"]};
}}
QWidget {{
    font-family: "Segoe UI", Arial, sans-serif;
    font-size: 12.5px;
    color: {t["text_primary"]};
}}
QScrollArea, QWidget#TransparentHost {{
    background-color: transparent;
    border: none;
}}

/* Top Command Bar */
QFrame#TopBar {{
    background-color: {t["bg_surface"]};
    border-bottom: 1px solid {t["border"]};
}}
QLabel#AppTitle {{
    font-size: 16px;
    font-weight: 700;
    color: {t["text_primary"]};
}}
QLabel#AppSubtitle {{
    font-size: 11px;
    font-weight: 500;
    color: {t["text_muted"]};
}}

/* Sidebar & Cards */
QFrame#Sidebar {{
    background-color: {t["bg_surface"]};
    border-right: 1px solid {t["border"]};
}}
QFrame#SectionCard {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border"]};
    border-radius: 10px;
}}
QFrame#FileCard {{
    background-color: {t["bg_subtle"]};
    border: 1.5px dashed {t["border_strong"]};
    border-radius: 9px;
}}
QFrame#FileCard:hover {{
    border-color: {t["accent"]};
    background-color: {t["bg_hover"]};
}}
QFrame#FileCard[loaded="true"] {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border_strong"]};
}}
QFrame#MetricCard {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border"]};
    border-radius: 10px;
}}
QFrame#MetricCard:hover {{
    border-color: {t["border_strong"]};
}}
QFrame#IsolateBanner {{
    background-color: {t["accent_soft"]};
    border: 1px solid {t["accent"]};
    border-radius: 8px;
}}

QLabel#SectionOverline {{
    font-size: 11px;
    font-weight: 700;
    color: {t["text_muted"]};
    letter-spacing: 0.5px;
}}

/* Inputs */
QLineEdit, QDateEdit, QComboBox {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border_strong"]};
    border-radius: 7px;
    padding: 5px 9px;
    min-height: 22px;
    color: {t["text_primary"]};
}}
QLineEdit:focus, QDateEdit:focus, QComboBox:focus {{
    border: 1.5px solid {t["accent"]};
}}
QComboBox::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 24px;
    background-color: {t["bg_subtle"]};
    border-left: 1px solid {t["border_strong"]};
    border-top-right-radius: 6px;
    border-bottom-right-radius: 6px;
}}
QComboBox::drop-down:hover {{
    background-color: {t["accent_soft"]};
}}
QComboBox QAbstractItemView {{
    background-color: {t["bg_surface"]};
    color: {t["text_primary"]};
    border: 1px solid {t["border_strong"]};
    border-radius: 6px;
    padding: 3px;
    outline: none;
    selection-background-color: {t["accent"]};
    selection-color: #FFFFFF;
}}
QComboBox QAbstractItemView::item {{
    padding: 5px 10px;
    min-height: 22px;
    color: {t["text_primary"]};
    background-color: {t["bg_surface"]};
}}
QComboBox QAbstractItemView::item:hover {{
    background-color: {t["bg_hover"]};
    color: {t["text_primary"]};
}}
QComboBox QAbstractItemView::item:selected {{
    background-color: {t["accent"]};
    color: #FFFFFF;
}}
QDateEdit::drop-down {{
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 28px;
    background-color: {t["accent_soft"]};
    border-left: 1px solid {t["border_strong"]};
    border-top-right-radius: 6px;
    border-bottom-right-radius: 6px;
}}
QDateEdit::drop-down:hover {{
    background-color: {t["accent"]};
}}

/* Calendar Popup */
QCalendarWidget QWidget {{
    background-color: {t["bg_surface"]};
    color: {t["text_primary"]};
}}
QCalendarWidget QWidget#qt_calendar_navigationbar {{
    background-color: {t["tbl_header_bg"]};
    padding: 4px;
}}
QCalendarWidget QToolButton {{
    background-color: transparent;
    color: #FFFFFF;
    font-weight: 700;
    border: none;
    border-radius: 4px;
    padding: 4px 8px;
}}
QCalendarWidget QToolButton:hover {{
    background-color: {t["accent"]};
}}
QCalendarWidget QAbstractItemView:enabled {{
    background-color: {t["bg_surface"]};
    color: {t["text_primary"]};
    selection-background-color: {t["accent"]};
    selection-color: #FFFFFF;
    font-size: 12px;
}}

/* Buttons */
QPushButton {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border_strong"]};
    border-radius: 7px;
    padding: 5px 12px;
    font-weight: 600;
    font-size: 12px;
    color: {t["text_primary"]};
}}
QPushButton:hover {{
    background-color: {t["bg_hover"]};
    border-color: {t["accent"]};
}}
QPushButton:pressed {{
    background-color: {t["bg_selected"]};
}}
QPushButton:disabled {{
    opacity: 0.5;
    color: {t["text_muted"]};
}}
QPushButton#PrimaryBtn {{
    background-color: {t["accent"]};
    color: #FFFFFF;
    border: 1px solid {t["accent"]};
    padding: 6px 14px;
}}
QPushButton#PrimaryBtn:hover {{
    background-color: {t["accent_hover"]};
    border-color: {t["accent_hover"]};
}}
QPushButton#SuccessBtn {{
    background-color: {t["success"]};
    color: #FFFFFF;
    border: 1px solid {t["success"]};
    padding: 7px 16px;
    font-size: 13px;
}}
QPushButton#SuccessBtn:hover {{
    background-color: {t["success_hover"]};
    border-color: {t["success_hover"]};
}}
QPushButton#GhostBtn {{
    background-color: transparent;
    border: 1px solid transparent;
    color: {t["text_secondary"]};
    padding: 4px 8px;
    font-weight: 600;
}}
QPushButton#GhostBtn:hover {{
    background-color: {t["bg_hover"]};
    color: {t["accent_text"]};
}}
QPushButton#ChipBtn {{
    background-color: {t["bg_subtle"]};
    border: 1px solid {t["border_strong"]};
    border-radius: 13px;
    padding: 3px 10px;
    font-size: 11.5px;
    font-weight: 600;
    color: {t["text_secondary"]};
}}
QPushButton#ChipBtn:hover {{
    border-color: {t["accent"]};
    color: {t["accent_text"]};
}}
QPushButton#ChipBtn:checked {{
    background-color: {t["accent"]};
    border-color: {t["accent"]};
    color: #FFFFFF;
}}
QPushButton#PresetBtn {{
    background-color: {t["bg_subtle"]};
    border: 1px solid {t["border"]};
    border-radius: 6px;
    padding: 4px 6px;
    font-size: 11.5px;
    font-weight: 600;
    color: {t["text_secondary"]};
}}
QPushButton#PresetBtn:hover {{
    background-color: {t["accent_soft"]};
    border-color: {t["accent"]};
    color: {t["accent_text"]};
}}
QPushButton#PresetBtn:checked {{
    background-color: {t["accent"]};
    border-color: {t["accent"]};
    color: #FFFFFF;
}}

/* Segmented Control */
QFrame#SegmentedWrap {{
    background-color: {t["bg_subtle"]};
    border: 1px solid {t["border"]};
    border-radius: 8px;
    padding: 2px;
}}
QPushButton#SegBtn {{
    background-color: transparent;
    border: none;
    border-radius: 6px;
    padding: 5px 8px;
    font-size: 11.5px;
    font-weight: 600;
    color: {t["text_secondary"]};
}}
QPushButton#SegBtn:hover {{
    color: {t["text_primary"]};
}}
QPushButton#SegBtn:checked {{
    background-color: {t["bg_surface"]};
    color: {t["accent_text"]};
    border: 1px solid {t["border_strong"]};
}}

/* Lists */
QListWidget {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border"]};
    border-radius: 8px;
    padding: 4px;
    outline: none;
}}
QListWidget::item {{
    padding: 5px 6px;
    border-radius: 6px;
    margin-bottom: 1px;
}}
QListWidget::item:hover {{
    background-color: {t["bg_hover"]};
}}
QListWidget::item:selected {{
    background-color: {t["bg_selected"]};
    color: {t["text_primary"]};
}}

/* Tabs */
QTabWidget::pane {{
    border: 1px solid {t["border"]};
    border-radius: 10px;
    background-color: {t["bg_surface"]};
    top: -1px;
}}
QTabBar::tab {{
    background-color: transparent;
    border: 1px solid transparent;
    border-bottom: none;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
    padding: 8px 16px;
    margin-right: 4px;
    font-weight: 600;
    font-size: 12.5px;
    color: {t["text_muted"]};
}}
QTabBar::tab:hover {{
    color: {t["text_primary"]};
    background-color: {t["bg_subtle"]};
}}
QTabBar::tab:selected {{
    background-color: {t["bg_surface"]};
    border: 1px solid {t["border"]};
    border-bottom: 2px solid {t["bg_surface"]};
    color: {t["accent_text"]};
}}
QTabBar QToolButton {{
    background-color: {t["bg_subtle"]};
    border: 1px solid {t["border"]};
    border-radius: 4px;
    color: {t["text_secondary"]};
}}

/* Tables */
QTableWidget {{
    background-color: {t["tbl_bg"]};
    gridline-color: {t["tbl_grid"]};
    border: none;
    selection-background-color: {t["bg_selected"]};
    selection-color: {t["text_primary"]};
    outline: none;
}}
QTableWidget::item {{
    padding: 2px 8px;
}}
QHeaderView::section {{
    background-color: {t["tbl_header_bg"]};
    color: {t["tbl_header_fg"]};
    font-weight: 700;
    font-size: 11.5px;
    padding: 6px 6px;
    border: 1px solid {t["tbl_header_border"]};
}}

/* CheckBox */
QCheckBox {{
    spacing: 6px;
    font-size: 12px;
    color: {t["text_secondary"]};
}}
QCheckBox::indicator {{
    width: 15px;
    height: 15px;
    border-radius: 4px;
    border: 1.5px solid {t["border_strong"]};
    background-color: {t["bg_surface"]};
}}
QCheckBox::indicator:checked {{
    background-color: {t["accent"]};
    border-color: {t["accent"]};
}}

/* Scrollbars */
QScrollBar {{
    background: {t["bg_subtle"]};
    border-radius: 6px;
}}
QScrollBar:vertical {{
    width: 14px;
    margin: 2px 1px;
}}
QScrollBar:horizontal {{
    height: 14px;
    margin: 1px 2px;
}}
QScrollBar::handle {{
    background: {t["text_muted"]};
    border-radius: 6px;
}}
QScrollBar::handle:vertical {{
    min-height: 36px;
}}
QScrollBar::handle:horizontal {{
    min-width: 36px;
}}
QScrollBar::handle:hover {{
    background: {t["accent"]};
}}
QScrollBar::handle:pressed {{
    background: {t["accent_hover"]};
}}
QScrollBar::add-page, QScrollBar::sub-page {{
    background: transparent;
}}
QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {{
    height: 0px;
}}
QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {{
    width: 0px;
}}
QAbstractScrollArea::corner {{
    background: {t["bg_subtle"]};
}}

/* Progress Bar */
QProgressBar {{
    border: none;
    border-radius: 3px;
    background-color: {t["border"]};
    height: 5px;
    text-align: center;
}}
QProgressBar::chunk {{
    background-color: {t["accent"]};
    border-radius: 3px;
}}

/* Splitter */
QSplitter::handle {{
    background-color: transparent;
}}

/* Status Bar */
QStatusBar {{
    background-color: {t["bg_surface"]};
    border-top: 1px solid {t["border"]};
    color: {t["text_muted"]};
    font-size: 11.5px;
}}
"""


class ReportTableDelegate(QStyledItemDelegate):
    """
    Custom table cell renderer that paints publication-grade structural rules:
      - Crisp thick vertical divider separating Input columns from Computed KPI columns
      - Crisp horizontal separator between Task groups (e.g., OD vs SSEG)
    """

    def __init__(self, divider_col, parent=None):
        super().__init__(parent)
        if isinstance(divider_col, (list, tuple, set)):
            self.divider_cols = set(divider_col)
        else:
            self.divider_cols = {divider_col}
        self.group_bottom_rows = set()
        self.divider_color = QColor("#0F172A")

    def set_theme(self, theme: dict):
        self.divider_color = QColor(theme["tbl_divider"])

    def set_group_bottom_rows(self, rows: set):
        self.group_bottom_rows = set(rows)

    def paint(self, painter: QPainter, option, index):
        super().paint(painter, option, index)
        rect: QRect = option.rect
        painter.save()
        # Thick vertical divider on left edge of any divider_col
        if index.column() in self.divider_cols:
            pen = QPen(self.divider_color, 2.2)
            painter.setPen(pen)
            painter.drawLine(rect.left() + 1, rect.top(), rect.left() + 1, rect.bottom())
        # Thick horizontal divider at bottom of each task group
        if index.row() in self.group_bottom_rows:
            pen = QPen(self.divider_color, 1.8)
            painter.setPen(pen)
            painter.drawLine(rect.left(), rect.bottom(), rect.right(), rect.bottom())
        painter.restore()
