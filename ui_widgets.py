import datetime
import os
from typing import List, Optional, Tuple

from PyQt6.QtCore import QPoint, QRect, QSize, Qt, QTimer, pyqtSignal
from PyQt6.QtGui import QCursor
from PyQt6.QtWidgets import (
    QButtonGroup,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLayout,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from ui_theme import svg_pixmap


class FlowLayout(QLayout):
    """Responsive wrapping layout for Task filter pill chips."""

    def __init__(self, parent=None, margin=0, h_spacing=6, v_spacing=6):
        super().__init__(parent)
        self._items = []
        self._h_space = h_spacing
        self._v_space = v_spacing
        self.setContentsMargins(margin, margin, margin, margin)

    def addItem(self, item):
        self._items.append(item)

    def count(self):
        return len(self._items)

    def itemAt(self, index):
        if 0 <= index < len(self._items):
            return self._items[index]
        return None

    def takeAt(self, index):
        if 0 <= index < len(self._items):
            return self._items.pop(index)
        return None

    def expandingDirections(self):
        return Qt.Orientation(0)

    def hasHeightForWidth(self):
        return True

    def heightForWidth(self, width):
        return self._do_layout(QRect(0, 0, width, 0), test_only=True)

    def setGeometry(self, rect):
        super().setGeometry(rect)
        self._do_layout(rect, test_only=False)

    def sizeHint(self):
        return self.minimumSize()

    def minimumSize(self):
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        m = self.contentsMargins()
        size += QSize(m.left() + m.right(), m.top() + m.bottom())
        return size

    def _do_layout(self, rect, test_only: bool) -> int:
        m = self.contentsMargins()
        eff = rect.adjusted(m.left(), m.top(), -m.right(), -m.bottom())
        x = eff.x()
        y = eff.y()
        line_h = 0

        for item in self._items:
            sz = item.sizeHint()
            next_x = x + sz.width() + self._h_space
            if next_x - self._h_space > eff.right() and line_h > 0:
                x = eff.x()
                y = y + line_h + self._v_space
                next_x = x + sz.width() + self._h_space
                line_h = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), sz))
            x = next_x
            line_h = max(line_h, sz.height())

        return y + line_h - rect.y() + m.bottom()


class SegmentedControl(QFrame):
    """Sleek pill-style segmented toggle switch."""

    changed = pyqtSignal(str)

    def __init__(self, options: List[Tuple[str, str]], parent=None):
        super().__init__(parent)
        self.setObjectName("SegmentedWrap")
        lay = QHBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.setSpacing(2)

        self._group = QButtonGroup(self)
        self._group.setExclusive(True)
        self._buttons = {}

        for idx, (label, value) in enumerate(options):
            btn = QPushButton(label)
            btn.setObjectName("SegBtn")
            btn.setCheckable(True)
            btn.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
            btn.setProperty("seg_value", value)
            if idx == 0:
                btn.setChecked(True)
            self._group.addButton(btn, idx)
            self._buttons[value] = btn
            lay.addWidget(btn, 1)

        self._group.buttonClicked.connect(self._on_click)

    def _on_click(self, btn):
        self.changed.emit(btn.property("seg_value"))

    def current_value(self) -> str:
        checked = self._group.checkedButton()
        return checked.property("seg_value") if checked else ""

    def set_value(self, value: str, emit: bool = False):
        if value in self._buttons:
            self._buttons[value].setChecked(True)
            if emit:
                self.changed.emit(value)


class FileDropCard(QFrame):
    """Modern compact file card supporting click-to-browse AND drag & drop .xlsx."""

    file_dropped = pyqtSignal(str)
    browse_clicked = pyqtSignal()

    def __init__(self, title: str, placeholder: str, accent: str = "#2563EB", parent=None):
        super().__init__(parent)
        self.setObjectName("FileCard")
        self.setAcceptDrops(True)
        self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.setProperty("loaded", "false")
        self._path = ""
        self._placeholder = placeholder
        self._accent = accent

        lay = QHBoxLayout(self)
        lay.setContentsMargins(10, 7, 10, 7)
        lay.setSpacing(9)

        self.icon_lbl = QLabel()
        self.icon_lbl.setFixedSize(28, 28)
        self.icon_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.icon_lbl.setPixmap(svg_pixmap("file", accent, 18))
        lay.addWidget(self.icon_lbl)

        text_col = QVBoxLayout()
        text_col.setSpacing(1)
        top_row = QHBoxLayout()
        top_row.setSpacing(6)
        self.lbl_title = QLabel(title)
        self.lbl_title.setStyleSheet("font-size: 11px; font-weight: 700; color: #64748B;")
        self.lbl_badge = QLabel("Chưa chọn")
        self.lbl_badge.setStyleSheet(
            "font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 8px; "
            "background: #E2E8F0; color: #475569;"
        )
        top_row.addWidget(self.lbl_title)
        top_row.addWidget(self.lbl_badge)
        top_row.addStretch(1)

        self.lbl_name = QLabel(placeholder)
        self.lbl_name.setStyleSheet("font-size: 12px; font-weight: 600;")
        self.lbl_name.setTextInteractionFlags(Qt.TextInteractionFlag.NoTextInteraction)
        text_col.addLayout(top_row)
        text_col.addWidget(self.lbl_name)
        lay.addLayout(text_col, 1)

        self.btn_pick = QPushButton("Đổi file")
        self.btn_pick.setObjectName("GhostBtn")
        self.btn_pick.setFixedHeight(26)
        self.btn_pick.clicked.connect(self.browse_clicked.emit)
        lay.addWidget(self.btn_pick)

    def path(self) -> str:
        return self._path

    def set_path(self, path: str, theme_dark: bool = False):
        self._path = path or ""
        if self._path and os.path.isfile(self._path):
            size_mb = os.path.getsize(self._path) / (1024 * 1024)
            mtime = datetime.datetime.fromtimestamp(os.path.getmtime(self._path))
            time_str = mtime.strftime("%H:%M %d/%m")
            is_od = "onedrive" in self._path.lower()
            src_tag = "☁ OneDrive" if is_od else "✓ Local"
            fname = os.path.basename(self._path)
            short = fname if len(fname) <= 34 else fname[:31] + "..."
            self.lbl_name.setText(short)
            self.setToolTip(
                f"Nguồn: {'OneDrive / SharePoint Sync' if is_od else 'Thư mục máy tính'}\n"
                f"Đường dẫn: {self._path}\n"
                f"Cập nhật lần cuối: {mtime.strftime('%H:%M:%S %d/%m/%Y')} ({size_mb:.2f} MB)\n"
                f"(Kéo thả file .xlsx vào đây để thay đổi)"
            )
            self.lbl_badge.setText(f"{src_tag} · {time_str}")
            if is_od:
                if theme_dark:
                    self.lbl_badge.setStyleSheet(
                        "font-size: 10px; font-weight: 700; padding: 1px 7px; border-radius: 8px; "
                        "background: #1E3A8A; color: #93C5FD;"
                    )
                else:
                    self.lbl_badge.setStyleSheet(
                        "font-size: 10px; font-weight: 700; padding: 1px 7px; border-radius: 8px; "
                        "background: #DBEAFE; color: #1D4ED8;"
                    )
            elif theme_dark:
                self.lbl_badge.setStyleSheet(
                    "font-size: 10px; font-weight: 700; padding: 1px 7px; border-radius: 8px; "
                    "background: #064E3B; color: #34D399;"
                )
            else:
                self.lbl_badge.setStyleSheet(
                    "font-size: 10px; font-weight: 700; padding: 1px 7px; border-radius: 8px; "
                    "background: #ECFDF5; color: #059669;"
                )
            self.setProperty("loaded", "true")
        else:
            self.lbl_name.setText(self._placeholder)
            self.setToolTip("Nhấp hoặc kéo thả file .xlsx vào đây")
            self.lbl_badge.setText("Chưa chọn")
            self.lbl_badge.setStyleSheet(
                "font-size: 10px; font-weight: 700; padding: 1px 6px; border-radius: 8px; "
                "background: #E2E8F0; color: #475569;"
            )
            self.setProperty("loaded", "false")
        self.style().unpolish(self)
        self.style().polish(self)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.browse_clicked.emit()
        super().mousePressEvent(event)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for u in event.mimeData().urls():
                if u.toLocalFile().lower().endswith(".xlsx"):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        for u in event.mimeData().urls():
            p = u.toLocalFile()
            if p.lower().endswith(".xlsx"):
                self.file_dropped.emit(p)
                event.acceptProposedAction()
                return


class MetricCard(QFrame):
    """Executive KPI card with icon badge, large tabular stat, and subtitle."""

    clicked = pyqtSignal()

    def __init__(self, title: str, icon_name: str, accent: str = "#2563EB", clickable: bool = False, parent=None):
        super().__init__(parent)
        self.setObjectName("MetricCard")
        self._clickable = clickable
        self._icon_name = icon_name
        self._accent = accent
        if clickable:
            self.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))

        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(0)

        self.accent_bar = QFrame()
        self.accent_bar.setFixedWidth(4)
        self.accent_bar.setStyleSheet(
            f"background-color: {accent}; border-top-left-radius: 10px; border-bottom-left-radius: 10px;"
        )
        outer.addWidget(self.accent_bar)

        body = QVBoxLayout()
        body.setContentsMargins(12, 9, 12, 9)
        body.setSpacing(3)

        hdr = QHBoxLayout()
        hdr.setSpacing(6)
        self.lbl_title = QLabel(title.upper())
        self.lbl_title.setStyleSheet("font-size: 10.5px; font-weight: 700; color: #64748B; letter-spacing: 0.4px;")
        self.lbl_icon = QLabel()
        self.lbl_icon.setPixmap(svg_pixmap(icon_name, accent, 15))
        hdr.addWidget(self.lbl_title)
        hdr.addStretch(1)
        hdr.addWidget(self.lbl_icon)
        body.addLayout(hdr)

        self.lbl_value = QLabel("—")
        self.lbl_value.setStyleSheet("font-size: 20px; font-weight: 700;")
        body.addWidget(self.lbl_value)

        self.lbl_sub = QLabel("")
        self.lbl_sub.setStyleSheet("font-size: 11px; color: #64748B;")
        body.addWidget(self.lbl_sub)

        outer.addLayout(body, 1)

    def set_data(self, value: str, subtitle: str = "", accent: Optional[str] = None, icon_name: Optional[str] = None):
        self.lbl_value.setText(value)
        self.lbl_sub.setText(subtitle)
        if accent:
            self._accent = accent
            self.accent_bar.setStyleSheet(
                f"background-color: {accent}; border-top-left-radius: 10px; border-bottom-left-radius: 10px;"
            )
            self.lbl_value.setStyleSheet(f"font-size: 20px; font-weight: 700; color: {accent};")
        if icon_name:
            self._icon_name = icon_name
        self.lbl_icon.setPixmap(svg_pixmap(self._icon_name, self._accent, 15))

    def mousePressEvent(self, event):
        if self._clickable and event.button() == Qt.MouseButton.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class ToastBanner(QFrame):
    """Non-blocking floating toast notification anchored at bottom-right."""

    def __init__(self, parent: QWidget):
        super().__init__(parent)
        self.setVisible(False)
        self.setAttribute(Qt.WidgetAttribute.WA_TransparentForMouseEvents, True)
        lay = QHBoxLayout(self)
        lay.setContentsMargins(14, 9, 14, 9)
        self.lbl = QLabel("")
        self.lbl.setStyleSheet("color: #FFFFFF; font-size: 12.5px; font-weight: 600;")
        lay.addWidget(self.lbl)

        self._timer = QTimer(self)
        self._timer.setSingleShot(True)
        self._timer.timeout.connect(self.hide)

    def show_toast(self, message: str, kind: str = "info", duration_ms: int = 3400):
        colors = {
            "info": "#1E293B",
            "success": "#059669",
            "warn": "#D97706",
            "error": "#DC2626",
        }
        bg = colors.get(kind, "#1E293B")
        self.setStyleSheet(f"background-color: {bg}; border-radius: 8px;")
        self.lbl.setText(message)
        self.adjustSize()
        self.reposition()
        self.setVisible(True)
        self.raise_()
        self._timer.start(duration_ms)

    def reposition(self):
        if not self.parent():
            return
        pw = self.parent().width()
        ph = self.parent().height()
        self.move(max(12, pw - self.width() - 22), max(68, ph - self.height() - 64))
