import ctypes
import dataclasses
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import webbrowser
import zipfile
from typing import List, Optional, Tuple

import requests
from PyQt6.QtCore import Qt, QThread, QTimer, pyqtSignal
from PyQt6.QtGui import QCursor
from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPlainTextEdit,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from ui_theme import svg_icon, svg_pixmap


GITHUB_REPO = "NDCLI/worklog-report-tools"
GITHUB_RELEASES_API = f"https://api.github.com/repos/{GITHUB_REPO}/releases/latest"
UPDATE_TEMP_DIR_NAME = "ReportTools_AutoUpdate"


@dataclasses.dataclass
class ReleaseInfo:
    version: str
    tag_name: str
    title: str
    notes: str
    html_url: str
    asset_name: str = ""
    asset_url: str = ""
    asset_size: int = 0
    published_at: str = ""


def parse_version(version_str: str) -> Tuple[Tuple[int, ...], str]:
    """Parse version string like 'v2.3.1-beta.1' into ((2, 3, 1), 'beta.1')."""
    s = (version_str or "").strip()
    if s.lower().startswith("v"):
        s = s[1:]
    main_part, _, pre_part = s.partition("-")
    nums: List[int] = []
    for token in re.split(r"[._]+", main_part):
        m = re.match(r"^(\d+)", token.strip())
        if m:
            nums.append(int(m.group(1)))
    if not nums:
        nums = [0]
    return tuple(nums), pre_part.strip().lower()


def is_newer_version(latest: str, current: str) -> bool:
    """Return True if `latest` is strictly newer than `current`."""
    l_nums, l_pre = parse_version(latest)
    c_nums, c_pre = parse_version(current)
    max_len = max(len(l_nums), len(c_nums))
    l_pad = l_nums + (0,) * (max_len - len(l_nums))
    c_pad = c_nums + (0,) * (max_len - len(c_nums))
    if l_pad != c_pad:
        return l_pad > c_pad
    # Same numeric version: final release > pre-release
    if c_pre and not l_pre:
        return True
    if not c_pre and l_pre:
        return False
    return l_pre > c_pre


def format_bytes(num_bytes: float) -> str:
    if num_bytes < 0:
        return "0 B"
    for unit in ("B", "KB", "MB", "GB"):
        if num_bytes < 1024.0 or unit == "GB":
            if unit == "B":
                return f"{int(num_bytes)} B"
            return f"{num_bytes:.1f} {unit}"
        num_bytes /= 1024.0
    return f"{num_bytes:.1f} GB"


def is_frozen_exe() -> bool:
    return bool(getattr(sys, "frozen", False))


def get_runtime_mode() -> str:
    """
    Detect how the app is currently running:
    - 'portable_onefile': PyInstaller --onefile portable executable
    - 'portable_onedir':  PyInstaller --onedir folder executable
    - 'source':           Running from Python source
    """
    if not is_frozen_exe():
        return "source"
    meipass = getattr(sys, "_MEIPASS", None)
    exe_dir = os.path.normcase(os.path.dirname(os.path.abspath(sys.executable)))
    if meipass:
        meipass_norm = os.path.normcase(os.path.abspath(meipass))
        internal_dir = os.path.normcase(os.path.join(exe_dir, "_internal"))
        if meipass_norm != exe_dir and meipass_norm != internal_dir:
            return "portable_onefile"
    return "portable_onedir"


def get_update_temp_dir() -> str:
    return os.path.join(tempfile.gettempdir(), UPDATE_TEMP_DIR_NAME)


def cleanup_old_update_files():
    """Remove leftover temporary staging files from previous update runs."""
    tmp_dir = get_update_temp_dir()
    if not os.path.isdir(tmp_dir):
        return
    try:
        shutil.rmtree(tmp_dir, ignore_errors=True)
    except Exception:
        pass


def _select_best_asset(assets: list) -> Tuple[str, str, int]:
    """
    Pick the best downloadable asset for Windows portable update.
    Prefers standalone portable .exe (non-setup), then .zip, then any .exe.
    """
    candidates = []
    for a in assets or []:
        name = (a.get("name") or "").strip()
        url = (a.get("browser_download_url") or "").strip()
        size = int(a.get("size") or 0)
        if not name or not url:
            continue
        low = name.lower()
        if low.endswith((".sha256", ".txt", ".md", ".sig", ".yml", ".yaml", ".json", ".blockmap")):
            continue
        is_setup = any(k in low for k in ("setup", "installer", "install"))
        if low.endswith(".exe"):
            score = 100
            if not is_setup:
                score += 50
            if "portable" in low:
                score += 20
            if "reporttools" in low or "report_tool" in low:
                score += 10
            candidates.append((score, name, url, size))
        elif low.endswith(".zip"):
            score = 80
            if "portable" in low:
                score += 15
            if "win" in low or "reporttools" in low:
                score += 5
            candidates.append((score, name, url, size))

    if not candidates:
        return "", "", 0
    candidates.sort(key=lambda x: x[0], reverse=True)
    _, name, url, size = candidates[0]
    return name, url, size


def fetch_latest_release(current_version: str, repo: str = GITHUB_REPO) -> Optional[ReleaseInfo]:
    """
    Query GitHub Releases API for the latest release.
    Returns ReleaseInfo if a release exists, or None if no releases have been published.
    Raises RuntimeError on network/API errors.
    """
    url = f"https://api.github.com/repos/{repo}/releases/latest"
    headers = {
        "Accept": "application/vnd.github+json",
        "User-Agent": f"ReportTools/{current_version}",
    }
    try:
        resp = requests.get(url, headers=headers, timeout=(6, 12))
    except requests.RequestException as exc:
        raise RuntimeError(f"Không thể kết nối máy chủ cập nhật:\n{exc}") from exc

    if resp.status_code == 404:
        return None
    if resp.status_code == 403 and resp.headers.get("X-RateLimit-Remaining") == "0":
        raise RuntimeError("Đã vượt giới hạn truy vấn GitHub API. Vui lòng thử lại sau vài phút.")
    if resp.status_code >= 400:
        raise RuntimeError(f"Máy chủ cập nhật trả về lỗi HTTP {resp.status_code}.")

    try:
        data = resp.json()
    except Exception as exc:
        raise RuntimeError(f"Dữ liệu phản hồi không hợp lệ: {exc}") from exc

    tag_name = (data.get("tag_name") or "").strip()
    if not tag_name:
        return None

    version = tag_name[1:] if tag_name.lower().startswith("v") else tag_name
    title = (data.get("name") or tag_name).strip()
    notes = (data.get("body") or "").strip()
    html_url = (data.get("html_url") or f"https://github.com/{repo}/releases").strip()
    published_at = (data.get("published_at") or "").strip()
    asset_name, asset_url, asset_size = _select_best_asset(data.get("assets") or [])

    return ReleaseInfo(
        version=version,
        tag_name=tag_name,
        title=title,
        notes=notes,
        html_url=html_url,
        asset_name=asset_name,
        asset_url=asset_url,
        asset_size=asset_size,
        published_at=published_at,
    )


def _is_valid_pe_exe(file_path: str) -> bool:
    try:
        if not os.path.isfile(file_path) or os.path.getsize(file_path) < 64 * 1024:
            return False
        with open(file_path, "rb") as f:
            return f.read(2) == b"MZ"
    except Exception:
        return False


def _ps_quote(value: str) -> str:
    """Safely single-quote a string literal for PowerShell."""
    return "'" + value.replace("'", "''") + "'"


def _build_portable_updater_ps1(
    target_exe: str,
    staged_exe: str,
    update_dir: str,
    current_pid: int,
    parent_pid: int,
) -> str:
    """
    Build a PowerShell script that:
    1. Waits for the running app (and PyInstaller bootloader parent if matching) to exit.
    2. Replaces the EXACT portable executable in-place without creating any duplicate in the target folder.
    3. Cleans up temporary staging files and PyInstaller environment variables.
    4. Relaunches the updated executable.
    """
    backup_exe = os.path.join(update_dir, "rollback_backup.exe")
    target_dir = os.path.dirname(target_exe)
    log_file = os.path.join(tempfile.gettempdir(), "ReportTools_AutoUpdate.log")

    return f"""$ErrorActionPreference = 'Stop'

$TargetExe  = {_ps_quote(target_exe)}
$TargetDir  = {_ps_quote(target_dir)}
$StagedExe  = {_ps_quote(staged_exe)}
$BackupExe  = {_ps_quote(backup_exe)}
$UpdateDir  = {_ps_quote(update_dir)}
$LogFile    = {_ps_quote(log_file)}
$CurrentPid = {int(current_pid)}
$ParentPid  = {int(parent_pid)}

function Write-Log([string]$Msg) {{
    try {{
        $ts = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss.fff')
        Add-Content -LiteralPath $LogFile -Value "[$ts] $Msg" -Encoding UTF8 -ErrorAction SilentlyContinue
    }} catch {{}}
}}

try {{ Set-Content -LiteralPath $LogFile -Value "" -Encoding UTF8 -ErrorAction SilentlyContinue }} catch {{}}
Write-Log "Updater started. Target=$TargetExe Staged=$StagedExe CurrentPid=$CurrentPid ParentPid=$ParentPid"

function Test-PidAlive([int]$ProcId) {{
    if ($ProcId -le 0) {{ return $false }}
    return [bool](Get-Process -Id $ProcId -ErrorAction SilentlyContinue)
}}

# Collect PIDs to wait for: always current process, plus parent ONLY if it is the same portable exe (PyInstaller bootloader)
$WaitPids = @($CurrentPid)
if ($ParentPid -gt 0 -and $ParentPid -ne $CurrentPid) {{
    try {{
        $pProc = Get-Process -Id $ParentPid -ErrorAction SilentlyContinue
        if ($pProc -and $pProc.Path -and ($pProc.Path.ToLowerInvariant() -eq $TargetExe.ToLowerInvariant())) {{
            $WaitPids += $ParentPid
        }}
    }} catch {{}}
}}
Write-Log "Waiting for PIDs: $($WaitPids -join ', ')"

# 1. Wait for app & bootloader to exit completely (up to 45s)
for ($i = 0; $i -lt 225; $i++) {{
    $anyAlive = $false
    foreach ($p in $WaitPids) {{
        if (Test-PidAlive $p) {{
            $anyAlive = $true
            break
        }}
    }}
    if (-not $anyAlive) {{ break }}
    Start-Sleep -Milliseconds 200
}}

# Brief pause so Windows & antivirus release image handles
Start-Sleep -Milliseconds 400

if (-not (Test-Path -LiteralPath $StagedExe)) {{
    Write-Log "ERROR: StagedExe missing: $StagedExe"
    exit 1
}}

$ExpectedSize = (Get-Item -LiteralPath $StagedExe).Length
Write-Log "All target PIDs exited. ExpectedSize=$ExpectedSize"

# Keep a temporary rollback backup inside %TEMP% (never inside the portable folder)
if (Test-Path -LiteralPath $TargetExe) {{
    try {{
        Copy-Item -LiteralPath $TargetExe -Destination $BackupExe -Force -ErrorAction SilentlyContinue
    }} catch {{}}
}}

# 2. Replace the exact portable .exe in-place (retry up to 30s if still locked)
$Replaced = $false
$LastErr = ""
for ($attempt = 0; $attempt -lt 60; $attempt++) {{
    try {{
        # Open-and-overwrite in-place via Copy-Item preserves exact filename and creates no copy in target folder
        Copy-Item -LiteralPath $StagedExe -Destination $TargetExe -Force -ErrorAction Stop
        $ActualSize = (Get-Item -LiteralPath $TargetExe).Length
        if ($ActualSize -eq $ExpectedSize) {{
            $Replaced = $true
            Write-Log "Replaced successfully on attempt $($attempt + 1). ActualSize=$ActualSize"
            break
        }}
    }} catch {{
        $LastErr = $_.Exception.Message
    }}
    Start-Sleep -Milliseconds 500
}}

if (-not $Replaced) {{
    Write-Log "ERROR: Replace failed after retries. LastErr=$LastErr"
    # Restore backup if target was damaged
    if ((Test-Path -LiteralPath $BackupExe) -and (-not (Test-Path -LiteralPath $TargetExe))) {{
        try {{
            Copy-Item -LiteralPath $BackupExe -Destination $TargetExe -Force -ErrorAction SilentlyContinue
        }} catch {{}}
    }}
    try {{
        Add-Type -AssemblyName PresentationFramework -ErrorAction SilentlyContinue
        [System.Windows.MessageBox]::Show(
            "Không thể ghi đè file portable đang chạy:`n$TargetExe`n`nChi tiết: $LastErr`nVui lòng đóng toàn bộ cửa sổ ứng dụng và thử lại.",
            "Lỗi cập nhật Report Tools",
            'OK',
            'Error'
        ) | Out-Null
    }} catch {{}}
    exit 2
}}

# 3. Remove temporary staged & backup files before relaunching
Remove-Item -LiteralPath $StagedExe -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $BackupExe -Force -ErrorAction SilentlyContinue

# 4. Clear inherited PyInstaller / Qt temp env vars so the new .exe extracts cleanly
foreach ($envName in @(
    '_MEIPASS', '_MEIPASS2', '_PYI_APPLICATION_HOME_DIR',
    '_PYI_ARCHIVE_FILE', '_PYI_PARENT_PROCESS_LEVEL', '_PYI_SPLASH_IPC',
    'QT_PLUGIN_PATH', 'QML2_IMPORT_PATH', 'TCL_LIBRARY', 'TK_LIBRARY'
)) {{
    Remove-Item "Env:$envName" -ErrorAction SilentlyContinue
}}

# 5. Launch the newly replaced portable executable
Write-Log "Launching updated exe: $TargetExe"
Start-Process -FilePath $TargetExe -WorkingDirectory $TargetDir

# 6. Clean up this script and update temp folder
try {{
    Remove-Item -LiteralPath $PSCommandPath -Force -ErrorAction SilentlyContinue
    if (Test-Path -LiteralPath $UpdateDir) {{
        Remove-Item -LiteralPath $UpdateDir -Recurse -Force -ErrorAction SilentlyContinue
    }}
}} catch {{}}
Write-Log "Updater finished."
"""


def launch_in_place_updater(staged_exe: str, target_exe: Optional[str] = None) -> None:
    """
    Spawn a hidden detached PowerShell updater that replaces `target_exe` in-place
    with `staged_exe`, deletes all temporary files, and relaunches `target_exe`.
    """
    if sys.platform != "win32":
        raise RuntimeError("Tự động cập nhật thay thế trực tiếp hiện chỉ hỗ trợ Windows.")

    if not target_exe:
        if not is_frozen_exe():
            raise RuntimeError("Ứng dụng đang chạy từ mã nguồn Python, không thể tự ghi đè file .exe.")
        target_exe = os.path.abspath(sys.executable)

    if not _is_valid_pe_exe(staged_exe):
        raise RuntimeError("File cập nhật tải về không hợp lệ hoặc bị hỏng.")

    # Verify target directory is writable before quitting the app
    target_dir = os.path.dirname(target_exe)
    if not os.access(target_dir, os.W_OK):
        raise RuntimeError(
            f"Không có quyền ghi vào thư mục chứa ứng dụng:\n{target_dir}\n"
            "Vui lòng chạy ứng dụng với quyền quản trị (Administrator) hoặc di chuyển bản portable sang thư mục khác."
        )

    update_dir = get_update_temp_dir()
    os.makedirs(update_dir, exist_ok=True)
    script_path = os.path.join(update_dir, "apply_update.ps1")

    ps_content = _build_portable_updater_ps1(
        target_exe=target_exe,
        staged_exe=os.path.abspath(staged_exe),
        update_dir=update_dir,
        current_pid=os.getpid(),
        parent_pid=os.getppid(),
    )
    # UTF-8 with BOM is required for Windows PowerShell 5.1 to handle Unicode paths properly
    with open(script_path, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write(ps_content)

    # Strip PyInstaller env vars for the child PowerShell process as well
    clean_env = os.environ.copy()
    for k in (
        "_MEIPASS",
        "_MEIPASS2",
        "_PYI_APPLICATION_HOME_DIR",
        "_PYI_ARCHIVE_FILE",
        "_PYI_PARENT_PROCESS_LEVEL",
        "_PYI_SPLASH_IPC",
        "QT_PLUGIN_PATH",
        "QML2_IMPORT_PATH",
    ):
        clean_env.pop(k, None)

    # NOTE: Never combine CREATE_NO_WINDOW with DETACHED_PROCESS on Windows —
    # they are mutually exclusive and cause powershell.exe to exit immediately without running -File!
    creationflags = (
        getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000)
        | getattr(subprocess, "CREATE_NEW_PROCESS_GROUP", 0x00000200)
    )
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    startupinfo.wShowWindow = 0  # SW_HIDE

    subprocess.Popen(
        [
            "powershell.exe",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-WindowStyle",
            "Hidden",
            "-File",
            script_path,
        ],
        cwd=tempfile.gettempdir(),
        env=clean_env,
        close_fds=True,
        creationflags=creationflags,
        startupinfo=startupinfo,
    )


class UpdateCheckThread(QThread):
    """Background worker to check GitHub Releases for a newer version."""

    update_available = pyqtSignal(object)  # ReleaseInfo
    up_to_date = pyqtSignal(str)  # latest or current version
    failed = pyqtSignal(str)

    def __init__(self, current_version: str, repo: str = GITHUB_REPO, parent=None):
        super().__init__(parent)
        self.current_version = current_version
        self.repo = repo

    def run(self):
        try:
            release = fetch_latest_release(self.current_version, self.repo)
            if not release:
                self.up_to_date.emit(self.current_version)
                return
            if is_newer_version(release.version, self.current_version):
                self.update_available.emit(release)
            else:
                self.up_to_date.emit(release.version)
        except Exception as exc:
            self.failed.emit(str(exc))


class UpdateDownloadThread(QThread):
    """Background worker to stream-download, extract (if zip), and verify the update binary."""

    status_changed = pyqtSignal(str)
    progress_changed = pyqtSignal(int, int, float)  # downloaded, total, speed_bps
    finished_ok = pyqtSignal(str)  # path to staged .exe
    cancelled = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(self, release: ReleaseInfo, current_version: str, parent=None):
        super().__init__(parent)
        self.release = release
        self.current_version = current_version
        self._cancel_requested = False

    def cancel(self):
        self._cancel_requested = True

    def run(self):
        tmp_dir = get_update_temp_dir()
        download_path = ""
        try:
            if not self.release.asset_url:
                raise RuntimeError("Bản phát hành này chưa đính kèm tệp cài đặt/portable (.exe hoặc .zip).")

            cleanup_old_update_files()
            os.makedirs(tmp_dir, exist_ok=True)

            safe_name = os.path.basename(self.release.asset_name or "ReportTools_update.exe")
            download_path = os.path.join(tmp_dir, safe_name)

            self.status_changed.emit("Đang kết nối máy chủ tải xuống...")
            headers = {"User-Agent": f"ReportTools/{self.current_version}"}
            with requests.get(
                self.release.asset_url,
                headers=headers,
                stream=True,
                allow_redirects=True,
                timeout=(10, 60),
            ) as resp:
                if resp.status_code >= 400:
                    raise RuntimeError(f"Tải bản cập nhật thất bại (HTTP {resp.status_code}).")

                total_size = int(resp.headers.get("Content-Length") or self.release.asset_size or 0)
                downloaded = 0
                start_ts = time.monotonic()
                last_emit_ts = 0.0

                self.status_changed.emit("Đang tải bản cập nhật...")
                with open(download_path, "wb") as out_f:
                    for chunk in resp.iter_content(chunk_size=128 * 1024):
                        if self._cancel_requested:
                            out_f.close()
                            try:
                                os.remove(download_path)
                            except Exception:
                                pass
                            self.cancelled.emit()
                            return
                        if not chunk:
                            continue
                        out_f.write(chunk)
                        downloaded += len(chunk)
                        now = time.monotonic()
                        if now - last_emit_ts >= 0.08 or (total_size > 0 and downloaded >= total_size):
                            elapsed = max(0.05, now - start_ts)
                            speed = downloaded / elapsed
                            self.progress_changed.emit(downloaded, total_size, speed)
                            last_emit_ts = now

            if self.release.asset_size > 0 and downloaded < self.release.asset_size:
                raise RuntimeError(
                    f"Tệp tải xuống chưa hoàn tất ({format_bytes(downloaded)} / {format_bytes(self.release.asset_size)})."
                )

            self.status_changed.emit("Đang kiểm tra tính toàn vẹn tệp tải xuống...")
            staged_exe = self._prepare_staged_exe(download_path, tmp_dir)

            if not _is_valid_pe_exe(staged_exe):
                raise RuntimeError("Tệp tải về không phải là tệp thực thi Windows (.exe) hợp lệ.")

            self.finished_ok.emit(staged_exe)
        except Exception as exc:
            if download_path and os.path.isfile(download_path):
                try:
                    os.remove(download_path)
                except Exception:
                    pass
            self.failed.emit(str(exc))

    def _prepare_staged_exe(self, download_path: str, tmp_dir: str) -> str:
        low = download_path.lower()
        if low.endswith(".exe"):
            return download_path

        if low.endswith(".zip"):
            self.status_changed.emit("Đang giải nén gói cập nhật...")
            extract_dir = os.path.join(tmp_dir, "extracted")
            os.makedirs(extract_dir, exist_ok=True)
            with zipfile.ZipFile(download_path, "r") as zf:
                zf.extractall(extract_dir)
            try:
                os.remove(download_path)
            except Exception:
                pass

            current_name = os.path.basename(sys.executable).lower() if is_frozen_exe() else "reporttools.exe"
            exe_candidates: List[Tuple[int, str]] = []
            for root, _, files in os.walk(extract_dir):
                for fn in files:
                    if fn.lower().endswith(".exe"):
                        full_p = os.path.join(root, fn)
                        score = os.path.getsize(full_p)
                        if fn.lower() == current_name:
                            score += 10**12
                        elif "reporttools" in fn.lower():
                            score += 10**11
                        exe_candidates.append((score, full_p))

            if not exe_candidates:
                raise RuntimeError("Không tìm thấy file .exe bên trong gói cập nhật (.zip).")
            exe_candidates.sort(key=lambda x: x[0], reverse=True)
            staged = os.path.join(tmp_dir, "staged_update.exe")
            shutil.move(exe_candidates[0][1], staged)
            shutil.rmtree(extract_dir, ignore_errors=True)
            return staged

        raise RuntimeError(f"Định dạng gói cập nhật chưa được hỗ trợ: {os.path.basename(download_path)}")


class UpdateBanner(QFrame):
    """Non-intrusive in-app notification bar shown when a newer version is detected."""

    update_clicked = pyqtSignal()
    details_clicked = pyqtSignal()
    dismissed = pyqtSignal()

    def __init__(self, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setObjectName("UpdateBanner")
        self.setVisible(False)
        self._release: Optional[ReleaseInfo] = None

        lay = QHBoxLayout(self)
        lay.setContentsMargins(16, 7, 14, 7)
        lay.setSpacing(10)

        self.lbl_icon = QLabel()
        self.lbl_icon.setFixedSize(18, 18)
        lay.addWidget(self.lbl_icon)

        self.lbl_msg = QLabel("")
        self.lbl_msg.setWordWrap(False)
        lay.addWidget(self.lbl_msg, 1)

        self.btn_details = QPushButton("Xem thay đổi")
        self.btn_details.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_details.setFixedHeight(27)
        self.btn_details.clicked.connect(self.details_clicked.emit)
        lay.addWidget(self.btn_details)

        self.btn_update = QPushButton("Cập nhật ngay")
        self.btn_update.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_update.setFixedHeight(27)
        self.btn_update.clicked.connect(self.update_clicked.emit)
        lay.addWidget(self.btn_update)

        self.btn_close = QPushButton("✕")
        self.btn_close.setToolTip("Ẩn thông báo này")
        self.btn_close.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_close.setFixedSize(26, 26)
        self.btn_close.clicked.connect(self._on_dismiss)
        lay.addWidget(self.btn_close)

    def show_release(self, release: ReleaseInfo, current_version: str):
        self._release = release
        size_txt = f" • {format_bytes(release.asset_size)}" if release.asset_size > 0 else ""
        self.lbl_msg.setText(
            f"<b>Có phiên bản mới v{release.version}</b> "
            f"(đang dùng v{current_version}{size_txt}) — Nhấn <b>Cập nhật ngay</b> để tự động tải và khởi động lại."
        )
        self.setVisible(True)

    def apply_theme(self, theme: dict):
        is_dark = theme.get("name") == "dark"
        bg = "#172554" if is_dark else "#EFF6FF"
        border = "#1D4ED8" if is_dark else "#93C5FD"
        fg = "#DBEAFE" if is_dark else "#1E3A8A"
        accent = theme.get("accent", "#2563EB")
        accent_hover = theme.get("accent_hover", "#1D4ED8")

        self.setStyleSheet(
            f"QFrame#UpdateBanner {{ background-color: {bg}; border-bottom: 1px solid {border}; }}"
        )
        self.lbl_icon.setPixmap(svg_pixmap("download", accent, 16))
        self.lbl_msg.setStyleSheet(f"color: {fg}; font-size: 12.5px; background: transparent;")
        self.btn_details.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {fg}; border: 1px solid {border}; "
            f"border-radius: 6px; padding: 2px 10px; font-size: 11.5px; font-weight: 600; }}"
            f"QPushButton:hover {{ background-color: {theme.get('bg_hover', '#DBEAFE')}; }}"
        )
        self.btn_update.setStyleSheet(
            f"QPushButton {{ background-color: {accent}; color: #FFFFFF; border: 1px solid {accent}; "
            f"border-radius: 6px; padding: 2px 12px; font-size: 11.5px; font-weight: 700; }}"
            f"QPushButton:hover {{ background-color: {accent_hover}; }}"
        )
        self.btn_close.setStyleSheet(
            f"QPushButton {{ background: transparent; color: {fg}; border: none; "
            f"border-radius: 4px; font-size: 12px; font-weight: 700; padding: 0; }}"
            f"QPushButton:hover {{ background-color: {border}; color: #FFFFFF; }}"
        )

    def _on_dismiss(self):
        self.setVisible(False)
        self.dismissed.emit()


class UpdateDialog(QDialog):
    """Modal dialog displaying release notes, live download progress, and auto-restart flow."""

    def __init__(
        self,
        release: ReleaseInfo,
        current_version: str,
        theme: dict,
        auto_start: bool = False,
        parent: Optional[QWidget] = None,
    ):
        super().__init__(parent)
        self.release = release
        self.current_version = current_version
        self.theme = theme
        self.download_thread: Optional[UpdateDownloadThread] = None
        self._is_updating = False
        self._is_restarting = False

        self.setWindowTitle(f"Cập nhật phiên bản mới — v{release.version}")
        self.setMinimumWidth(520)
        self.setMaximumWidth(620)
        self.setModal(True)

        self._build_ui()

        if auto_start:
            QTimer.singleShot(150, self.start_update)

    def _build_ui(self):
        lay = QVBoxLayout(self)
        lay.setContentsMargins(20, 18, 20, 18)
        lay.setSpacing(12)

        # Header row
        hdr = QHBoxLayout()
        hdr.setSpacing(12)

        icon_box = QLabel()
        icon_box.setFixedSize(40, 40)
        icon_box.setAlignment(Qt.AlignmentFlag.AlignCenter)
        icon_box.setStyleSheet("background-color: #2563EB; border-radius: 10px;")
        icon_box.setPixmap(svg_pixmap("download", "#FFFFFF", 20))
        hdr.addWidget(icon_box)

        title_col = QVBoxLayout()
        title_col.setSpacing(2)
        lbl_title = QLabel(f"Phiên bản mới v{self.release.version}")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: 700;")
        size_part = f" • Dung lượng: {format_bytes(self.release.asset_size)}" if self.release.asset_size > 0 else ""
        lbl_sub = QLabel(f"Phiên bản hiện tại: v{self.current_version} → v{self.release.version}{size_part}")
        lbl_sub.setStyleSheet(f"font-size: 12px; color: {self.theme.get('text_muted', '#64748B')};")
        title_col.addWidget(lbl_title)
        title_col.addWidget(lbl_sub)
        hdr.addLayout(title_col, 1)
        lay.addLayout(hdr)

        # Runtime mode info box
        mode = get_runtime_mode()
        mode_frame = QFrame()
        mode_frame.setStyleSheet(
            f"background-color: {self.theme.get('bg_subtle', '#DCE4EE')}; "
            f"border: 1px solid {self.theme.get('border', '#BCC8D8')}; border-radius: 8px;"
        )
        mf_lay = QVBoxLayout(mode_frame)
        mf_lay.setContentsMargins(12, 8, 12, 8)
        mf_lay.setSpacing(2)

        if mode in ("portable_onefile", "portable_onedir"):
            exe_path = os.path.abspath(sys.executable)
            lbl_mode = QLabel(
                "<b>Chế độ Portable:</b> Bản mới sẽ thay thế trực tiếp file đang chạy (không tạo bản sao mới) "
                "và tự động khởi động lại ứng dụng."
            )
            lbl_mode.setWordWrap(True)
            lbl_mode.setStyleSheet("font-size: 12px; border: none; background: transparent;")
            lbl_path = QLabel(f"Tệp mục tiêu: {exe_path}")
            lbl_path.setWordWrap(True)
            lbl_path.setStyleSheet(
                f"font-size: 11px; color: {self.theme.get('text_muted', '#64748B')}; border: none; background: transparent;"
            )
            mf_lay.addWidget(lbl_mode)
            mf_lay.addWidget(lbl_path)
        else:
            lbl_mode = QLabel(
                "<b>Đang chạy từ mã nguồn Python:</b> Tự động ghi đè chỉ áp dụng cho bản đóng gói Portable (.exe). "
                "Bạn vẫn có thể mở trang GitHub Release để tải bản mới."
            )
            lbl_mode.setWordWrap(True)
            lbl_mode.setStyleSheet("font-size: 12px; border: none; background: transparent;")
            mf_lay.addWidget(lbl_mode)
        lay.addWidget(mode_frame)

        # Release notes
        lbl_notes_hdr = QLabel("Nội dung cập nhật:")
        lbl_notes_hdr.setStyleSheet("font-size: 12px; font-weight: 700;")
        lay.addWidget(lbl_notes_hdr)

        self.txt_notes = QPlainTextEdit()
        self.txt_notes.setReadOnly(True)
        self.txt_notes.setFixedHeight(140)
        notes_text = self.release.notes or self.release.title or "Cải thiện hiệu năng và sửa lỗi."
        self.txt_notes.setPlainText(notes_text)
        lay.addWidget(self.txt_notes)

        # Progress section (shown once update starts)
        self.progress_wrap = QWidget()
        pw_lay = QVBoxLayout(self.progress_wrap)
        pw_lay.setContentsMargins(0, 4, 0, 0)
        pw_lay.setSpacing(6)

        self.lbl_status = QLabel("Sẵn sàng cập nhật.")
        self.lbl_status.setStyleSheet("font-size: 12px; font-weight: 600;")
        pw_lay.addWidget(self.lbl_status)

        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.progress_bar.setFixedHeight(18)
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat("%p%")
        pw_lay.addWidget(self.progress_bar)

        self.lbl_detail = QLabel("")
        self.lbl_detail.setStyleSheet(f"font-size: 11.5px; color: {self.theme.get('text_muted', '#64748B')};")
        pw_lay.addWidget(self.lbl_detail)

        self.progress_wrap.setVisible(False)
        lay.addWidget(self.progress_wrap)

        # Action buttons
        btn_row = QHBoxLayout()
        btn_row.setSpacing(8)

        self.btn_open_web = QPushButton("Mở trên GitHub")
        self.btn_open_web.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_open_web.clicked.connect(lambda: webbrowser.open(self.release.html_url))
        btn_row.addWidget(self.btn_open_web)

        btn_row.addStretch(1)

        self.btn_cancel = QPushButton("Để sau")
        self.btn_cancel.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_cancel.clicked.connect(self._on_cancel_clicked)
        btn_row.addWidget(self.btn_cancel)

        self.btn_start = QPushButton(" Cập nhật ngay")
        self.btn_start.setObjectName("PrimaryBtn")
        self.btn_start.setIcon(svg_icon("download", "#FFFFFF", 14))
        self.btn_start.setCursor(QCursor(Qt.CursorShape.PointingHandCursor))
        self.btn_start.clicked.connect(self.start_update)
        btn_row.addWidget(self.btn_start)

        lay.addLayout(btn_row)

    def start_update(self):
        if self._is_updating:
            return

        if not is_frozen_exe():
            ans = QMessageBox.information(
                self,
                "Đang chạy từ mã nguồn",
                "Bạn đang chạy ứng dụng từ mã nguồn Python (python main.py).\n"
                "Chức năng tự động tải và ghi đè chỉ hoạt động trên bản đóng gói Portable (.exe).\n\n"
                "Bạn có muốn mở trang tải bản phát hành trên GitHub không?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ans == QMessageBox.StandardButton.Yes:
                webbrowser.open(self.release.html_url)
            return

        if not self.release.asset_url:
            ans = QMessageBox.warning(
                self,
                "Chưa có tệp tải tự động",
                f"Phiên bản v{self.release.version} trên GitHub chưa đính kèm tệp .exe hoặc .zip.\n"
                "Mở trang GitHub Release để kiểm tra thủ công?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            )
            if ans == QMessageBox.StandardButton.Yes:
                webbrowser.open(self.release.html_url)
            return

        self._is_updating = True
        self.progress_wrap.setVisible(True)
        self.btn_start.setEnabled(False)
        self.btn_start.setText(" Đang cập nhật...")
        self.btn_cancel.setText("Hủy tải")
        self.progress_bar.setRange(0, 0)  # indeterminate until Content-Length arrives
        self.lbl_status.setText("Đang kết nối máy chủ tải xuống...")
        self.lbl_detail.setText(self.release.asset_name or "")

        self.download_thread = UpdateDownloadThread(self.release, self.current_version, self)
        self.download_thread.status_changed.connect(self._on_status)
        self.download_thread.progress_changed.connect(self._on_progress)
        self.download_thread.finished_ok.connect(self._on_download_finished)
        self.download_thread.cancelled.connect(self._on_download_cancelled)
        self.download_thread.failed.connect(self._on_download_failed)
        self.download_thread.start()

    def _on_status(self, msg: str):
        self.lbl_status.setText(msg)

    def _on_progress(self, downloaded: int, total: int, speed_bps: float):
        if total > 0:
            if self.progress_bar.maximum() == 0:
                self.progress_bar.setRange(0, 100)
            pct = min(100, int((downloaded * 100) / total))
            self.progress_bar.setValue(pct)
            self.lbl_detail.setText(
                f"Đã tải {format_bytes(downloaded)} / {format_bytes(total)} ({pct}%) — Tốc độ: {format_bytes(speed_bps)}/s"
            )
        else:
            self.lbl_detail.setText(f"Đã tải {format_bytes(downloaded)} — Tốc độ: {format_bytes(speed_bps)}/s")

    def _on_download_finished(self, staged_exe: str):
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(100)
        self.btn_cancel.setEnabled(False)
        self.lbl_status.setText("Tải hoàn tất! Đang thay thế bản portable và khởi động lại...")
        self.lbl_detail.setText("Ứng dụng sẽ tự đóng và mở lại phiên bản mới trong giây lát.")

        try:
            launch_in_place_updater(staged_exe)
        except Exception as exc:
            self._is_updating = False
            self.btn_start.setEnabled(True)
            self.btn_start.setText(" Thử lại")
            self.btn_cancel.setEnabled(True)
            self.btn_cancel.setText("Đóng")
            QMessageBox.critical(self, "Lỗi cập nhật", f"Không thể khởi chạy trình cập nhật:\n{exc}")
            return

        self._is_restarting = True
        QTimer.singleShot(550, self._quit_for_restart)

    def _quit_for_restart(self):
        self.accept()
        if self.parent() is not None and hasattr(self.parent(), "close"):
            self.parent().close()
        app = QApplication.instance()
        if app is not None:
            app.quit()

    def _on_download_cancelled(self):
        self._is_updating = False
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Đã hủy tải bản cập nhật.")
        self.lbl_detail.setText("")
        self.btn_start.setEnabled(True)
        self.btn_start.setText(" Cập nhật ngay")
        self.btn_cancel.setText("Đóng")

    def _on_download_failed(self, err_msg: str):
        self._is_updating = False
        self.progress_bar.setRange(0, 100)
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Cập nhật thất bại.")
        self.lbl_detail.setText(err_msg)
        self.btn_start.setEnabled(True)
        self.btn_start.setText(" Thử lại")
        self.btn_cancel.setText("Đóng")
        QMessageBox.critical(self, "Lỗi tải bản cập nhật", err_msg)

    def _on_cancel_clicked(self):
        if self._is_restarting:
            return
        if self.download_thread and self.download_thread.isRunning():
            self.lbl_status.setText("Đang hủy tải xuống...")
            self.btn_cancel.setEnabled(False)
            self.download_thread.cancel()
        else:
            self.reject()

    def closeEvent(self, event):
        if self._is_restarting:
            event.accept()
            return
        if self.download_thread and self.download_thread.isRunning():
            self.download_thread.cancel()
            self.download_thread.wait(2000)
        event.accept()
