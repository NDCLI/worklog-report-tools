import json
import os
import time
import urllib.parse
from typing import Callable, Dict, Optional

import requests
from PyQt6.QtCore import Qt, QThread, QTimer, QUrl, pyqtSignal
from PyQt6.QtWidgets import (
    QDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
)
from PyQt6.QtWebEngineCore import QWebEnginePage, QWebEngineProfile
from PyQt6.QtWebEngineWidgets import QWebEngineView


_SHARED_PROFILE: Optional[QWebEngineProfile] = None


def get_sp_profile(cache_dir: str) -> QWebEngineProfile:
    """Return a singleton persistent QWebEngineProfile for SharePoint login."""
    global _SHARED_PROFILE
    if _SHARED_PROFILE is None:
        os.makedirs(cache_dir, exist_ok=True)
        prof_dir = os.path.join(cache_dir, "sp_web_profile")
        http_dir = os.path.join(cache_dir, "sp_web_cache")
        os.makedirs(prof_dir, exist_ok=True)
        os.makedirs(http_dir, exist_ok=True)
        prof = QWebEngineProfile("WorklogSharePointProfile")
        prof.setPersistentStoragePath(prof_dir)
        prof.setCachePath(http_dir)
        prof.setPersistentCookiesPolicy(QWebEngineProfile.PersistentCookiesPolicy.ForcePersistentCookies)
        _SHARED_PROFILE = prof
    return _SHARED_PROFILE


def _cookies_path(cache_dir: str) -> str:
    return os.path.join(cache_dir, "sp_cookies.json")


def load_saved_sp_cookies(cache_dir: str) -> Dict[str, str]:
    p = _cookies_path(cache_dir)
    if not os.path.isfile(p):
        return {}
    try:
        data = json.load(open(p, "r", encoding="utf-8"))
        return data.get("cookies") or {}
    except Exception:
        return {}


def save_sp_cookies(cache_dir: str, host: str, cookies: Dict[str, str]):
    os.makedirs(cache_dir, exist_ok=True)
    try:
        with open(_cookies_path(cache_dir), "w", encoding="utf-8") as f:
            json.dump({"host": host, "cookies": cookies}, f, indent=2)
    except Exception:
        pass


def has_saved_sp_session(cache_dir: str) -> Optional[str]:
    p = _cookies_path(cache_dir)
    if not os.path.isfile(p):
        return None
    try:
        data = json.load(open(p, "r", encoding="utf-8"))
        ck = data.get("cookies") or {}
        if ck.get("FedAuth") or ck.get("rtFa"):
            return data.get("host") or "SharePoint"
    except Exception:
        pass
    return None


def clear_sp_session(cache_dir: str):
    prof = get_sp_profile(cache_dir)
    try:
        prof.cookieStore().deleteAllCookies()
    except Exception:
        pass
    for fn in ("sp_cookies.json", "sp_session.txt"):
        p = os.path.join(cache_dir, fn)
        if os.path.isfile(p):
            try:
                os.remove(p)
            except Exception:
                pass


def extract_sp_site_url(*urls: str) -> str:
    """Return the SharePoint site URL (NOT Excel Online) to open for clean sign-in."""
    for u in urls:
        u = (u or "").strip()
        if not u.startswith("http"):
            continue
        parsed = urllib.parse.urlparse(u)
        if not parsed.netloc:
            continue
        decoded = urllib.parse.unquote(parsed.path or "")
        idx = decoded.lower().find("/sites/")
        if idx >= 0:
            parts = [p for p in decoded[idx:].split("/") if p]
            if len(parts) >= 2:
                return f"{parsed.scheme}://{parsed.netloc}/sites/{parts[1]}"
        return f"{parsed.scheme}://{parsed.netloc}"
    return "https://www.office.com/login"


def build_download_candidates(raw_url: str):
    """Return ordered list of direct download URLs for a SharePoint link (including Doc.aspx?sourcedoc=...)."""
    url = (raw_url or "").strip()
    if not url:
        return []
    parsed = urllib.parse.urlparse(url)
    decoded_path = urllib.parse.unquote(parsed.path or "")
    qs = urllib.parse.parse_qs(parsed.query, keep_blank_values=True)
    candidates = []

    idx = decoded_path.lower().find("/sites/")
    site_prefix = ""
    site_rel = ""
    if idx >= 0:
        site_rel = decoded_path[idx:]
        parts = [p for p in site_rel.split("/") if p]
        if len(parts) >= 2 and parts[0].lower() == "sites":
            site_prefix = f"/sites/{parts[1]}"

    # Handle _layouts/15/Doc.aspx?sourcedoc={GUID}&file=... links
    sourcedoc_vals = qs.get("sourcedoc") or qs.get("UniqueId") or []
    sourcedoc = (sourcedoc_vals[0] if sourcedoc_vals else "").strip()
    guid = sourcedoc.strip("{}").strip()
    file_vals = qs.get("file") or []
    file_name = (file_vals[0] if file_vals else "").strip()

    if site_prefix and guid:
        # 1. SharePoint REST API GetFileById (works regardless of folder path)
        candidates.append(
            f"{parsed.scheme}://{parsed.netloc}{site_prefix}/_api/web/GetFileById('{guid}')/$value"
        )
        # 2. Classic SharePoint download.aspx with UniqueId
        uid_q = urllib.parse.urlencode({"UniqueId": guid})
        candidates.append(
            f"{parsed.scheme}://{parsed.netloc}{site_prefix}/_layouts/15/download.aspx?{uid_q}"
        )

    if site_prefix and site_rel and "_layouts/" not in site_rel.lower():
        # Standard direct path link: /sites/<site>/Shared Documents/file.xlsx
        escaped_rel = urllib.parse.quote(site_rel.replace("'", "''"), safe="/")
        candidates.append(
            f"{parsed.scheme}://{parsed.netloc}{site_prefix}/_api/web/"
            f"GetFileByServerRelativeUrl('{escaped_rel}')/$value"
        )
        src_q = urllib.parse.urlencode({"SourceUrl": site_rel}, quote_via=urllib.parse.quote)
        candidates.append(
            f"{parsed.scheme}://{parsed.netloc}{site_prefix}/_layouts/15/download.aspx?{src_q}"
        )
    elif site_prefix and file_name:
        # Fallback: try Shared Documents/<file_name>
        rel_guess = f"{site_prefix}/Shared Documents/{file_name}"
        escaped_rel = urllib.parse.quote(rel_guess.replace("'", "''"), safe="/")
        candidates.append(
            f"{parsed.scheme}://{parsed.netloc}{site_prefix}/_api/web/"
            f"GetFileByServerRelativeUrl('{escaped_rel}')/$value"
        )

    # Final fallback: Sharing link with download=1
    qs_dl = {k: list(v) for k, v in qs.items()}
    qs_dl.pop("web", None)
    qs_dl["download"] = ["1"]
    candidates.append(urllib.parse.urlunparse(parsed._replace(query=urllib.parse.urlencode(qs_dl, doseq=True))))
    return candidates


class SessionExpiredError(RuntimeError):
    """Raised when SharePoint cookies are missing or expired and re-login is needed."""


def download_sp_file_with_cookies(
    raw_url: str,
    dest_path: str,
    cookies: Dict[str, str],
    progress_cb: Optional[Callable[[int], None]] = None,
) -> str:
    """Stream-download an Excel file from SharePoint using captured FedAuth/rtFa cookies."""
    if not cookies or not (cookies.get("FedAuth") or cookies.get("rtFa")):
        raise SessionExpiredError("Chưa có phiên đăng nhập SharePoint.")

    candidates = build_download_candidates(raw_url)
    if not candidates:
        raise ValueError("Link SharePoint không hợp lệ.")

    os.makedirs(os.path.dirname(dest_path), exist_ok=True)
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Accept": "*/*",
        "Cache-Control": "no-cache, no-store, must-revalidate",
        "Pragma": "no-cache",
    }

    last_err = ""
    sp_timeout = (20, 120)
    max_retries = 3
    retryable_statuses = {429, 500, 502, 503, 504}

    for url in candidates:
        for attempt in range(max_retries):
            if attempt > 0:
                time.sleep(1.5 * attempt)
            resp = None
            tmp_path = dest_path + ".tmp"
            try:
                resp = requests.get(
                    url,
                    cookies=cookies,
                    headers=headers,
                    stream=True,
                    timeout=sp_timeout,
                    allow_redirects=False,
                )

                # Redirect to login or 401/403 -> session expired
                loc = (resp.headers.get("Location") or "").lower()
                if resp.status_code in (401, 403) or (
                    resp.status_code in (301, 302, 303, 307, 308)
                    and ("login.microsoftonline.com" in loc or "authenticate.aspx" in loc or "login.live.com" in loc)
                ):
                    resp.close()
                    raise SessionExpiredError("Phiên đăng nhập SharePoint đã hết hạn, cần đăng nhập lại.")

                # Follow non-login redirect once if needed
                if resp.status_code in (301, 302, 303, 307, 308) and loc:
                    redir_url = urllib.parse.urljoin(url, resp.headers["Location"])
                    resp.close()
                    resp = requests.get(
                        redir_url,
                        cookies=cookies,
                        headers=headers,
                        stream=True,
                        timeout=sp_timeout,
                        allow_redirects=True,
                    )

                if resp.status_code in retryable_statuses:
                    last_err = f"HTTP {resp.status_code}"
                    resp.close()
                    continue

                if resp.status_code != 200:
                    last_err = f"HTTP {resp.status_code}"
                    resp.close()
                    break

                total = int(resp.headers.get("Content-Length") or 0)
                written = 0
                first_chunk = b""
                with open(tmp_path, "wb") as f:
                    for chunk in resp.iter_content(chunk_size=65536):
                        if not chunk:
                            continue
                        if not first_chunk:
                            first_chunk = chunk[:4]
                        f.write(chunk)
                        written += len(chunk)
                        if progress_cb and total > 0:
                            progress_cb(min(99, int(written * 100 / total)))
                resp.close()
                resp = None

                if first_chunk.startswith(b"PK"):
                    os.replace(tmp_path, dest_path)
                    if progress_cb:
                        progress_cb(100)
                    return dest_path

                # Not a valid xlsx (probably HTML viewer/error page) -> try next candidate
                try:
                    os.remove(tmp_path)
                except Exception:
                    pass
                last_err = "Phản hồi không phải file .xlsx"
                break
            except SessionExpiredError:
                raise
            except requests.RequestException as exc:
                last_err = str(exc)
                if resp is not None:
                    try:
                        resp.close()
                    except Exception:
                        pass
                try:
                    if os.path.exists(tmp_path):
                        os.remove(tmp_path)
                except Exception:
                    pass
                continue

    raise RuntimeError(f"Không tải được file từ SharePoint ({last_err or 'lỗi không xác định'}).")


class SharePointLoginDialog(QDialog):
    """Embedded browser login window that captures SharePoint FedAuth/rtFa cookies and auto-closes."""

    def __init__(self, start_url: str, cache_dir: str, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Đăng nhập SharePoint / Microsoft 365")
        self.resize(960, 680)
        self.cache_dir = cache_dir
        self.captured_cookies: Dict[str, str] = {}
        self.logged_in_host = ""
        self._auto_closed = False

        lay = QVBoxLayout(self)
        lay.setContentsMargins(10, 10, 10, 10)
        lay.setSpacing(8)

        top = QHBoxLayout()
        self.lbl_info = QLabel("Đăng nhập tài khoản Microsoft công ty. Cửa sổ sẽ tự đóng ngay khi xác thực xong.")
        self.lbl_info.setStyleSheet("font-weight: 600; font-size: 12.5px;")
        top.addWidget(self.lbl_info, 1)

        self.btn_done = QPushButton("✓ Hoàn tất")
        self.btn_done.setObjectName("SuccessBtn")
        self.btn_done.clicked.connect(self._finish)
        top.addWidget(self.btn_done)
        lay.addLayout(top)

        self.prof = get_sp_profile(cache_dir)
        self.prof.cookieStore().cookieAdded.connect(self._on_cookie_added)
        self.prof.cookieStore().loadAllCookies()

        self.web = QWebEngineView(self)
        self.page = QWebEnginePage(self.prof, self.web)
        self.web.setPage(self.page)
        self.web.urlChanged.connect(self._on_url_changed)
        lay.addWidget(self.web, 1)

        self.web.setUrl(QUrl(start_url))

    def _on_cookie_added(self, cookie):
        dom = cookie.domain().lower()
        if "sharepoint.com" not in dom:
            return
        name = bytes(cookie.name()).decode("utf-8", "ignore")
        val = bytes(cookie.value()).decode("utf-8", "ignore")
        if name and val:
            self.captured_cookies[name] = val
            if not self.logged_in_host:
                self.logged_in_host = dom.lstrip(".")
            if self.captured_cookies.get("FedAuth") and self.captured_cookies.get("rtFa"):
                save_sp_cookies(self.cache_dir, self.logged_in_host, self.captured_cookies)
                self.lbl_info.setText(f"✓ Đã đăng nhập thành công ({self.logged_in_host})! Đang quay lại app...")
                self.lbl_info.setStyleSheet("font-weight: 700; font-size: 12.5px; color: #059669;")
                if not self._auto_closed:
                    self._auto_closed = True
                    QTimer.singleShot(500, self._finish)

    def _on_url_changed(self, qurl: QUrl):
        host = qurl.host().lower()
        if "sharepoint.com" in host:
            self.logged_in_host = qurl.host()

    def _finish(self):
        try:
            self.prof.cookieStore().cookieAdded.disconnect(self._on_cookie_added)
        except Exception:
            pass
        if self.captured_cookies.get("FedAuth") or self.captured_cookies.get("rtFa"):
            save_sp_cookies(self.cache_dir, self.logged_in_host or "SharePoint", self.captured_cookies)
            self.accept()
        else:
            self.reject()

    def closeEvent(self, event):
        try:
            self.prof.cookieStore().cookieAdded.disconnect(self._on_cookie_added)
        except Exception:
            pass
        super().closeEvent(event)


class SharePointDownloadThread(QThread):
    """Background HTTP downloader for 1-3 SharePoint .xlsx files using captured session cookies."""

    progress = pyqtSignal(int, str)
    finished_ok = pyqtSignal(str, str, str)
    need_login = pyqtSignal()
    failed = pyqtSignal(str)

    def __init__(self, wl_url: str, rw_url: str, cache_dir: str, dist_url: str = ""):
        super().__init__()
        self.wl_url = wl_url
        self.rw_url = rw_url
        self.dist_url = dist_url
        self.cache_dir = cache_dir

    def run(self):
        try:
            cookies = load_saved_sp_cookies(self.cache_dir)
            if not cookies or not (cookies.get("FedAuth") or cookies.get("rtFa")):
                self.need_login.emit()
                return

            wl_dest = ""
            rw_dest = ""
            dist_dest = ""
            if self.wl_url.strip():
                wl_dest = os.path.join(self.cache_dir, "Work_log_SharePoint.xlsx")
                self.progress.emit(10, "Đang tải Work_log từ SharePoint...")
                download_sp_file_with_cookies(
                    self.wl_url.strip(),
                    wl_dest,
                    cookies,
                    progress_cb=lambda p: self.progress.emit(10 + int(p * 0.30), f"Đang tải Work_log ({p}%)..."),
                )
            if self.rw_url.strip():
                rw_dest = os.path.join(self.cache_dir, "Inspection_And_Rework_SharePoint.xlsx")
                self.progress.emit(45, "Đang tải Inspection & Rework từ SharePoint...")
                download_sp_file_with_cookies(
                    self.rw_url.strip(),
                    rw_dest,
                    cookies,
                    progress_cb=lambda p: self.progress.emit(45 + int(p * 0.30), f"Đang tải Inspection & Rework ({p}%)..."),
                )
            if self.dist_url.strip():
                dist_dest = os.path.join(self.cache_dir, "Distribution_Sheet_SharePoint.xlsx")
                self.progress.emit(78, "Đang tải Distribution Sheet từ SharePoint...")
                download_sp_file_with_cookies(
                    self.dist_url.strip(),
                    dist_dest,
                    cookies,
                    progress_cb=lambda p: self.progress.emit(78 + int(p * 0.20), f"Đang tải Distribution Sheet ({p}%)..."),
                )
            self.progress.emit(100, "Đã tải xong từ SharePoint!")
            self.finished_ok.emit(wl_dest, rw_dest, dist_dest)
        except SessionExpiredError:
            self.need_login.emit()
        except Exception as exc:
            self.failed.emit(str(exc))
