from __future__ import annotations

import json
import re
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any, Callable
from tkinter import (
    BooleanVar,
    Button,
    Checkbutton,
    END,
    Entry,
    EXTENDED,
    Frame,
    IntVar,
    Label,
    Listbox,
    SINGLE,
    StringVar,
    Tk,
    Toplevel,
    filedialog,
    messagebox,
    ttk,
)
from tkinter.scrolledtext import ScrolledText
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode, unquote, urlparse
from urllib.request import Request, urlopen

try:
    import pystray
    from PIL import Image, ImageDraw
    TRAY_AVAILABLE = True
except Exception:
    TRAY_AVAILABLE = False


APP_TITLE = "Real-Debris STRM Maker"
SETTINGS_FILE = Path.home() / ".strm_maker_settings.json"
INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
REQUEST_TIMEOUT_SECONDS = 20
REQUEST_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) STRM Maker/1.0",
    "Accept": "*/*",
}
REAL_DEBRID_UNRESTRICT_URL = "https://api.real-debrid.com/rest/1.0/unrestrict/link"
REAL_DEBRID_DOWNLOADS_URL = "https://api.real-debrid.com/rest/1.0/downloads"
REAL_DEBRID_TORRENTS_URL = "https://api.real-debrid.com/rest/1.0/torrents"
WRITE_RETRY_COUNT = 3
WRITE_RETRY_DELAY_SECONDS = 0.25
LINK_RETRY_COUNT = 5
LINK_RETRY_DELAY_SECONDS = 2.0
SCHEDULER_POLL_SECONDS = 60  # Check every minute whether it is time to run


# ---------------------------------------------------------------------------
# Settings helpers
# ---------------------------------------------------------------------------

def load_settings() -> dict[str, Any]:
    try:
        return json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def save_settings(settings: dict[str, Any]) -> None:
    try:
        SETTINGS_FILE.write_text(
            json.dumps(settings, indent=2),
            encoding="utf-8",
        )
    except OSError:
        pass


def default_settings() -> dict[str, Any]:
    return {
        "output_folder": "",
        "real_debrid_token": "",
        "auto_check_enabled": False,
        "auto_check_interval_hours": 12,
        "auto_check_time": "",          # HH:MM or empty for "every N hours"
        "default_auto_folder": "",
        "keyword_rules": [],            # [{"keywords": ["WWE Raw"], "folder": "/path"}]
        "processed_torrent_ids": [],
    }


def merge_settings(loaded: dict[str, Any]) -> dict[str, Any]:
    result = default_settings()
    result.update(loaded)
    return result


# ---------------------------------------------------------------------------
# File / link utilities  (unchanged)
# ---------------------------------------------------------------------------

def clean_filename(name: str) -> str:
    name = unquote(name).strip()
    name = INVALID_FILENAME_CHARS.sub("_", name)
    name = re.sub(r"\s+", " ", name).strip(" .")
    return name or "stream"


def filename_from_url(link: str, index: int, use_url_names: bool) -> str:
    if not use_url_names:
        return f"stream_{index:03d}.strm"

    parsed = urlparse(link)
    path_name = Path(parsed.path).name
    stem = clean_filename(path_name)

    if stem.lower().endswith(".strm"):
        return stem

    if "." in stem:
        stem = stem.rsplit(".", 1)[0]

    if not stem or stem == "stream":
        host = clean_filename(parsed.netloc)
        stem = f"{host or 'stream'}_{index:03d}"

    return f"{stem}.strm"


def unique_path(folder: Path, filename: str) -> Path:
    candidate = folder / filename
    if not candidate.exists():
        return candidate

    stem = candidate.stem
    suffix = candidate.suffix
    counter = 2

    while True:
        candidate = folder / f"{stem} ({counter}){suffix}"
        if not candidate.exists():
            return candidate
        counter += 1


def write_text_with_retries(path: Path, content: str) -> str | None:
    last_error = ""
    for attempt in range(1, WRITE_RETRY_COUNT + 1):
        try:
            path.write_text(content, encoding="utf-8")
            return None
        except OSError as exc:
            last_error = str(exc)
            if attempt < WRITE_RETRY_COUNT:
                time.sleep(WRITE_RETRY_DELAY_SECONDS)

    return last_error or "Unknown write error"


def extract_magnets(text: str) -> list[str]:
    return [line.strip() for line in text.splitlines() if line.strip().startswith("magnet:?")]


def parse_api_error(exc: HTTPError) -> str:
    try:
        body = exc.read().decode("utf-8", "replace")
        data = json.loads(body)
        if isinstance(data, dict) and data.get("error"):
            return str(data["error"])
    except (OSError, json.JSONDecodeError):
        pass

    return str(exc)


# ---------------------------------------------------------------------------
# Real-Debrid API  (unchanged)
# ---------------------------------------------------------------------------

def real_debrid_request(token: str, url: str) -> Any:
    request = Request(
        url,
        headers={
            **REQUEST_HEADERS,
            "Authorization": f"Bearer {token}",
        },
        method="GET",
    )

    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        raise RuntimeError(parse_api_error(exc)) from exc


def unrestrict_real_debrid_link(link: str, token: str) -> list[dict[str, str]]:
    payload = urlencode({"link": link}).encode("utf-8")
    request = Request(
        REAL_DEBRID_UNRESTRICT_URL,
        data=payload,
        headers={
            **REQUEST_HEADERS,
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            data: Any = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        raise RuntimeError(parse_api_error(exc)) from exc

    if isinstance(data, list):
        results = data
    elif isinstance(data, dict):
        results = [data]
    else:
        raise RuntimeError("Real-Debrid returned an unexpected response.")

    resolved: list[dict[str, str]] = []
    for item in results:
        if not isinstance(item, dict):
            continue

        final_link = item.get("download") or item.get("link")
        if not final_link:
            continue

        resolved.append(
            {
                "link": str(final_link),
                "filename": str(item.get("filename") or ""),
            }
        )

    if not resolved:
        raise RuntimeError("Real-Debrid did not return a download link.")

    return resolved


def get_real_debrid_downloads(token: str) -> list[dict[str, Any]]:
    data = real_debrid_request(token, f"{REAL_DEBRID_DOWNLOADS_URL}?limit=5000")
    if not isinstance(data, list):
        raise RuntimeError("Real-Debrid returned an unexpected downloads list.")

    return [item for item in data if isinstance(item, dict)]


def get_real_debrid_torrents(token: str, limit: int = 100) -> list[dict[str, Any]]:
    data = real_debrid_request(token, f"{REAL_DEBRID_TORRENTS_URL}?limit={limit}")
    if not isinstance(data, list):
        raise RuntimeError("Real-Debrid returned an unexpected torrents list.")

    return [item for item in data if isinstance(item, dict)]


def get_real_debrid_torrent_info(token: str, torrent_id: str) -> dict[str, Any]:
    data = real_debrid_request(token, f"{REAL_DEBRID_TORRENTS_URL}/info/{torrent_id}")
    if not isinstance(data, dict):
        raise RuntimeError("Real-Debrid returned unexpected torrent details.")

    return data


def add_real_debrid_magnet(token: str, magnet: str) -> str:
    payload = urlencode({"magnet": magnet}).encode("utf-8")
    request = Request(
        f"{REAL_DEBRID_TORRENTS_URL}/addMagnet",
        data=payload,
        headers={
            **REQUEST_HEADERS,
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            data: Any = json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        raise RuntimeError(parse_api_error(exc)) from exc

    if not isinstance(data, dict) or not data.get("id"):
        raise RuntimeError("Real-Debrid did not return a torrent ID.")

    return str(data["id"])


VIDEO_EXTENSIONS = {
    ".mkv", ".mp4", ".avi", ".mov", ".wmv", ".m4v", ".ts", ".m2ts",
    ".mpg", ".mpeg", ".flv", ".webm", ".vob", ".divx", ".xvid",
    ".h264", ".h265", ".hevc", ".rmvb", ".3gp",
}


def select_video_files_real_debrid(token: str, torrent_id: str) -> None:
    """Select only video files in the torrent; fall back to all if none found."""
    torrent_info = get_real_debrid_torrent_info(token, torrent_id)
    files: list[dict[str, Any]] = torrent_info.get("files") or []

    video_ids = [
        str(f["id"])
        for f in files
        if isinstance(f, dict)
        and Path(str(f.get("path") or "")).suffix.lower() in VIDEO_EXTENSIONS
    ]

    file_selection = ",".join(video_ids) if video_ids else "all"

    payload = urlencode({"files": file_selection}).encode("utf-8")
    request = Request(
        f"{REAL_DEBRID_TORRENTS_URL}/selectFiles/{torrent_id}",
        data=payload,
        headers={
            **REQUEST_HEADERS,
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/x-www-form-urlencoded",
        },
        method="POST",
    )

    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS):
            return
    except HTTPError as exc:
        if exc.code == 202:
            return
        raise RuntimeError(parse_api_error(exc)) from exc


def find_real_debrid_download(link: str, token: str) -> list[dict[str, str]]:
    parsed = urlparse(link)
    code = parsed.path.rstrip("/").split("/")[-1]

    if not code:
        raise RuntimeError("Could not find the Real-Debrid download code in the URL.")

    for item in get_real_debrid_downloads(token):
        download = str(item.get("download") or "")
        original = str(item.get("link") or "")
        if code in download or code in original:
            final_link = download or original
            if final_link:
                return [
                    {
                        "link": final_link,
                        "filename": str(item.get("filename") or ""),
                    }
                ]

    raise RuntimeError("That Real-Debrid link was not found in your downloads list.")


def is_real_debrid_download_page(link: str) -> bool:
    parsed = urlparse(link)
    host = parsed.netloc.lower()
    return host.endswith("real-debrid.com") and parsed.path.startswith("/d/")


def resolve_redirect(link: str) -> list[dict[str, str]]:
    request = Request(link, headers=REQUEST_HEADERS, method="HEAD")
    try:
        with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
            final_link = response.geturl()
            content_type = response.headers.get("content-type", "")
            if final_link == link and "text/html" in content_type.lower():
                raise RuntimeError("The link opened as a web page, not a direct file.")
            return [{"link": final_link, "filename": ""}]
    except HTTPError as exc:
        if exc.code not in {403, 405, 501}:
            raise
    except URLError:
        raise

    fallback_headers = {**REQUEST_HEADERS, "Range": "bytes=0-0"}
    request = Request(link, headers=fallback_headers, method="GET")
    with urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
        final_link = response.geturl()
        content_type = response.headers.get("content-type", "")
        if final_link == link and "text/html" in content_type.lower():
            raise RuntimeError("The link opened as a web page, not a direct file.")
        return [{"link": final_link, "filename": ""}]


def resolve_links(link: str, real_debrid_token: str) -> list[dict[str, str]]:
    if real_debrid_token:
        try:
            return unrestrict_real_debrid_link(link, real_debrid_token)
        except RuntimeError:
            if is_real_debrid_download_page(link):
                return find_real_debrid_download(link, real_debrid_token)
            raise

    return resolve_redirect(link)


def short_size(byte_count: Any) -> str:
    try:
        size = float(byte_count)
    except (TypeError, ValueError):
        return "unknown size"

    units = ["B", "KB", "MB", "GB", "TB"]
    for unit in units:
        if size < 1024 or unit == units[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}"
        size /= 1024

    return "unknown size"


def torrent_label(torrent: dict[str, Any]) -> str:
    name = str(torrent.get("filename") or torrent.get("original_filename") or "Unnamed")
    status = str(torrent.get("status") or "unknown")
    progress = torrent.get("progress")
    added = str(torrent.get("added") or "")[:10]
    size = short_size(torrent.get("bytes") or torrent.get("original_bytes"))
    progress_text = f"{progress}%" if progress is not None else "?%"
    date_text = f" - {added}" if added else ""
    return f"{name} [{status} {progress_text}, {size}{date_text}]"


def torrent_links_from_info(torrent_info: dict[str, Any]) -> list[str]:
    links = torrent_info.get("links")
    if not isinstance(links, list):
        return []

    return [str(link) for link in links if str(link).strip()]


def get_torrent_links_with_retries(
    token: str,
    torrent_id: str,
    wait: Callable[[float], None] = time.sleep,
) -> tuple[list[str], str]:
    last_error = ""
    for attempt in range(1, LINK_RETRY_COUNT + 1):
        try:
            torrent_info = get_real_debrid_torrent_info(token, torrent_id)
            links = torrent_links_from_info(torrent_info)
            if links:
                return links, ""
            last_error = "no downloadable links yet"
        except (OSError, RuntimeError) as exc:
            last_error = str(exc)

        if attempt < LINK_RETRY_COUNT:
            wait(LINK_RETRY_DELAY_SECONDS)

    return [], last_error or "no downloadable links yet"


def resolve_links_with_retries(
    link: str,
    real_debrid_token: str,
    wait: Callable[[float], None] = time.sleep,
) -> tuple[list[dict[str, str]], str]:
    last_error = ""
    for attempt in range(1, LINK_RETRY_COUNT + 1):
        try:
            resolved_links = resolve_links(link, real_debrid_token)
            if resolved_links:
                return resolved_links, ""
            last_error = "no resolved links returned"
        except (HTTPError, URLError, TimeoutError, OSError, RuntimeError) as exc:
            last_error = str(exc)

        if attempt < LINK_RETRY_COUNT:
            wait(LINK_RETRY_DELAY_SECONDS)

    return [], last_error or "no resolved links returned"


def add_magnet_with_retries(
    token: str,
    magnet: str,
    wait: Callable[[float], None] = time.sleep,
) -> tuple[str, str]:
    last_error = ""
    for attempt in range(1, LINK_RETRY_COUNT + 1):
        try:
            torrent_id = add_real_debrid_magnet(token, magnet)
            select_video_files_real_debrid(token, torrent_id)
            return torrent_id, ""
        except (OSError, RuntimeError) as exc:
            last_error = str(exc)

        if attempt < LINK_RETRY_COUNT:
            wait(LINK_RETRY_DELAY_SECONDS)

    return "", last_error or "could not add magnet"


# ---------------------------------------------------------------------------
# Keyword rule matching
# ---------------------------------------------------------------------------

_SEPARATOR_RE = re.compile(r"[._\-]+")


def _normalize(text: str) -> str:
    """Replace dots, underscores, and hyphens with spaces for fuzzy matching."""
    return _SEPARATOR_RE.sub(" ", text).lower()


def match_keyword_rule(
    torrent_name: str,
    keyword_rules: list[dict[str, Any]],
) -> dict[str, Any] | None:
    """Return the first keyword rule that matches torrent_name, or None.

    Matching is case-insensitive and treats dots/underscores/hyphens as spaces
    so that e.g. 'WWE.Raw.S35E14' matches the keyword 'WWE Raw'.
    """
    name_norm = _normalize(torrent_name)
    for rule in keyword_rules:
        for kw in rule.get("keywords", []):
            if _normalize(kw.strip()) in name_norm:
                return rule
    return None


# ---------------------------------------------------------------------------
# Keyword Rule Dialog  (Add / Edit a single rule)
# ---------------------------------------------------------------------------

class KeywordRuleDialog(Toplevel):
    """Small dialog to add or edit a keyword rule."""

    def __init__(
        self,
        parent: Tk | Toplevel,
        rule: dict[str, Any] | None = None,
    ) -> None:
        super().__init__(parent)
        self.transient(parent)
        self.grab_set()
        self.title("Keyword Rule" if rule is None else "Edit Keyword Rule")
        self.resizable(False, False)
        self.result: dict[str, Any] | None = None

        existing_keywords = ", ".join(rule.get("keywords", [])) if rule else ""
        existing_folder = rule.get("folder", "") if rule else ""

        pad = {"padx": 10, "pady": 5}

        Label(self, text="Keywords (comma-separated):").grid(
            row=0, column=0, sticky="w", **pad
        )
        self._kw_var = StringVar(value=existing_keywords)
        kw_entry = Entry(self, textvariable=self._kw_var, width=46)
        kw_entry.grid(row=0, column=1, columnspan=2, sticky="ew", **pad)
        kw_entry.focus_set()

        Label(self, text="Destination Folder:").grid(
            row=1, column=0, sticky="w", **pad
        )
        self._folder_var = StringVar(value=existing_folder)
        Entry(self, textvariable=self._folder_var, width=38).grid(
            row=1, column=1, sticky="ew", **pad
        )
        Button(self, text="Browse…", command=self._browse).grid(
            row=1, column=2, sticky="w", padx=(0, 10)
        )

        hint = Label(
            self,
            text=(
                "Tip: use multiple keywords to match variations.\n"
                'e.g. "WWE Raw, Monday Night Raw"'
            ),
            fg="gray",
            justify="left",
        )
        hint.grid(row=2, column=0, columnspan=3, sticky="w", padx=10, pady=(0, 8))

        btn_frame = Frame(self)
        btn_frame.grid(row=3, column=0, columnspan=3, pady=(0, 10))
        Button(btn_frame, text="OK", width=10, command=self._ok).pack(
            side="left", padx=6
        )
        Button(btn_frame, text="Cancel", width=10, command=self.destroy).pack(
            side="left", padx=6
        )

        self.columnconfigure(1, weight=1)
        self.wait_window()

    def _browse(self) -> None:
        folder = filedialog.askdirectory(
            parent=self,
            title="Choose destination folder for this keyword",
            initialdir=self._folder_var.get() or str(Path.home()),
        )
        if folder:
            self._folder_var.set(folder)

    def _ok(self) -> None:
        raw_kw = self._kw_var.get().strip()
        folder = self._folder_var.get().strip()

        if not raw_kw:
            messagebox.showwarning("Keyword Rule", "Enter at least one keyword.", parent=self)
            return
        if not folder:
            messagebox.showwarning("Keyword Rule", "Choose a destination folder.", parent=self)
            return

        keywords = [k.strip() for k in raw_kw.split(",") if k.strip()]
        self.result = {"keywords": keywords, "folder": folder}
        self.destroy()


# ---------------------------------------------------------------------------
# Auto-Download Settings Window
# ---------------------------------------------------------------------------

class AutoDownloadSettingsWindow(Toplevel):
    """Settings window for the auto-download scheduler and keyword rules."""

    def __init__(self, parent: Tk, app: "StrmMakerApp") -> None:
        super().__init__(parent)
        self.app = app
        self.transient(parent)
        self.grab_set()
        self.title(f"{APP_TITLE} — Auto-Download Settings")
        self.minsize(620, 580)
        self.resizable(True, True)

        # Work on a copy of current rules so Cancel discards changes.
        self._rules: list[dict[str, Any]] = [
            dict(r) for r in app.settings.get("keyword_rules", [])
        ]

        self._build_ui()
        self.wait_window()

    # ------------------------------------------------------------------
    def _build_ui(self) -> None:
        s = self.app.settings
        outer = Frame(self, padx=14, pady=14)
        outer.pack(fill="both", expand=True)

        # ── Section: General auto-download folder ──────────────────────
        sec1 = ttk.LabelFrame(outer, text=" General Settings ", padding=10)
        sec1.pack(fill="x", pady=(0, 12))

        Label(sec1, text="Default auto-download folder:").grid(
            row=0, column=0, sticky="w"
        )
        self._default_folder_var = StringVar(value=s.get("default_auto_folder", ""))
        Entry(sec1, textvariable=self._default_folder_var, width=40).grid(
            row=0, column=1, sticky="ew", padx=(8, 4)
        )
        Button(sec1, text="Browse…", command=self._browse_default_folder).grid(
            row=0, column=2, sticky="w"
        )

        # ── Section: Scheduler ─────────────────────────────────────────
        sec2 = ttk.LabelFrame(outer, text=" Scheduler ", padding=10)
        sec2.pack(fill="x", pady=(0, 12))

        self._auto_enabled_var = BooleanVar(value=bool(s.get("auto_check_enabled", False)))
        Checkbutton(
            sec2,
            text="Enable automatic torrent scan",
            variable=self._auto_enabled_var,
            command=self._toggle_scheduler_fields,
        ).grid(row=0, column=0, columnspan=3, sticky="w")

        Label(sec2, text="Check interval (hours):").grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )
        self._interval_var = StringVar(value=str(s.get("auto_check_interval_hours", 12)))
        self._interval_entry = Entry(sec2, textvariable=self._interval_var, width=6)
        self._interval_entry.grid(row=1, column=1, sticky="w", padx=(8, 0), pady=(8, 0))

        Label(sec2, text="Daily check time (HH:MM, optional):").grid(
            row=2, column=0, sticky="w", pady=(8, 0)
        )
        self._check_time_var = StringVar(value=s.get("auto_check_time", ""))
        self._check_time_entry = Entry(
            sec2, textvariable=self._check_time_var, width=8
        )
        self._check_time_entry.grid(row=2, column=1, sticky="w", padx=(8, 0), pady=(8, 0))
        Label(
            sec2,
            text="Leave blank to check every N hours from start-up.",
            fg="gray",
        ).grid(row=2, column=2, sticky="w", padx=(6, 0), pady=(8, 0))

        sec2.columnconfigure(2, weight=1)

        # ── Section: Keyword Rules ─────────────────────────────────────
        sec3 = ttk.LabelFrame(outer, text=" Keyword Rules ", padding=10)
        sec3.pack(fill="both", expand=True, pady=(0, 12))

        Label(
            sec3,
            text=(
                "Add a rule for each show or keyword you want to auto-download.\n"
                "The first matching rule wins. Matching is case-insensitive substring search."
            ),
            fg="gray",
            justify="left",
        ).pack(anchor="w", pady=(0, 6))

        list_frame = Frame(sec3)
        list_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(list_frame, orient="vertical")
        self._rules_list = Listbox(
            list_frame,
            selectmode=SINGLE,
            yscrollcommand=scrollbar.set,
            height=8,
            activestyle="dotbox",
        )
        scrollbar.config(command=self._rules_list.yview)
        self._rules_list.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self._rules_list.bind("<Double-Button-1>", lambda _: self._edit_rule())

        btn_col = Frame(sec3)
        btn_col.pack(fill="x", pady=(6, 0))
        Button(btn_col, text="Add Rule", command=self._add_rule).pack(
            side="left", padx=(0, 6)
        )
        Button(btn_col, text="Edit Rule", command=self._edit_rule).pack(
            side="left", padx=(0, 6)
        )
        Button(btn_col, text="Remove Rule", command=self._remove_rule).pack(
            side="left", padx=(0, 6)
        )
        Button(btn_col, text="Move Up ▲", command=self._move_up).pack(
            side="left", padx=(0, 6)
        )
        Button(btn_col, text="Move Down ▼", command=self._move_down).pack(
            side="left"
        )

        # ── Bottom buttons ─────────────────────────────────────────────
        bottom = Frame(outer)
        bottom.pack(fill="x")
        Button(
            bottom,
            text="Save & Close",
            command=self._save_and_close,
        ).pack(side="right", padx=(8, 0))
        Button(bottom, text="Cancel", command=self.destroy).pack(side="right")

        # Populate the rules listbox
        self._refresh_rules_list()
        self._toggle_scheduler_fields()

    # ------------------------------------------------------------------
    def _refresh_rules_list(self) -> None:
        self._rules_list.delete(0, END)
        for rule in self._rules:
            kw_text = ", ".join(rule.get("keywords", []))
            folder = rule.get("folder", "")
            self._rules_list.insert(END, f"{kw_text}  →  {folder}")

    def _browse_default_folder(self) -> None:
        folder = filedialog.askdirectory(
            parent=self,
            title="Default auto-download folder",
            initialdir=self._default_folder_var.get() or str(Path.home()),
        )
        if folder:
            self._default_folder_var.set(folder)

    def _toggle_scheduler_fields(self) -> None:
        state = "normal" if self._auto_enabled_var.get() else "disabled"
        self._interval_entry.configure(state=state)
        self._check_time_entry.configure(state=state)

    def _add_rule(self) -> None:
        dlg = KeywordRuleDialog(self)
        if dlg.result:
            self._rules.append(dlg.result)
            self._refresh_rules_list()
            self._rules_list.selection_clear(0, END)
            self._rules_list.selection_set(END)

    def _edit_rule(self) -> None:
        sel = self._rules_list.curselection()
        if not sel:
            messagebox.showwarning("Edit Rule", "Select a rule to edit.", parent=self)
            return
        idx = sel[0]
        dlg = KeywordRuleDialog(self, rule=self._rules[idx])
        if dlg.result:
            self._rules[idx] = dlg.result
            self._refresh_rules_list()
            self._rules_list.selection_set(idx)

    def _remove_rule(self) -> None:
        sel = self._rules_list.curselection()
        if not sel:
            return
        idx = sel[0]
        del self._rules[idx]
        self._refresh_rules_list()

    def _move_up(self) -> None:
        sel = self._rules_list.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self._rules[idx - 1], self._rules[idx] = self._rules[idx], self._rules[idx - 1]
        self._refresh_rules_list()
        self._rules_list.selection_set(idx - 1)

    def _move_down(self) -> None:
        sel = self._rules_list.curselection()
        if not sel or sel[0] >= len(self._rules) - 1:
            return
        idx = sel[0]
        self._rules[idx + 1], self._rules[idx] = self._rules[idx], self._rules[idx + 1]
        self._refresh_rules_list()
        self._rules_list.selection_set(idx + 1)

    def _validate_interval(self) -> int | None:
        try:
            val = int(self._interval_var.get().strip())
            if val < 1:
                raise ValueError
            return val
        except ValueError:
            messagebox.showerror(
                "Invalid Interval",
                "Check interval must be a whole number of hours (minimum 1).",
                parent=self,
            )
            return None

    def _validate_check_time(self) -> str | None:
        raw = self._check_time_var.get().strip()
        if not raw:
            return ""
        try:
            datetime.strptime(raw, "%H:%M")
            return raw
        except ValueError:
            messagebox.showerror(
                "Invalid Time",
                "Daily check time must be in HH:MM format, e.g. 03:00.",
                parent=self,
            )
            return None

    def _save_and_close(self) -> None:
        interval = self._validate_interval()
        if interval is None:
            return

        check_time = self._validate_check_time()
        if check_time is None:
            return

        self.app.settings["default_auto_folder"] = self._default_folder_var.get().strip()
        self.app.settings["auto_check_enabled"] = self._auto_enabled_var.get()
        self.app.settings["auto_check_interval_hours"] = interval
        self.app.settings["auto_check_time"] = check_time
        self.app.settings["keyword_rules"] = self._rules

        save_settings(self.app.settings)
        self.app._restart_scheduler()
        self.app._update_next_check_label()
        self.destroy()


# ---------------------------------------------------------------------------
# Main Application
# ---------------------------------------------------------------------------

class StrmMakerApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title(APP_TITLE)
        self.root.minsize(720, 720)

        loaded = load_settings()
        self.settings: dict[str, Any] = merge_settings(loaded)

        self.output_folder = StringVar(value=self.settings.get("output_folder", ""))
        self.real_debrid_token = StringVar(value=self.settings.get("real_debrid_token", ""))
        self.use_url_names = BooleanVar(value=True)
        self.status = StringVar(value="Load recent torrents or add magnets.")
        self.next_check_label_var = StringVar(value="")
        self.torrents: list[dict[str, Any]] = []

        # Scheduler state
        self._scheduler_stop = threading.Event()
        self._scheduler_thread: threading.Thread | None = None
        self._last_auto_scan: float = 0.0
        self._auto_scan_running = False

        # System tray state
        self._tray_icon: "pystray.Icon | None" = None

        self.build_ui()
        self._setup_tray()

        # Start scheduler if enabled
        if self.settings.get("auto_check_enabled"):
            self._restart_scheduler()
        self._update_next_check_label()

    # ------------------------------------------------------------------
    # UI
    # ------------------------------------------------------------------

    def build_ui(self) -> None:
        outer = Frame(self.root, padx=14, pady=14)
        outer.pack(fill="both", expand=True)

        # ── Output folder ──────────────────────────────────────────────
        folder_row = Frame(outer)
        folder_row.pack(fill="x", pady=(0, 10))

        Label(folder_row, text="Output Folder").pack(anchor="w")
        folder_picker = Frame(folder_row)
        folder_picker.pack(fill="x", pady=(4, 0))

        self.folder_entry = Entry(folder_picker, textvariable=self.output_folder)
        self.folder_entry.pack(side="left", fill="x", expand=True)
        Button(folder_picker, text="Browse...", command=self.choose_folder).pack(
            side="left", padx=(8, 0)
        )

        # ── Token ──────────────────────────────────────────────────────
        token_row = Frame(outer)
        token_row.pack(fill="x", pady=(0, 10))

        Label(token_row, text="Real-Debrid API Token").pack(anchor="w")
        token_picker = Frame(token_row)
        token_picker.pack(fill="x", pady=(4, 0))

        self.token_entry = Entry(
            token_picker, textvariable=self.real_debrid_token, show="*"
        )
        self.token_entry.pack(side="left", fill="x", expand=True)
        Button(token_picker, text="Save Token", command=self.save_token).pack(
            side="left", padx=(8, 0)
        )

        Checkbutton(
            outer,
            text="Use filenames from links when possible",
            variable=self.use_url_names,
        ).pack(anchor="w", pady=(0, 6))

        # ── Auto-download settings button + next-check indicator ───────
        auto_bar = Frame(outer)
        auto_bar.pack(fill="x", pady=(0, 12))
        Button(
            auto_bar,
            text="⚙  Auto-Download Settings",
            command=self.open_auto_settings,
        ).pack(side="left")
        Button(
            auto_bar,
            text="▶  Run Keyword Scan Now",
            command=self.run_auto_scan,
        ).pack(side="left", padx=(8, 0))
        Label(auto_bar, textvariable=self.next_check_label_var, fg="gray").pack(
            side="left", padx=(16, 0)
        )

        # ── Torrent list ───────────────────────────────────────────────
        Label(outer, text="Recent Real-Debrid Torrents").pack(anchor="w")
        self.torrent_list = Listbox(outer, height=9, selectmode=EXTENDED)
        self.torrent_list.pack(fill="both", expand=True, pady=(4, 8))

        torrent_actions = Frame(outer)
        torrent_actions.pack(fill="x", pady=(0, 12))

        Button(
            torrent_actions,
            text="Load Recent Torrents",
            command=self.load_recent_torrents,
        ).pack(side="left")
        Button(
            torrent_actions,
            text="Create From Selected Torrent(s)",
            command=self.create_from_selected_torrent,
        ).pack(side="left", padx=(8, 0))

        # ── Magnet links ───────────────────────────────────────────────
        Label(outer, text="Magnet Links").pack(anchor="w")
        self.magnets_text = ScrolledText(outer, height=5, wrap="none", undo=True)
        self.magnets_text.pack(fill="both", expand=True, pady=(4, 8))

        magnet_actions = Frame(outer)
        magnet_actions.pack(fill="x", pady=(0, 12))

        Button(
            magnet_actions,
            text="Add Magnets to Real-Debrid",
            command=self.add_magnets_to_real_debrid,
        ).pack(side="left")
        Button(
            magnet_actions,
            text="Clear Magnets",
            command=self.clear_magnets,
        ).pack(side="left", padx=(8, 0))

        Label(outer, textvariable=self.status, anchor="w").pack(fill="x", pady=(12, 0))

    # ------------------------------------------------------------------
    # Settings helpers
    # ------------------------------------------------------------------

    def _save_settings(self) -> None:
        self.settings["output_folder"] = self.output_folder.get().strip()
        self.settings["real_debrid_token"] = self.real_debrid_token.get().strip()
        save_settings(self.settings)

    def open_auto_settings(self) -> None:
        AutoDownloadSettingsWindow(self.root, self)

    # ------------------------------------------------------------------
    # Scheduler
    # ------------------------------------------------------------------

    # ------------------------------------------------------------------
    # System tray
    # ------------------------------------------------------------------

    def _make_tray_image(self) -> "Image.Image":
        """Draw a simple icon: dark background, blue circle, white play triangle."""
        size = 64
        img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        # Dark rounded background
        draw.ellipse([0, 0, size - 1, size - 1], fill=(28, 28, 36, 255))
        # Blue circle ring
        draw.ellipse([4, 4, size - 5, size - 5], outline=(80, 160, 255, 255), width=4)
        # White play triangle
        margin = 18
        draw.polygon(
            [(margin + 4, margin), (size - margin, size // 2), (margin + 4, size - margin)],
            fill=(255, 255, 255, 255),
        )
        return img

    def _setup_tray(self) -> None:
        """Wire up minimize-to-tray and close-to-tray behaviour."""
        # Always intercept the X button — quit via tray menu instead
        self.root.protocol("WM_DELETE_WINDOW", self._on_close_button)
        # Intercept minimize
        self.root.bind("<Unmap>", self._on_unmap)

    def _on_close_button(self) -> None:
        """Closing the window hides to tray (if available) or quits."""
        if TRAY_AVAILABLE:
            self._hide_to_tray()
        else:
            self._quit_app()

    def _on_unmap(self, event: Any) -> None:
        """Fires when the window is minimized — send it to the tray instead."""
        if event.widget is not self.root:
            return
        if TRAY_AVAILABLE and self.root.wm_state() == "iconic":
            # Small delay so Windows finishes the minimize animation first
            self.root.after(150, self._hide_to_tray)

    def _hide_to_tray(self) -> None:
        """Hide the main window and show the system-tray icon."""
        if self._tray_icon is not None:
            return  # already in tray

        self.root.withdraw()

        if not TRAY_AVAILABLE:
            return

        menu = pystray.Menu(
            pystray.MenuItem("Show STRM Maker", self._restore_from_tray, default=True),
            pystray.MenuItem("Run Keyword Scan", self._tray_run_scan),
            pystray.Menu.SEPARATOR,
            pystray.MenuItem("Quit", self._tray_quit),
        )
        self._tray_icon = pystray.Icon(
            "strm_maker",
            self._make_tray_image(),
            "Real-Debris STRM Maker",
            menu,
        )
        threading.Thread(
            target=self._tray_icon.run,
            daemon=True,
            name="strm-tray",
        ).start()

    def _restore_from_tray(self, *_: Any) -> None:
        """Called from the tray icon — restore the main window."""
        if self._tray_icon:
            self._tray_icon.stop()
            self._tray_icon = None
        self.root.after(0, self._do_restore)

    def _do_restore(self) -> None:
        self.root.deiconify()
        self.root.lift()
        self.root.focus_force()

    def _tray_run_scan(self, *_: Any) -> None:
        """Trigger a keyword scan from the tray menu."""
        self._restore_from_tray()
        self.root.after(300, self.run_auto_scan)

    def _tray_quit(self, *_: Any) -> None:
        """Quit the app entirely from the tray menu."""
        if self._tray_icon:
            self._tray_icon.stop()
            self._tray_icon = None
        self.root.after(0, self._quit_app)

    def _quit_app(self) -> None:
        self._scheduler_stop.set()
        self.root.destroy()

    def _restart_scheduler(self) -> None:
        """Stop any existing scheduler thread and start a fresh one."""
        self._scheduler_stop.set()
        if self._scheduler_thread and self._scheduler_thread.is_alive():
            self._scheduler_thread.join(timeout=2)

        self._scheduler_stop.clear()
        self._scheduler_thread = threading.Thread(
            target=self._scheduler_loop,
            daemon=True,
            name="strm-scheduler",
        )
        self._scheduler_thread.start()

    def _scheduler_loop(self) -> None:
        """Background thread: fires auto scan when it is time."""
        while not self._scheduler_stop.wait(SCHEDULER_POLL_SECONDS):
            if not self.settings.get("auto_check_enabled"):
                continue

            if self._should_run_scan():
                self.root.after(0, self._trigger_auto_scan)

    def _should_run_scan(self) -> bool:
        check_time_str = self.settings.get("auto_check_time", "").strip()
        interval_hours = int(self.settings.get("auto_check_interval_hours", 12))

        now = datetime.now()

        if check_time_str:
            try:
                target = datetime.strptime(check_time_str, "%H:%M").replace(
                    year=now.year, month=now.month, day=now.day
                )
            except ValueError:
                return False

            # Within a 1-minute window of the target time, and haven't run today
            delta = abs((now - target).total_seconds())
            if delta > 60:
                return False

            # Avoid double-firing within the same minute
            if self._last_auto_scan and (time.time() - self._last_auto_scan) < 3600:
                return False

            return True

        # Interval-based: run every N hours since last scan
        elapsed_hours = (time.time() - self._last_auto_scan) / 3600
        return elapsed_hours >= interval_hours

    def _trigger_auto_scan(self) -> None:
        """Called on the main thread from the scheduler."""
        if self._auto_scan_running:
            return
        self.run_auto_scan()

    def _update_next_check_label(self) -> None:
        if not self.settings.get("auto_check_enabled"):
            self.next_check_label_var.set("Auto-scan: OFF")
            return

        check_time = self.settings.get("auto_check_time", "").strip()
        interval = int(self.settings.get("auto_check_interval_hours", 12))

        if check_time:
            self.next_check_label_var.set(f"Auto-scan: daily at {check_time}")
        else:
            if self._last_auto_scan:
                next_dt = datetime.fromtimestamp(
                    self._last_auto_scan + interval * 3600
                )
                self.next_check_label_var.set(
                    f"Next scan: {next_dt.strftime('%b %d %H:%M')}"
                )
            else:
                self.next_check_label_var.set(f"Auto-scan: every {interval}h")

    # ------------------------------------------------------------------
    # Keyword auto-scan logic
    # ------------------------------------------------------------------

    def run_auto_scan(self) -> None:
        """Scan recent torrents and auto-download STRM files for keyword matches."""
        token = self.token_or_warn()
        if not token:
            return

        keyword_rules: list[dict[str, Any]] = self.settings.get("keyword_rules", [])
        default_auto_folder = self.settings.get("default_auto_folder", "").strip()

        if not keyword_rules and not default_auto_folder:
            messagebox.showinfo(
                APP_TITLE,
                "No keyword rules or default folder configured.\n"
                "Open Auto-Download Settings to set them up.",
            )
            return

        self._auto_scan_running = True
        self.status.set("Auto-scan: fetching latest torrents…")
        self.root.update_idletasks()

        try:
            torrents = get_real_debrid_torrents(token, limit=10)
        except (OSError, RuntimeError) as exc:
            messagebox.showerror(APP_TITLE, f"Auto-scan failed to load torrents:\n{exc}")
            self.status.set("Auto-scan failed.")
            self._auto_scan_running = False
            return

        processed_ids: set[str] = set(self.settings.get("processed_torrent_ids", []))
        newly_processed: list[str] = []
        total_written = 0
        scan_log: list[str] = []

        for torrent in torrents:
            torrent_id = str(torrent.get("id") or "")
            if not torrent_id or torrent_id in processed_ids:
                continue

            torrent_status = str(torrent.get("status") or "").lower()
            if torrent_status not in {"downloaded", "seeding"}:
                continue  # skip torrents that aren't ready

            torrent_name = str(
                torrent.get("filename") or torrent.get("original_filename") or "Unnamed"
            )

            # Match against keyword rules
            matched_rule = match_keyword_rule(torrent_name, keyword_rules)
            if matched_rule:
                dest_folder_str = matched_rule.get("folder", "").strip()
                name_norm = _normalize(torrent_name)
                matched_kw = next(
                    (k for k in matched_rule.get("keywords", []) if _normalize(k) in name_norm),
                    "?",
                )
            elif default_auto_folder:
                dest_folder_str = default_auto_folder
                matched_kw = "(default)"
            else:
                newly_processed.append(torrent_id)  # mark as seen but skip
                continue

            dest_folder = Path(dest_folder_str)
            try:
                dest_folder.mkdir(parents=True, exist_ok=True)
            except OSError as exc:
                scan_log.append(f"SKIP {torrent_name}: could not create folder — {exc}")
                continue

            self.status.set(f"Auto-scan: processing «{torrent_name}»…")
            self.root.update_idletasks()

            torrent_links, error = get_torrent_links_with_retries(
                token, torrent_id, self.wait_with_updates
            )
            if not torrent_links:
                scan_log.append(f"SKIP {torrent_name}: {error}")
                continue

            written, failed, skipped = self.write_links_to_strm_files(
                torrent_links, dest_folder, token
            )
            total_written += len(written)
            newly_processed.append(torrent_id)

            if written:
                scan_log.append(
                    f"✓ {torrent_name}\n"
                    f"  keyword: {matched_kw}  →  {dest_folder_str}\n"
                    f"  {len(written)} .strm file(s) created"
                )
            if failed:
                for fname, ferr in failed:
                    scan_log.append(f"  ✗ write failed: {fname} — {ferr}")
            if skipped:
                for item in skipped:
                    scan_log.append(f"  ⚠ skipped: {item}")

        # Update processed IDs and save
        processed_ids.update(newly_processed)
        self.settings["processed_torrent_ids"] = list(processed_ids)
        self._last_auto_scan = time.time()
        self._save_settings()
        self._update_next_check_label()

        # Refresh torrent list in main window
        self.torrents = torrents
        self.torrent_list.delete(0, END)
        for t in torrents:
            self.torrent_list.insert(END, torrent_label(t))

        self._auto_scan_running = False
        summary = f"Auto-scan complete: {total_written} new .strm file(s) created."
        self.status.set(summary)

        if scan_log:
            self.show_report_window(
                "Auto-Download Scan Report",
                summary + "\n\n" + "\n\n".join(scan_log),
            )
        else:
            self.status.set(
                f"Auto-scan complete — no new matching torrents found "
                f"(checked {len(torrents)} torrent(s))."
            )

    # ------------------------------------------------------------------
    # Existing actions (updated to use _save_settings)
    # ------------------------------------------------------------------

    def choose_folder(self) -> None:
        folder = filedialog.askdirectory(
            title="Choose where to save .strm files",
            initialdir=self.output_folder.get() or str(Path.home()),
        )
        if folder:
            self.output_folder.set(folder)
            self._save_settings()

    def clear_magnets(self) -> None:
        self.magnets_text.delete("1.0", END)
        self.status.set("Paste one magnet link per line.")

    def save_token(self) -> None:
        token = self.real_debrid_token.get().strip()
        if not token:
            messagebox.showwarning(APP_TITLE, "Paste your Real-Debrid API token first.")
            return

        self._save_settings()
        self.status.set("Real-Debrid API token saved locally.")
        messagebox.showinfo(APP_TITLE, "Real-Debrid API token saved locally.")

    def token_or_warn(self) -> str:
        token = self.real_debrid_token.get().strip()
        if not token:
            messagebox.showwarning(APP_TITLE, "Paste your Real-Debrid API token first.")
            return ""
        return token

    def output_folder_or_warn(self) -> Path | None:
        folder_text = self.output_folder.get().strip()
        if not folder_text:
            messagebox.showwarning(APP_TITLE, "Choose an output folder first.")
            return None

        output_folder = Path(folder_text)
        try:
            output_folder.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not create output folder:\n{exc}")
            return None

        return output_folder

    def wait_with_updates(self, seconds: float) -> None:
        deadline = time.monotonic() + seconds
        while time.monotonic() < deadline:
            self.root.update()
            time.sleep(min(0.1, max(0, deadline - time.monotonic())))

    def load_recent_torrents(self) -> None:
        token = self.token_or_warn()
        if not token:
            return

        self.status.set("Loading recent Real-Debrid torrents…")
        self.root.update_idletasks()

        try:
            self.torrents = get_real_debrid_torrents(token)
        except (OSError, RuntimeError) as exc:
            messagebox.showerror(APP_TITLE, f"Could not load torrents:\n{exc}")
            self.status.set("Could not load torrents.")
            return

        self.torrent_list.delete(0, END)
        for torrent in self.torrents:
            self.torrent_list.insert(END, torrent_label(torrent))

        self._save_settings()
        self.status.set(f"Loaded {len(self.torrents)} recent torrent(s).")

    def add_magnets_to_real_debrid(self) -> None:
        token = self.token_or_warn()
        if not token:
            return

        magnets = extract_magnets(self.magnets_text.get("1.0", END))
        if not magnets:
            messagebox.showwarning(APP_TITLE, "Paste at least one magnet link first.")
            return

        added = 0
        failed: list[tuple[str, str]] = []
        for index, magnet in enumerate(magnets, start=1):
            self.status.set(
                f"Adding magnet {index} of {len(magnets)} "
                f"with up to {LINK_RETRY_COUNT} attempts..."
            )
            self.root.update_idletasks()

            torrent_id, error = add_magnet_with_retries(
                token, magnet, self.wait_with_updates
            )
            if torrent_id:
                added += 1
            else:
                failed.append((magnet, f"{error} after {LINK_RETRY_COUNT} attempts"))

        self._save_settings()

        try:
            self.torrents = get_real_debrid_torrents(token)
            self.torrent_list.delete(0, END)
            for torrent in self.torrents:
                self.torrent_list.insert(END, torrent_label(torrent))
        except (OSError, RuntimeError):
            pass

        self.status.set(f"Added {added} magnet(s); {len(failed)} failed.")
        if failed:
            failed_items = "\n\n".join(
                f"{magnet}\n{error}" for magnet, error in failed
            )
            self.show_report_window(
                "STRM Maker Magnet Report",
                f"Added {added} magnet(s), but {len(failed)} failed:\n\n{failed_items}",
            )
        else:
            messagebox.showinfo(APP_TITLE, f"Added {added} magnet(s) to Real-Debrid.")

    def create_from_selected_torrent(self) -> None:
        token = self.token_or_warn()
        if not token:
            return

        output_folder = self.output_folder_or_warn()
        if output_folder is None:
            return

        selection = self.torrent_list.curselection()
        if not selection:
            messagebox.showwarning(
                APP_TITLE, "Select one or more torrents from the list first."
            )
            return

        links: list[str] = []
        skipped: list[str] = []
        for position, selected_index in enumerate(selection, start=1):
            torrent = self.torrents[selected_index]
            torrent_name = str(
                torrent.get("filename") or torrent.get("original_filename") or "Unnamed"
            )
            torrent_id = str(torrent.get("id") or "")
            if not torrent_id:
                skipped.append(f"{torrent_name}: missing torrent ID")
                continue

            self.status.set(
                f"Loading torrent {position} of {len(selection)} "
                f"with up to {LINK_RETRY_COUNT} attempts..."
            )
            self.root.update_idletasks()

            torrent_links, error = get_torrent_links_with_retries(
                token, torrent_id, self.wait_with_updates
            )
            if torrent_links:
                links.extend(torrent_links)
            else:
                skipped.append(
                    f"{torrent_name}: {error} after {LINK_RETRY_COUNT} attempts"
                )

        if not links:
            skipped_items = "\n".join(skipped)
            messagebox.showwarning(
                APP_TITLE,
                "No downloadable links were found for the selected torrent(s)."
                f"\n\nSkipped before writing {len(skipped)} item(s):\n\n{skipped_items}",
            )
            self.status.set("Selected torrent(s) have no downloadable links.")
            return

        written, failed, resolve_skipped = self.write_links_to_strm_files(
            links, output_folder, token
        )
        self.finish_create_result(
            written, failed, [*skipped, *resolve_skipped], output_folder, token
        )

    def write_links_to_strm_files(
        self,
        links: list[str],
        output_folder: Path,
        real_debrid_token: str,
    ) -> tuple[list[Path], list[tuple[str, str]], list[str]]:
        written: list[Path] = []
        failed: list[tuple[str, str]] = []
        skipped: list[str] = []

        try:
            for index, link in enumerate(links, start=1):
                self.status.set(
                    f"Resolving {index} of {len(links)} "
                    f"with up to {LINK_RETRY_COUNT} attempts..."
                )
                self.root.update_idletasks()

                resolved_links, error = resolve_links_with_retries(
                    link, real_debrid_token, self.wait_with_updates
                )
                if not resolved_links:
                    skipped.append(f"{link}: {error} after {LINK_RETRY_COUNT} attempts")
                    continue

                for resolved_index, resolved in enumerate(resolved_links, start=1):
                    resolved_link = resolved["link"]
                    filename_hint = clean_filename(resolved.get("filename", ""))
                    filename_source = resolved_link

                    if filename_hint:
                        filename_source = f"https://real-debrid.local/{filename_hint}"

                    filename = filename_from_url(
                        filename_source,
                        index if resolved_index == 1 else len(written) + 1,
                        self.use_url_names.get(),
                    )
                    path = unique_path(output_folder, filename)
                    error = write_text_with_retries(path, f"{resolved_link}\n")
                    if error:
                        failed.append((path.name, error))
                    else:
                        written.append(path)
        except OSError as exc:
            messagebox.showerror(APP_TITLE, f"Could not write files:\n{exc}")
            return written, failed, skipped

        return written, failed, skipped

    def show_report_window(self, title: str, report: str) -> None:
        report_window = Toplevel(self.root)
        report_window.title(title)
        report_window.minsize(640, 420)

        text = ScrolledText(report_window, wrap="word", padx=10, pady=10)
        text.pack(fill="both", expand=True, padx=12, pady=(12, 8))
        text.insert("1.0", report)
        text.configure(state="disabled")

        Button(report_window, text="Close", command=report_window.destroy).pack(
            pady=(0, 12)
        )

    def finish_create_result(
        self,
        written: list[Path],
        failed: list[tuple[str, str]],
        skipped: list[str],
        output_folder: Path,
        real_debrid_token: str,
    ) -> None:
        self._save_settings()
        self.status.set(
            f"Created {len(written)} .strm file(s); {len(failed)} write failure(s)."
        )

        if failed or skipped:
            sections = []
            if failed:
                failed_files = "\n".join(filename for filename, _ in failed)
                sections.append(
                    f"Failed to write {len(failed)} file(s) after "
                    f"{WRITE_RETRY_COUNT} attempts:\n\n{failed_files}"
                )
            if skipped:
                skipped_items = "\n".join(skipped)
                sections.append(
                    f"Skipped before writing {len(skipped)} item(s) after "
                    f"{LINK_RETRY_COUNT} link attempt(s):\n\n{skipped_items}"
                )

            self.show_report_window(
                "STRM Maker Report",
                f"Created {len(written)} .strm file(s).\n"
                f"Failed writes: {len(failed)}\n"
                f"Skipped before writing: {len(skipped)}\n\n"
                + "\n\n".join(sections),
            )
        else:
            messagebox.showinfo(
                APP_TITLE,
                f"Created {len(written)} .strm file(s).",
            )


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main() -> None:
    root = Tk()
    StrmMakerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
