from __future__ import annotations

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import traceback
from dataclasses import dataclass, asdict
from datetime import datetime
from pathlib import Path
from typing import Optional, Dict, Any, List, Tuple


from PyQt6.QtCore import Qt, QThread, pyqtSignal, QSettings
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import (
    QApplication, QAbstractItemView, QFileDialog, QFormLayout,
    QHBoxLayout, QHeaderView, QLabel, QLineEdit, QMainWindow, QMenu,
    QMessageBox, QProgressBar, QPushButton, QSplitter, QStatusBar,
    QTableWidget, QTableWidgetItem, QTextEdit, QTreeWidget, QTreeWidgetItem,
    QVBoxLayout, QWidget
)


from pypdf import PdfReader, PdfWriter


APP_ORG = "dpeana"
APP_NAME = "BennetPDFManager"
CACHE_FILENAME = ".bennet_pdf_cache.json"
CACHE_SCHEMA_VERSION = 4
CONTENT_PAGE_LIMIT = 5
HEADER_STATE_KEY = "table/header_state"
WINDOW_GEOMETRY_KEY = "window/geometry"
WINDOW_STATE_KEY = "window/state"
SORT_SECTION_KEY = "table/sort_section"
SORT_ORDER_KEY = "table/sort_order"
YEAR_RE = re.compile(r"(?:19|20)\d{2}")
CURRENT_YEAR = datetime.now().year


@dataclass
class PDFRecord:
    path: str
    filename: str
    subdir: str
    title: str = ""
    author: str = ""
    subject: str = ""
    year: str = ""
    content: str = ""
    mtime: float = 0.0
    size: int = 0
    read_error: str = ""

    @property
    def display_title(self) -> str:
        return self.title or Path(self.path).stem


def is_within_directory(path: Path, root: Path) -> bool:
    try:
        path.resolve().relative_to(root.resolve())
        return True
    except ValueError:
        return False


def open_with_default_app(path: str) -> None:
    if sys.platform.startswith("win"):
        os.startfile(path)  # type: ignore[attr-defined]
    elif sys.platform == "darwin":
        subprocess.Popen(["open", path])
    else:
        subprocess.Popen(["xdg-open", path])


def reveal_in_file_manager(path: str) -> None:
    if sys.platform.startswith("win"):
        subprocess.Popen(["explorer", "/select,", os.path.normpath(path)])
    elif sys.platform == "darwin":
        subprocess.Popen(["open", "-R", path])
    else:
        subprocess.Popen(["xdg-open", str(Path(path).parent)])


def safe_str(x) -> str:
    return "" if x is None else str(x).strip()


def normalize_subdir(pdf_path: Path, home_dir: Path) -> str:
    rel = pdf_path.parent.relative_to(home_dir)
    return "(root)" if str(rel) == "." else str(rel)


def normalize_whitespace(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def parse_year_from_date_string(text: str) -> str:
    if not text:
        return ""
    m = YEAR_RE.search(str(text))
    return m.group(0) if m else ""


def meta_get(meta, attr: str, key: str) -> str:
    """Safely read a value from a pypdf DocumentInformation object or a dict."""
    value = None
    try:
        value = getattr(meta, attr, None)
    except Exception:
        value = None
    if value:
        return safe_str(value)
    try:
        if meta is not None and hasattr(meta, "get"):
            value = meta.get(key, None)
    except Exception:
        value = None
    return safe_str(value)


# ---------------------------------------------------------------------------
#  YEAR DETECTION
# ---------------------------------------------------------------------------

def _plausible_year(y: int) -> bool:
    return 1900 <= y <= CURRENT_YEAR


def extract_year_from_text(text: str) -> str:
    """Find the most likely PUBLICATION year in page text."""
    if not text:
        return ""

    # 1. Explicit publication phrases.
    priority_patterns = [
        r"\bpublished\s+(?:online\s+)?(?:on\s+)?(?:\w+\.?\s+\d{1,2},?\s+)?((?:19|20)\d{2})\b",
        r"\baccepted\s+(?:\w+\.?\s+\d{1,2},?\s+)?((?:19|20)\d{2})\b",
        r"\breceived\s+(?:\w+\.?\s+\d{1,2},?\s+)?((?:19|20)\d{2})\b",
        r"\b©\s*((?:19|20)\d{2})\b",
        r"\bcopyright\s*©?\s*((?:19|20)\d{2})\b",
    ]
    for pattern in priority_patterns:
        for m in re.finditer(pattern, text, flags=re.IGNORECASE):
            y = int(m.group(1))
            if _plausible_year(y):
                return str(y)

    # 2. "Month YYYY" or "DD Month YYYY" common in headers/footers
    months = (r"jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
              r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
              r"nov(?:ember)?|dec(?:ember)?")
    month_year_patterns = [
        re.compile(rf"\b\d{{1,2}}\s+(?:{months})\.?\s+((?:19|20)\d{{2}})\b", re.IGNORECASE),
        re.compile(rf"\b(?:{months})\.?\s+\d{{1,2}},?\s+((?:19|20)\d{{2}})\b", re.IGNORECASE),
        re.compile(rf"\b(?:{months})\.?\s+((?:19|20)\d{{2}})\b", re.IGNORECASE),
    ]
    for pat in month_year_patterns:
        candidates = [int(m.group(1)) for m in pat.finditer(text)]
        candidates = [y for y in candidates if _plausible_year(y)]
        if candidates:
            return str(min(candidates, key=lambda y: (-candidates.count(y), y)))

    # 3. Most frequent plausible year before references
    lines = text.splitlines()
    search_text = ""
    for line in lines:
        lower = line.strip().lower()
        if lower.startswith(("references", "bibliography", "works cited")):
            break
        search_text += line + "\n"

    all_years = [int(m.group(0)) for m in YEAR_RE.finditer(search_text)]
    all_years = [y for y in all_years if _plausible_year(y)]
    if all_years:
        return str(max(set(all_years), key=lambda y: (all_years.count(y), y)))

    return ""


def extract_year(meta, pdf_path: Path, first_page_text: str = "") -> str:
    """Decide the best year. Page text is trusted before file mtime."""
    # 1. Custom /Year tag written by this app on a previous save.
    if meta and hasattr(meta, "get"):
        try:
            y = parse_year_from_date_string(str(meta.get("/Year", "")))
            if y and _plausible_year(int(y)):
                return y
        except Exception:
            pass

    # 2. Publication year found in the document text.
    text_year = extract_year_from_text(first_page_text)
    if text_year:
        return text_year

    # 3. PDF embedded creation/modification date.
    candidates: List[str] = []
    for attr in ("creation_date", "modification_date"):
        try:
            value = getattr(meta, attr, None)
            if value is not None:
                if hasattr(value, "year") and _plausible_year(value.year):
                    candidates.append(str(value.year))
                else:
                    candidates.append(str(value))
        except Exception:
            pass
    for key in ("/CreationDate", "/ModDate"):
        try:
            if meta is not None and hasattr(meta, "get"):
                raw = meta.get(key, "")
                if raw:
                    candidates.append(str(raw))
        except Exception:
            pass
    for candidate in candidates:
        y = parse_year_from_date_string(candidate)
        if y and _plausible_year(int(y)):
            return y

    # 4. Last resort: file modification time.
    try:
        y = datetime.fromtimestamp(pdf_path.stat().st_mtime).year
        if _plausible_year(y):
            return str(y)
    except Exception:
        pass
    return ""


# ---------------------------------------------------------------------------
#  AUTHOR DETECTION
# ---------------------------------------------------------------------------

AUTHOR_BANNED = [
    "abstract", "introduction", "keywords", "journal", "conference", "proceedings",
    "doi:", "http", "www.", "received", "accepted", "published",
    "issn", "isbn", "arxiv", "vol.", "volume", "issue", "email", "@",
    "corresponding author", "supplementary material",
]


def is_all_caps_heading(text: str) -> bool:
    """Detect short all-caps section headers like 'COLD ATOMS'."""
    text = normalize_whitespace(text)
    if not text:
        return False
    letters = re.sub(r"[^A-Za-z]", "", text)
    if not letters:
        return False
    has_marker = bool(re.search(r"[\d\*\u2020\u2021†‡,]", text))
    return letters.isupper() and not has_marker and len(text.split()) <= 8


def clean_author_text(text: str) -> str:
    """Normalise an authors string: strip numbers/markers, unify separators."""
    text = normalize_whitespace(text)
    # Remove affiliation numbers and footnote symbols throughout
    text = re.sub(r"[\d\*\u2020\u2021\u00a7\u00b6‡†§¶]+", "", text)
    # Convert ' and ' / ' & ' / ';' into commas
    text = re.sub(r"\s+and\s+", ", ", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*&\s*", ", ", text)
    text = re.sub(r"\s*;\s*", ", ", text)
    # Tidy comma spacing and collapse duplicate commas
    text = re.sub(r"\s*,\s*", ", ", text)
    text = re.sub(r"(,\s*){2,}", ", ", text)
    text = re.sub(r"\s+", " ", text)
    return text.strip(" ,;.")


def _is_name_token(raw: str) -> bool:
    """True if a single token looks like part of a person's name."""
    clean = re.sub(r"[^A-Za-z\u00C0-\u024F\-\.']", "", raw)
    if not clean:
        return False
    low = clean.lower()
    if low in {"and", "of", "for", "the", "in", "on", "with", "by",
               "de", "da", "del", "van", "von", "der", "di", "la", "le", "du"}:
        return True
    if len(clean.replace(".", "")) <= 1:
        return True
    return clean[0].isupper()


def looks_like_author_line(line: str) -> bool:
    """Heuristic gate: does this line look like a list of author names?"""
    line = normalize_whitespace(line)
    if not line:
        return False

    if is_all_caps_heading(line):
        return False
    if is_affiliation_line(line):
        return False
    if is_body_text_line(line):
        return False

    lower = line.lower()
    hard_ban = any(token in lower for token in AUTHOR_BANNED)

    if len(line) > 350:
        return False
    if YEAR_RE.search(line):
        return False

    # Strip affiliation markers/numbers into spaces, then split into candidate names.
    clean_line = re.sub(r"[\d\*\u2020\u2021\u00a7\u00b6‡†§¶]+", " ", line)
    clean_line = normalize_whitespace(clean_line)
    parts = [p.strip() for p in re.split(r",|\band\b|;|&", clean_line, flags=re.IGNORECASE) if p.strip()]
    if not parts:
        return False

    valid_name_count = 0
    for part in parts:
        words = [w for w in re.split(r"\s+", part) if w and len(w) > 0]
        if not (1 <= len(words) <= 6):
            continue
        name_words = [w for w in words if _is_name_token(w)]
        # Tolerate 1 junk token per name part (key fix for APS-style lines)
        if words and len(name_words) >= max(1, len(words) - 1):
            valid_name_count += 1

    if valid_name_count == 0:
        return False
    if hard_ban and valid_name_count <= 1:
        return False
    return True


def is_affiliation_line(line: str) -> bool:
    lower = normalize_whitespace(line).lower()
    tokens = [
        "department of", "dept. of", "dept of", "university", "institute",
        "laboratory", "laboratories", "school of", "faculty", "college",
        "center for", "centre for", "division of", "academy", "hospital",
        "purdue university", "usa", "cambridge", "massachusetts",
    ]
    if any(t in lower for t in tokens):
        return True
    # Many affiliations end with a ZIP/postal code + country
    if re.search(r"\b\d{4,5}\b.*\b(usa|united states|uk|canada|germany|france|china|japan)\b", lower):
        return True
    return False


def is_body_text_line(line: str) -> bool:
    lower = normalize_whitespace(line).lower()
    words = lower.split()
    if not words:
        return False
    # Explicit paragraph starts
    if lower.startswith(("abstract", "introduction", "we ", "this ", "here ", "in this ", "the realization",
                         "the ", "to ", "as a ", "in a ")):
        return True
    # A long line ending in a sentence-terminal period (not an initial like "D.")
    if len(words) >= 12 and re.search(r"[a-z]{3,}\.$", lower):
        return True
    return False


def guess_author_from_first_page(first_page_text: str, title: str) -> str:
    if not first_page_text:
        return ""

    # Explicit "Author(s):" line wins.
    m = re.search(r"(?im)^\s*authors?\s*:\s*(.+)$", first_page_text)
    if m:
        cleaned = clean_author_text(m.group(1))
        if cleaned:
            return cleaned

    lines = [normalize_whitespace(l) for l in first_page_text.splitlines()]
    lines = [l for l in lines if l]
    if not lines:
        return ""

    title_norm = normalize_whitespace(title).lower()

    def is_stop_line(line: str) -> bool:
        """Lines that mark the end of the author block."""
        if is_affiliation_line(line):
            return True
        if is_body_text_line(line):
            return True
        lower = line.lower()
        if re.search(r"\b(received|accepted|published)\b", lower):
            return True
        if lower.startswith("(received") or lower.startswith("(dated"):
            return True
        return False

    def collect_from(start: int) -> List[str]:
        block: List[str] = []
        for line in lines[start:start + 15]:
            if is_stop_line(line):
                break
            if is_all_caps_heading(line):
                continue
            # Skip title continuation
            if title_norm and line.lower() in title_norm:
                continue
            if looks_like_author_line(line):
                block.append(line)
            elif block:
                break
        return block

    # Strategy 1: find title, authors are right after it.
    title_idx = None
    if title_norm:
        for i, line in enumerate(lines[:30]):
            ln = line.lower()
            if ln == title_norm or (len(title_norm) > 15 and (ln in title_norm or title_norm in ln)):
                title_idx = i
                break

    collected: List[str] = []
    if title_idx is not None:
        collected = collect_from(title_idx + 1)

    # Strategy 2: skip all-caps header, scan from top.
    if not collected:
        start = 1 if lines and is_all_caps_heading(lines[0]) else 0
        collected = collect_from(start)

    # Strategy 3: scan every line until first affiliation.
    if not collected:
        for i, line in enumerate(lines[:25]):
            if is_stop_line(line):
                break
            if looks_like_author_line(line):
                collected = collect_from(i)
                if collected:
                    break

    if collected:
        full = clean_author_text(" ".join(collected))
        parts = [p.strip() for p in full.split(",") if p.strip()]
        seen: List[str] = []
        for p in parts:
            if p not in seen:
                seen.append(p)
        return ", ".join(seen)

    return ""


def choose_best_author(meta_author: str, guessed_author: str) -> str:
    meta_author = clean_author_text(meta_author)
    guessed_author = clean_author_text(guessed_author)

    bad_meta_tokens = ["microsoft", "word", "acrobat", "scanner", "unknown",
                       "administrator", "user", "owner", "pdf", "latex"]

    meta_ok = (meta_author
               and any(ch.isalpha() for ch in meta_author)
               and not any(t in meta_author.lower() for t in bad_meta_tokens))

    if meta_ok:
        meta_names = meta_author.count(",") + 1
        guess_names = guessed_author.count(",") + 1 if guessed_author else 0
        if guess_names > meta_names + 1:
            return guessed_author
        return meta_author

    return guessed_author or meta_author


# ---------------------------------------------------------------------------
#  READ / WRITE
# ---------------------------------------------------------------------------

def read_pdf_record(pdf_path: Path, home_dir: Path) -> PDFRecord:
    stat = pdf_path.stat()
    rec = PDFRecord(
        path=str(pdf_path),
        filename=pdf_path.name,
        subdir=normalize_subdir(pdf_path, home_dir),
        mtime=stat.st_mtime,
        size=stat.st_size,
    )

    try:
        reader = PdfReader(str(pdf_path))

        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                rec.read_error = "Encrypted PDF"
                rec.title = pdf_path.stem
                rec.year = str(datetime.fromtimestamp(stat.st_mtime).year)
                return rec

        meta = reader.metadata

        meta_title = meta_get(meta, "title", "/Title")
        meta_author = meta_get(meta, "author", "/Author")
        rec.subject = meta_get(meta, "subject", "/Subject")

        rec.title = meta_title or pdf_path.stem

        parts: List[str] = []
        first_page_text = ""
        for idx, page in enumerate(reader.pages[: min(len(reader.pages), CONTENT_PAGE_LIMIT)]):
            try:
                t = page.extract_text() or ""
                if idx == 0:
                    first_page_text = t
                if t:
                    parts.append(t)
            except Exception:
                pass
        rec.content = "\n".join(parts)

        # Year detection
        rec.year = extract_year(meta, pdf_path, first_page_text or rec.content)

        # Author detection
        guessed_author = guess_author_from_first_page(first_page_text, rec.title)
        rec.author = choose_best_author(meta_author, guessed_author)

        if not rec.year:
            rec.year = str(datetime.fromtimestamp(stat.st_mtime).year)

    except Exception as e:
        rec.read_error = f"{type(e).__name__}: {e}"
        rec.title = pdf_path.stem
        try:
            rec.year = str(datetime.fromtimestamp(pdf_path.stat().st_mtime).year)
        except Exception:
            rec.year = ""

    return rec


def _year_to_pdf_date(year: str) -> Optional[str]:
    y = parse_year_from_date_string(year)
    if not y:
        return None
    return f"D:{y}0101000000"


def write_pdf_metadata(pdf_path: str, title: str, author: str, subject: str,
                       year: str = "") -> None:
    pdf = Path(pdf_path)
    tmp_path: Optional[Path] = None

    try:
        reader = PdfReader(str(pdf))
        if reader.is_encrypted:
            try:
                reader.decrypt("")
            except Exception:
                raise RuntimeError("Encrypted PDF cannot be modified.")

        writer = PdfWriter()
        for page in reader.pages:
            writer.add_page(page)

        existing: Dict[str, Any] = {}
        if reader.metadata:
            for k, v in dict(reader.metadata).items():
                if v is not None:
                    existing[str(k)] = str(v)

        existing.update({
            "/Title": title,
            "/Author": author,
            "/Subject": subject,
        })

        clean_year = parse_year_from_date_string(year)
        if clean_year:
            existing["/Year"] = clean_year
            pdf_date = _year_to_pdf_date(clean_year)
            if pdf_date:
                existing["/CreationDate"] = pdf_date

        writer.add_metadata(existing)

        with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False, dir=str(pdf.parent)) as tmp:
            tmp_path = Path(tmp.name)
            writer.write(tmp)

        shutil.move(str(tmp_path), str(pdf))

    except Exception:
        if tmp_path and tmp_path.exists():
            try:
                tmp_path.unlink()
            except Exception:
                pass
        raise


# ---------------------------------------------------------------------------
#  CACHE TRUST CHECK
# ---------------------------------------------------------------------------

def cached_record_is_trustworthy(cached: Dict[str, Any], stat) -> bool:
    if not cached:
        return False
    if cached.get("mtime") != stat.st_mtime or cached.get("size") != stat.st_size:
        return False
    author = safe_str(cached.get("author", ""))
    year = safe_str(cached.get("year", ""))
    if not author or not year:
        return False
    if is_all_caps_heading(author):
        return False
    if not YEAR_RE.fullmatch(year):
        return False
    return True


class ScanWorker(QThread):
    progress = pyqtSignal(int, int, str)
    finished_ok = pyqtSignal(list)
    failed = pyqtSignal(str)

    def __init__(self, home_dir: str, cache_records: List[Dict[str, Any]]):
        super().__init__()
        self.home_dir = Path(home_dir)
        self.cache_by_path = {r.get("path"): r for r in cache_records if isinstance(r, dict)}
        self._cancelled = False

    def cancel(self) -> None:
        self._cancelled = True

    def run(self) -> None:
        try:
            pdf_paths: List[Path] = []
            for root, dirs, files in os.walk(self.home_dir):
                if self._cancelled:
                    return
                dirs[:] = [d for d in dirs if d not in {"__pycache__", ".git", ".svn", ".hg"}]
                root_p = Path(root)
                for fn in files:
                    if fn.lower().endswith(".pdf"):
                        p = root_p / fn
                        if is_within_directory(p, self.home_dir):
                            pdf_paths.append(p)

            total = len(pdf_paths)
            out: List[Dict[str, Any]] = []

            for i, p in enumerate(pdf_paths, start=1):
                if self._cancelled:
                    return
                self.progress.emit(i, total, str(p))

                try:
                    stat = p.stat()
                except OSError:
                    continue

                cached = self.cache_by_path.get(str(p))
                if cached_record_is_trustworthy(cached, stat):
                    out.append(cached)
                    continue

                rec = read_pdf_record(p, self.home_dir)
                out.append(asdict(rec))

            self.finished_ok.emit(out)

        except Exception as e:
            self.failed.emit(f"{e}\n\n{traceback.format_exc()}")


class BennetPDFManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Bennet PDF Manager")
        self.resize(1450, 820)

        self.settings = QSettings(APP_ORG, APP_NAME)
        self.home_dir: str = self.settings.value("home_dir", "", type=str)

        self.records: List[PDFRecord] = []
        self.filtered: List[PDFRecord] = []
        self.current: Optional[PDFRecord] = None
        self.worker: Optional[ScanWorker] = None
        self.default_sort_section = 0
        self.default_sort_order = Qt.SortOrder.AscendingOrder

        self._build_ui()
        self.restore_ui_state()

        if self.home_dir and Path(self.home_dir).exists():
            self.dir_label.setText(f"Home: {self.home_dir}")
            self.load_cache_or_scan()
        else:
            self.status.showMessage("Choose a home directory to begin.")

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        outer = QVBoxLayout(central)

        top = QHBoxLayout()
        self.dir_label = QLabel("Home: (not set)")
        self.dir_label.setStyleSheet("color:#555;")

        self.search_box = QLineEdit()
        self.search_box.setPlaceholderText("Search (title, author, subject, contents)…")
        self.search_box.textChanged.connect(self.apply_filters)

        choose_btn = QPushButton("Choose Home…")
        choose_btn.clicked.connect(self.choose_home)

        rescan_btn = QPushButton("Rescan")
        rescan_btn.clicked.connect(self.rescan)

        clear_btn = QPushButton("Clear")
        clear_btn.clicked.connect(lambda: self.search_box.setText(""))

        top.addWidget(self.dir_label, 2)
        top.addWidget(QLabel("Search:"))
        top.addWidget(self.search_box, 3)
        top.addWidget(clear_btn)
        top.addWidget(choose_btn)
        top.addWidget(rescan_btn)
        outer.addLayout(top)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        self.folder_tree = QTreeWidget()
        self.folder_tree.setHeaderLabel("Folders")
        self.folder_tree.itemSelectionChanged.connect(self.apply_filters)
        splitter.addWidget(self.folder_tree)

        self.table = QTableWidget(0, 7)
        self.table.setHorizontalHeaderLabels(["Title", "Author", "Date", "Paper/Subject", "Folder", "File", "Match"])
        self.table.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.SelectionMode.SingleSelection)
        self.table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.table.setSortingEnabled(True)
        self.table.itemSelectionChanged.connect(self.on_selection)
        self.table.itemDoubleClicked.connect(self.open_selected)
        self.table.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.context_menu)

        hdr = self.table.horizontalHeader()
        hdr.setSectionsMovable(False)
        hdr.setStretchLastSection(False)
        hdr.setMinimumSectionSize(70)
        hdr.setDefaultSectionSize(150)
        for i in range(self.table.columnCount()):
            hdr.setSectionResizeMode(i, QHeaderView.ResizeMode.Interactive)

        self.table.setColumnWidth(0, 300)
        self.table.setColumnWidth(1, 260)
        self.table.setColumnWidth(2, 80)
        self.table.setColumnWidth(3, 240)
        self.table.setColumnWidth(4, 150)
        self.table.setColumnWidth(5, 180)
        self.table.setColumnWidth(6, 120)

        hdr.sortIndicatorChanged.connect(self.on_sort_changed)
        splitter.addWidget(self.table)

        right = QWidget()
        right_l = QVBoxLayout(right)

        right_l.addWidget(QLabel("<b>Metadata</b>"))

        form = QFormLayout()
        self.title_edit = QLineEdit()
        self.author_edit = QLineEdit()
        self.subject_edit = QLineEdit()
        self.year_edit = QLineEdit()
        self.year_edit.setPlaceholderText("e.g. 2016")
        self.year_edit.setMaxLength(4)
        form.addRow("Title:", self.title_edit)
        form.addRow("Author:", self.author_edit)
        form.addRow("Paper/Subject:", self.subject_edit)
        form.addRow("Date (year):", self.year_edit)
        right_l.addLayout(form)

        btn_row = QHBoxLayout()
        self.save_btn = QPushButton("Save")
        self.save_btn.clicked.connect(self.save_metadata)
        self.save_btn.setEnabled(False)

        self.open_btn = QPushButton("Open")
        self.open_btn.clicked.connect(self.open_selected)
        self.open_btn.setEnabled(False)

        btn_row.addWidget(self.save_btn)
        btn_row.addWidget(self.open_btn)
        right_l.addLayout(btn_row)

        right_l.addWidget(QLabel("<b>Path</b>"))
        self.path_label = QLabel("")
        self.path_label.setWordWrap(True)
        self.path_label.setStyleSheet("color:#666; font-size:11px;")
        right_l.addWidget(self.path_label)

        right_l.addWidget(QLabel(f"<b>Indexed text preview</b> (first {CONTENT_PAGE_LIMIT} page(s))"))
        self.preview = QTextEdit()
        self.preview.setReadOnly(True)
        right_l.addWidget(self.preview, 1)

        splitter.addWidget(right)

        splitter.setStretchFactor(0, 1)
        splitter.setStretchFactor(1, 5)
        splitter.setStretchFactor(2, 2)
        splitter.setSizes([250, 900, 340])

        outer.addWidget(splitter, 1)

        self.status = QStatusBar()
        self.setStatusBar(self.status)

        self.progress = QProgressBar()
        self.progress.setMaximumWidth(260)
        self.progress.setVisible(False)
        self.status.addPermanentWidget(self.progress)

        m = self.menuBar().addMenu("&File")
        a_choose = QAction("Choose Home…", self)
        a_choose.triggered.connect(self.choose_home)
        m.addAction(a_choose)

        a_rescan = QAction("Rescan", self)
        a_rescan.triggered.connect(self.rescan)
        m.addAction(a_rescan)

        m.addSeparator()
        a_exit = QAction("Exit", self)
        a_exit.triggered.connect(self.close)
        m.addAction(a_exit)

    def cache_path(self) -> Optional[Path]:
        if not self.home_dir:
            return None
        return Path(self.home_dir) / CACHE_FILENAME

    def restore_ui_state(self):
        geometry = self.settings.value(WINDOW_GEOMETRY_KEY)
        if geometry is not None:
            self.restoreGeometry(geometry)

        window_state = self.settings.value(WINDOW_STATE_KEY)
        if window_state is not None:
            self.restoreState(window_state)

        header_state = self.settings.value(HEADER_STATE_KEY)
        if header_state is not None:
            self.table.horizontalHeader().restoreState(header_state)

        sort_section = self.settings.value(SORT_SECTION_KEY, self.default_sort_section, type=int)
        sort_order_int = self.settings.value(SORT_ORDER_KEY, int(self.default_sort_order.value), type=int)
        sort_order = Qt.SortOrder(sort_order_int)
        self.table.sortItems(sort_section, sort_order)
        self.table.horizontalHeader().setSortIndicator(sort_section, sort_order)

    def save_ui_state(self):
        self.settings.setValue(WINDOW_GEOMETRY_KEY, self.saveGeometry())
        self.settings.setValue(WINDOW_STATE_KEY, self.saveState())
        self.settings.setValue(HEADER_STATE_KEY, self.table.horizontalHeader().saveState())

        hdr = self.table.horizontalHeader()
        self.settings.setValue(SORT_SECTION_KEY, hdr.sortIndicatorSection())
        self.settings.setValue(SORT_ORDER_KEY, int(hdr.sortIndicatorOrder().value))

    def read_cache_records(self) -> List[Dict[str, Any]]:
        cache = self.cache_path()
        if not cache or not cache.exists():
            return []
        try:
            data = json.loads(cache.read_text(encoding="utf-8"))
        except Exception:
            return []
        if isinstance(data, dict):
            if data.get("schema_version") != CACHE_SCHEMA_VERSION:
                return []
            records = data.get("records", [])
        else:
            return []
        if not isinstance(records, list):
            return []
        return [r for r in records if isinstance(r, dict)]

    def load_cache_or_scan(self):
        cache = self.cache_path()
        if cache and cache.exists():
            try:
                records = self.read_cache_records()
                if records:
                    self.records = [PDFRecord(**r) for r in records]
                    self.populate_folders()
                    self.apply_filters()
                    self.status.showMessage(
                        f"Loaded {len(self.records)} PDFs from cache. Rescan to refresh.", 7000)
                    return
            except Exception as e:
                self.status.showMessage(f"Cache load failed: {e}", 7000)
        self.rescan()

    def save_cache(self, records_as_dicts: List[Dict[str, Any]]):
        cache = self.cache_path()
        if not cache:
            return
        payload = {
            "schema_version": CACHE_SCHEMA_VERSION,
            "home_dir": self.home_dir,
            "records": records_as_dicts,
        }
        cache.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")

    def choose_home(self):
        start = self.home_dir or str(Path.home())
        folder = QFileDialog.getExistingDirectory(self, "Choose PDF Home Directory", start)
        if not folder:
            return
        self.home_dir = folder
        self.settings.setValue("home_dir", folder)
        self.dir_label.setText(f"Home: {folder}")
        self.rescan()

    def rescan(self):
        if not self.home_dir:
            self.choose_home()
            return
        if self.worker and self.worker.isRunning():
            return

        cache_records = self.read_cache_records()

        self.progress.setVisible(True)
        self.progress.setValue(0)
        self.status.showMessage("Indexing…")

        self.folder_tree.clear()
        self.table.setRowCount(0)
        self.clear_right_panel()

        self.worker = ScanWorker(self.home_dir, cache_records)
        self.worker.progress.connect(self.on_scan_progress)
        self.worker.finished_ok.connect(self.on_scan_done)
        self.worker.failed.connect(self.on_scan_failed)
        self.worker.start()

    def on_scan_progress(self, i: int, total: int, path: str):
        self.progress.setMaximum(max(total, 1))
        self.progress.setValue(i)
        self.status.showMessage(f"[{i}/{total}] {path}")

    def on_scan_done(self, records_as_dicts: List[Dict[str, Any]]):
        self.progress.setVisible(False)
        self.save_cache(records_as_dicts)
        self.records = [PDFRecord(**r) for r in records_as_dicts]
        self.populate_folders()
        self.apply_filters()
        self.status.showMessage(f"Indexed {len(self.records)} PDFs.", 6000)

    def on_scan_failed(self, msg: str):
        self.progress.setVisible(False)
        QMessageBox.critical(self, "Indexing failed", msg)

    def populate_folders(self):
        self.folder_tree.clear()
        root = QTreeWidgetItem([f"All ({len(self.records)})"])
        root.setData(0, Qt.ItemDataRole.UserRole, None)
        self.folder_tree.addTopLevelItem(root)

        counts: Dict[str, int] = {}
        for r in self.records:
            counts[r.subdir] = counts.get(r.subdir, 0) + 1

        for sd in sorted(counts, key=lambda s: s.lower()):
            it = QTreeWidgetItem([f"{sd} ({counts[sd]})"])
            it.setData(0, Qt.ItemDataRole.UserRole, sd)
            root.addChild(it)

        root.setExpanded(True)
        self.folder_tree.setCurrentItem(root)

    def current_folder(self) -> Optional[str]:
        item = self.folder_tree.currentItem()
        if not item:
            return None
        return item.data(0, Qt.ItemDataRole.UserRole)

    def score(self, r: PDFRecord, q: str) -> Tuple[int, str]:
        q = q.lower()
        labels: List[str] = []
        score = 0

        if q in (r.title or "").lower():
            score = max(score, 300)
            labels.append("Title")
        if q in (r.author or "").lower():
            score = max(score, 250)
            labels.append("Author")
        if q in (r.year or "").lower():
            score = max(score, 225)
            labels.append("Date")
        if q in (r.subject or "").lower():
            score = max(score, 200)
            labels.append("Subject")
        if q in (r.filename or "").lower():
            score = max(score, 175)
            labels.append("File")
        if q in (r.content or "").lower():
            score = max(score, 100)
            labels.append("Contents")

        if labels:
            score += len(labels)

        return score, ", ".join(labels)

    def apply_filters(self):
        folder = self.current_folder()
        q = self.search_box.text().strip()

        selected_path = self.current.path if self.current else None
        header_state = self.table.horizontalHeader().saveState()
        sort_section = self.table.horizontalHeader().sortIndicatorSection()
        sort_order = self.table.horizontalHeader().sortIndicatorOrder()

        scored: List[Tuple[int, str, PDFRecord]] = []
        for r in self.records:
            if folder is not None and r.subdir != folder:
                continue
            if not q:
                scored.append((0, "", r))
            else:
                s, label = self.score(r, q)
                if s > 0:
                    scored.append((s, label, r))

        if q:
            scored.sort(key=lambda x: (-x[0], x[2].display_title.lower(), x[2].filename.lower()))
        else:
            scored.sort(key=lambda x: (x[2].subdir.lower(), x[2].display_title.lower()))

        self.filtered = [r for _, _, r in scored]
        labels = {r.path: lab for _, lab, r in scored}
        self.populate_table(labels)

        self.table.horizontalHeader().restoreState(header_state)
        self.table.sortItems(sort_section, sort_order)
        self.table.horizontalHeader().setSortIndicator(sort_section, sort_order)

        if selected_path:
            self.select_record_by_path(selected_path)

        if q:
            self.status.showMessage(f"{len(self.filtered)} result(s) for '{q}'.")
        else:
            self.status.showMessage(f"Showing {len(self.filtered)} PDF(s).")

    def populate_table(self, labels: Dict[str, str]):
        self.table.setSortingEnabled(False)
        self.table.setRowCount(0)

        for r in self.filtered:
            row = self.table.rowCount()
            self.table.insertRow(row)

            cols = [r.display_title, r.author, r.year, r.subject, r.subdir, r.filename, labels.get(r.path, "")]
            for c, val in enumerate(cols):
                item = QTableWidgetItem(val)
                item.setData(Qt.ItemDataRole.UserRole, r.path)
                self.table.setItem(row, c, item)

        self.table.setSortingEnabled(True)

    def select_record_by_path(self, path: str):
        for row in range(self.table.rowCount()):
            item = self.table.item(row, 0)
            if item and item.data(Qt.ItemDataRole.UserRole) == path:
                self.table.selectRow(row)
                return

    def selected_record(self) -> Optional[PDFRecord]:
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return None
        item = self.table.item(rows[0].row(), 0)
        if not item:
            return None
        path = item.data(Qt.ItemDataRole.UserRole)
        for r in self.records:
            if r.path == path:
                return r
        return None

    def on_selection(self):
        r = self.selected_record()
        self.current = r
        if not r:
            self.clear_right_panel()
            return

        self.title_edit.setText(r.title)
        self.author_edit.setText(r.author)
        self.subject_edit.setText(r.subject)
        self.year_edit.setText(r.year)
        self.path_label.setText(r.path)
        self.preview.setPlainText((r.content or "")[:12000])

        self.save_btn.setEnabled(True)
        self.open_btn.setEnabled(True)

    def on_sort_changed(self, section: int, order: Qt.SortOrder):
        self.settings.setValue(SORT_SECTION_KEY, section)
        self.settings.setValue(SORT_ORDER_KEY, int(order.value))
        self.settings.setValue(HEADER_STATE_KEY, self.table.horizontalHeader().saveState())

    def clear_right_panel(self):
        self.current = None
        self.title_edit.clear()
        self.author_edit.clear()
        self.subject_edit.clear()
        self.year_edit.clear()
        self.path_label.clear()
        self.preview.clear()
        self.save_btn.setEnabled(False)
        self.open_btn.setEnabled(False)

    def open_selected(self):
        r = self.current or self.selected_record()
        if not r:
            return
        p = Path(r.path)
        if not p.exists():
            QMessageBox.warning(self, "Missing file", f"File not found:\n\n{r.path}")
            return
        if self.home_dir and not is_within_directory(p, Path(self.home_dir)):
            QMessageBox.critical(self, "Blocked", "Refusing to open outside the home directory.")
            return
        try:
            open_with_default_app(r.path)
        except Exception as e:
            QMessageBox.critical(self, "Open failed", f"{e}")

    def save_metadata(self):
        r = self.current
        if not r:
            return

        p = Path(r.path)
        if not p.exists():
            QMessageBox.warning(self, "Missing file", f"File not found:\n\n{r.path}")
            return
        if self.home_dir and not is_within_directory(p, Path(self.home_dir)):
            QMessageBox.critical(self, "Blocked", "Refusing to modify outside the home directory.")
            return

        new_title = self.title_edit.text().strip()
        new_author = self.author_edit.text().strip()
        new_subject = self.subject_edit.text().strip()
        new_year = self.year_edit.text().strip()

        if new_year and not YEAR_RE.fullmatch(new_year):
            QMessageBox.warning(self, "Invalid year",
                                "Please enter a 4-digit year (e.g. 2016), or leave it blank.")
            return

        try:
            write_pdf_metadata(r.path, new_title, new_author, new_subject, new_year)

            r.title = new_title or p.stem
            r.author = new_author
            r.subject = new_subject
            if new_year:
                r.year = new_year
            st = p.stat()
            r.mtime = st.st_mtime
            r.size = st.st_size
            r.read_error = ""

            records = self.read_cache_records()
            merged = {rec.get("path"): rec for rec in records if isinstance(rec, dict)}
            merged[r.path] = asdict(r)
            self.save_cache(list(merged.values()))

            self.apply_filters()
            self.select_record_by_path(r.path)
            self.status.showMessage("Metadata saved.", 3000)

        except Exception as e:
            QMessageBox.critical(self, "Save failed", f"Could not write metadata:\n\n{e}")

    def context_menu(self, pos):
        if not self.selected_record():
            return
        menu = QMenu(self)

        a_open = QAction("Open", self)
        a_open.triggered.connect(self.open_selected)
        menu.addAction(a_open)

        a_reveal = QAction("Show in Folder", self)
        a_reveal.triggered.connect(lambda: self._reveal_selected())
        menu.addAction(a_reveal)

        menu.addSeparator()
        a_rescan = QAction("Rescan", self)
        a_rescan.triggered.connect(self.rescan)
        menu.addAction(a_rescan)

        menu.exec(self.table.viewport().mapToGlobal(pos))

    def _reveal_selected(self):
        r = self.current or self.selected_record()
        if not r:
            return
        try:
            reveal_in_file_manager(r.path)
        except Exception as e:
            QMessageBox.critical(self, "Reveal failed", f"{e}")

    def closeEvent(self, event):
        self.save_ui_state()
        if self.worker and self.worker.isRunning():
            self.worker.cancel()
            self.worker.wait(1500)
        event.accept()


def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    app.setOrganizationName(APP_ORG)
    app.setApplicationName(APP_NAME)

    w = BennetPDFManager()
    w.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()