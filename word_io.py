"""Read book requests out of Word documents (.docx / .doc).

Unlike the xlsx input (clean columns), Word orders are free-text citation lists.
Books appear one per paragraph, mixed with professor-name headers and short
status/price annotation lines that a librarian adds afterwards. We classify each
paragraph by content (never by colour) into: empty / header / status / book, and
turn each book paragraph into a `Row` (reusing the dataclass from xlsx_io).
"""

import re
import unicodedata
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from xlsx_io import Row

_W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"

# Bidi / formatting marks that litter the example files and carry no meaning.
_BIDI_MARKS = "".join(
    chr(c) for c in (0x200E, 0x200F, 0x202A, 0x202B, 0x202C, 0x202D, 0x202E, 0xFEFF)
)
_DROP = {ord(c): None for c in _BIDI_MARKS}


def _clean(text: str) -> str:
    # Normalise exotic spaces (no-break, figure, narrow) to ASCII space, drop bidi marks.
    text = "".join(" " if unicodedata.category(ch) == "Zs" else ch for ch in text)
    text = text.translate(_DROP)
    return re.sub(r"\s+", " ", text).strip()


# --- paragraph extraction -------------------------------------------------

def _paragraphs_docx(path: str | Path) -> list[str]:
    with zipfile.ZipFile(path) as z:
        xml = z.read("word/document.xml")
    body = ET.fromstring(xml).find(_W + "body")
    out: list[str] = []
    if body is None:
        return out
    for p in body.iter(_W + "p"):
        parts: list[str] = []
        for node in p.iter():
            if node.tag == _W + "t":
                parts.append(node.text or "")
            elif node.tag in (_W + "tab", _W + "br", _W + "cr"):
                parts.append(" ")
        out.append(_clean("".join(parts)))
    return out


def _paragraphs_doc(path: str | Path) -> list[str]:
    """Read legacy binary .doc via Word COM automation (needs MS Word + pywin32)."""
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError as e:
        raise RuntimeError(
            "קריאת קובצי .doc דורשת את החבילה pywin32 ואת Microsoft Word מותקן. "
            "אפשרות חלופית: לשמור את הקובץ כ-.docx."
        ) from e

    pythoncom.CoInitialize()
    word = None
    doc = None
    try:
        word = win32.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(
            str(Path(path).resolve()), ReadOnly=True, AddToRecentFiles=False
        )
        return [_clean(p.Range.Text) for p in doc.Paragraphs]
    except Exception as e:  # noqa: BLE001 - surface a friendly message to the GUI
        raise RuntimeError(
            f"לא ניתן לקרוא את קובץ ה-.doc. ודא ש-Microsoft Word מותקן. ({e})"
        ) from e
    finally:
        if doc is not None:
            doc.Close(False)
        if word is not None:
            word.Quit()
        pythoncom.CoUninitialize()


# --- line classification --------------------------------------------------

_NUM_PREFIX = re.compile(r"^\s*\d+\s*[.)]\s*")

_STOPWORDS = {
    "in", "of", "the", "and", "for", "a", "an", "to", "on", "with",
    "de", "la", "le", "du", "des", "el", "von", "der", "und",
}

# 13-digit (97x...) or 10-digit (...X) ISBN, possibly hyphen/space separated.
_ISBN_RE = re.compile(r"ISBN(?:[-\s]*1[03])?\s*:?\s*([\dXx][\dXx\- ]{8,20})", re.IGNORECASE)

_STATUS_WORDS = re.compile(
    r"\b(lib|online|print|ebc|ebsco|jstor|stock|electronic|cairn|proquest|copy"
    r"|hardcover|paperback|unlimit\w*)\b|t&f",
    re.IGNORECASE,
)
_PRICE = re.compile(r"\d+\.\d{2}\b")
_HEBREW = re.compile("[֐-׿]")


def _find_isbn(text: str) -> str:
    m = _ISBN_RE.search(text)
    if not m:
        return ""
    digits = re.sub(r"[^0-9Xx]", "", m.group(1)).upper()
    return digits if len(digits) in (10, 13) else ""


def _is_status(text: str) -> bool:
    words = text.split()
    if not words or len(words) > 7:
        return False
    if _HEBREW.search(text):
        return True
    if _PRICE.search(text):
        return True
    return bool(_STATUS_WORDS.search(text))


def _is_header(text: str) -> bool:
    if text.endswith(":"):
        return True  # section lead-in, e.g. "Books for library by order of preference:"
    if _find_isbn(text) or _is_status(text):
        return False
    words = text.split()
    if not words or len(words) > 3 or "," in text:
        return False
    if any(ch.isdigit() for ch in text):
        return False
    core = [w.strip(".").strip() for w in words]
    if any(w.lower() in _STOPWORDS for w in core):
        return False
    return all(w[:1].isupper() for w in core if w)


def _strip_leading_number(text: str) -> str:
    return _NUM_PREFIX.sub("", text).strip()


# --- public API -----------------------------------------------------------

def read_rows(path: str | Path) -> list[Row]:
    suffix = Path(path).suffix.lower()
    if suffix == ".docx":
        paras = _paragraphs_docx(path)
    elif suffix == ".doc":
        paras = _paragraphs_doc(path)
    else:
        raise ValueError(f"Unsupported Word format: {suffix}")

    professor = Path(path).stem
    rows: list[Row] = []
    for idx, text in enumerate(paras, start=1):
        if not text:
            continue
        isbn = _find_isbn(text)
        if isbn:
            if rows and not rows[-1].isbn:
                rows[-1].isbn = isbn  # belongs to the preceding book record
            continue
        if _is_header(text) or _is_status(text):
            continue
        title = _strip_leading_number(text)
        if title:
            rows.append(
                Row(
                    professor=professor,
                    title=title,
                    author="",
                    publisher="",
                    source_row=idx,
                    isbn="",
                )
            )
    return rows
