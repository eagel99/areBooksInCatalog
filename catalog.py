import re
import time
from dataclasses import dataclass
from typing import Any
from urllib.parse import quote

import requests
from rapidfuzz import fuzz

from xlsx_io import Row

PRIMO_ENDPOINT = "https://primo.bgu.ac.il/primaws/rest/pub/pnxs"
PRIMO_DISCOVERY = "https://primo.bgu.ac.il/discovery/fulldisplay"
VID = "972BGU_INST:972BGU"

_BASE_PARAMS = {
    "blendFacetsSeparately": "false",
    "disableCache": "false",
    "getMore": "0",
    "inst": "972BGU_INST",
    "lang": "he",
    "limit": "10",
    "offset": "0",
    "otbRanking": "false",
    "pcAvailability": "true",
    "pfilter": "rtype,exact,books",
    "rtaLinks": "true",
    "scope": "MyInst_and_CI",
    "skipDelivery": "N",
    "sort": "rank",
    "tab": "Everything",
    "vid": VID,
}

_HEADERS = {
    "Accept": "application/json",
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    ),
}

REQUEST_DELAY_SEC = 0.4
MAX_ATTEMPTS = 3

TITLE_MIN = 75
AUTHOR_MIN = 50
COMPOSITE_MIN = 60

_EDITOR_TOKENS = re.compile(
    r"\(?\b(?:edited\s+by|editors?|eds?\.?)\b\)?",
    re.IGNORECASE,
)

_CAMEL_BOUNDARY = re.compile(r"([a-z])([A-Z])")


def _clean_author(s: str) -> str:
    if not s:
        return s
    s = _EDITOR_TOKENS.sub(" ", s)
    return re.sub(r"\s+", " ", s).strip()


def _normalize_query(s: str) -> str:
    return _CAMEL_BOUNDARY.sub(r"\1 \2", s)


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(_HEADERS)
    return s


def search(title: str, session: requests.Session) -> list[dict[str, Any]]:
    params = dict(_BASE_PARAMS)
    params["q"] = f"any,contains,{_normalize_query(title)}"

    for attempt in range(1, MAX_ATTEMPTS + 1):
        try:
            r = session.get(PRIMO_ENDPOINT, params=params, timeout=20)
        except requests.RequestException:
            if attempt == MAX_ATTEMPTS:
                raise
            time.sleep(2 ** attempt)
            continue

        if r.status_code in (429, 500, 502, 503, 504):
            if attempt == MAX_ATTEMPTS:
                r.raise_for_status()
            time.sleep(2 ** attempt)
            continue

        r.raise_for_status()
        time.sleep(REQUEST_DELAY_SEC)
        return r.json().get("docs", []) or []

    return []


def _first(d: dict, *path, default: str = "") -> str:
    cur: Any = d
    for p in path:
        if isinstance(cur, dict):
            cur = cur.get(p)
        else:
            return default
        if cur is None:
            return default
    if isinstance(cur, list):
        return str(cur[0]).strip() if cur else default
    return str(cur).strip() if cur else default


def _strip_marc(s: str) -> str:
    return s.split("$$", 1)[0].strip() if s else s


@dataclass
class Score:
    title: int
    author: int
    publisher: int
    composite: float


def score(row: Row, doc: dict[str, Any]) -> Score:
    doc_title = _first(doc, "pnx", "display", "title")
    doc_creator = _strip_marc(_first(doc, "pnx", "display", "creator"))
    doc_au = _strip_marc(_first(doc, "pnx", "addata", "au"))
    doc_aulast = _strip_marc(_first(doc, "pnx", "addata", "aulast"))
    doc_pub_disp = _first(doc, "pnx", "display", "publisher")
    doc_pub_add = _strip_marc(_first(doc, "pnx", "addata", "pub"))

    if doc_title:
        ts = fuzz.token_set_ratio(row.title, doc_title)
        # partial_ratio captures the "user gave only the main title; catalog
        # has a long subtitle" case. Only override when substring match is
        # very strong (>=90), to avoid loosening the title gate everywhere.
        pr = fuzz.partial_ratio(row.title.lower(), doc_title.lower())
        title_s = max(ts, pr) if pr >= 90 else ts
    else:
        title_s = 0

    row_author = _clean_author(row.author)
    # Only use primary author fields (creator, au); aulast alone is often
    # catalog-noise from a chapter contributor on edited volumes.
    primary_authors = [_clean_author(c) for c in (doc_creator, doc_au) if c]
    primary_authors = [c for c in primary_authors if c]
    aulast_clean = _clean_author(doc_aulast)
    if row_author and primary_authors:
        author_s = max(fuzz.token_set_ratio(row_author, c) for c in primary_authors)
    elif not row_author:
        author_s = 70
    elif aulast_clean:
        # Catalog has aulast but no full author/creator. Use it as a soft
        # check: if the user-supplied surname appears, score moderately;
        # otherwise treat as no-author (sentinel handled in pick_best).
        if fuzz.partial_ratio(row_author, aulast_clean) >= 80:
            author_s = 70
        else:
            author_s = -1
    else:
        # Catalog record has no creator info (e.g. edited volume).
        # Sentinel is handled in pick_best with a stricter title gate.
        author_s = -1

    pub_candidates = [c for c in (doc_pub_disp, doc_pub_add) if c]
    if row.publisher and pub_candidates:
        pub_s = max(fuzz.partial_ratio(row.publisher, c) for c in pub_candidates)
    elif not row.publisher:
        pub_s = 70
    else:
        pub_s = 0

    if author_s == -1:
        composite = (0.6 * title_s + 0.1 * pub_s) / 0.7
    else:
        composite = 0.6 * title_s + 0.3 * author_s + 0.1 * pub_s
    return Score(title=int(title_s), author=int(author_s), publisher=int(pub_s), composite=composite)


_LOCAL_CATEGORIES = ("Alma-P", "Alma-E")


def _is_local(doc: dict[str, Any]) -> bool:
    cats = (doc.get("delivery") or {}).get("deliveryCategory") or []
    return any(c in _LOCAL_CATEGORIES for c in cats)


NO_AUTHOR_TITLE_MIN = 85
STRONG_AUTHOR_MIN = 95
STRONG_AUTHOR_TITLE_MIN = 70


def pick_best(row: Row, docs: list[dict[str, Any]]) -> dict[str, Any] | None:
    candidates: list[tuple[bool, float, dict[str, Any]]] = []
    for d in docs:
        s = score(row, d)
        local = _is_local(d)
        if s.author == -1:
            # Catalog record has no creator (e.g. edited volume).
            # Restrict to local Alma records — remote indexes often have
            # blank creator on review/article records that aren't the book.
            if not local:
                continue
            if s.title < NO_AUTHOR_TITLE_MIN or s.composite < COMPOSITE_MIN:
                continue
        else:
            # Near-perfect author match relaxes the title threshold to 70.
            title_min = STRONG_AUTHOR_TITLE_MIN if s.author >= STRONG_AUTHOR_MIN else TITLE_MIN
            if s.title < title_min or s.composite < COMPOSITE_MIN:
                continue
            if row.author and s.author < AUTHOR_MIN:
                continue
        candidates.append((local, s.composite, d))
    if not candidates:
        return None
    candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
    return candidates[0][2]


_AVAIL_POSITIVE = {
    "available",
    "available_in_library",
    "check_holdings",
    "fulltext",
    "open_access",
    "not_restricted",
}


def is_available(doc: dict[str, Any]) -> bool:
    delivery = doc.get("delivery") or {}
    categories = delivery.get("deliveryCategory") or []
    availability = [str(a).lower() for a in (delivery.get("availability") or [])]

    if any(c in _LOCAL_CATEGORIES for c in categories):
        return True

    return any(a in _AVAIL_POSITIVE for a in availability)


def _values(d: dict, *path) -> list[str]:
    cur: Any = d
    for p in path:
        if isinstance(cur, dict):
            cur = cur.get(p)
        else:
            return []
        if cur is None:
            return []
    if isinstance(cur, list):
        return [str(x) for x in cur]
    return [str(cur)]


def _isbn_in_doc(isbn: str, doc: dict[str, Any]) -> bool:
    target = re.sub(r"[^0-9Xx]", "", isbn).upper()
    if not target:
        return False
    fields = _values(doc, "pnx", "addata", "isbn") + _values(doc, "pnx", "display", "identifier")
    for raw in fields:
        if target in re.sub(r"[^0-9Xx]", "", raw).upper():
            return True
    return False


_PUB_WORDS = re.compile(
    r"\b(press|university|univ|publications?|editions?|éditions?|books?|classics?"
    r"|directions|routledge|penguin|verso|liveright|seagull|nyrb|harvard|yale|cornell"
    r"|princeton|chicago|wisconsin|garland|duculot|pub|up)\b",
    re.IGNORECASE,
)
_YEAR = re.compile(r"\b(?:19|20)\d{2}\b")
_EDITOR_PARENS = re.compile(r"\(\s*(?:eds?\.?|editors?|edited by)\s*\)", re.IGNORECASE)


def _search_query(title: str) -> str:
    """Reduce a free-text citation to a high-recall title/author query.

    Word citations append a publisher + year (e.g. "U of Notre Dame P, 2018")
    that the catalog's keyword AND-search cannot match, returning zero hits.
    Drop trailing publisher/year clauses and stray years; keep title + author.
    """
    text = _EDITOR_PARENS.sub(" ", title)
    clauses = re.split(r"\.\s+", text.strip().rstrip("."))
    while len(clauses) > 1:
        last = clauses[-1].strip()
        if _YEAR.search(last) or (len(last.split()) <= 6 and _PUB_WORDS.search(last)):
            clauses.pop()
        else:
            break
    q = _YEAR.sub(" ", " ".join(clauses))
    q = re.sub(r"\s+", " ", q).strip(" .,;:")
    return q or title


def find_doc(row: Row, session: requests.Session) -> dict[str, Any] | None:
    """Locate the best catalog record for a row.

    When the row carries an ISBN (typical of Word input), do a precise ISBN
    lookup and require the result to actually carry that ISBN; otherwise fall
    back to the existing fuzzy title/author match on a trimmed query.
    """
    if row.isbn:
        docs = search(row.isbn, session)
        matches = [d for d in docs if _isbn_in_doc(row.isbn, d)]
        if matches:
            matches.sort(key=_is_local, reverse=True)
            return matches[0]
    docs = search(_search_query(row.title), session)
    return pick_best(row, docs)


def permalink(doc: dict[str, Any]) -> str:
    recordid = _first(doc, "pnx", "control", "recordid")
    if not recordid:
        return ""
    context = doc.get("context") or "L"
    adaptor = doc.get("adaptor") or "Local Search Engine"
    return (
        f"{PRIMO_DISCOVERY}"
        f"?docid={quote(recordid, safe='')}"
        f"&context={quote(context, safe='')}"
        f"&adaptor={quote(adaptor, safe='')}"
        f"&vid={quote(VID, safe='')}"
        f"&lang=he"
    )
