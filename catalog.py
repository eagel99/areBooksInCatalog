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


def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update(_HEADERS)
    return s


def search(title: str, session: requests.Session) -> list[dict[str, Any]]:
    params = dict(_BASE_PARAMS)
    params["q"] = f"any,contains,{title}"

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

    title_s = fuzz.token_set_ratio(row.title, doc_title) if doc_title else 0

    author_candidates = [c for c in (doc_creator, doc_au, doc_aulast) if c]
    if row.author and author_candidates:
        author_s = max(fuzz.token_set_ratio(row.author, c) for c in author_candidates)
    elif not row.author:
        author_s = 70
    else:
        author_s = 0

    pub_candidates = [c for c in (doc_pub_disp, doc_pub_add) if c]
    if row.publisher and pub_candidates:
        pub_s = max(fuzz.partial_ratio(row.publisher, c) for c in pub_candidates)
    elif not row.publisher:
        pub_s = 70
    else:
        pub_s = 0

    composite = 0.6 * title_s + 0.3 * author_s + 0.1 * pub_s
    return Score(title=int(title_s), author=int(author_s), publisher=int(pub_s), composite=composite)


_LOCAL_CATEGORIES = ("Alma-P", "Alma-E")


def _is_local(doc: dict[str, Any]) -> bool:
    cats = (doc.get("delivery") or {}).get("deliveryCategory") or []
    return any(c in _LOCAL_CATEGORIES for c in cats)


def pick_best(row: Row, docs: list[dict[str, Any]]) -> dict[str, Any] | None:
    candidates: list[tuple[bool, float, dict[str, Any]]] = []
    for d in docs:
        s = score(row, d)
        if s.title < 70 or s.composite < 60:
            continue
        candidates.append((_is_local(d), s.composite, d))
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
