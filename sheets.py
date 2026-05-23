"""
Google Sheets layer
"""

import random
import time
import unicodedata
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

# ── DB config ───────────────────────────────────────

CREDENTIALS_PATH = '/path/to/your/credentials.json'
SPREADSHEET_NAME = 'your_spreadsheet'

SCOPE = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive',
]

# Expected header rows:
FACT_HEADER = ['voting', 'date', 'member', 'fraction', 'result'] #sheet1
DIM_HEADER = ['voting', 'term', 'votingname', 'votingurl', 'tags'] #sheet2

"""
 Quota limit handling:
Google Sheets throttles to a per-minute request quota per user. Exceeding it returns HTTP 429 (also 503). 
A batch data-entry session can hit this. _with_retry transparently waits and retries ONLY those transient statuses
"""

RETRY_STATUSES = {429, 503}
RETRY_MAX_ATTEMPTS = 5
RETRY_BASE_DELAY = 1.0  # seconds; doubles each attempt, plus jitter


class QuotaExceeded(RuntimeError):
    """Raised when retries are exhausted on a 429/503."""


def _status_of(exc):
    resp = getattr(exc, "response", None)
    code = getattr(resp, "status_code", None)
    if isinstance(code, int):
        return code
    return None


def _with_retry(fn):
    delay = RETRY_BASE_DELAY
    last = None
    for attempt in range(RETRY_MAX_ATTEMPTS):
        try:
            return fn()
        except gspread.exceptions.APIError as e:
            if _status_of(e) not in RETRY_STATUSES:
                raise  # genuine error — do not mask it behind retries
            last = e
            if attempt == RETRY_MAX_ATTEMPTS - 1:
                break
            time.sleep(delay + random.uniform(0, 0.5))
            delay *= 2
    raise QuotaExceeded(
        f"Google Sheets rate limit not cleared after "
        f"{RETRY_MAX_ATTEMPTS} attempts; try again shortly."
    ) from last


_gc = None
_sheets = None  # (fact_sheet, dim_sheet)


def _client():
    global _gc
    if _gc is None:
        creds = Credentials.from_service_account_file(
            CREDENTIALS_PATH, scopes=SCOPE)
        _gc = gspread.authorize(creds)
    return _gc


def reset_sheets():
    global _gc, _sheets
    _gc = None
    _sheets = None


# ── helpers ───────

def _accent_fold(s):
    """Lowercase and strip"""
    s = str(s).strip().lower()
    return "".join(c for c in unicodedata.normalize("NFKD", s)
                   if not unicodedata.combining(c))


def normalize_tags(raw):
    out, seen = [], set()
    for part in str(raw).split(","):
        t = " ".join(part.split())  # trim + collapse internal whitespace
        if not t:
            continue
        key = t.casefold()
        if key in seen:
            continue
        seen.add(key)
        out.append(t)
    return out


def tag_match_keys(raw):
    return {t.casefold() for t in normalize_tags(raw)}


# ── Google Sheets helpers ──────────────────────────────────

def get_sheets():
    global _sheets
    if _sheets is None:
        wb = _with_retry(lambda: _client().open(SPREADSHEET_NAME))
        _sheets = (wb.sheet1, wb.get_worksheet(1))
    return _sheets


def verify_schema():
    fact_sheet, dim_sheet = get_sheets()
    problems = []
    fact_hdr = [h.strip() for h in _with_retry(lambda: fact_sheet.row_values(1))]
    dim_hdr = [h.strip() for h in _with_retry(lambda: dim_sheet.row_values(1))]
    if fact_hdr[:len(FACT_HEADER)] != FACT_HEADER:
        problems.append(
            f"Sheet 1 header is {fact_hdr or '(empty)'}; "
            f"expected {FACT_HEADER}")
    if dim_hdr[:len(DIM_HEADER)] != DIM_HEADER:
        problems.append(
            f"Sheet 2 header is {dim_hdr or '(empty)'}; "
            f"expected {DIM_HEADER}")
    if problems:
        raise ValueError("Spreadsheet schema mismatch:\n" + "\n".join(problems))


def delete_voting_from_sheets(voting_id):
    """Delete all rows matching voting_id from both sheets.

    Crash-safe by construction: there is never a moment where the sheet
    is empty on the server. Instead of clear()-then-rewrite (which leaves
    a data-loss window if the rewrite fails), this overwrites the kept
    rows in place starting at A1, then deletes only the now-redundant
    TRAILING rows. If the process dies mid-operation the worst case is a
    few duplicate trailing rows, never a wiped sheet.

    No-op deletes (nothing matches) touch the sheet zero times."""
    fact_sheet, dim_sheet = get_sheets()
    vid = str(voting_id)

    for sheet in (fact_sheet, dim_sheet):
        all_rows = _with_retry(lambda s=sheet: s.get_all_values())
        if not all_rows:
            continue
        header = all_rows[0]
        kept = [header] + [r for r in all_rows[1:] if r and str(r[0]) != vid]

        if len(kept) == len(all_rows):
            continue  # nothing matched — do not touch the sheet at all

        _with_retry(lambda s=sheet, k=kept: s.update(
            values=k, range_name="A1"))

        extra = len(all_rows) - len(kept)
        if extra > 0:
            start = len(kept) + 1  #
            end = len(all_rows)
            _with_retry(lambda s=sheet, a=start, b=end:
                        s.delete_rows(a, b))


def get_existing_voting_ids():
    fact_sheet, dim_sheet = get_sheets()
    ids = set()
    for sheet in (fact_sheet, dim_sheet):
        col = _with_retry(lambda s=sheet: s.col_values(1))  # col 1 = 'voting'
        ids.update(str(v) for v in col[1:] if v)  # skip header
    return ids


def check_duplicate(voting_id):
    return str(voting_id) in get_existing_voting_ids()


def append_to_sheets(voting_df, dim_row):
    fact_sheet, dim_sheet = get_sheets()
    _with_retry(lambda: fact_sheet.append_rows(voting_df.values.tolist()))
    _with_retry(lambda: dim_sheet.append_rows([dim_row]))


def load_db_summary():
    fact_sheet, dim_sheet = get_sheets()
    data = _with_retry(lambda: fact_sheet.get_all_values())
    if len(data) < 2:
        return [], []
    df = pd.DataFrame(data[1:], columns=data[0])

    # Dimension lookup keyed by voting id
    dims = {}
    tag_seen, all_tags = set(), []
    dim_data = _with_retry(lambda: dim_sheet.get_all_values())
    if len(dim_data) >= 2:
        dheader = dim_data[0]
        for row in dim_data[1:]:
            if not row:
                continue
            rec = dict(zip(dheader, row))
            tags = normalize_tags(rec.get('tags', ''))
            dims[str(rec.get('voting', ''))] = (
                rec.get('term', ''),
                rec.get('votingname', ''),
                tags,
            )
            for t in tags:
                k = t.casefold()
                if k not in tag_seen:
                    tag_seen.add(k)
                    all_tags.append(t)

    summary = (df.groupby('voting')
                 .agg(records=('voting', 'count'), date=('date', 'first'))
                 .reset_index()
                 .sort_values('date', ascending=False))

    out = []
    for _, r in summary.iterrows():
        vid = str(r['voting'])
        term, name, tags = dims.get(vid, ("", "", []))
        # Row layout: [term, voting_id, date, voting_name, tags, records]
        out.append([term, vid, r['date'], name,
                    ", ".join(tags), int(r['records'])])
    return out, sorted(all_tags, key=str.casefold)


def load_voting_detail(voting_id):
    fact_sheet, _ = get_sheets()
    data = _with_retry(lambda: fact_sheet.get_all_values())
    if len(data) < 2:
        return pd.DataFrame(columns=['member', 'fraction', 'result'])
    df = pd.DataFrame(data[1:], columns=data[0])
    df = df[df['voting'].astype(str) == str(voting_id)]
    keep = [c for c in ['member', 'fraction', 'result'] if c in df.columns]
    return df[keep].reset_index(drop=True)


def load_voting_meta(voting_id):
    _, dim_sheet = get_sheets()
    data = _with_retry(lambda: dim_sheet.get_all_values())
    empty = {'name': '', 'url': '', 'term': '', 'tags': []}
    if len(data) < 2:
        return empty
    header = data[0]
    vid = str(voting_id)
    for row in data[1:]:
        if not row:
            continue
        rec = dict(zip(header, row))
        if str(rec.get('voting', '')) == vid:
            return {
                'name': rec.get('votingname', ''),
                'url': rec.get('votingurl', ''),
                'term': rec.get('term', ''),
                'tags': normalize_tags(rec.get('tags', '')),
            }
    return empty
