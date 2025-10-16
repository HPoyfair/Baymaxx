# invoicing.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, Any, List, Tuple, Optional
import json
import uuid
import os
import re
import csv
from datetime import date, datetime

# ---------- internal data paths (app-local) ----------
DATA_DIR = Path(__file__).resolve().parent / "data"
INVOICES_DIR = DATA_DIR / "invoices"  # internal storage (unchanged)
SETTINGS_PATH = DATA_DIR / "invoicing_settings.json"

# ---------- user-visible default ----------
DEFAULT_USER_INVOICE_ROOT = Path.home() / "Baymaxx Invoices"


# ---------- small utils ----------
def _ensure_dirs() -> None:
    """Make sure app-local data dirs exist (data/invoices)."""
    INVOICES_DIR.mkdir(parents=True, exist_ok=True)
    DATA_DIR.mkdir(parents=True, exist_ok=True)


def _atomic_write_text(path: Path, text: str) -> None:
    """Write via a temp file and replace to avoid partial writes."""
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def _new_id() -> str:
    return str(uuid.uuid4())


def _recalc_totals(inv: Dict[str, Any]) -> None:
    """Recalculate amounts and totals in-place."""
    items: List[Dict[str, Any]] = inv.get("line_items", [])
    subtotal = 0.0
    for it in items:
        qty = float(it.get("qty", 0))
        price = float(it.get("unit_price", 0))
        it["amount"] = round(qty * price, 2)
        subtotal += it["amount"]
    inv.setdefault("totals", {})
    inv["totals"]["subtotal"] = round(subtotal, 2)
    tax_rate = float(inv.get("tax_rate", 0.0))
    inv["totals"]["tax"] = round(subtotal * tax_rate, 2)
    inv["totals"]["total"] = round(inv["totals"]["subtotal"] + inv["totals"]["tax"], 2)


# ---------- settings (remember user's chosen folder) ----------
def _load_settings() -> Dict[str, Any]:
    try:
        if SETTINGS_PATH.exists():
            return json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}


def _save_settings(d: Dict[str, Any]) -> None:
    _ensure_dirs()
    _atomic_write_text(SETTINGS_PATH, json.dumps(d, indent=2, ensure_ascii=False) + "\n")


def get_remembered_invoice_root() -> Optional[Path]:
    """Return the remembered user invoice root, if any and still valid."""
    s = _load_settings()
    p = s.get("invoice_root")
    if not p:
        return None
    pp = Path(p)
    return pp if pp.exists() and pp.is_dir() else None


def set_remembered_invoice_root(path: Path) -> None:
    s = _load_settings()
    s["invoice_root"] = str(path)
    _save_settings(s)


def ensure_invoice_root(parent: Optional[object] = None) -> Optional[Path]:
    """
    Ensure a user-visible folder exists for saving/exporting invoices.
    - If a remembered folder exists, return it.
    - Otherwise prompt to create the default folder (~/Baymaxx Invoices).
    - If declined, let the user pick a folder via directory dialog.
    Returns the chosen/created Path, or None if the user cancels.
    NOTE: `parent` can be a Tk widget/root; prompts are silent if Tk is absent.
    """
    # 1) already remembered?
    remembered = get_remembered_invoice_root()
    if remembered:
        return remembered

    # If the default already exists, use it and remember.
    if DEFAULT_USER_INVOICE_ROOT.exists():
        set_remembered_invoice_root(DEFAULT_USER_INVOICE_ROOT)
        return DEFAULT_USER_INVOICE_ROOT

    # Try to prompt via Tk if available
    askyesno = None
    askdirectory = None
    try:
        from tkinter import messagebox, filedialog
        askyesno = messagebox.askyesno
        askdirectory = filedialog.askdirectory
    except Exception:
        # non-GUI environment; create default automatically
        DEFAULT_USER_INVOICE_ROOT.mkdir(parents=True, exist_ok=True)
        set_remembered_invoice_root(DEFAULT_USER_INVOICE_ROOT)
        return DEFAULT_USER_INVOICE_ROOT

    # 2) Ask to create default folder
    create = askyesno(
        "Create Invoices Folder",
        f"Baymaxx needs a place to save invoices.\n\n"
        f"Create this folder?\n\n{DEFAULT_USER_INVOICE_ROOT}"
    )
    if create:
        try:
            DEFAULT_USER_INVOICE_ROOT.mkdir(parents=True, exist_ok=True)
            set_remembered_invoice_root(DEFAULT_USER_INVOICE_ROOT)
            return DEFAULT_USER_INVOICE_ROOT
        except Exception as e:
            # fallback: let user choose
            pass

    # 3) Let the user pick a different folder
    chosen = askdirectory(
        title="Choose a folder to save invoices",
        initialdir=str(Path.home())
    )
    if not chosen:
        return None  # user cancelled

    chosen_path = Path(chosen)
    try:
        chosen_path.mkdir(parents=True, exist_ok=True)
    except Exception:
        # If cannot create, bail
        return None

    set_remembered_invoice_root(chosen_path)
    return chosen_path


def invoice_output_dir() -> Path:
    """
    Return the active user-visible invoices folder (guaranteed to exist).
    If none is remembered, returns DEFAULT_USER_INVOICE_ROOT (creating it).
    This is safe to call in non-GUI contexts.
    """
    p = get_remembered_invoice_root()
    if p:
        return p
    # Non-GUI auto-create
    DEFAULT_USER_INVOICE_ROOT.mkdir(parents=True, exist_ok=True)
    set_remembered_invoice_root(DEFAULT_USER_INVOICE_ROOT)
    return DEFAULT_USER_INVOICE_ROOT


# ---------- invoice API ----------
def new_monthly_invoice(
    year: int,
    month: int,
    client_id: str | None = None,
    client_name_snapshot: str | None = None,
    tax_rate: float = 0.0,
) -> Dict[str, Any]:
    """
    Create an in-memory invoice dict for a monthly period.
    (Not saved until you call save_invoice().)
    """
    _ensure_dirs()
    inv: Dict[str, Any] = {
        "id": _new_id(),
        "type": "monthly",
        "period": {"year": int(year), "month": int(month)},  # 1..12
        "client_id": client_id,
        "client_name_snapshot": client_name_snapshot,
        "created_at": None,      # you can fill timestamps later
        "tax_rate": float(tax_rate),
        "line_items": [],        # [{description, qty, unit_price, amount}]
        "totals": {"subtotal": 0.0, "tax": 0.0, "total": 0.0},
        "notes": "",
    }
    return inv


def add_line_item(inv: Dict[str, Any], description: str, qty: float, unit_price: float) -> None:
    """Append a line item and recalc totals."""
    inv.setdefault("line_items", [])
    inv["line_items"].append({
        "description": description.strip(),
        "qty": float(qty),
        "unit_price": float(unit_price),
        "amount": 0.0,  # computed below
    })
    _recalc_totals(inv)


def set_client(inv: Dict[str, Any], client_id: str, client_name_snapshot: str) -> None:
    inv["client_id"] = client_id
    inv["client_name_snapshot"] = client_name_snapshot


def set_tax_rate(inv: Dict[str, Any], tax_rate: float) -> None:
    inv["tax_rate"] = float(tax_rate)
    _recalc_totals(inv)


def save_invoice(inv: Dict[str, Any]) -> Path:
    """
    Persist the invoice as data/invoices/<id>.json (atomic write).
    Returns the internal data path.
    (Use invoice_output_dir() for the user-visible folder when exporting.)
    """
    _ensure_dirs()
    if not inv.get("id"):
        inv["id"] = _new_id()
    _recalc_totals(inv)
    path = INVOICES_DIR / f"{inv['id']}.json"
    _atomic_write_text(path, json.dumps(inv, indent=2, ensure_ascii=False) + "\n")
    return path


def load_invoice(invoice_id: str) -> Dict[str, Any] | None:
    """Load a single invoice by id."""
    path = INVOICES_DIR / f"{invoice_id}.json"
    if not path.exists():
        return None
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return None


def list_invoices() -> List[Dict[str, Any]]:
    """
    Return a lightweight listing of invoices (id, type, period, client, total).
    """
    _ensure_dirs()
    out: List[Dict[str, Any]] = []
    for p in sorted(INVOICES_DIR.glob("*.json")):
        try:
            doc = json.loads(p.read_text(encoding="utf-8"))
            out.append({
                "id": doc.get("id"),
                "type": doc.get("type"),
                "period": doc.get("period"),
                "client_id": doc.get("client_id"),
                "client_name": doc.get("client_name_snapshot"),
                "total": (doc.get("totals") or {}).get("total", 0.0),
            })
        except Exception:
            # skip corrupt files
            continue
    return out


# ---------- CSV helpers (kind + source number) ----------
def _norm(s: str) -> str:
    """Normalize header names: lowercase + strip spaces/underscores/dashes."""
    return re.sub(r"[\s_\-]+", "", (s or "").strip().lower())


def _detect_kind(fieldnames: List[str]) -> str:
    """
    Decide 'messages' vs 'calls' using explicit columns:
      - messages if a 'NumSegments' column exists
      - calls     if a 'Duration'   column exists
    If both or neither are present, return 'unknown'.
    """
    if not fieldnames:
        return "unknown"

    normalized = {_norm(h) for h in fieldnames}
    has_numsegments = "numsegments" in normalized or "numofsegments" in normalized
    has_duration = "duration" in normalized

    if has_numsegments and not has_duration:
        return "messages"
    if has_duration and not has_numsegments:
        return "calls"
    return "unknown"


def sniff_csv(path: str | Path) -> Tuple[str, List[str]]:
    """
    Read the header of a CSV and return (kind, headers).
    kind ∈ {'messages','calls','unknown'} based on headers present.
    """
    p = Path(path)
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        headers = next(reader, [])
    return _detect_kind(headers), headers


def _clean_phone(raw: str) -> str:
    """Normalize a phone string to digits with optional leading '+'."""
    if not raw:
        return ""
    raw = raw.strip()
    lead_plus = raw.startswith("+")
    digits = re.sub(r"\D+", "", raw)
    return ("+" if lead_plus else "") + digits


def identify_source(path: str | Path) -> Dict[str, Any]:
    """
    Inspect a CSV file to determine:
      - kind: 'messages' or 'calls' (or 'unknown')
      - raw_number: value from the 'From' column (first non-empty in file)
      - number: normalized phone
    """
    p = Path(path)
    kind, headers = sniff_csv(p)

    # Find the "From" column (allow variants)
    candidate_names = {"from", "sender", "source", "callerid", "caller"}
    header_index = None
    normalized = [_norm(h) for h in headers]
    for i, hn in enumerate(normalized):
        if hn in candidate_names:
            header_index = i
            break

    raw_number = ""
    if header_index is not None:
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            next(reader, None)  # skip header
            for row in reader:
                if header_index < len(row):
                    raw = row[header_index].strip()
                    if raw:
                        raw_number = raw
                        break

    return {
        "kind": kind,                  # 'messages' | 'calls' | 'unknown'
        "raw_number": raw_number,      # as seen in CSV
        "number": _clean_phone(raw_number),  # normalized
        "headers": headers,            # actual header row
    }
# --- add near the top ---
from datetime import date, datetime

# ... keep your existing code above this line ...


# ---------- export naming helpers ----------
def monthly_filename_for_today() -> str:
    """
    Return a safe filename like '2025-10-15 Monthly.json'.
    (No slashes so it works on Windows/macOS/Linux.)
    """
    today = date.today()
    return f"{today:%Y-%m-%d} Monthly.json"


def monthly_output_path(root: Path | None = None) -> Path:
    """
    Return the full path to save a monthly invoice JSON under the user's invoice root.
    Ensures the root exists.
    """
    root = (root or invoice_output_dir())
    root.mkdir(parents=True, exist_ok=True)
    return root / monthly_filename_for_today()


# ---------- phone matching across clients/divisions/sites ----------
def _match_phone_in_clients(clients_doc: Dict[str, Any], number: str) -> Dict[str, Any] | None:
    """
    Given the loaded clients.json-like document and a normalized number,
    find a site whose phone matches. Returns a breadcrumb dict or None.
    Expected structure:
      {"clients":[
          {"id":.., "name":.., "divisions":[
              {"id":.., "name":.., "sites":[
                  {"id":.., "name":.., "phone":..}, ...
              ]}]}]}
    """
    if not number:
        return None

    for c in (clients_doc.get("clients") or []):
        if not isinstance(c, dict):
            continue
        cid, cname = c.get("id"), c.get("name", "")
        for d in (c.get("divisions") or []):
            if not isinstance(d, dict):
                continue
            did, dname = d.get("id"), d.get("name", "")
            for s in (d.get("sites") or []):
                if not isinstance(s, dict):
                    continue
                sid, sname = s.get("id"), s.get("name", "")
                phone_norm = _clean_phone(s.get("phone", ""))
                if phone_norm and phone_norm == number:
                    return {
                        "client_id": cid, "client_name": cname,
                        "division_id": did, "division_name": dname,
                        "site_id": sid, "site_name": sname,
                    }
    return None


# ---------- CSV wrapper expected by the UI ----------
def identify_csv_and_phone(path: str | Path, clients_doc: Dict[str, Any]) -> Dict[str, Any]:
    """
    Wrapper used by the UI. Detects file kind + source phone and
    tries to match it to a client/division/site.
    Returns:
      {
        "kind": "messages" | "calls" | "unknown",
        "phone": "+15551234567",
        "match": {client_id/name, division_id/name, site_id/name} | None
      }
    """
    src = identify_source(path)             # {kind, raw_number, number, headers}
    number = src.get("number", "")
    match = _match_phone_in_clients(clients_doc or {}, number) if number else None
    return {
        "kind": src.get("kind", "unknown"),
        "phone": number,
        "match": match
    }

# ---------- matching helpers (CSV -> site by last-4) ----------

def _digits_only(s: str) -> str:
    return re.sub(r"\D+", "", s or "")

def _match_site_by_last4(clients_doc, phone_digits: str) -> dict | None:
    """
    Find a site whose phone ends with the last 4 digits of `phone_digits`.
    Works with the current clients.json structure:
      client -> divisions[] -> sites[] (each site has 'phone').
    Returns a small breadcrumb dict or None.
    """
    if not phone_digits:
        return None
    last4 = phone_digits[-4:]

    # clients_doc may be a dict {"clients":[...]} or a list [...]
    candidates = clients_doc.get("clients") if isinstance(clients_doc, dict) else clients_doc
    if not isinstance(candidates, list):
        return None

    for c in candidates:
        client_name = (c or {}).get("name", "")
        for d in (c or {}).get("divisions", []) or []:
            division_name = (d or {}).get("name", "")
            for s in (d or {}).get("sites", []) or []:
                site_phone_digits = _digits_only((s or {}).get("phone", ""))
                if site_phone_digits.endswith(last4):
                    return {
                        "client_id": (c or {}).get("id"),
                        "client_name": client_name,
                        "division_id": (d or {}).get("id"),
                        "division_name": division_name,
                        "site_id": (s or {}).get("id"),
                        "site_name": (s or {}).get("name", ""),
                        "site_phone": site_phone_digits,
                        "matched_last4": last4,
                    }
    return None

def identify_csv_and_phone(path: str | Path, clients_doc=None) -> dict:
    """
    Convenience for the UI:
    - detect kind ('messages' | 'calls' | 'unknown')
    - pull 'From' number
    - try to match by last-4 digits to a site (if clients_doc provided)
    """
    base = identify_source(path)
    phone_digits = _digits_only(base.get("number") or base.get("raw_number") or "")
    match = _match_site_by_last4(clients_doc, phone_digits) if clients_doc else None

    return {
        "kind": base.get("kind", "unknown"),
        "phone": phone_digits,
        "match": match,         # dict | None
        "headers": base.get("headers", []),
    }
# ---------- month/year validation for CSVs ----------
def _ym_from_any_date(s: str) -> tuple[int | None, int | None]:
    """
    Extract (year, month) from any date-time-ish string by regexing 'YYYY-MM-DD'.
    Works for:
      - '2025-05-31T20:26:26-07:00'
      - '13:00:35 PDT 2025-05-31'
    Returns (None, None) if not found.
    """
    if not s:
        return (None, None)
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", s)
    if not m:
        return (None, None)
    y = int(m.group(1))
    mo = int(m.group(2))
    return (y, mo)


def check_csv_month_year(path: str | Path, kind: str, year: int, month: int) -> tuple[bool, dict]:
    """
    Scan the CSV and check if *all* rows fall within the given (year, month).
    kind: 'messages' or 'calls' (anything else -> unknown -> False)
    Returns (all_ok, stats) where stats = {'in': n_in, 'out': n_out, 'rows': n_total}
    """
    p = Path(path)
    try:
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            headers = next(reader, [])
            norm = [_norm(h) for h in headers]

            # choose the date column based on file kind
            if kind == "messages":
                # e.g., "SentDate"
                candidates = {"sentdate", "date", "timestamp"}
            elif kind == "calls":
                # e.g., "Start Time"
                candidates = {"starttime", "start", "calldate"}
            else:
                return (False, {"in": 0, "out": 0, "rows": 0})

            idx = None
            for i, hn in enumerate(norm):
                if hn in candidates:
                    idx = i
                    break
            if idx is None:
                # can't find a date column -> fail safe
                return (False, {"in": 0, "out": 0, "rows": 0})

            n_in = n_out = n_total = 0
            for row in reader:
                n_total += 1
                val = row[idx] if idx < len(row) else ""
                y, m = _ym_from_any_date(val)
                if y == year and m == month:
                    n_in += 1
                else:
                    n_out += 1

        return (n_out == 0 and n_total > 0, {"in": n_in, "out": n_out, "rows": n_total})
    except Exception:
        return (False, {"in": 0, "out": 0, "rows": 0})
# ---------- month/year row finder for preview highlighting ----------

def _date_col_index(headers: list[str], kind: str) -> int | None:
    """Return index of the date column we should check for this kind."""
    norm = lambda s: re.sub(r"[\s_\-]+", "", s.strip().lower())
    normed = [norm(h) for h in headers]
    if kind == "messages":
        # Twilio uses 'SentDate'
        targets = {"sentdate"}
    elif kind == "calls":
        # Twilio uses 'Start Time'
        targets = {"starttime"}
    else:
        targets = set()

    for i, n in enumerate(normed):
        if n in targets:
            return i
    return None


_date_re = re.compile(r"(\d{4})-(\d{2})-(\d{2})")

def _ym_from_cell(cell: str) -> tuple[int | None, int | None]:
    """Extract (year, month) from a cell by regex like 'YYYY-MM-DD'."""
    if not cell:
        return (None, None)
    m = _date_re.search(cell)
    if not m:
        return (None, None)
    y, mo, _ = m.groups()
    try:
        return (int(y), int(mo))
    except Exception:
        return (None, None)


def find_out_of_month_rows(path: str | Path, kind: str, year: int, month: int) -> list[tuple[int, str]]:
    """
    Return a list of (row_number, cell_value) for rows whose date is NOT in (year, month).
    row_number is 1-based counting the header as row 1 (so first data row is 2).
    """
    p = Path(path)
    try:
        with p.open("r", encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            headers = next(reader, [])
            ci = _date_col_index(headers, kind)
            out: list[tuple[int, str]] = []

            if ci is None:
                # No date column? mark every data row as out-of-range with reason
                for idx, row in enumerate(reader, start=2):
                    out.append((idx, "date-column-missing"))
                return out

            for idx, row in enumerate(reader, start=2):
                cell = row[ci] if ci < len(row) else ""
                y, m = _ym_from_cell(cell)
                if y is None or m is None or y != int(year) or m != int(month):
                    out.append((idx, cell))
            return out
    except Exception:
        # On any read error, just say everything is bad so the UI warns the user
        try:
            # best-effort to count rows for user feedback
            with p.open("r", encoding="utf-8-sig", newline="") as f2:
                n = sum(1 for _ in f2)
        except Exception:
            n = 0
        return [(i, "read-error") for i in range(2, max(2, n + 1))]
