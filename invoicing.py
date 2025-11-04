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


# ---------- required core helpers (paste into invoicing.py) ----------

def _ensure_line_items(inv: dict) -> list:
    """Make sure inv['line_items'] exists and is a list; return it."""
    if not isinstance(inv.get("line_items"), list):
        inv["line_items"] = []
    return inv["line_items"]

def recompute_totals(inv: dict) -> None:
    """Recalculate invoice total from line items (qty * unit_price)."""
    total = 0.0
    for li in inv.get("line_items", []):
        # ensure each line has an amount
        qty = float(li.get("qty", 0) or 0)
        price = float(li.get("unit_price", 0) or 0)
        amt = round(qty * price, 2)
        li["amount"] = amt
        total += amt
    inv["total"] = round(total, 2)

def add_item(inv: dict, description: str, qty: float, unit_price: float) -> None:
    """
    Append a line item and keep totals in sync.
    Called by add_voice_items_to_invoice / add_message_items_to_invoice.
    """
    items = _ensure_line_items(inv)
    qty = float(qty or 0)
    unit_price = float(unit_price or 0.0)
    amount = round(qty * unit_price, 2)

    items.append({
        "description": description or "",
        "qty": qty,
        "unit_price": unit_price,
        "amount": amount,
    })
    _recompute_totals(inv)

# (optional) public alias if any other code prefers this name
def add_line_item(inv: dict, description: str, qty: float, unit_price: float) -> None:
    _add_item(inv, description, qty, unit_price)
# --------------------------------------------------------------------


# ==== BEGIN: minimal line-item + totals shims (define only if missing) ====

# Public helper (create if missing)
if 'add_line_item' not in globals():
    def add_line_item(inv: dict, description: str, qty: float, unit_price: float) -> None:
        """
        Append a line item to inv['line_items'] and keep amount/total in sync.
        """
        qty = float(qty or 0)
        unit_price = float(unit_price or 0.0)
        amount = round(qty * unit_price, 2)

        li = {
            "description": description or "",
            "qty": qty,
            "unit_price": unit_price,
            "amount": amount,
        }
        inv.setdefault("line_items", []).append(li)
        # keep totals in sync
        if '_recompute_totals' in globals():
            _recompute_totals(inv)
        elif 'recompute_totals' in globals():
            recompute_totals(inv)
        else:
            # minimal inline total calc
            subtotal = round(sum(float(x.get("amount", 0) or 0) for x in inv.get("line_items", [])), 2)
            inv["total"] = subtotal

# Private alias expected by app.py or newer helpers
if '_add_item' not in globals():
    def _add_item(inv: dict, description: str, qty: float, unit_price: float) -> None:
        add_line_item(inv, description, qty, unit_price)

# Totals shim if project uses different name
if '_recompute_totals' not in globals():
    if 'recompute_totals' in globals():
        def _recompute_totals(inv: dict) -> None:
            recompute_totals(inv)
    else:
        def _recompute_totals(inv: dict) -> None:
            subtotal = round(sum(float(x.get("amount", 0) or 0) for x in inv.get("line_items", [])), 2)
            inv["total"] = subtotal

# ==== END: minimal line-item + totals shims ====



# ================== BEGIN: messages support & compat helpers ==================

# price fallback if not already defined elsewhere
try:
    UNIT_PRICE_SMS
except NameError:
    UNIT_PRICE_SMS = 0.14

# compat aliases if your project uses non-underscored names
if '_add_item' not in globals() and 'add_item' in globals():
    def _add_item(*a, **k): return add_item(*a, **k)
if '_recompute_totals' not in globals() and 'recompute_totals' in globals():
    def _recompute_totals(*a, **k): return recompute_totals(*a, **k)

# ---------------- phone resolution (re-usable) ----------------
def _normalize_site_key(s: str) -> str:
    u = (s or '').upper().strip()
    for suf in (' VOICE', ' SMS'):
        if u.endswith(suf):
            u = u[:-len(suf)].strip()
    if '–' in u:
        u = u.split('–', 1)[0].strip()
    u = ' '.join(u.split())
    return u

def _build_priority_phone_map(inv: dict) -> dict[str, str]:
    phones: dict[str, str] = {}
    # 1) UI-harvested matches
    sp = inv.get('site_phones') or {}
    if isinstance(sp, dict):
        for name, last4 in sp.items():
            last4 = str(last4 or '').strip()
            if last4.isdigit() and len(last4) == 4:
                phones[_normalize_site_key(name)] = last4

    # 2) clients.json (optional)
    load_clients = globals().get('_load_clients_doc') or globals().get('load_clients_doc')
    if callable(load_clients):
        try:
            clients_doc = load_clients(inv.get('clients_path'))
        except TypeError:
            clients_doc = load_clients()
        if isinstance(clients_doc, dict):
            for c in (clients_doc.get('clients') or []):
                for s in (c.get('sites') or []):
                    nm = (s.get('name') or '').strip()
                    ph = (s.get('phone') or '').strip()
                    if nm and ph.isdigit() and len(ph) == 4:
                        phones.setdefault(_normalize_site_key(nm), ph)
    return phones

def _lookup_last4(phones: dict[str, str], desc: str) -> str | None:
    if not desc:
        return None
    return phones.get(_normalize_site_key(desc))




# ---------------- CSV aggregation (month/year) ----------------
def _extract_row_datetime(row: dict, kind: str):
    """
    Try to parse a datetime from a CSV row for the given kind.

    Much more forgiving:
    - For calls: look for headers like 'Start Time', 'StartTime', 'Start', 'CallDate'
    - For messages: look for 'SentDate', 'Date', 'MessageDate', 'SendDate', 'Timestamp'
    - Then parse common formats, pulling out YYYY-MM-DD or MM/DD/YYYY wherever it appears.
    """
    from datetime import datetime
    import re

    # 1) choose the cell to inspect
    def _norm(h: str) -> str:
        return re.sub(r"[\s_\-]+", "", (h or "").strip().lower())

    if kind == "calls":
        wanted = {"starttime", "start", "calldate"}
    else:  # messages
        wanted = {"sentdate", "date", "messagedate", "senddate", "timestamp"}

    val = None
    for key, raw in row.items():
        if _norm(key) in wanted and raw not in (None, ""):
            val = str(raw).strip()
            break

    if not val:
        return None

    # 2) try ISO-ish first
    try:
        # handles '2025-10-31T13:00:53-07:00' and friends
        return datetime.fromisoformat(val.replace("Z", "+00:00"))
    except Exception:
        pass

    # 3) any 'YYYY-MM-DD' inside the string (e.g. '13:00:53 PDT 2025-10-31')
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", val)
    if m:
        y, mo, d = map(int, m.groups())
        try:
            return datetime(y, mo, d)
        except Exception:
            pass

    # 4) US-style 'MM/DD/YYYY' or 'M/D/YY'
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{2,4})", val)
    if m:
        mo, d, y = map(int, m.groups())
        if y < 100:  # '25' -> 2025
            y += 2000
        try:
            return datetime(y, mo, d)
        except Exception:
            pass

    # if we get here, we couldn't parse it
    return None


def _aggregate_rows_by_site(files_with_sites, kind: str, year: int, month: int) -> dict[str, int]:
    """files_with_sites: List[Tuple[path, site_name_or_None]]"""
    import csv
    from collections import defaultdict
    from pathlib import Path

    counts = defaultdict(int)
    for path, site_name in (files_with_sites or []):
        site = site_name or Path(path).stem
        try:
            with open(path, newline='', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    dt = _extract_row_datetime(row, kind)
                    if not dt:
                        continue
                    if dt.year == year and dt.month == month:
                        counts[site] += 1
        except Exception:
            # ignore unreadable files
            pass
    return dict(counts)

# ---------------- public API: messages ----------------


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


# ------------------ compat shims (safe to add once) ------------------

# If your file doesn't define UNIT_PRICE_SMS, default it here
try:
    UNIT_PRICE_SMS
except NameError:
    UNIT_PRICE_SMS = 0.14  # adjust if your SMS price is different

# If your aggregator is exported without a leading underscore, alias it
if '_aggregate_rows_by_site' not in globals() and 'aggregate_rows_by_site' in globals():
    def _aggregate_rows_by_site(files_with_sites, kind, year, month):
        return aggregate_rows_by_site(files_with_sites, kind, year, month)

# If your add/recompute helpers are exported without underscores, alias them
if '_add_item' not in globals() and 'add_item' in globals():
    def _add_item(*a, **k):
        return add_item(*a, **k)

if '_recompute_totals' not in globals() and 'recompute_totals' in globals():
    def _recompute_totals(*a, **k):
        return recompute_totals(*a, **k)

# ------------------ phone-resolution helpers ------------------

def _normalize_site_key(s: str) -> str:
    """Uppercase, trim VOICE/SMS suffix, and take left side of en dash."""
    u = (s or '').upper().strip()
    for suf in (' VOICE', ' SMS'):
        if u.endswith(suf):
            u = u[:-len(suf)].strip()
    if '–' in u:
        u = u.split('–', 1)[0].strip()
    # collapse internal whitespace
    u = ' '.join(u.split())
    return u

def _build_priority_phone_map(inv: dict) -> dict[str, str]:
    """
    Merge phones from inv['site_phones'] (UI-harvested) with clients.json,
    returning NAME(normalized)->'last4'.
    """
    phones: dict[str, str] = {}

    # 1) UI-provided site phones from the match column
    sp = inv.get('site_phones') or {}
    if isinstance(sp, dict):
        for name, last4 in sp.items():
            if last4 and str(last4).isdigit() and len(str(last4)) == 4:
                phones[_normalize_site_key(name)] = str(last4)

    # 2) Clients.json (if your loader exists)
    load_clients = (globals().get('_load_clients_doc') or
                    globals().get('load_clients_doc'))
    clients_doc = None
    if callable(load_clients):
        try:
            # Some codebases take an optional path arg; others take none
            clients_doc = load_clients(inv.get('clients_path'))
        except TypeError:
            clients_doc = load_clients()
    if isinstance(clients_doc, dict):
        for c in (clients_doc.get('clients') or []):
            for s in (c.get('sites') or []):
                nm = (s.get('name') or '').strip()
                ph = (s.get('phone') or '').strip()
                if nm and ph.isdigit() and len(ph) == 4:
                    phones.setdefault(_normalize_site_key(nm), ph)

    return phones

def _lookup_last4(phones: dict[str, str], desc: str) -> str | None:
    """Resolve the 4-digit phone for a given line-item description."""
    if not desc:
        return None
    key = _normalize_site_key(desc)
    return phones.get(key)




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
# ---------- starting invoice number (persisted in settings) ----------

def get_starting_invoice_number(default: int | None = None) -> int | None:
    s = _load_settings()
    val = s.get("starting_invoice_number", None)
    if isinstance(val, int):
        return val
    # coerce if stored as string
    try:
        return int(val)
    except Exception:
        return default

def set_starting_invoice_number(n: int | None) -> None:
    s = _load_settings()
    if n is None:
        # remove it if you want “blank”
        s.pop("starting_invoice_number", None)
    else:
        s["starting_invoice_number"] = int(n)
    _save_settings(s)

# ---------- Voice invoice helpers (row-count × fixed unit price) ----------

UNIT_PRICE_VOICE = 0.14  # USD per call (flat), per user spec

def _normalize_headers(headers: list[str]) -> list[str]:
    return [re.sub(r"[\s_\-]+", "", (h or "").strip().lower()) for h in headers]

def count_rows_calls_csv(path: str | Path, filter_year: int | None = None, filter_month: int | None = None) -> int:
    """
    Count rows in a Twilio *calls* CSV. If filter_year/month provided,
    only count rows whose date cell is within that (year, month).
    """
    p = Path(path)
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        headers = next(reader, [])
        if filter_year is None or filter_month is None:
            # Fast path: count all data rows
            return sum(1 for _ in reader)
        # Filtered path: locate date column appropriate for 'calls'
        ci = _date_col_index(headers, "calls")
        if ci is None:
            # If we can't find a date column, safest is count zero for filtered mode
            return 0
        n = 0
        for row in reader:
            cell = row[ci] if ci < len(row) else ""
            y, m = _ym_from_cell(cell)
            if y == filter_year and m == filter_month:
                n += 1
        return n

def build_voice_line_item(site_name: str | None, qty: int, unit_price: float = UNIT_PRICE_VOICE) -> dict:
    """
    Build a single line item dict (not yet added to invoice).
    Description style: '{Site Name} Voice' if site provided, else 'Voice'.
    If the site name already ends with VOICE, don't double it.
    """
    base = (site_name or "").strip()
    if base.upper().endswith("VOICE"):
        desc = base
    else:
        desc = f"{base} Voice" if base else "Voice"

    return {
        "description": desc,
        "qty": float(qty),
        "unit_price": float(unit_price),
    }



from pathlib import Path as _Path

def aggregate_voice_items_from_csvs(files_with_sites, year=None, month=None):
    by_site: dict[str, int] = {}
    for csv_path, site_name in files_with_sites:
        qty = count_rows_calls_csv(csv_path, year, month) if (year and month) else count_rows_calls_csv(csv_path)
        label = site_name or _Path(csv_path).stem      # <--- use filename if no site
        by_site[label] = by_site.get(label, 0) + int(qty)

    items: list[dict] = []
    for site_key, qty in sorted(by_site.items(), key=lambda kv: (kv[0] or "",)):
        items.append(build_voice_line_item(site_key or None, qty))
    return items



def _normalize_site_key(name: str) -> str:
    import re as _re
    if not name: return ""
    u = (name or "").upper().strip()
    u = u.replace("—", "-").replace("–", "-")
    u = _re.sub(r"\s+", " ", u)
    return u

def _build_priority_phone_map(inv: dict) -> dict[str,str]:
    phones: dict[str,str] = {}
    for k, v in (inv.get("site_phones") or {}).items():
        sv = str(v)
        if sv.isdigit() and len(sv) == 4:
            phones[_normalize_site_key(k)] = sv
    try:
        import json
        here = Path(__file__).resolve().parent
        cpath = here / "data" / "clients.json"
        if cpath.exists():
            data = json.loads(cpath.read_text(encoding="utf-8"))
            for c in data.get("clients", []):
                for d in (c.get("divisions") or []):
                    for s in (d.get("sites") or []):
                        n = (s.get("name") or "").strip()
                        ph = (s.get("phone") or "").strip()
                        if n and ph.isdigit() and len(ph) == 4:
                            phones.setdefault(_normalize_site_key(n), ph)
    except Exception:
        pass
    return phones

def _lookup_last4(phones: dict[str,str], desc: str) -> str | None:
    if not desc: return None
    key = _normalize_site_key(desc)
    # direct
    if key in phones:
        return phones[key]
    # try trimming VOICE/SMS
    for suf in (" VOICE"," SMS"):
        if key.endswith(suf):
            k = key[:-len(suf)].strip()
            if k in phones:
                return phones[k]
            key = k
            break
    # left side of dash
    if " - " in key:
        k = key.split(" - ",1)[0].strip()
        if k in phones:
            return phones[k]
    return None


def _normalize_site_key(name: str) -> str:
    import re as _re
    if not name: return ""
    u = (name or "").upper().strip()
    u = u.replace("—", "-").replace("–", "-")
    u = _re.sub(r"\s+", " ", u)
    return u

def _build_priority_phone_map(inv: dict) -> dict[str,str]:
    phones: dict[str,str] = {}
    for k, v in (inv.get("site_phones") or {}).items():
        sv = str(v)
        if sv.isdigit() and len(sv) == 4:
            phones[_normalize_site_key(k)] = sv
    try:
        import json
        here = Path(__file__).resolve().parent
        cpath = here / "data" / "clients.json"
        if cpath.exists():
            data = json.loads(cpath.read_text(encoding="utf-8"))
            for c in data.get("clients", []):
                for d in (c.get("divisions") or []):
                    for s in (d.get("sites") or []):
                        n = (s.get("name") or "").strip()
                        ph = (s.get("phone") or "").strip()
                        if n and ph.isdigit() and len(ph) == 4:
                            phones.setdefault(_normalize_site_key(n), ph)
    except Exception:
        pass
    return phones

def _lookup_last4(phones: dict[str,str], desc: str) -> str | None:
    if not desc: return None
    key = _normalize_site_key(desc)
    # direct
    if key in phones:
        return phones[key]
    # try trimming VOICE/SMS
    for suf in (" VOICE"," SMS"):
        if key.endswith(suf):
            k = key[:-len(suf)].strip()
            if k in phones:
                return phones[k]
            key = k
            break
    # left side of dash
    if " - " in key:
        k = key.split(" - ",1)[0].strip()
        if k in phones:
            return phones[k]
    return None
def add_voice_items_to_invoice(inv: Dict[str, Any],
                               files_with_sites: list[tuple[str | Path, str | None]],
                               year: int | None = None,
                               month: int | None = None,
                               unit_price: float = UNIT_PRICE_VOICE) -> Dict[str, Any]:
    """
    High-level helper: aggregate voice items from CSVs and append them to `inv`.
    Decorates each description with (-LAST4) when we can infer a phone for the site.
    Returns the modified invoice (same object).
    """
    items = aggregate_voice_items_from_csvs(files_with_sites, year, month)
    phones = _build_priority_phone_map(inv)
    for it in items:
        desc = str(it.get("description", "")).strip()
        # strip VOICE/SMS decoration that came from the file-type
        for suf in (" – VOICE"," – SMS"," VOICE"," SMS"):
            if desc.endswith(suf):
                desc = desc[:-len(suf)].strip()
        last4 = _lookup_last4(phones, desc)
        if last4:
            desc = f"{desc} (-{last4})"
        add_line_item(inv, desc, it.get("qty", 0), unit_price)
    return inv



def invoice_filename(inv: Dict[str, Any], ext: str) -> str:
    """Generate a human-friendly filename like Invoice-YYYY-MM-<id>.<ext>"""
    per = inv.get("period") or {}
    y = per.get("year")
    m = per.get("month")
    ym = f"{y}-{int(m):02d}" if y and m else "unknown"
    return f"Invoice-{ym}-{inv.get('id','')}.{ext.lstrip('.')}"

def export_invoice_csv(inv: Dict[str, Any], out_dir: str | Path | None = None) -> Path:
    """
    Export an invoice (from dict) to CSV with columns:
    Description,Qty,Unit Price,Amount,Subtotal,Tax,Total
    Saves to the remembered invoice_output_dir() if out_dir not provided.
    Returns the CSV path.
    """
    _recalc_totals(inv)
    if out_dir is None:
        out_dir = invoice_output_dir() or INVOICES_DIR
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / invoice_filename(inv, "csv")

    import io, csv as _csv
    buf = io.StringIO()
    w = _csv.writer(buf, lineterminator="\\n")
    w.writerow(["Description", "Qty", "Unit Price", "Amount"])
    for li in inv.get("line_items", []):
        w.writerow([li.get("description",""), li.get("qty",0), li.get("unit_price",0), li.get("amount",0)])
    w.writerow([])
    totals = inv.get("totals", {})
    w.writerow(["Subtotal", "", "", totals.get("subtotal", 0)])
    w.writerow(["Tax", "", "", totals.get("tax", 0)])
    w.writerow(["Total", "", "", totals.get("total", 0)])

    csv_path.write_text(buf.getvalue(), encoding="utf-8")
    return csv_path

# ---------- Simple PDF export (ReportLab) ----------

def export_invoice_pdf(inv: Dict[str, Any], out_dir: str | Path | None = None) -> Path:
    """
    Export the invoice as a simple PDF using reportlab (pip install reportlab).
    Layout: header, period, table (Description, Qty, Unit Price, Amount), totals.
    Saves into the remembered invoice folder if out_dir is None.
    Returns the PDF path.
    """
    try:
        from reportlab.lib.pagesizes import LETTER
        from reportlab.pdfgen import canvas
        from reportlab.lib.units import inch
    except Exception as e:
        raise RuntimeError("reportlab is required: pip install reportlab") from e

    _recalc_totals(inv)
    if out_dir is None:
        out_dir = invoice_output_dir() or INVOICES_DIR
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_path = out_dir / invoice_filename(inv, "pdf")

    # --- basic layout ---
    c = canvas.Canvas(str(pdf_path), pagesize=LETTER)
    width, height = LETTER
    x_margin = 0.75 * inch
    y = height - 0.75 * inch

    # Header
    c.setFont("Helvetica-Bold", 16)
    c.drawString(x_margin, y, "INVOICE")
    y -= 18

    per = inv.get("period") or {}
    period_text = f"Period: {per.get('year','')} - {int(per.get('month',0)):02d}" if per else ""
    c.setFont("Helvetica", 10)
    if period_text:
        c.drawString(x_margin, y, period_text)
        y -= 14
    inv_id = inv.get("id", "")
    c.drawString(x_margin, y, f"Invoice ID: {inv_id}")
    y -= 20

    # Client snapshot if present
    snap = inv.get("client_name_snapshot", "")
    if snap:
        c.drawString(x_margin, y, f"Client: {snap}")
        y -= 16

    # Table header
    c.setFont("Helvetica-Bold", 10)
    col_desc_x = x_margin
    col_qty_x  = x_margin + 4.6 * inch
    col_unit_x = x_margin + 5.4 * inch
    col_amt_x  = x_margin + 6.3 * inch
    c.drawString(col_desc_x, y, "Description")
    c.drawString(col_qty_x,  y, "Qty")
    c.drawString(col_unit_x, y, "Unit Price")
    c.drawString(col_amt_x,  y, "Amount")
    y -= 12
    c.line(x_margin, y, width - x_margin, y)
    y -= 8

    c.setFont("Helvetica", 10)
    for li in inv.get("line_items", []):
        if y < 1.3 * inch:
            c.showPage()
            y = height - 0.75 * inch
            c.setFont("Helvetica-Bold", 10)
            c.drawString(col_desc_x, y, "Description")
            c.drawString(col_qty_x,  y, "Qty")
            c.drawString(col_unit_x, y, "Unit Price")
            c.drawString(col_amt_x,  y, "Amount")
            y -= 12
            c.line(x_margin, y, width - x_margin, y)
            y -= 8
            c.setFont("Helvetica", 10)

        raw_desc = str(li.get("description", ""))
        # 🔹 Try to append (-LAST4) using your mapping helper
        try:
            desc = decorate_with_last4_kind(inv, raw_desc)
        except Exception:
            desc = raw_desc
        desc = desc[:80]

        # 🔹 FIXED: show whole quantity without stripping zeros
        qty_val = li.get("qty", 0)
        try:
            qty = str(int(round(float(qty_val))))
        except Exception:
            qty = str(qty_val)

        unit = f"{li.get('unit_price', 0):.2f}"
        amt  = f"{li.get('amount', 0):.2f}"

        c.drawString(col_desc_x, y, desc)
        c.drawRightString(col_qty_x + 0.5 * inch, y, qty)
        c.drawRightString(col_unit_x + 0.8 * inch, y, unit)
        c.drawRightString(col_amt_x + 0.8 * inch, y, amt)
        y -= 14

    y -= 6
    c.line(x_margin, y, width - x_margin, y)
    y -= 12

    totals = inv.get("totals", {})
    c.setFont("Helvetica-Bold", 10)
    c.drawRightString(col_unit_x + 0.8 * inch, y, "Subtotal:")
    c.drawRightString(col_amt_x + 0.8 * inch, y, f"{totals.get('subtotal', 0):.2f}")
    y -= 14
    c.drawRightString(col_unit_x + 0.8 * inch, y, "Tax:")
    c.drawRightString(col_amt_x + 0.8 * inch, y, f"{totals.get('tax', 0):.2f}")
    y -= 14
    c.drawRightString(col_unit_x + 0.8 * inch, y, "Total:")
    c.drawRightString(col_amt_x + 0.8 * inch, y, f"{totals.get('total', 0):.2f}")

    c.showPage()
    c.save()
    return pdf_path


# ---------- Excel template → PDF export (Windows, requires Excel) ----------
def _load_clients_doc(path: str | Path | None) -> dict:
    try:
        p = Path(path) if path else DATA_DIR / "clients.json"
        if p.exists():
            return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        pass
    return {}

def _find_client_address(clients_doc: dict, name_snapshot: str | None) -> list[str]:
    """Return up to 3 lines for the client's billing block (name + address)."""
    if not name_snapshot:
        return []
    # Try exact name match at top-level clients
    items = clients_doc.get("clients") or clients_doc.get("items") or []
    for c in items:
        if (c.get("name") or "").strip() == name_snapshot.strip():
            addr = (c.get("address") or "").strip()
            lines = [name_snapshot]
            if addr:
                for line in addr.splitlines():
                    if line.strip():
                        lines.append(line.strip())
            return lines[:3]
    # Fallback: just the snapshot name
    return [name_snapshot]

# === BEGIN: exporter decoration helpers (idempotent) ===
import re as _re_dec

def _infer_kind_and_base(desc: str) -> tuple[str|None, str]:
    """Strip trailing '— Voice/— SMS' variants, return (kind, base)."""
    if not isinstance(desc, str):
        return (None, "")
    s = desc.strip()
    # Common trailing junk variants
    tails = (
        " — Voice"," — VOICE"," – Voice"," – VOICE"," - Voice"," - VOICE",
        " — Sms"," — SMS"," – Sms"," – SMS"," - Sms"," - SMS",
    )
    changed = True
    while changed:
        changed = False
        for suf in tails:
            if s.endswith(suf):
                s = s[:-len(suf)].rstrip()
                changed = True
    up = s.upper()
    if up.endswith(" SMS"):
        return "SMS", s[:-3].rstrip()
    if up.endswith(" VOICE"):
        return "VOICE", s[:-5].rstrip()
    return None, s

def _phones_map_from_inv(inv: dict) -> dict[str, str]:
    """Combine your _build_priority_phone_map and inv['site_phones']; normalize to last-4."""
    phones: dict[str, str] = {}
    try:
        bpm = globals().get("_build_priority_phone_map")
        if callable(bpm):
            pm = bpm(inv) or {}
            for k, v in pm.items():
                if k:
                    phones[k] = str(v or "")[-4:]
    except Exception:
        pass
    if not phones and isinstance(inv.get("site_phones"), dict):
        for k, v in inv["site_phones"].items():
            if k and v:
                phones[k] = str(v)[-4:]
    return phones

def decorate_with_last4_kind(inv: dict, desc: str) -> str:
    """
    Return description with (-LAST4) when resolvable.

    Strategy:
      - Parse out a KIND ('VOICE'/'SMS') and a base site name.
      - If the description already ends with '(-dddd)', leave it alone.
      - If the base is just 'Voice' or 'SMS' (no site name), DO NOT decorate.
      - Otherwise try:
          1) exact match on '<base> KIND'
          2) exact match on '<base>'
          3) longest substring match against known site names
    """
    # nothing to do if it's not a real string
    if not isinstance(desc, str) or not desc.strip():
        return desc

    # already has (-1234) at the end
    if _re_dec.search(r"\(-\d{3,4}\)\s*$", desc):
        return desc

    kind, base = _infer_kind_and_base(desc)

    # no base text at all
    if not base:
        return desc

    # 🚫 generic 'Voice' / 'SMS' only (no site name) should never be decorated
    #   e.g. "Voice", "SMS", "voice", "sms"
    if base.upper() in {"VOICE", "SMS"} and " " not in base.strip():
        return desc

    phones = _phones_map_from_inv(inv)

    # 1) exact base+kind, then base
    for key in ([f"{base} {kind}"] if kind else []) + [base]:
        last4 = phones.get(key)
        if last4:
            return f"{base} (-{str(last4)[-4:]})"

    # 2) longest substring match as a last resort
    target = (f"{base} {kind}" if kind else base).upper()
    best_k, best_len = None, -1
    for k in phones.keys():
        ku = str(k).upper()
        if ku in target or target in ku:
            if len(ku) > best_len:
                best_k, best_len = k, len(ku)

    if best_k and phones.get(best_k):
        return f"{base} (-{phones[best_k][-4:]})"

    # couldn't resolve a phone – return undecorated base text
    return base



# === END: exporter decoration helpers ===



from pathlib import Path
from typing import Any, Dict, List



from pathlib import Path
from typing import Any, Dict, List

def export_invoice_pdf_via_template(inv: Dict[str, Any],
                                    template_path: str | Path,
                                    out_dir: str | Path | None = None,
                                    clients_path: str | Path | None = None) -> Path:
    """
    Export invoice to Excel (and PDF via Excel COM) with:
      - Date = TODAY → H5 (and H7), m/d/yyyy
      - Writes line items from inv["line_items"] (fallback to inv["items"]) starting at row 13
        A=Description, F=Qty, G=Unit, H=Amount (=F*G)
      - Keeps VOICE/SMS visible and appends (-LAST4) where resolvable
      - Fills Bill-To A8..A10 (uses your helpers if available)
      - Preserves SUBTOTAL/TOTAL rows (fills only if blank)
    """
    import openpyxl
    from openpyxl import load_workbook
    from datetime import datetime

    tpl = Path(template_path)
    if not tpl.exists():
        raise FileNotFoundError(f"Template not found: {tpl}")

    if out_dir is None:
        try:
            out_dir = invoice_output_dir() or INVOICES_DIR
        except NameError:
            out_dir = Path("invoices_output")
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    wb = load_workbook(tpl, data_only=False, keep_vba=True)
    ws = wb.active

    # Header
    inv_num = inv.get("human_number") or inv.get("starting_invoice_number") or inv.get("id")
    try:
        ws["G5"].value = inv_num
    except Exception:
        pass

    # Date: today into H5 and H7
    today = datetime.today()
    for cell in ("H5", "H7"):
        try:
            ws[cell].value = today
            ws[cell].number_format = "m/d/yyyy"
        except Exception:
            pass

    # Terms
    try:
        if not ws["H8"].value or str(ws["H8"].value).strip() == "":
            ws["H8"].value = "Due on Receipt"
    except Exception:
        pass

    # Bill-To lines (A8..A10)
    def _bill_to_lines() -> List[str]:
        try:
            load_clients = globals().get("_load_clients_doc")
            find_addr = globals().get("_find_client_address")
            if callable(load_clients) and callable(find_addr):
                doc = load_clients(inv.get("clients_path") if clients_path is None else clients_path)
                lines = find_addr(doc, inv.get("client_name_snapshot"))
                if isinstance(lines, list) and lines:
                    return [str(x) for x in lines][:3]
        except Exception:
            pass
        client = inv.get("client_snapshot") or inv.get("client") or {}
        name = (client.get("name") or "").strip()
        addr = (client.get("address") or "").strip()
        out: List[str] = []
        if name: out.append(name)
        if addr:
            for line in addr.splitlines():
                line = line.strip()
                if line: out.append(line)
        return out[:3]

    try:
        lines = _bill_to_lines()
        for i in range(3):
            ws[f"A{8+i}"].value = lines[i] if i < len(lines) else None
    except Exception:
        pass

    # Items
    row = 13
    line_items = inv.get("line_items", [])
    if not isinstance(line_items, list) or not line_items:
        line_items = inv.get("items", []) or []

    for li in line_items:
        # A=Desc, F=Qty, G=Unit, H=Amount
        try:
            raw_desc = li.get("description", "")
            ws[f"A{row}"].value = decorate_with_last4_kind(inv, raw_desc)
        except Exception:
            ws[f"A{row}"].value = li.get("description", "")
        try:
            ws[f"F{row}"].value = float(li.get("qty", 0) or 0.0)
            ws[f"G{row}"].value = float(li.get("unit_price", 0) or 0.0)
            if ws[f"H{row}"].value in (None, ""):
                ws[f"H{row}"].value = f"=F{row}*G{row}"
        except Exception:
            pass
        row += 1

    last_item_row = max(13, row - 1)

    # Locate labels
    def _find_label_row(label: str) -> int | None:
        L = label.strip().upper()
        for r in range(13, 200):
            for c in range(1, 8):  # A..G
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip().upper() == L:
                    return r
        return None

    subtotal_row = _find_label_row("SUBTOTAL")
    total_row = _find_label_row("TOTAL")

    # Clear trailing items but NOT subtotal/total block
    start_clear = last_item_row + 1
    stop_clear = subtotal_row if subtotal_row else start_clear
    if subtotal_row and start_clear < subtotal_row:
        for r in range(start_clear, subtotal_row):
            for c in ("A","F","G","H"):
                try:
                    ws[f"{c}{r}"].value = None
                except Exception:
                    pass

    # Ensure formulas if blank
    subtotal_formula = f"=SUM(H13:H{last_item_row})"
    try:
        if subtotal_row:
            if ws[f"H{subtotal_row}"].value in (None, ""):
                ws[f"H{subtotal_row}"].value = subtotal_formula
        else:
            ws["H16"].value = subtotal_formula
            subtotal_row = 16
    except Exception:
        pass
    try:
        if total_row:
            if ws[f"H{total_row}"].value in (None, ""):
                ws[f"H{total_row}"].value = f"=H{subtotal_row}"
        else:
            ws["H17"].value = f"=H{subtotal_row}"
    except Exception:
        pass

    # Save & export
    xlsm_path = out_dir / invoice_filename(inv, "xlsm")
    wb.save(xlsm_path)

    pdf_path = out_dir / invoice_filename(inv, "pdf")
    try:
        import win32com.client  # type: ignore
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        wb_com = excel.Workbooks.Open(str(xlsm_path))
        xlTypePDF = 0
        wb_com.ExportAsFixedFormat(xlTypePDF, str(pdf_path))
        wb_com.Close(False)
        excel.Quit()
        return pdf_path
    except Exception as e:
        raise RuntimeError(f"Excel export failed (install Excel + pywin32). Filled workbook at: {xlsm_path}") from e




# === BEGIN: message billing by segments (idempotent) ===
from typing import Any, Dict, Iterable, List, Tuple
import csv
from pathlib import Path as _Path_seg

def _ceil_div2(n: int) -> int:
    """1–2→1, 3–4→2, 5–6→3, 7–8→4, ..."""
    try:
        x = int(float(n))
    except Exception:
        return 1
    if x <= 0:
        return 0
    return (x + 1) // 2

def _extract_num_segments(row: Dict[str, Any]) -> int:
    """Robustly read NumSegments (handles common header variants)."""
    candidates = ("NumSegments", "Numsegments", "numsegments",
                  "Num_Segments", "NumSeg", "Numseg", "Segments", "segments")
    for k in candidates:
        if k in row and row[k] not in (None, "", "-"):
            try:
                return int(float(row[k]))
            except Exception:
                continue
    return 1

def _sum_billed_units_by_site(files_with_sites: List[Tuple[str | _Path_seg, str | None]],
                              year: int, month: int) -> Dict[str, int]:
    """
    Open each messages CSV and compute billed units per site:
      billed_units(row) = ceil(NumSegments/2) for rows in (year, month).
    Returns {site_name: total_billed_units}.
    """
    from collections import defaultdict
    totals: Dict[str, int] = defaultdict(int)

    for path, site_name in (files_with_sites or []):
        site = site_name or _Path_seg(path).stem
        try:
            with open(path, newline='', encoding='utf-8-sig') as f:
                reader = csv.DictReader(f)
                for row in reader:
                    # Use your existing helper to parse row timestamp for messages
                    try:
                        dt = _extract_row_datetime(row, "messages")
                    except Exception:
                        dt = None
                    if not dt or dt.year != int(year) or dt.month != int(month):
                        continue
                    seg = _extract_num_segments(row)
                    totals[site] += _ceil_div2(seg)
        except Exception:
            # swallow unreadable files silently
            pass
    return dict(totals)

def add_message_items_to_invoice(
    inv: dict,
    messages_with_sites: List[Tuple[str | _Path_seg, str | None]],
    year: int,
    month: int,
    unit_price: float = 0.14,  # per your new rule
) -> None:
    """
    Add one SMS line per site with:
      - description: "<Site Name> SMS" (export adds (-LAST4) while keeping SMS label)
      - qty: sum( ceil(NumSegments/2) ) for rows in (year, month)
      - unit_price: default 0.14
    """
    billed = _sum_billed_units_by_site(messages_with_sites, year, month)

    for site, qty in sorted(billed.items()):
        if qty <= 0:
            continue
        base = (site or "").strip()
        # 🔹 avoid double "SMS SMS"
        if base.upper().endswith("SMS"):
            desc = base
        else:
            desc = f"{base} SMS" if base else "SMS"

        try:
            _add_item(inv, desc, qty, unit_price)
        except NameError:
            inv.setdefault("line_items", []).append({
                "description": desc,
                "qty": qty,
                "unit_price": unit_price,
            })

    # Recompute totals in-memory (Excel still computes H=F*G)
    try:
        _recalc_totals(inv)
    except Exception:
        try:
            _recompute_totals(inv)
        except Exception:
            pass

# === END: message billing by segments ===
from pathlib import Path
import csv

# --- Full CSV → PDF (all rows, wide table) --------------------
from pathlib import Path
import csv

def export_invoice_csv_full_pdf(csv_path: str | Path,
                                out_dir: str | Path | None = None) -> Path:
    """
    Take an invoice CSV (whatever finalize_with_template/export_invoice_csv wrote)
    and render ALL rows into a single PDF as one big table.

    - First row is treated as a header and repeats on each page.
    - PDF is landscape to fit many columns.
    - Output filename: <csv_stem>-full.pdf in the same folder by default.

    Returns the PDF path.
    """
    try:
        from reportlab.lib.pagesizes import letter, landscape
        from reportlab.platypus import SimpleDocTemplate, Table, TableStyle
        from reportlab.lib import colors
    except Exception as e:
        raise RuntimeError(
            "reportlab is required to generate the full CSV table PDF "
            "(pip install reportlab)"
        ) from e

    csv_path = Path(csv_path)

    if out_dir is None:
        out_dir = csv_path.parent
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    # Read entire CSV
    with csv_path.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.reader(f)
        rows = list(reader)

    if not rows:
        raise ValueError(f"CSV appears to be empty: {csv_path}")

    # Normalize row lengths so the table is rectangular
    max_cols = max(len(r) for r in rows)
    norm_rows = [r + [""] * (max_cols - len(r)) for r in rows]

    # Output PDF path: same name + '-full.pdf'
    pdf_path = out_dir / (csv_path.stem + "-full.pdf")

    doc = SimpleDocTemplate(
        str(pdf_path),
        pagesize=landscape(letter),
        leftMargin=24, rightMargin=24,
        topMargin=24, bottomMargin=24,
    )

    table = Table(norm_rows, repeatRows=1)
    table.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
        ("FONTSIZE", (0, 0), (-1, 0), 9),
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.25, colors.grey),
        ("ALIGN", (1, 1), (-1, -1), "LEFT"),
    ]))

    doc.build([table])
    return pdf_path
