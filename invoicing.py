# invoicing.py
from __future__ import annotations

from pathlib import Path
from typing import Dict, Any, List, Tuple
import json
import uuid
import os
import re
import csv


# ---------- paths ----------
DATA_DIR = Path(__file__).resolve().parent / "data"
INVOICES_DIR = DATA_DIR / "invoices"


# ---------- small utils ----------
def _ensure_dirs() -> None:
    """Make sure data/invoices exists."""
    INVOICES_DIR.mkdir(parents=True, exist_ok=True)


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
    """Persist the invoice as data/invoices/<id>.json (atomic write)."""
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
    # keep leading + if present, then digits
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
