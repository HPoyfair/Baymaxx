# view_clients.py

from __future__ import annotations  # <- must be the first import in the file

import json
import os
import tempfile
import uuid
from pathlib import Path
from typing import Any, Dict, List

# NOTE: this module is pure data/IO; no tkinter imports needed here.

DATA_PATH = Path(__file__).resolve().parent / "data" / "clients.json"


def _ensure_file() -> None:
    """Make sure the data folder/file exist, and that the file starts valid."""
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not DATA_PATH.exists():
        # Valid JSON! (your original string had mismatched braces)
        DATA_PATH.write_text('{"version": 1, "clients": []}\n', encoding="utf-8")


def load_clients() -> Dict[str, Any]:
    """Return the whole JSON document as a Python dict."""
    _ensure_file()
    try:
        with DATA_PATH.open("r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        # Corrupt/bad JSON: fail soft to an empty structure.
        return {"version": 1, "clients": []}


def _atomic_write_text(path: Path, text: str) -> None:
    """
    Write text to a temp file and replace into place to avoid partial writes.

    Use a temp file in the SAME directory so the replace is atomic on Windows.
    """
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def save_clients(doc: Dict[str, Any]) -> None:
    """Persist the whole JSON document prettily."""
    _ensure_file()
    text = json.dumps(doc, indent=2, ensure_ascii=False) + "\n"
    _atomic_write_text(DATA_PATH, text)


# ---------------------- CRUD helpers ----------------------

def new_id() -> str:
    return str(uuid.uuid4())


def list_clients() -> List[Dict[str, Any]]:
    return load_clients().get("clients", [])


def add_client(name: str, notes: str = "") -> Dict[str, Any]:
    doc = load_clients()
    client = {"id": new_id(), "name": name, "notes": notes, "offices": []}
    doc["clients"].append(client)
    save_clients(doc)
    return client


def find_client(client_id: str) -> Dict[str, Any] | None:
    for c in list_clients():
        if c["id"] == client_id:
            return c
    return None


def update_client(client_id: str, **fields) -> bool:
    doc = load_clients()
    for c in doc["clients"]:
        if c["id"] == client_id:
            c.update(fields)
            save_clients(doc)
            return True
    return False


def delete_client(client_id: str) -> bool:
    doc = load_clients()
    before = len(doc["clients"])
    doc["clients"] = [c for c in doc["clients"] if c["id"] != client_id]
    if len(doc["clients"]) != before:
        save_clients(doc)
        return True
    return False


# ---------------------- Offices (nested) ----------------------

def add_office(client_id: str, name: str,
               phone: str = "", email: str = "") -> Dict[str, Any] | None:
    doc = load_clients()
    for c in doc["clients"]:
        if c["id"] == client_id:
            office = {"id": new_id(), "name": name, "phone": phone, "email": email}
            c["offices"].append(office)
            save_clients(doc)
            return office
    return None


def update_office(client_id: str, office_id: str, **fields) -> bool:
    doc = load_clients()
    for c in doc["clients"]:
        if c["id"] == client_id:
            for o in c["offices"]:
                if o["id"] == office_id:
                    o.update(fields)
                    save_clients(doc)
                    return True
    return False


def delete_office(client_id: str, office_id: str) -> bool:
    doc = load_clients()
    for c in doc["clients"]:
        if c["id"] == client_id:
            before = len(c["offices"])
            c["offices"] = [o for o in c["offices"] if o["id"] != office_id]
            if len(c["offices"]) != before:
                save_clients(doc)
                return True
    return False

if __name__ == "__main__":
    from pprint import pprint

    # reset file for a clean test (optional)
    if DATA_PATH.exists():
        DATA_PATH.unlink()

    c1 = add_client("Community Medical Group", notes="Main account")
    add_office(c1["id"], "North Clinic", phone="555-111-2222")
    add_office(c1["id"], "West Clinic", phone="555-333-4444", email="west@cmg.com")

    c2 = add_client("Fresno Digestive Health", notes="GI")

    pprint(list_clients())
