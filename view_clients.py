# view_clients.py
from __future__ import annotations

from pathlib import Path
from typing import Any, Dict, List
import json
import os
import uuid

DATA_PATH = Path(__file__).resolve().parent / "data" / "clients.json"


# --------- low-level IO ---------

def _ensure_file() -> None:
    DATA_PATH.parent.mkdir(parents=True, exist_ok=True)
    if not DATA_PATH.exists():
        DATA_PATH.write_text('{"version": 2, "clients": []}\n', encoding="utf-8")


def _atomic_write_text(path: Path, text: str) -> None:
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)


def new_id() -> str:
    return str(uuid.uuid4())


# --------- migration to 3-level model ---------

def _migrate_if_needed(doc: Dict[str, Any]) -> Dict[str, Any]:
    """
    Migrate legacy shape:
      client = {"name": ..., "suborgs":[{"name":..,"phone":..}]}
    into new shape:
      client = {"name": ..., "address":"", "divisions":[{"name":.., "sites":[{"name":..,"phone":..}]}]}
    """
    version = int(doc.get("version", 1))
    if version >= 2:
        return doc

    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        doc["clients"] = []
        doc["version"] = 2
        return doc

    for c in clients:
        if not isinstance(c, dict):
            continue
        # Already migrated?
        if "divisions" in c:
            continue

        # Legacy?
        suborgs = c.get("suborgs", [])
        if isinstance(suborgs, list):
            # Wrap legacy suborgs into a single division called "General"
            sites: List[Dict[str, Any]] = []
            for s in suborgs:
                if isinstance(s, dict):
                    sites.append({
                        "id": new_id(),
                        "name": s.get("name", ""),
                        "phone": s.get("phone", ""),
                    })
            c["divisions"] = [{
                "id": new_id(),
                "name": "General",
                "sites": sites,
            }]
            if "suborgs" in c:
                del c["suborgs"]

        # add address field if missing
        c.setdefault("address", "")

    doc["version"] = 2
    return doc


def load_clients() -> Dict[str, Any]:
    _ensure_file()
    try:
        doc = json.loads(DATA_PATH.read_text(encoding="utf-8"))
    except Exception:
        return {"version": 2, "clients": []}
    doc = _migrate_if_needed(doc)
    # persist migration if we upgraded
    if int(doc.get("version", 2)) == 2:
        _atomic_write_text(DATA_PATH, json.dumps(doc, indent=2, ensure_ascii=False) + "\n")
    return doc


def save_clients(doc: Dict[str, Any]) -> None:
    doc.setdefault("version", 2)
    _atomic_write_text(DATA_PATH, json.dumps(doc, indent=2, ensure_ascii=False) + "\n")


# --------- Client (top level) ---------

def list_clients() -> List[Dict[str, Any]]:
    doc = load_clients()
    clients = doc.get("clients", [])
    return clients if isinstance(clients, list) else []


def add_client(name: str, address: str = "") -> Dict[str, Any]:
    doc = load_clients()
    clients = doc.get("clients")
    if not isinstance(clients, list):
        clients = []
        doc["clients"] = clients

    client = {
        "id": new_id(),
        "name": name.strip(),
        "address": address.strip(),
        "divisions": [],          # list of {"id","name","sites":[ ... ]}
    }
    clients.append(client)
    save_clients(doc)
    return client


def find_client(client_id: str) -> Dict[str, Any] | None:
    for c in list_clients():
        if isinstance(c, dict) and c.get("id") == client_id:
            return c
    return None


def update_client(client_id: str, *, name: str | None = None, address: str | None = None) -> bool:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return False

    for c in clients:
        if isinstance(c, dict) and c.get("id") == client_id:
            changed = False
            if name is not None:
                c["name"] = name.strip()
                changed = True
            if address is not None:
                c["address"] = address.strip()
                changed = True
            if changed:
                save_clients(doc)
            return changed
    return False


def delete_client(client_id: str) -> bool:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return False
    for c in list(clients):
        if isinstance(c, dict) and c.get("id") == client_id:
            clients.remove(c)
            save_clients(doc)
            return True
    return False


# --------- Division (middle level) ---------

def add_division(client_id: str, name: str) -> Dict[str, Any] | None:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return None

    target = None
    for c in clients:
        if isinstance(c, dict) and c.get("id") == client_id:
            target = c
            break
    if target is None:
        return None

    divisions = target.get("divisions")
    if not isinstance(divisions, list):
        divisions = []
        target["divisions"] = divisions

    div = {"id": new_id(), "name": name.strip(), "sites": []}
    divisions.append(div)
    save_clients(doc)
    return div


def update_division(client_id: str, division_id: str, *, name: str | None = None) -> bool:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return False

    for c in clients:
        if not (isinstance(c, dict) and c.get("id") == client_id):
            continue
        divisions = c.get("divisions", [])
        if not isinstance(divisions, list):
            return False
        for d in divisions:
            if isinstance(d, dict) and d.get("id") == division_id:
                if name is None:
                    return False
                d["name"] = name.strip()
                save_clients(doc)
                return True
    return False


def delete_division(client_id: str, division_id: str) -> bool:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return False

    for c in clients:
        if not (isinstance(c, dict) and c.get("id") == client_id):
            continue
        divisions = c.get("divisions", [])
        if not isinstance(divisions, list):
            return False
        for d in list(divisions):
            if isinstance(d, dict) and d.get("id") == division_id:
                divisions.remove(d)
                save_clients(doc)
                return True
    return False


# --------- Site (bottom level, has phone) ---------

def add_site(client_id: str, division_id: str, name: str, phone: str = "") -> Dict[str, Any] | None:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return None

    for c in clients:
        if not (isinstance(c, dict) and c.get("id") == client_id):
            continue
        divisions = c.get("divisions", [])
        if not isinstance(divisions, list):
            return None
        for d in divisions:
            if isinstance(d, dict) and d.get("id") == division_id:
                sites = d.get("sites")
                if not isinstance(sites, list):
                    sites = []
                    d["sites"] = sites
                site = {"id": new_id(), "name": name.strip(), "phone": phone.strip()}
                sites.append(site)
                save_clients(doc)
                return site
    return None


def update_site(client_id: str, division_id: str, site_id: str,
                *, name: str | None = None, phone: str | None = None) -> bool:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return False

    for c in clients:
        if not (isinstance(c, dict) and c.get("id") == client_id):
            continue
        divisions = c.get("divisions", [])
        if not isinstance(divisions, list):
            return False
        for d in divisions:
            if not (isinstance(d, dict) and d.get("id") == division_id):
                continue
            sites = d.get("sites", [])
            if not isinstance(sites, list):
                return False
            for s in sites:
                if isinstance(s, dict) and s.get("id") == site_id:
                    changed = False
                    if name is not None:
                        s["name"] = name.strip()
                        changed = True
                    if phone is not None:
                        s["phone"] = phone.strip()
                        changed = True
                    if changed:
                        save_clients(doc)
                    return changed
    return False


def delete_site(client_id: str, division_id: str, site_id: str) -> bool:
    doc = load_clients()
    clients = doc.get("clients", [])
    if not isinstance(clients, list):
        return False

    for c in clients:
        if not (isinstance(c, dict) and c.get("id") == client_id):
            continue
        divisions = c.get("divisions", [])
        if not isinstance(divisions, list):
            return False
        for d in divisions:
            if not (isinstance(d, dict) and d.get("id") == division_id):
                continue
            sites = d.get("sites", [])
            if not isinstance(sites, list):
                return False
            for s in list(sites):
                if isinstance(s, dict) and s.get("id") == site_id:
                    sites.remove(s)
                    save_clients(doc)
                    return True
    return False


if __name__ == "__main__":
    _ensure_file()
    print("Path:", DATA_PATH)
    print("Exists?", DATA_PATH.exists())
    print("Contents:", load_clients())
