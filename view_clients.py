from __future__ import annotations
from pathlib import Path
import json
from typing import List, Dict, Any
import os
import tempfile
import uuid

def new_id() -> str:
    """returns a random  uuid"""
    return str(uuid.uuid4())


DATA_PATH = Path(__file__).resolve().parent / "data" / "clients.json" # the pathway for where the file should be





def _ensure_file() -> None:
    """makes sure the .json exisits and contains valid json"""

    DATA_PATH.parent.mkdir(parents=True, exist_ok= True) # this tests that the pathway has the actual json at the end

    if not DATA_PATH.exists():
        DATA_PATH.write_text('{"version": 1, "clients": []}\n', encoding ="utf-8") # this makes an empty json if at the end of the pathway nothing is there






def load_clients() -> Dict[str, Any]:
    """ read clients.json and return as a python dict if the file is missing or bad, return a safe empty structure"""
    try:
        text = DATA_PATH.read_text(encoding ="utf-8")
        return json.loads(text)
    except Exception:
        return{"version":1,"clients":[]}
    

def list_clients() -> List[Dict[str, Any]]:
    """ return the array of clientrs from the json doc
    if the key is missing or isent a list, return an empty list"""

    doc = load_clients() # load the whole json as a dict
    clients = doc.get("clients", []) # safly pull the clients key
    return clients if isinstance(clients, list) else []




def add_client(name: str) -> Dict[str, Any]:

    doc = load_clients()
    clients = doc.get("clients")

    if not isinstance(clients, list):
        clients = []
        doc["clients"] = clients

    client = {
        "id": new_id(),
        "name": name.strip(),
        "suborgs": []
    }
    clients.append(client)
    save_clients(doc)
    return client




def _atomic_write_text(path: Path, text:str) -> None:
    """writing to a tempfile before saving, this makes sure that it doesent corrupt the file if something goes wrong"""
    tmp = path.with_suffix(path.suffix + ".tmp") # .suffix returns the suffix -> json .with_suffix creates a new path with a different suffix, it does not modify the original path.suffix + ".tmp" is ".json.tmp" path.with_suffix(".json.tmp") becomes .../data/clients.json.tmp
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)

def save_clients(doc: Dict[str, Any]) -> None:

    _ensure_file()
    text = json.dumps(doc, indent = 2, ensure_ascii = False) + "\n" # creating the text by dumping the current contents of the dict into text indenting it
    _atomic_write_text(DATA_PATH, text) # passing the path and the text to atomic write, the temp file





def add_suborg(client_id: str, name: str, phone: str = "") -> Dict[str, Any] | None: # given the id name phone and return a dict or nothing
    """adding a suborg to the parent client"""
    doc = load_clients() #load the json into the dict
    clients = doc.get("clients") # getting the clients from current json

    if not isinstance(clients, list): # if clients isent a list make it into one and fill it with the clients from the json

        clients = []
        doc["clients"] = clients
    
    target = None

    for c in clients:
        if isinstance(c, dict) and c.get("id") == client_id: #checks to see if c is a dictionary, and whether the id of c is he same as the one we searching for
            target = c #this grabs that dict if it has the matching id
            break
    if target is None:
        return None
    

    suborgs = target.get("suborgs")
    if not isinstance(suborgs, list):
        suborgs= []
        target["suborgs"] = suborgs 

    suborg = {
        "id": new_id(),
        "name": name.strip(),
        "phone":phone.strip()
    }

    suborgs.append(suborg)
    save_clients(doc)

    return suborg




def find_client(client_id: str) -> Dict[str, Any] | None:
    doc = load_clients()
    clients = doc.get("clients", [])

    if not isinstance(clients, list):
        return None
    
    for c in clients:
        if isinstance(c, dict) and c.get("id") == client_id:
            return c
    return None












def update_client_name(client_id: str, new_name: str) -> bool:
    """Rename an existing client. Return True if updated, False otherwise."""

    doc = load_clients()
    clients = doc.get("clients", [])

    new_name = new_name.strip()

    if not new_name: # refuse to change an empty name
        return False

    for c in clients:
        if isinstance(c, dict) and  c.get("id") == client_id:
            c["name"] = new_name
            save_clients(doc)
            return True
    return False 

    



def delete_client(client_id: str) -> bool:
    """Remove a client by id. Return True if anything was deleted, False otherwise."""
    doc = load_clients()
    clients = doc.get("clients", [])
   

    for c in clients:
        if isinstance(c, dict) and c.get("id") == client_id:
            clients.remove(c)
            save_clients(doc)
        
        
    return False
    







def update_suborg(client_id: str, suborg_id: str, *,
                  name: str | None = None,
                  phone: str | None = None) -> bool:
    """Update fields on a specific suborg. Only update provided fields. Return True if updated."""
    raise NotImplementedError


def delete_suborg(client_id: str, suborg_id: str) -> bool:
    """Delete a suborg under a client by id. Return True if anything was removed."""
    raise NotImplementedError















if __name__ == "__main__":
    _ensure_file()
    print("Path:", DATA_PATH)
    print("Exists?", DATA_PATH.exists())
    print("Contents", load_clients())