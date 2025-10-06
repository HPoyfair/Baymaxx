from __future__ import annotations
from pathlib import Path
import json
from typing import Dict, Any
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


def _atomic_write_text(path: Path, text:str) -> None:
    """writing to a tempfile before saving, this makes sure that it doesent corrupt the file if something goes wrong"""
    tmp = path.with_suffix(path.suffix + ".tmp") # .suffix returns the suffix -> json .with_suffix creates a new path with a different suffix, it does not modify the original path.suffix + ".tmp" is ".json.tmp" path.with_suffix(".json.tmp") becomes .../data/clients.json.tmp
    tmp.write_text(text, encoding="utf-8")
    os.replace(tmp, path)

def save_clients(doc: Dict[str, Any]) -> None:

    _ensure_file()
    text = json.dumps(doc, indent = 2, ensure_ascii = False) + "\n" # creating the text by dumping the current contents of the dict into text indenting it
    _atomic_write_text(DATA_PATH, text) # passing the path and the text to atomic write, the temp file






if __name__ == "__main__":
    _ensure_file()
    print("Path:", DATA_PATH)
    print("Exists?", DATA_PATH.exists())
    print("Contents", load_clients())