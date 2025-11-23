from pathlib import Path
import json
from invoicing import identify_csv_and_phone, _load_clients_doc

def main():
    clients_doc = _load_clients_doc()

    # point this to wherever your Twilio CSVs are
    csv_dir = Path(r"C:\Users\HaydenP\Downloads\drive-download-2025...")
    csvs = sorted(csv_dir.glob("*.csv"))

    for p in csvs:
        info = identify_csv_and_phone(p, clients_doc)
        m = info.get("match") or {}
        print(
            p.name,
            "kind=", info.get("kind"),
            "phone=", info.get("phone"),
            "site=", m.get("site_name"),
            "division=", m.get("division_name"),
            "site_phone=", m.get("site_phone"),
        )

if __name__ == "__main__":
    main()
