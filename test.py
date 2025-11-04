from pathlib import Path
from invoicing import identify_csv_and_phone, count_rows_calls_csv, _load_clients_doc

BASE_DIR = Path(r"C:\Users\HaydenP\Downloads\drive-download-20251103T211628Z-1-001")
YEAR = 2025
MONTH = 10

def main():
    clients_doc = _load_clients_doc(None)
    total_unknown = 0

    print(f"=== Per-CSV call counts for {YEAR}-{MONTH:02d} (with site) ===")
    for csv_path in sorted(BASE_DIR.glob("*.csv")):
        info = identify_csv_and_phone(csv_path, clients_doc)
        if info["kind"] != "calls":
            continue

        match = info.get("match") or {}
        site = match.get("site_name")
        qty = count_rows_calls_csv(csv_path, YEAR, MONTH)

        label = "UNMATCHED → goes into 'Voice'" if not site else ""
        print(f"{csv_path.name:60} site={repr(site):35} calls={qty:5} {label}")

        if not site:
            total_unknown += qty

    print()
    print("Total calls from UNMATCHED voice CSVs:", total_unknown)

if __name__ == "__main__":
    main()
