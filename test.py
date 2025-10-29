from pathlib import Path
import invoicing as inv

inv_obj = inv.new_monthly_invoice(2025, 5)
inv_obj.setdefault("client", {})
inv_obj["client"]["name"] = 'Sante Foundation Medical Group ("SFMG")'
inv_obj["client"]["address"] = "7370 N Palm Ave STE 101\nFresno, CA 93711"
inv_obj["client"]["contact"] = "billing@client.com"
inv_obj["invoice_number"] = "24"
inv_obj["date"] = "9/16/2025"

inv.add_line_item(inv_obj, "CCHLS Voice (-9990)", 336, inv.UNIT_PRICE_VOICE)
inv.save_invoice(inv_obj)
inv.export_invoice_csv(inv_obj)

tpl = Path(__file__).parent / "invoice.xlsm"
print("Template exists:", tpl.exists(), tpl)

# IMPORTANT: call the V3 function:
paths = inv.export_invoice_with_excel_template_v3(
    inv_obj,
    template_path=str(tpl),
    out_dir=None,      # your usual invoices folder
    export_pdf=True    # Windows + Excel + pywin32 -> PDF; else .xlsm only
)
print(paths)
