# Standard library
from pathlib import Path
from datetime import datetime
import os
import sys
import csv

# Tkinter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Optional PIL (logo)
try:
    from PIL import Image, ImageTk
except Exception:
    Image = None
    ImageTk = None

# Local modules
import view_clients as clients


# --- Baymaxx: ensure we import the local invoicing.py, not a different site-package ---
import sys as _sys_local, os as _os_local, importlib as _importlib_local
_HERE = _os_local.path.dirname(__file__)
if _HERE not in _sys_local.path:
    _sys_local.path.insert(0, _HERE)

import invoicing as inv
_importlib_local.reload(inv)  # ensure latest file is used


CELL_MAP = {
    "start_row": 13,
    "desc_col": "A",
    "qty_col":  "F",   # push right
    "unit_col": "G",   # push right
    "amount_col": "H", # push right
    "subtotal_cell": "H16",
    "total_cell":    "H18",
    "date_cell":     "H7",  # DATE
    "terms_cell":    "H8",  # TERMS
    "invoice_no_cell":"G5",
    "billto_cell":   "A8",
}


# --- Fallback shim: if invoicing.finalize_with_template is missing, use classic pipeline ---
if not hasattr(inv, "finalize_with_template"):
    def _finalize_shim(inv_obj, template_path):
        internal = inv.save_invoice(inv_obj)
        try:
            csv_path = inv.export_invoice_csv(inv_obj)
        except Exception:
            csv_path = None
        pdf_path = None
        # Try Excel template -> PDF (Windows with Excel)
        if template_path and os.name == "nt":
            try:
                pdf_path = inv.export_invoice_pdf_via_template(inv_obj, template_path)
            except Exception:
                pdf_path = None
        if not pdf_path:
            try:
                pdf_path = inv.export_invoice_pdf(inv_obj)
            except Exception:
                pdf_path = None
        return {"json": str(internal), "csv": str(csv_path) if csv_path else None,
                "xlsm": None, "pdf": str(pdf_path) if pdf_path else None}

    inv.finalize_with_template = _finalize_shim  # patches module at runtime


# ---- Finalize shim: uses invoicing.finalize_with_template if present,
# ---- else reconstructs the same pipeline with existing functions.
def _finalize_shim(inv_module, inv_obj, template_path):
    fn = getattr(inv_module, "finalize_with_template", None)
    if fn:
        return fn(inv_obj, template_path)

    json_path = inv_module.save_invoice(inv_obj)
    try:
        csv_path = inv_module.export_invoice_csv(inv_obj)
    except Exception:
        csv_path = None

    xlsm_path = None
    pdf_path = None
    try:
        pdf_path = inv_module.export_invoice_pdf_via_template(
            inv_obj, template_path
        )
        try:
            out_dir = inv_module.invoice_output_dir()
        except Exception:
            from pathlib import Path as _P
            out_dir = _P(__file__).resolve().parent / "data" / "invoices"
        cand = out_dir / (inv_module.invoice_filename(inv_obj, "xlsm").replace(".xlsm","") + ".xlsm")
        if cand.exists():
            xlsm_path = cand
    except Exception:
        pass

    if not pdf_path:
        try:
            pdf_path = inv_module.export_invoice_pdf(inv_obj, template_path=template_path)
        except Exception:
            pdf_path = None

    return {"json": json_path, "csv": csv_path, "xlsm": xlsm_path, "pdf": pdf_path}

# --- Find the parent org (top-level client) from clients.json based on the sites in the invoice
def _infer_parent_from_clients(inv_obj, clients_path):
    import json, re
    try:
        with open(clients_path, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception:
        return None

    # Collect site names from line-item descriptions (before the "—")
    def _norm(s):
        s = (s or "").upper().strip()
        s = s.replace("—", "-").replace("–", "-")
        s = re.sub(r"\s+", " ", s)
        return s

    site_names = []
    for li in inv_obj.get("line_items", []):
        desc = (li.get("description") or "")
        left = desc.split("—", 1)[0].strip()
        u = _norm(left)
        # drop suffixes like " VOICE", " SMS"
        for suf in (" VOICE", " SMS", "- VOICE", "- SMS"):
            if u.endswith(suf):
                u = u[: -len(suf)].strip()
        site_names.append(u)

    # Look up which top-level client contains any of these sites
    for c in (data.get("clients") or []):
        for d in (c.get("divisions") or []):
            for s in (d.get("sites") or []):
                sn = _norm(s.get("name"))
                for target in site_names:
                    if sn == target or sn.startswith(target) or target in sn:
                        return c.get("name")
    return None


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Baymaxx — Invoicing Toolkit")
        self.geometry("760x520")

        # ---- grid: 2 cols, fixed left rail, flexible right content
        self.grid_columnconfigure(0, weight=0, minsize=180)   # sidebar width
        self.grid_columnconfigure(1, weight=1)                # content
        self.grid_rowconfigure(0, weight=0)                   # topbar
        self.grid_rowconfigure(1, weight=1)                   # main area

        # ---- top bar
        topbar = ttk.Frame(self, padding=(12, 8))
        topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(topbar, text="help").grid(row=0, column=0, sticky="w")

        # ---- left sidebar (buttons)
        sidebar = ttk.Frame(self, padding=(10, 10))
        sidebar.grid(row=1, column=0, sticky="nsw")
        ttk.Button(sidebar, text="New Month Invoice",
                   command=self.show_monthly_import).pack(fill="x", pady=6)
        ttk.Button(sidebar, text="View Past Invoices",
                   command=self.show_invoices).pack(fill="x", pady=6)
        ttk.Button(sidebar, text="New Individual Invoice").pack(fill="x", pady=6)
        ttk.Button(sidebar, text="View Clients",
                   command=self.open_clients_manager).pack(fill="x", pady=6)

        # ---- right content (logo centered)
        self.content = ttk.Frame(self, padding=16)
        self.content.grid(row=1, column=1, sticky="nsew")
        self.content.rowconfigure(0, weight=1)
        self.content.columnconfigure(0, weight=1)

        logo_path = Path(__file__).resolve().parent / "baymaxx.png"
        self.show_logo(logo_path)

        # Ensure an invoice root exists / is remembered
        self.after(0, lambda: inv.ensure_invoice_root(self))

    # ---------- Content swaps ----------
    def show_logo(self, logo_path: Path):
        for child in self.content.winfo_children():
            child.destroy()
        try:
            if Image and ImageTk and logo_path.exists():
                img = Image.open(logo_path)
                img.thumbnail((420, 420))
                self.logo_img = ImageTk.PhotoImage(img)
                ttk.Label(self.content, image=self.logo_img).grid(row=0, column=0)
            else:
                ttk.Label(self.content, text="Baymaxx", font=("", 28, "bold")).grid(row=0, column=0)
        except Exception:
            ttk.Label(self.content, text="Baymaxx", font=("", 28, "bold")).grid(row=0, column=0)

    def open_clients_manager(self):
        ClientsManager(self)

    def show_home(self) -> None:
        self.show_logo(Path(__file__).resolve().parent / "baymaxx.png")

    def show_monthly_import(self) -> None:
        for child in self.content.winfo_children():
            child.destroy()
        view = MonthlyImportView(self.content, on_back=self.show_home)
        view.grid(row=0, column=0, sticky="nsew")

    def show_invoices(self) -> None:
        for child in self.content.winfo_children():
            child.destroy()
        view = ViewInvoicesView(self.content, on_back=self.show_home)
        view.grid(row=0, column=0, sticky="nsew")


# ---------------- Clients Manager ----------------
class ClientsManager(tk.Toplevel):
    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Clients")
        self.geometry("720x420")
        self.transient(parent)
        self.grab_set()
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        cols = ("name", "address", "divisions")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        self.tree.heading("name", text="Name")
        self.tree.heading("address", text="Address")
        self.tree.heading("divisions", text="# Divisions")
        self.tree.column("name", width=320, anchor="w")
        self.tree.column("address", width=240, anchor="w")
        self.tree.column("divisions", width=100, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=0, column=1, sticky="ns", pady=(12, 6))
        self.tree.configure(yscrollcommand=ybar.set)

        btns = ttk.Frame(self, padding=(12, 6))
        btns.grid(row=1, column=0, columnspan=2, sticky="ew")
        for i in range(5):
            btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="Add", command=self.add_client).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Edit", command=self.edit_client).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_client).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="Divisions…", command=self.open_divisions).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(btns, text="Close", command=self.destroy).grid(row=0, column=4, sticky="ew", padx=4)

        self.tree.bind("<Double-1>", lambda e: self.edit_client())
        self.refresh()

    def selected_id(self) -> str | None:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for c in clients.list_clients():
            name = c.get("name", "")
            address = c.get("address", "")
            divs = c.get("divisions", [])
            count = len(divs) if isinstance(divs, list) else 0
            self.tree.insert("", tk.END, iid=c.get("id", ""), values=(name, address, count))

    def _client_dialog(self, title: str, init_name: str = "", init_address: str = "") -> tuple[str | None, str]:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(dlg, text="Name:").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        name_var = tk.StringVar(value=init_name)
        name_ent = ttk.Entry(dlg, textvariable=name_var, width=46)
        name_ent.grid(row=0, column=1, sticky="ew", padx=12, pady=(12, 4))

        ttk.Label(dlg, text="Address:").grid(row=1, column=0, sticky="w", padx=12, pady=(4, 8))
        addr_var = tk.StringVar(value=init_address)
        addr_ent = ttk.Entry(dlg, textvariable=addr_var, width=46)
        addr_ent.grid(row=1, column=1, sticky="ew", padx=12, pady=(4, 8))

        btns = ttk.Frame(dlg)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", padx=12, pady=(0, 12))
        ttk.Button(btns, text="OK", command=dlg.destroy).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="Cancel", command=lambda: (name_var.set("__CANCEL__"), dlg.destroy())).grid(row=0, column=1)

        name_ent.focus_set()
        dlg.wait_window(dlg)

        name = name_var.get().strip()
        if name == "__CANCEL__" or not name:
            return (None, "")
        return (name, addr_var.get().strip())

    # actions
    def add_client(self):
        name, addr = self._client_dialog("Add Client")
        if name is None:
            return
        clients.add_client(name, address=addr)
        self.refresh()

    def edit_client(self):
        cid = self.selected_id()
        if not cid:
            messagebox.showinfo("Edit Client", "Select a client first.")
            return
        c = clients.find_client(cid)
        if not c:
            messagebox.showerror("Edit Client", "Client not found.")
            return
        name, addr = self._client_dialog("Edit Client", init_name=c.get("name", ""), init_address=c.get("address", ""))
        if name is None:
            return
        if not clients.update_client(cid, name=name, address=addr):
            messagebox.showerror("Edit Client", "Update failed.")
        self.refresh()

    def delete_client(self):
        cid = self.selected_id()
        if not cid:
            messagebox.showinfo("Delete Client", "Select a client first.")
            return
        item = self.tree.item(cid)
        nm = item["values"][0] if item and item.get("values") else "(unnamed)"
        if not messagebox.askyesno("Delete Client", f"Delete '{nm}'?"):
            return
        if clients.delete_client(cid):
            self.refresh()
        else:
            messagebox.showerror("Delete Client", "Delete failed.")

    def open_divisions(self):
        cid = self.selected_id()
        if not cid:
            messagebox.showinfo("Divisions", "Select a client first.")
            return
        c = clients.find_client(cid)
        if not c:
            messagebox.showerror("Divisions", "Client not found.")
            return
        DivisionsManager(self, client_id=cid, client_name=c.get("name", "(unnamed)"))


# ---------------- Divisions Manager (middle) ----------------
class DivisionsManager(tk.Toplevel):
    def __init__(self, parent: tk.Toplevel, client_id: str, client_name: str):
        super().__init__(parent)
        self.client_id = client_id
        self.title(f"Divisions — {client_name}")
        self.geometry("640x400")
        self.transient(parent)
        self.grab_set()
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        cols = ("name", "sites")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        self.tree.heading("name", text="Name")
        self.tree.heading("sites", text="# Sites")
        self.tree.column("name", width=360, anchor="w")
        self.tree.column("sites", width=100, anchor="center")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=0, column=1, sticky="ns", pady=(12, 6))
        self.tree.configure(yscrollcommand=ybar.set)

        btns = ttk.Frame(self, padding=(12, 6))
        btns.grid(row=1, column=0, columnspan=2, sticky="ew")
        for i in range(5): 
            btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="Add", command=self.add_division).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Edit", command=self.edit_division).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_division).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="Sites…", command=self.open_sites).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(btns, text="Close", command=self.destroy).grid(row=0, column=4, sticky="ew", padx=4)

        self.tree.bind("<Double-1>", lambda e: self.edit_division())
        self.refresh()

    def selected_id(self) -> str | None:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        c = clients.find_client(self.client_id)
        if not c:
            return
        divisions = c.get("divisions", [])
        if not isinstance(divisions, list):
            return
        for d in divisions:
            n = d.get("name", "")
            sites = d.get("sites", [])
            cnt = len(sites) if isinstance(sites, list) else 0
            self.tree.insert("", tk.END, iid=d.get("id", ""), values=(n, cnt))

    def _name_dialog(self, title: str, init: str = "") -> str | None:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(dlg, text="Name:").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        var = tk.StringVar(value=init)
        ent = ttk.Entry(dlg, textvariable=var, width=40)
        ent.grid(row=0, column=1, sticky="ew", padx=12, pady=(12, 4))

        btns = ttk.Frame(dlg)
        btns.grid(row=1, column=0, columnspan=2, sticky="e", padx=12, pady=(0, 12))
        ttk.Button(btns, text="OK", command=dlg.destroy).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="Cancel", command=lambda: (var.set("__CANCEL__"), dlg.destroy())).grid(row=0, column=1)

        ent.focus_set()
        dlg.wait_window(dlg)
        name = var.get().strip()
        if name == "__CANCEL__" or not name:
            return None
        return name

    def add_division(self):
        name = self._name_dialog("Add Division")
        if name is None:
            return
        if not clients.add_division(self.client_id, name):
            messagebox.showerror("Add Division", "Add failed.")
        self.refresh()

    def edit_division(self):
        did = self.selected_id()
        if not did:
            messagebox.showinfo("Edit Division", "Select a division first.")
            return
        c = clients.find_client(self.client_id)
        if not c:
            messagebox.showerror("Edit Division", "Client not found.")
            return
        cur = None
        for d in c.get("divisions", []):
            if isinstance(d, dict) and d.get("id") == did:
                cur = d
                break
        if not cur:
            messagebox.showerror("Edit Division", "Division not found.")
            return
        name = self._name_dialog("Edit Division", init=cur.get("name", ""))
        if name is None:
            return
        if not clients.update_division(self.client_id, did, name=name):
            messagebox.showerror("Edit Division", "Update failed.")
        self.refresh()

    def delete_division(self):
        did = self.selected_id()
        if not did:
            messagebox.showinfo("Delete Division", "Select a division first.")
            return
        item = self.tree.item(did)
        nm = item["values"][0] if item and item.get("values") else "(unnamed)"
        if not messagebox.askyesno("Delete Division", f"Delete '{nm}'?"):
            return
        if not clients.delete_division(self.client_id, did):
            messagebox.showerror("Delete Division", "Delete failed.")
        self.refresh()

    def open_sites(self):
        did = self.selected_id()
        if not did:
            messagebox.showinfo("Sites", "Select a division first.")
            return
        c = clients.find_client(self.client_id)
        if not c:
            messagebox.showerror("Sites", "Client not found.")
            return
        dname = "(unnamed)"
        for d in c.get("divisions", []):
            if isinstance(d, dict) and d.get("id") == did:
                dname = d.get("name", "(unnamed)")
                break
        SitesManager(self, client_id=self.client_id, division_id=did, division_name=dname)


# ---------------- Sites Manager (bottom – has phone) ----------------
class SitesManager(tk.Toplevel):
    def __init__(self, parent: tk.Toplevel, client_id: str, division_id: str, division_name: str):
        super().__init__(parent)
        self.client_id = client_id
        self.division_id = division_id
        self.title(f"Sites — {division_name}")
        self.geometry("620x380")
        self.transient(parent)
        self.grab_set()
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        cols = ("name", "phone")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        self.tree.heading("name", text="Name")
        self.tree.heading("phone", text="Phone")
        self.tree.column("name", width=320, anchor="w")
        self.tree.column("phone", width=180, anchor="w")
        self.tree.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=0, column=1, sticky="ns", pady=(12, 6))
        self.tree.configure(yscrollcommand=ybar.set)

        btns = ttk.Frame(self, padding=(12, 6))
        btns.grid(row=1, column=0, columnspan=2, sticky="ew")
        for i in range(4):
            btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="Add", command=self.add_site).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Edit", command=self.edit_site).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_site).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="Close", command=self.destroy).grid(row=0, column=3, sticky="ew", padx=4)

        self.tree.bind("<Double-1>", lambda e: self.edit_site())
        self.refresh()

    def selected_id(self) -> str | None:
        sel = self.tree.selection()
        return sel[0] if sel else None

    def refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        c = clients.find_client(self.client_id)
        if not c:
            return
        for d in c.get("divisions", []):
            if not (isinstance(d, dict) and d.get("id") == self.division_id):
                continue
            sites = d.get("sites", [])
            if not isinstance(sites, list):
                return
            for s in sites:
                self.tree.insert("", tk.END, iid=s.get("id", ""), values=(s.get("name", ""), s.get("phone", "")))

    def _site_dialog(self, title: str, init_name: str = "", init_phone: str = "") -> tuple[str | None, str]:
        dlg = tk.Toplevel(self)
        dlg.title(title)
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        ttk.Label(dlg, text="Name:").grid(row=0, column=0, sticky="w", padx=12, pady=(12, 4))
        name_var = tk.StringVar(value=init_name)
        name_ent = ttk.Entry(dlg, textvariable=name_var, width=40)
        name_ent.grid(row=0, column=1, sticky="ew", padx=12, pady=(12, 4))

        ttk.Label(dlg, text="Phone:").grid(row=1, column=0, sticky="w", padx=12, pady=(4, 8))
        phone_var = tk.StringVar(value=init_phone)
        phone_ent = ttk.Entry(dlg, textvariable=phone_var, width=40)
        phone_ent.grid(row=1, column=1, sticky="ew", padx=12, pady=(4, 8))

        btns = ttk.Frame(dlg)
        btns.grid(row=2, column=0, columnspan=2, sticky="e", padx=12, pady=(0, 12))
        ttk.Button(btns, text="OK", command=dlg.destroy).grid(row=0, column=0, padx=6)
        ttk.Button(btns, text="Cancel",
                   command=lambda: (name_var.set("__CANCEL__"), dlg.destroy())).grid(row=0, column=1)

        name_ent.focus_set()
        dlg.wait_window(dlg)

        name = name_var.get().strip()
        if name == "__CANCEL__" or not name:
            return (None, "")
        phone = phone_var.get().strip()
        return (name, phone)

    def add_site(self):
        name, phone = self._site_dialog("Add Site")
        if name is None:
            return
        added = clients.add_site(self.client_id, self.division_id, name=name, phone=phone)
        if not added:
            messagebox.showerror("Add Site", "Add failed.")
        self.refresh()

    def edit_site(self):
        sid = self.selected_id()
        if not sid:
            messagebox.showinfo("Edit Site", "Select a site first.")
            return
        c = clients.find_client(self.client_id)
        if not c:
            messagebox.showerror("Edit Site", "Client not found.")
            return
        cur = None
        for d in c.get("divisions", []):
            if not (isinstance(d, dict) and d.get("id") == self.division_id):
                continue
            for s in d.get("sites", []):
                if isinstance(s, dict) and s.get("id") == sid:
                    cur = s
                    break
        if not cur:
            messagebox.showerror("Edit Site", "Site not found.")
            return
        name, phone = self._site_dialog("Edit Site", init_name=cur.get("name", ""), init_phone=cur.get("phone", ""))
        if name is None:
            return
        if not clients.update_site(self.client_id, self.division_id, sid, name=name, phone=phone):
            messagebox.showerror("Edit Site", "Update failed.")
        self.refresh()

    def delete_site(self):
        sid = self.selected_id()
        if not sid:
            messagebox.showinfo("Delete Site", "Select a site first.")
            return
        item = self.tree.item(sid)
        nm = item["values"][0] if item and item.get("values") else "(unnamed)"
        if not messagebox.askyesno("Delete Site", f"Delete '{nm}'?"):
            return
        if not clients.delete_site(self.client_id, self.division_id, sid):
            messagebox.showerror("Delete Site", "Delete failed.")
        self.refresh()


# ---------------- Right-pane views ----------------
class MonthlyImportView(ttk.Frame):
    
    def __init__(self, parent: tk.Widget, on_back):
        super().__init__(parent, padding=12)
        self.on_back = on_back
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)  # table row

        # Header
        hdr = ttk.Frame(self)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 6))
        ttk.Label(hdr, text="New Monthly Invoice", font=("", 14, "bold")).pack(side="left")
        ttk.Button(hdr, text="Back", command=self.on_back).pack(side="right")

        # --- controls row: month/year (left) + starting invoice (right) ---
        controls = ttk.Frame(self)
        controls.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        controls.columnconfigure(0, weight=1)   # left zone grows
        controls.columnconfigure(1, weight=0)   # right zone hugs content

        # left: month/year + Apply
        left = ttk.Frame(controls)
        left.grid(row=0, column=0, sticky="w")
        self.month_var = tk.StringVar(value=str(datetime.now().month))
        self.year_var  = tk.StringVar(value=str(datetime.now().year))

        ttk.Label(left, text="Month (1–12):").pack(side="left", padx=(0, 6))
        self.month_entry = ttk.Entry(left, textvariable=self.month_var, width=4)
        self.month_entry.pack(side="left")

        ttk.Label(left, text="Year:").pack(side="left", padx=(12, 6))
        self.year_entry = ttk.Entry(left, textvariable=self.year_var, width=6)
        self.year_entry.pack(side="left")

        ttk.Button(left, text="Apply", command=self._apply_month_year).pack(side="left", padx=(12, 0))

        # right: starting invoice number
        right = ttk.Frame(controls)
        right.grid(row=0, column=1, sticky="e")
        ttk.Label(right, text="Starting invoice #:").pack(side="left", padx=(0, 6))
        self.start_num_var = tk.StringVar()
        try:
            preset = inv.get_starting_invoice_number()
            if preset is not None:
                self.start_num_var.set(str(preset))
        except Exception:
            pass
        ttk.Entry(right, textvariable=self.start_num_var, width=10).pack(side="left")

        # Table of selected files
        cols = ("file", "type", "phone", "match")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for c, w, a in (("file", 420, "w"), ("type", 90, "center"), ("phone", 120, "center"), ("match", 260, "w")):
            self.tree.heading(c, text=c.capitalize())
            self.tree.column(c, width=w, anchor=a)
        self.tree.grid(row=2, column=0, sticky="nsew")

        # colors for validation
        self.tree.tag_configure("ok",  background="#e6ffed")   # light green
        self.tree.tag_configure("bad", background="#ffecec")   # light red
        self.tree.tag_configure("unk", background="")          # default

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=2, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=ybar.set)

        # double-click to preview
        self.tree.bind("<Double-1>", self._on_preview)

        # Buttons
        btns = ttk.Frame(self)
        btns.grid(row=3, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for i in range(3): btns.columnconfigure(i, weight=1)
        ttk.Button(btns, text="Add CSV Files…", command=self.add_files).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Remove Selected", command=self.remove_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Clear All", command=self.clear_all).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="Continue", command=self.on_continue).grid(row=0, column=3, sticky="ew", padx=4)

        # revalidate when month/year edits lose focus or hit Return
        self.month_entry.bind("<FocusOut>", lambda e: self._revalidate_all())
        self.year_entry.bind("<FocusOut>",  lambda e: self._revalidate_all())
        self.month_entry.bind("<Return>",   lambda e: self._revalidate_all())
        self.year_entry.bind("<Return>",    lambda e: self._revalidate_all())

    # ---- file ops ----
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select CSV files",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not paths:
            return
        clients_doc = clients.load_clients()
        for p in paths:
            pth = Path(p)
            try:
                info = inv.identify_csv_and_phone(pth, clients_doc)
                kind  = info.get("kind", "unknown")
                phone = info.get("phone", "")
                match = info.get("match") or {}
                match_str = self._format_match(match)
            except Exception as e:
                kind, phone, match_str = "error", "", f"Error: {e}"
            self.tree.insert("", tk.END, values=(str(pth), kind, phone, match_str))
        # validate after adding
        self._revalidate_all()

    def _format_match(self, match: dict) -> str:
        if not isinstance(match, dict) or not match:
            return ""
        parts = [match.get("client_name", ""), match.get("division_name", ""), match.get("site_name", "")]
        
        return " > ".join([p for p in parts if p])

    def remove_selected(self):
        for iid in self.tree.selection():
            self.tree.delete(iid)

    def clear_all(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)

    # ---------- preview helpers ----------
    def _open_csv_preview(self, path: Path, highlight_rows=None):
        headers, rows = [], []
        try:
            with path.open("r", encoding="utf-8-sig", newline="") as f:
                reader = csv.reader(f)
                headers = next(reader, [])
                for i, row in enumerate(reader):
                    if i >= 500:
                        break
                    rows.append(row)
        except Exception as e:
            messagebox.showerror("CSV Preview", f"Failed to read file:\n{path}\n\n{e}")
            return

        dlg = tk.Toplevel(self)
        dlg.title(f"Preview — {path.name}")
        dlg.geometry("900x520")
        dlg.transient(self)
        dlg.grab_set()
        dlg.columnconfigure(0, weight=1)
        dlg.rowconfigure(0, weight=1)

        tree = ttk.Treeview(dlg, show="headings")
        tree.grid(row=0, column=0, sticky="nsew")
        vbar = ttk.Scrollbar(dlg, orient="vertical", command=tree.yview)
        hbar = ttk.Scrollbar(dlg, orient="horizontal", command=tree.xview)
        vbar.grid(row=0, column=1, sticky="ns")
        hbar.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

        # Columns
        num_cols = max(len(headers), max((len(r) for r in rows), default=0))
        col_ids = [f"c{i}" for i in range(num_cols)]
        tree["columns"] = col_ids
        for i, cid in enumerate(col_ids):
            head = headers[i] if i < len(headers) else f"Column {i+1}"
            tree.heading(cid, text=head)
            tree.column(cid, width=140, stretch=True, anchor="w")

        # Highlight out-of-range rows if provided
        bad_preview_idx: set[int] = set()
        if highlight_rows:
            for file_row_num, _cell in highlight_rows:
                idx = file_row_num - 2  # header is line 1
                if 0 <= idx < len(rows):
                    bad_preview_idx.add(idx)

        for i, r in enumerate(rows):
            values = r + [""] * (num_cols - len(r))
            tags = ("bad",) if i in bad_preview_idx else ()
            tree.insert("", tk.END, values=values, tags=tags)

        try:
            tree.tag_configure("bad", background="#ffe5e5")
        except Exception:
            pass

        footer = ttk.Frame(dlg, padding=(8, 6))
        footer.grid(row=2, column=0, columnspan=2, sticky="ew")
        footer.columnconfigure(0, weight=1)
        if bad_preview_idx:
            ttk.Label(footer, text=f"Rows shown: {len(rows)} (max 500) — Out-of-range rows: {len(bad_preview_idx)}").grid(row=0, column=0, sticky="w")
        else:
            ttk.Label(footer, text=f"Rows shown: {len(rows)} (max 500)").grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="Close", command=dlg.destroy).grid(row=0, column=1, sticky="e")

    def _on_preview(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        path = Path(self.tree.set(iid, "file"))
        kind = (self.tree.set(iid, "type") or "").strip().lower()

        # If month/year are provided, compute the out-of-range rows
        highlight_rows = None
        try:
            y = int(getattr(self, "year_var", tk.StringVar(value="")).get() or 0)
            m = int(getattr(self, "month_var", tk.StringVar(value="")).get() or 0)
            if y and m and kind in {"messages", "calls"}:
                highlight_rows = inv.find_out_of_month_rows(path, kind, y, m)
        except Exception:
            highlight_rows = None

        self._open_csv_preview(path, highlight_rows=highlight_rows)

    # ---------- month/year validation ----------
    def _get_year_month(self) -> tuple[int | None, int | None]:
        try:
            y = int(self.year_var.get().strip())
            m = int(self.month_var.get().strip())
            if 1 <= m <= 12 and 2000 <= y <= 2100:
                return y, m
        except Exception:
            pass
        return (None, None)

    def _revalidate_all(self):
        y, m = self._get_year_month()
        for iid in self.tree.get_children():
            self._validate_row(iid, y, m)

    def _validate_row(self, iid, y, m):
        # clear tags
        self.tree.item(iid, tags=("unk",))
        if not y or not m:
            return
        path = self.tree.set(iid, "file")
        kind = self.tree.set(iid, "type")
        ok, _stats = inv.check_csv_month_year(path, kind, y, m)
        self.tree.item(iid, tags=("ok",) if ok else ("bad",))

    def _apply_month_year(self):
        # Validate month & year
        try:
            month = int(self.month_var.get())
            year = int(self.year_var.get())
            if not (1 <= month <= 12):
                raise ValueError
        except Exception:
            messagebox.showinfo("Month/Year", "Enter a valid month (1–12) and year (e.g., 2025)." )
            return

        # Read & persist starting invoice number (optional)
        raw = (self.start_num_var.get() or "").strip()
        if raw:
            try:
                start_no = int(raw)
                if start_no < 0:
                    raise ValueError
                inv.set_starting_invoice_number(start_no)
            except Exception:
                messagebox.showinfo("Starting invoice #", "Please enter a non-negative whole number.")
                return

        # Re-check each file for the selected month/year and recolor rows
        self._revalidate_all()

    # ---------- Continue: build & export invoice ----------
    def _site_from_match(self, text: str) -> str | None:
        parts = [p.strip() for p in (text or "").split(">") if p.strip()]
        return parts[-1] if parts else None

    def on_continue(self):
        """
        Verify CSVs for the selected month/year, build invoice,
        add Voice line items (qty = row count, $0.14 each) and Message line items likewise,
        then run a single finalize_with_template() that:
        - saves JSON (internal)
        - exports CSV (user folder)
        - fills Excel template (.xlsm) and exports PDF on Windows (Excel + pywin32)
        - falls back to ReportLab PDF elsewhere.
        Also decorates each line-item description with (-LAST4) from the match column,
        and sets BILL TO to the parent organization (name + address) from clients.json.
        """
        y, m = self._get_year_month()
        if not y or not m:
            messagebox.showerror("Month/Year required", "Enter Month and Year, then click Apply.")
            return

        bad = []
        calls_with_sites = []
        messages_with_sites = []
        for iid in self.tree.get_children():
            tags = set(self.tree.item(iid, "tags") or ())
            path = Path(self.tree.set(iid, "file"))
            kind = (self.tree.set(iid, "type") or "").strip().lower()
            match_text = self.tree.set(iid, "match") or ""
            site_name = self._site_from_match(match_text)

            if "ok" not in tags:
                bad.append(path.name)
            elif kind == "calls":
                calls_with_sites.append((str(path), site_name))
            elif kind == "messages":
                messages_with_sites.append((str(path), site_name))

        if bad:
            message = "Some files are not within the selected month/year:\n\n" + "\n".join(f"- {n}" for n in bad)
            messagebox.showerror("Validation failed", message)
            return

        # Build invoice
        inv_obj = inv.new_monthly_invoice(y, m)

        # Starting invoice number (optional)
        start_txt = (self.start_num_var.get() or "").strip()
        if start_txt.isdigit():
            inv_obj["starting_invoice_number"] = int(start_txt)
            try:
                inv.set_starting_invoice_number(int(start_txt))
            except Exception:
                pass

        # Add voice and message line items
        if calls_with_sites:
            inv.add_voice_items_to_invoice(inv_obj, calls_with_sites, y, m, inv.UNIT_PRICE_VOICE)
        if messages_with_sites:
            inv.add_message_items_to_invoice(inv_obj, messages_with_sites, y, m, inv.UNIT_PRICE_SMS)

        # Decorate descriptions with phone last4 from match column
        import re as _re, json
        site_phones = {}
        for iid in self.tree.get_children():
            match_text = self.tree.set(iid, "match") or ""
            site_name = self._site_from_match(match_text)
            if not site_name:
                continue
            m_last4 = _re.search(r"\((?:-|–)?(\d{4})\)", match_text)  # matches '(-1234)' or '(–1234)'
            if m_last4:
                site_phones[site_name] = m_last4.group(1)
        if site_phones:
            inv_obj["site_phones"] = site_phones

        # Infer top-level parent org for BILL TO from clients.json
        def _norm_name(s: str) -> str:
            s = (s or "").upper().strip()
            s = s.replace("—", "-").replace("–", "-")
            s = _re.sub(r"\s+", " ", s)
            for suf in (" VOICE", " SMS", "- VOICE", "- SMS"):
                if s.endswith(suf):
                    s = s[: -len(suf)].strip()
            return s

        here = Path(__file__).resolve().parent
        clients_path = here / "data" / "clients.json"
        parent_name = None
        try:
            data = json.loads(clients_path.read_text(encoding="utf-8"))
            target_sites = {_norm_name(s) for _, s in (calls_with_sites + messages_with_sites) if s}
            for c in (data.get("clients") or []):
                found = False
                for d in (c.get("divisions") or []):
                    for s in (d.get("sites") or []):
                        sn = _norm_name(s.get("name", ""))
                        if any(sn == t or sn.startswith(t) or t in sn for t in target_sites):
                            found = True
                            break
                    if found:
                        break
                if found:
                    parent_name = c.get("name")
                    break
        except Exception:
            parent_name = None

        if parent_name:
            inv_obj["client_name_snapshot"] = parent_name  # invoicing.py will fill name+address under BILL TO

        # Locate invoice template next to app.py
        template_candidates = [here / "invoice.xlsm", here / "Invoice Template HP.xlsm"]
        template_path = None
        for cand in template_candidates:
            if cand.exists():
                template_path = cand
                break
        if not template_path:
            messagebox.showwarning(
                "Template not found",
                "Couldn't find 'invoice.xlsm' or 'Invoice Template HP.xlsm' next to app.py.\n"
                "I'll still generate a PDF with the fallback layout."
            )
            template_path = here / "invoice.xlsm"  # finalize() will still run; Excel export will fallback

        # Run one-step pipeline to generate outputs
        try:
            paths = inv.finalize_with_template(inv_obj, str(template_path))

        except Exception as e:
            messagebox.showerror("Invoice error", f"Failed to create invoice:\n{e}")
            return

        # Notify success
        parts = []
        if paths.get("json"): parts.append(f"Saved JSON: {paths['json']}")
        if paths.get("csv"):  parts.append(f"Exported CSV: {paths['csv']}")
        if paths.get("xlsm"): parts.append(f"Filled Excel: {paths['xlsm']}")
        if paths.get("pdf"):  parts.append(f"Exported PDF: {paths['pdf']}")
        if not parts:
            parts = ["No files were produced."]
        messagebox.showinfo("Invoice created", "\n".join(parts))




class ViewInvoicesView(ttk.Frame):
    """Simple list of saved invoices using invoicing.list_invoices()."""
    def __init__(self, parent: tk.Widget, on_back):
        super().__init__(parent, padding=12)
        self.on_back = on_back
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        hdr = ttk.Frame(self)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(hdr, text="Past Invoices", font=("", 14, "bold")).pack(side="left")
        ttk.Button(hdr, text="Back", command=self.on_back).pack(side="right")

        cols = ("id", "type", "period", "client", "total")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
        self.tree.heading("id", text="ID")
        self.tree.heading("type", text="Type")
        self.tree.heading("period", text="Period")
        self.tree.heading("client", text="Client")
        self.tree.heading("total", text="Total")
        self.tree.column("id", width=220, anchor="w")
        self.tree.column("type", width=80, anchor="center")
        self.tree.column("period", width=120, anchor="center")
        self.tree.column("client", width=220, anchor="w")
        self.tree.column("total", width=100, anchor="e")
        self.tree.grid(row=1, column=0, sticky="nsew")

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=ybar.set)

        self.refresh()

    def refresh(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        try:
            lst = inv.list_invoices()
        except Exception as e:
            messagebox.showerror("Invoices", f"Failed to load invoices:\n{e}")
            return

        for item in lst:
            pid = item.get("id", "")
            ptype = item.get("type", "")
            period = item.get("period") or {}
            if isinstance(period, dict) and "year" in period and "month" in period:
                ptxt = f"{period['year']}-{int(period['month']):02d}"
            else:
                ptxt = ""
            client = item.get("client_name", "") or item.get("client_id", "")
            total = item.get("total", 0.0)
            self.tree.insert("", tk.END, values=(pid, ptype, ptxt, client, f"{total:,.2f}"))
# ---------------- Helpers for invoice edits & finalize ----------------

def load_clients_doc(path):
    """Load data/clients.json -> dict (return {} if missing/bad)."""
    try:
        import json
        from pathlib import Path
        p = Path(path)
        if not p.exists():
            return {}
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:
        return {}

def _normalize_site_key(s: str) -> str:
    """Uppercase, strip, remove VOICE/SMS suffix and any trailing ' - ' junk."""
    import re
    u = (s or "").upper().strip()
    u = u.replace("—", "-")
    for suf in (" VOICE", " SMS"):
        if u.endswith(suf):
            u = u[:-len(suf)].strip()
    # If there is an en dash/ hyphen separating base name, prefer left side
    if " - " in u:
        u = u.split(" - ", 1)[0].strip()
    if "–" in u:
        u = u.split("–", 1)[0].strip()
    # collapse internal whitespace
    u = re.sub(r"\s+", " ", u)
    return u

def decorate_descriptions_with_last4(inv_obj: dict, site_phones: dict) -> None:
    """
    Append ' (-LAST4)' to each description when we have a phone match.
    We DON'T remove VOICE/SMS text from names; we only add the suffix.
    """
    import re
    phones = {(_normalize_site_key(k)): v for k, v in (site_phones or {}).items()}
    for li in inv_obj.get("line_items", []) or []:
        desc = (li.get("description") or "").strip()
        if not desc:
            continue
        # Skip if already has (-####)
        if re.search(r"\(\-\d{4}\)\s*$", desc):
            continue
        base_key = _normalize_site_key(desc)
        last4 = phones.get(base_key)
        if not last4 and "–" in desc:
            base_key = _normalize_site_key(desc.split("–", 1)[0])
            last4 = phones.get(base_key)
        if last4 and len(str(last4)) == 4 and str(last4).isdigit():
            li["description"] = f"{desc} (-{last4})"

def infer_parent_billto_from_clients(inv_obj: dict, clients_doc: dict) -> None:
    """
    Look up the parent organization for any site in line_items and set inv_obj['client']
    to {'name': PARENT_NAME, 'address': ADDRESS, 'contact': optional}.
    Uses the first matched parent found among the invoice's sites.
    """
    if not clients_doc:
        return

    # Build a quick map: site_name_norm -> (parent_name, parent_address, parent_contact)
    mapping = {}  # normalized site -> (parent, address, contact)
    for c in clients_doc.get("clients", []) or []:
        parent_name = (c.get("name") or "").strip()
        parent_addr = (c.get("address") or "").strip()
        parent_contact = (c.get("contact") or "").strip()
        for div in c.get("divisions", []) or []:
            for s in div.get("sites", []) or []:
                site_nm = _normalize_site_key(s.get("name") or "")
                if site_nm:
                    mapping[site_nm] = (parent_name, parent_addr, parent_contact)

    # Scan the invoice line items until we find a site that maps to a parent
    client = inv_obj.setdefault("client", {})
    for li in inv_obj.get("line_items", []) or []:
        desc = (li.get("description") or "").strip()
        if not desc:
            continue
        key = _normalize_site_key(desc)
        hit = mapping.get(key)
        if not hit and "–" in desc:
            key = _normalize_site_key(desc.split("–", 1)[0])
            hit = mapping.get(key)
        if hit:
            client["name"], client["address"], client["contact"] = hit
            break

def _finalize_shim(inv_obj: dict, template_path: str) -> dict:
    """
    Save JSON, export CSV, then try to fill Excel template (and PDF if Windows+Excel),
    else fall back to ReportLab PDF. Returns paths dict with any of: json,csv,xlsm,xlsx,pdf.
    """
    paths = {}
    # 1) JSON (internal)
    try:
        p = inv.save_invoice(inv_obj)
        paths["json"] = str(p)
    except Exception:
        pass

    # 2) CSV
    try:
        p = inv.export_invoice_csv(inv_obj)
        paths["csv"] = str(p)
    except Exception:
        pass

    # 3) Excel template (+ PDF on Windows via pywin32) – support v3 or legacy name
    try:
        if hasattr(inv, "export_invoice_with_excel_template_v3"):
            out = inv.export_invoice_with_excel_template_v3(
                inv_obj,
                template_path=str(template_path),
                out_dir=None,
                export_pdf=True
            )
            # out may be a dict or tuple; normalize:
            if isinstance(out, dict):
                paths.update({k: str(v) for k, v in out.items() if v})
            else:
                # (xlsm_path, pdf_path) style
                if len(out) >= 1 and out[0]: paths["xlsm"] = str(out[0])
                if len(out) >= 2 and out[1]: paths["pdf"]  = str(out[1])
        elif hasattr(inv, "export_invoice_with_excel_template"):
            out = inv.export_invoice_with_excel_template(
                inv_obj,
                str(template_path)
            )
            if isinstance(out, dict):
                paths.update({k: str(v) for k, v in out.items() if v})
            else:
                if len(out) >= 1 and out[0]: paths["xlsm"] = str(out[0])
                if len(out) >= 2 and out[1]: paths["pdf"]  = str(out[1])
        elif hasattr(inv, "export_invoice_pdf_via_template"):
            # Some versions only return a PDF via template
            p = inv.export_invoice_pdf_via_template(inv_obj, str(template_path))
            if p:
                paths["pdf"] = str(p)
    except Exception:
        # Fall through to ReportLab PDF
        pass

    # 4) Fallback PDF (simple layout)
    if "pdf" not in paths:
        try:
            p = inv.export_invoice_pdf(inv_obj)
            if p:
                paths["pdf"] = str(p)
        except Exception:
            pass

    return paths


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
