# app.py
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from pathlib import Path
from PIL import Image, ImageTk
import csv

import view_clients as clients
import invoicing as inv


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
        try:
            self.logo_img = self.load_logo(logo_path, max_w=420, max_h=420)
            ttk.Label(self.content, image=self.logo_img).grid(row=0, column=0)
        except Exception:
            ttk.Label(self.content, text="Baymaxx", font=("", 28, "bold")).grid(row=0, column=0)
        self.after(0, lambda: inv.ensure_invoice_root(self))

    def load_logo(self, path: Path, max_w: int, max_h: int):
        img = Image.open(path)
        img.thumbnail((max_w, max_h), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def open_clients_manager(self):
        ClientsManager(self)

    def show_home(self) -> None:
        for child in self.content.winfo_children():
            child.destroy()
        logo_path = Path(__file__).resolve().parent / "baymaxx.png"
        try:
            self.logo_img = self.load_logo(logo_path, max_w=420, max_h=420)
            ttk.Label(self.content, image=self.logo_img).grid(row=0, column=0)
        except Exception:
            ttk.Label(self.content, text="Baymaxx", font=("", 28, "bold")).grid(row=0, column=0)

    def show_monthly_import(self) -> None:
        # clear -> create -> grid (so we don't destroy the new view)
        for child in self.content.winfo_children():
            child.destroy()
        view = MonthlyImportView(self.content, on_back=self.show_home)
        view.grid(row=0, column=0)

    def show_invoices(self) -> None:
        # clear -> create -> grid
        for child in self.content.winfo_children():
            child.destroy()
        view = ViewInvoicesView(self.content, on_back=self.show_home)
        view.grid(row=0, column=0)


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
        shead = self.tree.heading("sites", text="# Sites")
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
    """
    Right-pane view for importing CSVs for a monthly invoice.
    Lets the user add/remove CSV files and shows detected type + phone + match.
    Double-click a row to preview the CSV.
    Adds a month/year bar and colors rows by month/year validity.
    """
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

        # --- month/year bar ---
        bar = ttk.Frame(self)
        bar.grid(row=1, column=0, sticky="ew", pady=(0, 6))
        bar.columnconfigure(5, weight=1)

        # month dropdown
        self.month_var = tk.StringVar(value="10")  # default: October (adjust if you like)
        self.year_var = tk.StringVar(value=str(Path().stat().st_mtime_ns // 10**9 and __import__("datetime").datetime.now().year))

        ttk.Label(bar, text="Month (1–12):").grid(row=0, column=0, padx=(0, 6))
        self.month_entry = ttk.Entry(bar, textvariable=self.month_var, width=5)
        self.month_entry.grid(row=0, column=1)

        ttk.Label(bar, text="Year:").grid(row=0, column=2, padx=(12, 6))
        self.year_entry = ttk.Entry(bar, textvariable=self.year_var, width=6)
        self.year_entry.grid(row=0, column=3)

        ttk.Button(bar, text="Apply", command=self._revalidate_all).grid(row=0, column=4, padx=(12, 0))

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

        # also revalidate when month/year edits lose focus or hit Return
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

            iid = self.tree.insert("", tk.END, values=(str(pth), kind, phone, match_str))
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
    def _on_preview(self, _event=None):
        sel = self.tree.selection()
        if not sel:
            return
        iid = sel[0]
        path = self.tree.set(iid, "file")  # full path in the 'file' column
        if not path:
            return
        self._open_csv_preview(Path(path))



    def _open_csv_preview(self, path: Path, highlight_rows=None):
        # Read up to ~500 rows (preview only)
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

        # Configure row highlight tags
        try:
            tree.tag_configure("bad", background="#ffe5e5")   # light red
            tree.tag_configure("good", background="#e8ffe8")  # light green (optional)
        except Exception:
            pass

        # Columns
        num_cols = max(len(headers), max((len(r) for r in rows), default=0))
        col_ids = [f"c{i}" for i in range(num_cols)]
        tree["columns"] = col_ids
        for i, cid in enumerate(col_ids):
            head = headers[i] if i < len(headers) else f"Column {i+1}"
            tree.heading(cid, text=head)
            tree.column(cid, width=140, stretch=True, anchor="w")

        # Compute set of "bad" preview-row indices from file rows
        bad_preview_idx: set[int] = set()
        if highlight_rows:
            # highlight_rows contains 1-based CSV lines, header=1, first data row=2
            # our preview rows are 0-based for the first data row (line 2)
            for file_row_num, _cell in highlight_rows:
                idx = file_row_num - 2
                if 0 <= idx < len(rows):
                    bad_preview_idx.add(idx)

        # Insert rows with tags
        for i, r in enumerate(rows):
            values = r + [""] * (num_cols - len(r))
            tags = ()
            if bad_preview_idx:
                tags = ("bad",) if i in bad_preview_idx else ()
            tree.insert("", tk.END, values=values, tags=tags)

        # Footer
        footer = ttk.Frame(dlg, padding=(8, 6))
        footer.grid(row=2, column=0, columnspan=2, sticky="ew")
        footer.columnconfigure(0, weight=1)

        # Summary line if we know how many are bad
        if bad_preview_idx:
            ttk.Label(
                footer,
                text=f"Rows shown: {len(rows)} (max 500) — Out-of-range rows: {len(bad_preview_idx)}",
            ).grid(row=0, column=0, sticky="w")
        else:
            ttk.Label(footer, text=f"Rows shown: {len(rows)} (max 500)").grid(row=0, column=0, sticky="w")

        ttk.Button(footer, text="Close", command=dlg.destroy).grid(row=0, column=1, sticky="e")






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


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
