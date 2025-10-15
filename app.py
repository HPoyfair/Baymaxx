# app.py
import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from PIL import Image, ImageTk
import view_clients as clients

import invoicing as inv
from tkinter import filedialog



class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Baymaxx — Invoicing Toolkit")
        self.geometry("760x520")

        # ---- grid: 2 cols, fixed left rail, flexible right content
        self.grid_columnconfigure(0, weight=0, minsize=180)   # sidebar width
        self.grid_columnconfigure(1, weight=1)                 # content
        self.grid_rowconfigure(0, weight=0)                    # topbar
        self.grid_rowconfigure(1, weight=1)                    # main area

        # ---- top bar
        topbar = ttk.Frame(self, padding=(12, 8))
        topbar.grid(row=0, column=0, columnspan=2, sticky="ew")
        ttk.Label(topbar, text="help").grid(row=0, column=0, sticky="w")

        # ---- left sidebar (buttons)
        sidebar = ttk.Frame(self, padding=(10, 10))
        sidebar.grid(row=1, column=0, sticky="nsw")
        # stack buttons vertically with consistent spacing
        ttk.Button(sidebar, text="New Month Invoice",
           command=self.show_monthly_import).pack(fill="x", pady=6)

        ttk.Button(sidebar, text="View Past Invoices",
           command=self.show_invoices).pack(fill="x", pady=6)

        ttk.Button(sidebar, text="New Individual Invoice").pack(fill="x", pady=6)
        ttk.Button(sidebar, text="View Clients", command=self.open_clients_manager).pack(fill="x", pady=6)

        # ---- right content (logo centered)
        self.content = ttk.Frame(self, padding=16)
        self.content.grid(row=1, column=1, sticky="nsew")
        self.content.rowconfigure(0, weight=1)
        self.content.columnconfigure(0, weight=1)

        logo_path = Path(__file__).resolve().parent / "baymaxx.png"
        try:
            self.logo_img = self.load_logo(logo_path, max_w=420, max_h=420)
            # centered by default: no sticky needed
            ttk.Label(self.content, image=self.logo_img).grid(row=0, column=0)
        except Exception:
            ttk.Label(self.content, text="Baymaxx", font=("", 28, "bold")).grid(row=0, column=0)


    def load_logo(self, path: Path, max_w: int, max_h: int):
        img = Image.open(path)
        img.thumbnail((max_w, max_h), Image.LANCZOS)
        return ImageTk.PhotoImage(img)

    def open_clients_manager(self):
        ClientsManager(self)

    def _set_content(self, widget: tk.Widget) -> None:
        """Replace whatever is in the right pane with `widget`."""
        for child in self.content.winfo_children():
            child.destroy()
        widget.grid(row=0, column=0)  # centered by grid weights
    
    def show_home(self) -> None:
    # restore the logo content
        for child in self.content.winfo_children():
            child.destroy()
        logo_path = Path(__file__).resolve().parent / "baymaxx.png"
        try:
            self.logo_img = self.load_logo(logo_path, max_w=420, max_h=420)
            ttk.Label(self.content, image=self.logo_img).grid(row=0, column=0)
        except Exception:
            ttk.Label(self.content, text="Baymaxx", font=("", 28, "bold")).grid(row=0, column=0)

    def show_monthly_import(self) -> None:
        view = MonthlyImportView(self.content, on_back=self.show_home)
        self._set_content(view)

    def show_invoices(self) -> None:
        view = ViewInvoicesView(self.content, on_back=self.show_home)
        self._set_content(view)
    





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
        ttk.Button(btns, text="Cancel", command=lambda: (name_var.set("__CANCEL__"), dlg.destroy())).grid(row=0, column=1)

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
class MonthlyImportView(ttk.Frame):
    """
    Right-pane view for importing CSVs for a monthly invoice.
    Lets the user add/remove CSV files and shows detected type + phone + match.
    """
    def __init__(self, parent: tk.Widget, on_back):
        super().__init__(parent, padding=12)
        self.on_back = on_back
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        hdr = ttk.Frame(self)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(hdr, text="New Monthly Invoice", font=("", 14, "bold")).pack(side="left")
        ttk.Button(hdr, text="Back", command=self.on_back).pack(side="right")

        # Table of selected files
        cols = ("file", "type", "phone", "match")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        for c, w, a in (("file", 340, "w"), ("type", 90, "center"), ("phone", 120, "center"), ("match", 240, "w")):
            self.tree.heading(c, text=c.capitalize())
            self.tree.column(c, width=w, anchor=a)
        self.tree.grid(row=1, column=0, sticky="nsew")

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=1, column=1, sticky="ns")
        self.tree.configure(yscrollcommand=ybar.set)

        # Buttons
        btns = ttk.Frame(self)
        btns.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(8, 0))
        for i in range(3): btns.columnconfigure(i, weight=1)
        ttk.Button(btns, text="Add CSV Files…", command=self.add_files).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Remove Selected", command=self.remove_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Clear All", command=self.clear_all).grid(row=0, column=2, sticky="ew", padx=4)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select CSV files",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if not paths:
            return

        # load clients snapshot for matching
        clients_doc = clients.load_clients()

        for p in paths:
            pth = Path(p)
            # Ask invoicing helper to detect type + phone + match
            try:
                info = inv.identify_csv_and_phone(pth, clients_doc)
                # info example: {"kind": "sms"|"call"|"unknown", "phone": "+15551234567", "match": {client/division/site…}}
                kind = info.get("kind", "unknown")
                phone = info.get("phone", "")
                match = info.get("match") or {}
                match_str = self._format_match(match)
            except Exception as e:
                kind, phone, match_str = "error", "", f"Error: {e}"

            self.tree.insert(
                "", tk.END,
                values=(pth.name, kind, phone, match_str),
                tags=("path",)
            )
            # stash full path on the item for later processing
            last = self.tree.get_children()[-1]
            self.tree.set(last, column="file", value=str(pth))  # keep full path in file column

    def _format_match(self, match: dict) -> str:
        if not isinstance(match, dict) or not match:
            return ""
        # Build a short breadcrumb like: Client > Division > Site
        parts = [
            match.get("client_name", ""),
            match.get("division_name", ""),
            match.get("site_name", ""),
        ]
        return " > ".join([p for p in parts if p])

    def remove_selected(self):
        for iid in self.tree.selection():
            self.tree.delete(iid)

    def clear_all(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)


class ViewInvoicesView(ttk.Frame):
    """Simple list of saved invoices using invoicing.list_invoices()."""
    def __init__(self, parent: tk.Widget, on_back):
        super().__init__(parent, padding=12)
        self.on_back = on_back
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Header
        hdr = ttk.Frame(self)
        hdr.grid(row=0, column=0, sticky="ew", pady=(0, 8))
        ttk.Label(hdr, text="Past Invoices", font=("", 14, "bold")).pack(side="left")
        ttk.Button(hdr, text="Back", command=self.on_back).pack(side="right")

        # Table
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
        # Clear existing
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        # Load from invoicing module
        try:
            lst = inv.list_invoices()
        except Exception as e:
            messagebox.showerror("Invoices", f"Failed to load invoices:\n{e}")
            return

        # Populate
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
