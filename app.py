import tkinter as tk
from tkinter import ttk, messagebox
from pathlib import Path
from PIL import Image, ImageTk
import view_clients as clients



class App(tk.Tk):

    def __init__(self):
        super().__init__()

        # --- windo basics ---
        self.title("Baymaxx") #title
        self.geometry("700x500") #window size
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight= 0)
        self.columnconfigure(1, weight=1)
        self.rowconfigure(1, weight= 1)



        # -- wrapper --
        wrapper = tk.Frame(self, padx = 1, pady = 1)
        wrapper.grid(row = 1, column = 0, sticky = "nsew", padx = 10, pady = 10)
        wrapper.columnconfigure(0, weight =1)
        wrapper.rowconfigure(0, weight =1)

        # ---content frame----
        left_frame = ttk.Frame(wrapper, padding=20, borderwidth = 1, relief = "solid")
        left_frame.grid(row = 0, column= 0, sticky ="nsew")
        left_frame.grid_columnconfigure(0, weight = 1)

        # -- label and button ---
        self.status = tk.StringVar(value="Ready")
        #title = ttk.Label(left_frame , text = "testing", font =("",30, "bold"))
        #title.grid(row = 0, column =0, sticky="w")


        #-- counter label --
        self.counter = 0
        self.counter_label= ttk.Label(left_frame , text = f"count: {self.counter}")
        #self.counter_label.grid(row = 1, column = 0, sticky = "w", pady = (12, 0))

        # --- button ---

        btns = ttk.Frame(left_frame)
        btns.grid(row = 0, column =0, sticky= "nsew", pady = 8)

        btns.grid_columnconfigure(0, weight = 1)
        btns.grid_columnconfigure(1, weight = 1)
        btns.grid_columnconfigure(2, weight = 1)

        
        c_new_m_invoice = ttk.Button(btns , text = "New Month Invoice")
        view_old= ttk.Button(btns, text = "View Past Invoices")
        c_indiv= ttk.Button(btns , text = "New Individual Invoice")
        v_clients = ttk.Button(btns, text = "View Clients", command = self.open_clients_manager)
    
        c_new_m_invoice.grid(row=0, column =1,  pady = 5)
        view_old.grid(row=1, column =1,  pady = 5)
        c_indiv.grid(row=2, column =1,  pady = 5)
        v_clients.grid(row=3, column =1,  pady = 5)

        # -- right frame ---
        right_frame = ttk.Frame(self, padding = 20, borderwidth = 1, relief = "solid")
        right_frame.grid(row = 1 ,column = 1, sticky ="nsew", padx =10, pady = 10)


        # -- Top Bar ---

        topbar = ttk.Frame(self, padding = (12,8), borderwidth = 1, relief = "solid")
        topbar.grid(row= 0, column = 0, columnspan =2, sticky = "ew")


       





        #-- logo --
        logo_path = Path(__file__).resolve().parent / "baymaxx.png"
        try:
            

            self.logo_img = self.load_logo(logo_path, max_w = 360, max_h = 360)
            ttk.Label(right_frame, image = self.logo_img).grid(row =0, column = 0 , sticky = "se")


        except Exception as e:
            ttk.Label(right_frame, text = "logo missing").grid(row = 1, column = 1, sticky = "se")

        





        # ---status bar---
        statusbar = ttk.Frame(self, padding = (12,6))
        statusbar.grid(row= 0, column= 0, sticky = "n")

        help_label = ttk.Label(statusbar, text = "help").grid(row=0, column = 0, sticky= "nw")
       


   
    def load_logo(self,path, max_w =360, max_h = 360):
        img = Image.open(path)
        img.thumbnail((max_w, max_h), Image.LANCZOS)
        return ImageTk.PhotoImage(img)
    
    def open_clients_manager(self):
        ClientsManager(self)



class ClientsManager(tk.Toplevel):
    """Simple manager window to list/add/edit/delete clients stored in clients.json"""

    def __init__(self, parent: tk.Tk):
        super().__init__(parent)
        self.title("Clients")
        self.geometry("560x380")
        self.transient(parent)     # stay on top of the main window
        self.grab_set()            # modal-ish: focus this window
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        # ---- Tree (list of clients)
        cols = ("name", "offices", "notes")
        self.tree = ttk.Treeview(self, columns=cols, show="headings", height=12)
        self.tree.grid(row=0, column=0, sticky="nsew", padx=12, pady=(12, 6))

        self.tree.heading("name", text="Name")
        self.tree.heading("offices", text="# Offices")
        self.tree.heading("notes", text="Notes")
        self.tree.column("name", width=200, anchor="w")
        self.tree.column("offices", width=80, anchor="center")
        self.tree.column("notes", width=240, anchor="w")

        ybar = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        ybar.grid(row=0, column=1, sticky="ns", pady=(12, 6))
        self.tree.configure(yscrollcommand=ybar.set)

        # Double-click to edit
        self.tree.bind("<Double-1>", lambda e: self.edit_selected())

        # ---- Buttons
        btns = ttk.Frame(self, padding=(12, 6))
        btns.grid(row=1, column=0, columnspan=2, sticky="ew")
        for i in range(5):
            btns.columnconfigure(i, weight=1)

        ttk.Button(btns, text="Add", command=self.add_client).grid(row=0, column=0, sticky="ew", padx=4)
        ttk.Button(btns, text="Edit", command=self.edit_selected).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(btns, text="Delete", command=self.delete_selected).grid(row=0, column=2, sticky="ew", padx=4)
        ttk.Button(btns, text="Refresh", command=self.refresh).grid(row=0, column=3, sticky="ew", padx=4)
        ttk.Button(btns, text="Close", command=self.destroy).grid(row=0, column=4, sticky="ew", padx=4)

        self.refresh()

    # -------- data <-> tree helpers --------

    def refresh(self):
        """Reload clients.json and repopulate the tree."""
        for iid in self.tree.get_children():
            self.tree.delete(iid)

        data = clients.list_clients()
        for c in data:
            name = c.get("name", "")
            notes = c.get("notes", "")
            offices = len(c.get("offices", []))
            # store client_id as the item iid for easy retrieval
            self.tree.insert("", tk.END, iid=c["id"], values=(name, offices, notes))

  

  

def main():
    app = App()
    app.mainloop()



if(__name__ == "__main__"):
    main()