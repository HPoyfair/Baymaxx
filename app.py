import tkinter as tk
from tkinter import ttk



class App(tk.Tk):
    def __init__(self):
        super().__init__()

        # --- windo basics ---
        self.title("Baymaxx") #title
        self.geometry("600x400") #window size
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight= 1)

        # ---content frame----
        content = ttk.Frame(self, padding=20)
        content.grid(row = 0, column= 0, sticky ="nsew")

        # -- label and button ---
        self.status = tk.StringVar(value="Ready")
        title = ttk.Label(content, text = "testing", font =("",30, "bold"))
        title.grid(row=10, column =0, sticky="w")

        # --- button ---
        self.counter = 0
        inc_btn = ttk.Button(content, text = "inc", command = self.increment)
        inc_btn.grid(row=0, column =0, sticky= "e")

        # ---status bar---
        statusbar = ttk.Frame(self, padding = (12,6))
        statusbar.grid(row = 1, column = 0, sticky ="ew")
        statusbar.columnconfigure(0, weight = 1)
        ttk.Label(statusbar, textvariable = self.status, anchor ="w").grid(row = 0, column = 0, sticky = "e")


    def increment(self):
        
        self.counter += 1
        self.counter_label.config(text = f"Count: {self.counter}")
        self.status.set ("button clicked")



def main():
    app = App()
    app.mainloop()



if(__name__ == "__main__"):
    main()