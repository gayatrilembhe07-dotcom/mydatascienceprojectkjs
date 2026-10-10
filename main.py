import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from excel_manager import ExcelManager
import os

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory & Stock Management")
        self.root.geometry("1150x700")
        self.root.minsize(1000, 620)

        self.excel = ExcelManager()
        self.selected_id = None

        self.build_ui()
        self.load_products()
        self.refresh_dashboard()

    def build_ui(self):
        header = tk.Frame(self.root, bg="#1f4e78", height=70)
        header.pack(fill="x")
        tk.Label(header, text="INVENTORY & STOCK MANAGEMENT SYSTEM  ",
                 bg="#1f4e78", fg="white", font=("Arial", 20, "bold")).pack(pady=18)

        dashboard = tk.Frame(self.root, padx=12, pady=10)
        dashboard.pack(fill="x")
        self.cards = {}
        for title in ["Products", "Total Stock", "Low Stock", "Inventory Value"]:
            f = tk.Frame(dashboard, relief="ridge", bd=1, padx=20, pady=10)
            f.pack(side="left", expand=True, fill="x", padx=5)
            tk.Label(f, text=title, font=("Arial", 10, "bold")).pack()
            self.cards[title] = tk.Label(f, text="0", font=("Arial", 18, "bold"))
            self.cards[title].pack(pady=4)

        form = tk.LabelFrame(self.root, text="Product Details", padx=10, pady=10)
        form.pack(fill="x", padx=12, pady=5)

        labels = ["Product ID", "Product Name", "Category", "Quantity", "Price", "Supplier", "Reorder Level"]
        self.vars = {x: tk.StringVar() for x in labels}
        for i, label in enumerate(labels):
            r, c = divmod(i, 4)
            tk.Label(form, text=label).grid(row=r*2, column=c, sticky="w", padx=7, pady=(3,0))
            e = ttk.Entry(form, textvariable=self.vars[label], width=24)
            e.grid(row=r*2+1, column=c, padx=7, pady=(0,6), sticky="ew")
        for c in range(4):
            form.columnconfigure(c, weight=1)

        buttons = tk.Frame(self.root, padx=12, pady=5)
        buttons.pack(fill="x")
        ttk.Button(buttons, text="Add Product", command=self.add_product).pack(side="left", padx=4)
        ttk.Button(buttons, text="Update", command=self.update_product).pack(side="left", padx=4)
        ttk.Button(buttons, text="Delete", command=self.delete_product).pack(side="left", padx=4)
        ttk.Button(buttons, text="Clear", command=self.clear_form).pack(side="left", padx=4)
        ttk.Button(buttons, text="Refresh", command=self.load_products).pack(side="left", padx=4)
        ttk.Button(buttons, text="Open Excel File", command=self.open_excel).pack(side="right", padx=4)

        search_frame = tk.Frame(self.root, padx=12, pady=5)
        search_frame.pack(fill="x")
        tk.Label(search_frame, text="Search:").pack(side="left")
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", lambda *_: self.load_products())
        ttk.Entry(search_frame, textvariable=self.search_var, width=35).pack(side="left", padx=8)

        table_frame = tk.Frame(self.root, padx=12, pady=5)
        table_frame.pack(fill="both", expand=True)
        cols = ["Product ID","Product Name","Category","Quantity","Price","Supplier","Reorder Level","Stock Value","Status"]
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        widths = [100,180,130,80,90,160,110,110,100]
        for col, w in zip(cols, widths):
            self.tree.heading(col, text=col)
            self.tree.column(col, width=w, anchor="center")
        yscroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        xscroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        self.tree.pack(side="left", fill="both", expand=True)
        yscroll.pack(side="right", fill="y")
        xscroll.pack(side="bottom", fill="x")
        self.tree.bind("<<TreeviewSelect>>", self.select_product)

    def values(self):
        return {k: v.get().strip() for k,v in self.vars.items()}

    def validate(self):
        v = self.values()
        required = ["Product ID","Product Name","Quantity","Price","Reorder Level"]
        if any(not v[x] for x in required):
            messagebox.showwarning("Validation", "Please fill all required fields.")
            return None
        try:
            v["Quantity"] = int(v["Quantity"])
            v["Price"] = float(v["Price"])
            v["Reorder Level"] = int(v["Reorder Level"])
            if v["Quantity"] < 0 or v["Price"] < 0 or v["Reorder Level"] < 0:
                raise ValueError
        except ValueError:
            messagebox.showerror("Validation", "Quantity and reorder level must be whole numbers; price must be numeric.")
            return None
        return v

    def add_product(self):
        v = self.validate()
        if not v: return
        if self.excel.exists(v["Product ID"]):
            messagebox.showerror("Duplicate", "Product ID already exists.")
            return
        self.excel.add(v)
        self.clear_form()
        self.load_products()
        messagebox.showinfo("Success", "Product added successfully.")

    def update_product(self):
        if not self.selected_id:
            messagebox.showwarning("Select", "Select a product first.")
            return
        v = self.validate()
        if not v: return
        if v["Product ID"] != self.selected_id and self.excel.exists(v["Product ID"]):
            messagebox.showerror("Duplicate", "New Product ID already exists.")
            return
        self.excel.update(self.selected_id, v)
        self.clear_form()
        self.load_products()
        messagebox.showinfo("Success", "Product updated successfully.")

    def delete_product(self):
        if not self.selected_id:
            messagebox.showwarning("Select", "Select a product first.")
            return
        if messagebox.askyesno("Confirm", "Delete selected product?"):
            self.excel.delete(self.selected_id)
            self.clear_form()
            self.load_products()

    def clear_form(self):
        self.selected_id = None
        for v in self.vars.values(): v.set("")
        for item in self.tree.selection(): self.tree.selection_remove(item)

    def select_product(self, event=None):
        sel = self.tree.selection()
        if not sel: return
        vals = self.tree.item(sel[0], "values")
        self.selected_id = vals[0]
        for key, value in zip(["Product ID","Product Name","Category","Quantity","Price","Supplier","Reorder Level"], vals[:7]):
            self.vars[key].set(value)

    def load_products(self):
        query = self.search_var.get().lower().strip() if hasattr(self, "search_var") else ""
        for item in self.tree.get_children():
            self.tree.delete(item)
        for p in self.excel.get_all():
            text = " ".join(str(x).lower() for x in p[:7])
            if query and query not in text:
                continue
            qty, price, reorder = int(p[3]), float(p[4]), int(p[6])
            value = qty * price
            status = "LOW STOCK" if qty <= reorder else "OK"
            self.tree.insert("", "end", values=(p[0],p[1],p[2],qty,f"{price:.2f}",p[5],reorder,f"{value:.2f}",status))
        self.refresh_dashboard()

    def refresh_dashboard(self):
        data = self.excel.get_all()
        total_qty = sum(int(p[3]) for p in data)
        low = sum(1 for p in data if int(p[3]) <= int(p[6]))
        total_value = sum(int(p[3])*float(p[4]) for p in data)
        self.cards["Products"].config(text=str(len(data)))
        self.cards["Total Stock"].config(text=str(total_qty))
        self.cards["Low Stock"].config(text=str(low))
        self.cards["Inventory Value"].config(text=f"₹ {total_value:,.2f}")

    def open_excel(self):
        path = os.path.abspath(self.excel.filename)
        try:
            os.startfile(path)
        except AttributeError:
            import subprocess
            subprocess.Popen(["xdg-open", path])

if __name__ == "__main__":
    root = tk.Tk()
    InventoryApp(root)
    root.mainloop()
