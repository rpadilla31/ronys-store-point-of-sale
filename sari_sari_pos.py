import os
import tkinter as tk
from tkinter import ttk, messagebox
import sqlite3
import datetime as dt
import calendar as cal

DB_NAME = "store_v2.db"

# ------------------ DB SETUP ------------------
conn = sqlite3.connect(DB_NAME)
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS products (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    unit TEXT,
    description TEXT,
    unit_price REAL,
    selling_price REAL,
    income_price REAL,
    quantity INTEGER
)
""")
conn.commit()

def ensure_category_column():
    cursor.execute("PRAGMA table_info(products)")
    cols = [c[1] for c in cursor.fetchall()]
    if "category" not in cols:
        cursor.execute("ALTER TABLE products ADD COLUMN category TEXT")
        conn.commit()

def ensure_expiry_column():
    """Add expiry_date column if missing (stored as TEXT 'YYYY-MM-DD')."""
    cursor.execute("PRAGMA table_info(products)")
    cols = [c[1] for c in cursor.fetchall()]
    if "expiry_date" not in cols:
        cursor.execute("ALTER TABLE products ADD COLUMN expiry_date TEXT")
        conn.commit()

ensure_category_column()
ensure_expiry_column()

cursor.execute("""
CREATE TABLE IF NOT EXISTS sales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    product_id INTEGER,
    description TEXT,
    qty INTEGER,
    price_each REAL,
    total REAL,
    payment REAL,
    change REAL,
    created_at DATETIME DEFAULT CURRENT_TIMESTAMP
)
""")
conn.commit()


# ------------------ MAIN APP ------------------
class BigTabPOS:
    def __init__(self, root):
        self.root = root
        self.root.title("RoNy’s Sari-Sari Store Dashboard")
        self.root.geometry("1200x700")
        self.root.config(bg="#F5DEB3")

        # Configurable threshold for "expiring soon"
        self.expiry_threshold_days = 7
        self._maintenance_notified = False   # avoid multiple popups per open

        # Excel-ish Treeview theme
        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", background="white", foreground="black",
                        rowheight=25, fieldbackground="white", font=("Calibri", 11),
                        bordercolor="#D9D9D9", borderwidth=1, relief="solid")
        style.map("Treeview", background=[("selected", "#cce5ff")])
        style.configure("Treeview.Heading", font=("Calibri", 11, "bold"),
                        background="#D9EAD3", foreground="black", relief="raised")

        # Tabs
        bar = tk.Frame(self.root, bg="#8B0000"); bar.pack(fill="x")
        self.tabs = {
            "SELLING": tk.Button(bar, text="SELLING", font=("Poppins", 14, "bold"),
                                 bg="#8B0000", fg="white", relief="flat",
                                 activebackground="#600000",
                                 command=lambda: self.show_tab("SELLING")),
            "MAINTENANCE": tk.Button(bar, text="MAINTENANCE", font=("Poppins", 14, "bold"),
                                     bg="#8B0000", fg="white", relief="flat",
                                     activebackground="#600000",
                                     command=lambda: self.show_tab("MAINTENANCE")),
            "REPORT": tk.Button(bar, text="REPORT", font=("Poppins", 14, "bold"),
                                bg="#8B0000", fg="white", relief="flat",
                                activebackground="#600000",
                                command=lambda: self.show_tab("REPORT")),
        }
        for b in self.tabs.values():
            b.pack(side="left", fill="x", expand=True, padx=(0,1))

        self.content = tk.Frame(self.root, bg="#F5DEB3"); self.content.pack(fill="both", expand=True)
        self.active_tab = None
        self.show_tab("SELLING")

    def show_tab(self, name):
        self.active_tab = name
        for w in self.content.winfo_children(): w.destroy()
        for n, b in self.tabs.items(): b.config(bg="#8B0000" if n != name else "#A52A2A")
        if name == "SELLING": self.selling_tab()
        elif name == "MAINTENANCE":
            self._maintenance_notified = False  # reset per entry
            self.maintenance_tab()
        else: self.report_tab()

    # ---------- shared ----------
    def get_categories(self):
        """Return distinct non-empty categories only."""
        cursor.execute("""
            SELECT DISTINCT TRIM(category)
            FROM products
            WHERE TRIM(COALESCE(category,'')) <> ''
            ORDER BY 1
        """)
        return [r[0] for r in cursor.fetchall()]

    # ================= SELLING (Quick-Sell) =================
    def selling_tab(self):
        BRAND_BG = "#F5DEB3"; BRAND_DARK = "#8B0000"

        page = tk.Frame(self.content, bg=BRAND_BG); page.pack(fill="both", expand=True, padx=12, pady=10)
        tk.Label(page, text="🛒 SELLING (Quick-Sell)", font=("Poppins", 20, "bold"),
                 bg=BRAND_BG, fg=BRAND_DARK).pack(anchor="w", pady=(0,8))

        top = tk.Frame(page, bg=BRAND_BG); top.pack(fill="x", pady=(0,8))
        tk.Label(top, text="Category:", font=("Poppins", 12, "bold"),
                 bg=BRAND_BG, fg=BRAND_DARK).pack(side="left", padx=(0,8))

        # ---------- NEW: Search bar with auto-suggest ----------
        self._all_cats = self.get_categories()
        self.selected_category = ""  # current filter
        self._cat_suggest_win = None  # toplevel for suggestions

        self.cat_search_var = tk.StringVar(value="")
        self.cat_search = tk.Entry(top, textvariable=self.cat_search_var, font=("Poppins", 12), width=28)
        self.cat_search.pack(side="left")
        # live suggestions + live filter
        self.cat_search.bind("<KeyRelease>", self._on_cat_search)
        self.cat_search.bind("<Down>", self._cat_suggest_focus)
        self.cat_search.bind("<Return>", lambda e: self._apply_category(self.cat_search_var.get()))
        self.cat_search.bind("<Escape>", lambda e: self._clear_category_search())

        ttk.Button(top, text="Refresh", command=self._refresh_cats).pack(side="left", padx=8)

        cols = tk.Frame(page, bg=BRAND_BG); cols.pack(fill="both", expand=True)

        # LEFT: products
        left = tk.Frame(cols, bg=BRAND_BG); left.pack(side="left", fill="both", expand=True)
        tk.Label(left, text="Products (double-click to select)", font=("Poppins", 12, "bold"),
                 bg=BRAND_BG, fg=BRAND_DARK).pack(anchor="w")
        self.prod_tv = ttk.Treeview(left, columns=("ID","Desc","Price","Stock"), show="headings", height=22)
        for c,w in zip(("ID","Desc","Price","Stock"), (60, 520, 130, 90)):
            self.prod_tv.heading(c, text=c, anchor="center")
            self.prod_tv.column(c, width=w, anchor="center")
        self.prod_tv.pack(fill="both", expand=True, pady=(4,0))
        self.prod_tv.tag_configure("oddrow", background="#FFFFFF")
        self.prod_tv.tag_configure("evenrow", background="#F9F9F9")
        self.prod_tv.bind("<Double-1>", self._on_pick_product)
        self.prod_tv.bind("<<TreeviewSelect>>", self._on_pick_product)

        # RIGHT: single-item quick panel
        right = tk.Frame(cols, bg="#FFF3D6", bd=1, relief="solid"); right.pack(side="right", fill="y", padx=(10,0))
        tk.Label(right, text="🧾 SELECTED ITEM", font=("Poppins", 14, "bold"),
                 bg="#8B0000", fg="white").pack(fill="x", pady=(0,6))
        self.sel_name = tk.StringVar(value="—")
        tk.Label(right, textvariable=self.sel_name, font=("Poppins", 12),
                 bg="#FFF3D6").pack(pady=(0,6))

        form = tk.Frame(right, bg="#FFF3D6"); form.pack(padx=14, pady=6, anchor="w")
        self.sel_price = tk.DoubleVar(value=0.0)
        self.sel_stock = tk.IntVar(value=0)
        tk.Label(form, text="Price (₱):", font=("Poppins", 12, "bold"), bg="#FFF3D6", fg=BRAND_DARK)\
            .grid(row=0, column=0, sticky="e", padx=6, pady=4)
        tk.Label(form, textvariable=self.sel_price, font=("Poppins", 12), bg="#FFF3D6")\
            .grid(row=0, column=1, sticky="w", padx=6, pady=4)
        tk.Label(form, text="Stock:", font=("Poppins", 12, "bold"), bg="#FFF3D6", fg=BRAND_DARK)\
            .grid(row=1, column=0, sticky="e", padx=6, pady=4)
        tk.Label(form, textvariable=self.sel_stock, font=("Poppins", 12), bg="#FFF3D6")\
            .grid(row=1, column=1, sticky="w", padx=6, pady=4)

        tk.Label(form, text="Qty:", font=("Poppins", 12, "bold"), bg="#FFF3D6", fg=BRAND_DARK)\
            .grid(row=2, column=0, sticky="e", padx=6, pady=6)
        self.qty_var = tk.StringVar(value="")
        qty = tk.Entry(form, textvariable=self.qty_var, font=("Poppins", 12), width=8, justify="center")
        qty.grid(row=2, column=1, sticky="w", padx=6, pady=6)
        qty.bind("<KeyRelease>", lambda e: self._recompute())

        tk.Label(form, text="Total (₱):", font=("Poppins", 12, "bold"), bg="#FFF3D6", fg=BRAND_DARK)\
            .grid(row=3, column=0, sticky="e", padx=6, pady=6)
        self.total_var = tk.StringVar(value="0.00")
        tk.Label(form, textvariable=self.total_var, font=("Poppins", 16, "bold"),
                 bg="#FFF3D6", fg=BRAND_DARK).grid(row=3, column=1, sticky="w", padx=6, pady=6)

        tk.Label(form, text="Payment (₱):", font=("Poppins", 12, "bold"), bg="#FFF3D6", fg=BRAND_DARK)\
            .grid(row=4, column=0, sticky="e", padx=6, pady=(8,6))
        self.pay_var = tk.StringVar()
        pay = tk.Entry(form, textvariable=self.pay_var, font=("Poppins", 12), width=10, justify="center")
        pay.grid(row=4, column=1, sticky="w", padx=6, pady=(8,6))
        pay.bind("<KeyRelease>", lambda e: self._recompute())

        # quick cash
        qwrap = tk.Frame(right, bg="#FFF3D6"); qwrap.pack(padx=14, pady=(0,6), anchor="w")
        def qbtn(val):
            if val == "EXACT": self.pay_var.set(self.total_var.get())
            else: self.pay_var.set(str(val))
            self._recompute()
        for txt,val in [("Exact","EXACT"),(20,20),(50,50),(100,100),(200,200),(500,500)]:
            ttk.Button(qwrap, text=str(txt), command=lambda v=val: qbtn(v)).pack(side="left", padx=2)

        tk.Label(form, text="Change (₱):", font=("Poppins", 12, "bold"), bg="#FFF3D6", fg=BRAND_DARK)\
            .grid(row=5, column=0, sticky="e", padx=6, pady=(6,10))
        self.change_var = tk.StringVar(value="0.00")
        self.change_lbl = tk.Label(form, textvariable=self.change_var, font=("Poppins", 16, "bold"),
                                   bg="#FFF3D6", fg="#2E7D32")
        self.change_lbl.grid(row=5, column=1, sticky="w", padx=6, pady=(6,10))

        self.confirm_btn = ttk.Button(right, text="✅ Confirm Sale", command=self._confirm_quick_sale)
        self.confirm_btn.pack(pady=(0,12), ipadx=10)
        self.confirm_btn.config(state="disabled")
        self.root.bind("<Return>", lambda e: self.confirm_btn.invoke() if str(self.confirm_btn['state'])=="normal" else None)

        # state
        self.selected = None  # {id, desc, price, stock}
        self._load_products()

    # ---------- Category search helpers (auto-suggest) ----------
    def _on_cat_search(self, _=None):
        """Show suggestions + live filter products as user types."""
        typed = (self.cat_search_var.get() or "").strip()
        # live filter using partial category
        self.selected_category = typed
        self._load_products()

        # Suggestions
        items = self._all_cats
        if typed:
            low = typed.lower()
            items = [c for c in self._all_cats if low in c.lower()]
        self._show_cat_suggest(items)

    def _show_cat_suggest(self, items):
        # hide if no categories
        if not items:
            self._hide_cat_suggest()
            return

        # Create popover window if needed
        if self._cat_suggest_win is None or not self._cat_suggest_win.winfo_exists():
            self._cat_suggest_win = tk.Toplevel(self.root)
            self._cat_suggest_win.overrideredirect(True)
            self._cat_suggest_win.attributes("-topmost", True)
            self._cat_suggest_win.configure(bg="#D9D9D9", padx=1, pady=1)

            self._cat_list = tk.Listbox(self._cat_suggest_win,
                                        font=("Poppins", 11),
                                        activestyle="none",
                                        selectmode="single",
                                        relief="flat")
            self._cat_list.pack(fill="both", expand=True)
            self._cat_list.bind("<ButtonRelease-1>", self._pick_cat_from_suggest)
            self._cat_list.bind("<Return>", self._pick_cat_from_suggest)
            self._cat_list.bind("<Escape>", lambda e: self._hide_cat_suggest())
            self._cat_list.bind("<FocusOut>", lambda e: self._hide_cat_suggest())
            self._cat_list.bind("<Up>", self._cat_list_up)
            self._cat_list.bind("<Down>", self._cat_list_down)

        # Position just under the entry
        try:
            x = self.cat_search.winfo_rootx()
            y = self.cat_search.winfo_rooty() + self.cat_search.winfo_height()
            w = self.cat_search.winfo_width()
        except tk.TclError:
            return

        h_rows = max(1, min(8, len(items)))
        self._cat_suggest_win.geometry(f"{w}x{h_rows*24}+{x}+{y}")

        # Fill with items
        self._cat_list.delete(0, tk.END)
        for c in items:
            self._cat_list.insert(tk.END, c)
        if items:
            self._cat_list.selection_clear(0, tk.END)
            self._cat_list.selection_set(0)
            self._cat_list.activate(0)

        self._cat_suggest_win.deiconify()
        self._cat_suggest_win.update_idletasks()

    def _cat_suggest_focus(self, _=None):
        if self._cat_suggest_win and self._cat_suggest_win.winfo_exists():
            self._cat_list.focus_set()
            return "break"

    def _pick_cat_from_suggest(self, _=None):
        if not (self._cat_suggest_win and self._cat_suggest_win.winfo_exists()):
            return "break"
        sel = self._cat_list.curselection()
        if not sel:
            self._hide_cat_suggest(); return "break"
        cat = self._cat_list.get(sel[0])
        self.cat_search_var.set(cat)
        self._apply_category(cat)
        self._hide_cat_suggest()
        return "break"

    def _cat_list_up(self, _=None):
        sel = self._cat_list.curselection()
        if sel:
            i = max(0, sel[0]-1); self._cat_list.selection_clear(0, tk.END)
            self._cat_list.selection_set(i); self._cat_list.activate(i)
        return "break"

    def _cat_list_down(self, _=None):
        sel = self._cat_list.curselection()
        last = self._cat_list.size()-1
        if sel:
            i = min(last, sel[0]+1)
        else:
            i = 0
        self._cat_list.selection_clear(0, tk.END)
        self._cat_list.selection_set(i); self._cat_list.activate(i)
        return "break"

    def _hide_cat_suggest(self):
        if self._cat_suggest_win and self._cat_suggest_win.winfo_exists():
            self._cat_suggest_win.destroy()
        self._cat_suggest_win = None

    def _apply_category(self, cat):
        self.selected_category = (cat or "").strip()
        self._load_products()

    def _clear_category_search(self):
        self.cat_search_var.set("")
        self.selected_category = ""
        self._hide_cat_suggest()
        self._load_products()

    def _refresh_cats(self):
        self._all_cats = self.get_categories()
        # Refresh suggestions based on current typing
        self._on_cat_search()

    # ---------- product list handlers ----------
    def _load_products(self):
        cat = (self.selected_category or "").strip()
        sql = "SELECT id, description, selling_price, quantity FROM products WHERE 1=1"
        params = []
        if cat:
            sql += " AND TRIM(COALESCE(category,'')) LIKE ?"
            params.append(f"%{cat}%")
        sql += " ORDER BY description COLLATE NOCASE"

        self.prod_tv.delete(*self.prod_tv.get_children())
        cursor.execute(sql, params)
        for i,row in enumerate(cursor.fetchall()):
            self.prod_tv.insert("", "end", values=row, tags=("evenrow" if i%2==0 else "oddrow",))

        # reset quick panel
        self.selected = None
        self.sel_name.set("—"); self.sel_price.set(0.0); self.sel_stock.set(0)
        self.qty_var.set(""); self.total_var.set("0.00"); self.pay_var.set(""); self.change_var.set("0.00")
        self.confirm_btn.config(state="disabled")

    def _on_pick_product(self, _=None):
        sel = self.prod_tv.selection()
        if not sel: return
        pid, desc, price, stock = self.prod_tv.item(sel[0])["values"]
        self.selected = {"id": int(pid), "desc": desc, "price": float(price), "stock": int(stock)}
        self.sel_name.set(desc); self.sel_price.set(float(price)); self.sel_stock.set(int(stock))
        self.qty_var.set("1"); self.pay_var.set(""); self._recompute()

    def _recompute(self):
        if not self.selected:
            self.total_var.set("0.00"); self.change_var.set("0.00"); self.confirm_btn.config(state="disabled"); return
        try:
            q = max(0, int((self.qty_var.get() or "0").strip()))
        except:
            q = 0
        total = round(q * self.selected["price"], 2)
        self.total_var.set(f"{total:.2f}")

        try:
            pay = float((self.pay_var.get() or "0").strip())
        except:
            pay = 0.0
        ch = round(pay - total, 2)
        self.change_var.set(f"{ch:.2f}")
        self.change_lbl.config(fg="#2E7D32" if ch >= 0 else "#C62828")

        ok = (self.selected is not None) and (q>0) and (q<=self.selected["stock"]) and (pay>=total)
        self.confirm_btn.config(state="normal" if ok else "disabled")

    def _confirm_quick_sale(self):
        if not self.selected: return
        try:
            qty = int(self.qty_var.get())
        except:
            messagebox.showerror("Error", "Invalid quantity."); return
        if qty<=0 or qty>self.selected["stock"]:
            messagebox.showerror("Error", "Not enough stock / invalid qty."); return
        try:
            pay = float(self.pay_var.get())
        except:
            messagebox.showerror("Error", "Invalid payment."); return
        total = round(qty * self.selected["price"], 2)
        if pay < total:
            messagebox.showerror("Error", "Kulangi ang bayad."); return

        try:
            # update stock
            new_qty = self.selected["stock"] - qty
            cursor.execute("UPDATE products SET quantity=? WHERE id=?", (new_qty, self.selected["id"]))
            # save sale (explicit CURRENT_TIMESTAMP to guarantee created_at)
            cursor.execute(
                "INSERT INTO sales (product_id, description, qty, price_each, total, payment, change, created_at) "
                "VALUES (?,?,?,?,?,?,?, CURRENT_TIMESTAMP)",
                (self.selected["id"], self.selected["desc"], qty, self.selected["price"], total, pay, pay-total)
            )
            conn.commit()
        except Exception as e:
            messagebox.showerror("Database error", f"Nabigong mag-save:\n{e}")
            return

        messagebox.showinfo("Success", "Sale recorded. Stock updated.")
        # Go directly to report so the user sees the history and income
        self.show_tab("REPORT")

    # ================= MAINTENANCE =================
    def maintenance_tab(self):
        frame = tk.Frame(self.content, bg="#F5DEB3"); frame.pack(fill="both", expand=True)
        tk.Label(frame, text="🧰 PRODUCT MAINTENANCE", font=("Poppins", 20, "bold"),
                 fg="#8B0000", bg="#F5DEB3").pack(pady=15)

        # ---- Notification banner (counts)
        banner = tk.Frame(frame, bg="#FFF9C4", bd=1, relief="solid"); banner.pack(fill="x", padx=20, pady=(0,8))
        tk.Label(banner, text="🔔 Expiration Watch", font=("Poppins", 12, "bold"),
                 bg="#FFF9C4", fg="#8B0000").pack(side="left", padx=(10,8))

        chips = tk.Frame(banner, bg="#FFF9C4"); chips.pack(side="left", pady=6)
        def mk_chip(bg, fg):
            out = tk.Frame(chips, bg=bg)
            out.pack(side="left", padx=5)
            var = tk.StringVar(value="…")
            tk.Label(out, textvariable=var, font=("Calibri", 11, "bold"), bg=bg, fg=fg, padx=8, pady=2).pack()
            return var
        self.notif_labels = {
            "expired": mk_chip("#FFCDD2", "#8B0000"),  # red
            "soon": mk_chip("#FFF4CE", "#7A5C00"),     # amber
            "ok": mk_chip("#C8E6C9", "#1B5E20"),       # green
        }

        form = tk.Frame(frame, bg="#F5DEB3"); form.pack(pady=10)
        # NEW FIELD: Expiration (YYYY-MM-DD)
        fields = [("Category:","category"),("Quantity:","quantity"),("Unit:","unit"),
                  ("Description:","description"),("Original Price:","unit_price"),
                  ("Selling Price:","selling_price"),("Expiration (YYYY-MM-DD):","expiry_date")]
        self.entries = {}
        for i,(lbl,key) in enumerate(fields):
            tk.Label(form, text=lbl, font=("Poppins", 12, "bold"), bg="#F5DEB3", fg="#8B0000")\
              .grid(row=i, column=0, sticky="e", padx=10, pady=5)
            e = tk.Entry(form, font=("Poppins", 12)); e.grid(row=i, column=1, padx=10, pady=5)
            self.entries[key]=e

        btns = tk.Frame(form, bg="#F5DEB3"); btns.grid(row=len(fields), columnspan=2, pady=15)
        tk.Button(btns, text="Add Product", font=("Poppins", 12, "bold"),
                  bg="#8B0000", fg="white", command=self.add_product).pack(side="left", padx=10)
        tk.Button(btns, text="Delete Product", font=("Poppins", 12, "bold"),
                  bg="#8B0000", fg="white", command=self.delete_product).pack(side="left", padx=10)

        # ---- Table: added Expiry + Days Left; color-coding
        cols=("ID","Category","Unit","Description","Original Price","Selling Price","Qty","Expiry","Days Left")
        self.tree = ttk.Treeview(frame, columns=cols, show="headings", height=16)
        self.tree.pack(padx=20, pady=10, fill="both", expand=True)
        for c,w in zip(cols,(70,160,120,440,140,140,80,120,100)):
            self.tree.heading(c, text=c, anchor="center")
            self.tree.column(c, width=w, anchor="center", stretch=True)
        self.tree.tag_configure("oddrow", background="#FFFFFF")
        self.tree.tag_configure("evenrow", background="#F9F9F9")
        self.tree.tag_configure("expired", background="#FFEBEE") # light red
        self.tree.tag_configure("soon", background="#FFF7E0")    # light amber

        self.refresh_table()

        # one-time popup notification when entering maintenance
        self._maybe_notify_expiries()

    def _parse_date(self, s):
        s = (s or "").strip()
        if not s: return None
        try:
            return dt.datetime.strptime(s, "%Y-%m-%d").date()
        except:
            return None

    def add_product(self):
        d={k:v.get().strip() for k,v in self.entries.items()}
        if not d["description"]: return messagebox.showwarning("Warning","Description is required!")
        if not d["selling_price"]: return messagebox.showwarning("Warning","Selling Price is required!")
        if not d["quantity"]: d["quantity"]="0"
        # validate numerics
        try:
            up = float(d["unit_price"]) if d["unit_price"] else 0.0
            sp = float(d["selling_price"]); qty = int(d["quantity"])
        except:
            return messagebox.showerror("Error","Check numeric fields.")
        # validate expiry (optional)
        exp_iso = None
        if d.get("expiry_date"):
            exp = self._parse_date(d["expiry_date"])
            if not exp:
                return messagebox.showerror("Error","Expiration must be YYYY-MM-DD (e.g., 2025-12-31).")
            exp_iso = exp.isoformat()

        cursor.execute("""INSERT INTO products (category, unit, description, unit_price, selling_price, quantity, expiry_date)
                          VALUES (?,?,?,?,?,?,?)""",
                       (d.get("category",""), d.get("unit",""), d["description"], up, sp, qty, exp_iso))
        conn.commit(); self.refresh_table(); messagebox.showinfo("Success","Product added successfully!")

    def delete_product(self):
        sel = self.tree.selection()
        if not sel: return messagebox.showwarning("Warning","Select a product to delete!")
        pid = self.tree.item(sel[0])["values"][0]
        if messagebox.askyesno("Confirm Delete","Delete this product?"):
            cursor.execute("DELETE FROM products WHERE id=?", (pid,))
            conn.commit(); self.refresh_table(); messagebox.showinfo("Deleted","Product deleted successfully!")

    def refresh_table(self):
        """Reload table + update expiry counters and banner."""
        self.tree.delete(*self.tree.get_children())
        cursor.execute("""SELECT id, COALESCE(category,''), unit, description, unit_price, selling_price, quantity, expiry_date
                          FROM products ORDER BY description COLLATE NOCASE""")
        rows = cursor.fetchall()

        today = dt.date.today()
        expired_count = 0
        soon_count = 0
        ok_count = 0

        for i,row in enumerate(rows):
            pid, cat, unit, desc, up, sp, qty, exp_str = row
            exp_dt = self._parse_date(exp_str)
            days_left = ""
            tags = ["evenrow" if i%2==0 else "oddrow"]
            if exp_dt:
                delta = (exp_dt - today).days
                days_left = str(delta)
                if delta < 0:
                    expired_count += 1
                    tags.append("expired")
                elif delta <= self.expiry_threshold_days:
                    soon_count += 1
                    tags.append("soon")
                else:
                    ok_count += 1
            else:
                ok_count += 1

            self.tree.insert("", "end",
                             values=(pid, cat, unit, desc,
                                     f"{float(up or 0):.2f}",
                                     f"{float(sp or 0):.2f}",
                                     qty, (exp_dt.isoformat() if exp_dt else ""),
                                     days_left),
                             tags=tuple(tags))

        # update banner chips if present
        if hasattr(self, "notif_labels") and self.notif_labels:
            self.notif_labels["expired"].set(f"Expired: {expired_count}")
            self.notif_labels["soon"].set(f"Expiring ≤{self.expiry_threshold_days}d: {soon_count}")
            self.notif_labels["ok"].set(f"OK: {ok_count}")

        # store for popup logic
        self._last_expired_count = expired_count
        self._last_soon_count = soon_count

    def _maybe_notify_expiries(self):
        """Show a one-time popup when entering the Maintenance tab."""
        if self._maintenance_notified:
            return
        expired = getattr(self, "_last_expired_count", 0)
        soon = getattr(self, "_last_soon_count", 0)
        if expired > 0 or soon > 0:
            msg = []
            if expired > 0: msg.append(f"{expired} item(s) are EXPIRED.")
            if soon > 0: msg.append(f"{soon} item(s) will expire within {self.expiry_threshold_days} day(s).")
            messagebox.showwarning("Expiration Alerts", "\n".join(msg))
        self._maintenance_notified = True

    # ================= REPORT (date range + live search) =================
    def report_tab(self):
        rep = tk.Frame(self.content, bg="#F5DEB3"); rep.pack(fill="both", expand=True, padx=12, pady=10)
        tk.Label(rep, text="📊 SALES REPORT", font=("Poppins", 20, "bold"),
                 bg="#F5DEB3", fg="#8B0000").pack(anchor="w", pady=(0,10))

        top = tk.Frame(rep, bg="#F5DEB3"); top.pack(fill="x", pady=(0,8))
        self.rep_from_var = tk.StringVar(); self.rep_to_var = tk.StringVar()
        def open_picker():
            today = dt.date.today()
            try: s = dt.datetime.strptime(self.rep_from_var.get(), "%Y-%m-%d").date()
            except: s = today
            try: e = dt.datetime.strptime(self.rep_to_var.get(), "%Y-%m-%d").date()
            except: e = today
            rp = RangePicker(self.root, s, e); self.root.wait_window(rp.win)
            if rp.result:
                s,e = rp.result; self.rep_from_var.set(s.isoformat()); self.rep_to_var.set(e.isoformat())
                date_btn.config(text=f"📅 {s}  →  {e}"); self.load_sales()
        date_btn = ttk.Button(top, text="📅 Date Range", command=open_picker); date_btn.pack(side="left")

        tk.Label(top, text="   Search:", font=("Poppins", 11, "bold"),
                 bg="#F5DEB3", fg="#8B0000").pack(side="left")
        self.rep_kw_var = tk.StringVar()
        ent = tk.Entry(top, textvariable=self.rep_kw_var, font=("Poppins", 11), width=26)
        ent.pack(side="left", padx=(6,0)); ent.bind("<KeyRelease>", lambda e: self.load_sales())

        cols=("ID","When","Description","Qty","Price Each","Total","Payment","Change")
        self.rep_tv = ttk.Treeview(rep, columns=cols, show="headings", height=18); self.rep_tv.pack(fill="both", expand=True)
        for c,w in zip(cols,(60,170,520,80,120,120,120,120)):
            self.rep_tv.heading(c, text=c, anchor="center"); self.rep_tv.column(c, width=w, anchor="center")
        self.rep_tv.tag_configure("oddrow", background="#FFFFFF"); self.rep_tv.tag_configure("evenrow", background="#F9F9F9")

        # summary row (RIGHT-ALIGNED)
        sumrow = tk.Frame(rep, bg="#F5DEB3")
        sumrow.pack(fill="x", pady=(6,0))
        sumrow.grid_columnconfigure(0, weight=1)
        tk.Label(sumrow, text="", bg="#F5DEB3").grid(row=0, column=0, sticky="we")  # spacer
        tk.Label(sumrow, text="Total Income (₱):", font=("Poppins", 12, "bold"),
                 bg="#F5DEB3", fg="#8B0000").grid(row=0, column=1, sticky="e")
        self.rep_total_var = tk.StringVar(value="0.00")
        tk.Label(sumrow, textvariable=self.rep_total_var, font=("Poppins", 12),
                 bg="#F5DEB3").grid(row=0, column=2, sticky="e", padx=(6,12))

        today = dt.date.today()
        self.rep_from_var.set(today.isoformat()); self.rep_to_var.set(today.isoformat())
        date_btn.config(text=f"📅 {today}  →  {today}")
        self.load_sales()

    def load_sales(self):
        sql = """SELECT id, created_at, description, qty, price_each, total, payment, change
                 FROM sales WHERE 1=1"""
        params=[]
        df=(self.rep_from_var.get() or "").strip()
        dt_=(self.rep_to_var.get() or "").strip()
        kw=(self.rep_kw_var.get() or "").strip()
        if df: sql+=" AND DATE(created_at)>=DATE(?)"; params.append(df)
        if dt_: sql+=" AND DATE(created_at)<=DATE(?)"; params.append(dt_)
        if kw: sql+=" AND description LIKE ?"; params.append(f"%{kw}%")
        sql+=" ORDER BY created_at DESC, id DESC"

        self.rep_tv.delete(*self.rep_tv.get_children())
        income=0.0
        cursor.execute(sql, params)
        for i,row in enumerate(cursor.fetchall()):
            rid,when,desc,qty,pe,tot,pay,chg=row
            sale_total = float(tot or 0.0)
            if sale_total == 0.0:
                try:
                    sale_total = float(qty or 0) * float(pe or 0)
                except:
                    sale_total = 0.0
            income += sale_total

            self.rep_tv.insert("", "end",
                               values=(rid,when,desc,qty,
                                       f"{float(pe or 0):.2f}",
                                       f"{sale_total:.2f}",
                                       f"{float(pay or 0):.2f}",
                                       f"{float(chg or 0):.2f}"),
                               tags=("evenrow" if i%2==0 else "oddrow",))
        self.rep_total_var.set(f"{income:,.2f}")


# ---------- Single-month date-range picker ----------
class RangePicker:
    def __init__(self, master, start_date, end_date):
        self.master=master; self.tmp_start=start_date; self.tmp_end=end_date; self._result=None
        self.win=tk.Toplevel(master); self.win.title("Select date range"); self.win.config(bg="#F5DEB3")
        self.win.resizable(False,False); self.win.grab_set()
        wrap=tk.Frame(self.win,bg="#FFFFFF",bd=1,relief="solid"); wrap.pack(padx=10,pady=10)

        left=tk.Frame(wrap,bg="#FFFFFF"); left.grid(row=0,column=0,sticky="ns",padx=(8,8),pady=8)
        def preset(k):
            t=dt.date.today()
            if k=="today": s=e=t
            elif k=="yday": s=e=t-dt.timedelta(days=1)
            elif k=="last7": s=t-dt.timedelta(days=6); e=t
            elif k=="last30": s=t-dt.timedelta(days=29); e=t
            elif k=="thismonth": s=t.replace(day=1); e=(s.replace(month=s.month%12+1,year=s.year+(s.month//12))-dt.timedelta(days=1))
            else:
                first=t.replace(day=1); e=first-dt.timedelta(days=1); s=e.replace(day=1)
            self.tmp_start,self.tmp_end=s,e; self.anchor=s; self._hdr(); self._render()
        for lbl,key in [("Today","today"),("Yesterday","yday"),("Last 7 days","last7"),
                        ("Last 30 days","last30"),("This month","thismonth"),("Last month","lastmonth")]:
            ttk.Button(left,text=lbl,width=18,command=lambda k=key:preset(k)).pack(fill="x",pady=2)

        right=tk.Frame(wrap,bg="#FFFFFF"); right.grid(row=0,column=1,padx=(6,8),pady=8)
        hdr=tk.Frame(right,bg="#FFFFFF"); hdr.pack(fill="x",pady=(0,6))
        self.range_str=tk.StringVar(); tk.Label(hdr,textvariable=self.range_str,bg="#FFFFFF",font=("Poppins",11)).pack(side="left")
        ttk.Button(hdr,text="Clear",command=self._clear).pack(side="right")

        box=tk.Frame(right,bg="#FFFFFF",bd=1,relief="solid"); box.pack()
        nav=tk.Frame(box,bg="#FFFFFF"); nav.pack(fill="x")
        ttk.Button(nav,text="◀",width=3,command=lambda:self._shift(-1)).pack(side="left",padx=4,pady=4)
        self.lbl=tk.StringVar(); tk.Label(nav,textvariable=self.lbl,bg="#FFFFFF",font=("Poppins",11,"bold")).pack(side="left")
        ttk.Button(nav,text="▶",width=3,command=lambda:self._shift(1)).pack(side="left",padx=4)
        self.grid=tk.Frame(box,bg="#FFFFFF"); self.grid.pack(padx=6,pady=6)

        ftr=tk.Frame(right,bg="#FFFFFF"); ftr.pack(fill="x",pady=(8,0))
        ttk.Button(ftr,text="Cancel",command=self._cancel).pack(side="right",padx=4)
        ttk.Button(ftr,text="Apply",command=self._apply).pack(side="right",padx=4)

        self.anchor=self.tmp_start or dt.date.today(); self._hdr(); self._render()

    @property
    def result(self): return self._result
    def _cancel(self): self._result=None; self.win.destroy()
    def _apply(self):
        if not self.tmp_start: self._result=None
        else:
            s=self.tmp_start; e=self.tmp_end or s
            if e<s: s,e=e,s
            self._result=(s,e)
        self.win.destroy()
    def _clear(self): self.tmp_start=None; self.tmp_end=None; self._hdr(); self._render()
    def _shift(self,months):
        y=self.anchor.year+((self.anchor.month-1+months)//12)
        m=((self.anchor.month-1+months)%12)+1
        d=min(self.anchor.day,28)
        self.anchor=dt.date(y,m,d); self._render()
    def _hdr(self):
        if self.tmp_start and self.tmp_end:
            s,e=self.tmp_start,self.tmp_end
            if e<s: s,e=e,s
            self.range_str.set(f"{s}  ~  {e}")
        elif self.tmp_start: self.range_str.set(f"{self.tmp_start}  ~  …")
        else: self.range_str.set("Pick a start date")
    def _render(self):
        for w in self.grid.winfo_children(): w.destroy()
        y,m=self.anchor.year,self.anchor.month
        self.lbl.set(dt.date(y,m,1).strftime("%B %Y"))
        for i,wd in enumerate(["Sun","Mon","Tue","Wed","Thu","Fri","Sat"]):
            tk.Label(self.grid,text=wd,bg="#FFFFFF",width=4,font=("Poppins",9,"bold")).grid(row=0,column=i,padx=2,pady=(0,2))
        calobj=cal.Calendar(firstweekday=6); r=1
        for week in calobj.monthdayscalendar(y,m):
            c=0
            for day in week:
                if day==0:
                    tk.Label(self.grid,text=" ",width=4,bg="#FFFFFF").grid(row=r,column=c,padx=2,pady=2)
                else:
                    dtt=dt.date(y,m,day); bg="#F7F7F7"
                    if self.tmp_start and self.tmp_end:
                        s,e=self.tmp_start,self.tmp_end
                        if e<s: s,e=e,s
                        if s<=dtt<=e: bg="#D7D9FF"
                    elif self.tmp_start and dtt==self.tmp_start:
                        bg="#BFC2FF"
                    tk.Button(self.grid,text=f"{day:02d}",width=4,relief="flat",bg=bg,
                              command=lambda d=dtt:self._pick(d)).grid(row=r,column=c,padx=2,pady=2)
                c+=1
            r+=1
    def _pick(self,d):
        if not self.tmp_start: self.tmp_start=d; self.tmp_end=None
        elif self.tmp_start and not self.tmp_end: self.tmp_end=d
        else: self.tmp_start=d; self.tmp_end=None
        self.anchor=d; self._hdr(); self._render()


# ------------------ RUN (GUI locally, CI/headless on GitHub Actions) ------------------
if __name__ == "__main__":
    # If running in GitHub Actions (no GUI), just print a quick status and exit
    if os.environ.get("CI") == "true":
        print("✅ CI mode: App loaded (GUI not started). Database:", DB_NAME)
        # optional: do a quick DB check
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table'")
        print("✅ Tables:", [r[0] for r in cursor.fetchall()])
    else:
        root = tk.Tk()
        BigTabPOS(root)
        root.mainloop()
