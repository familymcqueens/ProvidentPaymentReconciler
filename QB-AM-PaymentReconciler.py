import csv
import sys
import os
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
from fuzzywuzzy import fuzz

# ==== Constants ====
INTEREST_INCOME_TYPE = 1
PRINCIPAL_INCOME_TYPE = 2
LATE_FEE_INCOME_TYPE = 3
INTEREST_EXPENSE_INCOME_TYPE = 4

# ==== Utility Functions ====
def convert_payment_type_to_string(payment_type):
    if payment_type == INTEREST_INCOME_TYPE:
        return "INTEREST INCOME"
    elif payment_type == PRINCIPAL_INCOME_TYPE:
        return "TOTAL NOTES RECEIVABLE"
    elif payment_type == LATE_FEE_INCOME_TYPE:
        return "LATE_FEE"
    elif payment_type == INTEREST_EXPENSE_INCOME_TYPE:
        return "INTEREST_EXPENSE"
    return "UNKNOWN"

def normalize_date(date_str):
    try:
        parts = date_str.split("/")
        if len(parts) == 3:
            month, day, year = parts
            year = int(year)
            if year < 100:
                year += 2000
            normalized = f"{int(month):02d}/{int(day):02d}/{year}"
            return normalized
    except Exception as e:
        print(f"[DEBUG] normalize_date: failed to normalize '{date_str}' ({e})")
        return date_str
    return date_str

# === NEW: Get date range helper ===
def get_date_range(data):
    if not data:
        return None, None
    dates = sorted([row["date"] for row in data])
    return dates[0], dates[-1]

# ==== File Readers ====
def read_automanager_file(filename):
    AMA = []
    print(f"[DEBUG] read_automanager_file: Opening '{filename}'")
    with open(filename, newline='') as f:
        reader = csv.reader(f)
        row_index = 0
        for row in reader:
            row_index += 1
            if not row or len(row) < 16:
                continue
            date, total, expense, name, payment_type, notes, invoice, check, part, balance, interest, principal, lot, vehicle, status, receipt = row[:16]
            date = normalize_date(date.strip())
            if not date:
                continue
            name_clean = name.replace('"', "").replace(",", "").strip().upper()
            notes_clean = notes.strip().upper()
            try:
                if payment_type == "PAYMENT":
                    if float(interest) != 0:
                        amt = f"{abs(float(interest)):.2f}"
                        AMA.append({"date": date, "name": name_clean, "notes": notes_clean,
                                    "payment_type": INTEREST_INCOME_TYPE, "amount": amt, "match_index": -1})
                    if float(principal) != 0:
                        amt = f"{abs(float(principal)):.2f}"
                        AMA.append({"date": date, "name": name_clean, "notes": notes_clean,
                                    "payment_type": PRINCIPAL_INCOME_TYPE, "amount": amt, "match_index": -1})
                elif payment_type == "LATE FEE":
                    amt = f"{abs(float(total)):.2f}"
                    AMA.append({"date": date, "name": name_clean, "notes": notes_clean,
                                "payment_type": LATE_FEE_INCOME_TYPE, "amount": amt, "match_index": -1})
            except Exception as e:
                print(f"[DEBUG] AMA row {row_index}: error processing amounts ({e}) - skipping")
    print(f"[DEBUG] read_automanager_file: Completed. Total AMA records built: {len(AMA)}")
    return AMA

def read_quickbooks_file(filename):
    QBA = []
    print(f"[DEBUG] read_quickbooks_file: Opening '{filename}'")
    with open(filename, newline='') as f:
        reader = csv.reader(f)
        row_index = 0
        for row in reader:
            row_index += 1
            if len(row) < 6:
                continue
            _, date, name, memo, amount, payment_type = row[:6]
            if date.strip() in ("Date", ""):
                continue
            date_norm = normalize_date(date)
            name_clean = name.replace('"', "").replace(",", "").strip().upper()
            memo_clean = memo.strip().upper()
            payment_type_clean = payment_type.strip().upper()
            try:
                amt = float(amount.replace(",", "").replace('"', ""))
            except Exception as e:
                print(f"[DEBUG] QBA row {row_index}: amount parse failed ({e}) - skipped")
                continue
            income_type = None
            if payment_type_clean == "INTEREST INCOME":
                income_type = INTEREST_INCOME_TYPE
            elif payment_type_clean == "TOTAL NOTES RECEIVABLE":
                income_type = PRINCIPAL_INCOME_TYPE
            elif payment_type_clean in ("LATE FEES", "LATE FEE"):
                income_type = LATE_FEE_INCOME_TYPE
            elif payment_type_clean == "INTEREST EXPENSE":
                income_type = INTEREST_EXPENSE_INCOME_TYPE
            else:
                continue
            amt_str = f"{abs(amt):.2f}"
            QBA.append({"date": date_norm, "name": name_clean, "memo": memo_clean,
                        "amount": amt_str, "payment_type": income_type, "match_index": -1})
    print(f"[DEBUG] read_quickbooks_file: Completed. Total QBA records built: {len(QBA)}")
    return QBA

# ==== Matching Logic ====
def reconcile(AMA, QBA):
    print("[DEBUG] reconcile: Finding exact matches between AutoManager and QuickBooks...")
    ai = 0
    total_attempts = 0
    matches_found = 0
    while ai < len(AMA):
        arow = AMA[ai]
        matched = False
        qi = 0
        while qi < len(QBA):
            qrow = QBA[qi]
            total_attempts += 1
            if (arow["name"] == qrow["name"] and arow["payment_type"] == qrow["payment_type"] and
                    arow["amount"] == qrow["amount"]):
                del AMA[ai]
                del QBA[qi]
                matched = True
                matches_found += 1
                break
            qi += 1
        if not matched:
            ai += 1
    print(f"[DEBUG] reconcile: Completed. total_attempts={total_attempts}, matches_found={matches_found}, "
          f"remaining_AMA={len(AMA)}, remaining_QBA={len(QBA)}")
    return AMA, QBA

# ==== Redirect print to GUI ====
class TextRedirector:
    def __init__(self, widget, tag):
        self.widget = widget
        self.tag = tag
    def write(self, s):
        if not s:
            return
        try:
            self.widget.config(state="normal")
            self.widget.insert(tk.END, s, (self.tag,))
            self.widget.see(tk.END)
            self.widget.config(state="disabled")
        except Exception:
            sys.__stdout__.write(s)
    def flush(self):
        pass

# ==== GUI ====
class ReconcileGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoManager vs QuickBooks Reconciliation")
        self.root.configure(bg="#f0f2f5")
        self.automanager_file = None
        self.quickbooks_file = None
        self.AMA = []
        self.QBA = []

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview.Heading", background="#4B9CD3", foreground="white",
                        font=('Arial', 10, 'bold'))
        style.map("Treeview.Heading", background=[('active', '#5dadec')])

        frame_top = tk.Frame(root, bg="#f0f2f5", pady=12)
        frame_top.pack(fill=tk.X, padx=20)

        self.btn_qb = tk.Button(frame_top, text="Select QuickBooks File", width=25,
                                command=self.load_quickbooks,
                                bg="#4B9CD3", fg="white", font=('Arial', 10, 'bold'))
        self.btn_qb.pack(side="left", padx=10)
        
        self.btn_am = tk.Button(frame_top, text="Select AutoManager File", width=25,
                                command=self.load_automanager,
                                bg="#4B9CD3", fg="white", font=('Arial', 10, 'bold'))
        self.btn_am.pack(side="left", padx=10)

        self.btn_start = tk.Button(frame_top, text="Start Reconciliation", width=25,
                                   command=self.start_reconciliation,
                                   bg="#4CAF50", fg="white", font=('Arial', 10, 'bold'))
        self.btn_start.pack(side="left", padx=10)

        frame_tables = tk.Frame(root, bg="#f0f2f5")
        frame_tables.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # ======================================
        # QuickBooks Table
        # ======================================
        frame_qb = tk.Frame(frame_tables, bg="#f0f2f5")
        frame_qb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        tk.Label(frame_qb, text="QuickBooks Entries (Unmatched)", bg="#f0f2f5",
                 font=('Arial', 12, 'bold')).pack(pady=5)

        # === NEW DATE RANGE LABEL ===
        self.label_qb_dates = tk.Label(frame_qb, text="Date Range: N/A", bg="#f0f2f5",
                                       font=('Arial', 10, 'italic'))
        self.label_qb_dates.pack(pady=(0, 5))

        qb_scrollbar_y = ttk.Scrollbar(frame_qb, orient=tk.VERTICAL)
        qb_scrollbar_x = ttk.Scrollbar(frame_qb, orient=tk.HORIZONTAL)

        self.tree_qb = ttk.Treeview(
            frame_qb,
            columns=("Date", "Name", "Type", "Amount", "Memo"),
            show="headings",
            yscrollcommand=qb_scrollbar_y.set,
            xscrollcommand=qb_scrollbar_x.set
        )
        for col in ("Date", "Name", "Type", "Amount", "Memo"):
            self.tree_qb.heading(col, text=col)
        self.tree_qb.column("Date", width=100, anchor="center")
        self.tree_qb.column("Name", width=200, anchor="w")
        self.tree_qb.column("Type", width=160, anchor="w")
        self.tree_qb.column("Amount", width=100, anchor="center")
        self.tree_qb.column("Memo", width=150, anchor="w")

        qb_scrollbar_y.config(command=self.tree_qb.yview)
        qb_scrollbar_x.config(command=self.tree_qb.xview)
        qb_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        qb_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_qb.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_qb.bind("<Double-1>", self.on_qb_double_click)

        # ======================================
        # AutoManager Table
        # ======================================
        frame_am = tk.Frame(frame_tables, bg="#f0f2f5")
        frame_am.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5)

        tk.Label(frame_am, text="AutoManager Entries (Unmatched)", bg="#f0f2f5",
                 font=('Arial', 12, 'bold')).pack(pady=5)

        # === NEW DATE RANGE LABEL ===
        self.label_am_dates = tk.Label(frame_am, text="Date Range: N/A", bg="#f0f2f5",
                                       font=('Arial', 10, 'italic'))
        self.label_am_dates.pack(pady=(0, 5))

        am_scrollbar_y = ttk.Scrollbar(frame_am, orient=tk.VERTICAL)
        am_scrollbar_x = ttk.Scrollbar(frame_am, orient=tk.HORIZONTAL)

        self.tree_am = ttk.Treeview(
            frame_am,
            columns=("Date", "Name", "Type", "Amount", "Notes"),
            show="headings",
            yscrollcommand=am_scrollbar_y.set,
            xscrollcommand=am_scrollbar_x.set
        )
        for col in ("Date", "Name", "Type", "Amount", "Notes"):
            self.tree_am.heading(col, text=col)
        self.tree_am.column("Date", width=100, anchor="center")
        self.tree_am.column("Name", width=200, anchor="w")
        self.tree_am.column("Type", width=160, anchor="w")
        self.tree_am.column("Amount", width=100, anchor="center")
        self.tree_am.column("Notes", width=150, anchor="w")

        am_scrollbar_y.config(command=self.tree_am.yview)
        am_scrollbar_x.config(command=self.tree_am.xview)
        am_scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
        am_scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree_am.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree_am.bind("<Double-1>", self.on_am_double_click)

        self.tree_am.tag_configure('partial_match', background='yellow')
        self.tree_qb.tag_configure('partial_match', background='yellow')
        
        # Yellow for amount match
        self.tree_am.tag_configure('match_amount', background='#fff2a8')
        self.tree_qb.tag_configure('match_amount', background='#fff2a8')

        # Teal for date match
        self.tree_am.tag_configure('match_date', background='#a8ffe4')
        self.tree_qb.tag_configure('match_date', background='#a8ffe4')

        # Pink for name match
        self.tree_am.tag_configure('match_name', background='#ffc7c7')
        self.tree_qb.tag_configure('match_name', background='#ffc7c7')

        # Console
        self.console = scrolledtext.ScrolledText(
            root, state='disabled', height=13,
            bg="#1e1e1e", fg="#f0f0f0", font=('Consolas', 9)
        )
        self.console.pack(fill=tk.BOTH, expand=False, padx=20, pady=10)
        sys.stdout = TextRedirector(self.console, "stdout")
        sys.stderr = TextRedirector(self.console, "stderr")

    # ==== Populate only one table ====
    def populate_single_table(self, source):
        if source == "AM":
            for item in self.tree_am.get_children():
                self.tree_am.delete(item)
            for row in self.AMA:
                self.tree_am.insert(
                    "", "end",
                    values=(row['date'], row['name'],
                            convert_payment_type_to_string(row['payment_type']),
                            row['amount'], row['notes'])
                )

        elif source == "QB":
            for item in self.tree_qb.get_children():
                self.tree_qb.delete(item)
            for row in self.QBA:
                self.tree_qb.insert(
                    "", "end",
                    values=(row['date'], row['name'],
                            convert_payment_type_to_string(row['payment_type']),
                            row['amount'], row['memo'])
                )

    # ==== Duplicate QBA Check ====
    def check_duplicate_qba_entries(self, QBA):
        seen = {}
        duplicates = []

        for q in QBA:
            key = (q["date"], q["name"], q["amount"], q["payment_type"])
            if key in seen:
                duplicates.append(q)
            else:
                seen[key] = True

        if duplicates:
            msg_lines = ["Duplicate QBA Entries Found:\n"]
            for d in duplicates:
                msg_lines.append(
                    f"Date: {d['date']}, Name: {d['name']}, "
                    f"Type: {convert_payment_type_to_string(d['payment_type'])}, "
                    f"Amount: {d['amount']}"
                )

            msg_lines.append("\nDo you want to proceed with reconciliation?")
            result = messagebox.askyesno("Duplicate QBA Entries Found", "\n".join(msg_lines))

            if not result:
                print("[DEBUG] User chose to stop processing due to duplicates.")
                return False

        return True

    # ==== Updated Load AutoManager ====
    def load_automanager(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if filename:
            self.automanager_file = filename
            print(f"AutoManager file selected: {os.path.basename(filename)}")

            self.AMA = read_automanager_file(filename)
            self.populate_single_table("AM")

            # === Update date range label ===
            start, end = get_date_range(self.AMA)
            if start and end:
                self.label_am_dates.config(text=f"Date Range: {start} → {end}")
            else:
                self.label_am_dates.config(text="Date Range: N/A")

    # ==== Updated Load QuickBooks ====
    def load_quickbooks(self):
        filename = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if filename:
            self.quickbooks_file = filename
            print(f"QuickBooks file selected: {os.path.basename(filename)}")

            self.QBA = read_quickbooks_file(filename)
            self.populate_single_table("QB")

            # === Update date range label ===
            start, end = get_date_range(self.QBA)
            if start and end:
                self.label_qb_dates.config(text=f"Date Range: {start} → {end}")
            else:
                self.label_qb_dates.config(text="Date Range: N/A")

    # ==== Reconciliation ====
    def start_reconciliation(self):
        if not self.automanager_file or not self.quickbooks_file:
            messagebox.showerror("Error", "Please select both AutoManager and QuickBooks files.")
            return

        self.root.state('zoomed')

        def run_reconciliation():
            self.AMA = read_automanager_file(self.automanager_file)
            self.QBA = read_quickbooks_file(self.quickbooks_file)

            if not self.check_duplicate_qba_entries(self.QBA):
                return

            self.AMA, self.QBA = reconcile(self.AMA, self.QBA)
            self.populate_tables()

        threading.Thread(target=run_reconciliation).start()

    def populate_tables(self):
        for item in self.tree_am.get_children():
            self.tree_am.delete(item)
        for item in self.tree_qb.get_children():
            self.tree_qb.delete(item)

        for row in self.AMA:
            self.tree_am.insert("", "end",
                                values=(row['date'], row['name'],
                                        convert_payment_type_to_string(row['payment_type']),
                                        row['amount'], row['notes']))

        for row in self.QBA:
            self.tree_qb.insert("", "end",
                                values=(row['date'], row['name'],
                                        convert_payment_type_to_string(row['payment_type']),
                                        row['amount'], row['memo']))

        self.am_items = {item_id: self.AMA[i] for i, item_id in enumerate(self.tree_am.get_children())}
        self.qb_items = {item_id: self.QBA[i] for i, item_id in enumerate(self.tree_qb.get_children())}

    # ==== Matching Interaction ====
    def on_am_double_click(self, event):
        self.clear_highlights()
        selection = self.tree_am.selection()
        if not selection:
            return
        item_id = selection[0]
        selected_am_entry = self.am_items.get(item_id)
        if selected_am_entry:
            self.highlight_partial_matches(selected_am_entry, self.tree_qb, self.qb_items, "quickbooks")

    def on_qb_double_click(self, event):
        self.clear_highlights()
        selection = self.tree_qb.selection()
        if not selection:
            return
        item_id = selection[0]
        selected_qb_entry = self.qb_items.get(item_id)
        if selected_qb_entry:
            self.highlight_partial_matches(selected_qb_entry, self.tree_am, self.am_items, "automanager")

    def clear_highlights(self):
        for item in self.tree_am.get_children():
            self.tree_am.item(item, tags=())
        for item in self.tree_qb.get_children():
            self.tree_qb.item(item, tags=())

    def highlight_partial_matches(self, selected_entry, target_tree, target_items_dict, source_type):
        # Clear previous styles
        self.clear_highlights()

        selected_name = selected_entry.get("name", "")
        selected_upper = selected_name.upper()

        for item_id, target_entry in target_items_dict.items():

            # Match checks
            amount_match = (selected_entry['amount'] == target_entry['amount'])
            date_match = (selected_entry['date'] == target_entry['date'])

            target_text = (
                target_entry.get("name") or
                target_entry.get("notes") or
                target_entry.get("memo") or
                ""
            )

            fuzzy_score = fuzz.token_sort_ratio(selected_name, target_text)
            name_match = fuzzy_score >= 75

            # Decide which tag applies (priority: amount → date → name)
            if amount_match:
                tag_to_use = "match_amount"
            elif date_match:
                tag_to_use = "match_date"
            elif name_match:
                tag_to_use = "match_name"
            else:
                continue

            # Apply row highlight ONLY — no text changes
            target_tree.item(item_id, tags=(tag_to_use,))

            print(
                f"Row {item_id} partial match ({tag_to_use}): "
                f"amount={amount_match}, date={date_match}, name_score={fuzzy_score}"
            )




def main():
    root = tk.Tk()
    gui = ReconcileGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()
