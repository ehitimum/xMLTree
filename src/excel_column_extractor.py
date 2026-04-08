import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import os
import threading


class ExcelColumnExtractorApp:
    def __init__(self, root):
        self.root = root
        # Only set window properties if root is a Tk or Toplevel window
        if hasattr(self.root, 'title') and callable(self.root.title):
            self.root.title("Excel Column Extractor")
        if hasattr(self.root, 'geometry') and callable(self.root.geometry):
            self.root.geometry("900x700")
        if hasattr(self.root, 'minsize') and callable(self.root.minsize):
            self.root.minsize(700, 500)

        self.file_paths = []
        self.all_columns = []
        self.column_map = {}
        self.all_iids = []
        self.output_dir = "Extracted_Output"
        self._search_after_id = None

        self._build_ui()

    def _build_ui(self):
        self.main_frame = ttk.Frame(self.root, padding="15")
        self.main_frame.pack(expand=True, fill="both")

        title_frame = ttk.Frame(self.main_frame)
        title_frame.pack(fill="x", pady=(0, 10))
        ttk.Label(title_frame, text="Excel Column Extractor", font=("Segoe UI", 18, "bold")).pack(side="left")
        ttk.Label(title_frame, text="v1.2", font=("Segoe UI", 10), foreground="gray").pack(side="left", padx=(10, 0), pady=(6, 0))

        file_frame = ttk.LabelFrame(self.main_frame, text="1. Select Files", padding="10")
        file_frame.pack(fill="x", pady=(0, 10))

        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill="x")

        ttk.Button(btn_frame, text="Browse Files", command=self.select_files).pack(side="left", padx=(0, 10))
        ttk.Button(btn_frame, text="Clear Files", command=self.clear_files).pack(side="left", padx=(0, 10))

        self.lbl_file_count = ttk.Label(btn_frame, text="No files selected", foreground="gray")
        self.lbl_file_count.pack(side="left", padx=(15, 0))

        self.file_listbox = tk.Listbox(file_frame, height=4, selectmode="extended", font=("Segoe UI", 9))
        self.file_listbox.pack(fill="x", pady=(8, 0))
        file_scrollbar = ttk.Scrollbar(self.file_listbox, orient="vertical", command=self.file_listbox.yview)
        file_scrollbar.pack(side="right", fill="y")
        self.file_listbox.config(yscrollcommand=file_scrollbar.set)

        col_frame = ttk.LabelFrame(self.main_frame, text="2. Select Columns to Extract", padding="10")
        col_frame.pack(fill="both", expand=True, pady=(0, 10))

        search_frame = ttk.Frame(col_frame)
        search_frame.pack(fill="x", pady=(0, 8))

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search_change)
        self.search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=40, font=("Segoe UI", 10))
        self.search_entry.pack(side="left", fill="x", expand=True, padx=(0, 10))

        ttk.Button(search_frame, text="Clear Search", command=self._clear_search).pack(side="left")

        sel_btn_frame = ttk.Frame(col_frame)
        sel_btn_frame.pack(fill="x", pady=(0, 5))

        ttk.Button(sel_btn_frame, text="Select All", command=self._select_all_columns).pack(side="left", padx=(0, 5))
        ttk.Button(sel_btn_frame, text="Deselect All", command=self._deselect_all_columns).pack(side="left")

        self.lbl_col_count = ttk.Label(sel_btn_frame, text="0 columns loaded", foreground="gray")
        self.lbl_col_count.pack(side="left", padx=(15, 0))

        self.lbl_selected_count = ttk.Label(sel_btn_frame, text="| 0 selected", foreground="gray")
        self.lbl_selected_count.pack(side="left", padx=(5, 0))

        tree_container = ttk.Frame(col_frame)
        tree_container.pack(fill="both", expand=True)

        self.col_tree = ttk.Treeview(tree_container, columns=("status", "name"), show="tree headings", selectmode="none")
        self.col_tree.heading("#0", text="", anchor="w")
        self.col_tree.heading("status", text="", anchor="center")
        self.col_tree.heading("name", text="Column Name", anchor="w")
        self.col_tree.column("#0", width=0, stretch=False, minwidth=0)
        self.col_tree.column("status", width=40, stretch=False, anchor="center")
        self.col_tree.column("name", width=500)

        vsb = ttk.Scrollbar(tree_container, orient="vertical", command=self.col_tree.yview)
        hsb = ttk.Scrollbar(tree_container, orient="horizontal", command=self.col_tree.xview)
        self.col_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.col_tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")

        self.col_tree.bind("<Button-1>", self._on_tree_click)
        self.col_tree.tag_configure("selected", background="#E8F5E9")

        out_frame = ttk.LabelFrame(self.main_frame, text="3. Generate Output", padding="10")
        out_frame.pack(fill="x", pady=(0, 10))

        out_btn_frame = ttk.Frame(out_frame)
        out_btn_frame.pack(fill="x")

        ttk.Label(out_btn_frame, text="Output folder:").pack(side="left", padx=(0, 5))
        self.lbl_output_dir = ttk.Label(out_btn_frame, text=self.output_dir, foreground="gray")
        self.lbl_output_dir.pack(side="left", padx=(0, 15))

        ttk.Button(out_btn_frame, text="Change Output Folder", command=self.change_output_dir).pack(side="left", padx=(0, 15))

        self.btn_extract = ttk.Button(out_btn_frame, text="Extract Columns", command=self.extract_columns)
        self.btn_extract.pack(side="left")

        self.progress = ttk.Progressbar(self.main_frame, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 5))

        self.lbl_status = ttk.Label(self.main_frame, text="Ready", foreground="gray")
        self.lbl_status.pack()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="Select Excel or CSV Files",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if files:
            self.file_paths = list(files)
            self._update_file_list()
            self._load_columns()

    def clear_files(self):
        self.file_paths = []
        self.all_columns = []
        self.column_map = {}
        self.all_iids = []
        self._update_file_list()
        self._clear_tree()
        self.lbl_file_count.config(text="No files selected")
        self.lbl_col_count.config(text="0 columns loaded")
        self.lbl_selected_count.config(text="| 0 selected")
        self.search_var.set("")

    def _update_file_list(self):
        self.file_listbox.delete(0, tk.END)
        for f in self.file_paths:
            self.file_listbox.insert(tk.END, os.path.basename(f))
        if self.file_paths:
            self.lbl_file_count.config(text=f"{len(self.file_paths)} file(s) selected")

    def _load_columns(self):
        self.all_columns = []
        self.column_map = {}
        self.all_iids = []

        seen = set()
        for path in self.file_paths:
            try:
                if path.lower().endswith('.csv'):
                    df = pd.read_csv(path, nrows=0)
                else:
                    df = pd.read_excel(path, nrows=0)
                for col in df.columns:
                    col_str = str(col).strip()
                    if col_str not in seen:
                        seen.add(col_str)
                        self.all_columns.append(col_str)
            except Exception as e:
                messagebox.showerror("Error", f"Failed to read {os.path.basename(path)}:\n{e}")
                return

        self._populate_tree()
        self.lbl_col_count.config(text=f"{len(self.all_columns)} columns loaded")
        self._update_selected_count()

    def _populate_tree(self):
        self.col_tree.delete(*self.col_tree.get_children())
        self.column_map = {}
        self.all_iids = []

        for i, col in enumerate(self.all_columns):
            iid = f"col_{i}"
            self.col_tree.insert("", "end", iid=iid, text="", values=("[ ]", col))
            self.column_map[iid] = col
            self.all_iids.append(iid)

    def _clear_tree(self):
        self.col_tree.delete(*self.col_tree.get_children())
        self.column_map = {}
        self.all_iids = []

    def _on_tree_click(self, event):
        item = self.col_tree.identify_row(event.y)
        if not item:
            return

        tags = list(self.col_tree.item(item, "tags"))
        if "selected" in tags:
            tags.remove("selected")
            self.col_tree.item(item, tags=(), values=("[ ]", self.column_map[item]))
        else:
            tags.append("selected")
            self.col_tree.item(item, tags=("selected",), values=("[x]", self.column_map[item]))
        self._update_selected_count()

    def _get_selected_columns(self):
        selected = []
        for iid in self.all_iids:
            tags = self.col_tree.item(iid, "tags")
            if tags and "selected" in tags:
                col = self.column_map.get(iid)
                if col:
                    selected.append(col)
        return selected

    def _update_selected_count(self):
        count = len(self._get_selected_columns())
        self.lbl_selected_count.config(text=f"| {count} selected")

    def _on_search_change(self, *args):
        if self._search_after_id:
            self.root.after_cancel(self._search_after_id)
        self._search_after_id = self.root.after(150, self._apply_search)

    def _apply_search(self):
        query = self.search_var.get().strip().lower()
        for iid in self.all_iids:
            col = self.column_map.get(iid, "")
            if not query or query in col.lower():
                self.col_tree.reattach(iid, "", "end")
            else:
                self.col_tree.detach(iid)
        self._search_after_id = None

    def _clear_search(self):
        self.search_var.set("")

    def _select_all_columns(self):
        query = self.search_var.get().strip().lower()
        for iid in self.all_iids:
            col = self.column_map.get(iid, "")
            if not query or query in col.lower():
                self.col_tree.item(iid, tags=("selected",), values=("[x]", col))
        self._update_selected_count()

    def _deselect_all_columns(self):
        query = self.search_var.get().strip().lower()
        for iid in self.all_iids:
            col = self.column_map.get(iid, "")
            if not query or query in col.lower():
                self.col_tree.item(iid, tags=(), values=("[ ]", col))
        self._update_selected_count()

    def change_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Folder")
        if directory:
            self.output_dir = directory
            self.lbl_output_dir.config(text=self.output_dir)

    def extract_columns(self):
        if not self.file_paths:
            messagebox.showwarning("Warning", "Please select files first.")
            return

        selected = self._get_selected_columns()
        if not selected:
            messagebox.showwarning("Warning", "Please select at least one column to extract.")
            return

        self.btn_extract.config(state="disabled")
        self.progress["value"] = 0
        self.lbl_status.config(text="Extracting...")
        threading.Thread(target=self._do_extract, args=(selected,), daemon=True).start()

    def _do_extract(self, selected_columns):
        try:
            if not os.path.exists(self.output_dir):
                os.makedirs(self.output_dir)

            total = len(self.file_paths)
            success_count = 0
            error_files = []

            selected_lower = {c.lower(): c for c in selected_columns}

            for i, path in enumerate(self.file_paths):
                try:
                    if path.lower().endswith('.csv'):
                        df = pd.read_csv(path, usecols=lambda c: str(c).strip().lower() in selected_lower)
                    else:
                        df = pd.read_excel(path)
                        matched = [c for c in df.columns if str(c).strip().lower() in selected_lower]
                        df = df[matched]

                    extracted_df = df.dropna(how="all")
                    base_name = os.path.basename(path)
                    name, ext = os.path.splitext(base_name)
                    save_path = os.path.join(self.output_dir, f"{name}_extracted.xlsx")
                    extracted_df.to_excel(save_path, index=False, engine="openpyxl")
                    success_count += 1

                except Exception as e:
                    error_files.append(f"{os.path.basename(path)}: {str(e)}")

                self.root.after(0, self._update_progress, i + 1, total)

            msg = f"Successfully extracted {success_count} file(s).\nSaved to: {self.output_dir}"
            if error_files:
                msg += f"\n\nErrors ({len(error_files)} file(s)):\n" + "\n".join(error_files[:10])
                if len(error_files) > 10:
                    msg += f"\n...and {len(error_files) - 10} more"

            self.root.after(0, lambda: self._extraction_complete(msg))

        except Exception as e:
            self.root.after(0, lambda: self._extraction_error(str(e)))

    def _update_progress(self, current, total):
        self.progress["value"] = (current / total) * 100
        self.lbl_status.config(text=f"Processing {current}/{total}...")

    def _extraction_complete(self, message):
        self.btn_extract.config(state="normal")
        self.lbl_status.config(text="Done!")
        messagebox.showinfo("Extraction Complete", message)

    def _extraction_error(self, error):
        self.btn_extract.config(state="normal")
        self.lbl_status.config(text="Error occurred")
        messagebox.showerror("Error", f"An error occurred during extraction:\n{error}")


def main():
    root = tk.Tk()
    app = ExcelColumnExtractorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
