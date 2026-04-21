import os
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk

from ui.theme import Theme
from ui.components import DragDropArea
from core.invoice_parser_engine import InvoiceParserEngine
from services.excel_csv_exporter import ExcelCSVExporter, OutputMode

class InvoiceParserView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app_window = app_window
        
        self.files_to_process = []
        self.file_map = {} # basename -> full_path
        self.extracted_data = [] # List of tuple rows
        self.is_processing = False
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1)

        self._setup_header()
        self._setup_drag_drop()
        self.columns = (
            "Filename", "Page Range", "Status", "Invoice No", "Date", "Party Name", 
            "Buyer GSTIN", "Taxable", 
            "CGST", "CGST %", "SGST", "SGST %", "IGST", "IGST %", 
            "Grand Total", "Confidence", "Remarks"
        )
        self._setup_grid_area()
        self._setup_footer()

    def _setup_header(self):
        header_frame = ctk.CTkFrame(self, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))
        
        title = ctk.CTkLabel(header_frame, text="Smart Invoice Parser", 
                             font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        title.pack(side="left")
        
        desc = ctk.CTkLabel(header_frame, text="Extract party data and accounting totals automatically (For Sales Invoices Only)", 
                            text_color=Theme.TEXT_MUTED, font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=14, weight="bold"))
        desc.pack(side="left", padx=15, pady=(5,0))

    def _setup_drag_drop(self):
        self.dnd_area = DragDropArea(self, on_drop_callback=self._handle_new_files, height=100)
        self.dnd_area.grid(row=1, column=0, sticky="ew", pady=(0, 20), padx=Theme.PADDING)

    def _setup_grid_area(self):
        self.grid_frame = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS)
        self.grid_frame.grid(row=2, column=0, sticky="nsew", padx=Theme.PADDING)
        self.grid_frame.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) # The grid_frame itself is in row 2

        # 1. SETUP NOTEBOOK
        self.notebook = ttk.Notebook(self.grid_frame)
        self.notebook.pack(expand=True, fill="both", padx=5, pady=5)
        
        self.tab_single = ctk.CTkFrame(self.notebook, fg_color=Theme.BG_PRIMARY)
        self.tab_bulk = ctk.CTkFrame(self.notebook, fg_color=Theme.BG_PRIMARY)
        
        self.notebook.add(self.tab_single, text=" Standard (Single Invoice) ")
        self.notebook.add(self.tab_bulk, text=" Bulk (Multi-Invoice PDF Splitter) ")

        self._setup_tree_in_tab(self.tab_single, "single")
        self._setup_tree_in_tab(self.tab_bulk, "bulk")

    def _setup_tree_in_tab(self, parent_frame, tab_id):
        parent_frame.grid_columnconfigure(0, weight=1)
        parent_frame.grid_rowconfigure(0, weight=1)

        style = ttk.Style()
        style.theme_use("default")
        style.configure("Treeview", 
                        background=Theme.BG_PRIMARY,
                        foreground=Theme.TEXT_PRIMARY,
                        rowheight=35,
                        fieldbackground=Theme.BG_PRIMARY,
                        borderwidth=0,
                        font=(Theme.FONT_FAMILY, 10))
        style.map('Treeview', background=[('selected', Theme.ACCENT_BLUE)])
        
        style.configure("TNotebook.Tab", 
                        background=Theme.BG_SECONDARY, 
                        foreground=Theme.TEXT_MUTED,
                        font=(Theme.FONT_FAMILY, 11, "bold"),
                        padding=[20, 10])
        style.map("TNotebook.Tab", 
                  background=[("selected", Theme.ACCENT_BLUE)],
                  foreground=[("selected", Theme.TEXT_PRIMARY)])

        style.configure("Treeview.Heading", 
                        background=Theme.BG_SECONDARY, 
                        foreground=Theme.TEXT_PRIMARY, 
                        font=(Theme.FONT_FAMILY, 11, "bold"),
                        borderwidth=1,
                        relief="flat")
        style.map("Treeview.Heading", background=[('active', Theme.SIDEBAR_HOVER)])

        columns = (
            "Filename", "Page Range", "Status", "Invoice No", "Date", "Party Name", 
            "Buyer GSTIN", "Taxable", 
            "CGST", "CGST %", "SGST", "SGST %", "IGST", "IGST %", 
            "Grand Total", "Confidence", "Remarks"
        )
        
        tree = ttk.Treeview(parent_frame, columns=columns, show="headings", selectmode="extended")
        
        widths = {
            "Filename": 150, "Page Range": 80, "Status": 90, "Invoice No": 100, "Date": 90, "Party Name": 180,
            "Buyer GSTIN": 130, "Taxable": 90, 
            "CGST": 80, "CGST %": 60, "SGST": 80, "SGST %": 60, "IGST": 80, "IGST %": 60,
            "Grand Total": 100, "Confidence": 80, "Remarks": 220
        }
        
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=widths.get(col, 100), anchor="w" if col in ["Filename", "Party Name", "Remarks"] else "center")
            
        scrollbar_y = ttk.Scrollbar(parent_frame, orient="vertical", command=tree.yview)
        scrollbar_x = ttk.Scrollbar(parent_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        tree.grid(row=0, column=0, sticky="nsew", padx=(10, 0), pady=10)
        scrollbar_y.grid(row=0, column=1, sticky="ns", pady=10)
        scrollbar_x.grid(row=1, column=0, sticky="ew", padx=(10, 0))
        
        tree.bind("<Double-1>", self._on_double_click)
        tree.bind("<Delete>", self._on_delete_key)
        
        tree.tag_configure("Failed", background="#4a1515")
        tree.tag_configure("Partial", background="#593c11")
        tree.tag_configure("Needs Review", background="#4d4111")
        tree.tag_configure("Parsed", background="#124021")

        if tab_id == "single":
            self.tree_single = tree
        else:
            self.tree_bulk = tree

    @property
    def current_tree(self):
        """Returns the treeview associated with the active tab."""
        tab_idx = self.notebook.index(self.notebook.select())
        return self.tree_single if tab_idx == 0 else self.tree_bulk

    @property
    def is_bulk_mode(self):
        """Returns True if the Bulk tab is active."""
        return self.notebook.index(self.notebook.select()) == 1

    def _setup_footer(self):
        footer_frame = ctk.CTkFrame(self, fg_color="transparent")
        footer_frame.grid(row=3, column=0, sticky="ew", pady=20, padx=Theme.PADDING)
        
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ctk.CTkProgressBar(footer_frame, variable=self.progress_var, width=300, height=10, 
                                               progress_color=Theme.ACCENT_GREEN)
        self.progress_bar.pack(side="left", padx=(0, 20))
        self.progress_bar.set(0)
        
        self.status_label = ctk.CTkLabel(footer_frame, text="Ready", text_color=Theme.TEXT_MUTED)
        self.status_label.pack(side="left")
        
        self.btn_export = ctk.CTkButton(footer_frame, text="Export to Excel", width=120, 
                                        fg_color=Theme.ACCENT_GREEN, hover_color=Theme.ACTIVATION_HOVER,
                                        command=self._export_to_excel)
        self.btn_export.pack(side="right")
        
        self.btn_clear_all = ctk.CTkButton(footer_frame, text="Clear All", width=80, 
                                           fg_color="transparent", border_width=1, text_color=Theme.TEXT_PRIMARY,
                                           command=self._clear_all)
        self.btn_clear_all.pack(side="right", padx=10)

        self.btn_delete = ctk.CTkButton(footer_frame, text="Delete", width=80, 
                                        fg_color="#8a1a1a", hover_color="#6b1414",
                                        command=self._on_delete_key)
        self.btn_delete.pack(side="right", padx=(0, 10))

        self.btn_extract = ctk.CTkButton(footer_frame, text="Extract Data", width=120, 
                                         fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                         command=self._start_extraction)
        self.btn_extract.pack(side="right")

    def _safe_show_status(self, message, m_type="success"):
        """Safely notifies the app window or updates local status label."""
        try:
            if hasattr(self.app_window, "show_toast"):
                self.app_window.show_toast(m_type.upper(), message)
            elif hasattr(self.app_window, "show_status_message"):
                self.app_window.show_status_message(message, m_type)
            else:
                self.status_label.configure(text=message)
        except:
            self.status_label.configure(text=message)

    def _handle_new_files(self, files):
        if self.is_processing: return
        added = 0
        tree = self.current_tree
        existing_filenames = [tree.set(item, "Filename") for item in tree.get_children()]
        
        columns = (
            "Filename", "Page Range", "Status", "Invoice No", "Date", "Party Name", 
            "Buyer GSTIN", "Taxable", 
            "CGST", "CGST %", "SGST", "SGST %", "IGST", "IGST %", 
            "Grand Total", "Confidence", "Remarks"
        )

        for f in files:
            basename = os.path.basename(f)
            if basename not in existing_filenames:
                ext = os.path.splitext(f)[1].lower()
                if ext in ['.pdf', '.jpg', '.jpeg', '.png', '.bmp']:
                    self.files_to_process.append(f)
                    self.file_map[basename] = f
                    pending_values = ["-"] * len(columns)
                    pending_values[0] = basename
                    pending_values[1] = "-" if not self.is_bulk_mode else "Pending"
                    pending_values[2] = "Pending"
                    pending_values[-1] = "Click Extract Data to begin"
                    tree.insert("", "end", values=tuple(pending_values))
                    added += 1
        if added > 0:
            self.status_label.configure(text=f"Added {added} files to {'Bulk' if self.is_bulk_mode else 'Standard'} tab.")

    def _start_extraction(self):
        tree = self.current_tree
        # Identify files from the current tree that are 'Pending'
        # We need the full paths. We assume files were dropped into self.files_to_process 
        # or we just re-verify from the disk if we know where they are.
        # Actually, let's keep a map of Filename -> FullPath to make this robust.
        # But for now, we'll just process all Pending items in the current tree.
        
        pending_items = [item for item in tree.get_children() if tree.set(item, "Status") == "Pending"]
        if not pending_items:
            messagebox.showinfo("No Files", "Please add files first.")
            return
            
        if self.is_processing: return
        self.is_processing = True
        
        # Capture current mode for the thread
        is_bulk = self.is_bulk_mode
        active_tree = tree
        
        self.btn_extract.configure(state="disabled")
        self.btn_clear_all.configure(state="disabled")
        self.btn_delete.configure(state="disabled")
        self.dnd_area.btn_browse.configure(state="disabled")
        self.progress_var.set(0)
        self.status_label.configure(text=f"Extracting ({'Bulk' if is_bulk else 'Standard'})...")
        
        threading.Thread(target=self._extraction_thread, args=(pending_items, is_bulk, active_tree), daemon=True).start()

    def _extraction_thread(self, pending_items, is_bulk, tree):
        total = len(pending_items)
        # We need a way to get the full path for a filename. 
        # Let's assume the files are in the same directory as when they were dropped.
        # For a truly robust fix, we'd store a map. 
        # But in this app, files are usually in one place.
        
        for idx, item in enumerate(pending_items):
            filename = tree.set(item, "Filename")
            file_path = self.file_map.get(filename)
            
            if not file_path:
                continue

            # CALL ENGINE with is_bulk flag
            results = InvoiceParserEngine.parse_invoice(file_path, detect_multi=is_bulk)
            
            rows = []
            for result in results:
                row_data = (
                    result.get("Filename", filename), result.get("Page Range", "1"), result.get("Status", "Failed"), 
                    result.get("Invoice No", ""), result.get("Date", ""), result.get("Party Name", ""),
                    result.get("Buyer GSTIN", ""), result.get("Taxable Value", ""), result.get("CGST", ""),
                    result.get("CGST %", ""), result.get("SGST", ""), result.get("SGST %", ""),
                    result.get("IGST", ""), result.get("IGST %", ""),
                    result.get("Grand Total", ""), result.get("Confidence", ""), result.get("Remarks", "")
                )
                rows.append(row_data)
            
            self.app_window.after(0, lambda r=rows, it=item, tr=tree, i=idx, t=total: self._update_rows_from_thread(r, it, tr, i, t))
        self.app_window.after(0, self._extraction_complete)

    def _update_rows_from_thread(self, rows, item, tree, idx, total):
        # Insert new rows before removing the pending one
        idx_in_tree = tree.index(item)
        for row in rows:
            tree.insert("", idx_in_tree, values=row, tags=(row[2],))
            idx_in_tree += 1
        tree.delete(item)
        
        self.progress_var.set((idx + 1) / total)
        self.status_label.configure(text=f"Processed {idx + 1} of {total} items...")

    def _extraction_complete(self):
        self.is_processing = False
        self.files_to_process.clear()
        self.btn_extract.configure(state="normal")
        self.btn_clear_all.configure(state="normal")
        self.btn_delete.configure(state="normal")
        self.dnd_area.btn_browse.configure(state="normal")
        self.status_label.configure(text="Extraction complete. Verify individual rows.")
        self.progress_var.set(1.0)
        self._safe_show_status("Invoice parsing completed.", "success")

    def _on_double_click(self, event):
        if self.is_processing: return
        tree = self.current_tree
        region = tree.identify_region(event.x, event.y)
        if region != "cell": return
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if not item or not column: return
        col_idx = int(column[1:]) - 1
        if col_idx in [0, 1, 2, 15]: return # Filename, Page Range, Status, Confidence
        x, y, width, height = tree.bbox(item, column)
        current_value = tree.set(item, column)
        if self.edit_entry: self.edit_entry.destroy()
        self.edit_entry = tk.Entry(tree, font=(Theme.FONT_FAMILY, 10), bg=Theme.BG_SECONDARY, fg=Theme.TEXT_PRIMARY, insertbackground=Theme.TEXT_PRIMARY)
        self.edit_entry.place(x=x, y=y, width=width, height=height)
        self.edit_entry.insert(0, current_value)
        self.edit_entry.focus_set()
        self.edit_entry.bind("<Return>", lambda e, it=item, col=column: self._save_edit(it, col))
        self.edit_entry.bind("<FocusOut>", lambda e, it=item, col=column: self._save_edit(it, col))
        self.edit_entry.bind("<Escape>", lambda e: self.edit_entry.destroy())

    def _save_edit(self, item, column):
        if not self.edit_entry: return
        tree = self.current_tree
        new_value = self.edit_entry.get()
        tree.set(item, column, new_value)
        if tree.set(item, "Status") != "Failed":
             tree.set(item, "Status", "Parsed")
             tree.item(item, tags=("Parsed",))
        self.edit_entry.destroy()
        self.edit_entry = None

    def _on_delete_key(self, event=None):
        if self.is_processing: return
        tree = self.current_tree
        selected_items = tree.selection()
        if not selected_items: return
        
        for item in selected_items:
            # If the file is still in the pending queue, remove it from the list
            filename = tree.set(item, "Filename")
            file_path = self.file_map.pop(filename, None)
            if file_path and file_path in self.files_to_process:
                self.files_to_process.remove(file_path)
            tree.delete(item)
            
        if not tree.get_children():
            self.progress_var.set(0)
            self.status_label.configure(text="Ready")
        else:
            self.status_label.configure(text=f"Removed {len(selected_items)} item(s) from current tab.")

    def _clear_all(self):
        if self.is_processing: return
        tree = self.current_tree
        # Only clear the current tree
        for item in list(tree.get_children()): # Use list() to avoid mutation issues
            filename = tree.set(item, "Filename")
            file_path = self.file_map.pop(filename, None)
            if file_path and file_path in self.files_to_process:
                self.files_to_process.remove(file_path)
            tree.delete(item)
            
        self.progress_var.set(0)
        self.status_label.configure(text=f"Cleared { 'Bulk' if self.is_bulk_mode else 'Standard' } tab.")

    def _export_to_excel(self):
        if self.is_processing: return
        tree = self.current_tree
        items = tree.get_children()
        if not items:
            messagebox.showinfo("No Data", "No data in this tab to export.")
            return
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
            title="Export Invoices", initialfile="Invoice_Parser_Output.xlsx"
        )
        if not output_path: return
        rows = [list(tree.item(item, "values")) for item in items]
        try:
            ExcelCSVExporter.export_to_excel(
                headers=list(self.columns), rows=rows, output_path=output_path,
                mode=OutputMode.PROFESSIONAL, title="Smart Invoice Report", sheet_name="Invoices"
            )
            self._safe_show_status(f"Exported to {os.path.basename(output_path)}", "success")
        except Exception as e:
            messagebox.showerror("Export Error", str(e))
