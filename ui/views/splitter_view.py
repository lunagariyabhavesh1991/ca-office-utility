import customtkinter as ctk
from ui.components import DragDropArea, SmartNamingFrame
from utils.file_manager import FileManager
from core.pdf_engine import PDFEngine
import threading
import os
from ui.theme import Theme

class SplitterView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        self.pdf_file = None
        
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        lbl_hdr = ctk.CTkLabel(self, text="PDF Splitter & Tools", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        
        left_panel = ctk.CTkFrame(self, fg_color="transparent")
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_panel.grid_rowconfigure(2, weight=1)
        left_panel.grid_columnconfigure(0, weight=1)
        
        self.dnd_area = DragDropArea(left_panel, title="Drop 1 PDF Here", on_drop_callback=self.on_file_dropped, height=120)
        self.dnd_area.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.lbl_selected = ctk.CTkLabel(left_panel, text="No PDF selected", text_color=Theme.TEXT_MUTED, font=(Theme.FONT_FAMILY, 13))
        self.lbl_selected.grid(row=1, column=0, sticky="w", pady=(0, 10))
        
        tools_frame = ctk.CTkFrame(left_panel, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS,
                                   border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR)
        tools_frame.grid(row=2, column=0, sticky="nsew")
        
        # Action selector
        self.action_var = ctk.StringVar(value="split_indiv")
        ctk.CTkRadioButton(tools_frame, text="Split into individual pages", variable=self.action_var, value="split_indiv").pack(anchor="w", padx=20, pady=(20, 10))
        
        ctk.CTkRadioButton(tools_frame, text="Extract specific pages", variable=self.action_var, value="extract").pack(anchor="w", padx=20, pady=10)
        
        ctk.CTkRadioButton(tools_frame, text="Remove specific pages", variable=self.action_var, value="remove").pack(anchor="w", padx=20, pady=10)
        
        self.ent_pages = ctk.CTkEntry(tools_frame, placeholder_text="e.g. 1, 3, 5-10", width=250)
        self.ent_pages.pack(anchor="w", padx=40, pady=(0, 10))
        
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        
        self.naming_frame = SmartNamingFrame(right_panel)
        self.naming_frame.pack(fill="x", pady=(0, 20))
        
        self.btn_process = ctk.CTkButton(right_panel, text="Process PDF", height=44, corner_radius=10,
                                         fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                         font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=15, weight="bold"), 
                                         command=self.start_process)
        self.btn_process.pack(fill="x", pady=10)
        
    def on_file_dropped(self, files):
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if pdf_files:
            self.pdf_file = pdf_files[0]
            self.lbl_selected.configure(text=f"Selected: {os.path.basename(self.pdf_file)}", text_color="white")
        else:
            self.app.show_toast("Error", "Please drop a valid PDF file.", is_error=True)
            
    def _parse_pages(self, page_str):
        pages = set()
        for part in page_str.split(','):
            part = part.strip()
            if not part: continue
            if '-' in part:
                 try:
                    start_str, end_str = part.split('-')
                    start, end = int(start_str), int(end_str)
                    pages.update(range(start, end + 1))
                 except ValueError: pass
            else:
                try: pages.add(int(part))
                except ValueError: pass
        return sorted(list(pages))
            
    def start_process(self):
        if not self.pdf_file:
            self.app.show_toast("Error", "Please select a PDF file first.", is_error=True)
            return

        naming_data = self.naming_frame.get_data()
        try:
            output_path = FileManager.generate_simple_output_path(
                output_dir=naming_data.get("output_dir", ""),
                output_filename=naming_data.get("output_filename", "")
            )
        except ValueError as e:
            self.app.show_toast("Error", str(e), is_error=True)
            return
            
        action = self.action_var.get()
        output_dir = os.path.dirname(output_path)
        
        # Overwrite check (Only for single file output modes like Extract or Remove)
        if action in ("extract", "remove") and os.path.exists(output_path):
            if not self.app.confirm("File Exists", f"A file named '{os.path.basename(output_path)}' already exists. Overwrite?"):
                return

        self.btn_process.configure(state="disabled", text="Processing...")
        threading.Thread(target=self._process_thread, args=(action, output_path, output_dir), daemon=True).start()
        
    def _process_thread(self, action, output_path, output_dir):
        try:
            if action == "split_indiv":
                PDFEngine.split_pdf(self.pdf_file, output_dir)
                msg = f"PDF split into individual pages in:\n{output_dir}"
            elif action == "extract":
                pages_str = self.ent_pages.get()
                pages = self._parse_pages(pages_str)
                PDFEngine.extract_pages(self.pdf_file, output_path, pages)
                msg = f"Pages extracted successfully to:\n{output_path}"
            elif action == "remove":
                pages_str = self.ent_pages.get()
                pages = self._parse_pages(pages_str)
                PDFEngine.remove_pages(self.pdf_file, output_path, pages)
                msg = f"Pages removed successfully. Saved to:\n{output_path}"
                
            self.after(500, lambda: self._on_success(msg))
        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))
            
    def _on_success(self, msg):
        self.btn_process.configure(state="normal", text="Process PDF")
        self.app.show_toast("Success", msg)
        
    def _on_error(self, error_msg):
        self.btn_process.configure(state="normal", text="Process PDF")
        self.app.show_toast("Process Failed", error_msg, is_error=True)
