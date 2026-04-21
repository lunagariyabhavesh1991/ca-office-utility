import customtkinter as ctk
from ui.components import DragDropArea, FileListFrame, SmartNamingFrame
from utils.file_manager import FileManager
from core.pdf_engine import PDFEngine
import threading
from ui.theme import Theme

class MergerView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        
        # Layout: Left column for files, Right column for naming/actions
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        # Header
        lbl_hdr = ctk.CTkLabel(self, text="PDF Merger", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        
        # Left Panel (Files)
        left_panel = ctk.CTkFrame(self, fg_color="transparent")
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_panel.grid_rowconfigure(1, weight=1)
        left_panel.grid_columnconfigure(0, weight=1)
        
        self.dnd_area = DragDropArea(left_panel, on_drop_callback=self.on_files_dropped)
        self.dnd_area.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.file_list = FileListFrame(left_panel)
        self.file_list.grid(row=1, column=0, sticky="nsew")
        
        # Right Panel (Settings & Actions)
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        
        self.naming_frame = SmartNamingFrame(right_panel)
        self.naming_frame.pack(fill="x", pady=(0, 20))
        
        self.btn_merge = ctk.CTkButton(right_panel, text="Merge PDFs", height=44, corner_radius=10,
                                       fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                       font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=15, weight="bold"), 
                                       command=self.start_merge)
        self.btn_merge.pack(fill="x", pady=10)
        
        self.progress = ctk.CTkProgressBar(right_panel, progress_color=Theme.ACCENT_BLUE)
        self.progress.pack(fill="x", pady=10)
        self.progress.set(0)
        
    def on_files_dropped(self, files):
        # Filter for PDFs
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if len(pdf_files) < len(files):
            self.app.show_toast("Warning", "Some non-PDF files were ignored.")
        self.file_list.add_files(pdf_files)
        
    def start_merge(self):
        files = self.file_list.get_files()
        if len(files) < 2:
            self.app.show_toast("Error", "Please add at least 2 PDF files.", is_error=True)
            return

        if not self.app.check_operation_allowed():
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
            
        self.progress.set(0)
        # Overwrite check
        import os
        if os.path.exists(output_path):
            if not self.app.confirm("File Exists", f"A file named '{os.path.basename(output_path)}' already exists. Do you want to overwrite it?"):
                return
                
        self.btn_merge.configure(state="disabled", text="Merging...")
        
        # Run in thread
        threading.Thread(target=self._merge_thread, args=(files, output_path), daemon=True).start()
        
    def _merge_thread(self, files, output_path):
        try:
            self.progress.set(0.5)
            PDFEngine.merge_pdfs(files, output_path)
            self.progress.set(1.0)
            
            # Show success back on main thread
            self.after(500, lambda: self._on_merge_success(output_path))
        except Exception as e:
            self.after(0, lambda: self._on_merge_error(str(e)))
            
    def _on_merge_success(self, output_path):
        self.btn_merge.configure(state="normal", text="Merge PDFs")
        self.progress.set(0)
        self.file_list.clear_files()
        self.app.show_toast("Success", f"Merged successfully!\nSaved to:\n{output_path}")
        
    def _on_merge_error(self, error_msg):
        self.btn_merge.configure(state="normal", text="Merge PDFs")
        self.progress.set(0)
        self.app.show_toast("Merge Failed", error_msg, is_error=True)
