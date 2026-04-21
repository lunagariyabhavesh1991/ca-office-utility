import customtkinter as ctk
from ui.components import DragDropArea, FileListFrame, SmartNamingFrame
from utils.file_manager import FileManager
from core.pdf_engine import PDFEngine
import threading

class CompressorView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        lbl_hdr = ctk.CTkLabel(self, text="PDF Compressor", font=ctk.CTkFont(size=24, weight="bold"))
        lbl_hdr.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        
        left_panel = ctk.CTkFrame(self, fg_color="transparent")
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_panel.grid_rowconfigure(1, weight=1)
        left_panel.grid_columnconfigure(0, weight=1)
        
        self.dnd_area = DragDropArea(left_panel, title="Drag & Drop PDFs to Compress", on_drop_callback=self.on_files_dropped)
        self.dnd_area.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.file_list = FileListFrame(left_panel)
        self.file_list.grid(row=1, column=0, sticky="nsew")
        
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        
        # Target Size Options
        target_frame = ctk.CTkFrame(right_panel)
        target_frame.pack(fill="x", pady=(0, 20))
        
        ctk.CTkLabel(target_frame, text="Target Compression Size:", font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=10, pady=(10, 5))
        
        self.target_var = ctk.StringVar(value="none")
        ctk.CTkRadioButton(target_frame, text="Default (Structural Only)", variable=self.target_var, value="none").pack(anchor="w", padx=20, pady=5)
        ctk.CTkRadioButton(target_frame, text="500 KB (Lossy)", variable=self.target_var, value="500").pack(anchor="w", padx=20, pady=5)
        ctk.CTkRadioButton(target_frame, text="1 MB (Lossy)", variable=self.target_var, value="1024").pack(anchor="w", padx=20, pady=5)
        
        custom_frame = ctk.CTkFrame(target_frame, fg_color="transparent")
        custom_frame.pack(fill="x", padx=20, pady=(5, 10))
        ctk.CTkRadioButton(custom_frame, text="Custom:", variable=self.target_var, value="custom").grid(row=0, column=0, sticky="w")
        self.ent_custom = ctk.CTkEntry(custom_frame, placeholder_text="KB", width=80)
        self.ent_custom.grid(row=0, column=1, padx=(10, 0))
        
        self.naming_frame = SmartNamingFrame(right_panel)
        self.naming_frame.pack(fill="x", pady=(0, 20))
        
        self.btn_compress = ctk.CTkButton(right_panel, text="Compress PDFs", height=40, font=ctk.CTkFont(size=15, weight="bold"), command=self.start_compress)
        self.btn_compress.pack(fill="x", pady=10)
        
        self.progress = ctk.CTkProgressBar(right_panel)
        self.progress.pack(fill="x", pady=10)
        self.progress.set(0)
        
        self.lbl_stats = ctk.CTkLabel(right_panel, text="", text_color="gray70", justify="left")
        self.lbl_stats.pack(fill="x", pady=5)
        
    def on_files_dropped(self, files):
        pdf_files = [f for f in files if f.lower().endswith('.pdf')]
        if len(pdf_files) < len(files):
            self.app.show_toast("Warning", "Only PDF files were added.")
        self.file_list.add_files(pdf_files)
        
    def start_compress(self):
        files = self.file_list.get_files()
        if not files:
            self.app.show_toast("Error", "Please add at least 1 PDF file.", is_error=True)
            return
            
        target_val = self.target_var.get()
        target_kb = None
        if target_val == "custom":
            try:
                target_kb = int(self.ent_custom.get())
            except ValueError:
                self.app.show_toast("Error", "Please enter a valid number for custom KB.", is_error=True)
                return
        elif target_val != "none":
            target_kb = int(target_val)
            
        naming_data = self.naming_frame.get_data()
        try:
            # For multiple files, we'll append a counter to the name if there's more than 1
            base_output_path = FileManager.generate_simple_output_path(
                output_dir=naming_data.get("output_dir", ""),
                output_filename=naming_data.get("output_filename", "")
            )
        except ValueError as e:
            self.app.show_toast("Error", str(e), is_error=True)
            return
            
        self.progress.set(0)
        self.lbl_stats.configure(text="Compressing...")
        self.btn_compress.configure(state="disabled", text="Compressing...")
        threading.Thread(target=self._compress_thread, args=(files, base_output_path, target_kb), daemon=True).start()
        
    def _compress_thread(self, files, base_output_path, target_kb):
        import os
        try:
            total = len(files)
            out_dir = os.path.dirname(base_output_path)
            base_name, ext = os.path.splitext(os.path.basename(base_output_path))
            
            orig_size_total = 0
            new_size_total = 0
            
            for i, f in enumerate(files):
                orig_size_total += os.path.getsize(f)
                
                if total > 1:
                    out_path = os.path.join(out_dir, f"{base_name}_{i+1}{ext}")
                else:
                    out_path = base_output_path
                    
                PDFEngine.compress_pdf_target(f, out_path, target_kb)
                new_size_total += os.path.getsize(out_path)
                
                self.progress.set((i + 1) / total)
                
            stats = f"Original Total: {orig_size_total / 1024:.1f} KB\nCompressed Total: {new_size_total / 1024:.1f} KB"
            self.after(500, lambda: self._on_success(out_dir, stats))
        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))
            
    def _on_success(self, out_dir, stats):
        self.btn_compress.configure(state="normal", text="Compress PDFs")
        self.progress.set(0)
        self.file_list.clear_files()
        self.lbl_stats.configure(text=stats)
        self.app.show_toast("Success", f"All files compressed successfully!\nSaved in:\n{out_dir}")
        
    def _on_error(self, error_msg):
        self.btn_compress.configure(state="normal", text="Compress PDFs")
        self.progress.set(0)
        self.lbl_stats.configure(text="")
        self.app.show_toast("Compression Failed", error_msg, is_error=True)
