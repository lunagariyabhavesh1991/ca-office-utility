import customtkinter as ctk
from ui.components import DragDropArea, FileListFrame, SmartNamingFrame
from utils.file_manager import FileManager
from core.pdf_engine import PDFEngine
from core.image_engine import ImageEngine
import threading
import os
from ui.theme import Theme

class CompressionCenterView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        lbl_hdr = ctk.CTkLabel(self, text="Compression Center", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        
        # Tabs for PDF and Image
        self.tabs = ctk.CTkTabview(self, height=500, segmented_button_selected_color=Theme.ACCENT_BLUE,
                                   segmented_button_selected_hover_color=Theme.ACCENT_HOVER,
                                   text_color=Theme.TEXT_PRIMARY)
        self.tabs.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        self.tab_pdf = self.tabs.add("PDF Compression")
        self.tab_img = self.tabs.add("Image Compression")
        
        # Setup PDF Tab
        self.setup_pdf_tab()
        # Setup Image Tab
        self.setup_img_tab()
        
        # Right Panel (Common Options & Actions) - Made scrollable
        right_panel = ctk.CTkScrollableFrame(self, fg_color="transparent", corner_radius=0)
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        
        # Target Size Options
        target_frame = ctk.CTkFrame(right_panel, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS,
                                    border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR)
        target_frame.pack(fill="x", pady=(0, 20))
        
        ctk.CTkLabel(target_frame, text="Target Compression Size:", font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold")).pack(anchor="w", padx=15, pady=(15, 10))
        
        self.target_var = ctk.StringVar(value="none")
        radio_kwargs = {"font": (Theme.FONT_FAMILY, 12), "hover_color": Theme.ACCENT_BLUE}
        ctk.CTkRadioButton(target_frame, text="Default (Auto Best)", variable=self.target_var, value="none", **radio_kwargs).pack(padx=20, pady=5)
        ctk.CTkRadioButton(target_frame, text="100 KB", variable=self.target_var, value="100", **radio_kwargs).pack(padx=20, pady=5)
        ctk.CTkRadioButton(target_frame, text="200 KB", variable=self.target_var, value="200", **radio_kwargs).pack(padx=20, pady=5)
        ctk.CTkRadioButton(target_frame, text="500 KB", variable=self.target_var, value="500", **radio_kwargs).pack(padx=20, pady=5)
        ctk.CTkRadioButton(target_frame, text="1 MB", variable=self.target_var, value="1024", **radio_kwargs).pack(padx=20, pady=5)
        
        custom_frame = ctk.CTkFrame(target_frame, fg_color="transparent")
        custom_frame.pack(fill="x", padx=20, pady=(5, 10))
        ctk.CTkRadioButton(custom_frame, text="Custom KB:", variable=self.target_var, value="custom").grid(row=0, column=0, sticky="w")
        self.ent_custom = ctk.CTkEntry(custom_frame, placeholder_text="KB", width=80)
        self.ent_custom.grid(row=0, column=1, padx=(10, 0))
        
        self.naming_frame = SmartNamingFrame(right_panel)
        self.naming_frame.pack(fill="x", pady=(0, 20))
        
        self.btn_compress = ctk.CTkButton(right_panel, text="Compress Files", height=44, corner_radius=10,
                                          fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                          font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=15, weight="bold"), 
                                          command=self.start_compress)
        self.btn_compress.pack(fill="x", pady=10)
        
        self.progress = ctk.CTkProgressBar(right_panel, progress_color=Theme.ACCENT_BLUE)
        self.progress.pack(fill="x", pady=10)
        self.progress.set(0)
        
        self.lbl_stats = ctk.CTkLabel(right_panel, text="", text_color="gray70", justify="left")
        self.lbl_stats.pack(fill="x", pady=5)

    def setup_pdf_tab(self):
        self.tab_pdf.grid_columnconfigure(0, weight=1)
        self.tab_pdf.grid_rowconfigure(1, weight=1)
        
        self.dnd_pdf = DragDropArea(self.tab_pdf, title="Drop PDF Files Here", on_drop_callback=self.on_pdf_dropped)
        self.dnd_pdf.grid(row=0, column=0, sticky="ew", pady=(10, 10), padx=10)
        
        self.pdf_list = FileListFrame(self.tab_pdf)
        self.pdf_list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

    def setup_img_tab(self):
        self.tab_img.grid_columnconfigure(0, weight=1)
        self.tab_img.grid_rowconfigure(1, weight=1)
        
        self.dnd_img = DragDropArea(self.tab_img, title="Drop Image Files (JPG/PNG) Here", on_drop_callback=self.on_img_dropped)
        self.dnd_img.grid(row=0, column=0, sticky="ew", pady=(10, 10), padx=10)
        
        self.img_list = FileListFrame(self.tab_img)
        self.img_list.grid(row=1, column=0, sticky="nsew", padx=10, pady=(0, 10))

    def on_pdf_dropped(self, files):
        accepted = [f for f in files if f.lower().endswith('.pdf')]
        if len(accepted) < len(files):
            self.app.show_toast("Warning", "Only PDF files accepted in this tab.")
        self.pdf_list.add_files(accepted)

    def on_img_dropped(self, files):
        accepted = [f for f in files if f.lower().endswith(('.jpg', '.jpeg', '.png'))]
        if len(accepted) < len(files):
            self.app.show_toast("Warning", "Only Images (JPG/PNG) accepted in this tab.")
        self.img_list.add_files(accepted)

    def start_compress(self):
        current_tab = self.tabs.get()
        if current_tab == "PDF Compression":
            files = self.pdf_list.get_files()
            is_pdf_mode = True
        else:
            files = self.img_list.get_files()
            is_pdf_mode = False

        if not files:
            self.app.show_toast("Error", f"Please add at least 1 file to compress in {current_tab}.", is_error=True)
            return

        if not self.app.check_operation_allowed():
            return

        target_val = self.target_var.get()
        target_kb = None
        if target_val == "custom":
            try:
                target_kb = int(self.ent_custom.get())
                if target_kb <= 0:
                    self.app.show_toast("Error", "Custom KB must be a positive number.", is_error=True)
                    return
            except ValueError:
                self.app.show_toast("Error", "Please enter a valid number for custom KB.", is_error=True)
                return
        elif target_val != "none":
            target_kb = int(target_val)
            
        naming_data = self.naming_frame.get_data()
        try:
            base_output_path = FileManager.generate_simple_output_path(
                output_dir=naming_data.get("output_dir", ""),
                output_filename=naming_data.get("output_filename", "")
            )
        except ValueError as e:
            self.app.show_toast("Error", str(e), is_error=True)
            return
            
        out_dir = os.path.dirname(base_output_path)
        out_name_base = os.path.splitext(os.path.basename(base_output_path))[0]
            
        self.progress.set(0)
        self.lbl_stats.configure(text="Compressing...")
        
        # Overwrite check (Simplified: check base path)
        if os.path.exists(base_output_path):
            if not self.app.confirm("File Exists", f"Output file '{os.path.basename(base_output_path)}' already exists. Overwrite?"):
                return

        self.btn_compress.configure(state="disabled", text="Processing...")
        
        threading.Thread(target=self._compress_thread, args=(files, is_pdf_mode, target_val, target_kb, out_dir, out_name_base), daemon=True).start()
        
    def _compress_thread(self, files, is_pdf_mode, target_mode, target_kb, out_dir, out_name_base):
        try:
            total_saved_bytes = 0
            original_bytes = 0
            
            for i, fp in enumerate(files):
                ext = os.path.splitext(fp)[1].lower()
                if len(files) == 1:
                    out_path = os.path.join(out_dir, f"{out_name_base}{ext}")
                else:
                    out_path = os.path.join(out_dir, f"{out_name_base}_{i+1}{ext}")
                    
                orig_size = os.path.getsize(fp)
                original_bytes += orig_size
                
                if is_pdf_mode:
                    def update_prog(p):
                        # p is the overall progress (0.0 to 1.0)
                        self.after(0, lambda: self.progress.set(p))
                        self.after(0, lambda: self.lbl_stats.configure(text=f"Compressing... {int(p*100)}%"))
                        
                    PDFEngine.compress_pdf_target(fp, out_path, mode=target_mode, target_kb=target_kb, 
                                                  progress_callback=update_prog, 
                                                  current_file_idx=i, total_files=len(files))
                else:
                    eff_target_kb = target_kb if target_kb else 500
                    ImageEngine.compress_image(fp, out_path, eff_target_kb)
                    
                if os.path.exists(out_path):
                    new_size = os.path.getsize(out_path)
                    total_saved_bytes += (orig_size - new_size)
                    
                self.progress.set((i + 1) / len(files))
                    
            orig_mb = original_bytes / (1024*1024)
            saved_mb = total_saved_bytes / (1024*1024)
            pct = (total_saved_bytes / original_bytes * 100) if original_bytes else 0
            final_mb = orig_mb - saved_mb
            
            msg = f"Original Size: {orig_mb:.2f} MB\nCompressed: {final_mb:.2f} MB\nSaved: {saved_mb:.2f} MB ({pct:.1f}%)"
            
            self.after(0, lambda: self._on_success(msg, out_dir))
        except Exception as e:
            self.after(0, lambda err=str(e): self._on_error(err))
            
    def _on_success(self, stats_msg, out_dir):
        self.btn_compress.configure(state="normal", text="Compress Files")
        self.progress.set(0)
        if self.tabs.get() == "PDF Compression":
             self.pdf_list.clear_files()
        else:
             self.img_list.clear_files()
        self.lbl_stats.configure(text=stats_msg)
        self.app.show_toast("Compression Successful", f"{stats_msg}\n\nSaved in: {out_dir}")
        
    def _on_error(self, error_msg):
        self.btn_compress.configure(state="normal", text="Compress Files")
        self.progress.set(0)
        self.lbl_stats.configure(text="")
        self.app.show_toast("Compression Failed", error_msg, is_error=True)
