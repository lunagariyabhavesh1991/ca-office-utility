import customtkinter as ctk
from ui.components import DragDropArea, FileListFrame, SmartNamingFrame
from utils.file_manager import FileManager
from core.image_engine import ImageEngine
import threading
import os
from ui.theme import Theme

class ImageToPdfView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        lbl_hdr = ctk.CTkLabel(self, text="Image to PDF Converter", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        
        left_panel = ctk.CTkFrame(self, fg_color="transparent")
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_panel.grid_rowconfigure(1, weight=1)
        left_panel.grid_columnconfigure(0, weight=1)
        
        self.dnd_area = DragDropArea(left_panel, title="Drag & Drop Images (JPG/PNG)", on_drop_callback=self.on_files_dropped)
        self.dnd_area.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.file_list = FileListFrame(left_panel)
        self.file_list.grid(row=1, column=0, sticky="nsew")
        
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        
        self.naming_frame = SmartNamingFrame(right_panel)
        self.naming_frame.pack(fill="x", pady=(0, 20))
        
        self.btn_convert = ctk.CTkButton(right_panel, text="Convert to PDF", height=44, corner_radius=10,
                                         fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                         font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=15, weight="bold"), 
                                         command=self.start_conversion)
        self.btn_convert.pack(fill="x", pady=10)
        
        self.progress = ctk.CTkProgressBar(right_panel, progress_color=Theme.ACCENT_BLUE)
        self.progress.pack(fill="x", pady=10)
        self.progress.set(0)
        
    def on_files_dropped(self, files):
        valid_exts = ('.jpg', '.jpeg', '.png', '.bmp', '.tiff')
        img_files = [f for f in files if f.lower().endswith(valid_exts)]
        if len(img_files) < len(files):
            self.app.show_toast("Warning", "Only image files were added.")
        self.file_list.add_files(img_files)
        
    def start_conversion(self):
        files = self.file_list.get_files()
        if not files:
            self.app.show_toast("Error", "Please add at least 1 image file.", is_error=True)
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
        if os.path.exists(output_path):
            if not self.app.confirm("File Exists", f"A file named '{os.path.basename(output_path)}' already exists. Overwrite?"):
                return

        self.btn_convert.configure(state="disabled", text="Converting...")
        threading.Thread(target=self._convert_thread, args=(files, output_path), daemon=True).start()
        
    def _convert_thread(self, files, output_path):
        try:
            self.progress.set(0.5)
            ImageEngine.images_to_pdf(files, output_path)
            self.progress.set(1.0)
            self.after(500, lambda: self._on_success(output_path))
        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))
            
    def _on_success(self, output_path):
        self.btn_convert.configure(state="normal", text="Convert to PDF")
        self.progress.set(0)
        self.file_list.clear_files()
        self.app.show_toast("Success", f"Converted successfully!\nSaved to:\n{output_path}")
        
    def _on_error(self, error_msg):
        self.btn_convert.configure(state="normal", text="Convert to PDF")
        self.progress.set(0)
        self.app.show_toast("Conversion Failed", error_msg, is_error=True)
