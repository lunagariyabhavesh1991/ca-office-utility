import customtkinter as ctk
from ui.components import DragDropArea
from core.ocr_engine import OCREngine
import threading
import os
from ui.theme import Theme

class OcrView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        self.target_file = None
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=2)
        self.grid_rowconfigure(1, weight=1)
        
        lbl_hdr = ctk.CTkLabel(self, text="OCR Text Extractor", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 20))
        
        # Left Panel (Input)
        left_panel = ctk.CTkFrame(self, fg_color="transparent")
        left_panel.grid(row=1, column=0, sticky="nsew", padx=(0, 10))
        left_panel.grid_rowconfigure(4, weight=1)
        
        self.dnd_area = DragDropArea(left_panel, title="Drop 1 PDF/Image Here", on_drop_callback=self.on_file_dropped, height=180)
        self.dnd_area.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        
        self.lbl_selected = ctk.CTkLabel(left_panel, text="No file selected", text_color=Theme.TEXT_MUTED, font=(Theme.FONT_FAMILY, 13))
        self.lbl_selected.grid(row=1, column=0, sticky="w", pady=(0, 10))
        
        self.btn_extract = ctk.CTkButton(left_panel, text="Extract Text", height=50, corner_radius=12,
                                         fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                         font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=16, weight="bold"), 
                                         command=self.start_extract)
        self.btn_extract.grid(row=2, column=0, sticky="ew", pady=(10, 0))
        
        self.progress = ctk.CTkProgressBar(left_panel, progress_color=Theme.ACCENT_BLUE)
        self.progress.grid(row=3, column=0, sticky="ew", pady=20)
        self.progress.set(0)
        
        # Right Panel (Output)
        right_panel = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=Theme.CORNER_RADIUS,
                                   border_width=Theme.BORDER_WIDTH, border_color=Theme.BORDER_COLOR)
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        right_panel.grid_rowconfigure(1, weight=1)
        
        top_bar = ctk.CTkFrame(right_panel, fg_color="transparent")
        top_bar.pack(fill="x", padx=10, pady=10)
        
        ctk.CTkLabel(top_bar, text="Extracted Text:", font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold")).pack(side="left")
        
        ctk.CTkButton(top_bar, text="Copy", width=70, height=32, corner_radius=8, command=self.copy_text).pack(side="right", padx=5)
        ctk.CTkButton(top_bar, text="Save As TXT", width=110, height=32, corner_radius=8, command=self.save_text).pack(side="right", padx=5)
        
        self.txt_output = ctk.CTkTextbox(right_panel, font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=13),
                                         fg_color=Theme.BG_PRIMARY, border_color=Theme.BORDER_COLOR, border_width=1)
        self.txt_output.pack(fill="both", expand=True, padx=15, pady=(0, 15))
        
    def on_file_dropped(self, files):
        valid_exts = ('.pdf', '.jpg', '.jpeg', '.png', '.bmp', '.tiff')
        target_files = [f for f in files if f.lower().endswith(valid_exts)]
        if target_files:
            self.target_file = target_files[0]
            if hasattr(self, 'lbl_selected'):
                self.lbl_selected.configure(text=f"Selected: {os.path.basename(self.target_file)}", text_color="white")
        else:
            self.app.show_toast("Error", "Please drop a valid PDF or Image file.", is_error=True)
            
    def start_extract(self):
        if not self.target_file:
            self.app.show_toast("Error", "Please select a file first.", is_error=True)
            return
            
        if not self.app.check_operation_allowed():
            return
            
        self.txt_output.delete("1.0", "end")
        self.txt_output.insert("end", "Initializing OCR Engine... (Please wait, this may take a moment on first run)\n")
        self.progress.set(0)
        self.btn_extract.configure(state="disabled", text="Extracting...")
        threading.Thread(target=self._extract_thread, daemon=True).start()
        
    def _extract_thread(self):
        try:
            def on_progress(pct):
                self.after(0, lambda: self.progress.set(pct / 100))
            
            # This call will trigger the 'ocr_models' download on first run
            ext = os.path.splitext(self.target_file)[1].lower()
            if ext == '.pdf':
                text = OCREngine.extract_text_from_pdf(self.target_file, progress_callback=on_progress)
            else:
                text = OCREngine.extract_text_from_image(self.target_file)
                on_progress(100)
            
            # Update UI with final text
            self.after(0, lambda: self.txt_output.delete("1.0", "end"))
            self.after(0, lambda: self.txt_output.insert("end", text))
            self.after(0, lambda: self._on_success("Extraction complete!"))
                
        except Exception as e:
            self.after(0, lambda err=str(e): self._on_error(err))
            
    def _on_success(self, msg):
        self.progress.stop()
        self.progress.set(1)
        self.btn_extract.configure(state="normal", text="Extract Text")
        
        text = self.txt_output.get("1.0", "end-1c").strip()
        if not text:
            self.app.show_toast("OCR Completed", "No readable text found.", is_error=True)
        else:
            self.app.show_toast("Success", f"{msg}\nText extracted successfully!")
        
    def _on_error(self, error_msg):
        self.progress.stop()
        self.progress.set(0)
        self.btn_extract.configure(state="normal", text="Extract Text")
        self.app.show_toast("Extraction Failed", error_msg, is_error=True)
        
        # Save error log next to target file
        if self.target_file:
            log_path = os.path.join(os.path.dirname(self.target_file), "ocr_error_log.txt")
            try:
                with open(log_path, 'w', encoding='utf-8') as f:
                    f.write(f"OCR Error Log\nTarget: {self.target_file}\nError: {error_msg}\n")
            except Exception:
                pass

    def copy_text(self):
        text = self.txt_output.get("1.0", "end-1c")
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.app.show_toast("Copied", "Text copied to clipboard.")
            
    def save_text(self):
        text = self.txt_output.get("1.0", "end-1c")
        if not text:
            self.app.show_toast("Error", "No text to save.", is_error=True)
            return
            
        import tkinter.filedialog as fd
        f = fd.asksaveasfilename(defaultextension=".txt", 
                                 filetypes=[("Text Files", "*.txt")],
                                 initialfile="Extracted_Text.txt")
        if f:
            with open(f, "w", encoding="utf-8") as out:
                out.write(text)
            self.app.show_toast("Saved", f"Text saved to {os.path.basename(f)}")
