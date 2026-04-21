import customtkinter as ctk
from ui.components import DragDropArea, FileListFrame, SmartNamingFrame
from utils.file_manager import FileManager
from core.pdf_engine import PDFEngine
import threading
import os
from ui.theme import Theme

class SecurityView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self.app = app_window
        self.pdf_file = None
        
        self.grid_columnconfigure(0, weight=2)
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        
        lbl_hdr = ctk.CTkLabel(self, text="Security & Watermark", font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
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
        self.action_var = ctk.StringVar(value="watermark")
        ctk.CTkRadioButton(tools_frame, text="Add Watermark", variable=self.action_var, value="watermark").grid(row=0, column=0, padx=20, pady=10, sticky="w")
        self.ent_wm = ctk.CTkEntry(tools_frame, placeholder_text="Watermark text (e.g. DRAFT)", width=250)
        self.ent_wm.grid(row=0, column=1, padx=20, pady=10, sticky="w")
        
        ctk.CTkRadioButton(tools_frame, text="Encrypt (Add Password)", variable=self.action_var, value="encrypt").grid(row=1, column=0, padx=20, pady=10, sticky="w")
        self.ent_enc_pw = ctk.CTkEntry(tools_frame, placeholder_text="Password", show="*", width=250)
        self.ent_enc_pw.grid(row=1, column=1, padx=20, pady=10, sticky="w")
        
        ctk.CTkRadioButton(tools_frame, text="Decrypt (Remove Password)", variable=self.action_var, value="decrypt").grid(row=2, column=0, padx=20, pady=10, sticky="w")
        self.ent_dec_pw = ctk.CTkEntry(tools_frame, placeholder_text="Current Password", show="*", width=250)
        self.ent_dec_pw.grid(row=2, column=1, padx=20, pady=10, sticky="w")
        
        right_panel = ctk.CTkFrame(self, fg_color="transparent")
        right_panel.grid(row=1, column=1, sticky="nsew", padx=(10, 0))
        
        self.naming_frame = SmartNamingFrame(right_panel)
        self.naming_frame.pack(fill="x", pady=(0, 20))
        
        self.btn_process = ctk.CTkButton(right_panel, text="Apply Security", height=44, corner_radius=10,
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
            
    def start_process(self):
        if not self.pdf_file:
            self.app.show_toast("Error", "Please select a PDF file first.", is_error=True)
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
            
        action = self.action_var.get()
        param = ""
        if action == "watermark":
            param = self.ent_wm.get().strip()
            if not param:
                self.app.show_toast("Error", "Please enter watermark text.", is_error=True)
                return
        elif action == "encrypt":
            param = self.ent_enc_pw.get()
            if not param:
                self.app.show_toast("Error", "Please enter a password for encryption.", is_error=True)
                return
        elif action == "decrypt":
            param = self.ent_dec_pw.get()
            if not param:
                self.app.show_toast("Error", "Please enter the current password for decryption.", is_error=True)
                return
                
        # Overwrite check
        if os.path.exists(output_path):
            if not self.app.confirm("File Exists", f"A file named '{os.path.basename(output_path)}' already exists. Do you want to overwrite it?"):
                return
                
        self.btn_process.configure(state="disabled", text="Processing...")
        threading.Thread(target=self._process_thread, args=(action, param, output_path), daemon=True).start()
        
    def _process_thread(self, action, param, output_path):
        try:
            if action == "watermark":
                PDFEngine.add_watermark(self.pdf_file, param, output_path)
                msg = "Watermark added successfully!"
            elif action == "encrypt":
                PDFEngine.encrypt_pdf(self.pdf_file, output_path, param)
                msg = "PDF encrypted successfully!"
            elif action == "decrypt":
                success = PDFEngine.decrypt_pdf(self.pdf_file, output_path, param)
                if not success:
                    raise Exception("Incorrect password or file is not encrypted.")
                msg = "Password removed successfully!"
                
            self.after(500, lambda: self._on_success(msg, output_path))
        except Exception as e:
            self.after(0, lambda: self._on_error(str(e)))
            
    def _on_success(self, msg, output_path):
        self.btn_process.configure(state="normal", text="Apply Security")
        self.app.show_toast("Success", f"{msg}\nSaved to:\n{output_path}")
        
    def _on_error(self, error_msg):
        self.btn_process.configure(state="normal", text="Apply Security")
        self.app.show_toast("Process Failed", error_msg, is_error=True)
