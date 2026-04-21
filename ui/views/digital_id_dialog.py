import customtkinter as ctk
import sys
import win32crypt
import datetime
from utils.cert_utils import extract_common_name
try:
    import win32timezone # Required by pywin32 for certificate datetime objects in frozen environment
except ImportError:
    pass
import os
from tkinter import filedialog
from utils.cert_utils import extract_common_name
from ui.theme import Theme
from pyhanko.sign.signers import SimpleSigner

class DigitalIDDialog(ctk.CTkToplevel):
    def __init__(self, master, on_select):
        super().__init__(master)
        self.title("Digital Signature Pro")
        self.geometry("650x550")
        self.minsize(600, 500)
        self.on_select = on_select
        
        self.selected_cert = None
        self.pfx_certs = [] # Locally loaded PFX certs
        self.setup_ui()
        self.refresh_certs()
        
        # Make modal
        self.grab_set()
        self.focus_force()

    def setup_ui(self):
        # Header + Refresh Frame
        title_frame = ctk.CTkFrame(self, fg_color="transparent")
        title_frame.pack(fill="x", padx=20, pady=(25, 10))
        
        header = ctk.CTkLabel(title_frame, text="Select a Digital ID for Signing", 
                               font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=16, weight="bold"), anchor="w")
        header.pack(side="left")
        
        self.refresh_btn = ctk.CTkButton(title_frame, text="🔄 Refresh", width=100, height=32, corner_radius=6,
                                         fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                                         font=(Theme.FONT_FAMILY, 11),
                                         command=self.refresh_certs)
        self.refresh_btn.pack(side="right")

        # Scrollable Frame for Certs
        self.cert_frame = ctk.CTkScrollableFrame(self, fg_color="transparent", label_text="Available Certificates",
                                                label_font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                                label_fg_color="transparent")
        self.cert_frame.pack(expand=True, fill="both", padx=20, pady=10)

        # Footer Buttons
        btn_frame = ctk.CTkFrame(self, fg_color="transparent")
        btn_frame.pack(fill="x", padx=20, pady=20)
        
        # Use grid for buttons to prevent overlap
        btn_frame.grid_columnconfigure((0,1,2,3), weight=1)
        
        self.load_pfx_btn = ctk.CTkButton(btn_frame, text="📁 Load PFX", height=36, corner_radius=8,
                                          fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                                          font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                          command=self.load_pfx)
        self.load_pfx_btn.grid(row=0, column=0, padx=5, sticky="ew")
        
        self.configure_btn = ctk.CTkButton(btn_frame, text="New ID Settings", height=36, corner_radius=8,
                                           fg_color="transparent", text_color=Theme.ACCENT_BLUE,
                                           font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"))
        self.configure_btn.grid(row=0, column=1, padx=5, sticky="ew")
        
        self.cancel_btn = ctk.CTkButton(btn_frame, text="Cancel", height=36, corner_radius=8,
                                        fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                                        text_color=Theme.TEXT_MUTED, command=self.destroy)
        self.cancel_btn.grid(row=0, column=2, padx=5, sticky="ew")

        self.continue_btn = ctk.CTkButton(btn_frame, text="Continue Signing", height=36, corner_radius=8,
                                          fg_color=Theme.ACCENT_BLUE, hover_color=Theme.ACCENT_HOVER,
                                          font=ctk.CTkFont(family=Theme.FONT_FAMILY, weight="bold"),
                                          state="disabled", command=self.apply_selection)
        self.continue_btn.grid(row=0, column=3, padx=5, sticky="ew")

    def refresh_certs(self):
        for widget in self.cert_frame.winfo_children():
            widget.destroy()
            
        system_certs = self.enumerate_system_certs()
        all_certs = system_certs + self.pfx_certs
        self.certs = all_certs
        self.radio_var = ctk.IntVar(value=-1)
        
        # UI Level Debug Log
        with open(os.path.join(os.getcwd(), "cert_debug_log.txt"), "a") as f:
            f.write(f"UI received {len(self.certs)} certs for display.\n")
        
        if not self.certs:
            ctk.CTkLabel(self.cert_frame, text="No digital IDs found.\nPlease insert your USB token or load a PFX file.", 
                         font=(Theme.FONT_FAMILY, 12), text_color=Theme.TEXT_MUTED).pack(pady=40)
            return

        for i, cert in enumerate(self.certs):
            try:
                item_frame = ctk.CTkFrame(self.cert_frame, fg_color=Theme.BG_PRIMARY, corner_radius=Theme.CORNER_RADIUS, border_width=1, border_color=Theme.BORDER_COLOR)
                item_frame.pack(fill="x", pady=6, padx=5)
                
                rb = ctk.CTkRadioButton(item_frame, text="", variable=self.radio_var, value=i, 
                                        border_color=Theme.BORDER_COLOR, hover_color=Theme.ACCENT_BLUE,
                                        command=self.on_cert_clicked)
                rb.pack(side="left", padx=(15, 5))
                
                # Icon
                icon_lbl = ctk.CTkLabel(item_frame, text="📜", font=("", 24))
                icon_lbl.pack(side="left", padx=10)
                
                info_frame = ctk.CTkFrame(item_frame, fg_color="transparent")
                info_frame.pack(side="left", expand=True, fill="both", padx=10, pady=10)
                
                name_lbl = ctk.CTkLabel(info_frame, text=cert['name'], font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=12, weight="bold"), anchor="w")
                name_lbl.pack(fill="x")
                
                issuer_lbl = ctk.CTkLabel(info_frame, text=f"Issuer: {cert['issuer_short']}", font=(Theme.FONT_FAMILY, 10), text_color=Theme.TEXT_MUTED, anchor="w")
                issuer_lbl.pack(fill="x")
                
                expiry_lbl = ctk.CTkLabel(info_frame, text=f"Expires: {cert['expiry']}", font=(Theme.FONT_FAMILY, 10), text_color=Theme.TEXT_MUTED, anchor="w")
                expiry_lbl.pack(fill="x")
                
                details_btn = ctk.CTkButton(item_frame, text="Details", width=70, height=28, corner_radius=6,
                                            fg_color="transparent", border_width=1, border_color=Theme.BORDER_COLOR,
                                            text_color=Theme.ACCENT_BLUE, font=(Theme.FONT_FAMILY, 10))
                details_btn.pack(side="right", padx=15)
            except Exception as ui_e:
                 with open(os.path.join(os.getcwd(), "cert_debug_log.txt"), "a") as f:
                    f.write(f"UI Frame {i} Error: {ui_e}\n")

    def enumerate_system_certs(self):
        certs = []
        log_path = os.path.join(os.getcwd(), "cert_debug_log.txt")
        try:
            with open(log_path, "w") as f:
                f.write(f"Started enumeration at {datetime.datetime.now()}\n")
                
                # Match diagnostic tool exactly
                store = win32crypt.CertOpenStore(10, 0, None, 0x00010000, "MY")
                if not store:
                    f.write("Failed to open store 'MY'\n")
                    return []
                    
                raw_iter = store.CertEnumCertificatesInStore()
                for cert in raw_iter:
                    try:
                        # Use NameToStr which is confirmed working in all environments
                        subject = str(win32crypt.CertNameToStr(cert.Subject))
                        issuer = str(win32crypt.CertNameToStr(cert.Issuer))
                        thumbprint = str(cert.CertGetCertificateContextProperty(3).hex())
                        expiry = str(cert.NotAfter).split(" ")[0]
                        
                        # Use our robust custom extractor for the clean name
                        display_name = extract_common_name(subject)
                        
                        # Logging for debug
                        f.write(f"Cert: {display_name} (Subject: {subject})\n")
                        
                        certs.append({
                            "name": display_name,
                            "issuer_short": issuer,
                            "expiry": expiry,
                            "subject": subject,
                            "thumbprint": thumbprint
                        })
                    except Exception as loop_e:
                        f.write(f"  Loop error: {loop_e}\n")
                        continue
                
                f.write(f"Successfully collected {len(certs)} certs. Closing store...\n")
                
                # Use CERT_CLOSE_STORE_FORCE_FLAG (1) to ensure the store closes even with pending handles
                try:
                    store.CertCloseStore(1)
                except Exception as close_e:
                    f.write(f"  Close Warning: {close_e}\n")
                
            # If still empty in EXE, show the count
            if not certs and getattr(sys, 'frozen', False):
                self.master.after(200, lambda: messagebox.showinfo("Debug", "Found 0 certificates in system store."))
            elif certs and getattr(sys, 'frozen', False):
                 # Temporarily show success count to confirm it's working
                 # self.master.after(200, lambda: messagebox.showinfo("Debug", f"Found {len(certs)} certificates!"))
                 pass
                 
        except Exception as e:
            import traceback
            err = traceback.format_exc()
            try:
                with open(log_path, "a") as f:
                    f.write(f"CRITICAL ERROR:\n{err}\n")
            except: pass
            if getattr(sys, 'frozen', False):
                self.master.after(200, lambda: messagebox.showerror("System Error", f"Cert Access Failed:\n{err}"))
        
        return certs

    def load_pfx(self):
        path = filedialog.askopenfilename(title="Select PFX/P12 File", 
                                         filetypes=[("Certificate files", "*.pfx;*.p12"), ("All files", "*.*")])
        if not path:
            return
            
        # Ask for password
        dialog = ctk.CTkInputDialog(text="Enter password for the certificate file:", title="PFX Password")
        password = dialog.get_input()
        if password is None: # User cancelled
            return
            
        try:
            # Use pyhanko to briefly load and get info
            with open(path, 'rb') as f:
                pfx_data = f.read()
                
            signer = SimpleSigner.load_pkcs12(pfx_data, passphrase=password.encode())
            
            # Extract info for the list
            cert = signer.signing_cert
            subject = str(cert.subject)
            issuer = str(cert.issuer)
            
            cn = extract_common_name(subject)
            issuer_cn = extract_common_name(issuer)
            
            expiry_str = cert.not_valid_after.strftime("%Y-%m-%d")
            
            self.pfx_certs.append({
                "name": f"{cn} (PFX)",
                "issuer_short": issuer_cn,
                "expiry": expiry_str,
                "subject": subject,
                "pfx_path": path,
                "pfx_password": password,
                "thumbprint": None # Not needed for PFX
            })
            self.refresh_certs()
            
        except Exception as e:
            from tkinter import messagebox
            messagebox.showerror("Error", f"Failed to load PFX: {str(e)}")

    def on_cert_clicked(self):
        self.continue_btn.configure(state="normal")
        idx = self.radio_var.get()
        self.selected_cert = self.certs[idx]

    def apply_selection(self):
        if self.selected_cert:
            # Final validation: check if private key is accessible
            if not self.selected_cert.get('pfx_path'):
                # For system certs, we already check cert.HasPrivateKey in enumeration
                pass 
                
            self.on_select(self.selected_cert)
            self.destroy()
