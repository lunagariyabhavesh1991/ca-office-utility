import customtkinter as ctk
from ui.theme import Theme
import threading
import time

def _BODY_FONT(size=12, bold=False):
    return ctk.CTkFont(family=Theme.FONT_FAMILY, size=size, weight="bold" if bold else "normal")

_MUTED = Theme.TEXT_MUTED
_ACCENT = Theme.ACCENT_BLUE

class LicenseServerView(ctk.CTkFrame):
    def __init__(self, master, app_window, **kwargs):
        super().__init__(master, fg_color=Theme.BG_PRIMARY, **kwargs)
        self.app = app_window
        self.mgr = getattr(self.app, 'seat_manager', None)
        
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(2, weight=1) 
        
        # Header
        hdr_frame = ctk.CTkFrame(self, fg_color="transparent")
        hdr_frame.pack(fill="x", padx=40, pady=(20, 10))
        
        lbl_hdr = ctk.CTkLabel(hdr_frame, text="Office Network Seats", 
                               font=ctk.CTkFont(family=Theme.FONT_FAMILY, size=24, weight="bold"))
        lbl_hdr.pack(side="left")
        
        self.lbl_status = ctk.CTkLabel(hdr_frame, text="● SCANNING LAN", text_color=_ACCENT,
                                       font=ctk.CTkFont(weight="bold"))
        self.lbl_status.pack(side="right", pady=5)

        # Seat Overview Card
        self.stats_frame = ctk.CTkFrame(self, fg_color=Theme.BG_SECONDARY, corner_radius=10)
        self.stats_frame.pack(fill="x", padx=40, pady=10)
        
        self.lbl_seats = ctk.CTkLabel(self.stats_frame, text="Network Seats: 0 / 0",
                                      font=ctk.CTkFont(size=18, weight="bold"))
        self.lbl_seats.pack(pady=15)
        
        self.lbl_msg = ctk.CTkLabel(self.stats_frame, text="Waiting for discovery...", font=_BODY_FONT(11), text_color=_MUTED)
        self.lbl_msg.pack(pady=(0, 15))

        # Active Users Table
        ctk.CTkLabel(self, text="Running Instances on Local Network", 
                     font=ctk.CTkFont(weight="bold")).pack(anchor="w", padx=45, pady=(10, 5))
        
        self.table_frame = ctk.CTkScrollableFrame(self, fg_color=Theme.BG_SECONDARY, height=250)
        self.table_frame.pack(fill="x", padx=40, pady=5)
        
        # Peer-to-Peer Info Note
        ctk.CTkLabel(self, text="💡 Note: Seat tracking is decentralized. All PCs coordinate automatically via LAN broadcasts.", 
                     font=_BODY_FONT(10), text_color=_MUTED).pack(anchor="w", padx=45, pady=10)

        self.update_loop()

    def update_loop(self):
        if not self.winfo_exists(): return
        
        if self.mgr:
            # Update Status text and color
            color = "#2ecc71" if self.mgr.has_seat else "#e74c3c"
            status = "LICENSE ACTIVE" if self.mgr.has_seat else "SEATS EXCEEDED"
            self.lbl_status.configure(text=f"● {status}", text_color=color)
            
            # Update Seat Count
            count = self.mgr.get_peer_count()
            self.lbl_seats.configure(text=f"Network Usage: {count} / {self.mgr.max_seats}")
            self.lbl_msg.configure(text=self.mgr.status_msg, text_color=color if self.mgr.is_over_limit else "white")
            
            # Update Table
            for widget in self.table_frame.winfo_children():
                widget.destroy()
            
            peers = self.mgr.get_active_list()
            if not peers:
                ctk.CTkLabel(self.table_frame, text="No active connections found on LAN", text_color="gray").pack(pady=20)
            else:
                for mid in peers:
                    row = ctk.CTkFrame(self.table_frame, fg_color=Theme.BG_PRIMARY if mid == self.mgr.machine_id else "transparent", corner_radius=6)
                    row.pack(fill="x", pady=2, padx=5)
                    
                    me_tag = " (This PC)" if mid == self.mgr.machine_id else ""
                    ctk.CTkLabel(row, text=f"💻 {mid}{me_tag}", width=300, anchor="w", 
                                font=_BODY_FONT(12, bold=(mid==self.mgr.machine_id))).pack(side="left", padx=15, pady=8)
                    
                    status_lbl = ctk.CTkLabel(row, text="Active", text_color="#2ecc71", font=_BODY_FONT(11))
                    status_lbl.pack(side="right", padx=20)

        self.after(3000, self.update_loop)
