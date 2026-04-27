import hashlib
import os
import subprocess
import json
import socket
import threading
import time
from datetime import datetime, timedelta

class LicenseManager:
    # SECRET SALT - DO NOT SHARE THIS
    _SALT = "BKL_CA_OFFICE_2024_SECRET_99"
    _TRIAL_DAYS = 30
    
    # Network Ports
    UDP_PORT = 5005
    TCP_PORT = 5006
    
    @staticmethod
    def get_machine_id():
        """Generates a unique Machine ID based on the hardware UUID."""
        try:
            # Try to get system UUID using wmic
            cmd = 'wmic csproduct get uuid'
            uuid_str = subprocess.check_output(cmd, shell=True).decode().split('\n')[1].strip()
            if not uuid_str or "UUID" in uuid_str:
                # Fallback to mac address if wmic fails
                import uuid
                uuid_str = str(uuid.getnode())
        except Exception:
            import uuid
            uuid_str = str(uuid.getnode())
            
        # Create a shorter, readable hash of the hardware ID
        short_hash = hashlib.sha256(uuid_str.encode()).hexdigest()[:12].upper()
        # Format as BKL-XXXX-XXXX
        return f"BKL-{short_hash[:4]}-{short_hash[4:8]}-{short_hash[8:12]}"

    @staticmethod
    def generate_valid_key(machine_id, expiry_date_str, member_no=None, seats=1):
        """
        Generates the correct license key for a given Machine ID and Expiry Date.
        Supports both Standalone (KEY-) and Office (OFFICE-) formats.
        """
        member_part = str(member_no).strip().upper() if member_no else ""
        seats_part = str(seats)
        combined = f"{machine_id}{member_part}{expiry_date_str}{seats_part}{LicenseManager._SALT}"
        
        full_hash = hashlib.sha256(combined.encode()).hexdigest().upper()
        
        if seats > 1:
            # Office Format: OFFICE-XXXX-XXXX-YYYYMMDD-SN
            return f"OFFICE-{full_hash[2:6]}-{full_hash[10:14]}-{expiry_date_str}-S{seats}"
        else:
            # Standalone Format: KEY-XXXX-XXXX-YYYYMMDD
            return f"KEY-{full_hash[2:6]}-{full_hash[10:14]}-{expiry_date_str}"

    @staticmethod
    def _get_appdata_path():
        path = os.path.join(os.environ.get('APPDATA', os.path.expanduser('~')), 'BKL_Office_Tools')
        if not os.path.exists(path):
            os.makedirs(path)
        return path

    @staticmethod
    def _get_license_file():
        return os.path.join(LicenseManager._get_appdata_path(), 'config.dat')

    @classmethod
    def get_status(cls):
        """
        Returns a dict with activation status and trial info.
        """
        file_path = cls._get_license_file()
        data = {"activated": False, "machine_id": cls.get_machine_id()}
        
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f:
                    data = json.load(f)
            except Exception:
                pass
        
        # Check activation status
        is_activated = data.get("activated", False)
        expiry_date_str = data.get("expiry_date", "")
        saved_key = data.get("license_key", "")
        member_no = data.get("member_no", "")
        seats = data.get("seats", 1)
        
        # TIME-TRAVEL PROTECTION
        is_time_cheating = False
        last_run_str = data.get("last_run_date", "")
        current_now = datetime.now()
        
        if last_run_str:
            try:
                last_run_dt = datetime.strptime(last_run_str, "%Y-%m-%d %H:%M:%S")
                # If current time is more than 2 hours behind last run, something is wrong
                if current_now < (last_run_dt - timedelta(hours=2)):
                    is_time_cheating = True
            except: pass
            
        # Update last run date
        data["last_run_date"] = current_now.strftime("%Y-%m-%d %H:%M:%S")
        with open(file_path, 'w') as f: json.dump(data, f)

        if is_time_cheating:
            return {
                "is_activated": False,
                "days_left": 0,
                "expired": True,
                "error": "System clock manipulation detected. Please correct your PC date.",
                "machine_id": cls.get_machine_id()
            }
        
        if is_activated and saved_key:
            expected_key = cls.generate_valid_key(cls.get_machine_id(), expiry_date_str, member_no, seats)
            if saved_key != expected_key:
                is_activated = False
        
        days_left = 0
        is_expired = False
        if is_activated and expiry_date_str:
            try:
                expiry_dt = datetime.strptime(expiry_date_str, "%Y%m%d")
                delta = expiry_dt - datetime.now()
                days_left = delta.days + 1
                if days_left <= 0:
                    is_expired = True
                    is_activated = False
            except Exception:
                is_activated = False
        
        if not is_activated and not is_expired:
            first_run = data.get("first_run_date")
            if not first_run:
                first_run = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                data["first_run_date"] = first_run
                with open(file_path, 'w') as f:
                    json.dump(data, f)
            
            first_run_dt = datetime.strptime(first_run, "%Y-%m-%d %H:%M:%S")
            trial_expiry_dt = first_run_dt + timedelta(days=cls._TRIAL_DAYS)
            trial_days_left = (trial_expiry_dt - datetime.now()).days + 1
            is_expired = trial_days_left <= 0
            days_left = max(0, trial_days_left)
        
        status = {
            "is_activated": is_activated,
            "days_left": days_left,
            "trial_limit": cls._TRIAL_DAYS,
            "expired": is_expired,
            "machine_id": cls.get_machine_id(),
            "expiry_date": expiry_date_str,
            "member_no": member_no,
            "is_ca": bool(member_no),
            "seats": seats,
            "is_office": seats > 1
        }
        
        # Add server IP if available
        status["server_ip"] = data.get("server_ip")
        return status

    @classmethod
    def activate(cls, key, member_no=None):
        """Attempts to activate the software locally (Standalone or Server)."""
        key = key.strip().upper()
        machine_id = cls.get_machine_id()
        member_no = str(member_no).strip().upper() if member_no else ""
        
        parts = key.split('-')
        seats = 1
        
        if key.startswith("OFFICE-") and len(parts) == 5:
            # OFFICE-HASH-HASH-YYYYMMDD-SN
            expiry_date_str = parts[3]
            try:
                seats = int(parts[4][1:]) # Strip 'S'
            except: return False
        elif key.startswith("KEY-") and len(parts) == 4:
            # KEY-HASH-HASH-YYYYMMDD
            expiry_date_str = parts[3]
        else:
            return False
            
        valid_key = cls.generate_valid_key(machine_id, expiry_date_str, member_no, seats)
        
        if key == valid_key:
            try:
                expiry_dt = datetime.strptime(expiry_date_str, "%Y%m%d")
                if expiry_dt < datetime.now(): return False
            except: return False

            file_path = cls._get_license_file()
            data = {}
            if os.path.exists(file_path):
                try:
                    with open(file_path, 'r') as f: data = json.load(f)
                except: pass
                    
            data.update({
                "activated": True,
                "license_key": key,
                "machine_id": machine_id,
                "member_no": member_no,
                "expiry_date": expiry_date_str,
                "seats": seats,
                "activation_date": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })
            with open(file_path, 'w') as f:
                json.dump(data, f)
            return True
        return False

    @staticmethod
    def save_server_ip(ip):
        file_path = LicenseManager._get_license_file()
        data = {}
        if os.path.exists(file_path):
            try:
                with open(file_path, 'r') as f: data = json.load(f)
            except: pass
        data["server_ip"] = ip
        with open(file_path, 'w') as f:
            json.dump(data, f)

class FloatingSeatManager:
    """
    Decentralized Peer-to-Peer License Seat Tracker.
    No central server required. All instances on the LAN broadcast their presence.
    Every running instance (server or client) broadcasts an ALIVE heartbeat and
    listens for others. Seats are counted by the number of unique machine_ids seen.
    """
    def __init__(self, machine_id, member_no=None, max_seats=0):
        self.machine_id = machine_id
        self.member_no = member_no
        self.max_seats = max_seats
        
        # Peer metadata: machine_id -> {"last_seen": float, "hostname": str, "ip": str}
        self.active_peers = {}
        self.running = False
        self.lock = threading.Lock()
        
        # Identify this machine
        self.hostname = socket.gethostname()
        self.local_ip = self._get_local_ip()
        
        # UI accessible status
        self.has_seat = True
        self.status_msg = "Checking network seats..."
        self.is_over_limit = False
        self.is_connected = False

    @staticmethod
    def _get_local_ip():
        """Get the LAN IP address of this machine."""
        try:
            # Connect to a non-routable address to determine the local IP
            s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
            s.settimeout(0.5)
            s.connect(("10.255.255.255", 1))
            ip = s.getsockname()[0]
            s.close()
            return ip
        except Exception:
            return "127.0.0.1"

    def start(self):
        if self.running: return
        self.running = True
        # Register ourselves immediately
        with self.lock:
            self.active_peers[self.machine_id] = {
                "last_seen": time.time(),
                "hostname": self.hostname,
                "ip": self.local_ip
            }
        threading.Thread(target=self._broadcast_loop, daemon=True).start()
        threading.Thread(target=self._listen_loop, daemon=True).start()
        threading.Thread(target=self._cleanup_loop, daemon=True).start()

    def stop(self):
        self.running = False

    def _broadcast_loop(self):
        """Periodically broadcast presence to let others know we are using a seat.
        Always broadcasts regardless of whether member_no is set — this ensures
        client PCs (without activation) are still visible on the network."""
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            while self.running:
                try:
                    msg = {
                        "action": "ALIVE",
                        "product": "BKL_CA_OFFICE",
                        "member_no": self.member_no or "",
                        "machine_id": self.machine_id,
                        "max_seats": self.max_seats,
                        "hostname": self.hostname,
                        "ip": self.local_ip
                    }
                    s.sendto(json.dumps(msg).encode(), ('<broadcast>', LicenseManager.UDP_PORT))
                except: pass
                time.sleep(8)

    def _listen_loop(self):
        """Listen for other instances on the LAN."""
        with socket.socket(socket.AF_INET, socket.SOCK_DGRAM) as s:
            s.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            s.setsockopt(socket.SOL_SOCKET, socket.SO_BROADCAST, 1)
            try:
                s.bind(('', LicenseManager.UDP_PORT))
                s.settimeout(2.0)
                while self.running:
                    try:
                        data, addr = s.recvfrom(2048)
                        msg = json.loads(data.decode())
                        
                        if msg.get("product") == "BKL_CA_OFFICE":
                            m_no = msg.get("member_no", "")
                            m_seats = msg.get("max_seats", 0)
                            mid = msg.get("machine_id", "")
                            m_hostname = msg.get("hostname", "Unknown")
                            m_ip = msg.get("ip", addr[0])
                            
                            if not mid:
                                continue
                            
                            # --- Network Discovery ---
                            # If we are a client (no member_no), adopt the first office we find
                            if not self.member_no and m_no:
                                self.member_no = m_no
                                self.max_seats = m_seats
                                self.is_connected = True
                            
                            # If the sender is the server (has member_no with seats),
                            # always update our max_seats to stay in sync
                            if m_no and m_seats > 0 and m_no == self.member_no:
                                self.max_seats = m_seats
                            
                            # --- Peer Registration ---
                            # Register this peer if:
                            #   1. They belong to the same network (same member_no), OR
                            #   2. They are a client still discovering (empty member_no)
                            #      AND we are the server (we have member_no)
                            same_network = (m_no == self.member_no) and self.member_no
                            client_discovering = (not m_no) and self.member_no
                            
                            if same_network or client_discovering:
                                with self.lock:
                                    self.active_peers[mid] = {
                                        "last_seen": time.time(),
                                        "hostname": m_hostname,
                                        "ip": m_ip
                                    }
                                    self.is_connected = True
                                    self._update_limit_status()
                    except socket.timeout:
                        continue
                    except: pass
            except Exception as e:
                print(f"P2P Listen Error: {e}")

    def _cleanup_loop(self):
        """Remove peers that haven't heartbeated recently."""
        while self.running:
            time.sleep(15)
            now = time.time()
            with self.lock:
                # Remove peers not seen for 30 seconds (3 missed heartbeats at 8s interval)
                expired = [mid for mid, info in self.active_peers.items() 
                          if now - info.get("last_seen", 0) > 30 and mid != self.machine_id]
                for mid in expired:
                    del self.active_peers[mid]
                
                self._update_limit_status()

    def _update_limit_status(self):
        """Re-evaluate if we are within seat limits."""
        # Our own self is always in the list
        self.active_peers[self.machine_id] = {
            "last_seen": time.time(),
            "hostname": self.hostname,
            "ip": self.local_ip
        }
        
        peer_ids = sorted(list(self.active_peers.keys()))
        instance_rank = peer_ids.index(self.machine_id) + 1 # 1-based rank
        total = len(peer_ids)
        
        # If max_seats is 0, we haven't discovered a network yet
        if self.max_seats <= 0:
            self.has_seat = True
            self.is_over_limit = False
            self.status_msg = "🔍 Searching for office network..."
            return
        
        self.is_over_limit = total > self.max_seats
        
        if self.is_over_limit:
            if instance_rank > self.max_seats:
                # We are the "extra" seat
                self.has_seat = False
                self.status_msg = f"⛔ Seat Limit Exceeded ({total}/{self.max_seats}). Functions Restricted."
            else:
                # We are among the first N, so we have a seat even if total > limit
                self.has_seat = True
                self.status_msg = f"⚠️ Network Full ({total}/{self.max_seats}), but you have an active seat."
        else:
            self.has_seat = True
            self.status_msg = f"✅ Network Shared ({total}/{self.max_seats} Seats in Use)"

    def get_peer_count(self):
        with self.lock:
            return len(self.active_peers)

    def get_active_list(self):
        """Returns list of dicts: [{"machine_id": ..., "hostname": ..., "ip": ..., "last_seen": ...}]"""
        with self.lock:
            result = []
            for mid, info in self.active_peers.items():
                result.append({
                    "machine_id": mid,
                    "hostname": info.get("hostname", "Unknown"),
                    "ip": info.get("ip", "?"),
                    "last_seen": info.get("last_seen", 0)
                })
            return result

