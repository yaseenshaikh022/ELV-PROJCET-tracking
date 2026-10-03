
import os
import sys
import sqlite3
import hashlib
import shutil
import re
import subprocess
import platform
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("openpyxl is missing. Run: python -m pip install openpyxl")

APP_NAME = "ELV Project Inventory Manager"
BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "ELV_Inventory_Data"
DATA_DIR.mkdir(exist_ok=True)
XLSX_FILE = DATA_DIR / "Inventory.xlsx"
DB_FILE = DATA_DIR / "InventoryUsers.db"
BACKUP_DIR = DATA_DIR / "Backups"
BACKUP_DIR.mkdir(exist_ok=True)

HEADERS = [
    "Sl. No.", "Product / Material", "Category", "Make", "Model / Part No.",
    "Qty", "Location", "Handover To", "Installed Place", "Serial Number",
    "MAC Address", "IP Address", "Asset Number", "IT Room", "Rack No.",
    "Switch Name / No.", "Switch Port", "Installation Date", "Status",
    "Testing Status", "Handover Status", "Warranty / AMC Expiry", "Remarks",
    "Created By", "Created Date", "Last Modified By", "Last Modified Date"
]
STATUS = ["Supplied / In Stock", "Partially Installed", "Installed", "Faulty", "Returned", "Missing"]
TESTING = ["Not Tested", "Pending", "Passed", "Failed"]
HANDOVER = ["Pending", "Handed Over", "Received", "N/A"]

FORM_FIELDS = [
    ("Product / Material","entry"), ("Category","entry"), ("Make","entry"),
    ("Model / Part No.","entry"), ("Qty","entry"), ("Location","entry"),
    ("Handover To","entry"), ("Installed Place","entry"), ("Serial Number","entry"),
    ("MAC Address","entry"), ("IP Address","entry"), ("Asset Number","entry"),
    ("IT Room","entry"), ("Rack No.","entry"), ("Switch Name / No.","entry"),
    ("Switch Port","entry"), ("Installation Date","entry"), ("Status","status"),
    ("Testing Status","testing"), ("Handover Status","handover"),
    ("Warranty / AMC Expiry","entry"), ("Remarks","text")
]

def ts():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")

def sha(s):
    return hashlib.sha256(s.encode("utf-8")).hexdigest()

def init_db():
    con = sqlite3.connect(DB_FILE)
    con.execute("""CREATE TABLE IF NOT EXISTS users(
        username TEXT PRIMARY KEY,
        password_hash TEXT NOT NULL,
        full_name TEXT NOT NULL,
        role TEXT NOT NULL,
        active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL)""")
    con.execute("""CREATE TABLE IF NOT EXISTS audit(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        timestamp TEXT NOT NULL, username TEXT NOT NULL, action TEXT NOT NULL,
        asset_number TEXT, serial_number TEXT, details TEXT)""")
    if con.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
        con.execute("INSERT INTO users VALUES(?,?,?,?,?,?)",
                    ("admin", sha("admin123"), "System Administrator", "Administrator", 1, ts()))
    con.commit(); con.close()

def audit(username, action, asset="", serial="", details=""):
    con = sqlite3.connect(DB_FILE)
    con.execute("INSERT INTO audit(timestamp,username,action,asset_number,serial_number,details) VALUES(?,?,?,?,?,?)",
                (ts(), username, action, asset or "", serial or "", details or ""))
    con.commit(); con.close()

def ensure_excel():
    if not XLSX_FILE.exists():
        wb = Workbook()
        ws = wb.active; ws.title = "Inventory"; ws.append(HEADERS)
        wb.create_sheet("User Register").append(["Username","Full Name","Role","Active","Created At"])
        wb.create_sheet("Audit Log").append(["ID","Timestamp","Username","Action","Asset Number","Serial Number","Details"])
        wb.save(XLSX_FILE)
    else:
        wb = load_workbook(XLSX_FILE)
        if "Inventory" not in wb.sheetnames:
            ws = wb.create_sheet("Inventory", 0); ws.append(HEADERS)
        else:
            ws = wb["Inventory"]
            current = [c.value for c in ws[1]]
            for h in HEADERS:
                if h not in current:
                    ws.cell(1, ws.max_column + 1, h)
        if "User Register" not in wb.sheetnames:
            wb.create_sheet("User Register").append(["Username","Full Name","Role","Active","Created At"])
        if "Audit Log" not in wb.sheetnames:
            wb.create_sheet("Audit Log").append(["ID","Timestamp","Username","Action","Asset Number","Serial Number","Details"])
        wb.save(XLSX_FILE)
    style_workbook()

def style_workbook():
    wb = load_workbook(XLSX_FILE)
    ws = wb["Inventory"]
    fill = PatternFill("solid", fgColor="1F4E78")
    font = Font(color="FFFFFF", bold=True)
    for c in ws[1]:
        c.fill=fill; c.font=font; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
    ws.freeze_panes="A2"; ws.auto_filter.ref=ws.dimensions
    widths=[8,26,18,18,24,8,22,22,30,24,20,18,20,18,12,20,14,18,20,18,18,20,35,20,20,22,22]
    for i,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(i)].width=w
    wb.save(XLSX_FILE)

def backup():
    if XLSX_FILE.exists():
        shutil.copy2(XLSX_FILE, BACKUP_DIR / f"Inventory_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")

def sync_excel():
    wb=load_workbook(XLSX_FILE)
    ur=wb["User Register"]
    if ur.max_row>1: ur.delete_rows(2,ur.max_row)
    con=sqlite3.connect(DB_FILE)
    users=con.execute("SELECT username,full_name,role,active,created_at FROM users ORDER BY username").fetchall()
    logs=con.execute("SELECT id,timestamp,username,action,asset_number,serial_number,details FROM audit ORDER BY id").fetchall()
    con.close()
    for r in users: ur.append([r[0],r[1],r[2],"Yes" if r[3] else "No",r[4]])
    al=wb["Audit Log"]
    if al.max_row>1: al.delete_rows(2,al.max_row)
    for r in logs: al.append(list(r))
    wb.save(XLSX_FILE)

def next_sl(ws):
    nums=[]
    for r in range(2,ws.max_row+1):
        try: nums.append(int(ws.cell(r,1).value))
        except: pass
    return max(nums,default=0)+1


def load_inventory_records():
    """Return inventory records as dictionaries from Inventory.xlsx."""
    if not XLSX_FILE.exists():
        return []
    wb = load_workbook(XLSX_FILE, data_only=True, read_only=True)
    ws = wb["Inventory"]
    headers = [c.value for c in ws[1]]
    records = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        if not any(v not in (None, "") for v in values):
            continue
        d = {}
        for i, h in enumerate(headers):
            if h:
                d[h] = values[i] if i < len(values) else ""
        records.append(d)
    wb.close()
    return records

def ping_host(ip, timeout_ms=1000):
    """
    Windows ICMP ping. Returns:
      (reachable: bool, latency_ms: float|None, ttl: int|None, raw: str)
    """
    ip = str(ip or "").strip()
    if not ip:
        return False, None, None, "No IP address"
    try:
        proc = subprocess.run(
            ["ping", "-n", "1", "-w", str(timeout_ms), ip],
            capture_output=True, text=True, encoding="cp1252", errors="replace",
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            timeout=max(2, timeout_ms / 1000 + 2)
        )
        raw = (proc.stdout or "") + (proc.stderr or "")
        reachable = proc.returncode == 0
        latency = None
        ttl = None

        # Works with common English Windows output, e.g. time<1ms and TTL=128.
        m = re.search(r"(?:time[=<]\s*)(\d+(?:[.,]\d+)?)\s*ms", raw, re.I)
        if m:
            latency = float(m.group(1).replace(",", "."))
        m = re.search(r"TTL[=\s]*(\d+)", raw, re.I)
        if m:
            ttl = int(m.group(1))

        # Some localized Windows builds still expose TTL=.
        return reachable, latency, ttl, raw.strip()
    except Exception as e:
        return False, None, None, str(e)

def ping_many(records, callback):
    for idx, record in enumerate(records):
        ip = str(record.get("IP Address") or "").strip()
        if not ip:
            result = (idx, False, None, None, "No IP address")
        else:
            ok, latency, ttl, raw = ping_host(ip)
            result = (idx, ok, latency, ttl, raw)
        callback(result)

class NetworkTester(tk.Toplevel):
    def __init__(self, parent, user):
        super().__init__(parent)
        self.parent = parent
        self.user = user
        self.title("Network Ping & TTL Tester")
        self.geometry("1050x650")
        self.minsize(900, 550)
        self.configure(padx=12, pady=12)
        self.build()

    def build(self):
        top = ttk.LabelFrame(self, text="Ping Tester", padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="IP Address").pack(side="left")
        self.ip_var = tk.StringVar()
        self.ip_entry = ttk.Entry(top, textvariable=self.ip_var, width=25)
        self.ip_entry.pack(side="left", padx=7)
        self.ip_entry.bind("<Return>", lambda e: self.test_one())
        ttk.Button(top, text="PING", command=self.test_one).pack(side="left", padx=3)
        ttk.Button(top, text="Clear", command=lambda: self.clear_one()).pack(side="left", padx=3)

        self.result_frame = ttk.LabelFrame(self, text="Live Result", padding=10)
        self.result_frame.pack(fill="x", pady=10)
        self.result_var = tk.StringVar(value="Enter an IP address and click PING.")
        self.result_label = ttk.Label(self.result_frame, textvariable=self.result_var, font=("Segoe UI", 12, "bold"))
        self.result_label.pack(anchor="w")

        inv = ttk.LabelFrame(self, text="Inventory Devices — double-click to test", padding=7)
        inv.pack(fill="both", expand=True)
        columns = ["Device Name", "Model", "IP Address", "Asset Number", "IT Room", "Switch", "Port", "Ping", "Latency", "TTL"]
        self.tree = ttk.Treeview(inv, columns=columns, show="headings")
        widths = {"Device Name":180,"Model":180,"IP Address":130,"Asset Number":150,"IT Room":130,
                  "Switch":140,"Port":80,"Ping":100,"Latency":90,"TTL":70}
        for c in columns:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=widths[c], minwidth=60)
        ys = ttk.Scrollbar(inv, orient="vertical", command=self.tree.yview)
        xs = ttk.Scrollbar(inv, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set, xscrollcommand=xs.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); ys.grid(row=0,column=1,sticky="ns"); xs.grid(row=1,column=0,sticky="ew")
        inv.rowconfigure(0,weight=1); inv.columnconfigure(0,weight=1)
        self.tree.bind("<Double-1>", self.test_selected)
        buttons = ttk.Frame(self); buttons.pack(fill="x", pady=8)
        ttk.Button(buttons, text="TEST SELECTED", command=self.test_selected).pack(side="left", padx=3)
        ttk.Button(buttons, text="TEST ALL IP DEVICES", command=self.test_all).pack(side="left", padx=3)
        ttk.Button(buttons, text="REFRESH INVENTORY", command=self.refresh).pack(side="left", padx=3)
        self.refresh()

    def clear_one(self):
        self.ip_var.set("")
        self.result_var.set("Enter an IP address and click PING.")

    def test_one(self):
        ip = self.ip_var.get().strip()
        if not ip:
            messagebox.showwarning("IP Address", "Enter an IP address.", parent=self)
            return
        self.result_var.set(f"Testing {ip} ...")
        def worker():
            ok, latency, ttl, raw = ping_host(ip)
            self.after(0, lambda: self.show_result(ip, ok, latency, ttl, raw))
        threading.Thread(target=worker, daemon=True).start()

    def show_result(self, ip, ok, latency, ttl, raw):
        status = "ONLINE" if ok else "OFFLINE"
        latency_text = f"{latency:.1f} ms" if latency is not None else "—"
        ttl_text = str(ttl) if ttl is not None else "—"
        self.result_var.set(f"{ip}   |   {status}   |   Latency: {latency_text}   |   TTL: {ttl_text}")
        audit(self.user["username"], "PING TEST", details=f"IP={ip}, status={status}, latency={latency_text}, TTL={ttl_text}")

    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        for d in load_inventory_records():
            ip = str(d.get("IP Address") or "").strip()
            if not ip:
                continue
            name = str(d.get("Product / Material") or "Unknown Device")
            self.tree.insert("", "end", values=(
                name, d.get("Model / Part No.",""), ip, d.get("Asset Number",""),
                d.get("IT Room",""), d.get("Switch Name / No.",""), d.get("Switch Port",""),
                "NOT TESTED", "—", "—"
            ))

    def test_selected(self, event=None):
        item = self.tree.focus()
        if not item:
            return
        vals = self.tree.item(item, "values")
        ip = vals[2]
        self.ip_var.set(ip)
        self.result_var.set(f"Testing {ip} ...")
        def worker():
            ok, latency, ttl, raw = ping_host(ip)
            self.after(0, lambda: self.update_row(item, ip, ok, latency, ttl, raw))
        threading.Thread(target=worker, daemon=True).start()

    def update_row(self, item, ip, ok, latency, ttl, raw):
        vals = list(self.tree.item(item, "values"))
        vals[7] = "ONLINE" if ok else "OFFLINE"
        vals[8] = f"{latency:.1f} ms" if latency is not None else "—"
        vals[9] = str(ttl) if ttl is not None else "—"
        self.tree.item(item, values=vals)
        self.show_result(ip, ok, latency, ttl, raw)

    def test_all(self):
        items = self.tree.get_children()
        if not items:
            messagebox.showinfo("Ping", "No inventory devices with IP addresses found.", parent=self)
            return
        self.result_var.set("Testing all inventory IP addresses...")
        def worker():
            for item in items:
                vals = self.tree.item(item, "values")
                ip = vals[2]
                ok, latency, ttl, raw = ping_host(ip)
                def apply(item=item, ip=ip, ok=ok, latency=latency, ttl=ttl, raw=raw):
                    self.update_row(item, ip, ok, latency, ttl, raw)
                self.after(0, apply)
        threading.Thread(target=worker, daemon=True).start()

class NetworkMap(tk.Toplevel):
    def __init__(self, parent, user):
        super().__init__(parent)
        self.parent = parent
        self.user = user
        self.title("ELV Network Map — IT Rooms")
        self.geometry("1350x850")
        self.minsize(1050,650)
        self.configure(padx=10,pady=10)
        self.tooltip = None
        self.nodes = []
        self.build()

    def build(self):
        bar = ttk.Frame(self); bar.pack(fill="x", pady=(0,7))
        ttk.Label(bar, text="Network Map", font=("Segoe UI",18,"bold")).pack(side="left")
        ttk.Button(bar,text="Refresh Map",command=self.draw_map).pack(side="right",padx=3)
        ttk.Button(bar,text="Ping All Devices",command=self.ping_all).pack(side="right",padx=3)
        ttk.Label(bar,text="Hover over a device node for Model, IP, Asset and Ping status.",foreground="#666").pack(side="left",padx=20)

        body=ttk.Frame(self); body.pack(fill="both",expand=True)
        left=ttk.LabelFrame(body,text="IT Rooms",padding=5); left.pack(side="left",fill="y")
        self.rooms=tk.Listbox(left,width=28,font=("Segoe UI",10))
        self.rooms.pack(fill="y",expand=True)
        self.rooms.bind("<<ListboxSelect>>",lambda e:self.draw_map())

        right=ttk.Frame(body); right.pack(side="left",fill="both",expand=True,padx=(8,0))
        self.canvas=tk.Canvas(right,background="#F5F7FA",highlightthickness=1,highlightbackground="#D0D7DE")
        ys=ttk.Scrollbar(right,orient="vertical",command=self.canvas.yview)
        xs=ttk.Scrollbar(right,orient="horizontal",command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=ys.set,xscrollcommand=xs.set)
        self.canvas.grid(row=0,column=0,sticky="nsew"); ys.grid(row=0,column=1,sticky="ns"); xs.grid(row=1,column=0,sticky="ew")
        right.rowconfigure(0,weight=1); right.columnconfigure(0,weight=1)
        self.canvas.bind("<Motion>",self.on_motion)
        self.records=[]
        self.load_rooms()
        self.draw_map()

    def load_rooms(self):
        self.records=load_inventory_records()
        rooms=[]
        for d in self.records:
            room=str(d.get("IT Room") or "").strip()
            if room and room not in rooms: rooms.append(room)
        rooms.sort()
        self.rooms.delete(0,"end")
        self.rooms.insert(0,"ALL IT ROOMS")
        for r in rooms:self.rooms.insert("end",r)
        if rooms:self.rooms.selection_set(0)

    def selected_room(self):
        sel=self.rooms.curselection()
        return "ALL IT ROOMS" if not sel or self.rooms.get(sel[0])=="ALL IT ROOMS" else self.rooms.get(sel[0])

    def device_name(self,d):
        # Prefer an explicit asset number/name, otherwise derive a readable device name.
        asset=str(d.get("Asset Number") or "").strip()
        product=str(d.get("Product / Material") or "").strip()
        return asset or product or "Unnamed Device"

    def device_icon_type(self,d):
        text=(str(d.get("Product / Material") or "")+" "+str(d.get("Category") or "")+" "+str(d.get("Model / Part No.") or "")).lower()
        if "camera" in text or "cctv" in text or "ptz" in text: return "CAM"
        if "nvr" in text or "dvr" in text: return "NVR"
        if "tv" in text or "display" in text or "led" in text: return "TV"
        if "rack" in text: return "RACK"
        if "switch" in text: return "SW"
        if "access point" in text or "wifi" in text or re.search(r"\bap\b",text): return "AP"
        if "server" in text: return "SRV"
        if "router" in text: return "RTR"
        if "firewall" in text: return "FW"
        return "DEV"

    def draw_map(self):
        self.canvas.delete("all"); self.nodes=[]
        room=self.selected_room()
        devices=[d for d in self.records if str(d.get("IP Address") or "").strip()]
        if room!="ALL IT ROOMS":
            devices=[d for d in devices if str(d.get("IT Room") or "").strip()==room]

        # IT room header
        title=f"IT ROOM: {room}"
        self.canvas.create_text(40,35,text=title,anchor="w",font=("Segoe UI",18,"bold"),fill="#1F2937")
        self.canvas.create_text(40,62,text=f"{len(devices)} IP-connected device(s) in view",anchor="w",font=("Segoe UI",10),fill="#667085")

        # Group devices by switch, then draw switch hubs and device nodes.
        groups={}
        for d in devices:
            sw=str(d.get("Switch Name / No.") or "Switch not specified").strip()
            groups.setdefault(sw,[]).append(d)

        y=115
        max_x=1100
        for switch, group in sorted(groups.items()):
            hub_x=150; hub_y=y+35
            self.canvas.create_rounded_rectangle = getattr(self.canvas, "create_round_rectangle", None)
            # Tk Canvas has no portable rounded rectangle; use rectangle.
            self.canvas.create_rectangle(hub_x-75,hub_y-25,hub_x+75,hub_y+25,fill="#1F4E78",outline="#163A5C")
            self.canvas.create_text(hub_x,hub_y,text=switch,fill="white",font=("Segoe UI",10,"bold"))
            cols=4
            for i,d in enumerate(group):
                col=i%cols; row=i//cols
                x=370+col*220; dy=35+row*110; yy=y+dy
                status="NOT TESTED"; ttl="—"
                dtype=self.device_icon_type(d)
                node_id=self.canvas.create_rectangle(x-32,yy-24,x+32,yy+24,fill="#FFFFFF",outline="#94A3B8",width=2)
                self.canvas.create_text(x,yy,text=dtype,font=("Segoe UI",8,"bold"),fill="#102A43")
                label=self.device_name(d)
                ip=str(d.get("IP Address") or "")
                self.canvas.create_text(x,yy+39,text=label[:24],font=("Segoe UI",8,"bold"),fill="#111827")
                self.canvas.create_text(x,yy+55,text=ip,font=("Segoe UI",7),fill="#475467")
                self.canvas.create_line(hub_x+75,hub_y,x-24,yy,fill="#98A2B3",width=2)
                self.nodes.append({"id":node_id,"x":x,"y":yy,"data":d,"status":status,"ttl":ttl})
            rows=(len(group)+cols-1)//cols
            y += max(150,rows*110+55)

        self.canvas.configure(scrollregion=(0,0,max(max_x,y+80),max(500,y+80)))

    def on_motion(self,event):
        x=self.canvas.canvasx(event.x); y=self.canvas.canvasy(event.y)
        hit=None
        for n in self.nodes:
            if (x-n["x"])**2+(y-n["y"])**2 <= 24**2:
                hit=n; break
        if hit:
            self.show_tooltip(event,hit)
        else:
            self.hide_tooltip()

    def show_tooltip(self,event,node):
        d=node["data"]
        text=(
            f"Device: {self.device_name(d)}\n"
            f"Model: {d.get('Model / Part No.','') or '—'}\n"
            f"IP: {d.get('IP Address','') or '—'}\n"
            f"Asset: {d.get('Asset Number','') or '—'}\n"
            f"Serial: {d.get('Serial Number','') or '—'}\n"
            f"IT Room: {d.get('IT Room','') or '—'}\n"
            f"Switch: {d.get('Switch Name / No.','') or '—'}\n"
            f"Port: {d.get('Switch Port','') or '—'}\n"
            f"Ping: {node['status']}\n"
            f"TTL: {node['ttl']}"
        )
        self.hide_tooltip()
        self.tooltip=tk.Toplevel(self)
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.attributes("-topmost",True)
        ttk.Label(self.tooltip,text=text,justify="left",padding=10,background="#FFFBEA").pack()
        self.tooltip.geometry(f"+{self.winfo_pointerx()+15}+{self.winfo_pointery()+15}")

    def hide_tooltip(self):
        if self.tooltip:
            try:self.tooltip.destroy()
            except:pass
            self.tooltip=None

    def ping_all(self):
        for n in self.nodes:
            ip=str(n["data"].get("IP Address") or "").strip()
            if not ip: continue
            def worker(n=n,ip=ip):
                ok,lat,ttl,raw=ping_host(ip)
                def apply():
                    n["status"]="ONLINE" if ok else "OFFLINE"
                    n["ttl"]=str(ttl) if ttl is not None else "—"
                    color="#22C55E" if ok else "#EF4444"
                    self.canvas.itemconfig(n["id"],fill=color)
                    audit(self.user["username"],"NETWORK MAP PING",details=f"IP={ip}, status={n['status']}, TTL={n['ttl']}, latency={lat}")
                self.after(0,apply)
            threading.Thread(target=worker,daemon=True).start()




class Dashboard(ttk.Frame):
    """Modern ELV Project Manager dashboard."""
    def __init__(self, root, user, open_module):
        super().__init__(root, padding=0)
        self.root=root; self.user=user; self.open_module=open_module
        self.sidebar_width=235
        self.build()

    def build(self):
        self.configure(style="App.TFrame")
        self._setup_styles()

        # Overall split: sidebar + content
        self.sidebar=ttk.Frame(self, style="Sidebar.TFrame", width=self.sidebar_width)
        self.sidebar.pack(side="left", fill="y")
        self.sidebar.pack_propagate(False)

        self.content=ttk.Frame(self, style="Content.TFrame")
        self.content.pack(side="left", fill="both", expand=True)

        self.build_sidebar()
        self.build_content()

    def _setup_styles(self):
        s=ttk.Style(self.root)
        try: s.theme_use("clam")
        except: pass
        s.configure("App.TFrame", background="#F4F7FB")
        s.configure("Sidebar.TFrame", background="#102A43")
        s.configure("Content.TFrame", background="#F4F7FB")
        s.configure("Nav.TButton", background="#102A43", foreground="#E8EEF5",
                    borderwidth=0, relief="flat", font=("Segoe UI",10,"bold"), padding=(12,9))
        s.map("Nav.TButton", background=[("active","#1769E0")], foreground=[("active","white")])
        s.configure("Card.TFrame", background="white", relief="solid", borderwidth=1)
        s.configure("Small.TLabel", background="#F4F7FB", foreground="#667085", font=("Segoe UI",9))
        s.configure("Section.TLabel", background="#F4F7FB", foreground="#0F172A", font=("Segoe UI",14,"bold"))
        s.configure("KPI.TLabel", background="white", foreground="#0F172A", font=("Segoe UI",19,"bold"))
        s.configure("KPIName.TLabel", background="white", foreground="#475467", font=("Segoe UI",8,"bold"))
        s.configure("CardTitle.TLabel", background="white", foreground="#0F172A", font=("Segoe UI",10,"bold"))
        s.configure("CardText.TLabel", background="white", foreground="#667085", font=("Segoe UI",8), wraplength=190)
        s.configure("SidebarTitle.TLabel", background="#102A43", foreground="white", font=("Segoe UI",15,"bold"))
        s.configure("SidebarSub.TLabel", background="#102A43", foreground="#AFC1D4", font=("Segoe UI",8))
        s.configure("SidebarSection.TLabel", background="#102A43", foreground="#7F9BB5", font=("Segoe UI",8,"bold"))
        s.configure("Session.TLabel", background="#0B2034", foreground="#D9E5EF", font=("Segoe UI",8))
        s.configure("White.TLabel", background="white", foreground="#0F172A")

    def build_sidebar(self):
        head=ttk.Frame(self.sidebar, style="Sidebar.TFrame", padding=(15,16,15,10)); head.pack(fill="x")
        ttk.Label(head,text="▣  ELV PROJECT",style="SidebarTitle.TLabel").pack(anchor="w")
        ttk.Label(head,text="PROJECT MANAGER",style="SidebarTitle.TLabel").pack(anchor="w")
        ttk.Label(head,text="ELV • CCTV • NETWORK • ASSET CONTROL",style="SidebarSub.TLabel").pack(anchor="w",pady=(4,0))

        nav=ttk.Frame(self.sidebar,style="Sidebar.TFrame"); nav.pack(fill="both",expand=True,padx=8,pady=8)
        items=[
            ("⌂","Dashboard",None),
            ("▣","Project Status","project"),
            ("□","Material Inventory","inventory"),
            ("▣","Asset Register","inventory"),
            ("◉","CCTV Assets","inventory"),
            ("◎","Network / IT","network"),
            ("⌘","Network Map","map"),
            ("◌","Ping / TTL Tester","network"),
            ("⚒","Installation","installation"),
            ("✓","Testing & Commissioning","testing"),
            ("⇄","Handover","handover"),
            ("▤","Documents","documents"),
            ("▥","Reports","reports"),
            ("♙","User Management","users"),
            ("⚙","Settings","settings"),
        ]
        ttk.Label(nav,text="PROJECT MODULES",style="SidebarSection.TLabel").pack(anchor="w",padx=7,pady=(5,6))
        for icon,title,target in items:
            b=ttk.Button(nav,text=f"{icon}   {title}",style="Nav.TButton",
                         command=lambda t=target: self.navigate(t))
            b.pack(fill="x",pady=1)

        # Current session
        self.session_box=ttk.Frame(self.sidebar,style="Session.TFrame",padding=10)
        self.session_box.pack(side="bottom",fill="x",padx=10,pady=10)
        self.session_text=tk.StringVar()
        ttk.Label(self.session_box,text="●  CURRENT SESSION",style="Session.TLabel").pack(anchor="w")
        ttk.Label(self.session_box,textvariable=self.session_text,style="Session.TLabel",justify="left").pack(anchor="w",pady=(6,0))
        self.update_session()

    def build_content(self):
        self.header=ttk.Frame(self.content,style="Content.TFrame",padding=(20,14,20,8)); self.header.pack(fill="x")
        left=ttk.Frame(self.header,style="Content.TFrame"); left.pack(side="left",fill="x",expand=True)
        ttk.Label(left,text="Project Dashboard",font=("Segoe UI",24,"bold"),background="#F4F7FB",foreground="#102A43").pack(anchor="w")
        ttk.Label(left,text=f"Welcome, {self.user['full_name']}  |  {self.user['role']}",style="Small.TLabel").pack(anchor="w",pady=(2,0))

        right=ttk.Frame(self.header,style="Content.TFrame"); right.pack(side="right")
        self.clock=tk.StringVar()
        ttk.Label(right,textvariable=self.clock,style="Small.TLabel").pack(side="left",padx=12)
        ttk.Button(right,text="Logout",command=self.logout).pack(side="left")

        # Scrollable content
        holder=ttk.Frame(self.content,style="Content.TFrame"); holder.pack(fill="both",expand=True,padx=18,pady=(0,12))
        self.canvas=tk.Canvas(holder,bg="#F4F7FB",highlightthickness=0)
        ys=ttk.Scrollbar(holder,orient="vertical",command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=ys.set)
        self.canvas.pack(side="left",fill="both",expand=True); ys.pack(side="right",fill="y")
        self.body=ttk.Frame(self.canvas,style="Content.TFrame")
        self.window_id=self.canvas.create_window((0,0),window=self.body,anchor="nw")
        self.body.bind("<Configure>",lambda e:self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>",lambda e:self.canvas.itemconfigure(self.window_id,width=max(e.width,900)))
        self.render_body()

    def render_body(self):
        for w in self.body.winfo_children(): w.destroy()
        stats=self.stats()

        # KPI cards
        kpi_frame=ttk.Frame(self.body,style="Content.TFrame"); kpi_frame.pack(fill="x",pady=(0,12))
        kpis=[
            ("TOTAL MATERIALS",stats["materials"],"#E7F0FF","▤"),
            ("INSTALLED",stats["installed"],"#E9FAF0","✓"),
            ("PENDING",stats["pending"],"#FFF5E5","◷"),
            ("IP DEVICES",stats["ip_devices"],"#F1EAFF","◎"),
            ("IT ROOMS",stats["it_rooms"],"#FFECEF","▣"),
            ("CCTV / NVR",stats["cctv"],"#E8FBFA","◉"),
            ("NETWORK DEVICES",stats["network"],"#FFF8DD","◇"),
            ("DOCUMENTS",stats["documents"],"#FDEBFF","▤"),
        ]
        for i,(name,value,bg,icon) in enumerate(kpis):
            card=tk.Frame(kpi_frame,bg="white",highlightbackground="#D9E2EC",highlightthickness=1)
            card.grid(row=0,column=i,sticky="nsew",padx=3)
            kpi_frame.columnconfigure(i,weight=1)
            tk.Label(card,text=icon,bg=bg,fg="#175CD3",font=("Segoe UI",18,"bold"),width=3).pack(side="left",fill="y")
            inner=tk.Frame(card,bg="white"); inner.pack(side="left",fill="both",expand=True,padx=6,pady=8)
            tk.Label(inner,text=name,bg="white",fg="#667085",font=("Segoe UI",7,"bold")).pack(anchor="w")
            tk.Label(inner,text=str(value),bg="white",fg="#101828",font=("Segoe UI",16,"bold")).pack(anchor="w")

        ttk.Label(self.body,text="Project Modules",style="Section.TLabel").pack(anchor="w",pady=(3,7))

        modules=[
            ("▣","#2F80ED","Project Status","Progress, activities and milestones","project"),
            ("□","#22C55E","Material Inventory","Supplied, received, issued and balance","inventory"),
            ("▣","#7C3AED","Asset Register","Installed assets and identity information","inventory"),
            ("◉","#EF4444","CCTV Assets","Cameras, NVRs, channels and installation","inventory"),
            ("◎","#0EA5E9","Network / IT","IP, IT rooms, switches, ports and status","network"),
            ("⌘","#14B8A6","Network Map","IT rooms, switches and device topology","map"),
            ("◌","#EAB308","Ping / TTL Tester","Live connectivity, latency and TTL","network"),
            ("⚒","#EC4899","Installation","Installation status and pending points","installation"),
            ("✓","#06B6D4","Testing & Commissioning","Testing, verification and commissioning","testing"),
            ("⇄","#8B5CF6","Handover","Handover status and documentation","handover"),
            ("▤","#F97316","Documents","Drawings, TDS, BOQ and related files","documents"),
            ("▥","#2563EB","Reports","Excel/PDF reports and summaries","reports"),
        ]
        grid=ttk.Frame(self.body,style="Content.TFrame"); grid.pack(fill="x")
        for i,(icon,accent,title,desc,target) in enumerate(modules):
            r=i//6; c=i%6
            card=tk.Frame(grid,bg="white",highlightbackground="#D9E2EC",highlightthickness=1,cursor="hand2")
            card.grid(row=r,column=c,sticky="nsew",padx=4,pady=4,ipady=3)
            grid.columnconfigure(c,weight=1)
            banner=tk.Frame(card,bg=accent,height=45); banner.pack(fill="x"); banner.pack_propagate(False)
            tk.Label(banner,text=icon,bg=accent,fg="white",font=("Segoe UI",22,"bold")).pack(pady=4)
            tk.Label(card,text=title,bg="white",fg="#102A43",font=("Segoe UI",9,"bold")).pack(pady=(8,2))
            tk.Label(card,text=desc,bg="white",fg="#667085",font=("Segoe UI",7),wraplength=150,justify="center").pack(padx=6)
            tk.Button(card,text="OPEN  ›",command=lambda t=target:self.navigate(t),
                      bg="white",fg=accent,relief="flat",font=("Segoe UI",8,"bold"),cursor="hand2").pack(pady=6)
            for widget in (card,banner):
                widget.bind("<Button-1>",lambda e,t=target:self.navigate(t))

        ttk.Label(self.body,text="Live Project Overview",style="Section.TLabel").pack(anchor="w",pady=(15,7))
        overview=ttk.Frame(self.body,style="Content.TFrame"); overview.pack(fill="x")

        progress=tk.Frame(overview,bg="white",highlightbackground="#D9E2EC",highlightthickness=1)
        progress.pack(side="left",fill="both",expand=True,padx=(0,5))
        tk.Label(progress,text=f"Project Progress  •  {stats['project']}%",bg="white",fg="#102A43",font=("Segoe UI",11,"bold")).pack(anchor="w",padx=12,pady=10)
        progress_items=[
            ("Material Supply",stats["supply"],"#22C55E"),
            ("Installation",stats["installation"],"#2F80ED"),
            ("Testing",stats["testing"],"#F97316"),
            ("Commissioning",stats["commissioning"],"#8B5CF6"),
            ("Handover",stats["handover"],"#EF4444"),
        ]
        for name,pct,color in progress_items:
            row=tk.Frame(progress,bg="white"); row.pack(fill="x",padx=12,pady=3)
            tk.Label(row,text=name,bg="white",fg="#475467",font=("Segoe UI",8),width=18,anchor="w").pack(side="left")
            bg=tk.Frame(row,bg="#EAECF0",height=9); bg.pack(side="left",fill="x",expand=True,padx=5); bg.pack_propagate(False)
            fg=tk.Frame(bg,bg=color,height=9); fg.place(relx=0,rely=0,relwidth=max(0,min(pct/100,1)),relheight=1)
            tk.Label(row,text=f"{pct}%",bg="white",fg="#344054",font=("Segoe UI",8,"bold"),width=6).pack(side="right")
        tk.Frame(progress,bg="white",height=8).pack()

        recent=tk.Frame(overview,bg="white",highlightbackground="#D9E2EC",highlightthickness=1,width=420)
        recent.pack(side="left",fill="both",padx=(5,0))
        tk.Label(recent,text="Recent Activity",bg="white",fg="#102A43",font=("Segoe UI",11,"bold")).pack(anchor="w",padx=12,pady=10)
        for line in self.recent_activity():
            tk.Label(recent,text=line,bg="white",fg="#475467",font=("Segoe UI",8),anchor="w").pack(fill="x",padx=12,pady=2)
        tk.Frame(recent,bg="white",height=8).pack()

        ttk.Label(self.body,text="Network Map — IT Room View",style="Section.TLabel").pack(anchor="w",pady=(15,7))
        self.network_preview(self.body,stats)

    def stats(self):
        rec=load_inventory_records()
        materials=len(rec)
        installed=sum(1 for d in rec if str(d.get("Status","")).lower()=="installed")
        pending=sum(1 for d in rec if str(d.get("Status","")).lower() in ("supplied / in stock","partially installed"))
        ips=sum(1 for d in rec if str(d.get("IP Address") or "").strip())
        rooms=len({str(d.get("IT Room") or "").strip() for d in rec if str(d.get("IT Room") or "").strip()})
        cctv=sum(1 for d in rec if any(k in (str(d.get("Product / Material") or "")+" "+str(d.get("Category") or "")).lower() for k in ("camera","cctv","nvr","dvr","ptz")))
        network=sum(1 for d in rec if any(k in (str(d.get("Product / Material") or "")+" "+str(d.get("Category") or "")).lower() for k in ("switch","router","firewall","access point","ap","server","nvr","dvr")))
        project=max(0,min(100,round((installed/max(materials,1))*100)))
        supply=100 if materials else 0
        installation=project
        testing=round(sum(1 for d in rec if str(d.get("Testing Status","")).lower()=="passed")/max(materials,1)*100)
        commissioning=round(sum(1 for d in rec if str(d.get("Status","")).lower()=="installed" and str(d.get("Testing Status","")).lower()=="passed")/max(materials,1)*100)
        handover=round(sum(1 for d in rec if str(d.get("Handover Status","")).lower() in ("handed over","received"))/max(materials,1)*100)
        # Documents are currently represented by rows with remarks/document-related values.
        documents=sum(1 for d in rec if any(k in str(d.get("Remarks","")).lower() for k in ("drawing","tds","document","report","boq")))
        return dict(materials=materials,installed=installed,pending=pending,ip_devices=ips,it_rooms=rooms,
                    cctv=cctv,network=network,documents=documents,project=project,supply=supply,
                    installation=installation,testing=testing,commissioning=commissioning,handover=handover)

    def recent_activity(self):
        try:
            con=sqlite3.connect(DB_FILE)
            rows=con.execute("SELECT timestamp,action,details FROM audit ORDER BY id DESC LIMIT 6").fetchall()
            con.close()
            if rows:
                return [f"• {r[0][11:19]}  {r[1]}  —  {str(r[2] or '')[:65]}" for r in rows]
        except: pass
        return ["• No activity recorded yet."]

    def network_preview(self,parent,stats):
        frame=tk.Frame(parent,bg="white",highlightbackground="#D9E2EC",highlightthickness=1,height=280)
        frame.pack(fill="x",pady=(0,20)); frame.pack_propagate(False)
        head=tk.Frame(frame,bg="white"); head.pack(fill="x",padx=12,pady=8)
        tk.Label(head,text="IT Room Network Topology",bg="white",fg="#102A43",font=("Segoe UI",11,"bold")).pack(side="left")
        tk.Button(head,text="Open Full Network Map",command=lambda:self.navigate("map"),bg="white",fg="#175CD3",relief="flat",font=("Segoe UI",8,"bold")).pack(side="right")

        canvas=tk.Canvas(frame,bg="#FBFCFE",highlightthickness=0)
        canvas.pack(fill="both",expand=True,padx=8,pady=(0,8))
        rec=[d for d in load_inventory_records() if str(d.get("IP Address") or "").strip()]
        rooms={}
        for d in rec:
            room=str(d.get("IT Room") or "IT Room Not Set").strip()
            rooms.setdefault(room,[]).append(d)
        if not rooms:
            canvas.create_text(30,80,text="Add IT Room, Switch and IP Address in Inventory to generate the network map.",
                               anchor="w",fill="#667085",font=("Segoe UI",9))
            return

        # Draw a compact room -> switch -> device diagram.
        x=80; y=35
        for room,devs in list(sorted(rooms.items()))[:3]:
            canvas.create_rectangle(x,y,x+170,y+45,fill="#E8F1FF",outline="#2F80ED",width=2)
            canvas.create_text(x+85,y+15,text="▣  "+room[:20],font=("Segoe UI",9,"bold"),fill="#102A43")
            canvas.create_text(x+85,y+33,text=f"{len(devs)} IP devices",font=("Segoe UI",7),fill="#667085")
            switches={}
            for d in devs: switches.setdefault(str(d.get("Switch Name / No.") or "Switch"),[]).append(d)
            sx=x; sy=y+85
            for sw,ds in list(switches.items())[:3]:
                canvas.create_rectangle(sx,sy,sx+120,sy+34,fill="#F1F8FF",outline="#7BA7E1")
                canvas.create_text(sx+60,sy+17,text="◇ "+sw[:16],font=("Segoe UI",8,"bold"),fill="#175CD3")
                canvas.create_line(x+85,y+45,sx+60,sy,fill="#98A2B3")
                for j,d in enumerate(ds[:5]):
                    dx=sx+j*50; dy=sy+55
                    dtype=self.device_icon_type(d)
                    canvas.create_oval(dx-15,dy-15,dx+15,dy+15,fill="#FFFFFF",outline="#CBD5E1",width=1)
                    canvas.create_text(dx,dy,text=dtype,fill="#102A43",font=("Segoe UI",6,"bold"))
                    canvas.create_text(dx,dy+24,text=str(d.get("IP Address",""))[-12:],fill="#667085",font=("Segoe UI",5))
                    canvas.create_line(sx+60,sy+34,dx,dy-15,fill="#CBD5E1")
                sx+=190
            x+=450
            if x>1000: x=80; y+=190

    def device_icon_type(self,d):
        text=(str(d.get("Product / Material") or "")+" "+str(d.get("Category") or "")+" "+str(d.get("Model / Part No.") or "")).lower()
        if "camera" in text or "cctv" in text or "ptz" in text: return "CAM"
        if "nvr" in text or "dvr" in text: return "NVR"
        if "tv" in text or "display" in text or "led" in text: return "TV"
        if "rack" in text: return "RACK"
        if "switch" in text: return "SW"
        if "access point" in text or "wifi" in text or re.search(r"\bap\b",text): return "AP"
        if "server" in text: return "SRV"
        if "router" in text: return "RTR"
        if "firewall" in text: return "FW"
        return "DEV"

    def navigate(self,target):
        if target is None: return
        if target=="network": NetworkTester(self.root,self.user)
        elif target=="map": NetworkMap(self.root,self.user)
        elif target=="inventory": self.open_inventory()
        elif target=="project": self.module_window("Project Status")
        elif target=="installation": self.module_window("Installation")
        elif target=="testing": self.module_window("Testing & Commissioning")
        elif target=="handover": self.module_window("Handover")
        elif target=="documents": self.module_window("Documents")
        elif target=="reports": self.module_window("Reports")
        elif target=="users":
            if self.user["role"].lower()=="administrator": UserWindow(self.root,self.user)
            else: messagebox.showwarning("Access Restricted","Only an Administrator can manage users.",parent=self.root)
        elif target=="settings": self.module_window("Settings")

    def open_inventory(self):
        win=tk.Toplevel(self.root)
        win.title(f"{APP_NAME} | Inventory")
        win.geometry("1450x900")
        MainApp(win,self.user).pack(fill="both",expand=True)

    def module_window(self,title):
        win=tk.Toplevel(self.root); win.title(f"{APP_NAME} | {title}"); win.geometry("1050x650")
        outer=ttk.Frame(win,padding=25); outer.pack(fill="both",expand=True)
        ttk.Label(outer,text=title,font=("Segoe UI",24,"bold")).pack(anchor="w")
        ttk.Label(outer,text="Module workspace",foreground="#667085").pack(anchor="w",pady=(0,20))
        ttk.Label(outer,text=f"{title} module is ready for project data integration.",font=("Segoe UI",12)).pack(anchor="w")
        ttk.Label(outer,text="The common project database and audit trail are already connected to this application.",
                  foreground="#667085",wraplength=800).pack(anchor="w",pady=8)
        ttk.Button(outer,text="Close",command=win.destroy).pack(anchor="w",pady=15)

    def update_session(self):
        # Dashboard session panel uses login audit timestamp if available.
        now=datetime.now()
        self.clock_update()
        self.session_text.set(
            f"User: {self.user['full_name']}\n"
            f"Role: {self.user['role']}\n"
            f"Date: {now.strftime('%d-%m-%Y')}\n"
            f"Time: {now.strftime('%H:%M:%S')}"
        )
        self.after(1000,self.update_session)

    def clock_update(self):
        if hasattr(self,"clock"):
            self.clock.set(datetime.now().strftime("%A  |  %d %b %Y  |  %H:%M:%S"))

    def logout(self):
        if messagebox.askyesno("Logout","Do you want to logout?",parent=self.root):
            audit(self.user["username"],"LOGOUT",details="Dashboard logout")
            sync_excel()
            self.open_module("login")


class LoginFrame(ttk.Frame):
    def __init__(self, root, on_login):
        super().__init__(root, padding=35)
        self.root=root; self.on_login=on_login
        self.columnconfigure(1,weight=1)
        ttk.Label(self,text="ELV PROJECT INVENTORY",font=("Segoe UI",22,"bold")).grid(row=0,column=0,columnspan=2,pady=(0,5))
        ttk.Label(self,text="Material • Asset • Network • Handover Management",foreground="#666").grid(row=1,column=0,columnspan=2,pady=(0,25))
        ttk.Label(self,text="Username").grid(row=2,column=0,sticky="w",pady=8)
        self.user=ttk.Entry(self,width=30); self.user.grid(row=2,column=1,sticky="ew",pady=8)
        ttk.Label(self,text="Password").grid(row=3,column=0,sticky="w",pady=8)
        self.pwd=ttk.Entry(self,show="*",width=30); self.pwd.grid(row=3,column=1,sticky="ew",pady=8)
        self.pwd.bind("<Return>",lambda e:self.login())
        ttk.Button(self,text="LOGIN",command=self.login).grid(row=4,column=0,columnspan=2,sticky="ew",pady=(20,8))
        ttk.Label(self,text="Default: admin / admin123",foreground="#777").grid(row=5,column=0,columnspan=2)
        self.user.focus_set()
    def login(self):
        u=self.user.get().strip(); p=self.pwd.get()
        con=sqlite3.connect(DB_FILE)
        row=con.execute("SELECT username,full_name,role FROM users WHERE username=? AND password_hash=? AND active=1",(u,sha(p))).fetchone()
        con.close()
        if not row:
            messagebox.showerror("Login Failed","Invalid username/password or inactive user.",parent=self.root); return
        audit(u,"LOGIN",details=f"Login date={datetime.now().strftime("%Y-%m-%d")}, login time={datetime.now().strftime("%H:%M:%S")}")
        self.on_login({"username":row[0],"full_name":row[1],"role":row[2]})

class UserWindow(tk.Toplevel):
    def __init__(self, parent, user):
        super().__init__(parent); self.title("User Register"); self.geometry("720x500"); self.user=user
        box=ttk.LabelFrame(self,text="Register User",padding=12); box.pack(fill="x",padx=12,pady=12)
        self.v={k:tk.StringVar() for k in ["Username","Full Name","Password","Role"]}
        labels=list(self.v)
        for i,l in enumerate(labels):
            ttk.Label(box,text=l).grid(row=i//2*2,column=i%2*2,sticky="w",padx=5,pady=5)
            e=ttk.Entry(box,textvariable=self.v[l],show="*" if l=="Password" else "")
            e.grid(row=i//2*2+1,column=i%2*2+1,sticky="ew",padx=5,pady=5)
        self.v["Role"].set("Engineer")
        ttk.Button(box,text="REGISTER USER",command=self.add).grid(row=4,column=0,columnspan=4,pady=8)
        for c in range(4): box.columnconfigure(c,weight=1)
        self.tree=ttk.Treeview(self,columns=("u","n","r","a","d"),show="headings")
        for c,h,w in [("u","Username",130),("n","Full Name",210),("r","Role",130),("a","Active",70),("d","Created",150)]:
            self.tree.heading(c,text=h); self.tree.column(c,width=w)
        self.tree.pack(fill="both",expand=True,padx=12,pady=(0,12)); self.refresh()
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        con=sqlite3.connect(DB_FILE); rows=con.execute("SELECT username,full_name,role,active,created_at FROM users ORDER BY username").fetchall(); con.close()
        for r in rows:self.tree.insert("", "end", values=(r[0],r[1],r[2],"Yes" if r[3] else "No",r[4]))
    def add(self):
        u,n,p,r=[self.v[x].get().strip() for x in self.v]
        if not u or not n or not p: messagebox.showwarning("Required","All fields are required.",parent=self); return
        try:
            con=sqlite3.connect(DB_FILE); con.execute("INSERT INTO users VALUES(?,?,?,?,?,?)",(u,sha(p),n,r or "Engineer",1,ts())); con.commit(); con.close()
        except sqlite3.IntegrityError:
            messagebox.showerror("Duplicate","Username already exists.",parent=self); return
        audit(self.user["username"],"USER REGISTERED",details=u); sync_excel()
        for x in self.v.values(): x.set("")
        self.v["Role"].set("Engineer"); self.refresh()
        messagebox.showinfo("Success","User registered.",parent=self)

class MainApp(ttk.Frame):
    def __init__(self, root, user):
        super().__init__(root,padding=10); self.root=root; self.user=user; self.edit_row=None
        self.session_start = datetime.now()
        self.session_date = self.session_start.strftime("%Y-%m-%d")
        self.session_time = self.session_start.strftime("%H:%M:%S")
        self.vars={}; self.build(); self.clear_form(); self.refresh()
        self.update_session_clock()
    def build(self):
        top=ttk.Frame(self); top.pack(fill="x",pady=(0,5))
        left=ttk.Frame(top); left.pack(side="left",fill="x",expand=True)
        ttk.Label(left,text="ELV PROJECT INVENTORY",font=("Segoe UI",21,"bold")).pack(side="left")
        ttk.Label(left,text=f"  |  {self.user['full_name']} ({self.user['role']})",foreground="#555").pack(side="left")

        session=ttk.LabelFrame(top,text="Current Session",padding=(8,3))
        session.pack(side="right",padx=(10,0))
        self.session_label=ttk.Label(session,text="Starting...",font=("Segoe UI",9))
        self.session_label.pack(side="left",padx=5)

        ttk.Button(top,text="Logout",command=self.logout).pack(side="right",padx=3)
        ttk.Button(top,text="Network Map",command=lambda:NetworkMap(self.root,self.user)).pack(side="right",padx=3)
        ttk.Button(top,text="Ping / TTL Tester",command=lambda:NetworkTester(self.root,self.user)).pack(side="right",padx=3)
        ttk.Button(top,text="Open Excel",command=self.open_excel).pack(side="right",padx=3)
        if self.user["role"].lower()=="administrator":
            ttk.Button(top,text="User Register",command=lambda:UserWindow(self.root,self.user)).pack(side="right",padx=3)

        form=ttk.LabelFrame(self,text="Material / Asset Details",padding=10); form.pack(fill="x")
        for i,(label,kind) in enumerate(FORM_FIELDS):
            r=(i//3)*2; c=(i%3)*2
            ttk.Label(form,text=label).grid(row=r,column=c,sticky="w",padx=5,pady=(3,1))
            if kind=="text":
                w=tk.Text(form,height=2); w.grid(row=r+1,column=c,columnspan=2,sticky="ew",padx=5,pady=(0,4))
            elif kind=="status":
                v=tk.StringVar(); w=ttk.Combobox(form,textvariable=v,values=STATUS,state="readonly"); w.grid(row=r+1,column=c,columnspan=2,sticky="ew",padx=5,pady=(0,4)); self.vars[label]=v
            elif kind=="testing":
                v=tk.StringVar(); w=ttk.Combobox(form,textvariable=v,values=TESTING,state="readonly"); w.grid(row=r+1,column=c,columnspan=2,sticky="ew",padx=5,pady=(0,4)); self.vars[label]=v
            elif kind=="handover":
                v=tk.StringVar(); w=ttk.Combobox(form,textvariable=v,values=HANDOVER,state="readonly"); w.grid(row=r+1,column=c,columnspan=2,sticky="ew",padx=5,pady=(0,4)); self.vars[label]=v
            else:
                v=tk.StringVar(); w=ttk.Entry(form,textvariable=v); w.grid(row=r+1,column=c,columnspan=2,sticky="ew",padx=5,pady=(0,4)); self.vars[label]=v
            if kind=="text": self.vars[label]=w
        for c in range(6): form.columnconfigure(c,weight=1)
        btn=ttk.Frame(form); btn.grid(row=16,column=0,columnspan=6,sticky="ew",pady=(7,0))
        for text,cmd in [("NEW / CLEAR",self.clear_form),("SAVE",self.save),("SAVE & NEXT ▶",self.save_next),("UPDATE SELECTED",self.update),("DELETE SELECTED",self.delete)]:
            ttk.Button(btn,text=text,command=cmd).pack(side="left",padx=3)
        ttk.Button(btn,text="REFRESH EXCEL",command=self.refresh_excel).pack(side="right",padx=3)

        box=ttk.LabelFrame(self,text="Inventory Register — double-click a row to edit",padding=7); box.pack(fill="both",expand=True,pady=(8,0))
        s=ttk.Frame(box); s.pack(fill="x",pady=(0,5))
        ttk.Label(s,text="Search").pack(side="left")
        self.search=tk.StringVar(); e=ttk.Entry(s,textvariable=self.search,width=55); e.pack(side="left",padx=6); e.bind("<KeyRelease>",lambda x:self.refresh())
        ttk.Button(s,text="Clear",command=lambda:(self.search.set(""),self.refresh())).pack(side="left")
        f=ttk.Frame(box); f.pack(fill="both",expand=True)
        self.tree=ttk.Treeview(f,columns=HEADERS,show="headings")
        for h in HEADERS:self.tree.heading(h,text=h); self.tree.column(h,width=125,minwidth=70)
        self.tree.column("Sl. No.",width=60); self.tree.column("Qty",width=55)
        self.tree.bind("<Double-1>",self.load)
        ys=ttk.Scrollbar(f,orient="vertical",command=self.tree.yview); xs=ttk.Scrollbar(f,orient="horizontal",command=self.tree.xview)
        self.tree.configure(yscrollcommand=ys.set,xscrollcommand=xs.set)
        self.tree.grid(row=0,column=0,sticky="nsew"); ys.grid(row=0,column=1,sticky="ns"); xs.grid(row=1,column=0,sticky="ew")
        f.rowconfigure(0,weight=1); f.columnconfigure(0,weight=1)
        self.status=ttk.Label(self,text="Ready",relief="sunken",anchor="w"); self.status.pack(fill="x",pady=(5,0))
    def get(self):
        d={}
        for k,w in self.vars.items(): d[k]=w.get("1.0","end").strip() if isinstance(w,tk.Text) else w.get().strip()
        return d
    def set(self,d):
        for k,w in self.vars.items():
            v=str(d.get(k,"") or "")
            if isinstance(w,tk.Text): w.delete("1.0","end"); w.insert("1.0",v)
            else:w.set(v)
    def clear_form(self):
        self.edit_row=None
        for w in self.vars.values():
            if isinstance(w,tk.Text): w.delete("1.0","end")
            else:w.set("")
        self.vars["Qty"].set("1"); self.vars["Status"].set(STATUS[0]); self.vars["Testing Status"].set(TESTING[0]); self.vars["Handover Status"].set(HANDOVER[0]); self.vars["Installation Date"].set(datetime.now().strftime("%Y-%m-%d"))
        self.status.config(text="New record mode.")
    def validate(self,d):
        miss=[x for x in ["Product / Material","Qty","Location"] if not d.get(x)]
        if miss: messagebox.showwarning("Required","Please fill: "+", ".join(miss),parent=self.root); return False
        try:
            if float(d["Qty"])<=0: raise ValueError
        except: messagebox.showwarning("Quantity","Qty must be a positive number.",parent=self.root); return False
        ip=d.get("IP Address","")
        if ip and not re.match(r"^(?:(?:25[0-5]|2[0-4]\d|1?\d?\d)\.){3}(?:25[0-5]|2[0-4]\d|1?\d?\d)$",ip):
            if not messagebox.askyesno("IP Address","IP address format looks unusual. Save anyway?",parent=self.root): return False
        return True
    def save(self):
        d=self.get()
        if not self.validate(d): return
        backup(); wb=load_workbook(XLSX_FILE); ws=wb["Inventory"]
        if self.edit_row is None:
            row=ws.max_row+1; d["Sl. No."]=next_sl(ws); d["Created By"]=self.user["username"]; d["Created Date"]=ts()
            action="CREATE"
        else:
            row=self.edit_row; d["Sl. No."]=ws.cell(row,1).value; d["Created By"]=ws.cell(row,24).value; d["Created Date"]=ws.cell(row,25).value; d["Last Modified By"]=self.user["username"]; d["Last Modified Date"]=ts(); action="UPDATE"
        for c,h in enumerate(HEADERS,1): ws.cell(row,c,d.get(h,""))
        wb.save(XLSX_FILE)
        audit(self.user["username"],action,d.get("Asset Number"),d.get("Serial Number"),d.get("Product / Material")); sync_excel()
        self.refresh(); self.status.config(text=f"{action} successful — Sl. No. {d['Sl. No.']}")
        return True
    def save_next(self):
        if self.save(): self.clear_form(); self.status.config(text="Saved. Ready for next material.")
    def update(self):
        if self.edit_row is None: messagebox.showinfo("Select Item","Double-click a row first.",parent=self.root); return
        self.save()
    def selected_sl(self):
        item=self.tree.focus()
        return self.tree.item(item,"values")[0] if item else None
    def delete(self):
        sl=self.selected_sl()
        if not sl: messagebox.showinfo("Select Item","Select an item first.",parent=self.root); return
        if not messagebox.askyesno("Confirm Delete",f"Delete Sl. No. {sl}?\nA backup will be created first.",parent=self.root): return
        backup(); wb=load_workbook(XLSX_FILE); ws=wb["Inventory"]; row=None
        for r in range(2,ws.max_row+1):
            if str(ws.cell(r,1).value)==str(sl): row=r; break
        if row:
            asset=ws.cell(row,13).value; serial=ws.cell(row,10).value; ws.delete_rows(row,1)
            for r in range(2,ws.max_row+1): ws.cell(r,1,r-1)
            wb.save(XLSX_FILE); audit(self.user["username"],"DELETE",asset,serial,f"Deleted Sl. No. {sl}"); sync_excel()
        self.clear_form(); self.refresh(); self.status.config(text=f"Deleted Sl. No. {sl}")
    def refresh(self):
        self.tree.delete(*self.tree.get_children())
        wb=load_workbook(XLSX_FILE,data_only=True); ws=wb["Inventory"]; q=self.search.get().lower() if hasattr(self,"search") else ""
        for row in ws.iter_rows(min_row=2,values_only=True):
            vals=list(row)+[""]*max(0,len(HEADERS)-len(row)); vals=vals[:len(HEADERS)]
            if q and q not in " ".join(str(x or "") for x in vals).lower(): continue
            self.tree.insert("", "end", values=vals)
    def load(self,event=None):
        item=self.tree.focus()
        if not item:return
        sl=self.tree.item(item,"values")[0]; wb=load_workbook(XLSX_FILE,data_only=True); ws=wb["Inventory"]
        for r in range(2,ws.max_row+1):
            if str(ws.cell(r,1).value)==str(sl):
                self.edit_row=r; self.set({h:ws.cell(r,c+1).value for c,h in enumerate(HEADERS)}); self.status.config(text=f"Editing Sl. No. {sl}"); return
    def update_session_clock(self):
        if not self.winfo_exists():
            return
        current=datetime.now()
        elapsed=int((current-self.session_start).total_seconds())
        hours, rem=divmod(elapsed,3600)
        minutes, seconds=divmod(rem,60)
        self.session_label.config(
            text=f"Login: {self.session_date} {self.session_time}  |  Now: {current.strftime('%Y-%m-%d %H:%M:%S')}  |  Session: {hours:02d}:{minutes:02d}:{seconds:02d}"
        )
        self._session_after=self.after(1000,self.update_session_clock)

    def logout(self):
        if not messagebox.askyesno("Logout","Do you want to logout from the ELV Inventory system?",parent=self.root):
            return
        elapsed=int((datetime.now()-self.session_start).total_seconds())
        audit(
            self.user["username"],
            "LOGOUT",
            details=f"Session date={self.session_date}, login={self.session_time}, duration_seconds={elapsed}"
        )
        sync_excel()
        # Return to login without destroying the main Tk root.
        for w in self.root.winfo_children():
            w.destroy()
        self.root.title(APP_NAME)
        LoginFrame(self.root,self._login_again).pack(fill="both",expand=True)

    def _login_again(self,user):
        for w in self.root.winfo_children():
            w.destroy()
        self.root.title(f"{APP_NAME} | {user['full_name']}")
        MainApp(self.root,user).pack(fill="both",expand=True)

    def refresh_excel(self):
        backup(); sync_excel(); self.refresh(); self.status.config(text=f"Excel synchronized: {XLSX_FILE}")
    def open_excel(self):
        try: os.startfile(XLSX_FILE)
        except Exception as e: messagebox.showerror("Excel",str(e),parent=self.root)


def start():
    init_db(); ensure_excel(); sync_excel()
    root=tk.Tk()
    root.title(APP_NAME)
    root.geometry("1550x920")
    root.minsize(1200,760)
    try: ttk.Style(root).theme_use("clam")
    except: pass

    state={"user":None}

    def show_login():
        for w in root.winfo_children(): w.destroy()
        root.title(APP_NAME)
        LoginFrame(root,show_dashboard).pack(fill="both",expand=True)

    def show_dashboard(user):
        state["user"]=user
        for w in root.winfo_children(): w.destroy()
        root.title(f"{APP_NAME} | Dashboard | {user['full_name']}")
        Dashboard(root,user,open_module).pack(fill="both",expand=True)

    def open_module(name):
        if name=="login":
            show_login()
        elif name=="network":
            NetworkTester(root,state["user"])
        elif name=="map":
            NetworkMap(root,state["user"])
        elif name=="inventory":
            win=tk.Toplevel(root); win.title(f"{APP_NAME} | Inventory"); win.geometry("1500x900")
            MainApp(win,state["user"]).pack(fill="both",expand=True)
        elif name=="project":
            win=tk.Toplevel(root); win.title("Project Status"); win.geometry("1050x650")
            ttk.Label(win,text="Project Status",font=("Segoe UI",24,"bold")).pack(anchor="w",padx=25,pady=25)
            ttk.Label(win,text="Use this workspace for project milestones, progress, activities and issues.",
                      font=("Segoe UI",11)).pack(anchor="w",padx=25)
        else:
            win=tk.Toplevel(root); win.title(name); win.geometry("1050x650")
            ttk.Label(win,text=name,font=("Segoe UI",24,"bold")).pack(anchor="w",padx=25,pady=25)
            ttk.Label(win,text="Module workspace connected to the ELV Project Manager.",
                      font=("Segoe UI",11)).pack(anchor="w",padx=25)

    root.protocol("WM_DELETE_WINDOW",root.destroy)
    show_login()
    root.mainloop()

if __name__=="__main__":
    start()
