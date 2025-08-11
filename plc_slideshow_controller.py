import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import configparser
import threading
import time
import os
import sys
import shutil
import ctypes

# Preload snap7.dll if running from PyInstaller
def preload_snap7():
    # Always load snap7.dll from the same folder as the executable/script
    dll_path = os.path.join(os.path.dirname(os.path.abspath(sys.argv[0])), "snap7.dll")
    try:
        ctypes.WinDLL(dll_path)
        print(f"[INFO] Loaded snap7.dll from: {dll_path}")
    except Exception as e:
        print(f"[ERROR] Failed to preload snap7.dll from {dll_path}: {e}")
        # don't raise here; let import handle missing snap7 gracefully


preload_snap7()

try:
    import snap7
    SNAP7_AVAILABLE = True
except Exception:
    SNAP7_AVAILABLE = False

try:
    import win32com.client
    import pythoncom
    PPT_AVAILABLE = True
except Exception:
    # pythoncom may be missing when pywin32 isn't installed
    win32com = None
    pythoncom = None
    PPT_AVAILABLE = False


class PLCSlideshowController:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PLC Slideshow Controller")
        self.root.geometry("900x620")
        # keep a neutral background compatible with ttk themes
        try:
            self.root.configure(bg='#f0f0f0')
        except Exception:
            pass

        # Simulation & state
        self.simulation_mode = tk.BooleanVar(value=False)
        self.last_start_bit = 0
        self.last_next_bit = 0

        # PLC & PPT state
        self.plc_client = None
        self.is_connected = False
        self.is_monitoring = False
        self.ppt_app = None
        self.ppt_presentation = None
        self.current_slideshow_index = -1

        # paths and storage
        self.app_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.slides_directory = os.path.join(self.app_directory, "slides")
        os.makedirs(self.slides_directory, exist_ok=True)

        self.slide_mappings = {}

        # config
        self.config = configparser.ConfigParser()
        self.load_config()

        # GUI
        self.setup_gui()
        self.load_slide_mappings()

    def load_config(self):
        config_path = os.path.join(self.app_directory, 'config.ini')
        if os.path.exists(config_path):
            self.config.read(config_path)
        else:
            self.config['PLC'] = {'ip_address': '192.168.1.100', 'db_number': '1'}
            with open(config_path, 'w') as f:
                self.config.write(f)

    def setup_gui(self):
        main = ttk.Frame(self.root)
        main.pack(fill='both', expand=True, padx=10, pady=10)
        ttk.Label(main, text="PLC Slideshow Controller", font=('Arial', 16, 'bold')).pack(pady=10)

        # Simulation
        sim_frame = ttk.LabelFrame(main, text="Simulation Mode")
        sim_frame.pack(fill='x', pady=5)
        ttk.Checkbutton(sim_frame, text="Enable Simulation", variable=self.simulation_mode).pack(anchor='w', padx=5)
        sim_inner = ttk.Frame(sim_frame)
        sim_inner.pack(fill='x', pady=5)
        ttk.Label(sim_inner, text="Index:").pack(side='left')
        self.sim_index = ttk.Entry(sim_inner, width=5)
        self.sim_index.pack(side='left', padx=5)
        ttk.Button(sim_inner, text="▶ Start", command=lambda: self.simulate_bits(start=True)).pack(side='left', padx=5)
        ttk.Button(sim_inner, text="⏭ Next", command=lambda: self.simulate_bits(next_cmd=True)).pack(side='left', padx=5)

        # Status
        status = ttk.LabelFrame(main, text="Status")
        status.pack(fill='x', pady=5)
        row = ttk.Frame(status)
        row.pack(fill='x')
        ttk.Label(row, text="Conn:").pack(side='left')
        self.conn_lbl = ttk.Label(row, text="Disconnected", foreground="red")
        self.conn_lbl.pack(side='left', padx=5)
        ttk.Label(row, text="Idx:").pack(side='left')
        self.idx_lbl = ttk.Label(row, text="0")
        self.idx_lbl.pack(side='left', padx=5)

        # PLC Controls
        cfg = ttk.LabelFrame(main, text="PLC Controls")
        cfg.pack(fill='x', pady=5)
        e1 = ttk.Frame(cfg)
        e1.pack(fill='x', pady=2)
        ttk.Label(e1, text="PLC IP:", width=10).pack(side='left')
        self.ip_e = ttk.Entry(e1)
        self.ip_e.insert(0, self.config['PLC']['ip_address'])
        self.ip_e.pack(side='left', fill='x', expand=True, padx=5)
        e2 = ttk.Frame(cfg)
        e2.pack(fill='x', pady=2)
        ttk.Label(e2, text="DB Num:", width=10).pack(side='left')
        self.db_e = ttk.Entry(e2, width=5)
        self.db_e.insert(0, self.config['PLC']['db_number'])
        self.db_e.pack(side='left', padx=5)
        btns = ttk.Frame(cfg)
        btns.pack(fill='x', pady=5)
        # Connect button
        self.connect_btn = ttk.Button(btns, text="Connect", command=self.toggle_connection)
        self.connect_btn.pack(side='left', padx=5)
        # Store a reference to Start Monitoring button instead of nametowidget
        self.start_monitor_btn = ttk.Button(btns, text="Start Monitoring", command=self.toggle_monitor, state='disabled')
        self.start_monitor_btn.pack(side='left')

        # Slide Mapping
        mapf = ttk.LabelFrame(main, text="Slide Mapping")
        mapf.pack(fill='both', expand=True, pady=5)
        ctrl = ttk.Frame(mapf)
        ctrl.pack(fill='x', pady=5)
        ttk.Label(ctrl, text="Index:").pack(side='left')
        self.map_idx = ttk.Entry(ctrl, width=5)
        self.map_idx.pack(side='left', padx=5)
        ttk.Button(ctrl, text="Add Slideshow", command=self.add_slide_mapping).pack(side='left', padx=5)
        ttk.Button(ctrl, text="Remove Selected", command=self.remove_slide_mapping).pack(side='left')
        cols = ('Index', 'File')
        self.tree = ttk.Treeview(mapf, columns=cols, show='headings')
        self.tree.heading('Index', text='Index')
        self.tree.heading('File', text='File')
        self.tree.pack(fill='both', expand=True)

        # Current Slideshow
        curf = ttk.LabelFrame(main, text="Current Slideshow")
        curf.pack(fill='x', pady=5)
        self.cur_lbl = ttk.Label(curf, text="None")
        self.cur_lbl.pack()

    def simulate_bits(self, start=False, next_cmd=False):
        try:
            idx = int(self.sim_index.get() or 0)
        except Exception:
            idx = 0
        if start:
            self.handle_start_pause(idx)
        if next_cmd:
            self.next_slide()
            # quick feedback pulse
            self.send_feedback_bit(True)
            self.root.after(200, lambda: self.send_feedback_bit(False))

    def toggle_connection(self):
        if self.is_connected:
            self.disconnect_plc()
        else:
            self.connect_plc()

    def connect_plc(self):
        if not SNAP7_AVAILABLE:
            messagebox.showerror("Error", "Install python-snap7")
            return
        try:
            ip = self.ip_e.get().strip()
            c = snap7.client.Client()
            c.connect(ip, 0, 1)
            self.plc_client = c
            self.is_connected = True
            self.conn_lbl.config(text="Connected", foreground="green")
            # enable start monitoring button safely using stored reference
            self.start_monitor_btn.config(state='normal')
            # optionally start monitoring automatically — uncomment next line if desired
            # self.toggle_monitor()
        except Exception as e:
            messagebox.showerror("Conn Err", str(e))

    def disconnect_plc(self):
        if self.is_monitoring:
            self.toggle_monitor()
        if self.plc_client:
            try:
                self.plc_client.disconnect()
            except Exception:
                pass
        self.is_connected = False
        self.conn_lbl.config(text="Disconnected", foreground="red")
        self.start_monitor_btn.config(state='disabled')

    def toggle_monitor(self):
        if self.is_monitoring:
            self.is_monitoring = False
            self.start_monitor_btn.config(text="Start Monitoring")
        else:
            self.is_monitoring = True
            self.start_monitor_btn.config(text="Stop Monitoring")
            t = threading.Thread(target=self.monitor_loop, daemon=True)
            t.start()

    def monitor_loop(self):
        """
        Background thread that reads PLC DB and controls PowerPoint.
        Must initialize COM in this thread before using win32com.
        """
        # Initialize COM for this thread if pythoncom is available
        if pythoncom:
            try:
                pythoncom.CoInitialize()
            except Exception as e:
                print(f"[COM Init Error] {e}")

        try:
            # prepare PowerPoint if available
            ppt = None
            if PPT_AVAILABLE:
                try:
                    ppt = win32com.client.Dispatch("PowerPoint.Application")
                    # make sure it is visible so SlideShowWindow works reliably
                    try:
                        ppt.Visible = True
                    except Exception:
                        pass
                except Exception as e:
                    print(f"[PPT Dispatch Error] {e}")
                    ppt = None

            while self.is_monitoring:
                try:
                    if not self.is_connected or not self.plc_client:
                        time.sleep(0.2)
                        continue

                    db = int(self.db_e.get())
                    # Read 2 bytes: first is control byte, second is index
                    data = self.plc_client.db_read(db, 0, 2)
                    if not data or len(data) < 2:
                        time.sleep(0.1)
                        continue

                    control_byte = int(data[0])
                    idx = int(data[1])

                    start_bit = control_byte & 0b00000001
                    next_bit = (control_byte >> 1) & 0b00000001

                    # handle start/pause edge
                    if start_bit == 1 and self.last_start_bit == 0:
                        # schedule GUI-safe call
                        self.root.after(0, lambda i=idx: self.handle_start_pause(i))

                    # handle next edge
                    if next_bit == 1 and self.last_next_bit == 0:
                        self.root.after(0, self.next_slide)
                        # feedback pulse
                        self.send_feedback_bit(True)
                        self.root.after(200, lambda: self.send_feedback_bit(False))

                    self.last_start_bit = start_bit
                    self.last_next_bit = next_bit

                    # update displayed index label in GUI thread
                    self.root.after(0, lambda i=idx: self.idx_lbl.config(text=str(i)))

                except Exception as e:
                    print(f"[Monitor Error] {e}")

                time.sleep(0.1)

        finally:
            # cleanup COM for this thread
            if pythoncom:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
            # don't quit global ppt_app here; we manage ppt per slideshow open/close in other methods

    def add_slide_mapping(self):
        try:
            idx = int(self.map_idx.get())
            assert 0 <= idx <= 255
        except Exception:
            messagebox.showerror("Error", "Enter index 0-255")
            return
        f = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx *.ppsx")])
        if not f:
            return
        dst = os.path.join(self.slides_directory, os.path.basename(f))
        try:
            shutil.copy2(f, dst)
        except Exception as e:
            messagebox.showerror("Copy Error", str(e))
            return
        self.slide_mappings[idx] = dst
        self.save_mappings()
        self.refresh_tree()

    def remove_slide_mapping(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(self.tree.item(sel[0])['values'][0])
        if idx in self.slide_mappings:
            del self.slide_mappings[idx]
            self.save_mappings()
            self.refresh_tree()

    def load_slide_mappings(self):
        p = os.path.join(self.app_directory, 'slide_mappings.txt')
        if os.path.exists(p):
            try:
                for L in open(p, encoding='utf-8'):
                    i, fp = L.strip().split('=', 1)
                    if os.path.exists(fp):
                        self.slide_mappings[int(i)] = fp
            except Exception:
                pass
        self.refresh_tree()

    def save_mappings(self):
        p = os.path.join(self.app_directory, 'slide_mappings.txt')
        try:
            with open(p, 'w', encoding='utf-8') as f:
                for i, fp in self.slide_mappings.items():
                    f.write(f"{i}={fp}\n")
        except Exception as e:
            print(f"[Save Mappings Error] {e}")

    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for i, fp in sorted(self.slide_mappings.items()):
            self.tree.insert('', "end", values=(i, os.path.basename(fp)))

    # PowerPoint control
    def open_ppt(self, path):
        """
        Open the given pptx file in PowerPoint and keep a reference in self.ppt_presentation.
        This uses the global COM instance controlled by win32com; ensure COM is initialized
        in the thread that calls this if calling from background threads.
        """
        try:
            if self.ppt_presentation:
                try:
                    self.ppt_presentation.Close()
                except Exception:
                    pass
                self.ppt_presentation = None

            if not PPT_AVAILABLE:
                messagebox.showerror("PPT Error", "pywin32 is not installed. PowerPoint control won't work.")
                return

            # create or reuse application object
            if not self.ppt_app:
                try:
                    # If called from main thread, normal dispatch is fine.
                    self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
                    self.ppt_app.Visible = True
                except Exception as e:
                    print(f"[open_ppt Dispatch Error] {e}")
                    self.ppt_app = None
                    messagebox.showerror("PPT Error", "Could not start PowerPoint via COM.")
                    return

            # Open presentation (WithWindow=True to ensure presentation has a window)
            try:
                self.ppt_presentation = self.ppt_app.Presentations.Open(path, WithWindow=True)
            except Exception as e:
                print(f"[Open PPT Error] {e}")
                messagebox.showerror("PPT Error", f"Could not open: {os.path.basename(path)}")
                self.ppt_presentation = None

        except Exception as e:
            print(f"[open_ppt Error] {e}")

    def start_slideshow(self):
        try:
            if self.ppt_presentation:
                self.ppt_presentation.SlideShowSettings.Run()
        except Exception as e:
            print(f"[Start Slideshow Error] {e}")

    def handle_start_pause(self, idx):
        path = self.slide_mappings.get(idx)
        if not path or not os.path.exists(path):
            self.cur_lbl.config(text=f"No PPT for idx {idx}")
            return
        if idx != self.current_slideshow_index:
            self.open_ppt(path)
            self.current_slideshow_index = idx
            self.cur_lbl.config(text=os.path.basename(path))
        self.start_slideshow()

    def next_slide(self):
        try:
            if self.ppt_presentation and getattr(self.ppt_presentation, 'SlideShowWindow', None):
                try:
                    self.ppt_presentation.SlideShowWindow.View.Next()
                except Exception as e:
                    print(f"[PPT Next Error] {e}")
            else:
                # If presentation isn't running, try a safe Next on application slideshow windows
                if PPT_AVAILABLE and self.ppt_app:
                    try:
                        # if there are SlideShowWindows, call Next on the first one
                        if self.ppt_app.SlideShowWindows.Count >= 1:
                            self.ppt_app.SlideShowWindows(1).View.Next()
                    except Exception as e:
                        print(f"[PPT Next alternate Error] {e}")
        except Exception as e:
            print(f"[next_slide Error] {e}")

    def send_feedback_bit(self, state: bool):
        try:
            if not self.plc_client:
                return
            db = int(self.db_e.get())
            # read 1 byte, modify bit 2 (mask 0b00000100)
            data = self.plc_client.db_read(db, 0, 1)
            if not data:
                return
            byte_val = int(data[0])
            if state:
                byte_val |= 0b00000100
            else:
                byte_val &= 0b11111011
            self.plc_client.db_write(db, 0, bytes([byte_val]))
        except Exception as e:
            print(f"[PLC Feedback Error] {e}")

    def run(self):
        if not SNAP7_AVAILABLE:
            messagebox.showwarning("Warn", "python-snap7 missing — PLC functionality disabled")
        if not PPT_AVAILABLE:
            # warn, but keep UI usable
            messagebox.showwarning("Warn", "pywin32 missing — PowerPoint control disabled")
        self.root.mainloop()


if __name__ == "__main__":
    PLCSlideshowController().run()
