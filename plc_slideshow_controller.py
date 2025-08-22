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
        raise


preload_snap7()

try:
    import snap7
    SNAP7_AVAILABLE = True
except ImportError:
    SNAP7_AVAILABLE = False

try:
    import win32com.client
    PPT_AVAILABLE = True
except ImportError:
    PPT_AVAILABLE = False


class PLCSlideshowController:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("PLC Slideshow Controller")
        self.root.geometry("800x600")
        self.root.configure(bg='#f0f0f0')

        self.simulation_mode = tk.BooleanVar(value=False)
        self.last_start_bit = 0
        self.last_next_bit = 0

        self.plc_client = None
        self.is_connected = False
        self.is_monitoring = False
        self.ppt_app = None
        self.ppt_presentation = None
        self.current_slideshow_index = -1

        self.app_directory = os.path.dirname(os.path.abspath(sys.argv[0]))
        self.slides_directory = os.path.join(self.app_directory, "slides")
        os.makedirs(self.slides_directory, exist_ok=True)

        self.slide_mappings = {}

        self.config = configparser.ConfigParser()
        self.load_config()

        self.setup_gui()
        self.load_slide_mappings()

    def load_config(self):
        config_path = os.path.join(self.app_directory, 'config.ini')
        if os.path.exists(config_path):
            self.config.read(config_path)
        else:
            self.config['PLC'] = {'ip_address':'192.168.1.100','db_number':'1'}
            with open(config_path,'w') as f:
                self.config.write(f)

    def setup_gui(self):
        main = ttk.Frame(self.root); main.pack(fill='both', expand=True, padx=10, pady=10)
        ttk.Label(main, text="PLC Slideshow Controller", font=('Arial', 16, 'bold')).pack(pady=10)

        # Simulation
        sim_frame = ttk.LabelFrame(main, text="Simulation Mode")
        sim_frame.pack(fill='x', pady=5)
        ttk.Checkbutton(sim_frame, text="Enable Simulation", variable=self.simulation_mode).pack(anchor='w', padx=5)
        sim_inner = ttk.Frame(sim_frame); sim_inner.pack(fill='x', pady=5)
        ttk.Label(sim_inner, text="Index:").pack(side='left')
        self.sim_index = ttk.Entry(sim_inner, width=5); self.sim_index.pack(side='left', padx=5)
        ttk.Button(sim_inner, text="▶ Start", command=lambda: self.simulate_bits(start=True)).pack(side='left', padx=5)
        ttk.Button(sim_inner, text="⏭ Next", command=lambda: self.simulate_bits(next_cmd=True)).pack(side='left', padx=5)

        # Status
        status = ttk.LabelFrame(main, text="Status"); status.pack(fill='x', pady=5)
        row = ttk.Frame(status); row.pack(fill='x')
        ttk.Label(row, text="Conn:").pack(side='left')
        self.conn_lbl = ttk.Label(row, text="Disconnected", foreground="red"); self.conn_lbl.pack(side='left', padx=5)
        ttk.Label(row, text="Idx:").pack(side='left'); self.idx_lbl = ttk.Label(row, text="0"); self.idx_lbl.pack(side='left', padx=5)

        # PLC Controls
        cfg = ttk.LabelFrame(main, text="PLC Controls"); cfg.pack(fill='x', pady=5)
        e1 = ttk.Frame(cfg); e1.pack(fill='x', pady=2)
        ttk.Label(e1, text="PLC IP:", width=10).pack(side='left')
        self.ip_e = ttk.Entry(e1); self.ip_e.insert(0, self.config['PLC']['ip_address']); self.ip_e.pack(side='left', fill='x', expand=True, padx=5)
        e2 = ttk.Frame(cfg); e2.pack(fill='x', pady=2)
        ttk.Label(e2, text="DB Num:", width=10).pack(side='left')
        self.db_e = ttk.Entry(e2, width=5); self.db_e.insert(0, self.config['PLC']['db_number']); self.db_e.pack(side='left', padx=5)
        btns = ttk.Frame(cfg); btns.pack(fill='x', pady=5)
        ttk.Button(btns, text="Connect", command=self.toggle_connection).pack(side='left', padx=5)
        ttk.Button(btns, text="Start Monitoring", command=self.toggle_monitor, state='disabled').pack(side='left')

        # Slide Mapping
        mapf = ttk.LabelFrame(main, text="Slide Mapping"); mapf.pack(fill='both', expand=True, pady=5)
        ctrl = ttk.Frame(mapf); ctrl.pack(fill='x', pady=5)
        ttk.Label(ctrl, text="Index:").pack(side='left')
        self.map_idx = ttk.Entry(ctrl, width=5); self.map_idx.pack(side='left', padx=5)
        ttk.Button(ctrl, text="Add Slideshow", command=self.add_slide_mapping).pack(side='left', padx=5)
        ttk.Button(ctrl, text="Remove Selected", command=self.remove_slide_mapping).pack(side='left')
        cols = ('Index', 'File')
        self.tree = ttk.Treeview(mapf, columns=cols, show='headings')
        self.tree.heading('Index', text='Index'); self.tree.heading('File', text='File')
        self.tree.pack(fill='both', expand=True)

        # Current Slideshow
        curf = ttk.LabelFrame(main, text="Current Slideshow"); curf.pack(fill='x', pady=5)
        self.cur_lbl = ttk.Label(curf, text="None"); self.cur_lbl.pack()

    def simulate_bits(self, start=False, next_cmd=False):
        idx = int(self.sim_index.get() or 0)
        if start:
            self.handle_start_pause(idx)
        if next_cmd:
            self.next_slide()
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
            self.root.nametowidget(".!frame.!labelframe2.!frame.!button2").config(state='normal')
        except Exception as e:
            messagebox.showerror("Conn Err", str(e))

    def disconnect_plc(self):
        if self.is_monitoring:
            self.toggle_monitor()
        if self.plc_client:
            self.plc_client.disconnect()
        self.is_connected = False
        self.conn_lbl.config(text="Disconnected", foreground="red")

    def toggle_monitor(self):
        btn = self.root.nametowidget(".!frame.!labelframe2.!frame.!button2")
        if self.is_monitoring:
            self.is_monitoring = False
            btn.config(text="Start Monitoring")
        else:
            self.is_monitoring = True
            btn.config(text="Stop Monitoring")
            threading.Thread(target=self.monitor_loop, daemon=True).start()

    def monitor_loop(self):
        while self.is_monitoring:
            try:
                db = int(self.db_e.get())
                data = self.plc_client.db_read(db, 0, 2)
                control_byte = data[0]
                idx = data[1]

                start_bit = control_byte & 0b00000001
                next_bit = (control_byte >> 1) & 0b00000001

                if start_bit == 1 and self.last_start_bit == 0:
                    self.root.after(0, lambda: self.handle_start_pause(idx))

                if next_bit == 1 and self.last_next_bit == 0:
                    self.root.after(0, self.next_slide)
                    self.root.after(0, lambda: self.send_feedback_bit(True))
                    self.root.after(200, lambda: self.send_feedback_bit(False))

                self.last_start_bit = start_bit
                self.last_next_bit = next_bit
                self.idx_lbl.config(text=str(idx))

            except Exception as e:
                print(f"[Monitor Error] {e}")

            time.sleep(0.1)

    def add_slide_mapping(self):
        try:
            idx = int(self.map_idx.get())
            assert 0 <= idx <= 255
        except:
            messagebox.showerror("Error", "Enter 0-255")
            return
        f = filedialog.askopenfilename(filetypes=[("PowerPoint Files", "*.pptx *.ppsx")])
        if not f:
            return
        dst = os.path.join(self.slides_directory, os.path.basename(f))
        shutil.copy2(f, dst)
        self.slide_mappings[idx] = dst
        self.save_mappings()
        self.refresh_tree()

    def remove_slide_mapping(self):
        sel = self.tree.selection()
        if not sel:
            return
        idx = int(self.tree.item(sel[0])['values'][0])
        del self.slide_mappings[idx]
        self.save_mappings()
        self.refresh_tree()

    def load_slide_mappings(self):
        p = os.path.join(self.app_directory, 'slide_mappings.txt')
        if os.path.exists(p):
            for L in open(p):
                i, fp = L.strip().split('=', 1)
                if os.path.exists(fp):
                    self.slide_mappings[int(i)] = fp
        self.refresh_tree()

    def save_mappings(self):
        p = os.path.join(self.app_directory, 'slide_mappings.txt')
        with open(p, 'w') as f:
            for i, fp in self.slide_mappings.items():
                f.write(f"{i}={fp}\n")

    def refresh_tree(self):
        for i in self.tree.get_children():
            self.tree.delete(i)
        for i, fp in sorted(self.slide_mappings.items()):
            self.tree.insert('', "end", values=(i, os.path.basename(fp)))

    # PowerPoint control
    def open_ppt(self, path):
        if self.ppt_presentation:
            self.ppt_presentation.Close()
        if not self.ppt_app:
            self.ppt_app = win32com.client.Dispatch("PowerPoint.Application")
            self.ppt_app.Visible = True
        self.ppt_presentation = self.ppt_app.Presentations.Open(path, WithWindow=True)

    def start_slideshow(self):
        if self.ppt_presentation:
            self.ppt_presentation.SlideShowSettings.Run()

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
        if self.ppt_presentation and self.ppt_presentation.SlideShowWindow:
            try:
                self.ppt_presentation.SlideShowWindow.View.Next()
            except Exception as e:
                print(f"[PPT Next Error] {e}")

    def send_feedback_bit(self, state: bool):
        try:
            db = int(self.db_e.get())
            data = self.plc_client.db_read(db, 0, 1)
            byte_val = data[0]
            if state:
                byte_val |= 0b00000100
            else:
                byte_val &= 0b11111011
            self.plc_client.db_write(db, 0, bytes([byte_val]))
        except Exception as e:
            print(f"[PLC Feedback Error] {e}")

    def run(self):
        if not SNAP7_AVAILABLE:
            messagebox.showwarning("Warn", "python-snap7 missing")
        if not PPT_AVAILABLE:
            messagebox.showwarning("Warn", "pywin32 missing")
        self.root.mainloop()


if __name__ == "__main__":
    PLCSlideshowController().run()
