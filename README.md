# PLC Slideshow Controller

A Windows-based application that reads commands from a Siemens S7 PLC (via `snap7.dll`) and controls a PowerPoint slideshow (.PPSX/.PPTX) accordingly.  
Supports **Start/Pause**, **Next Slide** commands, and **feedback bits** to the PLC.

---

## 📦 Features
- 🖥 Works with **Siemens S7 PLCs** via `snap7.dll`.
- 📂 Stores slides locally in the `slides` folder.
- ▶️ Start/Pause slideshow from PLC bit control.
- ⏭ Advance to the next slide on rising-edge PLC signal.
- 🔄 Feedback bit from PC to PLC for slide change confirmation.
- 🖱 Includes a built-in **Simulation Mode** (for testing without PLC).

---

## 📥 Download & Install

1. **Download the latest installer** from the [Releases](../../releases) page.
2. Run the installer — default install path:
   ```
   C:\Program Files\PLC Slideshow Controller
   ```
3. During installation:
   - The correct **64-bit `snap7.dll`** is automatically included.
   - A `slides` folder is created in the install directory.
   - A desktop shortcut is added for quick launch.

---

## 📂 Folder Structure After Installation
```
C:\Program Files\PLC Slideshow Controller\
│   plc_slideshow_controller.exe
│   snap7.dll
│
└── slides\
    └── (place your .PPTX or .PPSX files here)
```

---

## 🚀 Running the Application
1. Double-click **PLC Slideshow Controller** from your desktop or Start Menu.
2. In **Simulation Mode**, you can trigger PLC-like signals from the GUI.
3. In **Live Mode**, connect to the PLC (ensure IP, rack, slot are set correctly in the config).

---

## ⚙️ PLC Bit Mapping
- **Byte 0, Bit 0** → Start/Pause slideshow
- **Byte 0, Bit 1** → Next Slide command
- **Byte 0, Bit 2** → Feedback to PLC after slide change

---

## 🛠 Configuration
The application creates a `config.ini` file automatically on first run.  
You can edit this file to change PLC IP address, rack, slot, and other settings.

---

## ❗ Troubleshooting
- **"snap7.dll not found"** → Ensure `snap7.dll` is in the same folder as `plc_slideshow_controller.exe`.
- **Permission errors** → Run as Administrator if installed in Program Files and you need to save config there.
- **PLC not reachable** → Check network connection and PLC IP configuration.

---

## 📜 License
This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

## 👨‍💻 Author
Developed by [Your Name / Company]  
For support, contact: `youremail@example.com`
