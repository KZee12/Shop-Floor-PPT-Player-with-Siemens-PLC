
# Shop Floor PPT Player with Siemens PLC

This project enables automated PowerPoint (PPT) slide presentations on shop floor displays, controlled via a Siemens PLC. It integrates industrial automation with digital signage to enhance real-time communication and operational efficiency on the manufacturing floor.

## 🛠️ Features

- **Automated Slide Control**: Navigate through PowerPoint slides using PLC commands.  
- **Real-Time Integration**: Seamless communication between Siemens PLC and the presentation system.  
- **Customizable Configuration**: Easily adjustable settings to match specific shop floor requirements.  
- **Plug-and-Play Setup**: Minimal configuration needed for deployment.

## ⚙️ Requirements

- **Hardware**:  
  - Siemens PLC (e.g., S7-1200, S7-1500)  
  - PC or Raspberry Pi running Windows or Linux  
  - Display connected to the PC/Raspberry Pi  

- **Software**:  
  - Python 3.7+  
  - `python-pptx` library  
  - `snap7` library for Siemens PLC communication  
  - Microsoft PowerPoint installed on the PC  

## 📂 Project Structure

```
Shop-Floor-PPT-Player-with-Siemens-PLC/
│
├── slides/                       # Directory containing PPT files
├── config.ini                   # Configuration file for PLC settings
├── plc_slideshow_controller.py  # Main script to control PPT slides via PLC
├── slide_mappings.txt           # Mapping of PLC inputs to PPT slide actions
└── snap7.dll                    # Siemens PLC communication library (Windows)
```

## 🚀 Installation

1. **Clone the Repository**:

   ```bash
   git clone https://github.com/KZee12/Shop-Floor-PPT-Player-with-Siemens-PLC.git
   cd Shop-Floor-PPT-Player-with-Siemens-PLC
   ```

2. **Install Dependencies**:

   ```bash
   pip install python-pptx snap7
   ```

3. **Configure PLC Settings**:

   - Edit the `config.ini` file to match your PLC's IP address and rack/slot configuration.

4. **Prepare PowerPoint Slides**:

   - Place your PPT files in the `slides/` directory.  
   - Ensure the slides are named sequentially (e.g., `slide1.pptx`, `slide2.pptx`, etc.).

5. **Map PLC Inputs to Slide Actions**:

   - Edit the `slide_mappings.txt` file to define how PLC inputs correspond to slide actions (e.g., next slide, previous slide).

6. **Run the Controller Script**:

   ```bash
   python plc_slideshow_controller.py
   ```

   This script will start the slide presentation and listen for PLC inputs to control the slides.

## 🔧 Configuration

- **config.ini**:

  ```ini
  [PLC]
  ip_address = 192.168.0.1
  rack = 0
  slot = 1
  ```

- **slide_mappings.txt**:

  ```txt
  0 = next_slide
  1 = previous_slide
  2 = pause_slide
  3 = resume_slide
  ```

## 🖥️ Ready-to-Install Setup

If you prefer a ready-to-install setup file for quick deployment, please contact:  
**info@kurosystems.net**

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

Contributions are welcome! Please fork the repository, make your changes, and submit a pull request.

## 📞 Support

For issues or questions, please open an issue on the GitHub repository page.
