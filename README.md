 Keithley 2450 — I-V Studio

Professional I-V Characterisation and FET Analysis Desktop Application  
Developed by Rangaraajan Muralidaran

---

 Overview

Keithley 2450 I-V Studio is a Python/tkinter desktop application for instrument-controlled electrical characterisation using the Keithley 2450 Source Measure Unit (SMU). It provides a complete workflow from instrument connection through measurement execution to data export, with a dedicated FET Characterisation module for MOSFET analysis via dual-SMU TSP-Link.

---

 Features

- Voltage and Current Sweeps — Linear, logarithmic, and list-based
- Dual Sweep — Forward and reverse pass for hysteresis analysis
- Pulse Mode — Configurable on/off timing for pulsed measurements
- Stepper Mode — Node 2 steps a bias while the master SMU sweeps, producing a family of curves
- Live Plot — Real-time chart updates as data is acquired, with scientific notation axes
- FET Characterisation — Transfer Curve (Id vs Vgs) and Output Curve (Id vs Vds) via TSP-Link dual-SMU
- Q-Point Analysis — gm computation on Transfer Curves; load line + Pd hyperbola on Output Curves
- Auto-save — CSV and PNG exported automatically on sweep completion
- Flexible Axis Selector — X, Y1, and Y2 axes independently assignable to any measured channel

---

 Screenshots

> Transfer Curve — Id vs Vg (MOSFET)

> Output Curve — Id vs Vds with Q-Point load line overlay

(Add screenshots to the repository and update these links)

---

 System Requirements

| Component | Requirement |
|---|---|
| Operating System | Windows 10 / 11 (64-bit) |
| Python | 3.9 or later |
| VISA Runtime | NI-VISA 21+ or Keysight IO Libraries Suite |
| Instrument | Keithley 2450 SMU (firmware 1.7.0+) |
| FET Mode | Second Keithley 2450 + TSP-Link RJ-45 cable |
| Display | 1280 × 720 minimum, 1920 × 1080 recommended |

---

 Installation

 1. Install NI-VISA Runtime
Download and install the NI-VISA runtime from:  
https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html

This is required for USB communication with the instrument. After installing, open NI-MAX and verify the Keithley appears under Devices and Interfaces with a resource string like:
```
USB0::0x05E6::0x2450::04465297::INSTR
```

 2. Install Python
Download Python 3.12 (64-bit) from https://www.python.org/downloads/  
Important: Tick "Add Python to PATH" during installation.

 3. Install Python Dependencies
```bash
pip install pyvisa pyvisa-py numpy matplotlib
```

 4. Run the Application
```bash
python keithley_pro.py
```

---

 Building a Standalone .exe

To distribute the application without requiring Python on the target machine:

```bash
pip install pyinstaller
pyinstaller --onefile --windowed --name "IV_Studio" keithley_pro.py
```

The compiled executable will be located at:
```
dist\IV_Studio.exe
```

> Note: NI-VISA runtime must still be installed on the target PC — hardware drivers cannot be bundled.

---

 FET Mode — TSP-Link Setup

FET Characterisation requires two Keithley 2450 units connected via TSP-Link.

Physical Connections:
```
Gate   → SMU1 (Master, Node 1)  Rear FORCE HI / FORCE LO
Drain  → SMU2 (Slave,  Node 2)  Rear FORCE HI / FORCE LO
Source → GNDU                   Rear GNDU post
```

Set Node IDs on each instrument front panel:
```
Master: MENU → System → TSP-Link Node → 1
Slave:  MENU → System → TSP-Link Node → 2
```

Only the Master VISA address is needed in the application.

---

 Usage

 Main Window — I-V Sweep
1. Enter your VISA address in the header bar
2. Set Source Mode, sweep range, compliance, and NPLC
3. Select channels on the Axis Selector panel (right sidebar)
4. Press RUN SWEEP
5. Data auto-saves to `D:\KEITHLEY 2450\Data` on completion

 FET Characterisation
1. Click FET Characterisation ▶ in the header bar
2. Assign Gate, Drain, Source to their respective SMUs
3. Configure sweep parameters for the active terminal
4. Press RUN
5. Use Log Y Scale for subthreshold region analysis
6. Enable Q-Point Analysis to overlay load line and operating point

---

 Data Export

 Main Window CSV Format
```
 Keithley 2450  I-V Studio  v3.0
 Sample ID,  DUT_001
 Date/Time,  05/03/2026  14:32:17
 Mode,       Voltage Sweep

Time (s), Voltage (V), Current (A), Resistance (Ω), Power (W)
0.000000, 0.000000000, 0.000012450, ...
```

 FET Window CSV Format
```
 Keithley 2450  I-V Studio  v3.0  —  FET Characterisation
 Mode,  Transfer Curve (Id vs Vg)

Vg (V), Id  Vds = 1.5 V (A), Id  Vds = 3.5 V (A)
0.000000, 0.000000012, 0.000000015
```

---

 Troubleshooting

| Problem | Solution |
|---|---|
| VISA error / instrument not found | Verify resource string in NI-MAX. Check USB cable. |
| TSP-Link not online | Check cable. Set Node ID=1 on Master, Node ID=2 on Slave. |
| Current flat at compliance | Increase Compliance value or reduce sweep range. |
| Noisy Transfer Curve | Increase Gate Source Delay to 0.15–0.20 s. |
| Blank plot after sweep | Click Auto Scale. Confirm Current (A) is ticked on Y1. |
| Auto-save failed | Create folder `D:\KEITHLEY 2450\Data` or change `self.save_dir` in the script. |

---

 File Structure

```
keithley-iv-studio/
├── keithley_pro.py               Main application
├── Keithley_2450_-_Manual.pdf    Full user manual
├── Keithley_IV_Studio_Setup_Guide.txt   Installation guide
├── README.md                     This file
└── LICENSE                       MIT License
```

---

 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

 Author

Rangaraajan Muralidaran  
Keithley 2450 — I-V Studio  
Version 3.0 · March 2026
