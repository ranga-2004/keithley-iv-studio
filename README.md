# Keithley 2450 — Studio

**Professional I-V Characterisation and FET Analysis Desktop Application**  
Developed by Rangaraajan Muralidaran

---

## Overview

Keithley 2450 I-V Studio is a Python/tkinter desktop application for instrument-controlled electrical characterisation using the Keithley 2450 Source Measure Unit (SMU). It provides a complete workflow from instrument connection through measurement execution to data export, with three dedicated measurement modes selectable from a full-screen launcher.

---

## What's New in v4.0

| Feature | Detail |
|---|---|
| **Mode Selector** | Full-screen launcher — choose IV, FET, or DMM on startup |
| **Cycling & Stability** | Repeat sweeps N times with configurable rest, gradient colours, Mean±σ band |
| **Bias Stress** | Apply fixed voltage between cycles and log current drift over time |
| **Parameter Extraction** | Auto-compute Vth, SS, Ion/Ioff, hysteresis; cycle trend (ΔVth, Ion drift) |
| **Measurement Quality** | Per-point averaging, settling check with tolerance, compliance pre-check |
| **Ghost Overlay** | Previous run displayed as grey reference behind live data |
| **Compare CSV Manager** | Load multiple historical CSVs as dashed overlays; toggle/rename/recolour |
| **DMM Mode** | Live DC Voltage, DC Current, 2W/4W Resistance, Diode, Continuity, Power |
| **Dark / Light Theme** | Toggle per-window; PNGs always saved in light mode |
| **Persistent Config** | VISA address, save folder, sample ID, and theme saved across sessions |

---

## Features

**I-V Characterisation (IV mode)**
- Voltage and current sweeps — linear, logarithmic, and list-based
- Dual sweep — forward and reverse pass for hysteresis analysis
- Pulse mode — configurable on/off timing for pulsed I-V
- Stepper mode — Node 2 steps a gate/drain bias while the master SMU sweeps
- Cycling & Stability — N-cycle repeat with rest time, gradient colouring, Mean±σ band overlay
- Ghost overlay — previous full run shown as faded reference
- Compare CSV — load any historical CSV as a dashed overlay for direct comparison
- Parameter extraction — Vth (linear extrapolation), SS (mV/dec), Ion/Ioff, hysteresis area
- Live plot — real-time chart updates with engineering-notation axes and flexible X/Y1/Y2 selector

**FET Characterisation (FET mode)**
- Transfer Curve (Id vs Vgs) and Output Curve (Id vs Vds) via TSP-Link dual-SMU
- Q-Point Analysis — gm on Transfer Curves; load line + Pd hyperbola on Output Curves
- Cycling, bias stress, parameter extraction, and Compare CSV — all available in FET mode
- Per-terminal sweep configuration with auto/manual point count and dual-sweep support
- Log Y scale for subthreshold analysis

**Digital Multimeter (DMM mode)**
- Seven modes: DC Voltage, DC Current, 2W Resistance, 4W Resistance, Diode, Continuity, Power (V×I)
- Hold, Min/Max tracking, front/rear terminal selector, NPLC rate control
- Built-in wiring guide dialog for each mode

**Common**
- Auto-save CSV and PNG on measurement completion
- Dark and light theme with persistent preference
- Measurement Quality: per-point averaging (1–16×), settling check, compliance pre-check

---

## System Requirements

| Component | Requirement |
|---|---|
| Operating System | Windows 10 / 11 (64-bit) |
| Python | 3.9 or later |
| VISA Runtime | NI-VISA 21+ or Keysight IO Libraries Suite |
| Instrument | Keithley 2450 SMU (firmware 1.7.0+) |
| FET Mode | Second Keithley 2450 + TSP-Link RJ-45 cable |
| Display | 1280 × 720 minimum, 1920 × 1080 recommended |

---

## Installation

**1. Install NI-VISA Runtime**

Download and install from:  
https://www.ni.com/en/support/downloads/drivers/download.ni-visa.html

After installing, open NI-MAX and verify the Keithley appears under Devices and Interfaces with a resource string like:
```
USB0::0x05E6::0x2450::04465297::INSTR
```

**2. Install Python**

Download Python 3.12 (64-bit) from https://www.python.org/downloads/  
> Tick **"Add Python to PATH"** during installation.

**3. Install Python Dependencies**
```
pip install pyvisa pyvisa-py numpy matplotlib
```

**4. Run the Application**
```
python keithley_pro.py
```

The splash screen loads, then the mode selector appears. Choose **IV**, **FET**, or **DMM**.

---

## Building a Standalone .exe

To distribute without requiring Python on the target machine:
```
pip install pyinstaller
pyinstaller --onefile --windowed --name "IV_Studio" keithley_pro.py
```

Output:
```
dist\IV_Studio.exe
```

> NI-VISA runtime must still be installed on the target PC — hardware drivers cannot be bundled.

---

## FET Mode — TSP-Link Setup

FET Characterisation requires two Keithley 2450 units connected via TSP-Link.

**Physical Connections**
```
Gate   → SMU1 (Master, Node 1)  Rear FORCE HI / FORCE LO
Drain  → SMU2 (Slave,  Node 2)  Rear FORCE HI / FORCE LO
Source → GNDU                   Rear GNDU post
```

**Set Node IDs on each instrument front panel**
```
Master: MENU → System → TSP-Link Node → 1
Slave:  MENU → System → TSP-Link Node → 2
```

Only the Master VISA address is entered in the application.

---

## Usage

### Mode Selector
On launch, a full-screen selector presents three tiles — IV, FET, and DMM. Click a tile or its button to open that mode. Closing any mode window returns to the selector.

### I-V Sweep (IV Mode)
1. Enter your VISA address in the header bar
2. Set Source Mode, sweep range, compliance, and NPLC
3. Select measurement channels on the Axis Selector sidebar (right)
4. Optionally enable **Cycling** and/or **Parameter Extraction** in the left panel
5. Press **RUN SWEEP**
6. Data auto-saves as CSV + PNG on completion

### FET Characterisation (FET Mode)
1. Select **Transfer Curve** or **Output Curve** from the mode bar
2. Assign Gate, Drain, Source to their respective SMUs in the parameter table
3. Configure sweep parameters for each active terminal
4. Press **RUN**
5. Enable **Log Y Scale** for subthreshold analysis
6. Expand **Q-Point Analysis** panel to overlay load line and operating point

### DMM Mode
1. Enter your VISA address in the settings bar
2. Select measurement mode from the mode buttons (DCV, DCI, R 2W, etc.)
3. Click **► Connect** — live readings begin immediately
4. Use **HOLD** to freeze the display and **MIN/MAX** to track extremes

---

## Data Export

### Long-format CSV (single run or cycling disabled)
```
# Keithley 2450  I-V Studio  v4.0
# Sample ID,  DUT_001
# Date/Time,  05/03/2026  14:32:17
# Mode,       Voltage Sweep
# Extracted,  Vth = 2.14 V    SS = 87 mV/dec    Ion/Ioff = 1.23M

Time (s), Voltage (V), Current (A)
0.000000, 0.000000000, 0.000012450
...
```

### Wide-format CSV (cycling enabled, N > 1)
```
# Keithley 2450  I-V Studio  v4.0
# Cycles,  5

Voltage (V), Fresh  I (A), +20°C  I (A), Stressed  I (A), ...
0.0,         1.23e-12,     1.24e-12,      1.31e-12,        ...
```

### FET CSV
```
# Keithley 2450  I-V Studio  v4.0  —  FET Characterisation
# Mode,  Transfer Curve (Id vs Vg)

Vg (V), Id  Vds = 1.5 V (A), Id  Vds = 3.5 V (A)
0.000000, 0.000000012, 0.000000015
```

---

## Troubleshooting

| Problem | Solution |
|---|---|
| VISA error / instrument not found | Verify resource string in NI-MAX. Check USB cable. |
| TSP-Link not online | Check cable. Set Node ID=1 on Master, Node ID=2 on Slave. |
| Current flat at compliance | Increase Compliance value or reduce sweep range. |
| Noisy Transfer Curve | Increase Gate Source Delay to 0.15–0.20 s. |
| Blank plot after sweep | Click Auto Scale. Confirm Current (A) is ticked on Y1. |
| Auto-save failed | Create folder `D:\KEITHLEY 2450\Data` or update the path in Save Folder. |
| IV window appears doubled | Fixed in v4.0 — close and reopen the application. |
| FET Export CSV empty after cycling | Fixed in v4.0 — update to the latest `keithley_pro.py`. |
| Parameter Extraction returns N/A | Fixed in v4.0 — `_extract_vth_iv` was unreachable in prior builds. |

---

## File Structure

```
keithley-iv-studio/
├── keithley_pro.py                    Main application (v4.0)
├── Keithley 2450 - Manual.pdf         Full instrument user manual
├── Keithley_IV_Studio_Setup_Guide.txt Installation and wiring guide
├── README.md                          This file
└── LICENSE                            MIT License
```

---

## License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

## Author

Rangaraajan Muralidaran  
Keithley 2450 — I-V Studio  
**Version 4.0 · March 2026**
