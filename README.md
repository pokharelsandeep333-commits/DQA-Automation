# 🛡️ DQA Automation Suite

[![PowerShell](https://img.shields.io/badge/PowerShell-5.1-blue.svg?logo=powershell)](#)
[![WPF](https://img.shields.io/badge/UI-WPF%2FXAML-brightgreen.svg)](#)
[![C#](https://img.shields.io/badge/Core-C%23-purple.svg?logo=csharp)](#)
[![Status](https://img.shields.io/badge/Status-Production-success.svg)](#)

> **Automated Device Quality Assurance for DSU IT**

The **DQA Automation Suite** (`DQA_GUI_v10.0.ps1`) is a highly specialized, self-contained PowerShell-driven graphical application designed to streamline hardware diagnostics and quality assurance processes for Dakota State University (DSU) IT. 

By leveraging a custom WPF interface, inline C# execution, and deep system querying, this tool provides a rapid, reliable, and standardized method for technicians to validate laptop system health without requiring external databases or cumbersome dependencies.

---

## ✨ Key Features

- 🚀 **Zero-prompt Hardware Auto-detection:** Seamlessly scans and identifies system hardware (Network adapters, USB ports, Video ports, Touchscreens) using advanced `WMI/CIM` queries without manual input.
- 🎵 **Inline C# Core Audio API:** Bypasses native PowerShell limitations by directly interfacing with the Windows Core Audio API to manipulate master volume, auto-unmute, and comprehensively test audio output and microphone recording.
- 💾 **Self-modifying Persistence Engine:** Maintains an internal state by rewriting its own script logic during runtime, ensuring data persistence and technician profiles are saved without needing a SQL database or external config files.
- ⌨️ **Custom WPF Interactive Testers:** Features bespoke, interactive XAML canvas windows for:
  - **Keyboard Testing:** Native key-interception overlay to test every physical key.
  - **Touchscreen Validation:** A 5-second dynamic tap challenge to verify digitizer accuracy and responsiveness.
- 📊 **Robust Dashboard & CSV Export:** Built-in dashboard to search, filter, and edit past inspections, with secure single-click routing to export standardized CSV reports to connected USB drives.
- 🔐 **PIN-Based Authorization:** Administrative actions (like database wipes or record deletions) are safeguarded by SHA-256 hashed PIN verification.

---

## 📖 User Guide (For Technicians)

### Prerequisites
- **Windows 10/11**
- **PowerShell 5.1** (Run as Administrator recommended for full WMI access)

### 🛠️ Running an Inspection

1. **Launch the Application:** Right-click `DQA_GUI_v10.0.ps1` and select **Run with PowerShell**.
2. **Technician & Serial Entry:** The tool will auto-focus the email input. Select your email or type a new one. Enter the device Serial Number.
3. **Automated & Interactive Testing:** 
   - Click the **"Auto-detect hardware"** button to automatically scan for external monitors and USBs.
   - Use the **"Test Camera"**, **"Test Keyboard"**, and **"Test Touchscreen"** buttons to open interactive verification prompts.
   - Click **"Start Audio & Mic Test"** to automatically record 5 seconds of audio and play it back to test the internal speakers and microphone simultaneously.
4. **Cosmetic & Final Review:** Manually evaluate the physical condition (Palm Rest, Backplate, Base, Hinge) using the provided dropdowns. Use the **"Auto-fill Remaining"** button (Ctrl+A) to mark any untested standard components as "Operational".
5. **Save & Export:** 
   - Click **"SAVE INSPECTION"** (Ctrl+S). The dashboard will update instantly.
   - When finished with a batch of laptops, plug in a USB drive and click **"Export to USB"** (Ctrl+E) from the Dashboard tab.

### ⌨️ Keyboard Shortcuts
- `Ctrl + 1`: Switch to Inspection Tab
- `Ctrl + 2`: Switch to Dashboard Tab
- `Ctrl + S`: Save Inspection
- `Ctrl + R`: Reset Form
- `Ctrl + E`: Export Dashboard to CSV
- `Ctrl + F`: Focus Search Bar
- `Ctrl + A`: Auto-fill remaining fields

---

## 🏗️ Technical Architecture (For Maintainers)

The `DQA_GUI_v10.0.ps1` script is a monolithic, self-contained architecture that combines PowerShell, XAML, and C# to maximize deployment ease across various diagnostic environments.

### 1. WPF GUI & DPI Awareness
The graphical interface is built entirely with inline XAML loaded via `[Windows.Markup.XamlReader]::Load()`. To fix graphical blurring on high-resolution displays (common on modern laptops), the script dynamically imports `user32.dll` to call `SetProcessDPIAware()`, ensuring crisp rendering regardless of Windows scaling settings.

### 2. Self-Modifying State (`Update-ScriptData`)
To eliminate the need for an external database file that could be lost or corrupted on a diagnostic bench, the script uses a self-modifying engine. The `$Global:DatabaseCSV` and `$Global:SavedEmails` variables are hardcoded at the top of the file. When a user saves an inspection, the `Update-ScriptData` function reads the script's own source code, uses RegEx to find the data block, injects the updated CSV data, and overwrites the `.ps1` file in place.

### 3. Inline C# Core Audio API Integration
Native PowerShell cannot directly manipulate the master volume to an exact percentage or force an un-mute without third-party modules. The script embeds a C# class (`AudioHelper`) compiled at runtime via `Add-Type`. This class imports `ComImport` interfaces (`IAudioEndpointVolume`, `IMMDeviceEnumerator`) from the Windows Core Audio API, allowing the script to instantly set the system volume to 25% and unmute before executing the microphone recording test via `winmm.dll` (mciSendString).

### 4. Custom Hardware Testers
- **Keyboard Tester:** A custom XAML window that hooks into the `PreviewKeyDown` event. It intercepts `System.Windows.Input.Key` enumerations, maps them to the corresponding visual border on the screen, and turns them green (`#39FF14`). Default key behaviors (like Enter or Escape closing the window) are programmatically suppressed (`$e.Handled = $true`).
- **Touchscreen Tester:** A dynamic `<Canvas>` element that spawns colored ellipses in random, non-overlapping coordinates. A `DispatcherTimer` enforces a strict 5-second countdown. It differentiates between standard mouse clicks and actual touch inputs using the `TouchDown` event to ensure the physical digitizer is functioning, not just the mouse cursor.

### 5. Security & Authorization
Administrative actions, such as clearing the dashboard, require a PIN. The script does not store the PIN in plain text. Instead, it compares the SHA-256 hash of the user's input against `$Global:AdminPinHash`. 

*Developed for Dakota State University (DSU) IT.*
