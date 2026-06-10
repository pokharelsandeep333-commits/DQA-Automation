# Device Quality Assurance (DQA) Automation Tool

A complete, standalone PowerShell GUI designed to streamline hardware inspection and quality assurance workflows for IT support technicians. 

Originally developed to optimize help desk operations, this tool replaces manual laptop testing checklists with a streamlined dashboard that combines automated system detection, interactive hardware tests, and cosmetic grading into a single, user-friendly interface.

## ✨ Features

* **Automated System Detection:** Immediately detects and logs the device's Serial Number, Battery/Charging status, and Network Adapters the moment a technician is selected—no manual prompts required.
* **Interactive Hardware Diagnostics:** * **Audio/Mic Test:** Automatically unmutes the system, sets volume exactly to 50%, and records/plays back a 10-second mic test.
  * **Camera & Keyboard:** One-click launch buttons for local camera checks and web-based keyboard ghosting tests.
* **Self-Contained Local Database:** The script acts as its own database, rewriting itself to store the CSV array internally. This allows the tool to be entirely portable on a USB drive without requiring SQL or external file dependencies.
* **Custom CSV Exports:** Easily export session data to a local drive or USB. The export function is tailored for clean reporting, intentionally stripping out internal tracking columns (like `Id`, `start date`, and `Status`) to provide a clean data set.
* **Live Analytics Dashboard:** Track daily productivity with real-time metrics, including Total Runs, Pass/Fail ratios, and the Average Duration of each inspection.
* **Secure Environment:** Features a PIN-protected (default: `5555`) wipe function to clear local script memory between deployment batches.

## 🚀 Prerequisites

* **OS:** Windows 10 or Windows 11
* **Framework:** Windows PowerShell 5.1+
* **Permissions:** No administrative privileges required for basic operation (though some WMI queries may require elevation depending on group policy).

## 🛠️ Usage

1. Clone or download the repository to your local machine or a portable USB drive.
2. Right-click the `DQA_GUI.ps1` file and select **Run with PowerShell**.
3. **Pre-Check:** Ensure the device is connected to a network and plugged into power.
4. **Begin Inspection:** * Select or type a technician email in the dropdown.
   * *Note: Providing an email automatically triggers the hardware detection for Serial Number, Network, and Charging status.*
5. Run through the interactive tests (Camera, Audio, Keys) and log the physical/cosmetic conditions using the dropdowns.
6. Click **Save Inspection** to commit the record to the local dashboard.

## 📂 Exporting Data

Navigate to the **Dashboard & History** tab to view all saved inspections. Click **Export to CSV** to automatically generate a clean output file to your desktop or a connected USB drive. 

## 🛡️ Privacy & Security

This script is designed for internal network operations. Ensure that no sensitive technician emails or actual network credentials are hardcoded into the `$Global:SavedEmails` array if cloning or forking this repository for public use.

---
*Built for fast, consistent, and reliable IT deployments.*