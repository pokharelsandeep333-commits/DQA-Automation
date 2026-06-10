# Device Quality Assurance (DQA) Automated GUI

A PowerShell-based automated hardware testing and cosmetic inspection dashboard. This tool is designed to streamline device intake, automate diagnostic checks, and maintain a persistent database of all inspections. 

**Author:** Sandeep Prasad Pokharel

## Features
* **Auto-Trigger Hardware Detection:** Automatically pulls the device serial number, battery charging status, and network adapter status the moment a technician enters their email.
* **Interactive Diagnostic Tests:** One-click launches for Windows Camera, web-based keyboard checkers, and an asynchronous audio/microphone recording test.
* **Smart Autofill & Persistent Memory:** Automatically expands common email prefixes and remembers new technician emails for future sessions.
* **Live Session Tracking:** Calculates and logs the exact time spent inspecting each device.
* **Dashboard & Export:** Built-in dashboard to view all historical runs, filter by Pass/Fail status, and export results directly to a CSV file.

## Prerequisites
* **Windows OS** (Relies on WMI and Windows-specific diagnostic tools)
* **PowerShell 5.1** or newer
* **SQLite Module:** The script requires the accompanying `DQA_Database.psm1` to function.

## Installation & Setup
1. Download or clone this repository to your local machine or a portable USB drive.
2. Ensure both `DQA_GUI.ps1` and `DQA_Database.psm1` are in the same folder.
3. The script will automatically generate a `DQA_History.db` and `DQA_Emails.txt` file in the same directory upon first run.

## Usage
1. Right-click `DQA_GUI.ps1` and select **Run with PowerShell**.
2. Select your email from the dropdown (or type a new one) to trigger auto-detection.
3. Run through the interactive hardware tests.
4. Fill out the visual/cosmetic inspection fields.
5. Click **Save Inspection** to log the data and clear the form for the next device.

## Links
Github: https://github.com/pokharelsandeep333-commits
Linkedin: https://www.linkedin.com/in/sandeeppokharel333