# DQA Automation Suite

> **Automated Device Quality Assurance for DSU IT**

The **DQA Automation Suite** is a highly specialized PowerShell-driven graphical application designed to streamline hardware diagnostics and quality assurance processes for Dakota State University (DSU) IT. By leveraging a custom WPF interface and advanced system querying, this tool provides a rapid, reliable, and standardized method for technicians to validate system health.

## 🚀 Key Features

- **Zero-prompt Hardware Auto-detection:** Seamlessly scans and identifies system hardware without requiring manual input or interruptions.
- **Inline C# Core Audio API:** Directly interfaces with the Windows Core Audio API via embedded C# code to comprehensively test and validate audio devices.
- **Self-modifying Persistence Engine:** Employs an advanced self-updating mechanism to ensure script logic remains persistent and adaptable across sessions.
- **Comprehensive Data Export:** Generates robust, standardized CSV reports tailored specifically for DSU IT data management requirements.
- **Intuitive GUI:** Built with WPF/XAML, providing a clean, responsive, and user-friendly experience for IT staff.

## 🛠️ Technical Stack

- **PowerShell 5.1:** The core engine driving automation and script execution.
- **WPF/XAML:** Provides the modern, responsive graphical user interface.
- **C# COM Interop & Core Audio API:** Utilized inline for advanced, low-level hardware diagnostics not natively exposed to PowerShell.
- **WMI/CIM:** Extensively used for rapid, deep hardware telemetry and system health queries.

## ⚙️ How It Works

The tool initializes by bootstrapping a customized WPF graphical interface from within a PowerShell host. Upon launching an inspection, the suite initiates a series of automated hardware diagnostics using WMI/CIM queries and custom C# wrappers (like the Core Audio API). 

A unique **self-modifying script logic** allows the tool to maintain state and handle complex execution sequences gracefully. As diagnostics run, the results are collected and evaluated against defined quality standards, presenting real-time feedback to the technician.

## 🔐 Security & Best Practices

- **PIN-Based Authorization:** Administrative actions, such as critical data deletions or configuration resets, are safeguarded by a secure, PIN-based authorization layer to prevent accidental execution.
- **Secure CSV Export Routing:** Inspection results and sensitive device telemetry are routed and saved using secure, predictable paths, ensuring data integrity and compliance with internal protocols.

## 📖 Usage Instructions

1. **Launch the Application:** Run the `DQA_GUI_v10.0.ps1` script. 
2. **Initiate Inspection:** Click the **"Run Inspection"** button on the main dashboard to begin the zero-prompt automated hardware diagnostic sequence.
3. **Review Results:** Wait for the diagnostic process to complete. The dashboard will populate with real-time pass/fail metrics and detailed hardware telemetry.
4. **Export Data:** Click the **"Export Results"** button to generate a standardized CSV report. The file will be securely routed to the designated directory for archival or further analysis.
5. **Administrative Tasks (Optional):** If an administrative action (like data deletion) is necessary, navigate to the relevant section and enter the authorized PIN when prompted.

---
*Developed for Dakota State University (DSU) IT.*
