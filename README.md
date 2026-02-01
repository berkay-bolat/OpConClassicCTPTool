# OpConClassicCTPTool

AUTOMATED COMMISSIONING TEST PROTOCOL (CTP) TOOL

An advanced desktop utility designed to automate the analysis, tracking, and reporting of Commissioning Test Protocols (CTP) for industrial PLC projects (specifically OpCon Classic systems).

# OVERVIEW

Commissioning industrial automation systems often involves manually checking hundreds of IOs, manual functions, and sequences. This tool automates the ingestion of PLC export files (`.EXP`), allows engineers to track testing progress digitally, and generates professional Excel reports for customer handover.

# KEY FEATURES

Automated Parsing: Uses Regular Expressions (RegEx) to parse standard OpCon export files automatically.

Real-time Progress Tracking: Visualizes completion rates for IOs, Manual Operations, and Sequences with dynamic progress bars.

Excel Reporting: Exports comprehensive, formatted reports using `Pandas` and `OpenPyXL`, ready for client presentation.

User Accountability: Automatically tags every check/modification with the current system username to ensure traceability.

State Management: Save and load project progress via JSON files.

Modern UI: Features a responsive tabbed interface with multiple color themes.

Thread-Safe: Implements `QThread` and Signals effectively.

# INSTALLATION

Install Dependencies: pip install -r requirements.txt

Run The App: python ctptool.py

# USAGE

New Project: Click "New Project" and select the root folder containing your `.EXP` files.

Track Progress: Overview Tab -> Monitor global progress. | IO/Manual/Sequence Tabs: Mark items as `OK`, `X` (Fail), or `N/A`. Add comments where necessary.

Save/Load: Use "Save Data" to store your current session as a `.json` file.

Export: Click "Export Data" to generate an `.xlsx` file with formatted columns and progress summaries.

# DISCLAIMER

This app is developed for specific needs. Not allowed to share '.EXP' file samples with you to try out the app throughly.