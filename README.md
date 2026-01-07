Stock Count Generator
A Windows desktop application for generating warehouse stock counting sheets (Excel or PDF) from WMS inventory exports, with optimized aisle-based counting order designed for cold storage warehouses.

✨ Features
📦 Supports WMS inventory Excel exports (Chinese / English column names auto-detected)

🧊 Cold storage–oriented logic
C Zone (double-deep racks with aisle-based counting order)
D Zone (layer-first, aisle-increasing order)

🧾 Two counting modes
By Bin Range: generate full stock count sheets with empty bins included
By Product (SKU): generate count sheets only for selected products

📄 Export formats
PDF (A4, auto-paging, Chinese font support)
Excel (.xlsx)

🖥 User-friendly GUI
Search and filter SKUs
Zone selection (C / D / C+D)
Adjustable rack ranges (aisle / row / level)

🚀 Packaged as a Windows application
Custom application icon
Splash screen on startup
Installer-style deployment supported

🧠 Counting Logic Overview
C Zone (Freezer – Double-deep racks)
Bin format: C{Aisle}-{Row}-{Level}

Counting order:
Group aisles in pairs: (1,2), (3,4), …

Within each group:
Count level by level (bottom → top)
Rows counted in pairs (row 1&2, then 3&4, etc.)
Follow aisle-side priority to minimize walking distance

Special rule:
Aisles C3–C14 have maximum 9 rows
Other aisles follow the full configured row range
D Zone (Freezer – Shuttle racks)

Bin format: D-{Aisle}B-{Level}

Counting order:
Level first (low → high)
Aisle second (small → large)

📂 Input Requirements
Excel file exported from WMS inventory detail page
Supported columns (auto-detected):
SKU
Product Name (CN)
Product Name (EN)
Bin / Location
Quantity
Only bins with prefix Freezer(C,D) are processed

📤 Output
PDF
A4 format
Auto column width adjustment
Chinese font support
Generation timestamp in footer

Excel
One sheet per export
Frozen headers
Auto-adjusted column widths

🖱 How to Use
Export inventory detail Excel from WMS
Launch Stock Count Generator
Click “Select Excel” and load the file

Choose:
Mode: By Bin Range or By Product
Zone: C / D / C+D
Export format: PDF or Excel
Configure ranges or select SKUs
Click “Generate File”

🪟 Windows Packaging
The application is packaged using PyInstaller and can be distributed as:
A standalone .exe
A full Windows installer (via Inno Setup)
The packaged application includes:
Custom window & taskbar icon
Splash screen
Embedded resources (icons, fonts)

🛠 Tech Stack
Python 3
PySide6 (Qt for Python)
Pandas
ReportLab (PDF generation)
OpenPyXL (Excel export)
PyInstaller (Windows packaging)

📜 License

This project is intended for internal warehouse operations and automation workflows.
Licensing terms can be adapted for internal, commercial, or open-source use as needed.

👤 Author

Developed for real-world cold storage warehouse operations, focusing on efficiency, accuracy, and operator usability.
