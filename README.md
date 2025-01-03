# Label Maker Application

## Overview
The Label Maker Application is a Python-based tool designed to generate labels with customizable colors, QR codes, and order details. It allows users to load files, categorize them as envelopes or letters, assign colors to orders, and generate a PDF containing the labels.

---

## Features
- **File Selection:** Choose envelope and letter chip files for processing.
- **Label Management:** Automatically assigns order details based on filenames.
- **Color Customization:** Assign and update colors for each order.
- **QR Code Integration:** Add and manage QR codes linked to URLs.
- **PDF Generation:** Generates a multi-page PDF with 2x6 label layouts for printing.
- **Order History:** Displays the last 20 order details with color assignments.
- **Responsive UI:** Supports high-DPI scaling for better visibility.

---

## Requirements
- Python 3.8 or higher
- Libraries:
  - customtkinter
  - tkinter
  - PIL (Pillow)
  - qrcode
  - logging

Install dependencies using:
```bash
pip install customtkinter Pillow qrcode
```

---

## Installation
1. Clone the repository or download the source code.
2. Install the required Python packages as listed above.
3. Ensure the following files and folders are available:
   - `resources/`: Contains icons and fonts.
   - Font file `Arial_Bold.ttf` in the `resources` folder.
   - Icon file `scribe-icon.ico` in the `resources` folder.
   - Logo file `scribe-logo-final.webp` in the `resources` folder.

---

## Usage
1. Launch the application by running:
```bash
python label_maker.py
```
2. Select files using the "Select Envelope Chip Files" or "Select Letter Chip Files" buttons.
3. Assign colors to orders and optionally add QR codes.
4. Generate a PDF with labels by clicking the "Create PDF" button.
5. View or open the created PDF file directly from the application.
6. Reset data as needed using the "Reset" button.

---

## File Naming Convention
The application expects file names to include keywords like "Envelopes" or "Letters" to distinguish types. Examples:
```
Chris LaVigne T1 Copy 1 Envelopes-1-167.bin
Chris LaVigne T1 Copy 2 Letters-1-167.bin
```
If a file doesn't match this convention, the user will be prompted to categorize it.

---

## Output
- PDF labels are saved to:
```
G:\Shared drives\Scribe Workspace\Scribe Master Folder\Batch Labels
```
- Order history is stored in:
```
G:\Shared drives\Scribe Workspace\Scribe Master Folder\Scribe Label Maker\order_history.json
```

---

## Notes
- Ensure fonts and icons are placed correctly in the `resources` folder.
- DPI scaling is supported for better visibility on high-resolution displays.
- Logs are saved in `file_processing_errors.log` for debugging.

---

## License
This project is licensed under the MIT License. See the LICENSE file for details.

---

## Troubleshooting
- **PDF Not Created:** Ensure required fonts are available and check error logs.
- **Color Picker Doesn't Work:** Update tkinter and customtkinter.
- **QR Codes Missing:** Ensure URLs are valid and saved before generating the PDF.

---

## Acknowledgments
Special thanks to the open-source community for libraries used in this project.