# Simple Scanner App

A simple Windows application to scan documents using WIA (Windows Image Acquisition).
Built with Python and Tkinter.

## Download & Run
1. Download `SimpleScanner.exe` from the `dist` folder in this repository.
2. Run the executable. No installation required.

## Features
- Auto-detects default WIA scanner.
- Scans documents with a single click.
- Automatically saves scanned images with timestamps.

## Development

### Requirements
- Python 3.10+
- `pywin32`
- `pillow`
- `pyinstaller` (for building)

### Build from Source
```bash
pip install -r requirements.txt
pyinstaller --noconsole --onefile --name "SimpleScanner" --hidden-import=win32com.client scanner_app.py
```
