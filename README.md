# Battery Inventory

A lightweight web app for logging and managing battery inventory at Titan AES. Run it locally — it opens automatically in your browser.

## Features

- Add batteries by scanning barcodes (camera or USB scanner) or typing manually
- Track Titan ID, Manufacturer ID, OCV, weight, flag status (Pass / Suspect / Fail), and comments
- Auto-detects out-of-spec cells (>2σ from batch average) and highlights them
- Export to formatted Excel (.xlsx) or CSV
- Import from a previously exported Excel file
- Inline editing and undo-delete
- Batch name header and print layout
- Data persists in browser localStorage between sessions

## Requirements

```
python 3.8+
flask
openpyxl
```

Install dependencies:

```bash
pip install flask openpyxl
```

## Usage

```bash
python battery_inventory.py
```

The app opens automatically at http://localhost:5555.

Optionally place `Titanaes.png` in the same folder as the script to show the logo in the header.

## Excel Export Format

| Column | Description |
|--------|-------------|
| Titan ID | Internal cell number |
| Manufacturer ID | Barcode / serial from manufacturer |
| OCV (V) | Open-circuit voltage |
| Weight (g) | Cell weight |
| Flag | Pass / Suspect / Fail |
| Comments | Optional notes |
| Date Added | Timestamp of entry |
