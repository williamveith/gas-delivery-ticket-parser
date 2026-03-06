# Gas Delivery Ticket Parser

Extracts cylinder delivery information from **Linde Gas delivery ticket PDFs** and formats it for direct pasting into **Excel or Google Sheets**.

The script reads all PDF files in a directory, parses delivery ticket data, and copies a tab-separated table to the clipboard.

---

# Features

* Parses **multi-page delivery tickets**
* Extracts:

  * Serial Number
  * Gas Type
  * Part Number
  * Customer Account
  * Research Group
  * PO Number
  * Order Number
  * Delivery Date
* Automatically maps:

  * **Account → Research Group**
  * **Part Number → Gas**
* Copies results directly to the clipboard for **Ctrl+V into Excel / Google Sheets**
* Handles inconsistent PDF spacing using **layout-based PDF extraction**
* Supports **custom input directories via CLI**

---

# Output Format

The clipboard contains **tab-separated rows** in the format:

```
Serial No.    Gas    Part Number    Account    Group    PO No.    Order No.    Received
```

Example:

```
200005443208    Argon    AR 5.0UH-T-EZ    71659397    Akinwande    2026U18493    85964907    2026-02-03
200006712706    Argon    AR 5.0UH-T-EZ    71659397    Akinwande    2026U18493    85964907    2026-02-03
```

After running the script, simply **paste into Excel or Google Sheets**.

---

# Requirements

Python version required:

```
Python 3.14.x
```

The script enforces this at runtime.

Dependencies:

```
pypdf
pyperclip
```

Install using:

```bash
pip install -r requirements.txt
```

---

# Environment Setup

Create a virtual environment before running the script.

```bash
python3.14 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Dependencies listed in:

```
requirements.txt
```

---

# Usage

### Default directory

Place delivery ticket PDFs inside:

```
Delivery Documents/
```

Run:

```bash
python main.py
```

---

### Custom directory

You can provide a directory containing PDFs:

```bash
python main.py "test files"
```

or

```bash
python main.py ~/Downloads/tickets
```

---

# Project Structure

```
gas-delivery-parser/
│
├── main.py
├── lookups.py
├── requirements.txt
├── README.md
│
└── Delivery Documents/
    ├── Order_XXXX.pdf
    ├── Order_XXXX.pdf
```

---

# Lookup Tables

The file `lookups.py` contains dictionaries used during parsing:

* `account_to_group`
* `part_number_to_gas`

These convert internal identifiers into readable values. 

---

# Data Model

Two dataclasses represent the parsed ticket structure.

### Cylinder

```
Cylinder
├─ serial_number
├─ part_number
└─ gas
```

### DeliveryTicket

```
DeliveryTicket
├─ customer_number
├─ group
├─ order_number
├─ customer_po
├─ delivery_date
└─ shipped_cylinders
```

---

# Parsing Logic

The script performs the following steps:

1. Extracts text from each PDF page using **layout-based extraction**
2. Normalizes whitespace
3. Uses regex patterns to extract key fields:

| Field           | Pattern                           |
| --------------- | --------------------------------- |
| Customer Number | 8 digits                          |
| Order Number    | 8 digits + 1–2 letters + 8 digits |
| Customer PO     | YYYY + A/C/U + 5 digits           |
| Delivery Date   | `MM/DD/YY HH:MM`                  |

4. Detects cylinder items using:

```
PART_NUMBER CO DESCRIPTION
```

5. Associates **SHIP serial numbers** with the current part number.

---

# Example Workflow

```
Linde Delivery Ticket PDF
        ↓
Parse Ticket
        ↓
Extract Cylinders
        ↓
Map Account → Group
Map Part Number → Gas
        ↓
Generate Table
        ↓
Copy to Clipboard
        ↓
Paste into Spreadsheet
```

---

# Notes

* Only **SHIPPED cylinders** are extracted.
* Returned cylinders are ignored.
* If a part number is missing from `part_number_to_gas`, the gas field will be blank.
* Files are processed in **sorted order** for consistent output.

---

# Author

William Veith

