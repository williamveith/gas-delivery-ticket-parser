# Gas Delivery Ticket Parser

Extracts cylinder delivery information from **Linde Gas delivery ticket PDFs** and formats it for direct pasting into **Excel or Google Sheets**.

The script reads all PDF files in a directory, parses delivery ticket data, and copies a tab-separated table to the clipboard.

---

## Features

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
* Copies results directly to clipboard for **Ctrl+V into Excel / Google Sheets**
* Handles inconsistent PDF spacing using layout extraction

---

## Output Format

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

## Project Structure

```
delivery-ticket-parser/
│
├── main.py
├── lookups.py
├── README.md
│
└── Delivery Documents/
    ├── Order_XXXX.pdf
    ├── Order_XXXX.pdf
```

### `main.py`

Main parser and extraction script.

### `lookups.py`

Contains lookup dictionaries:

* `account_to_group`
* `part_number_to_gas`

These map internal account numbers and cylinder part numbers to readable values.

---

## Requirements

Python 3.10+

Install dependencies:

```bash
pip install pypdf pyperclip
```

---

## Usage

1. Place delivery ticket PDFs in:

```
Delivery Documents/
```

2. Run the script:

```bash
python main.py
```

3. The script will:

* Parse all PDFs
* Copy the output table to your clipboard
* Print the list of processed files

4. Paste directly into **Excel or Google Sheets**.

---

## Data Model

Two dataclasses represent the parsed data.

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

## Parsing Logic

The script:

1. Extracts text from each PDF page using **layout-based extraction**
2. Normalizes whitespace
3. Uses regex to locate:

| Field           | Pattern                           |
| --------------- | --------------------------------- |
| Customer Number | 8 digits                          |
| Order Number    | 8 digits + 1–2 letters + 8 digits |
| Customer PO     | YYYY + A/C/U + 5 digits           |
| Delivery Date   | `MM/DD/YY HH:MM`                  |

4. Detects cylinder items using the pattern:

```
PART_NUMBER CO DESCRIPTION
```

5. Associates **SHIP serial numbers** with the current part number.

---

## Example Workflow

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

## Notes

* Only **SHIPPED cylinders** are extracted.
* Returned cylinders are ignored.
* If a part number is missing from `part_number_to_gas`, the gas field will be blank.

---

## Author

William Veith
