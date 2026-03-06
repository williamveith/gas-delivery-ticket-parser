import re
import pathlib
from datetime import datetime
from pypdf import PdfReader
import pyperclip
from dataclasses import dataclass, field
from typing import List
from lookups import account_to_group, part_number_to_gas

@dataclass
class Cylinder:
    serial_number: str
    part_number: str
    gas: str

@dataclass
class DeliveryTicket:
    customer_number: str = ""
    group: str = ""
    order_number: str = ""
    customer_po: str = ""
    delivery_date: str = ""
    shipped_cylinders: List[Cylinder] = field(default_factory=list)


def extract_delivery_ticket_data(pdf_path: str | pathlib.Path) -> DeliveryTicket:
    reader = PdfReader(str(pdf_path))

    text = "\n".join(
        page.extract_text(extraction_mode="layout", layout_mode_space_vertically=False)
        or ""
        for page in reader.pages
    )

    text = text.replace("\xa0", " ")
    lines = [line.rstrip() for line in text.splitlines() if line.strip()]
    full_text = "\n".join(lines)

    data = DeliveryTicket()

    # CUSTOMER# -> always 8 digits
    m = re.search(
        r"\bCUSTOMER\s*#?\s*[:\-]?\s*(\d{8})\b", full_text, flags=re.IGNORECASE
    )
    if m:
        data.customer_number = m.group(1)
        data.group = account_to_group.get(m.group(1), "")

    # ORDER# -> 8 digits + 1-2 letters + 8 digits
    m = re.search(
        r"\bORDER\s*#?\s*[:\-]?\s*(\d{8}[A-Z]{1,2}\d{8})\b",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        data.order_number = m.group(1)[:8]

    # CUSTOMER PO -> 4-digit year + A/C/U + 5 digits
    m = re.search(
        r"\bCUSTOMER\s*P\.?\s*O\.?\s*#?\s*[:\-]?\s*((?:20\d{2})[ACU]\d{5})\b",
        full_text,
        flags=re.IGNORECASE,
    )
    if m:
        data.customer_po = m.group(1)

    # delivery date from header line like: 03/28/24 09:00
    m = re.search(r"\b(\d{2}/\d{2}/\d{2})\s+\d{2}:\d{2}\b", full_text)
    if m:
        data.delivery_date = (
            datetime.strptime(m.group(1), "%m/%d/%y").date().isoformat()
        )

    # Product line pattern:
    #   AR 5.0UH-T-EZ               CO     ARGON ...
    #   HE 5.0UH-T                  CO     HELIUM ...
    #   OX 5.0RS-T                  CO     OXYGEN ...
    #   SH 4.5SP-AS                 CO     SULFUR HEXAFLUORIDE ...
    #   NI HY4C-K                   CO     NIT-HYD ...
    #
    # Capture everything before the UOM "CO"
    item_line_re = re.compile(r"^\s*([A-Z]{1,4}\s+[A-Z0-9.\-]+)\s+CO\b")

    # Serial lines marked SHIP
    ship_sn_re = re.compile(r"^\s*SN\s*:\s*(\d+)\b.*\bSHIP\b", flags=re.IGNORECASE)

    # Lines that clearly end an item block
    block_end_re = re.compile(
        r"^\s*(DEFAULT\s+VOLUME:|BACKORDER\b|TOTAL\s+CYLINDERS\s+SHIPPED:)",
        flags=re.IGNORECASE,
    )

    current_part_number = None

    for line in lines:
        item_match = item_line_re.match(line)
        if item_match:
            current_part_number = item_match.group(1).strip()
            continue

        if block_end_re.match(line):
            # leave current_part_number alone until next item,
            # but no SN lines should matter past this section
            continue

        sn_match = ship_sn_re.match(line)
        if sn_match and current_part_number:
            data.shipped_cylinders.append(
                Cylinder(
                    serial_number=sn_match.group(1),
                    part_number=current_part_number,
                    gas=part_number_to_gas.get(current_part_number, ""),
                )
            )

    return data


directory = "/Users/main/Projects/pdf reader/test files"
path = pathlib.Path(directory)

filenames = ["Copied data from the following files. Paste in Excel (Ctrl+V)."]
rows = []
for pdf_file_path in path.glob("*.pdf"):
    data = extract_delivery_ticket_data(pdf_file_path)
    filenames.append(pdf_file_path.stem)
    for cylinder in data.shipped_cylinders:
        rows.append(
            [
                cylinder.serial_number,
                cylinder.gas,
                cylinder.part_number,
                data.customer_number,
                data.group,
                data.customer_po,
                data.order_number,
                data.delivery_date,
            ]
        )

text_to_copy = "\n".join(["\t".join(row) for row in rows])
print_confirmation = "\n\t".join(filenames)

pyperclip.copy(text_to_copy)
print(print_confirmation)