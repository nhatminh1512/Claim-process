---
name: overseas-travel-claim
description: >-
  Process Bossard Vietnam overseas travel expense claims from receipt photos,
  Vietnamese bank/card VND deduction screenshots (including dual-currency card
  alerts like Sacombankpay), and optional bank statement PDFs into a confirmed
  review table, a zip-patched Excel claim workbook (preserving layout, logo,
  merges, account codes, and formulas), and a date-grouped Word proof pack.
  Use whenever a user asks to prepare, extract, match, review, or export an
  overseas travel claim (`Template (oversea travel exp)`), including split
  Grab fare + toll deductions.
---

# Overseas Travel Expense Claim

End-to-end procedure for filing a Bossard Vietnam overseas travel expense claim without retyping receipt amounts or breaking the company Excel template.

---

## 1. Non-Negotiable Finance-Safety Rules

1. **Template fidelity (never re-save with `openpyxl`)**:
   - Inspect the uploaded `.xlsx` with `openpyxl.load_workbook(..., data_only=False)` read-only, then patch **only** the target worksheet XML and `xl/workbook.xml` (`fullCalcOnLoad="1"`) inside a copy of the `.xlsx` zip archive.
   - Every other zip member (`xl/media/*` logo, `xl/drawings/*`, merges, styles, number formats, GL account-code row, `Perdiem`, `Advances`, and other sheets) passes through byte-for-byte.
2. **Dynamic landmark discovery (never hardcode cell addresses)**:
   - Find the sheet whose name contains `"oversea"` (or whose top cells contain `"OVERSEA TRAVEL EXPENSE REPORT"`).
   - Locate `Description` header row/col, category header groups to its right (joining split hyphenated headers like `Accom-` + `modation`, `Entertain-` + `ment`), and split-label leaves (`Meals` → `Lunch`, `Transport` → `Oversea` / `Overseas`).
   - Foreign block: starts below `"Expenses in foreign currency"` and ends at `SUM(` row - 1. Only unhide rows that are actually written; preserve the template's hidden state on unused rows.
   - VND block: starts below `"Expenses in ..."` (non-foreign, e.g., `"Expenses in VND"` or legacy `"Expenses in MYR"`) and ends before `"Grand Total"`, excluding any `"Perdiem"` row.
   - Summary landmarks: `"Sub Total in ..."`, `"Total in VND"`, `"Exchange Rate"` row, `"Grand Total"` column, and optional header labels (`Name:`, `Period Covered From :`, `To :`, `Purpose of Trip:`, `Location:`, `Cost Cctr :`, `Currency Type:`, `Less: Advances...`).
3. **Locked expense columns (exactly 7)**:
   `Air Fare`, `Accommodation`, `Meals`, `Phone`, `Transport`, `Entertainment`, `Others`.
   Never invent a category name outside this list.
4. **Confirmation & export gate**:
   Block export until:
   - An `.xlsx` template is attached and every grid cell to be cleared/written contains no formula.
   - Every receipt has a finite amount (`0 <= amount < 1e12`), non-empty `purpose` (`<= 32,767` chars, no XML-illegal C0 control chars `[\x00-\x08\x0b\x0c\x0e-\x1f]`), and `column_confirmed == true`.
   - Every foreign receipt with a VND proof has a valid `vnd_amount` (`0 <= vnd_amount < 1e12`).
   - Zero unresolved pending VND proofs remain.
   - Foreign and VND slot counts do not exceed the template's row capacity.
5. **No double counting / no mixed columns**:
   - Within each leaf column, foreign rows must be **either all paired** (carrying `vnd_amount`) **or all unpaired**. Reject a column that mixes paired and unpaired foreign rows.
   - Fully paired columns rewrite `Total in VND` to `SUM(<col><vnd_start>:<col><total_row-1>)` (expanding any shared formula family first so sibling cells stay valid). This includes `Perdiem` while excluding the converted foreign subtotal.
   - Unpaired-only columns keep the template's original total formula.
6. **Weighted Exchange Rate**:
   - Over foreign receipts that have a valid `vnd_amount` (excluding pure VND rows and `Perdiem`), compute `round(sum(vnd_amount) / sum(foreign_amount), 3)` and write it across every expense grid column on the `Exchange Rate` row.

---

## 2. Document Classification & Extraction Rubric

When reading uploaded images (via vision model and/or OCR), extract structured metadata for each image:

| Field | Rule |
|---|---|
| `document_kind` | `foreign_invoice`, `vnd_deduction`, or `unclear` |
| `amount` | Printed **GRAND TOTAL** for invoices (including tolls/fees/tax; never subtotal or line item; never convert currency). Absolute value of **deducted/provisional VND** for bank/card proofs (e.g. `-192,510 VND` or `(tạm tính 533,997VND)` → `533997`; **never** account balance or credit limit such as `Số dư khả dụng` / `Hạn mức khả dụng`). `null` + `unreadable=true` if illegible. |
| `currency` | Normalize printed aliases (`RM` → `MYR`, `US$` → `USD`, `S$` → `SGD`, `RMB` → `CNY`, etc.) for bills; `VND` for bank/card proofs. |
| `currency_kind` | `"VND"` if the expense itself is VND (or a standalone `vnd_deduction` row); otherwise `"FOREIGN"`. |
| `foreign_amount` | *(Optional, on dual-currency bank/card alerts)* Foreign transaction amount printed in the alert (e.g., `83.43` from `giao dịch 83.43 MYR (tạm tính 533,997VND)`). |
| `foreign_currency` | *(Optional, on dual-currency bank/card alerts)* Normalized foreign currency printed in the alert (e.g., `MYR`). |
| `vendor` | Merchant-first: `"Grab"` if Grab/`Chuyến đi`/`Chi tiết giao dịch`/GrabMaps/ride summary or bank `Nội dung` mentions `GRAB`/`GRAB RIDES`; otherwise merchant printed on bill; fall back to bank name (`TPBank`, `Sacombank`, `Vietcombank`, etc.) only if no merchant appears. |
| `date` | Transaction date in `YYYY-MM-DD` (supports `27 Jul 2026`, `22 thg 11 2026`, `22/11/2026`, ISO). Never use screenshot status-bar date. |
| `time` | 24-hour `HH:MM` transaction time if printed (convert `SA`/`AM` and `CH`/`PM`, e.g. `4:53 CH` → `16:53`, `7:50 SA` → `07:50`). **Always ignore line 1 phone status-bar clock** (e.g., `09:01`, `17:30`). Treat `00:00` (TPBank settled placeholder) as `null`. |
| `card_last4` | 4-digit card suffix if visible (e.g. `0175` or `2071` from `Phương thức 0175`, `thẻ Amex Credit 2071`, `GIAO DICH THE 0175`, or icon `@ 2071`), else `null`. |
| `column` | Initial unconfirmed guess from the 7 locked columns (`Grab`/`Uber`/`Bolt`/`taxi`/tolls → `Transport`; hotels → `Accommodation`; food → `Meals`; else `null` if unclear). |
| `purpose_suggestion` | *(Optional draft for review)* Concise trip segment when pickup/dropoff landmarks are clearly visible on a ride screen (e.g. `Travel from <pickup> to <dropoff>`). Never finalize without user confirmation. |

### Classification Signals (in priority order)

1. **Bank / card alert markers win over embedded foreign currencies** (`vnd_deduction`):
   Real card alerts (e.g. Sacombankpay, TPBank, Vietcombank, Techcombank) often print **both** the foreign amount and the deducted VND (`giao dịch 83.43 MYR (tạm tính 533,997VND)`, `Hạn mức khả dụng: ...`). Whenever VND deduction / card alert labels appear alongside a VND charge amount, classify as `vnd_deduction` and extract the **VND charge** as `amount` (plus optional `foreign_amount` / `foreign_currency`).
   - Bank/card labels: `Số tiền giao dịch`, `Số dư sau giao dịch`, `Số dư khả dụng`, `Hạn mức khả dụng`, `Từ tài khoản`, `Chuyển tiền trong TPB`, `Giao dịch thẻ`, `Thanh toán thẻ`, `Chi tiêu thẻ`, `Số tiền bị trừ`, `Biên lai giao dịch`, `Thông báo giao dịch`, `tạm tính ... VND`, or VN bank names (`TPBank`, `Sacombank`, `Sacombankpay`, `Vietcombank`, `Techcombank`, `VPBank`, `MBBank`, `VietinBank`, `BIDV`, `Shinhan`).
2. **Explicit non-VND currency without bank/card deduction labels** → `foreign_invoice`:
   Includes both:
   - Ride/merchant detail screens (`Chi tiết giao dịch`, `Chuyến đi`, `Bạn đã thanh toán`, `Phương thức thanh toán`, `Mã đặt dịch vụ`, `GrabCar`, `GrabTaxi`, `JustGrab`, `Booking ID`, `Ride details`, `You paid`, `Trip fare`).
   - Ride-hailing map/summary screens (`Total RM...`, pickup/dropoff addresses with timestamps, `Standard | 6 seats`, `Xuất hóa đơn`, `Báo cáo sự cố`, `Đặt lại`, `Cá nhân`).
3. **Fractional amount under 1,000** tagged as `vnd_deduction` without explicit `VND` → reclassify as `foreign_invoice`.

---

## 3. Pairing & Split-Deduction Matching (Grab Fare + Toll)

Grab rides and similar expenses often charge the base fare at booking and a separate toll/adjustment fee later on arrival (e.g. `89.34 MYR` split across `83.43 MYR (533,997 VND)` + `5.91 MYR (47,170 VND)`, or `22.66 MYR` split across `14.42 MYR (100,602 VND)` + `8.24 MYR (61,883 VND)`).

### Plausible VND-per-Unit FX Rate Bands

| Currency | Plausible VND / unit | Currency | Plausible VND / unit |
|---|---|---|---|
| `MYR` / `RM` | `4,500 – 7,500` | `AUD` | `14,000 – 20,000` |
| `USD` | `22,000 – 30,000` | `GBP` | `29,000 – 40,000` |
| `EUR` | `24,000 – 33,000` | `HKD` | `2,800 – 3,900` |
| `SGD` | `16,000 – 22,000` | `TWD` | `700 – 1,000` |
| `THB` | `600 – 900` | `KRW` | `15 – 23` |
| `JPY` | `140 – 220` | `IDR` | `1.4 – 2.2` |
| `CNY` / `RMB` | `3,000 – 4,200` | Other | `100 – 40,000` |

### Matching Rules (1 to 3 Proofs per Foreign Bill)

1. **Priority 0 — Exact Foreign-Amount Match (when bank/card alerts print `foreign_amount`)**:
   - If 1–3 `vnd_deduction` proofs share the bill's normalized currency (`foreign_currency`), pass date/time/card/vendor checks, and `abs(sum(proof.foreign_amount) - bill.amount) <= 0.02`, treat this as a high-confidence match even if card FX spread/fees push the small toll's effective VND/foreign rate above the standard band (e.g. `47,170 / 5.91 ≈ 7,981` while `581,167 / 89.34 ≈ 6,505` is cleanly in band).
2. **Trip anchor rate estimation**:
   - Collect `implied_rate = vnd_amount / foreign_amount` from unambiguous 1:1 pairs (skipping any 1:1 pair where another date-compatible proof could combine with it inside the currency band). Use the median per currency as `reference_rate`.
3. **Hard vetoes** (never override with AI confidence):
   - Parsed dates `> 1` day apart.
   - Both items have `card_last4` and they differ.
   - Both items have valid non-midnight `time` (`HH:MM`) and are `> 180` minutes apart (proof ↔ bill, which accommodates timezone differences such as UTC+8 destination vs UTC+7 VN bank alert plus ride duration) or `> 150` minutes apart (proof ↔ proof in the same split group).
   - Conflicting non-bank merchants between proofs (unless `Transport` + toll keyword such as `toll`, `cau duong`, `etc`, `touch n go`, `tng`, `erp`).
   - Combined `sum(vnd_amounts) / foreign_amount` falls outside the currency's FX band (unless Priority 0 exact `foreign_amount` sum matched and total rate is plausible within `±25%` of the band).
4. **Agreement requirement**:
   - At least one of same/adjacent date (`<= 1` day), matching normalized vendor, or matching `card_last4` must hold for every proof in the group.
5. **Anti-false-merge guard** (prevents absorbing an unrelated stray debit onto a full-paid bill):
   - **When `reference_rate` exists**: combined group rate must be within `±15%` of `reference_rate`, and strictly closer to `reference_rate` than any single proof's rate (unless Priority 0 exact `foreign_amount` sum matched).
   - **When `reference_rate` is absent**: if no single proof is in-band on its own (e.g. 50/50 split), allow the merge; if one proof is already in-band (`single_rate`), allow merging only when `single_rate <= center + 0.08 * span` (understated fare), `group_rate <= low + 0.90 * span`, and `|group_rate - center| <= |single_rate - center| + 0.12 * span`.
6. **Primary vs extra proofs**:
   - The **largest** VND deduction in a matched group becomes `vnd_proof_path` (`vnd_primary_amount`); smaller ones (tolls/fees) go into `extra_proofs` with `amount_included = true` and `vnd_amount` summed into the row's `vnd_amount`.
   - When a user manually adds a 2nd/extra VND screenshot later, **extract its VND amount** and sum it into `vnd_amount` (never leave it ignored); subtract cleanly if removed.
   - Unmatched VND proofs when foreign bills exist go to `pending_proofs` (never silently dropped). If the claim has zero foreign invoices, VND proofs become standalone `VND` rows.

---

## 4. Optional Bank Statement PDF (`sao kê`) Reconciliation

When the user supplies a bank statement PDF instead of individual bank screenshots:

1. **Extract outgoing debits only**: Native PDF text first; scanned pages (`< 30` chars) fall back to local Tesseract OCR (`eng --psm 6`). Ignore credits, opening/closing balances, and column totals. Normalize `amount = abs(vnd)`, `date = YYYY-MM-DD`, `time = HH:MM` (null if `00:00`), `page`.
2. **Do not drop small debits prematurely**: Keep any debit that either matches a bill's 1:1 rate band **or** shares date (±1 day) and merchant/toll keywords with a bill and is smaller than the bill's max plausible VND.
3. **Strict matching**: Require rate band + date (±1 day) + time window + merchant token overlap (for secondary split debits on `Transport` rows, also accept toll keywords `toll`, `cau duong`, `etc`, `touch n go`, `tng`, `erp`).
4. **Red-circled crop**: Use the `crop-bank` subcommand of the embedded script below to crop each matched payment row (description + amount, excluding balance) and draw a red ellipse around the VND amount. Attach the largest debit crop as `vnd_proof_path` (`vnd_proof_source = "bank_statement"`) and split toll crops as `extra_proofs` (`source = "bank_statement"`, `amount_included = true`, caption `"Bank statement · DD/MM/YYYY · DESC · AMT VND"`).

---

## 5. Interactive Review & Confirmation Protocol

Before exporting files, present a concise Markdown review table to the user:

- Columns: `# | Date | Vendor | Foreign | VND Deducted (Breakdown) | Column (Suggested / Confirmed) | Purpose (Draft / Confirmed) | Proofs`
- Explicitly list:
  - Any unreadable amounts or `unclear` classifications.
  - Any possible duplicates (same amount + date/vendor).
  - Any unpaired `pending_proofs` (suggest matching foreign rows by date/vendor/rate/foreign_amount).
  - Any column mixing paired and unpaired foreign rows (must be resolved before export).
- Ask the user to confirm or edit `purpose`, `column`, optional trip header fields (if filling header), and any missing amounts. Only run the export tool once all gates in Section 1 pass.

---

## 6. Self-Contained Export & Bank-Crop Script

Save or run this standalone Python script (requires `openpyxl`, `lxml`, `defusedxml`, `python-docx`, `pillow`, `pymupdf`) to inspect the template, fill the workbook, build the Word proof pack, and optionally crop bank statement rows:

```bash
# Export filled Excel, Word proof pack, and combined ZIP:
uv run --with openpyxl --with lxml --with defusedxml --with python-docx --with pillow --with pymupdf \
  python claim_tool.py export --claim claim.json --out-dir ./output

# Crop a bank statement PDF row with a red circle around the VND amount:
uv run --with openpyxl --with lxml --with defusedxml --with python-docx --with pillow --with pymupdf \
  python claim_tool.py crop-bank --pdf statement.pdf --amount 192510 --date 2026-07-31 --description "GRAB RIDES" --page 1 --out crop.png
```

```python
#!/usr/bin/env python3
"""Self-contained Overseas Travel Claim exporter & bank-statement cropper."""

from __future__ import annotations

import argparse
import json
import math
import os
import posixpath
import re
import shutil
import unicodedata
import zipfile
from dataclasses import dataclass, field
from datetime import date, datetime
from io import BytesIO
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree.ElementTree import ParseError

import pymupdf
from defusedxml.ElementTree import fromstring as safe_fromstring
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.shared import Cm, Inches, Pt
from lxml import etree
from openpyxl import load_workbook
from openpyxl.formula.translate import Translator
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import range_boundaries
from openpyxl.utils.exceptions import InvalidFileException
from PIL import Image, ImageDraw, ImageOps

EXPENSE_COLUMNS = (
    "Air Fare",
    "Accommodation",
    "Meals",
    "Phone",
    "Transport",
    "Entertainment",
    "Others",
)
CURRENCY_ALIASES = {
    "RM": "MYR",
    "RINGGIT": "MYR",
    "US$": "USD",
    "$": "USD",
    "€": "EUR",
    "S$": "SGD",
    "A$": "AUD",
    "£": "GBP",
    "HK$": "HKD",
    "NT$": "TWD",
    "₩": "KRW",
    "RP": "IDR",
    "RMB": "CNY",
    "BAHT": "THB",
    "YEN": "JPY",
}
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
XML_ILLEGAL = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f]")


class ClaimToolError(ValueError):
    """Raised when validation, mapping, fill, or proof-pack generation fails."""


@dataclass(frozen=True)
class _ColumnGroup:
    start: int
    end: int
    header: str
    children: tuple[tuple[int, str], ...]


@dataclass(frozen=True)
class SheetMapping:
    sheet_name: str
    description_col: str
    leaf_columns: dict[str, str]
    amount_columns: list[str]
    data_row_start: int
    data_row_end: int
    sum_row: int
    vnd_data_row_start: int | None
    vnd_data_row_end: int | None
    total_row: int | None
    subtotal_row: int | None
    exchange_rate_row: int | None
    header_cells: dict[str, str] = field(default_factory=dict)


@dataclass
class ExtraProof:
    stored_path: str
    caption: str | None = None
    vnd_amount: float | None = None
    amount_included: bool = False
    source: str | None = None


@dataclass
class Receipt:
    stored_path: str
    purpose: str
    column: str
    amount: float
    currency_kind: str = "FOREIGN"
    currency: str | None = None
    date: str | None = None
    column_confirmed: bool = True
    vnd_amount: float | None = None
    vnd_primary_amount: float | None = None
    vnd_proof_path: str | None = None
    vnd_proof_source: str | None = None
    extra_proofs: list[ExtraProof] = field(default_factory=list)


def _q(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


def _safe_xml_parser() -> etree.XMLParser:
    return etree.XMLParser(
        resolve_entities=False,
        load_dtd=False,
        no_network=True,
        huge_tree=False,
        recover=False,
    )


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize_currency(value: str | None) -> str:
    raw = (value or "").strip().upper().replace(" ", "")
    return CURRENCY_ALIASES.get(raw, raw)


def _merged_end(ws: Any, row: int, col: int) -> int:
    end = col
    for merged in ws.merged_cells.ranges:
        if merged.min_row <= row <= merged.max_row and merged.min_col == col:
            end = max(end, merged.max_col)
    return end


def _category_for(header: str) -> str | None:
    low = header.lower().strip()
    if low.startswith("air fare") or low.startswith("airfare"):
        return "Air Fare"
    if low.startswith("accom"):
        return "Accommodation"
    if low.startswith("meal"):
        return "Meals"
    if low.startswith("phone"):
        return "Phone"
    if low.startswith("transport"):
        return "Transport"
    if low.startswith("entertain"):
        return "Entertainment"
    if low.startswith("other"):
        return "Others"
    return None


def _leaf_column(group: _ColumnGroup, category: str) -> int:
    if category == "Meals":
        for col, label in group.children:
            if label.lower() == "lunch":
                return col
        if group.children:
            raise ClaimToolError("Meals group has no Lunch split label")
        return group.start
    if category == "Transport":
        for col, label in group.children:
            if label.lower() in {"oversea", "overseas"}:
                return col
        if group.children:
            raise ClaimToolError("Transport group has no Oversea split label")
        return group.start
    return group.start


def _discover_header_cells(ws: Any, header_row: int, grand_total_col: int | None) -> dict[str, str]:
    """Locate optional header value cells dynamically from label text."""

    label_prefixes = {
        "name": ("name:",),
        "period_from": ("period covered from",),
        "period_to": ("to :", "to:"),
        "purpose_of_trip": ("purpose of trip",),
        "location": ("location:",),
        "cost_center": ("cost cctr", "cost center"),
        "currency_type": ("currency type",),
    }
    found: dict[str, str] = {}
    max_col = min(ws.max_column, 20)
    for r in range(1, min(ws.max_row, header_row + 5) + 1):
        for c in range(1, max_col + 1):
            lbl = _text(ws.cell(row=r, column=c).value).lower()
            if not lbl:
                continue
            for key, prefixes in label_prefixes.items():
                if key not in found and any(lbl.startswith(p) for p in prefixes):
                    target_col = _merged_end(ws, r, c) + 1
                    found[key] = f"{get_column_letter(target_col)}{r}"
    for r in range(header_row + 1, ws.max_row + 1):
        for c in range(1, max_col + 1):
            lbl = _text(ws.cell(row=r, column=c).value).lower()
            if lbl.startswith("less: advances") or lbl.startswith("advances"):
                target_col = None
                for vc in range(_merged_end(ws, r, c) + 1, ws.max_column + 1):
                    v = ws.cell(row=r, column=vc).value
                    if isinstance(v, (int, float)) and not isinstance(v, bool):
                        target_col = vc
                        break
                if target_col is None:
                    target_col = grand_total_col or (_merged_end(ws, r, c) + 1)
                found["advances"] = f"{get_column_letter(target_col)}{r}"
                break
        if "advances" in found:
            break
    return found


def inspect_overseas_mapping(xlsx_path: str | Path) -> SheetMapping:
    try:
        wb = load_workbook(xlsx_path, data_only=False)
    except (
        BadZipFile, InvalidFileException, KeyError, OSError, ParseError, TypeError, ValueError
    ) as exc:
        raise ClaimToolError("template is not a readable Excel workbook") from exc
    try:
        ws = next((wb[n] for n in wb.sheetnames if "oversea" in n.lower()), None)
        if ws is None:
            for candidate in wb.worksheets:
                for row in candidate.iter_rows(
                    min_row=1, max_row=min(candidate.max_row, 60), max_col=min(candidate.max_column, 20)
                ):
                    if any(
                        isinstance(c.value, str) and "OVERSEA TRAVEL EXPENSE REPORT" in c.value.upper()
                        for c in row
                    ):
                        ws = candidate
                        break
                if ws is not None:
                    break
        if ws is None:
            raise ClaimToolError("no overseas travel expense sheet found in workbook")

        header_row = desc_col = None
        for r in range(1, min(ws.max_row, 100) + 1):
            for c in range(1, min(ws.max_column, 20) + 1):
                if _text(ws.cell(row=r, column=c).value).lower() == "description":
                    header_row, desc_col = r, c
                    break
            if header_row is not None:
                break
        if header_row is None or desc_col is None:
            raise ClaimToolError("could not find Description header row")

        max_col = min(ws.max_column, 40)
        groups: list[_ColumnGroup] = []
        for col in range(desc_col + 1, max_col + 1):
            header = _text(ws.cell(row=header_row, column=col).value)
            if not header:
                continue
            end = _merged_end(ws, header_row, col)
            children = [
                (c, _text(ws.cell(row=header_row + 1, column=c).value))
                for c in range(col, end + 1)
                if _text(ws.cell(row=header_row + 1, column=c).value)
            ]
            if header.endswith("-") and len(children) == 1:
                header = header[:-1] + children[0][1]
                children = []
            groups.append(_ColumnGroup(col, end, header, tuple(children)))

        leaf_columns: dict[str, str] = {}
        for g in groups:
            cat = _category_for(g.header)
            if cat and cat not in leaf_columns:
                leaf_columns[cat] = get_column_letter(_leaf_column(g, cat))
        missing = [c for c in EXPENSE_COLUMNS if c not in leaf_columns]
        if missing:
            raise ClaimToolError("missing overseas column(s): " + ", ".join(missing))

        amount_cols = sorted(
            {c for g in groups if _category_for(g.header) is not None for c in range(g.start, g.end + 1)}
        )

        data_start = None
        for r in range(header_row + 1, ws.max_row + 1):
            if "expenses in foreign currency" in _text(ws.cell(row=r, column=1).value).lower():
                data_start = r + 1
                break
        if data_start is None:
            raise ClaimToolError("could not find foreign-currency expense section")

        sum_row = None
        for r in range(data_start + 1, ws.max_row + 1):
            if any(
                isinstance(ws.cell(row=r, column=c).value, str)
                and "SUM(" in ws.cell(row=r, column=c).value.lstrip("=").upper()
                for c in amount_cols
            ):
                sum_row = r
                break
        if sum_row is None or sum_row - 1 < data_start:
            raise ClaimToolError("could not find foreign-currency SUM row")

        vnd_start = vnd_end = grand_total_col = None
        label_row = next(
            (
                r
                for r in range(sum_row + 1, ws.max_row + 1)
                if (t := _text(ws.cell(row=r, column=1).value).lower()).startswith("expenses in")
                and "foreign" not in t
            ),
            None,
        )
        if label_row is not None:
            gt_row = None
            for r in range(label_row + 1, ws.max_row + 1):
                for c in range(1, ws.max_column + 1):
                    if _text(ws.cell(row=r, column=c).value).lower() == "grand total":
                        gt_row, grand_total_col = r, c
                        break
                if gt_row is not None:
                    break
            if gt_row is not None:
                s, e = label_row + 1, gt_row - 1
                for r in range(e, s - 1, -1):
                    if _text(ws.cell(row=r, column=1).value).lower().startswith("perdiem"):
                        e = r - 1
                        break
                if e >= s:
                    vnd_start, vnd_end = s, e

        total_row = subtotal_row = exchange_rate_row = None
        for r in range(1, ws.max_row + 1):
            lbl = _text(ws.cell(row=r, column=1).value).lower()
            desc_lbl = _text(ws.cell(row=r, column=desc_col).value).lower()
            if subtotal_row is None and lbl.startswith("sub total in"):
                subtotal_row = r
            elif total_row is None and lbl == "total in vnd":
                total_row = r
            if exchange_rate_row is None and desc_lbl.startswith("exchange rate"):
                exchange_rate_row = r

        header_cells = _discover_header_cells(ws, header_row, grand_total_col)
        return SheetMapping(
            sheet_name=ws.title,
            description_col=get_column_letter(desc_col),
            leaf_columns=leaf_columns,
            amount_columns=[get_column_letter(c) for c in amount_cols],
            data_row_start=data_start,
            data_row_end=sum_row - 1,
            sum_row=sum_row,
            vnd_data_row_start=vnd_start,
            vnd_data_row_end=vnd_end,
            total_row=total_row,
            subtotal_row=subtotal_row,
            exchange_rate_row=exchange_rate_row,
            header_cells=header_cells,
        )
    finally:
        wb.close()


def _sheet_xml_path(archive: zipfile.ZipFile, sheet_name: str) -> str:
    wb_root = safe_fromstring(archive.read("xl/workbook.xml"))
    rels_root = safe_fromstring(archive.read("xl/_rels/workbook.xml.rels"))
    sheet = next((s for s in wb_root.iter(_q(MAIN_NS, "sheet")) if s.get("name") == sheet_name), None)
    if sheet is None:
        raise ClaimToolError(f"worksheet {sheet_name!r} missing")
    rid = sheet.get(_q(REL_NS, "id"))
    target = next(
        (r.get("Target") for r in rels_root.iter(_q(PKG_REL_NS, "Relationship")) if r.get("Id") == rid),
        None,
    )
    if not target:
        raise ClaimToolError(f"worksheet {sheet_name!r} relationship missing")
    path = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join("xl", target))
    if ".." in PurePosixPath(path).parts or not path.startswith("xl/") or path not in archive.namelist():
        raise ClaimToolError(f"invalid worksheet path {path!r}")
    return path


def _find_row(root: etree._Element, row_num: int) -> etree._Element:
    sd = root.find(_q(MAIN_NS, "sheetData"))
    if sd is None:
        raise ClaimToolError("worksheet missing sheetData")
    row = next((r for r in sd.findall(_q(MAIN_NS, "row")) if r.get("r") == str(row_num)), None)
    if row is None:
        raise ClaimToolError(f"worksheet row {row_num} missing")
    return row


def _find_cell(row: etree._Element, ref: str) -> etree._Element | None:
    return next((c for c in row.findall(_q(MAIN_NS, "c")) if c.get("r") == ref), None)


def _get_cell(row: etree._Element, ref: str) -> etree._Element:
    existing = _find_cell(row, ref)
    if existing is not None:
        return existing
    m = re.match(r"([A-Za-z]+)", ref)
    if m is None:
        raise ClaimToolError(f"invalid cell ref {ref!r}")
    wanted = column_index_from_string(m.group(1).upper())
    cell = etree.Element(_q(MAIN_NS, "c"))
    cell.set("r", ref)
    for idx, child in enumerate(list(row)):
        if child.tag == _q(MAIN_NS, "c"):
            cm = re.match(r"([A-Za-z]+)", child.get("r", ""))
            if cm and column_index_from_string(cm.group(1).upper()) > wanted:
                row.insert(idx, cell)
                return cell
    row.append(cell)
    return cell


def _clear_rows(root: etree._Element, mapping: SheetMapping, start: int, end: int) -> None:
    sd = root.find(_q(MAIN_NS, "sheetData"))
    rows = {int(r.get("r")): r for r in sd.findall(_q(MAIN_NS, "row")) if r.get("r")}
    leaves = sorted(set(mapping.leaf_columns.values()))
    for rnum in range(start, end + 1):
        row = rows.get(rnum)
        if row is None:
            continue
        for ref in [f"{mapping.description_col}{rnum}", *(f"{leaf}{rnum}" for leaf in leaves)]:
            cell = _find_cell(row, ref)
            if cell is None:
                continue
            if cell.find(_q(MAIN_NS, "f")) is not None:
                raise ClaimToolError("template already has formulas in expense grid; upload an empty template")
            for ch in list(cell):
                cell.remove(ch)
            cell.attrib.pop("t", None)


def _set_inline_string(cell: etree._Element, text: str) -> None:
    if cell.find(_q(MAIN_NS, "f")) is not None:
        raise ClaimToolError(f"refusing to overwrite formula cell {cell.get('r')!r}")
    if XML_ILLEGAL.search(text) or len(text) > 32_767:
        raise ClaimToolError("invalid text for Excel")
    cell.set("t", "inlineStr")
    for ch in list(cell):
        cell.remove(ch)
    is_el = etree.SubElement(cell, _q(MAIN_NS, "is"))
    t_el = etree.SubElement(is_el, _q(MAIN_NS, "t"))
    t_el.text = text
    if text != text.strip() or "\n" in text or "\r" in text:
        t_el.set(XML_SPACE, "preserve")


def _set_number(cell: etree._Element, val: float) -> None:
    if cell.find(_q(MAIN_NS, "f")) is not None:
        return
    if not math.isfinite(val) or val < 0:
        raise ClaimToolError("refusing to write non-finite or negative number")
    cell.attrib.pop("t", None)
    for ch in list(cell):
        cell.remove(ch)
    v_el = etree.SubElement(cell, _q(MAIN_NS, "v"))
    v_el.text = format(float(val), ".15g")


def _set_formula(cell: etree._Element, formula_text: str) -> None:
    cell.attrib.pop("t", None)
    f_el = cell.find(_q(MAIN_NS, "f"))
    if f_el is None:
        for ch in list(cell):
            cell.remove(ch)
        f_el = etree.Element(_q(MAIN_NS, "f"))
        cell.insert(0, f_el)
    else:
        for attr in ("t", "ref", "si"):
            f_el.attrib.pop(attr, None)
        for ch in list(f_el):
            f_el.remove(ch)
    f_el.text = formula_text
    cached = cell.find(_q(MAIN_NS, "v"))
    if cached is not None:
        cell.remove(cached)


def _vnd_rewrite_columns(
    mapping: SheetMapping,
    foreign_entries: list[tuple[Receipt, float]],
    vnd_entries: list[tuple[Receipt, float]],
) -> list[str]:
    foreign_by_col: dict[str, list[Receipt]] = {}
    for r, _ in foreign_entries:
        foreign_by_col.setdefault(mapping.leaf_columns[r.column], []).append(r)
    vnd_cols: list[str] = []
    for r, _ in vnd_entries:
        col = mapping.leaf_columns[r.column]
        if col not in vnd_cols:
            vnd_cols.append(col)
    rewrite: list[str] = []
    for col in vnd_cols:
        rows = foreign_by_col.get(col, [])
        paired = [r for r in rows if r.vnd_amount is not None]
        unpaired = [r for r in rows if r.vnd_amount is None]
        if paired and unpaired:
            raise ClaimToolError(f"column {col} mixes paired and unpaired foreign rows")
        if not unpaired:
            rewrite.append(col)
    return rewrite


def _write_optional_headers(root: etree._Element, mapping: SheetMapping, header: dict[str, Any]) -> None:
    for key, raw_val in header.items():
        if raw_val is None or key not in mapping.header_cells:
            continue
        ref = mapping.header_cells[key]
        m = re.match(r"([A-Za-z]+)(\d+)", ref)
        if not m:
            continue
        row_el = _find_row(root, int(m.group(2)))
        cell = _get_cell(row_el, ref)
        if key == "advances":
            _set_number(cell, float(raw_val))
        else:
            text = str(raw_val).strip()
            if text:
                _set_inline_string(cell, text)


def fill_overseas(
    template_path: Path,
    mapping: SheetMapping,
    receipts: list[Receipt],
    dest: Path,
    header: dict[str, Any] | None = None,
) -> None:
    foreign_entries = [(r, r.amount) for r in receipts if r.currency_kind != "VND"]
    vnd_entries: list[tuple[Receipt, float]] = []
    for r in receipts:
        if r.currency_kind == "VND":
            vnd_entries.append((r, r.amount))
        elif r.vnd_amount is not None:
            vnd_entries.append((r, r.vnd_amount))

    if len(foreign_entries) > mapping.data_row_end - mapping.data_row_start + 1:
        raise ClaimToolError("foreign receipt count exceeds template capacity")
    if vnd_entries:
        if mapping.vnd_data_row_start is None or mapping.vnd_data_row_end is None:
            raise ClaimToolError("template has no Expenses in VND block")
        if len(vnd_entries) > mapping.vnd_data_row_end - mapping.vnd_data_row_start + 1:
            raise ClaimToolError("VND entry count exceeds template capacity")

    rewrite_cols = _vnd_rewrite_columns(mapping, foreign_entries, vnd_entries) if vnd_entries else []
    if rewrite_cols and mapping.total_row is None:
        raise ClaimToolError("template missing Total in VND row")

    shutil.copy2(template_path, dest)
    tmp = dest.with_name(dest.name + ".tmp")
    with zipfile.ZipFile(dest) as src:
        target_xml = _sheet_xml_path(src, mapping.sheet_name)
        root = etree.fromstring(src.read(target_xml), parser=_safe_xml_parser())

        _clear_rows(root, mapping, mapping.data_row_start, mapping.data_row_end)
        if mapping.vnd_data_row_start is not None and mapping.vnd_data_row_end is not None:
            _clear_rows(root, mapping, mapping.vnd_data_row_start, mapping.vnd_data_row_end)

        if header:
            _write_optional_headers(root, mapping, header)

        for entries, start in (
            (foreign_entries, mapping.data_row_start),
            (vnd_entries, mapping.vnd_data_row_start or 0),
        ):
            for idx, (r, amt) in enumerate(entries):
                rnum = start + idx
                row_el = _find_row(root, rnum)
                row_el.attrib.pop("hidden", None)
                _set_inline_string(_get_cell(row_el, f"{mapping.description_col}{rnum}"), r.purpose)
                _set_number(_get_cell(row_el, f"{mapping.leaf_columns[r.column]}{rnum}"), amt)

        if rewrite_cols and mapping.total_row is not None and mapping.vnd_data_row_start is not None:
            tot_row = _find_row(root, mapping.total_row)
            masters = [
                (c.get("r", ""), f.get("ref", ""), f.text or "")
                for c in tot_row.findall(_q(MAIN_NS, "c"))
                if (f := c.find(_q(MAIN_NS, "f"))) is not None and f.get("t") == "shared" and f.get("ref")
            ]
            for mref, rrange, mtext in masters:
                min_c, min_r, max_c, max_r = range_boundaries(rrange)
                origin = mtext if mtext.startswith("=") else "=" + mtext
                for rn in range(min_r, max_r + 1):
                    for cn in range(min_c, max_c + 1):
                        ref = f"{get_column_letter(cn)}{rn}"
                        cell = _find_cell(tot_row, ref)
                        if cell is not None and (cf := cell.find(_q(MAIN_NS, "f"))) is not None and cf.get("t") == "shared":
                            _set_formula(cell, Translator(origin, origin=mref).translate_formula(ref).lstrip("="))
            last_row = mapping.total_row - 1
            for col in rewrite_cols:
                _set_formula(
                    _get_cell(tot_row, f"{col}{mapping.total_row}"),
                    f"SUM({col}{mapping.vnd_data_row_start}:{col}{last_row})",
                )

        if mapping.exchange_rate_row is not None:
            tot_vnd = sum(r.vnd_amount for r, _ in foreign_entries if r.vnd_amount is not None)
            tot_for = sum(r.amount for r, _ in foreign_entries if r.vnd_amount is not None and r.amount > 0)
            if tot_for > 0:
                rate = round(tot_vnd / tot_for, 3)
                rate_row = _find_row(root, mapping.exchange_rate_row)
                for col in mapping.amount_columns:
                    _set_number(_get_cell(rate_row, f"{col}{mapping.exchange_rate_row}"), rate)

        for cell in root.iter(_q(MAIN_NS, "c")):
            if cell.find(_q(MAIN_NS, "f")) is not None:
                cached = cell.find(_q(MAIN_NS, "v"))
                if cached is not None:
                    cell.remove(cached)
                cell.attrib.pop("t", None)

        patched_sheet = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
        wb_root = etree.fromstring(src.read("xl/workbook.xml"), parser=_safe_xml_parser())
        calc = wb_root.find(_q(MAIN_NS, "calcPr"))
        if calc is None:
            calc = etree.SubElement(wb_root, _q(MAIN_NS, "calcPr"))
        calc.set("fullCalcOnLoad", "1")
        patched_wb = etree.tostring(wb_root, encoding="UTF-8", xml_declaration=True, standalone=True)

        with zipfile.ZipFile(tmp, "w") as out:
            for info in src.infolist():
                data = patched_sheet if info.filename == target_xml else (
                    patched_wb if info.filename == "xl/workbook.xml" else src.read(info.filename)
                )
                out.writestr(info, data)
    os.replace(tmp, dest)


def _parse_date(val: str | None) -> date | None:
    if not val:
        return None
    raw = val.strip().split(",")[0].strip()
    for cand, fmt in ((raw[:10], "%Y-%m-%d"), (raw, "%d/%m/%Y"), (raw, "%d-%m-%Y")):
        try:
            return datetime.strptime(cand, fmt).date()
        except ValueError:
            continue
    return None


def _fmt_money(val: float | None, *, vnd: bool = False) -> str:
    if val is None:
        return "—"
    dec = (0 if float(val).is_integer() else 2) if vnd else 2
    return f"{val:,.{dec}f}"


def _img_jpeg_stream(path: str | Path) -> tuple[BytesIO, int, int]:
    with Image.open(path) as opened:
        img = ImageOps.exif_transpose(opened).convert("RGBA")
    bg = Image.new("RGB", img.size, "white")
    bg.paste(img, mask=img.getchannel("A"))
    bg.thumbnail((1400, 1400), Image.Resampling.LANCZOS)
    buf = BytesIO()
    bg.save(buf, format="JPEG", quality=85, optimize=True)
    buf.seek(0)
    return buf, bg.width, bg.height


def build_proof_pack(receipts: list[Receipt], dest: Path) -> None:
    doc = Document()
    sec = doc.sections[0]
    sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
    sec.left_margin = sec.right_margin = Cm(1.6)
    sec.top_margin = sec.bottom_margin = Cm(1.5)
    doc.styles["Normal"].font.name = "Calibri"
    doc.styles["Normal"].font.size = Pt(10)

    doc.add_heading("Receipt Proof Pack", level=1).alignment = WD_ALIGN_PARAGRAPH.LEFT
    sub = doc.add_paragraph(f"Overseas travel expense claim · {len(receipts)} receipt(s) · grouped by bill date")
    sub.paragraph_format.space_after = Pt(4)

    groups: dict[date | None, list[Receipt]] = {}
    for r in receipts:
        groups.setdefault(_parse_date(r.date), []).append(r)

    for bill_date, rows in sorted(groups.items(), key=lambda item: (item[0] is None, item[0] or date.max)):
        hd = doc.add_heading(bill_date.strftime("%d/%m/%Y") if bill_date else "Undated bills", level=2)
        hd.paragraph_format.space_before = Pt(14)
        hd.paragraph_format.space_after = Pt(2)
        hd.paragraph_format.keep_with_next = True

        for r in rows:
            p = doc.add_paragraph()
            p.paragraph_format.keep_with_next = True
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run(r.purpose.strip() or "(no purpose)")
            run.bold = True
            run.font.size = Pt(11)

            cur = "VND" if r.currency_kind == "VND" else normalize_currency(r.currency)
            if cur == "VND" and r.currency_kind != "VND":
                cur = ""
            parts = [f"VND {_fmt_money(r.amount, vnd=True)}"] if r.currency_kind == "VND" else [
                f"{cur + ' ' if cur else 'Foreign amount '}{_fmt_money(r.amount)}".strip()
            ]
            if r.currency_kind != "VND" and r.vnd_amount is not None:
                parts.append(f"VND deducted {_fmt_money(r.vnd_amount, vnd=True)}")
            amt_p = doc.add_paragraph(" · ".join(parts))
            amt_p.paragraph_format.keep_with_next = True
            amt_p.paragraph_format.space_after = Pt(6)

            images: list[tuple[str, str]] = [("Foreign-currency bill", r.stored_path)]
            has_extra_vnd = any(ep.vnd_amount is not None for ep in r.extra_proofs)
            if r.vnd_proof_path:
                if r.vnd_proof_source == "bank_statement":
                    lbl = "Bank statement payment (VND amount circled)"
                elif has_extra_vnd and r.vnd_primary_amount is not None:
                    lbl = f"VND deduction proof 1 ({_fmt_money(r.vnd_primary_amount, vnd=True)} VND)"
                else:
                    lbl = "VND deduction proof"
                images.append((lbl, r.vnd_proof_path))
            for idx, ep in enumerate(r.extra_proofs, start=1):
                if ep.caption:
                    lbl = ep.caption
                elif ep.vnd_amount is not None:
                    num = idx + 1 if r.vnd_proof_path else idx
                    lbl = f"VND deduction proof {num} ({_fmt_money(ep.vnd_amount, vnd=True)} VND)"
                else:
                    lbl = f"Additional proof {idx + 1 if r.vnd_proof_path else idx}"
                images.append((lbl, ep.stored_path))

            full_width = r.vnd_proof_source == "bank_statement" or any(
                ep.caption and ep.caption.startswith("Bank statement") for ep in r.extra_proofs
            )
            cols = 1 if full_width or len(images) == 1 else 2
            max_w = Inches(5.8) if cols == 1 else Inches(3.15)
            max_h = Inches(6.4) if cols == 1 else Inches(5.0)
            tbl = doc.add_table(rows=(len(images) + cols - 1) // cols, cols=cols)
            tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
            for row_obj in tbl.rows:
                row_obj._tr.get_or_add_trPr().append(OxmlElement("w:cantSplit"))
            for idx, (cap, img_path) in enumerate(images):
                cell = tbl.rows[idx // cols].cells[idx % cols]
                stream, pw, ph = _img_jpeg_stream(img_path)
                w_in = max_w.inches
                h_in = w_in * ph / pw
                if h_in > max_h.inches:
                    h_in = max_h.inches
                    w_in = h_in * pw / ph
                cp = cell.paragraphs[0]
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cp.add_run().add_picture(stream, width=Inches(w_in))
                cap_p = cell.add_paragraph()
                cap_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cr = cap_p.add_run(cap)
                cr.italic = True
                cr.font.size = Pt(8)
            if cols == 2 and len(images) % 2 == 1 and len(images) > 1:
                tbl.rows[-1].cells[0].merge(tbl.rows[-1].cells[1])
            doc.add_paragraph().paragraph_format.space_after = Pt(0)
    doc.save(dest)


def _fold(s: str) -> str:
    nfd = unicodedata.normalize("NFD", s)
    clean = "".join(ch for ch in nfd if not unicodedata.combining(ch)).replace("đ", "d").replace("Đ", "D")
    return " ".join(re.sub(r"[^a-z0-9]+", " ", clean.casefold()).split())


def _money_val(tok: str) -> float | None:
    raw = re.sub(r"[A-Za-z()]", "", tok.strip().replace("\u00a0", "").replace(" ", ""))
    if not raw:
        return None
    neg = raw.startswith("-") or raw.endswith("-")
    raw = re.sub(r"[.,]00$", "", raw)
    digits = re.sub(r"\D", "", raw)
    return (-float(digits) if neg else float(digits)) if digits else None


def crop_bank_row(pdf_path: Path, amount: float, date_str: str | None, desc: str, page_hint: int | None, out: Path) -> None:
    doc = pymupdf.open(pdf_path)
    try:
        pages = list(range(1, doc.page_count + 1))
        if page_hint and page_hint in pages:
            pages.remove(page_hint)
            pages.insert(0, page_hint)
        tokens = [t for t in _fold(desc).split() if len(t) >= 3]
        target = abs(float(amount))
        best = None
        for pno in pages:
            page = doc[pno - 1]
            words = page.get_text("words")
            lines = []
            for blk in page.get_text("dict").get("blocks", []):
                for ln in blk.get("lines", []):
                    txt = "".join(sp.get("text", "") for sp in ln.get("spans", [])).strip()
                    if txt:
                        lines.append((pymupdf.Rect(ln["bbox"]), txt))
            lines.sort(key=lambda it: (it[0].y0, it[0].x0))
            groups: list[dict[str, Any]] = []
            for rect, txt in lines:
                cy = (rect.y0 + rect.y1) / 2
                placed = False
                for g in groups:
                    if abs(cy - (g["rect"].y0 + g["rect"].y1) / 2) <= max(4.0, rect.height * 0.7):
                        g["rect"] |= rect
                        g["texts"].append(txt)
                        placed = True
                        break
                if not placed:
                    groups.append({"rect": pymupdf.Rect(rect), "texts": [txt]})
            for g in groups:
                gwords = [w for w in words if g["rect"].y0 - 2 <= (w[1] + w[3]) / 2 <= g["rect"].y1 + 2]
                hits = [w for w in gwords if (v := _money_val(str(w[4]))) is not None and abs(abs(v) - target) <= 0.5]
                if not hits:
                    continue
                hits.sort(key=lambda w: (0 if str(w[4]).strip().startswith("-") or str(w[4]).strip().endswith("-") else 1, w[0]))
                aw = hits[0]
                folded = _fold(" ".join(g["texts"]))
                sc = 2.0 + (0.5 if pno == page_hint else 0.0) + min(1.5, sum(1 for t in tokens if t in folded) * 0.35)
                if date_str and date_str[:10] in folded:
                    sc += 1.0
                if best is None or sc > best[0]:
                    best = (sc, page, g["rect"], pymupdf.Rect(aw[:4]), gwords)
        if best is not None:
            _sc, page, lrect, arect, gwords = best
            drect = next((pymupdf.Rect(w[:4]) for w in sorted(gwords, key=lambda x: x[0]) if re.fullmatch(r"\d{1,2}[/.-]\d{1,2}[/.-]\d{2,4}", str(w[4]))), None)
            pr = page.rect
            left = max(pr.x0 + 6, (drect.x1 + 6) if drect else (arect.x0 - pr.width * 0.42))
            if left > arect.x0 - 12:
                left = max(pr.x0 + 6, arect.x0 - pr.width * 0.25)
            right = min(pr.x1 - 6, arect.x1 + 12)
            top, bottom = max(pr.y0 + 4, lrect.y0 - 7), min(pr.y1 - 4, lrect.y1 + 7)
            clip = pymupdf.Rect(left, top, right, bottom)
            pix = page.get_pixmap(matrix=pymupdf.Matrix(4.0, 4.0), clip=clip, alpha=False)
            img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
            box = ((arect.x0 - clip.x0) * 4 - 20, (arect.y0 - clip.y0) * 4 - 12, (arect.x1 - clip.x0) * 4 + 20, (arect.y1 - clip.y0) * 4 + 12)
            ImageDraw.Draw(img).ellipse(box, outline=(210, 30, 35), width=5)
            img.save(out, format="PNG", optimize=True)
            return
        raise ClaimToolError("could not locate payment row in native PDF text")
    finally:
        doc.close()


def _load_claim(claim_json: Path) -> tuple[Path, list[Receipt], dict[str, Any] | None]:
    data = json.loads(claim_json.read_text(encoding="utf-8"))
    template = Path(data["template_path"])
    if not template.is_file():
        raise ClaimToolError(f"template not found: {template}")
    if not data.get("receipts"):
        raise ClaimToolError("claim has no receipts")
    receipts: list[Receipt] = []
    for idx, raw in enumerate(data["receipts"], start=1):
        amt = raw.get("amount")
        if amt is None or not math.isfinite(amt) or amt < 0 or amt >= 1e12:
            raise ClaimToolError(f"receipt {idx} needs a finite amount")
        if not str(raw.get("purpose") or "").strip():
            raise ClaimToolError(f"receipt {idx} needs a purpose")
        col = raw.get("column")
        if col not in EXPENSE_COLUMNS or not raw.get("column_confirmed", False):
            raise ClaimToolError(f"receipt {idx} column must be valid and confirmed")
        extras = [
            ExtraProof(
                stored_path=ep["stored_path"],
                caption=ep.get("caption"),
                vnd_amount=ep.get("vnd_amount"),
                amount_included=bool(ep.get("amount_included", False)),
                source=ep.get("source"),
            )
            for ep in raw.get("extra_proofs") or []
        ]
        vnd_amt = raw.get("vnd_amount")
        vnd_pri = raw.get("vnd_primary_amount")
        if vnd_amt is None and (vnd_pri is not None or any(ep.amount_included and ep.vnd_amount is not None for ep in extras)):
            vnd_amt = float(vnd_pri or 0.0) + sum(
                float(ep.vnd_amount) for ep in extras if ep.amount_included and ep.vnd_amount is not None
            )
        r = Receipt(
            stored_path=raw["stored_path"],
            purpose=str(raw["purpose"]).strip(),
            column=col,
            amount=float(amt),
            currency_kind=raw.get("currency_kind", "FOREIGN"),
            currency=raw.get("currency"),
            date=raw.get("date"),
            column_confirmed=True,
            vnd_amount=vnd_amt,
            vnd_primary_amount=vnd_pri,
            vnd_proof_path=raw.get("vnd_proof_path"),
            vnd_proof_source=raw.get("vnd_proof_source"),
            extra_proofs=extras,
        )
        if r.currency_kind != "VND" and r.vnd_proof_path and r.vnd_amount is None:
            raise ClaimToolError(f"receipt {idx} has a VND proof but no vnd_amount")
        receipts.append(r)
    return template, receipts, data.get("header")


def main() -> None:
    ap = argparse.ArgumentParser(description="Overseas travel claim exporter & bank cropper")
    sub = ap.add_subparsers(dest="cmd", required=True)
    exp = sub.add_parser("export")
    exp.add_argument("--claim", type=Path, required=True)
    exp.add_argument("--out-dir", type=Path, required=True)
    crp = sub.add_parser("crop-bank")
    crp.add_argument("--pdf", type=Path, required=True)
    crp.add_argument("--amount", type=float, required=True)
    crp.add_argument("--date", default=None)
    crp.add_argument("--description", default="")
    crp.add_argument("--page", type=int, default=None)
    crp.add_argument("--out", type=Path, required=True)
    args = ap.parse_args()

    if args.cmd == "crop-bank":
        crop_bank_row(args.pdf, args.amount, args.date, args.description, args.page, args.out)
        print(args.out)
        return

    template, receipts, header = _load_claim(args.claim)
    args.out_dir.mkdir(parents=True, exist_ok=True)
    stem = template.stem
    xlsx_out = args.out_dir / f"{stem} - filled.xlsx"
    docx_out = args.out_dir / f"{stem} - proof pack.docx"
    zip_out = args.out_dir / f"{stem} - claim pack.zip"

    mapping = inspect_overseas_mapping(template)
    fill_overseas(template, mapping, receipts, xlsx_out, header=header)
    build_proof_pack(receipts, docx_out)
    with zipfile.ZipFile(zip_out, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.write(xlsx_out, arcname=xlsx_out.name)
        zf.write(docx_out, arcname=docx_out.name)
    print(json.dumps({"xlsx": str(xlsx_out), "docx": str(docx_out), "zip": str(zip_out)}))


if __name__ == "__main__":
    main()
```

---

## 7. Verification Checklist

After running `export`, always verify before delivering to the user:
1. Open `<stem> - filled.xlsx` with `openpyxl(..., data_only=False)`:
   - Purposes and amounts landed in the discovered foreign & VND rows (with written rows unhidden and unused rows untouched).
   - Paired columns have `SUM(<col><vnd_start>:<col><total_row-1>)` on `Total in VND` (no double counting).
   - `Exchange Rate` row matches `round(sum(vnd_amount) / sum(foreign_amount), 3)` across expense columns.
   - `xl/media/*` logo and `xl/workbook.xml` `fullCalcOnLoad="1"` are intact.
2. Open `<stem> - proof pack.docx` with `python-docx`:
   - `len(doc.inline_shapes)` equals total attached bill + primary + extra images.
   - Date headings (`dd/mm/yyyy`), normalized currency labels, purposes, and VND totals match the claim JSON.
3. If running the web app locally alongside `.env`, ensure shell env vars do not silently shadow `.env` (`load_dotenv(override=True)` or unset stale `DEEPSEEK_API_KEY` in the shell).
