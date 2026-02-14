from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tkinter import Tk, filedialog, messagebox

import openpyxl
import pdfplumber
from openai import OpenAI

MODEL_NAME = os.getenv("OPENAI_MODEL", "gpt-5")
GROUP3_LIMIT_UAH = Decimal(os.getenv("GROUP3_LIMIT_UAH", "9000000"))


@dataclass
class IncomeTransaction:
    date: datetime
    description: str
    amount_uah: Decimal
    source_file: str


def select_folder() -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Select folder with bank statement PDFs")
    root.destroy()
    if not folder:
        return None
    return Path(folder)


def extract_pdf_text(pdf_path: Path) -> str:
    chunks: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            chunks.append(page.extract_text() or "")
    return "\n".join(chunks)


def call_ai_extract_income(client: OpenAI, statement_text: str, source_file: str) -> list[IncomeTransaction]:
    schema_instruction = {
        "type": "array",
        "items": {
            "type": "object",
            "properties": {
                "date": {"type": "string", "description": "YYYY-MM-DD"},
                "description": {"type": "string"},
                "amount_uah": {"type": "number"},
            },
            "required": ["date", "description", "amount_uah"],
            "additionalProperties": False,
        },
    }

    prompt = f"""
You are an accounting extraction assistant for Ukraine FOP Group 3 tax calculation.

Task:
- Parse the bank statement text.
- Select only real business income from external customers.
- EXCLUDE own-account transfers, cash-in/out between own cards/accounts, currency exchange, refunds/reversals, and non-income operations.
- If uncertain, be conservative and exclude.
- Return JSON only as an array of objects with: date (YYYY-MM-DD), description, amount_uah.
- amount_uah must be positive number in UAH equivalent.

Source file: {source_file}

Statement text:
{statement_text[:120000]}
"""

    response = client.responses.create(
        model=MODEL_NAME,
        reasoning={"effort": "high"},
        input=[
            {
                "role": "system",
                "content": "Return strictly valid JSON with no markdown.",
            },
            {"role": "user", "content": prompt},
            {"role": "user", "content": f"JSON schema: {json.dumps(schema_instruction)}"},
        ],
    )

    raw = response.output_text.strip()
    data = json.loads(raw)
    txs: list[IncomeTransaction] = []
    for item in data:
        txs.append(
            IncomeTransaction(
                date=datetime.strptime(item["date"], "%Y-%m-%d"),
                description=item["description"].strip(),
                amount_uah=Decimal(str(item["amount_uah"])),
                source_file=source_file,
            )
        )
    return txs


def quarter_for_date(dt: datetime) -> str:
    q = ((dt.month - 1) // 3) + 1
    return f"Q{q} {dt.year}"


def build_report(transactions: list[IncomeTransaction], out_path: Path) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Selected Income"
    ws.append(["Date", "Quarter", "Description", "Amount UAH", "Source PDF"])

    for tx in sorted(transactions, key=lambda t: t.date):
        ws.append(
            [
                tx.date.date().isoformat(),
                quarter_for_date(tx.date),
                tx.description,
                float(tx.amount_uah),
                tx.source_file,
            ]
        )

    summary = wb.create_sheet("Tax Summary")
    summary.append(["Metric", "Value (UAH)"])

    by_quarter: dict[str, Decimal] = {}
    for tx in transactions:
        quarter = quarter_for_date(tx.date)
        by_quarter[quarter] = by_quarter.get(quarter, Decimal("0")) + tx.amount_uah

    sorted_quarters = sorted(
        by_quarter.keys(),
        key=lambda q: (int(q.split()[1]), int(q[1])),
    )

    ytd = Decimal("0")
    summary.append(["", ""])
    summary.append(["Quarter", "Eligible income"])
    for q in sorted_quarters:
        ytd += by_quarter[q]
        summary.append([q, float(by_quarter[q])])

    unified_tax = ytd * Decimal("0.05")
    military_tax = ytd * Decimal("0.01")

    summary.append(["", ""])
    summary.append(["Total eligible income (YTD)", float(ytd)])
    summary.append(["Unified Tax 5%", float(unified_tax)])
    summary.append(["Military tax 1%", float(military_tax)])
    summary.append(["Group 3 income limit", float(GROUP3_LIMIT_UAH)])
    summary.append(["Limit remaining", float(GROUP3_LIMIT_UAH - ytd)])
    summary.append(["Limit exceeded", "YES" if ytd > GROUP3_LIMIT_UAH else "NO"])

    wb.save(out_path)


def main() -> None:
    if not os.getenv("OPENAI_API_KEY"):
        messagebox.showerror("Configuration error", "OPENAI_API_KEY is not set.")
        return

    folder = select_folder()
    if folder is None:
        return

    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        messagebox.showwarning("No PDFs", "No PDF files found in selected folder.")
        return

    client = OpenAI()
    all_transactions: list[IncomeTransaction] = []

    for pdf in pdfs:
        text = extract_pdf_text(pdf)
        if not text.strip():
            continue
        txs = call_ai_extract_income(client, text, pdf.name)
        all_transactions.extend(txs)

    if not all_transactions:
        messagebox.showwarning(
            "No eligible income found",
            "No eligible income transactions were extracted. Please review source PDFs.",
        )
        return

    output_path = folder / "tax_report.xlsx"
    build_report(all_transactions, output_path)
    messagebox.showinfo("Done", f"Report generated:\n{output_path}")


if __name__ == "__main__":
    main()
