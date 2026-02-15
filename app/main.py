from __future__ import annotations

import json
import os
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal
from pathlib import Path
from tkinter import Tk, filedialog, messagebox, ttk

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


@dataclass
class UsageStats:
    input_tokens: int = 0
    output_tokens: int = 0
    total_tokens: int = 0
    calls: int = 0


def select_folder() -> Path | None:
    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    folder = filedialog.askdirectory(title="Оберіть папку з PDF-виписками")
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


def update_usage_stats(response, usage_stats: UsageStats) -> None:
    usage = getattr(response, "usage", None)
    if usage is None:
        return
    usage_stats.input_tokens += int(getattr(usage, "input_tokens", 0) or 0)
    usage_stats.output_tokens += int(getattr(usage, "output_tokens", 0) or 0)
    usage_stats.total_tokens += int(getattr(usage, "total_tokens", 0) or 0)
    usage_stats.calls += 1


def call_ai_extract_income(
    client: OpenAI,
    statement_text: str,
    source_file: str,
    usage_stats: UsageStats,
) -> list[IncomeTransaction]:
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
    update_usage_stats(response, usage_stats)

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


def build_report(transactions: list[IncomeTransaction], out_path: Path, usage_stats: UsageStats) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Відібрані надходження"
    ws.append(["Дата", "Квартал", "Опис", "Сума, грн", "PDF-файл"])

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

    summary = wb.create_sheet("Підсумок податків")
    summary.append(["Показник", "Значення (грн)"])

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
    summary.append(["Квартал", "Оподатковуваний дохід"])
    for q in sorted_quarters:
        ytd += by_quarter[q]
        summary.append([q, float(by_quarter[q])])

    unified_tax = ytd * Decimal("0.05")
    military_tax = ytd * Decimal("0.01")

    summary.append(["", ""])
    summary.append(["Загальний дохід (YTD)", float(ytd)])
    summary.append(["Єдиний податок 5%", float(unified_tax)])
    summary.append(["Військовий збір 1%", float(military_tax)])
    summary.append(["Ліміт доходу ФОП 3 групи", float(GROUP3_LIMIT_UAH)])
    summary.append(["Залишок до ліміту", float(GROUP3_LIMIT_UAH - ytd)])
    summary.append(["Ліміт перевищено", "ТАК" if ytd > GROUP3_LIMIT_UAH else "НІ"])

    token_log = wb.create_sheet("Лог токенів")
    token_log.append(["Метрика", "Значення"])
    token_log.append(["Кількість API-викликів", usage_stats.calls])
    token_log.append(["Вхідні токени", usage_stats.input_tokens])
    token_log.append(["Вихідні токени", usage_stats.output_tokens])
    token_log.append(["Усього токенів", usage_stats.total_tokens])

    wb.save(out_path)


def create_progress_window(total_files: int) -> tuple[Tk, ttk.Label, ttk.Progressbar]:
    window = Tk()
    window.title("Обробка банківських виписок")
    window.resizable(False, False)

    status = ttk.Label(window, text="Підготовка до обробки...")
    status.pack(padx=14, pady=(14, 10))

    progress = ttk.Progressbar(window, length=420, maximum=max(1, total_files), mode="determinate")
    progress.pack(padx=14, pady=(0, 14))
    progress["value"] = 0

    window.update_idletasks()
    window.update()
    return window, status, progress


def set_progress(window: Tk, label: ttk.Label, progress: ttk.Progressbar, value: int, text: str) -> None:
    progress["value"] = value
    label.configure(text=text)
    window.update_idletasks()
    window.update()


def main() -> None:
    if not os.getenv("OPENAI_API_KEY"):
        messagebox.showerror("Помилка конфігурації", "Змінна OPENAI_API_KEY не задана.")
        return

    folder = select_folder()
    if folder is None:
        return

    pdfs = sorted(folder.glob("*.pdf"))
    if not pdfs:
        messagebox.showwarning("Немає PDF", "У вибраній папці не знайдено PDF-файлів.")
        return

    client = OpenAI()
    usage_stats = UsageStats()
    all_transactions: list[IncomeTransaction] = []

    progress_window, status_label, progress_bar = create_progress_window(len(pdfs))
    try:
        for index, pdf in enumerate(pdfs, start=1):
            set_progress(
                progress_window,
                status_label,
                progress_bar,
                index - 1,
                f"Обробка {index}/{len(pdfs)}: читання {pdf.name}",
            )
            text = extract_pdf_text(pdf)
            if not text.strip():
                set_progress(
                    progress_window,
                    status_label,
                    progress_bar,
                    index,
                    f"Пропущено порожній PDF: {pdf.name}",
                )
                continue

            set_progress(
                progress_window,
                status_label,
                progress_bar,
                index - 1,
                f"Обробка {index}/{len(pdfs)}: AI-аналіз {pdf.name}",
            )
            txs = call_ai_extract_income(client, text, pdf.name, usage_stats)
            all_transactions.extend(txs)
            set_progress(
                progress_window,
                status_label,
                progress_bar,
                index,
                f"Завершено {index}/{len(pdfs)}: {pdf.name}",
            )

        if not all_transactions:
            messagebox.showwarning(
                "Надходження не знайдено",
                "Оподатковувані надходження не виявлено. Перевірте PDF-файли.",
            )
            return

        set_progress(progress_window, status_label, progress_bar, len(pdfs), "Формування Excel-звіту...")
        output_path = folder / "tax_report.xlsx"
        build_report(all_transactions, output_path, usage_stats)
        set_progress(progress_window, status_label, progress_bar, len(pdfs), "Готово")
        messagebox.showinfo("Готово", f"Звіт сформовано:\n{output_path}")
    finally:
        progress_window.destroy()


if __name__ == "__main__":
    main()
