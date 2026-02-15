# FOP Group 3 Tax Automation (Windows)

A desktop-first Python app that can be packaged into a Windows `.exe`.

## What it does

1. On launch, opens a folder picker for bank statement PDFs.
2. Extracts PDF text.
3. Sends statement text to OpenAI (ChatGPT 5.2 style reasoning) to identify **eligible income** for Ukrainian FOP Group 3 taxation.
4. Excludes:
   - own transfers between your accounts;
   - currency exchange transactions;
   - non-income operations.
5. Calculates:
   - Unified Tax (ЄП) = **5%** of eligible income;
   - Military tax = **1%** of eligible income.
6. Groups data by quarter and produces YTD values:
   - H1 = Q1 + Q2
   - 9M = Q1 + Q2 + Q3
   - 12M = Q1 + Q2 + Q3 + Q4
7. Exports a user-friendly Excel report in Ukrainian with separate sheets for selected income, tax summary, and token usage log.
8. Shows a local progress window (file-by-file status + progress bar) during processing.

---

## Quick start

### 1) Install dependencies

```bash
python -m venv .venv
source .venv/bin/activate  # Windows: .venv\\Scripts\\activate
pip install -r requirements.txt
```

### 2) Configure API key

Set environment variable:

```bash
export OPENAI_API_KEY="your_key"  # Windows PowerShell: $env:OPENAI_API_KEY="your_key"
```

### 3) Run

```bash
python app/main.py
```

The app asks for a folder with PDFs and generates `tax_report.xlsx` in that folder.

---

## Build Windows EXE

```bash
pyinstaller --noconfirm --onefile --windowed --name fop-tax-automation app/main.py
```

After build, executable appears in `dist/fop-tax-automation.exe`.

---

## Important notes

- This tool assists accounting workflows but does **not** replace professional tax advice.
- Always review selected transactions before submitting declarations.
- Income limit for FOP Group 3 can change annually. The app supports a configurable limit and highlights threshold status.

---

## Suggested improvements

- Add OCR fallback for scanned PDFs (`pytesseract` + `pdf2image`).
- Add local transaction parser rules before AI call (hybrid deterministic + AI).
- Add digital signature workflow and declaration draft export.
- Add audit log with per-transaction reasoning and confidence score.
- Add automatic annual income limit update from official tax data source API/page.
