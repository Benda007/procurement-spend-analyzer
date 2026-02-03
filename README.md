# Procurement Spend Analyzer 📊

Automated tool for analyzing supplier spend data and generating actionable insights for procurement professionals.
Designed to turn raw ERP/procurement exports into a clean dataset, summary metrics, and ready-to-use charts.

> Time saving idea: replace repetitive Excel spend analysis with a repeatable Python workflow (load → clean → analyze → report → visualize → export).

---

## 🎯 Problem

Procurement teams often spend hours every month manually analyzing spend data in Excel:
- Identifying top suppliers and spend patterns
- Calculating year-over-year (YoY) trends
- Building visualizations for stakeholder reporting
- Extracting category-level insights

---

## 💡 Solution

This tool automates the entire spend analysis workflow:
- Loads data from CSV / Excel (`.xlsx` + legacy `.xls`) exports
- Standardizes column names using a built-in synonym mapping (e.g., *vendor* → *supplier*)
- Cleans and validates spend + year formats (including EU/US number formats)
- Calculates key metrics: top suppliers, category breakdown, YoY trends
- Generates charts and saves a dashboard image (`spend_analysis.png`)
- Exports cleaned data to Excel (`cleaned_data.xlsx`) or CSV

---

## 🧰 Tech Stack

- Python 3.10+ (uses modern typing like `str | Path`)
- Pandas (data processing)
- Matplotlib (visualization)
- openpyxl (export to `.xlsx`)
- xlrd (read legacy `.xls` files)

---

## 📦 Installation

1) Create and activate a virtual environment (recommended but not mandatory):

```bash
python -m venv .venv
# Windows:
.venv\Scriptsactivate
# macOS/Linux:
source .venv/bin/activate
```

2) Install dependencies:

```bash
pip install -r requirements.txt
```

---

## ▶️ Usage

### Basic run (CSV / XLS / XLSX)

```bash
python spend_analyser.py path/to/your_file.csv
```

If you do not provide a file path, the script searches for files named like:
- `sample_data*.csv`
- `sample_data*.xls`
- `sample_data*.xlsx`

…and lets you choose one interactively:

```bash
python spend_analyser.py
```

### Optional arguments

- `--sep` / `-s` : set CSV separator manually (e.g. `;` for EU exports)
- `--export-format` / `-e` : choose export format (`xlsx` or `csv`)
- `--keep-negative` : keep negative spend values (credits/debit notes) instead of filtering them out

### Examples

```bash
# CSV with semicolon delimiter, export cleaned data to CSV
python spend_analyser.py data.csv --sep ";" --export-format csv
```

```bash
# Excel input, export cleaned data to Excel
python spend_analyser.py data.xlsx --export-format xlsx
```

```bash
# Keep negative spend values (credits / debit notes) for net spend analysis
python spend_analyser.py data.xlsx --keep-negative
```

---

## 📄 Input Data Requirements

### Required columns
- Supplier (e.g., `supplier`, `vendor`, `supplier name`, `counterpart`, ...)
- Spend (e.g., `spend`, `amount`, `value`, `cost`, ...)

### Optional columns
- Category (e.g., `category`, `group`, `type`)
- Year (e.g., `year`, `fiscal year`, `period`)

The script lowercases and trims headers, then maps them to internal names:
- `supplier`, `spend`, `category`, `year`

---

## 🧹 Data Cleaning Rules

### Spend parsing (EU + US formats)

The tool tries to handle typical formats such as:
- `1 234,56`
- `1.234,56`
- `1,234.56`
- currency symbols / spaces

Invalid amounts become `NaN` and are removed during cleaning.

### Year parsing

If the `year` column exists, it must be either:
- `YYYY` (e.g., `2024`)
- `YYYY-MM-DD` (e.g., `2024-03-15`)

Other formats are rejected with a clear error message.

### Negative spend

Default behavior: filters out negative spend values.
Use `--keep-negative` to keep credits/debit notes for net spend analysis.

---

## 📈 Outputs

After running, you get:

1) Console report with:
- total spend
- average transaction
- supplier count
- top suppliers / categories
- YoY growth (if year data exists)

2) Chart image:
- `spend_analysis.png`

3) Clean dataset export:
- `cleaned_data.xlsx` (default)
- or `cleaned_data.csv` (if `--export-format csv`)

---

## ✅ Notes for Portfolio Reviewers

- Focused on a repeatable, business-oriented workflow used in procurement analytics.
- Designed to be robust across typical ERP exports (encoding issues, separators, EU/US number formats).
- Produces both a clean dataset and an executive-friendly dashboard.
- Sample data provided in portfolio are random data users can train on. 

---

## 📜 License

MIT License
