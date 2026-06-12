# BOE Data Extractor

![Python](https://img.shields.io/badge/Python-3.9+-blue?style=flat-square&logo=python)
![pdfplumber](https://img.shields.io/badge/pdfplumber-0.10-orange?style=flat-square)
![openpyxl](https://img.shields.io/badge/openpyxl-3.1-green?style=flat-square)
![Fields](https://img.shields.io/badge/Fields%20Extracted-90-purple?style=flat-square)
![Status](https://img.shields.io/badge/Status-Production--tested-brightgreen?style=flat-square)

**Automated extraction of 90 finance-critical fields from Indian Customs Bill of Entry (BOE) PDFs into structured, analysis-ready Excel — tested in production at a top Indian e-commerce company on thousands of vendor invoices.**

---

## The problem

AP and finance teams at Indian importers receive Bill of Entry PDFs from customs house agents after every shipment. Each BOE is a dense, multi-page government document containing duty breakdowns, assessable values, IGST details, supplier data, and manifest information spread across 6 structured parts.

Copying these fields manually — Port Code, BE No, CTH, assessable value, BCD rate, IGST amount, MAWB No, OOC date — across hundreds of documents per month takes weeks, introduces transcription errors, and blocks critical downstream workflows: duty reconciliation, IGST ITC claims, and vendor payment processing.

This pipeline automates 100% of that extraction. Drop PDFs in a folder, run one command, get a clean Excel file with one row per line item.

---

## Real-world impact

> Processing time reduced from **45+ days per cycle to under 1 week — 85% faster.**
> Used in production for AP duty reconciliation, IGST Input Tax Credit (ITC) claims, and vendor payment processing at a major Indian e-commerce company.

---

## Input — raw Bill of Entry PDF

A standard Indian Customs BOE from ICEGATE — dense, multi-part, and completely unstructured for machine reading:

![BOE Input Sample](https://raw.githubusercontent.com/CadencePython/-boe-data-extractor/main/boe_input_sample.png)

A single BOE can have **multiple line items**, each with its own duties section:

![BOE Multi-Item](https://raw.githubusercontent.com/CadencePython/-boe-data-extractor/main/boe_input_multiitem.png)

---

## Output — structured Excel

Every field extracted, colour-coded by group, one row per line item:

![BOE Output Excel](https://raw.githubusercontent.com/CadencePython/-boe-data-extractor/main/boe_output_excel.png)

See the full sample: [`BOE_Extracted_sample.xlsx`](BOE_Extracted_sample.xlsx)

---

## What it handles

- **Multiple PDFs** in one run — processes every `.pdf` in the `Input/` folder
- **Multiple BOEs per PDF** — detects each BOE by its Part I page boundary
- **Multiple line items per BOE** — one Excel row per line item; header fields repeated across rows for easy filtering
- **Two BOE layout variants** — handles differences in blank row positions across customs house agents

---

## What it extracts (90 fields across 12 groups)

| Group | Key fields |
|---|---|
| Identification | Source PDF, BOE Index, Item Index |
| BE Header | Port Code, BE No, BE Date, BE Type, BE Mode, IEC/Br, GSTIN/TYPE, CB Code, G.WT, PKG |
| Parties | Importer Name & Address, CB Name, Supplier Name & Address |
| Origin & Routing | Country of Origin, Country of Consignment, Port of Loading, Port of Shipment |
| Manifest | IGM No, IGM Date, INW Date, MAWB No, HAWB No |
| Invoice (shared) | Invoice No, Date, PO No & Date, LC No & Date, Invoice Value, Currency, Freight, Insurance, Pay Terms, Trade Term, Valuation Method, Assessable Value |
| Item details | Item S.No, CTH, CETH, Item Description, Unit Price, Quantity, UQC, Amount |
| Valuation | Exchange Rate |
| Duty Summary (Part I) | BCD Total, ACD, SWS, NCCD, ADD, CVD, IGST Total, GCESS Total, Total Duty, Interest, Penalty, Fine, Total Amount Payable |
| Item duties (Part III) | Per-item: BCD (Notn No/SNo/Rate/Amount), SWS Rate/Amount, IGST (Notn No/SNo/Rate/Amount), GCESS (Notn No/SNo/Rate/Amount), CAIDC (Notn No/SNo/Rate/Amount) |
| OOC | OOC No, OOC Date |
| Invoice Summary | Invoice No, Invoice Amount, Currency |

---

## How to run

**1. Clone the repo**
```bash
git clone https://github.com/CadencePython/-boe-data-extractor.git
cd -boe-data-extractor
```

**2. Install dependencies**
```bash
pip install -r requirements.txt
```

**3. Add your BOE PDFs**
```
Drop all .pdf files into the Input/ folder
```

**4. Run**
```bash
python boe_extractor.py
```

**5. Collect output**
```
Output/BOE_Extracted_YYYYMMDD_HHMMSS.xlsx
One row per line item. Header-level fields repeated on every row for easy pivot/VLOOKUP.
```

---

## Requirements

```
pdfplumber==0.10.3
pandas==2.1.0
openpyxl==3.1.2
xlsxwriter==3.1.9
```

```bash
pip install -r requirements.txt
```

---

## Project structure

```
boe-data-extractor/
├── boe_extractor.py             ← main script (run this)
├── sample_data/
│   └── sample_boe.pdf           ← anonymised sample BOE PDF
├── sample_output/
│   └── BOE_Extracted_sample.xlsx
├── screenshots/
│   ├── boe_input_sample.png     ← Part I summary screenshot
│   ├── boe_input_multiitem.png  ← Part III multi-item screenshot
│   └── boe_output_excel.png     ← Excel output screenshot
├── requirements.txt
└── README.md
```

---

## Key technical decisions

**Why pdfplumber over PyPDF2 or Camelot?**
BOE PDFs contain mixed text-and-table layouts across 6 pages. Pdfplumber gives cell-level coordinate access and handles multi-column table structures that Camelot misreads and PyPDF2 cannot parse.

**Why rule-based extraction, not ML?**
ICEGATE BOE PDFs follow a nationally standardised format. Keyword-anchored positional extraction — find `5.IGST`, extract the value at row+N, col+C — is more accurate, faster, and fully auditable compared to general-purpose ML models. Every extracted field is explainable to a finance team.

**Why one row per line item?**
Each BOE can have multiple imported goods, each with its own CTH, duty calculation, and assessable value. Flattening to one row per item makes the output directly usable in Excel pivot tables and VLOOKUP-based reconciliation workflows.

**Why 90 fields?**
This covers the complete AP reconciliation workflow: duty payment (BCD/SWS/IGST amounts), ITC claims (IGST Notn No/SNo/Rate), vendor matching (Supplier Name/Address), logistics linking (MAWB/HAWB/IGM No), and audit trail (OOC No/Date, Exchange Rate).

---

## About this project

Simplified, anonymised recreation of automation built and deployed in a production e-commerce environment. All sample data is synthetically generated or publicly available — no real business or vendor data is included.

---

## Author

**Cadence Poppen** — Senior Data Analyst | Ex-Flipkart & Myntra

13 years of enterprise data experience in Indian e-commerce. Built and deployed automation pipelines trusted by CFO-level stakeholders.

📧 cadencepoppen@gmail.com
🔗 [LinkedIn](https://linkedin.com/in/cadencepoppen) · [GitHub](https://github.com/CadencePython) · [Fiverr](https://www.fiverr.com)
