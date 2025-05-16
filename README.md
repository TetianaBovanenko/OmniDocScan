# OmniPage Doc Tag Scraping

This project is a Python-based pipeline for OCR-powered tag extraction and document mapping using the OmniPage SDK, XML parsing, and Excel/SQL reporting.

🧠 The OCR engine is powered by the OmniPage SDK from **Tungsten Automation** (formerly Kofax), a leading provider of document capture and automation technologies.
![OmniPage Cover](omnipage_cover.svg)
---

## 📦 Installation & Setup

### ✅ 1. OmniPage SDK Installation (Required Only Once)

⚠️ You must have a valid OmniPage license to use this SDK.

- Install `OmniPage.exe`
- Activate license using:
  C:\Program Files\OmniPage\CSDK22\Bin\OPLicMgrUI.exe
- In your virtual environment:
  py -3.12 -m pip install omnipage-22.2-cp312-cp312-win_amd64.whl

  > Make sure the Python version matches the `.whl`

---

## 🧰 Project Structure

OmniPage/
├── .venv/                      # Virtual environment (excluded via .gitignore)
├── data/                       # Data and config files
│   ├── Tags.xlsx
│   ├── Docs.xlsx
│   ├── Doc-Tag.xlsx
│   ├── ENS_Syntax_Rosebank.txt
│   └── settings_new.sts
├── src/
│   ├── pdf_to_xml.py
│   ├── pdfprocessor.py
│   ├── SampleUtils.py
│   └── doctag_scraping.py
├── .env.example
├── .gitignore
├── requirements.txt
└── README.md

---

## ⚙️ Configuration

Update your `.env` file with project-specific paths:

BASE_FOLDER=data/test
ENS_SYNTAX_FILE=data/ENS_Syntax_Rosebank.txt
TAGS_PATH=data/Tags.xlsx
DOCS_PATH=data/Docs.xlsx
DOC_TAG_PATH=data/Doc-Tag.xlsx

✅ Copy `.env.example` to `.env` and modify values as needed.

---

## 🧠 How It Works

1. Converts PDFs to XML using OmniPage OCR.
2. Extracts tags from XML using regex-based ENS patterns.
3. Filters out tags listed in Docs.xlsx.
4. Matches tags with Doc-Tag.xlsx to retrieve actions.
5. Assigns statuses using Tags.xlsx.
6. Appends missing tags from Doc-Tag.xlsx with green highlight.
7. Exports results to Excel with hyperlinks and page references.

---

## 🚀 Run Sequence

# 1. Activate virtual environment
source .venv/bin/activate  # macOS/Linux
# OR
.venv\Scripts\activate     # Windows

# 2. Convert PDFs to XML (optional)
python src/pdf_to_xml.py

# 3. Run tag scraping
python src/doctag_scraping.py

> For concurrent conversion:
python src/pdftoxml_concurrent.py

---

## 📁 Output

Excel reports are saved in the same folder as the XMLs, named:
{folder}-Doc-Tag-Scraping.xlsx

Each row includes:
- Tag No
- DocumentNo (Hyperlinked to the original PDF)
- Page
- Action
- Status

Missing Doc-Tag mappings are highlighted in green.

---

## 🧪 Example ENS Tag Formats Supported

- 29-ET-902A/B/C → 29-ET-902A, 29-ET-902B, 29-ET-902C
- 26-KA-001A/B → 26-KA-001A, 26-KA-001B
- Supports slash expansions, suffixes, and ENS-style pattern filtering

---

## 🛠 Requirements

Install dependencies via pip:

pip install -r requirements.txt

requirements.txt includes:

openpyxl>=3.1.2
regex>=2023.12.25

Optional additions (for DB or `.env` support):

python-dotenv>=1.0.1
pyodbc>=5.0.1
sqlalchemy>=2.0.30

---

## 🚀 Enhancement Idea

Consider replacing Excel lookups with a live SQL Server database connection for:
- Centralized reference data
- Real-time updates
- Easier collaboration and automation

Implementation notes:
- Use sqlalchemy or pyodbc
- Replace load_workbook() calls with SELECT queries
- Store credentials securely in .env

---

For questions, improvements, or bug reports, feel free to open an issue.

