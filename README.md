# Excel PDF Missing File Checker (Logistics Automation)

## 📌 Overview
This Python automation script validates the presence of **Invoice and Waybill PDF files** based on data stored in an Excel sheet.  
It is designed for **logistics and courier operations** where missing documents can delay dispatch or filing.

The script reads invoice and waybill numbers from Excel, checks the corresponding PDF folder, and reports any missing files back into Excel.

---

## 🚀 Features
- Reads Invoice and Waybill numbers directly from Excel
- Cleans numeric values automatically (removes `.0` issue)
- Checks for missing PDF files in a specified folder
- Updates progress percentage live inside Excel
- Writes missing filenames back into Excel for easy review
- Displays a real-time progress bar in the terminal

---

## 🛠️ Technologies Used
- Python
- Windows COM Automation
- Microsoft Excel
- pywin32

---

## 📂 Excel Structure Used
| Column | Purpose |
|------|--------|
| A | Invoice Numbers |
| B | Waybill Numbers |
| J3 | Progress Indicator |
| S3 | Status Message |
| T5 ↓ | Missing File List |

---

## ⚙️ Configuration
Update these paths inside the script before running:

```python
EXCEL_PATH = r"D:\Automation\config.xlsx"
SHEET_NAME = "Filing"
PDF_FOLDER_PATH = r"D:\Automation\pdfs"
```
---

## ▶️ How to Run

#### 1️⃣ Install Dependencies
```python
pip install -r requirements.txt
```

#### 2️⃣ Run the Script
```python
python check_missing_pdfs.py
```

---

## ⚠️ Requirements
- Windows operating system.
- Microsoft Excel installed.
- Excel file should be closed before execution (recommended).
