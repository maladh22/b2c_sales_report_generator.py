# b2c_sales_report_generator.py
# B2C Sales Report Generator 📊

This Python script automates the creation of a **B2C monthly sales and refund report**. It processes a raw CSV file (named like `Dec 2024.csv`) exported from an e-commerce dashboard or ERP system, filters for transaction types, summarizes the data by SKU and location, and exports a well-formatted Excel report.

---

## 🔧 Features

- Extracts reporting period from the file name (e.g., "Dec 2024.csv")
- Processes **Shipment** and **Refund** transactions separately
- Groups data by **SKU**, **state**, **fulfillment channel**, etc.
- Automatically converts numeric columns and handles missing values
- Exports to a clean **multi-sheet Excel report** with tables and styling

---

## 📂 Input

- CSV file named in `Mon YYYY.csv` format (e.g., `Jan 2024.csv`)
- Expected columns:
  - `Transaction Type`, `Seller Gstin`, `Ship To State`, `Sku`, `Fulfillment Channel`
  - `Quantity`, `Tax Exclusive Gross`, `Cgst Tax`, `Sgst Tax`, `Igst Tax`, `Invoice Amount`

---

## 📄 Output

- `B2C_Sales_Report_Output.xlsx` with:
  - 📄 `Shipment` sheet (AFN & MFN sorted within states)
  - 📄 `Refund` sheet

---

## 🧠 Skills Demonstrated

- `pandas` for data preprocessing & aggregation
- `xlsxwriter` for Excel table formatting
- Automation of real-world reporting tasks
- Clean and modular Python code structure

---

## 🚀 How to Use

```bash
python b2c_sales_report_generator.py
