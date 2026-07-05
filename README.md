# 📄 Professional PDF Table Extractor

A lightweight Windows desktop application that intelligently extracts tables from PDF files and converts them into clean, professionally formatted Excel (`.xlsx`) workbooks.

Designed for users who frequently work with reports, statements, invoices, and tabular PDF documents, the application provides an easy-to-use interface while producing high-quality Excel outputs with minimal manual cleanup.

---

## Features

### Intelligent Table Extraction

* Automatically extracts structured tables from PDF documents.
* Uses a dual-strategy extraction approach for maximum accuracy:

  * **Grid-based extraction (`lines`)** for bordered tables.
  * **Text-based extraction (`text`)** as a fallback for borderless tables.

### Smart Multi-page Processing

* Detects and removes repeated table headers across multiple pages.
* Merges all extracted pages into a single, clean dataset.

### Automatic Output Location

* Saves the generated Excel file directly into the same folder as the selected PDF.
* Eliminates the need to manually choose an output directory.

### Professional Excel Formatting

The exported Excel file is automatically formatted with:

* Auto-adjusted column widths
* Professional dark blue header styling
* Frozen header row
* Auto-filter enabled
* Clean worksheet formatting for immediate use

### Responsive User Interface

* Background threading prevents the application from freezing during extraction.
* Suitable for processing large PDF files.

### Real-time Progress Tracking

* Progress bar with percentage updates.
* Status messages indicating the current processing stage.

---

## Built With

* **Python 3.10+**
* **CustomTkinter** – Modern desktop GUI
* **pdfplumber** – PDF table extraction
* **Pandas** – Data processing and cleaning
* **openpyxl** – Professional Excel formatting
* **threading** – Background processing for a responsive UI

---

## Installation

### Clone the repository

```bash
git clone https://github.com/yourusername/pdf-table-extractor.git
cd pdf-table-extractor
```

### Install dependencies

```bash
pip install customtkinter pdfplumber pandas openpyxl pyinstaller
```

---

## Run the Application

```bash
python app.py
```

---

## Build a Standalone Windows Executable

Create a single executable file that can be shared with users who do not have Python installed.

```bash
pyinstaller --noconsole --onefile app.py
```

The generated executable will be located in the `dist` folder.

---

## 📂 Output

After processing, the application generates an Excel file:

* In the **same directory** as the selected PDF
* Using the original filename with an `.xlsx` extension
* Fully formatted and ready for analysis or reporting

---

## Ideal Use Cases

* Financial reports
* Bank statements
* Business reports
* Sales reports
* Transaction records
* Invoices
* Any structured PDF table

---

## 📸 Screenshot

![Converter](https://raw.githubusercontent.com/kokoaung-da/pdf-to-excel-converter/main/pdf_to_excel_converter.png)

---

## Contributing

Contributions, feature requests, and bug reports are welcome.

If you have suggestions for improvements, feel free to open an issue or submit a pull request.

---

## ⭐ Support

If you find this project useful, consider giving it a **⭐ Star** on GitHub to support future development.
