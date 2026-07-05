# Professional PDF Table Extractor to Excel

**PDF Table Extractor** is a lightweight Windows desktop GUI application that intelligently extracts structural tables from PDF files and converts them into clean, beautifully formatted, professional Excel (`.xlsx`) files.

---

## Key Features

*   **Dual-Strategy Table Extraction:** Automatically falls back from grid line-based extraction (`lines`) to borderless text alignment extraction (`text`) to ensure maximum data capture.
*   **Smart Header Handling:** Intelligently detects and eliminates repetitive headers across multiple pages, consolidating data into a single clean table.
*   **Automated Save Location:** Enhances UX by automatically saving the output Excel file into the exact same folder as the source PDF—no folder picking required.
*   **Professional Excel Formatting:** Automatically adjusts column widths based on content length, applies a professional dark blue theme to headers, freezes the top row, and enables auto-filters.
*   **Asynchronous Threading:** Runs extraction processes on a background thread, preventing the GUI window from becoming "Not Responding" or freezing on large PDF files.
*   **Real-time Progress Bar:** Provides clear visual feedback with a progress bar and percentage status as it processes through PDF pages.

---

## Built With

*   **Python 3.10+**
*   **CustomTkinter** - For the modern and clean UI/UX design.
*   **pdfplumber** - For precise structural table extraction.
*   **Pandas** - For efficient data cleaning, structural alignment, and processing.
*   **openpyxl** - For programmatic, production-grade Excel styling.

---

## Installation & Setup

### 1. Clone the Repository
```bash
git clone [https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git](https://github.com/YOUR_USERNAME/YOUR_REPOSITORY_NAME.git)
cd YOUR_REPOSITORY_NAME
