# merge_excels.py

Merge every Excel workbook in a folder into **one tidy `xlsx` file**—all from the command line.

> **Quick-start**  
> ```bash
> python merge_excels.py /path/to/folder
> ```  
> Creates **`merged.xlsx`** in the current directory.

---

## Features
- **Folder-wide merge** – automatically detects every workbook that matches a glob pattern.  
- **Sheet selection** – pull from the first sheet (`0`) or any sheet name.  
- **Source tracing** – inserts a `__source_file` column so you always know where each row came from.  
- **Friendly CLI** – progress messages and helpful errors.  
- **Tiny footprint** – pure *pandas* + *openpyxl* (no heavy Excel automation).

---

## Installation
Requires **Python 3.8+**.

```bash
# 1. (Optional) Create a virtual environment
python -m venv .venv
source .venv/bin/activate   # Windows: .venv\Scripts\activate

# 2. Install dependencies
pip install pandas openpyxl
