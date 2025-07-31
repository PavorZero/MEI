# Tools for Processing and Analyzing Name Lists
This document describes Python scripts included in a repository aimed at cleaning, normalizing, comparing, and analyzing name lists. These tools are especially useful for event data, registration forms, and other text-based or spreadsheet sources.

📁 Contents

- limpol.py: Cleans Excel spreadsheets, removing line breaks, tabs, and duplicate spaces.
- normalize.py: Normalizes names in text lists (.txt) and exports to Excel.
- genders.py: Estimates gender based on first name (suffix pattern analysis).
- comparemei.py: Compares lists to identify matching names.
- distintctmei.py: Compares lists to identify different or missing names.

🔧 Requirements
Python 3.8+ and the following libraries:
- pandas
- openpyxl
- rapidfuzz

🧹 1. Excel Sheet Cleaner (limpol.py)
Cleans Excel spreadsheets (.xlsx) by removing line breaks and unnecessary spaces.
- python limpol.py
- input_excel_path = "yourfile.xlsx"
- aba_para_corrigir = "*"  # or sheet name
- coluna_para_corrigir = "*"  # or column name

🧽 2. Name Normalizer (normalize.py)
Reads a .txt file with names and generates a normalized .xlsx version.
- python normalizar.py
- arquivo_entrada = 'names.txt'
- arquivo_saida = 'normalized_names.xlsx'

🚻 3. Gender Identifier (genders.py)
Attempts to identify gender based on the first name using suffix analysis.
- python "genders.py"

🔁 4. Name List Comparator - Matching Names (comparemei.py)
Compares multiple name lists to find names that appear in more than one list, even with variations.
- python comparemei.py

🚫 5. Name List Comparator - Distinct Names (distintctmei.py)
Focuses on identifying names that are different between lists (non-matching).
- python distintctmei.py

📝 License
This project is licensed under the MIT License.

👤 Author
João Carlos

