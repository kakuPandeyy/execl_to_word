# Excel to Word

Fill a Word template from an Excel sheet. Every row in the sheet produces one finished `.docx` file, with each column's value placed exactly where you put its placeholder in the template.

Runs fully offline on your own computer. No accounts, no internet, no paid services.

## What you need

- Windows, Mac or Linux
- Python 3.9 or newer (from python.org; on Windows tick **Add Python to PATH** during install)

## Install (once)

```
pip install pandas openpyxl docxtpl
```

## Files in this folder

| File | Purpose |
|---|---|
| `excel_to_word.py` | The program |
| `run.bat` | Windows launcher. Double-click to open the window |
| `sample_data.xlsx` | Example Excel file (Name, Age, Status, Date) |
| `sample_template.docx` | Example Word template with placeholders |
| `README.md` | This file |

## Quick start

1. Double-click `run.bat` (or run `python excel_to_word.py`).
2. Choose the **Excel file**, the **Word template** and the **Output folder**.
3. Click **Generate**.
4. Open the output folder. You get one `.docx` per row.

Try it first with `sample_data.xlsx` and `sample_template.docx`. It should create three documents.

## Making your own template

1. Open your Word document.
2. Wherever a value should appear, type a placeholder in double curly braces, like `{{ name }}`.
3. Save it as a `.docx`.

Placeholder names come from your Excel column headers, in lowercase, with spaces replaced by underscores:

| Excel header | Placeholder |
|---|---|
| `Name` | `{{ name }}` |
| `Age` | `{{ age }}` |
| `Status` | `{{ status }}` |
| `Join Date` | `{{ join_date }}` |

Tips:
- Placeholders can sit in paragraphs, tables, headers and footers. Fonts, colours and layout are kept.
- Type each placeholder in one go, without switching formatting in the middle of it. If Word splits it into pieces, it may not be replaced. If that happens, delete it and retype it, or paste it as plain text.
- The first row of the Excel sheet must be the column headers.

## Output

- Each file is named after the `Name` column, for example `Ravi Kumar.docx`.
- If there is no `Name` column, files are called `row_1.docx`, `row_2.docx` and so on.
- Two rows with the same name never overwrite each other. The second becomes `Ravi Kumar_2.docx`.
- Characters Windows doesn't allow in file names are removed.

## How values are formatted

- Empty cells become blank (not `nan`).
- Whole numbers show without decimals (`28`, not `28.0`).
- Dates show as `dd-mm-yyyy`. To change this, edit `DATE_FORMAT` near the top of `excel_to_word.py`. For example, `"%d/%m/%Y"` or `"%d %B %Y"`.
- Fully empty rows in the sheet are skipped.

## Running without the window

```
python excel_to_word.py data.xlsx template.docx output_folder
```

## Common problems

| Problem | Fix |
|---|---|
| `'python' is not recognized` | Python is not on PATH. Reinstall it and tick **Add Python to PATH** |
| `No module named 'docxtpl'` | Run the `pip install` command above again |
| Values are blank in the output | The placeholder doesn't match a header. Check the WARNING in the log, which lists the columns it found |
| `{{ name }}` is still visible in the result | Word split the placeholder into parts. Retype it in one go |
| Permission error | The Excel or Word file is open. Close it and try again |
| Only the first sheet is read | By default the first sheet is used. Move your data to the first sheet |

## How it works

```
Excel (.xlsx) -> pandas reads rows -> headers cleaned (lowercase_with_underscores)
              -> for each row: load template -> replace {{ placeholders }} -> save .docx
```

Built with `pandas` and `openpyxl` (reading Excel), `docxtpl` (filling Word) and `tkinter` (the window).

## Ideas for later

- Repeating rows inside a Word table (docxtpl `{% for %}` loops)
- Images or signatures per row (docxtpl `InlineImage`)
- Combine all rows into one document (`docxcompose`)
- Export to PDF (LibreOffice headless or `docx2pdf`)
- A browser-based version using Streamlit
