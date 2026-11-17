"""
Excel -> Word template filler
------------------------------
Fills a Word template (with {{ placeholders }}) once for every row of an Excel sheet.

Template rules:
  * Put {{ column_name }} in the Word file where each value should go.
  * Column names are matched in lowercase, and spaces become underscores:
      Excel header "Name"       -> {{ name }}
      Excel header "Join Date"  -> {{ join_date }}

Run:
  python excel_to_word.py              (opens the window)
  python excel_to_word.py data.xlsx template.docx output_folder   (no window)

Install once:
  pip install pandas openpyxl docxtpl
"""
import re
import sys
from pathlib import Path

import pandas as pd
from docxtpl import DocxTemplate

DATE_FORMAT = "%d-%m-%Y"   # change to "%d/%m/%Y" or "%d %B %Y" if you prefer


def clean_key(header):
    """'Join Date ' -> 'join_date'"""
    return re.sub(r"\W+", "_", str(header).strip().lower()).strip("_")


def format_value(v):
    if pd.isna(v):
        return ""
    if isinstance(v, pd.Timestamp):
        return v.strftime(DATE_FORMAT)
    if isinstance(v, float) and v.is_integer():
        return str(int(v))          # 28.0 -> "28"
    return str(v)


def safe_filename(text):
    text = re.sub(r'[\\/:*?"<>|]+', "", str(text)).strip()
    return text[:80] or "row"


def generate(excel_path, template_path, out_dir, sheet=0, name_column=None, log=print):
    excel_path, template_path, out_dir = Path(excel_path), Path(template_path), Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    df = pd.read_excel(excel_path, sheet_name=sheet)
    df = df.dropna(how="all")                      # skip completely empty rows
    df.columns = [clean_key(c) for c in df.columns]

    # Check which placeholders in the template have no matching column
    placeholders = DocxTemplate(template_path).get_undeclared_template_variables()
    missing = sorted(placeholders - set(df.columns))
    if missing:
        log(f"WARNING: these placeholders have no matching Excel column and will be blank: {missing}")
        log(f"         Excel columns found: {list(df.columns)}")

    name_column = name_column or ("name" if "name" in df.columns else None)
    used, created = set(), []

    for i, (_, row) in enumerate(df.iterrows(), start=1):
        context = {k: format_value(v) for k, v in row.items()}

        base = safe_filename(context.get(name_column)) if name_column else f"row_{i}"
        fname, n = base, 2
        while fname.lower() in used:               # avoid overwriting duplicates
            fname, n = f"{base}_{n}", n + 1
        used.add(fname.lower())

        doc = DocxTemplate(template_path)
        doc.render(context)
        out = out_dir / f"{fname}.docx"
        doc.save(out)
        created.append(out)
        log(f"[{i}/{len(df)}] {out.name}")

    log(f"Done. {len(created)} file(s) saved in: {out_dir.resolve()}")
    return created


# ---------------------------------------------------------------- window ----
def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, scrolledtext

    root = tk.Tk()
    root.title("Excel to Word")
    root.geometry("640x420")

    excel_var, tpl_var, out_var = tk.StringVar(), tk.StringVar(), tk.StringVar()

    def pick(var, kind):
        if kind == "excel":
            p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        elif kind == "word":
            p = filedialog.askopenfilename(filetypes=[("Word", "*.docx")])
        else:
            p = filedialog.askdirectory()
        if p:
            var.set(p)
            if kind == "excel" and not out_var.get():
                out_var.set(str(Path(p).parent / "output"))

    for r, (label, var, kind) in enumerate([
        ("Excel file", excel_var, "excel"),
        ("Word template", tpl_var, "word"),
        ("Output folder", out_var, "folder"),
    ]):
        tk.Label(root, text=label, width=14, anchor="w").grid(row=r, column=0, padx=8, pady=6)
        tk.Entry(root, textvariable=var, width=55).grid(row=r, column=1, pady=6)
        tk.Button(root, text="Browse", command=lambda v=var, k=kind: pick(v, k)).grid(row=r, column=2, padx=8)

    box = scrolledtext.ScrolledText(root, height=14, width=76)
    box.grid(row=4, column=0, columnspan=3, padx=8, pady=8)

    def log(msg):
        box.insert(tk.END, msg + "\n"); box.see(tk.END); root.update()

    def go():
        if not (excel_var.get() and tpl_var.get() and out_var.get()):
            messagebox.showwarning("Missing", "Please choose the Excel file, template and output folder.")
            return
        box.delete("1.0", tk.END)
        try:
            files = generate(excel_var.get(), tpl_var.get(), out_var.get(), log=log)
            messagebox.showinfo("Finished", f"{len(files)} document(s) created.")
        except Exception as e:
            log(f"ERROR: {e}")
            messagebox.showerror("Error", str(e))

    tk.Button(root, text="Generate", width=20, bg="#2b6cb0", fg="white", command=go)\
        .grid(row=3, column=1, pady=6)
    root.mainloop()


if __name__ == "__main__":
    if len(sys.argv) == 4:
        generate(*sys.argv[1:])
    else:
        run_gui()
