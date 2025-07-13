## 🚀 XDO Generator

Easily create your XDO executable with a single command! (WINDOWS ONLY)

### What is this?

This tool automates the creation of the `XDO_METADATA` sheet and named ranges for **Oracle BI Publisher (XDO templates)**.  
It scans your Excel template for field tags (`G101`, `G102`, etc.), reads SQL fields from text files (`G1.txt`, `G2.txt`, ...), and links them to BI Publisher fields.  
Useful when building Excel-based reports in Oracle BI.

---

### 📄 Tagging Instructions (Excel Template)

See `example/template.xlsx` for how to tag your Excel file.

---

### 📁 Input Files

- `template.xlsx` — Excel file with tag markers like `G101`, `G201`, etc.
- `G1.txt`, `G2.txt`, ... — text files containing SQL models fields (comma-separated, code between `SELECT` and `FROM`)

---

### ✅ What It Does

- Creates or updates the `XDO_METADATA` sheet
- Maps each tag to a BI Publisher field:
  ```
  G101 → XDO_?XDOFIELD101? → <?FIELD_NAME?>
  ```
- Adds group loops:
  ```
  XDO_GROUP_?XDOG1? → <xsl:for-each select=".//G_1">
  ```
- Adds named ranges for each field and group
- Converts the workbook to `.xls` for BI Publisher compatibility

---

### 📦 Build the Executable

```sh
pyinstaller --onefile --name xdo_generator src/main.py
```

The compiled `xdo_generator.exe` will be placed in the `dist/` folder.

---

### ▶️ Usage

1. Place `xdo_generator.exe` in your project root
2. Add SQL files: `G1.txt`, `G2.txt`, etc.
3. Rename your Excel file to `template.xlsx`
4. Run:
   ```sh
   ./xdo_generator.exe
   ```

> Or run directly via Python:
> ```sh
> python src/main.py
> ```

---

### 📦 Requirements

- Python 3.x  
- Windows OS with Microsoft Excel (for `.xls` export)
- Install dependencies:
  ```sh
  pip install openpyxl pywin32
  ```

---
