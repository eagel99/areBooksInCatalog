# Book Catalog Checker (BGU)

Checks an xlsx of book requests against the BGU library Primo catalog and writes a yes/no + permalink report.

## For end users

Double-click `BookCatalogChecker.exe`. Pick the input xlsx. Pick where to save the output. Wait. Open the output.

Input xlsx layout (1 header row, then one book per row):

| Column | Meaning |
|---|---|
| A | Professor name |
| C | Book title |
| D | Author |
| G | Publisher |

Output xlsx columns: `מי ביקש את הספר`, `שם הספר`, `קיים בקטלוג` (כן / לא), `Permalink`.

## For developers

Requires Python 3.11+ on Windows.

```
pip install -r requirements.txt
python app.py
```

## Building the exe

```
build.bat
```

Output: `dist\BookCatalogChecker.exe`.
