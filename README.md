# Book Catalog Checker (BGU)

Checks a list of book requests against the BGU library Primo catalog and writes a yes/no + permalink report. Input can be an **Excel** sheet or a **Word** document.

## For end users

Double-click `BookCatalogChecker.exe`. Pick the input file (`.xlsx`, `.docx`, or `.doc`). Pick where to save the output. Wait. Open the output.

### Excel input (`.xlsx`)

1 header row, then one book per row:

| Column | Meaning |
|---|---|
| A | Professor name |
| C | Book title |
| D | Author |
| G | Publisher |

### Word input (`.docx` / `.doc`)

Free-text citation lists, one book per paragraph (e.g. `Author. Title. Publisher, Year` or a numbered `1. Author. (Year). Title.`). The reader automatically:

- skips professor/section heading lines and short status/price notes (e.g. `Lib online`, `Print 26.88`);
- strips list numbering and a trailing publisher/year so the title searches cleanly;
- uses an **ISBN** for an exact catalog lookup whenever one appears in the text.

The "who requested" column is filled with the **file name** for Word input. Reading `.doc` (the legacy binary format) requires **Microsoft Word installed** on the machine; if it isn't, save the file as `.docx` first.

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
