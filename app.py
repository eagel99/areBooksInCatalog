import os
import sys
import threading
import traceback
from pathlib import Path
from tkinter import Tk, StringVar, filedialog, messagebox, ttk, scrolledtext, END, DISABLED, NORMAL

from catalog import is_available, make_session, permalink, pick_best, search
from xlsx_io import Result, read_rows, write_results


class App:
    def __init__(self, root: Tk) -> None:
        self.root = root
        root.title("בדיקת ספרים בקטלוג BGU")
        root.geometry("680x440")
        root.minsize(680, 440)

        self.input_path = StringVar()
        self.output_path: str | None = None
        self.running = False

        frm = ttk.Frame(root, padding=12)
        frm.pack(fill="both", expand=True)

        ttk.Label(frm, text="קובץ הקלט (xlsx):").grid(row=0, column=0, sticky="w")
        ttk.Entry(frm, textvariable=self.input_path).grid(row=1, column=0, sticky="we", padx=(0, 8))
        self.browse_btn = ttk.Button(frm, text="עיון...", command=self.browse)
        self.browse_btn.grid(row=1, column=1)

        self.run_btn = ttk.Button(frm, text="הרץ בדיקה", command=self.start_run, state=DISABLED)
        self.run_btn.grid(row=2, column=0, columnspan=2, sticky="we", pady=(12, 6))

        self.progress = ttk.Progressbar(frm, mode="determinate")
        self.progress.grid(row=3, column=0, columnspan=2, sticky="we", pady=(6, 6))

        self.status = StringVar(value="מוכן")
        ttk.Label(frm, textvariable=self.status).grid(row=4, column=0, columnspan=2, sticky="w")

        self.log = scrolledtext.ScrolledText(frm, height=12, wrap="word")
        self.log.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(8, 6))

        self.open_btn = ttk.Button(frm, text="פתח קובץ פלט", command=self.open_output, state=DISABLED)
        self.open_btn.grid(row=6, column=0, columnspan=2, sticky="we")

        frm.columnconfigure(0, weight=1)
        frm.rowconfigure(5, weight=1)

        self.input_path.trace_add("write", self._on_input_change)

    def _on_input_change(self, *_: object) -> None:
        if self.input_path.get() and not self.running:
            self.run_btn.configure(state=NORMAL)
        else:
            self.run_btn.configure(state=DISABLED)

    def browse(self) -> None:
        path = filedialog.askopenfilename(
            title="בחר קובץ קלט",
            filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.input_path.set(path)

    def log_line(self, msg: str) -> None:
        self.log.insert(END, msg + "\n")
        self.log.see(END)

    def start_run(self) -> None:
        in_path = self.input_path.get().strip()
        if not in_path or not Path(in_path).is_file():
            messagebox.showerror("שגיאה", "קובץ קלט לא נמצא")
            return

        out_path = filedialog.asksaveasfilename(
            title="שמור קובץ פלט",
            defaultextension=".xlsx",
            initialfile=Path(in_path).stem + "_catalog_check.xlsx",
            filetypes=[("Excel", "*.xlsx")],
        )
        if not out_path:
            return

        self.running = True
        self.run_btn.configure(state=DISABLED)
        self.browse_btn.configure(state=DISABLED)
        self.open_btn.configure(state=DISABLED)
        self.log.delete("1.0", END)
        self.progress.configure(value=0)
        self.output_path = None

        t = threading.Thread(target=self._run, args=(in_path, out_path), daemon=True)
        t.start()

    def _run(self, in_path: str, out_path: str) -> None:
        try:
            rows = read_rows(in_path)
        except Exception as e:
            self._post(lambda: messagebox.showerror("שגיאה בקריאת הקובץ", str(e)))
            self._post(self._finish_failed)
            return

        if not rows:
            self._post(lambda: messagebox.showwarning("ריק", "לא נמצאו שורות עם שם ספר"))
            self._post(self._finish_failed)
            return

        n = len(rows)
        self._post(lambda: self.progress.configure(maximum=n, value=0))
        self._post(lambda: self.log_line(f"נמצאו {n} ספרים לבדיקה"))

        session = make_session()
        results: list[Result] = []
        found = 0
        not_found = 0

        for i, row in enumerate(rows, start=1):
            self._post(lambda i=i, n=n, t=row.title: self.status.set(f"מחפש ({i}/{n}): {t[:80]}"))
            try:
                docs = search(row.title, session)
                best = pick_best(row, docs)
                if best is not None:
                    exists = is_available(best)
                    link = permalink(best)
                    results.append(Result(row.professor, row.title, exists, link))
                    found += 1
                    self._post(lambda t=row.title, ex=exists: self.log_line(
                        f"  ✓ נמצא ({'כן' if ex else 'לא'}): {t[:90]}"
                    ))
                else:
                    results.append(Result(row.professor, row.title, False, ""))
                    not_found += 1
                    self._post(lambda t=row.title: self.log_line(f"  ✗ לא נמצא: {t[:90]}"))
            except Exception as e:
                results.append(Result(row.professor, row.title, False, ""))
                err = str(e)
                self._post(lambda t=row.title, err=err: self.log_line(f"  ! שגיאה ({t[:60]}): {err}"))

            self._post(lambda i=i: self.progress.configure(value=i))

        try:
            write_results(out_path, results)
        except Exception as e:
            err = str(e)
            tb = traceback.format_exc()
            self._post(lambda err=err, tb=tb: self.log_line(f"שגיאה בשמירה: {err}\n{tb}"))
            self._post(lambda err=err: messagebox.showerror("שגיאה בשמירה", err))
            self._post(self._finish_failed)
            return

        self.output_path = out_path
        self._post(lambda: self.status.set(f"הסתיים. נבדקו {n} ספרים. נמצאו {found}, לא נמצאו {not_found}."))
        self._post(lambda: self.log_line(f"\nנשמר: {out_path}"))
        self._post(self._finish_ok)

    def _post(self, fn) -> None:
        self.root.after(0, fn)

    def _finish_ok(self) -> None:
        self.running = False
        self.browse_btn.configure(state=NORMAL)
        self._on_input_change()
        self.open_btn.configure(state=NORMAL)

    def _finish_failed(self) -> None:
        self.running = False
        self.browse_btn.configure(state=NORMAL)
        self._on_input_change()
        self.status.set("נכשל")

    def open_output(self) -> None:
        if self.output_path and Path(self.output_path).is_file():
            os.startfile(self.output_path)


def main() -> None:
    root = Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
