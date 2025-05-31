import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
import difflib
import re
import openpyxl

class ExcelFunctionApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel関数支援アプリ 拡張版")

        self.df = None
        self.file_path = ""
        self.header_row = None

        # GUI要素
        tk.Button(root, text="📁 Excelファイルを選択\n（ここを押してファイルを選んでね）", command=self.load_file).pack(pady=5)

        tk.Label(root, text="▼ ヘッダー行の番号を入力してください\n（タイトルの行番号だよ。1番上の行なら「0」、2行目なら「1」）").pack()
        self.header_entry = tk.Entry(root)
        self.header_entry.pack()

        tk.Button(root, text="👁 データをプレビュー\n（ファイルの中身を少し見るボタン）", command=self.preview_data).pack(pady=5)

        self.tree = ttk.Treeview(root)
        self.tree.pack(pady=5)

        tk.Label(root, text="▼ どんなことをしたいか日本語で書いてね\n（例：「売上が100以上かつ来店回数が5以上の人を調べたい」）").pack()
        self.prompt_entry = tk.Entry(root, width=80)
        self.prompt_entry.pack()

        tk.Label(root, text="▼ 関数を書き込みたい場所を入力\n（例：「C2」みたいに、Excelのセルの場所を書くよ）").pack()
        self.cell_entry = tk.Entry(root, width=10)
        self.cell_entry.pack()

        tk.Label(root, text="▼ 関数を書き込む行の数を数字で入力してね\n（全部なら空白のままでOK！）").pack()
        self.rowcount_entry = tk.Entry(root, width=10)
        self.rowcount_entry.pack()

        tk.Button(root, text="🧮 日本語を関数に変換するよ", command=self.convert_to_formula).pack(pady=5)

        tk.Label(root, text="▼ できた関数がここに表示されるよ\n（ちゃんと見てね！）").pack()
        self.result_text = tk.Text(root, height=4, width=80)
        self.result_text.pack()

        tk.Button(root, text="💾 関数をExcelに書き込んで保存するよ", command=self.save_formula_to_excel).pack(pady=5)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return
        self.file_path = file_path
        try:
            self.df_raw = pd.read_excel(file_path, header=None)
            messagebox.showinfo("成功", "ファイルを読み込みました！")
        except Exception as e:
            messagebox.showerror("エラー", f"読み込み失敗: {e}")

    def preview_data(self):
        try:
            row = int(self.header_entry.get())
            self.header_row = row
            self.df = pd.read_excel(self.file_path, header=row)
            self.tree.delete(*self.tree.get_children())
            self.tree["columns"] = list(self.df.columns)
            self.tree["show"] = "headings"
            for col in self.df.columns:
                self.tree.heading(col, text=col)
            for index, row in self.df.head(5).iterrows():
                self.tree.insert("", "end", values=list(row))
        except Exception as e:
            messagebox.showerror("エラー", f"プレビュー失敗: {e}")

    def find_similar_column(self, word):
        if self.df is None:
            return None
        cols = list(self.df.columns)
        synonyms = {
            "名前": ["氏名", "名前", "顧客名"],
            "売上": ["売上", "売上高", "収入"],
            "来店回数": ["来店回数", "訪問回数", "回数"],
        }
        for key, syns in synonyms.items():
            if word in syns:
                word = key
                break

        matches = difflib.get_close_matches(word, cols, n=1, cutoff=0.6)
        if matches:
            return matches[0]
        return None

    def parse_conditions(self, prompt):
        and_parts = re.split(r"かつ|且つ|そして", prompt)
        or_parts = []
        for part in and_parts:
            or_split = re.split(r"または|もしくは", part)
            or_parts.append(or_split)
        return and_parts, or_parts

    def generate_formula(self, prompt):
        if self.df is None:
            return "Excelファイルを読み込み、ヘッダー行を設定してください。"

        cond_pattern = re.compile(r"(\w+?)が(\d+)以上")
        conds = cond_pattern.findall(prompt)

        conditions = []
        for colname, val in conds:
            col = self.find_similar_column(colname)
            if not col:
                return f"列「{colname}」が見つかりません。Excelのヘッダーを確認してください。"
            conditions.append(f'{col}>={val}')

        if "または" in prompt or "もしくは" in prompt:
            formula = "=OR(" + ", ".join(conditions) + ")"
        else:
            formula = "=AND(" + ", ".join(conditions) + ")"

        return formula if conditions else "条件が認識できませんでした。"

    def convert_to_formula(self):
        prompt = self.prompt_entry.get().strip()
        start_cell = self.cell_entry.get().strip().upper()
        row_count = self.rowcount_entry.get().strip()

        if not prompt or self.df is None:
            messagebox.showwarning("入力エラー", "日本語でやりたいことを入力してください。")
            return
        if not re.match(r"^[A-Z]+\d+$", start_cell):
            messagebox.showwarning("入力エラー", "開始セルは例のように入力してください（例：C2）。")
            return
        if row_count and not row_count.isdigit():
            messagebox.showwarning("入力エラー", "書き込み行数は数字で入力してください。")
            return

        formula = self.generate_formula(prompt)
        self.result_text.delete("1.0", tk.END)
        self.result_text.insert(tk.END, formula)
        self.generated_formula = formula
        self.start_cell = start_cell
        self.row_count = int(row_count) if row_count else len(self.df)

    def save_formula_to_excel(self):
        if not hasattr(self, "generated_formula") or not hasattr(self, "start_cell"):
            messagebox.showwarning("保存できません", "関数を生成し開始セルを指定してください。")
            return
        try:
            wb = openpyxl.load_workbook(self.file_path)
            ws = wb.active

            col_letters = re.findall(r"[A-Z]+", self.start_cell)[0]
            row_number = int(re.findall(r"\d+", self.start_cell)[0])
            col_number = 0
            for i, c in enumerate(reversed(col_letters)):
                col_number += (ord(c) - ord("A") + 1) * (26 ** i)

            for i in range(self.row_count):
                cell = ws.cell(row=row_number + i, column=col_number)
                cell.value = self.generated_formula

            new_path = self.file_path.replace(".xlsx", "_with_formula.xlsx")
            wb.save(new_path)
            messagebox.showinfo("保存完了", f"Excelファイルに書き込み保存しました。\n{new_path}")
        except Exception as e:
            messagebox.showerror("保存失敗", str(e))


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelFunctionApp(root)
    root.mainloop()
