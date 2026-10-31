import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import sqlite3
import sys
from pathlib import Path
from openpyxl import load_workbook

APP_NAME = "床號電話快速查詢系統"

# 資料庫放在 EXE 同一個資料夾
if getattr(sys, "frozen", False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent

DB_PATH = BASE_DIR / "bed_phone.db"


# =========================
# 資料庫初始化
# =========================

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS beds (
            bed_number TEXT PRIMARY KEY COLLATE NOCASE,
            call_number TEXT DEFAULT '',
            desk_number TEXT DEFAULT ''
        )
    """)
    conn.commit()
    return conn


# =========================
# 主程式
# =========================

class BedPhoneApp:

    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("900x700")
        self.root.minsize(800, 600)
        self.root.state("zoomed")

        self.conn = init_db()
        self.after_id = None

        self.create_interface()

    def create_interface(self):

        # 標題
        tk.Label(
            self.root,
            text="床號電話快速查詢",
            font=("Microsoft JhengHei UI", 22, "bold")
        ).pack(pady=15)

        tk.Label(
            self.root,
            text="輸入床號，自動顯示 Call 機及桌機號碼",
            font=("Microsoft JhengHei UI", 11)
        ).pack()

        # 床號輸入
        frame = tk.LabelFrame(
            self.root,
            text="床號輸入",
            padx=15,
            pady=15
        )
        frame.pack(fill="x", padx=25, pady=20)

        self.bed_var = tk.StringVar()

        self.entry = tk.Entry(
            frame,
            textvariable=self.bed_var,
            font=("Arial", 23),
            justify="center"
        )
        self.entry.pack(fill="x", ipady=8)

        self.entry.bind("<KeyRelease>", self.auto_search)
        self.entry.bind("<Return>", self.search)

        # 查詢結果
        result = tk.LabelFrame(
            self.root,
            text="查詢結果",
            padx=20,
            pady=20
        )
        result.pack(fill="both", expand=True, padx=25)

        self.status = tk.Label(
            result,
            text="請輸入床號",
            font=("Microsoft JhengHei UI", 12),
            fg="gray"
        )
        self.status.pack(pady=10)

        tk.Label(
            result,
            text="CALL 機號碼",
            font=("Microsoft JhengHei UI", 13)
        ).pack()

        self.call_label = tk.Label(
            result,
            text="—",
            font=("Arial", 26, "bold"),
            fg="#16843B"
        )
        self.call_label.pack(pady=10)

        ttk.Separator(result).pack(fill="x", pady=10)

        tk.Label(
            result,
            text="桌機電話",
            font=("Microsoft JhengHei UI", 13)
        ).pack()

        self.desk_label = tk.Label(
            result,
            text="—",
            font=("Arial", 26, "bold")
        )
        self.desk_label.pack(pady=10)

        # 功能按鈕
        buttons = tk.Frame(self.root)
        buttons.pack(pady=15)

        ttk.Button(
            buttons,
            text="匯入 Excel",
            command=self.import_excel
        ).pack(side="left", padx=5)

        ttk.Button(
            buttons,
            text="資料管理",
            command=self.manage_data
        ).pack(side="left", padx=5)

        ttk.Button(
            buttons,
            text="複製電話",
            command=self.copy_phone
        ).pack(side="left", padx=5)

        ttk.Button(
            buttons,
            text="清除",
            command=self.clear
        ).pack(side="left", padx=5)

        self.entry.focus_set()

    # =========================
    # 自動查詢
    # =========================

    def auto_search(self, event=None):

        if self.after_id:
            self.root.after_cancel(self.after_id)

        self.after_id = self.root.after(200, self.search)

    def search(self, event=None):

        self.after_id = None

        bed = self.bed_var.get().strip().upper()

        if not bed:
            self.status.config(text="請輸入床號", fg="gray")
            self.call_label.config(text="—")
            self.desk_label.config(text="—")
            return

        row = self.conn.execute(
            """
            SELECT call_number, desk_number
            FROM beds
            WHERE bed_number = ? COLLATE NOCASE
            """,
            (bed,)
        ).fetchone()

        if row:

            self.status.config(
                text=f"已找到床號：{bed}",
                fg="green"
            )

            self.call_label.config(
                text=row[0] or "未設定"
            )

            self.desk_label.config(
                text=row[1] or "未設定"
            )

        else:

            self.status.config(
                text=f"查無床號：{bed}",
                fg="red"
            )

            self.call_label.config(text="查無資料")
            self.desk_label.config(text="查無資料")

    # =========================
    # Excel 匯入
    # =========================

    def import_excel(self):

        filename = filedialog.askopenfilename(
            filetypes=[
                ("Excel 檔案", "*.xlsx")
            ]
        )

        if not filename:
            return

        try:

            wb = load_workbook(
                filename,
                read_only=True,
                data_only=True
            )

            ws = wb.active
            rows = ws.iter_rows(values_only=True)

            headers = next(rows, None)

            if not headers:
                raise ValueError("Excel 沒有資料")

            # 支援欄位名稱
            headers = [
                str(x or "").strip().lower().replace(" ", "")
                for x in headers
            ]

            aliases = {
                "bed": ["床號", "床位", "bed", "bednumber"],
                "call": ["call機號碼", "call機", "call", "callnumber"],
                "desk": ["桌機電話", "桌機", "電話", "desk", "deskphone"]
            }

            indexes = {}

            for key, names in aliases.items():
                indexes[key] = next(
                    (headers.index(n) for n in names if n in headers),
                    None
                )

            if any(v is None for v in indexes.values()):
                raise ValueError(
                    "Excel 第一列欄位必須包含：床號、Call機號碼、桌機電話"
                )

            count = 0

            with self.conn:

                for row in rows:

                    bed = str(
                        row[indexes["bed"]] or ""
                    ).strip().upper()

                    if not bed:
                        continue

                    call = str(
                        row[indexes["call"]] or ""
                    ).strip()

                    desk = str(
                        row[indexes["desk"]] or ""
                    ).strip()

                    self.conn.execute(
                        """
                        INSERT INTO beds
                        (bed_number, call_number, desk_number)
                        VALUES (?, ?, ?)
                        ON CONFLICT(bed_number)
                        DO UPDATE SET
                        call_number=excluded.call_number,
                        desk_number=excluded.desk_number
                        """,
                        (bed, call, desk)
                    )

                    count += 1

            wb.close()

            self.search()

            messagebox.showinfo(
                "匯入成功",
                f"成功匯入或更新 {count} 筆資料"
            )

        except Exception as e:

            messagebox.showerror(
                "匯入失敗",
                str(e)
            )

    # =========================
    # 資料管理
    # =========================

    def manage_data(self):

        win = tk.Toplevel(self.root)
        win.title("床號資料管理")
        win.geometry("650x450")

        tree = ttk.Treeview(
            win,
            columns=("bed", "call", "desk"),
            show="headings"
        )

        tree.heading("bed", text="床號")
        tree.heading("call", text="Call機號碼")
        tree.heading("desk", text="桌機電話")

        tree.column("bed", width=130)
        tree.column("call", width=180)
        tree.column("desk", width=200)

        tree.pack(fill="both", expand=True, padx=10, pady=10)

        def refresh():
            tree.delete(*tree.get_children())

            records = self.conn.execute(
                "SELECT bed_number, call_number, desk_number "
                "FROM beds ORDER BY bed_number"
            ).fetchall()

            for record in records:
                tree.insert("", "end", values=record)

        def add_or_edit(edit=False):

            selected = tree.selection() if edit else []

            if edit and not selected:
                messagebox.showwarning("提醒", "請先選擇資料")
                return

            old = tree.item(selected[0], "values") if selected else ("", "", "")

            bed = simpledialog.askstring(
                "床號",
                "請輸入床號",
                initialvalue=old[0],
                parent=win
            )

            if bed is None:
                return

            bed = bed.strip().upper()

            if not bed:
                messagebox.showwarning("提醒", "床號不可空白")
                return

            call = simpledialog.askstring(
                "Call機號碼",
                "請輸入 Call 機號碼",
                initialvalue=old[1],
                parent=win
            )

            if call is None:
                return

            desk = simpledialog.askstring(
                "桌機電話",
                "請輸入桌機電話",
                initialvalue=old[2],
                parent=win
            )

            if desk is None:
                return

            try:
                with self.conn:
                    if edit and bed.lower() != str(old[0]).lower():
                        self.conn.execute(
                            "DELETE FROM beds WHERE bed_number=?",
                            (old[0],)
                        )

                    self.conn.execute(
                        """
                        INSERT INTO beds
                        (bed_number, call_number, desk_number)
                        VALUES (?, ?, ?)
                        ON CONFLICT(bed_number)
                        DO UPDATE SET
                        call_number=excluded.call_number,
                        desk_number=excluded.desk_number
                        """,
                        (bed, call, desk)
                    )

                refresh()
                self.search()

            except sqlite3.IntegrityError:
                messagebox.showerror("錯誤", "床號資料重複")

        def delete():
            selected = tree.selection()

            if not selected:
                messagebox.showwarning("提醒", "請先選擇資料")
                return

            bed = tree.item(selected[0], "values")[0]

            if messagebox.askyesno("確認", f"確定刪除床號 {bed}？"):
                with self.conn:
                    self.conn.execute(
                        "DELETE FROM beds WHERE bed_number=?",
                        (bed,)
                    )

                refresh()
                self.search()

        controls = tk.Frame(win)
        controls.pack(pady=8)

        ttk.Button(
            controls,
            text="新增",
            command=lambda: add_or_edit(False)
        ).pack(side="left", padx=5)

        ttk.Button(
            controls,
            text="修改",
            command=lambda: add_or_edit(True)
        ).pack(side="left", padx=5)

        ttk.Button(
            controls,
            text="刪除",
            command=delete
        ).pack(side="left", padx=5)

        refresh()

    # =========================
    # 複製電話
    # =========================

    def copy_phone(self):

        phone = self.desk_label.cget("text")

        if phone in ("—", "查無資料", "未設定"):
            messagebox.showwarning("提醒", "目前沒有可複製的桌機電話")
            return

        self.root.clipboard_clear()
        self.root.clipboard_append(phone)

        messagebox.showinfo(
            "複製完成",
            f"已複製桌機電話：{phone}"
        )

    # =========================
    # 清除
    # =========================

    def clear(self):

        self.bed_var.set("")
        self.search()
        self.entry.focus_set()


# =========================
# 啟動
# =========================

if __name__ == "__main__":

    root = tk.Tk()

    app = BedPhoneApp(root)

    def close_app():
        app.conn.close()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", close_app)

    root.mainloop()
