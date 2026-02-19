"""
薬剤在庫管理・発注計算アプリケーション
- 在庫ファイル(Fメニュー)と使用予定ファイルを読み込み、
  返品候補リスト・発注候補リストを算出する。
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import csv
import os
import sys
import io
from datetime import date

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import (
    SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image,
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

import barcode
from barcode.writer import ImageWriter


# ---------------------------------------------------------------------------
# ユーティリティ
# ---------------------------------------------------------------------------

def read_csv_auto_encoding(filepath, skiprows=0):
    """CSV を Shift_JIS / CP932 / UTF-8 の順に試して読み込む。
    skiprows: 先頭から読み飛ばす行数（ヘッダー行の前にあるゴミ行）。
    戻り値: (headers: list[str], rows: list[dict])
    """
    encodings = ["cp932", "shift_jis", "utf-8", "utf-8-sig"]
    raw_lines = None
    for enc in encodings:
        try:
            with open(filepath, "r", encoding=enc, newline="") as f:
                raw_lines = f.readlines()
            break
        except (UnicodeDecodeError, UnicodeError):
            continue

    if raw_lines is None:
        raise UnicodeDecodeError(
            "auto", b"", 0, 1,
            "対応するエンコーディングが見つかりませんでした。"
        )

    # skiprows 分だけ先頭を捨てる
    raw_lines = raw_lines[skiprows:]
    reader = csv.DictReader(raw_lines)
    headers = reader.fieldnames or []
    rows = list(reader)
    return headers, rows


def to_number(value):
    """カンマ区切り文字列を float に変換する。変換不能なら 0.0。"""
    if value is None:
        return 0.0
    s = str(value).replace(",", "").replace("，", "").strip()
    if s == "":
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


# ---------------------------------------------------------------------------
# PDF 用フォント登録
# ---------------------------------------------------------------------------

def _register_japanese_font():
    """Windows 環境の日本語フォントを探して登録する。"""
    candidates = [
        # Windows
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "msgothic.ttc"),
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "meiryo.ttc"),
        os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", "YuGothM.ttc"),
        # macOS (開発用)
        "/System/Library/Fonts/ヒラギノ角ゴシック W3.ttc",
        "/System/Library/Fonts/Supplemental/Osaka.ttf",
        "/System/Library/Fonts/Hiragino Sans GB.ttc",
        "/Library/Fonts/Arial Unicode.ttf",
    ]
    for path in candidates:
        if os.path.exists(path):
            try:
                pdfmetrics.registerFont(TTFont("JaFont", path))
                return "JaFont"
            except Exception:
                continue
    # フォールバック: Helvetica（日本語は文字化けする可能性あり）
    return "Helvetica"


def _generate_barcode_image(code_str, width_mm=50, height_mm=12):
    """薬剤コードから CODE128 バーコード画像 (PNG bytes → ReportLab Image) を生成。"""
    code_str = str(code_str).strip()
    if not code_str:
        return None
    try:
        CODE128 = barcode.get_barcode_class("code128")
        writer = ImageWriter()
        bc = CODE128(code_str, writer=writer)
        buf = io.BytesIO()
        bc.write(buf, options={
            "module_width": 0.25,
            "module_height": 6.0,
            "font_size": 7,
            "text_distance": 2,
            "quiet_zone": 2,
        })
        buf.seek(0)
        return Image(buf, width=width_mm * mm, height=height_mm * mm)
    except Exception:
        return None


# ---------------------------------------------------------------------------
# メイン GUI
# ---------------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("薬剤在庫管理・発注計算")
        self.geometry("1100x700")
        self.minsize(900, 550)

        # アイコン設定（exe 化後も動くようにする）
        try:
            if getattr(sys, "frozen", False):
                base = sys._MEIPASS
            else:
                base = os.path.dirname(__file__)
            ico = os.path.join(base, "app.ico")
            if os.path.exists(ico):
                self.iconbitmap(ico)
        except Exception:
            pass

        # 内部状態
        self.inventory_path = None
        self.schedule_path = None
        self.surplus_data = []  # 返品候補
        self.shortage_data = []  # 発注候補
        self.surplus_sort_col = None
        self.surplus_sort_asc = True
        self.shortage_sort_col = None
        self.shortage_sort_asc = True

        self._build_ui()

    # ---- UI 構築 ----
    def _build_ui(self):
        # スタイル
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Header.TLabel", font=("Meiryo UI", 11, "bold"))
        style.configure("Path.TLabel", font=("Meiryo UI", 9), foreground="#555")
        style.configure("Treeview", font=("Meiryo UI", 10), rowheight=24)
        style.configure("Treeview.Heading", font=("Meiryo UI", 10, "bold"))
        style.configure("Status.TLabel", font=("Meiryo UI", 9))
        style.configure("Accent.TButton", font=("Meiryo UI", 10, "bold"))

        # --- 上部: ファイル選択エリア ---
        file_frame = ttk.LabelFrame(self, text="ファイル選択", padding=10)
        file_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        # 在庫ファイル行
        row1 = ttk.Frame(file_frame)
        row1.pack(fill=tk.X, pady=2)
        ttk.Label(row1, text="在庫ファイル（Fメニュー）:", style="Header.TLabel").pack(side=tk.LEFT)
        self.inv_btn = ttk.Button(row1, text="ファイルを選択…", command=self._select_inventory)
        self.inv_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.inv_label = ttk.Label(row1, text="未選択", style="Path.TLabel")
        self.inv_label.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)

        # 使用予定ファイル行
        row2 = ttk.Frame(file_frame)
        row2.pack(fill=tk.X, pady=2)
        ttk.Label(row2, text="使用予定ファイル:", style="Header.TLabel").pack(side=tk.LEFT)
        self.sch_btn = ttk.Button(row2, text="ファイルを選択…", command=self._select_schedule)
        self.sch_btn.pack(side=tk.LEFT, padx=(10, 0))
        self.sch_label = ttk.Label(row2, text="未選択", style="Path.TLabel")
        self.sch_label.pack(side=tk.LEFT, padx=(10, 0), fill=tk.X, expand=True)

        # 計算実行ボタン
        btn_frame = ttk.Frame(self)
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        self.calc_btn = ttk.Button(
            btn_frame, text="▶ 計算実行", style="Accent.TButton", command=self._run_calculation
        )
        self.calc_btn.pack(side=tk.LEFT)

        # --- 中央: タブ ---
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        # タブ1: 返品候補
        tab1 = ttk.Frame(self.notebook)
        self.notebook.add(tab1, text=" 返品候補リスト ")
        self._build_surplus_tab(tab1)

        # タブ2: 発注候補
        tab2 = ttk.Frame(self.notebook)
        self.notebook.add(tab2, text=" 発注候補リスト ")
        self._build_shortage_tab(tab2)

        # --- 下部: ステータスバー ---
        self.status_var = tk.StringVar(value="ファイルを選択してください。")
        status_bar = ttk.Label(
            self, textvariable=self.status_var, style="Status.TLabel",
            relief=tk.SUNKEN, anchor=tk.W, padding=(6, 3)
        )
        status_bar.pack(fill=tk.X, side=tk.BOTTOM, padx=0, pady=0)

    def _build_surplus_tab(self, parent):
        """返品候補タブの中身を作る。"""
        cols = ("code", "name", "unit", "stock", "scheduled", "surplus", "price")
        hdrs = ("コード", "薬品名", "単位", "在庫数", "使用予定量", "返品可能数", "薬価")
        widths = (100, 300, 60, 90, 90, 90, 100)

        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X, padx=5, pady=4)
        ttk.Label(toolbar, text="ソート:").pack(side=tk.LEFT)
        ttk.Button(toolbar, text="薬価 高い順", command=lambda: self._sort_surplus("price", False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="数量 多い順", command=lambda: self._sort_surplus("surplus", False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="CSVで保存", command=lambda: self._save_csv("surplus")).pack(side=tk.RIGHT, padx=2)

        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        self.surplus_tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")
        for c, h, w in zip(cols, hdrs, widths):
            self.surplus_tree.heading(c, text=h, command=lambda _c=c: self._sort_surplus_toggle(_c))
            anchor = tk.E if c in ("stock", "scheduled", "surplus", "price") else tk.W
            self.surplus_tree.column(c, width=w, anchor=anchor)
        vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.surplus_tree.yview)
        self.surplus_tree.configure(yscrollcommand=vsb.set)
        self.surplus_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    def _build_shortage_tab(self, parent):
        """発注候補タブの中身を作る。"""
        cols = ("code", "name", "unit", "stock", "scheduled", "shortage", "price", "cost")
        hdrs = ("コード", "薬品名", "単位", "在庫数", "使用予定量", "発注必要数", "薬価", "概算金額")
        widths = (100, 280, 60, 90, 90, 90, 100, 110)

        toolbar = ttk.Frame(parent)
        toolbar.pack(fill=tk.X, padx=5, pady=4)
        ttk.Label(toolbar, text="ソート:").pack(side=tk.LEFT)
        ttk.Button(toolbar, text="金額 高い順", command=lambda: self._sort_shortage("cost", False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="数量 多い順", command=lambda: self._sort_shortage("shortage", False)).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="CSVで保存", command=lambda: self._save_csv("shortage")).pack(side=tk.RIGHT, padx=2)
        ttk.Button(toolbar, text="発注PDFで保存", command=self._save_order_pdf).pack(side=tk.RIGHT, padx=2)

        container = ttk.Frame(parent)
        container.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))

        self.shortage_tree = ttk.Treeview(container, columns=cols, show="headings", selectmode="browse")
        for c, h, w in zip(cols, hdrs, widths):
            self.shortage_tree.heading(c, text=h, command=lambda _c=c: self._sort_shortage_toggle(_c))
            anchor = tk.E if c in ("stock", "scheduled", "shortage", "price", "cost") else tk.W
            self.shortage_tree.column(c, width=w, anchor=anchor)
        vsb = ttk.Scrollbar(container, orient=tk.VERTICAL, command=self.shortage_tree.yview)
        self.shortage_tree.configure(yscrollcommand=vsb.set)
        self.shortage_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    # ---- ファイル選択 ----
    def _select_inventory(self):
        path = filedialog.askopenfilename(
            title="在庫ファイルを選択",
            filetypes=[("CSV ファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if path:
            self.inventory_path = path
            self.inv_label.config(text=os.path.basename(path))
            self.status_var.set(f"在庫ファイル: {os.path.basename(path)} を選択しました。")

    def _select_schedule(self):
        path = filedialog.askopenfilename(
            title="使用予定ファイルを選択",
            filetypes=[("CSV ファイル", "*.csv"), ("すべてのファイル", "*.*")]
        )
        if path:
            self.schedule_path = path
            self.sch_label.config(text=os.path.basename(path))
            self.status_var.set(f"使用予定ファイル: {os.path.basename(path)} を選択しました。")

    # ---- 計算 ----
    def _run_calculation(self):
        # バリデーション
        if not self.inventory_path:
            messagebox.showwarning("未選択", "在庫ファイルが選択されていません。")
            return
        if not self.schedule_path:
            messagebox.showwarning("未選択", "使用予定ファイルが選択されていません。")
            return

        try:
            inv_headers, inv_rows = read_csv_auto_encoding(self.inventory_path, skiprows=0)
        except UnicodeDecodeError as e:
            messagebox.showerror("読み込みエラー", f"在庫ファイルの文字コードを判定できませんでした。\n{e}")
            return
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"在庫ファイルの読み込みに失敗しました。\n{e}")
            return

        try:
            sch_headers, sch_rows = read_csv_auto_encoding(self.schedule_path, skiprows=4)
        except UnicodeDecodeError as e:
            messagebox.showerror("読み込みエラー", f"使用予定ファイルの文字コードを判定できませんでした。\n{e}")
            return
        except Exception as e:
            messagebox.showerror("読み込みエラー", f"使用予定ファイルの読み込みに失敗しました。\n{e}")
            return

        # カラム存在チェック
        inv_key = "レセコンコード"
        sch_key = "薬剤ｺｰﾄﾞ"

        required_inv = {inv_key, "薬品名", "単位", "在庫数"}
        required_sch = {sch_key, "薬剤名", "薬価", "使用予定量"}

        missing_inv = required_inv - set(inv_headers)
        if missing_inv:
            messagebox.showerror(
                "カラムエラー",
                f"在庫ファイルに必要なカラムがありません:\n{', '.join(missing_inv)}\n\n"
                f"検出されたカラム:\n{', '.join(inv_headers)}"
            )
            return

        missing_sch = required_sch - set(sch_headers)
        if missing_sch:
            messagebox.showerror(
                "カラムエラー",
                f"使用予定ファイルに必要なカラムがありません:\n{', '.join(missing_sch)}\n\n"
                f"検出されたカラム:\n{', '.join(sch_headers)}"
            )
            return

        # 在庫辞書を構築 {コード: row}
        inv_dict = {}
        for row in inv_rows:
            code = str(row.get(inv_key, "")).strip()
            if code:
                inv_dict[code] = row

        # 使用予定辞書を構築 {コード: row}
        sch_dict = {}
        for row in sch_rows:
            code = str(row.get(sch_key, "")).strip()
            if code:
                sch_dict[code] = row

        # マージ対象: 両方に存在するコード + 片方だけのコード
        all_codes = set(inv_dict.keys()) | set(sch_dict.keys())

        surplus_list = []
        shortage_list = []

        for code in all_codes:
            inv_row = inv_dict.get(code)
            sch_row = sch_dict.get(code)

            stock = to_number(inv_row.get("在庫数")) if inv_row else 0.0
            scheduled = to_number(sch_row.get("使用予定量")) if sch_row else 0.0
            price = to_number(sch_row.get("薬価")) if sch_row else 0.0

            name = ""
            unit = ""
            if inv_row:
                name = inv_row.get("薬品名", "")
                unit = inv_row.get("単位", "")
            elif sch_row:
                name = sch_row.get("薬剤名", "")
                unit = ""

            diff = stock - scheduled  # 正=余剰, 負=不足

            if diff > 0:
                surplus_list.append({
                    "code": code,
                    "name": name,
                    "unit": unit,
                    "stock": stock,
                    "scheduled": scheduled,
                    "surplus": diff,
                    "price": price,
                })
            elif diff < 0:
                shortage_val = abs(diff)
                shortage_list.append({
                    "code": code,
                    "name": name,
                    "unit": unit,
                    "stock": stock,
                    "scheduled": scheduled,
                    "shortage": shortage_val,
                    "price": price,
                    "cost": shortage_val * price,
                })

        self.surplus_data = surplus_list
        self.shortage_data = shortage_list

        # 初期ソート: 返品候補→薬価高い順、発注候補→概算金額高い順
        self.surplus_data.sort(key=lambda r: r["price"], reverse=True)
        self.shortage_data.sort(key=lambda r: r["cost"], reverse=True)

        self._refresh_surplus_tree()
        self._refresh_shortage_tree()

        total_surplus = len(surplus_list)
        total_shortage = len(shortage_list)
        total_cost = sum(r["cost"] for r in shortage_list)
        self.status_var.set(
            f"計算完了 — 返品候補: {total_surplus} 件 / "
            f"発注候補: {total_shortage} 件（概算合計: ¥{total_cost:,.0f}）"
        )
        messagebox.showinfo(
            "計算完了",
            f"返品候補: {total_surplus} 件\n"
            f"発注候補: {total_shortage} 件\n"
            f"発注概算合計: ¥{total_cost:,.0f}"
        )

    # ---- ツリー更新 ----
    def _refresh_surplus_tree(self):
        tree = self.surplus_tree
        tree.delete(*tree.get_children())
        for r in self.surplus_data:
            tree.insert("", tk.END, values=(
                r["code"],
                r["name"],
                r["unit"],
                f'{r["stock"]:,.1f}',
                f'{r["scheduled"]:,.1f}',
                f'{r["surplus"]:,.1f}',
                f'{r["price"]:,.2f}',
            ))

    def _refresh_shortage_tree(self):
        tree = self.shortage_tree
        tree.delete(*tree.get_children())
        for r in self.shortage_data:
            tree.insert("", tk.END, values=(
                r["code"],
                r["name"],
                r["unit"],
                f'{r["stock"]:,.1f}',
                f'{r["scheduled"]:,.1f}',
                f'{r["shortage"]:,.1f}',
                f'{r["price"]:,.2f}',
                f'{r["cost"]:,.0f}',
            ))

    # ---- ソート ----
    def _sort_surplus(self, col, ascending):
        self.surplus_data.sort(key=lambda r: r.get(col, 0), reverse=not ascending)
        self._refresh_surplus_tree()

    def _sort_surplus_toggle(self, col):
        if self.surplus_sort_col == col:
            self.surplus_sort_asc = not self.surplus_sort_asc
        else:
            self.surplus_sort_col = col
            self.surplus_sort_asc = False  # 最初は降順
        self._sort_surplus(col, self.surplus_sort_asc)

    def _sort_shortage(self, col, ascending):
        self.shortage_data.sort(key=lambda r: r.get(col, 0), reverse=not ascending)
        self._refresh_shortage_tree()

    def _sort_shortage_toggle(self, col):
        if self.shortage_sort_col == col:
            self.shortage_sort_asc = not self.shortage_sort_asc
        else:
            self.shortage_sort_col = col
            self.shortage_sort_asc = False
        self._sort_shortage(col, self.shortage_sort_asc)

    # ---- CSV 保存 ----
    def _save_csv(self, kind):
        if kind == "surplus":
            data = self.surplus_data
            columns = ["code", "name", "unit", "stock", "scheduled", "surplus", "price"]
            headers = ["コード", "薬品名", "単位", "在庫数", "使用予定量", "返品可能数", "薬価"]
            default_name = "返品候補リスト.csv"
        else:
            data = self.shortage_data
            columns = ["code", "name", "unit", "stock", "scheduled", "shortage", "price", "cost"]
            headers = ["コード", "薬品名", "単位", "在庫数", "使用予定量", "発注必要数", "薬価", "概算金額"]
            default_name = "発注候補リスト.csv"

        if not data:
            messagebox.showinfo("データなし", "保存するデータがありません。先に計算を実行してください。")
            return

        path = filedialog.asksaveasfilename(
            title="CSVファイルを保存",
            defaultextension=".csv",
            initialfile=default_name,
            filetypes=[("CSV ファイル", "*.csv")]
        )
        if not path:
            return

        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for r in data:
                    writer.writerow([r[c] for c in columns])
            self.status_var.set(f"保存完了: {os.path.basename(path)}")
            messagebox.showinfo("保存完了", f"ファイルを保存しました。\n{path}")
        except Exception as e:
            messagebox.showerror("保存エラー", f"ファイルの保存に失敗しました。\n{e}")

    # ---- 発注 PDF 保存 ----
    def _save_order_pdf(self):
        if not self.shortage_data:
            messagebox.showinfo("データなし", "保存するデータがありません。先に計算を実行してください。")
            return

        today_str = date.today().strftime("%Y%m%d")
        default_name = f"発注リスト_{today_str}.pdf"
        path = filedialog.asksaveasfilename(
            title="発注PDFを保存",
            defaultextension=".pdf",
            initialfile=default_name,
            filetypes=[("PDF ファイル", "*.pdf")],
        )
        if not path:
            return

        try:
            self._build_order_pdf(path)
            self.status_var.set(f"PDF保存完了: {os.path.basename(path)}")
            messagebox.showinfo("保存完了", f"発注PDFを保存しました。\n{path}")
        except Exception as e:
            messagebox.showerror("PDF保存エラー", f"PDFの作成に失敗しました。\n{e}")

    def _build_order_pdf(self, filepath):
        """発注候補データから PDF を生成する。"""
        font_name = _register_japanese_font()
        today_str = date.today().strftime("%Y年%m月%d日")

        doc = SimpleDocTemplate(
            filepath,
            pagesize=A4,
            leftMargin=15 * mm,
            rightMargin=15 * mm,
            topMargin=15 * mm,
            bottomMargin=15 * mm,
        )

        styles = getSampleStyleSheet()
        style_title = ParagraphStyle(
            "TitleJa", parent=styles["Title"],
            fontName=font_name, fontSize=16, leading=22,
        )
        style_normal = ParagraphStyle(
            "NormalJa", parent=styles["Normal"],
            fontName=font_name, fontSize=9, leading=12,
        )
        style_header = ParagraphStyle(
            "HeaderJa", parent=styles["Normal"],
            fontName=font_name, fontSize=9, leading=12,
            textColor=colors.white,
        )
        style_date = ParagraphStyle(
            "DateJa", parent=styles["Normal"],
            fontName=font_name, fontSize=10, leading=14,
        )

        elements = []

        # タイトル
        elements.append(Paragraph("発注リスト", style_title))
        elements.append(Spacer(1, 2 * mm))
        elements.append(Paragraph(f"作成日: {today_str}", style_date))
        elements.append(Spacer(1, 6 * mm))

        # テーブルデータ構築
        header_row = [
            Paragraph("日付", style_header),
            Paragraph("医薬品名", style_header),
            Paragraph("発注必要数", style_header),
            Paragraph("バーコード", style_header),
        ]
        table_data = [header_row]

        for r in self.shortage_data:
            bc_img = _generate_barcode_image(r["code"], width_mm=40, height_mm=10)
            bc_cell = bc_img if bc_img else Paragraph(str(r["code"]), style_normal)

            row = [
                Paragraph(today_str, style_normal),
                Paragraph(str(r["name"]), style_normal),
                Paragraph(f'{r["shortage"]:,.1f}', style_normal),
                bc_cell,
            ]
            table_data.append(row)

        # カラム幅
        page_w = A4[0] - 30 * mm  # 左右マージン除外
        col_widths = [28 * mm, page_w - 28 * mm - 25 * mm - 45 * mm, 25 * mm, 45 * mm]

        table = Table(table_data, colWidths=col_widths, repeatRows=1)
        table.setStyle(TableStyle([
            # ヘッダー
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#3a3a3a")),
            ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
            ("FONTNAME", (0, 0), (-1, 0), font_name),
            ("FONTSIZE", (0, 0), (-1, 0), 9),
            # データ行
            ("FONTNAME", (0, 1), (-1, -1), font_name),
            ("FONTSIZE", (0, 1), (-1, -1), 9),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#f5f5f5")]),
            # 枠線
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            # 位置
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("ALIGN", (2, 0), (2, -1), "RIGHT"),
            ("TOPPADDING", (0, 0), (-1, -1), 3),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 4),
            ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ]))

        elements.append(table)

        # 合計行
        total_items = len(self.shortage_data)
        total_cost = sum(r["cost"] for r in self.shortage_data)
        elements.append(Spacer(1, 4 * mm))
        elements.append(Paragraph(
            f"合計 {total_items} 品目 / 概算合計金額: ¥{total_cost:,.0f}",
            style_date,
        ))

        doc.build(elements)


# ---------------------------------------------------------------------------
# エントリーポイント
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
