"""
公文序號改名工具
將掃描PDF的右下角序號章（例如 202604001）辨識出來，批次改名
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import re

try:
    import fitz  # pymupdf
except ImportError:
    fitz = None

try:
    from PIL import Image, ImageEnhance, ImageFilter, ImageOps
    import pytesseract
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
    OCR_AVAILABLE = True
except ImportError:
    OCR_AVAILABLE = False


class PDFRenamer:
    def __init__(self, root):
        self.root = root
        self.root.title("公文序號改名工具 v1.0")
        self.root.geometry("950x620")
        self.root.resizable(True, True)
        self.pdf_files = []
        self.results = []
        self.setup_ui()
        self.check_dependencies()

    def check_dependencies(self):
        missing = []
        if fitz is None:
            missing.append("pymupdf")
        if not OCR_AVAILABLE:
            missing.append("pytesseract / Pillow")
        if missing:
            msg = (
                "缺少以下套件，請在命令提示字元執行安裝指令：\n\n"
                f"pip install {' '.join(missing)}\n\n"
                "安裝完成後重新啟動程式。"
            )
            messagebox.showerror("缺少套件", msg)

    # ──────────────────────────────────────────
    # UI 建立
    # ──────────────────────────────────────────
    def setup_ui(self):
        # ── 頂部按鈕列 ──
        ctrl = tk.Frame(self.root, pady=8, padx=10)
        ctrl.pack(fill=tk.X)

        tk.Button(ctrl, text="📁  選擇資料夾", width=16,
                  command=self.select_folder).pack(side=tk.LEFT, padx=4)
        tk.Button(ctrl, text="📄  選擇PDF檔案", width=16,
                  command=self.select_files).pack(side=tk.LEFT, padx=4)

        ttk.Separator(ctrl, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=8)

        self.scan_btn = tk.Button(ctrl, text="🔍  開始辨識", width=14,
                                  command=self.start_scan, state=tk.DISABLED,
                                  bg="#2196F3", fg="white", font=("", 10, "bold"))
        self.scan_btn.pack(side=tk.LEFT, padx=4)

        self.rename_btn = tk.Button(ctrl, text="✅  確認改名", width=14,
                                    command=self.confirm_rename, state=tk.DISABLED,
                                    bg="#4CAF50", fg="white", font=("", 10, "bold"))
        self.rename_btn.pack(side=tk.LEFT, padx=4)

        self.progress_var = tk.StringVar(value="")
        tk.Label(ctrl, textvariable=self.progress_var, fg="#555").pack(side=tk.LEFT, padx=12)

        # ── 進度條 ──
        self.pb = ttk.Progressbar(self.root, mode="determinate")
        self.pb.pack(fill=tk.X, padx=10)

        # ── 說明 ──
        hint = tk.Label(self.root,
                        text="💡 雙擊「辨識序號」欄位可手動修改  ·  綠色=成功辨識  紅色=未辨識（需手動補上）",
                        fg="#666", font=("", 9))
        hint.pack(anchor=tk.W, padx=12, pady=(4, 0))

        # ── 表格 ──
        table_frame = tk.Frame(self.root)
        table_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)

        cols = ("原始檔名", "辨識序號", "新檔名", "狀態")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings")
        self.tree.heading("原始檔名", text="原始檔名")
        self.tree.heading("辨識序號", text="辨識序號（可雙擊修改）")
        self.tree.heading("新檔名", text="新檔名")
        self.tree.heading("狀態", text="狀態")
        self.tree.column("原始檔名", width=280)
        self.tree.column("辨識序號", width=160)
        self.tree.column("新檔名", width=200)
        self.tree.column("狀態", width=120, anchor=tk.CENTER)

        vsb = ttk.Scrollbar(table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscroll=vsb.set, xscroll=hsb.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

        self.tree.tag_configure("ok", foreground="#1B8A1B")
        self.tree.tag_configure("warn", foreground="#CC3300")
        self.tree.tag_configure("manual", foreground="#0055AA")
        self.tree.bind("<Double-1>", self.edit_serial)

        # ── 狀態列 ──
        self.status_var = tk.StringVar(value="請選擇資料夾或PDF檔案開始")
        tk.Label(self.root, textvariable=self.status_var,
                 bd=1, relief=tk.SUNKEN, anchor=tk.W, padx=6).pack(
            fill=tk.X, side=tk.BOTTOM)

    # ──────────────────────────────────────────
    # 檔案選擇
    # ──────────────────────────────────────────
    def select_folder(self):
        folder = filedialog.askdirectory(title="選擇包含PDF的資料夾")
        if not folder:
            return
        self.pdf_files = sorted([
            os.path.join(folder, f)
            for f in os.listdir(folder)
            if f.lower().endswith(".pdf")
        ])
        count = len(self.pdf_files)
        self.status_var.set(f"資料夾：{folder}  ·  找到 {count} 個PDF")
        self.scan_btn.config(state=tk.NORMAL if count else tk.DISABLED)
        self._clear_table()

    def select_files(self):
        files = filedialog.askopenfilenames(
            title="選擇PDF檔案", filetypes=[("PDF 檔案", "*.pdf")])
        if not files:
            return
        self.pdf_files = sorted(files)
        self.status_var.set(f"已選擇 {len(self.pdf_files)} 個PDF檔案")
        self.scan_btn.config(state=tk.NORMAL)
        self._clear_table()

    def _clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        self.results = []
        self.rename_btn.config(state=tk.DISABLED)

    # ──────────────────────────────────────────
    # OCR 辨識
    # ──────────────────────────────────────────
    def start_scan(self):
        if not self.pdf_files:
            return
        self.scan_btn.config(state=tk.DISABLED)
        self.rename_btn.config(state=tk.DISABLED)
        self._clear_table()
        self.pb["maximum"] = len(self.pdf_files)
        self.pb["value"] = 0
        thread = threading.Thread(target=self._process_all, daemon=True)
        thread.start()

    def _process_all(self):
        total = len(self.pdf_files)
        for i, filepath in enumerate(self.pdf_files):
            self.root.after(0, lambda i=i: (
                self.progress_var.set(f"辨識中 {i+1} / {total}"),
                self.pb.config(value=i+1)
            ))
            filename = os.path.basename(filepath)
            serial = self._extract_serial(filepath)
            new_name = f"{serial}.pdf" if serial else filename
            status = "✓ 已辨識" if serial else "⚠ 未辨識"
            tag = "ok" if serial else "warn"
            result = {
                "filepath": filepath,
                "original": filename,
                "serial": serial or "",
                "new_name": new_name,
            }
            self.results.append(result)
            self.root.after(0, lambda r=result, t=tag, s=status: self.tree.insert(
                "", tk.END,
                values=(r["original"], r["serial"], r["new_name"], s),
                tags=(t,)
            ))

        ok = sum(1 for r in self.results if r["serial"])
        self.root.after(0, lambda: (
            self.progress_var.set(f"辨識完成！  成功 {ok} / 未辨識 {total - ok}"),
            self.scan_btn.config(state=tk.NORMAL),
            self.rename_btn.config(state=tk.NORMAL),
            self.status_var.set(f"共 {total} 個檔案  ·  辨識成功 {ok}  ·  未辨識 {total - ok}（請雙擊補上序號）")
        ))

    def _extract_serial(self, filepath):
        """從PDF第一頁右下角辨識9位數序號"""
        if fitz is None or not OCR_AVAILABLE:
            return None
        try:
            doc = fitz.open(filepath)
            page = doc[0]
            w, h = page.rect.width, page.rect.height

            # 擷取右下角 35% x 28% 的區域
            crop = fitz.Rect(w * 0.65, h * 0.72, w, h)
            mat = fitz.Matrix(3.5, 3.5)   # 放大以提高辨識率
            pix = page.get_pixmap(matrix=mat, clip=crop)
            doc.close()

            img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
            img = img.convert("L")                        # 灰階
            img = ImageOps.autocontrast(img, cutoff=2)    # 自動對比
            enhancer = ImageEnhance.Sharpness(img)
            img = enhancer.enhance(2.5)

            # 二值化
            img = img.point(lambda x: 0 if x < 150 else 255, "1")

            cfg = "--psm 6 -c tessedit_char_whitelist=0123456789"
            text = pytesseract.image_to_string(img, config=cfg)
            clean = re.sub(r"\D", "", text)

            # 找9位數序號（格式 YYYYMM + 3位流水號）
            matches = re.findall(r"20\d{7}", clean)
            if matches:
                return matches[0]

            # 寬鬆匹配：任意9位數
            matches = re.findall(r"\d{9}", clean)
            if matches:
                return matches[0]

            return None
        except Exception as e:
            print(f"[錯誤] {filepath}: {e}")
            return None

    # ──────────────────────────────────────────
    # 手動編輯序號
    # ──────────────────────────────────────────
    def edit_serial(self, event):
        col = self.tree.identify_column(event.x)
        if col != "#2":   # 只允許編輯「辨識序號」欄
            return
        selection = self.tree.selection()
        if not selection:
            return
        item = selection[0]
        idx = self.tree.index(item)
        values = self.tree.item(item, "values")

        dlg = tk.Toplevel(self.root)
        dlg.title("修改序號")
        dlg.geometry("320x130")
        dlg.resizable(False, False)
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(dlg, text=f"檔案：{values[0]}", wraplength=300,
                 font=("", 9), fg="#555").pack(pady=(10, 2), padx=10, anchor=tk.W)
        tk.Label(dlg, text="序號（9位數）：").pack(anchor=tk.W, padx=14)

        entry_var = tk.StringVar(value=values[1])
        entry = tk.Entry(dlg, textvariable=entry_var, width=20, font=("", 12))
        entry.pack(padx=14)
        entry.select_range(0, tk.END)
        entry.focus()

        def save():
            serial = entry_var.get().strip()
            if not re.fullmatch(r"\d{9}", serial):
                messagebox.showwarning("格式錯誤", "請輸入9位數字（例如 202604001）", parent=dlg)
                return
            self.results[idx]["serial"] = serial
            self.results[idx]["new_name"] = f"{serial}.pdf"
            self.tree.item(item, values=(
                values[0], serial, f"{serial}.pdf", "✎ 手動修改"), tags=("manual",))
            dlg.destroy()

        tk.Button(dlg, text="確認", command=save, bg="#2196F3",
                  fg="white", width=10).pack(pady=8)
        dlg.bind("<Return>", lambda e: save())

    # ──────────────────────────────────────────
    # 批次改名
    # ──────────────────────────────────────────
    def confirm_rename(self):
        to_rename = [r for r in self.results if r["serial"]]
        skipped = [r for r in self.results if not r["serial"]]

        if not to_rename:
            messagebox.showwarning("無法改名", "沒有任何檔案有辨識到序號。\n請雙擊表格手動輸入序號。")
            return

        msg = (
            f"確定要改名嗎？\n\n"
            f"✅ 將改名：{len(to_rename)} 個檔案\n"
            f"⚠️  跳過（無序號）：{len(skipped)} 個檔案"
        )
        if not messagebox.askyesno("確認改名", msg):
            return

        success, errors = 0, []
        for r in to_rename:
            try:
                old = r["filepath"]
                new = os.path.join(os.path.dirname(old), r["new_name"])
                if os.path.abspath(old) == os.path.abspath(new):
                    success += 1
                    continue
                if os.path.exists(new):
                    errors.append(f"「{r['new_name']}」已存在，跳過")
                    continue
                os.rename(old, new)
                r["filepath"] = new
                r["original"] = r["new_name"]
                success += 1
            except Exception as e:
                errors.append(f"{r['original']}: {e}")

        # 重新整理表格
        for item in self.tree.get_children():
            self.tree.delete(item)
        for r in self.results:
            tag = "ok" if r["serial"] else "warn"
            status = "✅ 已改名" if r["serial"] else "⚠ 跳過"
            self.tree.insert("", tk.END,
                             values=(r["new_name"], r["serial"], r["new_name"], status),
                             tags=(tag,))

        result_msg = f"完成！成功改名 {success} 個檔案。"
        if errors:
            result_msg += "\n\n注意事項：\n" + "\n".join(errors)
        messagebox.showinfo("改名完成", result_msg)
        self.status_var.set(f"改名完成：{success} 個成功  ·  {len(errors)} 個錯誤")


# ──────────────────────────────────────────
if __name__ == "__main__":
    root = tk.Tk()
    app = PDFRenamer(root)
    root.mainloop()
