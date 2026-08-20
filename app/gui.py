# -*- coding: utf-8 -*-
"""Tkinter 图形界面。"""
import datetime as dtm
import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import excel_writer
from . import processor
from . import word_writer

STEPS = 7


class App:
    """主窗口。"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("蚊媒监测数据处理工具")
        self.root.minsize(780, 620)
        self.queue = queue.Queue()
        self.worker = None
        self._build()
        self.root.after(100, self._poll)

    # ------------------------------------------------------------------ 界面
    def _build(self):
        pad = {"padx": 8, "pady": 6}
        frm = ttk.Frame(self.root, padding=12)
        frm.pack(fill="both", expand=True)

        # 源文件
        row0 = ttk.Frame(frm)
        row0.pack(fill="x", **pad)
        ttk.Label(row0, text="源Excel文件：").pack(side="left")
        self.var_file = tk.StringVar()
        ent_file = ttk.Entry(row0, textvariable=self.var_file)
        ent_file.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row0, text="浏览…", command=self._pick_file).pack(side="left")

        # 目标日期
        row1 = ttk.Frame(frm)
        row1.pack(fill="x", **pad)
        ttk.Label(row1, text="目标日期：").pack(side="left")
        now = dtm.date.today()
        self.var_year = tk.StringVar(value=str(now.year))
        self.var_month = tk.StringVar(value=str(now.month))
        self.var_day = tk.StringVar(value=str(now.day))
        cb_year = ttk.Combobox(row1, textvariable=self.var_year, width=7,
                               values=[str(y) for y in range(2000, 2036)])
        cb_month = ttk.Combobox(row1, textvariable=self.var_month, width=4,
                                values=[str(m) for m in range(1, 13)])
        cb_day = ttk.Combobox(row1, textvariable=self.var_day, width=4,
                              values=[str(d) for d in range(1, 32)])
        for w, label in ((cb_year, "年"), (cb_month, "月"), (cb_day, "日")):
            w.pack(side="left", padx=(10, 0))
            ttk.Label(row1, text=label).pack(side="left")

        # 排除字段
        row2 = ttk.Frame(frm)
        row2.pack(fill="x", **pad)
        ttk.Label(row2, text="排除字段（可选）：").pack(side="left")
        self.var_exclude = tk.StringVar()
        ent_excl = ttk.Entry(row2, textvariable=self.var_exclude)
        ent_excl.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Label(row2, text="如：荔湾区（删除所有含该字段的记录）").pack(side="left")

        # 输出目录
        row3 = ttk.Frame(frm)
        row3.pack(fill="x", **pad)
        ttk.Label(row3, text="输出目录：").pack(side="left")
        self.var_outdir = tk.StringVar()
        ent_dir = ttk.Entry(row3, textvariable=self.var_outdir)
        ent_dir.pack(side="left", fill="x", expand=True, padx=4)
        ttk.Button(row3, text="浏览…", command=self._pick_dir).pack(side="left")

        # 按钮与进度
        row4 = ttk.Frame(frm)
        row4.pack(fill="x", **pad)
        self.btn_start = ttk.Button(row4, text="开始处理", command=self._start)
        self.btn_start.pack(side="left")
        self.pbar = ttk.Progressbar(row4, maximum=STEPS, length=360)
        self.pbar.pack(side="left", padx=12, fill="x", expand=True)
        self.lbl_step = ttk.Label(row4, text="就绪")
        self.lbl_step.pack(side="left")

        # 日志
        row5 = ttk.Frame(frm)
        row5.pack(fill="both", expand=True, **pad)
        self.txt = tk.Text(row5, height=14, state="disabled", wrap="word")
        scroll = ttk.Scrollbar(row5, command=self.txt.yview)
        self.txt.configure(yscrollcommand=scroll.set)
        self.txt.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

    def _pick_file(self):
        path = filedialog.askopenfilename(
            title="选择源Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")],
        )
        if path:
            self.var_file.set(path)
            if not self.var_outdir.get().strip():
                self.var_outdir.set(os.path.dirname(path))

    def _pick_dir(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.var_outdir.set(path)

    def _log(self, msg):
        self.txt.configure(state="normal")
        self.txt.insert("end", msg + "\n")
        self.txt.see("end")
        self.txt.configure(state="disabled")

    # ------------------------------------------------------------------ 处理
    def _validate(self):
        """返回 (源文件, 年, 月, 日, 排除字段, 输出目录) 或抛出 ValueError。"""
        src = self.var_file.get().strip()
        if not src:
            raise ValueError("请先选择源Excel文件。")
        if not os.path.isfile(src):
            raise ValueError(f"源文件不存在：{src}")
        year = self.var_year.get().strip()
        month = self.var_month.get().strip()
        day = self.var_day.get().strip()
        try:
            year, month, day = int(year), int(month), int(day)
        except ValueError:
            raise ValueError("年、月、日必须为整数。")
        import calendar
        if not (2000 <= year <= 2100):
            raise ValueError(f"年份 {year} 超出合理范围（2000~2100）。")
        if not (1 <= month <= 12):
            raise ValueError(f"月份 {month} 不合法（应为 1~12）。")
        if not (1 <= day <= calendar.monthrange(year, month)[1]):
            raise ValueError(f"日期 {year}-{month}-{day} 不合法。")
        exclude = self.var_exclude.get().strip()
        outdir = self.var_outdir.get().strip() or os.path.dirname(src)
        if not os.path.isdir(outdir):
            try:
                os.makedirs(outdir, exist_ok=True)
            except OSError as exc:
                raise ValueError(f"无法创建输出目录：{outdir}（{exc}）")
        return src, year, month, day, exclude, outdir

    def _start(self):
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("提示", "正在处理中，请稍候…")
            return
        try:
            src, year, month, day, exclude, outdir = self._validate()
        except ValueError as exc:
            messagebox.showerror("输入有误", str(exc))
            return
        self.btn_start.configure(state="disabled")
        self.pbar["value"] = 0
        self._log("=" * 50)
        self._log("开始处理…")
        self.worker = threading.Thread(
            target=self._work, args=(src, year, month, day, exclude, outdir), daemon=True
        )
        self.worker.start()

    def _work(self, src, year, month, day, exclude, outdir):
        try:
            def log(msg, step=None):
                self.queue.put(("log", msg, step))

            result = processor.process_all(src, year, month, day, exclude or None, log=log)

            target = result.target
            mm, dd = f"{target.month:02d}", f"{target.day:02d}"
            p_path = excel_writer.process_excel_path(outdir, target)
            w_path = excel_writer.word_path(outdir, target)
            s_path = excel_writer.summary_excel_path(outdir, target)

            log("正在生成计算过程Excel（BI/ADI 全部中间表）…", 6)
            excel_writer.build_process_excel(result, p_path)

            log("正在生成日报Word…", 6)
            word_writer.build_word(result, w_path)

            log("正在生成监测点汇总Excel…", 6)
            excel_writer.build_summary_excel(result, s_path)

            log("全部处理完成！", 7)
            self.queue.put(("done", (p_path, w_path, s_path)))
        except processor.ProcessingError as exc:
            self.queue.put(("error", str(exc)))
        except Exception as exc:  # noqa: BLE001
            import traceback
            self.queue.put(("error", f"发生未预期错误：{exc}\n{traceback.format_exc()}"))

    def _poll(self):
        try:
            while True:
                kind, payload = self.queue.get_nowait()
                if kind == "log":
                    msg, step = payload
                    self._log(msg)
                    if step:
                        self.pbar["value"] = step
                        self.lbl_step.configure(text=msg)
                elif kind == "done":
                    p_path, w_path, s_path = payload
                    self.btn_start.configure(state="normal")
                    self.lbl_step.configure(text="完成")
                    messagebox.showinfo(
                        "处理完成",
                        "生成的文件如下：\n\n"
                        f"1. 计算过程Excel：\n{p_path}\n\n"
                        f"2. 日报Word：\n{w_path}\n\n"
                        f"3. 监测点汇总Excel：\n{s_path}",
                    )
                elif kind == "error":
                    msg = payload
                    self.btn_start.configure(state="normal")
                    self.lbl_step.configure(text="出错")
                    self._log(msg)
                    messagebox.showerror("处理出错", msg.split("\n")[0] if msg else "未知错误")
        except queue.Empty:
            pass
        self.root.after(100, self._poll)


def run():
    root = tk.Tk()
    App(root)
    root.mainloop()
