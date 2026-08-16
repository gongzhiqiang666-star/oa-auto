#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""交易汇总工具 —— 图形界面入口（Windows / macOS 均可运行）。

功能：
  1. 下载模板：一键生成并保存《交易汇总模板.xlsx》
  2. 导入模板：选择填好的模板文件
  3. 开始计算：自动汇总，结果自动保存为“原文件名_汇总.xlsx”
  4. 结果下载：打开文件夹 / 另存为
"""

from __future__ import annotations

import os
import queue
import shutil
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import core

APP_TITLE = "交易汇总工具"
APP_VERSION = "1.0.0"

# 界面配色
BG = "#F5F7FA"
PRIMARY = "#2F6FDB"
CARD = "#FFFFFF"
TEXT = "#1F2937"
MUTED = "#6B7280"
GREEN = "#16A34A"


def resource_path(name: str) -> Path:
    """兼容 PyInstaller 打包后的资源路径。"""
    base = getattr(sys, "_MEIPASS", Path(__file__).parent)
    return Path(base) / name


class App:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.input_path: Path | None = None
        self.result_path: Path | None = None
        self.running = False
        self.msg_queue: queue.Queue = queue.Queue()

        root.title(f"{APP_TITLE} v{APP_VERSION}")
        root.configure(bg=BG)
        root.minsize(720, 560)

        try:
            icon = resource_path("交易汇总工具.ico")
            if icon.exists():
                root.iconbitmap(str(icon))
        except Exception:
            pass

        self._build_ui()
        self.root.after(120, self._poll_queue)
        self.log(f"欢迎使用{APP_TITLE} v{APP_VERSION}")
        self.log("步骤：① 下载模板 → ② 填写模板 → ③ 导入模板 → ④ 开始计算 → ⑤ 下载结果")

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        outer = tk.Frame(self.root, bg=BG)
        outer.pack(fill="both", expand=True, padx=18, pady=14)

        # 标题
        title = tk.Label(outer, text="📊 交易汇总工具", bg=BG, fg=TEXT,
                         font=("Microsoft YaHei UI", 17, "bold"))
        title.pack(anchor="w")
        sub = tk.Label(outer, text="把交易明细 Excel 汇总到指定模板，自动生成汇总结果与问题备注",
                       bg=BG, fg=MUTED, font=("Microsoft YaHei UI", 10))
        sub.pack(anchor="w", pady=(2, 12))

        # ---- 步骤 1：下载模板 ----
        card1 = self._section(outer, "① 下载模板", "生成一份空白的《交易汇总模板.xlsx》，含交易明细、汇总模板、使用说明")
        row1 = tk.Frame(card1, bg=CARD)
        row1.pack(fill="x", pady=(10, 12))
        self.btn_template = self._button(row1, "下载模板", self.download_template)
        self.template_hint = tk.Label(row1, text="尚未下载", bg=CARD, fg=MUTED,
                                      font=("Microsoft YaHei UI", 9), anchor="w")
        self.template_hint.pack(side="left", padx=(12, 0))

        # ---- 步骤 2：导入模板 ----
        card2 = self._section(outer, "② 导入模板", "选择填好交易明细的 Excel 文件（.xlsx）")
        row2 = tk.Frame(card2, bg=CARD)
        row2.pack(fill="x", pady=(10, 12))
        self.btn_choose = self._button(row2, "选择文件…", self.choose_file)
        self.file_hint = tk.Label(row2, text="尚未选择文件", bg=CARD, fg=MUTED,
                                  font=("Microsoft YaHei UI", 9), anchor="w")
        self.file_hint.pack(side="left", padx=(12, 0), fill="x", expand=True)

        # ---- 步骤 3：开始计算 ----
        card3 = self._section(outer, "③ 开始计算", "结果自动保存到原文件旁边：原文件名_汇总.xlsx")
        row3 = tk.Frame(card3, bg=CARD)
        row3.pack(fill="x", pady=(10, 12))
        self.btn_calc = self._button(row3, "开始计算", self.start_calc, primary=True)
        self.calc_hint = tk.Label(row3, text="等待开始", bg=CARD, fg=MUTED,
                                  font=("Microsoft YaHei UI", 9), anchor="w")
        self.calc_hint.pack(side="left", padx=(12, 0), fill="x", expand=True)

        # ---- 步骤 4：结果下载 ----
        card4 = self._section(outer, "④ 下载结果", "计算完成后结果已自动保存；可打开文件夹或另存到其他位置")
        row4 = tk.Frame(card4, bg=CARD)
        row4.pack(fill="x", pady=(10, 12))
        self.btn_open = self._button(row4, "打开文件夹", self.open_folder, enabled=False)
        self.btn_saveas = self._button(row4, "另存为…", self.save_as, enabled=False)
        self.result_hint = tk.Label(row4, text="暂无结果", bg=CARD, fg=MUTED,
                                    font=("Microsoft YaHei UI", 9), anchor="w")
        self.result_hint.pack(side="left", padx=(12, 0), fill="x", expand=True)

        # ---- 运行日志 ----
        log_box = tk.Frame(outer, bg=BG)
        log_box.pack(fill="both", expand=True, pady=(14, 0))
        log_title = tk.Label(log_box, text="运行日志", bg=BG, fg=TEXT,
                             font=("Microsoft YaHei UI", 10, "bold"))
        log_title.pack(anchor="w", pady=(0, 4))
        self.log_text = tk.Text(log_box, height=9, bg="#0F172A", fg="#E2E8F0",
                                font=("Consolas", 10), relief="flat", wrap="word",
                                state="disabled", padx=10, pady=8)
        scroll = ttk.Scrollbar(log_box, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

    def _section(self, parent, title, desc):
        card = tk.Frame(parent, bg=CARD, highlightbackground="#E5E7EB", highlightthickness=1)
        card.pack(fill="x", pady=(0, 10))
        head = tk.Label(card, text=title, bg=CARD, fg=PRIMARY,
                        font=("Microsoft YaHei UI", 12, "bold"))
        head.pack(anchor="w", padx=14, pady=(10, 0))
        tip = tk.Label(card, text=desc, bg=CARD, fg=MUTED, font=("Microsoft YaHei UI", 9))
        tip.pack(anchor="w", padx=14)
        return card

    def _button(self, parent, text, command, primary=False, enabled=True):
        btn = tk.Button(
            parent, text=text, command=command,
            bg=PRIMARY if primary else "#E8EEFB",
            fg="#FFFFFF" if primary else PRIMARY,
            activebackground="#2459B8" if primary else "#D9E4F8",
            activeforeground="#FFFFFF" if primary else PRIMARY,
            font=("Microsoft YaHei UI", 10, "bold"),
            relief="flat", bd=0, cursor="hand2", padx=16, pady=7,
            state="normal" if enabled else "disabled",
        )
        btn.pack(side="left", padx=(14, 0), pady=(0, 0))
        return btn

    # ------------------------------------------------------------ 功能逻辑
    def log(self, msg: str):
        self.msg_queue.put(msg)

    def _poll_queue(self):
        try:
            while True:
                msg = self.msg_queue.get_nowait()
                if isinstance(msg, tuple):
                    kind, payload = msg
                    if kind == "RESULT":
                        self._finish_ok(payload)
                    elif kind == "ERROR":
                        self._finish_error(payload)
                else:
                    self.log_text.configure(state="normal")
                    self.log_text.insert("end", msg + "\n")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
        except queue.Empty:
            pass
        self.root.after(120, self._poll_queue)

    def download_template(self):
        default_dir = self._default_save_dir()
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="保存模板",
            initialdir=default_dir,
            initialfile="交易汇总模板.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        try:
            saved = core.save_template(Path(path))
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("下载失败", f"模板生成失败：\n{exc}", parent=self.root)
            self.log(f"✗ 模板生成失败：{exc}")
            return
        self.template_hint.configure(text=f"已保存：{saved}", fg=GREEN)
        self.log(f"✓ 模板已下载：{saved}")
        messagebox.showinfo("下载成功", f"模板已保存到：\n{saved}\n\n请用 Excel 打开并填写交易明细后，再回到本工具导入。",
                            parent=self.root)

    def choose_file(self):
        path = filedialog.askopenfilename(
            parent=self.root,
            title="选择填好的模板文件",
            filetypes=[("Excel 工作簿", "*.xlsx;*.xlsm"), ("所有文件", "*.*")],
        )
        if not path:
            return
        self.input_path = Path(path)
        self.file_hint.configure(text=str(self.input_path), fg=PRIMARY)
        self.calc_hint.configure(text="已选择文件，可以开始计算", fg=PRIMARY)
        self.log(f"✓ 已导入文件：{self.input_path}")

    def start_calc(self):
        if self.running:
            return
        if not self.input_path:
            messagebox.showwarning("提示", "请先导入模板文件（第②步）。", parent=self.root)
            return
        if not self.input_path.exists():
            messagebox.showerror("错误", f"文件不存在：\n{self.input_path}", parent=self.root)
            return
        if self.input_path.suffix.lower() not in (".xlsx", ".xlsm"):
            messagebox.showwarning("提示", "请选择 .xlsx 格式的 Excel 文件。", parent=self.root)
            return
        self.running = True
        self.btn_calc.configure(state="disabled", text="计算中…")
        self.calc_hint.configure(text="正在计算，请稍候…", fg=PRIMARY)
        threading.Thread(target=self._worker, args=(self.input_path,), daemon=True).start()

    def _worker(self, in_path: Path):
        try:
            out = core.calculate(in_path, log=self.log)
            self.msg_queue.put(("RESULT", out))
        except Exception as exc:  # noqa: BLE001
            self.msg_queue.put(("ERROR", str(exc)))

    def _finish_ok(self, out: Path):
        self.running = False
        self.result_path = out
        self.btn_calc.configure(state="normal", text="开始计算")
        self.calc_hint.configure(text="✓ 计算完成", fg=GREEN)
        self.result_hint.configure(text=str(out), fg=GREEN)
        self.btn_open.configure(state="normal")
        self.btn_saveas.configure(state="normal")
        self.log(f"✓ 结果已自动保存：{out}")
        answer = messagebox.askyesno(
            "计算完成",
            f"汇总完成！\n结果已自动保存到：\n{out}\n\n是否立即打开所在文件夹？",
            parent=self.root,
        )
        if answer:
            self.open_folder()

    def _finish_error(self, msg: str):
        self.running = False
        self.btn_calc.configure(state="normal", text="开始计算")
        self.calc_hint.configure(text="计算失败", fg="#DC2626")
        self.log(f"✗ 计算失败：{msg}")
        messagebox.showerror("计算失败", msg, parent=self.root)

    def open_folder(self):
        if not self.result_path:
            return
        target = self.result_path
        try:
            if sys.platform == "win32":
                os.startfile(str(target.parent))  # noqa: S606
            elif sys.platform == "darwin":
                subprocess.Popen(["open", "-R", str(target)])
            else:
                subprocess.Popen(["xdg-open", str(target.parent)])
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("打开失败", str(exc), parent=self.root)

    def save_as(self):
        if not self.result_path:
            return
        path = filedialog.asksaveasfilename(
            parent=self.root,
            title="另存结果",
            initialdir=str(self.result_path.parent),
            initialfile=self.result_path.name,
            defaultextension=".xlsx",
            filetypes=[("Excel 工作簿", "*.xlsx")],
        )
        if not path:
            return
        try:
            shutil.copyfile(self.result_path, path)
        except Exception as exc:  # noqa: BLE001
            messagebox.showerror("另存失败", str(exc), parent=self.root)
            return
        self.result_hint.configure(text=f"已另存：{path}", fg=GREEN)
        self.log(f"✓ 结果已另存为：{path}")
        messagebox.showinfo("另存成功", f"结果已保存到：\n{path}", parent=self.root)

    @staticmethod
    def _default_save_dir() -> str:
        desktop = Path.home() / "Desktop"
        if desktop.exists():
            return str(desktop)
        return str(Path.home())


def main():
    if sys.platform == "win32":
        try:
            import ctypes

            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
