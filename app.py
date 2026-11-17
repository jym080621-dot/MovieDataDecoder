"""MovieDataDecoder desktop application."""
from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from decoder import NO_DETAILS, VERSION, convert, discover


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"MovieDataDecoder v{VERSION}")
        self.geometry("760x485")
        self.minsize(640, 450)
        self.protocol("WM_DELETE_WINDOW", self.close)
        self.configure(background="#EEF3F8")
        self.option_add("*Font", ("Microsoft YaHei UI", 10))
        theme = ttk.Style(self)
        theme.theme_use("clam")
        theme.configure("TFrame", background="#EEF3F8")
        theme.configure("TLabel", background="#EEF3F8", foreground="#223449")
        theme.configure("TButton", padding=(10, 6))
        theme.configure("Primary.TButton", background="#245DA1", foreground="white", padding=(15, 8))
        theme.map("Primary.TButton", background=[("disabled", "#AAB8C8"), ("active", "#1D4D88")])

        self.results = []
        self.work = queue.Queue()
        self.busy = False
        self.written = None
        self.origin = None
        self.source = tk.StringVar(value="尚未选择导出结果")
        self.note = tk.StringVar(value="选择影院明细 CSV，或包含该文件的 ZIP / 文件夹。")
        self.status = tk.StringVar(value="等待选择文件")

        main = ttk.Frame(self, padding=26)
        main.pack(fill="both", expand=True)
        main.columnconfigure(0, weight=1)
        ttk.Label(main, text="MovieDataDecoder", font=("Microsoft YaHei UI", 24, "bold")).grid(sticky="w")
        ttk.Label(main, text="每家影院一行，查看这部电影的票房", font=("Microsoft YaHei UI", 11)).grid(sticky="w", pady=(3, 20))
        toolbar = ttk.Frame(main)
        toolbar.grid(sticky="ew")
        self.pick_file = ttk.Button(toolbar, text="选择文件…", command=self.select_file)
        self.pick_file.pack(side="left")
        self.pick_folder = ttk.Button(toolbar, text="选择文件夹…", command=self.select_folder)
        self.pick_folder.pack(side="left", padx=10)
        ttk.Label(main, textvariable=self.source, wraplength=600).grid(sticky="w", pady=(11, 5))
        self.choice = ttk.Combobox(main, state="disabled", width=45)
        self.choice.grid(sticky="ew")
        self.choice.bind("<<ComboboxSelected>>", self.describe)
        ttk.Label(main, textvariable=self.note, wraplength=600).grid(sticky="w", pady=(8, 10))
        ttk.Separator(main).grid(sticky="ew", pady=16)
        actions = ttk.Frame(main)
        actions.grid(sticky="ew")
        self.generate = ttk.Button(actions, text="导出逐影院票房…", style="Primary.TButton", command=self.start, state="disabled")
        self.generate.pack(side="left")
        self.open_button = ttk.Button(actions, text="打开 Excel", command=self.open_report, state="disabled")
        self.open_button.pack(side="left", padx=10)
        self.progress = ttk.Progressbar(main, mode="indeterminate")
        self.progress.grid(pady=(15, 5), sticky="ew")
        ttk.Label(main, textvariable=self.status, wraplength=600).grid(sticky="w")
        self.timer = self.after(100, self.poll)

    def select_file(self):
        path = filedialog.askopenfilename(parent=self, title="选择电影的影院明细", filetypes=[
            ("影院明细或压缩包", "*.csv *.zip"), ("影院明细 CSV", "*_cinema_days.csv"), ("ZIP 压缩包", "*.zip"),
        ])
        if path:
            self.load(Path(path))

    def select_folder(self):
        path = filedialog.askdirectory(parent=self, title="选择 MovieData 结果文件夹", mustexist=True)
        if path:
            self.load(Path(path))

    def load(self, path):
        try:
            items = discover(path)
            if not items:
                raise ValueError(NO_DETAILS)
        except Exception as exc:
            messagebox.showerror("无法读取结果", str(exc), parent=self)
            return
        self.results = items
        self.origin = path
        self.source.set(str(path))
        self.choice.configure(values=[item.location for item in items], state="readonly")
        self.choice.current(0)
        self.status.set(f"找到 {len(items):,} 份电影影院明细")
        self.describe()

    def describe(self, _event=None):
        selected = self.choice.current()
        if selected < 0:
            return
        self.note.set("Excel 会按已确认票房排序。区间未核查齐的影院保留已确认金额，完整区间票房留空。")
        self.generate.configure(state="normal")

    def start(self):
        if self.busy or self.choice.current() < 0:
            return
        data = self.results[self.choice.current()]
        folder = self.origin if self.origin.is_dir() else self.origin.parent
        path = filedialog.asksaveasfilename(parent=self, title="保存 Excel 表格", initialdir=folder,
                                            initialfile=data.filename, defaultextension=".xlsx",
                                            filetypes=[("Excel 工作簿", "*.xlsx")])
        if path:
            self.begin_conversion(data, Path(path))

    def begin_conversion(self, data, output):
        self.busy = True
        for control in (self.pick_file, self.pick_folder, self.choice, self.generate, self.open_button):
            control.configure(state="disabled")
        self.status.set("正在转换…")
        self.progress.start(15)

        def worker():
            try:
                rows = convert(data, output, emit=lambda text: self.work.put(("progress", text)))
                self.work.put(("done", rows, output))
            except Exception as exc:
                self.work.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def poll(self):
        try:
            while True:
                event, *info = self.work.get_nowait()
                if event == "progress":
                    self.status.set(info[0])
                    continue
                self.progress.stop()
                self.progress.configure(value=0)
                self.busy = False
                self.pick_file.configure(state="normal")
                self.pick_folder.configure(state="normal")
                self.choice.configure(state="readonly")
                self.describe()
                if event == "error":
                    self.status.set("转换失败，请检查源文件或输出路径。")
                    messagebox.showerror("转换失败", info[0], parent=self)
                else:
                    rows, self.written = info
                    self.status.set(f"已导出 {rows:,} 家影院的票房 · {self.written.name}")
                self.open_button.configure(state="normal" if self.written else "disabled")
        except queue.Empty:
            pass
        self.timer = self.after(100, self.poll)

    def open_report(self):
        if self.written:
            try:
                os.startfile(self.written)
            except (AttributeError, OSError) as exc:
                messagebox.showinfo("文件已保存", f"请从此路径打开：\n{self.written}\n\n{exc}", parent=self)

    def close(self):
        if self.busy:
            self.status.set("正在保存工作簿，请等转换完成后关闭。")
            return
        self.after_cancel(self.timer)
        self.destroy()


def main():
    if os.name == "nt":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except (AttributeError, OSError):
            pass
    App().mainloop()


if __name__ == "__main__":
    main()
