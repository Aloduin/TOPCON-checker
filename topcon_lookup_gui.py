"""Tkinter desktop interface for TOPCON patient/image lookup."""

from __future__ import annotations

import os
import queue
import threading
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from topcon_lookup import (
    CopyResult,
    LookupResult,
    copy_selected_images,
    load_ids_file,
    lookup_patient_ids,
    parse_id_text,
    write_feedback_workbook,
)


class LookupApp(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=16)
        self.master = master
        self.pack(fill="both", expand=True)
        project_dir = Path.cwd()
        self.data_dir = tk.StringVar(
            value=str((project_dir / "data").resolve())
            if (project_dir / "data").is_dir()
            else ""
        )
        self.image_root = tk.StringVar()
        self.output_root = tk.StringVar(value=str((project_dir / "output").resolve()))
        self.allow_leading_zero = tk.BooleanVar(value=False)
        self.status_text = tk.StringVar(value="请选择全量数据目录并输入患者 ID。")
        self.lookup_result: LookupResult | None = None
        self.copy_result: CopyResult | None = None
        self.last_workbook: Path | None = None
        self.worker_queue: queue.Queue[tuple[str, object]] = queue.Queue()
        self.detail_paths: dict[str, str] = {}
        self._build_ui()

    def _build_ui(self) -> None:
        title = ttk.Label(self, text="TOPCON 数据核查与图片提取", style="Title.TLabel")
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 12))

        self._path_row(1, "全量数据目录", self.data_dir, self._choose_data_dir)
        self._path_row(2, "图片根目录", self.image_root, self._choose_image_root)
        ttk.Label(self, text="留空时与全量数据目录相同", foreground="#666666").grid(
            row=3, column=1, sticky="w", pady=(0, 6)
        )
        self._path_row(4, "结果输出目录", self.output_root, self._choose_output_root)

        id_frame = ttk.LabelFrame(self, text="患者 ID（可从 Excel 直接粘贴）", padding=8)
        id_frame.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(10, 8))
        id_frame.columnconfigure(0, weight=1)
        id_frame.rowconfigure(0, weight=1)
        self.id_text = tk.Text(id_frame, height=8, wrap="word", undo=True)
        self.id_text.grid(row=0, column=0, sticky="nsew")
        id_scroll = ttk.Scrollbar(id_frame, orient="vertical", command=self.id_text.yview)
        id_scroll.grid(row=0, column=1, sticky="ns")
        self.id_text.configure(yscrollcommand=id_scroll.set)
        ttk.Button(id_frame, text="从 TXT/CSV/XLSX 导入", command=self._load_ids).grid(
            row=1, column=0, sticky="w", pady=(8, 0)
        )

        options = ttk.Frame(self)
        options.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        ttk.Checkbutton(
            options,
            text="患者号未精确命中时，允许忽略前导零的唯一匹配",
            variable=self.allow_leading_zero,
        ).pack(side="left")
        ttk.Label(
            options,
            text="（默认关闭，防止误配）",
            foreground="#9C6500",
        ).pack(side="left", padx=(4, 0))

        action_frame = ttk.Frame(self)
        action_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self.run_button = ttk.Button(
            action_frame,
            text="查询、拷贝图片并生成 Excel",
            style="Accent.TButton",
            command=self._start_lookup,
        )
        self.run_button.pack(side="left")
        self.open_report_button = ttk.Button(
            action_frame,
            text="打开反馈表",
            command=self._open_report,
            state="disabled",
        )
        self.open_report_button.pack(side="left", padx=(8, 0))
        self.open_folder_button = ttk.Button(
            action_frame,
            text="打开结果目录",
            command=self._open_result_folder,
            state="disabled",
        )
        self.open_folder_button.pack(side="left", padx=(8, 0))
        self.progress = ttk.Progressbar(action_frame, mode="indeterminate", length=180)
        self.progress.pack(side="right")

        ttk.Label(self, textvariable=self.status_text).grid(
            row=8, column=0, columnspan=3, sticky="w", pady=(0, 8)
        )

        notebook = ttk.Notebook(self)
        notebook.grid(row=9, column=0, columnspan=3, sticky="nsew")
        self.summary_tree = self._make_tree(
            notebook,
            "查询汇总",
            ["查询ID", "状态", "患者ID", "姓名", "索引数", "找到文件", "缺失"],
            [150, 190, 150, 120, 70, 80, 70],
        )
        self.detail_tree = self._make_tree(
            notebook,
            "图片明细（双击打开）",
            ["查询ID", "患者ID", "拍摄日期", "图片名", "文件状态", "拷贝状态"],
            [140, 140, 110, 150, 130, 170],
        )
        self.detail_tree.bind("<Double-1>", self._open_selected_image)

        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)
        self.rowconfigure(9, weight=2)

    def _path_row(
        self,
        row: int,
        label: str,
        variable: tk.StringVar,
        command,
    ) -> None:
        ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=3)
        ttk.Entry(self, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=3)
        ttk.Button(self, text="选择…", command=command).grid(
            row=row, column=2, sticky="e", padx=(8, 0), pady=3
        )

    def _make_tree(
        self,
        notebook: ttk.Notebook,
        title: str,
        columns: list[str],
        widths: list[int],
    ) -> ttk.Treeview:
        frame = ttk.Frame(notebook, padding=4)
        notebook.add(frame, text=title)
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)
        tree = ttk.Treeview(frame, columns=columns, show="headings")
        for column, width in zip(columns, widths):
            tree.heading(column, text=column)
            tree.column(column, width=width, minwidth=60, stretch=True)
        tree.grid(row=0, column=0, sticky="nsew")
        scroll_y = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        scroll_x.grid(row=1, column=0, sticky="ew")
        tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        return tree

    def _choose_data_dir(self) -> None:
        path = filedialog.askdirectory(title="选择医院全量拷贝目录")
        if path:
            self.data_dir.set(path)

    def _choose_image_root(self) -> None:
        path = filedialog.askdirectory(title="选择图片根目录")
        if path:
            self.image_root.set(path)

    def _choose_output_root(self) -> None:
        path = filedialog.askdirectory(title="选择核查结果输出目录")
        if path:
            self.output_root.set(path)

    def _load_ids(self) -> None:
        path = filedialog.askopenfilename(
            title="选择患者 ID 文件",
            filetypes=[
                ("支持的 ID 文件", "*.txt *.csv *.xlsx *.xlsm"),
                ("所有文件", "*.*"),
            ],
        )
        if not path:
            return
        try:
            ids = load_ids_file(path)
        except (OSError, ValueError) as exc:
            messagebox.showerror("导入失败", str(exc), parent=self)
            return
        self.id_text.delete("1.0", "end")
        self.id_text.insert("1.0", "\n".join(ids))
        self.status_text.set(f"已导入 {len(ids)} 个唯一患者 ID。")

    def _start_lookup(self) -> None:
        data_dir = self.data_dir.get().strip()
        output_root = self.output_root.get().strip()
        image_root = self.image_root.get().strip() or data_dir
        ids = parse_id_text(self.id_text.get("1.0", "end"))
        if not data_dir or not Path(data_dir).is_dir():
            messagebox.showwarning("缺少数据目录", "请选择有效的全量数据目录。", parent=self)
            return
        if not output_root:
            messagebox.showwarning("缺少输出目录", "请选择结果输出目录。", parent=self)
            return
        if not ids:
            messagebox.showwarning("缺少患者 ID", "请粘贴或导入至少一个患者 ID。", parent=self)
            return

        self.run_button.configure(state="disabled")
        self.open_report_button.configure(state="disabled")
        self.open_folder_button.configure(state="disabled")
        self.progress.start(12)
        self.status_text.set("正在解析索引、扫描图片并生成反馈，请稍候……")
        worker = threading.Thread(
            target=self._lookup_worker,
            args=(data_dir, image_root, output_root, ids, self.allow_leading_zero.get()),
            daemon=True,
        )
        worker.start()
        self.after(100, self._poll_worker)

    def _lookup_worker(
        self,
        data_dir: str,
        image_root: str,
        output_root: str,
        ids: list[str],
        allow_leading_zero: bool,
    ) -> None:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")
            job_dir = Path(output_root).resolve() / f"TOPCON核查_{timestamp}"
            image_output = job_dir / "图片"
            workbook_path = job_dir / "核查反馈.xlsx"
            result = lookup_patient_ids(
                data_dir,
                ids,
                image_root=image_root,
                allow_leading_zero_match=allow_leading_zero,
            )
            copy_result = copy_selected_images(result, image_output)
            write_feedback_workbook(result, workbook_path, copy_result=copy_result)
            self.worker_queue.put(
                ("ok", (result, copy_result, workbook_path))
            )
        except Exception as exc:  # surfaced to the operator in the UI
            self.worker_queue.put(("error", exc))

    def _poll_worker(self) -> None:
        try:
            status, payload = self.worker_queue.get_nowait()
        except queue.Empty:
            self.after(100, self._poll_worker)
            return
        self.progress.stop()
        self.run_button.configure(state="normal")
        if status == "error":
            self.status_text.set("核查失败。")
            messagebox.showerror("核查失败", str(payload), parent=self)
            return
        result, copy_result, workbook_path = payload
        self.lookup_result = result
        self.copy_result = copy_result
        self.last_workbook = workbook_path
        self.open_report_button.configure(state="normal")
        self.open_folder_button.configure(state="normal")
        self._populate_preview()
        copy_stats = copy_result.stats
        self.status_text.set(
            f"完成：查询 {result.stats['requested_ids']} 个 ID，找到 {result.stats['image_files_found']} 个图片文件，"
            f"复制/已存在 {copy_stats['copied'] + copy_stats['already_present'] + copy_stats['renamed_conflicts']} 个。"
        )
        messagebox.showinfo(
            "核查完成",
            f"反馈表和筛选图片已保存到：\n{workbook_path.parent}",
            parent=self,
        )

    def _populate_preview(self) -> None:
        assert self.lookup_result is not None
        assert self.copy_result is not None
        for tree in (self.summary_tree, self.detail_tree):
            tree.delete(*tree.get_children())
        self.detail_paths.clear()
        for row in self.lookup_result.summaries:
            self.summary_tree.insert(
                "",
                "end",
                values=(
                    row.requested_id,
                    row.match_status,
                    row.patient_id,
                    row.patient_name,
                    row.image_index_count,
                    row.image_files_found,
                    row.image_files_missing,
                ),
            )
        copies = {}
        for record in self.copy_result.records:
            copies.setdefault((record.query_order, record.image_filename), record)
        for row in self.lookup_result.images:
            copy_record = copies.get((row.query_order, row.image_filename))
            item = self.detail_tree.insert(
                "",
                "end",
                values=(
                    row.requested_id,
                    row.patient_id,
                    row.capture_date,
                    row.image_filename,
                    row.file_status,
                    copy_record.copy_status if copy_record else "未拷贝",
                ),
            )
            if copy_record and copy_record.destination_path:
                self.detail_paths[item] = copy_record.destination_path
            elif row.image_path and "\n" not in row.image_path:
                self.detail_paths[item] = row.image_path

    def _open_selected_image(self, _event=None) -> None:
        selected = self.detail_tree.selection()
        if not selected:
            return
        path = self.detail_paths.get(selected[0])
        if not path or not Path(path).is_file():
            messagebox.showwarning("无法打开", "该记录没有可打开的图片文件。", parent=self)
            return
        os.startfile(path)

    def _open_report(self) -> None:
        if self.last_workbook and self.last_workbook.is_file():
            os.startfile(self.last_workbook)

    def _open_result_folder(self) -> None:
        if self.last_workbook and self.last_workbook.parent.is_dir():
            os.startfile(self.last_workbook.parent)


def main() -> None:
    try:
        import ctypes

        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass
    root = tk.Tk()
    root.title("TOPCON 数据核查与图片提取")
    root.geometry("1120x780")
    root.minsize(900, 650)
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    style.configure("Title.TLabel", font=("Microsoft YaHei", 16, "bold"))
    style.configure("Accent.TButton", font=("Microsoft YaHei", 10, "bold"))
    LookupApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
