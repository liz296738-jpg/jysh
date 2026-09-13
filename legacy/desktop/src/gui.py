"""Tkinter desktop interface for the offline employment pre-review workflow."""

from __future__ import annotations

from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .review_session import ISSUE_CATEGORIES, ReviewController, ReviewSession, filter_results
from .rules import ReviewStage, StudentAuditResult


_NAVY = "#17324D"
_BLUE = "#2E6FA3"
_PAPER = "#F7F9FC"
_RED = "#B33A3A"
_INK = "#1D2A36"
_MUTED = "#64748B"
_STAGE_LABELS = {ReviewStage.INITIAL: "初审", ReviewStage.FINAL: "终审"}
_FILTERS = ("全部异常", "企业信息", "电话", "单位行业", "单位性质", "工作职位")


def application_root() -> Path:
    """Find the writable deployment directory in source and PyInstaller modes."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


class EmploymentPreReviewApp(tk.Tk):
    def __init__(self, project_root: str | Path | None = None) -> None:
        super().__init__()
        self.project_root = Path(project_root) if project_root else application_root()
        self.reference_path = self.project_root / "企业参考库.xlsx"
        self.controller = ReviewController(self.reference_path)
        self.session: ReviewSession | None = None
        self.title("就业信息辅助预审工具")
        self.geometry("1220x760")
        self.minsize(1000, 650)
        self.configure(background=_PAPER)
        self._configure_style()
        self.container = ttk.Frame(self, padding=0, style="Page.TFrame")
        self.container.pack(fill="both", expand=True)
        self.show_home()

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("Page.TFrame", background=_PAPER)
        style.configure("Header.TFrame", background=_NAVY)
        style.configure("Title.TLabel", background=_NAVY, foreground="white", font=("Microsoft YaHei UI", 22, "bold"))
        style.configure("Subtitle.TLabel", background=_NAVY, foreground="#C9D8E6", font=("Microsoft YaHei UI", 10))
        style.configure("Body.TLabel", background=_PAPER, foreground=_INK, font=("Microsoft YaHei UI", 10))
        style.configure("Muted.TLabel", background=_PAPER, foreground=_MUTED, font=("Microsoft YaHei UI", 10))
        style.configure("Section.TLabel", background=_PAPER, foreground=_INK, font=("Microsoft YaHei UI", 14, "bold"))
        style.configure("StatName.TLabel", background="white", foreground=_MUTED, font=("Microsoft YaHei UI", 9))
        style.configure("StatValue.TLabel", background="white", foreground=_NAVY, font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 12, "bold"), padding=(24, 13), background=_BLUE, foreground="white")
        style.map("Primary.TButton", background=[("active", _NAVY)])
        style.configure("Secondary.TButton", font=("Microsoft YaHei UI", 10), padding=(13, 8))
        style.configure("Treeview", font=("Microsoft YaHei UI", 10), rowheight=30, background="white", fieldbackground="white", foreground=_INK)
        style.configure("Treeview.Heading", font=("Microsoft YaHei UI", 10, "bold"), background="#E8EEF5", foreground=_INK, relief="flat")
        style.map("Treeview", background=[("selected", "#DDEBF7")], foreground=[("selected", _INK)])

    def _clear(self) -> None:
        for widget in self.container.winfo_children():
            widget.destroy()

    def _header(self, parent: ttk.Frame, title: str, subtitle: str) -> None:
        header = ttk.Frame(parent, style="Header.TFrame", padding=(42, 25))
        header.pack(fill="x")
        ttk.Label(header, text=title, style="Title.TLabel").pack(anchor="w")
        ttk.Label(header, text=subtitle, style="Subtitle.TLabel").pack(anchor="w", pady=(5, 0))

    def show_home(self) -> None:
        self.session = None
        self._clear()
        self._header(self.container, "就业信息辅助预审工具", "本地离线 · 数据不上传 · 辅助审核")
        body = ttk.Frame(self.container, padding=(80, 60), style="Page.TFrame")
        body.pack(fill="both", expand=True)
        ttk.Label(body, text="开始一份新的预审", style="Section.TLabel").pack(anchor="w")
        status = "已加载" if self.reference_path.exists() else "未找到"
        status_color = _BLUE if self.reference_path.exists() else _RED
        library = ttk.Frame(body, padding=(18, 16), style="Page.TFrame")
        library.pack(anchor="w", fill="x", pady=(22, 32))
        ttk.Label(library, text="参考库状态", style="Muted.TLabel").grid(row=0, column=0, sticky="w")
        ttk.Label(library, text="企业参考库.xlsx", style="Body.TLabel").grid(row=1, column=0, sticky="w", pady=(5, 0))
        label = tk.Label(library, text=status, fg=status_color, bg=_PAPER, font=("Microsoft YaHei UI", 10, "bold"))
        label.grid(row=1, column=1, sticky="w", padx=(18, 0), pady=(5, 0))
        buttons = ttk.Frame(body, style="Page.TFrame")
        buttons.pack(anchor="w")
        ttk.Button(buttons, text="初审预审", style="Primary.TButton", command=lambda: self._choose_and_audit("initial")).grid(row=0, column=0, padx=(0, 18))
        ttk.Button(buttons, text="终审预审", style="Primary.TButton", command=lambda: self._choose_and_audit("final")).grid(row=0, column=1)

    def _choose_and_audit(self, stage: str) -> None:
        selected = filedialog.askopenfilename(
            parent=self,
            title=f"选择{ '初审' if stage == 'initial' else '终审' }Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not selected:
            return
        if not self.reference_path.exists():
            messagebox.showerror("无法审核", "未找到企业参考库.xlsx，请先在程序目录放置参考库。", parent=self)
            return
        try:
            self.session = self.controller.audit(selected, stage)
        except Exception as error:
            messagebox.showerror("无法审核", str(error), parent=self)
            return
        self.show_results()

    def _stat(self, parent: ttk.Frame, column: int, name: str, value: int) -> None:
        card = tk.Frame(parent, bg="white", padx=16, pady=12, highlightthickness=1, highlightbackground="#E4EAF1")
        card.grid(row=0, column=column, sticky="nsew", padx=(0 if column == 0 else 8, 0))
        ttk.Label(card, text=name, style="StatName.TLabel").pack(anchor="w")
        ttk.Label(card, text=str(value), style="StatValue.TLabel").pack(anchor="w", pady=(5, 0))
        parent.columnconfigure(column, weight=1)

    def show_results(self) -> None:
        if self.session is None:
            self.show_home()
            return
        self._clear()
        stage_label = _STAGE_LABELS[self.session.review_stage]
        self._header(self.container, f"{stage_label}预审结果", f"当前审核阶段：{stage_label}  ·  {self.session.input_path.name}")
        body = ttk.Frame(self.container, padding=(26, 20), style="Page.TFrame")
        body.pack(fill="both", expand=True)
        stats = ttk.Frame(body, style="Page.TFrame")
        stats.pack(fill="x")
        self._stat(stats, 0, "本批总数", self.session.total_count)
        self._stat(stats, 1, "正常", self.session.normal_count)
        self._stat(stats, 2, "有问题", self.session.issue_count)
        for offset, category in enumerate(ISSUE_CATEGORIES, start=3):
            self._stat(stats, offset, f"{category}异常", self.session.category_count(category))
        if self.session.review_stage is ReviewStage.FINAL:
            ttk.Label(body, text="本工具仅校验结构化就业数据，就业协议/PDF材料仍需人工审核。", style="Muted.TLabel").pack(anchor="w", pady=(16, 8))

        controls = ttk.Frame(body, style="Page.TFrame")
        controls.pack(fill="x", pady=(16, 8))
        ttk.Label(controls, text="按异常类型查看", style="Body.TLabel").pack(side="left")
        self.filter_var = tk.StringVar(value="全部异常")
        selector = ttk.Combobox(controls, textvariable=self.filter_var, values=_FILTERS, state="readonly", width=14)
        selector.pack(side="left", padx=(10, 0))
        selector.bind("<<ComboboxSelected>>", lambda _event: self._refresh_tree())
        ttk.Button(controls, text="导出审核结果", style="Secondary.TButton", command=self._export).pack(side="right")
        ttk.Button(controls, text="返回首页", style="Secondary.TButton", command=self.show_home).pack(side="right", padx=(0, 8))

        pane = ttk.PanedWindow(body, orient="vertical")
        pane.pack(fill="both", expand=True)
        upper = ttk.Frame(pane, style="Page.TFrame")
        lower = ttk.Frame(pane, style="Page.TFrame")
        pane.add(upper, weight=3)
        pane.add(lower, weight=2)
        columns = ("index", "name", "student_id", "company", "status", "count", "summary")
        self.tree = ttk.Treeview(upper, columns=columns, show="headings", selectmode="browse")
        headings = ("序号", "姓名", "学号", "企业名称", "总体结果", "异常数量", "异常摘要")
        widths = (60, 90, 130, 190, 90, 90, 440)
        for column, heading, width in zip(columns, headings, widths):
            self.tree.heading(column, text=heading)
            self.tree.column(column, width=width, minwidth=60, stretch=column == "summary")
        scrollbar = ttk.Scrollbar(upper, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        self.tree.bind("<<TreeviewSelect>>", self._show_detail)
        ttk.Label(lower, text="异常详情", style="Section.TLabel").pack(anchor="w", pady=(12, 6))
        self.detail = tk.Text(lower, height=8, wrap="word", font=("Microsoft YaHei UI", 10), background="white", foreground=_INK, relief="solid", borderwidth=1, state="disabled")
        self.detail.pack(fill="both", expand=True)
        self._refresh_tree()

    def _visible_results(self) -> tuple[StudentAuditResult, ...]:
        if self.filter_var.get() == "全部异常":
            return filter_results(self.session.results, "全部异常")
        return filter_results(self.session.results, self.filter_var.get())

    def _refresh_tree(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        for result in self._visible_results():
            summary = "；".join(issue.message for issue in result.issues)
            self.tree.insert(
                "",
                "end",
                iid=str(result.row_number),
                values=(
                    result.row_number - 1,
                    result.student_name,
                    result.student_id,
                    result.company_name,
                    "正常" if result.is_ok else "有问题",
                    len(result.issues),
                    summary,
                ),
            )
        self._set_detail("选择一名学生查看异常项目、学生填写值与历史参考值。")

    def _show_detail(self, _event=None) -> None:
        selection = self.tree.selection()
        if not selection:
            return
        row_number = int(selection[0])
        result = next(result for result in self.session.results if result.row_number == row_number)
        if not result.issues:
            self._set_detail("该学生记录正常。")
            return
        lines: list[str] = []
        for issue in result.issues:
            lines.extend(
                [
                    f"异常项目：{issue.field}",
                    f"学生填写值：{issue.student_value or '—'}",
                    f"历史参考值：{issue.reference_value or '—'}",
                    f"异常原因：{issue.message}",
                    "",
                ]
            )
        self._set_detail("\n".join(lines).rstrip())

    def _set_detail(self, value: str) -> None:
        self.detail.configure(state="normal")
        self.detail.delete("1.0", "end")
        self.detail.insert("1.0", value)
        self.detail.configure(state="disabled")

    def _export(self) -> None:
        try:
            output_path = self.controller.export(self.session)
        except Exception as error:
            messagebox.showerror("导出失败", str(error), parent=self)
            return
        messagebox.showinfo("导出完成", f"审核结果已保存至：\n{output_path}", parent=self)


def run() -> None:
    EmploymentPreReviewApp().mainloop()
