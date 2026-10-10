from __future__ import annotations

import json
import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import END, LEFT, RIGHT, filedialog, messagebox
from tkinter import ttk
import tkinter as tk

from tkinterdnd2 import DND_FILES, TkinterDnD

from merger import MergeCancelled, discover_files, is_generated_output, merge_files


APP_NAME = "合并文档 V5"
BG = "#F3F7FA"
CARD = "#FFFFFF"
INK = "#163247"
MUTED = "#6A7F8E"
ACCENT = "#147D92"
ACCENT_DARK = "#0D6274"
LINE = "#C9D9E2"


class MergeApp:
    def __init__(self, root: TkinterDnD.Tk) -> None:
        self.root = root
        self.folder: Path | None = None
        self.dropped_files: list[Path] | None = None
        self.output_path: Path | None = None
        self.cancel_event = threading.Event()
        self.events: queue.Queue[tuple[str, object]] = queue.Queue()
        self.busy = False
        self._progress_animation = 0

        root.title(APP_NAME)
        root.iconbitmap(str(Path(getattr(sys, "_MEIPASS", Path(__file__).parent)) / "assets" / "app_icon.ico"))
        root.geometry("820x760")
        root.minsize(760, 700)
        root.configure(bg=BG)
        root.protocol("WM_DELETE_WINDOW", self._close)

        self._setup_style()
        self._build_ui()
        self.root.after(100, self._poll_events)

    def _setup_style(self) -> None:
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("TFrame", background=BG)
        style.configure("Card.TFrame", background=CARD)
        style.configure("Title.TLabel", background=BG, foreground=INK, font=("Microsoft YaHei UI", 22, "bold"))
        style.configure("Sub.TLabel", background=BG, foreground=MUTED, font=("Microsoft YaHei UI", 10))
        style.configure("Card.TLabel", background=CARD, foreground=INK, font=("Microsoft YaHei UI", 10))
        style.configure("Status.TLabel", background=CARD, foreground=MUTED, font=("Microsoft YaHei UI", 9))
        style.configure(
            "Accent.TButton",
            font=("Microsoft YaHei UI", 10, "bold"),
            foreground="white",
            background=ACCENT,
            padding=(18, 9),
        )
        style.map("Accent.TButton", background=[("active", ACCENT_DARK), ("disabled", "#9BB4BE")])
        style.configure("TButton", font=("Microsoft YaHei UI", 10), padding=(14, 8))
        style.configure("Horizontal.TProgressbar", troughcolor="#DDE8ED", background=ACCENT, thickness=14)

    def _build_ui(self) -> None:
        shell = ttk.Frame(self.root, padding=(28, 22, 28, 22))
        shell.pack(fill="both", expand=True)

        ttk.Label(shell, text=APP_NAME, style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            shell,
            text="拖入多个 .xls 文件，或选择文件夹；读取清单后点击“开始合并”",
            style="Sub.TLabel",
        ).pack(anchor="w", pady=(2, 16))

        card = ttk.Frame(shell, style="Card.TFrame", padding=18)
        card.pack(fill="x")

        self.drop = tk.Canvas(card, height=122, bg="#F7FBFC", highlightthickness=0, cursor="hand2")
        self.drop.pack(fill="x")
        self.drop.bind("<Configure>", self._draw_drop_zone)
        self.drop.bind("<Button-1>", lambda _event: self._choose_folder())
        self.drop.drop_target_register(DND_FILES)
        self.drop.dnd_bind("<<Drop>>", self._on_drop)

        self.folder_var = tk.StringVar(value="尚未选择文件夹")
        ttk.Label(card, textvariable=self.folder_var, style="Status.TLabel").pack(anchor="w", pady=(10, 0))

        buttons = ttk.Frame(card, style="Card.TFrame")
        buttons.pack(fill="x", pady=(14, 0))
        self.choose_button = ttk.Button(buttons, text="选择文件夹", command=self._choose_folder)
        self.choose_button.pack(side=LEFT)
        self.start_button = ttk.Button(
            buttons,
            text="开始合并",
            style="Accent.TButton",
            command=self._start_merge,
            state="disabled",
        )
        self.start_button.pack(side=LEFT, padx=(10, 0))
        self.cancel_button = ttk.Button(buttons, text="取消", command=self._cancel, state="disabled")
        self.cancel_button.pack(side=LEFT, padx=(10, 0))
        self.open_button = ttk.Button(
            buttons,
            text="打开输出文件夹",
            command=self._open_output,
            state="disabled",
        )
        self.open_button.pack(side=RIGHT)

        self.progress_host = ttk.Frame(shell, height=78)
        self.progress_host.pack_propagate(False)
        progress_card = ttk.Frame(self.progress_host, style="Card.TFrame", padding=(18, 14))
        self.status_var = tk.StringVar(value="等待选择文件夹")
        self.count_var = tk.StringVar(value="0 / 0")
        progress_card.columnconfigure(0, weight=1)
        ttk.Label(progress_card, textvariable=self.status_var, style="Card.TLabel").grid(
            row=0,
            column=0,
            sticky="w",
        )
        ttk.Label(progress_card, textvariable=self.count_var, style="Status.TLabel").grid(
            row=0,
            column=1,
            sticky="e",
        )
        self.progress = ttk.Progressbar(progress_card, mode="determinate", maximum=100)
        self.progress.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(10, 0))

        self.log_card = ttk.Frame(shell, style="Card.TFrame", padding=(18, 14))
        self.log_card.pack(fill="both", expand=True, pady=(14, 0))
        log_card = self.log_card
        log_header = ttk.Frame(log_card, style="Card.TFrame")
        log_header.pack(fill="x")
        ttk.Label(log_header, text="运行日志", style="Card.TLabel").pack(side=LEFT)
        ttk.Button(log_header, text="清空日志", command=self._clear_log).pack(side=RIGHT)

        log_body = ttk.Frame(log_card, style="Card.TFrame")
        log_body.pack(fill="both", expand=True, pady=(8, 0))
        scrollbar = ttk.Scrollbar(log_body)
        scrollbar.pack(side=RIGHT, fill="y")
        self.log_text = tk.Text(
            log_body,
            height=12,
            bg="#102A3A",
            fg="#DDECF2",
            insertbackground="white",
            relief="flat",
            padx=12,
            pady=10,
            font=("Consolas", 9),
            wrap="word",
            yscrollcommand=scrollbar.set,
        )
        self.log_text.pack(side=LEFT, fill="both", expand=True)
        scrollbar.configure(command=self.log_text.yview)

    def _draw_drop_zone(self, event: tk.Event) -> None:
        self.drop.delete("all")
        width = max(100, event.width)
        self.drop.create_rectangle(10, 10, width - 10, 112, outline=LINE, width=2, dash=(7, 5))
        self.drop.create_text(
            width / 2,
            45,
            text="拖拽多个.xls文件或文件夹",
            fill=INK,
            font=("Microsoft YaHei UI", 15, "bold"),
        )
        self.drop.create_text(
            width / 2,
            78,
            text="或点击此区域选择文件夹",
            fill=MUTED,
            font=("Microsoft YaHei UI", 10),
        )

    def _timestamp(self) -> str:
        return datetime.now().strftime("%H:%M:%S")

    def _log(self, message: str) -> None:
        self.log_text.insert(END, f"[{self._timestamp()}] {message}\n")
        self.log_text.see(END)

    def _clear_log(self) -> None:
        self.log_text.delete("1.0", END)

    def _show_progress(self) -> None:
        self._progress_animation += 1
        animation = self._progress_animation
        self.progress_host.pack(fill="x", pady=(14, 0), before=self.log_card)
        card = self.progress.master
        card.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.04, relheight=1)

        def expand(step: int) -> None:
            if animation != self._progress_animation:
                return
            fraction = step / 12
            width = 0.04 + 0.96 * (1 - (1 - fraction) ** 3)
            card.place_configure(relwidth=width)
            if step < 12:
                self.root.after(18, expand, step + 1)

        self.root.after(18, expand, 1)

    def _hide_progress(self) -> None:
        self._progress_animation += 1
        self.progress.master.place_forget()
        self.progress_host.pack_forget()

    def _choose_folder(self) -> None:
        if self.busy:
            return
        selected = filedialog.askdirectory(title="选择包含 .xls 文件的文件夹")
        if selected:
            self._set_folder(Path(selected))

    def _on_drop(self, event: tk.Event) -> None:
        if self.busy:
            return
        try:
            paths = self.root.tk.splitlist(event.data)
        except tk.TclError:
            paths = [event.data]
        dropped = [Path(path) for path in paths]
        if len(dropped) == 1 and dropped[0].is_dir():
            self._set_folder(dropped[0])
            return
        if not dropped or not all(path.is_file() for path in dropped):
            messagebox.showwarning(APP_NAME, "请拖入一个文件夹，或同一文件夹中的多个 .xls 文件。")
            return
        eligible = [path for path in dropped if path.suffix.lower() == ".xls" and not path.name.startswith("~$")]
        if not eligible:
            messagebox.showwarning(APP_NAME, "拖入的文件中没有 .xls 文件。")
            return
        if len({path.parent for path in eligible}) != 1:
            messagebox.showwarning(APP_NAME, "请拖入同一文件夹中的 .xls 文件。")
            return
        files = sorted(set(eligible), key=lambda path: path.name.casefold())
        self._set_folder(files[0].parent, selected_files=files)
        if len(dropped) > len(eligible):
            self._log(f"已忽略 {len(dropped) - len(eligible)} 个非 .xls 或临时文件")

    def _current_files(self) -> tuple[list[Path], list[Path]]:
        assert self.folder is not None
        if self.dropped_files is None:
            return discover_files(self.folder)
        sources = [path for path in self.dropped_files if not is_generated_output(path)]
        generated = [path for path in self.dropped_files if is_generated_output(path)]
        return sources, generated

    def _set_folder(self, folder: Path, *, selected_files: list[Path] | None = None) -> None:
        self.folder = folder
        self.dropped_files = selected_files
        self.output_path = None
        self.folder_var.set(str(folder) if selected_files is None else f"已拖入 {len(selected_files)} 个 .xls 文件：{folder}")
        self.open_button.configure(state="disabled")
        self._log(f"已选择：{folder}" if selected_files is None else f"已拖入：{len(selected_files)} 个 .xls 文件（{folder}）")
        try:
            files, old_outputs = self._current_files()
        except Exception as exc:
            self._hide_progress()
            self.status_var.set("读取文件夹失败")
            self.count_var.set("0 / 0")
            self.progress.configure(value=0, maximum=1)
            self.start_button.configure(state="disabled")
            self._log(f"读取失败：{exc}")
            return

        total = len(files)
        self.progress.configure(value=0, maximum=max(total, 1))
        self.count_var.set(f"0 / {total}")
        if total:
            self._show_progress()
            self.status_var.set(f"已读取 {total} 个 Excel 文件，请确认后点击“开始合并”")
            self.start_button.configure(state="normal")
            self._log(f"扫描完成：发现 {total} 个可处理的 .xls 文件")
        else:
            self._hide_progress()
            self.status_var.set("未发现可处理的 .xls 文件")
            self.start_button.configure(state="disabled")
            self._log("扫描完成：未发现可处理的 .xls 文件")
        if old_outputs:
            self._log(f"另发现 {len(old_outputs)} 个历史合并结果，开始时将提醒并排除")

    def _start_merge(self) -> None:
        if self.busy or self.folder is None:
            return
        try:
            files, old_outputs = self._current_files()
        except Exception as exc:
            messagebox.showerror(APP_NAME, str(exc))
            return
        if old_outputs:
            names = "\n".join(path.name for path in old_outputs[:5])
            if len(old_outputs) > 5:
                names += f"\n……另有 {len(old_outputs) - 5} 个"
            confirmed = messagebox.askyesno(
                APP_NAME,
                f"检测到以前生成的合并结果：\n\n{names}\n\n继续时会自动排除这些文件，是否继续？",
            )
            if not confirmed:
                self._log("用户取消：检测到历史合并结果")
                return
        if not files:
            messagebox.showwarning(APP_NAME, "没有可处理的 .xls 文件。")
            return

        self.busy = True
        self.cancel_event.clear()
        self.output_path = None
        self.progress.configure(value=0, maximum=len(files))
        self.count_var.set(f"0 / {len(files)}")
        self.status_var.set("正在合并")
        self.choose_button.configure(state="disabled")
        self.start_button.configure(state="disabled")
        self.cancel_button.configure(state="normal")
        self.open_button.configure(state="disabled")
        self._log(f"开始处理 {len(files)} 个文件（按文件名从小到大）")

        threading.Thread(target=self._worker, args=(files,), daemon=True).start()

    def _worker(self, files: list[Path]) -> None:
        assert self.folder is not None

        def progress(done: int, total: int, _path: Path, ok: bool, message: str) -> None:
            self.events.put(("progress", (done, total, ok, message)))

        try:
            result = merge_files(files, self.folder, cancel=self.cancel_event, progress=progress)
        except MergeCancelled as exc:
            self.events.put(("cancelled", str(exc)))
        except Exception as exc:
            self.events.put(("error", str(exc) or exc.__class__.__name__))
        else:
            self.events.put(("done", result))

    def _poll_events(self) -> None:
        while True:
            try:
                kind, payload = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == "progress":
                done, total, _ok, message = payload
                self.progress.configure(value=done, maximum=total)
                self.count_var.set(f"{done} / {total}")
                self._log(message)
            elif kind == "done":
                result = payload
                self.output_path = result.output_path
                self.status_var.set("合并完成")
                self._log(
                    f"完成：合并 {result.merged_files} 个，跳过 {result.skipped_files} 个，"
                    f"共 {result.data_rows} 条明细"
                )
                self._log(f"输出：{result.output_path}")
                self._finish(enable_output=True)
            elif kind == "cancelled":
                self.status_var.set("已取消")
                self._log(str(payload))
                self._finish(enable_output=False)
            elif kind == "error":
                self.status_var.set("处理失败")
                self._log(f"失败：{payload}")
                messagebox.showerror(APP_NAME, str(payload))
                self._finish(enable_output=False)
        self.root.after(100, self._poll_events)

    def _finish(self, *, enable_output: bool) -> None:
        self.busy = False
        self.choose_button.configure(state="normal")
        self.start_button.configure(state="normal" if self.folder else "disabled")
        self.cancel_button.configure(state="disabled")
        self.open_button.configure(state="normal" if enable_output else "disabled")

    def _cancel(self) -> None:
        if self.busy:
            self.cancel_event.set()
            self.cancel_button.configure(state="disabled")
            self.status_var.set("正在取消")
            self._log("正在取消，请稍候……")

    def _open_output(self) -> None:
        if self.output_path and self.output_path.exists():
            subprocess.Popen(["explorer.exe", "/select,", str(self.output_path)])

    def _close(self) -> None:
        if self.busy and not messagebox.askyesno(APP_NAME, "任务仍在进行，确定退出吗？"):
            return
        self.cancel_event.set()
        self.root.destroy()


def run_self_test(folder: Path, report_path: Path) -> int:
    try:
        files, _generated = discover_files(folder)
        result = merge_files(files, folder)
        payload = {
            "ok": True,
            "output": str(result.output_path),
            "total_files": result.total_files,
            "merged_files": result.merged_files,
            "skipped_files": result.skipped_files,
            "data_rows": result.data_rows,
            "issues": [{"file": issue.path.name, "reason": issue.reason} for issue in result.issues],
        }
        code = 0
    except Exception as exc:
        payload = {"ok": False, "error": str(exc) or exc.__class__.__name__}
        code = 1
    report_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return code


def main() -> int:
    if len(sys.argv) == 4 and sys.argv[1] == "--self-test":
        return run_self_test(Path(sys.argv[2]), Path(sys.argv[3]))
    root = TkinterDnD.Tk()
    MergeApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
