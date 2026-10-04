"""WordMerge V0.1.1 - drag multiple DOCX files and merge them."""

from __future__ import annotations

import importlib
import os
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from uuid import uuid4
from tkinter import (
    BOTH,
    END,
    LEFT,
    RIGHT,
    VERTICAL,
    BooleanVar,
    StringVar,
    Tk,
    filedialog,
    messagebox,
    ttk,
)


APP_NAME = "Word組合工"
VERSION = "V0.1.1"
SUPPORTED_EXTENSIONS = {".docx"}


def ensure_dependencies(root: Tk) -> bool:
    """Install optional third-party packages after getting user approval."""
    requirements = {
        "docx": "python-docx>=1.1.0",
        "docxcompose": "docxcompose>=1.4.0",
        "tkinterdnd2": "tkinterdnd2>=0.4.2",
    }
    missing = []
    for module_name, package_name in requirements.items():
        try:
            importlib.import_module(module_name)
        except ImportError:
            missing.append(package_name)

    if not missing:
        return True

    package_text = "\n".join(f"• {name}" for name in missing)
    approved = messagebox.askyesno(
        "需要安裝元件",
        "第一次使用需要安裝下列 Python 元件：\n\n"
        f"{package_text}\n\n是否現在自動安裝？",
        parent=root,
    )
    if not approved:
        return False

    try:
        creation_flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        subprocess.run(
            [sys.executable, "-m", "pip", "install", *missing],
            check=True,
            capture_output=True,
            text=True,
            creationflags=creation_flags,
        )
        importlib.invalidate_caches()
        return True
    except subprocess.CalledProcessError as exc:
        detail = (exc.stderr or exc.stdout or str(exc)).strip()
        messagebox.showerror(
            "安裝失敗",
            "無法安裝必要元件。請確認網路連線，或在命令提示字元執行：\n\n"
            f"{sys.executable} -m pip install {' '.join(missing)}\n\n{detail}",
            parent=root,
        )
        return False


def merge_docx_files(
    source_files: list[str], output_file: str, page_break_between: bool
) -> None:
    """Merge DOCX files with docxcompose, preserving styles and media."""
    from docx import Document
    from docxcompose.composer import Composer

    if not source_files:
        raise ValueError("沒有可合併的來源文件。")

    resolved_sources = [Path(path).resolve(strict=True) for path in source_files]
    output_path = Path(output_file).resolve()
    if output_path.suffix.lower() != ".docx":
        raise ValueError("輸出檔必須使用 .docx 副檔名。")
    if output_path in resolved_sources:
        raise ValueError("輸出檔不能覆蓋來源文件。")

    master = Document(str(resolved_sources[0]))
    composer = Composer(master)

    for path in resolved_sources[1:]:
        if page_break_between:
            master.add_page_break()
        composer.append(Document(str(path)))

    output_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = output_path.with_name(
        f".{output_path.stem}_merging_{uuid4().hex}{output_path.suffix}"
    )
    try:
        composer.save(str(temporary))
        os.replace(temporary, output_path)
    finally:
        if temporary.exists():
            try:
                temporary.unlink()
            except OSError:
                pass


class WordMergeApp:
    def __init__(self, root: Tk, dnd_files_token: str | None = None) -> None:
        self.root = root
        self.files: list[str] = []
        self.dnd_files_token = dnd_files_token
        self.page_break_var = BooleanVar(value=True)
        self.status_var = StringVar(value="請拖曳 Word 檔案到下方區域")
        self.is_busy = False

        self._configure_window()
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

    def _configure_window(self) -> None:
        self.root.title(f"{APP_NAME}  {VERSION}")
        self.root.geometry("780x560")
        self.root.minsize(640, 450)
        self.root.option_add("*Font", ("Microsoft JhengHei UI", 10))

        style = ttk.Style(self.root)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Microsoft JhengHei UI", 18, "bold"))
        style.configure("Hint.TLabel", foreground="#56616f")
        style.configure("Drop.TFrame", background="#f4f7f9", relief="solid", borderwidth=1)
        style.configure("Accent.TButton", font=("Microsoft JhengHei UI", 11, "bold"))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill=BOTH, expand=True)

        top = ttk.Frame(outer)
        top.pack(fill="x")
        ttk.Label(top, text=APP_NAME, style="Title.TLabel").pack(side=LEFT)
        ttk.Label(top, text=VERSION, foreground="#6b7280").pack(side=RIGHT, anchor="s")

        ttk.Label(
            outer,
            text="依清單順序將多個 .docx 文件合併成一個 Word 檔",
            style="Hint.TLabel",
        ).pack(fill="x", pady=(3, 14))

        toolbar = ttk.Frame(outer)
        toolbar.pack(fill="x", pady=(0, 8))
        self.add_button = ttk.Button(toolbar, text="＋ 加入檔案", command=self.add_files)
        self.add_button.pack(side=LEFT)
        self.remove_button = ttk.Button(toolbar, text="移除", command=self.remove_selected)
        self.remove_button.pack(side=LEFT, padx=(8, 0))
        self.clear_button = ttk.Button(toolbar, text="清空", command=self.clear_files)
        self.clear_button.pack(side=LEFT, padx=(8, 0))
        self.up_button = ttk.Button(toolbar, text="▲ 上移", command=lambda: self.move_selected(-1))
        self.up_button.pack(side=RIGHT)
        self.down_button = ttk.Button(toolbar, text="▼ 下移", command=lambda: self.move_selected(1))
        self.down_button.pack(side=RIGHT, padx=(0, 8))

        list_frame = ttk.Frame(outer, style="Drop.TFrame", padding=8)
        list_frame.pack(fill=BOTH, expand=True)

        self.tree = ttk.Treeview(
            list_frame,
            columns=("number", "name", "folder"),
            show="headings",
            selectmode="extended",
        )
        self.tree.heading("number", text="順序")
        self.tree.heading("name", text="檔名")
        self.tree.heading("folder", text="資料夾")
        self.tree.column("number", width=58, minwidth=52, anchor="center", stretch=False)
        self.tree.column("name", width=280, minwidth=160)
        self.tree.column("folder", width=360, minwidth=180)

        scroll = ttk.Scrollbar(list_frame, orient=VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(side=LEFT, fill=BOTH, expand=True)
        scroll.pack(side=RIGHT, fill="y")

        self.tree.bind("<Delete>", lambda _event: self.remove_selected())
        self.tree.bind("<Control-a>", self.select_all)

        if self.dnd_files_token:
            try:
                self.tree.drop_target_register(self.dnd_files_token)
                self.tree.dnd_bind("<<Drop>>", self.on_drop)
            except Exception:
                pass

        options = ttk.Frame(outer)
        options.pack(fill="x", pady=(10, 0))
        ttk.Checkbutton(
            options,
            text="每份文件從新的一頁開始",
            variable=self.page_break_var,
        ).pack(side=LEFT)
        self.count_label = ttk.Label(options, text="共 0 份文件")
        self.count_label.pack(side=RIGHT)

        bottom = ttk.Frame(outer)
        bottom.pack(fill="x", pady=(14, 0))
        self.status_label = ttk.Label(bottom, textvariable=self.status_var, style="Hint.TLabel")
        self.status_label.pack(side=LEFT, fill="x", expand=True)
        self.merge_button = ttk.Button(
            bottom,
            text="合併並輸出 Word",
            command=self.start_merge,
            style="Accent.TButton",
            padding=(18, 8),
        )
        self.merge_button.pack(side=RIGHT)

    def add_files(self) -> None:
        if self.is_busy:
            return
        paths = filedialog.askopenfilenames(
            parent=self.root,
            title="選取要合併的 Word 檔案",
            filetypes=[("Word 文件", "*.docx"), ("所有檔案", "*.*")],
        )
        if paths:
            self.add_paths(list(paths))

    def on_drop(self, event) -> str:
        if self.is_busy:
            return "break"
        try:
            paths = list(self.root.tk.splitlist(event.data))
        except Exception:
            paths = [event.data]
        self.add_paths(paths)
        return "break"

    def add_paths(self, paths: list[str]) -> None:
        if self.is_busy:
            return
        normalized_existing = {os.path.normcase(os.path.abspath(p)) for p in self.files}
        added = 0
        rejected = []

        for raw_path in paths:
            path = os.path.abspath(os.path.expanduser(str(raw_path).strip().strip("{}")))
            if not os.path.isfile(path) or Path(path).suffix.lower() not in SUPPORTED_EXTENSIONS:
                rejected.append(Path(path).name or path)
                continue
            normalized = os.path.normcase(path)
            if normalized in normalized_existing:
                continue
            self.files.append(path)
            normalized_existing.add(normalized)
            added += 1

        self.refresh_tree()
        if added:
            self.status_var.set(f"已加入 {added} 份文件，可用上移、下移調整合併順序")
        if rejected:
            messagebox.showwarning(
                "部分檔案未載入",
                "目前僅支援 .docx 文件，下列項目已略過：\n\n" + "\n".join(rejected[:12]),
                parent=self.root,
            )

    def refresh_tree(self, selected_indices: list[int] | None = None) -> None:
        self.tree.delete(*self.tree.get_children())
        for index, path in enumerate(self.files):
            file_path = Path(path)
            self.tree.insert(
                "",
                END,
                iid=str(index),
                values=(index + 1, file_path.name, str(file_path.parent)),
            )
        self.count_label.configure(text=f"共 {len(self.files)} 份文件")
        if selected_indices:
            valid = [str(i) for i in selected_indices if 0 <= i < len(self.files)]
            self.tree.selection_set(valid)
            if valid:
                self.tree.see(valid[0])

    def selected_indices(self) -> list[int]:
        return sorted(int(item) for item in self.tree.selection())

    def remove_selected(self) -> None:
        if self.is_busy:
            return
        indices = self.selected_indices()
        if not indices:
            return
        for index in reversed(indices):
            del self.files[index]
        next_index = min(indices[0], len(self.files) - 1)
        self.refresh_tree([next_index] if next_index >= 0 else None)
        self.status_var.set("已移除選取文件")

    def clear_files(self) -> None:
        if self.is_busy:
            return
        if not self.files:
            return
        if messagebox.askyesno("清空清單", "確定要移除清單內全部文件嗎？", parent=self.root):
            self.files.clear()
            self.refresh_tree()
            self.status_var.set("清單已清空")

    def move_selected(self, direction: int) -> None:
        if self.is_busy:
            return
        indices = self.selected_indices()
        if not indices:
            return
        if direction < 0:
            for index in indices:
                if index > 0 and index - 1 not in indices:
                    self.files[index - 1], self.files[index] = self.files[index], self.files[index - 1]
            new_indices = [max(0, i - 1) for i in indices]
        else:
            for index in reversed(indices):
                if index < len(self.files) - 1 and index + 1 not in indices:
                    self.files[index + 1], self.files[index] = self.files[index], self.files[index + 1]
            new_indices = [min(len(self.files) - 1, i + 1) for i in indices]
        self.refresh_tree(new_indices)

    def select_all(self, _event=None) -> str:
        self.tree.selection_set(self.tree.get_children())
        return "break"

    def start_merge(self) -> None:
        if self.is_busy:
            return
        if not self.files:
            messagebox.showinfo("尚未載入文件", "請先拖曳或加入 Word 文件。", parent=self.root)
            return

        default_name = f"合併文件_{datetime.now():%Y%m%d}.docx"
        output_path = filedialog.asksaveasfilename(
            parent=self.root,
            title="儲存合併後的 Word 文件",
            defaultextension=".docx",
            initialfile=default_name,
            filetypes=[("Word 文件", "*.docx")],
        )
        if not output_path:
            return

        source_norm = {os.path.normcase(os.path.abspath(p)) for p in self.files}
        if os.path.normcase(os.path.abspath(output_path)) in source_norm:
            messagebox.showerror("輸出位置錯誤", "輸出檔不能覆蓋清單中的來源文件。", parent=self.root)
            return

        self.set_busy(True)
        self.status_var.set("正在合併文件，請稍候……")
        source_files = list(self.files)
        page_break = self.page_break_var.get()
        threading.Thread(
            target=self.merge_worker,
            args=(source_files, output_path, page_break),
            daemon=True,
        ).start()

    def merge_worker(self, source_files: list[str], output_path: str, page_break: bool) -> None:
        try:
            merge_docx_files(source_files, output_path, page_break)
        except Exception as exc:
            self.root.after(0, self.merge_failed, str(exc))
        else:
            self.root.after(0, self.merge_succeeded, output_path)

    def merge_succeeded(self, output_path: str) -> None:
        self.set_busy(False)
        self.status_var.set(f"合併完成：{Path(output_path).name}")
        open_now = messagebox.askyesno(
            "合併完成",
            f"已輸出：\n{output_path}\n\n是否立即開啟檔案？",
            parent=self.root,
        )
        if open_now:
            try:
                os.startfile(output_path)  # type: ignore[attr-defined]
            except Exception as exc:
                messagebox.showwarning("無法開啟", str(exc), parent=self.root)

    def merge_failed(self, error: str) -> None:
        self.set_busy(False)
        self.status_var.set("合併失敗")
        messagebox.showerror(
            "合併失敗",
            "請確認來源文件未損壞、未被鎖定，且輸出檔未在 Word 中開啟。\n\n"
            f"錯誤內容：{error}",
            parent=self.root,
        )

    def set_busy(self, busy: bool) -> None:
        self.is_busy = busy
        state = "disabled" if busy else "normal"
        for widget in (
            self.add_button,
            self.remove_button,
            self.clear_button,
            self.up_button,
            self.down_button,
            self.merge_button,
        ):
            widget.configure(state=state)

    def on_close(self) -> None:
        if self.is_busy:
            messagebox.showinfo(
                "正在合併",
                "文件仍在合併中，完成後即可關閉程式。",
                parent=self.root,
            )
            return
        self.root.destroy()


def main() -> None:
    bootstrap = Tk()
    bootstrap.withdraw()
    if not ensure_dependencies(bootstrap):
        bootstrap.destroy()
        return
    bootstrap.destroy()

    try:
        from tkinterdnd2 import DND_FILES, TkinterDnD

        root = TkinterDnD.Tk()
        dnd_token = DND_FILES
    except Exception:
        root = Tk()
        dnd_token = None

    WordMergeApp(root, dnd_token)
    root.mainloop()


if __name__ == "__main__":
    main()
