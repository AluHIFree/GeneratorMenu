"""Графический интерфейс генератора ежедневных меню."""

from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from generator import (
    MODE_TITLES,
    MenuMode,
    default_data_dir,
    default_template_path,
    find_default_files,
    generate_menus,
)

MONTH_LABELS = [
    (1, "Январь"),
    (2, "Февраль"),
    (3, "Март"),
    (4, "Апрель"),
    (5, "Май"),
    (6, "Июнь"),
    (7, "Июль"),
    (8, "Август"),
    (9, "Сентябрь"),
    (10, "Октябрь"),
    (11, "Ноябрь"),
    (12, "Декабрь"),
]

MODE_BUTTONS = {
    "Завтраки": "breakfast",
    "Завтраки и обеды": "school",
    "Интернат": "internat",
}
MODE_HINTS = {
    "breakfast": "Только завтрак (обед в файле пустой)",
    "school": "Завтрак и обед",
    "internat": "Завтрак, завтрак 2, обед, полдник, ужин, ужин 2",
}


class App(ctk.CTk):
    def __init__(self, start_dir: Path | None = None) -> None:
        super().__init__()
        self.title("Генератор ежедневных меню")
        self.geometry("820x660")
        self.minsize(720, 580)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self._busy = False
        self._mode: MenuMode = "school"
        self._start_dir = start_dir or default_data_dir("school")

        self.tm_var = tk.StringVar()
        self.kp_var = tk.StringVar()
        self.out_var = tk.StringVar()
        self.overwrite_var = tk.BooleanVar(value=True)
        self.month_vars = {m: tk.BooleanVar(value=True) for m, _ in MONTH_LABELS}

        self._build()
        self._apply_mode("school")

    def _build(self) -> None:
        pad = {"padx": 12, "pady": 6}
        root = ctk.CTkFrame(self)
        root.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(
            root, text="Генератор ежедневных меню", font=ctk.CTkFont(size=20, weight="bold")
        ).pack(anchor="w", **pad)

        mode_row = ctk.CTkFrame(root, fg_color="transparent")
        mode_row.pack(fill="x", **pad)
        ctk.CTkLabel(mode_row, text="Режим:").pack(side="left")
        self.mode_seg = ctk.CTkSegmentedButton(
            mode_row,
            values=list(MODE_BUTTONS.keys()),
            command=self._on_mode_change,
        )
        self.mode_seg.set("Завтраки и обеды")
        self.mode_seg.pack(side="left", padx=(10, 0))
        self.mode_hint = ctk.CTkLabel(mode_row, text="", text_color="gray")
        self.mode_hint.pack(side="left", padx=(14, 0))

        self._path_row(root, "Типовое меню (tm*-sm.xlsx)", self.tm_var, self._browse_tm)
        self._path_row(root, "Календарь питания (kp*.xlsx)", self.kp_var, self._browse_kp)
        self._path_row(root, "Папка для ежедневных меню", self.out_var, self._browse_out)

        months_frame = ctk.CTkFrame(root)
        months_frame.pack(fill="x", **pad)
        ctk.CTkLabel(months_frame, text="Месяцы:").grid(row=0, column=0, sticky="w", padx=8, pady=6)
        btns = ctk.CTkFrame(months_frame, fg_color="transparent")
        btns.grid(row=0, column=1, sticky="e", padx=8)
        ctk.CTkButton(btns, text="Все", width=70, command=lambda: self._set_months(True)).pack(
            side="left", padx=2
        )
        ctk.CTkButton(btns, text="Сбросить", width=90, command=lambda: self._set_months(False)).pack(
            side="left", padx=2
        )

        grid = ctk.CTkFrame(months_frame, fg_color="transparent")
        grid.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=(0, 8))
        for i, (num, label) in enumerate(MONTH_LABELS):
            ctk.CTkCheckBox(grid, text=label, variable=self.month_vars[num], width=100).grid(
                row=i // 6, column=i % 6, sticky="w", padx=4, pady=2
            )

        opts = ctk.CTkFrame(root, fg_color="transparent")
        opts.pack(fill="x", **pad)
        ctk.CTkCheckBox(opts, text="Перезаписывать существующие файлы", variable=self.overwrite_var).pack(
            side="left"
        )

        actions = ctk.CTkFrame(root, fg_color="transparent")
        actions.pack(fill="x", **pad)
        self.run_btn = ctk.CTkButton(actions, text="Сгенерировать", command=self._start_generate)
        self.run_btn.pack(side="left")
        self.status = ctk.CTkLabel(actions, text="")
        self.status.pack(side="left", padx=12)

        ctk.CTkLabel(root, text="Журнал:").pack(anchor="w", padx=12, pady=(8, 0))
        self.log_box = ctk.CTkTextbox(root, height=220)
        self.log_box.pack(fill="both", expand=True, padx=12, pady=(4, 12))

    def _path_row(self, parent, label, var, browse_cmd) -> None:
        frame = ctk.CTkFrame(parent)
        frame.pack(fill="x", padx=12, pady=4)
        ctk.CTkLabel(frame, text=label).pack(anchor="w", padx=8, pady=(6, 0))
        row = ctk.CTkFrame(frame, fg_color="transparent")
        row.pack(fill="x", padx=8, pady=(2, 8))
        ctk.CTkEntry(row, textvariable=var).pack(side="left", fill="x", expand=True)
        ctk.CTkButton(row, text="Обзор…", width=90, command=browse_cmd).pack(side="left", padx=(8, 0))

    def _on_mode_change(self, value: str) -> None:
        mode = MODE_BUTTONS.get(value, "school")
        self._apply_mode(mode)  # type: ignore[arg-type]

    def _apply_mode(self, mode: MenuMode) -> None:
        self._mode = mode
        data_dir = default_data_dir(mode)
        tm, kp = find_default_files(data_dir)
        self.tm_var.set(str(tm) if tm else "")
        self.kp_var.set(str(kp) if kp else "")
        self.out_var.set(str(data_dir / "output"))
        self.mode_hint.configure(text=MODE_HINTS.get(mode, ""))
        self._log(f"Режим: {MODE_TITLES.get(mode, mode)} — каталог {data_dir}")

    def _set_months(self, value: bool) -> None:
        for var in self.month_vars.values():
            var.set(value)

    def _browse_tm(self) -> None:
        path = filedialog.askopenfilename(
            title="Типовое меню",
            initialdir=self._initial_dir(self.tm_var),
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if path:
            self.tm_var.set(path)

    def _browse_kp(self) -> None:
        path = filedialog.askopenfilename(
            title="Календарь питания",
            initialdir=self._initial_dir(self.kp_var),
            filetypes=[("Excel", "*.xlsx"), ("Все файлы", "*.*")],
        )
        if path:
            self.kp_var.set(path)

    def _browse_out(self) -> None:
        path = filedialog.askdirectory(
            title="Папка для ежедневных меню",
            initialdir=self._initial_dir(self.out_var),
        )
        if path:
            self.out_var.set(path)

    def _initial_dir(self, var: tk.StringVar) -> str:
        raw = var.get().strip()
        if raw:
            p = Path(raw)
            return str(p if p.is_dir() else p.parent)
        return str(default_data_dir(self._mode))

    def _log(self, msg: str) -> None:
        self.log_box.insert("end", msg + "\n")
        self.log_box.see("end")

    def _start_generate(self) -> None:
        if self._busy:
            return

        tm = Path(self.tm_var.get().strip())
        kp = Path(self.kp_var.get().strip())
        out = Path(self.out_var.get().strip())
        template = default_template_path(self._mode)
        months = {m for m, var in self.month_vars.items() if var.get()}

        if not tm.is_file():
            messagebox.showerror("Ошибка", "Укажите файл типового меню (tm*-sm.xlsx)")
            return
        if not kp.is_file():
            messagebox.showerror("Ошибка", "Укажите файл календаря питания (kp*.xlsx)")
            return
        if not template.is_file():
            messagebox.showerror("Ошибка", f"Не найден шаблон:\n{template}")
            return
        if not months:
            messagebox.showerror("Ошибка", "Выберите хотя бы один месяц")
            return

        self._busy = True
        self.run_btn.configure(state="disabled")
        self.status.configure(text="Идёт генерация…")
        self.log_box.delete("1.0", "end")

        def worker() -> None:
            try:
                result = generate_menus(
                    tm,
                    kp,
                    out,
                    template,
                    mode=self._mode,
                    months=months,
                    overwrite=self.overwrite_var.get(),
                    log=lambda m: self.after(0, self._log, m),
                )
            except Exception as exc:  # noqa: BLE001
                self.after(0, self._done, False, str(exc))
                return
            summary = (
                f"Создано: {len(result.created)}, "
                f"пропущено: {len(result.skipped)}, "
                f"ошибок: {len(result.errors)}"
            )
            self.after(0, self._done, not result.errors, summary)

        threading.Thread(target=worker, daemon=True).start()

    def _done(self, ok: bool, message: str) -> None:
        self._busy = False
        self.run_btn.configure(state="normal")
        self.status.configure(text=message)
        if ok:
            messagebox.showinfo("Готово", message)
        else:
            messagebox.showwarning("Завершено с ошибками", message)


def run_app(start_dir: Path | None = None) -> None:
    app = App(start_dir=start_dir)
    app.mainloop()
