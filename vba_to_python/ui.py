from __future__ import annotations

import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
from typing import Callable, Optional

from .actions import (
    change_value,
    io_change,
    list_excel_files,
    select_folder_path,
    update_files,
    value_find,
)
from .constants import EXCEL_EXTENSIONS


class VBAConversionApp:
    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("VBA to Python Toolkit")

        self.control_path: Optional[Path] = None
        self.status_var = tk.StringVar(value="컨트롤 통합문서를 먼저 선택하세요.")
        self.control_var = tk.StringVar(value="(선택되지 않음)")

        self.action_buttons: list[tk.Button] = []

        self._build_layout()

    # ------------------------------------------------------------------ UI
    def _build_layout(self) -> None:
        padding = {"padx": 10, "pady": 5}

        header = tk.Label(self.root, text="VBA 작업을 Python으로 실행합니다.", font=("Segoe UI", 12, "bold"))
        header.grid(row=0, column=0, columnspan=2, sticky="w", **padding)

        control_label = tk.Label(self.root, text="컨트롤 통합문서:")
        control_label.grid(row=1, column=0, sticky="e", **padding)

        control_value = tk.Label(self.root, textvariable=self.control_var, anchor="w", width=50)
        control_value.grid(row=1, column=1, sticky="w", **padding)

        select_button = tk.Button(
            self.root,
            text="컨트롤 통합문서 선택",
            command=self._choose_control_workbook,
            width=25,
        )
        select_button.grid(row=2, column=0, columnspan=2, **padding)

        actions = [
            ("Select Folder Path", self._handle_select_folder),
            ("List Excel Files", self._handle_list_files),
            ("Update Files", self._handle_update_files),
            ("IO Change", self._handle_io_change),
            ("Value Find", self._handle_value_find),
            ("Change Value", self._handle_change_value),
        ]

        for idx, (label, handler) in enumerate(actions, start=3):
            button = tk.Button(self.root, text=label, width=25, command=handler, state=tk.DISABLED)
            button.grid(row=idx, column=0, columnspan=2, **padding)
            self.action_buttons.append(button)

        status_label = tk.Label(self.root, textvariable=self.status_var, anchor="w")
        status_label.grid(row=idx + 1, column=0, columnspan=2, sticky="we", **padding)

    # ----------------------------------------------------------- UI helpers
    def _ensure_control_selected(self) -> Path:
        if not self.control_path:
            raise RuntimeError("컨트롤 통합문서를 먼저 선택하세요.")
        return self.control_path

    def _choose_control_workbook(self) -> None:
        filetypes = [
            ("Excel", [f"*{ext}" for ext in EXCEL_EXTENSIONS]),
            ("All files", "*.*"),
        ]
        selected = filedialog.askopenfilename(
            title="컨트롤 통합문서 선택",
            filetypes=filetypes,
        )
        if not selected:
            return
        self.control_path = Path(selected)
        self.control_var.set(str(self.control_path))
        self.status_var.set("버튼을 눌러 원하는 작업을 실행하세요.")
        for button in self.action_buttons:
            button.config(state=tk.NORMAL)

    def _run_action(self, label: str, func: Callable[[], str]) -> None:
        try:
            self._toggle_buttons(tk.DISABLED)
            message = func()
            self.status_var.set(message)
            messagebox.showinfo("완료", message, parent=self.root)
        except Exception as exc:  # noqa: BLE001
            error_message = f"{label} 실행 중 오류가 발생했습니다.\n{exc}"
            self.status_var.set(error_message)
            messagebox.showerror("오류", error_message, parent=self.root)
        finally:
            self._toggle_buttons(tk.NORMAL)

    def _toggle_buttons(self, state: str) -> None:
        for button in self.action_buttons:
            button.config(state=state if self.control_path else tk.DISABLED)

    # --------------------------------------------------------------- Actions
    def _handle_select_folder(self) -> None:
        def action() -> str:
            control_path = self._ensure_control_selected()
            folder = filedialog.askdirectory(title="대상 폴더 선택")
            if not folder:
                raise RuntimeError("폴더 선택이 취소되었습니다.")
            count = select_folder_path(control_path, Path(folder))
            return f"{count}개의 Excel 파일을 불러왔습니다."

        self._run_action("Select Folder Path", action)

    def _handle_list_files(self) -> None:
        def action() -> str:
            control_path = self._ensure_control_selected()
            count = list_excel_files(control_path)
            return f"{count}개의 Excel 파일 목록을 갱신했습니다."

        self._run_action("List Excel Files", action)

    def _handle_update_files(self) -> None:
        def action() -> str:
            control_path = self._ensure_control_selected()
            count = update_files(control_path)
            return f"{count}개의 파일에서 ID를 갱신했습니다."

        self._run_action("Update Files", action)

    def _handle_io_change(self) -> None:
        def action() -> str:
            control_path = self._ensure_control_selected()
            count = io_change(control_path)
            return f"{count}개의 파일에서 IO 문자열을 치환했습니다."

        self._run_action("IO Change", action)

    def _handle_value_find(self) -> None:
        def action() -> str:
            control_path = self._ensure_control_selected()
            matches = value_find(control_path)
            return f"{matches}건의 결과를 data_update 시트에 기록했습니다."

        self._run_action("Value Find", action)

    def _handle_change_value(self) -> None:
        def action() -> str:
            control_path = self._ensure_control_selected()
            success, failure, errors = change_value(control_path)
            if errors:
                error_text = "\n".join(errors[:5])
                if len(errors) > 5:
                    error_text += f"\n... 외 {len(errors) - 5}건"
                raise RuntimeError(f"{failure}건 실패\n{error_text}")
            return f"{success}개의 셀 값을 변경했습니다."

        self._run_action("Change Value", action)

    # --------------------------------------------------------------- Entrypoint
    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = VBAConversionApp()
    app.run()


if __name__ == "__main__":
    main()

