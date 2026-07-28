from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .models import NormalizeRequest, SheetConfig
from .normalize import normalize_workbook, safe_text


class SheetMappingDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, input_path: Path, sheet_name: str, initial: SheetConfig | None = None):
        super().__init__(parent)
        self.title(f"Map columns — {sheet_name}")
        self.geometry("1100x720")
        self.minsize(900, 600)
        self.transient(parent)
        self.grab_set()
        self.result: SheetConfig | None = None
        self.input_path = input_path
        self.sheet_name = sheet_name
        self.initial = initial or SheetConfig(sheet_name=sheet_name)

        self.header_row_var = tk.IntVar(value=self.initial.header_row)
        self.start_row_var = tk.IntVar(value=self.initial.start_row)
        self.column_vars: dict[str, tk.StringVar] = {}
        self.description_list: tk.Listbox
        self.preview_tree: ttk.Treeview
        self._build()
        self._load_preview()

    def _build(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)
        top = ttk.LabelFrame(root, text="1. Row settings", padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Header row").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(top, from_=1, to=500, textvariable=self.header_row_var, width=8, command=self._load_preview).grid(row=0, column=1, padx=(6, 20))
        ttk.Label(top, text="First data row").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(top, from_=1, to=500, textvariable=self.start_row_var, width=8).grid(row=0, column=3, padx=(6, 20))
        ttk.Button(top, text="Refresh preview", command=self._load_preview).grid(row=0, column=4)

        mapping = ttk.LabelFrame(root, text="2. Column mapping", padding=10)
        mapping.pack(fill="x", pady=10)
        fields = [
            ("item_no", "Item / sequence"),
            ("unit", "Unit"),
            ("qty", "Qty"),
            ("material", "Material"),
            ("labor", "Labor"),
            ("amount", "Amount"),
        ]
        for idx, (key, label) in enumerate(fields):
            ttk.Label(mapping, text=label).grid(row=0, column=idx, sticky="w", padx=4)
            var = tk.StringVar(value="— None —")
            self.column_vars[key] = var
            combo = ttk.Combobox(mapping, textvariable=var, state="readonly", width=16)
            combo.grid(row=1, column=idx, padx=4, pady=(3, 0))
            setattr(self, f"{key}_combo", combo)

        body = ttk.Panedwindow(root, orient="horizontal")
        body.pack(fill="both", expand=True)
        desc_frame = ttk.LabelFrame(body, text="3. Description columns (select one or more)", padding=8)
        preview_frame = ttk.LabelFrame(body, text="Preview", padding=8)
        body.add(desc_frame, weight=1)
        body.add(preview_frame, weight=4)

        self.description_list = tk.Listbox(desc_frame, selectmode="extended", exportselection=False, width=30)
        self.description_list.pack(fill="both", expand=True)
        ttk.Label(desc_frame, text="Use Ctrl/Shift to select multiple columns.", wraplength=220).pack(fill="x", pady=(6, 0))

        self.preview_tree = ttk.Treeview(preview_frame, show="headings", height=18)
        xbar = ttk.Scrollbar(preview_frame, orient="horizontal", command=self.preview_tree.xview)
        ybar = ttk.Scrollbar(preview_frame, orient="vertical", command=self.preview_tree.yview)
        self.preview_tree.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        self.preview_tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        preview_frame.rowconfigure(0, weight=1)
        preview_frame.columnconfigure(0, weight=1)

        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Save mapping", command=self._save, style="Accent.TButton").pack(side="right", padx=(0, 8))

    def _load_preview(self) -> None:
        try:
            wb = load_workbook(self.input_path, read_only=True, data_only=True)
            ws = wb[self.sheet_name]
            max_col = min(ws.max_column, 40)
            header_row = max(1, self.header_row_var.get())
            choices = ["— None —"]
            desc_choices = []
            for col in range(1, max_col + 1):
                header = safe_text(ws.cell(header_row, col).value)
                label = f"{get_column_letter(col)} — {header or '(blank)'}"
                choices.append(label)
                desc_choices.append(label)
            for key in self.column_vars:
                combo = getattr(self, f"{key}_combo")
                combo["values"] = choices
            self.description_list.delete(0, "end")
            for label in desc_choices:
                self.description_list.insert("end", label)

            initial_map = {
                "item_no": self.initial.item_no_col,
                "unit": self.initial.unit_col,
                "qty": self.initial.qty_col,
                "material": self.initial.material_col,
                "labor": self.initial.labor_col,
                "amount": self.initial.amount_col,
            }
            for key, col in initial_map.items():
                if col and col <= max_col:
                    self.column_vars[key].set(choices[col])
            for col in self.initial.description_cols:
                if 1 <= col <= max_col:
                    self.description_list.selection_set(col - 1)

            columns = ["ROW"] + [get_column_letter(c) for c in range(1, max_col + 1)]
            self.preview_tree.delete(*self.preview_tree.get_children())
            self.preview_tree["columns"] = columns
            for col in columns:
                self.preview_tree.heading(col, text=col)
                self.preview_tree.column(col, width=70 if col != "ROW" else 50, stretch=False)
            preview_start = max(1, header_row - 2)
            preview_end = min(ws.max_row, preview_start + 34)
            for row in range(preview_start, preview_end + 1):
                values = [row] + [safe_text(ws.cell(row, c).value)[:80] for c in range(1, max_col + 1)]
                self.preview_tree.insert("", "end", values=values)
            wb.close()
        except Exception as exc:
            messagebox.showerror("Preview error", str(exc), parent=self)

    @staticmethod
    def _column_from_label(label: str) -> int | None:
        if not label or label.startswith("—"):
            return None
        letters = label.split(" — ", 1)[0]
        value = 0
        for char in letters:
            value = value * 26 + ord(char.upper()) - 64
        return value

    def _save(self) -> None:
        description_cols = [index + 1 for index in self.description_list.curselection()]
        if not description_cols:
            messagebox.showwarning("Missing mapping", "Select at least one Description column.", parent=self)
            return
        self.result = SheetConfig(
            sheet_name=self.sheet_name,
            header_row=self.header_row_var.get(),
            start_row=self.start_row_var.get(),
            item_no_col=self._column_from_label(self.column_vars["item_no"].get()),
            description_cols=description_cols,
            unit_col=self._column_from_label(self.column_vars["unit"].get()),
            qty_col=self._column_from_label(self.column_vars["qty"].get()),
            material_col=self._column_from_label(self.column_vars["material"].get()),
            labor_col=self._column_from_label(self.column_vars["labor"].get()),
            amount_col=self._column_from_label(self.column_vars["amount"].get()),
        )
        self.destroy()


class BOQStudioApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BOQ Studio — MS-1")
        self.geometry("950x700")
        self.minsize(820, 600)
        self.input_path: Path | None = None
        self.sheet_vars: dict[str, tk.BooleanVar] = {}
        self.configs: dict[str, SheetConfig] = {}
        self.events: queue.Queue = queue.Queue()
        self._build()
        self.after(100, self._poll_events)

    def _build(self) -> None:
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Title.TLabel", font=("Segoe UI", 20, "bold"))
        style.configure("Subtitle.TLabel", font=("Segoe UI", 10))
        style.configure("Accent.TButton", font=("Segoe UI", 10, "bold"))

        root = ttk.Frame(self, padding=18)
        root.pack(fill="both", expand=True)
        ttk.Label(root, text="BOQ Studio", style="Title.TLabel").pack(anchor="w")
        ttk.Label(root, text="Normalize breakdown sheets into OKD-compatible Excel sheets.", style="Subtitle.TLabel").pack(anchor="w", pady=(0, 14))

        file_frame = ttk.LabelFrame(root, text="1. Input BOQ", padding=10)
        file_frame.pack(fill="x")
        self.file_label = ttk.Label(file_frame, text="No file selected", wraplength=680)
        self.file_label.pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="Choose Excel file…", command=self.choose_file).pack(side="right")

        sheets_frame = ttk.LabelFrame(root, text="2. Select breakdown sheets and map columns", padding=10)
        sheets_frame.pack(fill="both", expand=True, pady=12)
        toolbar = ttk.Frame(sheets_frame)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Select all", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(toolbar, text="Clear", command=lambda: self._set_all(False)).pack(side="left", padx=6)
        ttk.Label(toolbar, text="Double-click a sheet to configure its columns.").pack(side="right")

        self.sheet_tree = ttk.Treeview(sheets_frame, columns=("use", "sheet", "mapping"), show="headings", selectmode="browse")
        self.sheet_tree.heading("use", text="Use")
        self.sheet_tree.heading("sheet", text="Sheet")
        self.sheet_tree.heading("mapping", text="Mapping")
        self.sheet_tree.column("use", width=70, anchor="center")
        self.sheet_tree.column("sheet", width=360)
        self.sheet_tree.column("mapping", width=420)
        self.sheet_tree.pack(fill="both", expand=True)
        self.sheet_tree.bind("<Double-1>", self.configure_selected)
        self.sheet_tree.bind("<Button-1>", self.toggle_sheet)

        action = ttk.Frame(root)
        action.pack(fill="x")
        self.run_button = ttk.Button(action, text="Normalize to Desktop", command=self.run_normalize, style="Accent.TButton", state="disabled")
        self.run_button.pack(side="right")

        progress_frame = ttk.LabelFrame(root, text="3. Progress", padding=10)
        progress_frame.pack(fill="x", pady=(12, 0))
        self.status_label = ttk.Label(progress_frame, text="Ready")
        self.status_label.pack(anchor="w")
        self.progress = ttk.Progressbar(progress_frame, mode="determinate", maximum=100)
        self.progress.pack(fill="x", pady=(6, 0))

    def choose_file(self) -> None:
        filename = filedialog.askopenfilename(title="Choose BOQ Excel file", filetypes=[("Excel workbook", "*.xlsx"), ("All files", "*.*")])
        if not filename:
            return
        self.input_path = Path(filename)
        try:
            wb = load_workbook(self.input_path, read_only=True, data_only=False)
            names = wb.sheetnames
            wb.close()
        except Exception as exc:
            messagebox.showerror("Cannot open workbook", str(exc), parent=self)
            return
        self.file_label.configure(text=str(self.input_path))
        self.sheet_vars.clear()
        self.configs.clear()
        self.sheet_tree.delete(*self.sheet_tree.get_children())
        for name in names:
            selected = not any(word in name.casefold() for word in ("sum", "summary", "ปร.6", "ปร.5", "log"))
            self.sheet_vars[name] = tk.BooleanVar(value=selected)
            self.sheet_tree.insert("", "end", iid=name, values=("✓" if selected else "", name, "Not configured"))
        self.run_button.configure(state="normal")
        self.status_label.configure(text=f"Loaded {len(names)} sheets")

    def _set_all(self, value: bool) -> None:
        for name, var in self.sheet_vars.items():
            var.set(value)
            self._refresh_row(name)

    def toggle_sheet(self, event) -> None:
        region = self.sheet_tree.identify("region", event.x, event.y)
        column = self.sheet_tree.identify_column(event.x)
        row = self.sheet_tree.identify_row(event.y)
        if region == "cell" and column == "#1" and row:
            self.sheet_vars[row].set(not self.sheet_vars[row].get())
            self._refresh_row(row)

    def configure_selected(self, _event=None) -> None:
        selection = self.sheet_tree.selection()
        if not selection or not self.input_path:
            return
        name = selection[0]
        dialog = SheetMappingDialog(self, self.input_path, name, self.configs.get(name))
        self.wait_window(dialog)
        if dialog.result:
            self.configs[name] = dialog.result
            self.sheet_vars[name].set(True)
            self._refresh_row(name)

    def _refresh_row(self, name: str) -> None:
        cfg = self.configs.get(name)
        mapping = "Configured" if cfg else "Not configured"
        self.sheet_tree.item(name, values=("✓" if self.sheet_vars[name].get() else "", name, mapping))

    def run_normalize(self) -> None:
        if not self.input_path:
            return
        selected = [name for name, var in self.sheet_vars.items() if var.get()]
        if not selected:
            messagebox.showwarning("No sheets", "Select at least one breakdown sheet.", parent=self)
            return
        missing = [name for name in selected if name not in self.configs]
        if missing:
            messagebox.showwarning("Mapping required", "Double-click and configure every selected sheet first:\n\n" + "\n".join(missing), parent=self)
            return

        desktop = Path.home() / "Desktop"
        stem = self.input_path.stem.strip()
        output = desktop / f"{stem}_OKD.xlsx"
        suffix = 2
        while output.exists():
            output = desktop / f"{stem}_OKD_{suffix}.xlsx"
            suffix += 1
        request = NormalizeRequest(self.input_path, output, [self.configs[name] for name in selected])
        self.run_button.configure(state="disabled")
        self.progress["value"] = 0
        self.status_label.configure(text="Starting…")

        def worker() -> None:
            try:
                def on_progress(done: int, total: int, message: str) -> None:
                    self.events.put(("progress", done, total, message))
                results = normalize_workbook(request, on_progress)
                self.events.put(("done", output, results))
            except Exception as exc:
                self.events.put(("error", str(exc)))

        threading.Thread(target=worker, daemon=True).start()

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, done, total, message = event
                    self.progress["value"] = (done / max(1, total)) * 100
                    self.status_label.configure(text=message)
                elif event[0] == "done":
                    _, output, results = event
                    self.progress["value"] = 100
                    ok = sum(1 for r in results if r.status == "OK")
                    warning = sum(1 for r in results if r.status == "WARNING")
                    error = sum(1 for r in results if r.status == "ERROR")
                    self.status_label.configure(text=f"Finished — OK {ok}, Warning {warning}, Error {error}")
                    self.run_button.configure(state="normal")
                    messagebox.showinfo("BOQ Studio finished", f"Output created:\n{output}\n\nOK: {ok}\nWarnings: {warning}\nErrors: {error}", parent=self)
                elif event[0] == "error":
                    self.run_button.configure(state="normal")
                    self.status_label.configure(text="Failed")
                    messagebox.showerror("Normalization failed", event[1], parent=self)
        except queue.Empty:
            pass
        self.after(100, self._poll_events)


def main() -> None:
    app = BOQStudioApp()
    app.mainloop()


if __name__ == "__main__":
    main()
