from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .detection import auto_detect_config, mapping_summary
from .models import NormalizeRequest, SheetConfig
from .normalize import OUTPUT_HEADERS, normalize_workbook, normalized_preview_rows, safe_text
from .templates import default_template_path, load_template, save_template, template_directory

NONE_LABEL = "— None —"


class SheetMappingDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, input_path: Path, sheet_name: str, initial: SheetConfig | None = None):
        super().__init__(parent)
        self.title(f"Map columns — {sheet_name}")
        self.geometry("1220x790")
        self.minsize(980, 650)
        self.transient(parent)
        self.grab_set()
        self.result: SheetConfig | None = None
        self.input_path = input_path
        self.sheet_name = sheet_name
        self.initial = initial or SheetConfig(sheet_name=sheet_name)
        self.current_max_col = 0

        self.header_row_var = tk.IntVar(value=self.initial.header_row)
        self.start_row_var = tk.IntVar(value=self.initial.start_row)
        self.column_vars: dict[str, tk.StringVar] = {}
        self.description_list: tk.Listbox
        self.preview_tree: ttk.Treeview
        self.normalized_tree: ttk.Treeview
        self._build()
        self._load_preview(self.initial)

    def _build(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        top = ttk.LabelFrame(root, text="1. Row settings", padding=10)
        top.pack(fill="x")
        ttk.Label(top, text="Header row").grid(row=0, column=0, sticky="w")
        ttk.Spinbox(top, from_=1, to=5000, textvariable=self.header_row_var, width=8).grid(row=0, column=1, padx=(6, 20))
        ttk.Label(top, text="First data row").grid(row=0, column=2, sticky="w")
        ttk.Spinbox(top, from_=1, to=5000, textvariable=self.start_row_var, width=8).grid(row=0, column=3, padx=(6, 20))
        ttk.Button(top, text="Refresh source", command=self._load_preview).grid(row=0, column=4, padx=(0, 8))
        ttk.Button(top, text="Auto Detect", command=self._auto_detect, style="Accent.TButton").grid(row=0, column=5)

        mapping = ttk.LabelFrame(root, text="2. Column mapping", padding=10)
        mapping.pack(fill="x", pady=10)
        fields = [
            ("item_no", "Item No. / Existing WBS"),
            ("unit", "Unit"),
            ("qty", "Qty"),
            ("material", "Material"),
            ("labor", "Labor"),
            ("amount", "Amount"),
        ]
        for idx, (key, label) in enumerate(fields):
            ttk.Label(mapping, text=label).grid(row=0, column=idx, sticky="w", padx=4)
            var = tk.StringVar(value=NONE_LABEL)
            self.column_vars[key] = var
            combo = ttk.Combobox(mapping, textvariable=var, state="readonly", width=18)
            combo.grid(row=1, column=idx, padx=4, pady=(3, 0), sticky="ew")
            setattr(self, f"{key}_combo", combo)
            mapping.columnconfigure(idx, weight=1)

        body = ttk.Panedwindow(root, orient="horizontal")
        body.pack(fill="both", expand=True)
        desc_frame = ttk.LabelFrame(body, text="3. Description columns", padding=8)
        preview_frame = ttk.LabelFrame(body, text="4. Preview", padding=8)
        body.add(desc_frame, weight=1)
        body.add(preview_frame, weight=5)

        self.description_list = tk.Listbox(desc_frame, selectmode="extended", exportselection=False, width=30)
        self.description_list.pack(fill="both", expand=True)
        ttk.Label(desc_frame, text="Select one or more columns. Values are merged with spaces.", wraplength=220).pack(fill="x", pady=(6, 0))

        notebook = ttk.Notebook(preview_frame)
        notebook.pack(fill="both", expand=True)
        source_tab = ttk.Frame(notebook)
        normalized_tab = ttk.Frame(notebook)
        notebook.add(source_tab, text="Source")
        notebook.add(normalized_tab, text="Normalized")

        self.preview_tree = self._create_tree(source_tab)
        self.normalized_tree = self._create_tree(normalized_tab)

        preview_actions = ttk.Frame(preview_frame)
        preview_actions.pack(fill="x", pady=(8, 0))
        ttk.Button(preview_actions, text="Refresh normalized preview", command=self._load_normalized_preview).pack(side="right")

        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=(10, 0))
        ttk.Button(buttons, text="Cancel", command=self.destroy).pack(side="right")
        ttk.Button(buttons, text="Save mapping", command=self._save, style="Accent.TButton").pack(side="right", padx=(0, 8))
        ttk.Separator(buttons, orient="vertical").pack(side="right", fill="y", padx=8)
        ttk.Button(buttons, text="Save template…", command=self._save_template).pack(side="right")
        ttk.Button(buttons, text="Load template…", command=self._load_template).pack(side="right", padx=(0, 6))

    @staticmethod
    def _create_tree(parent: ttk.Frame) -> ttk.Treeview:
        tree = ttk.Treeview(parent, show="headings", height=18)
        xbar = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        ybar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=xbar.set, yscrollcommand=ybar.set)
        tree.grid(row=0, column=0, sticky="nsew")
        ybar.grid(row=0, column=1, sticky="ns")
        xbar.grid(row=1, column=0, sticky="ew")
        parent.rowconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        return tree

    def _read_sheet(self):
        wb = load_workbook(self.input_path, read_only=False, data_only=True)
        return wb, wb[self.sheet_name]

    def _load_preview(self, config: SheetConfig | None = None) -> None:
        try:
            wb, ws = self._read_sheet()
            max_col = min(ws.max_column, 60)
            self.current_max_col = max_col
            header_row = max(1, self.header_row_var.get())
            choices = [NONE_LABEL]
            desc_choices: list[str] = []
            for col in range(1, max_col + 1):
                header = safe_text(ws.cell(header_row, col).value)
                label = f"{get_column_letter(col)} — {header or '(blank)'}"
                choices.append(label)
                desc_choices.append(label)
            for key in self.column_vars:
                getattr(self, f"{key}_combo")["values"] = choices

            previous = config or self._config_from_widgets(require_description=False)
            self.description_list.delete(0, "end")
            for label in desc_choices:
                self.description_list.insert("end", label)
            self._apply_config_to_widgets(previous, choices)

            columns = ["ROW"] + [get_column_letter(c) for c in range(1, max_col + 1)]
            self._configure_tree(self.preview_tree, columns, default_width=82)
            preview_start = max(1, header_row - 2)
            preview_end = min(ws.max_row, preview_start + 44)
            for row in range(preview_start, preview_end + 1):
                values = [row] + [safe_text(ws.cell(row, c).value)[:100] for c in range(1, max_col + 1)]
                self.preview_tree.insert("", "end", values=values)
            wb.close()
            self._load_normalized_preview(show_errors=False)
        except Exception as exc:
            messagebox.showerror("Preview error", str(exc), parent=self)

    @staticmethod
    def _configure_tree(tree: ttk.Treeview, columns: list[str], default_width: int = 110) -> None:
        tree.delete(*tree.get_children())
        tree["columns"] = columns
        for col in columns:
            tree.heading(col, text=col)
            width = 58 if col in {"ROW", "WBS Level"} else default_width
            if col in {"Description", "WBS-1", "WBS-2", "WBS-3", "WBS-4"}:
                width = 180
            tree.column(col, width=width, stretch=False)

    def _auto_detect(self) -> None:
        try:
            wb, ws = self._read_sheet()
            cfg = auto_detect_config(ws, self.sheet_name, self._config_from_widgets(require_description=False))
            wb.close()
            self.header_row_var.set(cfg.header_row)
            self.start_row_var.set(cfg.start_row)
            self._load_preview(cfg)
            messagebox.showinfo("Auto Detect", "Suggested mapping applied. Review the source and normalized previews before saving.", parent=self)
        except Exception as exc:
            messagebox.showerror("Auto Detect failed", str(exc), parent=self)

    def _load_normalized_preview(self, show_errors: bool = True) -> None:
        try:
            cfg = self._config_from_widgets(require_description=True)
            wb, ws = self._read_sheet()
            rows = normalized_preview_rows(ws, cfg, limit=80)
            wb.close()
            self._configure_tree(self.normalized_tree, OUTPUT_HEADERS, default_width=105)
            for row in rows:
                values = ["" if value is None else value for value in row]
                self.normalized_tree.insert("", "end", values=values)
            if not rows and show_errors:
                messagebox.showwarning("No normalized rows", "No output rows were detected. Check First data row and column mapping.", parent=self)
        except ValueError as exc:
            if show_errors:
                messagebox.showwarning("Mapping incomplete", str(exc), parent=self)
        except Exception as exc:
            if show_errors:
                messagebox.showerror("Normalized preview error", str(exc), parent=self)

    @staticmethod
    def _column_from_label(label: str) -> int | None:
        if not label or label.startswith("—"):
            return None
        letters = label.split(" — ", 1)[0]
        value = 0
        for char in letters:
            value = value * 26 + ord(char.upper()) - 64
        return value

    def _config_from_widgets(self, require_description: bool = True) -> SheetConfig:
        description_cols = [index + 1 for index in self.description_list.curselection()]
        if require_description and not description_cols:
            raise ValueError("Select at least one Description column.")
        return SheetConfig(
            sheet_name=self.sheet_name,
            header_row=max(1, self.header_row_var.get()),
            start_row=max(1, self.start_row_var.get()),
            item_no_col=self._column_from_label(self.column_vars["item_no"].get()),
            description_cols=description_cols,
            unit_col=self._column_from_label(self.column_vars["unit"].get()),
            qty_col=self._column_from_label(self.column_vars["qty"].get()),
            material_col=self._column_from_label(self.column_vars["material"].get()),
            labor_col=self._column_from_label(self.column_vars["labor"].get()),
            amount_col=self._column_from_label(self.column_vars["amount"].get()),
        )

    def _apply_config_to_widgets(self, cfg: SheetConfig, choices: list[str] | None = None) -> None:
        self.header_row_var.set(cfg.header_row)
        self.start_row_var.set(cfg.start_row)
        choices = choices or list(getattr(self, "item_no_combo")["values"])
        column_map = {
            "item_no": cfg.item_no_col,
            "unit": cfg.unit_col,
            "qty": cfg.qty_col,
            "material": cfg.material_col,
            "labor": cfg.labor_col,
            "amount": cfg.amount_col,
        }
        for key, col in column_map.items():
            self.column_vars[key].set(choices[col] if col and col < len(choices) else NONE_LABEL)
        self.description_list.selection_clear(0, "end")
        for col in cfg.description_cols:
            if 1 <= col <= self.description_list.size():
                self.description_list.selection_set(col - 1)

    def _save_template(self) -> None:
        try:
            cfg = self._config_from_widgets(require_description=True)
        except ValueError as exc:
            messagebox.showwarning("Mapping incomplete", str(exc), parent=self)
            return
        suggested = default_template_path(self.input_path, self.sheet_name)
        filename = filedialog.asksaveasfilename(
            parent=self,
            title="Save mapping template",
            initialdir=suggested.parent,
            initialfile=suggested.name,
            defaultextension=".json",
            filetypes=[("BOQ mapping", "*.json")],
        )
        if filename:
            save_template(Path(filename), cfg)
            messagebox.showinfo("Template saved", filename, parent=self)

    def _load_template(self) -> None:
        filename = filedialog.askopenfilename(
            parent=self,
            title="Load mapping template",
            initialdir=template_directory(),
            filetypes=[("BOQ mapping", "*.json"), ("All files", "*.*")],
        )
        if not filename:
            return
        try:
            cfg = load_template(Path(filename), sheet_name=self.sheet_name)
            self.header_row_var.set(cfg.header_row)
            self.start_row_var.set(cfg.start_row)
            self._load_preview(cfg)
        except Exception as exc:
            messagebox.showerror("Cannot load template", str(exc), parent=self)

    def _save(self) -> None:
        try:
            self.result = self._config_from_widgets(require_description=True)
        except ValueError as exc:
            messagebox.showwarning("Missing mapping", str(exc), parent=self)
            return
        # Auto-save a private copy so the same workbook/sheet opens quickly next time.
        try:
            save_template(default_template_path(self.input_path, self.sheet_name), self.result)
        except OSError:
            pass
        self.destroy()


class BOQStudioApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("BOQ Studio — MS-1.1")
        self.geometry("980x720")
        self.minsize(840, 620)
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
        self.file_label = ttk.Label(file_frame, text="No file selected", wraplength=700)
        self.file_label.pack(side="left", fill="x", expand=True)
        ttk.Button(file_frame, text="Choose Excel file…", command=self.choose_file).pack(side="right")

        sheets_frame = ttk.LabelFrame(root, text="2. Select breakdown sheets and map columns", padding=10)
        sheets_frame.pack(fill="both", expand=True, pady=12)
        toolbar = ttk.Frame(sheets_frame)
        toolbar.pack(fill="x", pady=(0, 6))
        ttk.Button(toolbar, text="Select all", command=lambda: self._set_all(True)).pack(side="left")
        ttk.Button(toolbar, text="Clear", command=lambda: self._set_all(False)).pack(side="left", padx=6)
        ttk.Label(toolbar, text="Double-click a sheet to Auto Detect, review and save mapping.").pack(side="right")

        self.sheet_tree = ttk.Treeview(sheets_frame, columns=("use", "sheet", "mapping"), show="headings", selectmode="browse")
        self.sheet_tree.heading("use", text="Use")
        self.sheet_tree.heading("sheet", text="Sheet")
        self.sheet_tree.heading("mapping", text="Mapping")
        self.sheet_tree.column("use", width=70, anchor="center")
        self.sheet_tree.column("sheet", width=280)
        self.sheet_tree.column("mapping", width=580)
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
            auto_path = default_template_path(self.input_path, name)
            mapping = "Not configured"
            if auto_path.exists():
                try:
                    self.configs[name] = load_template(auto_path, sheet_name=name)
                    mapping = mapping_summary(self.configs[name])
                except Exception:
                    pass
            self.sheet_tree.insert("", "end", iid=name, values=("✓" if selected else "", name, mapping))
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
        mapping = mapping_summary(cfg) if cfg else "Not configured"
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
