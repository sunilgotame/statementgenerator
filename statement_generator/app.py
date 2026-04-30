from __future__ import annotations

from datetime import date, timedelta
import json
from pathlib import Path
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .exchange_rate import ExchangeRateLookupError, NRB_FOREX_DOCS, NRB_FOREX_PAGE
from .exporters import (
    TemplateCatalog,
    build_payload,
    default_output_name,
    export_certificate,
    export_normal_certificate,
    export_normal_statement,
    export_statement,
    resolve_exchange_rate,
    scan_template_directory,
)
from .generator import StatementConfig, generate_statement, names_from_text
from .selftest import run_tests
from .utils import format_amount, parse_iso_date, safe_filename


DEFAULT_TEMPLATE_DIR = Path(r"D:\Finance Doc\Format")
LEGACY_WEB_GENERATOR_PATH = Path(r"D:\Finance Doc\2026\Rubi\nepali_bank_statement_generator.html")
PERSISTENT_RULES_SCHEMA_VERSION = 2
REQUIRED_HOLIDAY_DATES = {"2025-10-23"}
DESCRIPTION_MODE_OPTIONS = {
    "Label + Name": "label_plus_name",
    "Label Only": "label_only",
    "Name Only": "name_only",
}


class StatementGeneratorApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Nepali Statement Generator Desktop")
        self.geometry("1540x940")
        self.minsize(1360, 800)

        self.generated_result = None
        self.catalog = TemplateCatalog([], [])
        self.legacy_holiday_dates = self._load_legacy_holiday_dates()
        self.custom_holiday_dates: set[str] = set(self.legacy_holiday_dates)
        self.excluded_saturday_dates: set[str] = set()
        self.holiday_manager_window: tk.Toplevel | None = None
        self.holiday_tree: ttk.Treeview | None = None
        self.holiday_view_var = tk.StringVar(value="All")
        self.holiday_status_var = tk.StringVar(value="")
        self.holiday_edit_date_var = tk.StringVar()
        self.holiday_edit_type_var = tk.StringVar(value="Holiday")
        self.summary_vars: dict[str, tk.StringVar] = {}
        self.vars: dict[str, tk.StringVar] = {}
        self._build_variables()
        self._configure_style()
        self._build_ui()
        self.protocol("WM_DELETE_WINDOW", self._on_app_close)
        self._load_persistent_rules()
        for key in ("start_date", "end_date"):
            self.vars[key].trace_add("write", lambda *_args: self.refresh_holiday_display())
        self.refresh_templates()
        self.refresh_holiday_display()

    def _build_variables(self) -> None:
        today = date.today().isoformat()
        defaults = {
            "bank_name": "Nepal Bank Limited",
            "branch_name": "Main Branch, Kathmandu",
            "customer_name": "Customer Name",
            "customer_address": "Customer Address",
            "account_number": "0000000000",
            "account_type": "Saving Account",
            "member_id": "",
            "currency": "NPR",
            "reference_no": "",
            "opening_date": "2022-05-10",
            "start_date": "2025-01-14",
            "end_date": "2026-01-14",
            "opening_balance": "1500000",
            "target_closing_balance": "2550000",
            "interest_rate": "8",
            "tax_rate": "6",
            "cheque_start": "10000001",
            "first_date_description": "Opening Balance",
            "last_date_description": "Balance C/F",
            "seed": "",
            "deposit_text": "Cash Deposit",
            "withdrawal_text": "Cheque Withdrawal",
            "interest_text": "Interest Posted",
            "tax_text": "Tax Deducted",
            "deposit_mode": "Label + Name",
            "withdrawal_mode": "Label + Name",
            "template_dir": str(DEFAULT_TEMPLATE_DIR),
            "manual_rate": "",
            "rate_mode": "Auto (NRB)",
            "rate_type": "sell",
            "today": today,
        }
        for key, value in defaults.items():
            self.vars[key] = tk.StringVar(value=value)

        for key in (
            "final_balance",
            "issue_date",
            "row_count",
            "target_gap",
            "total_deposits",
            "total_withdrawals",
            "total_interest",
            "total_tax",
            "seed_used",
        ):
            self.summary_vars[key] = tk.StringVar(value="-")

    def _configure_style(self) -> None:
        style = ttk.Style(self)
        if "vista" in style.theme_names():
            style.theme_use("vista")
        style.configure("Title.TLabel", font=("Segoe UI Semibold", 18))
        style.configure("Section.TLabelframe.Label", font=("Segoe UI Semibold", 10))
        style.configure("SummaryValue.TLabel", font=("Segoe UI Semibold", 14))

    def _build_ui(self) -> None:
        outer = ttk.Frame(self, padding=12)
        outer.pack(fill="both", expand=True)

        header = ttk.Frame(outer)
        header.pack(fill="x")
        ttk.Label(header, text="Nepali Statement Generator Desktop", style="Title.TLabel").pack(side="left")
        ttk.Label(
            header,
            text="Windows 11 desktop app with Excel and Word template export",
            foreground="#4b5563",
        ).pack(side="left", padx=(12, 0))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(10, 12))
        ttk.Button(actions, text="Generate Statement", command=self.generate_statement).pack(side="left")
        ttk.Button(actions, text="Holidays & Saturdays", command=self.open_holiday_manager).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Save Profile", command=self.save_profile).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Load Profile", command=self.load_profile).pack(side="left", padx=(8, 0))
        ttk.Button(actions, text="Run Self Tests", command=self.run_self_tests).pack(side="left", padx=(8, 0))

        content = ttk.Panedwindow(outer, orient="horizontal")
        content.pack(fill="both", expand=True)

        left = ttk.Frame(content, padding=(0, 0, 8, 0))
        right = ttk.Frame(content)
        content.add(left, weight=3)
        content.add(right, weight=4)

        notebook = ttk.Notebook(left)
        notebook.pack(fill="both", expand=True)

        tab_account = ttk.Frame(notebook, padding=10)
        tab_statement = ttk.Frame(notebook, padding=10)
        tab_texts = ttk.Frame(notebook, padding=10)
        tab_export = ttk.Frame(notebook, padding=10)
        notebook.add(tab_account, text="Account")
        notebook.add(tab_statement, text="Statement Rules")
        notebook.add(tab_texts, text="Texts & Names")
        notebook.add(tab_export, text="Export")

        self._build_account_tab(tab_account)
        self._build_statement_tab(tab_statement)
        self._build_texts_tab(tab_texts)
        self._build_export_tab(tab_export)
        self._build_preview(right)

    def _make_entry(self, parent: ttk.Frame, row: int, label: str, key: str, width: int = 28) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.vars[key], width=width).grid(row=row, column=1, sticky="ew", pady=4)

    def _build_account_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        labels = [
            ("Bank Name", "bank_name"),
            ("Branch", "branch_name"),
            ("Customer Name", "customer_name"),
            ("Customer Address", "customer_address"),
            ("Account Number", "account_number"),
            ("Account Type", "account_type"),
            ("Member ID", "member_id"),
            ("Currency", "currency"),
            ("Reference No.", "reference_no"),
            ("Opening Date (YYYY-MM-DD)", "opening_date"),
        ]
        for row, (label, key) in enumerate(labels):
            self._make_entry(parent, row, label, key)

    def _build_statement_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        labels = [
            ("Start Date (YYYY-MM-DD)", "start_date"),
            ("End Date (YYYY-MM-DD)", "end_date"),
            ("Opening Balance", "opening_balance"),
            ("Target Closing Balance", "target_closing_balance"),
            ("Interest Rate (%)", "interest_rate"),
            ("Tax Rate (%)", "tax_rate"),
            ("Cheque Start No.", "cheque_start"),
            ("First Date Description", "first_date_description"),
            ("Last Date Description", "last_date_description"),
            ("Optional Seed", "seed"),
        ]
        for row, (label, key) in enumerate(labels):
            self._make_entry(parent, row, label, key)

        note = (
            "Generation rules used here:\n"
            "- deposit amounts stay random instead of increasing from top to bottom\n"
            "- transaction dates are spread across months with mixed day numbers\n"
            "- interest and tax posting dates are reserved for system rows only\n"
            "- issue date becomes the next valid business day after the last transaction\n"
            "- blocked dates come from the Holiday & Saturday manager"
        )
        ttk.Label(parent, text=note, justify="left", foreground="#374151").grid(
            row=len(labels),
            column=0,
            columnspan=2,
            sticky="w",
            pady=(10, 0),
        )

    def _build_texts_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        self._make_entry(parent, 0, "Deposit Text", "deposit_text")
        self._make_entry(parent, 1, "Withdrawal Text", "withdrawal_text")
        self._make_entry(parent, 2, "Interest Text", "interest_text")
        self._make_entry(parent, 3, "Tax Text", "tax_text")

        ttk.Label(parent, text="Deposit Description Style").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Combobox(
            parent,
            textvariable=self.vars["deposit_mode"],
            values=list(DESCRIPTION_MODE_OPTIONS.keys()),
            state="readonly",
        ).grid(row=4, column=1, sticky="ew", pady=4)

        ttk.Label(parent, text="Withdrawal Description Style").grid(row=5, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Combobox(
            parent,
            textvariable=self.vars["withdrawal_mode"],
            values=list(DESCRIPTION_MODE_OPTIONS.keys()),
            state="readonly",
        ).grid(row=5, column=1, sticky="ew", pady=4)

        ttk.Label(parent, text="Deposit Names").grid(row=6, column=0, sticky="nw", padx=(0, 8), pady=4)
        self.deposit_names_text = tk.Text(parent, height=6, width=28, wrap="word")
        self.deposit_names_text.grid(row=6, column=1, sticky="ew", pady=4)
        self.deposit_names_text.insert("1.0", "Self\nKaruna\nKrishna\nManisha")

        ttk.Label(parent, text="Withdrawal Names").grid(row=7, column=0, sticky="nw", padx=(0, 8), pady=4)
        self.withdrawal_names_text = tk.Text(parent, height=6, width=28, wrap="word")
        self.withdrawal_names_text.grid(row=7, column=1, sticky="ew", pady=4)
        self.withdrawal_names_text.insert("1.0", "Self\nKabita Thapa\nKamala Pandey")

        ttk.Label(parent, text="Blocked Holidays & Saturdays").grid(row=8, column=0, sticky="nw", padx=(0, 8), pady=4)
        holiday_wrap = ttk.Frame(parent)
        holiday_wrap.grid(row=8, column=1, sticky="ew", pady=4)
        holiday_wrap.columnconfigure(0, weight=1)
        self.holiday_text = tk.Text(holiday_wrap, height=8, width=28, wrap="word")
        self.holiday_text.grid(row=0, column=0, sticky="ew")
        self.holiday_text.insert("1.0", "Holiday manager will show old holiday dates and active Saturdays here.")
        self.holiday_text.configure(state="disabled")
        holiday_buttons = ttk.Frame(holiday_wrap)
        holiday_buttons.grid(row=1, column=0, sticky="ew", pady=(6, 0))
        ttk.Button(holiday_buttons, text="Show Holidays & Saturdays", command=self.open_holiday_manager).pack(side="left")
        ttk.Button(holiday_buttons, text="Refresh List", command=self.refresh_holiday_display).pack(side="left", padx=(8, 0))

    def _build_export_tab(self, parent: ttk.Frame) -> None:
        parent.columnconfigure(1, weight=1)
        ttk.Label(parent, text="Template Folder").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        folder_frame = ttk.Frame(parent)
        folder_frame.grid(row=0, column=1, sticky="ew", pady=4)
        folder_frame.columnconfigure(0, weight=1)
        ttk.Entry(folder_frame, textvariable=self.vars["template_dir"]).grid(row=0, column=0, sticky="ew")
        ttk.Button(folder_frame, text="Browse", command=self.browse_template_dir).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(folder_frame, text="Refresh", command=self.refresh_templates).grid(row=0, column=2, padx=(6, 0))

        ttk.Label(parent, text="Exchange Rate Mode").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Combobox(parent, textvariable=self.vars["rate_mode"], values=["Auto (NRB)", "Manual"], state="readonly").grid(
            row=1, column=1, sticky="ew", pady=4
        )

        ttk.Label(parent, text="Rate Type").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Combobox(parent, textvariable=self.vars["rate_type"], values=["sell", "buy"], state="readonly").grid(
            row=2, column=1, sticky="ew", pady=4
        )

        ttk.Label(parent, text="Manual USD/NPR Rate").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=self.vars["manual_rate"]).grid(row=3, column=1, sticky="ew", pady=4)

        note = (
            "Auto rate source uses Nepal Rastra Bank official forex data.\n"
            f"Rates page: {NRB_FOREX_PAGE}\n"
            f"API docs: {NRB_FOREX_DOCS}"
        )
        ttk.Label(parent, text=note, justify="left", foreground="#374151").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(8, 8)
        )

        lists = ttk.Frame(parent)
        lists.grid(row=5, column=0, columnspan=2, sticky="nsew", pady=(4, 0))
        lists.columnconfigure(0, weight=1)
        lists.columnconfigure(1, weight=1)
        parent.rowconfigure(5, weight=1)

        statement_box = ttk.LabelFrame(lists, text="Export Statement to Excel", style="Section.TLabelframe")
        statement_box.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        statement_box.rowconfigure(0, weight=1)
        statement_box.columnconfigure(0, weight=1)
        self.statement_list = tk.Listbox(statement_box, exportselection=False, height=12)
        self.statement_list.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Button(statement_box, text="Export Selected Statement", command=self.export_selected_statement).grid(
            row=1, column=0, sticky="ew", padx=6, pady=(0, 6)
        )
        self.statement_list.bind("<Double-Button-1>", lambda _event: self.export_selected_statement())

        certificate_box = ttk.LabelFrame(lists, text="Export Balance Certificate to Word", style="Section.TLabelframe")
        certificate_box.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
        certificate_box.rowconfigure(0, weight=1)
        certificate_box.columnconfigure(0, weight=1)
        self.certificate_list = tk.Listbox(certificate_box, exportselection=False, height=12)
        self.certificate_list.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Button(certificate_box, text="Export Selected Certificate", command=self.export_selected_certificate).grid(
            row=1, column=0, sticky="ew", padx=6, pady=(0, 6)
        )
        self.certificate_list.bind("<Double-Button-1>", lambda _event: self.export_selected_certificate())

        fallback_box = ttk.LabelFrame(parent, text="Extra Normal Export", style="Section.TLabelframe")
        fallback_box.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(10, 0))
        ttk.Label(
            fallback_box,
            text="Use these if template export shows an Office/template error. These create simple new Excel and Word files.",
            foreground="#374151",
        ).pack(anchor="w", padx=8, pady=(8, 6))
        button_row = ttk.Frame(fallback_box)
        button_row.pack(fill="x", padx=8, pady=(0, 8))
        ttk.Button(button_row, text="Export Normal Excel", command=self.export_normal_statement_file).pack(side="left")
        ttk.Button(button_row, text="Export Normal Word", command=self.export_normal_certificate_file).pack(side="left", padx=(8, 0))

    def _build_preview(self, parent: ttk.Frame) -> None:
        summary = ttk.LabelFrame(parent, text="Summary", style="Section.TLabelframe", padding=10)
        summary.pack(fill="x")
        fields = [
            ("Final Balance", "final_balance"),
            ("Issue Date", "issue_date"),
            ("Rows", "row_count"),
            ("Target Gap", "target_gap"),
            ("Total Deposits", "total_deposits"),
            ("Total Withdrawals", "total_withdrawals"),
            ("Interest", "total_interest"),
            ("Tax", "total_tax"),
            ("Seed Used", "seed_used"),
        ]
        for index, (label, key) in enumerate(fields):
            row = index // 3
            column = (index % 3) * 2
            ttk.Label(summary, text=label).grid(row=row, column=column, sticky="w", padx=(0, 8), pady=4)
            ttk.Label(summary, textvariable=self.summary_vars[key], style="SummaryValue.TLabel").grid(
                row=row,
                column=column + 1,
                sticky="w",
                padx=(0, 18),
                pady=4,
            )

        table_wrap = ttk.LabelFrame(parent, text="Statement Preview", style="Section.TLabelframe", padding=10)
        table_wrap.pack(fill="both", expand=True, pady=(10, 0))
        columns = ("date", "description", "cheque", "debit", "credit", "balance")
        self.tree = ttk.Treeview(table_wrap, columns=columns, show="headings", height=22)
        headings = {
            "date": "Date",
            "description": "Description",
            "cheque": "Cheque No.",
            "debit": "Debit",
            "credit": "Credit",
            "balance": "Balance",
        }
        widths = {"date": 120, "description": 360, "cheque": 120, "debit": 120, "credit": 120, "balance": 140}
        for column in columns:
            self.tree.heading(column, text=headings[column])
            self.tree.column(column, width=widths[column], anchor="w" if column in {"date", "description"} else "e")
        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=yscroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        table_wrap.rowconfigure(0, weight=1)
        table_wrap.columnconfigure(0, weight=1)

        log_wrap = ttk.LabelFrame(parent, text="Status Log", style="Section.TLabelframe", padding=10)
        log_wrap.pack(fill="both", expand=False, pady=(10, 0))
        self.log_text = tk.Text(log_wrap, height=9, wrap="word")
        self.log_text.pack(fill="both", expand=True)
        self.append_log("Ready. Generate a statement, then choose Excel and Word formats to export.")

    def append_log(self, message: str) -> None:
        self.log_text.insert("end", message.rstrip() + "\n")
        self.log_text.see("end")

    def _get_statement_period(self) -> tuple[date, date] | None:
        try:
            start = parse_iso_date(self.vars["start_date"].get())
            end = parse_iso_date(self.vars["end_date"].get())
        except Exception:
            return None
        if start > end:
            return None
        return start, end

    def _load_legacy_holiday_dates(self) -> set[str]:
        if not LEGACY_WEB_GENERATOR_PATH.exists():
            return set(REQUIRED_HOLIDAY_DATES)
        try:
            content = LEGACY_WEB_GENERATOR_PATH.read_text(encoding="utf-8", errors="ignore")
        except OSError:
            return set(REQUIRED_HOLIDAY_DATES)
        match = re.search(r"const\s+HOLIDAYS_AND_SATURDAYS\s*=\s*\[(.*?)\];", content, flags=re.DOTALL)
        if not match:
            return set(REQUIRED_HOLIDAY_DATES)

        holidays: set[str] = set()
        for date_text in re.findall(r"\d{4}-\d{2}-\d{2}", match.group(1)):
            try:
                parsed = parse_iso_date(date_text)
            except Exception:
                continue
            if parsed.weekday() != 5:
                holidays.add(parsed.isoformat())
        holidays.update(REQUIRED_HOLIDAY_DATES)
        return holidays

    def _state_file_path(self) -> Path:
        if getattr(sys, "frozen", False):
            base_dir = Path(sys.executable).resolve().parent
        else:
            argv0 = Path(sys.argv[0]).resolve() if sys.argv and sys.argv[0] else Path.cwd()
            base_dir = argv0.parent if argv0.suffix else Path.cwd()
        return base_dir / "statement_generator_state.json"

    def _clean_date_strings(self, values: object, saturday_only: bool = False) -> set[str]:
        cleaned: set[str] = set()
        if not isinstance(values, list):
            return cleaned
        for item in values:
            text = str(item).strip()
            if not text:
                continue
            try:
                parsed = parse_iso_date(text)
            except Exception:
                continue
            if saturday_only and parsed.weekday() != 5:
                continue
            cleaned.add(parsed.isoformat())
        return cleaned

    def _load_persistent_rules(self) -> None:
        path = self._state_file_path()
        if not path.exists():
            return
        try:
            payload = json.loads(path.read_text(encoding="utf-8"))
        except Exception as error:
            self.append_log(f"Could not read saved holiday settings: {error}")
            return

        schema_version = int(payload.get("schema_version", 1))
        if "custom_holidays" in payload:
            self.custom_holiday_dates = self._clean_date_strings(payload.get("custom_holidays", []))
        elif payload.get("holidays"):
            legacy_lines = [item.strip() for item in str(payload.get("holidays", "")).splitlines() if item.strip()]
            self.custom_holiday_dates = self._clean_date_strings(legacy_lines)

        self.excluded_saturday_dates = self._clean_date_strings(payload.get("excluded_saturdays", []), saturday_only=True)

        if schema_version < PERSISTENT_RULES_SCHEMA_VERSION:
            self.custom_holiday_dates.update(REQUIRED_HOLIDAY_DATES)
            self._save_persistent_rules(show_error=False)

    def _save_persistent_rules(self, show_error: bool = True) -> bool:
        payload = {
            "schema_version": PERSISTENT_RULES_SCHEMA_VERSION,
            "custom_holidays": sorted(self.custom_holiday_dates),
            "excluded_saturdays": sorted(self.excluded_saturday_dates),
        }
        path = self._state_file_path()
        try:
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(json.dumps(payload, indent=2), encoding="utf-8")
        except Exception as error:
            if show_error:
                messagebox.showerror(
                    "Save Error",
                    f"Could not save permanent holiday settings.\n\n{error}",
                    parent=self.holiday_manager_window or self,
                )
            self.append_log(f"Could not save holiday settings: {error}")
            return False
        return True

    def _on_app_close(self) -> None:
        self._save_persistent_rules(show_error=False)
        self.destroy()

    def _auto_saturday_strings(self) -> list[str]:
        period = self._get_statement_period()
        if period is None:
            return []
        start, end = period
        end = end + timedelta(days=31)
        dates: list[str] = []
        current = start
        while current <= end:
            if current.weekday() == 5 and current.isoformat() not in self.excluded_saturday_dates:
                dates.append(current.isoformat())
            current += timedelta(days=1)
        return dates

    def _blocked_rule_rows(self, view: str | None = None) -> list[tuple[str, str, str]]:
        rows: list[tuple[str, str, str]] = []
        for holiday in sorted(self.custom_holiday_dates):
            rows.append((f"holiday:{holiday}", holiday, "Holiday"))
        for saturday in self._auto_saturday_strings():
            rows.append((f"saturday:{saturday}", saturday, "Saturday"))
        if view == "Holiday":
            rows = [row for row in rows if row[2] == "Holiday"]
        elif view == "Saturday":
            rows = [row for row in rows if row[2] == "Saturday"]
        return rows

    def _blocked_dates(self) -> set[date]:
        blocked = {parse_iso_date(item) for item in self.custom_holiday_dates}
        blocked.update(parse_iso_date(item) for item in self._auto_saturday_strings())
        return blocked

    def refresh_holiday_display(self) -> None:
        if not hasattr(self, "holiday_text"):
            return
        rows = self._blocked_rule_rows()
        period = self._get_statement_period()
        header = "Blocked dates for current statement period.\n"
        if period is not None:
            header += f"Period: {period[0].isoformat()} to {period[1].isoformat()}\n"
        header += (
            f"Holidays: {len(self.custom_holiday_dates)} | "
            f"Active Saturdays: {len(self._auto_saturday_strings())} | "
            f"Total Blocked: {len(rows)}\n\n"
        )
        if rows:
            preview = "\n".join(f"{date_text}  [{rule_type}]" for _rule_id, date_text, rule_type in rows[:40])
            if len(rows) > 40:
                preview += f"\n... and {len(rows) - 40} more"
        else:
            preview = "No blocked dates."
        self.holiday_text.configure(state="normal")
        self.holiday_text.delete("1.0", "end")
        self.holiday_text.insert("1.0", header + preview)
        self.holiday_text.configure(state="disabled")
        self._refresh_holiday_tree()

    def _set_holiday_view(self, view: str, sync_type: bool = True) -> None:
        self.holiday_view_var.set(view)
        if sync_type and view in {"Holiday", "Saturday"}:
            self.holiday_edit_type_var.set(view)
        self._refresh_holiday_tree()

    def open_holiday_manager(self) -> None:
        if self.holiday_manager_window is not None and self.holiday_manager_window.winfo_exists():
            self.holiday_manager_window.deiconify()
            self.holiday_manager_window.lift()
            self._refresh_holiday_tree()
            return

        window = tk.Toplevel(self)
        window.title("Holiday & Saturday Manager")
        window.geometry("760x560")
        window.minsize(700, 500)
        self.holiday_manager_window = window
        window.protocol("WM_DELETE_WINDOW", self._close_holiday_manager)

        note = ttk.Label(
            window,
            text=(
                "Holidays are loaded from your old HTML list, and Saturdays are generated automatically from the current "
                "statement period. Use the view buttons below to see All, only Holidays, or only Saturdays."
            ),
            justify="left",
            foreground="#374151",
        )
        note.pack(fill="x", padx=12, pady=(12, 8))

        view_bar = ttk.Frame(window, padding=(12, 0, 12, 8))
        view_bar.pack(fill="x")
        ttk.Label(view_bar, text="View").pack(side="left")
        for view_name in ("All", "Holiday", "Saturday"):
            ttk.Radiobutton(
                view_bar,
                text=view_name,
                value=view_name,
                variable=self.holiday_view_var,
                command=lambda selected=view_name: self._set_holiday_view(selected),
                style="Toolbutton",
            ).pack(side="left", padx=(8, 0))
        ttk.Label(view_bar, textvariable=self.holiday_status_var, foreground="#4b5563").pack(side="left", padx=(12, 0))

        tree_wrap = ttk.Frame(window, padding=(12, 0, 12, 8))
        tree_wrap.pack(fill="both", expand=True)
        tree_wrap.columnconfigure(0, weight=1)
        tree_wrap.rowconfigure(0, weight=1)

        self.holiday_tree = ttk.Treeview(tree_wrap, columns=("date", "type"), show="headings", height=16)
        self.holiday_tree.heading("date", text="Date")
        self.holiday_tree.heading("type", text="Type")
        self.holiday_tree.column("date", width=180, anchor="w")
        self.holiday_tree.column("type", width=160, anchor="w")
        self.holiday_tree.grid(row=0, column=0, sticky="nsew")
        holiday_scroll = ttk.Scrollbar(tree_wrap, orient="vertical", command=self.holiday_tree.yview)
        holiday_scroll.grid(row=0, column=1, sticky="ns")
        self.holiday_tree.configure(yscrollcommand=holiday_scroll.set)
        self.holiday_tree.bind("<<TreeviewSelect>>", lambda _event: self._load_selected_holiday_rule())

        editor = ttk.LabelFrame(window, text="Add / Modify Blocked Date", padding=12)
        editor.pack(fill="x", padx=12, pady=(0, 12))
        ttk.Label(editor, text="Date (YYYY-MM-DD)").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(editor, textvariable=self.holiday_edit_date_var, width=24).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(editor, text="Type").grid(row=0, column=2, sticky="w", padx=(16, 8), pady=4)
        ttk.Combobox(
            editor,
            textvariable=self.holiday_edit_type_var,
            values=["Holiday", "Saturday"],
            state="readonly",
            width=14,
        ).grid(row=0, column=3, sticky="w", pady=4)

        button_row = ttk.Frame(editor)
        button_row.grid(row=1, column=0, columnspan=4, sticky="w", pady=(8, 0))
        ttk.Button(button_row, text="Add", command=self.add_holiday_rule).pack(side="left")
        ttk.Button(button_row, text="Modify Selected", command=self.update_selected_holiday_rule).pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="Delete Selected", command=self.delete_selected_holiday_rule).pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="Restore All Saturdays", command=self.restore_all_saturdays).pack(side="left", padx=(8, 0))
        ttk.Button(button_row, text="Close", command=self._close_holiday_manager).pack(side="left", padx=(8, 0))

        self._refresh_holiday_tree()

    def _refresh_holiday_tree(self) -> None:
        if self.holiday_tree is None or not self.holiday_tree.winfo_exists():
            return
        current_view = self.holiday_view_var.get().strip() or "All"
        self.holiday_tree.delete(*self.holiday_tree.get_children())
        rows = self._blocked_rule_rows(current_view)
        for rule_id, date_text, rule_type in rows:
            self.holiday_tree.insert("", "end", iid=rule_id, values=(date_text, rule_type))
        holidays = len(self._blocked_rule_rows("Holiday"))
        saturdays = len(self._blocked_rule_rows("Saturday"))
        self.holiday_status_var.set(
            f"Holidays: {holidays} | Saturdays: {saturdays} | Showing: {len(rows)}"
        )

    def _load_selected_holiday_rule(self) -> None:
        if self.holiday_tree is None:
            return
        selection = self.holiday_tree.selection()
        if not selection:
            return
        item = self.holiday_tree.item(selection[0], "values")
        if len(item) >= 2:
            self.holiday_edit_date_var.set(item[0])
            self.holiday_edit_type_var.set(item[1])

    def _validate_rule_date(self, date_text: str, rule_type: str) -> str:
        parsed = parse_iso_date(date_text)
        if rule_type == "Saturday" and parsed.weekday() != 5:
            raise ValueError("Selected Saturday date must actually be a Saturday.")
        return parsed.isoformat()

    def _remove_rule(self, rule_id: str) -> None:
        rule_type, date_text = rule_id.split(":", 1)
        if rule_type == "holiday":
            self.custom_holiday_dates.discard(date_text)
        elif rule_type == "saturday":
            self.excluded_saturday_dates.add(date_text)

    def _apply_rule(self, date_text: str, rule_type: str) -> None:
        normalized = self._validate_rule_date(date_text, rule_type)
        if rule_type == "Holiday":
            self.custom_holiday_dates.add(normalized)
        else:
            self.excluded_saturday_dates.discard(normalized)

    def add_holiday_rule(self) -> None:
        try:
            self._apply_rule(self.holiday_edit_date_var.get().strip(), self.holiday_edit_type_var.get().strip() or "Holiday")
        except Exception as error:
            messagebox.showerror("Holiday Manager", str(error), parent=self.holiday_manager_window or self)
            return
        self._save_persistent_rules()
        self.refresh_holiday_display()
        self.append_log(f"Added blocked date {self.holiday_edit_date_var.get().strip()} as {self.holiday_edit_type_var.get().strip()}.")

    def update_selected_holiday_rule(self) -> None:
        if self.holiday_tree is None:
            return
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("Holiday Manager", "Select a holiday or Saturday to modify.", parent=self.holiday_manager_window or self)
            return
        try:
            updated_type = self.holiday_edit_type_var.get().strip() or "Holiday"
            updated_date = self._validate_rule_date(self.holiday_edit_date_var.get().strip(), updated_type)
            self._remove_rule(selection[0])
            self._apply_rule(updated_date, updated_type)
        except Exception as error:
            messagebox.showerror("Holiday Manager", str(error), parent=self.holiday_manager_window or self)
            return
        self._save_persistent_rules()
        self.refresh_holiday_display()
        self.append_log(f"Modified blocked date rule to {self.holiday_edit_date_var.get().strip()} [{self.holiday_edit_type_var.get().strip()}].")

    def delete_selected_holiday_rule(self) -> None:
        if self.holiday_tree is None:
            return
        selection = self.holiday_tree.selection()
        if not selection:
            messagebox.showwarning("Holiday Manager", "Select a holiday or Saturday to delete.", parent=self.holiday_manager_window or self)
            return
        item = self.holiday_tree.item(selection[0], "values")
        self._remove_rule(selection[0])
        self._save_persistent_rules()
        self.refresh_holiday_display()
        if len(item) >= 2:
            self.append_log(f"Removed blocked date {item[0]} [{item[1]}].")

    def restore_all_saturdays(self) -> None:
        self.excluded_saturday_dates.clear()
        self._save_persistent_rules()
        self.refresh_holiday_display()
        self.append_log("Restored all Saturdays for the current statement period into the blocked-date list.")

    def _close_holiday_manager(self) -> None:
        if self.holiday_manager_window is not None and self.holiday_manager_window.winfo_exists():
            self.holiday_manager_window.destroy()
        self.holiday_manager_window = None
        self.holiday_tree = None

    def browse_template_dir(self) -> None:
        selected = filedialog.askdirectory(initialdir=self.vars["template_dir"].get() or str(DEFAULT_TEMPLATE_DIR))
        if selected:
            self.vars["template_dir"].set(selected)
            self.refresh_templates()

    def refresh_templates(self) -> None:
        directory = Path(self.vars["template_dir"].get().strip())
        self.catalog = scan_template_directory(directory)
        self.statement_list.delete(0, "end")
        self.certificate_list.delete(0, "end")
        for entry in self.catalog.statement_templates:
            self.statement_list.insert("end", entry.name)
        for entry in self.catalog.certificate_templates:
            self.certificate_list.insert("end", entry.name)
        self.append_log(
            f"Loaded {len(self.catalog.statement_templates)} statement templates and {len(self.catalog.certificate_templates)} certificate templates from {directory}"
        )

    def _parse_optional_date(self, value: str) -> date | None:
        value = value.strip()
        if not value:
            return None
        return parse_iso_date(value)

    def collect_config(self) -> StatementConfig:
        return StatementConfig(
            bank_name=self.vars["bank_name"].get().strip(),
            branch_name=self.vars["branch_name"].get().strip(),
            customer_name=self.vars["customer_name"].get().strip(),
            customer_address=self.vars["customer_address"].get().strip(),
            account_number=self.vars["account_number"].get().strip(),
            account_type=self.vars["account_type"].get().strip(),
            member_id=self.vars["member_id"].get().strip(),
            currency=self.vars["currency"].get().strip() or "NPR",
            reference_no=self.vars["reference_no"].get().strip(),
            opening_date=self._parse_optional_date(self.vars["opening_date"].get()),
            start_date=parse_iso_date(self.vars["start_date"].get()),
            end_date=parse_iso_date(self.vars["end_date"].get()),
            opening_balance=float(self.vars["opening_balance"].get()),
            target_closing_balance=float(self.vars["target_closing_balance"].get()),
            interest_rate=float(self.vars["interest_rate"].get()),
            tax_rate=float(self.vars["tax_rate"].get()),
            cheque_start=int(self.vars["cheque_start"].get()),
            deposit_text=self.vars["deposit_text"].get().strip(),
            withdrawal_text=self.vars["withdrawal_text"].get().strip(),
            interest_text=self.vars["interest_text"].get().strip(),
            tax_text=self.vars["tax_text"].get().strip(),
            first_date_description=self.vars["first_date_description"].get().strip(),
            last_date_description=self.vars["last_date_description"].get().strip(),
            deposit_names=names_from_text(self.deposit_names_text.get("1.0", "end")),
            withdrawal_names=names_from_text(self.withdrawal_names_text.get("1.0", "end")),
            deposit_name_mode=DESCRIPTION_MODE_OPTIONS[self.vars["deposit_mode"].get()],
            withdrawal_name_mode=DESCRIPTION_MODE_OPTIONS[self.vars["withdrawal_mode"].get()],
            holiday_dates=self._blocked_dates(),
            seed=int(self.vars["seed"].get()) if self.vars["seed"].get().strip() else None,
        )

    def generate_statement(self) -> None:
        try:
            config = self.collect_config()
            self.generated_result = generate_statement(config)
        except Exception as error:
            messagebox.showerror("Generation Error", str(error))
            self.append_log(f"Generation failed: {error}")
            return

        self.tree.delete(*self.tree.get_children())
        for row in self.generated_result.rows:
            self.tree.insert(
                "",
                "end",
                values=(
                    row.date.isoformat(),
                    row.description,
                    row.cheque_no,
                    format_amount(row.debit) if row.debit else "",
                    format_amount(row.credit) if row.credit else "",
                    format_amount(row.balance),
                ),
            )

        gap = self.generated_result.final_balance - config.target_closing_balance
        self.summary_vars["final_balance"].set(f"Rs. {format_amount(self.generated_result.final_balance)}")
        self.summary_vars["issue_date"].set(self.generated_result.issue_date.isoformat())
        self.summary_vars["row_count"].set(str(len(self.generated_result.rows)))
        self.summary_vars["target_gap"].set(f"Rs. {format_amount(gap)}")
        self.summary_vars["total_deposits"].set(f"Rs. {format_amount(self.generated_result.summary.total_deposits)}")
        self.summary_vars["total_withdrawals"].set(f"Rs. {format_amount(self.generated_result.summary.total_withdrawals)}")
        self.summary_vars["total_interest"].set(f"Rs. {format_amount(self.generated_result.summary.total_interest)}")
        self.summary_vars["total_tax"].set(f"Rs. {format_amount(self.generated_result.summary.total_tax)}")
        self.summary_vars["seed_used"].set(str(self.generated_result.seed))
        self.append_log(
            f"Generated {len(self.generated_result.rows)} rows. Final balance Rs. {format_amount(self.generated_result.final_balance)}. Issue date {self.generated_result.issue_date.isoformat()}."
        )

    def _selected_template(self, kind: str):
        if kind == "statement":
            selection = self.statement_list.curselection()
            entries = self.catalog.statement_templates
        else:
            selection = self.certificate_list.curselection()
            entries = self.catalog.certificate_templates
        if not selection:
            return None
        return entries[selection[0]]

    def _resolve_rate(self):
        mode = "manual" if self.vars["rate_mode"].get() == "Manual" else "auto"
        manual = float(self.vars["manual_rate"].get()) if self.vars["manual_rate"].get().strip() else None
        return resolve_exchange_rate(
            self.generated_result.issue_date,
            mode=mode,
            manual_rate=manual,
            rate_type=self.vars["rate_type"].get().strip() or "sell",
        )

    def _build_export_payload(self):
        config = self.collect_config()
        rate = self._resolve_rate()
        payload = build_payload(config, self.generated_result, rate)
        return config, rate, payload

    def export_selected_statement(self) -> None:
        template = self._selected_template("statement")
        if self.generated_result is None:
            messagebox.showwarning("Generate First", "Generate the statement before exporting.")
            return
        if template is None:
            messagebox.showwarning("Select Format", "Select a statement Excel format first.")
            return
        try:
            config, rate, payload = self._build_export_payload()
        except Exception as error:
            messagebox.showerror("Rate Error", str(error))
            self.append_log(f"Could not resolve exchange rate: {error}")
            return

        default_name = default_output_name("statement", config.customer_name, template.name, self.generated_result.issue_date, template.path.suffix)
        output_path = filedialog.asksaveasfilename(
            title="Save Statement",
            defaultextension=template.path.suffix,
            initialfile=default_name,
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if not output_path:
            return
        try:
            export_statement(template.path, Path(output_path), payload)
        except Exception as error:
            messagebox.showerror("Export Error", str(error))
            self.append_log(f"Statement export failed: {error}")
            return

        self.append_log(
            f"Statement exported with format '{template.name}' to {output_path}. Exchange rate used: {rate.rate:.2f} ({rate.source_label})."
        )
        messagebox.showinfo("Export Complete", f"Statement exported successfully.\n\n{output_path}")

    def export_selected_certificate(self) -> None:
        template = self._selected_template("certificate")
        if self.generated_result is None:
            messagebox.showwarning("Generate First", "Generate the statement before exporting.")
            return
        if template is None:
            messagebox.showwarning("Select Format", "Select a Word certificate format first.")
            return
        try:
            config, rate, payload = self._build_export_payload()
        except Exception as error:
            messagebox.showerror("Rate Error", str(error))
            self.append_log(f"Could not resolve exchange rate: {error}")
            return

        default_name = default_output_name("certificate", config.customer_name, template.name, self.generated_result.issue_date, template.path.suffix)
        output_path = filedialog.asksaveasfilename(
            title="Save Balance Certificate",
            defaultextension=template.path.suffix,
            initialfile=default_name,
            filetypes=[("Word files", "*.docx *.doc"), ("All files", "*.*")],
        )
        if not output_path:
            return
        try:
            export_certificate(template.path, Path(output_path), payload)
        except Exception as error:
            messagebox.showerror("Export Error", str(error))
            self.append_log(f"Certificate export failed: {error}")
            return

        self.append_log(
            f"Balance certificate exported with format '{template.name}' to {output_path}. Exchange rate used: {rate.rate:.2f} ({rate.source_label})."
        )
        messagebox.showinfo("Export Complete", f"Balance certificate exported successfully.\n\n{output_path}")

    def export_normal_statement_file(self) -> None:
        if self.generated_result is None:
            messagebox.showwarning("Generate First", "Generate the statement before exporting.")
            return
        try:
            config, rate, payload = self._build_export_payload()
        except Exception as error:
            messagebox.showerror("Rate Error", str(error))
            self.append_log(f"Could not resolve exchange rate: {error}")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Normal Statement Excel",
            defaultextension=".xlsx",
            initialfile=default_output_name("statement_normal", config.customer_name, "standard", self.generated_result.issue_date, ".xlsx"),
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if not output_path:
            return
        try:
            export_normal_statement(Path(output_path), payload)
        except Exception as error:
            messagebox.showerror("Normal Export Error", str(error))
            self.append_log(f"Normal Excel export failed: {error}")
            return

        self.append_log(
            f"Normal Excel statement exported to {output_path}. Exchange rate used: {rate.rate:.2f} ({rate.source_label})."
        )
        messagebox.showinfo("Export Complete", f"Normal Excel statement exported successfully.\n\n{output_path}")

    def export_normal_certificate_file(self) -> None:
        if self.generated_result is None:
            messagebox.showwarning("Generate First", "Generate the statement before exporting.")
            return
        try:
            config, rate, payload = self._build_export_payload()
        except Exception as error:
            messagebox.showerror("Rate Error", str(error))
            self.append_log(f"Could not resolve exchange rate: {error}")
            return

        output_path = filedialog.asksaveasfilename(
            title="Save Normal Balance Certificate",
            defaultextension=".docx",
            initialfile=default_output_name("certificate_normal", config.customer_name, "standard", self.generated_result.issue_date, ".docx"),
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if not output_path:
            return
        try:
            export_normal_certificate(Path(output_path), payload)
        except Exception as error:
            messagebox.showerror("Normal Export Error", str(error))
            self.append_log(f"Normal Word export failed: {error}")
            return

        self.append_log(
            f"Normal Word balance certificate exported to {output_path}. Exchange rate used: {rate.rate:.2f} ({rate.source_label})."
        )
        messagebox.showinfo("Export Complete", f"Normal Word balance certificate exported successfully.\n\n{output_path}")

    def _profile_payload(self) -> dict:
        return {
            "vars": {key: var.get() for key, var in self.vars.items() if key != "today"},
            "deposit_names": self.deposit_names_text.get("1.0", "end").strip(),
            "withdrawal_names": self.withdrawal_names_text.get("1.0", "end").strip(),
            "custom_holidays": sorted(self.custom_holiday_dates),
            "excluded_saturdays": sorted(self.excluded_saturday_dates),
        }

    def _apply_profile_payload(self, payload: dict) -> None:
        for key, value in payload.get("vars", {}).items():
            if key in self.vars:
                self.vars[key].set(value)
        self.deposit_names_text.delete("1.0", "end")
        self.deposit_names_text.insert("1.0", payload.get("deposit_names", ""))
        self.withdrawal_names_text.delete("1.0", "end")
        self.withdrawal_names_text.insert("1.0", payload.get("withdrawal_names", ""))
        if "custom_holidays" in payload:
            self.custom_holiday_dates = self._clean_date_strings(payload.get("custom_holidays", []))
        elif payload.get("holidays"):
            legacy_lines = [item.strip() for item in str(payload.get("holidays", "")).splitlines() if item.strip()]
            self.custom_holiday_dates = self._clean_date_strings(legacy_lines)
        else:
            self.custom_holiday_dates = set(self.legacy_holiday_dates)
        self.excluded_saturday_dates = self._clean_date_strings(payload.get("excluded_saturdays", []), saturday_only=True)
        self.holiday_view_var.set("All")
        self._save_persistent_rules(show_error=False)
        self.refresh_holiday_display()

    def save_profile(self) -> None:
        path = filedialog.asksaveasfilename(
            title="Save Profile",
            defaultextension=".json",
            initialfile=f"{safe_filename(self.vars['customer_name'].get() or 'profile')}_statement_profile.json",
            filetypes=[("JSON files", "*.json")],
        )
        if not path:
            return
        Path(path).write_text(json.dumps(self._profile_payload(), indent=2), encoding="utf-8")
        self.append_log(f"Saved profile to {path}")

    def load_profile(self) -> None:
        path = filedialog.askopenfilename(title="Load Profile", filetypes=[("JSON files", "*.json")])
        if not path:
            return
        payload = json.loads(Path(path).read_text(encoding="utf-8"))
        self._apply_profile_payload(payload)
        self.refresh_templates()
        self.append_log(f"Loaded profile from {path}")

    def run_self_tests(self) -> None:
        result = run_tests()
        if result.wasSuccessful():
            self.append_log("Self tests passed.")
            messagebox.showinfo("Self Tests", "All self tests passed.")
        else:
            self.append_log("Self tests failed. Check terminal output for details.")
            messagebox.showwarning("Self Tests", "Some self tests failed. Check the test output.")


def main() -> None:
    app = StatementGeneratorApp()
    app.mainloop()


if __name__ == "__main__":
    main()
