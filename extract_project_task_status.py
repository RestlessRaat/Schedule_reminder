#!/usr/bin/env python3
"""Extract leaf-task status from a Microsoft Project exported XLSX file."""

from __future__ import annotations

import argparse
import ctypes
import json
import math
import re
import sys
import zipfile
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path, PurePosixPath
from typing import Any
from xml.etree import ElementTree as ET


DEFAULT_SHEET_NAME = "任务_表"
SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_INPUT_DIR = SCRIPT_DIR / "Data"
DEFAULT_OUTPUT_DIR = SCRIPT_DIR / "outputs"
REQUIRED_COLUMNS = {
    "标识号",
    "活动",
    "任务模式",
    "名称",
    "工期",
    "开始时间",
    "完成时间",
    "前置任务",
    "大纲级别",
    "备注",
}
DATE_RE = re.compile(
    r"^\s*(?P<year>\d{4})年(?P<month>\d{1,2})月(?P<day>\d{1,2})日"
    r"(?:\s+(?P<hour>\d{1,2}):(?P<minute>\d{2}))?\s*$"
)
ISO_DATE_ARG_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
CELL_REF_RE = re.compile(r"^([A-Z]+)(\d+)$")
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PACKAGE_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

STATUS_LATE = "滞后"
STATUS_IN_PROGRESS = "进行中"
STATUS_NOT_STARTED = "未开始"
STATUS_UNKNOWN = "未知"
STATUS_ORDER = [STATUS_LATE, STATUS_IN_PROGRESS, STATUS_NOT_STARTED, STATUS_UNKNOWN]
TASK_FILTER_ATTENTION = "滞后+进行中"
TASK_FILTER_ALL = "全部"
TASK_FILTER_OPTIONS = [
    TASK_FILTER_ATTENTION,
    STATUS_LATE,
    STATUS_IN_PROGRESS,
    STATUS_NOT_STARTED,
    TASK_FILTER_ALL,
]
TASK_TABLE_COLUMNS = (
    "status",
    "name",
    "days_late",
    "start_date",
    "finish_date",
    "duration_text",
    "task_id",
    "parent_path",
)
TASK_TABLE_HEADINGS = {
    "status": "状态",
    "name": "任务名称",
    "days_late": "滞后天数",
    "start_date": "开始日期",
    "finish_date": "完成日期",
    "duration_text": "工期",
    "task_id": "任务ID",
    "parent_path": "父级路径",
}
TASK_TABLE_WIDTHS = {
    "status": 80,
    "name": 280,
    "days_late": 80,
    "start_date": 100,
    "finish_date": 100,
    "duration_text": 80,
    "task_id": 70,
    "parent_path": 460,
}


def select_input_xlsx() -> Path | None:
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception as exc:
        print(f"File selection dialog is not available: {exc}", file=sys.stderr)
        return None

    initial_dir = DEFAULT_INPUT_DIR if DEFAULT_INPUT_DIR.exists() else SCRIPT_DIR
    root = tk.Tk()
    root.withdraw()
    try:
        root.attributes("-topmost", True)
        selected = filedialog.askopenfilename(
            title="选择 Microsoft Project 导出的 Excel 文件",
            initialdir=str(initial_dir),
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("All files", "*.*"),
            ],
        )
    finally:
        root.destroy()

    if not selected:
        return None
    return Path(selected)


def default_output_path_for(input_path: Path) -> Path:
    return DEFAULT_OUTPUT_DIR / f"{input_path.stem}_task_status.json"


def clean_value(value: Any) -> Any:
    """Convert empty values into JSON-friendly nulls."""
    if value is None:
        return None
    if isinstance(value, float) and math.isnan(value):
        return None
    if hasattr(value, "item"):
        try:
            return value.item()
        except (TypeError, ValueError):
            pass
    return value


def clean_text(value: Any) -> str | None:
    value = clean_value(value)
    if value is None:
        return None
    text = str(value).strip()
    return text or None


def clean_int(value: Any, field_name: str) -> int:
    value = clean_value(value)
    if value is None:
        raise ValueError(f"Missing required numeric value for {field_name}.")
    try:
        return int(value)
    except (TypeError, ValueError) as exc:
        raise ValueError(f"Invalid numeric value for {field_name}: {value!r}") from exc


def parse_project_date(value: Any) -> date | None:
    value = clean_value(value)
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    text = str(value).strip()
    if not text:
        return None

    if isinstance(value, (int, float)):
        serial = float(value)
        if 1 <= serial <= 100000:
            return date(1899, 12, 30) + timedelta(days=int(serial))

    match = DATE_RE.match(text)
    if match:
        parts = {key: int(val) for key, val in match.groupdict(default="0").items()}
        return date(parts["year"], parts["month"], parts["day"])

    for fmt in ("%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M", "%Y-%m-%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            continue
    return None


def iso_date(value: date | None) -> str | None:
    return value.isoformat() if value else None


def column_index(cell_ref: str) -> int:
    match = CELL_REF_RE.match(cell_ref)
    if not match:
        raise ValueError(f"Invalid cell reference: {cell_ref!r}")
    letters = match.group(1)
    index = 0
    for letter in letters:
        index = index * 26 + (ord(letter) - ord("A") + 1)
    return index - 1


def xml_text(element: ET.Element | None) -> str | None:
    if element is None:
        return None
    parts = [text for text in element.itertext() if text]
    return "".join(parts) if parts else None


def normalize_workbook_target(target: str) -> str:
    if target.startswith("/"):
        return str(PurePosixPath(target.lstrip("/")))
    return str(PurePosixPath("xl") / target)


def list_xlsx_sheets(xlsx_path: Path) -> dict[str, str]:
    with zipfile.ZipFile(xlsx_path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))

    rels = {
        rel.attrib["Id"]: normalize_workbook_target(rel.attrib["Target"])
        for rel in rels_root.findall(f"{{{NS_PACKAGE_REL}}}Relationship")
        if "Id" in rel.attrib and "Target" in rel.attrib
    }

    sheets: dict[str, str] = {}
    for sheet in workbook_root.findall(f".//{{{NS_MAIN}}}sheet"):
        name = sheet.attrib.get("name")
        rel_id = sheet.attrib.get(f"{{{NS_REL}}}id")
        if name and rel_id and rel_id in rels:
            sheets[name] = rels[rel_id]
    return sheets


def load_shared_strings(archive: zipfile.ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []

    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for item in root.findall(f"{{{NS_MAIN}}}si"):
        strings.append(xml_text(item) or "")
    return strings


def parse_cell_value(cell: ET.Element, shared_strings: list[str]) -> Any:
    cell_type = cell.attrib.get("t")

    if cell_type == "inlineStr":
        return xml_text(cell.find(f"{{{NS_MAIN}}}is"))

    value_node = cell.find(f"{{{NS_MAIN}}}v")
    value_text = value_node.text if value_node is not None else None
    if value_text is None:
        return None

    if cell_type == "s":
        try:
            return shared_strings[int(value_text)]
        except (IndexError, ValueError):
            return value_text
    if cell_type == "b":
        return value_text == "1"
    if cell_type in {"str", "e"}:
        return value_text

    try:
        number = float(value_text)
    except ValueError:
        return value_text
    if number.is_integer():
        return int(number)
    return number


def read_xlsx_rows(xlsx_path: Path, sheet_name: str) -> list[list[Any]]:
    sheets = list_xlsx_sheets(xlsx_path)
    if sheet_name not in sheets:
        raise ValueError(
            f"Worksheet {sheet_name!r} not found. Available sheets: {', '.join(sheets)}"
        )

    with zipfile.ZipFile(xlsx_path) as archive:
        shared_strings = load_shared_strings(archive)
        root = ET.fromstring(archive.read(sheets[sheet_name]))

    rows: list[list[Any]] = []
    for row in root.findall(f".//{{{NS_MAIN}}}sheetData/{{{NS_MAIN}}}row"):
        values: list[Any] = []
        for cell in row.findall(f"{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r")
            if not ref:
                continue
            zero_based_col = column_index(ref)
            while len(values) <= zero_based_col:
                values.append(None)
            values[zero_based_col] = parse_cell_value(cell, shared_strings)
        rows.append(values)
    return rows


def rows_to_records(rows: list[list[Any]], sheet_name: str) -> list[dict[str, Any]]:
    if not rows:
        raise ValueError(f"Worksheet {sheet_name!r} is empty.")

    headers = [clean_text(value) for value in rows[0]]
    header_set = {header for header in headers if header}
    missing = sorted(REQUIRED_COLUMNS - header_set)
    if missing:
        raise ValueError(f"Missing required columns in {sheet_name!r}: {', '.join(missing)}")

    records: list[dict[str, Any]] = []
    for row in rows[1:]:
        if not any(clean_value(value) is not None for value in row):
            continue
        record: dict[str, Any] = {}
        for index, header in enumerate(headers):
            if not header:
                continue
            record[header] = row[index] if index < len(row) else None
        records.append(record)
    return records


def choose_sheet_name(xlsx_path: Path, requested_name: str | None) -> str:
    sheets = list_xlsx_sheets(xlsx_path)
    sheet_names = list(sheets)

    if requested_name:
        if requested_name in sheets:
            return requested_name
        raise ValueError(
            f"Worksheet {requested_name!r} not found. Available sheets: {', '.join(sheet_names)}"
        )

    if DEFAULT_SHEET_NAME in sheets:
        return DEFAULT_SHEET_NAME

    candidates: list[str] = []
    for sheet_name in sheet_names:
        try:
            rows = read_xlsx_rows(xlsx_path, sheet_name)
        except Exception:
            continue
        header = {clean_text(value) for value in rows[0]} if rows else set()
        if REQUIRED_COLUMNS.issubset(header):
            candidates.append(sheet_name)

    if len(candidates) == 1:
        return candidates[0]
    if candidates:
        raise ValueError(
            "Multiple sheets look like task tables. Specify --sheet-name. "
            f"Candidates: {', '.join(candidates)}"
        )

    raise ValueError(
        f"Worksheet {DEFAULT_SHEET_NAME!r} not found, and no sheet contains the required task columns. "
        f"Available sheets: {', '.join(sheet_names)}"
    )


def read_task_table(xlsx_path: Path, sheet_name: str) -> list[dict[str, Any]]:
    return rows_to_records(read_xlsx_rows(xlsx_path, sheet_name), sheet_name)


def classify_status(start_date: date | None, finish_date: date | None, analysis_date: date) -> str:
    if start_date is None or finish_date is None:
        return STATUS_UNKNOWN
    if analysis_date > finish_date:
        return STATUS_LATE
    if start_date <= analysis_date <= finish_date:
        return STATUS_IN_PROGRESS
    if analysis_date < start_date:
        return STATUS_NOT_STARTED
    return STATUS_UNKNOWN


def build_tasks(task_rows: list[dict[str, Any]], analysis_date: date) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    records: list[dict[str, Any]] = []
    stack_by_level: dict[int, dict[str, Any]] = {}

    for index, row in enumerate(task_rows):
        task_id = clean_int(row["标识号"], "标识号")
        outline_level = clean_int(row["大纲级别"], "大纲级别")
        start = parse_project_date(row["开始时间"])
        finish = parse_project_date(row["完成时间"])
        status = classify_status(start, finish, analysis_date)
        next_level = None
        if index + 1 < len(task_rows):
            next_level = clean_int(task_rows[index + 1]["大纲级别"], "大纲级别")
        is_summary = next_level is not None and next_level > outline_level

        parent_levels = sorted(level for level in stack_by_level if level < outline_level)
        parent_chain = [stack_by_level[level] for level in parent_levels]
        parent = parent_chain[-1] if parent_chain else None
        top_parent = parent_chain[0] if parent_chain else None

        task = {
            "task_id": task_id,
            "name": clean_text(row["名称"]),
            "duration_text": clean_text(row["工期"]),
            "start_text": clean_text(row["开始时间"]),
            "finish_text": clean_text(row["完成时间"]),
            "predecessors_text": clean_text(row["前置任务"]),
            "outline_level": outline_level,
            "task_mode": clean_text(row["任务模式"]),
            "active": clean_text(row["活动"]),
            "notes": clean_text(row["备注"]),
            "start_date": iso_date(start),
            "finish_date": iso_date(finish),
            "is_summary": is_summary,
            "status": status,
            "days_late": (analysis_date - finish).days if status == STATUS_LATE and finish else None,
            "parent_task_id": parent["task_id"] if parent else None,
            "parent_chain_ids": [item["task_id"] for item in parent_chain],
            "parent_chain_names": [item["name"] for item in parent_chain],
            "task_path": " > ".join(
                [*(item["name"] for item in parent_chain if item["name"]), clean_text(row["名称"]) or ""]
            ),
            "top_parent_task_id": top_parent["task_id"] if top_parent else None,
            "top_parent_name": top_parent["name"] if top_parent else None,
        }

        records.append(task)

        stale_levels = [level for level in stack_by_level if level >= outline_level]
        for level in stale_levels:
            del stack_by_level[level]
        stack_by_level[outline_level] = {"task_id": task_id, "name": task["name"]}

    leaf_tasks = [task for task in records if not task["is_summary"]]
    summary_tasks = [task for task in records if task["is_summary"]]
    return leaf_tasks, summary_tasks


def summarize(
    leaf_tasks: list[dict[str, Any]],
    summary_tasks: list[dict[str, Any]],
) -> dict[str, Any]:
    status_counts = Counter(task["status"] for task in leaf_tasks)

    return {
        "total_task_count": len(leaf_tasks) + len(summary_tasks),
        "summary_task_count": len(summary_tasks),
        "leaf_task_count": len(leaf_tasks),
        "leaf_task_status_counts": {status: status_counts.get(status, 0) for status in STATUS_ORDER},
        "by_top_parent": group_leaf_tasks(leaf_tasks, parent_kind="top"),
        "by_immediate_parent": group_leaf_tasks(leaf_tasks, parent_kind="immediate"),
    }


def group_leaf_tasks(
    leaf_tasks: list[dict[str, Any]],
    parent_kind: str,
) -> list[dict[str, Any]]:
    parent_summary: dict[str, dict[str, Any]] = defaultdict(
        lambda: {
            "parent_task_id": None,
            "parent_name": None,
            "parent_path": None,
            "leaf_task_count": 0,
            "late_count": 0,
            "in_progress_count": 0,
            "status_counts": {status: 0 for status in STATUS_ORDER},
        }
    )

    for task in leaf_tasks:
        if parent_kind == "top":
            parent_id = task["top_parent_task_id"] or task["parent_task_id"] or task["task_id"]
            parent_name = task["top_parent_name"] or (
                task["parent_chain_names"][0] if task["parent_chain_names"] else task["name"]
            )
            parent_path = parent_name
        elif parent_kind == "immediate":
            parent_id = task["parent_task_id"] or task["task_id"]
            parent_name = task["parent_chain_names"][-1] if task["parent_chain_names"] else task["name"]
            parent_path = " > ".join(task["parent_chain_names"]) or task["name"]
        else:
            raise ValueError(f"Unknown parent grouping kind: {parent_kind}")

        parent_key = str(parent_id)
        bucket = parent_summary[parent_key]
        bucket["parent_task_id"] = parent_id
        bucket["parent_name"] = parent_name
        bucket["parent_path"] = parent_path
        bucket["leaf_task_count"] += 1
        bucket["status_counts"][task["status"]] += 1
        if task["status"] == STATUS_LATE:
            bucket["late_count"] += 1
        elif task["status"] == STATUS_IN_PROGRESS:
            bucket["in_progress_count"] += 1

    return sorted(
        parent_summary.values(),
        key=lambda item: (-item["late_count"], -item["in_progress_count"], str(item["parent_name"])),
    )


def make_output(
    xlsx_path: Path,
    sheet_name: str,
    analysis_date: date,
    leaf_tasks: list[dict[str, Any]],
    summary_tasks: list[dict[str, Any]],
    include_summary_tasks: bool,
) -> dict[str, Any]:
    output = {
        "source_file": str(xlsx_path),
        "sheet_name": sheet_name,
        "analysis_date": analysis_date.isoformat(),
        "summary": summarize(leaf_tasks, summary_tasks),
        "leaf_tasks": leaf_tasks,
    }
    if include_summary_tasks:
        output["summary_tasks"] = summary_tasks
    return output


def print_task_list(title: str, tasks: list[dict[str, Any]]) -> None:
    print(f"\n{title} ({len(tasks)}):")
    if not tasks:
        print("  无")
        return

    for task in tasks:
        parent_path = " > ".join(task["parent_chain_names"]) or "无"
        finish_text = task["finish_date"] or task["finish_text"] or "未知"
        days_late = task["days_late"]
        late_suffix = f", 滞后 {days_late} 天" if days_late is not None else ""
        print(
            f"  [{task['status']}] {task['name']} | 父级路径: {parent_path} "
            f"| 完成: {finish_text}{late_suffix}"
        )


def print_console_summary(result: dict[str, Any]) -> None:
    summary = result["summary"]
    counts = summary["leaf_task_status_counts"]
    print(f"Source: {result['source_file']}")
    print(f"Sheet: {result['sheet_name']}")
    print(f"Analysis date: {result['analysis_date']}")
    print(
        "Tasks: "
        f"total={summary['total_task_count']}, "
        f"summary={summary['summary_task_count']}, "
        f"leaf={summary['leaf_task_count']}"
    )
    print(
        "Leaf status: "
        + ", ".join(f"{status}={counts.get(status, 0)}" for status in STATUS_ORDER)
    )

    late_tasks = [task for task in result["leaf_tasks"] if task["status"] == STATUS_LATE]
    in_progress_tasks = [
        task for task in result["leaf_tasks"] if task["status"] == STATUS_IN_PROGRESS
    ]
    late_tasks.sort(key=lambda task: (-(task["days_late"] or 0), task["task_id"]))
    in_progress_tasks.sort(key=lambda task: (task["finish_date"] or "", task["task_id"]))

    print_task_list("滞后叶子任务", late_tasks)
    print_task_list("进行中叶子任务", in_progress_tasks)


def task_sort_key(task: dict[str, Any]) -> tuple[Any, ...]:
    status = task.get("status")
    task_id = task.get("task_id") or 0
    if status == STATUS_LATE:
        return (0, -(task.get("days_late") or 0), task_id)
    if status == STATUS_IN_PROGRESS:
        return (1, task.get("finish_date") or "", task_id)
    if status == STATUS_NOT_STARTED:
        return (2, task.get("start_date") or "", task_id)
    return (3, task_id)


def filter_leaf_tasks_for_table(
    result: dict[str, Any],
    filter_name: str = TASK_FILTER_ATTENTION,
) -> list[dict[str, Any]]:
    if filter_name == TASK_FILTER_ATTENTION:
        statuses = {STATUS_LATE, STATUS_IN_PROGRESS}
    elif filter_name == TASK_FILTER_ALL:
        statuses = None
    else:
        statuses = {filter_name}

    tasks = result.get("leaf_tasks", [])
    if statuses is not None:
        tasks = [task for task in tasks if task.get("status") in statuses]
    return sorted(tasks, key=task_sort_key)


def task_parent_path(task: dict[str, Any]) -> str:
    return " > ".join(task.get("parent_chain_names") or []) or "无"


def task_copy_path(task: dict[str, Any]) -> str:
    return task.get("task_path") or task.get("name") or ""


def task_table_values(task: dict[str, Any]) -> tuple[str, ...]:
    days_late = task.get("days_late")
    return (
        str(task.get("status") or ""),
        str(task.get("name") or ""),
        "" if days_late is None else str(days_late),
        str(task.get("start_date") or ""),
        str(task.get("finish_date") or ""),
        str(task.get("duration_text") or ""),
        str(task.get("task_id") or ""),
        task_parent_path(task),
    )


def show_success_popup(result: dict[str, Any], output_path: Path) -> None:
    summary = result["summary"]
    counts = summary["leaf_task_status_counts"]
    message = "\n".join(
        [
            f"分析日期: {result['analysis_date']}",
            f"叶子任务总数: {summary['leaf_task_count']}",
            f"滞后任务: {counts.get(STATUS_LATE, 0)}",
            f"进行中任务: {counts.get(STATUS_IN_PROGRESS, 0)}",
            f"未开始任务: {counts.get(STATUS_NOT_STARTED, 0)}",
            "",
            f"JSON 输出: {output_path}",
        ]
    )

    try:
        ctypes.windll.user32.MessageBoxW(
            None,
            message,
            "Project 任务状态分析完成",
            0x00000040,
        )
    except Exception as exc:
        print(f"Popup notification skipped: {exc}", file=sys.stderr)


def show_error_popup(message: str) -> None:
    try:
        ctypes.windll.user32.MessageBoxW(
            None,
            message,
            "Project 任务状态分析错误",
            0x00000010,
        )
    except Exception as exc:
        print(f"Error popup skipped: {exc}", file=sys.stderr)


def report_error(message: str, exit_code: int, no_popup: bool = False) -> int:
    print(message, file=sys.stderr)
    if not no_popup:
        show_error_popup(message)
    return exit_code


def show_results_window(result: dict[str, Any], output_path: Path) -> None:
    try:
        import tkinter as tk
        from tkinter import ttk

        summary = result["summary"]
        counts = summary["leaf_task_status_counts"]

        root = tk.Tk()
        root.title("Project 任务状态分析结果")
        root.geometry("1200x700")
        root.minsize(900, 520)
        root.columnconfigure(0, weight=1)
        root.rowconfigure(2, weight=1)

        summary_text = (
            f"分析日期: {result['analysis_date']}    "
            f"叶子任务: {summary['leaf_task_count']}    "
            f"滞后: {counts.get(STATUS_LATE, 0)}    "
            f"进行中: {counts.get(STATUS_IN_PROGRESS, 0)}    "
            f"JSON: {output_path}"
        )
        ttk.Label(root, text=summary_text, padding=(12, 10, 12, 6), anchor="w").grid(
            row=0,
            column=0,
            sticky="ew",
        )

        toolbar = ttk.Frame(root, padding=(12, 0, 12, 8))
        toolbar.grid(row=1, column=0, sticky="ew")
        toolbar.columnconfigure(2, weight=1)

        ttk.Label(toolbar, text="状态筛选:").grid(row=0, column=0, sticky="w")
        filter_var = tk.StringVar(value=TASK_FILTER_ATTENTION)
        filter_combo = ttk.Combobox(
            toolbar,
            textvariable=filter_var,
            values=TASK_FILTER_OPTIONS,
            state="readonly",
            width=14,
        )
        filter_combo.grid(row=0, column=1, padx=(8, 0), sticky="w")

        table_frame = ttk.Frame(root, padding=(12, 0, 12, 8))
        table_frame.grid(row=2, column=0, sticky="nsew")
        table_frame.columnconfigure(0, weight=1)
        table_frame.rowconfigure(0, weight=1)

        tree = ttk.Treeview(
            table_frame,
            columns=TASK_TABLE_COLUMNS,
            show="headings",
            selectmode="browse",
        )
        for column in TASK_TABLE_COLUMNS:
            tree.heading(column, text=TASK_TABLE_HEADINGS[column])
            tree.column(
                column,
                width=TASK_TABLE_WIDTHS[column],
                minwidth=60,
                stretch=column in {"name", "parent_path"},
                anchor="w",
            )

        tree.tag_configure(STATUS_LATE, foreground="#B00020")
        tree.tag_configure(STATUS_IN_PROGRESS, foreground="#075985")
        tree.tag_configure(STATUS_NOT_STARTED, foreground="#475569")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

        footer = ttk.Frame(root, padding=(12, 0, 12, 12))
        footer.grid(row=3, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        status_var = tk.StringVar(value="")
        ttk.Label(footer, textvariable=status_var, anchor="w").grid(row=0, column=0, sticky="ew")

        item_to_task: dict[str, dict[str, Any]] = {}

        def refresh_table(*_: Any) -> None:
            item_to_task.clear()
            for item_id in tree.get_children():
                tree.delete(item_id)

            tasks = filter_leaf_tasks_for_table(result, filter_var.get())
            for task in tasks:
                item_id = tree.insert(
                    "",
                    "end",
                    values=task_table_values(task),
                    tags=(task.get("status") or "",),
                )
                item_to_task[item_id] = task
            status_var.set(f"当前显示 {len(tasks)} 条任务")

        def copy_selected_task_path() -> None:
            selected = tree.selection()
            if not selected:
                status_var.set("请先选择一条任务")
                return
            task = item_to_task.get(selected[0])
            if not task:
                status_var.set("未找到选中任务")
                return
            text = task_copy_path(task)
            root.clipboard_clear()
            root.clipboard_append(text)
            status_var.set("已复制任务路径")

        filter_combo.bind("<<ComboboxSelected>>", refresh_table)
        tree.bind("<Double-1>", lambda _event: copy_selected_task_path())

        ttk.Button(footer, text="复制任务路径", command=copy_selected_task_path).grid(
            row=0,
            column=1,
            padx=(12, 8),
            sticky="e",
        )
        ttk.Button(footer, text="关闭", command=root.destroy).grid(row=0, column=2, sticky="e")

        refresh_table()
        root.after(100, root.lift)
        root.mainloop()
    except Exception as exc:
        print(f"Result table window skipped: {exc}", file=sys.stderr)
        show_success_popup(result, output_path)


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extract and analyze leaf-task status from a Microsoft Project exported XLSX file."
    )
    parser.add_argument(
        "input_xlsx",
        nargs="?",
        type=Path,
        help=(
            "Path to the Project exported .xlsx file. "
            "When omitted, a file selection dialog is shown."
        ),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=None,
        help="Path to write the JSON result. Defaults to outputs/<input filename>_task_status.json.",
    )
    parser.add_argument(
        "--sheet-name",
        default=None,
        help=f"Task sheet name. Defaults to {DEFAULT_SHEET_NAME!r} when present.",
    )
    parser.add_argument(
        "--date",
        default=date.today().isoformat(),
        help="Analysis date in YYYY-MM-DD format. Defaults to today.",
    )
    parser.add_argument(
        "--include-summary-tasks",
        action="store_true",
        help="Include summary task records in the JSON output.",
    )
    parser.add_argument(
        "--no-popup",
        action="store_true",
        help="Disable the success popup notification.",
    )
    return parser.parse_args(argv)


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv or sys.argv[1:])
    try:
        if not ISO_DATE_ARG_RE.match(args.date):
            raise ValueError
        analysis_date = date.fromisoformat(args.date)
    except ValueError:
        return report_error(
            f"Invalid --date value: {args.date!r}. Expected YYYY-MM-DD.",
            2,
            args.no_popup,
        )

    input_path = args.input_xlsx or select_input_xlsx()
    if input_path is None:
        return report_error("No input file selected. Analysis cancelled.", 2, args.no_popup)

    xlsx_path = input_path.expanduser().resolve()
    output_path = (args.output or default_output_path_for(xlsx_path)).expanduser().resolve()

    if not xlsx_path.exists():
        return report_error(f"Input file not found: {xlsx_path}", 2, args.no_popup)

    try:
        sheet_name = choose_sheet_name(xlsx_path, args.sheet_name)
        df = read_task_table(xlsx_path, sheet_name)
        leaf_tasks, summary_tasks = build_tasks(df, analysis_date)
        result = make_output(
            xlsx_path=xlsx_path,
            sheet_name=sheet_name,
            analysis_date=analysis_date,
            leaf_tasks=leaf_tasks,
            summary_tasks=summary_tasks,
            include_summary_tasks=args.include_summary_tasks,
        )
        output_path.parent.mkdir(parents=True, exist_ok=True)
        output_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception as exc:
        return report_error(f"Failed to extract task status: {exc}", 1, args.no_popup)

    print_console_summary(result)
    print(f"\nWrote JSON: {output_path}")
    if not args.no_popup:
        show_results_window(result, output_path)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
