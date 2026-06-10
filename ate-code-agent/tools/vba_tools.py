"""
VBA tools for reading and writing IG-XL VBA modules.
Uses pywin32 and requires Windows + Excel.
"""

import os
import ssl
import httpx
from pathlib import Path

from safety.rules import SAFETY_RULES
from tools.git_tools import git_commit_backup

os.environ["PYTHONHTTPSVERIFY"] = "0"
os.environ["CURL_CA_BUNDLE"] = ""
os.environ["REQUESTS_CA_BUNDLE"] = ""
ssl._create_default_https_context = ssl._create_unverified_context

try:
    import win32com.client  # type: ignore
except Exception:  # pragma: no cover - environment dependent
    win32com = None


def _require_windows_excel() -> None:
    if os.name != "nt" or win32com is None:
        raise RuntimeError("VBA tools require Windows with pywin32 and Excel installed")


def get_vba_module_names(excel_path: str) -> list[str]:
    """List all VBA module names in the workbook."""
    _require_windows_excel()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        return [comp.Name for comp in wb.VBProject.VBComponents]
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


def read_vba_module(excel_path: str, module_name: str) -> str:
    """Read full source code of one VBA module."""
    _require_windows_excel()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        for component in wb.VBProject.VBComponents:
            if component.Name == module_name:
                line_count = component.CodeModule.CountOfLines
                if line_count == 0:
                    return ""
                return component.CodeModule.Lines(1, line_count)
        raise ValueError(f"VBA module '{module_name}' not found in {excel_path}")
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()


def write_vba_module(excel_path: str, module_name: str, new_code: str) -> dict:
    """Replace the entire source code of a VBA module."""
    _require_windows_excel()

    if SAFETY_RULES["git_backup_enabled"]:
        git_commit_backup(excel_path, f"VBA backup before edit: {module_name}")

    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    wb = None
    try:
        wb = excel.Workbooks.Open(str(Path(excel_path).resolve()))
        component = wb.VBProject.VBComponents(module_name)
        cm = component.CodeModule
        if cm.CountOfLines > 0:
            cm.DeleteLines(1, cm.CountOfLines)
        if new_code.strip():
            cm.InsertLines(1, new_code)
        wb.Save()
        lines = len(new_code.splitlines())
        return {"module": module_name, "lines_written": lines, "status": "ok"}
    finally:
        try:
            if wb is not None:
                wb.Close(SaveChanges=False)
        except Exception:
            pass
        excel.Quit()
