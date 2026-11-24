# -*- coding: utf-8 -*-
"""
tabular_reader.py
-----------------
Bộ đọc bảng tính đa định dạng cho các file kết xuất từ TCT (HĐĐT), TMS, eTax, SAP, MISA, FAST:
Tự động nhận diện theo nội dung thực (magic bytes) thay vì chỉ tin vào đuôi file:
1. Excel hiện đại (.xlsx - ZIP magic PK\\x03\\x04) -> openpyxl
2. Excel 97-2003 (.xls - OLE2 magic \\xd0\\xcf\\x11\\xe0) -> xlrd
3. SpreadsheetML XML 2003 đổi đuôi .xls (<?xml ... Workbook) -> ElementTree
4. Bảng HTML đổi đuôi .xls (<html / <table / <!doctype) -> HTMLParser
5. Văn bản phân tách Tab / Phẩy / Chấm phẩy đổi đuôi .xls hoặc .csv (hỗ trợ cả UTF-16 SAP/TMS, UTF-8-SIG, UTF-8, CP1258) -> csv.reader
"""

from __future__ import annotations

import csv
import datetime
import io
import os
import xml.etree.ElementTree as ET
from html.parser import HTMLParser
from typing import Any, Dict, List

import openpyxl


class _HtmlTableParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.rows: List[List[str]] = []
        self._cur_row: List[str] = []
        self._in_cell = False
        self._buf: List[str] = []

    def handle_starttag(self, tag: str, attrs) -> None:
        t = tag.lower()
        if t == "tr":
            self._cur_row = []
        elif t in ("td", "th"):
            self._in_cell = True
            self._buf = []

    def handle_endtag(self, tag: str) -> None:
        t = tag.lower()
        if t in ("td", "th"):
            if self._in_cell:
                self._cur_row.append("".join(self._buf).strip())
                self._in_cell = False
        elif t == "tr":
            if self._cur_row:
                self.rows.append(self._cur_row)
                self._cur_row = []

    def handle_data(self, data: str) -> None:
        if self._in_cell:
            self._buf.append(data)


def _decode_bytes(raw: bytes) -> str:
    if raw.startswith((b"\xff\xfe", b"\xfe\xff")):
        return raw.decode("utf-16", errors="replace")
    if b"\x00" in raw[:200]:
        try:
            return raw.decode("utf-16-le", errors="replace")
        except Exception:
            pass
    for enc in ("utf-8-sig", "utf-8", "cp1258", "cp1252", "latin-1"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    return raw.decode("utf-8", errors="replace")


def _read_xlsx(file_path: str) -> Dict[str, List[List[Any]]]:
    wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
    out: Dict[str, List[List[Any]]] = {}
    try:
        for name in wb.sheetnames:
            ws = wb[name]
            out[name] = [list(r) for r in ws.iter_rows(values_only=True)]
    finally:
        wb.close()
    return out


def _read_xls_biff(file_path: str) -> Dict[str, List[List[Any]]]:
    import xlrd

    book = xlrd.open_workbook(file_path)
    out: Dict[str, List[List[Any]]] = {}
    for sheet in book.sheets():
        rows: List[List[Any]] = []
        for r_idx in range(sheet.nrows):
            row_vals: List[Any] = []
            for c_idx in range(sheet.ncols):
                cell = sheet.cell(r_idx, c_idx)
                val = cell.value
                if cell.ctype == xlrd.XL_CELL_DATE:
                    try:
                        dt_tuple = xlrd.xldate_as_tuple(val, book.datemode)
                        val = datetime.datetime(*dt_tuple)
                    except Exception:
                        pass
                row_vals.append(val)
            rows.append(row_vals)
        out[sheet.name] = rows
    return out


def _read_spreadsheetml_xml(text: str) -> Dict[str, List[List[Any]]]:
    root = ET.fromstring(text.strip())
    out: Dict[str, List[List[Any]]] = {}
    idx = 0
    for ws in root.iter():
        if ws.tag.split("}")[-1].lower() == "worksheet":
            idx += 1
            sname = next((v for k, v in ws.attrib.items() if k.split("}")[-1].lower() == "name"), f"Sheet{idx}")
            rows: List[List[Any]] = []
            for row_el in ws.iter():
                if row_el.tag.split("}")[-1].lower() == "row":
                    row_vals: List[Any] = []
                    for cell_el in row_el:
                        if cell_el.tag.split("}")[-1].lower() != "cell":
                            continue
                        idx_attr = next((v for k, v in cell_el.attrib.items() if k.split("}")[-1].lower() == "index"), None)
                        if idx_attr and str(idx_attr).isdigit():
                            target_pos = int(idx_attr) - 1
                            while len(row_vals) < target_pos:
                                row_vals.append("")
                        data_txt = ""
                        for sub in cell_el:
                            if sub.tag.split("}")[-1].lower() == "data":
                                data_txt = sub.text or ""
                                break
                        row_vals.append(data_txt)
                    rows.append(row_vals)
            out[sname] = rows
    return out


def _read_delimited_text(text: str) -> Dict[str, List[List[Any]]]:
    sample = text[:8192]
    tab_cnt = sample.count("\t")
    comma_cnt = sample.count(",")
    semi_cnt = sample.count(";")
    if tab_cnt >= comma_cnt and tab_cnt >= semi_cnt and tab_cnt > 0:
        delim = "\t"
    elif semi_cnt > comma_cnt and semi_cnt > 0:
        delim = ";"
    else:
        delim = ","
    reader = csv.reader(io.StringIO(text), delimiter=delim)
    return {"Sheet1": [list(r) for r in reader]}


def read_tabular_sheets(file_path: str) -> Dict[str, List[List[Any]]]:
    """
    Đọc mọi định dạng file bảng tính (.xlsx, .xls, HTML-xls, XML-xls, TSV/CSV-xls, .csv)
    trả về dict {sheet_name: list_of_rows}.
    """
    if not os.path.isfile(file_path):
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    with open(file_path, "rb") as f:
        head = f.read(16)

    if head.startswith(b"PK\x03\x04"):
        return _read_xlsx(file_path)

    if head.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
        return _read_xls_biff(file_path)

    with open(file_path, "rb") as f:
        raw = f.read()
    text = _decode_bytes(raw)
    stripped = text.lstrip()
    low_head = stripped[:600].lower()

    if "<?xml" in low_head and "<workbook" in low_head:
        try:
            res = _read_spreadsheetml_xml(stripped)
            if res:
                return res
        except Exception:
            pass

    if any(tag in low_head for tag in ("<html", "<!doctype html", "<table", "<tr", "<meta")):
        parser = _HtmlTableParser()
        parser.feed(text)
        if parser.rows:
            return {"Sheet1": parser.rows}

    return _read_delimited_text(text)


def read_first_sheet_rows(file_path: str) -> List[List[Any]]:
    sheets = read_tabular_sheets(file_path)
    if not sheets:
        return []
    first_key = next(iter(sheets.keys()))
    return sheets[first_key]
