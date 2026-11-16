"""Điền template Excel có placeholder {{ ... }} và tự nhân dòng hàng.

Quy ước trong template:
  * Ô chứa {{ ... }} được thay bằng giá trị. Nếu cả ô chỉ là một placeholder, giữ nguyên kiểu số.
  * Dòng nào chứa {{ it.xxx }} là dòng mẫu của mặt hàng; sẽ được nhân lên theo số mặt hàng.
  * Công thức dùng thẻ {r} (số dòng hiện tại), {first} và {last} (dòng hàng đầu và cuối).
    Các tham chiếu ô thường (vd =G11-G12) tự dịch xuống theo số dòng hàng thêm vào.
"""
import re
from copy import copy

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from jinja2 import Environment

_env = Environment(autoescape=False, keep_trailing_newline=True)
_single = re.compile(r"^\{\{\s*((?:(?!\{\{|\}\}).)+?)\s*\}\}\Z", re.S)
_ref = re.compile(r"(\$?[A-Z]{1,3})(\$?)(\d+)")


def _render(value, ctx):
    if isinstance(value, CellRichText):
        parts = []
        for b in value:
            if isinstance(b, TextBlock):
                parts.append(TextBlock(b.font, str(_render(b.text, ctx))))
            else:
                parts.append(str(_render(b, ctx)))
        return CellRichText(*parts)
    if not isinstance(value, str) or "{{" not in value:
        return value
    m = _single.match(value)
    if m:
        return _env.compile_expression(m.group(1))(**ctx)
    return _env.from_string(value).render(**ctx)


def _shift_formula(f, item_row, extra):
    if not extra:
        return f

    def sub(m):
        row = int(m.group(3))
        if row > item_row:
            row += extra
        return f"{m.group(1)}{m.group(2)}{row}"
    return _ref.sub(sub, f)


def _insert_rows(ws, idx, amount, tmpl_row):
    old_merges = [(m.min_col, m.min_row, m.max_col, m.max_row) for m in ws.merged_cells.ranges]
    for m in list(ws.merged_cells.ranges):
        ws.unmerge_cells(str(m))
    old_max = ws.max_row
    heights = {r: ws.row_dimensions[r].height for r in range(1, old_max + 1)}
    ws.insert_rows(idx, amount)
    for r in range(1, old_max + amount + 1):
        if r < idx:
            h = heights.get(r)
        elif r < idx + amount:
            h = heights.get(tmpl_row)
        else:
            h = heights.get(r - amount)
        ws.row_dimensions[r].height = h
    for c1, r1, c2, r2 in old_merges:
        if r1 >= idx:
            ws.merge_cells(start_column=c1, start_row=r1 + amount, end_column=c2, end_row=r2 + amount)
        else:
            ws.merge_cells(start_column=c1, start_row=r1, end_column=c2, end_row=r2)
            if r1 == tmpl_row == r2:
                for k in range(amount):
                    ws.merge_cells(start_column=c1, start_row=idx + k, end_column=c2, end_row=idx + k)
    for k in range(amount):
        for col in range(1, ws.max_column + 1):
            src = ws.cell(row=tmpl_row, column=col)
            dst = ws.cell(row=idx + k, column=col)
            dst._style = copy(src._style)


def fill_sheet(ws, ctx):
    items = ctx["items"]
    item_row = None
    for row in ws.iter_rows():
        for c in row:
            if isinstance(c.value, str) and "{{" in c.value and re.search(r"\{\{\s*it\.", c.value):
                item_row = c.row
                break
        if item_row:
            break
    extra = 0
    tmpl_cells = {}
    if item_row:
        extra = len(items) - 1
        tmpl_cells = {c.column: c.value for c in ws[item_row] if c.value is not None}
        if extra > 0:
            _insert_rows(ws, item_row + 1, extra, item_row)
    first = item_row
    last = item_row + extra if item_row else None

    for row in list(ws.iter_rows()):
        for c in row:
            if c.value is None:
                continue
            if item_row and item_row <= c.row <= last:
                continue  # dòng hàng xử lý riêng bên dưới
            if isinstance(c.value, str) and c.value.startswith("="):
                f = _shift_formula(c.value, item_row, extra) if item_row else c.value
                if item_row:
                    f = f.replace("{first}", str(first)).replace("{last}", str(last))
                c.value = f
            else:
                c.value = _render(c.value, ctx)

    if item_row:
        for k, it in enumerate(items):
            r = item_row + k
            for col, tv in tmpl_cells.items():
                cell = ws.cell(row=r, column=col)
                if isinstance(tv, str) and tv.startswith("="):
                    cell.value = tv.replace("{r}", str(r))
                else:
                    cell.value = _render(tv, {**ctx, "it": it})


def render_workbook(template_path, out_path, ctx):
    wb = openpyxl.load_workbook(template_path, rich_text=True)
    for ws in wb.worksheets:
        fill_sheet(ws, ctx)
    wb.save(out_path)
