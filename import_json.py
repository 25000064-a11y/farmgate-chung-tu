"""Nhập JSON của 1 lô vào master.xlsx (Shipments, Line_Items, Execution, và Parties nếu đối tác mới).

Cách dùng:  python3 import_json.py lots/OP01943.json [--master master.xlsx] [--recalc]
Quy tắc:  giá trị null trong JSON KHÔNG ghi đè ô đang có; giá trị khác null sẽ ghi (kèm báo thay đổi).
"""
import argparse
import datetime as dt
import json
import shutil
import subprocess
import sys
from pathlib import Path

import openpyxl
from jsonschema import Draft202012Validator

from lot_schema import SCHEMA

MAX_SHIP_ROW = 21     # khớp vùng công thức/validation trong master
MAX_LINE_ROW = 41


class ImportError_(Exception):
    pass


def _hdr(ws):
    return {str(c.value).split(" ")[0].strip(): c.column for c in ws[1] if c.value}


def _date(v):
    if v in (None, ""):
        return None
    try:
        return dt.datetime.strptime(v, "%Y-%m-%d")
    except ValueError:
        raise ImportError_(f"Ngày '{v}' sai định dạng, cần YYYY-MM-DD")


def _find_row(ws, op, last):
    first_empty = None
    for r in range(2, last + 1):
        v = ws.cell(r, 1).value
        if v == op:
            return r, True
        if v in (None, "") and first_empty is None:
            first_empty = r
    if first_empty is None:
        raise ImportError_(f"Sheet {ws.title} hết dòng trống (tối đa {last - 1} lô)")
    return first_empty, False


def _set(ws, row, col_map, field, value, changes, label=None, conv=None):
    if value is None or field not in col_map:
        return
    if conv:
        value = conv(value)
    cell = ws.cell(row, col_map[field])
    old = cell.value
    if old != value:
        changes.append((f"{ws.title}.{label or field}", old, value))
        cell.value = value


def _pct(v):
    return v / 100 if v > 1 else v


def _next_seq(wb, party_code):
    """Giống công thức next_order_seq trong Parties."""
    last = 0
    for r in range(2, wb["Parties"].max_row + 1):
        if wb["Parties"].cell(r, 1).value == party_code:
            last = wb["Parties"].cell(r, 18).value or 0
    ws = wb["Shipments"]
    for r in range(2, MAX_SHIP_ROW + 1):
        if ws.cell(r, 2).value == party_code:
            last = max(last, ws.cell(r, 5).value or 0)
        if ws.cell(r, 3).value == party_code:
            last = max(last, ws.cell(r, 7).value or 0)
    return last + 1


def _ensure_party(wb, code, info, role, changes, warnings):
    ws = wb["Parties"]
    for r in range(2, ws.max_row + 1):
        if ws.cell(r, 1).value == code:
            return
    if not info or not info.get("legal_name"):
        raise ImportError_(f"Đối tác '{code}' chưa có trong sheet Parties và JSON không có thông tin để thêm. "
                           f"Hãy thêm 1 dòng vào Parties rồi chạy lại.")
    hm = _hdr(ws)
    row = next((r for r in range(2, ws.max_row + 2) if ws.cell(r, 1).value in (None, "")), ws.max_row + 1)
    ws.cell(row, hm["party_code"]).value = code
    ws.cell(row, hm["role"]).value = role
    for k, v in info.items():
        if k in hm and v:
            ws.cell(row, hm[k]).value = v
    ws.cell(row, hm["sign_name"]).value = info["legal_name"]
    for c in range(1, ws.max_column + 1):   # copy style/công thức từ dòng trên
        src = ws.cell(row - 1, c)
        if src.has_style:
            ws.cell(row, c)._style = src._style
    nxt = ws.cell(row - 1, ws.max_column).value
    if isinstance(nxt, str) and nxt.startswith("="):
        ws.cell(row, ws.max_column).value = nxt.replace(str(row - 1), str(row))
    changes.append(("Parties", None, f"THÊM đối tác mới {code} - {info['legal_name']}"))
    warnings.append(f"Đối tác mới {code}: kiểm tra lại sheet Parties (tên ký, vendor_code, số TK ngân hàng...)")


def import_lot(data, master_path):
    errs = sorted(Draft202012Validator(SCHEMA).iter_errors(data), key=lambda e: list(e.path))
    if errs:
        raise ImportError_("JSON sai cấu trúc:\n  " + "\n  ".join(f"{'/'.join(map(str, e.path))}: {e.message}" for e in errs))
    op = data.get("op_code")
    if not op:
        raise ImportError_("Thiếu op_code")
    wb = openpyxl.load_workbook(master_path)
    changes, warnings = [], []

    products = {wb["Products"].cell(r, 1).value for r in range(2, wb["Products"].max_row + 1)} - {None}
    for ln in data["lines"]:
        if ln["product_code"] not in products:
            raise ImportError_(f"Sản phẩm '{ln['product_code']}' chưa có trong sheet Products "
                               f"(đang có: {', '.join(sorted(products))}). Thêm sản phẩm vào Products trước.")

    sup, buy = data.get("supplier_code"), data.get("buyer_code")
    ws = wb["Shipments"]
    row, existed = _find_row(ws, op, MAX_SHIP_ROW)
    hm = _hdr(ws)
    if not existed:
        ws.cell(row, 1).value = op
        changes.append(("Shipments", None, f"THÊM lô {op} (dòng {row})"))
    if not sup and not existed:
        raise ImportError_("Thiếu supplier_code cho lô mới")
    if not buy and not existed:
        raise ImportError_("Thiếu buyer_code cho lô mới")
    if sup:
        _ensure_party(wb, sup, data.get("supplier"), "SUPPLIER", changes, warnings)
    if buy:
        _ensure_party(wb, buy, data.get("buyer"), "BUYER", changes, warnings)
    _set(ws, row, hm, "supplier_code", sup, changes)
    _set(ws, row, hm, "buyer_code", buy, changes)
    sup = sup or ws.cell(row, hm["supplier_code"]).value
    buy = buy or ws.cell(row, hm["buyer_code"]).value

    pc, sc = data.get("purchase_contract") or {}, data.get("sales_contract") or {}
    _set(ws, row, hm, "purchase_contract_date", pc.get("date"), changes, conv=_date)
    _set(ws, row, hm, "purchase_contract_seq", pc.get("seq"), changes)
    _set(ws, row, hm, "sales_contract_seq", sc.get("seq"), changes)
    _set(ws, row, hm, "purchase_incoterm", pc.get("incoterm"), changes)
    _set(ws, row, hm, "purchase_incoterm_place", pc.get("incoterm_place"), changes)
    _set(ws, row, hm, "purchase_pay_type", pc.get("pay_type"), changes)
    _set(ws, row, hm, "purchase_deposit_pct", pc.get("deposit_pct"), changes, conv=_pct)
    _set(ws, row, hm, "purchase_etd_text", pc.get("etd_text"), changes)
    _set(ws, row, hm, "purchase_pod", pc.get("pod"), changes)
    _set(ws, row, hm, "purchase_surveyor_cert", pc.get("surveyor_cert"), changes)
    _set(ws, row, hm, "sales_contract_date", sc.get("date"), changes, conv=_date)
    for k_json, k_col in (("incoterm", "sales_incoterm"), ("incoterm_place", "sales_incoterm_place"),
                          ("pay_type", "sales_pay_type"), ("pay_days", "sales_pay_days"),
                          ("delivery_period", "sales_delivery_period"), ("loading_port", "sales_loading_port"),
                          ("place_of_delivery", "sales_place_of_delivery"), ("valid_until", "sales_valid_until"),
                          ("certificates", "sales_certificates"), ("preservation", "sales_preservation"),
                          ("other", "sales_other")):
        _set(ws, row, hm, k_col, sc.get(k_json), changes)
    _set(ws, row, hm, "sales_deposit_pct", sc.get("deposit_pct"), changes, conv=_pct)

    # số thứ tự hợp đồng: giữ số đã có; lô mới -> số kế tiếp của đối tác
    for seq_col, party, has_date in (("purchase_contract_seq", sup, pc.get("date")),
                                     ("sales_contract_seq", buy, sc.get("date"))):
        if ws.cell(row, hm[seq_col]).value in (None, "") and party and has_date:
            n = _next_seq(wb, party)
            ws.cell(row, hm[seq_col]).value = n
            changes.append((f"Shipments.{seq_col}", None, f"{n} (số kế tiếp của {party})"))

    # ---- Line_Items
    wl = wb["Line_Items"]
    lh = _hdr(wl)
    for r in range(2, MAX_LINE_ROW + 1):          # xoá dòng cũ của lô này (nhập lại toàn bộ)
        if wl.cell(r, 1).value == op:
            for c in ("op_code", "line_no", "product_code", "cartons", "buy_price_usd_kg", "sell_price_usd_kg"):
                wl.cell(r, lh[c]).value = None
    free = [r for r in range(2, MAX_LINE_ROW + 1) if wl.cell(r, 1).value in (None, "")]
    if len(free) < len(data["lines"]):
        raise ImportError_("Line_Items hết dòng trống")
    for i, (ln, r) in enumerate(zip(data["lines"], free), 1):
        wl.cell(r, lh["op_code"]).value = op
        wl.cell(r, lh["line_no"]).value = i
        wl.cell(r, lh["product_code"]).value = ln["product_code"]
        wl.cell(r, lh["cartons"]).value = ln.get("cartons")
        wl.cell(r, lh["buy_price_usd_kg"]).value = ln.get("buy_price_usd_kg")
        wl.cell(r, lh["sell_price_usd_kg"]).value = ln.get("sell_price_usd_kg")
        changes.append(("Line_Items", None, f"dòng {i}: {ln['product_code']} {ln.get('cartons')} ctn, "
                                            f"mua {ln.get('buy_price_usd_kg')} / bán {ln.get('sell_price_usd_kg')}"))
        for k in ("cartons", "buy_price_usd_kg", "sell_price_usd_kg"):
            if ln.get(k) is None:
                warnings.append(f"Line_Items dòng {i} ({ln['product_code']}): thiếu {k}")

    # ---- Execution
    we = wb["Execution"]
    erow, eexisted = _find_row(we, op, MAX_SHIP_ROW)
    eh = _hdr(we)
    if not eexisted:
        we.cell(erow, 1).value = op
    bk = data.get("booking") or {}
    _set(we, erow, eh, "stuffing_date", data.get("stuffing_date"), changes, conv=_date)
    _set(we, erow, eh, "etd", bk.get("etd"), changes, conv=_date)
    for k in ("vessel_voyage", "booking_no", "port_of_loading", "place_of_delivery", "port_of_discharge",
              "container_no", "seal_no"):
        _set(we, erow, eh, k, bk.get(k), changes)

    # ---- kiểm tra thiếu
    need_ship = ["purchase_contract_date", "purchase_pay_type", "purchase_incoterm", "purchase_incoterm_place",
                 "sales_contract_date", "sales_pay_type", "sales_incoterm", "sales_incoterm_place"]
    for k in need_ship:
        if ws.cell(row, hm[k]).value in (None, ""):
            warnings.append(f"Shipments.{k} còn trống")
    for k in ("stuffing_date", "etd", "vessel_voyage", "booking_no", "port_of_loading", "port_of_discharge"):
        if we.cell(erow, eh[k]).value in (None, ""):
            warnings.append(f"Execution.{k} còn trống" + (" (cần người dùng cung cấp)" if k == "stuffing_date" else ""))
    for f in data.get("uncertain_fields") or []:
        warnings.append(f"AI không chắc: {f} -> kiểm tra với bản gốc")
    return wb, changes, warnings


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("json_file")
    ap.add_argument("--master", default="master.xlsx")
    ap.add_argument("--out", help="Ghi ra file khác thay vì ghi đè master")
    ap.add_argument("--recalc", action="store_true", help="Cho LibreOffice tính lại công thức (để xem trước được trên Mac)")
    ap.add_argument("--dry-run", action="store_true", help="Chỉ in thay đổi, không lưu")
    a = ap.parse_args(argv)
    data = json.loads(Path(a.json_file).read_text(encoding="utf8"))
    try:
        wb, changes, warnings = import_lot(data, a.master)
    except ImportError_ as e:
        print("LỖI:", e)
        return 1
    print(f"== Thay đổi trong master cho lô {data.get('op_code')} ==")
    for where, old, new in changes:
        print(f"  {where}: {'(trống)' if old in (None, '') else old} -> {new}")
    if warnings:
        print("== CẦN KIỂM TRA ==")
        for w in warnings:
            print("  !", w)
    if a.dry_run:
        return 0
    out = a.out or a.master
    if out == a.master:
        shutil.copy(a.master, a.master + ".bak")
    wb.save(out)
    if a.recalc:
        rc = next(iter(Path("/root/.claude/skills").glob("**/xlsx/scripts/recalc.py")), None)
        if rc:
            subprocess.run([sys.executable, str(rc), out, "60"], check=False)
        else:
            print("(không có recalc.py - bỏ qua; mở bằng Excel sẽ tự tính)")
    print("Đã lưu:", out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
