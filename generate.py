#!/usr/bin/env python3
"""Sinh bộ chứng từ của một lô hàng từ master.xlsx.

Cách chạy:
    python3 generate.py OP01943
    python3 generate.py OP01943 --master ../master.xlsx --out output

Kết quả nằm trong output/<mã OP>/ (8 chứng từ: hợp đồng mua, hợp đồng bán, PL/CI/SI phía NCC,
PL/CI phía Buyer, PoD). Chứng từ cần ngày đóng hàng sẽ bị bỏ qua nếu Execution chưa có lô đó.
"""
import argparse
import sys
from pathlib import Path

from docxtpl import DocxTemplate
from jinja2 import Environment

import common
from xlsx_engine import render_workbook

HERE = Path(__file__).parent
TPL = HERE / "templates"


def render_docx(tpl, out, ctx):
    d = DocxTemplate(str(tpl))
    d.render(ctx, jinja_env=Environment(autoescape=True))
    d.save(str(out))


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("op_code")
    ap.add_argument("--master", default=str(HERE / "master.xlsx"))
    ap.add_argument("--out", default=str(HERE / "output"))
    a = ap.parse_args()

    try:
        ctx = common.build_context(a.master, a.op_code)
    except common.MasterError as e:
        sys.exit(f"LỖI DỮ LIỆU: {e}")

    out = Path(a.out) / a.op_code
    out.mkdir(parents=True, exist_ok=True)
    sup, buy = ctx["sup"]["code"], ctx["buy"]["code"]
    pay = ctx["pay"]
    made, skipped = [], []

    def job(name, ok, reason, fn):
        if not ok:
            skipped.append(f"{name}: {reason}")
            return
        fn()
        made.append(name)

    purchase_ok = "unsupported_purchase" not in pay
    sales_ok = "unsupported_sales" not in pay
    post_ok = ctx["has_exe"]
    no_exe = "chưa có ngày đóng hàng trong sheet Execution"

    job(f"{a.op_code} SC {sup} - FGS.docx", purchase_ok,
        f"chưa có mẫu cho điều khoản thanh toán phía NCC ({pay.get('unsupported_purchase')})",
        lambda: render_docx(TPL / "Hop_dong_mua_NCC.docx", out / f"{a.op_code} SC {sup} - FGS.docx", ctx))
    job(f"{a.op_code} SC FGS - {buy}.docx",
        sales_ok and "unsupported_sales_contract" not in pay,
        f"chưa có mẫu hợp đồng cho điều khoản thanh toán phía Buyer ({pay.get('unsupported_sales') or pay.get('unsupported_sales_contract')})",
        lambda: render_docx(TPL / "Hop_dong_ban_Buyer.docx", out / f"{a.op_code} SC FGS - {buy}.docx", ctx))
    job(f"{a.op_code} PL + CI + SI {sup}-FGS.xlsx", purchase_ok and post_ok,
        no_exe if not post_ok else "chưa có mẫu điều khoản thanh toán phía NCC",
        lambda: render_workbook(TPL / "PL_CI_SI_NCC.xlsx", out / f"{a.op_code} PL + CI + SI {sup}-FGS.xlsx", ctx))
    job(f"{a.op_code} INV + PKL - {buy} FGS.xlsx", sales_ok and post_ok,
        no_exe if not post_ok else "chưa có mẫu điều khoản thanh toán phía Buyer",
        lambda: render_workbook(TPL / "PL_CI_Buyer.xlsx", out / f"{a.op_code} INV + PKL - {buy} FGS.xlsx", ctx))
    job(f"{a.op_code} Proof of delivery {sup}.docx", post_ok, no_exe,
        lambda: render_docx(TPL / "PoD_NCC.docx", out / f"{a.op_code} Proof of delivery {sup}.docx", ctx))

    print(f"Đã tạo trong {out}:")
    for m in made:
        print("  +", m)
    for s_ in skipped:
        print("  - bỏ qua", s_)
    print(f"Tổng mua {ctx['tot']['buy_fmt']} USD, tổng bán {ctx['tot']['sell_fmt']} USD, "
          f"lợi nhuận gộp {ctx['margin']:,.2f} USD.")


if __name__ == "__main__":
    main()
