"""AI đọc hợp đồng / booking của 1 lô và tạo file JSON (chưa động tới master).

Cần:  pip3 install anthropic python-docx openpyxl jsonschema   và biến môi trường ANTHROPIC_API_KEY

Ví dụ:
  python3 extract.py OP01943 --purchase "HD mua HP.docx" --sales "HD ban JT.pdf" \
        --booking "booking.pdf" --stuffing 09/11/2026 --supplier HP --buyer JT
  -> lots/OP01943.json   (mở ra xem/sửa, rồi:  python3 import_json.py lots/OP01943.json)

Thêm --import để tự nhập vào master luôn, --generate để sinh luôn chứng từ.
Hỗ trợ file: .docx .pdf .png .jpg .xlsx .txt  (bản scan là ảnh/PDF đều đọc được)
"""
import argparse
import base64
import datetime as dt
import json
import os
import subprocess
import sys
from pathlib import Path

import openpyxl

from lot_schema import SCHEMA

DEFAULT_MODEL = os.environ.get("FARMGATE_MODEL", "claude-sonnet-5-5")

SYSTEM = """Bạn là trợ lý nhập liệu cho bộ phận xuất khẩu của Farmgate Pte. Ltd. (Singapore, công ty thương mại).
Bạn đọc hợp đồng mua (Farmgate mua của nhà cung cấp - NCC), hợp đồng bán (Farmgate bán cho Buyer) và booking
confirmation rồi điền dữ liệu của MỘT lô hàng bằng công cụ `record_lot`.

Quy tắc bắt buộc:
- Chỉ ghi thứ nhìn thấy trong tài liệu. Không thấy thì để null. Tuyệt đối không đoán hay suy ra.
- Ngày: YYYY-MM-DD. Số tiền/giá: số thuần (không ký hiệu). Tỷ lệ cọc 50% -> 0.5.
- Giá mua (buy_price_usd_kg) lấy từ HỢP ĐỒNG MUA; giá bán (sell_price_usd_kg) lấy từ HỢP ĐỒNG BÁN. Cùng một mặt hàng
  ghép hai giá vào cùng một dòng `lines`. Số carton lấy theo hợp đồng (nếu hai hợp đồng khác nhau, ghi carton
  của hợp đồng mua và đưa 'lines[n].cartons' vào uncertain_fields).
- product_code phải là một trong các mã trong danh sách Products được cung cấp. Không khớp được thì ghi mã bạn
  thấy trong tài liệu (ví dụ LP2) và đưa vào uncertain_fields.
- pay_type: tiền cọc + còn lại khi có bộ chứng từ scan -> DEPOSIT_BALANCE_ON_DOCS; cọc + còn lại khi hàng đến ->
  DEPOSIT_BALANCE_ON_ARRIVAL; thanh toán trong N ngày kể từ ngày B/L -> DAYS_FROM_BL (điền pay_days).
- etd_text: chép nguyên văn điều khoản ETD của hợp đồng mua (ví dụ "Within November 2026.").
- surveyor_cert: YES nếu hợp đồng mua có dòng yêu cầu surveyor/inspection certificate, ngược lại NO.
- booking.etd là ETD thực tế trên booking; vessel_voyage ví dụ "ESL OMAN V.02647W"; ports viết HOA như trên booking.
- Không điền stuffing_date (người dùng cung cấp). Không điền container_no/seal_no nếu booking chưa có.
- uncertain_fields: liệt kê tên trường bạn không chắc hoặc chữ mờ/khó đọc. evidence: trích dẫn ngắn (dưới 15 từ)
  cho giá, số lượng, ngày, cảng để người kiểm tra đối chiếu nhanh.
- supplier/buyer: chỉ điền thông tin pháp nhân (tên, địa chỉ, mã số, đại diện, ngân hàng) đọc từ hợp đồng."""


def docx_text(path):
    import docx
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    d = docx.Document(path)
    out = []
    for el in d.element.body.iterchildren():
        if el.tag.endswith("}p"):
            t = Paragraph(el, d).text.strip()
            if t:
                out.append(t)
        elif el.tag.endswith("}tbl"):
            for row in Table(el, d).rows:
                cells, prev = [], None
                for c in row.cells:
                    if c._tc is not prev:
                        cells.append(c.text.strip().replace("\n", " / "))
                    prev = c._tc
                out.append(" | ".join(cells))
    return "\n".join(out)


def xlsx_text(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    out = []
    for ws in wb:
        out.append(f"## {ws.title}")
        for row in ws.iter_rows(values_only=True):
            if any(v is not None for v in row):
                out.append(" | ".join("" if v is None else str(v) for v in row))
    return "\n".join(out)


def file_blocks(label, path):
    p = Path(path)
    ext = p.suffix.lower()
    head = {"type": "text", "text": f"=== {label}: {p.name} ==="}
    if ext == ".docx":
        return [head, {"type": "text", "text": docx_text(p)}]
    if ext == ".xlsx":
        return [head, {"type": "text", "text": xlsx_text(p)}]
    if ext == ".txt":
        return [head, {"type": "text", "text": p.read_text(encoding="utf8", errors="replace")}]
    data = base64.standard_b64encode(p.read_bytes()).decode()
    if ext == ".pdf":
        return [head, {"type": "document", "source": {"type": "base64", "media_type": "application/pdf", "data": data}}]
    if ext in (".png", ".jpg", ".jpeg"):
        mt = "image/png" if ext == ".png" else "image/jpeg"
        return [head, {"type": "image", "source": {"type": "base64", "media_type": mt, "data": data}}]
    raise SystemExit(f"Không hỗ trợ định dạng {ext} ({p.name})")


def master_hints(master_path):
    wb = openpyxl.load_workbook(master_path, data_only=True)
    prods = [f"{r[0]} ({r[3]})" for r in wb["Products"].iter_rows(min_row=2, values_only=True) if r[0]]
    parties = [f"{r[0]} = {r[2]} [{r[1]}]" for r in wb["Parties"].iter_rows(min_row=2, values_only=True) if r[0]]
    return prods, parties


def parse_date(s):
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    raise ValueError(f"Ngày '{s}' sai định dạng (dùng dd/mm/yyyy)")


def build_request(a):
    prods, parties = master_hints(a.master)
    content = [{"type": "text", "text":
                f"Lô hàng: {a.op_code}\nMã đối tác có sẵn:\n- " + "\n- ".join(parties) +
                "\nSản phẩm có sẵn (Products):\n- " + "\n- ".join(prods) +
                (f"\nNCC của lô này: {a.supplier}" if a.supplier else "") +
                (f"\nBuyer của lô này: {a.buyer}" if a.buyer else "") +
                "\nHãy đọc các tài liệu sau và gọi record_lot."}]
    for label, path in (("HỢP ĐỒNG MUA (NCC)", a.purchase), ("HỢP ĐỒNG BÁN (Buyer)", a.sales), ("BOOKING", a.booking)):
        if path:
            content += file_blocks(label, path)
    if len(content) == 1:
        raise SystemExit("Cần ít nhất một tài liệu (--purchase / --sales / --booking)")
    return {
        "model": a.model, "max_tokens": 4000, "system": SYSTEM,
        "tools": [{"name": "record_lot", "description": "Ghi dữ liệu trích xuất của một lô hàng",
                   "input_schema": SCHEMA}],
        "tool_choice": {"type": "tool", "name": "record_lot"},
        "messages": [{"role": "user", "content": content}],
    }


def run_extraction(a, req=None):
    """Gọi AI và trả về dict dữ liệu lô (đã gắn op_code, mã đối tác, ngày đóng hàng)."""
    import anthropic
    req = req or build_request(a)
    resp = anthropic.Anthropic().messages.create(**req)
    tool = next((b for b in resp.content if b.type == "tool_use"), None)
    if tool is None:
        raise RuntimeError("AI không trả về dữ liệu có cấu trúc. Thử lại hoặc kiểm tra tài liệu.")
    data = dict(tool.input)
    data["op_code"] = a.op_code
    if a.supplier:
        data["supplier_code"] = a.supplier
    if a.buyer:
        data["buyer_code"] = a.buyer
    if a.stuffing:
        data["stuffing_date"] = parse_date(a.stuffing)
    return data


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("op_code")
    ap.add_argument("--purchase", help="Hợp đồng mua với NCC")
    ap.add_argument("--sales", help="Hợp đồng bán cho Buyer")
    ap.add_argument("--booking", help="Booking confirmation")
    ap.add_argument("--stuffing", help="Ngày đóng hàng dd/mm/yyyy (bạn cung cấp)")
    ap.add_argument("--supplier", help="Mã NCC trong Parties, ví dụ HP")
    ap.add_argument("--buyer", help="Mã Buyer trong Parties, ví dụ JT")
    ap.add_argument("--master", default="master.xlsx")
    ap.add_argument("--model", default=DEFAULT_MODEL)
    ap.add_argument("--out", help="File JSON đầu ra (mặc định lots/<OP>.json)")
    ap.add_argument("--dry-run", action="store_true", help="Chỉ in những gì sẽ gửi, không gọi API")
    ap.add_argument("--import", dest="do_import", action="store_true", help="Nhập luôn vào master sau khi trích xuất")
    ap.add_argument("--generate", action="store_true", help="Sinh luôn chứng từ (kèm --import)")
    a = ap.parse_args()

    req = build_request(a)
    if a.dry_run:
        for b in req["messages"][0]["content"]:
            if b["type"] == "text":
                print(b["text"][:1500], "\n---")
            else:
                print(f"[{b['type']} base64 {len(b['source']['data'])} ký tự]\n---")
        return 0
    if not os.environ.get("ANTHROPIC_API_KEY"):
        raise SystemExit("Chưa có ANTHROPIC_API_KEY (export ANTHROPIC_API_KEY=sk-ant-...). "
                         "Hoặc dùng --dry-run để kiểm tra trước.")
    try:
        data = run_extraction(a, req)
    except (RuntimeError, ValueError) as ex:
        raise SystemExit(str(ex))
    out = Path(a.out or f"lots/{a.op_code}.json")
    out.parent.mkdir(exist_ok=True)
    out.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf8")
    print("Đã lưu:", out)
    if data.get("uncertain_fields"):
        print("AI không chắc:", ", ".join(data["uncertain_fields"]))
    if a.do_import:
        rc = subprocess.run([sys.executable, "import_json.py", str(out), "--master", a.master, "--recalc"]).returncode
        if rc == 0 and a.generate:
            subprocess.run([sys.executable, "generate.py", a.op_code, "--master", a.master])
    else:
        print(f"Kiểm tra file JSON rồi chạy:  python3 import_json.py {out} --master {a.master}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
