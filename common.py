"""Đọc master.xlsx và dựng dữ liệu (context) để điền chứng từ.

Chương trình KHÔNG tin các ô công thức trong master (vì file có thể chưa được Excel tính lại);
mọi giá trị tính toán được tính lại ở đây từ các ô nhập liệu.
"""
import datetime as dt
import re
from decimal import Decimal, ROUND_HALF_UP

import openpyxl

MONTHS = ["January", "February", "March", "April", "May", "June", "July",
          "August", "September", "October", "November", "December"]

ONES = ["zero", "one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten",
        "eleven", "twelve", "thirteen", "fourteen", "fifteen", "sixteen", "seventeen", "eighteen", "nineteen"]
TENS = ["", "", "twenty", "thirty", "forty", "fifty", "sixty", "seventy", "eighty", "ninety"]


class MasterError(Exception):
    pass


# ----------------------------------------------------------------- số thành chữ
def _below_1000(n, british):
    parts = []
    if n >= 100:
        parts.append(ONES[n // 100] + " hundred")
        n %= 100
        if n and british:
            parts.append("and")
    if n >= 20:
        t = TENS[n // 10]
        parts.append(t + ("-" + ONES[n % 10] if n % 10 else ""))
    elif n > 0 or not parts:
        if n > 0:
            parts.append(ONES[n])
    return " ".join(parts)


def words(n, british=False):
    """Số nguyên không âm -> chữ tiếng Anh thường. british=True thêm 'and' sau 'hundred'."""
    n = int(n)
    if n == 0:
        return "zero"
    out = []
    for size, name in ((10 ** 9, "billion"), (10 ** 6, "million"), (1000, "thousand")):
        if n >= size:
            out.append(_below_1000(n // size, british) + " " + name)
            n %= size
    if n:
        if out and british and n < 100:
            out.append("and")
        out.append(_below_1000(n, british))
    return " ".join(out)


def money_words(amount):
    """Trả về chữ cho phần đô-la và (nếu có) cent: 'twenty-nine thousand fifty' hoặc '... and fifty cents'."""
    amt = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    dollars = int(amt)
    cents = int((amt - dollars) * 100)
    text = words(dollars)
    if cents:
        text += " and " + words(cents) + " cents"
    return text


def title_words(text):
    return re.sub(r"[A-Za-z]+", lambda m: m.group(0).capitalize(), text)


# ----------------------------------------------------------------- định dạng
def fmt_int_or_2(x):
    x = float(x)
    return f"{x:,.0f}" if abs(x - round(x)) < 1e-9 else f"{x:,.2f}"


def fmt2(x):
    return f"{float(x):,.2f}"


def fmt_price(x):
    s = f"{float(x):.4f}".rstrip("0")
    if s.endswith("."):
        s += "0"
    if len(s.split(".")[1]) < 2:
        s += "0" * (2 - len(s.split(".")[1]))
    return s


def r2(x):
    return float(Decimal(str(x)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


def d_dmy(d):
    return f"{d.day:02d}{d.month:02d}{d.year}"


def d_text_a(d):  # November 09, 2026
    return f"{MONTHS[d.month - 1]} {d.day:02d}, {d.year}"


def d_text_b(d):  # 09 November, 2026
    return f"{d.day:02d} {MONTHS[d.month - 1]}, {d.year}"


def d_text_upper(d):  # 18 NOVEMBER, 2026
    return f"{d.day:02d} {MONTHS[d.month - 1].upper()}, {d.year}"


def d_slash(d):  # 29/10/2026
    return f"{d.day:02d}/{d.month:02d}/{d.year}"


def month_upper(d):  # NOVEMBER 2026
    return f"{MONTHS[d.month - 1].upper()} {d.year}"


def add_months(d, n):
    y = d.year + (d.month - 1 + n) // 12
    m = (d.month - 1 + n) % 12 + 1
    return dt.date(y, m, 1)


def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    raise MasterError(f"Giá trị ngày không hợp lệ: {v!r} (cần nhập kiểu ngày trong Excel)")


def s(v):
    return "" if v is None else str(v).strip()


# ----------------------------------------------------------------- đọc master
def _table(ws):
    heads = [(s(c.value).split(" ")[0] if c.value else None) for c in ws[1]]
    rows = []
    for row in ws.iter_rows(min_row=2):
        rec = {h: c.value for h, c in zip(heads, row) if h}
        if any(v not in (None, "") for k, v in rec.items() if not k.startswith(("next_", "production_month", "expiry_month"))
               and k not in ("net_kg_per_ctn", "gross_kg_per_ctn", "net_kg", "gross_kg", "buy_amount_usd",
                             "sell_amount_usd", "margin_usd", "sell_price_usd_mt")):
            rows.append(rec)
    return rows


def load_master(path):
    wb = openpyxl.load_workbook(path, data_only=False)
    need = ["Parties", "Products", "Shipments", "Line_Items", "Execution"]
    for n in need:
        if n not in wb.sheetnames:
            raise MasterError(f"Thiếu sheet {n} trong {path}")
    return {n: _table(wb[n]) for n in need}


def _party(parties, code, what):
    for p in parties:
        if s(p.get("party_code")) == code:
            return p
    raise MasterError(f"{what}: không tìm thấy mã đối tác '{code}' trong sheet Parties")


def _party_ctx(p):
    reg_label = s(p.get("reg_label"))
    reg_no = s(p.get("reg_no"))
    legal = s(p.get("legal_name"))
    short_label = "UEN" if reg_label.lower().startswith("unique") else reg_label
    rep = s(p.get("rep_name"))
    title = s(p.get("rep_title"))
    return {
        "code": s(p.get("party_code")),
        "legal_name": legal,
        "legal_name_nodot": legal.rstrip("."),
        "address": s(p.get("address")),
        "address_upper": s(p.get("address")).upper(),
        "reg_label": reg_label,
        "reg_no": reg_no,
        "reg_short": f"{short_label} {reg_no}".strip() if reg_no else "",
        "reg_line": (f"\n{reg_label}: {reg_no}" if reg_no else ""),
        "vendor_code": s(p.get("vendor_code")),
        "rep_name": rep,
        "rep_title": title,
        "rep_name_title": (rep.title() + (" - " + title if title else "")).strip(),
        "email": s(p.get("email")),
        "tel": s(p.get("tel")),
        "sign_name": s(p.get("sign_name")) or legal,
        "bank_beneficiary": s(p.get("bank_beneficiary")),
        "bank_beneficiary_nodot": s(p.get("bank_beneficiary")).rstrip("."),
        "bank_account": s(p.get("bank_account_usd")),
        "swift": s(p.get("swift")),
        "bank_name": s(p.get("bank_name")),
        "bank_name_upper": s(p.get("bank_name")).upper(),
        "bank_address": s(p.get("bank_address")),
        "bank_address_upper": s(p.get("bank_address")).upper(),
    }


def _short_place(text):
    t = text.title().replace("United Arab Emirates", "UAE")
    t = re.sub(r"Ho Chi Minh,", "Ho Chi Minh City,", t)
    return t


def number_words_days(n):
    return f"{words(n)} ({n})"


def build_context(master_path, op_code):
    m = load_master(master_path)
    ship = next((r for r in m["Shipments"] if s(r.get("op_code")) == op_code), None)
    if ship is None:
        raise MasterError(f"Không thấy lô {op_code} trong sheet Shipments")
    exe = next((r for r in m["Execution"] if s(r.get("op_code")) == op_code), None)
    lines = sorted((r for r in m["Line_Items"] if s(r.get("op_code")) == op_code),
                   key=lambda r: (r.get("line_no") or 0))
    if not lines:
        raise MasterError(f"{op_code}: chưa có dòng hàng nào trong sheet Line_Items")

    sup = _party_ctx(_party(m["Parties"], s(ship["supplier_code"]), f"{op_code} NCC"))
    buy = _party_ctx(_party(m["Parties"], s(ship["buyer_code"]), f"{op_code} Buyer"))
    fgs = _party_ctx(_party(m["Parties"], "FGS", "Farmgate"))

    prod_by_code = {s(p["product_code"]): p for p in m["Products"]}
    items = []
    for i, ln in enumerate(lines, start=1):
        code = s(ln["product_code"])
        p = prod_by_code.get(code)
        if p is None:
            raise MasterError(f"{op_code} dòng {i}: sản phẩm '{code}' chưa có trong sheet Products")
        for fld in ("cartons", "buy_price_usd_kg", "sell_price_usd_kg"):
            if ln.get(fld) in (None, ""):
                raise MasterError(f"{op_code} dòng {i} ({code}): thiếu {fld}")
        cartons = int(ln["cartons"])
        net_pc = float(p["net_kg_per_ctn"])
        gross_pc = float(p["gross_kg_per_ctn"])
        net_kg = cartons * net_pc
        gross_kg = cartons * gross_pc
        bp = float(ln["buy_price_usd_kg"])
        sp = float(ln["sell_price_usd_kg"])
        buy_amt = r2(net_kg * bp)
        sell_amt = r2(net_kg * sp)
        items.append({
            "no": f"{i:02d}", "no1": str(i), "code": code,
            "sup_name": s(p["sup_name"]), "sup_pack": s(p["sup_pack"]),
            "sup_pack_x": s(p["sup_pack"]).replace("×", "x"),
            "sup_desc": f'{s(p["sup_name"])} ({s(p["sup_pack"])})',
            "buy_desc": f'{s(p["buy_name"])} ({s(p["buy_pack"])})',
            "sc_name": s(p["sc_name"]), "sc_pack": s(p["sc_pack"]),
            "sc_desc": f'{s(p["sc_name"])} ({s(p["sc_pack"])})',
            "hs8": s(p["hs_code_8"]), "hs12": s(p["hs_code_12"]),
            "cartons": cartons, "net_per_ctn": net_pc, "gross_per_ctn": gross_pc,
            "net_kg": net_kg, "gross_kg": gross_kg,
            "buy_price": bp, "sell_price": sp,
            "buy_amount": buy_amt, "sell_amount": sell_amt,
            "cartons_fmt": f"{cartons:,}", "qty_fmt": f"{net_kg:,.0f}",
            "buy_price_fmt": fmt_price(bp), "buy_amount_fmt": fmt_int_or_2(buy_amt),
            "sell_price_fmt": fmt_price(sp), "sell_amount_fmt": fmt_int_or_2(sell_amt),
            "sell_price_mt_fmt": fmt2(sp * 1000), "mt_fmt": fmt2(net_kg / 1000),
            "sell_amount_fmt2": fmt2(sell_amt),
            "sc_line": None,
        })

    tot_cartons = sum(i["cartons"] for i in items)
    tot_net = sum(i["net_kg"] for i in items)
    tot_gross = sum(i["gross_kg"] for i in items)
    tot_buy = r2(sum(i["buy_amount"] for i in items))
    tot_sell = r2(sum(i["sell_amount"] for i in items))
    hs8 = ", ".join(dict.fromkeys(i["hs8"] for i in items))
    hs12 = ", ".join(dict.fromkeys(i["hs12"] for i in items))
    tot = {
        "cartons": tot_cartons, "net_kg": tot_net, "gross_kg": tot_gross,
        "buy": tot_buy, "sell": tot_sell,
        "cartons_fmt": f"{tot_cartons:,}", "qty_fmt": f"{tot_net:,.0f}",
        "buy_fmt": fmt_int_or_2(tot_buy), "sell_fmt": fmt_int_or_2(tot_sell), "sell_fmt2": fmt2(tot_sell),
        "mt_fmt": fmt2(tot_net / 1000),
        "hs8": hs8, "hs12": hs12,
        "cartons_words": words(tot_cartons, british=True).upper(),
        "buy_words_upper": "US DOLLARS " + money_words(tot_buy).upper() + " ONLY.",
        "sell_words_upper": "US DOLLARS " + money_words(tot_sell).upper() + " ONLY.",
        "buy_words_sentence": money_words(tot_buy).capitalize() + " US Dollars only.",
        "sell_words_title": "Say US Dollars " + title_words(money_words(tot_sell)) + " only.",
    }

    # ---- hợp đồng
    pc_date = to_date(ship.get("purchase_contract_date"))
    sc_date = to_date(ship.get("sales_contract_date"))
    if ship.get("purchase_contract_seq") in (None, "") or pc_date is None:
        raise MasterError(f"{op_code}: thiếu ngày/số thứ tự hợp đồng mua")
    if ship.get("sales_contract_seq") in (None, "") or sc_date is None:
        raise MasterError(f"{op_code}: thiếu ngày/số thứ tự hợp đồng bán")
    pc_no = f'FGSG/{sup["code"]}/{d_dmy(pc_date)}/{int(ship["purchase_contract_seq"]):02d}'
    sc_no = f'FGSG/{buy["code"]}/{d_dmy(sc_date)}/{int(ship["sales_contract_seq"]):02d}'

    pc_inc, pc_place = s(ship["purchase_incoterm"]), s(ship["purchase_incoterm_place"])
    sc_inc, sc_place = s(ship["sales_incoterm"]), s(ship["sales_incoterm_place"])

    # ---- thanh toán
    pay = {}
    ptype = s(ship.get("purchase_pay_type"))
    pct = ship.get("purchase_deposit_pct")
    if ptype == "DEPOSIT_BALANCE_ON_DOCS":
        if pct in (None, ""):
            raise MasterError(f"{op_code}: thiếu purchase_deposit_pct")
        dep = r2(tot_buy * float(pct))
        bal = r2(tot_buy - dep)
        p1 = round(float(pct) * 100)
        pay.update({
            "dep_pct": p1, "bal_pct": 100 - p1, "dep_fmt": fmt_int_or_2(dep), "bal_fmt": fmt_int_or_2(bal),
            "ci_term_ncc": (f"Term payment: Farmgate pays a {p1}% deposit (USD {fmt_int_or_2(dep)}) upon signing the "
                            f"contract, and the remaining {100 - p1}% (USD {fmt_int_or_2(bal)}) by T/T upon receipt of "
                            f"the scanned shipping documents from the Seller, for a total of USD {fmt_int_or_2(tot_buy)}, "
                            f"{pc_inc} {pc_place.split(',')[0]}, Incoterms 2020."),
        })
    else:
        pay["unsupported_purchase"] = ptype or "(trống)"
    stype = s(ship.get("sales_pay_type"))
    if stype == "DAYS_FROM_BL":
        days = int(ship.get("sales_pay_days") or 0)
        if not days:
            raise MasterError(f"{op_code}: thiếu sales_pay_days")
        pay["days_words"] = number_words_days(days)
        pay["ci_term_buy"] = (f"Term payment: Payment shall be made within {number_words_days(days)} calendar days "
                              f"from the Bill of Lading (On Board) date by telegraphic transfer (T/T)")
    elif stype == "DEPOSIT_BALANCE_ON_ARRIVAL":
        sp_pct = ship.get("sales_deposit_pct")
        if sp_pct in (None, ""):
            raise MasterError(f"{op_code}: thiếu sales_deposit_pct")
        p1 = round(float(sp_pct) * 100)
        dep = r2(tot_sell * float(sp_pct))
        pay["ci_term_buy"] = (f"Term payment: The Buyer pays a {p1}% deposit (USD {fmt_int_or_2(dep)}). The remaining "
                              f"balance (USD {fmt_int_or_2(tot_sell - dep)}) shall be paid by T/T on arrival at Port of "
                              f"Destination.")
        pay["unsupported_sales_contract"] = stype
    else:
        pay["unsupported_sales"] = stype or "(trống)"

    # ---- mô tả hàng cho hợp đồng bán
    sc_lines = []
    for n, it in enumerate(items, start=1):
        pack_short = it["sc_pack"].replace(" packing", "s")  # '2 x 10 kg tin packing' -> '2 x 10 kg tins'
        sc_lines.append(f'({n}) {it["sc_name"]}, packed in {pack_short}; {it["cartons"]} cartons; '
                        f'net weight {it["net_kg"]:,.2f} kgs. Unit price: USD {fmt2(it["sell_price"] * 1000)}/MT '
                        f'{sc_inc} {s(ship.get("sales_place_of_delivery"))}.')
    if len(items) == 1:
        sc_lines[0] = sc_lines[0][4:]
    else:
        sc_lines.append(f'Total: {tot_cartons} cartons; net weight {tot_net:,.2f} kgs.')
    sc_ctx = {
        "no": sc_no, "date": sc_date, "date_slash": d_slash(sc_date),
        "incoterm": sc_inc, "place": sc_place, "place_upper": sc_place.upper(),
        "place_of_delivery": s(ship.get("sales_place_of_delivery")),
        "loading_port": s(ship.get("sales_loading_port")),
        "delivery_period": s(ship.get("sales_delivery_period")),
        "valid_until": s(ship.get("sales_valid_until")),
        "certificates": s(ship.get("sales_certificates")),
        "preservation": s(ship.get("sales_preservation")),
        "cert_part": (f'; Certificates: {s(ship.get("sales_certificates"))}' if s(ship.get("sales_certificates")) else ""),
        "pres_part": (f'; Preservation: {s(ship.get("sales_preservation"))}' if s(ship.get("sales_preservation")) else ""),
        "other": s(ship.get("sales_other")),
        "description": "\n".join(sc_lines),
        "spec": ", ".join(it["code"] for it in items),
        "packaging": "; ".join(f'{it["sc_name"]} – {it["sc_pack"]} ({it["cartons"]} cartons)' for it in items)
        if len(items) > 1 else f'{items[0]["sc_pack"]} ({items[0]["cartons"]} cartons)',
    }
    pc_ctx = {
        "no": pc_no, "date": pc_date, "date_b": d_text_b(pc_date),
        "incoterm": pc_inc, "place": pc_place, "place_upper": pc_place.upper(),
        "etd_text": s(ship.get("purchase_etd_text")), "pod": s(ship.get("purchase_pod")),
        "surveyor": s(ship.get("purchase_surveyor_cert")).upper() == "YES",
        "surveyor_line": ("+ Quality and Weight Certificate issued by an independent surveyor: 1 Original + 2 Copies"
                          if s(ship.get("purchase_surveyor_cert")).upper() == "YES" else ""),
    }

    # ---- vận chuyển (có thể chưa có nếu mới ở giai đoạn hợp đồng)
    stuffing = to_date(exe.get("stuffing_date")) if exe else None
    etd = to_date(exe.get("etd")) if exe else None
    ctx = {
        "op": op_code, "sup": sup, "buy": buy, "fgs": fgs, "items": items, "tot": tot, "pay": pay,
        "pc": pc_ctx, "sc": sc_ctx, "has_exe": bool(exe and stuffing),
        "pod": {
            "fgs_rep": s(ship.get("pod_farmgate_rep")) or fgs["rep_name"],
            "fgs_title": s(ship.get("pod_farmgate_title")),
            "fgs_pos_line": (f'Position: {s(ship.get("pod_farmgate_title"))}\n' if s(ship.get("pod_farmgate_title")) else ""),
        },
        "margin": r2(tot_sell - tot_buy),
    }
    if exe and stuffing:
        prod_override = to_date(exe.get("production_month_override"))
        prod = (prod_override or stuffing).replace(day=1)
        exp = add_months(prod, 24)
        container, seal = s(exe.get("container_no")), s(exe.get("seal_no"))
        ctx["doc"] = {
            "date_a": d_text_a(stuffing), "date_b": d_text_b(stuffing), "dmy": d_dmy(stuffing),
            "pl_ncc": f'FGSG-{sup["code"]}/PL/{d_dmy(stuffing)}',
            "ci_ncc": f'FGSG-{sup["code"]}/CI/{d_dmy(stuffing)}',
            "si_ncc": f'FGSG-{sup["code"]}/SI/{d_dmy(stuffing)}',
            "pod_no": f'FGSG-{sup["code"]}/{d_dmy(stuffing)}',
            "pl_buy": f'FGSG-{buy["code"]}/PL/{d_dmy(stuffing)}',
            "ci_buy": f'FGSG-{buy["code"]}/CI/{d_dmy(stuffing)}',
        }
        ctx["ship"] = {
            "stuffing": stuffing, "vessel": s(exe.get("vessel_voyage")),
            "pol": s(exe.get("port_of_loading")), "place": s(exe.get("place_of_delivery")),
            "pod": s(exe.get("port_of_discharge")),
            "pol_pod": _short_place(s(exe.get("port_of_loading"))),
            "pod_short": _short_place(s(exe.get("port_of_discharge"))),
            "etd_upper": d_text_upper(etd) if etd else "",
            "container": container, "seal": seal,
            "container_blank": container or "__________",
            "cont_block": (f"CONTAINER NO.: {container or '__________'}\nSEAL NO.: {seal or '__________'}"
                           if (container or seal) else "CONTAINER NO./ SEAL NO.: __________"),
            "prod_upper": month_upper(prod), "exp_upper": month_upper(exp),
            "booking_no": s(exe.get("booking_no")),
        }
    return ctx
