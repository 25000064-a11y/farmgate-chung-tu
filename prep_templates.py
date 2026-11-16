"""Chạy 1 lần: biến 5 template gốc của bạn thành template có placeholder (thư mục templates/).

Template gốc giữ nguyên trong templates_original/. Nếu sau này bạn đổi bố cục template, chỉ cần sửa
file trong templates/ (giữ nguyên các placeholder {{ ... }}), không cần sửa code.
"""
import shutil
import sys
from copy import deepcopy
from pathlib import Path

import docx
import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

HERE = Path(__file__).parent
ORIG = HERE / "templates_original"
OUT = HERE / "templates"
OUT.mkdir(exist_ok=True)

SRC = {
    "ncc_xlsx": "PL_CI_SI_NCC.xlsx",
    "buyer_xlsx": "PL_CI_Buyer.xlsx",
    "pod": "Proof_of_delivery_NCC.docx",
    "purchase": "Hop_dong_mua_NCC.docx",
    "sales": "Hop_dong_ban_Buyer.docx",
}

# ===================================================================== XLSX
BANK_NCC = ("Payments shall be made via bank transfer to the following account details:\n"
            "·         Beneficiary name:  {{ sup.bank_beneficiary }}\n"
            "·         Account number (USD): {{ sup.bank_account }} -  Swift code: {{ sup.swift }}\n"
            "·         Bank name: {{ sup.bank_name }}.\n"
            "·         Bank address: {{ sup.bank_address }}")
BANK_FGS = ("Payments shall be made via bank transfer to the following account details:\n"
            "·         Beneficiary name:  {{ fgs.bank_beneficiary }}\n"
            "·         Account number (USD): {{ fgs.bank_account }} -  Swift code: {{ fgs.swift }}\n"
            "·         Bank name: {{ fgs.bank_name }}\n"
            "·         Bank address: {{ fgs.bank_address }}")


def fonts_of(cell):
    """Lấy font của các đoạn trong ô rich-text: trả về (font_đậm, font_thường)."""
    bold = reg = None
    for b in cell.value:
        if isinstance(b, TextBlock):
            if b.font.b and bold is None:
                bold = b.font
            if not b.font.b and reg is None:
                reg = b.font
    return bold, reg


def rich(*parts):
    """parts: (font, text) hoặc text thường."""
    return CellRichText(*[TextBlock(f, t) if f is not None else t for f, t in parts])


def clone_style(ws, src, dst):
    ws[dst]._style = deepcopy(ws[src]._style)


def prep_ncc():
    wb = openpyxl.load_workbook(ORIG / SRC["ncc_xlsx"], rich_text=True)
    pkl, inv, si = wb["PKL"], wb["INV"], wb["SI"]
    seller = ("SELLER:    {{ sup.legal_name }}\nAddress: {{ sup.address }}.\n"
              "{{ sup.reg_label }}: {{ sup.reg_no }}")
    buyer = ("BUYER:      {{ fgs.legal_name }}\nVENDOR CODE: {{ fgs.vendor_code }}\n"
             "Address: {{ fgs.address }}. \n{{ fgs.reg_label }}: {{ fgs.reg_no }}.")
    fb, fr = fonts_of(pkl["A5"])
    contract_rich = rich((fb, "CONTRACT NO: "), (fr, "{{ pc.no }}"), (fb, "                "), (fr, "\n"),
                         (fb, "{{ ship.cont_block }}"))
    contract = "CONTRACT NO: {{ pc.no }}                \n{{ ship.cont_block }}"
    bb, br = fonts_of(inv["A15"])
    bank_ncc_rich = rich((None, "Payments shall be made via bank transfer to the following account details:\n"
                                "·         Beneficiary name:  "),
                         (bb, "{{ sup.bank_beneficiary }}\n"), (br, "·         Account number (USD): "),
                         (bb, "{{ sup.bank_account }} "), (br, "-  Swift code: "), (bb, "{{ sup.swift }}\n"),
                         (br, "·         Bank name: "), (bb, "{{ sup.bank_name }}.\n"),
                         (br, "·         Bank address: {{ sup.bank_address }}"))
    place_lbl = "PORT OF LOADING: \nPLACE OF DELIVERY: \nDEPARTURE DATE:"
    place_val = "{{ ship.pol }}\n{{ ship.place }}\n{{ ship.etd_upper }}"
    desc = "{{ it.sup_desc }}\nHS CODE: {{ it.hs8 }}"
    origin = ("ORIGIN: VIETNAM\nPRODUCTION DATE : {{ ship.prod_upper }}\nEXPIRY DATE : {{ ship.exp_upper }}\n"
              "HS CODE: {{ tot.hs8 }}")
    ref = "Contract No.: {{ pc.no }}"

    # ---- PKL
    pkl["A3"] = seller
    pkl["F3"] = "{{ doc.pl_ncc }}\nDate: {{ doc.date_a }}\n" + ref
    pkl["A4"] = buyer
    pkl["A5"] = contract_rich
    pkl["A6"] = place_lbl
    pkl["C6"] = place_val
    pkl["A8"] = desc
    pkl["D8"] = "{{ it.cartons }}"
    pkl["E8"] = "{{ it.net_per_ctn }}"
    pkl["G8"] = "=E{r}*D{r}"
    pkl["H8"] = "=D{r}*K{r}"
    pkl["K7"] = "GROSS WEIGHT\n(KGS/CTN)"
    clone_style(pkl, "H7", "K7")
    pkl["K8"] = "{{ it.gross_per_ctn }}"
    clone_style(pkl, "E8", "K8")
    pkl.column_dimensions["K"].width = 14
    pkl["A9"] = origin
    for col in "DGH":
        pkl[f"{col}10"] = f"=SUM({col}{{first}}:{col}{{last}})"
    pkl["E12"] = "{{ sup.sign_name }}"

    # ---- INV
    inv["A4"] = seller
    inv["F4"] = "{{ doc.ci_ncc }}\nDate: {{ doc.date_a }}\n" + ref
    inv["A5"] = buyer
    inv["A6"] = contract
    inv["A7"] = place_lbl
    inv["B7"] = place_val
    inv["A9"] = desc
    inv["C9"] = "{{ it.cartons }}"
    inv["D9"] = "{{ it.net_kg }}"
    inv["F9"] = "{{ it.buy_price }}"
    inv["G9"] = "=F{r}*D{r}"
    inv["A10"] = origin
    inv["A11"] = "TOTAL:  {{ pc.incoterm }} {{ pc.place_upper }}"
    inv["G11"] = "=SUM(G{first}:G{last})"
    inv["A12"] = "{{ pay.ci_term_ncc }}"
    inv["C13"] = "=SUM(C{first}:C{last})"
    inv["D13"] = "=SUM(D{first}:D{last})"
    inv["A14"] = "Say US Dollars:  {{ tot.buy_words_upper }}"
    inv["A15"] = bank_ncc_rich
    inv["D16"] = "{{ sup.sign_name }}"

    # ---- SI
    si["A3"] = "SHIPPER:   {{ sup.legal_name }}\nAddress: {{ sup.address }}.\n{{ sup.reg_label }}: {{ sup.reg_no }}"
    si["F3"] = "{{ doc.si_ncc }}\nDate: {{ doc.date_a }}\n" + ref
    si["A4"] = ("CONSIGNEE:  {{ fgs.legal_name }}\nVENDOR CODE: {{ fgs.vendor_code }}\n"
                "Address: {{ fgs.address }}. \n{{ fgs.reg_label }}: {{ fgs.reg_no }}.")
    si["A5"] = contract_rich
    si["A6"] = place_lbl
    si["C6"] = place_val
    si["A8"] = desc
    si["D8"] = "{{ it.cartons }}"
    si["E8"] = "{{ it.net_per_ctn }}"
    si["G8"] = "=E{r}*D{r}"
    si["H8"] = "{{ it.gross_kg }}"
    si["A9"] = origin
    for col in "DGH":
        si[f"{col}10"] = f"=SUM({col}{{first}}:{col}{{last}})"
    si["E12"] = "{{ sup.sign_name }}"
    wb.save(OUT / "PL_CI_SI_NCC.xlsx")


def prep_buyer():
    wb = openpyxl.load_workbook(ORIG / SRC["buyer_xlsx"], rich_text=True)
    pkl, inv = wb["PKL"], wb["INV"]
    seller = ("SELLER:                                       {{ fgs.legal_name }}\nVENDOR CODE: {{ fgs.vendor_code }}\n"
              "Address: {{ fgs.address }}. {{ fgs.reg_label }}: {{ fgs.reg_no }}.")
    buyer = " {{ buy.legal_name }}\nAddress: {{ buy.address_upper }}{{ buy.reg_line }}"
    contract = "SALES CONTRACT NO: {{ sc.no }}                    \n{{ ship.cont_block }}"
    fb, fr = fonts_of(pkl["A5"])
    contract_rich = rich((fb, "SALES CONTRACT NO: {{ sc.no }}                    "), (fr, "\n"),
                         (fb, "{{ ship.cont_block }}"))
    bb, br = fonts_of(inv["A15"])
    bank_fgs_rich = rich((br, "Payments shall be made via bank transfer to the following account details:\n"
                              "·         Beneficiary name:  "),
                         (bb, "{{ fgs.bank_beneficiary }}\n"), (br, "·         Account number (USD): "),
                         (bb, "{{ fgs.bank_account }} "), (br, "-  Swift code: "), (bb, "{{ fgs.swift }}\n"),
                         (br, "·         Bank name: "), (bb, "{{ fgs.bank_name }}\n"),
                         (br, "·         Bank address: {{ fgs.bank_address }}"))
    title_font = [b.font for b in inv["B1"].value if isinstance(b, TextBlock) and b.font.sz == 22][0]
    small_font = [b.font for b in inv["B1"].value if isinstance(b, TextBlock) and b.font.sz == 10][0]
    place_lbl = "NAME OF VESSEL: \nPORT OF LOADING: \nPLACE OF DELIVERY: \nDEPARTURE DATE:"
    place_val = "{{ ship.vessel }}\n{{ ship.pol }}\n{{ ship.place }}\n{{ ship.etd_upper }}"
    desc = "{{ it.buy_desc }}"
    origin = ("ORIGIN: VIETNAM\nPRODUCTION DATE : {{ ship.prod_upper }}\nEXPIRY DATE : {{ ship.exp_upper }}\n"
              "HS CODE: {{ tot.hs12 }}")
    head = "Address: {{ fgs.address }}. {{ fgs.reg_label }}: {{ fgs.reg_no }}"
    ref = "Sales Contract No: {{ sc.no }}"

    pkl["A1"] = "Address: {{ fgs.address }}.\n{{ fgs.reg_label }}: {{ fgs.reg_no }}"
    pkl["A3"] = seller
    pkl["F3"] = "{{ doc.pl_buy }}\nDate: {{ doc.date_b }}\n" + ref
    pkl["A4"] = "BUYER:"
    pkl["B4"] = buyer
    pkl["A5"] = contract_rich
    pkl["A6"] = place_lbl
    pkl["C6"] = place_val
    pkl["A8"] = desc
    pkl["D8"] = "{{ it.cartons }}"
    pkl["E8"] = "{{ it.net_per_ctn }}"
    pkl["G8"] = "=E{r}*D{r}"
    pkl["H8"] = "=D{r}*K{r}"
    pkl["K7"] = "GROSS WEIGHT\n(KGS/CTN)"
    clone_style(pkl, "H7", "K7")
    pkl["K8"] = "{{ it.gross_per_ctn }}"
    clone_style(pkl, "E8", "K8")
    pkl.column_dimensions["K"].width = 14
    pkl["A9"] = origin
    for col in "DGH":
        pkl[f"{col}10"] = f"=SUM({col}{{first}}:{col}{{last}})"
    pkl["A11"] = "SAY IN WORD: {{ tot.cartons_words }} CARTONS ONLY./"
    pkl["A12"] = "COUNTRY OF ORIGIN: VIETNAM"
    pkl["F13"] = "{{ fgs.sign_name }}"

    inv["B1"] = rich((small_font, head + "\n"), (title_font, "COMMERCIAL INVOICE"))
    inv["A4"] = seller
    inv["E4"] = "{{ doc.ci_buy }}\nDate: {{ doc.date_b }}\n" + ref
    inv["B5"] = buyer
    inv["A6"] = contract
    inv["A7"] = place_lbl
    inv["B7"] = place_val
    inv["F8"] = "UNIT PRICE {{ sc.incoterm }} (USD/KGS)"
    inv["A9"] = desc
    inv["C9"] = "{{ it.cartons }}"
    inv["D9"] = "{{ it.net_kg }}"
    inv["F9"] = "{{ it.sell_price }}"
    inv["G9"] = "=F{r}*D{r}"
    inv["A10"] = origin
    inv["A11"] = "TOTAL:  {{ sc.incoterm }} {{ sc.place_upper }}"
    inv["G11"] = "=SUM(G{first}:G{last})"
    inv["A12"] = "{{ pay.ci_term_buy }}"
    inv["G12"] = 0
    inv["C13"] = "=SUM(C{first}:C{last})"
    inv["D13"] = "=SUM(D{first}:D{last})"
    inv["A14"] = "Say US Dollars:  {{ tot.sell_words_upper }}"
    inv["A15"] = bank_fgs_rich
    inv["E16"] = "{{ fgs.sign_name }}"
    wb.save(OUT / "PL_CI_Buyer.xlsx")


# ===================================================================== DOCX
def all_paragraphs(doc):
    body = doc.element.body
    for p in body.iter(qn("w:p")):
        yield Paragraph(p, doc)


def replace_in_par(p, old, new):
    from docx.text.run import Run
    runs = [Run(r, p) for r in p._p.iter(qn("w:r"))]  # gồm cả run nằm trong hyperlink
    full = "".join(r.text for r in runs)
    i = full.find(old)
    if i < 0:
        return False
    j = i + len(old)
    pos = 0
    first = None
    for r in runs:
        t = r.text
        a, b = pos, pos + len(t)
        if b > i and a < j:
            rs, re_ = max(i, a) - a, min(j, b) - a
            if first is None:
                r.text = t[:rs] + new + t[re_:]
                first = r
            else:
                r.text = t[:rs] + t[re_:]
        pos = b
    return True


def rep(doc, old, new, count=1, scope=None):
    """Thay `old` bằng `new` trong đúng `count` đoạn văn (kiểm tra để không sót)."""
    paras = list(all_paragraphs(doc)) if scope is None else scope
    n = 0
    for p in paras:
        while replace_in_par(p, old, new):
            n += 1
            if old in new:
                break
    if n != count:
        raise SystemExit(f"[prep] '{old[:60]}' thay được {n} lần, mong đợi {count}")


def cell_paras(table, r, c):
    return list(table.cell(r, c).paragraphs)


def add_loop_rows(table, row_idx, seq="items", var="it"):
    tr = table.rows[row_idx]._tr

    def tagged(tag):
        el = deepcopy(tr)
        paras = list(el.iter(qn("w:p")))
        for p in paras:
            for child in list(p):
                if child.tag != qn("w:pPr"):
                    p.remove(child)
        run = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = tag
        run.append(t)
        paras[0].append(run)
        return el

    tr.addprevious(tagged(f"{{%tr for {var} in {seq} %}}"))
    tr.addnext(tagged("{%tr endfor %}"))


def wrap_paragraph_if(doc, startswith, cond):
    """Bọc đoạn văn trong {%p if cond %} ... {%p endif %} để bỏ hẳn đoạn khi không có dữ liệu."""
    target = [p for p in all_paragraphs(doc) if p.text.strip().startswith(startswith)]
    if len(target) != 1:
        raise SystemExit(f"[prep] không tìm đúng 1 đoạn bắt đầu bằng {startswith!r}")
    el = target[0]._p

    def tagged(tag):
        t = deepcopy(el)
        for child in list(t):
            if child.tag != qn("w:pPr"):
                t.remove(child)
        run = OxmlElement("w:r")
        tx = OxmlElement("w:t")
        tx.text = tag
        run.append(tx)
        t.append(run)
        return t
    el.addprevious(tagged(f"{{%p if {cond} %}}"))
    el.addnext(tagged("{%p endif %}"))


def prep_pod():
    d = docx.Document(ORIG / SRC["pod"])
    rep(d, "FGSG-HP/07102026", "{{ doc.pod_no }}")
    rep(d, "Principal Contract FGSG/HP/25092026/03 signed on 25 September, 2026",
        "Principal Contract {{ pc.no }} signed on {{ pc.date_b }}")
    rep(d, "Delivery Date: 07 October, 2026", "Delivery Date: {{ doc.date_b }}")
    # Party A
    rep(d, "FARMGATE PTE. LTD", "{{ fgs.legal_name_nodot }}", count=2)
    rep(d, "68 Circular Road, #02-01 Singapore 049422", "{{ fgs.address }}")
    rep(d, "Unique Entity Number: 202528980G", "{{ fgs.reg_label }}: {{ fgs.reg_no }}")
    rep(d, "Tel: 0766311868", "Tel: {{ fgs.tel }}")
    rep(d, "Representative: HOANG THI MINH CHAU", "Representative: {{ pod.fgs_rep }}")
    rep(d, "Position: Head of Global Finance\nEmail: exportsg@techcoop.vn",
        "{{ pod.fgs_pos_line }}Email: {{ fgs.email }}")
    rep(d, "HOANG THI MINH CHAU", "{{ pod.fgs_rep }}")
    rep(d, "Head of Global Finance", "{{ pod.fgs_title }}")
    # Party B
    rep(d, "HAU PHAT PRODUCTION TRADING COMPANY LIMITED", "{{ sup.legal_name }}", count=2)
    rep(d, "Hoang Hoa Tham Street, Group 3, Binh Giang 1 Quarter, Phuoc Long Ward, Dong Nai City, Vietnam",
        "{{ sup.address }}")
    rep(d, "Business Registration Code: 3801283877", "Business Registration Code: {{ sup.reg_no }}")
    rep(d, "Representative: LE THI HIEN", "Representative: {{ sup.rep_name }}")
    rep(d, "Position: Director", "Position: {{ sup.rep_title }}")
    rep(d, "lethihienbgm@gmail.com", "{{ sup.email }}")
    rep(d, "LE THI HIEN", "{{ sup.rep_name }}")
    rep(d, "Director", "{{ sup.rep_title }}", count=1)
    # logistics
    t = d.tables[2]
    rep(d, "MRSU0438483", "{{ ship.container_blank }}")
    rep(d, "07 October, 2026", "{{ doc.date_b }}")
    rep(d, "FOB Ho Chi Minh City, Vietnam – Incoterms® 2020", "{{ pc.incoterm }} {{ pc.place }} – Incoterms® 2020")
    rep(d, "Ho Chi Minh City, Vietnam", "{{ pc.place }}", scope=cell_paras(t, 3, 1))
    rep(d, "Ho Chi Minh City, Vietnam", "{{ ship.pol_pod }}", scope=cell_paras(t, 4, 1))
    rep(d, "Jebel Ali, UAE", "{{ ship.pod_short }}")
    # goods
    g = d.tables[3]
    rep(d, "01", "{{ it.no }}", scope=cell_paras(g, 1, 0))
    rep(d, "Cashew Kernels LP2", "{{ it.sup_name }}", scope=cell_paras(g, 1, 1))
    rep(d, "(2 x 10 KGs Vacuum Tin/Carton)", "({{ it.sup_pack_x }})", scope=cell_paras(g, 1, 1))
    rep(d, "2 x 10 KGs Vacuum Tin/Carton", "{{ it.sup_pack_x }}", scope=cell_paras(g, 1, 2))
    rep(d, "15,000", "{{ it.qty_fmt }}", scope=cell_paras(g, 1, 3))
    rep(d, "750", "{{ it.cartons_fmt }}", scope=cell_paras(g, 1, 5))
    rep(d, "1.55", "{{ it.buy_price_fmt }}", scope=cell_paras(g, 1, 6))
    rep(d, "23,250", "{{ it.buy_amount_fmt }}", scope=cell_paras(g, 1, 7))
    add_loop_rows(g, 1)
    d.save(OUT / "PoD_NCC.docx")


def prep_purchase():
    d = docx.Document(ORIG / SRC["purchase"])
    rep(d, "FGSG/HP/09112026/07", "{{ pc.no }}")
    rep(d, "09 November, 2026", "{{ pc.date_b }}", count=1)
    rep(d, "HAU PHAT PRODUCTION TRADING COMPANY LIMITED", "{{ sup.legal_name }}",
        scope=[p for p in all_paragraphs(d) if p.text.startswith("THE SELLER")])
    rep(d, "Address: Hoang Hoa Tham Street, Group 3, Binh Giang 1 Quarter, Phuoc Long Ward, Dong Nai City, Viet Nam.",
        "Address: {{ sup.address }}.")
    rep(d, "Business registration code: 3801283877", "Business registration code: {{ sup.reg_no }}")
    rep(d, "lethihienbgm@gmail.com", "{{ sup.email }}")
    rep(d, "Position: Le Thi Hien - Director", "Position: {{ sup.rep_name_title }}")
    rep(d, "FARMGATE PTE. LTD.", "{{ fgs.legal_name }}", count=2)
    rep(d, "Address: 68 Circular Road, #02-01, Singapore 049422", "Address: {{ fgs.address }}")
    rep(d, "Unique Entity Number: 202528980G", "{{ fgs.reg_label }}: {{ fgs.reg_no }}")
    rep(d, "Represented by: Nguyen Ngoc Thanh", "Represented by: {{ fgs.rep_name_title }}")
    rep(d, "exportsg@techcoop.vn", "{{ fgs.email }}")
    # tổng giá trị
    paras = list(all_paragraphs(d))
    p_total = [p for p in paras if p.text.strip().startswith("Total contract value") or "contract value" in p.text]
    rep(d, "FOB", "{{ pc.incoterm }}", scope=p_total)
    rep(d, "Ho Chi Minh City, Vietnam", "{{ pc.place }}", scope=p_total)
    rep(d, "42,390", "{{ tot.buy_fmt }}", scope=p_total)
    rep(d, "Forty-two thousand three hundred ninety US Dollars only.", "{{ tot.buy_words_sentence }}", count=2)
    rep(d, "As soon as possible.", "{{ pc.etd_text }}")
    p_pol = [p for p in paras if p.text.strip().startswith("- POL")]
    rep(d, "Ho Chi Minh City, Vietnam.", "{{ pc.place }}.", scope=p_pol)
    rep(d, "Khor Fakkan, UAE", "{{ pc.pod }}")
    rep(d, "(Deposit 20%)", "(Deposit {{ pay.dep_pct }}%)")
    rep(d, "USD 8,478", "USD {{ pay.dep_fmt }}")
    rep(d, "(Balance 80%)", "(Balance {{ pay.bal_pct }}%)")
    rep(d, "USD 33,912", "USD {{ pay.bal_fmt }}")
    p_tot2 = [p for p in paras if p.text.strip().startswith("Total amount")]
    rep(d, "FOB", "{{ pc.incoterm }}", scope=p_tot2)
    rep(d, "42,390", "{{ tot.buy_fmt }}", scope=p_tot2)
    rep(d, "+ Quality and Weight Certificate issued by an independent surveyor: 1 Original + 2 Copies",
        "{{ pc.surveyor_line }}")
    # bảng hàng
    t0 = d.tables[0]
    rep(d, "1", "{{ it.no1 }}", scope=cell_paras(t0, 1, 0))
    rep(d, "Cashew Kernels LP (Grade 2)\n(2 × 10 KGs Tin with plastic inside/Carton)",
        "{{ it.sup_name }} ({{ it.sup_pack }})", scope=cell_paras(t0, 1, 1))
    rep(d, "27,000", "{{ it.qty_fmt }}", scope=cell_paras(t0, 1, 3))
    rep(d, "1,350", "{{ it.cartons_fmt }}", scope=cell_paras(t0, 1, 4))
    rep(d, "1.57", "{{ it.buy_price_fmt }}", scope=cell_paras(t0, 1, 5))
    rep(d, "42,390", "{{ it.buy_amount_fmt }}", scope=cell_paras(t0, 1, 6))
    rep(d, "27,000", "{{ tot.qty_fmt }}", scope=cell_paras(t0, 2, 3))
    rep(d, "1,350", "{{ tot.cartons_fmt }}", scope=cell_paras(t0, 2, 4))
    rep(d, "42,390", "{{ tot.buy_fmt }}", scope=cell_paras(t0, 2, 6))
    add_loop_rows(t0, 1)
    # ngân hàng
    t1 = d.tables[1]
    rep(d, "MILITARY COMMERCIAL JOINT STOCK BANK – PHUOC LONG BRANCH", "{{ sup.bank_name_upper }}", scope=cell_paras(t1, 0, 1))
    rep(d, "QUARTER 5, PHUOC BINH WARD, DONG NAI CITY, VIET NAM", "{{ sup.bank_address_upper }}", scope=cell_paras(t1, 1, 1))
    rep(d, "HAU PHAT PRODUCTION TRADING COMPANY LIMITED", "{{ sup.bank_beneficiary_nodot }}", scope=cell_paras(t1, 2, 1))
    rep(d, "755556677", "{{ sup.bank_account }}", scope=cell_paras(t1, 3, 1))
    rep(d, "MSCBVNVX", "{{ sup.swift }}", scope=cell_paras(t1, 4, 1))
    rep(d, "HAU PHAT PRODUCTION TRADING CO., LTD", "{{ sup.sign_name }}")
    d.save(OUT / "Hop_dong_mua_NCC.docx")


def prep_sales():
    d = docx.Document(ORIG / SRC["sales"])
    rep(d, "FGSG/JT/23102026/01", "{{ sc.no }}")
    rep(d, "23/10/2026", "{{ sc.date_slash }}")
    t0 = d.tables[0]
    rep(d, "FARMGATE PTE. LTD.", "{{ fgs.legal_name }}", scope=cell_paras(t0, 1, 1))
    rep(d, "JALEEL TRADERS LLC", "{{ buy.legal_name }}", scope=cell_paras(t0, 1, 2))
    rep(d, "UEN 202528980G", "{{ fgs.reg_short }}", scope=cell_paras(t0, 2, 1))
    rep(d, "TRN 100209727500003", "{{ buy.reg_short }}", scope=cell_paras(t0, 2, 2))
    rep(d, "68 Circular Road, #02-01, Singapore 049422", "{{ fgs.address }}", scope=cell_paras(t0, 3, 1))
    rep(d, "P.O. Box 3262, Dubai, U.A.E.", "{{ buy.address }}", scope=cell_paras(t0, 3, 2))
    t1 = d.tables[1]
    rep(d, "1", "{{ it.no1 }}", scope=cell_paras(t1, 1, 0))
    rep(d, "Cashew LP2 (2 x 10 kg tin packing)", "{{ it.sc_desc }}", scope=cell_paras(t1, 1, 1))
    rep(d, "Ton", "MT", scope=cell_paras(t1, 1, 2))
    rep(d, "2,180.00", "{{ it.sell_price_mt_fmt }}", scope=cell_paras(t1, 1, 3))
    rep(d, "15.00", "{{ it.mt_fmt }}", scope=cell_paras(t1, 1, 4))
    rep(d, "32,700.00", "{{ it.sell_amount_fmt2 }}", scope=cell_paras(t1, 1, 5))
    rep(d, "15.00", "{{ tot.mt_fmt }}", scope=cell_paras(t1, 2, 4))
    rep(d, "32,700.00", "{{ tot.sell_fmt2 }}", scope=cell_paras(t1, 2, 5))
    rep(d, "Say US Dollars Thirty-Two Thousand Seven Hundred only.", "{{ tot.sell_words_title }}")
    rep(d, "Description: Cashew LP2, packed in 2 x 10 kg tins; 750 cartons; net weight 15,000.00 kgs. "
           "Unit price: USD 2,180.00/MT CFR Jebel Ali Port, UAE.", "Description: {{ sc.description }}")
    rep(d, "Quality / Specification: LP2; Certificates: ",
        "Quality / Specification: {{ sc.spec }}{{ sc.cert_part }}")
    rep(d, "Packaging: 2 x 10 kg tin packing (750 cartons); Preservation: ",
        "Packaging: {{ sc.packaging }}{{ sc.pres_part }}")
    rep(d, "Other (if any): ", "Other (if any): {{ sc.other }}")
    wrap_paragraph_if(d, "Other (if any):", "sc.other")
    rep(d, "Incoterms® 2020 rule: CFR Jebel Ali Port, UAE.", "Incoterms® 2020 rule: {{ sc.incoterm }} {{ sc.place_of_delivery }}.")
    rep(d, "Place of delivery: Jebel Ali Port, UAE; Loading port: Cat Lai Port, Ho Chi Minh City, Vietnam; "
           "Discharge port: Jebel Ali Port, UAE.",
        "Place of delivery: {{ sc.place_of_delivery }}; Loading port: {{ sc.loading_port }}; "
        "Discharge port: {{ sc.place_of_delivery }}.")
    rep(d, "Delivery period: November 2026;", "Delivery period: {{ sc.delivery_period }};")
    rep(d, "~ USD 32,700.00", "~ USD {{ tot.sell_fmt2 }}")
    rep(d, "thirty (30)", "{{ pay.days_words }}",
        scope=[p for p in all_paragraphs(d) if "calendar days from the Bill of Lading" in p.text])
    rep(d, "Standard Chartered Bank (Singapore) Limited", "{{ fgs.bank_name }}")
    rep(d, "6 Battery Road, #08-01, Singapore 049909", "{{ fgs.bank_address }}")
    rep(d, "FARMGATE PTE. LTD.", "{{ fgs.bank_beneficiary }}", scope=[p for p in all_paragraphs(d)
                                                                     if "Account holder" in p.text])
    rep(d, "0100105580", "{{ fgs.bank_account }}")
    rep(d, "SCBLSG22XXX", "{{ fgs.swift }}")
    rep(d, "31 January 2027", "{{ sc.valid_until }}")
    add_loop_rows(t1, 1)
    d.save(OUT / "Hop_dong_ban_Buyer.docx")


if __name__ == "__main__":
    for step in (prep_ncc, prep_buyer, prep_pod, prep_purchase, prep_sales):
        step()
        print("ok", step.__name__)
