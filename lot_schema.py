"""Schema JSON của 1 lô hàng (OP). Dùng chung cho extract.py (AI) và import_json.py (nhập vào master)."""

PAY_TYPES = ["DEPOSIT_BALANCE_ON_DOCS", "DEPOSIT_BALANCE_ON_ARRIVAL", "DAYS_FROM_BL"]
INCOTERMS = ["FOB", "CFR", "CIF", "CIP", "EXW", "FCA", "DAP"]

_s = {"type": ["string", "null"]}
_n = {"type": ["number", "null"]}
_d = {"type": ["string", "null"], "description": "Ngày dạng YYYY-MM-DD"}

PARTY = {
    "type": ["object", "null"],
    "description": "Thông tin đối tác đọc từ hợp đồng (chỉ dùng để thêm đối tác MỚI vào sheet Parties)",
    "properties": {k: _s for k in (
        "legal_name", "address", "reg_label", "reg_no", "rep_name", "rep_title", "email", "tel",
        "bank_beneficiary", "bank_account_usd", "swift", "bank_name", "bank_address")},
}

SCHEMA = {
    "type": "object",
    "properties": {
        "op_code": {"type": ["string", "null"], "description": "Mã lô, ví dụ OP01943"},
        "supplier_code": {**_s, "description": "Mã đối tác NCC trong sheet Parties (HP...)"},
        "buyer_code": {**_s, "description": "Mã đối tác Buyer trong sheet Parties (JT...)"},
        "supplier": PARTY,
        "buyer": PARTY,
        "purchase_contract": {
            "type": ["object", "null"],
            "properties": {
                "seq": {"type": ["integer", "null"], "description": "Số thứ tự cuối của số hợp đồng, ví dụ FGSG/HP/02112026/05 -> 5. null nếu hợp đồng chưa có số"},
                "date": _d,
                "incoterm": {"type": ["string", "null"], "enum": INCOTERMS + [None]},
                "incoterm_place": {**_s, "description": "Ví dụ: Ho Chi Minh City, Vietnam"},
                "pay_type": {"type": ["string", "null"], "enum": PAY_TYPES + [None]},
                "deposit_pct": {**_n, "description": "Tỷ lệ cọc, 50% -> 0.5"},
                "etd_text": {**_s, "description": "Nguyên văn điều khoản ETD, ví dụ 'Within November 2026.'"},
                "pod": {**_s, "description": "Cảng dỡ hàng ghi trong hợp đồng mua"},
                "surveyor_cert": {"type": ["string", "null"], "enum": ["YES", "NO", None],
                                  "description": "Hợp đồng có dòng Surveyor certificate không"},
            },
        },
        "sales_contract": {
            "type": ["object", "null"],
            "properties": {
                "seq": {"type": ["integer", "null"], "description": "Số thứ tự cuối của số hợp đồng, ví dụ FGSG/HP/02112026/05 -> 5. null nếu hợp đồng chưa có số"},
                "date": _d,
                "incoterm": {"type": ["string", "null"], "enum": INCOTERMS + [None]},
                "incoterm_place": _s,
                "pay_type": {"type": ["string", "null"], "enum": PAY_TYPES + [None]},
                "deposit_pct": _n,
                "pay_days": {**_n, "description": "Số ngày thanh toán kể từ ngày B/L (DAYS_FROM_BL)"},
                "delivery_period": _s,
                "loading_port": _s,
                "place_of_delivery": _s,
                "valid_until": _s,
                "certificates": _s,
                "preservation": _s,
                "other": _s,
            },
        },
        "lines": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "product_code": {"type": "string", "description": "Mã sản phẩm trong sheet Products (LP2, WW320...)"},
                    "cartons": {"type": ["integer", "null"]},
                    "buy_price_usd_kg": _n,
                    "sell_price_usd_kg": _n,
                },
                "required": ["product_code"],
            },
        },
        "booking": {
            "type": ["object", "null"],
            "properties": {
                "vessel_voyage": _s, "booking_no": _s, "etd": _d,
                "port_of_loading": _s, "place_of_delivery": _s, "port_of_discharge": _s,
                "container_no": _s, "seal_no": _s,
            },
        },
        "stuffing_date": {**_d, "description": "Ngày đóng hàng - do người dùng cung cấp, không đoán"},
        "uncertain_fields": {"type": "array", "items": {"type": "string"},
                             "description": "Tên các trường AI không chắc, cần người kiểm tra"},
        "evidence": {"type": ["object", "null"],
                     "description": "Trích dẫn ngắn (<15 từ) cho các trường quan trọng: giá, số lượng, ngày, cảng"},
    },
    "required": ["lines"],
}
