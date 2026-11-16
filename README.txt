BO TAO CHUNG TU XUAT KHAU FARMGATE SG  (huong dan ngan)
=======================================================
Cai dat 1 lan:
    pip3 install openpyxl python-docx docxtpl jinja2 jsonschema anthropic

QUY TRINH 1 LO HANG (vi du OP01943)
-----------------------------------
Buoc 1 - AI doc hop dong/booking -> file JSON   (can ANTHROPIC_API_KEY)
    export ANTHROPIC_API_KEY=sk-ant-...
    python3 extract.py OP01943 --purchase "HD mua HP.docx" --sales "HD ban JT.pdf" \
            --booking "booking.pdf" --stuffing 09/11/2026 --supplier HP --buyer JT
    -> lots/OP01943.json. MO FILE RA KIEM TRA (muc uncertain_fields = cho AI khong chac).
    Kiem tra truoc khi goi AI:  them --dry-run

Buoc 2 - Nhap JSON vao master.xlsx (co the nhap tay thang vao master neu khong dung AI)
    python3 import_json.py lots/OP01943.json --recalc
    -> in ra tung o thay doi + danh sach "CAN KIEM TRA". master cu duoc luu thanh master.xlsx.bak

Buoc 3 - Sinh chung tu
    python3 generate.py OP01943
    -> output/OP01943/  (5 file = 8 chung tu)

Lam 1 phat:  python3 extract.py ... --import --generate

LUU Y
- Ngay dong hang (--stuffing) ban cung cap; AI khong tu doan. Ngay tren moi chung tu = ngay nay.
- Dien JSON null = giu nguyen o dang co trong master (khong ghi de).
- So thu tu hop dong: neu hop dong da co so (vd FGSG/HP/02112026/05) AI doc ra "seq"; neu chua co
  thi he thong tu lay so ke tiep cua doi tac (HP dang 7 -> 8, JT 1 -> 2).
- Doi tac moi: them mot dong vao sheet Parties (hoac de AI doc thong tin va tu them, nho kiem tra
  lai vendor_code, tai khoan ngan hang). San pham moi: them vao sheet Products truoc.
- Hop dong mua chi ho tro thanh toan DEPOSIT_BALANCE_ON_DOCS; hop dong ban chi ho tro DAYS_FROM_BL.
  Loai khac se bao "skip", khong tao hop dong do.
- Sua noi dung mau: sua file trong templates_original/ roi chay: python3 prep_templates.py
