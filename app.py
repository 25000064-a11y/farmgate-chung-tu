"""Web app tạo chứng từ xuất khẩu Farmgate SG.

Chạy thử trên máy:   python3 app.py   (mở http://localhost:8000)
Trên host:           gunicorn app:app
Biến môi trường:     ANTHROPIC_API_KEY (bắt buộc cho nút "Đọc bằng AI"), APP_PASSWORD (khuyên dùng),
                     FARMGATE_MODEL (tuỳ chọn)
Web không lưu gì trên server: mỗi lần dùng, bạn gửi master.xlsx lên và tải bản cập nhật về.
"""
import argparse
import hmac
import io
import json
import os
import shutil
import tempfile
import zipfile
from pathlib import Path

from flask import Flask, Response, jsonify, request, send_file

import common
import extract as ex
import generate as gen
import import_json as ij

HERE = Path(__file__).parent
DEFAULT_MASTER = HERE / "master.xlsx"
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 40 * 1024 * 1024


@app.before_request
def _auth():
    pw = os.environ.get("APP_PASSWORD")
    if not pw or request.path == "/healthz":
        return None
    a = request.authorization
    if a and hmac.compare_digest((a.password or ""), pw):
        return None
    return Response("Cần mật khẩu", 401, {"WWW-Authenticate": 'Basic realm="Farmgate"'})


@app.errorhandler(Exception)
def _err(e):
    code = getattr(e, "code", 500) if hasattr(e, "code") and isinstance(getattr(e, "code"), int) else 500
    return jsonify(error=f"{type(e).__name__}: {e}"), code


@app.get("/healthz")
def healthz():
    return "ok"


@app.get("/")
def index():
    p = HERE / "web" / "index.html"
    return send_file(p if p.exists() else HERE / "index.html")


@app.get("/api/master_default")
def master_default():
    return send_file(DEFAULT_MASTER, as_attachment=True, download_name="master.xlsx")


def _save_master(tmp):
    f = request.files.get("master")
    dst = Path(tmp) / "master.xlsx"
    if f and f.filename:
        f.save(dst)
    else:
        shutil.copy(DEFAULT_MASTER, dst)
    return dst


@app.post("/api/extract")
def api_extract():
    if not os.environ.get("ANTHROPIC_API_KEY"):
        return jsonify(error="Server chưa có ANTHROPIC_API_KEY. Dùng 'Nhập tay' hoặc báo người quản trị."), 400
    with tempfile.TemporaryDirectory() as tmp:
        master = _save_master(tmp)
        paths = {}
        for k in ("purchase", "sales", "booking"):
            f = request.files.get(k)
            if f and f.filename:
                p = Path(tmp) / (k + Path(f.filename).suffix.lower())
                f.save(p)
                paths[k] = str(p)
        if not paths:
            return jsonify(error="Chưa chọn file hợp đồng/booking nào"), 400
        stuffing = request.form.get("stuffing") or None
        a = argparse_ns(op_code=request.form.get("op_code", "").strip(), supplier=request.form.get("supplier") or None,
                        buyer=request.form.get("buyer") or None, stuffing=stuffing, master=str(master),
                        model=os.environ.get("FARMGATE_MODEL", ex.DEFAULT_MODEL), **{k: paths.get(k) for k in ("purchase", "sales", "booking")})
        if not a.op_code:
            return jsonify(error="Nhập mã OP trước"), 400
        data = ex.run_extraction(a)
        return jsonify(data=data)


def argparse_ns(**kw):
    return argparse.Namespace(**kw)


def _build(tmp):
    """Nhập dữ liệu vào master (bản tạm) và sinh chứng từ. Trả về dict kết quả."""
    data = json.loads(request.form["data"])
    master = _save_master(tmp)
    try:
        wb, changes, warnings = ij.import_lot(data, str(master))
    except ij.ImportError_ as e:
        raise ValueError(str(e))
    wb.save(master)
    out = Path(tmp) / "docs"
    try:
        made, skipped, ctx = gen.generate_lot(str(master), data["op_code"], out)
    except common.MasterError as e:
        raise ValueError(f"Thiếu/sai dữ liệu: {e}")
    return dict(data=data, master=master, out=out / data["op_code"], made=made, skipped=skipped, ctx=ctx,
                changes=[f"{w}: {n}" for w, o, n in changes], warnings=warnings)


@app.post("/api/check")
def api_check():
    with tempfile.TemporaryDirectory() as tmp:
        try:
            r = _build(tmp)
        except ValueError as e:
            return jsonify(error=str(e)), 400
        t = r["ctx"]["tot"]
        return jsonify(made=r["made"], skipped=r["skipped"], warnings=r["warnings"], changes=r["changes"],
                       totals=dict(cartons=t["cartons_fmt"], net_kg=t["qty_fmt"], buy=t["buy_fmt"],
                                   sell=t["sell_fmt"], margin=f"{r['ctx']['margin']:,.2f}"))


@app.post("/api/generate")
def api_generate():
    with tempfile.TemporaryDirectory() as tmp:
        try:
            r = _build(tmp)
        except ValueError as e:
            return jsonify(error=str(e)), 400
        op = r["data"]["op_code"]
        buf = io.BytesIO()
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for f in sorted(r["out"].iterdir()):
                z.write(f, f"{op}/{f.name}")
            z.write(r["master"], "master.xlsx")
            z.writestr(f"lots/{op}.json", json.dumps(r["data"], ensure_ascii=False, indent=2))
            if r["skipped"]:
                z.writestr(f"{op}/_BO_QUA.txt", "\n".join(r["skipped"]))
        buf.seek(0)
        return send_file(buf, mimetype="application/zip", as_attachment=True, download_name=f"{op}_chung_tu.zip")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 8000)), debug=False)
