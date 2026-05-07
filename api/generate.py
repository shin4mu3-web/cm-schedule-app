"""
CMスケ表 - PDF→Excel変換 APIエンドポイント
Vercel Python Serverless Function (Flask)
"""

from flask import Flask, request, send_file, jsonify
import pdfplumber
from openpyxl import load_workbook
import io
import re
import os
import tempfile
from datetime import time, datetime, timedelta

app = Flask(__name__)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), "template.xlsx")
DOW_JA = "月火水木金土日"
DOW_TO_COL = {0: 1, 1: 5, 2: 9, 3: 13, 4: 17, 5: 21, 6: 25}


# ── PDF解析 ──────────────────────────────────────

def zen2han(s: str) -> str:
    return s.translate(str.maketrans(
        "ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏｐｑｒｓｔｕｖｗｘｙｚ０１２３４５６７８９",
        "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789"
    ))


def parse_meisai(pdf_bytes: bytes) -> dict:
    with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)
        all_tables = [t for p in pdf.pages for t in p.extract_tables()]

    def ex(pattern, default=""):
        m = re.search(pattern, full_text)
        return m.group(1).strip() if m else default

    meta = {
        "contract_code": zen2han(ex(r"契約コード\s*：\s*([A-Za-zＡ-Ｚ０-９0-9]+")),
        "sponsor":       ex(r"スポンサー\s*：\s*(.+?)\s+A単価"),
        "agency":        ex(r"広告会社\s*：\s*(.+?)\s+評価欄帯域"),
        "person":        ex(r"外勤\s*：\s*(.+)"),
        "product":       ex(r"商品名\s*：\s*(.+?)\s+枠取りパターン"),
        "seconds":       int(ex(r"枠取り秒数\s*：\s*(\d+)", "15")),
    }

    # 放送局
    sm = re.search(r"^([^\s/]+／[^\s]+)\s+\d+／\d+", full_text, re.MULTILINE)
    meta["station"] = sm.group(1).split("／")[0] if sm else ""

    # 契約期間
    pm = re.search(r"契約期間\s*：\s*(\d{4}/\d{2}/\d{2})～(\d{4}/\d{2}/\d{2})", full_text)
    if not pm:
        raise ValueError("契約期間が取得できません")
    meta["period_start"] = datetime.strptime(pm.group(1), "%Y/%m/%d")
    meta["period_end"]   = datetime.strptime(pm.group(2), "%Y/%m/%d")

    # 総本数
    tm = re.search(r"枠取りパターン[^\n]+?(\d+)本\s*$", full_text, re.MULTILINE)
    meta["total_count"] = int(tm.group(1)) if tm else 0

    # スケジュール（テーブルから）
    meta["schedule"] = _parse_schedule(all_tables)
    return meta


def _parse_schedule(tables: list) -> dict:
    schedule = {}
    day_re   = re.compile(r"≪(\d{2}/\d{2})\s*[月火水木金土日]≫")
    entry_re = re.compile(r"ﾚ\s*(\d{4})([TS])\s*([ABCS])\s*(\d+)\"")

    for table in tables:
        current_days = {}
        for row in table:
            if not row:
                continue
            day_cols = {
                ci: day_re.search(str(cell)).group(1)
                for ci, cell in enumerate(row)
                if cell and day_re.search(str(cell))
            }
            if day_cols:
                current_days = day_cols
                for d in current_days.values():
                    schedule.setdefault(d, [])
                continue
            for ci, cell in enumerate(row):
                if not cell or ci not in current_days:
                    continue
                em = entry_re.search(str(cell))
                if not em:
                    continue
                raw = int(em.group(1))
                past = raw >= 2400
                if past:
                    raw -= 2400
                h, m = raw // 100, raw % 100
                schedule[current_days[ci]].append({
                    "time_obj":      time(h, m),
                    "sb_pt":         "PT" if em.group(2) == "T" else "SB",
                    "seconds":       int(em.group(4)),
                    "past_midnight": past,
                })
    return schedule


# ── Excel生成 ─────────────────────────────────────

def _excel_time(entry):
    t = entry["time_obj"]
    if entry["past_midnight"]:
        return datetime(1900, 1, 1, t.hour, t.minute)
    return t


def _cm_abbr(date_str: str, year: int, materials: list) -> str:
    dt = datetime(year, int(date_str[:2]), int(date_str[3:]))
    for mat in materials:
        pm = re.match(r"(\d+)/(\d+)[〜~](\d+)/(\d+)", mat.get("period", ""))
        if pm:
            if datetime(year, int(pm.group(1)), int(pm.group(2))) <= dt <= \
               datetime(year, int(pm.group(3)), int(pm.group(4))):
                return mat["abbr"]
    return materials[0]["abbr"] if materials else ""


def build_excel(meta: dict, materials: list, station_name: str,
                station_person: str, doc_date: str) -> bytes:
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb.active
    year = meta["period_start"].year

    # 日付
    if doc_date:
        try:
            ws.cell(2, 22, datetime.strptime(doc_date, "%Y/%m/%d"))
        except ValueError:
            pass

    # 放送局
    loc = station_name or meta.get("station", "")
    if station_person:
        loc += f"　{station_person}様"
    ws.cell(3, 2, loc)

    ws.cell(4, 2, meta["sponsor"])
    ws.cell(4, 10, f"契約No.{meta['contract_code']}")

    ps, pe = meta["period_start"], meta["period_end"]
    ws.cell(5, 2, (
        f"{ps.year-2000}年{ps.month}月{ps.day}日({DOW_JA[ps.weekday()]})"
        f"～{pe.month}月{pe.day}日({DOW_JA[pe.weekday()]})"
    ))
    ws.cell(5, 18, meta["total_count"])

    for i, mat in enumerate(materials):
        r = 8 + i
        ws.cell(r, 1,  meta["sponsor"])
        ws.cell(r, 3,  mat.get("name", ""))
        ws.cell(r, 9,  int(mat.get("seconds", meta["seconds"])))
        ws.cell(r, 10, mat.get("abbr", ""))
        ws.cell(r, 12, mat.get("material", "OL"))
        ws.cell(r, 14, mat.get("status", "送"))
        ws.cell(r, 16, mat.get("period", ""))

    schedule = meta["schedule"]
    if not schedule:
        out = io.BytesIO()
        wb.save(out)
        return out.getvalue()

    date_objs = {d: datetime(year, int(d[:2]), int(d[3:])) for d in schedule}
    sorted_d  = sorted(date_objs, key=lambda d: date_objs[d])
    first_dt, last_dt = date_objs[sorted_d[0]], date_objs[sorted_d[-1]]
    first_mon = first_dt - timedelta(days=first_dt.weekday())

    weeks, cur = [], first_mon
    while cur <= last_dt:
        week = []
        for i in range(7):
            d = cur + timedelta(days=i)
            ds = f"{d.month:02d}/{d.day:02d}"
            week.append((ds, d, schedule.get(ds, [])))
        if any(e for _, _, e in week):
            weeks.append(week)
        cur += timedelta(days=7)

    crow = 13
    for week in weeks:
        for ds, dt, entries in week:
            if entries:
                ws.cell(crow, DOW_TO_COL[dt.weekday()],
                        f"{dt.month}/{dt.day}({DOW_JA[dt.weekday()]})")
        mx = max(len(e) for _, _, e in week)
        for idx in range(mx):
            row = crow + 1 + idx
            for ds, dt, entries in week:
                if idx >= len(entries):
                    continue
                e   = entries[idx]
                col = DOW_TO_COL[dt.weekday()]
                ws.cell(row, col,     _excel_time(e))
                ws.cell(row, col + 1, e["sb_pt"])
                ws.cell(row, col + 2, e["seconds"])
                ws.cell(row, col + 3, _cm_abbr(ds, year, materials))
        crow += 1 + mx + 1

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# ── Flask ルート ──────────────────────────────────

@app.route("/api/generate", methods=["POST"])
def generate():
    meisai_file = request.files.get("meisai")
    if not meisai_file:
        return jsonify({"error": "明細リストPDFが必要です"}), 400

    try:
        meta = parse_meisai(meisai_file.read())
    except Exception as e:
        return jsonify({"error": f"PDF解析エラー: {str(e)}"}), 422

    import json as _json
    materials_raw = request.form.get("materials", "[]")
    try:
        materials = _json.loads(materials_raw)
    except Exception:
        materials = []

    excel_bytes = build_excel(
        meta=meta,
        materials=materials,
        station_name=request.form.get("station", ""),
        station_person=request.form.get("person", ""),
        doc_date=request.form.get("date", ""),
    )

    sponsor_short = re.sub(r"[^\w]", "", meta.get("sponsor", "output"))[:20]
    filename = f"CMスケ表_{sponsor_short}.xlsx"

    return send_file(
        io.BytesIO(excel_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename,
    )


@app.route("/api/health")
def health():
    return jsonify({"status": "ok"})
