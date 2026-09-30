# -*- coding: utf-8 -*-
"""
OCARCARE Grand Prix — build docs/data.js từ báo cáo 17h + file mapping.

Cách dùng:
    python tools/build_data.py                 # lấy file mới nhất trong input/ và mapping/
    python tools/build_data.py --report "đường_dẫn.xlsx" --date 2026-09-29

Chỉ xuất số liệu ĐÃ TỔNG HỢP theo đơn vị (không có tên KH, SĐT, biển số...).
Yêu cầu: pip install openpyxl
"""
import argparse, datetime as dt, glob, json, os, re, sys, unicodedata
from collections import defaultdict

try:
    import openpyxl
except ImportError:
    sys.exit("Thiếu thư viện openpyxl. Chạy: pip install openpyxl")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# ---------------------------------------------------------------- CẤU HÌNH THỂ LỆ
CONFIG = {
    "program": {
        "name": "Chào sân OCARCARE — Launch Sprint",
        "start": "2026-09-21",
        "end": "2026-10-31",
        "channel_keyword": "OPES",          # lọc cột "Danh mục Kênh bán hàng"
    },
    "early": {"tiers": [
        {"from": 1,   "to": 300, "reward": 300000, "label": "300 HĐBH đầu tiên"},
        {"from": 301, "to": 500, "reward": 200000, "label": "200 HĐBH tiếp theo"},
        {"from": 501, "to": 600, "reward": 100000, "label": "100 HĐBH cuối cùng"},
    ]},
    "individual": {"threshold": 125_000_000, "prizes": [5_000_000, 4_000_000, 3_000_000],
                   "note": "Sẽ cập nhật sau 31/10/2026"},
    "branch": {"threshold": 250_000_000, "prizes": [7_000_000, 5_000_000, 3_000_000]},
    "hub": {"threshold": 1_250_000_000, "prizes": [15_000_000, 10_000_000],
            "milestones": [
                {"value": 1_250_000_000, "label": "Điều kiện"},
                {"value": 1_825_000_000, "label": "+5 triệu"},
                {"value": 2_500_000_000, "label": "+10 triệu"}],
            "manager": {"threshold": 400_000_000,
                        "prizes": [6_000_000, 5_000_000, 4_000_000, 3_000_000],
                        "note": "Chưa phân tách được cấp phòng/nhóm Hub — công bố theo kết quả chính thức"}},
}

# ---------------------------------------------------------------- tiện ích
def norm(s):
    if s is None:
        return ""
    s = unicodedata.normalize("NFC", str(s)).replace("\n", " ")
    return re.sub(r"\s+", " ", s).strip().lower()

def to_date(v):
    if v is None or v == "":
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    s = str(v).strip()
    for fmt in ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y %H:%M:%S", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None

def to_num(v):
    if v is None or v == "":
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(",", "").replace(" ", "")
    try:
        return float(s)
    except ValueError:
        return 0.0

def latest(pattern):
    files = [f for f in glob.glob(pattern) if not os.path.basename(f).startswith("~$")]
    if not files:
        return None
    return max(files, key=os.path.getmtime)

def find_col(rows, *keys, start=0):
    """Tìm cột có tiêu đề chứa tất cả keys (đã chuẩn hoá) trong các dòng tiêu đề."""
    keys = [norm(k) for k in keys]
    for r in rows:
        for j, v in enumerate(r):
            if j < start:
                continue
            n = norm(v)
            if n and all(k in n for k in keys):
                return j
    return None

# ---------------------------------------------------------------- đọc mapping
def load_mapping(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hdr = rows[:3]
    c_type = find_col(hdr, "cn/pgd")
    c_name = find_col(hdr, "tên cn-pgd")
    c_code = find_col(hdr, "branch code", "mã cn")
    c_pname = find_col(hdr, "tên cn quản lý")
    c_pcode = find_col(hdr, "mã cn mẹ")
    c_region = find_col(hdr, "miền")
    c_prov = find_col(hdr, "tỉnh/thành phố")
    c_pvi = find_col(hdr, "tên đơn vị pvi phụ trách")
    need = dict(type=c_type, name=c_name, code=c_code, pname=c_pname, pcode=c_pcode)
    miss = [k for k, v in need.items() if v is None]
    if miss:
        sys.exit(f"File mapping thiếu cột: {miss}")

    def g(r, c):
        return (str(r[c]).strip() if c is not None and c < len(r) and r[c] is not None else "")

    m = {}
    for r in rows[3:]:
        code = g(r, c_code).upper()
        if not code.startswith("VN"):
            continue
        m[code] = dict(code=code, type=g(r, c_type).upper(), name=g(r, c_name),
                       pname=g(r, c_pname), pcode=g(r, c_pcode).upper(),
                       region=g(r, c_region), province=g(r, c_prov), pvi=g(r, c_pvi))
    return m

def unit_of(rec, mapping):
    """Trả về (loại, mã đơn vị thi đua, tên) theo quy tắc SPEC §3."""
    t = rec["type"]
    if "TTKD" in t or "OTO" in rec["name"].upper():
        return "HUB", rec["code"], rec["name"]
    if t == "TTTC":
        return "TTTC", rec["code"], rec["name"]
    pcode = rec["pcode"] or rec["code"]
    pname = rec["pname"] or rec["name"]
    return "CN", pcode, pname

# ---------------------------------------------------------------- đọc báo cáo
def load_report(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    ws = None
    for s in wb.worksheets:
        if "oto" in s.title.lower() or "nvu" in s.title.lower():
            ws = s
            break
    ws = ws or wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    hi = None
    for i, r in enumerate(rows[:40]):
        ns = [norm(v) for v in r]
        if "số đơn bảo hiểm" in ns and any("mã chi nhánh bank" in n for n in ns):
            hi = i
            break
    if hi is None:
        sys.exit("Không tìm thấy dòng tiêu đề (cần có 'Số đơn bảo hiểm' và 'Mã chi nhánh Bank').")
    hdr = rows[hi:hi + 2]
    c = dict(
        pol=find_col([rows[hi]], "số đơn bảo hiểm"),
        sdbs=find_col([rows[hi]], "số đơn sửa đổi bổ sung"),
        issued=find_col([rows[hi]], "ngày cấp"),
        channel=find_col([rows[hi]], "danh mục kênh bán hàng"),
        fee=find_col([rows[hi]], "tổng phí bh"),
        bank=find_col([rows[hi]], "mã chi nhánh bank"),
    )
    if c["fee"] is None or c["bank"] is None or c["pol"] is None or c["issued"] is None:
        sys.exit(f"Báo cáo thiếu cột bắt buộc: {c}")
    out = []
    for r in rows[hi + 3:]:
        if c["pol"] >= len(r):
            continue
        pol = r[c["pol"]]
        if not pol or not str(pol).strip() or "/" not in str(pol):
            continue
        out.append(dict(
            pol=str(pol).strip(),
            sdbs=str(r[c["sdbs"]] or "").strip() if c["sdbs"] is not None else "",
            issued=to_date(r[c["issued"]]),
            channel=str(r[c["channel"]] or "") if c["channel"] is not None else "",
            fee=to_num(r[c["fee"]]),
            bank=str(r[c["bank"]] or "").strip().upper(),
        ))
    # dòng SĐBS để trống Mã chi nhánh Bank → kế thừa từ đơn gốc
    origin = {x["pol"]: x["bank"] for x in out if x["bank"] and not x["sdbs"]}
    for x in out:
        if not x["bank"]:
            x["bank"] = origin.get(x["pol"], "")
    return out

# ---------------------------------------------------------------- tổng hợp
def rank(items):
    items.sort(key=lambda x: (-x["premium"], -x["policies"], x["name"]))
    for i, x in enumerate(items, 1):
        x["rank"] = i
    return items

def aggregate(recs, mapping, cutoff=None):
    """cutoff: chỉ tính đơn có ngày cấp < cutoff (để tính 'hôm qua')."""
    units = {}
    seen = set()
    unmapped = {"premium": 0.0, "policies": 0, "codes": set()}
    for x in recs:
        if cutoff and x["issued"] >= cutoff:
            continue
        rec = mapping.get(x["bank"])
        new_pol = not x["sdbs"] and x["pol"] not in seen
        if new_pol:
            seen.add(x["pol"])
        if not rec:
            unmapped["premium"] += x["fee"]
            unmapped["policies"] += 1 if new_pol else 0
            unmapped["codes"].add(x["bank"] or "(trống)")
            continue
        kind, uid, uname = unit_of(rec, mapping)
        u = units.get(uid)
        if not u:
            base = mapping.get(uid, rec)
            u = units[uid] = dict(id=uid, kind=kind, name=uname,
                                  region=base["region"] or rec["region"],
                                  province=base["province"] or rec["province"],
                                  pvi=base["pvi"] or rec["pvi"],
                                  premium=0.0, policies=0)
        u["premium"] += x["fee"]
        u["policies"] += 1 if new_pol else 0
    return units, unmapped, len(seen)

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--report")
    ap.add_argument("--mapping")
    ap.add_argument("--date", help="Ngày báo cáo YYYY-MM-DD (mặc định lấy từ tên file)")
    ap.add_argument("--out", default=os.path.join(ROOT, "docs"))
    a = ap.parse_args()

    report = a.report or latest(os.path.join(ROOT, "input", "*.xls*"))
    mfile = a.mapping or latest(os.path.join(ROOT, "mapping", "*.xls*"))
    if not report:
        sys.exit("Không có file báo cáo trong thư mục input/")
    if not mfile:
        sys.exit("Không có file mapping trong thư mục mapping/")
    print(f"• Báo cáo : {os.path.basename(report)}")
    print(f"• Mapping : {os.path.basename(mfile)}")

    mapping = load_mapping(mfile)
    recs_all = load_report(report)

    P = CONFIG["program"]
    start, end = dt.date.fromisoformat(P["start"]), dt.date.fromisoformat(P["end"])
    kw = P["channel_keyword"].lower()
    recs = [x for x in recs_all if x["issued"] and start <= x["issued"] <= end
            and (not x["channel"] or kw in x["channel"].lower())]
    print(f"• Dòng dữ liệu: {len(recs_all)} | trong kỳ & kênh {P['channel_keyword']}: {len(recs)}")

    # ngày báo cáo
    rdate = None
    if a.date:
        rdate = dt.date.fromisoformat(a.date)
    else:
        mt = re.search(r"(\d{2})(\d{2})(20\d{2})", os.path.basename(report))
        if mt:
            rdate = dt.date(int(mt.group(3)), int(mt.group(2)), int(mt.group(1)))
    if not rdate:
        rdate = max((x["issued"] for x in recs), default=dt.date.today())

    now_u, unmapped, total_pol = aggregate(recs, mapping)
    prev_u, _, prev_pol = aggregate(recs, mapping, cutoff=rdate)

    def board(kind_filter, thr):
        cur = rank([dict(u) for u in now_u.values() if u["kind"] in kind_filter])
        prv = {u["id"]: u for u in rank([dict(u) for u in prev_u.values() if u["kind"] in kind_filter])}
        for u in cur:
            p = prv.get(u["id"])
            u["prev_rank"] = p["rank"] if p else None
            u["today_premium"] = round(u["premium"] - (p["premium"] if p else 0))
            u["today_policies"] = u["policies"] - (p["policies"] if p else 0)
            u["qualified"] = u["premium"] >= thr
            u["crossed_today"] = u["qualified"] and not (p and p["premium"] >= thr)
            u["premium"] = round(u["premium"])
        return cur

    branches = board({"CN", "TTTC"}, CONFIG["branch"]["threshold"])
    hubs = board({"HUB"}, CONFIG["hub"]["threshold"])
    # luôn hiển thị đủ 2 Hub kể cả khi chưa có đơn
    for code, rec in mapping.items():
        if unit_of(rec, mapping)[0] == "HUB" and code not in {h["id"] for h in hubs}:
            hubs.append(dict(id=code, kind="HUB", name=rec["name"], region=rec["region"],
                             province=rec["province"], pvi=rec["pvi"], premium=0, policies=0,
                             prev_rank=None, today_premium=0, today_policies=0,
                             qualified=False, crossed_today=False))
    rank(hubs)

    # tra cứu: mọi CN/PGD/TTTC/TTKD → đơn vị thi đua
    lookup = []
    for code, rec in mapping.items():
        kind, uid, uname = unit_of(rec, mapping)
        lookup.append(dict(n=rec["name"], t=rec["type"] or kind, u=uid, un=uname, k=kind,
                           p=rec["province"]))

    # luỹ kế theo ngày
    daily = []
    d = start
    seen = set()
    cum_prem, cum_pol = 0.0, 0
    by_day = defaultdict(list)
    for x in recs:
        by_day[x["issued"]].append(x)
    while d <= min(rdate, end):
        for x in by_day.get(d, []):
            cum_prem += x["fee"]
            if not x["sdbs"] and x["pol"] not in seen:
                seen.add(x["pol"])
                cum_pol += 1
        daily.append(dict(date=d.isoformat(), policies=cum_pol, premium=round(cum_prem)))
        d += dt.timedelta(days=1)

    def group(key):
        g = defaultdict(lambda: dict(premium=0.0, policies=0))
        for u in now_u.values():
            k = u[key] or "Khác"
            g[k]["premium"] += u["premium"]
            g[k]["policies"] += u["policies"]
        return sorted([dict(name=k, premium=round(v["premium"]), policies=v["policies"])
                       for k, v in g.items()], key=lambda x: -x["premium"])

    total_prem = sum(x["fee"] for x in recs)
    today_recs = [x for x in recs if x["issued"] == rdate]
    data = dict(
        meta=dict(
            report_date=rdate.isoformat(), report_time="17:00",
            generated_at=dt.datetime.now().strftime("%Y-%m-%d %H:%M"),
            source=os.path.basename(report),
            total_policies=total_pol, total_premium=round(total_prem),
            today_policies=total_pol - prev_pol,
            today_premium=round(sum(x["fee"] for x in today_recs)),
            active_units=len([b for b in branches if b["policies"] > 0]),
            unmapped=dict(premium=round(unmapped["premium"]), policies=unmapped["policies"]),
        ),
        config=CONFIG,
        branches=branches, hubs=hubs, lookup=lookup, daily=daily,
        regions=group("region"), pvi_units=group("pvi"),
    )

    os.makedirs(os.path.join(a.out, "history"), exist_ok=True)
    js = "window.RACE_DATA = " + json.dumps(data, ensure_ascii=False, separators=(",", ":")) + ";\n"
    with open(os.path.join(a.out, "data.js"), "w", encoding="utf-8") as f:
        f.write(js)
    with open(os.path.join(a.out, "history", f"data-{rdate.isoformat()}.json"), "w", encoding="utf-8") as f:
        json.dump({**data["meta"], **{"branches": [{k: b[k] for k in ("id", "name", "premium", "policies", "rank")} for b in branches],
                                  "hubs": [{k: b[k] for k in ("id", "name", "premium", "policies", "rank")} for b in hubs]}},
                  f, ensure_ascii=False, indent=1)

    # ---- báo cáo kiểm tra
    print(f"• Ngày báo cáo: {rdate:%d/%m/%Y}")
    print(f"• Tổng: {total_pol} HĐBH | {total_prem:,.0f}đ (chưa VAT) | hôm nay +{data['meta']['today_policies']} HĐ")
    print(f"• Chi nhánh/TTTC có đơn: {len(branches)} | Hub: " +
          ", ".join(f"{h['name']} {h['premium']:,.0f}đ" for h in hubs))
    chk = sum(b["premium"] for b in branches) + sum(h["premium"] for h in hubs) + unmapped["premium"]
    if abs(chk - total_prem) > 5:
        print(f"  ⚠ Lệch tổng: {chk:,.0f} vs {total_prem:,.0f}")
    if unmapped["premium"] or unmapped["policies"]:
        print(f"  ⚠ {unmapped['policies']} HĐ có Mã chi nhánh Bank chưa có trong mapping: {sorted(unmapped['codes'])}")
    print("✔ Đã ghi docs/data.js")

if __name__ == "__main__":
    main()
