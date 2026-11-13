"""세트 기록 — 폰 앱에서 복사한 텍스트를 세트기록.csv에 합치고, 일차 · 종목별 흐름을 보여 준다.

2026-11-13 사용자 요청: "2일차 관련 데이터 빼내 줘" 하면 그 일차의 모든 기록을 대화창에 흐름으로.
폰 기록은 폰 브라우저(localStorage wk-sets)에만 있다 — 앱 일차 화면의 'PC로 보낼 세트 기록 복사'를
Claude에게 붙여 넣으면 add로 여기 쌓는다. 같은 날짜 · 종목 · 세트는 새 값으로 덮는다.

  python 세트기록.py add 붙여넣은.txt      (또는 표준 입력)
  python 세트기록.py show --day 2          (2일차 전부 · 종목별 사이클 흐름)
  python 세트기록.py show --ex 벤치         (이름에 '벤치'가 든 종목 · 모든 일차)
  python 세트기록.py show --day 2 --ex 플라이
  python 세트기록.py pull                  (구글 시트에서 받아 세트기록.csv를 시트와 같게 — 세션 시작 훅이 부른다)

2026-11-13 자동 저장: 앱이 기록할 때마다 구글 시트(구글시트_세트기록.gs)로 보내고, 세션을 열면 pull.
시트 주소는 .sync_url (git 제외 — 공개 저장소라서)에만 둔다.
"""
import argparse
import csv
import sys
import urllib.request
from collections import OrderedDict
from pathlib import Path

CSV = Path(__file__).resolve().parent / "세트기록.csv"
COLS = ["날짜", "사이클", "일차", "종목", "세트", "무게", "횟수", "달성"]


def load():
    if not CSV.exists():
        return []
    with CSV.open(encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def save(rows):
    rows.sort(key=lambda r: (r["날짜"], r["일차"], r["종목"], int(r["세트"] or 0)))
    with CSV.open("w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=COLS)
        w.writeheader()
        w.writerows(rows)


def parse_paste(text, strict=True):
    out = []
    for line in text.splitlines():
        line = line.rstrip("\r")
        if not line.strip() or line.startswith("#") or line.startswith("날짜\t"):
            continue
        p = line.split("\t")
        if len(p) < 8:
            if not strict:
                continue
            sys.exit(f"[중단] 칸이 8개가 아닌 줄: {line!r} — 앱 복사 텍스트를 그대로 붙여야 한다(탭 구분)")
        out.append(dict(zip(COLS, p[:8])))
    return out


def cmd_add(a):
    text = Path(a.file).read_text(encoding="utf-8") if a.file else sys.stdin.read()
    new = parse_paste(text)
    rows = load()
    key = lambda r: (r["날짜"], r["종목"], r["세트"])
    idx = {key(r): i for i, r in enumerate(rows)}
    added = changed = 0
    for r in new:
        k = key(r)
        if k in idx:
            if rows[idx[k]] != r:
                rows[idx[k]] = r
                changed += 1
        else:
            idx[k] = len(rows)
            rows.append(r)
            added += 1
    save(rows)
    dates = sorted({r["날짜"] for r in new})
    print(f"세트기록.csv — 새 세트 {added} · 바뀐 세트 {changed} · 받은 날짜 {', '.join(dates) or '없음'} · 전체 {len(rows)}세트")


def cmd_pull(a):
    """시트가 원본이다 — 받은 내용으로 세트기록.csv를 통째로 맞춘다. 실패하면 csv는 그대로 두고 한 줄만 알린다."""
    uf = CSV.parent / ".sync_url"
    url = uf.read_text(encoding="utf-8").strip() if uf.exists() else ""
    if not url:
        print("세트기록 자동 저장: 아직 설정 안 됨 (.sync_url 없음)")
        return
    try:
        with urllib.request.urlopen(url, timeout=15) as r:
            text = r.read().decode("utf-8")
    except Exception as e:
        print(f"세트기록 자동 저장: 시트에서 못 받음 ({type(e).__name__}) — 세트기록.csv는 지난 것 그대로")
        return
    if not text.startswith("날짜	"):
        print("세트기록 자동 저장: 시트 응답이 표가 아님 (배포 권한 '모든 사용자' 확인) — csv 그대로")
        return
    new = parse_paste(text, strict=False)
    old = load()
    key = lambda r: (r["날짜"], r["종목"], r["세트"])
    ok, nk = {key(r): r for r in old}, {key(r): r for r in new}
    if not new and old:
        print(f"세트기록 자동 저장: 시트가 비어 있음 — csv {len(old)}세트는 지우지 않고 둔다")
        return
    add = len(nk.keys() - ok.keys())
    chg = sum(1 for k in nk if k in ok and ok[k] != nk[k])
    gone = len(ok.keys() - nk.keys())
    try:
        save(list(new))
    except PermissionError:
        print("세트기록 자동 저장: 세트기록.csv가 엑셀에 열려 있어 못 씀 — 엑셀을 닫고 '세트기록 받아 줘'라고 말하기")
        return
    last = max((r["날짜"] for r in new), default="-")
    msg = f"세트기록 자동 저장: 세트기록.csv {len(new)}세트 (새 {add} · 바뀜 {chg} · 지움 {gone}) · 마지막 기록 {last}"
    if write_view():
        msg += " · 세트기록_보기.xlsx 갱신"
    else:
        msg += " · 세트기록_보기.xlsx는 엑셀에 열려 있어 못 고침"
    print(msg)


VIEW = CSV.parent.parent / "세트기록_보기.xlsx"   # 2026-11-13 폴더 정리 — 스크립트 · csv · .sync_url은 도구/, 보는 파일만 운동/에


def write_view():
    """폴더에서 바로 열어 보는 비교 표(엑셀) — 세션 시작 때마다 새로 쓴다. 엑셀에 열려 있으면 건너뛴다"""
    try:
        import openpyxl
        from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
        from openpyxl.utils import get_column_letter
    except ImportError:
        return False
    bl = card_order(list(blocks(load())))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    # 2026-11-13 사용자 요청: 일차마다 시트 하나(1 · 2 · 3 · 4 · 5 · 7 · 8 — 휴식일 6 · 9 빼고) · 종목은 앱 카드 순서 · 그래프는 한 시트에 모아 둔다
    for day in train_days(bl):
        sheet_day(wb.create_sheet(f"{day}일차"), day, [b for b in bl if b["day"] == day])
    charts(wb, bl)
    try:
        wb.save(VIEW)
        return True
    except PermissionError:
        return False


APP = CSV.parent.parent / "docs" / "index.html"   # 앱 배포본 — 카드 순서는 여기서 읽는다(build_mobile.py가 워크북에서 만든 것)


def _app_days():
    """앱의 9일 사이클 일차별 종목 이름 목록(카드 순서). 사이클 교대 종목(정지 벤치 ↔ 인클라인 등)은
    짝 종목 바로 뒤에 끼운다. 앱 파일이 없거나 못 읽으면 빈 dict — 그러면 예전처럼 이름순"""
    import json
    try:
        t = APP.read_text(encoding="utf-8")
        D, _ = json.JSONDecoder().raw_decode(t[t.index("const D = ") + len("const D = "):])
    except (OSError, ValueError):
        return {}
    out = {}
    for d, day in (D.get("back") or {}).items():
        ex = (day.get("ex") or []) + [x for _, g in (day.get("groups") or []) for x in g]
        names = [e.get("name") or "" for e in ex]
        for a in D.get("cycAlt") or []:
            if str(a[0]) == d and a[2] not in names:
                k = names.index(a[1]) + 1 if a[1] in names else len(names)
                names.insert(k, a[2])
        out[d] = names
    return out


def train_days(bl):
    """시트로 만들 일차 — 앱에서 종목이 있는 날(휴식일 빼고) + 기록에만 있는 날"""
    days = [d for d, n in _app_days().items() if n]
    days += [b["day"] for b in bl if b["day"] not in days]
    return sorted(set(days), key=lambda d: (not str(d).isdigit(), int(d) if str(d).isdigit() else 0, str(d)))


def card_order(bl):
    """블록을 일차 → 앱 카드 순서로. 같은 종목의 그립 · 방식 블록은 나란히. 지금 앱에 없는 종목(빠진 종목)은 그 일차 맨 뒤"""
    order = _app_days()

    def key(ib):
        i, b = ib
        names = order.get(str(b["day"]), [])
        pos = names.index(b["base"]) if b["base"] in names else len(names)
        d = int(b["day"]) if str(b["day"]).isdigit() else 99
        return (d, pos, i)
    return [b for _, b in sorted(enumerate(bl), key=key)]


def sheet_day(ws, day, bl):
    """일차 시트 하나 — 종목 블록을 카드 순서대로 위에서 아래로"""
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    if not bl:
        ws["A1"] = f"{day}일차 — 아직 기록 없음(다음 {day}일차 운동부터 채워진다)"
        ws["A1"].font = Font(bold=True, size=11)
        ws.column_dimensions["A"].width = 60
        return
    N = max([b["n"] for b in bl] + [1])
    ws["A1"] = f"{day}일차 세트 기록 — 앱 카드 순서 — 세션을 열 때마다 자동으로 새로 쓴다(직접 고치지 말 것) · 같은 그립 · 같은 방식(FST-7)끼리만 비교 · 진전 기준 = 추정 1RM"
    ws["A1"].font = Font(bold=True, size=11)
    head = ["날짜", "사이클"] + [f"{k}세트" for k in range(1, N + 1)] + \
           ["추정1RM", "±1RM", "최고무게", "±무게", "총횟수", "±횟수", "볼륨", "±볼륨"]
    ncol = len(head)
    T_FILL = PatternFill("solid", fgColor="DDE3EA")
    H_FILL = PatternFill("solid", fgColor="F2F4F7")
    OK_FILL = PatternFill("solid", fgColor="E3F2E6")
    UP, DN, MUTE = Font(color="1E7B34", bold=True), Font(color="C0392B", bold=True), Font(color="8A8F98")
    thin = Side(style="thin", color="D5D9E0")
    box = Border(bottom=thin)
    center = Alignment(horizontal="center", vertical="center")
    r = 3
    for b in bl:
        grip = f"   ‹그립 {b['grip']}› — 같은 그립끼리 비교" if b["grip"] else ""
        ws.cell(r, 1, f"{b['base']}{grip}").font = Font(bold=True, size=12)
        for c in range(1, ncol + 1):
            ws.cell(r, c).fill = T_FILL
        r += 1
        for c, h in enumerate(head, 1):
            x = ws.cell(r, c, h)
            x.font, x.fill, x.alignment, x.border = Font(bold=True, color="555B66"), H_FILL, center, box
        r += 1
        for s in b["sess"]:
            ws.cell(r, 1, s["dt"]).alignment = center
            ws.cell(r, 2, int(s["cyc"]) if str(s["cyc"]).isdigit() else s["cyc"]).alignment = center
            for k in range(1, N + 1):
                x = ws.cell(r, 2 + k)
                x.alignment = center
                if k in s["sets"]:
                    x.value = _cell(s["sets"][k])
                    if s["sets"][k]["달성"] == "O":
                        x.fill = OK_FILL
                elif k <= b["n"]:
                    x.value, x.font = "·", MUTE
            c0 = 3 + N
            vals = [(s["e1"], s["d_e1"], "0.0", "+0.0;-0.0"),
                    (s["top"], s["d_top"], "General", None),
                    (s["tot"], s["d_tot"], "0", "+0;-0"),
                    (s["vol"], s["d_vol"], "#,##0", "+#,##0;-#,##0")]
            for i, (v, dv, f1, f2) in enumerate(vals):
                a = ws.cell(r, c0 + 2 * i, round(v, 1) if v is not None else "-")
                a.number_format, a.alignment = f1, center
                d = ws.cell(r, c0 + 2 * i + 1)
                d.alignment = center
                if s["first"]:
                    if i == 0:
                        d.value, d.font = "첫 기록", MUTE
                elif dv is None:
                    if i == 1 and s["wtxt"]:
                        d.value = s["wtxt"]
                    else:
                        d.value, d.font = "—", MUTE
                elif abs(dv) < 1e-9:
                    d.value, d.font = "=", MUTE
                else:
                    dv = round(dv, 1)
                    d.value = int(dv) if dv == int(dv) else dv
                    d.number_format = f2 or ("+0;-0" if dv == int(dv) else "+0.0;-0.0")
                    d.font = UP if dv > 0 else DN
            r += 1
        if b["summary"]:
            ws.cell(r, 1, b["summary"]).font = Font(italic=True, color="1E5BA8")
            r += 1
        r += 1
    # 핀 무게('25 + 핀 + 1×12✓')처럼 긴 칸이 잘리지 않게 내용 폭에 맞춘다(한글 두 칸)
    sw = max([10] + [_w(_cell(x)) + 2 for b in bl for s in b["sess"] for x in s["sets"].values()])
    ww = max([8] + [_w(s["wtxt"]) + 2 for b in bl for s in b["sess"]])
    widths = [11, 7] + [sw] * N + [9, 8, 9, ww, 8, 7, 9, 8]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"


def charts(wb, bl):
    """'그래프' 시트 — 블록(일차 · 종목 · 그립 · 방식)마다 작은 선 그래프 하나, 2열 격자.
    지표 하나 · 축 하나: 추정 1RM(kg). 핀 무게라 1RM이 없는 종목은 총횟수. 두 번 이상 기록된 블록만.
    점 데이터는 숨긴 '_데이터' 시트에. y축은 0이 아니라 값 근처에서 시작해야 변화가 보인다"""
    from openpyxl.chart import LineChart, Reference
    from openpyxl.chart.label import DataLabel, DataLabelList
    from openpyxl.chart.layout import Layout, ManualLayout
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.drawing.line import LineProperties
    from openpyxl.styles import Font
    gs = wb.create_sheet("그래프")
    ds = wb.create_sheet("_데이터")
    ds.sheet_state = "hidden"
    BLUE = "2A78D6"
    use = [b for b in bl if len(b["sess"]) >= 2]
    gs["A1"] = ("종목별 흐름 — 추정 1RM(핀 무게 종목은 총횟수) · 같은 일차 · 같은 그립 · 같은 방식끼리 · "
                f"두 번 이상 기록된 {len(use)}개" + (f" (한 번뿐인 {len(bl) - len(use)}개는 다음 기록부터)" if len(bl) > len(use) else ""))
    gs["A1"].font = Font(bold=True, size=11)
    if not use:
        gs["A3"] = "아직 같은 조건으로 두 번 이상 기록된 종목이 없다 — 다음 사이클부터 선이 생긴다"
        return
    dc = 1
    for i, b in enumerate(use):
        e1 = all(s["e1"] for s in b["sess"])
        key, label, unit = ("e1", "추정 1RM", "kg") if e1 else ("tot", "총횟수", "회")
        ds.cell(1, dc, "날짜")
        ds.cell(1, dc + 1, label)
        for j, s in enumerate(b["sess"], 2):
            ds.cell(j, dc, f"{s['dt']} {s['cyc']}사이클")
            v = s[key]
            ds.cell(j, dc + 1, round(v, 1) if v is not None else None)
        n = len(b["sess"])
        ch = LineChart()
        grip = f" ‹{b['grip']}›" if b["grip"] else ""
        ch.title = f"{b['day']}일차 · {b['base']}{grip} — {label}"
        ch.add_data(Reference(ds, min_col=dc + 1, min_row=1, max_row=n + 1), titles_from_data=True)
        ch.set_categories(Reference(ds, min_col=dc, min_row=2, max_row=n + 1))
        sr = ch.series[0]
        sr.graphicalProperties.line.solidFill = BLUE
        sr.graphicalProperties.line.width = 25400          # 2pt
        sr.marker.symbol = "circle"
        sr.marker.size = 8
        sr.marker.graphicalProperties.solidFill = BLUE
        sr.marker.graphicalProperties.line.solidFill = "FFFFFF"
        sr.smooth = False
        ch.legend = None                                    # 선 하나 — 제목이 이름이다
        vals = [s[key] for s in b["sess"] if s[key] is not None]
        lo, hi = min(vals), max(vals)
        pad = max((hi - lo) * 0.25, 1)
        ch.y_axis.scaling.min = max(0, int(lo - pad))
        ch.y_axis.scaling.max = int(hi + pad) + 1
        ch.y_axis.title = unit
        ch.y_axis.majorGridlines.spPr = GraphicalProperties(ln=LineProperties(solidFill="E3E5E8", w=6350))
        ch.y_axis.majorUnit = max(1, round((ch.y_axis.scaling.max - ch.y_axis.scaling.min) / 5))
        # 숫자는 마지막 점에만 — 지금 값이 어디인지
        off = dict(showSerName=False, showCatName=False, showLegendKey=False, showPercent=False, showBubbleSize=False)
        dl = DataLabel(idx=n - 1, showVal=True, dLblPos="t", **off)
        sr.dLbls = DataLabelList(dLbl=[dl], showVal=False, **off)
        ch.x_axis.delete = False
        ch.y_axis.delete = False
        ch.y_axis.number_format = "0"
        ch.width, ch.height = 15, 7.5
        ch.layout = Layout(manualLayout=ManualLayout(x=0.08, y=0.13, w=0.88, h=0.68))
        gs.add_chart(ch, f"{'A' if i % 2 == 0 else 'J'}{3 + (i // 2) * 16}")
        dc += 3


def fmt_set(r):
    w = r["무게"] or "?"
    reps = r["횟수"] or ("목표" if r["달성"] == "O" else "?")
    return f"{w}×{reps}" + ("" if r["달성"] == "O" else "(미체크)" if r["횟수"] else "")


import datetime
import unicodedata

DOW = "월화수목금토일"


def _w(t):
    """터미널 표시 폭 — 한글은 두 칸"""
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in t)


def _pad(t, n, right=False):
    sp = " " * max(0, n - _w(t))
    return sp + t if right else t + sp


def _num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _cell(r):
    """한 세트 칸 — 52×9 · 달성(✓)이면 뒤에 ✓ · 횟수 없이 ✓만이면 52×목표"""
    w = r["무게"] or "?"
    if r["횟수"]:
        return f"{w}×{r['횟수']}" + ("✓" if r["달성"] == "O" else "")
    return f"{w}×목표✓" if r["달성"] == "O" else f"{w}×?"


def _arrow(now, before, unit=""):
    if now is None or before is None:
        return "—"
    d = now - before
    if abs(d) < 1e-9:
        return "="
    return f"{'▲' if d > 0 else '▼'}{abs(d):g}{unit}"


def _diff(a, b):
    return None if a is None or b is None else a - b


def blocks(rows):
    """종목(그립 · 방식 꼬리표 포함)마다 세션 목록과 계산값 — 터미널 표와 엑셀 표가 같이 쓴다.
    등 그립 회전 종목은 '랫풀다운 [오버 넓게 · 양손]'처럼 그립이 이름에 붙어 온다(앱 2026-11-13) —
    그립마다 따로 표 · 같은 그립끼리만 '지난번 대비'. 같은 종목의 그립 표는 나란히"""
    by_ex = OrderedDict()
    for r in sorted(rows, key=lambda r: (r["일차"], r["종목"].split(" [")[0], r["날짜"], int(r["세트"] or 0))):
        by_ex.setdefault((r["일차"], r["종목"]), OrderedDict()).setdefault(
            (r["날짜"], r["사이클"]), {})[int(r["세트"] or 0)] = r
    for (day, ex), ss in by_ex.items():
        base, _, grip = ex.partition(" [")
        out, prev = [], None
        for (d, cyc), sets in ss.items():
            ws = [_num(x["무게"]) for x in sets.values()]
            top = max([x for x in ws if x is not None], default=None)
            reps = [int(x["횟수"]) for x in sets.values() if x["횟수"].isdigit()]
            tot = sum(reps) if reps else None
            vol = sum((_num(x["무게"]) or 0) * int(x["횟수"]) for x in sets.values()
                      if _num(x["무게"]) is not None and x["횟수"].isdigit()) or None
            # 핀 · 손목 단계처럼 숫자가 아닌 무게 — 첫 세트 글자 그대로 비교한다
            raw = next((x["무게"] for x in sets.values() if x["무게"]), "")
            # 추정 1RM(Epley) — 그날 가장 좋은 세트. 무게와 횟수를 숫자 하나로: 강해졌으면 오른다
            e1 = max(((_num(x["무게"]) or 0) * (1 + int(x["횟수"]) / 30) for x in sets.values()
                      if _num(x["무게"]) and x["횟수"].isdigit()), default=None)
            dt = f"{int(d[5:7])}/{int(d[8:])}"
            try:
                dt += f"({DOW[datetime.date.fromisoformat(d).weekday()]})"
            except ValueError:
                pass
            s = {"d": d, "dt": dt, "cyc": cyc or "-", "sets": sets, "top": top, "tot": tot, "vol": vol,
                 "raw": raw, "e1": e1, "first": prev is None, "wtxt": "",
                 "d_e1": None, "d_top": None, "d_tot": None, "d_vol": None}
            if prev:
                s["d_e1"] = _diff(round(e1, 1) if e1 else None, round(prev["e1"], 1) if prev["e1"] else None)
                s["d_top"], s["d_tot"], s["d_vol"] = _diff(top, prev["top"]), _diff(tot, prev["tot"]), _diff(vol, prev["vol"])
                if s["d_top"] is None and raw != prev["raw"]:
                    s["wtxt"] = f"{prev['raw'] or '?'} → {raw or '?'}"
            out.append(s)
            prev = s
        first, last = out[0], out[-1]
        summary = ""
        if len(out) > 1 and first["e1"] and last["e1"]:
            summary = (f"→ 첫 기록 대비 추정 1RM {first['e1']:.1f} → {last['e1']:.1f} "
                       f"({(last['e1'] / first['e1'] - 1) * 100:+.1f}%) · {len(out)}회")
        elif len(out) > 1 and first["top"] is not None and last["top"] is not None:
            vv = f" · 볼륨 {(last['vol'] / first['vol'] - 1) * 100:+.0f}%" if first["vol"] and last["vol"] else ""
            summary = f"→ 첫 기록 대비 최고 무게 {_arrow(last['top'], first['top'], 'kg')}{vv} ({len(out)}회)"
        yield {"day": day, "ex": ex, "base": base, "grip": grip.rstrip("]"),
               "n": max(max(x["sets"]) for x in out), "sess": out, "summary": summary}


def cmd_show(a):
    rows = [r for r in load()
            if (not a.day or r["일차"] == a.day) and (not a.ex or a.ex in r["종목"])]
    if not rows:
        print("해당 기록 없음" + ("" if CSV.exists() else " — 세트기록.csv가 아직 없다"))
        return
    for b in card_order(list(blocks(rows))):
        n = b["n"]
        cw = max([6] + [_w(_cell(x)) for s in b["sess"] for x in s["sets"].values()])
        head = (_pad("날짜", 10) + _pad("사이클", 7) + "│ " + " ".join(_pad(f"{k}세트", cw) for k in range(1, n + 1))
                + " │ " + _pad("추정1RM", 8, True) + _pad("최고", 6, True) + _pad("총횟수", 8, True)
                + _pad("볼륨", 8, True) + " │ 지난번 대비")
        print()
        print(f"■ {b['day']}일차 · {b['base']}" + (f"  ‹그립 {b['grip']}› — 같은 그립끼리 비교" if b["grip"] else ""))
        print("  " + head)
        print("  " + "─" * _w(head))
        prev = None
        for s in b["sess"]:
            cells = " ".join(_pad(_cell(s["sets"][k]) if k in s["sets"] else "·", cw) for k in range(1, n + 1))
            if prev:
                wcmp = (_arrow(s["top"], prev["top"], "kg") if s["d_top"] is not None
                        else "=" if not s["wtxt"] else s["wtxt"])
                e1c = _arrow(round(s["e1"], 1) if s["e1"] else None, round(prev["e1"], 1) if prev["e1"] else None)
                cmp = ((f"1RM {e1c} · " if s["e1"] and prev["e1"] else "")
                       + f"무게 {wcmp} · 횟수 {_arrow(s['tot'], prev['tot'])} · 볼륨 {_arrow(s['vol'], prev['vol'])}")
            else:
                cmp = "첫 기록"
            print("  " + _pad(s["dt"], 10) + _pad(s["cyc"], 7) + "│ " + cells + " │ "
                  + _pad(f"{s['e1']:.1f}" if s["e1"] else "-", 8, True)
                  + _pad(f"{s['top']:g}" if s["top"] is not None else "-", 6, True)
                  + _pad(str(s["tot"]) if s["tot"] is not None else "-", 8, True)
                  + _pad(f"{s['vol']:,.0f}" if s["vol"] else "-", 8, True) + " │ " + cmp)
            prev = s
        if b["summary"]:
            print("  " + b["summary"])


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("add")
    p.add_argument("file", nargs="?")
    p.set_defaults(fn=cmd_add)
    p = sub.add_parser("pull")
    p.set_defaults(fn=cmd_pull)
    p = sub.add_parser("show")
    p.add_argument("--day")
    p.add_argument("--ex")
    p.set_defaults(fn=cmd_show)
    a = ap.parse_args()
    a.fn(a)
