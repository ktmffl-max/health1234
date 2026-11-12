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
        msg += " · 세트기록_보기.txt 갱신"
    print(msg)


VIEW = CSV.parent / "세트기록_보기.txt"


def write_view():
    """폴더에서 바로 열어 보는 비교 표 — show 전체 출력을 그대로 파일로(메모장에서 줄이 맞게 고정폭 글꼴 권장)"""
    import io
    from contextlib import redirect_stdout
    buf = io.StringIO()
    with redirect_stdout(buf):
        cmd_show(argparse.Namespace(day=None, ex=None))
    try:
        VIEW.write_text("세트 기록 비교 표 — 세션을 열 때마다 자동으로 새로 쓴다(직접 고치지 말 것)\n"
                        "같은 그립 · 같은 방식(FST-7)끼리만 '지난번 대비'를 잰다. 진전 기준 = 추정 1RM\n"
                        + buf.getvalue(), encoding="utf-8-sig")
        return True
    except PermissionError:
        return False


def fmt_set(r):
    w = r["무게"] or "?"
    reps = r["횟수"] or ("목표" if r["달성"] == "O" else "?")
    return f"{w}×{reps}" + ("" if r["달성"] == "O" else "(미체크)" if r["횟수"] else "")


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


def cmd_show(a):
    rows = [r for r in load()
            if (not a.day or r["일차"] == a.day) and (not a.ex or a.ex in r["종목"])]
    if not rows:
        print("해당 기록 없음" + ("" if CSV.exists() else " — 세트기록.csv가 아직 없다"))
        return
    by_ex = OrderedDict()
    # 등 그립 회전 종목은 '랫풀다운 [오버 넓게 · 양손]'처럼 그립이 이름에 붙어 온다(앱 2026-11-13) —
    # 그립마다 따로 표를 만들어 같은 그립끼리만 '지난번 대비'를 잰다. 같은 종목의 그립 표는 나란히
    for r in sorted(rows, key=lambda r: (r["일차"], r["종목"].split(" [")[0], r["날짜"], int(r["세트"] or 0))):
        by_ex.setdefault((r["일차"], r["종목"]), OrderedDict()).setdefault(
            (r["날짜"], r["사이클"]), {})[int(r["세트"] or 0)] = r
    for (day, ex), sess in by_ex.items():
        n = max(max(v) for v in sess.values())
        cw = max([6] + [_w(_cell(r)) for v in sess.values() for r in v.values()])
        head = (_pad("날짜", 10) + _pad("사이클", 7) + "│ " + " ".join(_pad(f"{k}세트", cw) for k in range(1, n + 1))
                + " │ " + _pad("추정1RM", 8, True) + _pad("최고", 6, True) + _pad("총횟수", 8, True) + _pad("볼륨", 8, True) + " │ 지난번 대비")
        print()
        base, _, grip = ex.partition(" [")
        print(f"■ {day}일차 · {base}" + (f"  ‹그립 {grip.rstrip(']')}› — 같은 그립끼리 비교" if grip else ""))
        print("  " + head)
        print("  " + "─" * _w(head))
        prev = first = None
        for (d, cyc), sets in sess.items():
            ws = [_num(r["무게"]) for r in sets.values()]
            top = max([x for x in ws if x is not None], default=None)
            reps = [int(r["횟수"]) for r in sets.values() if r["횟수"].isdigit()]
            tot = sum(reps) if reps else None
            vol = sum(_num(r["무게"]) * int(r["횟수"]) for r in sets.values()
                      if _num(r["무게"]) is not None and r["횟수"].isdigit()) or None
            # 핀 · 손목 단계처럼 숫자가 아닌 무게 — 첫 세트 글자 그대로 비교한다
            raw = next((r["무게"] for r in sets.values() if r["무게"]), "")
            # 추정 1RM(Epley) — 그날 가장 좋은 세트. 무게와 횟수를 숫자 하나로: 강해졌으면 오른다
            e1 = max((_num(r["무게"]) * (1 + int(r["횟수"]) / 30) for r in sets.values()
                      if _num(r["무게"]) and r["횟수"].isdigit()), default=None)
            cur = {"top": top, "tot": tot, "vol": vol, "raw": raw, "e1": e1}
            dt = f"{int(d[5:7])}/{int(d[8:])}"
            try:
                import datetime
                dt += f"({DOW[datetime.date.fromisoformat(d).weekday()]})"
            except ValueError:
                pass
            cells = " ".join(_pad(_cell(sets[k]) if k in sets else "·", cw) for k in range(1, n + 1))
            if prev:
                wcmp = (_arrow(top, prev['top'], 'kg') if top is not None and prev['top'] is not None
                        else "=" if raw == prev['raw'] else f"{prev['raw'] or '?'} → {raw or '?'}")
                e1c = _arrow(round(e1, 1) if e1 else None, round(prev['e1'], 1) if prev['e1'] else None)
                cmp = (f"1RM {e1c} · " if e1 and prev['e1'] else "") + f"무게 {wcmp} · 횟수 {_arrow(tot, prev['tot'])} · 볼륨 {_arrow(vol, prev['vol'])}"
            else:
                cmp = "첫 기록"
            print("  " + _pad(dt, 10) + _pad(cyc or "-", 7) + "│ " + cells + " │ "
                  + _pad(f"{e1:.1f}" if e1 else "-", 8, True)
                  + _pad(f"{top:g}" if top is not None else "-", 6, True)
                  + _pad(str(tot) if tot is not None else "-", 8, True)
                  + _pad(f"{vol:,.0f}" if vol else "-", 8, True) + " │ " + cmp)
            first = first or cur
            prev = cur
        if len(sess) > 1 and first["e1"] and prev["e1"]:
            print(f"  → 첫 기록 대비 추정 1RM {first['e1']:.1f} → {prev['e1']:.1f} "
                  f"({(prev['e1'] / first['e1'] - 1) * 100:+.1f}%) · {len(sess)}회")
        elif len(sess) > 1 and first["top"] is not None and prev["top"] is not None:
            vv = f" · 볼륨 {(prev['vol'] / first['vol'] - 1) * 100:+.0f}%" if first["vol"] and prev["vol"] else ""
            print(f"  → 첫 기록 대비 최고 무게 {_arrow(prev['top'], first['top'], 'kg')}{vv} ({len(sess)}회)")


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
