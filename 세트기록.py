"""세트 기록 — 폰 앱에서 복사한 텍스트를 세트기록.csv에 합치고, 일차 · 종목별 흐름을 보여 준다.

2026-11-13 사용자 요청: "2일차 관련 데이터 빼내 줘" 하면 그 일차의 모든 기록을 대화창에 흐름으로.
폰 기록은 폰 브라우저(localStorage wk-sets)에만 있다 — 앱 일차 화면의 'PC로 보낼 세트 기록 복사'를
Claude에게 붙여 넣으면 add로 여기 쌓는다. 같은 날짜 · 종목 · 세트는 새 값으로 덮는다.

  python 세트기록.py add 붙여넣은.txt      (또는 표준 입력)
  python 세트기록.py show --day 2          (2일차 전부 · 종목별 사이클 흐름)
  python 세트기록.py show --ex 벤치         (이름에 '벤치'가 든 종목 · 모든 일차)
  python 세트기록.py show --day 2 --ex 플라이
"""
import argparse
import csv
import sys
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


def parse_paste(text):
    out = []
    for line in text.splitlines():
        line = line.rstrip("\r")
        if not line.strip() or line.startswith("#") or line.startswith("날짜\t"):
            continue
        p = line.split("\t")
        if len(p) < 8:
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


def fmt_set(r):
    w = r["무게"] or "?"
    reps = r["횟수"] or ("목표" if r["달성"] == "O" else "?")
    return f"{w}×{reps}" + ("" if r["달성"] == "O" else "(미체크)" if r["횟수"] else "")


def cmd_show(a):
    rows = [r for r in load()
            if (not a.day or r["일차"] == a.day) and (not a.ex or a.ex in r["종목"])]
    if not rows:
        print("해당 기록 없음" + ("" if CSV.exists() else " — 세트기록.csv가 아직 없다(앱에서 복사해 붙여 넣기부터)"))
        return
    by_ex = OrderedDict()
    for r in sorted(rows, key=lambda r: (r["일차"], r["날짜"], int(r["세트"] or 0))):
        by_ex.setdefault((r["일차"], r["종목"]), OrderedDict()).setdefault(
            (r["날짜"], r["사이클"]), []).append(r)
    for (day, ex), sess in by_ex.items():
        print(f"\n■ {day}일차 · {ex}")
        prev = None
        for (d, cyc), sets in sess.items():
            vol = sum(float(s["무게"]) * float(s["횟수"]) for s in sets
                      if s["무게"].replace(".", "", 1).isdigit() and s["횟수"].isdigit())
            top = max((float(s["무게"]) for s in sets if s["무게"].replace(".", "", 1).isdigit()), default=None)
            trend = ""
            if top is not None and prev is not None:
                dd = top - prev
                trend = f"  ({'+' if dd > 0 else ''}{dd:g}kg)" if dd else "  (=)"
            if top is not None:
                prev = top
            line = " · ".join(fmt_set(s) for s in sets)
            print(f"  {d} [{cyc or '-'}사이클]  {line}" + (f"  | 볼륨 {vol:g}" if vol else "") + trend)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    sub = ap.add_subparsers(dest="cmd", required=True)
    p = sub.add_parser("add")
    p.add_argument("file", nargs="?")
    p.set_defaults(fn=cmd_add)
    p = sub.add_parser("show")
    p.add_argument("--day")
    p.add_argument("--ex")
    p.set_defaults(fn=cmd_show)
    a = ap.parse_args()
    a.fn(a)
