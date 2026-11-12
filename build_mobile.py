#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
운동프로그램 워크북 → 모바일 HTML 변환기

사용법:
    python build_mobile.py                       # 폴더에서 가장 최근 워크북을 자동으로 찾음
    python build_mobile.py 운동프로그램_재편_v14.xlsx
    python build_mobile.py 파일.xlsx -o 어디에저장.html

이 스크립트와 워크북을 같은 폴더에 두고 인자 없이 실행하는 것이 가장 편합니다.
엑셀을 고친 뒤 이 명령만 다시 돌리면 모바일 파일이 갱신됩니다.
종목 추가 · 삭제 · 세트 변경 · 메모 수정 전부 자동으로 따라옵니다.

필요 패키지: openpyxl  (pip install openpyxl)

주의 — 엑셀에서 수정한 뒤에는 반드시 엑셀로 한 번 저장해야 합니다.
      openpyxl은 수식의 '계산된 값'을 읽는데, 그 값은 엑셀이 저장할 때 기록됩니다.
"""

import argparse, datetime, json, re, sys
from pathlib import Path

import voice_build

try:
    import openpyxl
except ImportError:
    sys.exit("openpyxl이 필요합니다:  pip install openpyxl")

# 콘솔 코드페이지가 cp949면 '—' 같은 문자에서 출력이 터진다.
# HTML은 이미 쓰인 뒤라 결과물은 멀쩡한데 화면엔 traceback만 남으므로 미리 막는다.
for _stream in (sys.stdout, sys.stderr):
    try:
        _stream.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass


# ── 시트 이름 규약 ────────────────────────────────────────────────
# 재택 시트는 'D<일차>_<내용>' (8일 주기) 또는 '<요일>_<내용>' (구 주 단위) 둘 다 잡는다.
# 2026-08-28에 8일 주기로 옮기며 D1~D8 형태가 현행이 됐고, 요일 형태는 옛 워크북 호환용이다.
WEEKDAYS = "월화수목금토일"
DAY_SHEET = re.compile(r"^(?:D(\d+)|([월화수목금토일]))_")
# 복귀판은 주 6회 요일 고정 (2026-10-02 개편). 현장 근무가 끝나 요일을 박을 수 있게 됐고,
# 큰 부위 셋을 둘로 쪼개 가슴 · 등 · 하체가 주 2회가 됐다. 금요일만 쉰다.
# 등은 재택과 같이 움직임 평면으로 가른다 — 일요일 수평 · 목요일 수직.
# 무거운 로우를 데드(수) 다음날에서 떼어 내려고 재택과 순서를 반대로 뒀다.
# 순서는 토요일부터다 — 사용자가 주의 출발을 토요일로 잡았다.
# 2026-11-01 v22 — 요일 고정에서 9일 사이클(7세션 + 휴식 2)로. 라벨은 일차 번호다.
# 중량 · 주차는 그대로 날짜(토요일 넘김)로 세고, 오늘 몇 일차인지만 사이클 기준일에서 센다.
# 시트 이름이 '사이클_<일차>_<내용>'이면 전부 잡는다 — 6 · 9일차는 시트가 없어 휴식이다.
CYCLE_SHEET = re.compile(r"^사이클_(\d+)_")
CYCLE_LEN = 9


def cycle_days(wb):
    found = []
    for name in wb.sheetnames:
        m = CYCLE_SHEET.match(name)
        if m:
            found.append((m.group(1), name))
    return sorted(found, key=lambda p: int(p[0]))


def is_back(ws):
    return ws.title.startswith(("복귀", "사이클_", "분할_"))


# 2026-11-01 — 앱 왼쪽 버튼(옛 재택 D1~D9)을 주 6일 분할(분할_<요일>_ 시트 · 일요일 휴식)로 바꿨다.
# 재택 시트는 워크북에 남아 있지만 앱에는 안 싣는다(사용자 결정)
SPLIT_SHEET = re.compile(r"^분할_([월화수목금토일])_")


def split_days(wb):
    found = []
    for name in wb.sheetnames:
        m = SPLIT_SHEET.match(name)
        if m:
            found.append((m.group(1), name))
    return sorted(found, key=lambda p: WEEKDAYS.index(p[0]))


def home_days(wb):
    """재택 시트를 [(라벨, 시트명)] 로. '복귀_'는 앞이 D숫자도 요일도 아니라 걸리지 않는다.

    8일 주기 시트('D5_당기기B')는 라벨이 '5', 구 요일 시트('목_당기기B')는 '목'이다.
    정렬 키만 갈라두면 나머지 코드는 라벨을 문자열로만 다루므로 손댈 곳이 없다."""
    found = []
    for name in wb.sheetnames:
        m = DAY_SHEET.match(name)
        if m:
            found.append((m.group(1) or m.group(2), name))
    return sorted(found, key=lambda p: int(p[0]) if p[0].isdigit()
                  else WEEKDAYS.index(p[0]))

# '메인:' 줄에 적히는 이름 → 증량을 추적하는 리프트.
# 복귀판은 같은 리프트를 다른 기구로 친다 — 수요일 시티드 바벨 밀리터리 프레스가 밀리터리 자리다
# (2026-10-07 스미스 오버헤드 프레스에서 교체 — 옛 이름은 대체 종목으로 돌아갈 때를 위해 남긴다).
LIFT_KEYS = {"스쿼트": "squat", "데드리프트": "dead",
             "벤치프레스": "bench", "밀리터리 프레스": "press",
             "시티드 바벨 밀리터리 프레스": "press",
             "오버헤드 프레스 (스탠딩 바벨)": "press",   # 2026-11-05 시티드 → 스탠딩
             "스미스 오버헤드 프레스": "press"}

# 볼륨 램프를 중량보다 몇 주 앞세울지.
# 2026-08-10 결정 — 중량은 2주차 그대로 두고 볼륨만 3주차(138세트)부터 간다.
# 0으로 되돌리면 볼륨과 중량이 같은 주차를 쓴다.
RAMP_WEEK_OFFSET = 1

# head — 증량기록 표 머리글. 그 오른쪽 '미달 시 반복수' 열이 그 리프트의 기록이다
LIFT_META = {                      # 시작중량_설정 행 / 증량기록 행
    "squat": {"ko": "스쿼트",        "start_row": 7,  "inc_row": 6, "accent": "var(--sq)", "head": "스쿼트 목표"},
    "dead":  {"ko": "데드리프트",     "start_row": 8,  "inc_row": 5, "accent": "var(--dl)", "head": "데드 목표"},
    "bench": {"ko": "벤치프레스",     "start_row": 9,  "inc_row": 7, "accent": "var(--bp)", "head": "벤치 목표"},
    "press": {"ko": "밀리터리 프레스", "start_row": 10, "inc_row": 8, "accent": "var(--mp)", "head": "프레스 목표"},
    # 2026-10-23 사용자 결정 — 보조 종목 둘을 매주 +1kg 선형으로. 원판식 기구라 1kg 격자가 그대로 맞는다.
    # 메인 카드(웜업 램프 · 봉 원판 그림)는 없고 종목표의 그 줄 중량만 주차를 따라간다 — acc
    # 2026-10-23 같은 날 둘째 묶음 — 주기는 종목마다(E열 · 1~3주). 스미스 · 머신은 원판만 센 무게다
    # (봉 무게를 모른다 · 원판만 세면 '5로 나눈 나머지 4' 규칙이 봉과 상관없이 맞는다)
}
_ACC = [  # 키, 앱 이름, 요일, 종목표 이름 — 시작중량_설정 11행 · 증량기록 9행부터 이 순서
    ("row",   "해머 로우",          "일", "해머 스트랭스 로우"),   # 2026-11-04 '(오버그립)' 뗌
    ("incline", "인클라인 머신",     "월", "인클라인 벤치프레스 머신"),
    ("lp",    "레그 프레스",        "토", "레그 프레스"),
    ("lph",   "레그 프레스 발 높게", "수", "레그 프레스 (발 높게)"),
    ("rdl",   "루마니안",           "수", "루마니안 데드리프트"),
    ("hip",   "힙 쓰러스트",        "화", "힙 쓰러스트"),
    ("ohs",   "오버헤드 슈러그",     "목", "오버헤드 슈러그(스미스)"),
    ("scap",  "스캡 슈러그",        "일", "해머 스트랭스 스캡 슈러그(중립 그립)"),
    ("prot",  "프로트랙션",         "월", "스미스 프로트랙션"),
    # 셋째 묶음 — 봉만 있는 EZ바 + 원판으로 바꾼다. 봉 무게를 몰라 봉 포함 총무게(지금 15)로 적는다
    ("lying", "라잉 트라이셉스",     "화", "라잉 트라이셉스 익스텐션"),
    ("rcurl", "리버스컬",           "화", "리버스컬"),
    ("incb",  "인클라인 바벨",       "목", "인클라인 벤치프레스(바벨)"),
    # 2026-11-07 사용자 결정 — 4일차 핵 스쿼트 3×10 · 매 칸 +2 · 14칸 80(원판만). 시작중량_설정 23행 · 증량기록 설정 21행(표는 22행 머리글로 내려감)
    ("hack",  "핵 스쿼트",          "4일차", "핵 스쿼트"),
]
#  추적을 멈춘 종목 — 행 번호가 순서로 정해지니 목록에선 빼지 않고 여기서 건너뛴다
#  prot: 2026-10-26 월 스미스 프로트랙션 → 체스트 프레스 머신 프로트랙션(토요일과 같은 범위 + 핀 방식)
#  incline: 2026-11-09 8일차 인클라인 벤치프레스 머신 → 인클라인 덤벨 프레스(머신은 늘어난 위치가 안 걸림 · 사용자) — 덤벨 더블 프로그레션
_RETIRED = {"prot", "incline"}
# 2026-11-10 사용자 결정 — 핵 스쿼트 2사이클 쉬고(왼 슬개골 아래 · 그동안 4일차 1번 = 레그 익스텐션 등척 홀드)
# 사이클 4(17주차)에 60에서 다시 — 회차마다 +5 · 80에 닿으면 원래 +2(못 끼우는 값 +1). 쉬는 두 칸(15 · 16주차)은 80 그대로.
# 엑셀 증량기록 AL32~AL52가 같은 식이다. 자리 바꾸기는 사이클_종목교대 '4사이클부터' 줄
RESTART = {"hack": {"hold": 15, "at": 17, "base": 60, "inc": 5, "cap": 80}}
for _i, (_k, _ko, _d, _n) in enumerate(_ACC):
    if _k in _RETIRED:
        continue
    LIFT_META[_k] = {"ko": _ko, "start_row": 11 + _i, "inc_row": 9 + _i, "accent": "var(--ink)",
                     "head": _ko + " 목표", "acc": True, "day": _d}

# 비율 종목 — 무게 = 그 리프트 목표 × 증량기록 C열 비율 (2026-11-09 인클라인 바벨 = 벤치 × 0.8)
RATIO_OF = {"incb": "bench"}

# 추적 보조 종목 — 종목표의 이 이름 줄이 그 리프트다('메인:' 줄 없이)
TRACK_KEYS = {_n: _k for _k, _ko, _d, _n in _ACC if _k not in _RETIRED}

# 2026-10-24 사용자 결정 — 케이블 · 레터럴은 반복 범위 그대로 두고 칸만 쪽낸다(선형 아님).
# 전 세트 상한이면 앱 버튼으로 한 단계 — 무게 칸 글자 자체가 단계다('55 + 핀 + 1').
# pin   = 스택(케이블 · 머신): 칸 → 칸 + 핀 → + 핀 + 1 → + 핀 + 2.5 → 핀 한 칸 내리고 원판 뺌
# wrist = 덤벨: 덤벨 → + 손목 1 → 다음 덤벨 (모래주머니 한쪽 1kg 고정)
# assist = 어시스트 머신: 스택이 도와주는 무게라 거꾸로 간다 — 한 칸 가벼운 칸에 핀으로 얹었다가 덜어 낸다
#          45 → 40 + 핀 + 2.5 → 40 + 핀 + 1 → 40 + 핀 → 40 → …
# bar    = 봉 + 원판: +1kg (원판 격자 — 5로 나눈 나머지 4는 못 끼워 하나 더) · kg = +1kg · half = +0.5kg
MICRO = {
    "체스트 프레스 머신 프로트랙션": "pin", "시티드 케이블 로우": "pin", "원암 케이블 로우": "pin",
    "케이블 오버헤드 익스텐션 (로프)": "pin", "원암 케이블 레터럴 레이즈": "pin",
    "원암 케이블 리버스 플라이": "pin", "레터럴 레이즈": "pin",      # 월 7번은 머신
    "케이블 힙 외전": "pin", "랫풀다운 (와이드 오버그립)": "pin",
    "랫풀다운(언더그립 좁게 안쪽으로)": "pin", "디클라인 케이블 플라이": "pin",
    "암풀다운": "pin", "페이스풀(로프)": "pin", "로프 푸쉬다운": "pin",
    "케이블 드래그 컬": "pin", "벤트오버 레터럴 레이즈 머신": "pin",
    "시티드 레터럴 레이즈": "wrist",
    "케이블 외회전 (앉아서 · 팔 옆구리)": "pin", "케이블 외회전 (팔 90°)": "pin",   # 2026-11-05 회전근개 3 · 7일차
    "펙덱 플라이": "pin",
    "리버스 펙덱": "pin", "레그 컬": "pin", "원 레그 라잉 컬": "pin",   # 2026-10-27 토 시티드 → 라잉 복귀
    "레그 익스텐션": "pin", "머신 숄더 프레스": "pin", "프리쳐 컬": "pin",
    "레그 익스텐션 등척 홀드 (무릎 60°)": "pin",   # 2026-11-10 4일차 1번(핵 스쿼트 대신 · 일시)
    "인클라인 덤벨 컬": "wrist", "해머 컬": "wrist",
    "인클라인 플라이 머신": "pin",   # 2026-10-26 월 3번 덤벨 → 머신(옆 센터)
    "인클라인 덤벨 프로트랙션": "wrist", "덤벨 숄더 프레스": "wrist",
    "인클라인 덤벨 로우(오버그립)": "wrist",
    "딥스": "assist",
    "친업 (어시스트 가능)": "assist",
    "와이드 그립 업라이트 로우": "bar", "백 익스텐션 (45도)": "kg",
    "하이로우 머신(언더그립)": "kg",   # 원판식 머신 (2026-10-25 사용자 확인)
    # 리스트 컬은 원판 끼우는 봉으로 바꾸며 15~20 (봉 무게 모름 — 봉 포함 총무게). 넥은 2.5에서 +1이 40%라 0.5씩
    "시티드 리스트 컬(손등 아래)": "kg", "시티드 리스트 컬(손등 위)": "kg",
    "넥 익스텐션 (플레이트)": "half", "넥 플렉션 (플레이트)": "half",
    # 2026-11-01 v22 9일 사이클에서 새로 생긴 이름
    "원암 오버헤드 케이블 익스텐션": "pin", "페이스풀 (로프)": "pin", "로우 투 하이 케이블 플라이": "pin",
    "시티드 레그 컬": "pin", "랫풀다운 (뉴트럴 그립)": "pin", "랫풀다운(언더그립)": "pin",
    "원암 로우 케이블 컬": "pin", "스탠딩 케이블 크런치": "pin",
    "업라이트 로우 (EZ바 · 넓은 그립)": "bar", "푸쉬다운 (일자 바)": "pin",
    # 2026-11-04 등 그립 회전 시작 — 이름에서 그립 · 원암 표기를 뗀 새 이름(3 · 5 · 8일차 랫풀다운은 한 이름)
    "랫풀다운": "pin", "케이블 로우": "pin", "인클라인 덤벨 로우": "wrist", "하이로우 머신": "kg",
    "인클라인 덤벨 프레스": "wrist",   # 2026-11-09 8일차 4번(머신 대체)
    "케이블 플라이 (가슴 높이)": "pin",   # 2026-11-12 2일차 FST-7(머신 체스트 프레스 자리)
    "인클라인 사이드 덤벨 레터럴 레이즈": "wrist",   # 2026-11-12 8일차 9번(레터럴 레이즈 대체)
}

# 2026-10-24 사용자 결정 — 예열 1세트(본세트 50% × 6~8 · 세트 수에 안 셈). 그날 처음 쓰는 관절 ·
# 늘어난 자세 종목 · 탈 났던 곳만. 메인 넷은 웜업 램프가 따로 있어 뺀다. (요일, 종목) — 같은 이름이 다른 날에도 있다
# 2026-11-01 v22 9일 사이클로 옮김 — {(일차, 종목): 예열 방법}. 인클라인 바벨은 8일차 웜업 램프가 따로 있어 뺐다.
# 같은 날 2026-11-01 사용자 결정 — '갑자기 다른 부위로 넘어가는 자리'도 예열. 방법이 None이면 기본(50% × 6~8)이다.
# 방법 글자는 종목 카드 · 타이머 띠 · '다음 ▸' 줄에 그대로 뜬다(음성은 '예열 한 세트 먼저'만)
WU_DEFAULT = "50% × 6~8"
WARMUP = {
    ("1", "레그 컬"): None,                                       # 패드로 늘어난 자세 · 첫 무릎 굽힘
    ("1", "넥 익스텐션 (플레이트)"): "무게 없이 머리만 앞뒤로 × 15씩 — 그날 첫 목 · 가동범위부터",
    ("2", "딥스"): "어시스트 많이 걸고 × 8 · 맨 아래 깊이를 천천히 — 어깨 늘어나는 구간 확인",
    ("3", "해머 스트랭스 로우"): None,
    ("4", "힙 쓰러스트"): None,
    # 2026-11-12 사용자 — 4일차 라잉 ↔ 7일차 오버헤드 자리 바꿈(4일차 오버헤드 ↔ 로프 푸쉬다운 FST-7 7라운드 · 7일차 라잉 단독)
    ("4", "케이블 오버헤드 익스텐션 (로프)"): "로프 오버헤드 40% × 12~15 — 하체 뒤 차가운 팔꿈치 · 푸쉬다운은 따로 안 함",
    ("4", "페이스풀 (로프)"): "아주 가볍게 × 15 — 하체 뒤 그날 첫 어깨",
    # 2026-11-11 랫풀다운 → 원암 덤벨 로우(사용자) — 여전히 그날 첫 당기기 · 오른쪽 대원근 땡김 이력
    ("5", "원암 덤벨 로우"): "50% × 8 (좌우) — 그날 첫 당기기 · 맨 아래 늘릴 때 오른쪽 겨드랑이 뒤(대원근) 확인",
    ("5", "인클라인 덤벨 컬"): None,                              # 어깨 뒤 늘어난 자세 · 오른쪽 앞 어깨
    ("7", "루마니안 데드리프트"): None,                           # 햄스트링 근위부 이력
    ("7", "라잉 트라이셉스 익스텐션"): None,                     # 하체 70분 뒤 차가운 팔꿈치
    ("8", "업라이트 로우 (EZ바 · 넓은 그립)"): None, ("8", "인클라인 덤벨 컬"): None,  # 오른쪽 앞 어깨 이력
}


# 2026-11-10 사용자 요청 — 무릎(왼 슬개골 아래) 부하 종목은 세션 타이머에 노랑 경고 띠 — 그 종목 하기 전 휴식 · 세트 · 세트 사이 휴식 동안.
# 카드에도 같은 글자. 기준은 통증 3/10 · 다음 날 아침 더 나쁘지 않을 것 (일차, 종목) — 교대표 종목(4일차 핵 스쿼트)도 여기서 찾는다
_KN = "무릎 · 통증 3/10 이하 · 다음 날 아침 확인"
KNEE = {
    ("1", "스쿼트"): _KN + " · 넘으면 다음 증량 미루고 같은 무게 한 번 더",
    ("1", "레그 프레스"): _KN + " · 아프면 깊이 → 무게 → 내리는 속도 순으로",
    ("1", "불가리안 스플릿 스쿼트 (상체 숙여서)"): _KN + " · 왼다리 앞일 때 무릎 안쪽 모임 · 아프면 보폭 넓게",
    ("1", "레그 익스텐션"): _KN + " · 3초 내리기 · 맨 위 잠그지 않기",
    ("4", "레그 익스텐션 등척 홀드 (무릎 60°)"): "무릎 · 버티는 동안 통증 3/10 이하 · 잠그지 않기",
    # 2026-11-10 사용자 — 11/5 80kg 때 뒤꿈치가 거의 뜨고 앞꿈치로 밀었다(9/4 스쿼트와 같은 '자세 무너짐') → 뒤꿈치가 실패 기준
    ("4", "핵 스쿼트"): "무릎 · 뒤꿈치 뜨기 시작하면 그 세트 끝 · 발 반 발 위 · 통증 3/10 이하 · 다음 날 아침 확인",
    ("4", "레그 프레스 (발 높게)"): _KN,
    ("7", "레그 익스텐션"): _KN + " · 3초 내리기 · 맨 위 잠그지 않기",
}


# 2026-11-07 사용자 요청 — 세션 시작 전 웜업 타이머(세션 타이머와 따로 · 그 일차 화면 안 버튼).
# 엑셀 웜업 칸(종목표 위 '·' 줄)의 순서를 그대로 옮긴 것이다 — 엑셀 웜업 줄을 바꾸면 여기도 같이.
# 수행 · 휴식 초는 처음 정한 값이라 실제와 다르면 고친다(블랙번 300초 = 사용자 실측 약 5분 · 2026-11-07).
#   {"n","w","r","s","wk","rt"} = 일반 항목 · {"ramp": 1} = 그날 메인 리프트 램프(앱 카드와 같은 ramp())
#   {"ss": [항목, 항목], "s": 라운드} = 슈퍼세트(항목의 "of" 종목 무게 × "pct" · rt는 그 항목 뒤 쉼)
#   {"pctOf": 종목 이름 또는 "#2"(2번 행 — 8일차 교대 종목), "rows": [[비율|"bar", 반복, 휴식초], …]}
_SF = {"n": "숏풋 (왼발 위주)", "mg": "왼발 아치 · 5초 버티기", "w": "맨발", "r": "10", "s": 1, "wk": 60, "rt": 10}
_SOL = {"n": "가자미근 스트레칭 (왼쪽 · 무릎 굽혀)", "mg": "종아리 · 45초 버티기", "w": "벽", "r": "", "s": 2, "wk": 45, "rt": 15}
_ER = {"n": "케이블 외회전", "mg": "회전근개", "w": "매우 가볍게", "r": "15 (좌우)", "s": 2, "wk": 60, "rt": 20}
_BB = {"n": "블랙번", "mg": "하부 승모 · 회전근개 · 한계 안 감", "w": "1", "r": "", "s": 1, "wk": 300, "rt": 30}
# 2026-11-08 사용자 — 하체 날(1 · 4 · 7일차) 고관절 동적 2종. 정적 버티기 · 깊은 햄스트링 숙이기는 뺌(스쿼트 · 데드 전 · 햄스트링 근위부 이력)
_HIP = [{"n": "90/90 고관절 전환", "mg": "고관절 내·외회전 · 버티지 않음", "w": "맨몸", "r": "5 (좌우)", "s": 1, "wk": 50, "rt": 10},
        {"n": "하프 닐링 장요근 흔들기", "mg": "장요근 · 엉덩이 조이고 1~2초씩", "w": "맨몸", "r": "8 (좌우)", "s": 1, "wk": 50, "rt": 15}]
_LAT = lambda n:{"n": "랫풀다운 (가볍게)", "mg": "광배", "w": "가볍게", "r": "10~12", "s": n, "wk": 35, "rt": 45}
WU_SESSION = {
    # 2026-11-07 사용자 — 스쿼트 전 레그 익스텐션 ↔ 라잉 레그 컬 가볍게 2라운드(쉬지 않고 번갈아 · 본세트 무게의 50% — 35%는 너무 가볍다고 사용자 · 5회 이상 남김)
    "1": [_SF, _SOL, *_HIP,
          {"ss": [{"of": "레그 익스텐션", "n": "레그 익스텐션 (가볍게)", "mg": "대퇴사두 · 5회 이상 남김", "pct": 0.5, "r": "12~15", "wk": 40, "rt": 10},
                  {"of": "레그 컬", "n": "라잉 레그 컬 (가볍게)", "mg": "햄스트링 · 5회 이상 남김", "pct": 0.5, "r": "12~15", "wk": 40, "rt": 20}], "s": 2},
          {"ramp": 1}],
    # 2026-11-07 사용자 — 벤치 램프 앞 머신 체스트 프레스 가볍게 2세트(본세트 무게의 50% · 고정 궤도로 오른쪽 앞 어깨 먼저 확인 · 5회 이상 남김)
    # 2026-11-12 사용자 — 2일차 머신 체스트 프레스 본세트가 케이블 플라이 FST-7로 바뀌어 예열은 옛 본세트 32 × 50% = 16 고정(2 · 8일차)
    "2": [_ER, _BB,
          {"ss": [{"w": "16", "n": "머신 체스트 프레스 (가볍게)", "mg": "흉근 · 5회 이상 남김", "pct": 0.5, "r": "12~15", "wk": 40, "rt": 30}], "s": 2},
          {"ramp": 1}],
    "3": [_BB, _LAT(3)],
    "4": [{"n": "글루트 브릿지 (맨몸)", "mg": "둔근", "w": "맨몸", "r": "15", "s": 1, "wk": 45, "rt": 15},
          {"n": "왼발목 무릎 벽 밀기", "mg": "발목", "w": "맨발", "r": "10~15", "s": 1, "wk": 60, "rt": 15},
          *_HIP,
          # 2026-11-07 사용자 — 레그 익스텐션 가볍게 2세트(4일차엔 그 종목이 없어 1일차 무게 칸 × 50%) + 빈 썰매 핵 스쿼트 × 10(발 위치 · 깊이 확인)
          # 2026-11-10 사용자 — 사이클 1~3은 4일차 1번이 등척 홀드라 빈 썰매 · 핵 램프를 건너뛴다("cf": 그 사이클부터 나옴)
          {"ss": [{"of": "레그 익스텐션", "ofDay": "1", "n": "레그 익스텐션 (가볍게)", "mg": "대퇴사두 · 5회 이상 남김", "pct": 0.5, "r": "12~15", "wk": 40, "rt": 30}], "s": 2},
          {"n": "핵 스쿼트 (빈 썰매)", "mg": "발 반 발 위 · 왼 앞꿈치 쏠림 확인", "w": "원판 없이", "r": "10", "s": 1, "wk": 35, "rt": 45, "cf": 4},
          {"pctOf": "핵 스쿼트", "rows": [[0.5, 8, 60], [0.75, 4, 90]], "cf": 4}],
    "5": [_ER, _BB],
    # 2026-11-07 사용자 — 데드 램프 앞 라잉 레그 컬 가볍게 2세트(양측 · 1일차 레그 컬 무게 칸 × 50% · 햄스트링 근위부 이력)
    "7": [_SF, _SOL, *_HIP,
          {"ss": [{"of": "레그 컬", "ofDay": "1", "n": "라잉 레그 컬 (가볍게)", "mg": "햄스트링 · 5회 이상 남김", "pct": 0.5, "r": "12~15", "wk": 40, "rt": 30}], "s": 2},
          {"ramp": 1}],
    # 2026-11-07 사용자 — 8일차에도 바벨 램프 앞 머신 체스트 프레스 가볍게 2세트. 8일차엔 그 종목이 없어 2일차 무게 칸을 쓴다(ofDay)
    "8": [_BB, _LAT(2),
          {"ss": [{"w": "16", "n": "머신 체스트 프레스 (가볍게)", "mg": "흉근 · 5회 이상 남김", "pct": 0.5, "r": "12~15", "wk": 40, "rt": 30}], "s": 2},
          {"pctOf": "#2", "rows": [["bar", 5, 45], [0.6, 5, 60], [0.75, 3, 90], [0.9, 1, 120]]}],
}
# 세션 안에 있는 램프 — 5일차 프레스 램프는 원암 레터럴 뒤 · 프레스 직전이라 세션 타이머 쪽에 끼운다
WU_INSESSION = {
    "5": {"before": "오버헤드 프레스 (스탠딩 바벨)",
          # 2026-11-07 사용자 요청 — 빈 봉과 본세트 사이에 한 세트(본세트 × 75% · 1kg 격자 · 20 이하면 건너뜀) — 프레스가 오르면 같이 오른다
          "items": [{"n": "프레스 램프", "mg": "웜업 램프", "w": "20", "r": "5", "s": 2, "wk": 25, "rt": 60},
                    {"n": "프레스 램프 (75%)", "mg": "웜업 램프", "pct": 0.75, "r": "3", "s": 1, "wk": 15, "rt": 90}]},
}


def pick_latest(folder):
    """폴더에서 가장 최근에 수정된 워크북을 고른다. 엑셀 임시파일(~$)은 무시."""
    # 2026-11-13 같은 폴더에 세트기록_보기.xlsx(세트기록.py가 세션마다 새로 씀)가 생겼다 — 프로그램 워크북만
    cands = [p for p in Path(folder).glob("운동프로그램*.xlsx") if not p.name.startswith("~$")]
    if not cands:
        sys.exit(f"{folder} 안에 .xlsx 파일이 없습니다. 워크북을 같은 폴더에 두세요.")
    return max(cands, key=lambda p: p.stat().st_mtime)


def s(v):
    """셀 값을 문자열로. None은 빈 문자열."""
    if v is None:
        return ""
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


# ── 홈 화면 앱(PWA) 자산 ─────────────────────────────────────────
# manifest가 있어야 폰이 이걸 웹페이지가 아니라 앱으로 취급한다.
# 없으면 홈 화면에서 열어도 위에 주소창이 남는다.

BG = (0x15, 0x18, 0x1C)      # --bg   본문 배경과 동일
FG = (0xDD, 0xA5, 0x1B)      # --bp   벤치프레스 노랑. 어두운 배경에서 잘 보인다


def _png(size, px):
    """RGB 픽셀 함수를 8bit PNG 바이트로. 표준 zlib만 사용."""
    import zlib, struct
    raw = bytearray()
    for y in range(size):
        raw.append(0)                       # 필터 타입 None
        for x in range(size):
            raw.extend(px(x, y))

    def chunk(tag, data):
        return (struct.pack(">I", len(data)) + tag + data
                + struct.pack(">I", zlib.crc32(tag + data) & 0xFFFFFFFF))

    return (b"\x89PNG\r\n\x1a\n"
            + chunk(b"IHDR", struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(bytes(raw), 9))
            + chunk(b"IEND", b""))


def barbell_icon(size):
    """바벨 아이콘. 안드로이드가 원형으로 잘라내도 남도록 중앙 80% 안에 그린다."""
    def rect(x0, y0, x1, y1):
        return (x0 * size, y0 * size, x1 * size, y1 * size)

    shapes = [
        rect(0.16, 0.470, 0.84, 0.530),   # 바
        rect(0.26, 0.330, 0.33, 0.670),   # 안쪽 원판 (좌)
        rect(0.67, 0.330, 0.74, 0.670),   # 안쪽 원판 (우)
        rect(0.20, 0.390, 0.25, 0.610),   # 바깥 원판 (좌)
        rect(0.75, 0.390, 0.80, 0.610),   # 바깥 원판 (우)
    ]

    def px(x, y):
        for x0, y0, x1, y1 in shapes:
            if x0 <= x < x1 and y0 <= y < y1:
                return FG
        return BG

    return _png(size, px)


# Chrome은 manifest만으로는 '홈 화면에 추가'(단순 바로가기)까지만 준다.
# '앱 설치'로 뜨게 하려면 service worker가 있어야 하고, 덤으로 오프라인에서도 열린다.
# 네트워크 우선 — 신호가 있으면 항상 최신을 보고, 끊기면 캐시로 떨어진다.
# 워크북을 고칠 때마다 VERSION이 바뀌어 옛 캐시가 폐기되므로 낡은 표가 남지 않는다.
SW_JS = """\
const VERSION = "__VERSION__";
const CACHE = "workout-" + VERSION;
const ASSETS = ["./", "./index.html", "./manifest.webmanifest",
                "./icon-192.png", "./icon-512.png"];

/* 말소리는 몇 MB짜리 한 덩어리이고, 워크북을 고쳐도 하는 말은 거의 그대로다.
   버전 캐시에 같이 넣으면 표 한 줄 고칠 때마다 폰이 그걸 통째로 다시 받는다.
   그래서 파일 이름에 내용 해시를 박아 따로 보관한다 — 하는 말이 바뀔 때만
   이름이 바뀌고, 그때만 새로 받는다. */
const VOICE_CACHE = "workout-voice";
const VOICE = "__VOICE__";

self.addEventListener("install", e => {
  e.waitUntil(Promise.all([
    caches.open(CACHE).then(c => c.addAll(ASSETS)),
    VOICE ? caches.open(VOICE_CACHE)
              .then(c => c.match(VOICE).then(hit => hit || c.add(VOICE)))
              .catch(() => null)
          : null
  ]).then(() => self.skipWaiting()));
});

self.addEventListener("activate", e => {
  e.waitUntil(caches.keys()
    .then(ks => Promise.all(ks.filter(k => k !== CACHE && k !== VOICE_CACHE)
                              .map(k => caches.delete(k))))
    .then(() => caches.open(VOICE_CACHE))
    .then(c => c.keys().then(rs => Promise.all(      /* 지난 말소리는 버린다 */
      rs.filter(r => !VOICE || !r.url.endsWith(VOICE.slice(2)))
        .map(r => c.delete(r)))))
    .then(() => self.clients.claim()));
});

self.addEventListener("fetch", e => {
  if (e.request.method !== "GET") return;
  /* 말소리만은 캐시 우선이다. 파일 이름이 곧 내용이라 한 번 받으면 다시 받을
     이유가 없고, 네트워크 우선으로 두면 앱을 열 때마다 몇 MB가 샌다. */
  if (e.request.url.indexOf("/voice/") >= 0) {
    e.respondWith(caches.open(VOICE_CACHE).then(c =>
      c.match(e.request).then(hit => hit || fetch(e.request).then(res => {
        c.put(e.request, res.clone());
        return res;
      }))));
    return;
  }
  e.respondWith(
    fetch(e.request).then(res => {
      const copy = res.clone();
      caches.open(CACHE).then(c => c.put(e.request, copy));
      return res;
    }).catch(() => caches.match(e.request).then(r => r || caches.match("./index.html")))
  );
});
"""


def write_pwa_assets(site, html, voice_src):
    """manifest·아이콘·service worker를 배포 폴더에 쓴다. 아이콘은 내용이 고정이라 없을 때만 만든다."""
    manifest = {
        "name": "운동 프로그램",
        "short_name": "운동",
        "start_url": "./",
        "scope": "./",
        "display": "standalone",     # 주소창 없이 전체화면
        "orientation": "portrait",
        "background_color": "#15181C",
        "theme_color": "#15181C",
        "icons": [
            {"src": "icon-192.png", "sizes": "192x192", "type": "image/png",
             "purpose": "any maskable"},
            {"src": "icon-512.png", "sizes": "512x512", "type": "image/png",
             "purpose": "any maskable"},
        ],
    }
    (site / "manifest.webmanifest").write_text(
        json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")

    made = []
    for px_size in (192, 512):
        p = site / f"icon-{px_size}.png"
        if not p.exists():
            p.write_bytes(barbell_icon(px_size))
            made.append(p.name)

    import hashlib
    version = hashlib.sha1(html.encode("utf-8")).hexdigest()[:12]
    (site / "sw.js").write_text(
        SW_JS.replace("__VERSION__", version)
             .replace("__VOICE__", "./" + voice_src if voice_src else ""),
        encoding="utf-8")
    return made, version


def num(v, default=0):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def half(v, default=0):
    """검산 표시용 숫자. 간접은 0.5 단위라 정수로 자르면 11.5가 11이 된다."""
    x = round(num(v, default), 1)
    return int(x) if x == int(x) else x


def rest_ramp(v):
    """세트별 휴식 칸 — "150/170/190/215/240" (초, 세트 순).

    세트가 쌓이면 피로도 쌓인다. 5x5가 무너지는 자리는 늘 4~5세트째이므로
    앞 세트에서 덜어 뒤 세트에 준다. 비어 있으면 None = '휴식(초)' 한 값을
    전 세트에 쓴다(종전 동작).

    H열('휴식(초)')은 이 램프의 평균으로 둔다 — 그래야 I열 '종목 분'
    수식 E*(G+H)/60 이 손대지 않고도 그대로 맞는다."""
    t = s(v)
    if not t:
        return None
    out = [int(round(num(x))) for x in t.split("/") if s(x)]
    return out or None


def is_continuation(t):
    """소제목은 짧고 마침표가 없다. 길거나 문장으로 끝나면 앞 항목의 뒷줄로 본다."""
    return len(t) > 40 or t.endswith((".", "다", "음", "함"))


def is_ramp_cell(t):
    """메인 리프트 웜업 램프의 중량 칸 — 숫자이거나 '생략'. 안내문이 아니다."""
    return t == "생략" or re.fullmatch(r"\d+(\.\d+)?", t) is not None


def find_header(ws, max_row=40):
    """A열이 '#'인 행 = 종목표 머리글 행."""
    for r in range(1, max_row + 1):
        if s(ws.cell(r, 1).value) == "#":
            return r
    return None


# 증량표 칸 → 보조 추적 키. 사이클 시트는 이름이 아니라 D열 수식으로 고른다 —
# 사용자가 숫자로 고정한 칸(1일차 레그 프레스 162 · 3일차 해머 로우 61, 2026-11-01)은 추적하지 않는다
ACC_BY_ROW = {11 + _i: _k for _i, (_k, _ko, _d, _n) in enumerate(_ACC) if _k not in _RETIRED}
ACC_REF = re.compile(r"^=시작중량_설정!\$E\$(\d+)$")


def parse_day(ws, wsf=None):
    """요일 시트 하나를 파싱해 dict로. 재택 · 복귀 양쪽 레이아웃 모두 처리.
    wsf = 같은 시트의 수식 판(있으면 D열 증량표 연결로 보조 추적을 고른다)."""
    day = {"t": "", "sub": "", "ex": [], "groups": None, "total": 0,
           "main": None, "mains": [], "notes": [], "pre": []}

    # 제목 — 'A요일 — 가슴' 에서 뒷부분만
    title = s(ws.cell(1, 1).value)
    day["t"] = title.split("—", 1)[1].strip() if "—" in title else title
    day["t"] = re.sub(r"\s*\(.*?\)\s*$", "", day["t"])

    hdr = find_header(ws)
    if hdr is None:
        return day

    # 메인 리프트 — 'A4: 메인: 벤치프레스'. 종목표 위 어디에 있어도 되고,
    # 복귀 토요일처럼 한 세션에 둘이면 줄을 두 번 쓴다 (웜업 램프 블록마다 하나).
    # [키, 종목표에서 찾을 이름] — 이름이 같은 행이 그 리프트의 본세트 줄이다.
    for r in range(2, hdr):
        cell = s(ws.cell(r, 1).value)
        if not cell.startswith("메인:"):
            continue
        label = cell.split(":", 1)[1].strip()
        key = LIFT_KEYS.get(label)
        if key:
            day["mains"].append([key, label])
    day["main"] = day["mains"][0][0] if day["mains"] else None
    main_of = {label: key for key, label in day["mains"]}

    # 종목표 위쪽 블록 (웜업 · 관절 준비 등) — 세트 카운트 대상이 아닌 안내문
    pre_head, pre_bul = None, []
    SKIP = ("중량(Kg)", "운동 구성")

    def flush():
        """모아둔 블록을 내보낸다.

        불릿(·)이 한 줄도 없으면 제목만 있는 블록이다. v18의
        '웜업 — 스캡 풀업 …: 팔은 편 채 …' 처럼 한 줄에 다 적은 형태가 여기 해당하는데,
        예전에는 이런 줄이 통째로 버려져 앱에서 웜업이 사라졌다. 콜론에서 갈라
        앞은 제목, 뒤는 내용으로 쓴다."""
        if pre_bul:
            day["pre"].append([pre_head or "웜업", pre_bul])
        elif pre_head:
            head, _, rest = pre_head.partition(":")
            day["pre"].append([head.strip(), [rest.strip()] if rest.strip() else []])

    for rr in range(3, hdr):
        t = s(ws.cell(rr, 1).value)
        if (not t or t in SKIP or t.startswith("메인:") or t.startswith("↑")
                or is_ramp_cell(t)):
            continue
        if t.startswith("·"):
            pre_bul.append(t.lstrip("· ").strip())
        elif pre_bul and is_continuation(t):
            pre_bul[-1] += " " + t
        else:
            flush()
            pre_head, pre_bul = t.strip("[]"), []
    flush()

    # 머리글 → 열 번호 지도 (복귀 시트는 '간접 자극' 열이 없음)
    cols = {}
    for c in range(1, ws.max_column + 1):
        cols[s(ws.cell(hdr, c).value)] = c
    C = lambda name: cols.get(name)

    groups, current, flat = [], None, []
    r = hdr + 1
    while r <= ws.max_row:
        a, b = s(ws.cell(r, 1).value), s(ws.cell(r, C("종목")).value)
        d = s(ws.cell(r, C("중량(Kg)")).value)

        if d == "세트 합계":                      # 표 끝
            day["total"] = int(num(ws.cell(r, C("세트")).value))
            r += 1
            break
        if a and not b:                           # 그룹 머리글 (둔근 / 팔 / 복부)
            current = [a, []]
            groups.append(current)
            r += 1
            continue
        if b:
            item = {
                "n":   int(num(a)) if a else 0,
                "name": b,
                "mg":  s(ws.cell(r, C("근육군")).value),
                "w":   d or "—",
                "s":   int(num(ws.cell(r, C("세트")).value)),
                "r":   s(ws.cell(r, C("횟수")).value),
                # 인터벌 타이머용. 이 열이 없는 워크북(복귀 시트)은 0 = 타이머 없음
                "wk":  int(num(ws.cell(r, C("수행(초)")).value)) if C("수행(초)") else 0,
                "rt":  int(num(ws.cell(r, C("휴식(초)")).value)) if C("휴식(초)") else 0,
                # 세트별 휴식 램프. 열이 없거나 비면 None → rt 하나를 전 세트에
                "rl":  rest_ramp(ws.cell(r, C("세트별 휴식(초)")).value) if C("세트별 휴식(초)") else None,
                # 슈퍼세트 태그. 같은 태그가 붙은 이웃 종목끼리 세트를 번갈아 돈다
                "ss":  s(ws.cell(r, C("슈퍼세트")).value) if C("슈퍼세트") else "",
                "memo": s(ws.cell(r, C("메모")).value) if C("메모") else "",
                "alt": s(ws.cell(r, C("대체 종목")).value) if C("대체 종목") else "",
                "ind": s(ws.cell(r, C("간접 자극")).value) if C("간접 자극") else "",
                # 이 종목이 몇 주차부터 나오는지. 열이 없는 워크북은 전부 1 = 램프 없음
                "rw": int(num(ws.cell(r, C("복귀 주차")).value, 1)) if C("복귀 주차") else 1,
                # 마지막 세트에 붙는 드롭 세트 시간(초). 열이 없거나 비면 0 = 드롭 없음.
                # 세트 수(E열)는 그대로다 — 드롭은 같은 세트의 연장이지 새 세트가 아니다
                "drop": int(num(ws.cell(r, C("드롭(초)")).value)) if C("드롭(초)") else 0,
            }
            # 메인 리프트 중량은 주차에 따라 바뀌므로 상단 카드로 넘긴다.
            # 복귀 시트는 이름이 '메인:' 줄과 같고, 재택 시트는 1번 '… (본세트)' 행이다.
            lift = main_of.get(item["name"])
            # 재택 시트에도 같은 이름이 있다(45kg · 8~10) — 추적은 복귀판에서만
            if wsf is not None:
                m = ACC_REF.match(s(wsf.cell(r, C("중량(Kg)")).value))
                trk = ACC_BY_ROW.get(int(m.group(1))) if m else None
            else:
                trk = TRACK_KEYS.get(item["name"]) if ws.title.startswith("복귀") else None
            wk_ = (ws.title.split("_")[1], item["name"])
            if is_back(ws) and wk_ in WARMUP:
                item["wu"] = True
                item["wut"] = WARMUP[wk_] or WU_DEFAULT
            if is_back(ws) and item["name"] in MICRO:
                item["mic"] = MICRO[item["name"]]
            if is_back(ws) and wk_ in KNEE:
                item["kn"] = KNEE[wk_]
            if trk:
                # 중량은 앱이 주차로 낸다 — 엑셀 캐시값은 C4가 늦으면 지난 주 값이다
                item["lift"], item["trk"] = trk, True
            elif not lift and day["main"] and item["n"] == 1 and "본세트" in item["name"]:
                lift = day["main"]
            if lift:
                item["w"] = "위 참조"
                item["lift"] = lift
            # 백오프 = 직전 사이클의 톱세트 중량. 주차를 옮기면 같이 따라와야 하므로
            # 엑셀 수식의 캐시값을 쓰지 않고 리프트에 'lo'(몇 사이클 전)로 매단다
            # 백오프 = 톱세트 − 5kg(= 2주 전 중량). 주차를 옮기면 같이 따라와야 하므로
            # 엑셀 수식의 캐시값을 쓰지 않고 리프트에 'lo'(톱보다 몇 kg 아래)로 매단다
            # 복귀판은 2026-10-11부터 톱세트 × 0.82(1kg 반올림) — 톱 −5kg는 94%라 백오프가 아니었다.
            # 엑셀 D열 =MAX(20,MROUND(톱*0.82,1))과 같은 식이어야 한다
            elif "백오프" in item["name"] and day["main"]:
                item["lift"] = day["main"]
                if is_back(ws):
                    item["pct"] = 0.82
                else:
                    item["lo"] = 5
            (current[1] if current else flat).append(item)
        r += 1

    day["groups"] = groups or None
    day["ex"] = flat
    if day["total"] == 0:
        allx = flat + [e for _, g in groups for e in g]
        day["total"] = sum(e["s"] for e in allx)

    # 표 아래 A열 텍스트 = 해설. '·'로 시작하면 항목, 아니면 소제목
    head, bullets = None, []
    for rr in range(r, ws.max_row + 1):
        t = s(ws.cell(rr, 1).value)
        if not t:
            continue
        if t.startswith("·"):
            bullets.append(t.lstrip("· ").strip())
        elif bullets and is_continuation(t):  # 셀에서 줄바꿈된 문장의 뒷부분
            bullets[-1] += " " + t
        else:
            if bullets:                       # 소제목 없이 먼저 나온 항목도 살린다
                day["notes"].append([head or "구성 근거", bullets])
            head, bullets = t.strip("[]"), []
    if bullets:
        day["notes"].append([head or "구성 근거", bullets])

    # 부제 — 2행 안내문이 있으면 사용, 없으면 근육군 요약
    sub = s(ws.cell(2, 1).value)
    mgs, seen = [], set()
    for e in flat + [x for _, g in groups for x in g]:
        if e["mg"] and e["mg"] not in seen:
            seen.add(e["mg"]); mgs.append(e["mg"])
    day["sub"] = f"{sub} · {day['total']}세트" if sub else \
                 f"{' · '.join(mgs[:4])} · {day['total']}세트"
    return day


# 판정 문자열 → 색. 하한 미달은 노랑(정보)에 그치고 빨강은 상한 초과에만 준다 —
# 볼륨을 늘리라는 신호가 아니라 줄이라는 신호에만 경보를 붙인다.
# 재택 시트는 성장/유지/유지 미만/상한, 복귀 시트는 아직 충분/하한/보완 필요를 쓴다.
VERDICT_TAG = {"성장": "ok", "유지": "ok", "충분": "ok",
               "상한": "bad", "과다 점검": "bad"}


def parse_check(wb, sheet="세트검산", adj_col=None):
    """근육군별 검산표. adj_col을 주면 그 열을 '현장 일 보정' 등급으로 함께 읽는다.

    재택 시트는 G열이 합산 라벨이라 adj_col을 주지 않고, 복귀 시트만 G열에 보정 등급이 있다.
    J열(주당 환산)은 주기가 7일이 아닌 재택 시트에만 있다 — 복귀판은 주 단위라 환산이 없어 0이 실린다."""
    ws = wb[sheet]
    rows = []
    for r in range(5, ws.max_row + 1):
        name = s(ws.cell(r, 1).value)
        if not name or name.startswith("직접 세트 총계") or name == "기준선":
            break
        direct, ind = num(ws.cell(r, 2).value), num(ws.cell(r, 3).value)
        real = num(ws.cell(r, 4).value, direct + ind)
        verdict = s(ws.cell(r, 5).value)
        tag = VERDICT_TAG.get(verdict, "lo")
        rows.append([name, half(direct), half(ind), half(real), tag, verdict,
                     s(ws.cell(r, 6).value),
                     s(ws.cell(r, adj_col).value) if adj_col else "",
                     half(ws.cell(r, 10).value)])
    return rows


def parse_budget(wb, sheet="세트검산"):
    """세트검산 상단(G1:L3)의 총량 예산 블록. 블록이 없으면 None — 복귀 시트에는 아직 없다.

    엑셀은 가끔 열지만 폰은 매일 본다. INVALID가 워크북 안에만 있으면 아무도 안 본다."""
    ws = wb[sheet]
    if s(ws["K3"].value) != "상태":
        return None
    iv = lambda c: int(num(ws[c].value))          # 세트는 정수다 — 239.0으로 띄우지 않는다
    return {"cycle": iv("H2"), "cap": iv("J2"), "cur": iv("L2"), "left": iv("H3"),
            "weekly": num(ws["J3"].value), "state": s(ws["L3"].value) or "VALID"}


def parse_notes_block(ws, start_row, default_head="메모"):
    out, head, bullets = [], None, []
    for r in range(start_row, ws.max_row + 1):
        t = s(ws.cell(r, 1).value)
        if not t:
            continue
        if t.startswith("·") or t.startswith("→"):
            bullets.append(t.lstrip("·→ ").strip())
        elif bullets and is_continuation(t):
            bullets[-1] += " " + t
        else:
            if bullets:
                out.append([head or default_head, bullets])
            head, bullets = t.strip("[]"), []
    if bullets:
        out.append([head or default_head, bullets])
    return out


def parse_plan(wb, sheet="주간계획"):
    """주간계획. 열 위치는 머리글로 찾는다 — 6일판은 '메인 리프트' 열이 끼어들어 세트가 E열로 밀렸다."""
    ws = wb[sheet]
    hdr = next((r for r in range(1, 12)
                if s(ws.cell(r, 1).value) in ("일차", "요일")), 4)
    cols = {s(ws.cell(hdr, c).value): c for c in range(1, ws.max_column + 1)}
    c_body = cols.get("내용", 2)
    c_sets = cols.get("세트") or cols.get("예상 세트") or 4

    rows, r = [], hdr + 1
    while r <= ws.max_row:
        d = s(ws.cell(r, 1).value)
        if not d or d == "합계":
            break
        body, sets = s(ws.cell(r, c_body).value), int(num(ws.cell(r, c_sets).value))
        # openpyxl로 워크북을 저장하면 수식 캐시가 비고, 여기 세트가 전부 0으로 읽힌다.
        # 그대로 빌드하면 개요 화면이 0세트인 앱이 조용히 배포된다 — 2026-09-04에 실제로 겪었다.
        if sets == 0 and body and "휴식" not in body:
            print(f"[중단] 주간계획 {d}일차 '{body}' 세트가 0입니다 — 엑셀 수식 캐시가 비었습니다.")
            sys.exit("        워크북을 엑셀로 열어 Ctrl+Alt+F9 (전체 재계산) 후 저장하고 다시 실행하세요.")
        rows.append([d, body, sets])
        r += 1

    return {"sub": s(ws.cell(2, 1).value), "rows": rows,
            "total": sum(x[2] for x in rows),
            "notes": parse_notes_block(ws, r + 1, "배치 논리")}


def parse_ramp(wb):
    """볼륨램프 — 주차별로 어느 종목이 복귀하는지.

    실제 세트를 정하는 것은 요일 시트 K열('복귀 주차')이고, 이 시트는 그 요약과 근거다.
    시트가 없거나 표를 못 찾으면 None (= 항상 설계 볼륨 100%).
    """
    if "볼륨램프" not in wb.sheetnames:
        return None
    ws = wb["볼륨램프"]
    hdr = next((r for r in range(1, 12) if s(ws.cell(r, 1).value) == "주차"
                and s(ws.cell(r, 2).value) == "총 세트"), None)
    if hdr is None:
        return None

    rows, last = [], hdr
    for r in range(hdr + 1, ws.max_row + 1):
        label = s(ws.cell(r, 1).value)
        if not label:                      # 머리글 아래 빈 줄은 넘기고, 표가 끝나면 멈춘다
            if rows:
                break
            continue
        digits = re.findall(r"\d+", label)
        total = ws.cell(r, 2).value
        if not digits or not isinstance(total, (int, float)):
            break
        # '2주 (현재)' 같은 상태 표기는 떼어낸다 — 어느 주차인지는 앱이 직접 계산해 표시한다
        rows.append({"week": int(digits[0]),
                     "label": re.sub(r"\s*\((?:완료|현재)\)", "", label),
                     "total": int(total), "delta": s(ws.cell(r, 3).value),
                     "items": s(ws.cell(r, 4).value)})
        last = r
    if not rows:
        return None
    return {"sub": s(ws.cell(2, 1).value), "rows": rows,
            "offset": RAMP_WEEK_OFFSET,
            "notes": parse_notes_block(ws, last + 1, "설계 논리")}


def parse_home_anchor(wb):
    """재택 1일차 기준일 — 시작중량_설정의 전용 칸. 라벨로 찾는다.

    2026-09-13 전까지는 증량기록 '날짜' 열의 마지막 값을 썼는데, 그 열이 주차별
    월요일로 채워지면서 기준이 미래 날짜가 돼 버렸다. 칸을 따로 뒀다."""
    ws = wb["시작중량_설정"]
    for r in range(1, ws.max_row + 1):
        if s(ws.cell(r, 1).value) == "재택 1일차 기준일":
            v = ws.cell(r, 3).value
            if hasattr(v, "strftime"):
                return {"date": v.strftime("%Y-%m-%d"),
                        "cycle": int(num(ws.cell(4, 3).value, 1))}
    return None


def parse_anchor(wb):
    """증량기록 '날짜 (1일차)' 열 — 그 사이클의 1일차 날짜. 달력의 기준점이다.

    여러 줄이 채워져 있으면 가장 늦은 날짜를 쓴다. 중간에 쉬어서 주기가 밀리면
    다음 사이클 줄에 날짜를 적는 것만으로 기준이 다시 맞는다.
    비어 있으면 None — 앱이 '오늘이 몇 일차인지' 직접 묻는다.
    """
    if "증량기록" not in wb.sheetnames:
        return None
    ws = wb["증량기록"]
    hdr = next((r for r in range(1, 60)
                if s(ws.cell(r, 1).value) in ("주차", "사이클")), None)
    if hdr is None:
        return None
    c_date = next((c for c in range(1, ws.max_column + 1)
                   if s(ws.cell(hdr, c).value).startswith("날짜")), None)
    if not c_date:
        return None
    best = None
    for r in range(hdr + 1, ws.max_row + 1):
        label = s(ws.cell(r, 1).value)
        if not label:
            if best:
                break
            continue
        m = re.match(r"(\d+)", label)
        v = ws.cell(r, c_date).value
        if m and isinstance(v, datetime.datetime):
            cur = {"date": v.strftime("%Y-%m-%d"), "cycle": int(m.group(1))}
            if best is None or cur["date"] > best["date"]:
                best = cur
    return best


def parse_back_plan(wb):
    ws = wb["복귀_주간계획"]
    hdr = None
    for r in range(1, 12):
        # 주 6회판은 요일이 고정이라 머리글이 '요일'이다. 옛 주 4회판('슬롯')도 같이 받는다.
        if s(ws.cell(r, 1).value) in ("요일", "슬롯"):
            hdr = r; break
    rows = []
    if hdr:
        for r in range(hdr + 1, hdr + 9):
            d = s(ws.cell(r, 1).value)
            if not d or d == "합계":
                break
            rows.append([d, s(ws.cell(r, 2).value), s(ws.cell(r, 3).value),
                         s(ws.cell(r, 4).value), int(num(ws.cell(r, 5).value))])
    return {"sub": " · ".join(x for x in
                              [s(ws.cell(2, 1).value), s(ws.cell(3, 1).value)] if x),
            "rows": rows, "total": sum(x[4] for x in rows),
            "notes": parse_notes_block(ws, (hdr or 4) + 10, "배치 논리")}


def parse_cycle_plan(wb):
    """사이클_계획 4~13행 표 → [일차, 구성, 세트, 본운동 분]. 아래 [배치 원칙] 블록부터 해설."""
    ws = wb["사이클_계획"]
    rows = []
    for r in range(5, 5 + CYCLE_LEN):
        v = ws.cell(r, 1).value
        if v is None:
            break
        d = str(int(v)) if isinstance(v, (int, float)) else s(v)
        sets = int(num(ws.cell(r, 3).value))
        body = s(ws.cell(r, 2).value)
        if sets == 0 and body and "휴식" not in body:
            print(f"[중단] 사이클_계획 {d}일차 '{body}' 세트가 0입니다 — 엑셀 수식 캐시가 비었습니다.")
            sys.exit("        워크북을 엑셀로 열어 Ctrl+Alt+F9 (전체 재계산) 후 저장하고 다시 실행하세요.")
        rows.append([d, body, sets, int(num(ws.cell(r, 4).value))])
    start = next((r for r in range(15, ws.max_row + 1) if s(ws.cell(r, 1).value).startswith("[")), 44)
    return {"sub": s(ws.cell(2, 1).value), "rows": rows,
            "total": sum(x[2] for x in rows), "mins": sum(x[3] for x in rows),
            "notes": parse_notes_block(ws, start, "배치 원칙")}


def parse_cycle_check(wb):
    """사이클_계획 부위별 표('부위' 머리글) → [부위, 사이클당, 주당, MAV, 태그, 판정, 구성]."""
    ws = wb["사이클_계획"]
    hdr = next((r for r in range(10, 40) if s(ws.cell(r, 1).value) == "부위"), None)
    if hdr is None:
        return []
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        name = s(ws.cell(r, 1).value)
        if not name or name.startswith(("·", "[")):
            break
        cyc, wk, mav = num(ws.cell(r, 2).value), num(ws.cell(r, 3).value), ws.cell(r, 4).value
        if name == "합계":
            rows.append([name, half(cyc), half(wk), "", "", "", ""])
            continue
        if name.startswith("검산"):
            rows.append([name, half(cyc), "", "", "ok" if s(ws.cell(r, 3).value) == "일치" else "bad",
                         s(ws.cell(r, 3).value), ""])
            continue
        has = isinstance(mav, (int, float))
        desc = s(ws.cell(r, 5).value)
        # 'MAV 14는 등 전체(광배 + 상부 26)에 적용' — 그 줄의 MAV는 묶음 합계로 판정한다(줄 하나로 재면 미달로 보인다)
        grp = re.search(r"전체\([^)]*?(\d+)\)", desc)
        base = round(int(grp.group(1)) * 7 / CYCLE_LEN, 1) if grp else wk
        tag = ("ok" if base >= mav else "lo") if has else ""
        verdict = (("충분" if base >= mav else "미달") + (f" (합산 {half(base)})" if grp else "")) if has else ""
        rows.append([name, half(cyc), half(wk), half(mav) if has else "", tag, verdict, desc])
    return rows


def parse_cycle_anchor(wb):
    """시작중량_설정의 사이클 운용 칸 — 1일차 기준일(날짜) · 1사이클 중량 칸 · 첫 디로드 사이클 · 디로드 간격(사이클).
    2026-11-01부터 중량은 사이클마다 한 칸. 기준일이 없으면 None — 앱이 오늘 일차를 묻는다."""
    ws = wb["시작중량_설정"]
    want = {"사이클 1일차 기준일": "date", "1사이클 중량 칸": "base",
            "첫 디로드 사이클": "dlFirst", "디로드 간격(사이클)": "dlEvery"}
    got = {}
    for r in range(1, ws.max_row + 1):
        k = want.get(s(ws.cell(r, 1).value))
        if not k:
            continue
        v = ws.cell(r, 3).value
        if k == "date":
            got[k] = v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else None
        elif hasattr(v, "toordinal"):
            # 위 날짜 칸 아래에 행을 끼워 넣어 서식이 날짜로 따라왔다 — 14가 1900-01-14로 읽힌다. 일련번호로 되돌린다
            # (엑셀 1900 윤년 버그 — 60 미만 일련번호는 기준이 1899-12-31이다)
            ep = datetime.date(1899, 12, 31) if v.toordinal() < datetime.date(1900, 3, 1).toordinal() \
                else datetime.date(1899, 12, 30)
            got[k] = v.toordinal() - ep.toordinal()
        else:
            got[k] = int(num(v))
    if not got.get("date"):
        return None
    return {"date": got["date"], "base": got.get("base") or int(num(ws.cell(4, 3).value, 1)),
            "dlFirst": got.get("dlFirst") or 0, "dlEvery": got.get("dlEvery") or 0}


GRIP_SHEET = "사이클_그립회전"
# 2026-11-04 '원암 케이블 로우' → '케이블 로우'로 이름에서 원암을 뗐지만 시트 수행 초(120)는 여전히 원암(좌우) 기준이다
ONEARM_BASE = {"케이블 로우"}


def grip_display_name(name):
    """회전 중에는 이름 속 그립 표기가 틀린 말이 된다 — '랫풀다운(언더그립)'이 뉴트럴인 사이클. 화면 · 음성용으로 뗀다."""
    t = re.sub(r"\s*\((?:와이드 오버그립|오버그립|언더그립|뉴트럴 그립|오버 그립|언더 그립)\)", "", name)
    t = re.sub(r"^(?:와이드\s*)?(?:오버그립|언더그립)\s+", "", t)
    return t.strip() or name


def parse_grip_rotation(wb):
    """사이클_그립회전 — {start, rows:[[일차, 종목, 주기, [그립…], 메모, 원암축]]}. 시트가 없으면 None.
    원암축(H열, 2026-11-01): '양손 먼저' → 0 · '원암 먼저' → 1 · 빈칸 → None. 그립(3사이클)과 따로
    2사이클마다 원암 ↔ 양손. 원암은 언더 차례에만 — 오버 · 뉴트럴(넓은 패럴렐 · 팔꿈치 45°)은 늘 양손(2026-11-01 S1)."""
    if GRIP_SHEET not in wb.sheetnames:
        return None
    ws = wb[GRIP_SHEET]
    start = None
    for r in range(1, 6):
        if s(ws.cell(r, 1).value) == "회전 시작 사이클":
            start = int(num(ws.cell(r, 3).value, 0)) or None
    hdr = next((r for r in range(1, 15) if s(ws.cell(r, 1).value) == "일차"), None)
    if not start or hdr is None:
        return None
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if not isinstance(d, (int, float)):
            break
        per = max(1, min(3, int(num(ws.cell(r, 3).value, 3))))
        cells = [s(ws.cell(r, 4 + i).value) for i in range(per)]
        # ' | ' 뒤는 자세 큐(2026-11-01) — 화면에만 뜨고 음성 · 원암 판정은 앞부분(그립)으로
        grips = [c.split("|")[0].strip() for c in cells]
        cues = [c.split("|", 1)[1].strip() if "|" in c else "" for c in cells]
        if not all(grips):
            sys.exit(f"[중단] {GRIP_SHEET} {r}행 '{s(ws.cell(r, 2).value)}' — 주기 {per}인데 회전 칸이 비었습니다.")
        h = s(ws.cell(r, 8).value)
        if h and h not in ("양손 먼저", "원암 먼저"):
            sys.exit(f"[중단] {GRIP_SHEET} H{r} '{h}' — '양손 먼저' · '원암 먼저' · 빈칸 중 하나여야 합니다.")
        arm = {"양손 먼저": 0, "원암 먼저": 1}.get(h)
        rows.append([str(int(d)), s(ws.cell(r, 2).value), per, grips, s(ws.cell(r, 7).value), arm, cues])
    # 회전별 순서 바꾸기(2026-11-04) — '회전별 순서 바꾸기' 머리글 아래 [일차, 회전, 종목 A, 종목 B]. 그 회전에만 두 자리를 맞바꾼다
    order = []
    oh = next((r for r in range(1, ws.max_row + 1) if s(ws.cell(r, 1).value) == "회전별 순서 바꾸기"), None)
    if oh is not None:
        for r in range(oh + 2, ws.max_row + 1):
            d = ws.cell(r, 1).value
            if not isinstance(d, (int, float)):
                break
            order.append([str(int(d)), int(num(ws.cell(r, 2).value)), s(ws.cell(r, 3).value), s(ws.cell(r, 4).value)])
    return {"start": start, "rows": rows, "order": order}


ALT_SHEET = "사이클_종목교대"
# 사이클마다 갈아 끼우는 칸 — 일차 시트 종목(짝수 사이클)과 교대표 종목(홀수 사이클)이 이 칸들만 서로 바꾼다.
# 번호(#) · 세트는 자리 그대로다(세트가 다르면 일차 합계가 틀어지니 빌드에서 막는다)
ALT_KEYS = ("name", "mg", "w", "s", "r", "wk", "rt", "rl", "ss", "memo", "alt", "ind", "drop",
            "lift", "trk", "pct", "lo", "wu", "wut", "mic", "rw", "kn")
# 교대표 중량 칸 수식 — '=시작중량_설정!$E$9' 또는 '=MAX(20,MROUND(시작중량_설정!$E$9*0.95,1))'
ALT_REF = re.compile(r"^=(?:MAX\(20,\s*MROUND\()?시작중량_설정!\$E\$(\d+)(?:\*([\d.]+),\s*1\)\))?$")


def parse_cycle_alt(wb, wbf):
    """사이클_종목교대 (2026-11-02) — [{일차, 시트 종목, 홀짝, 변형}]. 시트가 없으면 [].
    같은 자리 · 같은 세트에서 사이클마다 종목만 바뀐다. 표의 종목이 '차례' 칸 사이클(홀수/짝수)에,
    일차 시트 종목이 나머지 사이클에 들어간다. 중량 칸이 시작중량_설정 E열을 그대로 가리키면
    그 리프트의 사다리를 따라간다 — 메인 넷은 pct(백오프와 같은 길) · 추적 보조는 trk."""
    if ALT_SHEET not in wb.sheetnames:
        return []
    ws, wf = wb[ALT_SHEET], wbf[ALT_SHEET]
    hdr = next((r for r in range(1, 15) if s(ws.cell(r, 1).value) == "일차"), None)
    if hdr is None:
        return []
    cols = {s(ws.cell(hdr, c).value): c for c in range(1, ws.max_column + 1)}
    need = ("일차 시트 종목", "차례", "종목", "근육군", "중량(Kg)", "세트", "횟수",
            "수행(초)", "휴식(초)", "메모", "대체 종목")
    miss = [k for k in need if k not in cols]
    if miss:
        sys.exit(f"[중단] {ALT_SHEET} 머리글에 {', '.join(miss)} 칸이 없습니다.")
    C = lambda k, r: ws.cell(r, cols[k]).value
    by_row = {m["start_row"]: k for k, m in LIFT_META.items()}
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        d = ws.cell(r, 1).value
        if not isinstance(d, (int, float)):
            break
        turn = s(C("차례", r))
        # 2026-11-10 'N사이클부터' — 번갈지 않고 그 사이클부터 계속 표 종목(4일차 핵 스쿼트 복귀)
        mf = re.match(r"^(\d+)사이클부터$", turn)
        if turn not in ("홀수 사이클", "짝수 사이클") and not mf:
            sys.exit(f"[중단] {ALT_SHEET} {r}행 차례 '{turn}' — '홀수 사이클' · '짝수 사이클' · 'N사이클부터' 중 하나여야 합니다.")
        v = {k: None for k in ALT_KEYS}
        v.update({"name": s(C("종목", r)), "mg": s(C("근육군", r)), "s": int(num(C("세트", r))),
                  "r": s(C("횟수", r)), "wk": int(num(C("수행(초)", r))), "rt": int(num(C("휴식(초)", r))),
                  "memo": s(C("메모", r)), "alt": s(C("대체 종목", r)), "ind": "", "drop": 0, "rw": 1,
                  "w": s(C("중량(Kg)", r)) or "—"})
        if "슈퍼세트" in cols:
            v["ss"] = s(C("슈퍼세트", r)) or None
        m = ALT_REF.match(s(wf.cell(r, cols["중량(Kg)"]).value))
        k = by_row.get(int(m.group(1))) if m else None
        if k and LIFT_META[k].get("acc"):
            v["lift"], v["trk"] = k, True
        elif k:
            v["lift"], v["pct"] = k, float(m.group(2) or 1)
        if v["name"] in MICRO:
            v["mic"] = MICRO[v["name"]]
        if (str(int(d)), v["name"]) in KNEE:
            v["kn"] = KNEE[(str(int(d)), v["name"])]
        rows.append({"day": str(int(d)), "base": s(C("일차 시트 종목", r)),
                     "side": "from" if mf else "odd" if turn == "홀수 사이클" else "even",
                     "from": int(mf.group(1)) if mf else None, "v": v})
    return rows


def say_grip(t):
    """그립 칸을 읽을 말로 — '오버 넓게 (양손)' → '오버 그립 넓게 양손'. 앱 쪽 키는 '그립:' + 칸 글자 그대로."""
    t = re.sub(r"[()（）·/]", " ", t)
    t = re.sub(r"\s+", " ", t).strip()
    return re.sub(r"^(오버|언더|뉴트럴)(?=\s|$)", r"\1 그립", t)


def parse_split_check(wb):
    """분할_주간계획 부위별 표('부위' 머리글 · 세트 · 구성) — 주 단위라 환산이 없다."""
    ws = wb["분할_주간계획"]
    hdr = next((r for r in range(1, 40) if s(ws.cell(r, 1).value) == "부위"), None)
    if hdr is None:
        return []
    rows = []
    for r in range(hdr + 1, ws.max_row + 1):
        name = s(ws.cell(r, 1).value)
        if not name or name.startswith(("·", "[")):
            break
        n = half(ws.cell(r, 2).value)
        if name.startswith("검산"):
            v = s(ws.cell(r, 3).value)
            rows.append([name, n, "", "", "ok" if v == "일치" else "bad", v, ""])
        else:
            rows.append([name, n, "", "", "", "", "" if name == "합계" else s(ws.cell(r, 3).value)])
    return rows


def parse_prep(wb):
    """생수배송 대비 두 시트(복귀_캐리 · 복귀_스텝박스) → 앱 '배송' 탭. 세트 카운트 밖이다.

    표는 머리글 첫 칸('항목' · '요일' · '주차')으로 찾는다 — 행 번호를 박지 않는다."""
    if "복귀_캐리" not in wb.sheetnames or "복귀_스텝박스" not in wb.sheetnames:
        return None

    def table(ws, head, cols, start=1):
        r0 = next((r for r in range(start, ws.max_row + 1)
                   if s(ws.cell(r, 1).value) == head), None)
        if r0 is None:
            return [], start
        out, r = [], r0 + 1
        while s(ws.cell(r, 1).value):
            out.append([ws.cell(r, c).value for c in cols])
            r += 1
        return out, r

    def iso(v):
        return v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else s(v)

    c, t = wb["복귀_캐리"], wb["복귀_스텝박스"]
    # 집 — 진짜 생수팩 + 아파트 계단. 시트가 없으면 앱에서 그 칸만 빠진다
    home = None
    if "복귀_집계단" in wb.sheetnames:
        h = wb["복귀_집계단"]
        hrows, hend = table(h, "시작일", (1, 2, 3, 4, 5))
        home = {"t": s(h.cell(1, 1).value), "sub": s(h.cell(2, 1).value),
                "rows": [[iso(d), s(w), s(am), s(pm), s(q)] for d, w, am, pm, q in hrows],
                "notes": parse_notes_block(h, hend + 1, "규칙")}
    rules, _ = table(c, "항목", (1, 2, 6))
    prog, _ = table(c, "회차", (1, 2, 3, 4, 5, 6, 7))
    # 진행표 바로 밑 '· …' 주석 줄은 빈 줄 없이 붙어 있어 표에 딸려 들어온다
    foot = [r[0] for r in prog if not isinstance(r[0], (int, float))]
    prog = [r for r in prog if isinstance(r[0], (int, float))]
    days, _ = table(t, "요일", (1, 2, 9))
    ramp, end2 = table(t, "회차", (1, 2, 3, 4, 5, 6, 7, 8, 9))
    sfoot = [r[0] for r in ramp if not isinstance(r[0], (int, float))]
    ramp = [r for r in ramp if isinstance(r[0], (int, float))]
    return {
        "carry": {
            "t": s(c.cell(1, 1).value), "sub": s(c.cell(2, 1).value),
            "rules": [[s(a), s(b), s(x)] for a, b, x in rules],
            # 걷기(초) G열 — 2026-10-25 덤벨이 30 다음 35라 무게 대신 시간으로 칸을 쪼갠다. 비면 40
            "prog": [[int(num(w)), iso(d), num(kg), num(both), s(act), s(memo), int(num(sec, 40)) or 40]
                     for w, d, kg, both, act, memo, sec in prog],
            "foot": s(foot[0]).lstrip("· ") if foot else "",
        },
        "stairs": {
            "t": s(t.cell(1, 1).value), "sub": s(t.cell(2, 1).value),
            "days": [[s(a), s(b), s(x)] for a, b, x in days],
            # [회차, 날짜, 메인 kg, 메인 분, 메인 세트, 짧은 날 kg, 짧은 날 분, 짧은 날 세트, 짧은 날 손(스트랩)]
            "ramp": [[int(num(r[0])), iso(r[1])] + [num(v) for v in r[2:8]] + [s(r[8])] for r in ramp],
            "foot": s(sfoot[0]).lstrip("· ") if sfoot else "",
            "notes": parse_notes_block(t, end2 + 1, "규칙"),
        },
        "home": home,
    }


def parse_config(wb):
    st, gr = wb["시작중량_설정"], wb["증량기록"]
    lifts = {}
    for key, m in LIFT_META.items():
        lifts[key] = {
            "ko": m["ko"], "accent": m["accent"],
            "start": num(st.cell(m["start_row"], 4).value),
            "prev":  num(st.cell(m["start_row"], 2).value),
            "inc":   num(gr.cell(m["inc_row"], 3).value),
            "sets":  int(num(re.sub(r"\D+", " ", s(st.cell(m["start_row"], 6).value))
                             .split()[-1], 1)),
            # 본세트 반복 — F열 '8회 × 4세트'의 첫 숫자. 벤치(2026-10-02) · 프레스(2026-10-21)가
            # 8회로 바뀌었는데 카드가 5회로 고정돼 있었다
            "reps":  int(num(re.sub(r"\D+", " ", s(st.cell(m["start_row"], 6).value))
                             .split()[0], 5)),
            # 증량 주기(주). 프레스만 2 — 봉 전체 1kg이 최소 계단이라 매주 올리면 두 배 반이 된다
            "every": max(1, int(num(gr.cell(m["inc_row"], 5).value, 1))),
            # 사다리가 출발하는 주차(F열). 비면 표 첫 줄 — 보조 둘은 13주차에 들어왔다
            "from":  int(num(gr.cell(m["inc_row"], 6).value, 0)) or None,
            "acc":   bool(m.get("acc")),
            "day":   m.get("day", ""),
        }
        # 2026-11-09 사용자 결정 — 인클라인 바벨은 따로 증량하지 않고 벤치 목표 × 비율(증량기록 C20 · 0.8).
        # 1kg 내림 · 못 끼우는 값(5로 나눈 나머지 4)은 1kg 더 내림 — 엑셀 AJ열과 같은 식
        if key in RESTART:
            lifts[key]["rs"] = RESTART[key]
        if key in RATIO_OF:
            lifts[key]["of"] = RATIO_OF[key]
            lifts[key]["ratio"] = num(gr.cell(m["inc_row"], 3).value, 0.8)
    # 표의 줄 수와 첫 줄의 번호. 2026-09-13 주 단위 전환으로 표가 6주차에서 시작한다 —
    # '몇 번째 줄'과 '몇 주차'가 더는 같지 않으므로 base를 같이 실어 보낸다
    weeks, base = 0, 0
    # 머리글('주차') 다음 줄부터 센다 — 머리글 자체도 '주차'로 끝나 첫 줄로 잡힌다.
    # 2026-10-23 설정 두 줄(보조 추적)이 들어와 표가 12행으로 내려갔다
    hdr = next((r for r in range(1, 60)
                if s(gr.cell(r, 1).value) in ("주차", "사이클")), 9)
    for r in range(hdr + 1, hdr + 120):
        lab = s(gr.cell(r, 1).value)
        if lab.endswith(("사이클", "주차")):
            if not weeks:
                base = int(num(re.sub(r"\D+", " ", lab).strip(), 1))
            weeks += 1
        elif weeks:
            break
    return {"lifts": lifts, "weeks": weeks or 12, "base": base or 1,
            "week": int(num(st.cell(4, 3).value, 1))}


def parse_back_anchor(wb):
    """복귀 운용 기준 — 시작일 · 시작 사이클 · 사이클 전진(주).

    라벨로 찾는다. 시작중량_설정은 LIFT_META가 행 번호를 박아 쓰는 시트라
    블록을 아래쪽에 두고 행 번호에 의존하지 않는다. 칸이 없으면 None —
    그때는 앱이 복귀 달력 대신 '기준일을 채우라'는 안내를 띄운다."""
    ws = wb["시작중량_설정"]
    want = {"프로그램 시작일": "date", "주차 기준일": "cdate",
            "기준 주차": "cycle", "증량 전진(주)": "step", "디로드 시작일": "deload",
            "디로드 간격(주)": "dlevery"}
    got = {}
    for r in range(1, ws.max_row + 1):
        key = want.get(s(ws.cell(r, 1).value))
        if not key:
            continue
        v = ws.cell(r, 3).value
        if key == "deload":     # 날짜 한 칸, 또는 '2026-10-17, 2026-12-26' 텍스트 — 디로드 주가 시작되는 토요일
            got[key] = ([v.strftime("%Y-%m-%d")] if hasattr(v, "strftime") else
                        ["%04d-%02d-%02d" % tuple(map(int, m))
                         for m in re.findall(r"(\d{4})[-./](\d{1,2})[-./](\d{1,2})", s(v))])
            continue
        got[key] = v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else int(num(v))
    if not got.get("date") or not got.get("cdate") or not got.get("cycle"):
        return None
    return {"date": got["date"], "cdate": got["cdate"],
            "cycle": got["cycle"], "step": got.get("step") or 2,
            "deload": got.get("deload") or [], "dlevery": got.get("dlevery") or 0}


def parse_progress_notes(wb):
    """표 아래 해설. 표 길이가 바뀌므로 머리글을 찾아 표 끝 다음 줄부터 훑는다."""
    ws = wb["증량기록"]
    hdr = next((r for r in range(1, 60)
                if s(ws.cell(r, 1).value) in ("주차", "사이클")), 9)
    r = hdr + 1
    while r <= ws.max_row and re.match(r"\d+", s(ws.cell(r, 1).value)):
        r += 1
    return parse_notes_block(ws, r, "운용 규칙")


def parse_log(wb):
    """증량기록의 실제 기록 — 미달 시 반복수 · 메모 · 다음날 피로 · 관절 이상.

    앱이 엑셀과 같은 규칙(공란·O = 다음 사이클 증량, 그 밖의 입력 = 유지)으로
    중량을 내므로 엑셀에 적힌 기록은 여기서 실려 가고, 폰에서 적은 기록이 그 위에 얹힌다.
    값이 하나도 없는 사이클은 넣지 않는다 — 없는 데이터는 없는 채로 둔다.
    열은 머리글로 찾는다: '<종목> 목표' 바로 오른쪽의 '미달 시 …' 열이 그 종목의 기록이다."""
    if "증량기록" not in wb.sheetnames:
        return {}
    ws = wb["증량기록"]
    hdr = next((r for r in range(1, 60)
                if s(ws.cell(r, 1).value) in ("주차", "사이클")), None)
    if hdr is None:
        return {}
    heads = {c: s(ws.cell(hdr, c).value) for c in range(1, ws.max_column + 1)}
    cols = {}
    for c, h in heads.items():
        for key, m in LIFT_META.items():
            if h == m["head"] and heads.get(c + 1, "").startswith("미달"):
                cols[key] = c + 1
        if h == "메모":
            cols["memo"] = c
        elif h.startswith("다음날 피로"):
            cols["fat"] = c
        elif h.startswith("관절"):
            cols["joint"] = c
        elif h.startswith("스트랩"):        # 2026-11-12 AN열 — 스트랩 쓴 종목 · 세트(판정과 무관)
            cols["strap"] = c
    out, seen = {}, False
    for r in range(hdr + 1, ws.max_row + 1):
        m = re.match(r"(\d+)", s(ws.cell(r, 1).value))
        if not m:
            if seen:
                break
            continue
        seen = True
        row = {f: s(ws.cell(r, c).value) for f, c in cols.items()}
        row = {f: v for f, v in row.items() if v}
        if row:
            out[int(m.group(1))] = row
    return out


def all_ex(day):
    """요일 하나의 종목 전부. 그룹으로 묶인 시트와 평평한 시트 양쪽을 같이 다룬다."""
    return (day.get("ex") or []) + [x for _, g in (day.get("groups") or []) for x in g]


def build_voice(site, data):
    """'다음 종목' 안내에 쓸 말소리를 굽고 스프라이트 한 덩어리로 묶는다.

    통째로 구울 수 있는 건 통째로 굽는다 — 종목 이름, 근육군, 그리고
    '맨몸'처럼 숫자가 아닌 칸. 반대로 숫자는 조각으로만 굽는다.
    중량은 주차마다 오르고 앱에서도 고칠 수 있어서 미리 구워 둘 수가 없다.
    그래서 0~99와 백 단위, '점오', '킬로 · 세트 · 회 · 에서'를 따로 두고
    앱이 그 순간의 값으로 이어 붙인다 — voice_build.py 머리말 참조.

    SAPI가 없어 못 구우면 지난 번 결과(sprite.json)를 그대로 쓴다.
    """
    say = {}
    for row in ((data.get("gripRot") or {}).get("rows") or []):
        for g in row[3]:
            say.setdefault("그립:" + g, say_grip(g))
        if row[5] is not None:
            say.setdefault("팔:원암", "원암, 좌우")
            say.setdefault("팔:양손", "양손")
    for src in ("home", "back"):
        for day in data.get(src, {}).values():
            # 사이클 종목 교대 — 오늘 안 얹힌 쪽 이름도 미리 굽는다(앱이 사이클마다 바꿔 읽는다)
            for e in [x for e0 in all_ex(day) for x in [e0] + list((e0.get("cv") or {}).values())]:
                if e.get("gname"):
                    say.setdefault(e["gname"], voice_build.say_name(e["gname"]))
                for raw, fn in ((e.get("name") or "", voice_build.say_name),
                                (e.get("mg") or "", voice_build.say_mg)):
                    if raw and raw not in say:
                        say[raw] = fn(raw)
                # 숫자·범위는 앱이 조립한다. 여기서 굽는 건 예외 표기뿐이다
                for raw, fn in ((str(e.get("r") or ""), voice_build.say_reps),
                                (str(e.get("w") or ""), voice_build.say_weight)):
                    if (raw and raw not in say and raw != "위 참조"
                            and not voice_build.is_plain_number(raw)):
                        say[raw] = fn(raw)

    for n in range(100):                                  # 0~99
        say.setdefault(voice_build.sino(n), voice_build.sino(n))
    for h in range(1, 10):                                # 백 ~ 구백
        say.setdefault(voice_build.hundreds(h), voice_build.hundreds(h))
    for w in list(voice_build.FRAC.values()) + ["점"] + voice_build.UNITS:
        say.setdefault(w, w)
    # 드롭이 걸린 종목의 마지막 세트 시작에 '운동' 뒤에 붙는 한 마디
    say.setdefault("드롭세트", "드롭 세트")
    # 2026-10-24 — 예열 1세트 종목 안내 꼬리 · 미세 단계 무게('5 + 핀 + 1' · '8 + 손목 1')를 읽는 토막
    say.setdefault("예열 한 세트 먼저", "예열 한 세트 먼저")
    for w in ("더하기 핀", "더하기 손목", "더하기"):
        say.setdefault(w, w)

    say = {k: v for k, v in say.items() if v}
    # 삐 소리 뒤에 붙는 네 마디는 이미 구워져 있다. 다시 굽지 않고 그대로 넣는다 —
    # 새로 구우면 지금 쓰는 말투와 미묘하게 달라진다.
    old = {k: site / "voice" / (k + ".wav") for k in ("prep", "work", "rest", "done")}
    old = {k: v for k, v in old.items() if v.exists()}

    src, offs = voice_build.build(site, say, old)
    keep = site / "voice" / "sprite.json"
    if src:
        voice = {"src": src, "map": offs}
        keep.write_text(json.dumps(voice, ensure_ascii=False), encoding="utf-8")
        return voice
    if keep.exists():          # 이번엔 못 구웠다 — 지난 번 것을 지킨다
        try:
            return json.loads(keep.read_text(encoding="utf-8"))
        except ValueError:
            pass
    return None


def build(xlsx_path, out_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    home = split_days(wb)
    if not home:
        sys.exit("주 6일 분할 시트를 찾을 수 없습니다 — '분할_월_하체A' 처럼 '분할_<요일>_<내용>'이어야 합니다.")
    cyc = cycle_days(wb)
    if not cyc:
        sys.exit("9일 사이클 시트를 찾을 수 없습니다 — '사이클_1_하체A' 처럼 '사이클_<일차>_<내용>'이어야 합니다.")
    wbf = openpyxl.load_workbook(xlsx_path)          # 수식 판 — 증량표 연결 확인용

    data = {
        "config": parse_config(wb),
        "home":   {label: parse_day(wb[sheet], wbf[sheet]) for label, sheet in home},
        "back":   {label: parse_day(wb[sheet], wbf[sheet]) for label, sheet in cyc},
        "check":  parse_split_check(wb),
        "budget": None,
        "backBudget": None,
        "backCheck": parse_cycle_check(wb),
        "plan":   parse_plan(wb, "분할_주간계획"),
        "ramp":   None,
        "backPlan": parse_cycle_plan(wb),
        "prep":   parse_prep(wb),
        "progressNotes": parse_progress_notes(wb),
        "log":    parse_log(wb),
        "anchor": None,
        "backAnchor": parse_back_anchor(wb),
        "cycAnchor": parse_cycle_anchor(wb),
        "source": Path(xlsx_path).name,
    }
    # 램프가 실제로 걸려 있는지는 요일 시트 K열이 정한다.
    # 전부 1이면 모든 종목이 1주차부터 나오는 것이므로 앱에 램프를 띄우지 않는다 —
    # 100%만 적힌 표는 매일 보는 화면에서 잡음이다. 시트에는 그대로 남아 있다.
    ret = [e.get("rw", 1) for d in data["home"].values() for e in all_ex(d)]
    if data["ramp"] and ret and max(ret) > 1:
        data["ramp"]["fullWeek"] = max(ret)
    else:
        data["ramp"] = None

    # 휴식일은 주간계획에서 가져와 끼워 넣는다
    for d, content, sets in data["plan"]["rows"]:
        if sets == 0 and d not in data["home"]:
            data["home"][d] = {"t": content or "휴식", "sub": "", "rest": True}
    order = [d for d, _, _ in data["plan"]["rows"]]
    data["homeOrder"] = order or [l for l, _ in home]
    order = [str(i) for i in range(1, CYCLE_LEN + 1)]
    for d in order:
        if d not in data["back"]:
            data["back"][d] = {"t": "휴식", "sub": "", "rest": True}
    data["back"] = {d: data["back"][d] for d in order}
    data["backOrder"] = order
    data["wuPlan"], data["wuIn"] = WU_SESSION, WU_INSESSION
    # 등 그립 회전(2026-11-01) — 회전표 종목을 일차 시트 종목에 붙인다. 이름이 안 맞으면 빌드에서 바로 알린다
    gr = parse_grip_rotation(wb)
    data["gripRot"] = gr
    if gr:
        for d, rot, a, b in gr["order"]:
            flat = (data["back"].get(d) or {}).get("ex") or []
            for nm in (a, b):
                if sum(1 for e in flat if e.get("name") == nm) != 1:
                    sys.exit(f"[중단] {GRIP_SHEET} 회전별 순서 바꾸기 '{d}일차 {nm}' — 그 일차 시트(묶음 밖)에 같은 이름이 하나여야 합니다.")
            if not 1 <= rot <= 3:
                sys.exit(f"[중단] {GRIP_SHEET} 회전별 순서 바꾸기 '{d}일차' — 회전은 1~3이어야 합니다.")
        for d, name, per, grips, memo, arm, cues in gr["rows"]:
            hit = [e for e in all_ex(data["back"].get(d, {})) if e.get("name") == name]
            if not hit:
                sys.exit(f"[중단] {GRIP_SHEET} '{d}일차 {name}' — 그 일차 시트에 같은 이름의 종목이 없습니다(글자까지 같아야 함).")
            for e in hit:
                e["grip"] = {"p": per, "g": grips}
                if arm is not None:
                    e["grip"]["a"] = arm
                if any(cues):
                    e["grip"]["c"] = cues
                if "원암" in name or name in ONEARM_BASE:   # 시트 수행 초가 원암(좌우) 기준인 종목
                    e["grip"]["b"] = 1
                e["gname"] = grip_display_name(name)
    # 사이클 종목 교대(2026-11-02) — 교대표 종목을 일차 시트 종목 자리에 붙인다. 앱이 오늘 사이클로 한쪽을 얹는다
    alts = parse_cycle_alt(wb, wbf)
    data["cycAlt"] = [[a["day"], a["base"], a["v"]["name"], a["side"], a["from"]] for a in alts]
    for a in alts:
        hit = [e for e in all_ex(data["back"].get(a["day"], {})) if e.get("name") == a["base"]]
        if len(hit) != 1:
            sys.exit(f"[중단] {ALT_SHEET} '{a['day']}일차 {a['base']}' — 그 일차 시트에 같은 이름의 종목이 "
                     f"{'없습니다(글자까지 같아야 함)' if not hit else '둘 이상입니다'}.")
        e, v = hit[0], a["v"]
        if e.get("grip") or e.get("cv"):
            sys.exit(f"[중단] {ALT_SHEET} '{a['day']}일차 {a['base']}' — 그립 회전이나 다른 교대가 이미 걸린 종목입니다.")
        if a["side"] != "from" and v["s"] != e["s"]:   # 'N사이클부터'는 한 번 바뀌고 끝이라 세트가 달라도 된다
            sys.exit(f"[중단] {ALT_SHEET} '{v['name']}' {v['s']}세트 ≠ 일차 시트 '{a['base']}' {e['s']}세트 — "
                     "같은 자리는 같은 세트여야 일차 합계가 맞습니다.")
        if v["ss"] is None:
            v["ss"] = e.get("ss", "")
        if a["side"] == "from":
            e["cv"] = {"on": v, "off": {k: e.get(k) for k in ALT_KEYS}}
            e["cvf"] = a["from"]
            continue
        other = "even" if a["side"] == "odd" else "odd"
        e["cv"] = {a["side"]: v, other: {k: e.get(k) for k in ALT_KEYS}}
        # 추적 보조가 사이클 걸러 한 번이 되면 미달 칸 사이에 빈 칸(안 한 사이클)이 끼인다 —
        # 사다리가 앞 두 칸을 보게 한다(엑셀 증량기록도 같은 식 · 2026-11-02)
        for x in (v, e["cv"][other]):
            if x.get("trk") and x.get("lift") in data["config"]["lifts"]:
                data["config"]["lifts"][x["lift"]]["lb"] = 2
    # 보조 추적 종목이 어느 일차에 있는지 — 증량 탭이 일차별로 묶는다. 없는 종목은 '' (추적 안 함)
    for k, L in data["config"]["lifts"].items():
        if L.get("acc"):
            L["day"] = next((d for d in order for e0 in all_ex(data["back"][d])
                             for e in [e0] + list((e0.get("cv") or {}).values())
                             if e.get("lift") == k and e.get("trk")), "")

    # 배포용 사본. 파일명을 index.html로 고정해야 사이트 주소 루트가 열리고,
    # 워크북이 v15, v16으로 올라가도 폰 홈화면 바로가기가 안 깨진다.
    site = Path(xlsx_path).resolve().parent / "docs"
    site.mkdir(exist_ok=True)

    # 말소리는 표보다 먼저 만든다 — 어느 구간이 무슨 말인지가 HTML에 실린다
    data["voice"] = build_voice(site, data)

    html = TEMPLATE.replace("/*__DATA__*/ null",
                            json.dumps(data, ensure_ascii=False))
    Path(out_path).write_text(html, encoding="utf-8")
    (site / "index.html").write_text(html, encoding="utf-8")
    made, version = write_pwa_assets(site, html,
                                     data["voice"] and data["voice"]["src"])

    print(f"완료 → {out_path}")
    print(f"       {site / 'index.html'}  (배포용, 캐시버전 {version})")
    if made:
        print(f"       아이콘 생성: {', '.join(made)}")
    train = sum(1 for _, _, n in data["plan"]["rows"] if n)
    if data.get("gripRot"):
        print(f"  등 그립 회전: {len(data['gripRot']['rows'])}종목 · {data['gripRot']['start']}사이클부터")
    for d, base, name, side, frm in data.get("cycAlt") or []:
        print(f"  사이클 종목 교대: {d}일차 {base} → {name}({frm}사이클부터)" if side == "from" else
              f"  사이클 종목 교대: {d}일차 {base} ↔ {name}({'홀수' if side == 'odd' else '짝수'} 사이클)")
    print(f"  훈련 {train}일 · 사이클당 {data['plan']['total']}세트 "
          f"(주당 환산 {round(data['plan']['total'] * 7 / len(data['plan']['rows']))}) "
          f"· {data['config']['weeks']}사이클 계획")
    if data["ramp"]:
        vw = data["config"]["week"] + RAMP_WEEK_OFFSET
        print("  볼륨램프(종목 복귀): " + " / ".join(
            f"{r['label']} {r['total']}세트" for r in data["ramp"]["rows"]))
        print(f"       중량 {data['config']['week']}주차 · 볼륨 {vw}주차"
              + (f" (앞세움 +{RAMP_WEEK_OFFSET})" if RAMP_WEEK_OFFSET else ""))
    if data["log"]:
        print(f"  증량기록: 엑셀에 {len(data['log'])}사이클 기록 있음 "
              f"({', '.join(str(c) + '사이클' for c in sorted(data['log']))})")
    else:
        print("  증량기록: 엑셀 쪽은 비어 있음 — 폰에서 적은 기록은 앱 증량 탭에서 옮겨 적는다")
    if data["budget"]:
        b = data["budget"]
        print(f"  예산 {b['state']} · 상한 {b['cap']} · 현재 {b['cur']} · 잔여 {b['left']}"
              f" · 실제 주당 {b['weekly']} ({b['cycle']}일 주기)")
        if b["left"] < 0:
            print(f"  [중단 권고] 총량 {abs(b['left'])}세트 초과 — 어딘가에서 빼기 전에는 추가 금지")
    for row in data["check"] + data["backCheck"]:
        if row[4] == "bad":
            print(f"  [주의] {row[0]}: {row[5]}")


# ── HTML 템플릿 ──────────────────────────────────────────────────
TEMPLATE = r"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#15181C">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="운동">
<meta name="robots" content="noindex, nofollow">
<link rel="manifest" href="manifest.webmanifest">
<link rel="apple-touch-icon" href="icon-192.png">
<link rel="icon" type="image/png" sizes="512x512" href="icon-512.png">
<title>운동 프로그램</title>
<script>
/* file:// 로 연 로컬 사본에는 navigator.serviceWorker 가 없으므로 조용히 넘어간다 */
if ("serviceWorker" in navigator) {
  addEventListener("load", function () {
    navigator.serviceWorker.register("sw.js").catch(function () {});
  });
}
</script>
<link rel="stylesheet" as="style" crossorigin href="https://cdn.jsdelivr.net/gh/orioncactus/pretendard@v1.3.9/dist/web/variable/pretendardvariable-dynamic-subset.min.css">
<style>
:root{
  --bg:#15181C; --surface:#1E232A; --raised:#272E37; --line:#333B45;
  --ink:#F0EEE9; --mute:#8B95A1; --dim:#5F6874;
  --sq:#C93A42; --dl:#2B62B0; --bp:#DDA51B; --mp:#2E9161;
  --chrome:#B9C0C8; --pad:16px;
}
*{box-sizing:border-box;-webkit-tap-highlight-color:transparent}
html,body{margin:0;padding:0}
body{background:var(--bg);color:var(--ink);
  font-family:"Pretendard Variable",Pretendard,-apple-system,BlinkMacSystemFont,"Apple SD Gothic Neo","Segoe UI",Roboto,"Noto Sans KR",sans-serif;
  font-size:15px;line-height:1.55;letter-spacing:-.01em;padding-bottom:56px;-webkit-text-size-adjust:100%}
.num{font-variant-numeric:tabular-nums;font-feature-settings:"tnum"}
header{position:sticky;top:0;z-index:40;background:rgba(21,24,28,.94);
  backdrop-filter:saturate(160%) blur(12px);border-bottom:1px solid var(--line);
  padding:calc(env(safe-area-inset-top) + 10px) var(--pad) 0}
.brand{display:flex;align-items:baseline;gap:8px;margin-bottom:10px}
.brand h1{font-size:15px;font-weight:800;margin:0;letter-spacing:-.02em}
.brand span{font-size:11px;color:var(--dim);font-weight:600}
.modes{display:flex;gap:6px;margin-left:auto}
.modes button{font:inherit;font-size:11.5px;font-weight:700;padding:5px 11px;border-radius:999px;
  background:transparent;color:var(--mute);border:1px solid var(--line);cursor:pointer}
.modes button[aria-pressed="true"]{background:var(--ink);color:var(--bg);border-color:var(--ink)}
.weekbar{display:flex;align-items:center;gap:12px;padding-bottom:12px}
.wk-step{width:38px;height:38px;flex:0 0 auto;border-radius:11px;border:1px solid var(--line);
  background:var(--surface);color:var(--ink);font-size:20px;line-height:1;cursor:pointer;
  display:grid;place-items:center;font-family:inherit}
.wk-step:active{background:var(--raised)}
.wk-step[disabled]{opacity:.3}
.wk-read{flex:1;text-align:center}
.wk-read b{font-size:23px;font-weight:800;letter-spacing:-.03em}
.wk-read i{font-style:normal;font-size:11px;color:var(--dim);display:block;margin-top:-2px;letter-spacing:.06em}
.wk-track{display:flex;gap:3px;padding:0 0 10px}
.wk-track span{flex:1;height:3px;border-radius:2px;background:#2A313A}
.wk-track span.on{background:var(--ink)}
nav{display:flex;overflow-x:auto;scrollbar-width:none;border-top:1px solid var(--line);
  margin:0 calc(var(--pad)*-1);padding:0 var(--pad)}
nav::-webkit-scrollbar{display:none}
nav button{font:inherit;font-size:13px;font-weight:700;white-space:nowrap;background:none;
  border:none;border-bottom:2px solid transparent;color:var(--dim);padding:11px 13px;cursor:pointer}
nav button[aria-pressed="true"]{color:var(--ink);border-bottom-color:var(--ink)}
main{padding:18px var(--pad) 0;max-width:640px;margin:0 auto}
.daytitle{font-size:22px;font-weight:800;letter-spacing:-.03em;margin:0 0 3px}
.daysub{font-size:12.5px;color:var(--mute);margin:0 0 18px;line-height:1.5}
.sub-h{font-size:15px;font-weight:800;letter-spacing:-.02em;margin:30px 0 3px}
/* 재적응 구간 알림 — 이 주차에 실제로 수행하는 세트가 설계값과 다르다는 표시 */
.rampbar{display:flex;flex-direction:column;gap:2px;background:var(--surface);
  border:1px solid var(--line);border-left:3px solid var(--bp);border-radius:11px;
  padding:10px 13px;margin:0 0 18px}
.rampbar b{font-size:12.5px;font-weight:800;color:var(--bp)}
.rampbar b span{font-weight:600;color:var(--mute)}
.rampbar span{font-size:11.5px;color:var(--mute);line-height:1.5}
.lift{background:var(--surface);border:1px solid var(--line);border-radius:16px;overflow:hidden;margin-bottom:20px}
.lift-top{padding:15px 16px 13px;display:flex;align-items:flex-end;gap:14px;border-left:4px solid var(--accent)}
.lift-name{flex:1;min-width:0}
.lift-name .eyebrow{font-size:10px;font-weight:800;letter-spacing:.14em;color:var(--accent);margin-bottom:3px}
.lift-name h3{margin:0;font-size:17px;font-weight:800;letter-spacing:-.02em}
.lift-name p{margin:2px 0 0;font-size:12px;color:var(--mute)}
.bigw{text-align:right;line-height:.9}
.bigw b{font-size:42px;font-weight:800;letter-spacing:-.045em}
.bigw i{font-style:normal;font-size:13px;font-weight:700;color:var(--mute);margin-left:2px}
.bar{padding:0 16px 14px;border-left:4px solid var(--accent)}
.bar-label{font-size:10px;font-weight:800;letter-spacing:.12em;color:var(--dim);margin-bottom:7px}
.stack{display:flex;align-items:center;gap:3px;min-height:44px}
.sleeve{width:26px;height:5px;background:var(--chrome);border-radius:2px;opacity:.55}
.pl{border-radius:2px;flex:0 0 auto}
.stack em{font-style:normal;font-size:12px;color:var(--dim);margin-left:6px}
.ramp{border-top:1px solid var(--line);display:grid;grid-template-columns:1fr auto auto;font-size:13.5px}
.ramp div{padding:9px 16px;border-bottom:1px solid rgba(51,59,69,.5)}
.ramp .h{font-size:10px;font-weight:800;letter-spacing:.1em;color:var(--dim);padding-top:10px;padding-bottom:6px}
.ramp .r-w{font-weight:700}
.ramp .r-x{color:var(--mute);text-align:right;font-size:12.5px}
.ramp .skip{color:var(--dim);font-weight:500}
.ramp .work{background:var(--raised)}
.ramp .work .r-w{color:var(--accent)}
.ramp>div:nth-last-child(-n+3){border-bottom:none}
.grouphead{font-size:10.5px;font-weight:800;letter-spacing:.14em;color:var(--dim);margin:22px 0 9px}
.ex{background:var(--surface);border:1px solid var(--line);border-radius:13px;padding:13px 14px;margin-bottom:9px}
.ex.flash{animation:exflash 1.6s ease-out}
@keyframes exflash{0%,35%{border-color:var(--mp);box-shadow:0 0 0 2px var(--mp)}100%{border-color:var(--line);box-shadow:none}}
.ex-h{display:flex;gap:10px;align-items:flex-start}
.ex-n{font-size:11px;font-weight:800;color:var(--dim);width:15px;flex:0 0 auto;padding-top:2px}
.ex-t{flex:1;min-width:0}
.ex-t h4{margin:0;font-size:15.5px;font-weight:700;letter-spacing:-.02em}
.ex-t .mg{font-size:11.5px;color:var(--mute);margin-top:1px}
.ex-t .mg .ss,.ex-t .mg .drop{color:var(--mp);font-weight:800}
.ex-p{text-align:right;flex:0 0 auto}
.ex-p b{font-size:16px;font-weight:800;letter-spacing:-.02em;display:block}
.ex-p span{font-size:11.5px;color:var(--mute)}
/* 이번 주에 쉬는 종목 — 흐리게 두고 언제 돌아오는지만 남긴다 */
.ex.hold{opacity:.5}
/* 설계 세트 0(프로그램에서 빼둔 종목) — 흐름에서 빼고 하단 접기 안에 모은다 */
.holdbox{margin-top:12px;padding:11px 14px;border:1px dashed var(--line);border-radius:13px}
.holdbox>summary{color:var(--mute)}
.holdbox>summary i{font-style:normal;font-weight:600;color:var(--dim)}
.holdbox[open]>summary{margin-bottom:11px}
.holdbox .ex{margin-bottom:0}
.holdbox .ex+.ex{margin-top:8px}
.ex-p b.hold-w{font-size:12px;font-weight:800;letter-spacing:0;color:var(--bp)}
.ex-memo{font-size:12.5px;color:var(--mute);margin-top:9px;padding-top:9px;border-top:1px solid var(--line);line-height:1.5}
details{margin-top:9px}
details summary{list-style:none;font-size:11.5px;font-weight:700;color:var(--dim);cursor:pointer;display:flex;align-items:center;gap:5px}
details summary::-webkit-details-marker{display:none}
details summary::after{content:"＋";font-size:12px}
details[open] summary::after{content:"−"}
details p{margin:7px 0 0;font-size:12.5px;color:var(--mute);line-height:1.6}
.total{display:flex;justify-content:space-between;align-items:baseline;padding:12px 14px;
  border:1px dashed var(--line);border-radius:13px;margin-top:12px}
.total span{font-size:11px;font-weight:800;letter-spacing:.12em;color:var(--dim)}
.total b{font-size:19px;font-weight:800}
.notes{margin:22px 0 0}
.notes h5{font-size:12.5px;font-weight:800;margin:0 0 8px}
.notes ul{margin:0 0 18px;padding:0;list-style:none}
.notes li{font-size:12.5px;color:var(--mute);line-height:1.65;padding-left:12px;position:relative;margin-bottom:5px}
.notes li::before{content:"";position:absolute;left:0;top:9px;width:4px;height:1px;background:var(--dim)}
.tbl{width:100%;border-collapse:collapse;font-size:12.5px;margin-bottom:6px}
.tbl th{font-size:10px;font-weight:800;letter-spacing:.08em;color:var(--dim);text-align:right;
  padding:0 0 8px;border-bottom:1px solid var(--line);white-space:nowrap}
.tbl th:first-child{text-align:left}
.tbl td{padding:9px 0;border-bottom:1px solid rgba(51,59,69,.45);text-align:right;white-space:nowrap}
.tbl td:first-child{text-align:left;white-space:normal;font-weight:600;padding-right:8px}
.tbl td.wrap{white-space:normal;text-align:left;color:var(--mute);font-size:11.5px;padding-left:8px}
/* 오른쪽 정렬 열 바로 뒤에 오는 왼쪽 정렬 머리글은 여백이 없으면 앞 글자와 붙는다 */
.tbl th.wrap{text-align:left;padding-left:8px}
.judge{font-size:10.5px;font-weight:800;padding:2px 7px;border-radius:999px;border:1px solid;
  white-space:nowrap;display:inline-block}
.j-ok{color:var(--mp);border-color:rgba(46,145,97,.45)}
.j-lo{color:var(--bp);border-color:rgba(221,165,27,.45)}
.j-bad{color:var(--sq);border-color:rgba(201,58,66,.45)}
.budget{display:flex;flex-wrap:wrap;gap:4px 14px;align-items:baseline;
  border:1px solid var(--line);border-radius:10px;padding:10px 12px;margin:0 0 12px}
.budget b{font-size:13px;letter-spacing:.04em}
.budget span{font-size:12px;color:var(--mute)}
.budget span i{font-style:normal;color:var(--ink);font-weight:700}
.budget.ok{border-color:rgba(46,145,97,.45)}
.budget.ok b{color:var(--mp)}
.budget.bad{border-color:rgba(201,58,66,.55);background:rgba(201,58,66,.07)}
.budget.bad b{color:var(--sq)}
.budget .warn{flex-basis:100%;font-size:12px;color:var(--sq);font-weight:700}
/* 달력·주별 세션에서 쓰는 경고와 참고 — 글자만 굵게, 배경은 부모가 쥔다 */
.warn{margin-top:8px;font-size:12px;line-height:1.5;color:var(--sq);font-weight:700}
.cal-note{margin-top:8px;font-size:12px;line-height:1.5;color:var(--mute)}
.footnote{font-size:11.5px;color:var(--dim);margin-top:18px;line-height:1.6}
.rest{opacity:.42}
[data-lift]{cursor:pointer}
.ex-w[data-exkey]{border-bottom:1px dotted var(--dim);cursor:pointer;padding-bottom:1px}
.ex-w.edited{color:var(--bp);border-bottom-color:var(--bp)}
.ex-go.mic{display:flex;flex-wrap:wrap;gap:6px;align-items:center}
.mic-cap{flex:1 1 180px;font-size:11px;line-height:1.45;color:var(--dim)}
#tm-rec .mic-cap{flex-basis:100%;text-align:center}
.mg .wu{color:var(--sq)}
.grip{margin-top:4px;font-size:13px;font-weight:700;color:var(--bp)}
.grip i{font-style:normal;font-weight:600;color:var(--mute)}
.tm-grip{display:block;margin-top:6px;font-size:18px;font-weight:800;color:var(--bp)}
.tm-cue{display:block;margin-top:4px;font-size:15px;font-weight:700;color:var(--sq)}
.grip.cue{color:var(--sq)}
.tm-wu{display:inline-block;margin-top:8px;padding:6px 12px;border-radius:14px;font-size:13.5px;line-height:1.4;
  max-width:100%;text-align:center;font-weight:800;color:var(--bg);background:var(--sq)}
.mic-dn{padding-left:10px;padding-right:10px}
/* 2026-11-10 무릎 경고(노랑) — 카드 한 줄 · 타이머 띠 */
.kn{margin-top:4px;font-size:13px;font-weight:700;color:var(--bp)}
.tm-kn{display:inline-block;margin-top:8px;padding:6px 12px;border-radius:14px;font-size:13.5px;line-height:1.4;
  max-width:100%;text-align:center;font-weight:800;color:#1a1400;background:#f2c200}
.wk-ed{margin-left:6px;cursor:pointer;border-style:solid;vertical-align:middle}
.wk-ed.edited{color:var(--bp);border-color:var(--bp)}
.wk-ed .w-input{width:56px;font-size:12px;padding:2px 5px}
#tm-wk{margin-top:10px}
/* 세트 기록 (2026-11-13) */
details.slog{margin:8px 0 0;border-top:1px dashed var(--line);padding-top:6px}
details.slog summary{font-size:12px;font-weight:700;color:var(--mute);cursor:pointer}
details.slog summary b{color:var(--ink)}
.sl-row{display:flex;align-items:center;gap:6px;margin-top:6px;font-size:13px}
.sl-n{width:18px;text-align:right;font-weight:800;color:var(--mute);font-variant-numeric:tabular-nums}
.sl-row input{font:inherit;font-size:15px;padding:5px 7px;border:1px solid var(--line);border-radius:7px;background:var(--surface);color:var(--ink);font-variant-numeric:tabular-nums}
.sl-w{width:110px}.sl-r{width:62px}
.sl-row i{font-style:normal;color:var(--mute)}
.sl-ok{font:inherit;font-weight:800;min-width:42px;padding:5px 0;border-radius:7px;border:1px solid var(--line);background:transparent;color:var(--mute);cursor:pointer}
.sl-ok.on{background:var(--mp);border-color:var(--mp);color:var(--bg)}
.sl-date{font-size:12px;color:var(--mute);margin:8px 0}
.sl-prev{font-size:12.5px;line-height:1.45;margin:8px 0 0;color:var(--ink);font-variant-numeric:tabular-nums}
.sl-prev b{color:var(--bp);font-weight:800}
.sl-prev span{color:var(--mute);font-size:11.5px}
.sl-prev.dim{color:var(--dim);margin-top:3px;font-size:11.5px}
.sl-prev.dim b{color:var(--dim)}
.sl-prev.tm{text-align:center;margin:4px 0 2px;font-size:13px;color:var(--mute)}
.sl-prev.tm b{color:var(--ink)}
#tm-slog{margin-top:10px;display:flex;flex-direction:column;align-items:center}
#tm-slog .sl-row{justify-content:center}
.tm-left{display:flex;align-items:center;gap:8px;min-width:0}
#tm-ctx{min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}
/* 2026-11-13 나가기 = 다이얼 오른쪽 아래(사용자 지정 자리 · ✕와 멀리) */
.tm-dialwrap{position:relative;width:100%;display:flex;justify-content:center}
#tm-min{position:absolute;right:0;bottom:0;font:inherit;font-size:13px;font-weight:800;color:var(--ink);background:var(--surface);
  border:1px solid var(--line);border-radius:999px;padding:7px 11px;cursor:pointer}
#tm-mini{position:fixed;left:16px;right:16px;bottom:14px;z-index:150;font:inherit;font-size:14px;font-weight:800;
  padding:13px 16px;border:0;border-radius:999px;background:var(--ink);color:var(--bg);box-shadow:0 4px 18px rgba(0,0,0,.25);
  font-variant-numeric:tabular-nums;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;cursor:pointer}
#tm-mini[hidden]{display:none}
.cal-cell{position:relative}
.cal-cell .sl-dot{position:absolute;top:4px;right:5px;width:6px;height:6px;border-radius:50%;background:var(--mp)}
.bigw.edited b{color:var(--bp)}
.ovnote{display:block;font-size:10px;font-weight:700;color:var(--bp);margin-top:5px;line-height:1.3}
.w-input{font:inherit;font-weight:700;background:var(--raised);color:var(--ink);
  border:1px solid var(--dim);border-radius:8px;padding:5px 8px;width:96px;
  text-align:right;font-size:16px}
.bigw .w-input{width:120px;font-size:22px}
.ovbox{margin-top:22px;border:1px dashed var(--bp);border-radius:13px;padding:13px 14px}
.ovbox h5{font-size:12.5px;font-weight:800;margin:0;color:var(--bp)}
.ovbox h5 span{font-weight:600;color:var(--mute);font-size:11px}
.ovbox ul{margin:8px 0 11px;padding:0;list-style:none}
.ovbox li{font-size:12.5px;color:var(--mute);line-height:1.65}
.ovbox button{font:inherit;font-size:12px;font-weight:700;padding:7px 12px;border-radius:9px;
  background:transparent;color:var(--sq);border:1px solid var(--sq);cursor:pointer}
.ovbox button.plain{color:var(--ink);border-color:var(--line);background:var(--raised)}
.log-text{margin-top:10px}
.log-text summary{font-size:11.5px;color:var(--dim);cursor:pointer}
.log-text pre{font:inherit;font-size:12px;color:var(--mute);white-space:pre-wrap;word-break:break-all;
  margin:8px 0 0;padding:9px 10px;background:var(--bg);border-radius:8px;user-select:text;-webkit-user-select:text}

/* ── 증량 기록 ──────────────────────────────────────────────────── */
.logcard{background:var(--surface);border:1px solid var(--line);border-radius:16px;padding:2px 14px 4px;margin-bottom:18px}
.logcard-h{display:flex;align-items:baseline;gap:8px;padding:12px 0 7px}
.logcard-h b{font-size:17px;font-weight:800;letter-spacing:-.02em}
.logcard-h span{font-size:12px;color:var(--mute)}
.log-row{display:flex;align-items:center;justify-content:space-between;gap:10px;padding:10px 0;
  border-top:1px solid rgba(51,59,69,.45)}
.log-l{flex:1;min-width:0}
.log-l b{display:block;font-size:14px;font-weight:700;letter-spacing:-.02em}
.log-l span{display:block;font-size:11.5px;color:var(--mute);margin-top:1px;line-height:1.4}
.log-l em{display:block;font-style:normal;font-size:11px;font-weight:700;color:var(--sq);margin-top:3px;line-height:1.4}
.seg{display:flex;flex-wrap:wrap;justify-content:flex-end;align-items:center;gap:6px;flex:0 0 auto;max-width:58%}
.seg button{font:inherit;font-size:12.5px;font-weight:700;padding:7px 11px;border-radius:9px;
  background:var(--raised);color:var(--mute);border:1px solid var(--line);cursor:pointer}
.seg button.on{color:var(--bg);border-color:var(--ink);background:var(--ink)}
.seg button.on.ok{color:#fff;border-color:var(--mp);background:var(--mp)}
.seg button.on.bad{color:#fff;border-color:var(--sq);background:var(--sq)}
.seg.fat{gap:4px}
.seg.fat button{padding:7px 0;width:32px;text-align:center}
.log-in{font:inherit;font-size:14px;font-weight:700;background:var(--raised);color:var(--ink);
  border:1px solid var(--line);border-radius:9px;padding:7px 9px;width:118px;text-align:right}
.log-in.wide{width:50%;text-align:left;font-weight:600}
.lm{display:block;font-style:normal;font-size:10px;font-weight:800;margin-top:2px;line-height:1.2;white-space:normal}
.lm.ok{color:var(--mp)}
.lm.bad{color:var(--sq)}
.tbl tr[data-logrow]{cursor:pointer}
.lift-log{padding:11px 16px 13px;border-top:1px solid var(--line);display:flex;align-items:center;
  justify-content:space-between;gap:10px;background:var(--surface)}
.lift-log-h{flex:1;min-width:0;font-size:12px;font-weight:800;color:var(--mute);letter-spacing:.02em;line-height:1.4}
.lift-log-h span{font-weight:600;color:var(--dim)}
.lift-log-h em{display:block;font-style:normal;font-size:11px;color:var(--sq);margin-top:3px;font-weight:700}

/* ── 달력 ──────────────────────────────────────────────────────── */
.cal-bar{display:flex;align-items:center;gap:8px;margin:2px 0 12px}
.cal-bar h3{flex:1;margin:0;font-size:17px;font-weight:800;letter-spacing:-.02em}
.cal-bar button{font:inherit;font-size:13px;font-weight:700;padding:7px 12px;border-radius:9px;
  background:var(--surface);color:var(--ink);border:1px solid var(--line);cursor:pointer}
.cal-grid{display:grid;grid-template-columns:repeat(7,1fr);gap:3px}
.cal-dow{text-align:center;font-size:11px;font-weight:800;color:var(--dim);padding:2px 0 4px}
.cal-dow.sun{color:var(--sq)}
.cal-cell{min-height:62px;border-radius:8px;padding:5px 4px;background:var(--surface);
  border:1px solid transparent;border-left:3px solid var(--acc,var(--line));
  display:flex;flex-direction:column;gap:1px;cursor:pointer;overflow:hidden}
.cal-cell.pad,.cal-cell.blank{background:transparent;border-color:transparent;cursor:default}
.cal-cell.blank b{color:var(--dim);font-weight:600}
.cal-cell.rest{background:transparent;border-left-color:var(--line);border:1px dashed var(--line)}
.cal-cell.today{border-color:var(--ink);background:var(--raised)}
.cal-cell b{font-size:13px;font-weight:800;font-variant-numeric:tabular-nums;line-height:1.1}
.cal-cell.sun b{color:var(--sq)}
.cal-cell i{font-style:normal;font-size:9.5px;font-weight:800;color:var(--acc,var(--mute));
  letter-spacing:-.02em}
.cal-cell span{font-size:9.5px;line-height:1.25;color:var(--mute);word-break:keep-all}
.cal-cell em{font-style:normal;font-size:9px;font-weight:800;color:var(--dim);margin-top:auto}
.cal-key{display:flex;flex-wrap:wrap;gap:5px 12px;margin:13px 2px 0}
.cal-key span{display:inline-flex;align-items:center;gap:5px;font-size:11px;
  font-weight:700;color:var(--mute)}
.cal-key i{width:9px;height:9px;border-radius:3px;flex:none}
.cal-key b{font-weight:700;color:var(--dim)}
.cal-now{background:var(--surface);border:1px solid var(--line);border-radius:12px;
  padding:13px 14px;margin:0 0 13px}
.cal-now b{display:block;font-size:19px;font-weight:800;letter-spacing:-.02em}
.cal-now p{margin:3px 0 0;font-size:12.5px;color:var(--mute)}
.cal-fix{margin-top:11px}
.cal-fix summary{font-size:12px;font-weight:700;color:var(--mute);cursor:pointer}
.cal-days{display:flex;flex-wrap:wrap;gap:6px;margin-top:9px}
.cal-days button{font:inherit;font-size:12px;font-weight:700;padding:8px 11px;border-radius:9px;
  background:var(--raised);color:var(--ink);border:1px solid var(--line);cursor:pointer}
.cal-days button.on{background:var(--ink);color:var(--bg);border-color:var(--ink)}

/* ── 인터벌 타이머 ─────────────────────────────────────────────── */
.tm-go{display:inline-flex;align-items:center;gap:7px;font:inherit;font-size:12px;font-weight:700;
  padding:6px 12px;border-radius:999px;background:var(--raised);color:var(--ink);
  border:1px solid var(--line);cursor:pointer}
.tm-go b{font-variant-numeric:tabular-nums;color:var(--mute);font-weight:600}
.ex-go{margin:10px 0 0}
.tm-none{display:inline-block;font-size:11.5px;font-weight:600;color:var(--dim);
  padding:5px 10px;border-radius:999px;border:1px dashed var(--line)}
.day-go i{font-style:normal;font-weight:600;opacity:.8}
.day-go.wu-go{background:var(--surface);color:var(--sq);border:1.5px solid var(--sq);margin-bottom:10px}
.day-go{display:block;width:100%;margin:0 0 15px;padding:14px;border:0;border-radius:13px;
  background:var(--mp);color:#fff;font:inherit;font-size:15px;font-weight:800;cursor:pointer}
#tm{position:fixed;inset:0;z-index:200;background:var(--bg);display:flex;flex-direction:column;
  padding:calc(env(safe-area-inset-top) + 12px) 18px calc(env(safe-area-inset-bottom) + 18px)}
#tm[hidden]{display:none}
.tm-top{display:flex;align-items:center;justify-content:space-between;gap:10px;
  font-size:12.5px;font-weight:700;color:var(--mute)}
.tm-top button{font:inherit;font-size:19px;line-height:1;background:none;border:0;
  color:var(--mute);padding:2px 6px;cursor:pointer}
.tm-btns{display:flex;align-items:center;gap:2px;flex:none}
#tm-mute.off{opacity:.4}
.tm-body{flex:1;display:flex;flex-direction:column;align-items:center;justify-content:safe center;
  text-align:center;min-height:0;overflow-y:auto;overscroll-behavior:contain}
.tm-body>*{flex:none}
#tm-name{margin:0;font-size:25px;font-weight:800;line-height:1.25;letter-spacing:-.02em}
#tm-meta{margin:5px 0 18px;font-size:14px;color:var(--mute);font-variant-numeric:tabular-nums}
#tm-meta .tm-w,.tm-up .tm-w{border-bottom:1px dotted currentColor;cursor:pointer;padding-bottom:1px}
#tm-meta .tm-w.edited,.tm-up .tm-w.edited{color:var(--bp);border-bottom-color:var(--bp)}
#tm-meta .w-input{width:106px;font-size:15px;padding:4px 7px}
.tm-up .w-input{width:88px;font-size:13px;padding:3px 6px}
.tm-dial{width:min(72vw,286px,32vh);aspect-ratio:1;border-radius:50%;display:flex;flex-direction:column;
  align-items:center;justify-content:center;background:var(--surface);
  border:7px solid var(--phase,var(--line));transition:border-color .25s}
.tm-dial span{font-size:15px;font-weight:800;letter-spacing:.03em;color:var(--phase,var(--mute))}
.tm-dial b{font-size:clamp(46px,15vw,66px);font-weight:800;letter-spacing:-.03em;
  font-variant-numeric:tabular-nums;line-height:1.15}
#tm-set{margin:18px 0 0;font-size:15px;font-weight:700;color:var(--mute)}
.tm-ctl{display:grid;grid-template-columns:1fr 1.25fr 1fr;gap:8px;margin-top:10px}
.tm-ctl button{font:inherit;font-size:14px;font-weight:700;padding:16px 4px;border-radius:12px;
  background:var(--surface);color:var(--ink);border:1px solid var(--line);cursor:pointer}
.tm-ctl button.primary{background:var(--ink);color:var(--bg);border-color:var(--ink)}
.tm-up{margin:13px 0 0;text-align:center;font-size:12.5px;color:var(--dim);
  line-height:1.55;min-height:3.2em}   /* 두 줄분을 늘 비워 둔다 — 종목이 바뀔 때
                                          한 줄이 늘어나며 다이얼이 들썩이지 않게 */
.tm-up b{color:var(--mute);font-weight:700;font-size:13.5px}
#tm-rec{margin-top:12px;min-height:38px;display:flex;flex-direction:column;align-items:center;gap:6px}
#tm-rec .seg{justify-content:center;max-width:none}
#tm-rec .ex-go{margin:0;justify-content:center}
.tm-rec-h{font-size:12px;font-weight:700;color:var(--mute)}
.tm-up i{display:block;font-style:normal;font-variant-numeric:tabular-nums}
</style>
</head>
<body>
<header>
  <div class="brand">
    <h1>운동 프로그램</h1><span class="num" id="ver"></span>
    <div class="modes">
      <button id="m-home" aria-pressed="true">주6일</button>
      <button id="m-back" aria-pressed="false">9일</button>
    </div>
  </div>
  <div class="weekbar" id="weekbar">
    <button class="wk-step" id="wk-down" aria-label="이전 사이클">−</button>
    <div class="wk-read"><b class="num" id="wk-num">1</b><i id="wk-of">CYCLE</i></div>
    <button class="wk-step" id="wk-up" aria-label="다음 사이클">+</button>
  </div>
  <div class="wk-track" id="wk-track"></div>
  <nav id="nav"></nav>
</header>
<main id="app"></main>
<div id="tm" hidden>
  <div class="tm-top"><span class="tm-left"><span id="tm-ctx"></span></span><span class="tm-btns"
    ><button id="tm-mute" aria-label="음성 켜기/끄기">&#128266;</button
    ><button id="tm-x" aria-label="타이머 닫기">&#10005;</button></span></div>
  <div class="tm-body">
    <h3 id="tm-name"></h3>
    <p id="tm-meta"></p>
    <div class="tm-dialwrap"><div class="tm-dial" id="tm-dial"><span id="tm-phase"></span><b id="tm-time">00:00</b></div
      ><button id="tm-min" aria-label="타이머는 돌려 둔 채 잠깐 나가기">&#8601; 나가기</button></div>
    <p id="tm-set"></p>
    <div id="tm-wk"></div>
    <div id="tm-rec"></div>
    <div id="tm-slog"></div>
  </div>
  <div class="tm-ctl">
    <button id="tm-prev">&#9664; 이전</button>
    <button id="tm-play" class="primary">일시정지</button>
    <button id="tm-next">건너뛰기 &#9654;</button>
  </div>
  <p class="tm-up" id="tm-up"></p>
</div>
<button id="tm-mini" hidden></button>
<script>
const D = /*__DATA__*/ null;

/* 중량 격자. 헬스장 작은 원판이 2.5 · 1 · 0.5kg 각 2개 = 한쪽에 하나씩뿐이라
   한쪽 끝수는 0 / 0.5 / 1 / 1.5 / 2.5 / 3 / 3.5 / 4 만 만들 수 있다.
   그래서 총 중량이 5로 나눈 나머지 4(24 · 29 · 34 …)면 못 끼운다 — 1kg 올린다.
   엑셀의 MROUND(...,1)+(MOD(...,5)=4)와 같은 값을 내야 한다. 어긋나면 폰과 시트가 갈린다 */
const R = v => { const n=Math.round(v); return n%5===4 ? n+1 : n; };
const fmt = v => (Math.round(v*1000)/1000).toString();
const WEEKS = D.config.weeks;
/* 표가 6주차에서 시작한다 — '몇 번째 줄'과 '몇 주차'가 다르므로 첫 줄 번호를 따로 쥔다 */
const WBASE = D.config.base||1, WLAST = WBASE+WEEKS-1;
const LIFT = D.config.lifts;

/* ── 사이클 종목 교대 (2026-11-02 · 엑셀 '사이클_종목교대') ──────────────
   같은 자리 · 같은 세트에서 사이클마다 종목만 바뀐다(8일차 인클라인 바벨 ↔ 정지 벤치). 종목 객체의 cv에
   양쪽(odd · even)이 다 실려 있고, 오늘 사이클 쪽을 객체에 그대로 얹는다 — 카드 · 타이머 · 음성 · 덮어쓰기가
   전부 얹힌 이름과 중량을 본다. 사이클 번호는 그립 회전과 같은 셈(1일차 기준일에서 9일씩 · 디로드도 한 번).
   덮어쓰기 정리보다 먼저 돌아야 해서 아래 달력 함수를 못 쓰고 날짜 셈을 따로 둔다(기준일 키 wk-canchor 같음) */
/* 시간표 탭(2026-11-05)이 다른 사이클을 계산하는 동안만 잠깐 그 사이클 번호를 쥔다 — 평소엔 null(오늘) */
let VIEW_R=null;
function altRound(){
  if(VIEW_R!=null) return VIEW_R;
  let a=(D.cycAnchor&&D.cycAnchor.date)||null;
  try{ const o=localStorage.getItem("wk-canchor"); if(o) a=o; }catch(e){}
  if(!a) return null;
  const p=a.split("-").map(Number), t=new Date();
  const n=Math.round((new Date(t.getFullYear(),t.getMonth(),t.getDate())-new Date(p[0],p[1]-1,p[2]))/86400000);
  return n<0 ? null : Math.floor(n/D.backOrder.length)+1;
}
let ALT_ON=null;                     /* 지금 얹힌 쪽 — "odd" · "even" (기준일 전이면 even = 일차 시트 그대로) */
const altEach = f => { for(const d in (D.back||{})){ const day=D.back[d]; if(!day||day.rest) continue;
  (day.ex||[]).concat(day.groups?day.groups.reduce((a,g)=>a.concat(g[1]),[]):[]).forEach(e=>{ if(e.cv) f(e,d); }); } };
/* 2026-11-10 'N사이클부터'(e.cvf) — 번갈지 않고 그 사이클부터 on(4일차 핵 스쿼트 복귀). 종목마다 쪽이 달라 각자 e.cvs에 쥔다 */
const altSide = (e,r) => e.cvf ? ((r!=null&&r>=e.cvf)?"on":"off") : ((r!=null&&r%2===1)?"odd":"even");
function applyAlt(){
  const r=altRound(); let ch=false;
  ALT_ON=(r!=null&&r%2===1)?"odd":"even";
  altEach(e=>{ const sd=altSide(e,r); if(e.cvs!==sd){ e.cvs=sd; Object.assign(e,e.cv[sd]); ch=true; } });
  return ch;
}
applyAlt();
/* ── 회전별 순서 바꾸기 (2026-11-04 · 사이클_그립회전 아래 표) ── 그 회전에만 두 종목 자리와 번호를 맞바꾼다.
   일차 시트 순서(ex0)를 원본으로 쥐고 매번 거기서 다시 만든다 — 사이클이 넘어가면 저절로 원래대로 */
function applyOrd(){
  const G=D.gripRot, O=(G&&G.order)||[]; if(!O.length) return;
  const r=altRound(), rot=(r!=null&&r>=G.start)?((r-G.start)%3)+1:null;
  const days={}; O.forEach(o=>{ (days[o[0]]=days[o[0]]||[]).push(o); });
  for(const d in days){ const day=D.back[d]; if(!day||day.rest) continue;
    if(!day.ex0){ day.ex0=day.ex.slice(); day.ex0.forEach(e=>{ e.n0=e.n; }); }
    const ex=day.ex0.slice(); ex.forEach(e=>{ e.n=e.n0; });
    days[d].forEach(o=>{ if(o[1]!==rot) return;
      const i=ex.findIndex(e=>e.name===o[2]), j=ex.findIndex(e=>e.name===o[3]); if(i<0||j<0) return;
      const t=ex[i]; ex[i]=ex[j]; ex[j]=t; const n=ex[i].n; ex[i].n=ex[j].n; ex[j].n=n; });
    day.ex=ex; }
}
applyOrd();
/* 오늘 안 얹힌 쪽 이름의 덮어쓰기(수행 초 등)는 지우지 않는다 — 다음 사이클에 다시 쓴다 */
function altIdle(key){
  const p=key.split("|"); if(p[0]!=="back"||!ALT_ON) return false;
  const day=D.back[p[1]]; if(!day||day.rest) return false;
  const oth=e=>e.cvf?(e.cvs==="on"?"off":"on"):(ALT_ON==="odd"?"even":"odd");
  return (day.ex||[]).concat(day.groups?day.groups.reduce((a,g)=>a.concat(g[1]),[]):[])
    .some(e=>e.cv&&e.cv[oth(e)].name===p[2]&&e.name!==p[2]);
}

/* ── 이 기기에서만 적용되는 중량 덮어쓰기 ─────────────────────────
   엑셀 원본은 건드리지 않는다. {수정값, 그때의 엑셀값}을 localStorage에 두고,
   엑셀 쪽 값이 바뀐 채 재배포되면 해당 덮어쓰기는 자동으로 버린다. */
const OVK = "wk-ov";
let OV = {lifts:{}, ex:{}};
try{ const o=JSON.parse(localStorage.getItem(OVK)||"null"); if(o&&o.lifts&&o.ex) OV=o; }catch(e){}
/* 2026-11-04 등 종목 이름에서 그립 · 원암 표기를 뗐다(그립 회전 시작). 폰에 적어 둔 무게 · 수행 초는
   '모드|일차|이름' 키라 이름이 바뀌면 아래 정리에서 버려진다 — 그 전에 새 이름 키로 옮긴다. 새 키에 이미 값이 있으면 그대로 둔다 */
const RENAMED={"back|3|랫풀다운 (뉴트럴 그립)":"back|3|랫풀다운","back|3|해머 스트랭스 로우(오버그립)":"back|3|해머 스트랭스 로우",
  "back|3|인클라인 덤벨 로우(오버그립)":"back|3|인클라인 덤벨 로우","back|3|원암 케이블 로우":"back|3|케이블 로우",
  "back|5|랫풀다운 (와이드 오버그립)":"back|5|랫풀다운","back|8|랫풀다운(언더그립)":"back|8|랫풀다운",
  "back|8|하이로우 머신(언더그립)":"back|8|하이로우 머신",
  "back|5|시티드 바벨 밀리터리 프레스":"back|5|오버헤드 프레스 (스탠딩 바벨)"};
function renameKeys(o){ let ch=false;
  for(const k in RENAMED){ if(o&&o[k]){ if(!o[RENAMED[k]]) o[RENAMED[k]]=o[k]; delete o[k]; ch=true; } }
  return ch; }
if(renameKeys(OV.ex)) try{ localStorage.setItem(OVK, JSON.stringify(OV)); }catch(e){}
function exByKey(key){
  const p=key.split("|"), day=(p[0]==="home"?D.home:D.back)[p[1]];
  if(!day||day.rest) return null;
  const all=(day.ex||[]).concat(day.groups?day.groups.reduce((a,g)=>a.concat(g[1]),[]):[]);
  return all.find(x=>x.name===p[2])||null;
}
const saveOV = () => { try{ localStorage.setItem(OVK, JSON.stringify(OV)); }catch(e){} };
let ovStale=false;
for(const k in OV.lifts){ if(!LIFT[k]||LIFT[k].start!==OV.lifts[k].base){ delete OV.lifts[k]; ovStale=true; } }
for(const k in OV.ex){ if(altIdle(k)) continue; const e=exByKey(k); if(!e||e.w!==OV.ex[k].base){ delete OV.ex[k]; ovStale=true; } }
if(ovStale) saveOV();
/* ── 수행 시간(초) 덮어쓰기 — 2026-10-25 사용자 요청(촉박한 종목을 엑셀 안 열고 늘리기).
   무게와 같은 규칙이다: {폰 값, 그때의 엑셀 G열 값}. 엑셀 쪽이 바뀐 채 재배포되면 버린다.
   불러올 때 종목 객체의 wk에 바로 얹는다 — 카드 · 타이머가 전부 그 값을 쓴다(wk0 = 엑셀 값) */
const OVTK="wk-ovt";
let OVT={};
try{ const o=JSON.parse(localStorage.getItem(OVTK)||"null"); if(o&&typeof o==="object") OVT=o; }catch(e){}
if(renameKeys(OVT)) try{ localStorage.setItem(OVTK, JSON.stringify(OVT)); }catch(e){}
const saveOVT = () => { try{ localStorage.setItem(OVTK, JSON.stringify(OVT)); }catch(e){} };
(function(){ let ch=false;
  ["home","back"].forEach(m=>{ const days=D[m]||{};
    for(const t in days){ const d=days[t]; if(!d||d.rest) continue;
      (d.ex||[]).concat(d.groups?d.groups.reduce((a,g)=>a.concat(g[1]),[]):[]).forEach(e=>{ e.wk0=e.wk; }); } });
  for(const k in OVT){ if(altIdle(k)) continue; const e=exByKey(k);
    if(!e||!e.wk0||e.wk0!==OVT[k].base){ delete OVT[k]; ch=true; } else e.wk=OVT[k].wk; }
  if(ch) saveOVT(); })();
function setWk(key,v){
  const e=exByKey(key); if(!e||!e.wk0) return;
  v=Math.max(5,Math.round(v));
  if(v===e.wk0) delete OVT[key]; else OVT[key]={wk:v, base:e.wk0};
  e.wk=v; saveOVT();
}
const startOf = k => OV.lifts[k] ? OV.lifts[k].start : LIFT[k].start;

/* ── 증량 기록 — 이 폰에 남는 실측 ────────────────────────────────
   엑셀 증량기록과 같은 규칙이다: 공란·O는 다음 사이클에 증량, 미달 반복수(또는 X)를
   적으면 그 중량을 한 사이클 더 든다. 엑셀에 적힌 값(D.log)이 바탕이고 폰에서 적은
   값이 그 위에 얹힌다. 엑셀과 같은 값은 폰에 따로 두지 않는다 — 옮겨 적고 나면
   폰 쪽은 저절로 비고, 남은 것만 '이 폰에만 있는 기록'으로 뜬다. */
const LGK="wk-log";
let LOG={};
try{ const o=JSON.parse(localStorage.getItem(LGK)||"null"); if(o&&typeof o==="object") LOG=o; }catch(e){}
const XL = D.log||{};
const LOGF={fat:"다음날 피로",joint:"관절 이상",strap:"스트랩",memo:"메모"};
const xlVal = (c,f) => (XL[c]&&XL[c][f]!=null) ? String(XL[c][f]) : "";
function saveLog(){ try{ localStorage.setItem(LGK,JSON.stringify(LOG)); }catch(e){} }
(function(){ let ch=false;
  for(const c in LOG){
    for(const f in LOG[c]) if(LOG[c][f]===xlVal(c,f)){ delete LOG[c][f]; ch=true; }
    if(!Object.keys(LOG[c]).length){ delete LOG[c]; ch=true; } }
  if(ch) saveLog(); })();
function logVal(c,f){ const l=LOG[c]&&LOG[c][f]; return l!==undefined ? l : xlVal(c,f); }
function setLog(c,f,v){
  v=String(v==null?"":v).trim();
  if(!LOG[c]) LOG[c]={};
  if(v===xlVal(c,f)) delete LOG[c][f]; else LOG[c][f]=v;
  if(!Object.keys(LOG[c]).length) delete LOG[c];
  saveLog();
}
/* ── 세트 기록 (2026-11-13 사용자 요청) ─────────────────────────────
   날짜마다 · 종목마다 · 세트마다 무게 · 횟수 · 달성(✓). 이 폰에만 저장된다 —
   'PC로 보낼 세트 기록 복사'로 텍스트를 뽑아 Claude에게 붙여 넣으면 PC의 세트기록.csv에 쌓인다.
   ✓만 누르고 횟수를 비우면 '목표 범위 달성'이다. 무게를 비우고 ✓를 누르면 그날 계획 무게가 들어간다 */
const SKEY="wk-sets"; let SETS={};
try{ const o=JSON.parse(localStorage.getItem(SKEY)||"null"); if(o&&typeof o==="object") SETS=o; }catch(e){}
function sSave(){ try{ localStorage.setItem(SKEY,JSON.stringify(SETS)); }catch(e){} }
let LOGDATE=null;              /* 달력 칸을 눌러 들어오면 {day, date} */
const SOPEN=new Set();         /* 펼쳐 둔 종목 — 다시 그려도 펼친 채로 */
const sGet = (date,name) => { const x=SETS[date]; return (x&&x.ex&&x.ex[name])||[]; };
function sPut(date,day,name,i,patch){
  const x=SETS[date]||(SETS[date]={ex:{}}); x.day=day;
  const c=cycDayOf(date); if(c) x.round=c.round;
  const L=x.ex[name]||(x.ex[name]=[]); while(L.length<=i) L.push({});
  Object.assign(L[i],patch);
  for(const f in L[i]) if(L[i][f]===""||L[i][f]==null) delete L[i][f];
  while(L.length&&!Object.keys(L[L.length-1]).length) L.pop();
  if(!L.length) delete x.ex[name];
  if(!Object.keys(x.ex).length) delete SETS[date];
  sSave(); syncMark(date);
}
/* ── 자동 저장 (2026-11-13) ── 바뀐 날짜의 세트 전부를 구글 시트 웹 앱으로 보낸다(그 날짜를 통째로 갈아끼움).
   주소는 이 폰에만(wk-sync-url) — 공개 저장소에 넣지 않는다. 못 보낸 날짜는 wk-sync-dirty에 남아 다음에 다시.
   no-cors라 응답 내용은 못 읽는다 — 네트워크가 통하면 보낸 것으로 친다. PC는 세션 시작 때 세트기록.py pull */
const SYU="wk-sync-url", SYD="wk-sync-dirty", SYT="wk-sync-at";
let SYNC_URL="", SDIRTY=[], SYNC_AT="", syncTimer=null, syncBusy=false;
try{ SYNC_URL=localStorage.getItem(SYU)||""; SDIRTY=JSON.parse(localStorage.getItem(SYD)||"[]")||[]; SYNC_AT=localStorage.getItem(SYT)||""; }catch(e){}
function syncSaveState(){ try{ localStorage.setItem(SYD,JSON.stringify(SDIRTY)); localStorage.setItem(SYT,SYNC_AT); }catch(e){} }
function syncMark(date){ if(SDIRTY.indexOf(date)<0) SDIRTY.push(date); syncSaveState();
  clearTimeout(syncTimer); syncTimer=setTimeout(sSync,3000); syncPaint(); }
/* 등 그립 회전 종목은 이름 뒤에 그 사이클의 그립을 붙여 보낸다(2026-11-13) — '랫풀다운 [오버 넓게]'.
   PC 표가 같은 그립끼리만 비교한다. 원암 축이 있으면 '· 원암/양손'까지 붙는다(gripOf().t) */
/* 사이클 교대로 같은 이름이 3세트 ↔ FST-7로 바뀌는 종목(2일차 플라이 두 자리 · 2026-11-12)은 FST-7인 사이클에
   '[FST-7]'을 붙인다 — 3세트 기록과 한 표에서 비교되지 않게. 늘 FST-7인 종목(교대 없음)은 붙이지 않는다.
   그 날짜 사이클의 교대를 잠깐 얹어(VIEW_R · reAlt — 시간표와 같은 방식) 보고 바로 되돌린다 */
function sTags(day,round,names){
  const out={}, d=D.back&&D.back[day]; if(!d||d.rest||!round) return out;
  const sv=VIEW_R; VIEW_R=round; reAlt();
  try{ names.forEach(n=>{ const e=allEx(d).find(x=>x.name===n); if(!e) return;
    const t=[]; const g=e.grip?gripOf(e):null; if(g&&g.t) t.push(g.t);
    if(e.cv&&e.s===7&&/한계/.test(e.r||"")) t.push("FST-7");
    if(t.length) out[n]=t.join(" · "); }); }
  finally{ VIEW_R=sv; reAlt(); }
  return out;
}
/* ── 지난번 · 지지난번 (2026-11-13 사용자 요청) ─────────────────────
   종목 카드 = 지난번 전부 + 지지난번 한 줄(흐리게) · 타이머 = 지난번 그 세트 번호만.
   '같은 것'은 PC 비교표와 같은 기준: 같은 일차 + 같은 이름 + 같은 꼬리표(그립 · FST-7, sTags).
   꼬리표는 (일차, 사이클)마다 한 번 계산해 둔다 — sTags가 교대를 잠깐 얹었다 되돌리므로
   그리는 도중이 아니라 그리기 전에(sWarm) 채운다. 폰 기록(wk-sets)만 본다 */
const TAGC={};
function tagsFor(day,round){
  const k=day+"|"+round; if(TAGC[k]) return TAGC[k];
  const d=D.back&&D.back[day]; if(!d||d.rest||!round) return (TAGC[k]={});
  return (TAGC[k]=sTags(day,round,allEx(d).map(e=>e.name)));
}
function sWarm(day,date){
  const c=cycDayOf(date); if(c) tagsFor(day,c.round);
  for(const d in SETS) if(SETS[d].day===day&&SETS[d].round) tagsFor(day,SETS[d].round);
}
const sE1 = L => { let m=null; L.forEach(x=>{ const w=parseFloat(x.w), r=parseInt(x.r,10);
  if(/^\s*\d+(\.\d+)?\s*$/.test(x.w||"")&&w>0&&r>0){ const v=w*(1+r/30); if(m==null||v>m) m=v; } }); return m; };
function sHist(day,name,date,k){
  const c=cycDayOf(date), cur=c?(tagsFor(day,c.round)[name]||""):"", out=[];
  Object.keys(SETS).filter(d=>d<date&&SETS[d].day===day).sort().reverse().some(d=>{
    const x=SETS[d], L=(x.ex&&x.ex[name])||[]; if(!L.some(y=>Object.keys(y).length)) return false;
    if((tagsFor(day,x.round)[name]||"")!==cur) return false;
    out.push({date:d, round:x.round, sets:L, e1:sE1(L)}); return out.length>=k; });
  return out;
}
const sCellTxt = x => !x||!Object.keys(x).length ? "·" : x.r ? `${x.w||"?"}×${x.r}${x.ok?"✓":""}` : x.ok ? `${x.w||"?"}×목표✓` : `${x.w||"?"}×?`;
function sHistHTML(day,name,date){
  const H=sHist(day,name,date,2), tag=(()=>{ const c=cycDayOf(date); return c?(tagsFor(day,c.round)[name]||""):""; })();
  if(!H.length) return `<div class="sl-prev dim">지난번 기록 없음${tag?` — ${esc(tag)} 차례는 이번이 처음`:""}</div>`;
  const line=(h,lab,cls)=>`<div class="sl-prev${cls}"><b>${lab}</b> <span>(${+h.date.slice(5,7)}/${+h.date.slice(8)} · ${h.round||"?"}사이클)</span> ${
    h.sets.map(sCellTxt).map(esc).join(" · ")}${h.e1?` · 1RM ${h.e1.toFixed(1)}`:""}</div>`;
  return line(H[0],"지난번","")+(H[1]?line(H[1],"지지난번"," dim"):"");
}
function sRowsOf(d){ const x=SETS[d], out=[]; if(!x) return out;
  const tg=sTags(x.day,x.round,Object.keys(x.ex));
  for(const n in x.ex){ const t=tg[n], nm=t?`${n} [${t}]`:n;
    x.ex[n].forEach((s,i)=>{ if(Object.keys(s).length)
      out.push([d,x.round||"",x.day||"",nm,i+1,s.w||"",s.r||"",s.ok?"O":""]); }); }
  return out; }
function sSync(){
  if(!SYNC_URL||!SDIRTY.length||syncBusy||!navigator.onLine) return;
  const dates=SDIRTY.slice(), rows=[].concat(...dates.map(sRowsOf));
  syncBusy=true;
  fetch(SYNC_URL,{method:"POST",mode:"no-cors",headers:{"Content-Type":"text/plain"},body:JSON.stringify({dates,rows})})
    .then(()=>{ SDIRTY=SDIRTY.filter(d=>dates.indexOf(d)<0);
      const n=new Date(); SYNC_AT=String(n.getHours()).padStart(2,"0")+":"+String(n.getMinutes()).padStart(2,"0"); syncSaveState(); })
    .catch(()=>{})
    .finally(()=>{ syncBusy=false; syncPaint(); if(SDIRTY.length&&SDIRTY.some(d=>dates.indexOf(d)<0)) setTimeout(sSync,1000); });
}
function syncText(){
  if(!SYNC_URL) return "자동 저장 꺼짐 — 아래 설정";
  if(SDIRTY.length) return `자동 저장 대기 ${SDIRTY.length}일${navigator.onLine?"":" (오프라인)"}`;
  return SYNC_AT?`자동 저장됨 ✓ ${SYNC_AT}`:"자동 저장 켜짐 ✓";
}
function syncPaint(){ const el=document.getElementById("sl-sync"); if(el) el.textContent=syncText(); }
window.addEventListener("online",sSync);
document.addEventListener("visibilitychange",()=>{ if(document.visibilityState==="visible") sSync(); });
/* 2026-11-13 등 이름에 그립 꼬리표가 붙었다 — 시트의 옛 이름 줄을 갈아끼우게 한 번 전부 다시 보낸다 */
try{ if(localStorage.getItem("wk-sync-v")!=="3"){ SDIRTY=Array.from(new Set(SDIRTY.concat(Object.keys(SETS)))); syncSaveState();
  localStorage.setItem("wk-sync-v","3"); } }catch(e){}
setTimeout(sSync,1500);
const sPlan = w => String(w||"").replace(/\s*kg$/,"");
function sToggle(date,day,name,i,pw){
  const s=sGet(date,name)[i]||{};
  if(s.ok) sPut(date,day,name,i,{ok:""});
  else sPut(date,day,name,i,{ok:1, w:s.w||sPlan(pw)});
}
/* 그 일차의 기록 날짜 — 달력에서 고른 날 · 아니면 오늘부터 거꾸로 가장 가까운 그 일차 · 없으면 앞으로 */
function logDateFor(day){
  if(LOGDATE&&LOGDATE.day===day) return LOGDATE.date;
  const t=ymd(new Date());
  for(let i=0;i<CLEN;i++){ const a=dayAdd(t,-i), c=cycDayOf(a); if(c&&c.day===day) return a; }
  for(let i=1;i<=CLEN;i++){ const a=dayAdd(t,i), c=cycDayOf(a); if(c&&c.day===day) return a; }
  return t;
}
function sRowsHTML(e,i0,i1,date,pw,a){
  const L=sGet(date,e.name); let h="";
  for(let i=i0;i<i1;i++){ const x=L[i]||{};
    h+=`<div class="sl-row"><span class="sl-n">${i+1}</span>
      <input class="sl-w" type="text" inputmode="decimal" data-${a}="w" data-si="${i}" value="${escAttr(x.w||"")}" placeholder="${escAttr(sPlan(pw))}"><i>kg</i>
      <input class="sl-r" type="text" inputmode="numeric" data-${a}="r" data-si="${i}" value="${escAttr(x.r||"")}" placeholder="${escAttr(e.r||"")}"><i>회</i>
      <button class="sl-ok${x.ok?" on":""}" data-${a}="ok" data-si="${i}">&#10003;</button></div>`; }
  return h;
}
const sBox = (e,date,day,pw,inner) => `<div class="sl-box" data-sn="${escAttr(e.name)}" data-sd="${date}" data-sday="${escAttr(day)}" data-spw="${escAttr(pw)}">${inner}</div>`;
const sDone = (date,name) => sGet(date,name).filter(x=>x.ok||x.r).length;
function setLogHTML(e,n){
  if(mode!=="back"||!n||!D.back[tab]) return "";
  const date=logDateFor(tab), pw=tmW(e);
  return sHistHTML(tab,e.name,date)+`<details class="slog" data-sn="${escAttr(e.name)}"${SOPEN.has(e.name)?" open":""}><summary>세트 기록 · <b class="num">${sDone(date,e.name)}/${n}</b></summary>${
    sBox(e,date,tab,pw,sRowsHTML(e,0,n,date,pw,"sl"))}</details>`;
}
/* PC로 — 탭 구분 한 줄 = 한 세트. Claude가 받아서 세트기록.csv에 합친다(같은 날짜 · 종목 · 세트는 덮어씀) */
function sExport(){
  const out=["#세트기록 v1","날짜\t사이클\t일차\t종목\t세트\t무게\t횟수\t달성"];
  Object.keys(SETS).sort().forEach(d=>{ sRowsOf(d).forEach(r=>out.push(r.join("\t"))); });
  return out.join("\n");
}
const isDone = v => v===""||/^[oO○]$/.test(v);        /* 엑셀: OR(칸="",칸="O") */
const missed = (k,c) => !isDone(logVal(c,k));
/* w사이클 중량 = 시작 + 증량폭 × (그 앞에서 증량된 사이클 수). 미달한 사이클은 다음으로 안 올린다 */
function stepsTo(k,w){ let n=0; for(let c=1;c<w;c++) if(!missed(k,c)) n++; return n; }
/* 사다리를 한 칸씩 걸어 올라간다. 못 끼우는 값을 만나면 그 자리에서 1kg 올리고
   다음 칸은 올린 값에서 이어지므로, 시작중량 + 폭 × 칸수로는 재현되지 않는다 —
   엑셀 증량기록도 같은 방식이라 두 곳이 같은 숫자를 내려면 여기서도 걸어야 한다 */
function ladder(k,w,base){
  const L=LIFT[k], ev=L.every||1, f0=L.from||WBASE;   /* 보조 둘은 13주차에서 출발한다 */
  /* 2026-11-10 다시 출발(rs · 핵 스쿼트) — hold~at-1 칸은 그대로 · at 칸에 base · 그 뒤 cap까지 +inc · 닿으면 원래 증량폭.
     엑셀 증량기록 AL열 같은 식 */
  const S=L.rs;
  if(S&&w>=S.at){ let v=S.base;
    for(let c=S.at+1;c<=w;c++){ if(missed(k,c-1)) continue; v = v<S.cap ? Math.min(S.cap,v+S.inc) : R(v+L.inc); }
    return v; }
  if(S&&w>=S.hold) return ladder(k,S.hold-1,base);
  let v=base;
  for(let c=f0+1;c<=w;c++){
    if(missed(k,c-1)) continue;              /* 앞 주가 미달이면 그 중량을 한 번 더 */
    /* 사이클 걸러 하는 추적 종목(lb 2 · 2026-11-02 인클라인 바벨)은 앞 칸이 안 한 사이클이라 비어 있다 —
       그 앞 칸 미달도 본다(엑셀 증량기록 같은 식) */
    if(L.lb===2&&c-2>f0-1&&missed(k,c-2)) continue;
    if((c-f0)%ev===0) v=R(v+L.inc);          /* 주기가 2면 격주에만 오른다(프레스) */
  }
  return v;
}
/* 비율 종목(2026-11-09 인클라인 바벨 = 벤치 × 0.8) — 1kg 내림 · 못 끼우는 값은 1kg 더 내림(엑셀 AJ열 같은 식).
   폰에서 무게를 고치면 비율이 바뀐다(OV.lifts[k].ratio) */
const ratioOf = k => (OV.lifts[k]&&OV.lifts[k].ratio) || LIFT[k].ratio;
const fromRatio = (b,r) => { let v=Math.floor(b*r+1e-9); if(v%5===4) v--; return v; };
const workWeight = (k,w) => LIFT[k].of ? fromRatio(workWeight(LIFT[k].of,w), ratioOf(k)) : ladder(k,w,startOf(k));
const xlWeight   = (k,w) => LIFT[k].of ? fromRatio(xlWeight(LIFT[k].of,w), LIFT[k].ratio) : ladder(k,w,LIFT[k].start);
/* 역산 — 이번 사이클 무게를 만드는 시작중량. 사다리가 칸마다 보정되므로 나눗셈이 안 된다.
   못 끼우는 값을 입력하면 가장 가까운 끼울 수 있는 무게로 붙는다 */
function startFor(k,w,n){
  let best=Math.round(n), bd=Infinity;
  for(let s=1;s<=Math.round(n)+1;s++){ const d=Math.abs(ladder(k,w,s)-n); if(d<bd){ bd=d; best=s; } }
  return best;
}
const clampW = w => Math.min(WLAST, Math.max(WBASE, w||WBASE));
/* w사이클 직전까지 같은 종목이 몇 사이클 연속 미달인지. 3이면 엑셀 규칙 "10% 디로드 후 재상승" —
   자동으로 내리지는 않는다. 중량을 탭해서 직접 고치면 웜업·원판이 따라온다 */
/* '원암'(2026-11-04) = 해머 로우 회전 2 원암 날 — 무게는 유지하지만 연속 미달 셈에서는 건너뛴다(엑셀 증량기록 A56) */
function missRun(k,w){ let n=0; for(let c=w-1;c>=1;c--){ if(logVal(c,k)==="원암") continue; if(!missed(k,c)) break; n++; } return n; }
function deloadNote(k,w){
  const cut = () => fmt(Math.max(20,R(workWeight(k,w)*0.9)));
  if(missRun(k,w)>=3)   return `앞 3주 연속 미달 — 규칙은 이 주 10% 디로드(${cut()}kg)`;
  if(missRun(k,w+1)>=3) return `3주 연속 미달 — 규칙은 다음 주 10% 디로드(${cut()}kg)`;
  return "";
}
/* 사이클 c의 1일차 날짜. 달력 기준점에서 주기 길이만큼 되감거나 앞으로 센다 */
function cycleDate(c){
  if(!ANCHOR) return "";
  const d=ymdParse(ANCHOR.date); d.setDate(d.getDate()+(c-ANCHOR.cycle)*CYCLE); return ymd(d);
}

/* ── 볼륨 램프 — 종목 복귀 방식 ───────────────────────────────────
   계수로 전 종목을 한꺼번에 깎던 방식을 버렸다. 종목마다 '복귀 주차'(엑셀 K열)가 있고,
   그 주차 전에는 종목 자체가 빠진다. 나오는 종목은 설계 세트를 그대로 한다.
   엑셀 J열과 같은 식: IF(볼륨 주차 >= 복귀 주차, 설계 세트, 0).
   볼륨 주차는 중량 주차보다 offset만큼 앞서간다 (2026-08-10: 중량 2주차 · 볼륨 3주차).
   복귀 모드에는 주차 개념이 없으므로 적용하지 않는다. */
const RAMP = D.ramp;
const rampOn = () => !!(RAMP && mode==="home");
const vWeek = w => w + (RAMP ? RAMP.offset : 0);
const setsOf = (e,w) => vWeek(w) >= (e.rw||1) ? e.s : 0;
const rampRow = w => RAMP.rows.filter(r=>r.week<=vWeek(w)).pop() || RAMP.rows[0];

function ramp(work){
  const pct=[.4,.6,.75,.9], reps=[5,5,3,2], sets=[2,1,1,1];
  const raw = pct.map(p=>Math.max(20,R(work*p)));
  return raw.map((w,i)=>({w:(i>0 && w<=raw[i-1])?null:w, reps:reps[i], sets:sets[i]}));
}
const PLATES=[{kg:25,c:"#C93A42",h:42,w:11},{kg:20,c:"#2B62B0",h:42,w:11},
  {kg:15,c:"#DDA51B",h:40,w:9},{kg:10,c:"#2E9161",h:37,w:8},
  {kg:5,c:"#E8E6E1",h:30,w:8},
  {kg:2.5,c:"#3D454F",h:24,w:7,n:1},{kg:1,c:"#B9C0C8",h:19,w:6,n:1},
  {kg:0.5,c:"#8B95A1",h:16,w:6,n:1}];
function load(total){
  let side=(total-20)/2, out=[];
  if(side<0.001) return out;
  /* 작은 원판(2.5 · 1 · 0.5)은 한쪽에 하나씩뿐이라 n으로 막는다 */
  for(const p of PLATES){ let left=p.n||99;
    while(side>=p.kg-0.0005 && left-->0){ out.push(p); side=Math.round((side-p.kg)*1000)/1000; } }
  return out;
}
function stackHTML(total){
  const ps=load(total);
  const body = ps.length
    ? ps.map(p=>`<div class="pl" style="background:${p.c};width:${p.w}px;height:${p.h}px"></div>`).join("")
      + `<em>한쪽 ${fmt((total-20)/2)}kg</em>`
    : `<em>빈 봉 20kg</em>`;
  return `<div class="stack"><div class="sleeve"></div>${body}</div>`;
}
/* name — 시트의 '메인:' 줄에 적힌 이름. 같은 리프트를 다른 기구로 치는 날이 있어
   (복귀 수요일의 시티드 바벨 밀리터리 프레스 = 밀리터리 자리) 카드에는 그 날의 이름을 쓴다 */
function liftCard(key,week,sets,name){
  const L=LIFT[key]; if(!L) return "";
  const work=workWeight(key,week), rows=ramp(work);
  /* 본세트 수는 종목표 1번 행이 정한다 — 메인 리프트는 어느 주차에도 빠지지 않는다 */
  const ws=(sets==null?L.sets:sets);
  if(ws===0) return `<section class="lift" style="--accent:${L.accent}"><div class="lift-top"><div class="lift-name">
    <div class="eyebrow">MAIN LIFT · 디로드</div><h3>${esc(name||L.ko)}</h3>
    <p class="num">이번 주는 건너뛴다 — 다음 주 ${fmt(workWeight(key,Math.min(WLAST,week+1)))}kg</p></div></div></section>`;
  const rp=rows.map(r=>`
    <div class="r-w ${r.w?'':'skip'}">${r.w?fmt(r.w)+'<span style="font-size:11px;color:var(--mute)"> kg</span>':'생략'}</div>
    <div class="r-x num">${r.reps}회</div><div class="r-x num">${r.sets}세트</div>`).join("");
  return `<section class="lift" style="--accent:${L.accent}">
    <div class="lift-top">
      <div class="lift-name"><div class="eyebrow">MAIN LIFT</div><h3>${esc(name||L.ko)}</h3>
        <p class="num">5년 전 ${fmt(L.prev)}kg · 본세트 ${L.reps}회 × ${ws}세트${ws!==L.sets?` <span style="color:var(--dim)">(설계 ${L.sets})</span>`:""}</p></div>
      <div class="bigw ${OV.lifts[key]?'edited':''}" data-lift="${key}"><b class="num">${fmt(work)}</b><i>kg</i>${
        OV.lifts[key]?`<span class="ovnote">수정됨 · 엑셀 기준 ${fmt(xlWeight(key,week))}kg</span>`:""}</div>
    </div>
    <div class="bar"><div class="bar-label">한쪽 원판 구성</div>${stackHTML(work)}</div>
    <div class="ramp">
      <div class="h">웜업 램프</div><div class="h" style="text-align:right">반복</div><div class="h" style="text-align:right">세트</div>
      ${rp}
      <div class="work r-w">${fmt(work)}<span style="font-size:11px;color:var(--mute)"> kg</span></div>
      <div class="work r-x num">${L.reps}회</div><div class="work r-x num">${ws}세트</div>
    </div>
    ${dlOn()?`<div class="lift-log"><div class="lift-log-h">디로드 주 — 결과를 적지 않는다 <span>· 지난주와 같은 ${week}주차 칸이라 적으면 지난주 기록을 덮는다</span></div></div>`
      :liftLogHTML(key,week)}</section>`;
}
/* 본세트 결과 한 줄 — 세션 끝나고 그 자리에서 적는다. 증량 탭의 카드와 같은 저장소다 */
function liftLogHTML(key,week){
  const dl=deloadNote(key,week), held=missRun(key,week)>0;
  return `<div class="lift-log"><div class="lift-log-h">이번 칸 결과 <span>· ${week}주차 칸${
      held?` · 앞 주 미달로 중량 유지`:""}</span>${dl?`<em>${esc(dl)}</em>`:""}</div>${logCtl(key,week)}</div>`;
}
/* 한 종목 · 한 사이클 기록 조각 — [달성] [미달 N회]. 같은 버튼을 다시 누르면 비운다(미기록) */
function logCtl(k,c){
  const v=logVal(c,k), done=v!==""&&isDone(v), miss=v!==""&&!isDone(v);
  const reps=/^\d+$/.test(v)?v:"";
  return `<div class="seg" data-logk="${k}" data-logc="${c}">
    <button class="${done?'on ok':''}" data-logv="O">달성</button>
    <button class="${miss?'on bad':''}" data-logv="X">미달</button>
    ${miss?`<input class="log-in num" type="text" inputmode="numeric" data-logrep="1" value="${escAttr(reps)}" placeholder="마지막 세트 회">`:""}
  </div>`;
}
const esc = t => String(t).replace(/&/g,"&amp;").replace(/</g,"&lt;");
const escAttr = t => esc(t).replace(/"/g,"&quot;");
/* 백오프 행의 중량 — 톱세트 − lo(kg) 또는 톱세트 × pct(1kg 반올림 · 엑셀 MROUND와 같다).
   엑셀 캐시값이 아니라 지금 주차에서 낸다 */
/* 추적 보조 종목(해머 로우 · 인클라인 머신)은 그 줄 자체가 리프트다 — 사다리 값 그대로 */
const exW = e => (e.trk&&LIFT[e.lift]) ? (workWeight(e.lift,week)===0 ? "빈 바" : fmt(workWeight(e.lift,week))+"kg")
  : ((e.lo||e.pct)&&e.lift&&LIFT[e.lift])
  ? fmt(Math.max(20, e.pct ? Math.round(workWeight(e.lift,week)*e.pct)
                           : workWeight(e.lift,week)-e.lo))+"kg" : e.w;
/* 비율 램프(웜업 계획의 pctOf — 8일차 2번 바벨 · 4일차 핵 스쿼트) — 타이머와 카드 표가 같은 계산을 쓴다 (2026-11-09) */
function pctRamp(x,e){
  const base=(e.lift&&LIFT[e.lift])?workWeight(e.lift,week):parseFloat(tmW(e)); if(!(base>0)) return [];
  return x.rows.map(r=>({bar:r[0]==="bar", w:r[0]==="bar"?20:Math.max(20,R(base*r[0])), reps:r[1], rt:r[2]}));
}
function pctRampHTML(e,prefix,n){
  const day=String(prefix).split("|")[1], P=mode==="back"&&D.wuPlan?D.wuPlan[day]:null; if(!P) return "";
  const x=P.find(y=>y.pctOf&&(y.pctOf==="#2"?e.n===2:y.pctOf===e.name)); if(!x) return "";
  const rows=pctRamp(x,e); if(!rows.length) return "";
  const kg=v=>`${fmt(v)}<span style="font-size:11px;color:var(--mute)"> kg</span>`;
  const top=parseFloat((e.lift&&LIFT[e.lift])?workWeight(e.lift,week):tmW(e));
  return `<div class="ramp">
      <div class="h">웜업 램프</div><div class="h" style="text-align:right">반복</div><div class="h" style="text-align:right">휴식</div>
      ${rows.map(r=>`<div class="r-w">${r.bar?"빈 봉 "+kg(20):kg(r.w)}</div><div class="r-x num">${r.reps}회</div><div class="r-x num">${r.rt}초</div>`).join("")}
      <div class="work r-w">${kg(top)}</div><div class="work r-x num">${esc(e.r)}회</div><div class="work r-x num">${n}세트</div>
    </div>`;
}
function exHTML(e,prefix,sets,timed){
  const key=prefix+"|"+e.name, ov=OV.ex[key];
  const w=ov&&!e.trk?ov.w:exW(e), canEdit=e.w!=="위 참조"&&!e.lo&&!e.pct&&!e.trk;
  /* 추적 종목의 중량을 고치면 시작중량을 역산한다(메인 카드와 같은 덮어쓰기) */
  const wAttr = e.trk ? ` data-lift="${e.lift}"` : canEdit ? ` data-exkey="${escAttr(key)}"` : "";
  const wEd = e.trk ? !!OV.lifts[e.lift] : !!ov;
  const n=(sets==null?e.s:sets);
  /* 이번 주에 안 하는 종목 — 목록에서 빼지 않고 흐리게 남긴다.
     사라진 게 아니라 몇 주차에 돌아오는지가 보여야 한다. 메모·대체 종목은 그때 읽으면 된다.
     설계 세트(e.s)까지 0이면 램프 대기가 아니라 프로그램에서 빼둔 종목이다 —
     엑셀도 조건부서식($E=0)으로 같은 줄을 연회색 처리한다. 되살릴 수 있게 행은 남긴다. */
  if(n===0) return `<article class="ex hold"><div class="ex-h">
      <div class="ex-n num">${e.n}</div>
      <div class="ex-t"><h4>${esc(e.name)}</h4><div class="mg">${esc(e.mg)}</div></div>
      <div class="ex-p">${e.s===0 ? `<b class="hold-w">보류</b><span class="num">${esc(e.r)}</span>`
        : `<b class="hold-w">${dlOn()?"디로드 — 쉼":e.rw+"주차 복귀"}</b><span class="num">설계 ${e.s} × ${esc(e.r)}</span>`}
      </div></div></article>`;
  const gp = mode==="back" ? gripOf(e) : null;
  return `<article class="ex"><div class="ex-h">
      <div class="ex-n num">${e.n}</div>
      <div class="ex-t"><h4>${esc(gp&&e.gname?e.gname:e.name)}</h4>${gp?`<div class="grip">그립 · ${esc(gp.t)} <i>(회전 ${gp.n})</i></div>${gp.cue?`<div class="grip cue">▸ ${esc(gp.cue)}</div>`:""}`:""}${e.cv&&e.cvf?(e.cvs==="on"?"":`<div class="grip">${e.cvf-1}사이클까지 <i>(${e.cvf}사이클부터 ${esc(e.cv.on.name)})</i></div>`):e.cv&&ALT_ON?`<div class="grip">사이클 교대 · ${ALT_ON==="odd"?"홀수":"짝수"} 사이클 <i>(다음 사이클은 ${esc(e.cv[ALT_ON==="odd"?"even":"odd"].name)})</i></div>`:""}${e.kn?`<div class="kn">⚠ ${esc(e.kn)}</div>`:""}${e.wu?`<div class="grip" style="color:var(--sq)">예열 · ${esc(e.wut||"50% × 6~8")}</div>`:""}<div class="mg">${esc(e.mg)}${e.ind?` · 간접 ${esc(e.ind)}`:""}${e.ss?` · <b class="ss">슈퍼세트 ${esc(e.ss)}</b>`:""}${e.drop?` · <b class="drop">마지막 세트 드롭</b>`:""}${e.wu?` · <b class="wu">예열 1세트</b>`:""}</div></div>
      <div class="ex-p"><b class="num">${n} × ${esc(e.r)}</b><span class="num ex-w ${wEd?'edited':''}"${wAttr}>${esc(w)}</span></div></div>
    ${e.wk?`<div class="ex-go"><button class="tm-go" data-tm="${escAttr(e.name)}">&#9654; 타이머<b>${armWk(e,e.wk)}초${armWk(e,e.wk)!==e.wk?` (${gp&&gp.arm}${gp&&gp.arm==="원암"?" · 좌우":""})`:""} · 휴식 ${restLabel(e)}${e.drop?` · 드롭 +${e.drop}초`:""}</b></button><span class="tm-none wk-ed${OVT[key]?" edited":""}" data-wkkey="${escAttr(key)}">수행 ${e.wk}초 &#9998;</span></div>`
      :timed?`<div class="ex-go"><span class="tm-none">타이머 없음 — 엑셀 G·H열(수행 초 · 휴식 초)이 비어 있다</span></div>`:""}
    ${pctRampHTML(e,prefix,n)}
    ${micHTML(e,key,w)}
    ${e.trk&&LIFT[e.lift]&&!dlOn()?liftLogHTML(e.lift,week):""}
    ${setLogHTML(e,n)}
    ${e.memo?`<div class="ex-memo">${esc(e.memo)}</div>`:""}
    ${e.alt?`<details><summary>대체 종목</summary><p>${esc(e.alt)}</p></details>`:""}</article>`;
}
/* ── 미세 단계 (케이블 · 레터럴) ──────────────────────────────────
   반복 범위는 그대로고, 전 세트가 상한이면 한 단계 올린다. 단계는 무게 칸 글자 그 자체다 —
   '55 + 핀 + 1'. 바꾼 값은 다른 수정과 같이 이 폰 덮어쓰기에 남고, 엑셀에 옮겨 적으면 저절로 빠진다.
   핀 무게는 재지 않는다 — 늘 같은 핀이라 지난주와 비교하는 데 숫자가 필요 없다.
   마지막 단계 다음(스택 한 칸 · 다음 덤벨)은 기계마다 칸 크기가 달라 숫자를 직접 받는다 */
const MIC_RE = {
  assist: /^(\d+(?:\.\d+)?)\s*(?:kg)?\s*(?:\+\s*핀\s*(?:\+\s*(\d+(?:\.\d+)?)\s*)?)?$/,
  pin:   /^(\d+(?:\.\d+)?)\s*(?:kg)?\s*(?:\+\s*핀\s*(?:\+\s*(\d+(?:\.\d+)?)\s*)?)?$/,
  wrist: /^(\d+(?:\.\d+)?)\s*(?:kg)?\s*(?:\+\s*손목\s*(\d+(?:\.\d+)?)\s*)?$/,
};
function micStep(kind,w,dir){
  if(kind==="bar"||kind==="kg"||kind==="half"){
    const m=/^(\d+(?:\.\d+)?)\s*(?:kg)?$/.exec(String(w).trim()); if(!m) return null;
    const st=kind==="half"?.5:1, f=kind==="bar"?(v=>v%5===4?v+dir:v):(v=>v);
    const v=f(+m[1]+dir*st);
    return v>0?{w:fmt(v)}:null;
  }
  const m=String(w).match(MIC_RE[kind]); if(!m) return null;
  const b=m[1], pinOn=/핀/.test(w), x=m[2]?Number(m[2]):0;
  if(kind==="pin"){
    const L=[b, b+" + 핀", b+" + 핀 + 1", b+" + 핀 + 2.5"];
    const i=!pinOn?0:x===0?1:x<2.5?2:3;
    if(dir<0) return i>0?{w:L[i-1]}:null;
    return i<3?{w:L[i+1]}:{wrap:true, w:b+" + 핀", say:"핀을 스택 한 칸 아래로 · 원판 빼기 — 새 칸 숫자로 고치세요"};
  }
  if(kind==="assist"){
    const L=[b+" + 핀 + 2.5", b+" + 핀 + 1", b+" + 핀", b];
    const i=!pinOn?3:x===0?2:x<2.5?1:0;
    if(dir<0) return i>0?{w:L[i-1]}:null;
    return i<3?{w:L[i+1]}:{wrap:true, w:b+" + 핀 + 2.5", say:"어시스트 한 칸 가볍게 · 핀 + 2.5 — 새 칸 숫자로 고치세요"};
  }
  /* 모래주머니는 한쪽 1kg 고정(조절형 아님 · 2026-10-24) — 단계가 하나뿐이다 */
  if(dir<0) return x>0?{w:b}:null;
  return x===0?{w:b+" + 손목 1"}:{wrap:true, w:b, say:"다음 덤벨로 · 손목 빼기 — 덤벨 무게로 고치세요"};
}
/* 조건형 상한 (2026-10-24 사용자 결정) — 숫자로 막지 않는다. ▲ 누르기 전에 이 중 하나라도 있으면 무게 대신 반복으로 */
const MIC_CAP=`<span class="mic-cap">반동 · 가동범위 줄어듦 · 다른 근육이 먼저 지침 · 관절 통증 — 하나라도 있으면 ▲ 대신 반복</span>`;
function micHTML(e,key,w){
  if(!e.mic||e.trk) return "";
  const up=micStep(e.mic,w,1), dn=micStep(e.mic,w,-1);
  if(!up) return `<div class="ex-go mic"><span class="tm-none">미세 단계 — 무게를 '${e.mic==="wrist"?"8 + 손목 1":e.mic==="bar"||e.mic==="kg"||e.mic==="half"?"20":"55 + 핀 + 1"}' 꼴로 적으면 버튼이 뜬다</span></div>`;
  const top=(String(e.r).match(/~\s*(\d+)/)||[])[1];
  return `<div class="ex-go mic"><button class="tm-go" data-mic="${escAttr(key)}">&#9650; ${top?`전 세트 ${top}회 →`:"상한 채움 →"}<b>${esc(up.wrap?"다음 칸":up.w)}</b></button>${
    dn?`<button class="tm-go mic-dn" data-micdn="${escAttr(key)}" aria-label="한 단계 내리기">&#9660;</button>`:""}${MIC_CAP}</div>`;
}
/* 2026-10-27 사용자 요청 — 월 원암 케이블 레터럴 무게가 수요일을 따라간다(엑셀 월 D22 = 수 D36).
   한 방향이다: 수를 고치면 월도 같은 값으로 · 월만 고치면 그 날만 바뀐다(다음에 수를 고치면 다시 맞춰진다) */
/* 2026-11-01 v22 9일 사이클 — 원암 케이블 레터럴은 5일차 하나뿐이라 묶을 것이 없다 */
const EX_LINK = {};
function setExW(key,v){
  const e=exByKey(key); if(!e) return;
  [key].concat(EX_LINK[key]||[]).forEach(k=>{ const x=exByKey(k); if(!x) return;
    if(!v||v===x.w) delete OV.ex[k]; else OV.ex[k]={w:v, base:x.w}; });
  saveOV();
}
function micGo(key,dir,tmEl){
  const e=exByKey(key); if(!e) return;
  const nx=micStep(e.mic, OV.ex[key]?OV.ex[key].w:e.w, dir); if(!nx) return;
  if(!nx.wrap){ setExW(key,nx.w); render(); if(TM) tmPaint(); return; }
  /* 타이머에서는 위 무게 칸에서 바로 고친다 — 뒤 화면의 카드는 가려져 있다 */
  if(tmEl) return editEx(tmEl,key,()=>{ tmEl.innerHTML=""; tmPaint(); },nx.w,nx.say);
  const el=[...document.querySelectorAll("[data-exkey]")].find(x=>x.dataset.exkey===key);
  if(el) editEx(el,key,null,nx.w,nx.say);
}
const notesHTML = ns => !ns||!ns.length ? "" :
  `<div class="notes">`+ns.map(([h,ls])=>`<h5>${esc(h)}</h5><ul>${ls.map(l=>`<li>${esc(l)}</li>`).join("")}</ul>`).join("")+`</div>`;

const allEx = d => (d.ex||[]).concat(d.groups?d.groups.reduce((a,g)=>a.concat(g[1]),[]):[]);

function rampBar(all,S,total,done,week){
  const wait={};
  all.forEach(e=>{ if(e.s>0 && S(e)===0) wait[e.rw]=(wait[e.rw]||0)+1; });
  const when=Object.keys(wait).map(Number).sort((a,b)=>a-b)
    .map(k=>`${k}주차에 ${wait[k]}종목`).join(" · ");
  return `<div class="rampbar"><b>볼륨 ${vWeek(week)}주차 · ${done}세트 <span>/ 설계 ${total}</span></b>
    <span>${when?`대기 중 — ${when} 복귀.`:""} 나오는 종목은 설계 세트 그대로 한다</span></div>`;
}

function dayHTML(d,week){
  if(d.rest) return `<h2 class="daytitle">${esc(d.t)}</h2><p class="daysub">${esc(d.sub||"휴식일")}</p>`;
  const prefix=mode+"|"+tab;
  const on=rampOn(), dl=dlOn();
  const S = e => effSets(e,week);
  if(mode==="back"&&D.back[tab]===d) sWarm(tab,logDateFor(tab));
  const all=allEx(d), done=all.reduce((a,e)=>a+S(e),0);
  /* 타이머가 실제로 걸 수 있는 종목 — 수행 초가 있어야 한다 */
  const timed = allEx(d).filter(e=>e.s!==0&&S(e)>0&&e.wk>0);
  const one = e => exHTML(e,prefix,S(e),timed.length>0);
  /* 설계 세트 0 = 프로그램에서 빼둔 종목(옮겼거나 제외했거나).
     목록 흐름에서 빼고 하단 접기에 모은다 — 폰에서는 오늘 할 것만 보이고,
     엑셀 행은 그대로라 세트수만 되돌리면 다시 흐름으로 돌아온다.
     램프 대기(설계는 있는데 이번 주 0)는 몇 주차 복귀인지가 정보라 그대로 남긴다. */
  const held = all.filter(e=>e.s===0), live = list => list.filter(e=>e.s!==0);
  const body = d.groups
    ? d.groups.map(([g,list])=>live(list).length
        ? `<div class="grouphead">${esc(g)}</div>`+live(list).map(one).join("") : "").join("")
    : `<div class="grouphead">운동 구성</div>`+live(d.ex).map(one).join("");
  const holdBox = held.length ? `<details class="holdbox"><summary>보류 ${held.length}종목
    <i>· 세트 합계에서 빠짐</i></summary>${held.map(one).join("")}</details>` : "";
  const bar = dl ? `<div class="rampbar"><b>디로드 사이클 · ${done}세트 <span>/ 설계 ${d.total}</span></b>
      <span>세트만 절반 — 중량은 지난 사이클(${week}주차 칸) 그대로. 데드는 건너뛰고, 이 사이클은 증량 기록을 적지 않는다</span></div>`
    : (on && done!==d.total) ? rampBar(all,S,d.total,done,week) : "";
  /* 메인 리프트 카드의 본세트 수는 종목표에서 그 리프트가 붙은 행과 같은 값을 쓴다.
     복귀 토요일처럼 메인이 둘이면 카드도 둘이다 — 웜업 램프가 종목마다 따로 필요하다 */
  const cards = (d.mains||[]).map(m=>{
    const row = all.find(e=>e.lift===m[0]);
    return liftCard(m[0],week,row?S(row):null,m[1]);
  }).join("");
  /* 버튼에 적는 세트 수는 '오늘 할 세트'가 아니라 '타이머가 돌릴 세트'다.
     새 종목에 G·H열을 안 채우면 여기서 수가 어긋나 바로 눈에 띈다. */
  const tSets = timed.reduce((a,e)=>a+S(e),0);
  const go = timed.length
    ? `<button class="day-go" data-tm="">&#9654; 세션 타이머 시작 · ${tSets}세트${
        tSets!==done?` <i>(시간 없는 ${done-tSets}세트 제외)</i>`:""}</button>` : "";
  const wst=wuSteps(d), wsec=wst.reduce((a,x)=>a+x.t,0);
  const wgo = wst.length ? `<button class="day-go wu-go" data-wtm="1">&#9654; 웜업 타이머 · ${wuItems(d).length}항목 · 약 ${Math.round(wsec/60)}분</button>` : "";
  const ld = mode==="back"&&D.back[tab] ? logDateFor(tab) : null;
  const sl = ld ? `<p class="sl-date">세트 기록 날짜 · <b>${+ld.slice(5,7)}/${+ld.slice(8)}(${DOWK(ld)})</b> — 다른 날이면 달력에서 그 칸을 누른다 ·
    <button class="tm-go" id="sl-copy">PC로 보낼 세트 기록 복사</button> · <b id="sl-sync">${syncText()}</b></p>
    <details class="cal-fix"><summary>자동 저장 설정</summary>
      <p class="footnote">구글 시트 웹 앱 주소(…/exec)를 한 번만 붙여 넣는다. 이 폰에만 저장되고, 붙여 넣는 순간 지금까지 기록 전부를 보낸다.</p>
      <div class="cal-days"><input id="sl-url" class="log-in wide" type="url" placeholder="https://script.google.com/macros/s/…/exec" value="${escAttr(SYNC_URL)}">
        <button id="sl-url-save">저장</button>${SYNC_URL?`<button id="sl-url-del">끄기</button>`:""}</div></details>` : "";
  return `<h2 class="daytitle">${esc(d.t)}</h2><p class="daysub">${esc(d.sub)}</p>${bar}${wgo}${go}${sl}
    ${cards}${notesHTML(d.pre)}${body}
    <div class="total"><span>세트 합계</span><b class="num">${done!==d.total?`${done} <i style="font-style:normal;font-weight:600;color:var(--mute);font-size:13px">/ 설계 ${d.total}</i>`:d.total}</b></div>${holdBox}${notesHTML(d.notes)}`;
}
/* ── 달력 ─────────────────────────────────────────────────────────
   8일 주기는 요일과 어긋나므로 '오늘 몇 일차인지'가 날짜로만 풀린다.
   기준점은 엑셀 증량기록의 '날짜 (1일차)'. 하루 빠뜨려 밀렸으면
   앱에서 '오늘 = N일차'로 다시 맞춘다 — 그 값은 이 폰에만 저장된다. */
const DOW="일월화수목금토";
const ymd = d => d.getFullYear()+"-"+String(d.getMonth()+1).padStart(2,"0")
                 +"-"+String(d.getDate()).padStart(2,"0");
const ymdParse = t => { const p=t.split("-").map(Number); return new Date(p[0],p[1]-1,p[2]); };
const dayGap = (a,b) => Math.round((ymdParse(a)-ymdParse(b))/86400000);
const AKEY="wk-anchor";
let ANCHOR=D.anchor||null;
try{ const o=JSON.parse(localStorage.getItem(AKEY)||"null"); if(o&&o.date) ANCHOR=o; }catch(e){}
const saveAnchor = () => { try{ localStorage.setItem(AKEY,JSON.stringify(ANCHOR)); }catch(e){} };

/* ── 복귀 달력 ─────────────────────────────────────────────────────
   복귀는 주기가 아니라 주차가 단위다 — 토요일에 주차가 넘어가고(2026-10-04: 훈련 주가
   토요일에 출발하므로 월요일 넘김을 버렸다. 넘기는 요일은 엑셀 '프로그램 시작일'의 요일이다),
   중량은 그 주차가 정한다(2026-09-13 전환 뒤로 매주 한 칸).
   사용자가 손으로 셀 일이 없게 하는 것이 이 블록의 전부다. */
const BA = D.backAnchor || null;
const BSKEY="wk-bshift", BCKEY="wk-bsess";
let BSHIFT=0;  try{ BSHIFT=Number(localStorage.getItem(BSKEY)||0)||0; }catch(e){}
let BSESS={};  try{ BSESS=JSON.parse(localStorage.getItem(BCKEY)||"{}")||{}; }catch(e){}
const bSave = () => { try{ localStorage.setItem(BSKEY,String(BSHIFT));
  localStorage.setItem(BCKEY,JSON.stringify(BSESS)); }catch(e){} };
const DOWK = t => "일월화수목금토"[ymdParse(t).getDay()];
const dayAdd = (t,n) => { const d=ymdParse(t); d.setDate(d.getDate()+n); return ymd(d); };
/* 주차와 사이클은 기준이 다르다 — 주차는 프로그램 시작일에서 세는 숫자일 뿐이고,
   사이클은 중량이 매달린 숫자라 사이클 기준일에서 따로 센다. 하나로 묶으면
   주차를 맞추려다 중량까지 따라 움직인다. BSHIFT(미루기)는 사이클에만 건다. */
/* 디로드 주 — 엑셀 시작중량_설정 '디로드 시작일'(토요일 날짜 목록). 그 주는 주차로 세지 않는다
   (사용자 지정 2026-10-04 — '주차로 새지 말고 디로드 주라고만'). 주차 · 중량 칸 · 엑셀 C4가
   전부 그 주에 멈추므로 세 숫자가 디로드 뒤에도 같이 간다. 세트는 절반으로 보인다.
   주차 기준일보다 앞선 디로드는 기준일에 이미 반영돼 있으므로 중량 계산에서는 그 뒤의 것만 뺀다 */
/* 적힌 날짜 + '디로드 간격(주)'로 마지막 날짜에서 계속 반복(증량 6주 + 디로드 1주 = 7).
   앞당긴 디로드는 그 날짜를 적으면 거기서 다시 센다. 80주 앞까지만 만든다 — 중량 표는 그 전에 끝난다 */
const DLD = (()=>{ const L=((BA&&BA.deload)||[]).slice().sort(), ev=(BA&&BA.dlevery)|0;
  if(L.length&&ev>0){ let t=L[L.length-1]; for(let i=0;i<Math.ceil(80/ev);i++){ t=dayAdd(t,ev*7); L.push(t); } }
  return L; })();
/* 2026-11-01 사용자 결정 — 중량은 9일 사이클마다 한 칸. 같은 종목이 9일에 한 번 오므로 한 세션에 한 칸이고,
   날짜 주(7일)로 세던 때처럼 두 칸을 한 번에 건너뛰는 일이 없다. 엑셀 시작중량_설정 '1사이클 중량 칸'(base)에서
   지난 사이클 수만큼 오른다. 디로드 = '첫 디로드 사이클'부터 '디로드 간격(사이클)'마다 — 그 사이클은 칸이 멈추고 세트 절반.
   반환 이름(wk · cycle · from)은 옛 주차 코드와 맞춘 것이다: wk = 표시 칸, cycle = 중량 칸, from = 그 사이클 1일차 */
const CA = D.cycAnchor || null;
function backWeekOf(t){
  if(!CA||!CANCHOR) return null;
  const n=dayGap(t,CANCHOR); if(n<0) return null;
  const k=Math.floor(n/CLEN);
  const isDl = j => CA.dlFirst>0 && j>=CA.dlFirst-1 &&
    (CA.dlEvery>0 ? (j-(CA.dlFirst-1))%CA.dlEvery===0 : j===CA.dlFirst-1);
  let before=0; for(let j=0;j<k;j++) if(isDl(j)) before++;
  const deload=isDl(k), wk=CA.base+k-before-(deload?1:0);
  return {wk, cycle:clampW(wk), from:dayAdd(CANCHOR,k*CLEN), deload, round:k+1};
}
const wkLabel = w => w.deload ? "디로드 사이클" : w.wk+"주차 칸";
const bToday = () => backWeekOf(ymd(new Date()));
/* 지금 보는 화면이 '이번 주 디로드'인가 — 화살표로 다른 주를 넘겨 보면 평소 세트로 보인다 */
const dlOn = () => { const b=bToday();
  return !!(b&&b.deload&&week===Math.max(b.cycle,xlCycle())); };
const isDeadRow = e => e.lift==="dead" || (/데드리프트/.test(e.name)&&!/루마니안/.test(e.name));
/* 복귀_증량원칙 디로드 규칙: 메인 리프트는 내림(5→2 · 4→2 · 3→1), 데드는 건너뜀, 보조는 반올림(3→2 · 4→2 · 6→3) */
const dlSets = e => e.s<=0 ? e.s : isDeadRow(e) ? 0
  : e.lift&&!e.trk ? Math.max(1,Math.floor(e.s/2)) : Math.max(1,Math.round(e.s/2));
const effSets = (e,w) => rampOn() ? setsOf(e,w) : dlOn() ? dlSets(e) : e.s;
/* 지금 쓰는 사이클보다 큰 사이클이 처음 시작되는 날. 손으로 먼저 올려 둔 경우에도
   '다음'이 과거를 가리키지 않는다. 마지막 사이클이면 null */
function backNextAfter(eff){
  if(!CANCHOR) return null;
  for(let j=0;j<WEEKS+20;j++){
    const from=dayAdd(CANCHOR,j*CLEN), w=backWeekOf(from);
    if(w&&w.cycle>eff) return {cycle:w.cycle, from};
  }
  return null;
}
/* 복귀 화면의 사이클 — 달력이 낸 값과 엑셀 '현재 사이클'(C4) 중 큰 쪽을 쓴다.
   손으로 먼저 올렸으면 그 뜻을 따르고, 달력이 앞서면 올릴 때가 됐다는 뜻이다.
   중량 표가 C4에 매달려 있으므로(백오프 · 일차 시트 D열) 둘을 벌려 두면 엑셀과 폰이 갈린다 */
const xlCycle = () => clampW(D.config.week);
function syncBackWeek(){
  const c=backWeekOf(ymd(new Date()));
  if(c) week=Math.max(c.cycle, xlCycle());
}

/* 날짜 → 그 날의 요일 일차(주 6일 분할). 중량 칸은 9일 사이클 달력과 같은 값을 쓴다 */
function cycleOf(t){ return {day: DOWK(t), cycle: week}; }
const PLANMAP={}; D.plan.rows.forEach(([d,c,n])=>{ PLANMAP[d]={c:c,n:n}; });
const shortLabel = d => (PLANMAP[d]?PLANMAP[d].c:d).replace(/\s*\([^)]*\)/g,"").trim();
/* 일차별 색. 메인 리프트가 있는 날은 그 종목 색(빨강 스쿼트 · 파랑 데드 ·
   노랑 벤치 · 초록 밀리터리)을 쓰고, 없는 날은 남은 색을 순서대로 받는다.
   이웃한 일차끼리 색이 겹치지 않는 것이 목적이다 — 마지막 일차 다음이 1일차라
   둘이 같은 색이면 사이클이 넘어간 건지 알 수 없다. */
const EXTRA=["#2BA6BF","#8A6CE0","#CE6BA5","#B08442"];
const DAYACC={};
(function(){ let i=0;
  D.homeOrder.forEach(d=>{
    const h=D.home[d];
    if(!h||h.rest) return;
    DAYACC[d] = (h.main&&LIFT[h.main]) ? LIFT[h.main].accent : EXTRA[i++%EXTRA.length];
  });
})();
const accentOf = d => DAYACC[d]||"";

let calYM=null;
function calHTML(){
  const now=new Date(), tKey=ymd(now), td=DOWK(tKey), P=PLANMAP[td]||{};
  const ym=calYM||{y:now.getFullYear(),m:now.getMonth()};
  const first=new Date(ym.y,ym.m,1), last=new Date(ym.y,ym.m+1,0);
  const head=`<div class="cal-now"><b style="color:${accentOf(td)||"var(--ink)"}">오늘 ${td}요일 · ${esc(P.c||"휴식")}</b>${P.n?` · ${P.n}세트`:""}
    ${P.n?`<button class="day-go" data-hcal="${td}">${td}요일 운동 보기</button>`:""}</div>`;
  const cells=[];
  for(let i=0;i<first.getDay();i++) cells.push(`<div class="cal-cell pad"></div>`);
  for(let dn=1;dn<=last.getDate();dn++){
    const d=new Date(ym.y,ym.m,dn), key=ymd(d), wd=DOWK(key), Q=PLANMAP[wd]||{}, rest=!Q.n, acc=rest?"":accentOf(wd);
    cells.push(`<div class="cal-cell${key===tKey?" today":""}${rest?" rest":""}${d.getDay()===0?" sun":""}"
      ${rest?"":`data-hcal="${wd}"`} style="${acc?`--acc:${acc}`:""}">
      <b>${dn}</b><i style="color:${acc||"var(--mute)"}">${wd}</i><span>${esc(rest?"휴식":shortLabel(wd))}</span></div>`);
  }
  return `<h2 class="daytitle">달력</h2>
  <p class="daysub">주 6일 분할 — 월~토 · 일요일 휴식. 요일이 곧 그날 할 운동이다(칸을 누르면 그 요일로). 중량은 9일 사이클 쪽과 같은 칸을 쓴다.</p>
  ${head}
  <div class="cal-bar"><button id="cal-prev">‹</button>
    <h3>${ym.y}년 ${ym.m+1}월</h3>
    <button id="cal-today">오늘</button><button id="cal-next">›</button></div>
  <div class="cal-grid">${[...DOW].map((w,i)=>`<div class="cal-dow${i===0?" sun":""}">${w}</div>`).join("")}
    ${cells.join("")}</div>`;
}
/* ── 9일 사이클 (2026-11-01 v22) ─────────────────────────────────────
   오늘이 몇 일차인지는 '사이클 1일차 기준일'에서 센다. 중량도 같은 기준일에서 사이클마다 한 칸(위 backWeekOf).
   쉬어서 일차가 밀리면 일차만 다시 누른다 — 지금 몇 번째 사이클인지는 지켜서 중량이 되돌아가지 않는다 */
const CAKEY="wk-canchor";
let CANCHOR=(D.cycAnchor&&D.cycAnchor.date)||null;
try{ const o=localStorage.getItem(CAKEY); if(o) CANCHOR=o; }catch(e){}
const CLEN=D.backOrder.length;
function cycDayOf(t){
  if(!CANCHOR) return null;
  const n=dayGap(t,CANCHOR); if(n<0) return null;
  const k=Math.floor(n/CLEN);
  return {day:D.backOrder[n%CLEN], round:k+1, start:dayAdd(CANCHOR,k*CLEN)};
}
/* ── 등 그립 회전 (2026-11-01) ── 엑셀 '사이클_그립회전'. 종목 · 세트는 그대로, 그립만 사이클마다 돈다.
   회전 = (오늘의 사이클 − 시작 사이클) ÷ 주기의 나머지. 시작 전에는 null — 시트 이름의 그립 그대로다.
   일차를 다시 눌러도 사이클 번호는 지켜지므로(setCycToday) 회전 순서가 꼬이지 않는다 */
const GR = D.gripRot || null;
function gripOf(e){
  if(!GR||!e||!e.grip) return null;
  const c=cycDayOf(ymd(new Date())), rr=VIEW_R!=null?VIEW_R:(c?c.round:null); if(rr==null) return null;
  const k=rr-GR.start; if(k<0) return null;
  const g=e.grip.g[k%e.grip.p];
  /* 원암 ↔ 양손(H열) — 2사이클 교대. 오버 그립 차례는 늘 양손. H열이 없으면 그립 칸 글자의 (원암)/(양손) */
  let arm=null;
  /* 원암은 언더 차례에만(2026-11-01 S1) — 오버 · 뉴트럴(넓은 패럴렐 · 팔꿈치 45° = 중부)은 늘 양손 */
  if(e.grip.a!=null) arm = /^(오버|뉴트럴|패럴렐)/.test(g) ? "양손" : ((k+e.grip.a)%2 ? "원암" : "양손");
  else arm = /원암/.test(g) ? "원암" : /양손/.test(g) ? "양손" : null;
  const cue=(e.grip.c&&e.grip.c[k%e.grip.p])||"";
  return {n:(k%3)+1, g:g, arm:arm, ax:e.grip.a!=null, t:(e.grip.a!=null ? g+" · "+arm : g), cue:cue};
}
/* 원암 차례면 수행 시간 2배(좌우) · 최소 120초(2026-11-01 사용자 — 원암 레터럴 · 리버스 플라이와 같은 선).
   회전표 종목이 원암인 사이클에만 건다. 원래 원암 종목(시트 초가 좌우 기준)은 원암 차례면 120초 이상,
   양손 차례면 절반 · 최소 40초 */
function armWk(e,wk){
  const gp = mode==="back" ? gripOf(e) : null;
  if(!gp||!gp.arm||!wk) return wk;
  const b1=!!e.grip.b;
  if(gp.arm==="원암") return Math.max(120, b1 ? wk : wk*2);
  if(gp.arm==="양손"&&b1) return 40;              /* 원암 케이블 로우 120초 → 양손 시티드 40초(보통 양손 종목 선) */
  return wk;
}
const exName = e => gripOf(e)&&e.gname ? e.gname : e.name;
/* 중량 칸 w를 쓰는 사이클 번호(디로드 사이클은 건너뜀). 기준일 앞 칸이거나 표 밖이면 null */
function roundOfWeek(w){
  if(!CANCHOR||!CA) return null;
  for(let j=0;j<WEEKS+20;j++){
    const b=backWeekOf(dayAdd(CANCHOR,j*CLEN));
    if(b&&!b.deload&&b.cycle===w) return b.round;
    if(b&&b.cycle>w) break;
  }
  return null;
}
function setCycToday(d){
  const i=D.backOrder.indexOf(d); if(i<0) return;
  const t=ymd(new Date()), c=cycDayOf(t), r=c?c.round:1;
  CANCHOR=dayAdd(t,-i-(r-1)*CLEN);
  try{ localStorage.setItem(CAKEY,CANCHOR); }catch(e){}
}
const CPLAN={}; ((D.backPlan&&D.backPlan.rows)||[]).forEach(r=>{ CPLAN[r[0]]={c:r[1],n:r[2],m:r[3]}; });
const cShort = d => (CPLAN[d]?CPLAN[d].c:"").replace(/\s*\([^)]*\)/g,"").split("·")[0].trim();
const CACC = d => { const h=D.back[d]; return h&&h.main&&LIFT[h.main]?LIFT[h.main].accent:""; };
function cycFix(c){
  return `<details class="cal-fix"${CANCHOR?"":" open"}><summary>${c?`일차가 밀렸다면 — 오늘을 다른 일차로 (지금 ${c.day}일차)`:"오늘은 몇 일차인가"}</summary>
    <div class="cal-days">${D.backOrder.map(d=>
      `<button data-csetday="${d}"${c&&c.day===d?' class="on"':""}>${d}일차 ${esc(cShort(d)||"휴식")}</button>`).join("")}</div>
    ${D.cycAnchor&&CANCHOR!==D.cycAnchor.date?`<div class="cal-days"><button data-csetday="__xl">엑셀 기준으로 되돌리기</button></div>`:""}
    <p class="footnote">쉬어서 일차가 밀리면 여기서 오늘 일차를 다시 누른다 — 이 폰에만 남는다.
      지금이 몇 번째 사이클인지는 그대로 두므로 중량은 되돌아가지 않고, 밀린 만큼 다음 칸이 늦게 온다.
      영구히 맞추려면 엑셀 시작중량_설정 '사이클 1일차 기준일'.</p></details>`;
}
function backCalHTML(){
  const now=new Date(), tKey=ymd(now), cur=backWeekOf(tKey), cd=cycDayOf(tKey);
  const keys=Object.keys(LIFT).filter(k=>!LIFT[k].acc);
  const ym=calYM||{y:now.getFullYear(),m:now.getMonth()};
  const first=new Date(ym.y,ym.m,1), last=new Date(ym.y,ym.m+1,0);
  const eff = cur ? Math.max(cur.cycle, xlCycle()) : xlCycle();

  const today = cd ? (()=>{ const P=CPLAN[cd.day]||{}, rest=!P.n;
      return `<p><b style="color:${CACC(cd.day)||"var(--ink)"}">오늘 ${cd.day}일차 · ${esc(P.c||"휴식")}</b>${rest?"":` · ${P.n}세트 · ${P.m}분`}</p>
        ${rest?"":`<button class="day-go" data-bcal="${cd.day}">${cd.day}일차 운동 보기</button>`}`; })()
    : CANCHOR&&dayGap(tKey,CANCHOR)<0
    ? `<p><b>사이클 시작 전</b> — ${+CANCHOR.slice(5,7)}/${+CANCHOR.slice(8)}(${DOWK(CANCHOR)})이 1일차 ${esc(cShort("1"))}다.</p>
       <button class="day-go" data-bcal="1">1일차 운동 미리 보기</button>`
    : `<p><b>오늘이 몇 일차인지 아직 모른다</b> — 아래에서 한 번만 누르면 그 뒤로는 날짜로 센다.</p>`;

  const md = t => `${+t.slice(5,7)}/${+t.slice(8)}`;
  let card="";
  if(cur){
    const nx=backNextAfter(eff), to=dayAdd(cur.from,CLEN-1);
    const up = keys.map(k=>{ const d=workWeight(k,eff)-workWeight(k,Math.max(WBASE,eff-1));
      return (d>0?"+":"")+fmt(d); }).join(" · ");
    card=`<p class="num" style="margin-top:10px"><b>${cur.round}사이클 · ${cur.deload?"디로드":eff+"주차 칸 중량"}</b> · ${md(cur.from)}(${DOWK(cur.from)}) ~ ${md(to)}(${DOWK(to)})</p>
      ${cur.deload?`<p><b>디로드 사이클</b> — 세트 절반 · 중량은 지난 사이클 그대로 · 데드 건너뜀. 다음 사이클에 ${cur.wk+1}칸으로 오른다</p>`
        :`<p class="num" style="color:var(--mute)">지난 사이클 대비 ${up}kg</p>`}
      <p>${keys.map(k=>`${esc(LIFT[k].ko)} <b class="num">${fmt(workWeight(k,eff))}</b>kg`).join(" · ")}</p>
      ${nx?`<p style="color:var(--mute)">다음 사이클(${md(nx.from)} ${DOWK(nx.from)}) — ${keys.map(k=>`${esc(LIFT[k].ko)} ${fmt(workWeight(k,nx.cycle))}`).join(" · ")}kg</p>`:""}`;
  } else if(CA){
    card=`<p class="num" style="margin-top:10px;color:var(--mute)">1사이클은 ${CA.base}주차 칸 중량으로 시작한다 — ${keys.map(k=>`${esc(LIFT[k].ko)} ${fmt(workWeight(k,clampW(CA.base)))}`).join(" · ")}kg</p>`;
  }
  const head=`<div class="cal-now">${today}${card}${cycFix(cd)}</div>`;

  const cells=[];
  for(let i=0;i<first.getDay();i++) cells.push(`<div class="cal-cell pad"></div>`);
  for(let dn=1;dn<=last.getDate();dn++){
    const d=new Date(ym.y,ym.m,dn), key=ymd(d), w=backWeekOf(key), c=cycDayOf(key);
    const P=c?CPLAN[c.day]||{}:{}, rest=c&&!P.n, acc=c&&!rest?CACC(c.day):"";
    /* 사이클 첫날 칸에 그 사이클의 중량 칸을 적는다 */
    const badge = w && c && c.day===D.backOrder[0] ? (w.deload?"디로드":`${w.wk}칸`) : "";
    cells.push(`<div class="cal-cell${c?"":" blank"}${key===tKey?" today":""}${rest?" rest":""}${d.getDay()===0?" sun":""}"
      ${c&&!rest?`data-bcal="${c.day}" data-bdate="${key}"`:""} style="${acc?`--acc:${acc}`:""}">
      ${SETS[key]?`<span class="sl-dot"></span>`:""}<b>${dn}</b>${c?`<i style="color:${acc||(rest?"var(--mute)":"var(--ink)")}">${c.day}일차</i>
        <span>${esc(rest?"휴식":cShort(c.day))}</span>`:""}${badge?`<em>${badge}</em>`:""}</div>`);
  }

  return `<h2 class="daytitle">달력</h2>
  <p class="daysub"><b>9일 사이클</b> — 7세션 + 휴식 2일이라 요일과 어긋난다. 오늘 할 일차는 날짜가 정한다(칸을 누르면 그 일차로).
    <b>중량은 사이클마다 한 칸</b> 오른다 — 1일차 칸에 그 사이클의 칸 번호가 적힌다.</p>
  ${head}
  <div class="cal-bar"><button id="cal-prev">‹</button>
    <h3>${ym.y}년 ${ym.m+1}월</h3>
    <button id="cal-today">오늘</button><button id="cal-next">›</button></div>
  <div class="cal-grid">${[...DOW].map((w,i)=>`<div class="cal-dow${i===0?" sun":""}">${w}</div>`).join("")}
    ${cells.join("")}</div>
  <div class="cal-key">${D.backOrder.map(d=>{ const acc=CPLAN[d]&&CPLAN[d].n?CACC(d):"";
    return `<span><i style="background:${acc||"transparent"};${acc?"":"border:1px dashed var(--line)"}"></i>${d} ${esc(cShort(d)||"휴식")}${
      D.back[d]&&D.back[d].main&&LIFT[D.back[d].main]?` <b>${esc(LIFT[D.back[d].main].ko)}</b>`:""}</span>`; }).join("")}</div>`;
}
/* 복귀 주차 보정 — 디로드나 통째로 쉰 주를 뒤로 미룬다. 이 폰에만 남는다 */
function backFix(cur){
  return `<details class="cal-fix"><summary>${cur?`주차가 안 맞으면 — 지금 ${wkLabel(cur)}`:"기준일 보정"}</summary>
    <div class="cal-days"><button data-bshift="1">한 주 미루기</button>
      <button data-bshift="-1">한 주 당기기</button>
      ${BSHIFT?`<button data-bshift="0">엑셀 기준으로 (${BSHIFT>0?"+":""}${BSHIFT}주 보정 중)</button>`:""}</div>
    <p class="footnote"><b>계획한 디로드에는 누르지 않는다</b> — 엑셀 '시작중량_설정'의 <b>디로드 시작일</b>에
      적은 주는 앱이 알아서 멈춘다(주차로 세지 않고 '디로드 주'로 표시). 여기서 또 누르면 두 번 밀린다.
      누르는 것은 계획 밖에 <b>통째로 쉰 주</b>뿐이다 — 주차 표시는 그대로 두고 중량만 한 주 늦춘다.
      영구히 맞추려면 '시작중량_설정'의 <b>주차 기준일</b>을 한 주 뒤로 옮긴다.</p></details>`;
}
function anchorFix(cur){
  return `<details class="cal-fix"><summary>${cur?"주기가 밀렸다면 — 오늘을 다른 일차로":"오늘은 몇 일차인가"}</summary>
    <div class="cal-days">${D.homeOrder.map(d=>
      `<button data-setday="${d}"${cur&&cur.day===d?' class="on"':""}>${d}일차 ${esc(shortLabel(d))}</button>`).join("")}</div>
    ${D.anchor?`<div class="cal-days"><button data-setday="__xl">엑셀 기준으로 되돌리기</button></div>`:""}</details>`;
}
/* 오늘을 N일차로 맞춘다 — 기준 날짜를 거꾸로 계산해 저장한다 */
function setToday(d){
  const L=D.homeOrder.length, i=D.homeOrder.indexOf(d);
  if(i<0) return;
  const back=new Date(); back.setDate(back.getDate()-i);
  const cur=cycleOf(ymd(new Date()));
  ANCHOR={date:ymd(back), cycle:(cur?cur.cycle:(D.config.week||1))};
  saveAnchor();
}

/* 검산 — 재택과 복귀가 다른 표를 쓴다. 복귀판은 주 단위라 환산이 없고,
   대신 현장 일이 채워주는 정도를 '보정' 열로 함께 보여준다. */
/* 주기 길이. 주간계획 행 수에서 오므로 8일 → 9일로 바뀌어도 화면 문구가 따라간다. */
const CYCLE = (D.homeOrder||[]).length || 8;
function cycCheckHTML(){
  const rows=D.backCheck||[];
  if(!rows.length) return `<h2 class="daytitle">세트 검산</h2><p class="daysub">표가 없습니다.</p>`;
  return `<h2 class="daytitle">세트 검산</h2>
  <p class="daysub">사이클_계획 부위별 표 그대로 — 직접 세트 · 주당 환산 = 사이클 × 7/9 · MAV 최소(주당)와 비교한다.
    간접은 미리 정한 규칙 하나(넓은 그립 로우 → 후면 0.33)만 쓴다.</p>
  <table class="tbl"><thead><tr><th>부위</th><th>사이클</th><th>주당</th><th>MAV</th><th>판정</th></tr></thead><tbody>
  ${rows.map(([m,c,w,v,tag,vd])=>`<tr><td${/합계|검산|실질/.test(m)?' style="font-weight:800"':""}>${esc(m)}</td>
    <td class="num" style="font-weight:800">${c}</td><td class="num">${w}</td>
    <td class="num" style="color:var(--mute)">${v===""?"—":v}</td>
    <td>${vd?`<span class="judge j-${tag}">${esc(vd)}</span>`:""}</td></tr>`).join("")}
  </tbody></table>
  <div class="notes"><h5>구성</h5><ul>${rows.filter(r=>r[6]).map(r=>`<li><b style="color:var(--ink)">${esc(r[0])}</b> — ${esc(r[6])}</li>`).join("")}</ul></div>`;
}
function splitCheckHTML(){
  const rows=D.check||[];
  if(!rows.length) return `<h2 class="daytitle">세트 검산</h2><p class="daysub">표가 없습니다.</p>`;
  return `<h2 class="daytitle">세트 검산</h2>
  <p class="daysub">분할_주간계획 부위별 표 그대로 — 주 단위 직접 세트다.</p>
  <table class="tbl"><thead><tr><th>부위</th><th>주간 세트</th><th style="text-align:left">구성</th></tr></thead><tbody>
  ${rows.map(([m,n,,,tag,vd,desc])=>`<tr><td${/합계|검산|실질/.test(m)?' style="font-weight:800"':""}>${esc(m)}</td>
    <td class="num" style="font-weight:800">${n}</td>
    <td class="wrap" style="color:var(--mute)">${vd?`<span class="judge j-${tag}">${esc(vd)}</span> `:""}${esc(desc||"")}</td></tr>`).join("")}
  </tbody></table>`;
}
function checkHTML(){
  if(mode==="back") return cycCheckHTML();
  return splitCheckHTML();
  const back = mode==="back", rows = back ? D.backCheck : D.check;
  if(!rows||!rows.length) return `<h2 class="daytitle">세트 검산</h2><p class="daysub">표가 없습니다.</p>`;
  const adj = back && rows.some(r=>r[7]);
  /* 주당 환산(J열)은 주기가 7일이 아닌 재택 시트에만 있다. 판정이 이 값으로 내려지므로 같이 띄운다. */
  const wk  = rows.some(r=>r[8]), span = (adj?4:3) + (wk?1:0);
  return `<h2 class="daytitle">세트 검산</h2>
  <p class="daysub">${back
    ? "세트는 근육군 단위로 센다. 복귀판은 주 6회 · 주 단위라 이 숫자가 곧 주간 세트다. 실질 = 직접 + 간접 추정."
    : `세트는 근육군 단위로 센다. 실질 = 직접 + 간접 추정 — 이건 사이클당(${CYCLE}일) 세트다. 판정은 주당 환산(×7/${CYCLE})으로 내린다.`}</p>
  ${(b=>b?`<div class="budget ${b.left<0?"bad":"ok"}">
    <b>${b.state}</b>
    <span>상한 <i>${b.cap}</i></span><span>현재 <i>${b.cur}</i></span>
    <span>잔여 <i>${b.left}</i></span><span>실제 주당 <i>${b.weekly}</i></span>
    <span>${b.cycle}일 주기</span>
    ${b.left<0?`<div class="warn">다른 종목에서 ${Math.abs(b.left)}세트를 빼기 전에는 추가 금지</div>`:""}
  </div>`:"")(back?D.backBudget:D.budget)}
  <table class="tbl"><thead><tr><th>근육군</th><th>직접</th><th>간접</th><th>실질</th>${wk?"<th>주당</th>":""}<th>판정</th>${adj?"<th>보정</th>":""}</tr></thead><tbody>
  ${rows.map(([m,dr,i,t,tag,v,,g,w])=>`<tr><td>${esc(m)}</td><td class="num">${dr}</td>
    <td class="num" style="color:var(--mute)">${i}</td><td class="num" style="font-weight:800">${t}</td>
    ${wk?`<td class="num">${w||"—"}</td>`:""}
    <td><span class="judge j-${tag}">${esc(v)}</span></td>${adj?`<td class="num" style="color:var(--mute)">${esc(g||"—")}</td>`:""}</tr>`).join("")}
  <tr><td style="font-weight:800">직접 세트 총계</td><td class="num" style="font-weight:800">${rows.reduce((a,r)=>a+r[1],0)}</td><td colspan="${span}"></td></tr>
  </tbody></table>
  ${adj?`<div class="notes"><h5>현장 일 보정</h5><ul>
    <li><b style="color:var(--ink)">높음</b> — 현장 일이 하루치를 얹는다. 판정이 하한 미달이어도 실제로는 채워진다.</li>
    <li><b style="color:var(--ink)">중간</b> — 부분적으로만 채워진다. 판정을 절반쯤 믿는다.</li>
    <li><b style="color:var(--ink)">없음</b> — 현장 일이 안 건드린다. 판정을 액면대로 믿어야 하고, 볼륨을 쓸 곳도 여기다.</li></ul></div>`:""}
  <div class="notes"><h5>간접 출처</h5><ul>${rows.filter(r=>r[6]&&r[6]!=="—")
    .map(r=>`<li><b style="color:var(--ink)">${esc(r[0])}</b> — ${esc(r[6])}</li>`).join("")}</ul></div>
  ${!back&&RAMP?`<p class="footnote">이 표는 설계 볼륨 기준이다 — 램프가 끝난 ${RAMP.rows[RAMP.rows.length-1].week}주차부터의 값.
    이번 주에 실제로 하는 세트는 요일 탭과 주간 탭에서 본다.</p>`:""}`;
}
function rampHTML(){
  if(!RAMP) return "";
  const now=rampRow(week);
  return `<h3 class="sub-h">볼륨 램프 — 종목 복귀</h3><p class="daysub">${esc(RAMP.sub||"")}</p>
  <table class="tbl"><thead><tr><th>주차</th><th>총 세트</th><th>증가폭</th><th class="wrap">이 주차에 복귀하는 종목</th></tr></thead><tbody>
  ${RAMP.rows.map(r=>`<tr style="${r===now?'background:var(--raised)':''}">
    <td>${esc(r.label)}${r===now?' <span class="judge j-lo">현재</span>':''}</td>
    <td class="num" style="font-weight:800">${r.total}</td>
    <td class="num" style="color:var(--mute)">${esc(r.delta)}</td>
    <td class="wrap" style="color:var(--mute)">${esc(r.items)}</td></tr>`).join("")}
  </tbody></table>${notesHTML(RAMP.notes)}`;
}
function planHTML(){
  const p=D.plan;
  /* 요일 카드와 같은 방식으로 다시 센다 — 종목 단위로 빠지므로 합계에 계수를 곱하는 식이 아니다 */
  const wkSets = d => { const day=D.home[d];
    return (!day||day.rest) ? 0 : allEx(day).reduce((a,e)=>a+setsOf(e,week),0); };
  const sum=rampOn()?p.rows.reduce((a,[d,,n])=>a+(n?wkSets(d):0),0):p.total;
  const on=rampOn()&&sum!==p.total;
  const rows=p.rows.map(([d,c,n])=>`<tr class="${n?'':'rest'}"><td>${esc(d)}</td><td class="wrap">${esc(c)}</td>
    <td class="num">${n||"—"}</td>${on?`<td class="num" style="font-weight:800">${n?wkSets(d):"—"}</td>`:""}</tr>`).join("");
  return `<h2 class="daytitle">${CYCLE}일 주기 계획</h2><p class="daysub">${esc(p.sub)}</p>
  ${on?`<div class="rampbar"><b>볼륨 ${vWeek(week)}주차 · ${sum}세트 <span>/ 설계 ${p.total}</span></b>
      <span>오른쪽 열이 이번 주에 실제로 하는 세트다. 중량은 ${week}주차 그대로</span></div>`:""}
  <table class="tbl"><thead><tr><th>일차</th><th style="text-align:left">내용</th><th>${on?"설계":"세트"}</th>${on?"<th>이번 주</th>":""}</tr></thead><tbody>
  ${rows}
  <tr><td style="font-weight:800">합계</td><td></td><td class="num" style="font-weight:800">${p.total}</td>${on?`<td class="num" style="font-weight:800">${sum}</td>`:""}</tr>
  </tbody></table>${rampHTML()}${notesHTML(p.notes)}`;
}
function backPlanHTML(){
  const p=D.backPlan, c=cycDayOf(ymd(new Date()));
  return `<h2 class="daytitle">9일 사이클 계획</h2><p class="daysub">${esc(p.sub)}</p>
  <table class="tbl"><thead><tr><th>일차</th><th style="text-align:left">구성</th><th>세트</th><th>분</th></tr></thead><tbody>
  ${p.rows.map(r=>`<tr class="${r[2]?'':'rest'}" style="${c&&c.day===r[0]?'background:var(--raised)':''}"${r[2]?` data-bcal="${r[0]}"`:""}><td>${esc(r[0])}${c&&c.day===r[0]?' <span class="judge j-lo">오늘</span>':''}</td>
    <td class="wrap" style="font-weight:600;color:var(--ink)">${esc(r[1])}</td><td class="num">${r[2]||"—"}</td><td class="num">${r[3]||"—"}</td></tr>`).join("")}
  <tr><td style="font-weight:800">합계</td><td></td><td class="num" style="font-weight:800">${p.total}</td><td class="num" style="font-weight:800">${p.mins}</td></tr>
  </tbody></table>
  <p class="footnote">주당 환산 ${Math.round(p.total*7/CLEN*10)/10}세트 · 분은 본운동(스트레칭 · 웜업 30분과 샤워 20분은 따로). 줄을 누르면 그 일차로 간다.</p>
  ${gripTableHTML()}
  ${notesHTML(p.notes)}`;
}
function gripTableHTML(){
  if(!GR) return "";
  const c=cycDayOf(ymd(new Date())), k=c?c.round-GR.start:-1;
  const startDay=CANCHOR?dayAdd(CANCHOR,(GR.start-1)*CLEN):"";
  const now = k>=0 ? `지금 ${c.round}사이클 · 회전 ${(k%3)+1}` :
    `${GR.start}사이클${startDay?`(${+startDay.slice(5,7)}/${+startDay.slice(8)})`:""}부터 시작 — 그 전에는 시트 이름의 그립 그대로`;
  const cur = r => k>=0 && (k%3)===r;
  return `<h3 class="sub-h">등 그립 회전</h3><p class="daysub">${now}. 종목 · 세트는 그대로, 그립만 사이클마다 돈다. 종목 카드 · 세션 타이머 · 다음 종목 음성에 자동으로 뜬다.</p>
  <table class="tbl"><thead><tr><th>일차</th><th style="text-align:left">종목</th>${[0,1,2].map(i=>`<th style="${cur(i)?"color:var(--bp)":""}">회전 ${i+1}</th>`).join("")}</tr></thead><tbody>
  ${GR.rows.map(([d,n,per,g,memo,arm,cues])=>`<tr><td>${esc(d)}</td><td class="wrap">${esc(n)}${arm!=null?`<br><b style="font-size:12px">원암 ↔ 양손 2사이클 교대 (${arm?"원암":"양손"} 먼저 · 원암은 언더 차례만)${
      k>=0?` · <span style="color:var(--bp)">이번: ${/^(오버|뉴트럴|패럴렐)/.test(g[k%per])?"양손":((k+arm)%2?"원암":"양손")}</span>`:""}</b>`:""}${memo?`<br><span style="color:var(--mute);font-size:12px">${esc(memo)}</span>`:""}</td>${per===3 ? [0,1,2].map(i=>
      `<td class="wrap" style="${cur(i)?"font-weight:800;color:var(--bp)":""}">${esc(g[i])}${cues&&cues[i]?`<br><span style="color:var(--sq);font-size:12px">▸ ${esc(cues[i])}</span>`:""}</td>`).join("")
    : `<td colspan="3" class="wrap">${per===1?esc(g[0])+" (고정)":`사이클마다 번갈아 — ${g.map(esc).join(" ↔ ")}`}${
        k>=0?` · <b style="color:var(--bp)">이번: ${esc(g[k%per])}</b>`:""}</td>`}</tr>`).join("")}
  </tbody></table>`;
}
/* 생수배송 대비 — 복귀_캐리 · 복귀_스텝박스. 스텝박스 주차는 캐리 1회차 토요일부터 날짜로 센다.
   캐리는 회차(토 · 수) 단위 — 오늘 이후 첫 회차가 오늘(캐리 날) 또는 다음 캐리다 */
/* 캐리 진행표 실제 무게 · 메모 — 폰에서 적는 칸. 증량 기록(wk-log)과 같은 규칙:
   엑셀 E·F열과 같아지면 지운다(옮겨 적었다는 뜻), 빈 값으로 고치면 엑셀 값으로 돌아간다 */
const PGK="wk-prep";
let PLOG={};
try{ const o=JSON.parse(localStorage.getItem(PGK)||"null"); if(o&&typeof o==="object") PLOG=o; }catch(e){}
const pXl=(w,f)=>{ const r=D.prep&&D.prep.carry.prog.find(x=>x[0]===+w); return r?String(r[f==="w"?4:5]||""):""; };
function savePLog(){ try{ localStorage.setItem(PGK,JSON.stringify(PLOG)); }catch(e){} }
(function(){ let ch=false;
  for(const w in PLOG){ for(const f in PLOG[w]) if(PLOG[w][f]===pXl(w,f)){ delete PLOG[w][f]; ch=true; }
    if(!Object.keys(PLOG[w]).length){ delete PLOG[w]; ch=true; } }
  if(ch) savePLog(); })();
const pVal=(w,f)=>{ const l=PLOG[w]&&PLOG[w][f]; return l!==undefined?l:pXl(w,f); };
function setPLog(w,f,v){
  v=String(v==null?"":v).trim().replace(/kg$/i,"");
  if(!PLOG[w]) PLOG[w]={};
  if(v===pXl(w,f)) delete PLOG[w][f]; else PLOG[w][f]=v;
  if(!Object.keys(PLOG[w]).length) delete PLOG[w];
  savePLog();
}
/* 오늘 할 것 — 표 · 타이머가 같은 값을 쓴다. 타이머에는 종목 두 개짜리 세션으로 넘긴다
   (캐리 40초 × 4 · 휴식 120초 → 스텝박스 N분 × M · 휴식 60초). 캐리 마지막 휴식이 스텝박스로 넘어가는 시간이다 */
function prepToday(){
  const P=D.prep, C=P.carry, S=P.stairs, now=new Date();
  const start=C.prog.length?ymdParse(C.prog[0][1]):now;
  const wk=Math.max(1,Math.floor((ymdParse(ymd(now))-start)/864e5/7)+1);
  const dow="일월화수목금토"[now.getDay()];
  const day=S.days.find(r=>r[0]===dow), kind=day?day[1]:"";
  const main=kind.startsWith("메인"), off=!kind||kind.startsWith("없음");
  const tday0=ymd(now), cr=C.prog.find(r=>r[1]>=tday0)||C.prog[C.prog.length-1];
  /* 디로드 주에는 스텝박스도 한 주 멈춘다 — 그 주는 지난주 분량. 캐리는 엑셀 진행표가 이미 멈춰 둔다 */
  const bt=bToday(), dlNow=!!(bt&&bt.deload);
  const p0=C.prog.length?C.prog[0][1]:ymd(now), tday=ymd(now);
  const swk=Math.max(1,wk-DLD.filter(x=>x>=p0&&x<=tday).length);
  /* 스텝박스는 메인 날(토 · 수)마다 한 칸 — 오늘까지 지난 마지막 칸. 짧은 날은 그 칸의 짧은 날 값. 디로드 주 칸은 엑셀이 멈춰 둔다 */
  const sr=S.ramp.filter(r=>r[1]<=tday).pop()||S.ramp[0];
  const carry=cr?{name:"파머스 워크",mg:"캐리",w:pVal(cr[0],"w")||String(cr[2]),r:"",s:4,wk:cr[6]||40,rt:120,drop:0}:null;
  /* 쉬는 날에 스텝박스만 따로 돌리면 짧은 날 분량이다 */
  /* 짧은 날 손 칸(I열)이 차 있으면 이름에 붙인다 — 타이머 화면 · 오늘 줄에서 스트랩을 잊지 않게 */
  const strap=!main&&sr&&sr[8]?sr[8]+" 착용":"";
  const stairs=sr?{name:"스텝박스"+(strap?" · "+strap:main?" · 맨손":""),mg:"스텝",w:fmt(main?sr[2]:sr[5]),r:"",
                   s:main?sr[4]:sr[7],wk:(main?sr[3]:sr[6])*60,rt:60,drop:0}:null;
  const ex=off?[]:[main&&carry,stairs].filter(Boolean);
  const one=e=>e?{t:"생수배송 대비 · "+e.name,ex:[e]}:null;
  /* 집 — 오늘까지 시작한 마지막 단계. 아침 · 저녁 칸에 오늘 요일이 들어 있으면 하는 날이다 */
  const hr=P.home?P.home.rows.filter(r=>r[0]<=tday).pop():null;
  const home=hr?{w:hr[1],q:hr[4],am:hr[2].includes(dow),pm:hr[3].includes(dow)}:null;
  return {wk,swk,dlNow,dow,day,main,off,cr,sr,strap,home,d:{t:"생수배송 대비 · "+dow,ex},carry:one(carry),stairs:one(stairs)};
}
function prepHTML(){
  const P=D.prep, C=P.carry, S=P.stairs;
  const T=prepToday(), {wk,swk,dlNow,dow,day,main,off,cr,sr,strap,home,d}=T, H=P.home;
  const homeLine = !home ? "" : home.am||home.pm
      ? `<span>집 (${[home.am&&"아침",home.pm&&"저녁"].filter(Boolean).join(" · ")}) — 한쪽 ${esc(home.w)} · ${esc(home.q)}${dlNow?" · 디로드 주라 왕복 1번":""}</span>`
      : `<span>집 — 오늘은 쉼</span>`;
  const today = off ? `<b>오늘(${dow}) · 쉬는 날</b><span>${esc(day?day[2]:"")}</span>`
    : `<b>오늘(${dow}) · ${dlNow?"디로드 주 — 캐리 · 스텝박스 무게와 분량 멈춤":swk+"주차"}</b><span>${main&&cr?`캐리 한 손 ${fmt(cr[2])}kg · ${cr[6]||40}초 × 4 (휴식 120초) → `:""}스텝박스 ${
        main?`${sr[3]}분 × ${sr[4]}`:`${sr[6]}분 × ${sr[7]}`} · 덤벨 한 손 ${fmt(main?sr[2]:sr[5])}kg ${strap?`<b style="color:var(--sq)">· ${esc(strap)}</b>`:"· 맨손"} (휴식 60초)</span>`;
  /* 세션 끝 휴식은 타이머도 빼므로 마지막 종목의 휴식 하나를 덜어 센다 */
  const mins=x=>{ const L=x.ex[x.ex.length-1]; return Math.round((x.ex.reduce((a,e)=>a+e.s*(e.wk+e.rt),0)-(L?L.rt:0))/60); };
  const sub=`background:var(--surface);color:var(--ink);flex:1;margin:0`;
  const part=(k,x,lab)=>x?`<button class="day-go" style="${sub}" data-prep-tm="${k}">&#9654; ${lab}만<i style="display:block">${
    x.ex[0].s} × ${x.ex[0].wk>=60?x.ex[0].wk/60+"분":x.ex[0].wk+"초"} · 약 ${mins(x)}분</i></button>`:"";
  const go = (off ? "" : `<button class="day-go" style="margin-top:12px" data-prep-tm="today">&#9654; 오늘 세션 · 약 ${mins(d)}분</button>`)
    + `<div style="display:flex;gap:8px;margin:${off?"12px":"0"} 0 15px">${part("carry",T.carry,"캐리")}${part("stairs",T.stairs,"스텝박스")}</div>`;
  const md=t=>{ const d=ymdParse(t); return `${d.getMonth()+1}/${d.getDate()}`; };
  const hr0=X=>X.rows.filter(r=>r[0]<=ymd(new Date())).pop();
  /* 디로드 주(시작 토요일부터 7일) 안의 회차 — 무게가 멈춘 줄이라 번호에서 뺀다 */
  const inDL=t=>DLD.some(x=>{ const g=(ymdParse(t)-ymdParse(x))/864e5; return g>=0&&g<7; });
  return `<h2 class="daytitle">생수배송 대비</h2><p class="daysub">목표: 한 손 35kg(2L 묶음 3팩) 컨트롤 · 세트 합계 밖</p>
  <div class="rampbar">${today}${homeLine}</div>${go}
  <h5 style="font-size:13px;font-weight:800;margin:22px 0 4px">캐리 — 무겁게 들고 걷기</h5><p class="daysub">${esc(C.sub)}</p>
  <table class="tbl"><thead><tr><th>회차</th><th>날짜</th><th>한 손</th><th>양손</th><th>실제</th><th class="wrap">메모</th></tr></thead><tbody>
  ${C.prog.map(r=>`<tr${r===cr?' style="background:var(--surface)"':""}><td>${inDL(r[1])?"디로드":r[0]-C.prog.filter(x=>x[0]<r[0]&&inDL(x[1])).length}${r===cr?(r[1]===ymd(new Date())?" · 오늘":" · 다음"):""}</td>
    <td class="num">${md(r[1])} ${"일월화수목금토"[ymdParse(r[1]).getDay()]}</td><td class="num" style="font-weight:800;color:var(--ink)">${fmt(r[2])}<span style="font-weight:400;color:var(--mute)"> · ${r[6]||40}초</span></td>
    <td class="num" style="color:var(--mute)">${fmt(r[3])}</td>
    <td class="num"><input class="log-in" style="width:58px" type="text" inputmode="decimal" data-prepf="w" data-prepw="${r[0]}" value="${escAttr(pVal(r[0],"w"))}" placeholder="—"></td>
    <td class="wrap"><input class="log-in wide" style="width:100%;min-width:70px" type="text" data-prepf="m" data-prepw="${r[0]}" value="${escAttr(pVal(r[0],"m"))}" placeholder="—"></td></tr>`).join("")}
  </tbody></table>${C.foot?`<p class="daysub">${esc(C.foot)}</p>`:""}
  ${Object.keys(PLOG).length?`<button class="day-go" id="prep-copy" style="background:var(--surface);color:var(--ink)">엑셀에 옮길 값 복사 (복귀_캐리 E·F열)</button>`:""}
  <table class="tbl"><tbody>${C.rules.map(r=>`<tr><td>${esc(r[0])}</td><td class="wrap" style="color:var(--ink)">${esc(r[1])}<br><span style="color:var(--mute)">${esc(r[2])}</span></td></tr>`).join("")}</tbody></table>
  <h5 style="font-size:13px;font-weight:800;margin:22px 0 4px">스텝박스 + 덤벨</h5><p class="daysub">${esc(S.sub)}</p>
  <table class="tbl"><thead><tr><th>회차</th><th>날짜</th><th>메인 (토·수) 맨손</th><th>짧은 날 (일·월·목)</th></tr></thead><tbody>
  ${S.ramp.map(r=>`<tr${r===sr?' style="background:var(--surface)"':""}><td>${inDL(r[1])?"디로드":r[0]-S.ramp.filter(x=>x[0]<r[0]&&inDL(x[1])).length}${r===sr?(r[1]>ymd(new Date())?" · 다음":" · 지금"):""}</td>
    <td class="num">${md(r[1])} ${"일월화수목금토"[ymdParse(r[1]).getDay()]}</td>
    <td class="num"><b style="color:var(--ink)">${fmt(r[2])}kg</b> · ${r[3]}분 × ${r[4]}</td><td class="num" style="color:var(--mute)">${fmt(r[5])}kg · ${r[6]}분 × ${r[7]}${r[8]?` · <b style="color:var(--sq)">${esc(r[8])}</b>`:""}</td></tr>`).join("")}
  </tbody></table>${S.foot?`<p class="daysub">${esc(S.foot)}</p>`:""}
  <table class="tbl"><tbody>${S.days.map(r=>`<tr class="${r[1].startsWith("없음")?"rest":""}"><td>${esc(r[0])}${r[0]===dow?" · 오늘":""}</td>
    <td class="wrap" style="color:var(--ink)">${esc(r[1])}</td><td class="wrap">${esc(r[2])}</td></tr>`).join("")}</tbody></table>
  ${notesHTML(S.notes)}
  ${H?`<h5 style="font-size:13px;font-weight:800;margin:22px 0 4px">집 — 생수팩 + 계단</h5><p class="daysub">${esc(H.sub)}</p>
  <table class="tbl"><thead><tr><th>시작</th><th class="wrap">한쪽</th><th>아침</th><th>저녁</th></tr></thead><tbody>
  ${H.rows.map(r=>`<tr${hr0(H)===r?' style="background:var(--surface)"':""}><td class="num">${md(r[0])}</td><td class="wrap" style="color:var(--ink)">${esc(r[1])}<br><span style="color:var(--mute)">${esc(r[4])}</span></td>
    <td class="wrap">${esc(r[2])}</td><td class="wrap">${esc(r[3])}</td></tr>`).join("")}
  </tbody></table>${notesHTML(H.notes)}`:""}`;
}
/* 그 주 다음으로 무게가 실제로 오르는 주차 — 주기 2 · 3인 보조 종목에서 '언제 오르나'가 정보다 */
function nextUp(k,c){ const v=workWeight(k,c); for(let w=c+1;w<=WLAST;w++) if(workWeight(k,w)!==v) return w; return 0; }
function progressHTML(week){
  /* 메인 넷은 표로 · 보조(2026-10-23 · 9종목)는 요일별 목록으로 — 13열 표는 폰에서 안 읽힌다 */
  const all=Object.keys(LIFT), keys=all.filter(k=>!LIFT[k].acc), acc=all.filter(k=>LIFT[k].acc);
  const c=week, fat=logVal(c,"fat");
  const kgOf = (k,w) => { const v=workWeight(k,w); return v===0&&LIFT[k].acc ? "빈 바" : fmt(v)+"kg"; };
  /* 날짜·'지금'은 재택 일차 회전에서 나온다. 복귀는 달력이 주차를 세므로 붙이지 않는다 */
  const bk=true, cur=null, date=null;
  const card=`<section class="logcard">
    <div class="logcard-h"><b>${c}주차 기록</b><span>${date?`${+date.slice(5,7)}/${+date.slice(8)} 시작`:""}${
      cur&&cur.cycle===c?" · 지금":""}</span></div>
    ${(()=>{ const row = k => { const dl=deloadNote(k,c), L=LIFT[k], nx=L.acc?nextUp(k,c):0;
      return `<div class="log-row" style="--accent:${L.accent}">
      <div class="log-l"><b>${esc(L.ko)}</b><span class="num">${kgOf(k,c)} × ${L.reps}회 × ${L.sets}세트${
        missRun(k,c)>0?" · 앞 주 미달로 유지":""}${L.acc&&L.every>1?` · ${L.every}주마다${nx?` · 다음 ${nx}주차`:""}`:""}</span>${dl?`<em>${esc(dl)}</em>`:""}</div>
      ${logCtl(k,c)}</div>`; };
      const days=D.backOrder.filter(d=>acc.some(k=>LIFT[k].day===d));
      return keys.map(row).join("") + days.map(d=>
        `<div class="log-row" style="padding-top:14px"><div class="log-l"><b style="color:var(--mute)">보조 · ${d}일차</b></div></div>`+
        acc.filter(k=>LIFT[k].day===d&&c>=(LIFT[k].from||WBASE)).map(row).join("")).join(""); })()}
    <div class="log-row"><div class="log-l"><b>다음날 피로</b><span>0 가뿐 · 3 뻐근하지만 훈련 가능 · 5 탈진. 다음 세션 아침에</span></div>
      <div class="seg fat">${[0,1,2,3,4,5].map(n=>`<button class="${fat===String(n)?'on':''}" data-logf="fat" data-logc="${c}" data-logv="${n}">${n}</button>`).join("")}</div></div>
    <div class="log-row"><div class="log-l"><b>관절 이상</b><span>없으면 비워둔다 · 부위와 0~10</span></div>
      <input class="log-in wide" type="text" data-logf="joint" data-logc="${c}" value="${escAttr(logVal(c,"joint"))}" placeholder="예: 무릎 3"></div>
    <div class="log-row"><div class="log-l"><b>스트랩</b><span>쓴 종목 · 세트 · 중간에 감았으면 '중간' — 증량 판정과 무관</span></div>
      <input class="log-in wide" type="text" data-logf="strap" data-logc="${c}" value="${escAttr(logVal(c,"strap"))}" placeholder="예: 데드 3세트 · 랫풀다운 4세트 중간"></div>
    <div class="log-row"><div class="log-l"><b>메모</b></div>
      <input class="log-in wide" type="text" data-logf="memo" data-logc="${c}" value="${escAttr(logVal(c,"memo"))}" placeholder="—"></div>
  </section>`;

  const mark = v => v==="" ? "" : isDone(v) ? `<i class="lm ok">달성</i>`
    : `<i class="lm bad">${/^\d+$/.test(v)?v+"회":"미달"}</i>`;
  const rows=[];
  for(let w=WBASE;w<=WLAST;w++){
    const j=logVal(w,"joint"), f=logVal(w,"fat");
    rows.push(`<tr data-logrow="${w}" style="${w===week?'background:var(--raised)':''}"><td class="num">${w}</td>`+
      keys.map(k=>`<td class="num">${w<(LIFT[k].from||WBASE)?"—":fmt(workWeight(k,w))+mark(logVal(w,k))}</td>`).join("")+
      `<td class="num">${esc(f)}${j?`<i class="lm bad">${esc(j)}</i>`:""}</td></tr>`);
  }
  const mine=phoneOnly(), lines=logLines();
  return `<h2 class="daytitle">증량 기록</h2>
  <p class="daysub">엑셀 증량기록과 같은 규칙 — 달성·공란은 다음 주에 오르고, 미달을 적으면 그 중량을 한 번 더 든다.
    위 −/+ 로 주차를 옮긴다. 기억으로 소급해 채우지 말 것.${bk
      ? " 9일 사이클마다 한 칸씩 오른다 — 올리기 전에 이 칸의 결과를 먼저 적는다."
      : ""}</p>
  ${card}
  <table class="tbl"><thead><tr><th>주차</th>${keys.map(k=>`<th>${esc(LIFT[k].ko)}</th>`).join("")}<th>피로</th></tr></thead>
  <tbody>${rows.join("")}</tbody></table>
  <p class="footnote">표의 중량은 위 기록을 반영한 값이다 — 미달 뒤 주는 오르지 않는다. 줄을 누르면 그 주를 적는다.</p>
  ${mine.length?`<div class="ovbox"><h5>이 폰에만 있는 기록 <span>· 엑셀 증량기록에 옮겨 적으면 여기서 사라진다</span></h5>
    <ul>${mine.map(t=>`<li>${esc(t)}</li>`).join("")}</ul>
    <button id="log-copy" class="plain">전체 기록 텍스트로 복사</button>
    <details class="log-text"><summary>복사가 안 되면 여기서 길게 눌러 복사</summary><pre>${esc(lines.join("\n"))}</pre></details></div>`:""}
  <div class="notes"><h5>주당 증량폭</h5><ul>${all.filter(k=>!LIFT[k].acc||LIFT[k].day).map(k=>
    `<li>${esc(LIFT[k].ko)} ${fmt(LIFT[k].inc)}kg${LIFT[k].every>1?` · ${LIFT[k].every}주마다`:""}${LIFT[k].acc?` · ${LIFT[k].day}일차`:""}</li>`).join("")}</ul><p class="footnote">한쪽 원판이 2.5 · 1 · 0.5kg 하나씩이라 총 중량이 5로 나눈 나머지 4면 못 끼운다 — 그 칸은 자동으로 1kg 올라가므로 폭이 칸마다 조금씩 다르다.</p></div>
  ${notesHTML(D.progressNotes)}`;
}
/* 사이클별 한 줄 — 엑셀 증량기록에 옮겨 적기 위한 텍스트. 기록이 있는 사이클만 */
function logLines(){
  const keys=Object.keys(LIFT), out=[];
  for(let c=WBASE;c<=WLAST;c++){
    const p=[];
    keys.forEach(k=>{ const v=logVal(c,k); if(v==="") return;
      p.push(LIFT[k].ko+" "+(isDone(v)?"O":(/^\d+$/.test(v)?v+"회":"X"))); });
    for(const f in LOGF){ const v=logVal(c,f); if(v) p.push(LOGF[f]+" "+v); }
    if(p.length) out.push(`${c}사이클${cycleDate(c)?" ("+cycleDate(c)+")":""} · `+p.join(" · "));
  }
  return out;
}
/* 폰에 있고 엑셀엔 없는(또는 다른) 칸 */
function phoneOnly(){
  const items=[];
  Object.keys(LOG).map(Number).sort((a,b)=>a-b).forEach(c=>{ for(const f in LOG[c]){
    const x=xlVal(c,f), v=LOG[c][f];
    const name=LIFT[f]?LIFT[f].ko:LOGF[f]||f;
    items.push(`${c}사이클 ${name} — ${v===""?"비움":(LIFT[f]&&isDone(v))?"달성(O)":v}${x?` (엑셀에는 ${x})`:""}`);
  }});
  return items;
}
function logCopy(btn){
  const text=logLines().join("\n");
  const ok=()=>{ btn.textContent="복사됨"; setTimeout(()=>{ btn.textContent="전체 기록 텍스트로 복사"; },1500); };
  if(navigator.clipboard&&navigator.clipboard.writeText){ navigator.clipboard.writeText(text).then(ok,()=>{ btn.textContent="복사 실패 — 아래에서 길게 눌러 복사"; }); return; }
  try{ const ta=document.createElement("textarea"); ta.value=text; document.body.appendChild(ta); ta.select();
    document.execCommand("copy"); ta.remove(); ok(); }catch(e){ btn.textContent="복사 실패 — 아래에서 길게 눌러 복사"; }
}

/* ── 중량 편집 UI ── 숫자를 탭 → 입력 → 확인(엔터·바깥 탭)으로 저장.
   빈 값으로 확인하면 엑셀 값으로 되돌아간다. */
function grabInput(el, value, numeric, onCommit, after){
  if(el.querySelector("input")) return;
  el.innerHTML=`<input class="w-input" type="text"${numeric?' inputmode="decimal"':''}`+
    ` value="${escAttr(value)}" placeholder="—">`;
  const inp=el.querySelector("input");
  inp.focus(); inp.select();
  let done=false;
  const finish=commit=>{ if(done) return; done=true; if(commit) onCommit(inp.value.trim());
    render(); if(after) after(); };
  inp.onblur=()=>finish(true);
  inp.onkeydown=ev=>{ if(ev.key==="Enter") finish(true); else if(ev.key==="Escape") finish(false); };
}
function editLift(el,k,after){
  k=k||el.dataset.lift;
  grabInput(el, fmt(workWeight(k,week)), true, v=>{
    const n=parseFloat(v);
    if(v===""){ delete OV.lifts[k]; }
    else if(!isNaN(n)&&n>0&&LIFT[k].of){
      const b=workWeight(LIFT[k].of,week), r=b>0 ? Math.round(n/b*1000)/1000 : LIFT[k].ratio;
      if(Math.abs(r-LIFT[k].ratio)<0.001) delete OV.lifts[k];
      else OV.lifts[k]={ratio:r, start:LIFT[k].start, base:LIFT[k].start};
    }
    else if(!isNaN(n)&&n>0){
      // 이번 사이클 무게를 입력받아 시작중량을 역산 — 이후 사이클·원판·웜업이 전부 따라온다.
      // 미달로 유지된 사이클은 기록이 이미 빼고 세므로 여기서도 같은 계단 수를 쓴다
      const start=startFor(k,week,n);
      if(Math.abs(start-LIFT[k].start)<0.001) delete OV.lifts[k];
      else OV.lifts[k]={start:start, base:LIFT[k].start};
    }
    saveOV();
  }, after);
}
function editEx(el,key,after,fill,hint){
  key=key||el.dataset.exkey;
  const e=exByKey(key);
  if(!e) return;
  const cur=fill!=null?fill:OV.ex[key]?OV.ex[key].w:e.w;
  /* 마지막 단계 다음은 새 숫자를 받는다 — 안 고치고 닫으면 아무것도 안 바꾼다 */
  grabInput(el, cur==="—"?"":cur, false, v=>{ if(fill!=null&&v===fill) return; setExW(key,v); }, after);
  const inp=el.querySelector("input");
  if(inp&&hint){ inp.title=hint; inp.placeholder=hint;
    el.insertAdjacentHTML("beforeend",`<span class="ovnote">${esc(hint)}</span>`); }
}
function ovHTML(){
  const items=[];
  for(const k in OV.lifts)
    items.push(OV.lifts[k].ratio ? `${LIFT[k].ko} — 비율 ${LIFT[k].ratio} → ${OV.lifts[k].ratio}`
                                 : `${LIFT[k].ko} — 시작중량 ${fmt(OV.lifts[k].base)} → ${fmt(OV.lifts[k].start)}kg`);
  for(const k in OV.ex){
    const p=k.split("|");
    items.push(`${p[0]==="back"?"복귀 ":""}${p[1]} · ${p[2]} — ${OV.ex[k].base} → ${OV.ex[k].w}`);
  }
  for(const k in OVT){
    const p=k.split("|");
    items.push(`${p[0]==="back"?"복귀 ":""}${p[1]} · ${p[2]} — 수행 ${OVT[k].base}초 → ${OVT[k].wk}초 (G열)`);
  }
  if(!items.length) return "";
  return `<div class="ovbox"><h5>이 폰에서 수정한 값 <span>· 엑셀 원본은 그대로이니 나중에 옮겨 적으세요</span></h5>
    <ul>${items.map(t=>`<li>${esc(t)}</li>`).join("")}</ul>
    <button id="ov-reset">전부 엑셀 값으로 되돌리기</button></div>`;
}
/* ── 인터벌 타이머 ────────────────────────────────────────────────
   엑셀 G·H열(수행 초 · 휴식 초)을 그대로 읽어 종목 → 세트 순으로 걷는다.
   남은 시간은 setInterval 누적이 아니라 목표 시각과의 차로 낸다 —
   화면이 꺼져 콜백이 밀려도 돌아왔을 때 숫자가 어긋나지 않는다.
   화면 잠금은 Wake Lock으로 막는다. 브라우저가 백그라운드에서 소리를
   막을 수 있으므로 화면을 켜둔 채 쓰는 것이 전제다. */
const PH={prep:["준비","var(--bp)"],work:["운동","var(--mp)"],rest:["휴식","var(--sq)"]};
const PREP=5;
/* 2026-11-07 사용자 요청 — 예열 1세트를 타이머에 따로 넣는다(전에는 앞 종목 휴식에 끼워 넣어야 했다).
   예열 종목(e.wu)의 1세트 앞에 '예열' 운동 구간 + 휴식을 둔다. 50% × 6~8은 본세트보다 짧으니
   수행 초는 그 종목 수행 초와 60초 중 짧은 쪽, 뒤 휴식은 60초. 시간 탭 합계에도 그대로 들어간다. */
const WU_WK_MAX=60, WU_REST=60;
const mmss=v=>{const m=Math.floor(v/60),x=v%60;return (m<10?"0":"")+m+":"+(x<10?"0":"")+x;};
/* 세트별 휴식. e.rl(엑셀 '세트별 휴식(초)')이 있으면 그 배열을, 없으면 e.rt
   하나를 전 세트에 쓴다. 복귀 램프로 세트가 배열보다 늘어난 경우에는 마지막
   값을 이어 쓴다 — 오름차순이므로 '뒤로 갈수록 길어진다'가 유지된다. */
const restAt=(e,k)=>{const L=e.rl; return (L&&L.length)?L[Math.min(k,L.length)-1]:e.rt;};
const restLabel=e=>{const L=e.rl;
  return (L&&L.length>1)?mmss(L[0])+" → "+mmss(L[L.length-1]):mmss(e.rt);};
let TM=null, tmInt=null, tmAC=null, tmWL=null, tmLeft=-1;

function tmSteps(d,week,from,Sf){
  const S=Sf||(e=>effSets(e,week));
  const live=allEx(d).filter(e=>e.s!==0&&S(e)>0&&e.wk>0);
  /* 슈퍼세트 — 같은 태그가 붙은 이웃 종목끼리 한 묶음이 되어 세트를 번갈아
     돈다 (A1 → A2 → 휴식 → A1 → A2 → …). 휴식(초)의 뜻은 그대로 "이 종목
     세트 뒤에 오는 쉼"이라, 묶음 앞쪽에는 전환 시간을, 마지막 종목에는
     라운드 간 휴식을 적으면 된다. 세트 수가 서로 달라도 된다 — 모자란 쪽은
     그 라운드에서 빠진다. */
  const groups=[];
  live.forEach(e=>{ const g=groups[groups.length-1];
    if(g&&e.ss&&g[0].ss===e.ss) g.push(e); else groups.push([e]); });
  const st=[], IN=mode==="back"&&D.wuIn?D.wuIn[dayKey(d)]:null;
  groups.forEach(g=>{ const n=Math.max.apply(null,g.map(S));
    for(let k=1;k<=n;k++) g.forEach(e=>{
      if(k>S(e)) return;
      /* 드롭이 걸린 종목(엑셀 O열)은 마지막 세트만 그 시간만큼 길게 잡고 표시를
         남긴다 — 세트 시작 때 '드롭세트'라고 읊고 화면에도 적는다 */
      const dp=(k===S(e)&&e.drop>0);
      if(k===1&&IN&&IN.before===e.name) IN.items.forEach(x=>{
        let w=x.w;
        if(x.pct){ const m=(d.mains||[])[0]; if(!m||!LIFT[m[0]]) return;
                   const v=Math.max(20,R(workWeight(m[0],week)*x.pct)); if(v<=20) return; w=String(v); }
        const pe={name:x.n,mg:x.mg,w:w,r:x.r,wk:x.wk,of:e};
        for(let j=1;j<=x.s;j++){ st.push({p:"work",t:x.wk,e:pe,set:j,sets:x.s,ramp:true});
                                 st.push({p:"rest",t:x.rt,e:pe,set:j,sets:x.s,ramp:true}); } });
      if(k===1&&e.wu){
        st.push({p:"work",t:Math.min(WU_WK_MAX,armWk(e,e.wk)),e:e,set:1,sets:S(e),wu:true});
        st.push({p:"rest",t:WU_REST,e:e,set:1,sets:S(e),wu:true});
      }
      st.push({p:"work",t:armWk(e,e.wk)+(dp?e.drop:0),e:e,set:k,sets:S(e),drop:dp});
      const rt=restAt(e,k);
      if(rt>0) st.push({p:"rest",t:rt,e:e,set:k,sets:S(e)});
    });});
  while(st.length&&st[st.length-1].p==="rest") st.pop();   /* 세션 끝의 휴식은 뺀다 */
  let cut=0;
  if(from){
    /* 슈퍼세트는 어느 짝을 눌러도 그 묶음의 처음부터 시작한다 —
       라운드 중간에서 시작하면 짝이 어긋난다 */
    const fe=live.find(e=>e.name===from), g=fe?groups.find(x=>x.indexOf(fe)>=0):null;
    const j=st.findIndex(x=>x.e.name===(g?g[0].name:from)); if(j>0) cut=j;
  }
  const out=st.slice(cut);
  if(out.length) out.unshift({p:"prep",t:PREP,e:out[0].e,set:out[0].set,sets:out[0].sets});
  out.forEach(x=>{ x.li=live.indexOf(x.e.of||x.e)+1; x.ln=live.length; });
  return out;
}
function tmW(e){
  if(e.w==="위 참조"&&e.lift) return fmt(workWeight(e.lift,week))+"kg";
  if(e.lo||e.pct||e.trk) return exW(e);
  const ov=OV.ex[mode+"|"+tab+"|"+e.name], w=String(ov?ov.w:e.w);
  return /^[\d.~\s]+$/.test(w) ? w.replace(/\s+$/,"")+"kg" : w;
}
/* -- 화면이 꺼져도 알림이 울리게 하는 두 가지 --------------------
   (1) 무음 유지음: 사람 귀에 안 들리는 아주 작은 소리를 계속 흘린다.
       크롬은 '소리를 내는 탭'을 얼리지 않고, 타이머도 분당 1회까지 늦추지 않는다.
       화면을 꺼도 오디오 스레드는 계속 돈다 — 웹 라디오가 꺼진 화면에서
       계속 나오는 것과 같은 원리다.
   (2) 미리 예약: 알림음을 JS가 그 시각에 깨어나서 울리는 게 아니라
       오디오 시계에 90초 앞까지 미리 걸어둔다. JS가 얼마나 늦어지든
       예약된 소리는 하드웨어가 정시에 낸다.
   진동은 예약이 안 되고 화면이 꺼져 있으면 브라우저가 막으므로 보조 수단이다. */
const LOOKAHEAD=90000;
let tmKeep=null, tmPend=[], tmSchedAt=0, tmIdle=null;

/* 말소리도 삐 소리와 같은 오디오 시계에 건다.
   기기 TTS(speechSynthesis)를 안 쓰는 이유: 화면이 꺼지면 크롬이
   not-allowed로 통째로 거부한다 — 정작 필요한 순간에 안 나온다.
   미리 구운 wav라 화면이 꺼져 있어도 예약대로 난다. */
/* 말은 한 파일에 죽 이어 붙여 두고(스프라이트) 필요한 구간만 잘라 쓴다.
   조각을 파일로 두면 종목 이름·숫자까지 200개가 넘어 첫 방문에 요청이
   200번 난다. V.map["랫풀다운"] = [시작초, 길이초] — 그게 전부다. */
const V=D.voice||null;
let vBuf=null, vBusy=false;
let voiceOn=true; try{ voiceOn=localStorage.getItem("wk-voice")!=="0"; }catch(e){}
function tmVoiceLoad(){
  if(!tmAC||!V||vBuf||vBusy) return;
  vBusy=true;
  fetch(V.src).then(r=>r.arrayBuffer()).then(b=>tmAC.decodeAudioData(b))
    .then(buf=>{ vBuf=buf; if(TM&&TM.run) tmSchedule(); })
    .catch(()=>{ vBusy=false; });                    /* 못 받으면 삐 소리만 */
}
const tmHas = k => !!(V&&V.map[k]);
const tmDur = k => tmHas(k) ? V.map[k][1] : 0;
function tmSay(kind,at){
  if(!voiceOn||!vBuf||!tmHas(kind)) return null;
  const sl=V.map[kind];
  try{
    const src=tmAC.createBufferSource();
    src.buffer=vBuf; src.connect(tmAC.destination);
    src.start(at,sl[0],sl[1]);
    return src;
  }catch(e){ return null; }
}
/* 조각을 차례로 건다. ","는 쉼표 — 소리는 없고 사이만 벌린다.
   낱말 사이를 너무 붙이면 한 덩어리로 뭉쳐 들려 무슨 말인지 안 잡힌다. */
const VGAP=.13, VPAUSE=.3;
function tmSayAll(keys,at){
  const n=[]; let t=at;
  keys.forEach(k=>{
    if(k===","){ t+=VPAUSE; return; }
    const s=tmSay(k,t);
    if(s){ n.push(s); t+=tmDur(k)+VGAP; }
  });
  return n;
}

/* -- 숫자를 조각 이름으로 --------------------------------------
   중량은 주차마다 오르고 앱에서 고칠 수도 있어서 문장을 미리 구워 둘 수가
   없다. 그래서 0~99와 백 단위만 구워 두고 여기서 이어 붙인다.
   규칙은 voice_build.py의 sino()/num_keys()와 한 글자도 다르면 안 된다 —
   한쪽만 고치면 앱이 조각을 못 찾아 그 부분이 통째로 조용해진다. */
const KO1=["","일","이","삼","사","오","육","칠","팔","구"];
const KOF={"5":"점오","25":"점이오","75":"점칠오","125":"점일이오"};
function koTens(n){
  if(n===0) return "영";
  const t=Math.floor(n/10), o=n%10;
  return (t===0?"":t===1?"십":KO1[t]+"십")+KO1[o];
}
function koNum(s){
  const m=/^(\d+)(?:\.(\d+))?$/.exec(String(s).trim());
  if(!m) return null;
  const n=+m[1], f=(m[2]||"").replace(/0+$/,"");
  if(n>=1000) return null;
  const k=[], h=Math.floor(n/100);
  if(h){ k.push((h===1?"":KO1[h])+"백"); if(n%100) k.push(koTens(n%100)); }
  else k.push(koTens(n));
  if(f){
    if(KOF[f]) k.push(KOF[f]);
    else { k.push("점"); for(const c of f) k.push(koTens(+c)); }
  }
  return k;
}
/* '40' → 사십 킬로 · '30~35' → 삼십 에서 삼십오 킬로 · 숫자가 아니면 null */
function koRange(s,unit){
  const p=String(s).trim().split("~");
  if(p.length>2) return null;
  const a=koNum(p[0]); if(!a) return null;
  if(p.length===1) return a.concat([unit]);
  const b=koNum(p[1]); if(!b) return null;
  return a.concat(["에서"],b,[unit]);
}
/* 숫자 꼴이면 조립하고, 아니면 통째로 구워 둔 칸('맨몸')을 찾는다 */
function koCell(s,unit){
  return koMic(s) || koRange(s,unit) || (tmHas(s)?[s]:[]);
}
/* 미세 단계 무게 — '5 + 핀 + 1' → 오 킬로 · 더하기 핀 · 더하기 · 일, '8 + 손목 1' → 팔 킬로 · 더하기 손목 · 일 */
function koMic(s){
  const m=/^(\d+(?:\.\d+)?)\s*\+\s*(핀|손목)\s*(?:\+?\s*(\d+(?:\.\d+)?))?$/.exec(String(s).trim());
  if(!m||!koNum(m[1])) return null;
  const k=koNum(m[1]).concat(["킬로","더하기 "+m[2]]);
  if(m[3]){ if(m[2]==="핀") k.push("더하기"); k.push.apply(k,koNum(m[3])||[]); }
  return k;
}
/* 다음 종목 안내. 값은 그 순간 화면에 떠 있는 것을 그대로 읽으므로
   주차 증량도 앱에서 고친 중량도 따로 손댈 데가 없다. */
function tmNextKeys(e,sets,set){
  const gp = mode==="back" ? gripOf(e) : null;
  const k=["다음 종목",",",gp?exName(e):e.name];
  if(gp){ k.push(",","그립:"+gp.g); if(gp.ax) k.push("팔:"+gp.arm); }
  if(e.mg) k.push(",",e.mg);
  k.push(",");
  k.push.apply(k,(koNum(String(sets))||[]).concat(["세트"]));
  k.push(",");
  k.push.apply(k,koCell(String(e.r||""),"회"));
  k.push(",");
  k.push.apply(k,koCell(tmW(e).replace(/kg$/,""),"킬로"));
  if(e.wu&&set===1) k.push(",","예열 한 세트 먼저");   /* 2026-11-06 슈퍼세트는 라운드마다 종목이 바뀌어 매번 읊었다 — 1세트 앞에서만 */
  return k.filter(x=>x===","||tmHas(x));
}
const tmLen = keys => keys.reduce(
  (a,k)=>a+(k===","?VPAUSE:tmDur(k)+VGAP), 0);
/* 이 휴식 다음에 종목이 바뀔 때만 읊는다. 같은 종목의 다음 세트로
   이어질 뿐이면 조용하다 — 매 세트 떠들면 그냥 소음이다.
   휴식이 끝나갈 때 나는 카운트다운에 말이 묻히지 않도록, 그 전에 다 끝낼 수
   있을 때만 입을 뗀다. 5초짜리 '준비'에서 읊지 않는 것도 같은 이유다 —
   시작 버튼을 누른 그 순간엔 화면을 보고 있다. */
function tmNextSay(idx,at,off){
  if(!TM||idx==null||!voiceOn||!vBuf) return [];
  const cur=TM.st[idx], nx=TM.st[idx+1];
  if(!cur||!nx||nx.e===cur.e) return [];
  const k=tmNextKeys(nx.e,nx.sets,nx.set);
  /* 카운트다운이 시작되기 1초 전까지는 끝나야 한다. 리드가 길어진 만큼
     이 여유도 같이 늘어난다 — 안 그러면 안내가 딸깍 위에 얹힌다. */
  if(off+tmLen(k)>cur.t-(tmLead(cur,nx)+1)) return [];
  return tmSayAll(k,at);
}

const tmVib=p=>{ try{ if(document.visibilityState==="visible"&&navigator.vibrate) navigator.vibrate(p); }catch(e){} };

function tmKeepStart(){
  try{
    const b=tmAC.createBuffer(1,tmAC.sampleRate,tmAC.sampleRate), ch=b.getChannelData(0);
    for(let i=0;i<ch.length;i++) ch[i]=(Math.random()*2-1)*2e-4;   /* -74dB — 안 들린다 */
    tmKeep=tmAC.createBufferSource(); tmKeep.buffer=b; tmKeep.loop=true;
    tmKeep.connect(tmAC.destination); tmKeep.start();
  }catch(e){}
}
/* quiet=true — 컨텍스트만 만들고 말소리를 미리 디코딩한다. 소리는 안 낸다.
   페이지가 열릴 때 한 번 이렇게 불러두면 첫 '준비'부터 말이 나온다. */
function tmAudio(quiet){
  clearTimeout(tmIdle);
  if(!tmAC){
    try{ tmAC=new (window.AudioContext||window.webkitAudioContext)(); }
    catch(e){ tmAC=null; return null; }
    tmVoiceLoad();
  }
  if(!quiet){
    if(tmAC.state==="suspended") tmAC.resume();
    if(!tmKeep) tmKeepStart();
    tmVoiceLoad();
  }
  return tmAC;
}
function tmTone(f,dur,vol,at){
  try{
    const o=tmAC.createOscillator(), g=tmAC.createGain();
    o.type="sine"; o.frequency.value=f;
    g.gain.setValueAtTime(vol,at); g.gain.exponentialRampToValueAtTime(.0001,at+dur);
    o.connect(g); g.connect(tmAC.destination);
    o.start(at); o.stop(at+dur+.05);
    return o;
  }catch(e){ return null; }
}
/* 삐 소리와 말 사이의 간격. 삐 소리는 여운이 길어서 바짝 붙이면 말이
   그 위에 얹혀 안 들린다 — 소리가 잦아든 뒤에 입을 뗀다. */
const VAFTER=.36;
function tmCueAt(kind,wall,idx){
  if(!tmAC) return;
  const at=tmAC.currentTime+Math.max(0,(wall-Date.now())/1000), n=[];
  let say=[];
  /* 삐 소리가 먼저 귀를 잡고, 뒤이어 무엇인지 말한다 */
  if(kind==="work"){ n.push(tmTone(1046,.35,.6,at)); n.push(tmSay("work",at+.35+VAFTER));
                     /* 드롭이 걸린 마지막 세트 — '운동' 뒤에 '드롭세트'를 잇는다.
                        실패 뒤 쉬지 않고 무게를 내려 이어가라는 뜻이다 */
                     const seg=TM&&TM.st[idx];
                     if(seg&&seg.drop) say=tmSayAll(["드롭세트"],at+.35+VAFTER+tmDur("work")+VPAUSE); }
  else if(kind==="rest"){ const beep=.22+.15;   /* 두 번째 삐 소리가 끝나는 시각 */
                          n.push(tmTone(660,.15,.5,at)); n.push(tmTone(660,.15,.5,at+.22));
                          n.push(tmSay("rest",at+beep+VAFTER));
                          /* 다음이 새 종목이면 휴식 시작에 한 번 읊는다 —
                             쉬는 동안 옮겨 갈 기구와 무게를 미리 알라고 */
                          const off=beep+VAFTER+tmDur("rest")+.5;
                          say=tmNextSay(idx,at+off,off); }
  else if(kind==="prep"){ n.push(tmTone(784,.12,.4,at)); n.push(tmSay("prep",at+.12+VAFTER)); }
  /* 카운트다운은 두 단이다. 리드 시작에 낮고 긴 소리로 한 번 알리고(준비 시작),
     그 뒤로는 조용한 딸깍이 이어지다가, 마지막 3초만 높고 또렷해진다.
     전부 같은 소리로 울리면 몇 초 남았는지 귀로 구별이 안 돼 결국 화면을 보게 된다. */
  else if(kind==="lead")  n.push(tmTone(523,.22,.38,at));
  else if(kind==="tick")  n.push(tmTone(660,.05,.16,at));
  else if(kind==="tick3") n.push(tmTone(880,.07,.30,at));
  else { n.push(tmTone(1046,.5,.6,at)); n.push(tmSay("done",at+.5+VAFTER)); }
  tmPend.push({at:wall,kind:kind,nodes:n.filter(Boolean).concat(say),say:say});
}
/* 읊는 중인 안내를 끊는다. 건너뛰기로 다른 종목에 가 놓고 이전 안내를
   계속 듣고 있으면 틀린 무게를 듣게 된다. 삐 소리는 짧아서 그냥 둔다. */
function tmHush(){
  tmPend.forEach(x=>(x.say||[]).forEach(o=>{ try{ o.stop(0); }catch(e){} }));
}
/* 아직 안 난 예약을 모두 취소한다. 나고 있는 소리는 자르지 않는다 */
function tmDrop(){
  const now=Date.now(), keep=[];
  tmPend.forEach(x=>{
    if(x.at<=now+250){ if(x.at>now-8000) keep.push(x); return; }
    x.nodes.forEach(o=>{ try{ o.stop(0); }catch(e){} });
  });
  tmPend=keep;
}
/* 카운트다운을 몇 초 전부터 시작할지.

   운동으로 들어갈 때만 길게 센다 — 3초는 통보지 준비 시간이 아니어서
   바 밑으로 들어가려면 그 전에 알아야 한다. 반대로 운동 → 휴식 전환에는
   3초 그대로다. 세트를 치는 중에 딸깍거리면 도움이 아니라 방해다.

   리드는 끝나는 구간의 길이에 비례시킨다. 휴식이 길수록 준비할 것이 많은
   종목이기 때문이다 — 스쿼트 180초에는 15초가 맞고, 레터럴 60초에 15초를
   주면 휴식의 4분의 1이 카운트다운이 된다. */
function tmLead(seg,next){
  if(!next||next.p!=="work") return 3;
  const t=seg.t, lead = t>=150 ? 15 : t>=105 ? 10 : t>=60 ? 5 : 3;
  return Math.min(lead,Math.max(3,t-2));   /* 준비(5초)처럼 짧은 구간은 잘라 맞춘다 */
}
/* 지금부터 90초 안에 울릴 것을 전부 오디오 시계에 건다 */
function tmSchedule(){
  tmDrop(); tmSchedAt=Date.now();
  if(!TM||!TM.run||!tmAC) return;
  const now=Date.now(), lim=now+LOOKAHEAD;
  const add=(k,at,idx)=>{ if(at>now+250&&at<=lim) tmCueAt(k,at,idx); };
  let end=TM.endAt;
  for(let i=TM.i;i<TM.st.length;i++){
    if(i>TM.i) end+=TM.st[i].t*1000;
    const nx=TM.st[i+1], lead=tmLead(TM.st[i],nx);
    if(lead>3) add("lead",end-lead*1000);
    for(let k=lead-1;k>=4;k--) add("tick",end-k*1000);
    for(let k=3;k>=1;k--) add("tick3",end-k*1000);
    /* end에 시작되는 건 i+1번 구간이다 — 안내는 그 구간을 기준으로 만든다 */
    add(nx?nx.p:"done",end,i+1);
    if(end>lim) break;
  }
}
/* 방금 이 구간에 들어왔다 — 예약이 아니라 즉시 */
function tmEnter(){
  if(!TM) return;
  const p=TM.st[TM.i].p;
  tmHush();
  tmCueAt(p,Date.now()+20,TM.i);
  tmVib(p==="work"?[350]:p==="rest"?[140,90,140]:[90]);
  tmSchedule();
}
/* 잠금화면·알림창에 종목과 세트를 띄우고 거기서 일시정지할 수 있게 한다 */
function tmMedia(){
  if(!("mediaSession" in navigator)||!TM) return;
  const s=TM.st[TM.i];
  try{
    navigator.mediaSession.metadata=new MediaMetadata({
      title:s.e.name, artist:(s.wu?(s.p==="work"?"예열":"예열 뒤 휴식"):PH[s.p][0]+" · 세트 "+s.set+" / "+s.sets), album:TM.title});
    navigator.mediaSession.playbackState=TM.run?"playing":"paused";
    navigator.mediaSession.setActionHandler("play",tmPause);
    navigator.mediaSession.setActionHandler("pause",tmPause);
    navigator.mediaSession.setActionHandler("nexttrack",()=>tmGo(1));
    navigator.mediaSession.setActionHandler("previoustrack",()=>tmGo(-1));
  }catch(e){}
}
async function tmLock(){ try{ if("wakeLock" in navigator && !tmWL) tmWL=await navigator.wakeLock.request("screen"); }catch(e){} }
async function tmUnlock(){ try{ if(tmWL){ const w=tmWL; tmWL=null; await w.release(); } }catch(e){} }
document.addEventListener("visibilitychange",()=>{
  if(document.visibilityState!=="visible"||!TM) return;
  /* 돌아왔다 — 화면을 지금 시각에 맞추고 예약도 다시 건다 */
  if(tmAC&&tmAC.state==="suspended") tmAC.resume();
  if(TM.run){ tmLock(); tmLoop(); tmSchedule(); }
});

/* 종목 한 줄 요약. 어느 근육을 쓰는지 · 몇 kg · 몇 회 — 상단과 하단이
   같은 함수를 쓰므로 한쪽만 형식이 어긋날 일이 없다.
   가운데 무게는 눌러서 그 자리에서 고친다 — 봉에 실제로 끼운 게 다를 때
   타이머를 닫고 목록으로 돌아갔다 오지 않아도 되게. 저장되는 곳은
   목록에서 고칠 때와 같은 덮어쓰기(OV)이므로 값이 두 벌로 갈리지 않는다.
   '위 참조'(메인 리프트 본세트)는 종목이 아니라 시작중량을 고쳐야
   웜업 램프·원판까지 따라오므로 lift 쪽으로 보낸다. */
function tmKey(e){
  return ((e.w==="위 참조"||e.trk)&&e.lift) ? "lift|"+e.lift : "ex|"+mode+"|"+tab+"|"+e.name;
}
/* 타이머 안 기록 — 지금 종목에 맞는 것 하나. 메인 · 선형 추적은 달성/미달(증량 탭과 같은 저장소),
   케이블 · 레터럴은 미세 단계 ▲▼. 속성 이름을 바꿔 둔다(data-t…) — render()의 bindEdits가
   같은 이름이면 타이머 버튼에도 onclick을 달아 아래 위임 처리와 두 번 눌린다 */
function tmRec(e){
  if(e.mic&&!e.trk){ const key=mode+"|"+tab+"|"+e.name;
    return micHTML(e,key,OV.ex[key]?OV.ex[key].w:e.w).replace(/data-mic/g,"data-tmic"); }
  const k=e.lift;
  if(!k||!LIFT[k]||e.lo||e.pct||!(e.w==="위 참조"||e.trk)||dlOn()) return "";
  return (`<div class="tm-rec-h">이번 칸 결과 · ${week}주차 칸</div>`+logCtl(k,week)).replace(/data-log/g,"data-tlog");
}
function tmRecSet(html){
  const el=document.getElementById("tm-rec"), a=document.activeElement;
  if(a&&a.tagName==="INPUT"&&el.contains(a)) return;      /* 미달 회수 치는 중 */
  el.innerHTML=html;
}
function tmMeta(e){
  const key=tmKey(e), p=key.split("|");
  const ov = p[0]==="lift" ? OV.lifts[p[1]] : OV.ex[p.slice(1).join("|")];
  const w=`<span class="tm-w${ov?" edited":""}" data-tmw="${escAttr(key)}">${esc(tmW(e))}</span>`;
  return [esc(e.mg), w, e.r?esc(e.r)+(/초$/.test(e.r)?"":"회"):""].filter(Boolean).join(" · ");
}
function tmEditW(el){
  const p=el.dataset.tmw.split("|");
  /* 입력칸을 먼저 걷어내고 다시 그린다 — 그냥 tmPaint를 부르면
     아래 tmSet이 '아직 고치는 중'으로 보고 입력칸을 그대로 둔다 */
  const after=()=>{ el.innerHTML=""; tmPaint(); };
  if(p[0]==="lift") editLift(el, p[1], after);
  else editEx(el, p.slice(1).join("|"), after);
}
/* 입력칸이 열려 있는 줄은 건드리지 않는다 — 타이머가 120ms마다 다시 그리므로
   그냥 덮어쓰면 두 글자 치는 사이에 입력이 날아간다. */
function tmSet(id,html){
  const el=document.getElementById(id);
  if(el.querySelector("input")) return;
  el.innerHTML=html;
}
function tmPaint(){
  if(!TM) return;
  const s=TM.st[TM.i], nx=TM.st[TM.i+1], ph=PH[s.p];
  document.getElementById("tm-dial").style.setProperty("--phase",ph[1]);
  document.getElementById("tm-phase").textContent=(s.p==="work"?(s.wu?"예열":s.ramp?"램프":s.wt?"웜업":ph[0]):ph[0])+(s.drop?" · 드롭":"");
  document.getElementById("tm-ctx").textContent=TM.title+" · "+s.li+"/"+s.ln+" 종목";
  /* 그립 회전 종목은 이름 아래에 이번 사이클 그립을 크게 — 타이머에서 나가지 않고 본다 */
  const gNow = mode==="back" ? gripOf(s.e) : null;
  document.getElementById("tm-name").innerHTML=esc(gNow?exName(s.e):s.e.name)+(gNow?`<span class="tm-grip">그립 · ${esc(gNow.t)}</span>${gNow.cue?`<span class="tm-cue">▸ ${esc(gNow.cue)}</span>`:""}`:"");
  tmSet("tm-meta",tmMeta(s.e));
  document.getElementById("tm-set").textContent=(s.wu?(s.p==="work"?"예열 세트":"예열 뒤 휴식 · 다음은 1세트")+" · 본세트 "+s.sets
    :s.ramp?"램프 "+s.set+" / "+s.sets+" · 다음은 "+s.e.of.name:"세트 "+s.set+" / "+s.sets)+(s.drop?" · 드롭세트":"");
  /* 예열 — 이 종목 1세트 앞 준비 시간이거나 1세트 중이면 큰 띠로. 다음 종목 예고 줄에도 붙인다 */
  const wuNow = s.wu&&s.p==="work" || s.p==="prep"&&nx&&nx.wu, wuNext = nx&&nx.e!==s.e&&nx.e.wu&&nx.set===1;
  if(wuNow) document.getElementById("tm-set").insertAdjacentHTML("beforeend",
    `<br><span class="tm-wu">예열 — ${esc(s.e.wut||"50% × 6~8")}</span>`);
  /* 무릎 경고(2026-11-10) — 그 종목 세트 중 · 세트 사이 휴식 · 그 종목 앞 휴식(다음이 그 종목)이면 노랑 띠 */
  const knE = s.p==="work" ? s.e : (nx ? nx.e : null);
  if(knE&&knE.kn&&!s.wt&&!s.ramp) document.getElementById("tm-set").insertAdjacentHTML("beforeend",
    `<br><span class="tm-kn">⚠ ${esc(knE.kn)}</span>`);
  tmRecSet(tmRec(s.e));
  tmSlog(s);
  /* 수행 시간 — 운동 중이면 지금 세트, 준비 · 휴식 중이면 곧 시작할 세트의 종목 */
  const wt = s.p==="work" ? s.e : (nx&&nx.p==="work" ? nx.e : null);
  tmSet("tm-wk", wt&&wt.wk0 ? tmWkHTML(wt) : "");
  document.getElementById("tm-time").textContent=
    mmss(TM.run?Math.max(0,Math.ceil((TM.endAt-Date.now())/1000)):Math.ceil(TM.rest/1000));
  document.getElementById("tm-play").textContent=TM.run?"일시정지":"계속";
  /* 종목이 바뀐다 — 쉬는 동안 눈으로도 확인하게 상단과 같은 형식으로 적는다.
     읊어 주는 내용과 같다. 소리를 껐거나 못 들었을 때의 몫이기도 하다.
     여기 무게도 눌러 고칠 수 있다 — 다음 종목 준비는 쉬는 동안 하는 것이니 */
  tmSet("tm-up", (nx&&nx.e!==s.e)
    ? "다음 ▸ <b>"+esc(mode==="back"?exName(nx.e):nx.e.name)+"</b>"+((g=>g?' <b style="color:var(--bp)">· 그립 '+esc(g.t)+"</b>":"")(mode==="back"?gripOf(nx.e):null))+"<i>"+tmMeta(nx.e)+" · "+nx.sets+"세트"+(wuNext?'<br><b style="color:var(--sq)">예열 1세트 — '+esc(nx.e.wut||"50% × 6~8")+'</b>':"")+"</i>"
    : !nx ? "마지막 세트"
    : s.p==="prep" ? "이어서 ▸ 운동 "+mmss(nx.t)+" × "+nx.sets+"세트"
    : nx.p==="rest" ? "다음 ▸ 휴식 "+mmss(nx.t)
    : "다음 ▸ "+nx.set+"세트째"+(nx.drop?" · 드롭세트 — 내릴 무게 준비":""));
}
/* ±15초는 이번 세션에서 이 종목의 남은 세트(지금 세트 포함)에만 건다. 계속 쓰려면 '저장' —
   그때 폰 덮어쓰기(OVT)에 들어가고, 다음 세션부터 그 시간으로 시작한다 */
function tmWkHTML(e){
  const key=mode+"|"+tab+"|"+e.name, sv=TM.wk[e.name]!=null?TM.wk[e.name]:e.wk;
  const save = sv!==e.wk ? `<button class="tm-go" data-twksave="${escAttr(e.name)}">이 종목 ${sv}초로 저장</button>` : "";
  return `<div class="ex-go mic" style="justify-content:center"><button class="tm-go mic-dn" data-twk="-15">&minus;15초</button><span class="tm-rec-h">수행 ${sv}초${
    OVT[key]&&sv===e.wk?" · 폰 값":""}</span><button class="tm-go mic-dn" data-twk="15">+15초</button>${save}</div>`;
}
function tmWkAdj(d){
  if(!TM) return;
  const s=TM.st[TM.i], nx=TM.st[TM.i+1];
  const e = s.p==="work" ? s.e : (nx&&nx.p==="work" ? nx.e : null); if(!e) return;
  const cur=TM.wk[e.name]!=null?TM.wk[e.name]:e.wk, v=Math.max(5,cur+d), dd=v-cur; if(!dd) return;
  TM.wk[e.name]=v;
  for(let i=TM.i;i<TM.st.length;i++){ const x=TM.st[i]; if(x.p==="work"&&!x.wu&&x.e===e) x.t=Math.max(1,x.t+dd); }
  if(s.p==="work"&&!s.wu&&s.e===e){
    if(TM.run) TM.endAt=Math.max(Date.now()+1000, TM.endAt+dd*1000);
    else TM.rest=Math.max(1000, TM.rest+dd*1000);
  }
  tmLeft=-1; tmPaint(); if(TM.run) tmSchedule();
}
/* 나가 있는 동안(TM.min) 다른 탭을 봐도 타이머는 시작한 일차 · 칸으로 그린다 */
function tmLoop(){
  if(TM&&TM.min&&(tab!==TM.tab||week!==TM.week)){
    const t=tab, w=week; tab=TM.tab; week=TM.week;
    try{ tmLoop0(); } finally { tab=t; week=w; }
  } else tmLoop0();
  if(TM&&TM.min){ tmMiniPaint(); tmFollow(); }
}
/* 2026-11-13 사용자 요청 — 나가 있는 동안 타이머 종목이 바뀌면 그 카드로 따라간다.
   그 일차 화면을 보고 있을 때만 · 숫자 치는 중이면 안 움직인다(다음 틱에 다시 본다) */
function tmFollow(){
  const s=TM.st[TM.i], nm=s&&s.e?(s.e.of||s.e).name:null;
  if(!nm||nm===TM.focus||tab!==TM.tab) return;
  const a=document.activeElement; if(a&&a.tagName==="INPUT") return;
  TM.focus=nm; SOPEN.add(nm);
  const d=[...document.querySelectorAll("details.slog")].find(x=>x.dataset.sn===nm); if(d) d.open=true;
  tmGoCard(nm,true);
}
/* 세트 기록 한 줄 — 운동 중이면 지금 세트, 휴식이면 방금 끝낸 세트. 치는 중이면 안 건드린다 */
function tmSlog(s){
  const el=document.getElementById("tm-slog");
  const on = TM.date&&s.e&&s.p!=="prep"&&!s.ramp&&!s.wu&&!s.wt&&!s.e.of;
  const key = on ? s.e.name+"|"+s.set+"|"+JSON.stringify(sGet(TM.date,s.e.name)[s.set-1]||{}) : "";
  const H = on ? sHist(TM.tab,s.e.name,TM.date,1)[0] : null, hx = H ? H.sets[s.set-1] : null;
  const prev = on ? `<div class="sl-prev tm">${H?`지난번 ${s.set}세트: <b>${esc(sCellTxt(hx))}</b> <span>(${+H.date.slice(5,7)}/${+H.date.slice(8)})</span>`:"지난번 기록 없음"}</div>` : "";
  if(el.dataset.k===key) return;
  const a=document.activeElement; if(a&&el.contains(a)&&a.tagName==="INPUT") return;
  el.dataset.k=key;
  el.innerHTML = on ? `<div class="tm-rec-h">세트 기록 · ${s.set}/${s.sets}${s.p==="rest"?" · 방금 세트":""}</div>`+prev+
    sBox(s.e,TM.date,TM.tab,tmW(s.e),sRowsHTML(s.e,s.set-1,s.set,TM.date,tmW(s.e),"tsl")) : "";
}
function tmMiniPaint(){
  const b=document.getElementById("tm-mini");
  if(!TM||!TM.min){ b.hidden=true; return; }
  const s=TM.st[TM.i], left=TM.run?Math.max(0,Math.ceil((TM.endAt-Date.now())/1000)):Math.ceil(TM.rest/1000);
  const t=`⏱ ${mmss(left)} · ${TM.run?PH[s.p][0]:"일시정지"} · ${s.e.name} ${s.set}/${s.sets} — 눌러서 타이머로`;
  if(b.textContent!==t) b.textContent=t;
  /* 2026-11-13 사용자 요청 — 다이얼과 같은 색: 운동 초록 · 휴식 빨강 · 준비 노랑 · 일시정지는 원래 흰색 */
  const bg = TM.run ? PH[s.p][1] : "";
  if(b.style.background!==bg){ b.style.background=bg; b.style.color = !TM.run ? "" : s.p==="prep" ? "var(--bg)" : "#fff"; }
}
function tmMin(){
  if(!TM) return;
  const s=TM.st[TM.i], nm=s&&s.e?(s.e.of||s.e).name:null;
  TM.min=true; document.getElementById("tm").hidden=true; document.body.style.overflow="";
  tab=TM.tab; week=TM.week; if(TM.date) LOGDATE={day:TM.tab,date:TM.date};
  if(nm) SOPEN.add(nm);
  TM.focus=nm;
  render(); tmMiniPaint(); document.getElementById("tm-mini").hidden=false;
  tmGoCard(nm);
}
/* 2026-11-13 사용자 요청 — 나가기 하면 지금 종목 카드의 맨 위가 고정 머리글 바로 아래에 오게. 잠깐 테두리 반짝 */
function tmGoCard(nm,soft){
  if(!nm) return;
  const d=[...document.querySelectorAll("details.slog")].find(x=>x.dataset.sn===nm);
  const card=d?d.closest("article.ex"):null; if(!card) return;
  /* soft(자동 따라가기) — 카드가 이미 화면에 다 보이면 스크롤 없이 반짝이기만 */
  const h=document.querySelector("header"), top0=h?h.getBoundingClientRect().height:0, r=card.getBoundingClientRect();
  if(soft&&r.top>=top0&&r.bottom<=window.innerHeight-70){ card.classList.remove("flash"); void card.offsetWidth; card.classList.add("flash"); return; }
  const go=()=>{ const h=document.querySelector("header"), top=h?h.getBoundingClientRect().height:0;
    window.scrollTo({top:Math.max(0,card.getBoundingClientRect().top+window.scrollY-top-10)}); };
  go(); requestAnimationFrame(go);
  card.classList.remove("flash"); void card.offsetWidth; card.classList.add("flash");
}
function tmMax(){
  if(!TM) return;
  TM.min=false; document.getElementById("tm-mini").hidden=true;
  if(tab!==TM.tab||week!==TM.week){ tab=TM.tab; week=TM.week; render(); }
  document.getElementById("tm").hidden=false; document.body.style.overflow="hidden";
  document.getElementById("tm-slog").dataset.k=""; tmLeft=-1; tmPaint();
}
function tmLoop0(){
  if(!TM||!TM.run) return;
  const now=Date.now(); let moved=false;
  while(TM.endAt<=now&&TM.i<TM.st.length-1){ TM.i++; TM.endAt+=TM.st[TM.i].t*1000; moved=true; }
  if(TM.endAt<=now){ tmVib([300,120,300,120,300]); tmClose(); return; }
  if(moved){
    tmPaint(); tmMedia(); tmLeft=-1;
    tmVib(TM.st[TM.i].p==="work"?[350]:TM.st[TM.i].p==="rest"?[140,90,140]:[90]);
  }
  /* 예약은 5초마다 다시 건다. 구간 전환 때만 걸면 휴식 150초처럼
     예약 창(90초)보다 긴 구간의 알림이 통째로 빠진다. 다시 걸면서
     오디오 시계와 벽시계 사이의 드리프트도 같이 잡힌다. */
  if(moved||now-tmSchedAt>5000) tmSchedule();
  const left=Math.max(0,Math.ceil((TM.endAt-now)/1000));
  if(left!==tmLeft){
    tmLeft=left;
    document.getElementById("tm-time").textContent=mmss(left);
  }
}
function tmStart(d,from){
  tmRun(tmSteps(d,week,from), d.t);
}
/* ── 웜업 타이머 (2026-11-07) ─────────────────────────────────────
   세션 시작 전 웜업을 세션 타이머와 같은 엔진으로 따로 돌린다. 항목은 빌드의 WU_SESSION.
   램프 무게는 화면 카드와 같은 계산(메인 = ramp() · 핵 · 8일차 바벨 = 그 종목 무게 × 비율).
   끝의 휴식은 남긴다 — 마지막이 램프면 그게 본세트 전 휴식이다. */
const dayKey = d => Object.keys(D.back||{}).find(k=>D.back[k]===d);
function wuItems(d){
  const P=mode==="back"&&D.wuPlan?D.wuPlan[dayKey(d)]:null; if(!P) return [];
  const RR=[60,60,90,120], out=[], cyc=altRound();
  P.forEach(x=>{
    if(x.cf&&(cyc==null||cyc<x.cf)) return;          /* 2026-11-10 그 사이클부터 나오는 웜업 항목 */
    if(x.ramp){
      const m=(d.mains||[])[0]; if(!m||!LIFT[m[0]]) return;
      ramp(workWeight(m[0],week)).filter(r=>r.w).forEach((r,i)=>out.push(
        {name:m[1]+" 램프 "+(i+1),mg:"웜업 램프",w:String(r.w),r:String(r.reps),s:r.sets,wk:r.reps*3+10,rt:RR[i]||90}));
    } else if(x.pctOf){
      const e=allEx(d).find(y=>x.pctOf==="#2"?y.n===2:y.name===x.pctOf); if(!e) return;
      pctRamp(x,e).forEach((r,i)=>out.push({name:exName(e)+" 램프 "+(i+1),mg:"웜업 램프",
        w:r.bar?"빈 봉 20":String(r.w),r:String(r.reps),s:1,wk:r.reps*3+10,rt:r.rt}));
    } else if(x.ss){
      const g="ss"+out.length;
      x.ss.forEach(y=>{ const dd=y.ofDay?D.back[y.ofDay]:d, e=dd&&!dd.rest?allEx(dd).find(z=>z.name===y.of):null;
        /* 다른 일차 무게 칸이면 그 일차의 폰 덮어쓰기를 본다(tmW는 지금 탭 기준) */
        const ov=e&&y.ofDay?OV.ex[mode+"|"+y.ofDay+"|"+e.name]:null;
        const b=e?parseFloat(y.ofDay?(ov?ov.w:e.w):tmW(e)):NaN, w=y.w?y.w:(b>0?String(Math.max(1,R(b*y.pct))):"가볍게");  /* y.w = 고정 무게(2026-11-12 머신 체스트 프레스 예열) */
        out.push({name:y.n,mg:y.mg,w:w,r:y.r,s:x.s,wk:y.wk,rt:y.rt,ssg:g}); });
    } else out.push({name:x.n,mg:x.mg||"웜업",w:x.w,r:x.r,s:x.s,wk:x.wk,rt:x.rt});
  });
  return out;
}
function wuSteps(d){
  const it=wuItems(d), st=[], G=[];
  it.forEach(e=>{ const g=G[G.length-1]; if(g&&e.ssg&&g[0].ssg===e.ssg) g.push(e); else G.push([e]); });
  /* 슈퍼세트 묶음은 라운드마다 번갈아 — 세션 타이머와 같은 걸음 */
  G.forEach(g=>{ for(let k=1;k<=g[0].s;k++) g.forEach(e=>{
    st.push({p:"work",t:e.wk,e:e,set:k,sets:e.s,wt:true});
    if(e.rt>0) st.push({p:"rest",t:e.rt,e:e,set:k,sets:e.s,wt:true}); }); });
  if(st.length) st.unshift({p:"prep",t:PREP,e:st[0].e,set:1,sets:st[0].sets});
  st.forEach(x=>{ x.li=it.indexOf(x.e)+1; x.ln=it.length; });
  return st;
}
function tmRun(st,title){
  if(!st.length) return;
  tmAudio();
  TM={st:st,i:0,endAt:Date.now()+st[0].t*1000,rest:0,run:true,title:title,wk:{},
      tab:tab, week:week, date:(mode==="back"&&D.back[tab]&&!D.back[tab].rest)?logDateFor(tab):null, min:false};
  document.getElementById("tm-mini").hidden=true;
  if(TM.date) sWarm(TM.tab,TM.date);
  document.getElementById("tm").hidden=false;
  document.body.style.overflow="hidden";
  tmLock(); tmPaint(); tmMedia(); tmEnter(); tmLeft=-1;
  clearInterval(tmInt); tmInt=setInterval(tmLoop,120);
}
function tmGo(delta){
  if(!TM) return;
  const j=TM.i+delta;
  if(j<0||j>=TM.st.length) return;
  TM.i=j; TM.endAt=Date.now()+TM.st[j].t*1000; TM.rest=TM.st[j].t*1000;
  tmLeft=-1; tmPaint(); tmMedia();
  if(TM.run) tmEnter(); else tmDrop();
}
function tmPause(){
  if(!TM) return;
  if(TM.run){ TM.rest=Math.max(0,TM.endAt-Date.now()); TM.run=false; tmUnlock();
              tmHush(); tmDrop(); }
  else { TM.endAt=Date.now()+TM.rest; TM.run=true; tmLock(); tmSchedule(); }
  tmLeft=-1; tmPaint(); tmMedia();
}
function tmClose(){
  clearInterval(tmInt); tmInt=null; TM=null; tmLeft=-1; tmUnlock(); tmHush(); tmDrop();
  document.getElementById("tm-mini").hidden=true;
  /* 유지음은 끊는다 — 타이머를 닫고도 계속 흘리면 배터리만 먹는다.
     다만 '완료' 말소리가 남아 있을 수 있어 2.5초 뒤에 끈다. */
  clearTimeout(tmIdle);
  tmIdle=setTimeout(()=>{ try{
    if(tmKeep){ tmKeep.stop(); tmKeep=null; }
    if(tmAC) tmAC.suspend();
  }catch(e){} }, 2500);
  try{ if("mediaSession" in navigator){ navigator.mediaSession.metadata=null;
    navigator.mediaSession.playbackState="none"; } }catch(e){}
  document.getElementById("tm").hidden=true;
  document.body.style.overflow="";
}

function bindEdits(){
  document.querySelectorAll("[data-lift]").forEach(el=>{ el.onclick=()=>editLift(el); });
  document.querySelectorAll("[data-exkey]").forEach(el=>{ el.onclick=()=>editEx(el); });
  document.querySelectorAll("[data-mic]").forEach(el=>{ el.onclick=()=>micGo(el.dataset.mic,1); });
  document.querySelectorAll("[data-micdn]").forEach(el=>{ el.onclick=()=>micGo(el.dataset.micdn,-1); });
  const rb=document.getElementById("ov-reset");
  if(rb) rb.onclick=()=>{ OV={lifts:{},ex:{}}; saveOV();
    for(const k in OVT){ const e=exByKey(k); if(e) e.wk=e.wk0; } OVT={}; saveOVT(); render(); };
  /* 카드의 '수행 N초' — 숫자로 고친다. 비우면 엑셀 값으로 */
  document.querySelectorAll("[data-wkkey]").forEach(el=>{ el.onclick=()=>{
    const k=el.dataset.wkkey, e=exByKey(k); if(!e) return;
    grabInput(el, String(e.wk), true, v=>{ const n=parseInt(v,10);
      if(v==="") setWk(k,e.wk0); else if(n>0) setWk(k,n); }); }; });
  const cp=document.getElementById("cal-prev"), cn=document.getElementById("cal-next"),
        ct=document.getElementById("cal-today");
  const shift=n=>{ const now=new Date(), b=calYM||{y:now.getFullYear(),m:now.getMonth()};
    const d=new Date(b.y,b.m+n,1); calYM={y:d.getFullYear(),m:d.getMonth()}; render(); };
  if(cp) cp.onclick=()=>shift(-1);
  if(cn) cn.onclick=()=>shift(1);
  if(ct) ct.onclick=()=>{ calYM=null; render(); };
  /* 복귀 주차 보정 — 0이면 엑셀 기준으로 되돌린다 */
  document.querySelectorAll("[data-bshift]").forEach(el=>{ el.onclick=()=>{
    const v=Number(el.dataset.bshift);
    BSHIFT = v===0 ? 0 : BSHIFT+v;
    if(BSHIFT<0 && !backWeekOf(ymd(new Date()))) BSHIFT-=v;   /* 기준일 앞으로는 안 간다 */
    bSave(); syncBackWeek(); render(); }; });
  /* 주별 세션 체크 — 다시 누르면 해제 */
  document.querySelectorAll("[data-bsess]").forEach(el=>{ el.onclick=()=>{
    const f=el.dataset.bsess, s=el.dataset.bslot;
    const w=BSESS[f]||(BSESS[f]={});
    if(w[s]) delete w[s]; else w[s]=1;
    if(!Object.keys(w).length) delete BSESS[f];
    bSave(); render(); }; });
  document.querySelectorAll("[data-csetday]").forEach(el=>{ el.onclick=()=>{
    const v=el.dataset.csetday;
    if(v==="__xl"){ CANCHOR=D.cycAnchor&&D.cycAnchor.date; try{ localStorage.removeItem(CAKEY); }catch(e){} }
    else setCycToday(v);
    render(); }; });
  document.querySelectorAll("[data-hcal]").forEach(el=>{ el.onclick=()=>{
    tab=el.dataset.hcal; render(); window.scrollTo({top:0}); }; });
  document.querySelectorAll("[data-bcal]").forEach(el=>{ el.onclick=()=>{
    LOGDATE = el.dataset.bdate ? {day:el.dataset.bcal, date:el.dataset.bdate} : null;
    tab=el.dataset.bcal; render(); window.scrollTo({top:0}); }; });
  const su=document.getElementById("sl-url-save");
  if(su) su.onclick=()=>{ const v=document.getElementById("sl-url").value.trim();
    if(!/^https:\/\/script\.google(usercontent)?\.com\//.test(v)){ su.textContent="주소 확인"; return; }
    SYNC_URL=v; try{ localStorage.setItem(SYU,v); }catch(e){}
    SDIRTY=Object.keys(SETS); syncSaveState(); sSync(); render(); };
  const sd=document.getElementById("sl-url-del");
  if(sd) sd.onclick=()=>{ SYNC_URL=""; try{ localStorage.removeItem(SYU); }catch(e){} render(); };
  const slc=document.getElementById("sl-copy");
  if(slc) slc.onclick=()=>{
    const text=sExport(), n=text.split("\n").length-2;
    const ok=()=>{ slc.textContent=`복사됨 · ${n}세트`; setTimeout(()=>{ slc.textContent="PC로 보낼 세트 기록 복사"; },1800); };
    try{ navigator.clipboard.writeText(text).then(ok,()=>{ slc.textContent="복사 실패"; }); }catch(e){ slc.textContent="복사 실패"; } };
  document.querySelectorAll("[data-setday]").forEach(el=>{ el.onclick=()=>{
    const v=el.dataset.setday;
    if(v==="__xl"){ ANCHOR=D.anchor; try{ localStorage.removeItem(AKEY); }catch(e){} }
    else setToday(v);
    render(); }; });
  document.querySelectorAll("[data-cal]").forEach(el=>{ el.onclick=()=>{
    const c=Number(el.dataset.cyc);
    if(c>=WBASE&&c<=WLAST) week=c;
    tab=el.dataset.cal; render(); window.scrollTo({top:0}); }; });
  /* data-tm="" = 세션 처음부터 · data-tm="종목명" = 그 종목부터 */
  document.querySelectorAll("[data-tm]").forEach(el=>{ el.onclick=()=>{
    const d=(mode==="home"?D.home:D.back)[tab];
    if(d&&!d.rest) tmStart(d,el.dataset.tm||null); }; });
  document.querySelectorAll("[data-wtm]").forEach(el=>{ el.onclick=()=>{
    const d=D.back[tab]; if(mode==="back"&&d&&!d.rest) tmRun(wuSteps(d),"웜업 · "+d.t); }; });
  document.querySelectorAll("[data-prep-tm]").forEach(el=>{ el.onclick=()=>{
    const t=prepToday(), k=el.dataset.prepTm, x=k==="today"?t.d:t[k];
    if(x&&x.ex.length) tmStart(x,null); }; });
  document.querySelectorAll("[data-prepf]").forEach(el=>{
    el.onchange=()=>{ setPLog(el.dataset.prepw,el.dataset.prepf,el.value); render(); };
    el.onkeydown=ev=>{ if(ev.key==="Enter") el.blur(); }; });
  const pc=document.getElementById("prep-copy");
  if(pc) pc.onclick=()=>{
    const text=Object.keys(PLOG).sort((a,b)=>a-b).map(w=>`${w}회차\t${pVal(w,"w")}\t${pVal(w,"m")}`).join("\n");
    const ok=()=>{ pc.textContent="복사됨"; setTimeout(()=>{ pc.textContent="엑셀에 옮길 값 복사 (복귀_캐리 E·F열)"; },1500); };
    try{ navigator.clipboard.writeText(text).then(ok,()=>{ pc.textContent="복사 실패"; }); }catch(e){ pc.textContent="복사 실패"; } };
  /* 증량 기록 — 달성/미달 토글, 미달 반복수, 피로 0~5, 관절·메모 */
  document.querySelectorAll("[data-logv]:not([data-logf])").forEach(b=>{ b.onclick=()=>{
    const seg=b.closest(".seg"), k=seg.dataset.logk, c=seg.dataset.logc, v=b.dataset.logv;
    const cur=logVal(c,k), same = v==="O" ? (cur!==""&&isDone(cur)) : (cur!==""&&!isDone(cur));
    setLog(c,k, same?"":v); render();
    if(!same&&v==="X"){ const i=document.querySelector(`.seg[data-logk="${k}"][data-logc="${c}"] input`); if(i) i.focus(); }
  }; });
  document.querySelectorAll("[data-logrep]").forEach(i=>{
    i.onchange=()=>{ const seg=i.closest(".seg"), v=i.value.replace(/\D/g,"");
      setLog(seg.dataset.logc, seg.dataset.logk, v||"X"); render(); };
    i.onkeydown=ev=>{ if(ev.key==="Enter") i.blur(); };
  });
  document.querySelectorAll("[data-logf]").forEach(el=>{
    const c=el.dataset.logc, f=el.dataset.logf;
    if(el.tagName==="BUTTON") el.onclick=()=>{ const v=el.dataset.logv; setLog(c,f, logVal(c,f)===v?"":v); render(); };
    else { el.onchange=()=>{ setLog(c,f,el.value); render(); };
           el.onkeydown=ev=>{ if(ev.key==="Enter") el.blur(); }; }
  });
  document.querySelectorAll("[data-logrow]").forEach(tr=>{ tr.onclick=ev=>{
    if(ev.target.closest("input,button")) return;
    week=Number(tr.dataset.logrow); render(); }; });
  const lc=document.getElementById("log-copy"); if(lc) lc.onclick=()=>logCopy(lc);
}

let mode="back", week=clampW(D.config.week), tab=D.backOrder[0];
const tabsFor = m => (m==="home"
  ? D.homeOrder.filter(d=>!(D.home[d]&&D.home[d].rest)).map(d=>[d,d]).concat([["cal","달력"],["plan","주간"],["prog","증량"],["check","검산"]])
  : D.backOrder.filter(d=>!(D.back[d]&&D.back[d].rest)).map(d=>[d,d+"일"]).concat([["time","시간"]],D.prep?[["prep","배송"]]:[],[["bcal","달력"],["bplan","사이클"],["prog","증량"],["check","검산"]]));

function reAlt(){
  if(applyAlt()) altEach((e,d)=>{ e.wk0=e.wk; const o=OVT["back|"+d+"|"+e.name]; if(o&&o.base===e.wk0) e.wk=o.wk; });
  applyOrd();
}
/* ── 시간표 (2026-11-05 사용자 요청) ── 폰에서 고친 수행 초 · 그립 회전 원암 · 사이클 교대 · 디로드 세트까지
   얹은 채로 세션 타이머와 같은 계산(tmSteps — 준비 5초 + 세트 · 휴식, 마지막 휴식 뺌)을 일차마다 돌린다.
   다른 사이클은 VIEW_R로 잠깐 그 사이클을 얹어 계산하고 바로 오늘로 되돌린다 */
const TTK="wk-tt";
let TT={wu:30,sh:20,sw:0}, TT_R=null;
try{ Object.assign(TT,JSON.parse(localStorage.getItem(TTK)||"{}")); }catch(e){}
function ttCycle(r){
  const start=CANCHOR?dayAdd(CANCHOR,(r-1)*CLEN):null, bw=start?backWeekOf(start):null, dl=!!(bw&&bw.deload);
  VIEW_R=r; reAlt();
  const S=e=>dl?dlSets(e):e.s;
  const days=D.backOrder.map((d,i)=>{ const day=D.back[d], date=start?dayAdd(start,i):null;
    if(!day||day.rest) return {d,date,rest:true,t:(day&&day.t)||"휴식"};
    const st=tmSteps(day,week,null,S), per=new Map();
    st.forEach(x=>{ const k=x.e; if(!per.has(k)) per.set(k,{name:exName(k),n:k.n,sets:S(k),secs:0,arm:(gripOf(k)||{}).arm||""}); per.get(k).secs+=x.t; });
    const ex=[...per.values()], secs=st.reduce((a,x)=>a+x.t,0)+Math.max(0,ex.length-1)*TT.sw;
    return {d,date,t:day.t,sets:allEx(day).reduce((a,e)=>a+(e.s>0?S(e):0),0),secs,ex}; });
  VIEW_R=null; reAlt();
  return {r,start,dl,days};
}
const hm = m => { m=Math.round(m); const h=Math.floor(m/60), x=m%60; return h?h+"시간"+(x?" "+x+"분":""):x+"분"; };
function timeHTML(){
  const t0=ymd(new Date()), c=cycDayOf(t0), cur=c?c.round:1;
  const r=Math.max(1,TT_R||cur), C=ttCycle(r), md=t=>t?`${+t.slice(5,7)}/${+t.slice(8)}(${DOWK(t)})`:"";
  const fix=(TT.wu|0)+(TT.sh|0), live=C.days.filter(x=>!x.rest);
  const wm = x => Math.round(x.secs/60);
  let sumW=0;
  const rows=C.days.map(x=>{
    if(x.rest) return `<tr class="rest"><td>${x.d}</td><td class="wrap">${md(x.date)} · 휴식</td><td class="num">—</td><td class="num">—</td><td class="num">—</td></tr>`;
    const w=wm(x), xl=CPLAN[x.d]?+CPLAN[x.d].m||0:0, dm=w-xl, today=r===cur&&c&&c.day===x.d;
    sumW+=w;
    return `<tr style="${today?'background:var(--raised)':''}"><td>${x.d}${today?'<br><span class="judge j-lo">오늘</span>':''}</td>
      <td class="wrap"><details><summary style="font-weight:700;color:var(--ink)">${md(x.date)} · ${esc(x.t)} <span style="font-weight:500;color:var(--mute)">· ${x.sets}세트</span></summary>
      <div style="font-size:12.5px;line-height:1.7;margin-top:4px">${x.ex.map(e=>`${e.n}. ${esc(e.name)}${e.arm==="원암"?' <b style="color:var(--bp)">원암</b>':''} · ${e.sets}세트 · <b>${Math.round(e.secs/60*10)/10}분</b>`).join("<br>")}
      ${TT.sw?`<br><span style="color:var(--mute)">종목 사이 전환 ${TT.sw}초 × ${Math.max(0,x.ex.length-1)}</span>`:""}
      ${xl&&Math.abs(dm)>1?`<br><span style="color:var(--mute)">엑셀 계획 운동 ${xl}분보다 ${dm>0?dm+"분 길다":(-dm)+"분 짧다"}</span>`:""}</div></details></td>
      <td class="num"><b>${hm(w)}</b></td><td class="num">${hm(fix)}</td>
      <td class="num" style="font-weight:800;color:var(--ink)">${hm(w+fix)}</td></tr>`; });
  const td=r===cur&&c?C.days.find(x=>x.d===c.day&&!x.rest):null;
  let endLine="";
  if(td){ const w=wm(td), now=new Date(), end=new Date(now.getTime()+(w+fix)*60000);
    endLine=`<div class="rampbar"><b>오늘 ${td.d}일차 · 총 ${hm(w+fix)}</b><span>운동 ${hm(w)} · 웜업 + 샤워 ${hm(fix)} (스트레칭 · 웜업 ${TT.wu|0}분 + 샤워 ${TT.sh|0}분)<br>지금 시작하면 ${end.getHours()}시 ${String(end.getMinutes()).padStart(2,"0")}분에 끝</span></div>`; }
  const inp=(k,l,u)=>`<label style="display:inline-flex;align-items:center;gap:4px;margin:0 10px 6px 0;font-size:13px">${l}<input data-tt="${k}" type="number" inputmode="numeric" min="0" value="${TT[k]|0}" style="width:56px;padding:4px 6px;font:inherit;border:1px solid var(--line);border-radius:8px;background:var(--bg);color:var(--ink)">${u}</label>`;
  return `<h2 class="daytitle">시간표</h2>
  <div style="display:flex;align-items:center;justify-content:space-between;gap:8px;margin:6px 0 10px">
    <button class="tm-go" data-ttr="${r-1}" ${r<=1?"disabled":""} style="min-width:44px">◀</button>
    <div style="text-align:center"><b style="font-size:18px">${r}사이클${C.dl?' · 디로드':''}</b><br><span style="font-size:12.5px;color:var(--mute)">${md(C.start)} ~ ${md(C.start?dayAdd(C.start,CLEN-1):null)}${r===cur?" · 이번 사이클":r===cur+1?" · 다음 사이클":""}</span></div>
    <button class="tm-go" data-ttr="${r+1}" style="min-width:44px">▶</button></div>
  ${r!==cur?`<p style="text-align:center;margin:0 0 10px"><button class="tm-go" data-ttr="${cur}">이번 사이클로</button></p>`:""}
  ${endLine}
  <table class="tbl"><thead><tr><th>일차</th><th style="text-align:left">날짜 · 구성 (눌러서 종목별)</th><th>운동</th><th>웜업<br>+ 샤워</th><th>총</th></tr></thead><tbody>
  ${rows.join("")}
  <tr><td style="font-weight:800">합계</td><td class="wrap" style="color:var(--mute)">${live.length}일 · ${live.reduce((a,x)=>a+x.sets,0)}세트</td><td class="num" style="font-weight:800">${hm(sumW)}</td><td class="num" style="font-weight:800">${hm(fix*live.length)}</td><td class="num" style="font-weight:800">${hm(sumW+fix*live.length)}</td></tr>
  </tbody></table>
  <div style="margin:12px 0 4px">${inp("wu","스트레칭 · 웜업","분")}${inp("sh","샤워","분")}${inp("sw","종목 사이 전환","초")}</div>
  <p class="footnote">운동 = 세션 타이머와 같은 계산(폰에서 고친 수행 초 · 그립 회전 원암 120초 · 케이블 로우 양손 40초 · 8일차 교대 종목 · 디로드 절반 세트 · 준비 5초 · 마지막 휴식 뺌 · 전환 초). 예열 1세트는 운동에 들어간다(예열 ≤60초 + 휴식 60초). 5일차 프레스 램프는 세션 타이머 안이라 운동에 들어간다. 웜업 + 샤워 = 위 칸 두 개(웜업 타이머 · 메인 리프트 램프 포함으로 셈). 총 = 운동 + 웜업 + 샤워. 기구 대기는 안 들어간다.</p>`;
}function ttBind(){
  document.querySelectorAll("[data-ttr]").forEach(b=>b.onclick=()=>{ TT_R=Math.max(1,+b.dataset.ttr); render(); });
  document.querySelectorAll("[data-tt]").forEach(i=>i.onchange=()=>{ TT[i.dataset.tt]=Math.max(0,parseInt(i.value,10)||0);
    try{ localStorage.setItem(TTK,JSON.stringify(TT)); }catch(e){} render(); });
}
function render(){
  /* 켜 둔 채 사이클이 넘어갔거나 일차 기준을 되돌렸으면 교대 종목을 다시 얹는다 —
     수행 초 덮어쓰기는 불러올 때와 같은 규칙으로 새 이름 쪽에 다시 붙인다 */
  reAlt();
  const nav=document.getElementById("nav");
  nav.innerHTML=tabsFor(mode).map(([k,l])=>`<button data-k="${k}" aria-pressed="${k===tab}">${l}</button>`).join("");
  nav.querySelectorAll("button").forEach(b=>b.onclick=()=>{tab=b.dataset.k;LOGDATE=null;render();window.scrollTo({top:0});});
  /* 사이클 바는 양쪽 모드에 다 띄운다 — 중량 사다리는 재택 · 복귀가 같은 칸을 쓴다.
     다른 것은 올리는 속도뿐이다(재택 10일 주기마다 · 복귀 매주 한 칸). */
  document.getElementById("weekbar").style.display = "flex";
  document.getElementById("wk-track").style.display = "flex";
  const wkn=document.getElementById("wk-num"), dlh=dlOn();
  /* 2026-11-01 사용자 요청 — 사이클 번호와 중량 칸을 둘 다. '14'만 보이면 14사이클로 읽힌다.
     사이클 = 지금 보고 있는 중량 칸을 쓰는 사이클(−/+로 칸을 옮기면 같이 따라간다) · 디로드 중이면 오늘 사이클 */
  const rnd = dlh ? (bToday()||{}).round : roundOfWeek(week);
  wkn.textContent=(rnd?rnd+"사이클 · ":"")+(dlh?"디로드":week+"칸");
  wkn.style.fontSize=dlh||rnd?"19px":"";
  document.getElementById("wk-of").textContent=`중량 칸 / ${WLAST}`
    + (dlh?" · 세트 절반":" · 사이클마다 한 칸");
  document.getElementById("wk-track").innerHTML=
    Array.from({length:WEEKS},(_,i)=>`<span class="${i<week-WBASE+1?'on':''}"></span>`).join("");
  document.getElementById("wk-down").disabled=week<=WBASE;
  document.getElementById("wk-up").disabled=week>=WLAST;
  const set = mode==="home"?D.home:D.back;
  let html;
  if(tab==="cal") html=calHTML();
  else if(tab==="check") html=checkHTML();
  else if(tab==="plan") html=planHTML();
  else if(tab==="prog") html=progressHTML(week);
  else if(tab==="bcal") html=backCalHTML();
  else if(tab==="bplan") html=backPlanHTML();
  else if(tab==="prep") html=prepHTML();
  else if(tab==="time") html=timeHTML();
  else html = set[tab] ? dayHTML(set[tab],week) : planHTML();
  document.getElementById("app").innerHTML = html + ovHTML() +
    `<p class="footnote">원본: ${esc(D.source)} · 중량은 시트와 동일한 계산식(1kg 격자 · 못 끼우는 값은 +1kg, 웜업 40/60/75/90%)으로 산출됩니다. 중량 숫자를 탭하면 이 폰에서만 고칠 수 있습니다.</p>`;
  bindEdits();
  if(tab==="time") ttBind();
  try{ const t=ymd(new Date()), c=mode==="back"?cycDayOf(t):{day:DOWK(t)};
    localStorage.setItem("wk-prog", JSON.stringify(
      {mode,week,tab,on:t,cd:c?c.day:null})); }catch(e){}
}
const tmMuteBtn=document.getElementById("tm-mute");
const tmMutePaint=()=>{ tmMuteBtn.textContent=voiceOn?"\uD83D\uDD0A":"\uD83D\uDD07";
  tmMuteBtn.classList.toggle("off",!voiceOn); };
tmMutePaint();
try{ tmAudio(true); }catch(e){}     /* 말소리 미리 디코딩 — 출력은 없다 */
tmMuteBtn.onclick=()=>{
  voiceOn=!voiceOn;
  try{ localStorage.setItem("wk-voice",voiceOn?"1":"0"); }catch(e){}
  tmMutePaint();
  if(!voiceOn) tmHush();            /* 끄면 하던 말도 그 자리에서 멈춘다 */
  if(TM&&TM.run) tmSchedule();      /* 이미 걸어둔 예약을 새 설정으로 다시 건다 */
};
/* 두 줄 다 tmPaint가 매번 새로 그리므로 개별 onclick은 붙자마자 사라진다 */
document.getElementById("tm").addEventListener("click",ev=>{
  const el=ev.target.closest("[data-tmw]");
  if(el) return tmEditW(el);
  const wb=ev.target.closest("[data-twk]");
  if(wb) return tmWkAdj(+wb.dataset.twk);
  const wsv=ev.target.closest("[data-twksave]");
  if(wsv&&TM){ const n=wsv.dataset.twksave;
    setWk(mode+"|"+tab+"|"+n, TM.wk[n]); delete TM.wk[n]; render(); return tmPaint(); }
  const mb=ev.target.closest("[data-tmic],[data-tmicdn]");
  if(mb){ const up=mb.dataset.tmic!=null;
    return micGo(up?mb.dataset.tmic:mb.dataset.tmicdn, up?1:-1, document.querySelector("#tm-meta [data-tmw]")); }
  const lb=ev.target.closest("[data-tlogv]");
  if(lb){ const seg=lb.closest(".seg"), k=seg.dataset.tlogk, c=seg.dataset.tlogc, v=lb.dataset.tlogv;
    const cur=logVal(c,k), same = v==="O" ? (cur!==""&&isDone(cur)) : (cur!==""&&!isDone(cur));
    setLog(c,k, same?"":v); render(); tmPaint();
    if(!same&&v==="X"){ const i=document.querySelector("#tm-rec input"); if(i) i.focus(); } }
});
document.getElementById("tm").addEventListener("change",ev=>{
  const i=ev.target; if(!i.dataset||!i.dataset.tlogrep) return;
  const seg=i.closest(".seg"); setLog(seg.dataset.tlogc, seg.dataset.tlogk, i.value.replace(/\D/g,"")||"X");
  i.blur(); render(); tmPaint();
});
document.getElementById("tm").addEventListener("keydown",ev=>{
  if(ev.key==="Enter"&&ev.target.dataset&&ev.target.dataset.tlogrep) ev.target.blur(); });
document.getElementById("tm-x").onclick=tmClose;
document.getElementById("tm-min").onclick=tmMin;
document.getElementById("tm-mini").onclick=tmMax;
function slCount(box){
  const det=box.closest("details.slog"); if(!det) return;
  const b=det.querySelector("summary b"); if(b) b.textContent=sDone(box.dataset.sd,box.dataset.sn)+"/"+box.querySelectorAll(".sl-row").length;
}
function slClick(ev,a){
  const btn=ev.target.closest(`[data-${a}="ok"]`); if(!btn) return false;
  const box=btn.closest(".sl-box"), i=+btn.dataset.si, B=box.dataset;
  sToggle(B.sd,B.sday,B.sn,i,B.spw);
  const x=sGet(B.sd,B.sn)[i]||{};
  btn.classList.toggle("on",!!x.ok);
  const w=box.querySelector(`[data-${a}="w"][data-si="${i}"]`); if(w&&document.activeElement!==w) w.value=x.w||"";
  slCount(box); return true;
}
function slChange(ev,a){
  const el=ev.target, f=el.dataset&&el.dataset[a]; if(f!=="w"&&f!=="r") return false;
  const box=el.closest(".sl-box"), B=box.dataset;
  sPut(B.sd,B.sday,B.sn,+el.dataset.si,{[f]:el.value.replace(/\s*kg$/i,"").trim()});
  slCount(box); return true;
}
const appEl=document.getElementById("app");
appEl.addEventListener("click",ev=>{ if(slClick(ev,"sl")) ev.preventDefault(); });
appEl.addEventListener("change",ev=>{ slChange(ev,"sl"); });
appEl.addEventListener("keydown",ev=>{ if(ev.key==="Enter"&&ev.target.dataset&&ev.target.dataset.sl) ev.target.blur(); });
appEl.addEventListener("toggle",ev=>{ const d=ev.target;
  if(d.classList&&d.classList.contains("slog")){ if(d.open) SOPEN.add(d.dataset.sn); else SOPEN.delete(d.dataset.sn); } },true);
const tmEl=document.getElementById("tm");
tmEl.addEventListener("click",ev=>{ slClick(ev,"tsl"); });
tmEl.addEventListener("change",ev=>{ slChange(ev,"tsl"); });
tmEl.addEventListener("keydown",ev=>{ if(ev.key==="Enter"&&ev.target.dataset&&ev.target.dataset.tsl) ev.target.blur(); });
document.getElementById("tm-play").onclick=tmPause;
document.getElementById("tm-next").onclick=()=>tmGo(1);
document.getElementById("tm-prev").onclick=()=>tmGo(-1);
document.getElementById("ver").textContent=(D.source.match(/v\d+/)||[""])[0];
document.getElementById("wk-up").onclick=()=>{if(week<WLAST){week++;render();}};
document.getElementById("wk-down").onclick=()=>{if(week>WBASE){week--;render();}};
document.getElementById("m-home").onclick=()=>{mode="home";
  const td=DOWK(ymd(new Date())); tab=D.home[td]&&!D.home[td].rest?td:"cal"; syncBackWeek();
  document.getElementById("m-home").setAttribute("aria-pressed","true");
  document.getElementById("m-back").setAttribute("aria-pressed","false");render();};
document.getElementById("m-back").onclick=()=>{mode="back";
  const c=cycDayOf(ymd(new Date())); tab=c&&!(D.back[c.day]&&D.back[c.day].rest)?c.day:"bcal"; syncBackWeek();
  document.getElementById("m-back").setAttribute("aria-pressed","true");
  document.getElementById("m-home").setAttribute("aria-pressed","false");render();};
try{
  /* 2026-11-01 v22 — 9일 사이클이 기본 화면이다. 재택(옛 판)은 같은 날 일부러 열어 둔 경우만 이어 본다.
     사이클 쪽은 오늘 할 일차로 간다 — 날짜가 바뀌었거나 오늘 일차가 달라졌으면(일차를 다시 눌렀으면) 옮긴다.
     기준일이 아직 없으면 달력으로 보내 한 번 누르게 한다. 주차 · 중량은 오늘 날짜가 정한다 */
  const st=JSON.parse(localStorage.getItem("wk-prog")||"null"), t0=ymd(new Date());
  const cd=cycDayOf(t0), okTab=k=>tabsFor("back").some(([x])=>x===k);
  if(st&&st.mode==="home"&&st.on===t0){ mode="home"; tab=st.tab||DOWK(t0);
    if(!tabsFor("home").some(([x])=>x===tab)) tab="cal"; }
  else {
    mode="back";
    const moved=!st||st.mode!=="back"||st.on!==t0||st.cd!==(cd?cd.day:null);
    if(!cd) tab="bcal";
    else if(moved) tab=okTab(cd.day)?cd.day:"bcal";
    else tab=okTab(st.tab)?st.tab:(okTab(cd.day)?cd.day:"bcal");
  }
}catch(e){ mode="back"; tab="bcal"; }
document.getElementById("m-home").setAttribute("aria-pressed",mode==="home");
document.getElementById("m-back").setAttribute("aria-pressed",mode==="back");
syncBackWeek();
render();
</script>
</body>
</html>
"""


if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="운동프로그램 워크북 → 모바일 HTML")
    ap.add_argument("xlsx", nargs="?", help="워크북 경로 (생략하면 폴더에서 자동 탐색)")
    ap.add_argument("-o", "--out", help="출력 HTML 경로 (기본: 같은 이름 _모바일.html)")
    a = ap.parse_args()
    if a.xlsx:
        src = Path(a.xlsx)
        if not src.exists():
            sys.exit(f"파일을 찾을 수 없습니다: {src}")
    else:
        src = pick_latest(Path(__file__).resolve().parent)
        print(f"워크북: {src.name}")
    out = Path(a.out) if a.out else src.with_name(src.stem + "_모바일.html")
    build(src, out)
