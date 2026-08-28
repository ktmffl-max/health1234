# -*- coding: utf-8 -*-
"""말소리 조각을 굽고 한 파일로 이어 붙인다.

왜 이렇게 하나 —
  앱은 기기 TTS를 안 쓴다. 화면이 꺼지면 크롬이 speechSynthesis를 통째로
  거부해서, 정작 폰을 주머니에 넣은 순간에 안 나오기 때문이다. 대신 미리
  구운 소리를 오디오 시계에 예약한다(build_mobile.py의 tmCueAt 참조).

  그런데 '다음 종목' 안내는 종목 이름·세트·횟수·중량이 매번 다르다.
  문장마다 wav를 굽는 건 불가능하다 — 중량은 주차마다 오르고 앱에서 고칠
  수도 있으니 경우의 수가 끝이 없다. 그래서 조각을 굽는다:
  종목 이름 하나하나, 근육군 하나하나, 숫자 0~99와 백 단위, 그리고
  '킬로 · 세트 · 회 · 에서' 같은 토막. 재생할 때 그 순간 화면에 떠 있는
  값으로 조각을 이어 붙인다. 그래서 증량도 앱 내 수정도 그냥 따라온다.

  조각을 파일 200개로 두면 첫 방문에 요청이 200번 난다. 그래서 하나의
  wav로 이어 붙이고(sprite) 어느 구간이 무슨 말인지만 표로 넘긴다.
  Web Audio의 start(when, offset, duration)이 구간만 잘라 낸다.

  파일 이름에 내용 해시를 박는다. 워크북을 고쳐 앱을 다시 빌드해도 말소리가
  그대로면 이름도 그대로라, 폰이 몇 MB를 다시 받지 않는다.

윈도우 SAPI(Heami)로 굽는다. SAPI가 없는 환경에서는 조용히 건너뛰고
이미 있는 스프라이트를 그대로 쓴다 — 빌드가 깨지지 않는다.
"""

import hashlib
import re
import subprocess
import wave
from array import array
from pathlib import Path

RATE = 16000          # SAPI Heami 기본. 기존 prep/work/rest/done도 같은 규격이다
GAP = 0.07            # 조각 사이 숨. 이어 붙인 말이 뭉치지 않을 만큼만
PAD = 0.05            # 잘라낸 앞뒤로 남기는 여유
SILENCE = 20          # 이 진폭 아래를 묵음으로 본다 (16비트, 최대 32767)

CACHE = ".voice"      # 조각 wav 보관. 다시 빌드해도 이미 구운 건 안 굽는다

# ── 숫자를 한글로 ────────────────────────────────────────────────
# SAPI에 "62"를 던지면 뭐라 읽을지 보장이 안 된다. 한글로 바꿔 던지면
# 읽는 법이 하나뿐이라 어긋날 여지가 없다. 아래 규칙은 앱 쪽 koNum()과
# 똑같이 맞춰져 있다 — 한쪽만 고치면 조각을 못 찾는다.
ONES = ["", "일", "이", "삼", "사", "오", "육", "칠", "팔", "구"]
FRAC = {"5": "점오", "25": "점이오", "75": "점칠오", "125": "점일이오"}
UNITS = ["다음 종목", "킬로", "세트", "회", "에서", "좌우"]


def sino(n):
    """0~99를 한글로. 10은 '일십'이 아니라 '십'이다."""
    if n == 0:
        return "영"
    t, o = divmod(n, 10)
    return ("" if t == 0 else "십" if t == 1 else ONES[t] + "십") + ONES[o]


def hundreds(h):
    return ("" if h == 1 else ONES[h]) + "백"


def num_keys(text):
    """'62.5' → ['육십이', '점오'].  숫자 꼴이 아니면 None."""
    m = re.fullmatch(r"(\d+)(?:\.(\d+))?", text.strip())
    if not m:
        return None
    n, frac = int(m.group(1)), (m.group(2) or "").rstrip("0")
    if n >= 1000:
        return None
    keys = []
    if n >= 100:
        keys.append(hundreds(n // 100))
        if n % 100:
            keys.append(sino(n % 100))
    else:
        keys.append(sino(n))
    if frac:
        if frac in FRAC:
            keys.append(FRAC[frac])
        else:
            keys.append("점")
            keys += [sino(int(c)) for c in frac]
    return keys


def num_words(text):
    """숫자를 소리 나는 대로 이어 붙인 한 덩어리. 조각 이름이 아니라 '읽을 말'."""
    k = num_keys(text)
    return "".join(k) if k else None


def spell(text):
    """문장 속 숫자를 전부 한글로. '75 10개 채움' → '칠십오 십개 채움'"""
    return re.sub(r"\d+(?:\.\d+)?",
                  lambda m: num_words(m.group(0)) or m.group(0), text)


# ── 읽을 말 다듬기 ──────────────────────────────────────────────
def say_name(s):
    """종목 이름. 괄호 안은 대체 설명이라 읽지 않는다 —
    '친업 (어시스트 가능)' → '친업'"""
    s = re.sub(r"[(\[（].*?[)\]）]", " ", s)
    s = s.replace("·", " ").replace("/", " ").replace("—", " ")
    return re.sub(r"\s+", " ", s).strip(" -—·")


def say_mg(s):
    """근육군. 여기서는 괄호 안이 정보다 — '승모(중부)' → '승모 중부'"""
    s = re.sub(r"[(\[（）)\]]", " ", s).replace("·", " ").replace("/", " ")
    return re.sub(r"\s+", " ", s).strip(" -—·")


def say_reps(s):
    """워크북 횟수 칸의 예외 표기. 숫자·범위는 앱이 조립하므로
    여기 오는 건 '10 (좌우)' 같은 것뿐이다."""
    m = re.match(r"^\s*(\d+(?:\.\d+)?)(?:\s*~\s*(\d+(?:\.\d+)?))?\s*(.*)$", s)
    if not m or not m.group(1):
        return say_mg(s)
    head = num_words(m.group(1))
    if m.group(2):
        head += " 에서 " + num_words(m.group(2))
    return re.sub(r"\s+", " ", (head + " 회 " + say_mg(m.group(3) or "")).strip())


def say_weight(s):
    """중량 칸의 예외 표기. '맨몸', '75 10개 채움' 따위."""
    return spell(say_mg(s))


def is_plain_number(s):
    """숫자 하나 또는 '30~35' 같은 범위. 이런 건 앱이 조각으로 조립한다 —
    주차마다 오르고 앱에서 고칠 수 있으니 통째로 구워 둘 수가 없다."""
    return re.fullmatch(r"\d+(?:\.\d+)?(?:\s*~\s*\d+(?:\.\d+)?)?",
                        s.strip()) is not None


# ── SAPI로 굽기 ─────────────────────────────────────────────────
PS = """
param([string]$List, [string]$OutDir)
Add-Type -AssemblyName System.Speech
$s = New-Object System.Speech.Synthesis.SpeechSynthesizer
$ko = $s.GetInstalledVoices() | Where-Object { $_.VoiceInfo.Culture.Name -eq 'ko-KR' } |
      Select-Object -First 1
if ($ko) { $s.SelectVoice($ko.VoiceInfo.Name) } else { exit 3 }
$fmt = New-Object System.Speech.AudioFormat.SpeechAudioFormatInfo(
    16000,
    [System.Speech.AudioFormat.AudioBitsPerSample]::Sixteen,
    [System.Speech.AudioFormat.AudioChannel]::Mono)
foreach ($line in (Get-Content -LiteralPath $List -Encoding UTF8)) {
  if (-not $line) { continue }
  $p = $line -split \"`t\", 2
  if ($p.Count -lt 2) { continue }
  $s.SetOutputToWaveFile((Join-Path $OutDir ($p[0] + \".wav\")), $fmt)
  $s.Speak($p[1])
}
$s.SetOutputToNull()
$s.Dispose()
"""


def frag_id(text):
    return hashlib.sha1(text.encode("utf-8")).hexdigest()[:12]


def synth(missing, cache):
    """missing: {조각파일이름: 읽을 말}. 파워셸을 한 번만 띄워 전부 굽는다 —
    조각마다 프로세스를 새로 띄우면 200개에 2분이 넘는다."""
    if not missing:
        return True
    cache.mkdir(parents=True, exist_ok=True)
    lst = cache / "_list.tsv"
    lst.write_text("".join("%s\t%s\n" % (k, v) for k, v in missing.items()),
                   encoding="utf-8-sig")
    ps = cache / "_say.ps1"
    ps.write_text(PS, encoding="utf-8-sig")
    try:
        r = subprocess.run(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                            "-File", str(ps), "-List", str(lst), "-OutDir", str(cache)],
                           capture_output=True, timeout=1800)
    except (OSError, subprocess.SubprocessError):
        return False
    return r.returncode == 0


# ── wav 다루기 ──────────────────────────────────────────────────
def read_pcm(path):
    try:
        with wave.open(str(path), "rb") as w:
            if (w.getnchannels() != 1 or w.getsampwidth() != 2
                    or w.getframerate() != RATE):
                return None
            a = array("h")
            a.frombytes(w.readframes(w.getnframes()))
    except (OSError, wave.Error):
        return None
    return a


def trim(a):
    """앞뒤 묵음을 걷어낸다. SAPI가 조각마다 붙이는 1초 가까운 빈 구간을
    그대로 두면 이어 붙였을 때 말과 말 사이가 하염없이 벌어진다.

    기준은 최대 진폭 대비가 아니라 고정값이다. 비례로 잡으면 'ㅋ, ㅅ, ㅎ'
    처럼 조용하게 시작하는 첫소리가 통째로 잘려 나가 무슨 말인지 알아듣기
    어려워진다. SAPI가 넣는 빈 구간은 진짜 0에 가까우므로 아주 낮은
    고정값으로도 충분히 걸러진다."""
    if not a:
        return a
    th = SILENCE
    i, j = 0, len(a) - 1
    while i < len(a) and abs(a[i]) < th:
        i += 1
    while j > i and abs(a[j]) < th:
        j -= 1
    if i > j:
        return array("h")
    pad = int(PAD * RATE)
    return a[max(0, i - pad):min(len(a), j + pad)]


def build(site, phrases, extra):
    """phrases: {찾을 이름: 읽을 말}, extra: {찾을 이름: 이미 있는 wav 경로}

    돌려주는 것 — (스프라이트 상대경로, {찾을 이름: [시작초, 길이초]}).
    구울 수 없으면 (None, None)이라 부르는 쪽이 옛 스프라이트를 지킬 수 있다."""
    site = Path(site)
    cache = site.parent / CACHE
    ids = {k: frag_id(v) for k, v in phrases.items()}
    missing = {i: phrases[k] for k, i in ids.items()
               if not (cache / (i + ".wav")).exists()}
    if missing and not synth(missing, cache):
        return None, None

    order = sorted(extra) + sorted(phrases)
    pcm, offs, at = [], {}, 0
    gap = array("h", bytes(int(GAP * RATE) * 2))
    for key in order:
        src = Path(extra[key]) if key in extra else cache / (ids[key] + ".wav")
        a = read_pcm(src)
        if a is None:
            continue
        a = trim(a)
        if len(a) < RATE // 50:          # 0.02초도 안 되면 말이 아니다
            continue
        offs[key] = [round(at / RATE, 6), round(len(a) / RATE, 6)]
        pcm.append(a)
        pcm.append(gap)
        at += len(a) + len(gap)
    if not offs:
        return None, None

    body = b"".join(x.tobytes() for x in pcm)
    name = "sprite-%s.wav" % hashlib.sha1(body).hexdigest()[:10]
    out = site / "voice" / name
    out.parent.mkdir(parents=True, exist_ok=True)
    if not out.exists():
        with wave.open(str(out), "wb") as w:
            w.setnchannels(1)
            w.setsampwidth(2)
            w.setframerate(RATE)
            w.writeframes(body)
    for old in out.parent.glob("sprite-*.wav"):   # 지난 스프라이트는 치운다
        if old.name != name:
            try:
                old.unlink()
            except OSError:
                pass
    return "voice/" + name, offs
