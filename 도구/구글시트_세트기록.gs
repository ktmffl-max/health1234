/* 세트 기록 중계 — 폰 앱 → 이 구글 시트 → PC(세트기록.py pull) (2026-11-13)
 *
 * 설치 (한 번만):
 *  1. 구글 시트 새로 만들기 (이름 아무거나, 예: 세트기록)
 *  2. 메뉴 확장 프로그램 → Apps Script → 이 파일 내용 전부 붙여 넣고 저장
 *  3. 배포 → 새 배포 → 유형 '웹 앱'
 *       실행 사용자: 나 · 액세스 권한: 모든 사용자 → 배포 → 권한 허용
 *  4. 나온 '웹 앱 URL'(…/exec)을 Claude에게 알려 주고, 폰 앱 '자동 저장 설정' 칸에도 붙여 넣기
 *
 * 앱은 바뀐 날짜의 세트 전부를 보낸다 → 그 날짜 줄을 지우고 새로 넣는다(그래서 지운 세트도 반영).
 * 이 주소는 아는 사람만 쓸 수 있다 — 공개 저장소에 올리지 말 것.
 */
const SHEET = '세트기록';
const COLS = ['날짜', '사이클', '일차', '종목', '세트', '무게', '횟수', '달성'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.getRange(1, 1, 1, COLS.length).setValues([COLS]); }
  return s;
}

function text_(v) {
  return v instanceof Date ? Utilities.formatDate(v, 'Asia/Seoul', 'yyyy-MM-dd') : String(v);
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const dates = new Set(body.dates || []);
    const s = sheet_();
    const vals = s.getDataRange().getValues().slice(1).map(r => r.map(text_));
    const rows = vals.filter(r => r[0] && !dates.has(r[0]))
      .concat((body.rows || []).map(r => COLS.map((_, i) => String(r[i] == null ? '' : r[i]))));
    rows.sort((a, b) => (a[0] + a[2] + a[3]).localeCompare(b[0] + b[2] + b[3]) || (+a[4]) - (+b[4]));
    s.clearContents();
    s.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    if (rows.length) s.getRange(2, 1, rows.length, COLS.length).setNumberFormat('@').setValues(rows);
    return ContentService.createTextOutput('ok ' + rows.length);
  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  const vals = sheet_().getDataRange().getValues().map(r => r.map(text_).join('\t'));
  return ContentService.createTextOutput(vals.join('\n')).setMimeType(ContentService.MimeType.TEXT);
}
