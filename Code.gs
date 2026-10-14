/**
 * @OnlyCurrentDoc
 * 침실2 배치 웹앱 저장소 (휴먼시아 5단지 51형)
 * 이 시트의 '배치' 탭: A2 = 배치 전체(JSON), B2 = 저장 번호(올라가기만 함), C2 = 마지막 저장 시각
 * 웹앱 배포: 실행 사용자 = 나, 액세스 = 모든 사용자 → 받은 /exec 주소를 index.html 의 EXEC_URL 에 넣기
 */
const TAB = '배치';
const MAX_ITEMS = 200;
const MAX_LEN = 45000; // 시트 한 칸 한도(5만 자)보다 작게

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(TAB);
  if (!sh) {
    sh = ss.insertSheet(TAB);
    sh.getRange('A1:C1').setValues([['배치(JSON)', '저장 번호', '저장 시각']]);
  }
  return sh;
}

function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// 읽기: 지금 저장된 배치와 저장 번호
function doGet() {
  const row = sheet_().getRange('A2:B2').getValues()[0];
  let layout = null;
  try { layout = row[0] ? JSON.parse(row[0]) : null; } catch (e) { layout = null; }
  return out_({ ok: true, rev: Number(row[1]) || 0, layout: layout });
}

// 저장: 링크 가진 누구나 배치 저장 가능(사용자 결정). 모양만 검사하고 마지막 저장이 이김
function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return out_({ ok: false, error: 'bad_json' }); }
  const L = body && body.layout;
  if (!valid_(L)) return out_({ ok: false, error: 'bad_layout' });
  const s = JSON.stringify(L);
  if (s.length > MAX_LEN) return out_({ ok: false, error: 'too_big' });
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet_();
    const rev = (Number(sh.getRange('B2').getValue()) || 0) + 1;
    sh.getRange('A2:C2').setValues([[s, rev, new Date()]]);
    return out_({ ok: true, rev: rev });
  } finally {
    lock.releaseLock();
  }
}

function valid_(L) {
  if (!L || typeof L !== 'object' || !L.room || !Array.isArray(L.items)) return false;
  if (L.items.length > MAX_ITEMS) return false;
  const num = v => typeof v === 'number' && isFinite(v) && v >= 0 && v < 100000;
  const R = L.room;
  if (!['W', 'D', 'H', 'doorOff', 'doorW', 'winOff', 'winW', 'winSill', 'winHead'].every(k => num(R[k]))) return false;
  return L.items.every(it => it && typeof it.name === 'string' && it.name.length <= 60 &&
    ['id', 'w', 'd', 'h', 'x', 'y', 'rot'].every(k => num(it[k])) &&
    typeof it.type === 'string' && /^#[0-9a-fA-F]{6}$/.test(String(it.color)));
}
