/**
 * 한능검 심화 기출 500 — 학생 제출 기록용 Apps Script
 * 사용법: 구글 시트 > 확장 프로그램 > Apps Script 에 이 코드를 붙여 넣고
 *        ① setup 실행 → ② 배포 > 새 배포 > 웹 앱 (실행: 나 / 액세스: 모든 사용자)
 *        ③ 나온 웹 앱 URL을 앱의 config.json "scriptUrl" 에 넣기
 */
const SH = { LOG: '제출기록', CODES: '학생코드', SUM: '학생별 요약', QS: '문항별 오답', PASTE: '결과코드 붙여넣기' };
const LOG_HEAD = ['접수 시각', '기록 ID', '개인 코드', '이름(코드표)', '유형', '내용', '점수', '급수',
  '정답 수', '문항 수', '정답률', '걸린 시간(분)', '틀린 문항', '앱에서 푼 시각', '경로'];
const MODE = { mock: '모의고사', era: '시대별', random: '랜덤10', timer: '타이머', wrong: '오답노트' };

/* ---------- 최초 1회: 시트 만들기 ---------- */
function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone('Asia/Seoul');
  const mk = (name, head, widths) => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold').setBackground('#1f3a8a').setFontColor('#ffffff');
      sh.setFrozenRows(1);
      (widths || []).forEach((w, i) => w && sh.setColumnWidth(i + 1, w));
    }
    return sh;
  };
  mk(SH.CODES, ['개인 코드', '이름', '반', '메모'], [110, 100, 70, 200]);
  mk(SH.LOG, LOG_HEAD, [140, 110, 90, 90, 80, 200, 50, 60, 60, 60, 60, 80, 260, 140, 70]);
  mk(SH.SUM, ['개인 코드', '이름', '반', '제출 횟수', '모의고사 횟수', '모의고사 최고점', '모의고사 평균', '최근 모의고사', '연습 정답률', '마지막 제출'], [90, 90, 60, 70, 80, 90, 90, 140, 80, 140]);
  mk(SH.QS, ['문항', '회차', '번호', '틀린 횟수'], [80, 60, 60, 80]);
  mk(SH.PASTE, ['결과 코드 (학생이 보낸 HK1-… 코드를 A열에 한 줄씩)', '처리 결과'], [520, 200]);
  const s1 = ss.getSheetByName('시트1') || ss.getSheetByName('Sheet1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
  ss.setActiveSheet(ss.getSheetByName(SH.CODES));
  SpreadsheetApp.getUi().alert('준비 완료!\n\n1) "학생코드" 시트에 개인 코드와 이름을 적어 주세요.\n2) 배포 > 새 배포 > 웹 앱으로 배포하세요.');
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('한능검 앱')
    .addItem('요약 새로고침', 'refreshSummary')
    .addItem('결과 코드 반영', 'importPastedCodes')
    .addItem('처음 설정(setup)', 'setup')
    .addToUi();
  try { refreshSummary(); } catch (e) {}
}

/* ---------- 웹 앱 ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action === 'check') return json({ ok: !!codeInfo(p.code) });
  return json({ ok: true, app: 'hanguksa-500' });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const r = JSON.parse(e.postData.contents);
    return json(addRecord(r, '자동'));
  } catch (err) {
    return json({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

function addRecord(r, via) {
  if (!r || !r.rid || !r.code) return { ok: false, error: 'bad record' };
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH.LOG);
  if (sh.getLastRow() > 1) {
    const hit = sh.getRange(2, 2, sh.getLastRow() - 1, 1).createTextFinder(String(r.rid)).matchEntireCell(true).findNext();
    if (hit) return { ok: true, dup: true };
  }
  const info = codeInfo(r.code);
  const total = Number(r.total) || 0, cor = Number(r.correct) || 0;
  sh.appendRow([
    new Date(), safe(r.rid), safe(String(r.code).toUpperCase()), info ? safe(info.name) : '미등록 코드',
    MODE[r.mode] || safe(r.mode), safe(r.label), r.mode === 'mock' ? Number(r.score) || 0 : '',
    r.mode === 'mock' ? safe(r.grade) : '', cor, total, total ? Math.round(cor / total * 1000) / 10 : '',
    Math.round((Number(r.sec) || 0) / 6) / 10, safe(r.wrong), safe(r.at), via
  ]);
  return { ok: true };
}

/* ---------- 학생이 결과 코드로 낸 경우 ---------- */
function importPastedCodes() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SH.PASTE);
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const vals = sh.getRange(2, 1, n, 2).getValues();
  let added = 0;
  vals.forEach((row, i) => {
    const c = String(row[0] || '').trim();
    if (!c || row[1]) return;
    let msg;
    try {
      if (c.indexOf('HK1-') !== 0) throw new Error('HK1- 로 시작하지 않음');
      const r = JSON.parse(Utilities.newBlob(Utilities.base64Decode(c.slice(4))).getDataAsString('UTF-8'));
      const res = addRecord(r, '코드');
      msg = res.dup ? '이미 있음' : (res.ok ? '반영됨' : '오류: ' + res.error);
      if (res.ok && !res.dup) added++;
    } catch (err) { msg = '오류: ' + err.message; }
    sh.getRange(i + 2, 2).setValue(msg);
  });
  refreshSummary();
  SpreadsheetApp.getActive().toast(added + '건 반영했어요.');
}

/* ---------- 요약 시트 다시 만들기 ---------- */
function refreshSummary() {
  const ss = SpreadsheetApp.getActive();
  const log = ss.getSheetByName(SH.LOG);
  if (!log || log.getLastRow() < 2) return;
  const rows = log.getRange(2, 1, log.getLastRow() - 1, LOG_HEAD.length).getValues();
  const codes = codeTable();
  const by = {}, qWrong = {};
  rows.forEach(r => {
    const code = String(r[2]);
    const o = by[code] || (by[code] = { n: 0, mock: [], pc: 0, pt: 0, last: null, lastMock: '' });
    o.n++;
    if (!o.last || r[0] > o.last) o.last = r[0];
    if (r[4] === '모의고사') { o.mock.push(Number(r[6]) || 0); o.lastMock = r[5] + ' ' + r[6] + '점'; }
    else { o.pc += Number(r[8]) || 0; o.pt += Number(r[9]) || 0; }
    String(r[12] || '').split(',').filter(Boolean).forEach(id => qWrong[id] = (qWrong[id] || 0) + 1);
  });
  const sum = ss.getSheetByName(SH.SUM);
  const out = Object.keys(by).sort().map(code => {
    const o = by[code], info = codes[code] || {};
    const avg = o.mock.length ? Math.round(o.mock.reduce((a, b) => a + b, 0) / o.mock.length * 10) / 10 : '';
    return [code, info.name || '미등록', info.cls || '', o.n, o.mock.length, o.mock.length ? Math.max.apply(null, o.mock) : '',
      avg, o.lastMock, o.pt ? Math.round(o.pc / o.pt * 1000) / 10 + '%' : '', o.last];
  });
  if (sum.getLastRow() > 1) sum.getRange(2, 1, sum.getLastRow() - 1, 10).clearContent();
  if (out.length) sum.getRange(2, 1, out.length, 10).setValues(out);
  const qs = ss.getSheetByName(SH.QS);
  const qOut = Object.keys(qWrong).map(id => { const p = id.split('-'); return [id, Number(p[0]), Number(p[1]), qWrong[id]]; })
    .sort((a, b) => b[3] - a[3]);
  if (qs.getLastRow() > 1) qs.getRange(2, 1, qs.getLastRow() - 1, 4).clearContent();
  if (qOut.length) qs.getRange(2, 1, qOut.length, 4).setValues(qOut);
}

/* ---------- 도우미 ---------- */
function codeTable() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SH.CODES);
  const t = {};
  if (!sh || sh.getLastRow() < 2) return t;
  sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => {
    const c = String(r[0] || '').trim().toUpperCase();
    if (c) t[c] = { name: String(r[1] || ''), cls: String(r[2] || '') };
  });
  return t;
}
function codeInfo(code) { return codeTable()[String(code || '').trim().toUpperCase()] || null; }
function safe(v) { const s = String(v == null ? '' : v).slice(0, 2000); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function json(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
