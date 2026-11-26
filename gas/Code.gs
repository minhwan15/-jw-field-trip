/**
 * 정왕고 체험학습 작성기 — 담임 검토 서버 (구글 시트 + Apps Script)
 *
 * 설치(1회)
 *  1) 새 구글 시트 → 확장 프로그램 → Apps Script → 이 코드를 붙여넣고 저장
 *  2) 위쪽 함수 선택에서 setup 선택 → 실행 → 권한 허용
 *     ('신청' 탭, '설정' 탭, 드라이브 PDF 폴더가 만들어집니다)
 *  3) 배포 → 새 배포 → 유형: 웹 앱
 *       - 다음 사용자 인증 정보로 실행: 나
 *       - 액세스 권한이 있는 사용자: 모든 사용자
 *     → 배포 → 나온 웹 앱 URL(…/exec)을 앱 config.js에 넣습니다
 *  4) '설정' 탭에 반별 담임 비밀번호를 적습니다
 *  5) (선택) 담임 기록함 연동: 시트 메뉴 [체험학습] → [기록함 연동 설정]에서 기록함 주소와 연동 키 입력
 *
 * 코드를 고친 뒤에는 배포 → 배포 관리 → 수정(연필) → 버전: 새 버전 → 배포 (URL은 그대로 유지)
 */
const SHEET = '신청', CONF = '설정', FOLDER_NAME = '체험학습_검토_PDF';
const HEAD = ['ID', '구분', '학년', '반', '번호', '이름', '기간', '목적지', '상태', '보완사항', '요청시각', '검토시각', '검토자', '요청횟수', 'PIN', 'PDF파일ID', '데이터', '기록함', '자산ID', '수정내역'];
const C = {}; HEAD.forEach((h, i) => C[h] = i);
const DOCNAME = { app: '신청서', rep: '결과보고서' };

function setup() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEAD); sh.setFrozenRows(1);
    sh.hideColumns(C['PIN'] + 1, 3); // PIN·PDF파일ID·데이터 숨김
  }
  const cf = ss.getSheetByName(CONF) || ss.insertSheet(CONF);
  if (cf.getLastRow() === 0) {
    cf.appendRow(['학년', '반', '담임 비밀번호', '비고']);
    for (let i = 1; i <= 12; i++) cf.appendRow([3, i, '', '3학년 ' + i + '반 담임 비밀번호']);
    cf.setFrozenRows(1);
  }
  folder_();
}

function folder_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('FOLDER');
  if (id) { try { return DriveApp.getFolderById(id); } catch (e) { } }
  const f = DriveApp.createFolder(FOLDER_NAME);
  p.setProperty('FOLDER', f.getId());
  return f;
}

function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return "'" + Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm'); } // 앞의 ' 는 시트가 날짜로 바꾸지 않게 하는 표시
function salt_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('SALT');
  if (!s) { s = Utilities.getUuid(); p.setProperty('SALT', s); }
  return s;
}
function hash_(pin) {
  return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt_() + '|' + String(pin)));
}
function str_(v, n) { return String(v == null ? '' : v).trim().slice(0, n || 200); }

function doGet() { return out_({ ok: true, app: 'jw-field-trip' }); }

function doPost(e) {
  let b;
  try { b = JSON.parse(e.postData.contents); } catch (x) { return out_({ ok: false, msg: '잘못된 요청입니다.' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    const f = { submit: submit_, draft: draft_, status: status_, mine: mine_, mypdf: mypdf_, tlist: tlist_, tget: tget_, treview: treview_, tedit: tedit_, tdelete: tdelete_, tresetpin: tresetpin_ }[b.action];
    return out_(f ? f(b) : { ok: false, msg: '알 수 없는 요청입니다.' });
  } catch (err) {
    return out_({ ok: false, msg: String(err && err.message || err) });
  } finally { lock.releaseLock(); }
}

function rows_() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET);
  const v = sh.getDataRange().getValues();
  return { sh: sh, v: v };
}
function student_(b) {
  const n = x => { const t = str_(x, 3); return /^\d+$/.test(t) ? String(parseInt(t, 10)) : t; };
  const s = { grade: n(b.grade), cls: n(b.cls), num: n(b.num), name: str_(b.name, 20).replace(/\s+/g, ''), pin: str_(b.pin, 20) };
  if (!/^\d{1,2}$/.test(s.grade) || !/^\d{1,2}$/.test(s.cls) || !/^\d{1,3}$/.test(s.num) || !s.name) throw new Error('학년·반·번호·이름을 확인하세요.');
  if (!/^\d{4}$/.test(s.pin)) throw new Error('비밀번호는 숫자 4자리입니다.');
  return s;
}
function same_(r, s) { return String(r[C['학년']]) === s.grade && String(r[C['반']]) === s.cls && String(r[C['번호']]) === s.num && String(r[C['이름']]) === s.name; }

/* 학생: 검토 요청(같은 학생·같은 서류·같은 시작일이면 덮어쓰기) */
const DRAFT = '작성 중';
const RESET = 'RESET';
function claim_(v, i, h, phone) { // 담임이 초기화한 행: 이번에 입력한 비밀번호·휴대폰으로 새로 정함
  const sh = rows_().sh; v[i][C['PIN']] = h; sh.getRange(i + 1, C['PIN'] + 1).setValue(h);
  if (phone) { let d = {}; try { d = JSON.parse(String(v[i][C['데이터']] || '{}').replace(/^'/, '')); } catch (e) { }
    if (d && typeof d === 'object') { d.phone = String(phone); const j = JSON.stringify(d); v[i][C['데이터']] = j; sh.getRange(i + 1, C['데이터'] + 1).setValue("'" + j); } }
}
function guard_(v, s, b) { // 같은 학생의 기존 행이 있으면 비밀번호·휴대폰이 같아야 함
  const h = hash_(s.pin);
  for (let i = 1; i < v.length; i++) {
    if (String(v[i][C['학년']]) === s.grade && String(v[i][C['반']]) === s.cls && String(v[i][C['번호']]) === s.num) {
      if (!same_(v[i], s)) throw new Error('같은 학번으로 다른 이름의 신청이 있습니다. 담임 선생님께 문의하세요.');
      if (v[i][C['PIN']] === RESET) claim_(v, i, h, b.phone);
      if (v[i][C['PIN']] !== h || (b.phone && !phoneOk_(v[i], b.phone))) throw new Error('처음 정한 비밀번호 또는 휴대폰 번호와 다릅니다.');
    }
  }
  return h;
}
function dropDraft_(sh, s, prevId, keepId) { // 기간을 바꿔 저장하면 예전 '작성 중' 행은 지움
  prevId = str_(prevId, 80);
  if (!prevId || prevId === keepId) return;
  const v = sh.getDataRange().getValues();
  for (let i = v.length - 1; i >= 1; i--) if (v[i][C['ID']] === prevId && same_(v[i], s) && v[i][C['상태']] === DRAFT) { sh.deleteRow(i + 1); return; }
}
/* 학생: 보관함에 작성 중으로 저장(담임 목록에는 안 보임) */
function draft_(b) {
  const s = student_(b);
  const doc = b.doc === 'rep' ? 'rep' : 'app';
  const start = str_(b.start, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start)) throw new Error('체험학습 날짜를 먼저 정하세요.');
  const id = [doc, s.grade, s.cls, s.num, start].join('-');
  const t = rows_(), sh = t.sh, v = t.v, h = guard_(v, s, b);
  let ri = -1;
  for (let i = 1; i < v.length; i++) if (v[i][C['ID']] === id) { ri = i; break; }
  if (ri > 0 && v[ri][C['상태']] !== DRAFT) return { ok: true, id: id, kept: true, status: String(v[ri][C['상태']]) };
  const row = ["'" + id, DOCNAME[doc], s.grade, s.cls, s.num, "'" + s.name, "'" + str_(b.period, 60), "'" + str_(b.dest, 100),
    DRAFT, '', now_(), '', '', 0, "'" + h, '', "'" + str_(JSON.stringify(b.data || {}), 40000)];
  if (ri > 0) sh.getRange(ri + 1, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
  dropDraft_(sh, s, b.prevId, id);
  return { ok: true, id: id, status: DRAFT };
}

function submit_(b) {
  const s = student_(b);
  const doc = b.doc === 'rep' ? 'rep' : 'app';
  const start = str_(b.start, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(start)) throw new Error('체험학습 날짜를 확인하세요.');
  if (!b.pdf || String(b.pdf).length > 30e6) throw new Error('PDF가 없거나 너무 큽니다.');
  const id = [doc, s.grade, s.cls, s.num, start].join('-');
  const t = rows_(); const sh = t.sh, v = t.v;
  let ri = -1;
  for (let i = 1; i < v.length; i++) if (v[i][C['ID']] === id) { ri = i; break; }
  const h = guard_(v, s, b);
  const fname = DOCNAME[doc] + '_' + s.grade + '-' + s.cls + '-' + s.num + '_' + s.name + '_' + start + '.pdf';
  const blob = Utilities.newBlob(Utilities.base64Decode(String(b.pdf)), 'application/pdf', fname);
  const file = folder_().createFile(blob);
  if (ri > 0 && v[ri][C['PDF파일ID']]) { try { DriveApp.getFileById(v[ri][C['PDF파일ID']]).setTrashed(true); } catch (e) { } }
  const sent = ri > 0 && v[ri][C['상태']] !== DRAFT;
  const cnt = sent ? (Number(v[ri][C['요청횟수']]) || 1) + 1 : 1;
  const row = ["'" + id, DOCNAME[doc], s.grade, s.cls, s.num, "'" + s.name, "'" + str_(b.period, 60), "'" + str_(b.dest, 100),
    cnt > 1 ? '재검토 대기' : '검토 대기', ri > 0 && v[ri][C['보완사항']] ? "'" + v[ri][C['보완사항']] : '', now_(), ri > 0 && v[ri][C['검토시각']] ? "'" + txt_(v[ri][C['검토시각']]) : '', ri > 0 && v[ri][C['검토자']] ? "'" + v[ri][C['검토자']] : '',
    cnt, "'" + h, "'" + file.getId(), "'" + str_(JSON.stringify(b.data || {}), 40000)];
  if (ri > 0) sh.getRange(ri + 1, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  saveAssets_(sh, ri > 0 ? ri + 1 : sh.getLastRow(), ri > 0 ? v[ri][C['자산ID']] : '', b.assets, id);
  dropDraft_(sh, s, b.prevId, id);
  return { ok: true, id: id, status: String(row[C['상태']]) };
}

/* 학생: 내 요청 현황 */
function digits_(x) { return String(x == null ? '' : x).replace(/\D/g, ''); }
function phoneOk_(r, phone) { // 신청서에 적은 학생 휴대폰과 일치해야 함(적지 않은 경우는 통과)
  let d = {}; try { d = JSON.parse(r[C['데이터']] || '{}'); } catch (e) { }
  const saved = digits_(d.phone);
  return !saved || saved === digits_(phone);
}
function myRows_(b) {
  const s = student_(b);
  if (!/^01\d{8,9}$/.test(digits_(b.phone))) throw new Error('휴대폰 번호를 확인하세요.');
  const v = rows_().v, h = hash_(s.pin), out = [];
  let wrong = false;
  for (let i = 1; i < v.length; i++) {
    if (!same_(v[i], s)) continue;
    if (v[i][C['PIN']] === RESET) claim_(v, i, h, b.phone);
    if (v[i][C['PIN']] !== h || !phoneOk_(v[i], b.phone)) { wrong = true; continue; }
    out.push(v[i]);
  }
  if (!out.length && wrong) throw new Error('비밀번호 또는 휴대폰 번호가 다릅니다.');
  return out;
}
function mine_(b) { // 학생: 보낸 신청 내용을 작성칸으로 다시 불러오기
  const r = myRows_(b).filter(x => x[C['ID']] === str_(b.id, 80))[0];
  if (!r) throw new Error('불러올 신청을 찾을 수 없습니다.');
  let data = {}; try { data = JSON.parse(r[C['데이터']] || '{}'); } catch (e) { }
  return { ok: true, item: pub_(r), data: data };
}
function mypdf_(b) { // 학생: 담임이 승인한 서류 PDF(서명·사진 포함, 제출 당시 그대로) 받기
  const r = myRows_(b).filter(x => x[C['ID']] === str_(b.id, 80))[0];
  if (!r) throw new Error('서류를 찾을 수 없습니다.');
  if (r[C['상태']] !== '승인') throw new Error('담임 선생님이 승인한 뒤에 저장할 수 있어요.');
  let pdf = '';
  try { pdf = Utilities.base64Encode(DriveApp.getFileById(r[C['PDF파일ID']]).getBlob().getBytes()); } catch (e) { throw new Error('서류 파일을 찾을 수 없습니다. 담임 선생님께 문의하세요.'); }
  return { ok: true, item: pub_(r), pdf: pdf };
}
function status_(b) {
  const list = myRows_(b).map(pub_);
  list.sort((a, c) => String(c.at).localeCompare(String(a.at)));
  return { ok: true, list: list };
}
function txt_(x) { return x instanceof Date ? Utilities.formatDate(x, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(x == null ? '' : x); }
function pub_(r) {
  return { id: r[C['ID']], doc: r[C['구분']], grade: r[C['학년']], cls: r[C['반']], num: r[C['번호']], name: r[C['이름']], period: r[C['기간']], dest: r[C['목적지']],
    status: r[C['상태']], comment: r[C['보완사항']], at: txt_(r[C['요청시각']]), reviewedAt: txt_(r[C['검토시각']]), count: r[C['요청횟수']], days: days_(r), edits: String(r[C['수정내역']] || '') };
}
function days_(r) { try { return String(JSON.parse(String(r[C['데이터']] || '{}')).days || ''); } catch (e) { return ''; } }

/* 담임 인증: 설정 탭의 반별 비밀번호 */
function teacher_(b) {
  const n = x => { const t = str_(x, 3); return /^\d+$/.test(t) ? String(parseInt(t, 10)) : t; };
  const pw = str_(b.tpw, 100), grade = n(b.grade), cls = n(b.cls);
  if (!pw) throw new Error('비밀번호를 입력하세요.');
  const v = SpreadsheetApp.getActive().getSheetByName(CONF).getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const g = String(v[i][0]).trim(), c = String(v[i][1]).trim(), p = String(v[i][2]).trim();
    if (!p || p !== pw) continue;
    if (g === grade && c === cls) return { grade: g, cls: c, label: g + '-' + c + ' 담임' };
  }
  throw new Error('학년·반 또는 비밀번호가 맞지 않습니다.');
}
function inScope_(t, r) { return String(r[C['학년']]) === t.grade && String(r[C['반']]) === t.cls; }

function tlist_(b) {
  const t = teacher_(b), v = rows_().v, list = [];
  for (let i = 1; i < v.length; i++) if (v[i][C['ID']] && v[i][C['상태']] !== DRAFT && inScope_(t, v[i])) list.push(pub_(v[i]));
  list.sort((a, c) => String(c.at).localeCompare(String(a.at)));
  return { ok: true, who: t.label, list: list };
}
function find_(t, id) {
  const r = rows_();
  for (let i = 1; i < r.v.length; i++) if (r.v[i][C['ID']] === id) {
    if (!inScope_(t, r.v[i])) throw new Error('다른 반 요청입니다.');
    return { sh: r.sh, i: i, r: r.v[i] };
  }
  throw new Error('요청을 찾을 수 없습니다.');
}
function saveAssets_(sh, rowNo, oldId, assets, id) { // 서명·사진(담임이 대신 고칠 때 다시 그리기용)
  if (!assets) return;
  const hd = sh.getRange(1, C['자산ID'] + 1);
  if (hd.getValue() !== '자산ID') { hd.setValue('자산ID'); try { sh.hideColumns(C['자산ID'] + 1); } catch (e) { } }
  if (oldId) { try { DriveApp.getFileById(String(oldId)).setTrashed(true); } catch (e) { } }
  const f = folder_().createFile(Utilities.newBlob(JSON.stringify(assets), 'application/json', id + '_자료.json'));
  sh.getRange(rowNo, C['자산ID'] + 1).setValue("'" + f.getId());
}
function tget_(b) {
  const t = teacher_(b), f = find_(t, str_(b.id, 80));
  let pdf = '';
  try { pdf = Utilities.base64Encode(DriveApp.getFileById(f.r[C['PDF파일ID']]).getBlob().getBytes()); } catch (e) { }
  let data = {};
  try { data = JSON.parse(f.r[C['데이터']] || '{}'); } catch (e) { }
  let assets = null;
  if (b.assets && f.r[C['자산ID']]) { try { assets = JSON.parse(DriveApp.getFileById(String(f.r[C['자산ID']])).getBlob().getDataAsString()); } catch (e) { } }
  return { ok: true, item: pub_(f.r), pdf: b.nopdf ? '' : pdf, data: data, assets: assets };
}
function tresetpin_(b) { // 담임이 학생 비밀번호 초기화: 다음 로그인 때 입력한 비밀번호·휴대폰이 새 값이 됨
  const t = teacher_(b), f = find_(t, str_(b.id, 80)), r = f.r, v = f.sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < v.length; i++) if (String(v[i][C['학년']]) === String(r[C['학년']]) && String(v[i][C['반']]) === String(r[C['반']]) && String(v[i][C['번호']]) === String(r[C['번호']])) { f.sh.getRange(i + 1, C['PIN'] + 1).setValue(RESET); n++; }
  return { ok: true, n: n };
}
function tdelete_(b) { // 담임이 학생 서류를 삭제(시트 행 + 드라이브 PDF·자료 휴지통)
  const t = teacher_(b), f = find_(t, str_(b.id, 80));
  [C['PDF파일ID'], C['자산ID']].forEach(c => { const v = f.r[c]; if (v) { try { DriveApp.getFileById(String(v)).setTrashed(true); } catch (e) { } } });
  f.sh.deleteRow(f.i + 1);
  return { ok: true };
}
function tedit_(b) { // 담임이 학생 서류 내용을 직접 고쳐 저장(상태는 그대로)
  const t = teacher_(b), f = find_(t, str_(b.id, 80));
  if (f.r[C['상태']] === DRAFT) throw new Error('학생이 아직 작성 중인 서류입니다.');
  if (!b.pdf || String(b.pdf).length > 30e6) throw new Error('PDF가 없거나 너무 큽니다.');
  const r = f.r, fname = r[C['구분']] + '_' + r[C['학년']] + '-' + r[C['반']] + '-' + r[C['번호']] + '_' + r[C['이름']] + '_담임수정.pdf';
  const file = folder_().createFile(Utilities.newBlob(Utilities.base64Decode(String(b.pdf)), 'application/pdf', fname));
  if (r[C['PDF파일ID']]) { try { DriveApp.getFileById(String(r[C['PDF파일ID']])).setTrashed(true); } catch (e) { } }
  const data = "'" + str_(JSON.stringify(b.data || {}), 40000);
  f.sh.getRange(f.i + 1, C['기간'] + 1, 1, 2).setValues([["'" + str_(b.period, 60), "'" + str_(b.dest, 100)]]);
  f.sh.getRange(f.i + 1, C['PDF파일ID'] + 1, 1, 2).setValues([["'" + file.getId(), data]]);
  f.sh.getRange(f.i + 1, C['검토자'] + 1).setValue("'" + t.label + ' (내용 수정 ' + String(now_()).replace(/^'/, '') + ')');
  const ch = (Array.isArray(b.changes) ? b.changes : []).slice(0, 30).map(x => '· ' + str_(x, 300)).join('\n');
  if (ch) {
    const hd = f.sh.getRange(1, C['수정내역'] + 1);
    if (hd.getValue() !== '수정내역') hd.setValue('수정내역');
    const prev = String(r[C['수정내역']] || '');
    f.sh.getRange(f.i + 1, C['수정내역'] + 1).setValue("'" + ('[' + String(now_()).replace(/^'/, '') + ' 담임 수정]\n' + ch + (prev ? '\n\n' + prev : '')).slice(0, 4000));
  }
  let sync = null;
  if (r[C['상태']] === '승인') {
    const nr = f.sh.getRange(f.i + 1, 1, 1, HEAD.length).getValues()[0];
    sync = pushClassCalendar_(nr);
    if (!sync.off) f.sh.getRange(f.i + 1, C['기록함'] + 1).setValue("'" + (sync.ok ? '✔ ' : '✖ ') + sync.msg + ' (' + String(now_()).replace(/^'/, '') + ')');
  }
  return { ok: true, status: String(r[C['상태']]), sync: sync };
}
/* ── 담임 기록함(class-calendar) 연동: 승인 시 해당 반 기록함에 체험학습 기록 등록 ── */
const HOLIDAYS = ['2026-12-25', '2027-01-01'];
function onOpen() {
  SpreadsheetApp.getUi().createMenu('체험학습').addItem('기록함 연동 설정', 'setupClassCalendar').addItem('기록함 연동 끄기', 'disableClassCalendar').addToUi();
}
function setupClassCalendar() {
  const ui = SpreadsheetApp.getUi(), p = PropertiesService.getScriptProperties();
  const u = ui.prompt('기록함 연동 (1/2)', '담임 기록함 주소를 입력하세요. 예: https://class-calendar.vercel.app', ui.ButtonSet.OK_CANCEL);
  if (u.getSelectedButton() !== ui.Button.OK) return;
  const url = String(u.getResponseText()).trim().replace(/\/+$/, '');
  if (!/^https:\/\/[^\s/]+$/.test(url)) { ui.alert('주소 형식이 맞지 않습니다. https://로 시작하는 기본 주소만 입력하세요.'); return; }
  const k = ui.prompt('기록함 연동 (2/2)', '기록함 Vercel 환경변수 EXPERIENCE_INTAKE_KEY와 같은 연동 키를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (k.getSelectedButton() !== ui.Button.OK) return;
  const key = String(k.getResponseText()).trim();
  if (key.length < 16) { ui.alert('연동 키는 16자 이상이어야 합니다.'); return; }
  p.setProperty('CC_URL', url); p.setProperty('CC_KEY', key);
  ui.alert('저장했습니다. 이제 담임이 승인하면 해당 반 기록함에 체험학습 기록이 자동으로 등록됩니다.');
}
function disableClassCalendar() { const p = PropertiesService.getScriptProperties(); p.deleteProperty('CC_URL'); p.deleteProperty('CC_KEY'); SpreadsheetApp.getUi().alert('기록함 연동을 껐습니다.'); }
function schoolDays_(s, e) { // 주말·휴일을 뺀 수업일(나이스 입력 날짜)
  const out = [], p = x => { const m = String(x).split('-').map(Number); return Date.UTC(m[0], m[1] - 1, m[2]); };
  for (let t = p(s); t <= p(e) && out.length < 60; t += 864e5) {
    const d = new Date(t), k = d.toISOString().slice(0, 10), w = d.getUTCDay();
    if (w !== 0 && w !== 6 && HOLIDAYS.indexOf(k) < 0) out.push(k);
  }
  return out;
}
function pushClassCalendar_(r) {
  const p = PropertiesService.getScriptProperties(), url = p.getProperty('CC_URL'), key = p.getProperty('CC_KEY');
  if (!url || !key) return { ok: false, off: true, msg: '기록함 연동 꺼짐' };
  let d = {}; try { d = JSON.parse(r[C['데이터']] || '{}'); } catch (e) { }
  const g = String(r[C['학년']]), c = String(r[C['반']]).padStart(2, '0'), n = String(r[C['번호']]).padStart(2, '0');
  const doc = r[C['구분']] === '결과보고서' ? 'rep' : 'app', half = d.type === 'half';
  const start = half ? d.hdate : d.start, end = half ? d.hdate : (d.end || d.start);
  if (!start) return { ok: false, msg: '기간 정보가 없어 쉬운출결에 보내지 못했습니다' };
  let pdf = '';
  try { const bytes = DriveApp.getFileById(r[C['PDF파일ID']]).getBlob().getBytes(); if (bytes.length <= 3 * 1024 * 1024) pdf = Utilities.base64Encode(bytes); } catch (e) { }
  const purpose = String(d.title || d.purpose || '').trim(), form = String(d.form || '').trim();
  const payload = { intakeId: String(r[C['ID']]), doc: doc, username: g + c, studentId: g + c + n, studentName: String(r[C['이름']]),
    startDate: start, endDate: end, attendanceDates: half ? [start] : schoolDays_(start, end), location: String(d.dest || ''),
    reason: (form ? form + (purpose ? ' - ' : '') : '') + purpose + (half ? ' (반일 ' + (d.ampm === 'pm' ? '오후' : '오전') + ')' : ''),
    approvedAt: String(now_()).replace(/^'/, ''), pdf: pdf };
  try {
    const res = UrlFetchApp.fetch(url + '/api/experience-intake', { method: 'post', contentType: 'application/json', headers: { 'x-intake-key': key }, payload: JSON.stringify(payload), muteHttpExceptions: true, followRedirects: false });
    let j = {}; try { j = JSON.parse(res.getContentText()); } catch (e) { }
    if (res.getResponseCode() === 200 && j.ok) return { ok: true, msg: j.created ? '쉬운출결에 새 체험학습 기록 등록' : '쉬운출결 체험학습 기록 갱신' };
    return { ok: false, msg: '쉬운출결 전송 실패: ' + (j.message || j.error || ('HTTP ' + res.getResponseCode())) };
  } catch (e) { return { ok: false, msg: '쉬운출결에 연결하지 못했습니다' }; }
}

function treview_(b) {
  const t = teacher_(b), f = find_(t, str_(b.id, 80));
  const st = b.status === '승인' ? '승인' : '보완 요청';
  const cm = str_(b.comment, 2000);
  if (st === '보완 요청' && !cm) throw new Error('보완할 내용을 적어 주세요.');
  f.sh.getRange(f.i + 1, C['상태'] + 1, 1, 1).setValue(st);
  f.sh.getRange(f.i + 1, C['보완사항'] + 1, 1, 1).setValue(cm ? "'" + cm : '');
  f.sh.getRange(f.i + 1, C['검토시각'] + 1, 1, 2).setValues([[now_(), "'" + t.label]]);
  let sync = null;
  if (st === '승인') {
    const hd = f.sh.getRange(1, C['기록함'] + 1);
    if (hd.getValue() !== '기록함') hd.setValue('기록함');
    sync = pushClassCalendar_(f.r);
    if (!sync.off) f.sh.getRange(f.i + 1, C['기록함'] + 1).setValue("'" + (sync.ok ? '✔ ' : '✖ ') + sync.msg + ' (' + String(now_()).replace(/^'/, '') + ')');
  }
  return { ok: true, status: st, sync: sync };
}
