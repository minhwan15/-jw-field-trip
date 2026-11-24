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
 *
 * 코드를 고친 뒤에는 배포 → 배포 관리 → 수정(연필) → 버전: 새 버전 → 배포 (URL은 그대로 유지)
 */
const SHEET = '신청', CONF = '설정', FOLDER_NAME = '체험학습_검토_PDF';
const HEAD = ['ID', '구분', '학년', '반', '번호', '이름', '기간', '목적지', '상태', '보완사항', '요청시각', '검토시각', '검토자', '요청횟수', 'PIN', 'PDF파일ID', '데이터'];
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
    const f = { submit: submit_, status: status_, tlist: tlist_, tget: tget_, treview: treview_ }[b.action];
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
  const h = hash_(s.pin);
  if (ri > 0) {
    if (!same_(v[ri], s)) throw new Error('같은 학번으로 다른 이름의 요청이 있습니다. 담임 선생님께 문의하세요.');
    if (v[ri][C['PIN']] !== h) throw new Error('처음 요청할 때 정한 비밀번호와 다릅니다.');
  }
  const fname = DOCNAME[doc] + '_' + s.grade + '-' + s.cls + '-' + s.num + '_' + s.name + '_' + start + '.pdf';
  const blob = Utilities.newBlob(Utilities.base64Decode(String(b.pdf)), 'application/pdf', fname);
  const file = folder_().createFile(blob);
  if (ri > 0 && v[ri][C['PDF파일ID']]) { try { DriveApp.getFileById(v[ri][C['PDF파일ID']]).setTrashed(true); } catch (e) { } }
  const cnt = ri > 0 ? (Number(v[ri][C['요청횟수']]) || 1) + 1 : 1;
  const row = ["'" + id, DOCNAME[doc], s.grade, s.cls, s.num, "'" + s.name, "'" + str_(b.period, 60), "'" + str_(b.dest, 100),
    cnt > 1 ? '재검토 대기' : '검토 대기', ri > 0 && v[ri][C['보완사항']] ? "'" + v[ri][C['보완사항']] : '', now_(), ri > 0 && v[ri][C['검토시각']] ? "'" + txt_(v[ri][C['검토시각']]) : '', ri > 0 && v[ri][C['검토자']] ? "'" + v[ri][C['검토자']] : '',
    cnt, "'" + h, "'" + file.getId(), "'" + str_(JSON.stringify(b.data || {}), 40000)];
  if (ri > 0) sh.getRange(ri + 1, 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, id: id, status: String(row[C['상태']]) };
}

/* 학생: 내 요청 현황 */
function status_(b) {
  const s = student_(b);
  const v = rows_().v, h = hash_(s.pin), list = [];
  let wrongPin = false;
  for (let i = 1; i < v.length; i++) {
    if (!same_(v[i], s)) continue;
    if (v[i][C['PIN']] !== h) { wrongPin = true; continue; }
    list.push(pub_(v[i]));
  }
  if (!list.length && wrongPin) throw new Error('비밀번호가 다릅니다.');
  list.sort((a, c) => String(c.at).localeCompare(String(a.at)));
  return { ok: true, list: list };
}
function txt_(x) { return x instanceof Date ? Utilities.formatDate(x, 'Asia/Seoul', 'yyyy-MM-dd HH:mm') : String(x == null ? '' : x); }
function pub_(r) {
  return { id: r[C['ID']], doc: r[C['구분']], grade: r[C['학년']], cls: r[C['반']], num: r[C['번호']], name: r[C['이름']], period: r[C['기간']], dest: r[C['목적지']],
    status: r[C['상태']], comment: r[C['보완사항']], at: txt_(r[C['요청시각']]), reviewedAt: txt_(r[C['검토시각']]), count: r[C['요청횟수']] };
}

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
  for (let i = 1; i < v.length; i++) if (v[i][C['ID']] && inScope_(t, v[i])) list.push(pub_(v[i]));
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
function tget_(b) {
  const t = teacher_(b), f = find_(t, str_(b.id, 80));
  let pdf = '';
  try { pdf = Utilities.base64Encode(DriveApp.getFileById(f.r[C['PDF파일ID']]).getBlob().getBytes()); } catch (e) { }
  let data = {};
  try { data = JSON.parse(f.r[C['데이터']] || '{}'); } catch (e) { }
  return { ok: true, item: pub_(f.r), pdf: pdf, data: data };
}
function treview_(b) {
  const t = teacher_(b), f = find_(t, str_(b.id, 80));
  const st = b.status === '승인' ? '승인' : '보완 요청';
  const cm = str_(b.comment, 2000);
  if (st === '보완 요청' && !cm) throw new Error('보완할 내용을 적어 주세요.');
  f.sh.getRange(f.i + 1, C['상태'] + 1, 1, 1).setValue(st);
  f.sh.getRange(f.i + 1, C['보완사항'] + 1, 1, 1).setValue(cm ? "'" + cm : '');
  f.sh.getRange(f.i + 1, C['검토시각'] + 1, 1, 2).setValues([[now_(), "'" + t.label]]);
  return { ok: true, status: st };
}
