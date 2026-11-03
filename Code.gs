/**
 * 한걸음 행동노트 (ABC 기록 · 빈도 기록 · 지속시간 기록)
 * - Google Apps Script 웹앱 서버 코드
 * - 기록법마다 구글 스프레드시트 1개, 학생마다 시트 탭 1개를 자동 생성합니다.
 */

const APP_TITLE = '한걸음 행동노트';
const DAYS_KO = ['일', '월', '화', '수', '목', '금', '토'];
const INT_LABEL = { 1: '약', 2: '중', 3: '강' };

/** 기록법별 파일 이름 · 헤더 · 셀 서식 · 내부 키 */
const METHODS = {
  abc: {
    label: 'ABC 기록',
    fileName: '행동기록_ABC',
    headers: ['기록ID', '날짜', '시각', '요일', '학생', '장소/활동', '선행사건(A)', '행동(B)', '강도(1~3)', '후속결과(C)', '메모', '입력시각'],
    keys:    ['id', 'date', 'time', 'dow', 'student', 'place', 'antecedent', 'behavior', 'intensity', 'consequence', 'memo', 'stamp'],
    formats: ['@', '@', '@', '@', '@', '@', '@', '@', '0', '@', '@', '@']
  },
  freq: {
    label: '빈도 기록',
    fileName: '행동기록_빈도',
    headers: ['기록ID', '날짜', '시각', '요일', '학생', '표적행동', '강도(1~3)', '강도', '장소/활동', '메모', '입력시각'],
    keys:    ['id', 'date', 'time', 'dow', 'student', 'behavior', 'intensity', 'intensityLabel', 'place', 'memo', 'stamp'],
    formats: ['@', '@', '@', '@', '@', '@', '0', '@', '@', '@', '@']
  },
  dur: {
    label: '지속시간 기록',
    fileName: '행동기록_지속시간',
    headers: ['기록ID', '날짜', '요일', '학생', '행동', '시작시각', '종료시각', '지속시간(초)', '지속시간', '장소/활동', '메모', '입력시각'],
    keys:    ['id', 'date', 'dow', 'student', 'behavior', 'start', 'end', 'seconds', 'durationText', 'place', 'memo', 'stamp'],
    formats: ['@', '@', '@', '@', '@', '@', '@', '0', '@', '@', '@', '@']
  }
};

const DEFAULT_CONFIG = {
  students: [
    { name: '예시학생', abc: ['소리 지르기', '자리 이탈', '물건 던지기'], freq: ['소리 지르기', '자리 이탈', '물건 던지기'], dur: ['자리 이탈', '울기', '책상에 엎드리기'] }
  ],
  antecedents: ['과제 제시', '어려운 과제', '활동 전환', '선호물 제거', '요구 거절', '교사 관심 없음', '또래 상호작용', '기다리기', '소음/감각 자극', '지시 반복'],
  consequences: ['언어적 제지', '언어적 재지시', '과제 중단/철회', '관심 제공', '선호물 제공', '계획된 무시', '타임아웃', '신체적 촉구', '또래 반응(웃음 등)', '자리 이동', '진정 공간 이동'],
  places: ['국어', '수학', '통합/특별실', '체육', '점심시간', '쉬는 시간', '등교', '하교']
};

function props_() { return PropertiesService.getScriptProperties(); }

/* ───────────── 웹앱 진입점 ───────────── */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ───────────── 설정 ───────────── */
function getConfig_() {
  const raw = props_().getProperty('CONFIG');
  const base = JSON.parse(JSON.stringify(DEFAULT_CONFIG));
  if (!raw) return base;
  let cfg;
  try { cfg = Object.assign(base, JSON.parse(raw)); } catch (e) { cfg = base; }
  // 값이 빠져 있어도 화면이 멈추지 않도록 정리
  cfg.students = (Array.isArray(cfg.students) ? cfg.students : []).map(s => ({
    name: String((s && s.name) || ''), abc: list_(s && s.abc), freq: list_(s && s.freq), dur: list_(s && s.dur)
  }));
  ['antecedents', 'consequences', 'places'].forEach(k => { cfg[k] = list_(cfg[k]); });
  // 예전 기본값 '언어적 꾸중' → '언어적 제지', '언어적 재지시'로 바꿈
  const ci = cfg.consequences.indexOf('언어적 꾸중');
  if (ci > -1) {
    cfg.consequences.splice(ci, 1, '언어적 제지');
    if (cfg.consequences.indexOf('언어적 재지시') === -1) cfg.consequences.splice(ci + 1, 0, '언어적 재지시');
  }
  return cfg;
}

/**
 * 문제 점검용: 스크립트 편집기에서 이 함수를 선택해 [실행]하세요.
 * 처음 실행하면 권한 승인 창이 뜨며, 승인하면 권한 문제도 함께 해결됩니다.
 * 결과는 아래 [실행 로그]에 표시됩니다.
 */
function diagnose() {
  const log = (ok, msg) => console.log((ok ? '✅ ' : '❌ ') + msg);
  try { const st = getAppState(); log(true, '설정 불러오기 OK (학생 ' + st.config.students.length + '명)'); }
  catch (e) { log(false, '설정 불러오기 실패: ' + e.message); }
  try { HtmlService.createHtmlOutputFromFile('Index'); log(true, 'Index.html 파일 찾음'); }
  catch (e) { log(false, 'Index.html 파일을 찾지 못함 → 파일 이름이 정확히 Index 인지 확인: ' + e.message); }
  Object.keys(METHODS).forEach(k => {
    const id = props_().getProperty('SS_' + k);
    if (!id) { log(true, METHODS[k].label + ' 시트: 아직 없음(첫 기록 때 자동 생성)'); return; }
    try { log(true, METHODS[k].label + ' 시트 열기 OK: ' + SpreadsheetApp.openById(id).getUrl()); }
    catch (e) { log(false, METHODS[k].label + ' 시트 열기 실패: ' + e.message); }
  });
}

function list_(a) {
  const arr = Array.isArray(a) ? a : String(a || '').split(/\n/);
  const out = [];
  arr.forEach(x => {
    const v = String(x).trim();
    if (v && out.indexOf(v) === -1) out.push(v);
  });
  return out;
}

function getAppState(student) {
  const p = props_();
  const files = {};
  Object.keys(METHODS).forEach(k => {
    const id = p.getProperty('SS_' + k);
    files[k] = id ? 'https://docs.google.com/spreadsheets/d/' + id + '/edit' : '';
  });
  const config = getConfig_();
  const out = { config: config, files: files };
  // 앱을 열 때 마지막으로 보던 학생의 기록도 같이 보내 서버 왕복을 줄임
  if (student && config.students.some(s => s.name === student)) {
    try { out.packed = getPacked(['abc', 'freq', 'dur'], student); out.packedFor = student; } catch (e) { /* 기록은 나중에 따로 */ }
  }
  return out;
}

function saveConfig(cfg) {
  cfg = cfg || {};
  const seen = {};
  const clean = {
    students: (cfg.students || []).map(s => ({
      name: String(s.name || '').trim(),
      abc: list_(s.abc), freq: list_(s.freq), dur: list_(s.dur)
    })).filter(s => s.name && !seen[s.name] && (seen[s.name] = true)),
    antecedents: list_(cfg.antecedents),
    consequences: list_(cfg.consequences),
    places: list_(cfg.places)
  };
  props_().setProperty('CONFIG', JSON.stringify(clean));
  return getAppState();
}

/* ───────────── 스프레드시트 / 시트 탭 ───────────── */
function getSpreadsheet_(method) {
  const key = 'SS_' + method;
  const id = props_().getProperty(key);
  if (id) {
    try { return SpreadsheetApp.openById(id); } catch (e) { /* 삭제된 경우 새로 만듦 */ }
  }
  const ss = SpreadsheetApp.create(METHODS[method].fileName);
  props_().setProperty(key, ss.getId());
  return ss;
}

function sheetName_(name) {
  return String(name || '').replace(/[\[\]\*\/\\\?:]/g, '_').trim().slice(0, 90) || '이름없음';
}

function getStudentSheet_(ss, method, student) {
  const name = sheetName_(student);
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const def = METHODS[method];
  sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, def.headers.length)
    .setValues([def.headers]).setFontWeight('bold').setBackground('#e7ecff');
  sh.setFrozenRows(1);
  // 처음 생성 시 생기는 빈 기본 시트(시트1/Sheet1) 정리
  ss.getSheets().forEach(s => {
    if (s.getSheetId() !== sh.getSheetId() && s.getLastRow() === 0 && /^(Sheet|시트)\s?\d*$/.test(s.getName())) {
      ss.deleteSheet(s);
    }
  });
  return sh;
}

function appendRow_(method, student, values) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = getSpreadsheet_(method);
    const sh = getStudentSheet_(ss, method, student);
    const def = METHODS[method];
    const row = sh.getLastRow() + 1;
    const rng = sh.getRange(row, 1, 1, def.headers.length);
    rng.setNumberFormats([def.formats]); // 날짜·시각이 자동 변환되지 않도록 텍스트 서식
    rng.setValues([values]);
    const url = ss.getUrl() + '#gid=' + sh.getSheetId();
    cacheUpdate_(method, student, o => {
      const row = o.k.map(k => { const v = values[def.keys.indexOf(k)]; return (k === 'intensity' || k === 'seconds') ? (Number(v) || 0) : String(v === undefined ? '' : v); });
      o.v.push(row); o.url = o.url || url;
      const di = o.k.indexOf('date'), ti = o.k.indexOf(method === 'dur' ? 'start' : 'time');
      o.v.sort((a, b) => (a[di] + a[ti]).localeCompare(b[di] + b[ti]));
    });
    return { ok: true, id: values[0], url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── 캐시 (기록 조회 속도 향상) ───────────── */
const CACHE_TTL = 21600; // 6시간(최대값). 앱으로 저장·삭제하면 캐시도 즉시 고쳐지고, 시트를 직접 고쳤을 땐 🔄 버튼으로 반영
function cacheKey_(m, s) { return 'rec_' + m + '_' + Utilities.base64EncodeWebSafe(Utilities.newBlob(String(s)).getBytes()).slice(0, 200); }
const CHUNK = 30000;
function cachePut_(key, str) {
  try {
    const n = Math.ceil(str.length / CHUNK);
    if (n > 150) return;                       // 너무 크면 캐시 생략
    const o = {};
    for (let i = 0; i < n; i++) o[key + '_' + i] = str.substr(i * CHUNK, CHUNK);
    o[key + '_n'] = String(n);
    CacheService.getScriptCache().putAll(o, CACHE_TTL);
  } catch (e) { /* noop */ }
}
function cacheGet_(key) {
  try {
    const c = CacheService.getScriptCache(), n = Number(c.get(key + '_n'));
    if (!n) return null;
    const ks = []; for (let i = 0; i < n; i++) ks.push(key + '_' + i);
    const got = c.getAll(ks), parts = [];
    for (let i = 0; i < n; i++) { if (got[ks[i]] == null) return null; parts.push(got[ks[i]]); }
    return parts.join('');
  } catch (e) { return null; }
}
/** 캐시가 있으면 지우지 않고 내용만 고쳐서 다시 저장 (다음 조회도 빠르게) */
function cacheUpdate_(m, s, fn) {
  const key = cacheKey_(m, s), hit = cacheGet_(key);
  if (!hit) return;
  try { const o = JSON.parse(hit); fn(o); cachePut_(key, JSON.stringify(o)); } catch (e) { clearCache_(m, s); }
}
function clearCache_(m, s) { try { CacheService.getScriptCache().remove(cacheKey_(m, s) + '_n'); } catch (e) { /* noop */ } }

/* ───────────── 시간 도우미 ───────────── */
function tz_() { return Session.getScriptTimeZone() || 'Asia/Seoul'; }
function fmt_(d, p) { return Utilities.formatDate(d, tz_(), p); }
function dowOf_(dateStr) {
  const p = String(dateStr).split('-').map(Number);
  return DAYS_KO[new Date(p[0], p[1] - 1, p[2]).getDay()];
}
function validDate_(s) { return /^\d{4}-\d{2}-\d{2}$/.test(String(s || '')) ? String(s) : ''; }
function validTime_(s) { return /^\d{2}:\d{2}(:\d{2})?$/.test(String(s || '')) ? String(s) : ''; }
function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }
function need_(v, label) { if (!String(v || '').trim()) throw new Error(label + '을(를) 입력해 주세요.'); }
function fmtSec_(sec) {
  sec = Math.max(0, Math.round(sec));
  const h = Math.floor(sec / 3600), m = Math.floor((sec % 3600) / 60), s = sec % 60;
  const pad = n => (n < 10 ? '0' : '') + n;
  return h ? h + ':' + pad(m) + ':' + pad(s) : pad(m) + ':' + pad(s);
}

/* ───────────── 저장 ───────────── */
function saveABC(r) {
  need_(r.student, '학생'); need_(r.behavior, '행동(B)');
  const now = new Date();
  const date = validDate_(r.date) || fmt_(now, 'yyyy-MM-dd');
  const time = validTime_(r.time) || fmt_(now, 'HH:mm');
  const i = Number(r.intensity);
  return appendRow_('abc', r.student, [
    newId_(), date, time, dowOf_(date), r.student, r.place || '', r.antecedent || '', r.behavior,
    i >= 1 && i <= 3 ? i : '', r.consequence || '', r.memo || '', fmt_(now, 'yyyy-MM-dd HH:mm:ss')
  ]);
}

function saveFrequency(r) {
  need_(r.student, '학생'); need_(r.behavior, '표적행동');
  const i = Number(r.intensity);
  if (!(i >= 1 && i <= 3)) throw new Error('강도는 1~3 중 하나여야 합니다.');
  const when = r.ts ? new Date(Number(r.ts)) : new Date();
  const date = fmt_(when, 'yyyy-MM-dd');
  return appendRow_('freq', r.student, [
    newId_(), date, fmt_(when, 'HH:mm:ss'), dowOf_(date), r.student, r.behavior, i, INT_LABEL[i],
    r.place || '', r.memo || '', fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss')
  ]);
}

function saveDuration(r) {
  need_(r.student, '학생'); need_(r.behavior, '행동');
  const s = new Date(Number(r.startTs)), e = new Date(Number(r.endTs));
  if (isNaN(s) || isNaN(e) || e < s) throw new Error('시작/종료 시각이 올바르지 않습니다.');
  const sec = Math.round((e - s) / 1000);
  const date = fmt_(s, 'yyyy-MM-dd');
  return appendRow_('dur', r.student, [
    newId_(), date, dowOf_(date), r.student, r.behavior, fmt_(s, 'HH:mm:ss'), fmt_(e, 'HH:mm:ss'),
    sec, fmtSec_(sec), r.place || '', r.memo || '', fmt_(new Date(), 'yyyy-MM-dd HH:mm:ss')
  ]);
}

function deleteRecord(method, student, id) {
  const ssId = props_().getProperty('SS_' + method);
  if (!ssId) throw new Error('기록 파일이 없습니다.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = SpreadsheetApp.openById(ssId).getSheetByName(sheetName_(student));
    if (!sh || sh.getLastRow() < 2) throw new Error('기록을 찾을 수 없습니다.');
    const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) {
      if (String(ids[i][0]) === String(id)) { sh.deleteRow(i + 2); cacheUpdate_(method, student, o => { o.v = o.v.filter(r => String(r[0]) !== String(id)); }); return { ok: true }; }
    }
    throw new Error('기록을 찾을 수 없습니다.');
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── 조회 ───────────── */
/** 기록 읽기(압축 형식) : { k:[열 이름], v:[[값…]…], url } */
function readPacked_(method, student, noCache) {
  const def = METHODS[method];
  if (!def) throw new Error('알 수 없는 기록법입니다.');
  const keys = def.keys.filter(k => k !== 'stamp' && k !== 'student');
  const empty = { k: keys, v: [], url: '' };
  const ssId = props_().getProperty('SS_' + method);
  if (!ssId || !student) return empty;
  const key = cacheKey_(method, student);
  if (!noCache) { const hit = cacheGet_(key); if (hit) { try { return JSON.parse(hit); } catch (e) { /* 다시 읽기 */ } } }

  let ss;
  try { ss = SpreadsheetApp.openById(ssId); } catch (e) { return empty; }
  const sh = ss.getSheetByName(sheetName_(student));
  if (!sh || sh.getLastRow() < 2) { empty.url = ss.getUrl(); return empty; }
  const idx = keys.map(k => def.keys.indexOf(k));
  const vals = sh.getRange(2, 1, sh.getLastRow() - 1, def.keys.length).getValues();
  const di = keys.indexOf('date'), ti = keys.indexOf(method === 'dur' ? 'start' : 'time');
  const v = [];
  vals.forEach(r => {
    const a = idx.map((j, n) => {
      let x = r[j]; const k = keys[n];
      if (x instanceof Date) x = fmt_(x, k === 'date' ? 'yyyy-MM-dd' : 'HH:mm:ss');
      if (k === 'intensity' || k === 'seconds') return Number(x) || 0;
      return (x === null || x === undefined) ? '' : String(x);
    });
    if (validDate_(a[di])) v.push(a);
  });
  v.sort((a, b) => (a[di] + a[ti]).localeCompare(b[di] + b[ti]));
  const out = { k: keys, v: v, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
  cachePut_(key, JSON.stringify(out));
  return out;
}

/** 화면용: 여러 기록법을 한 번에 (서버 왕복 1번) */
function getPacked(methods, student, noCache) {
  const out = {};
  (methods || []).forEach(m => { out[m] = readPacked_(m, student, noCache); });
  return out;
}
