/**
 * 한걸음 행동노트 (ABC 기록 · 빈도 기록 · 지속시간 기록)
 * - Google Apps Script 웹앱 서버 코드
 * - 기록법마다 구글 스프레드시트 1개, 학생마다 시트 탭 1개를 자동 생성합니다.
 * - AI 분석은 OpenRouter API(https://openrouter.ai/api/v1/chat/completions)를 사용합니다.
 */

const APP_TITLE = '한걸음 행동노트';
const OPENROUTER_URL = 'https://openrouter.ai/api/v1/chat/completions';
const DAYS_KO = ['일', '월', '화', '수', '목', '금', '토'];
const INT_LABEL = { 1: '약', 2: '중', 3: '강' };
const REPORT_SHEET = 'AI분석기록';

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
  model: '~anthropic/claude-sonnet-latest',
  // 분석할 때 고를 수 있는 모델 (한 줄: '모델ID | 별칭') — 아래 ID는 OpenRouter 문서에 나온 '최신 별칭'
  favorites: ['~anthropic/claude-sonnet-latest | Claude Sonnet (최신)', '~anthropic/claude-opus-latest | Claude Opus (최신·고품질)', '~openai/gpt-sol-latest | OpenAI GPT Sol (최신)'],
  fallbacks: [],      // 첫 모델이 실패하면 차례로 시도할 예비 모델
  autoAllowed: [],    // 자동 선택(openrouter/auto) 때 후보로 허용할 모델 패턴 (예: anthropic/*)
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
  cfg.model = String(cfg.model || DEFAULT_CONFIG.model);
  cfg.fast = cfg.fast !== false;
  ['favorites', 'fallbacks', 'autoAllowed'].forEach(k => { cfg[k] = list_(cfg[k]); });
  return cfg;
}

/**
 * 문제 점검용: 스크립트 편집기에서 이 함수를 선택해 [실행]하세요.
 * 처음 실행하면 권한 승인 창이 뜨며, 승인하면 권한 문제도 함께 해결됩니다.
 * 결과는 아래 [실행 로그]에 표시됩니다.
 */
function diagnose() {
  const log = (ok, msg) => console.log((ok ? '✅ ' : '❌ ') + msg);
  try { const st = getAppState(); log(true, '설정 불러오기 OK (학생 ' + st.config.students.length + '명, API 키 ' + (st.hasApiKey ? '있음' : '없음') + ')'); }
  catch (e) { log(false, '설정 불러오기 실패: ' + e.message); }
  try { HtmlService.createHtmlOutputFromFile('Index'); log(true, 'Index.html 파일 찾음'); }
  catch (e) { log(false, 'Index.html 파일을 찾지 못함 → 파일 이름이 정확히 Index 인지 확인: ' + e.message); }
  Object.keys(METHODS).forEach(k => {
    const id = props_().getProperty('SS_' + k);
    if (!id) { log(true, METHODS[k].label + ' 시트: 아직 없음(첫 기록 때 자동 생성)'); return; }
    try { log(true, METHODS[k].label + ' 시트 열기 OK: ' + SpreadsheetApp.openById(id).getUrl()); }
    catch (e) { log(false, METHODS[k].label + ' 시트 열기 실패: ' + e.message); }
  });
  if (props_().getProperty('OPENROUTER_API_KEY')) {
    try { log(true, 'OpenRouter 연결 OK: ' + testOpenRouter()); }
    catch (e) { log(false, 'OpenRouter 연결 실패: ' + e.message); }
  } else log(true, 'OpenRouter API 키 미설정(앱의 ⚙ 설정에서 저장)');
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
  const out = { config: config, hasApiKey: !!p.getProperty('OPENROUTER_API_KEY'), files: files };
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
    model: String(cfg.model || DEFAULT_CONFIG.model).trim(),
    fast: cfg.fast !== false,
    favorites: list_(cfg.favorites),
    fallbacks: list_(cfg.fallbacks).slice(0, 3),
    autoAllowed: list_(cfg.autoAllowed),
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

function saveApiKey(key) {
  key = String(key || '').trim();
  if (key) props_().setProperty('OPENROUTER_API_KEY', key);
  else props_().deleteProperty('OPENROUTER_API_KEY');
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

/** 서버 내부용(AI 분석 등): 객체 배열로 */
function getRecords(method, student, days, noCache) {
  const p = readPacked_(method, student, noCache);
  let rows = p.v.map(a => { const o = {}; p.k.forEach((k, i) => { o[k] = a[i]; }); return o; });
  days = Number(days) || 0;
  if (days > 0) {
    const cut = fmt_(new Date(Date.now() - (days - 1) * 86400000), 'yyyy-MM-dd');
    rows = rows.filter(o => o.date >= cut);
  }
  return { rows: rows, url: p.url };
}

/** 마지막으로 저장된 AI 분석 결과 (대시보드를 열 때 바로 보여주기용) */
function getLastReport(method, student) {
  const ssId = props_().getProperty('SS_' + method);
  if (!ssId) return null;
  let sh;
  try { sh = SpreadsheetApp.openById(ssId).getSheetByName(REPORT_SHEET); } catch (e) { return null; }
  if (!sh || sh.getLastRow() < 2) return null;
  const n = Math.min(300, sh.getLastRow() - 1), start = sh.getLastRow() - n + 1;
  const vals = sh.getRange(start, 1, n, 5).getValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][1]) === String(student)) {
      const w = vals[i][0];
      return { when: w instanceof Date ? fmt_(w, 'yyyy-MM-dd HH:mm') : String(w), period: String(vals[i][2]), model: String(vals[i][3]),
               markdown: String(vals[i][4]), url: sh.getParent().getUrl() + '#gid=' + sh.getSheetId() };
    }
  }
  return null;
}

/* ───────────── AI (OpenRouter) ───────────── */
const AUTO_MODEL = 'openrouter/auto';
function idOf_(line) { return String(line || '').split('|')[0].trim(); }

function callOpenRouter_(messages, maxTokens, modelOverride) {
  const key = props_().getProperty('OPENROUTER_API_KEY');
  if (!key) throw new Error('설정 탭에서 OpenRouter API 키를 먼저 저장해 주세요.');
  const cfg = getConfig_();
  const model = idOf_(modelOverride) || cfg.model;
  const body = { model: model, messages: messages, temperature: 0.4, max_tokens: maxTokens || 4000 };
  if (model === AUTO_MODEL) {
    // 자동 선택: OpenRouter가 요청 내용에 맞는 모델을 고름. 후보를 제한했으면 그 안에서만.
    if (cfg.autoAllowed.length) body.plugins = [{ id: 'auto-router', allowed_models: cfg.autoAllowed }];
  } else {
    // 예비 모델: 첫 모델이 오류(혼잡·중단 등)면 OpenRouter가 다음 모델로 자동 전환
    const fb = cfg.fallbacks.map(idOf_).filter(x => x && x !== model && x !== AUTO_MODEL);
    if (fb.length) body.models = [model].concat(fb);
  }
  if (cfg.fast) body.provider = { sort: 'throughput' }; // 빠른 제공자 우선
  const res = UrlFetchApp.fetch(OPENROUTER_URL, {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + key, 'HTTP-Referer': 'https://script.google.com', 'X-OpenRouter-Title': 'Behavior Recorder' },
    payload: JSON.stringify(body),
    muteHttpExceptions: true
  });
  const code = res.getResponseCode();
  let json = {};
  try { json = JSON.parse(res.getContentText()); } catch (e) { /* noop */ }
  if (code !== 200 || json.error) {
    const msg = (json.error && json.error.message) || res.getContentText().slice(0, 300);
    const hint = code === 401 ? ' (API 키를 확인해 주세요)' : code === 402 ? ' (OpenRouter 크레딧이 부족합니다)' :
      (code === 400 || code === 404) ? ' (모델 ID를 확인하거나 다른 모델을 골라 주세요)' : '';
    throw new Error('OpenRouter 오류 ' + code + ' [' + model + ']: ' + msg + hint);
  }
  let content = json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content;
  if (Array.isArray(content)) content = content.map(c => c.text || '').join('');
  if (!content) throw new Error('AI 응답이 비어 있습니다. 다른 모델을 사용해 보세요.');
  // json.model = 실제로 답한 모델 (자동 선택·예비 모델 전환 시 요청한 것과 다를 수 있음)
  return { text: content, model: json.model || model, requested: model };
}

function testOpenRouter(model) {
  const r = callOpenRouter_([{ role: 'user', content: '연결 테스트입니다. "연결 성공"이라고만 답하세요.' }], 30, model);
  return r.model + ' → ' + r.text.trim();
}

/** OpenRouter 전체 모델 목록 (10분 캐시). 가격은 100만 토큰당 달러로 환산 */
function listModels(noCache) {
  const ck = 'or_models_v1';
  if (!noCache) { const hit = cacheGet_(ck); if (hit) { try { return JSON.parse(hit); } catch (e) { /* 다시 받기 */ } } }
  const key = props_().getProperty('OPENROUTER_API_KEY');
  const res = UrlFetchApp.fetch('https://openrouter.ai/api/v1/models', {
    headers: key ? { Authorization: 'Bearer ' + key } : {}, muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) throw new Error('모델 목록을 받지 못했습니다 (' + res.getResponseCode() + ')');
  const data = (JSON.parse(res.getContentText()).data) || [];
  const perM = v => { const n = Number(v); return isNaN(n) ? null : Math.round(n * 1e6 * 100) / 100; };
  const out = data.map(m => ({
    id: m.id, name: m.name || m.id,
    pin: perM(m.pricing && m.pricing.prompt), pout: perM(m.pricing && m.pricing.completion),
    ctx: m.context_length || 0, created: m.created || 0
  })).filter(m => m.id);
  const json = JSON.stringify(out);
  cachePut_(ck, json);
  return out;
}

const SYSTEM_PROMPT = [
  '당신은 특수교육, 응용행동분석(ABA), 긍정적 행동지원(PBS)에 정통한 행동지원 전문가입니다.',
  '교사가 학교에서 직접 관찰·기록한 데이터를 근거로, 교사가 바로 활용할 수 있는 분석과 조언을 한국어로 작성합니다.',
  '원칙:',
  '- 반드시 제공된 데이터에 근거하세요. 데이터로 확인되지 않는 내용은 "추정" 또는 "가설"이라고 분명히 밝히세요.',
  '- 의학적·심리학적 진단을 내리지 마세요. 필요하면 전문가 연계를 권하세요.',
  '- 학생을 낙인찍는 표현을 피하고, 행동은 관찰 가능한 용어로 서술하세요.',
  '- 기록 수가 적으면(예: 10건 미만) 결론의 신뢰도가 낮다고 명시하세요.',
  '출력 형식(마크다운, 아래 제목을 그대로 사용):',
  '## 1. 데이터 요약',
  '## 2. 행동 패턴 분석  (언제·어디서·무엇 다음에 많이 일어나는지, 강도/지속시간의 추세와 변화)',
  '## 3. 행동 기능 가설  (관심 획득 / 회피·도피 / 물건·활동 획득 / 감각 자극 중 무엇으로 보이는지, 근거와 확신 정도)',
  '## 4. 중재 조언',
  '### 선행사건 중재(예방)',
  '### 대체행동 교수',
  '### 후속결과 중재',
  '### 위기 상황 대응(강도가 높을 때)',
  '## 5. 학부모 상담 팁  (상담 시작 문장 예시, 데이터를 보여주는 방법, 가정과 연계할 수 있는 일, 피해야 할 표현)',
  '## 6. 다음 관찰 제안  (데이터의 빈틈과 추가로 기록하면 좋은 것)',
  '전체 분량은 A4 2~3쪽 이내로, 구체적이고 실행 가능한 문장으로 작성하세요.'
].join('\n');

function buildSummary_(method, rows, student, rawLimit) {
  const others = getConfig_().students.map(s => s.name).filter(n => n && n !== student);
  const names = [{ n: student, to: '대상 학생' }].concat(others.map(n => ({ n: n, to: '다른 학생' })))
    .filter(x => x.n).sort((a, b) => b.n.length - a.n.length);
  const anon = t => { let s = String(t || ''); names.forEach(x => { s = s.split(x.n).join(x.to); }); return s; };

  const top = (fn, n, split) => {
    const m = {};
    rows.forEach(r => {
      const v = fn(r);
      (split ? String(v).split(',') : [v]).map(x => String(x).trim()).filter(Boolean)
        .forEach(x => { m[x] = (m[x] || 0) + 1; });
    });
    return Object.keys(m).sort((a, b) => m[b] - m[a]).slice(0, n || 10)
      .map(k => anon(k) + '(' + m[k] + ')').join(', ') || '없음';
  };
  const hour = r => { const h = parseInt(r.time || r.start || '', 10); return isNaN(h) ? '' : h + '시'; };
  const dates = rows.map(r => r.date).filter((d, i, a) => a.indexOf(d) === i);
  const L = [];
  L.push('기록법: ' + METHODS[method].label);
  L.push('기간: ' + rows[0].date + ' ~ ' + rows[rows.length - 1].date + ' (기록이 있는 날 ' + dates.length + '일, 총 ' + rows.length + '건)');

  if (method === 'abc') {
    L.push('행동별: ' + top(r => r.behavior));
    L.push('선행사건별: ' + top(r => r.antecedent, 12, true));
    L.push('후속결과별: ' + top(r => r.consequence, 12, true));
    L.push('장소/활동별: ' + top(r => r.place));
    L.push('시간대별: ' + top(hour, 24));
    L.push('요일별: ' + top(r => r.dow, 7));
    L.push('A→B→C 조합 상위: ' + top(r => (r.antecedent || '?') + ' → ' + r.behavior + ' → ' + (r.consequence || '?'), 8));
    const withI = rows.filter(r => r.intensity);
    if (withI.length) L.push('강도 기록: ' + [1, 2, 3].map(i => INT_LABEL[i] + ' ' + withI.filter(r => r.intensity === i).length).join(', '));
  } else if (method === 'freq') {
    L.push('행동별 빈도: ' + top(r => r.behavior));
    L.push('강도 분포: ' + [1, 2, 3].map(i => INT_LABEL[i] + ' ' + rows.filter(r => r.intensity === i).length).join(', '));
    const behs = rows.map(r => r.behavior).filter((b, i, a) => a.indexOf(b) === i);
    behs.forEach(b => {
      const rs = rows.filter(r => r.behavior === b);
      const avg = rs.reduce((s, r) => s + r.intensity, 0) / rs.length;
      L.push('  - ' + anon(b) + ': ' + [1, 2, 3].map(i => INT_LABEL[i] + ' ' + rs.filter(r => r.intensity === i).length).join(' ') + ', 평균강도 ' + avg.toFixed(2));
    });
    L.push('일별(건수/평균강도): ' + dates.map(d => {
      const rs = rows.filter(r => r.date === d);
      return d + '(' + rs.length + '건/' + (rs.reduce((s, r) => s + r.intensity, 0) / rs.length).toFixed(1) + ')';
    }).join(', '));
    L.push('시간대별: ' + top(hour, 24));
    L.push('요일별: ' + top(r => r.dow, 7));
    L.push('장소/활동별: ' + top(r => r.place));
  } else {
    const behs = rows.map(r => r.behavior).filter((b, i, a) => a.indexOf(b) === i);
    behs.forEach(b => {
      const rs = rows.filter(r => r.behavior === b);
      const tot = rs.reduce((s, r) => s + r.seconds, 0);
      const max = Math.max.apply(null, rs.map(r => r.seconds));
      L.push('  - ' + anon(b) + ': ' + rs.length + '회, 총 ' + fmtSec_(tot) + ', 평균 ' + fmtSec_(tot / rs.length) + ', 최장 ' + fmtSec_(max));
    });
    L.push('일별 총 지속시간: ' + dates.map(d => d + '(' + fmtSec_(rows.filter(r => r.date === d).reduce((s, r) => s + r.seconds, 0)) + ')').join(', '));
    L.push('시작 시간대별: ' + top(hour, 24));
    L.push('요일별: ' + top(r => r.dow, 7));
    L.push('장소/활동별: ' + top(r => r.place));
  }

  const recent = rows.slice(-(rawLimit || 150));
  const raw = recent.map(r => {
    if (method === 'abc') return [r.date + '(' + r.dow + ') ' + r.time, '장소:' + r.place, 'A:' + r.antecedent, 'B:' + r.behavior,
      '강도:' + (r.intensity ? INT_LABEL[r.intensity] : '-'), 'C:' + r.consequence, '메모:' + r.memo].map(anon).join(' | ');
    if (method === 'freq') return [r.date + '(' + r.dow + ') ' + r.time, r.behavior, '강도:' + INT_LABEL[r.intensity], '장소:' + r.place, '메모:' + r.memo].map(anon).join(' | ');
    return [r.date + '(' + r.dow + ') ' + r.start + '~' + r.end, r.behavior, '지속:' + fmtSec_(r.seconds), '장소:' + r.place, '메모:' + r.memo].map(anon).join(' | ');
  });
  return '[요약 통계]\n' + L.join('\n') + '\n\n[개별 기록 (최근 ' + recent.length + '건)]\n' + raw.join('\n');
}

function analyzeWithAI(method, student, days, behavior, mode, model) {
  const short = mode !== 'full';
  let rows = getRecords(method, student, days).rows;
  if (behavior) rows = rows.filter(r => r.behavior === behavior);
  if (!rows.length) throw new Error('선택한 기간에 분석할 기록이 없습니다.');
  const note = method === 'abc' ? '' :
    '\n\n참고: 이 데이터는 ' + METHODS[method].label + '이라 선행사건·후속결과 정보가 없습니다. 기능 가설은 시간대·장소 등 제한된 근거로만 세우고, ABC 기록을 함께 수집하도록 권해 주세요.';
  const lengthRule = short
    ? '\n\n[작성 분량] 간단 모드입니다. 제목 구성은 그대로 두되, 각 제목 아래에 핵심만 2~4개의 짧은 글머리표로 쓰고 전체를 A4 1쪽 이내로 작성하세요.'
    : '';
  const user = '아래는 교사가 학교에서 수집한 ' + METHODS[method].label + ' 데이터입니다. 학생 실명은 "대상 학생"으로 익명 처리되었습니다.' +
    (behavior ? ' 분석 대상 행동: ' + behavior + '.' : '') + note + lengthRule + '\n\n' + buildSummary_(method, rows, student, short ? 60 : 150) +
    '\n\n위 데이터를 근거로 지정된 형식에 따라 분석해 주세요.';
  const r = callOpenRouter_([{ role: 'system', content: SYSTEM_PROMPT }, { role: 'user', content: user }], short ? 1800 : 4000, model);
  let url = '';
  try { url = saveReport(method, student, days, r.text, r.model + (short ? ' · 간단' : ' · 자세히')).url; } catch (e) { /* 저장 실패해도 결과는 보여줌 */ }
  return { markdown: r.text, model: r.model, requested: r.requested, count: rows.length, url: url };
}

function saveReport(method, student, days, markdown, model) {
  const ss = getSpreadsheet_(method);
  let sh = ss.getSheetByName(REPORT_SHEET);
  if (!sh) {
    sh = ss.insertSheet(REPORT_SHEET);
    sh.getRange(1, 1, 1, 5).setValues([['분석일시', '학생', '분석 기간', '모델', '분석 내용']])
      .setFontWeight('bold').setBackground('#fff3bf');
    sh.setFrozenRows(1);
    sh.setColumnWidth(5, 700);
  }
  sh.appendRow([fmt_(new Date(), 'yyyy-MM-dd HH:mm'), student, Number(days) ? '최근 ' + days + '일' : '전체', model || '', markdown]);
  sh.getRange(sh.getLastRow(), 5).setWrap(true);
  return { ok: true, url: ss.getUrl() + '#gid=' + sh.getSheetId() };
}
