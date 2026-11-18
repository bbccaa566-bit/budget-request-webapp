// 예산 요구안 작성 웹앱 — 화면 동작
// 화면 출력은 textContent/DOM 생성만 사용 (innerHTML 미사용). 서버는 127.0.0.1 로컬 서버뿐.
'use strict';

const METHODS = ['계획 산출', '올해 집행액 기준', '계획 없음', '미산정'];
const CHECKS = [
  ['C1', '산출방식·필수 값'], ['C2', '미산정 사업에 금액 추측 없음'], ['C3', '확인 필요 표시'],
  ['C4', '근거 자료 표시'], ['C5', '원본 세부사업 누락 없음'], ['C6', '원본 금액 일치'], ['C7', '증감 사유 표시'],
];
const AUTOSAVE_KEY = 'budget-draft-autosave-v1';
const FORMAT = 'budget-draft/v2';
const FORMATS_OK = ['budget-draft/v1', FORMAT];
const newItem = (o = {}) => ({ 항목: '', 단가: null, 수량: null, 횟수: 1, 단위: '', ...o });

const $ = (id) => document.getElementById(id);
const el = (tag, props = {}, ...kids) => {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (k === 'class') e.className = v;
    else if (k === 'text') e.textContent = v;
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else if (v !== undefined && v !== null && v !== false) e.setAttribute(k, v === true ? '' : v);
  }
  for (const c of kids) if (c != null) e.append(c);
  return e;
};
const fmt = (n) => (n === null || n === undefined || n === '' || Number.isNaN(n)) ? '' : Number(n).toLocaleString('ko-KR');
const nowIso = () => {
  const d = new Date(), p = (x) => String(x).padStart(2, '0');
  const tz = -d.getTimezoneOffset(), s = tz >= 0 ? '+' : '-';
  return `${d.getFullYear()}-${p(d.getMonth() + 1)}-${p(d.getDate())}T${p(d.getHours())}:${p(d.getMinutes())}:${p(d.getSeconds())}${s}${p(Math.floor(Math.abs(tz) / 60))}:${p(Math.abs(tz) % 60)}`;
};
const ymd = () => nowIso().slice(0, 10).replace(/-/g, '');

// ───────── 상태 ─────────
const state = {
  draft: null,        // 초안 (서버로 보내고 .json으로 저장되는 것)
  xlsx: null,         // { name, buf }
  parsed: null,       // /api/parse 결과
  mapping: {}, headerRow: null,
  selected: null,     // 선택한 행 번호 (0부터)
  check: null,        // 마지막 /api/check 결과
  reached: 1,
};

const year = () => Number(state.draft?.year || $('yearInput').value);
const KA = () => `${year() - 1} 배정액`;
const KE = () => `${year() - 1} 집행액`;
const memoLabel = () => state.draft?.memo_name || (state.draft?.memo_text ? '사업계획 메모' : '');

// ───────── 단계 이동 ─────────
function go(step) {
  state.reached = Math.max(state.reached, step);
  for (let i = 1; i <= 4; i++) $('step' + i).classList.toggle('hidden', i !== step);
  document.querySelectorAll('.step').forEach((b) => {
    const s = Number(b.dataset.step);
    b.classList.toggle('active', s === step);
    b.disabled = s > state.reached || (s === 2 && !state.xlsx) || (s >= 3 && !state.draft?.rows);
  });
  $('summaryBar').classList.toggle('hidden', step < 3);
  if (step === 3) renderEdit();
  if (step === 4) renderStep4();
  window.scrollTo(0, 0);
}
document.querySelectorAll('.step').forEach((b) => b.addEventListener('click', () => go(Number(b.dataset.step))));

// ───────── ① 자료 올리기 ─────────
(function initStep1() {
  $('yearInput').value = new Date().getFullYear() + 1;
  for (let m = 1; m <= 12; m++) $('monthInput').append(el('option', { value: m, text: m === 12 ? '12월 (연간 확정)' : `${m}월 말 누계` }));
  $('fltMethod').append(...METHODS.map((m) => el('option', { value: m, text: m })));
  for (const id of ['yearInput', 'monthInput', 'unitInput']) $(id).addEventListener('change', updateStep1Ready);
  $('thrInput').addEventListener('input', () => {
    if (!state.draft) return;
    state.draft.reason_threshold = readThreshold();
    changed(false);
  });

  const drop = $('xlsxDrop');
  ['dragover', 'dragenter'].forEach((t) => drop.addEventListener(t, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((t) => drop.addEventListener(t, () => drop.classList.remove('over')));
  drop.addEventListener('drop', (e) => { e.preventDefault(); if (e.dataTransfer.files[0]) loadXlsx(e.dataTransfer.files[0]); });
  $('xlsxInput').addEventListener('change', (e) => e.target.files[0] && loadXlsx(e.target.files[0]));

  $('memoInput').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    if (!f) return;
    if (f.size > 1024 * 1024) { $('step1Err').textContent = '메모 파일은 1MB 이하만 올릴 수 있습니다.'; return; }
    $('memoText').value = await f.text();
    $('memoText').dataset.name = f.name;
    $('memoLabel').textContent = `✔ ${f.name}`;
    $('memoInput').closest('.drop').classList.add('done');
  });
  $('memoText').addEventListener('input', () => { if (!$('memoInput').files[0]) delete $('memoText').dataset.name; });
  $('toStep2').addEventListener('click', () => { renderStep2(); go(2); });
  $('draftInput').addEventListener('change', (e) => e.target.files[0] && loadDraftFile(e.target.files[0]));
  showResume();
  setTimeout(renderPhrases, 0);
})();

async function loadXlsx(file) {
  $('step1Err').textContent = '';
  if (!/\.xlsx$/i.test(file.name)) { $('step1Err').textContent = '.xlsx 파일만 올릴 수 있습니다.'; return; }
  if (file.size > 10 * 1024 * 1024) { $('step1Err').textContent = '엑셀 파일은 10MB 이하만 올릴 수 있습니다.'; return; }
  state.xlsx = { name: file.name, buf: await file.arrayBuffer() };
  state.mapping = {}; state.headerRow = null;
  const res = await apiParse();
  if (!res) return;
  const ok = res.rows.length > 0;
  $('xlsxLabel').textContent = ok ? `✔ ${file.name} — 세부사업 ${res.rows.length}개 인식`
    : `⚠ ${file.name} — 세부사업을 찾지 못했습니다. 다음 단계에서 머리글 행·열을 지정하거나 다른 파일을 올리세요.`;
  $('xlsxDrop').classList.toggle('done', ok);
  $('step1Err').textContent = res.errors.length ? res.errors.join('\n') : '';
  updateStep1Ready();
}

function readThreshold() {
  const v = parseNum($('thrInput').value);
  return (v === null || Number.isNaN(v) || v < 0) ? null : v;
}

function updateStep1Ready() {
  $('toStep2').disabled = !(state.xlsx && state.parsed && $('yearInput').value && $('monthInput').value && $('unitInput').value);
}

async function apiParse(buf = state.xlsx.buf, headerRow = state.headerRow, mapping = state.mapping) {
  const q = new URLSearchParams();
  if (headerRow) q.set('header_row', headerRow);
  for (const [k, v] of Object.entries(mapping || {})) if (v) q.set('map_' + k, v);
  try {
    const r = await fetch('/api/parse?' + q, { method: 'POST', body: buf, headers: { 'Content-Type': 'application/octet-stream' } });
    const j = await r.json();
    if (!r.ok) throw new Error(j.error || '엑셀을 읽지 못했습니다.');
    if (buf === state.xlsx?.buf) { state.parsed = j; state.headerRow = j.header_row; state.mapping = j.mapping; }
    return j;
  } catch (err) {
    $('step1Err').textContent = err.message.includes('fetch') ? '서버에 연결할 수 없습니다. 실행 창이 켜져 있는지 확인하세요.' : err.message;
    return null;
  }
}

// ───────── ② 열 맞추기 ─────────
function renderStep2() {
  const p = state.parsed;
  const hr = $('headerRow');
  hr.replaceChildren(...Array.from({ length: 10 }, (_, i) => el('option', { value: i + 1, text: `${i + 1}행` })));
  hr.value = p.header_row;
  document.querySelectorAll('.mapSel').forEach((s) => {
    s.replaceChildren(el('option', { value: '', text: '— 선택 —' }),
      ...p.columns.map((c, i) => el('option', { value: i + 1, text: `${colLetter(i + 1)}열: ${c || '(빈 머리글)'}` })));
    s.value = p.mapping[s.dataset.key] || '';
  });
  const tbl = $('previewTbl');
  tbl.replaceChildren(...p.preview.map((row) => el('tr', {}, ...row.map((c) => el('td', { text: c })))));
  $('parseErrors').replaceChildren(...p.errors.map((e) => el('li', { text: e })));
  $('parseOk').textContent = p.errors.length ? '' : `세부사업 ${p.rows.length}개를 읽었습니다. 배정액 합계 ${fmt(p.rows.reduce((a, r) => a + r.배정액, 0))}`;
  $('toStep3').disabled = p.errors.length > 0 || p.rows.length === 0;
}
const colLetter = (n) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };

$('headerRow').addEventListener('change', async (e) => {
  state.headerRow = Number(e.target.value); state.mapping = {};
  if (await apiParse()) renderStep2();
});
document.querySelectorAll('.mapSel').forEach((s) => s.addEventListener('change', async () => {
  const m = {};
  document.querySelectorAll('.mapSel').forEach((x) => { if (x.value) m[x.dataset.key] = Number(x.value); });
  state.mapping = m;
  if (await apiParse()) renderStep2();
}));
$('back1').addEventListener('click', () => go(1));
$('toStep3').addEventListener('click', () => {
  if (state.draft?.rows?.length && !confirm('표를 새로 만들면 지금까지 편집한 내용이 사라집니다. 계속할까요?')) return;
  buildDraft();
  go(3);
});

function buildDraft() {
  const p = state.parsed, y = Number($('yearInput').value), m = Number($('monthInput').value);
  const memo = $('memoText').value.trim();
  const memoName = memo ? ($('memoText').dataset.name || '사업계획 메모(붙여넣기)') : '';
  state.draft = {
    format: FORMAT,
    title: $('titleInput').value.trim() || `${y}년 예산 요구안 초안`,
    year: y, unit: $('unitInput').value, base_month: m,
    base_note: m < 12 ? `올해 집행액은 ${y - 1}년 ${m}월 말 누계` : `올해 집행액은 ${y - 1}년 연간 확정치`,
    sources: [state.xlsx.name, ...(memoName ? [memoName] : [])],
    source_name: state.xlsx.name, memo_name: memoName, memo_text: memo,
    source_hash: p.source_hash, header_row: state.headerRow, mapping: state.mapping,
    original: p.original, source_verified: true, reason_threshold: readThreshold(),
    practice: !!state.practice,
    started_at: nowIso(), exported_at: null, rows: [],
  };
  const ka = KA(), ke = KE();
  state.draft.rows = p.rows.map((o) => {
    const r = { 비목: o.비목, 세부사업: o.세부사업, 산출방식: '', 단가: null, 수량: null, 단위: '', 산출: [],
      [ka]: o.배정액, [ke]: o.집행액, '산출 근거': '', '요구 사유': '', '근거 자료': '', '확인 필요': '', _new: false, _auto: '' };
    setMethod(r, '올해 집행액 기준');
    return r;
  });
  state.selected = null;
  scheduleCheck(0);
  autosave();
}

// ───────── 산출방식 규칙 (PRD F5·F6) ─────────
function autoNote(r) {
  const d = state.draft, parts = [];
  const memoCited = memoLabel() && String(r['근거 자료'] || '').includes(memoLabel());
  if (r.산출방식 === '올해 집행액 기준') {
    if (Number(d.base_month) < 12) parts.push(`연간 집행액 아님(${d.base_month}월 말 누계) — 연말 추정치 확인 필요`);
    if (!memoCited) parts.push('계획 메모에 언급 없음 — 내년 계속 여부 확인');
    if (!parts.length) parts.push('계획 메모에 단가·수량 없음 — 올해 집행액 적용이 맞는지 확인');
  } else if (r.산출방식 === '계획 없음') {
    parts.push('요구액 0 — 예비 수요가 없는지 확인');
  } else if (r.산출방식 === '미산정') {
    parts.push('금액 미정 — 견적·산출 근거 필요(요구액 합계에 미포함)');
  }
  if ((r.비목 || '미분류') === '미분류' && r.산출방식) parts.push('비목 지정 필요');
  return parts.join('; ');
}
function refreshAuto(r, force = false) {
  const cur = String(r['확인 필요'] || '');
  const next = autoNote(r);
  if (force || !cur.trim() || cur === r._auto) r['확인 필요'] = next;
  r._auto = next;
}
function setMethod(r, m) {
  const ke = KE();
  // 계획 산출의 산출 줄은 다른 방식으로 바꿔도 보관했다가 되돌리면 복원
  if (r.산출방식 === '계획 산출' && m !== '계획 산출' && r.산출?.length) { r._keptItems = r.산출; }
  r.산출 = [];
  r.산출방식 = m;
  if (m === '계획 산출') {
    r.산출 = r._keptItems?.length ? r._keptItems : [newItem()];
    delete r._keptItems;
    r.단가 = null; r.수량 = null; r.단위 = '';
  } else if (m === '올해 집행액 기준') { r.단가 = r[ke] ?? 0; r.수량 = 1; r.단위 = '식'; }
  else if (m === '계획 없음') { r.단가 = 0; r.수량 = 1; r.단위 = '식'; }
  else if (m === '미산정') { r.단가 = null; r.수량 = null; r.단위 = ''; }
  refreshTexts(r);
}

// ───────── 자동 문구 (요구 사유·산출 근거·근거 자료) ─────────
// 규칙 기반(AI 미사용). 칸이 비었거나 마지막 자동 문구 그대로일 때만 채운다 → 직접 쓴 문구는 바꾸지 않음.
const AUTO_KEYS = { '산출 근거': '_autoBasis', '요구 사유': '_autoReason', '근거 자료': '_autoSource' };
function memoLineFor(name) {
  const text = state.draft?.memo_text || '';
  if (!text || !name) return '';
  const lines = text.split(/\r?\n/).map((l) => l.trim()).filter((l) => l && !l.startsWith('#'));
  let hit = lines.find((l) => l.includes(name));
  if (!hit) {
    const words = name.split(/[\s·,()]+/).filter((w) => w.length >= 2);
    if (words.length) hit = lines.find((l) => words.every((w) => l.includes(w)));
  }
  return hit ? hit.replace(/^[-*•·○□\d.)\s]+/, '').slice(0, 80) : '';
}
function autoTexts(r) {
  const d = state.draft, y1 = year() - 1, ka = KA(), ke = KE();
  const req = reqAmount(r), alloc = Number(r[ka] || 0), line = memoLineFor(r.세부사업);
  let basis = '';
  if (r.산출방식 === '계획 산출') {
    const items = itemsOf(r);
    if (items.length) {
      // 입력이 끝난 줄만 문구로 (빈 줄은 건너뜀)
      const done = items.filter((it) => itemAmount(it) !== null);
      basis = done.map(itemText).join('; ') + (req !== null && items.length > 1 ? ` (합계 ${fmt(req)})` : '');
    } else if (req !== null) basis = `${fmt(r.단가)} × ${fmt(r.수량)}${r.단위 || ''} = ${fmt(req)}`;
  } else if (r.산출방식 === '올해 집행액 기준') {
    basis = `${y1}년 집행액 ${fmt(r[ke] ?? 0)} 적용` + (Number(d.base_month) < 12 ? `(${d.base_month}월 말 누계)` : '');
  } else if (r.산출방식 === '계획 없음') basis = '내년 계획 없음';
  else if (r.산출방식 === '미산정') basis = '금액 미정';

  let reason;
  if (r.산출방식 === '미산정' || req === null) reason = '금액 미정 — 견적·산출 근거 확보 후 반영';
  else if (r.산출방식 === '계획 없음') reason = alloc ? `내년 사업 계획 없음 — 전년 배정 ${fmt(alloc)} 전액 감액` : '내년 사업 계획 없음';
  else if (!alloc) reason = `신규 사업 — ${fmt(req)} 요구`;
  else {
    const diff = req - alloc, pct = ((diff / alloc) * 100).toFixed(1);
    reason = diff > 0 ? `전년 배정 대비 ${fmt(diff)} 증액(+${pct}%)` : diff < 0 ? `전년 배정 대비 ${fmt(-diff)} 감액(${pct}%)` : '전년 배정과 동일';
  }
  if (line) reason += ` — 계획: ${line}`;

  const parts = [];
  if (!r._new && d.source_name) parts.push(`${d.source_name} > ${r.세부사업}`);
  if (line && memoLabel()) parts.push(`${memoLabel()} > ${line}`);
  return { '산출 근거': basis, '요구 사유': reason, '근거 자료': parts.join('; ') };
}
function refreshTexts(r, force = null) {
  const t = autoTexts(r);
  for (const [k, ak] of Object.entries(AUTO_KEYS)) {
    const cur = String(r[k] ?? '');
    if (force === k || force === true || !cur.trim() || cur === r[ak]) r[k] = t[k];
    r[ak] = t[k];
  }
  refreshAuto(r);
}
const isAuto = (r, k) => String(r[k] ?? '') !== '' && r[k] === r[AUTO_KEYS[k]];

// 자주 쓰는 문구
const PHRASE_KEY = 'budget-phrases-v1';
const DEFAULT_PHRASES = ['계약 단가 인상 반영', '사업 범위 확대', '물가 상승분 반영', '전년 집행 실적 반영 감액',
  '사업 종료에 따른 감액', '신규 수요 반영', '법정·의무 지출', '일회성 경비 제외', '인원 증가 반영', '견적 확보 후 확정 예정'];
function myPhrases() { try { return JSON.parse(localStorage.getItem(PHRASE_KEY) || '[]'); } catch { return []; } }
function saveMyPhrases(list) { try { localStorage.setItem(PHRASE_KEY, JSON.stringify(list)); } catch { /* 저장 불가 */ } }
function renderPhrases() {
  const mine = myPhrases();
  const chip = (p, own) => {
    const c = el('span', { class: 'chip' + (own ? ' mine' : ''), role: 'button', tabindex: 0, title: '요구 사유에 넣기', text: p });
    c.addEventListener('click', (e) => {
      if (e.target.classList.contains('x')) return;
      const r = state.draft.rows[state.selected]; if (!r) return;
      const cur = String(r['요구 사유'] || '').trim();
      r['요구 사유'] = cur ? `${cur}; ${p}` : p;
      $('dReason').value = r['요구 사유']; updateTags(r); changed(false);
    });
    if (own) c.append(el('span', { class: 'x', text: '×', title: '내 문구 삭제', onclick: () => { saveMyPhrases(myPhrases().filter((x) => x !== p)); renderPhrases(); } }));
    return c;
  };
  $('phraseBox').replaceChildren(...DEFAULT_PHRASES.map((p) => chip(p, false)), ...mine.map((p) => chip(p, true)));
}
$('phraseAdd').addEventListener('click', () => {
  const v = $('phraseInput').value.trim();
  if (!v) return;
  const mine = myPhrases();
  if (!mine.includes(v) && !DEFAULT_PHRASES.includes(v)) { mine.push(v); saveMyPhrases(mine); }
  $('phraseInput').value = ''; renderPhrases();
});
$('phraseInput').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); $('phraseAdd').click(); } });
document.querySelectorAll('button[data-auto]').forEach((b) => b.addEventListener('click', () => {
  const r = state.draft.rows[state.selected]; if (!r) return;
  refreshTexts(r, b.dataset.auto);
  renderDetail(); updateRowText(r); changed(false);
}));
function updateTags(r) {
  $('tagBasis').classList.toggle('hidden', !isAuto(r, '산출 근거'));
  $('tagReason').classList.toggle('hidden', !isAuto(r, '요구 사유'));
  $('tagSource').classList.toggle('hidden', !isAuto(r, '근거 자료'));
}
$('fillAll').addEventListener('click', () => {
  let n = 0;
  state.draft.rows.forEach((r) => {
    const before = JSON.stringify([r['산출 근거'], r['요구 사유'], r['근거 자료']]);
    refreshTexts(r);
    if (JSON.stringify([r['산출 근거'], r['요구 사유'], r['근거 자료']]) !== before) n++;
  });
  changed(true);
  $('fillAll').textContent = `빈 칸 자동 문구 채우기 (${n}개 행 갱신)`;
  setTimeout(() => { $('fillAll').textContent = '빈 칸 자동 문구 채우기'; }, 2500);
});
// 계산 (서버 budget_core.request_amount 와 같은 규칙): 계획 산출은 산출 줄 합계, 줄마다 단가 × 수량 × 횟수(비면 1)
const blank = (v) => v === null || v === undefined || v === '';
const itemsOf = (r) => (r.산출방식 === '계획 산출' && Array.isArray(r.산출) && r.산출.length) ? r.산출 : [];
const itemAmount = (it) => (blank(it.단가) || blank(it.수량)) ? null : Number(it.단가) * Number(it.수량) * (blank(it.횟수) ? 1 : Number(it.횟수));
const itemText = (it) => {
  const a = itemAmount(it);
  const parts = [fmt(it.단가), fmt(it.수량) + (it.단위 || '')];
  if (!blank(it.횟수) && Number(it.횟수) !== 1) parts.push(fmt(it.횟수) + '회');
  return `${it.항목 ? it.항목 + ' ' : ''}${parts.join(' × ')}${a === null ? '' : ' = ' + fmt(a)}`;
};
const reqAmount = (r) => {
  const items = itemsOf(r);
  if (items.length) { const a = items.map(itemAmount); return a.some((x) => x === null) ? null : a.reduce((s, x) => s + x, 0); }
  return (blank(r.단가) || blank(r.수량)) ? null : Number(r.단가) * Number(r.수량);
};
const parseNum = (s) => { const t = String(s).replace(/,/g, '').trim(); if (!t) return null; const n = Number(t); return Number.isFinite(n) ? n : NaN; };

// ───────── ③ 산출 편집 ─────────
function renderEdit() {
  renderTable();
  renderDetail();
  renderMemo();
  renderLive();
}
function renderTable() {
  const d = state.draft, ka = KA(), ke = KE();
  const probRows = new Set((state.check?.problems || []).map((p) => p.row).filter(Boolean));
  const onlyCheck = $('fltCheck').checked, fm = $('fltMethod').value;
  const head = el('tr', {}, ...['번호', '비목', '세부사업', '산출방식', '단가', '수량', '횟수', '단위', `${d.year} 요구액`, `${d.year - 1} 배정액`, `${d.year - 1} 집행액`, '확인']
    .map((h, i) => el('th', { class: [4, 5, 6, 8, 9, 10].includes(i) ? 'num' : '', text: h })));
  const trs = [head];
  d.rows.forEach((r, i) => {
    if (onlyCheck && !String(r['확인 필요'] || '').trim()) return;
    if (fm && r.산출방식 !== fm) return;
    const mi = METHODS.indexOf(r.산출방식);
    const items = itemsOf(r);
    const plan = r.산출방식 === '계획 산출';
    // 계획 산출이 한 줄이면 표에서 바로 입력, 여러 줄이면 아래 상세에서 입력
    const target = plan ? items[0] : r;
    const sel = el('select', { 'aria-label': '산출방식', onchange: (e) => {
      setMethod(r, e.target.value);
      if (e.target.value === '계획 산출') state.selected = i;
      changed(true);
    } }, ...METHODS.map((m) => el('option', { value: m, text: m })));
    sel.value = r.산출방식;
    const numInput = (key) => {
      const editable = plan && key !== null;
      const val = key === '횟수' && !plan ? '' : (target?.[key] ?? '');
      const inp = el('input', { class: 'n', type: 'text', inputmode: 'decimal', value: val, disabled: !editable, 'aria-label': key,
        placeholder: key === '횟수' && plan ? '1' : '' });
      inp.addEventListener('input', () => {
        const v = parseNum(inp.value);
        target[key] = Number.isNaN(v) ? null : v;
        inp.classList.toggle('bad', Number.isNaN(v));
        tr.querySelector('.reqamt').textContent = fmt(reqAmount(r));
        textsChanged(r, i);
        if (state.selected === i) renderItems();
        changed(false);
      });
      inp.addEventListener('blur', () => { if (!Number.isNaN(parseNum(inp.value))) inp.value = target[key] ?? ''; });
      return inp;
    };
    let calcCells;
    if (plan && items.length > 1) {
      calcCells = [el('td', { colspan: 4, class: 'multi', text: `산출 ${items.length}줄 — 행을 누르면 아래에서 편집` })];
    } else {
      const unit = el('input', { class: 'u', type: 'text', value: target?.단위 || '', disabled: !plan, maxlength: 10, 'aria-label': '단위' });
      unit.addEventListener('input', () => { target.단위 = unit.value; textsChanged(r, i); if (state.selected === i) renderItems(); changed(false); });
      calcCells = [el('td', {}, numInput('단가')), el('td', {}, numInput('수량')), el('td', {}, numInput('횟수')), el('td', {}, unit)];
    }
    const nameCell = el('td', { class: r._new ? '' : 'lock' }, r.세부사업 || '(이름 없음)', r._new ? el('span', { class: 'newtag', text: '신규' }) : null);
    const tr = el('tr', { class: [i === state.selected ? 'selected' : '', probRows.has(i + 1) ? 'hasprob' : ''].join(' ') },
      el('td', { class: 'num', text: i + 1 }),
      el('td', { class: r._new ? '' : 'lock', text: r.비목 }),
      nameCell,
      el('td', { class: `m-${mi}` }, sel),
      ...calcCells,
      el('td', { class: 'num reqamt', text: fmt(reqAmount(r)) }),
      el('td', { class: 'num lock', text: fmt(r[ka]) }), el('td', { class: 'num lock', text: fmt(r[ke]) }),
      el('td', { class: 'flag', title: r['확인 필요'] || '', text: String(r['확인 필요'] || '').trim() ? '⚠' : '' }));
    tr.addEventListener('click', () => {
      if (state.selected === i) return;
      $('editTbl').querySelectorAll('tr.selected').forEach((x) => x.classList.remove('selected'));
      tr.classList.add('selected');
      state.selected = i;
      renderDetail();
    });
    trs.push(tr);
  });
  if (trs.length === 1) trs.push(el('tr', {}, el('td', { colspan: 12, text: '조건에 맞는 행이 없습니다.' })));
  $('editTbl').replaceChildren(...trs);
}

// 상세: 산출 줄 편집 (계획 산출)
function renderItems() {
  const r = state.draft.rows[state.selected];
  const box = $('itemsBox');
  if (!r || r.산출방식 !== '계획 산출') { box.classList.add('hidden'); return; }
  box.classList.remove('hidden');
  const head = el('tr', {}, ...['항목', '단가', '수량', '횟수', '단위', '금액', ''].map((h, k) => el('th', { class: [1, 2, 3, 5].includes(k) ? 'num' : '', text: h })));
  const rows = r.산출.map((it, k) => {
    const amt = el('td', { class: 'num', text: fmt(itemAmount(it)) });
    const inp = (key, num, attrs = {}) => {
      const x = el('input', { type: 'text', class: num ? 'n' : '', value: it[key] ?? '', 'aria-label': `${k + 1}번째 줄 ${key}`, ...attrs });
      if (num) x.setAttribute('inputmode', 'decimal');
      x.addEventListener('input', () => {
        if (num) { const v = parseNum(x.value); it[key] = Number.isNaN(v) ? null : v; x.classList.toggle('bad', Number.isNaN(v)); }
        else it[key] = x.value;
        amt.textContent = fmt(itemAmount(it));
        $('itemsTotal').textContent = `합계 ${fmt(reqAmount(r))}`;
        textsChanged(r, state.selected);
        renderTable();
        changed(false);
      });
      return x;
    };
    const del = el('button', { class: 'mini', text: '삭제', title: '이 산출 줄 삭제', disabled: r.산출.length === 1 });
    del.addEventListener('click', () => { r.산출.splice(k, 1); refreshTexts(r); changed(true); });
    return el('tr', {}, el('td', {}, inp('항목', false, { maxlength: 40, placeholder: '예: 강사료' })),
      el('td', {}, inp('단가', true)), el('td', {}, inp('수량', true)), el('td', {}, inp('횟수', true, { placeholder: '1' })),
      el('td', {}, inp('단위', false, { maxlength: 10, placeholder: '명·대' })), amt, el('td', {}, del));
  });
  $('itemsTbl').replaceChildren(head, ...rows);
  $('itemsTotal').textContent = `합계 ${fmt(reqAmount(r))}`;
}
$('itemAdd').addEventListener('click', () => {
  const r = state.draft.rows[state.selected]; if (!r) return;
  r.산출.push(newItem());
  refreshTexts(r);
  changed(true);
  const ins = $('itemsTbl').querySelectorAll('tr:last-child input');
  if (ins[0]) ins[0].focus();
});
// 숫자가 바뀌면 자동 문구(직접 쓴 것 제외)를 다시 만들고, 그 행이 선택돼 있으면 상세 칸도 갱신
function textsChanged(r, i) {
  refreshTexts(r);
  if (state.selected === i) {
    for (const [id, k] of [['dBasis', '산출 근거'], ['dReason', '요구 사유'], ['dSource', '근거 자료'], ['dCheck', '확인 필요']]) {
      if (document.activeElement !== $(id)) $(id).value = r[k] || '';
    }
    updateTags(r);
  }
  const tr = $('editTbl').querySelectorAll('tr')[[...$('editTbl').querySelectorAll('tr')].findIndex((x) => Number(x.firstChild?.textContent) === i + 1)];
  if (tr) tr.lastElementChild.textContent = String(r['확인 필요'] || '').trim() ? '⚠' : '';
}

function renderDetail() {
  const box = $('detail');
  const r = state.draft.rows[state.selected];
  if (!r) { box.classList.add('hidden'); return; }
  box.classList.remove('hidden');
  box.classList.toggle('isnew', !!r._new);
  $('detailTitle').textContent = `${state.selected + 1}. ${r.세부사업 || '(이름 없음)'} — ${r.산출방식}`;
  $('dCat').value = r.비목 || ''; $('dName').value = r.세부사업 || '';
  $('dBasis').value = r['산출 근거'] || ''; $('dReason').value = r['요구 사유'] || '';
  $('dSource').value = r['근거 자료'] || ''; $('dCheck').value = r['확인 필요'] || '';
  updateTags(r);
  renderItems();
  renderReasonHint(r);
}
// 요구 사유 칸 아래: 이 사업이 증감 사유 점검(C7) 대상인지 알려 줌
function renderReasonHint(r) {
  const st = reasonStatus(r);
  $('reasonHint').textContent = st === '사유 없음' ? '⚠ 증감 기준을 넘는 사업입니다. 요구 사유를 적어야 내보낼 수 있습니다.'
    : st === '보강 필요' ? '⚠ 증감 기준을 넘는 사업입니다. 자동 문구 그대로이니 구체적 사유를 덧붙여 주세요(경고).'
      : '';
}
function reasonStatus(r) {
  const n = state.check?.summary?.by_project?.[state.draft.rows.indexOf(r)];
  return n && n.세부사업 === r.세부사업 ? n.사유점검 : '';
}
const bindDetail = (id, key, opts = {}) => $(id).addEventListener('input', () => {
  const r = state.draft.rows[state.selected];
  if (!r) return;
  r[key] = $(id).value;
  if (opts.texts) {
    refreshTexts(r);
    for (const [fid, k] of [['dBasis', '산출 근거'], ['dReason', '요구 사유'], ['dSource', '근거 자료'], ['dCheck', '확인 필요']]) $(fid).value = r[k] || '';
  } else if (opts.auto) { refreshAuto(r); $('dCheck').value = r['확인 필요']; }
  if (opts.row) updateRowText(r);
  updateTags(r);
  changed(false);
});
bindDetail('dCat', '비목', { auto: true, row: true });
bindDetail('dName', '세부사업', { row: true, texts: true });
bindDetail('dBasis', '산출 근거');
bindDetail('dReason', '요구 사유');
bindDetail('dSource', '근거 자료', { auto: true });
bindDetail('dCheck', '확인 필요', { row: true });
function updateRowText(r) {
  const tr = $('editTbl').querySelector('tr.selected');
  if (!tr) return;
  tr.children[1].textContent = r.비목;
  tr.children[2].firstChild.textContent = r.세부사업 || '(이름 없음)';
  tr.lastElementChild.textContent = String(r['확인 필요'] || '').trim() ? '⚠' : '';
  $('detailTitle').textContent = `${state.selected + 1}. ${r.세부사업 || '(이름 없음)'} — ${r.산출방식}`;
}
$('dAuto').addEventListener('click', () => {
  const r = state.draft.rows[state.selected]; if (!r) return;
  refreshAuto(r, true); $('dCheck').value = r['확인 필요']; updateRowText(r); changed(false);
});
$('dMemoRef').addEventListener('click', () => {
  const r = state.draft.rows[state.selected]; if (!r) return;
  if (!memoLabel()) { alert('올린 사업계획 메모가 없습니다.'); return; }
  const sel = window.getSelection();
  let part = '';
  if (sel && sel.rangeCount && $('memoView').contains(sel.anchorNode)) part = sel.toString().trim().split('\n')[0].slice(0, 40);
  const ref = `${memoLabel()} > ${part || '(절 제목 입력)'}`;
  r['근거 자료'] = r['근거 자료'] ? `${r['근거 자료']}; ${ref}` : ref;
  refreshAuto(r);
  renderDetail(); updateRowText(r); changed(false);
});
$('dDelete').addEventListener('click', () => {
  const r = state.draft.rows[state.selected];
  if (!r || !r._new) return;
  if (!confirm(`신규 사업 '${r.세부사업}'을 삭제할까요?`)) return;
  state.draft.rows.splice(state.selected, 1);
  state.selected = null; changed(true);
});
$('addRow').addEventListener('click', () => {
  const d = state.draft, n = d.rows.filter((r) => r._new).length + 1;
  const r = { 비목: '미분류', 세부사업: `신규 사업 ${n}`, 산출방식: '', 단가: null, 수량: null, 단위: '',
    [KA()]: 0, [KE()]: 0, '산출 근거': '', '요구 사유': '', '근거 자료': '', '확인 필요': '', _new: true, _auto: '' };
  setMethod(r, '미산정');
  d.rows.push(r);
  state.selected = d.rows.length - 1;
  $('fltCheck').checked = false; $('fltMethod').value = '';
  changed(true);
  $('dName').focus(); $('dName').select();
});
$('fltCheck').addEventListener('change', renderEdit);
$('fltMethod').addEventListener('change', renderEdit);
$('back2').addEventListener('click', () => go(state.xlsx ? 2 : 1));
$('toStep4').addEventListener('click', () => go(4));

// 메모 패널 + 검색
function renderMemo() {
  const text = state.draft.memo_text || '';
  const q = $('memoSearch').value.trim();
  const view = $('memoView');
  if (!text) { view.textContent = '올린 메모가 없습니다.'; return; }
  if (!q) { view.textContent = text; return; }
  const nodes = []; const low = text.toLowerCase(), ql = q.toLowerCase();
  let i = 0, j;
  while ((j = low.indexOf(ql, i)) !== -1) { nodes.push(document.createTextNode(text.slice(i, j)), el('mark', { text: text.slice(j, j + q.length) })); i = j + q.length; }
  nodes.push(document.createTextNode(text.slice(i)));
  view.replaceChildren(...nodes);
  const first = view.querySelector('mark'); if (first) first.scrollIntoView({ block: 'nearest' });
}
$('memoSearch').addEventListener('input', renderMemo);

// ───────── 변경 → 검사·자동 저장 ─────────
function changed(rerender) {
  if (rerender) renderEdit();
  scheduleCheck();
  autosave();
}
let checkTimer = null, checkSeq = 0;
function scheduleCheck(delay = 250) {
  clearTimeout(checkTimer);
  checkTimer = setTimeout(runCheck, delay);
}
async function runCheck() {
  const seq = ++checkSeq;
  try {
    const r = await fetch('/api/check', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(state.draft) });
    const j = await r.json();
    if (seq !== checkSeq) return;
    if (!r.ok) throw new Error(j.error);
    state.check = j;
  } catch (err) {
    state.check = { ok: false, problems: [{ code: '-', row: null, msg: '검사 서버에 연결할 수 없습니다. 실행 창이 켜져 있는지 확인하세요.' }], warnings: [], summary: null };
  }
  renderSummaryBar(); renderLive();
  document.querySelectorAll('#editTbl tr').forEach((tr, k) => {
    const n = Number(tr.firstChild?.textContent);
    if (k > 0 && n) tr.classList.toggle('hasprob', state.check.problems.some((p) => p.row === n));
  });
  const sr = state.draft?.rows?.[state.selected]; if (sr && !$('detail').classList.contains('hidden')) renderReasonHint(sr);
  if (!$('step4').classList.contains('hidden')) renderStep4();
}
function renderSummaryBar() {
  const s = state.check?.summary, d = state.draft;
  $('sbAllocLbl').textContent = `${d.year - 1} 배정`;
  $('sbReqLbl').textContent = `${d.year} 요구(초안)`;
  $('sbUnit').textContent = `단위: ${d.unit}`;
  if (!s) return;
  $('sbAlloc').textContent = fmt(s.배정);
  $('sbReq').textContent = fmt(s.요구);
  $('sbDiff').textContent = `${s.증감 > 0 ? '+' : ''}${fmt(s.증감)}${s.증감률 === null ? '' : ` (${(s.증감률 * 100).toFixed(1)}%)`}`;
  $('sbDiff').classList.toggle('neg', s.증감 < 0);
  $('sbUndet').textContent = s.미산정;
  $('sbCheck').textContent = s.확인필요;
  const failed = new Set(state.check.problems.map((p) => p.code));
  $('sbTests').textContent = `${CHECKS.filter(([c]) => !failed.has(c)).length}/${CHECKS.length} ${state.check.ok ? '✔' : '✖'}`;
  $('sbTests').classList.toggle('neg', !state.check.ok);
}
function problemItem(p) {
  return el('li', { class: p.row ? 'link' : '', text: p.msg, onclick: () => { if (p.row) { state.selected = p.row - 1; $('fltCheck').checked = false; $('fltMethod').value = ''; go(3); } } });
}
function renderLive() {
  const ps = state.check?.problems || [];
  $('liveProblems').replaceChildren(...ps.map(problemItem));
}

// ───────── ④ 검사·내보내기 ─────────
function renderStep4() {
  const c = state.check;
  if (!c) { scheduleCheck(0); return; }
  $('checkList').replaceChildren(...CHECKS.map(([code, name]) => {
    const ps = c.problems.filter((p) => p.code === code);
    return el('li', { class: ps.length ? 'fail' : 'pass' }, el('span', { class: 'st', text: ps.length ? '✖' : '✔' }), `${code} ${name}`,
      ps.length ? el('ul', {}, ...ps.map(problemItem)) : null);
  }), ...c.problems.filter((p) => p.code === '-').map((p) => el('li', { class: 'fail' }, el('span', { class: 'st', text: '✖' }), p.msg)));
  $('warnList').replaceChildren(...(c.warnings || []).map((w) => el('li', { text: '⚠ ' + w })),
    ...(c.reason_notes || []).map((p) => { const li = problemItem(p); li.classList.add('sub'); return li; }));
  const s = c.summary, d = state.draft;
  const thr = d.reason_threshold;
  $('thrNote').textContent = `증감 사유 기준: ${blank(thr) || thr === null ? '증감이 있는 모든 사업' : `전년 배정 대비 ±${thr}% 이상`} (①단계에서 바꿀 수 있음)`;
  if (s?.by_project) {
    const head = el('tr', {}, ...['번호', '비목', '세부사업', '산출방식', `${d.year - 1} 배정액`, `${d.year} 요구액(초안)`, '증감', '증감률', '사유 점검']
      .map((h, i) => el('th', { class: [4, 5, 6, 7].includes(i) ? 'num' : '', text: h })));
    const body = s.by_project.map((x) => el('tr', { class: 'link', title: '눌러서 이 사업 편집', onclick: () => { state.selected = x.번호 - 1; $('fltCheck').checked = false; $('fltMethod').value = ''; go(3); } },
      el('td', { class: 'num', text: x.번호 }), el('td', { text: x.비목 }), el('td', { text: x.세부사업 }),
      el('td', { class: `m-${METHODS.indexOf(x.산출방식)}`, text: x.산출방식 }),
      el('td', { class: 'num', text: fmt(x.배정) }), el('td', { class: 'num', text: x.요구 === null ? '미정' : fmt(x.요구) }),
      el('td', { class: 'num' + (x.증감 < 0 ? ' neg' : ''), text: x.증감 === null ? '' : fmt(x.증감) }),
      el('td', { class: 'num', text: x.요구 === null ? '' : x.배정 ? `${(x.증감률 * 100).toFixed(1)}%` : '신규' }),
      el('td', { class: 'reason-' + (x.사유점검 === '사유 없음' ? 'bad' : x.사유점검 ? 'warn' : 'ok'), text: x.사유점검 || '' })));
    $('projTbl').replaceChildren(head, ...body);
  }
  if (s) {
    const rows = [el('tr', {}, ...['비목', `${d.year - 1} 배정액`, `${d.year} 요구액(초안)`, '증감', '증감률', '미산정', '확인 필요'].map((h, i) => el('th', { class: i ? 'num' : '', text: h })))];
    const line = (name, a, r, u, k, cls) => el('tr', { class: cls || '' }, el('td', { text: name }), el('td', { class: 'num', text: fmt(a) }), el('td', { class: 'num', text: fmt(r) }),
      el('td', { class: 'num' + (r - a < 0 ? ' neg' : ''), text: fmt(r - a) }), el('td', { class: 'num', text: a ? `${(((r - a) / a) * 100).toFixed(1)}%` : '신규' }),
      el('td', { class: 'num', text: u }), el('td', { class: 'num', text: k }));
    s.by_category.forEach((x) => rows.push(line(x.비목, x.배정, x.요구, x.미산정, x.확인필요)));
    rows.push(line('합계', s.배정, s.요구, s.미산정, s.확인필요, 'total'));
    $('sumTbl').replaceChildren(...rows);
  }
  $('exportBtn').disabled = !c.ok;
}
$('back3').addEventListener('click', () => go(3));
$('exportBtn').addEventListener('click', async () => {
  $('exportErr').textContent = '';
  state.draft.exported_at = nowIso();
  try {
    const r = await fetch('/api/export', { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(state.draft) });
    if (!r.ok) { const j = await r.json(); state.check = j.problems ? j : state.check; renderStep4(); throw new Error(j.error); }
    const blob = await r.blob();
    const tot = JSON.parse(decodeURIComponent(r.headers.get('X-Check-Total') || '%7B%7D'));
    download(blob, `${state.draft.year}_예산요구안_초안_${ymd()}.xlsx`);
    $('exportErr').classList.add('ok');
    $('exportErr').textContent = `내려받았습니다. 서버 검산: 배정 ${fmt(tot.배정)} / 요구 ${fmt(tot.요구)} / 증감 ${fmt(tot.증감)} ${state.draft.unit} (초안, 확인 필요)`;
    autosave();
  } catch (err) {
    $('exportErr').classList.remove('ok');
    $('exportErr').textContent = err.message.includes('fetch') ? '서버에 연결할 수 없습니다.' : err.message;
  }
});
$('saveDraft').addEventListener('click', () => {
  const safe = (state.draft.title || '예산요구안').replace(/[\\/:*?"<>|]/g, '_').slice(0, 40);
  download(new Blob([JSON.stringify(state.draft, null, 2)], { type: 'application/json' }), `초안_${safe}_${ymd()}.json`);
});
function download(blob, name) {
  const a = el('a', { href: URL.createObjectURL(blob), download: name });
  document.body.append(a); a.click(); a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 5000);
}

// 원본 다시 대조
$('reverifyInput').addEventListener('change', async (e) => {
  const f = e.target.files[0]; if (!f) return;
  const d = state.draft;
  const res = await apiParse(await f.arrayBuffer(), d.header_row, d.mapping);
  e.target.value = '';
  if (!res) return;
  if (res.errors.length) { $('reverifyLabel').textContent = '✖ ' + res.errors.join(' '); return; }
  const same = res.source_hash === d.source_hash;
  d.original = res.original; d.source_verified = true; d.source_hash = res.source_hash;
  $('reverifyLabel').textContent = same ? `✔ ${f.name} — 초안 작성 때와 같은 원본입니다.`
    : `⚠ ${f.name} — 초안 작성 때와 다른 파일입니다. 아래 금액 대조 결과를 확인하세요.`;
  changed(false);
});

// ───────── 초안 불러오기·자동 저장 ─────────
async function loadDraftFile(file) {
  $('step1Err').textContent = '';
  try {
    if (file.size > 5 * 1024 * 1024) throw new Error('초안 파일이 너무 큽니다.');
    const d = JSON.parse(await file.text());
    if (!FORMATS_OK.includes(d.format) || !Array.isArray(d.rows)) throw new Error('이 웹앱에서 저장한 초안 파일이 아닙니다.');
    d.source_verified = false; // 이 PC에서 원본 파일과 아직 대조하지 않음
    openDraft(d);
  } catch (err) {
    $('step1Err').textContent = err.message.startsWith('Unexpected') ? '초안 파일을 읽을 수 없습니다.' : err.message;
  }
}
// v1 초안(계획 산출 = 단가·수량 한 줄) → v2(산출 줄 목록)
function migrate(d) {
  if (d.format === FORMAT) return d;
  for (const r of d.rows) {
    if (!Array.isArray(r.산출)) r.산출 = [];
    if (r.산출방식 === '계획 산출' && !r.산출.length) {
      r.산출 = [newItem({ 단가: r.단가, 수량: r.수량, 단위: r.단위 || '' })];
      r.단가 = null; r.수량 = null; r.단위 = '';
    }
  }
  if (!('reason_threshold' in d)) d.reason_threshold = null;
  d.format = FORMAT;
  return d;
}
function openDraft(d) {
  migrate(d);
  state.draft = d; state.xlsx = null; state.parsed = null; state.selected = null; state.check = null;
  state.reached = 4;
  $('yearInput').value = d.year; $('monthInput').value = d.base_month; $('unitInput').value = d.unit; $('titleInput').value = d.title || '';
  $('thrInput').value = d.reason_threshold ?? '';
  setPractice(!!d.practice);
  $('reverifyLabel').textContent = '원본 엑셀을 올리면 초안의 원본 값과 대조합니다';
  scheduleCheck(0);
  go(3);
}
function autosave() {
  try { localStorage.setItem(AUTOSAVE_KEY, JSON.stringify({ saved_at: nowIso(), draft: state.draft })); } catch { /* 저장 불가 환경 */ }
}
function showResume() {
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem(AUTOSAVE_KEY) || 'null'); } catch { saved = null; }
  if (!saved?.draft?.rows) { $('welcomeBox').classList.remove('hidden'); return; }
  $('resumeBox').classList.remove('hidden');
  $('resumeText').textContent = `자동 저장된 초안이 있습니다: ${saved.draft.title} (세부사업 ${saved.draft.rows.length}개, ${saved.saved_at.slice(0, 16).replace('T', ' ')} 저장)`;
  $('resumeBtn').onclick = () => { $('resumeBox').classList.add('hidden'); openDraft(saved.draft); };
  $('resumeDiscard').onclick = () => { try { localStorage.removeItem(AUTOSAVE_KEY); } catch { /* 무시 */ } $('resumeBox').classList.add('hidden'); };
}

// ───────── 처음 쓰는 사람: 샘플 연습·도움말 ─────────
function setPractice(on) {
  state.practice = on;
  $('practiceTag').classList.toggle('hidden', !on);
}
async function startPractice() {
  if (state.draft?.rows?.length && !confirm('샘플로 연습을 시작하면 지금 작성 중인 화면을 샘플로 바꿉니다(자동 저장 초안도 덮어씀). 계속할까요?')) return;
  $('step1Err').textContent = '';
  try {
    const [x, m] = await Promise.all([fetch('/sample/budget.xlsx'), fetch('/sample/memo.md')]);
    if (!x.ok || !m.ok) throw new Error('샘플 파일을 찾지 못했습니다.');
    const file = new File([await x.arrayBuffer()], '[연습]2026_예산집행현황_샘플.xlsx');
    $('memoText').value = await m.text();
    $('memoText').dataset.name = '[연습]2027_사업계획_메모.md';
    $('memoLabel').textContent = '✔ [연습]2027_사업계획_메모.md';
    $('yearInput').value = 2027; $('monthInput').value = 10; $('unitInput').value = '천원';
    $('titleInput').value = '[연습] 2027년 ○○부 예산 요구안 초안';
    $('thrInput').value = 10;
    state.draft = null;
    setPractice(true);
    $('welcomeBox').classList.add('hidden');
    await loadXlsx(file);
  } catch (err) {
    $('step1Err').textContent = err.message.includes('fetch') ? '서버에 연결할 수 없습니다. 실행 창이 켜져 있는지 확인하세요.' : err.message;
  }
}
$('practiceBtn').addEventListener('click', startPractice);
$('practiceBtn2').addEventListener('click', startPractice);
$('xlsxInput').addEventListener('change', () => setPractice(false));
$('xlsxDrop').addEventListener('drop', () => setPractice(false));
const toggleHelp = (open) => {
  const h = $('help');
  h.classList.toggle('hidden', open === undefined ? !h.classList.contains('hidden') : !open);
  if (!h.classList.contains('hidden')) h.scrollIntoView({ block: 'start' });
};
$('helpBtn').addEventListener('click', () => toggleHelp());
$('helpBtn2').addEventListener('click', () => toggleHelp(true));
$('helpClose').addEventListener('click', () => toggleHelp(false));
