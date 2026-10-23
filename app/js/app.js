/*
 * 화면 렌더링과 입력 처리. 계산은 모두 HR.engine에 맡긴다.
 */
(function () {
  'use strict';

  const E = HR.engine;
  const GRADES = E.GRADES;
  const STORAGE_KEY = 'hr-kpi-eval-v1';
  const UI_KEY = 'hr-kpi-eval-ui';
  const XLSX_URL = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  let state = load();
  let results = null;
  const ui = Object.assign({ view: 'dashboard', team: '', q: '', grade: '', modal: null }, loadUi());

  // ---------- 저장 ----------
  function load() {
    try {
      const raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        const saved = JSON.parse(raw);
        const migrating = !(saved.version >= 2);
        const s = normalize(saved);
        if (migrating) localStorage.setItem(STORAGE_KEY, JSON.stringify(s));
        return s;
      }
    } catch (e) { /* 저장소를 쓸 수 없으면 데모 데이터로 시작 */ }
    return HR.data.demoState();
  }

  // 예전 버전에서 저장된 설정에 새 항목이 빠져 있으면 기본값으로 채운다
  function normalize(s) {
    const defaults = HR.data.defaultSettings();
    s.settings = Object.assign({}, defaults, s.settings || {});
    s.teams = s.teams || [];
    s.employees = s.employees || [];
    // v2: P등급 조정값을 최종 인상률(연봉 삭감)로 확정 — 이전 버전에서 저장된 해석값을 바꾼다
    if (!(s.version >= 2)) {
      s.settings.pAbsolute = true;
      s.version = 2;
    }
    return s;
  }

  function save() {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
    } catch (e) {
      toast('브라우저 저장소에 저장하지 못했습니다.');
    }
  }

  function loadUi() {
    try { return { view: JSON.parse(localStorage.getItem(UI_KEY) || '{}').view || 'dashboard' }; } catch (e) { return {}; }
  }

  function saveUi() {
    try { localStorage.setItem(UI_KEY, JSON.stringify({ view: ui.view })); } catch (e) { /* 무시 */ }
  }

  // ---------- 포맷 ----------
  const isNum = (v) => typeof v === 'number' && isFinite(v);
  const esc = (v) => String(v == null ? '' : v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const trim = (v) => String(+v.toFixed(2));
  // 음수는 표 안의 다른 값(σ−1% 등)과 같은 빼기 기호(−)로 표시
  const minus = (s) => s.replace(/^-/, '−');
  const pct = (v, d = 1) => (isNum(v) ? minus((v * 100).toFixed(d)) + '%' : '—');
  const pctShort = (v) => (isNum(v) ? minus(trim(v * 100)) + '%' : '—');
  const won = (v) => (isNum(v) ? minus(Math.round(v).toLocaleString('ko-KR')) + '원' : '—');
  const num = (v) => (isNum(v) ? minus(Math.round(v).toLocaleString('ko-KR')) : '—');
  const eok = (v) => (isNum(v) ? (+v.toFixed(2)).toLocaleString('ko-KR') + '억' : '—');
  const signed = (v) => (isNum(v) ? (v > 0 ? '+' : v < 0 ? '−' : '') + trim(Math.abs(v) * 100) + '%' : '—');
  const sigmaLabel = (v) => (!isNum(v) ? '—' : v === 0 ? 'σ' : 'σ' + signed(v));
  const deductLabel = (v) => (v ? '−' + trim(v * 100) + '%p' : '0');

  function gradeChip(g) {
    return g ? `<span class="grade grade-${g}">${g}</span>` : '<span class="grade grade-none">—</span>';
  }

  function teamById(id) {
    return state.teams.find((t) => t.id === id);
  }

  function teamName(id) {
    const t = teamById(id);
    return t ? t.name : '(팀 없음)';
  }

  function compaLabels() {
    const b = state.settings.compaBands.map((v) => trim(v * 100) + '%');
    return [b[0] + ' 초과'].concat(b.slice(1).map((v, i) => v + ' 초과 ~ ' + b[i] + ' 이하'), [b[b.length - 1] + ' 이하']);
  }

  function bandLabels(bands) {
    return bands.map((b, i) => {
      const lo = trim(b.min * 100) + '%';
      if (i === 0) return lo + ' 이상';
      const hi = trim(bands[i - 1].min * 100) + '%';
      if (i === bands.length - 1) return hi + ' 미만';
      return lo + ' ~ ' + hi + ' 미만';
    });
  }

  // ---------- 입력칸 ----------
  function fieldValue(type, v) {
    if (type === 'pct') return isNum(v) ? trim(v * 100) : '';
    if (type === 'num' || type === 'int') return isNum(v) ? String(v) : '';
    if (type === 'won') return isNum(v) ? Math.round(v).toLocaleString('ko-KR') : '';
    return v == null ? '' : v;
  }

  function input(kind, id, path, value, type, cls = '', extra = '') {
    const key = `${kind}:${id}:${path}`;
    const inputType = type === 'date' ? 'date' : 'text';
    const mode = type === 'text' || type === 'date' ? '' : ' inputmode="decimal"';
    return `<input class="cell-input ${cls}" type="${inputType}"${mode} data-key="${esc(key)}" data-kind="${kind}" data-id="${esc(id)}" data-path="${esc(path)}" data-type="${type}" value="${esc(fieldValue(type, value))}" ${extra}>`;
  }

  function select(kind, id, path, value, options, cls = '') {
    const key = `${kind}:${id}:${path}`;
    const opts = options.map((o) => {
      const [v, label] = Array.isArray(o) ? o : [o, o];
      return `<option value="${esc(v)}"${v === value ? ' selected' : ''}>${esc(label)}</option>`;
    }).join('');
    return `<select class="cell-input ${cls}" data-key="${esc(key)}" data-kind="${kind}" data-id="${esc(id)}" data-path="${esc(path)}" data-type="select">${opts}</select>`;
  }

  function checkbox(kind, id, path, value) {
    const key = `${kind}:${id}:${path}`;
    return `<input type="checkbox" data-key="${esc(key)}" data-kind="${kind}" data-id="${esc(id)}" data-path="${esc(path)}" data-type="bool"${value ? ' checked' : ''}>`;
  }

  function parseValue(type, el) {
    const raw = el.value.trim();
    switch (type) {
      case 'pct': {
        if (raw === '') return null;
        const n = parseFloat(raw.replace(/[%,\s]/g, ''));
        return isNaN(n) ? undefined : Math.round(n * 100) / 10000;
      }
      case 'num': {
        if (raw === '') return null;
        const n = parseFloat(raw.replace(/[,\s]/g, ''));
        return isNaN(n) ? undefined : n;
      }
      case 'int': {
        const n = parseInt(raw.replace(/[,\s]/g, '') || '0', 10);
        return isNaN(n) || n < 0 ? undefined : n;
      }
      case 'won': {
        if (raw === '') return null;
        const n = parseInt(raw.replace(/[^0-9]/g, ''), 10);
        return isNaN(n) ? undefined : n;
      }
      case 'bool': return el.checked;
      default: return el.value;
    }
  }

  function targetOf(kind, id) {
    if (kind === 'emp') return state.employees.find((e) => e.id === id);
    if (kind === 'team') return state.teams.find((t) => t.id === id);
    if (kind === 'set') return state.settings;
    return null;
  }

  function setPath(obj, path, value) {
    const parts = path.split('.');
    let o = obj;
    for (let i = 0; i < parts.length - 1; i++) {
      if (o[parts[i]] == null) o[parts[i]] = {};
      o = o[parts[i]];
    }
    o[parts[parts.length - 1]] = value;
  }

  // ---------- 렌더링 ----------
  let renderQueued = false;
  function scheduleRender() {
    if (renderQueued) return;
    renderQueued = true;
    // change 이벤트 뒤 포커스가 다음 칸으로 옮겨진 다음에 다시 그린다
    setTimeout(() => { renderQueued = false; render(); }, 0);
  }

  function render() {
    const active = document.activeElement;
    const focusKey = active && active.dataset ? active.dataset.key : null;
    const caret = focusKey && active.selectionStart != null ? active.selectionStart : null;

    results = E.evaluateAll(state);
    document.getElementById('cycle-label').textContent = state.settings.cycleName + ' 평가';
    document.querySelectorAll('.nav-btn').forEach((b) => b.classList.toggle('active', b.dataset.view === ui.view));

    const views = { dashboard, teams, people, raise, roster, settings };
    document.getElementById('view').innerHTML = (views[ui.view] || dashboard)();
    document.getElementById('modal-root').innerHTML = ui.modal ? evalSheet(ui.modal) : '';

    if (focusKey) {
      const el = document.querySelector(`[data-key="${CSS.escape(focusKey)}"]`);
      if (el) {
        el.focus();
        if (caret != null && el.setSelectionRange && el.type !== 'date') {
          try { el.setSelectionRange(caret, caret); } catch (e) { /* 일부 입력칸은 지원하지 않음 */ }
        }
      }
    }
  }

  function pageHead(title, desc, actions = '') {
    const demo = state.demo ? ' <span class="tag info">데모 데이터</span>' : '';
    return `<div class="page-head"><div><h1 class="page-title">${title}${demo}</h1><p class="page-desc">${desc}</p></div><div class="head-actions">${actions}</div></div>`;
  }

  function eligibleEmployees() {
    return state.employees.filter((e) => e.eligible);
  }

  function filteredEmployees() {
    return eligibleEmployees().filter((e) => {
      if (ui.team && e.teamId !== ui.team) return false;
      if (ui.q && !e.name.includes(ui.q.trim())) return false;
      if (ui.grade) {
        const g = results.employees[e.id].finalGrade;
        if (ui.grade === 'none' ? g : g !== ui.grade) return false;
      }
      return true;
    });
  }

  function filterBar(count) {
    const teamOpts = [['', '전체 팀']].concat(state.teams.map((t) => [t.id, t.name]));
    const gradeOpts = [['', '전체 등급']].concat(GRADES.map((g) => [g, '종합 ' + g]), [['none', '미산출']]);
    const opt = (list, cur) => list.map(([v, l]) => `<option value="${v}"${v === cur ? ' selected' : ''}>${esc(l)}</option>`).join('');
    return `<div class="toolbar no-print">
      <input type="search" placeholder="이름 검색" data-ui="q" data-key="ui:q" value="${esc(ui.q)}">
      <select data-ui="team">${opt(teamOpts, ui.team)}</select>
      <select data-ui="grade">${opt(gradeOpts, ui.grade)}</select>
      <span class="spacer"></span><span class="count">${count}명</span>
    </div>`;
  }

  // ---------- 대시보드 ----------
  function dashboard() {
    const s = state.settings;
    const emps = eligibleEmployees();
    const done = emps.filter((e) => results.employees[e.id].complete);
    const salarySum = done.reduce((a, e) => a + e.salary, 0);
    const incSum = done.reduce((a, e) => a + results.employees[e.id].increase, 0);
    const avgRaise = done.length ? done.reduce((a, e) => a + results.employees[e.id].finalRaise, 0) / done.length : null;
    const overrides = emps.filter((e) => results.employees[e.id].overridden);

    const counts = {};
    GRADES.forEach((g) => { counts[g] = 0; });
    done.forEach((e) => { counts[results.employees[e.id].finalGrade]++; });
    const maxCount = Math.max(1, ...GRADES.map((g) => counts[g]));

    const revRate = E.rate(s.company.revenueActual, s.company.revenueTarget);
    const profRate = E.rate(s.company.profitActual, s.company.profitTarget);
    const bar = (label, r, a, t) => `<div class="progress-row">
      <div class="progress-label"><span>${label} <span class="muted">${eok(a)} / ${eok(t)}</span></span><b>${pct(r)}</b></div>
      <div class="progress"><i style="width:${isNum(r) ? Math.min(100, r * 100) : 0}%"></i></div></div>`;

    const bars = GRADES.map((g) => {
      const c = counts[g];
      const share = done.length ? c / done.length : 0;
      return `<div class="bar-col"><span class="tip">종합 ${g}등급 · ${c}명 (${pct(share)}) · 기본 인상률 ${pctShort(s.baseRates[g])}</span>
        <span class="bar-val">${c}</span><div class="bar" style="height:${(c / maxCount) * 140}px"></div></div>`;
    }).join('');

    const teamRows = state.teams.map((t) => {
      const r = results.teams[t.id];
      const members = emps.filter((e) => e.teamId === t.id);
      const fin = members.filter((e) => results.employees[e.id].complete);
      const avg = fin.length ? fin.reduce((a, e) => a + results.employees[e.id].finalRaise, 0) / fin.length : null;
      return `<tr><td>${esc(t.division)}</td><td>${esc(t.name)}</td><td class="num">${pct(r.rate)}</td><td class="center">${gradeChip(r.grade)}</td>
        <td class="num">${members.length}명</td><td class="num">${pct(avg)}</td></tr>`;
    }).join('');

    const missing = emps.filter((e) => !results.employees[e.id].complete);
    const missingList = missing.slice(0, 8).map((e) => `<tr><td class="name-cell"><b>${esc(e.name)}</b><span>${esc(teamName(e.teamId))}</span></td>
      <td>${results.employees[e.id].missing.map((m) => `<span class="tag warn">${m} 없음</span>`).join(' ')}</td></tr>`).join('');

    return pageHead('대시보드', `${esc(s.cycleName)} · 평가기준일 ${esc(s.evalDate)}`, `<button class="btn primary" data-action="export-xlsx">결과 엑셀 내보내기</button>`) + `
      <div class="grid grid-4">
        <div class="card stat"><div class="stat-label">평가 대상자</div><div class="stat-value">${emps.length}명</div><div class="stat-note">팀 ${state.teams.length}개</div></div>
        <div class="card stat"><div class="stat-label">산출 완료</div><div class="stat-value">${done.length}명</div><div class="stat-note">${emps.length ? pct(done.length / emps.length, 0) : '—'} 완료</div></div>
        <div class="card stat"><div class="stat-label">평균 인상률</div><div class="stat-value">${pct(avgRaise)}</div><div class="stat-note">인건비 기준 ${salarySum ? pct(incSum / salarySum) : '—'}</div></div>
        <div class="card stat"><div class="stat-label">연간 인건비 증가</div><div class="stat-value">${isNum(incSum) ? (incSum / 1e8).toFixed(2) + '억' : '—'}</div><div class="stat-note">부사장 조정 ${overrides.length}건</div></div>
      </div>
      <div class="grid grid-2">
        <div class="card"><h2>종합 등급 분포<span class="sub">산출 완료 ${done.length}명</span></h2>
          <div class="bars">${bars}</div>
          <div class="bar-axis">${GRADES.map((g) => `<span>${gradeChip(g)}</span>`).join('')}</div>
        </div>
        <div class="card"><h2>전사 실적마감<span class="sub">참고 지표</span></h2>
          ${bar('전사 매출', revRate, s.company.revenueActual, s.company.revenueTarget)}
          ${bar('전사 영업이익', profRate, s.company.profitActual, s.company.profitTarget)}
          <p class="hint">전사 실적은 평가표의 실적마감 항목으로, 등급 산출에는 직접 쓰이지 않습니다. 값은 평가 기준 설정에서 입력합니다.</p>
        </div>
      </div>
      <div class="grid grid-2">
        <div class="card"><h2>팀별 현황</h2><div class="table-wrap"><table>
          <thead><tr><th>사업부</th><th>팀</th><th class="num">팀 달성률</th><th class="center">팀 등급</th><th class="num">인원</th><th class="num">평균 인상률</th></tr></thead>
          <tbody>${teamRows}</tbody></table></div></div>
        <div class="card"><h2>입력이 필요한 대상자<span class="sub">${missing.length}명</span></h2>
          ${missing.length ? `<div class="table-wrap"><table><tbody>${missingList}</tbody></table></div>${missing.length > 8 ? `<p class="hint">외 ${missing.length - 8}명</p>` : ''}` : '<p class="muted">모든 대상자의 인상률이 산출되었습니다.</p>'}
        </div>
      </div>`;
  }

  // ---------- 팀 평가 ----------
  function teams() {
    const emps = eligibleEmployees();
    const rows = state.teams.map((t) => {
      const r = results.teams[t.id];
      const n = emps.filter((e) => e.teamId === t.id).length;
      return `<tr>
        <td>${input('team', t.id, 'division', t.division, 'text', 'mid')}</td>
        <td>${input('team', t.id, 'name', t.name, 'text', 'mid')}</td>
        <td>${input('team', t.id, 'goal', t.goal, 'text', 'wide')}</td>
        <td class="bl">${input('team', t.id, 'revenueTarget', t.revenueTarget, 'num')}</td>
        <td>${input('team', t.id, 'revenueActual', t.revenueActual, 'num')}</td>
        <td class="num">${pct(r.revenueRate)}</td>
        <td class="bl">${input('team', t.id, 'profitTarget', t.profitTarget, 'num')}</td>
        <td>${input('team', t.id, 'profitActual', t.profitActual, 'num')}</td>
        <td class="num">${pct(r.profitRate)}</td>
        <td class="num bl"><b>${pct(r.rate)}</b></td>
        <td class="center">${gradeChip(r.grade)}</td>
        <td class="num">${n}명</td>
        <td><button class="btn small danger" data-action="del-team" data-id="${t.id}">삭제</button></td>
      </tr>`;
    }).join('');

    const labels = bandLabels(state.settings.teamBands);
    return pageHead('팀 평가', '팀 달성률 = (매출 달성률 + 영업이익 달성률) ÷ 2. 단위는 억원입니다.', '<button class="btn" data-action="add-team">팀 추가</button>') + `
      <div class="table-wrap"><table>
        <thead>
          <tr><th colspan="3"></th><th colspan="3" class="group">팀 매출</th><th colspan="3" class="group">팀 영업이익</th><th colspan="4" class="group">팀 평가</th></tr>
          <tr><th>사업부</th><th>팀</th><th>팀 목표</th><th class="bl">목표</th><th>실적</th><th class="num">달성률</th><th class="bl">목표</th><th>실적</th><th class="num">달성률</th><th class="num bl">팀 달성률</th><th class="center">등급</th><th class="num">인원</th><th></th></tr>
        </thead>
        <tbody>${rows || '<tr><td colspan="13" class="muted">팀이 없습니다.</td></tr>'}</tbody>
      </table></div>
      <p class="hint">팀 등급 구간: ${state.settings.teamBands.map((b, i) => `${b.grade} ${labels[i]}`).join(' · ')}</p>`;
  }

  // ---------- 개인 평가 ----------
  function people() {
    const list = filteredEmployees();
    const u = state.settings.deductionUnits;
    const rows = list.map((e) => {
      const r = results.employees[e.id];
      const ev = e.evaluation || {};
      const d = e.deductions || {};
      const src = { final: '최종', head: '부문장', self: '본인' }[r.kpiSource] || '';
      return `<tr>
        <td class="name-cell"><b>${esc(e.name)}</b><span>${esc(e.level)}</span></td>
        <td>${esc(teamName(e.teamId))}</td>
        <td>${input('emp', e.id, 'kpi', e.kpi, 'text', 'wide')}</td>
        <td class="bl">${input('emp', e.id, 'evaluation.selfRate', ev.selfRate, 'pct')}</td>
        <td>${input('emp', e.id, 'evaluation.headRate', ev.headRate, 'pct')}</td>
        <td>${input('emp', e.id, 'evaluation.finalRate', ev.finalRate, 'pct', '', 'placeholder="자동"')}</td>
        <td class="bl">${input('emp', e.id, 'deductions.late', d.late, 'int', 'narrow')}</td>
        <td>${input('emp', e.id, 'deductions.leave', d.leave, 'int', 'narrow')}</td>
        <td>${input('emp', e.id, 'deductions.training', d.training, 'int', 'narrow')}</td>
        <td class="num">${deductLabel(r.deduction)}</td>
        <td class="num bl" title="${src} 평가 기준">${pct(r.personalScore)}</td>
        <td class="center">${gradeChip(r.personalGrade)}</td>
        <td class="center">${gradeChip(r.teamGrade)}</td>
        <td class="center">${gradeChip(r.finalGrade)}</td>
        <td><button class="btn small" data-action="open-sheet" data-id="${e.id}">평가표</button></td>
      </tr>`;
    }).join('');

    return pageHead('개인 평가', '개인 점수 = 개인 KPI 점수(최종 → 부문장 → 본인 순으로 적용) − 감점. 종합 등급은 팀 등급과 개인 등급의 매트릭스로 정합니다.') +
      filterBar(list.length) + `
      <div class="table-wrap"><table>
        <thead>
          <tr><th colspan="3"></th><th colspan="3" class="group">개인 KPI 평가 (%)</th><th colspan="4" class="group">감점요인 (회)</th><th colspan="4" class="group">결과</th><th></th></tr>
          <tr><th>이름</th><th>팀</th><th>개인 KPI</th>
            <th class="bl">본인</th><th>부문장</th><th>최종</th>
            <th class="bl" title="1회당 ${trim(u.late * 100)}%p">지각</th><th title="1회당 ${trim(u.leave * 100)}%p">초과연차</th><th title="1회당 ${trim(u.training * 100)}%p">교육미이수</th><th class="num">감점</th>
            <th class="num bl">개인 점수</th><th class="center">개인</th><th class="center">팀</th><th class="center">종합</th><th></th></tr>
        </thead>
        <tbody>${rows || '<tr><td colspan="15" class="muted">조건에 맞는 대상자가 없습니다.</td></tr>'}</tbody>
      </table></div>`;
  }

  // ---------- 연봉 인상 ----------
  function raise() {
    const list = filteredEmployees();
    const labels = compaLabels();
    let cur = 0;
    let next = 0;
    const rows = list.map((e) => {
      const r = results.employees[e.id];
      if (isNum(r.newSalary)) { cur += e.salary; next += r.newSalary; }
      const needReason = r.overridden && !String(e.vpReason || '').trim();
      return `<tr>
        <td class="name-cell"><b>${esc(e.name)}</b><span>${esc(e.level)}</span></td>
        <td>${esc(teamName(e.teamId))}</td>
        <td class="center">${gradeChip(r.finalGrade)}</td>
        <td class="num bl">${num(e.salary)}</td>
        <td class="num">${num(r.market)}</td>
        <td class="num">${pct(r.compa)}</td>
        <td class="muted">${r.compaBand != null ? labels[r.compaBand] : '—'}</td>
        <td class="num bl">${pctShort(r.sigma)}</td>
        <td class="num">${r.finalGrade === 'P' && state.settings.pAbsolute ? signed(r.adjust) : sigmaLabel(r.adjust)}</td>
        <td class="num"><b>${pct(r.calcRaise)}</b></td>
        <td class="bl">${input('emp', e.id, 'vpOverride', e.vpOverride, 'pct', '', 'placeholder="—"')}</td>
        <td>${input('emp', e.id, 'vpReason', e.vpReason, 'text', 'wide', 'placeholder="조정 시 사유 필수"')}${needReason ? ' <span class="tag danger">사유 필요</span>' : ''}</td>
        <td class="num bl"><b>${pct(r.finalRaise)}</b></td>
        <td class="num">${num(r.newSalary)}</td>
        <td class="num">${num(r.increase)}</td>
      </tr>`;
    }).join('');

    return pageHead('연봉 인상', '산출 인상률 = 종합등급별 기본 인상률(σ) + 연봉 수준(현재연봉 ÷ 업계평균)별 조정. 금액 단위는 원입니다.', '<button class="btn primary" data-action="export-xlsx">결과 엑셀 내보내기</button>') +
      filterBar(list.length) + `
      <div class="table-wrap"><table>
        <thead>
          <tr><th colspan="3"></th><th colspan="4" class="group">연봉 수준</th><th colspan="3" class="group">산출</th><th colspan="2" class="group">부사장 검토</th><th colspan="3" class="group">최종</th></tr>
          <tr><th>이름</th><th>팀</th><th class="center">종합</th>
            <th class="num bl">현재연봉</th><th class="num">업계평균</th><th class="num">연봉 수준</th><th>구간</th>
            <th class="num bl">σ</th><th class="num">조정</th><th class="num">산출 인상률</th>
            <th class="bl">조정 인상률(%)</th><th>사유</th>
            <th class="num bl">최종 인상률</th><th class="num">인상 연봉</th><th class="num">증가액</th></tr>
        </thead>
        <tbody>${rows || '<tr><td colspan="15" class="muted">조건에 맞는 대상자가 없습니다.</td></tr>'}</tbody>
        <tfoot><tr><td colspan="3">합계 (산출 완료자)</td><td class="num bl">${num(cur)}</td><td colspan="8"></td>
          <td class="num bl">${cur ? pct((next - cur) / cur) : '—'}</td><td class="num">${num(next)}</td><td class="num">${num(next - cur)}</td></tr></tfoot>
      </table></div>`;
  }

  // ---------- 대상자 관리 ----------
  function roster() {
    const levels = state.settings.market.map((m) => m.level);
    const teamOpts = state.teams.map((t) => [t.id, `${t.name} (${t.division})`]);
    const rows = state.employees.map((e) => {
      const r = results.employees[e.id];
      const lvOpts = levels.includes(e.level) ? levels : [e.level].concat(levels);
      return `<tr>
        <td>${input('emp', e.id, 'name', e.name, 'text', 'mid')}</td>
        <td>${select('emp', e.id, 'teamId', e.teamId, teamOpts, 'wide')}</td>
        <td>${select('emp', e.id, 'level', e.level, lvOpts)}</td>
        <td>${input('emp', e.id, 'salary', e.salary, 'won', 'mid')}</td>
        <td>${input('emp', e.id, 'hireDate', e.hireDate, 'date', 'mid')}</td>
        <td>${input('emp', e.id, 'lastRaiseDate', e.lastRaiseDate, 'date', 'mid')}</td>
        <td class="num">${isNum(r.stayDays) ? r.stayDays + '일' : '—'}</td>
        <td class="center">${checkbox('emp', e.id, 'eligible', e.eligible)}</td>
        <td><button class="btn small danger" data-action="del-emp" data-id="${e.id}">삭제</button></td>
      </tr>`;
    }).join('');

    return pageHead('대상자 관리', `엑셀 평가지의 '대상자' 시트를 그대로 가져올 수 있습니다. 연봉체류일 = 평가기준일(${esc(state.settings.evalDate)}) − 기존 인상일.`,
      '<button class="btn" data-action="add-emp">대상자 추가</button><button class="btn primary" data-action="import-xlsx">엑셀에서 가져오기</button>') + `
      ${state.demo ? '<div class="callout info">지금은 가상 인물로 만든 데모 데이터입니다. 실제 평가 파일을 가져오면 데모 데이터는 교체됩니다.</div>' : ''}
      <div class="table-wrap"><table>
        <thead><tr><th>성명</th><th>팀</th><th>직급</th><th>현재연봉(원)</th><th>입사일</th><th>기존 인상일</th><th class="num">연봉체류일</th><th class="center">대상</th><th></th></tr></thead>
        <tbody>${rows || '<tr><td colspan="9" class="muted">대상자가 없습니다.</td></tr>'}</tbody>
      </table></div>
      <p class="hint">총 ${state.employees.length}명 · 평가 대상 ${eligibleEmployees().length}명</p>`;
  }

  // ---------- 설정 ----------
  function settings() {
    const s = state.settings;
    const bandTable = (key) => {
      const labels = bandLabels(s[key]);
      return `<div class="scroll-x"><table class="matrix"><thead><tr><th>등급</th><th>하한 (%)</th><th>구간</th></tr></thead><tbody>${s[key].map((b, i) => `<tr>
        <td>${gradeChip(b.grade)}</td>
        <td>${i === s[key].length - 1 ? '<span class="muted">그 외</span>' : input('set', '', `${key}.${i}.min`, b.min, 'pct')}</td>
        <td class="muted">${labels[i]}</td></tr>`).join('')}</tbody></table></div>`;
    };
    const bandsOk = (key) => s[key].every((b, i) => i === 0 || b.min <= s[key][i - 1].min || i === s[key].length - 1);
    const compaOk = s.compaBands.every((v, i) => i === 0 || v < s.compaBands[i - 1]);

    const combined = `<div class="scroll-x"><table class="matrix"><thead><tr><th>개인 / 팀</th>${GRADES.map((g) => `<th>${gradeChip(g)}</th>`).join('')}</tr></thead><tbody>
      ${GRADES.map((p) => `<tr><td>${gradeChip(p)}</td>${GRADES.map((t) => `<td>${select('set', '', `combinedMatrix.${p}.${t}`, s.combinedMatrix[p][t], GRADES)}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>`;

    const labels = compaLabels();
    const compa = `<div class="scroll-x"><table class="matrix"><thead><tr><th>연봉 수준</th>${GRADES.map((g) => `<th>${gradeChip(g)}</th>`).join('')}</tr></thead><tbody>
      <tr><td class="muted">기본 인상률 σ</td>${GRADES.map((g) => `<td>${input('set', '', `baseRates.${g}`, s.baseRates[g], 'pct')}</td>`).join('')}</tr>
      ${s.compaMatrix.map((row, i) => `<tr><td>${labels[i]}</td>${GRADES.map((g) => `<td>${input('set', '', `compaMatrix.${i}.${g}`, row[g], 'pct')}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>`;

    const market = `<table class="matrix"><thead><tr><th>직급</th><th>업계평균 연봉 (원)</th></tr></thead><tbody>
      ${s.market.map((m, i) => `<tr><td>${esc(m.level)}</td><td>${input('set', '', `market.${i}.salary`, m.salary, 'won', 'mid')}</td></tr>`).join('')}
      </tbody></table>`;

    return pageHead('평가 기준 설정', '엑셀 평가지의 기준표를 옮겨 둔 값입니다. 바꾸면 모든 결과가 바로 다시 계산됩니다.') + `
      <div class="grid grid-2">
        <div class="card"><h2>평가 회차</h2>
          <div class="form-row"><label>회차 이름</label>${input('set', '', 'cycleName', s.cycleName, 'text', 'wide')}</div>
          <div class="form-row"><label>평가기준일</label>${input('set', '', 'evalDate', s.evalDate, 'date', 'mid')}</div>
          <h2 style="margin-top:18px">전사 실적마감 <span class="sub">억원</span></h2>
          <div class="form-row"><label>전사 매출 목표 / 실적</label>${input('set', '', 'company.revenueTarget', s.company.revenueTarget, 'num', 'mid')}${input('set', '', 'company.revenueActual', s.company.revenueActual, 'num', 'mid')}</div>
          <div class="form-row"><label>전사 영업이익 목표 / 실적</label>${input('set', '', 'company.profitTarget', s.company.profitTarget, 'num', 'mid')}${input('set', '', 'company.profitActual', s.company.profitActual, 'num', 'mid')}</div>
        </div>
        <div class="card"><h2>감점 기준 <span class="sub">1회당 차감 %p</span></h2>
          <div class="form-row"><label>지각</label>${input('set', '', 'deductionUnits.late', s.deductionUnits.late, 'pct')}</div>
          <div class="form-row"><label>초과연차사용</label>${input('set', '', 'deductionUnits.leave', s.deductionUnits.leave, 'pct')}</div>
          <div class="form-row"><label>교육미이수</label>${input('set', '', 'deductionUnits.training', s.deductionUnits.training, 'pct')}</div>
          <p class="hint">엑셀 평가지에는 감점 단위가 정의되어 있지 않아 임의로 정한 기본값입니다. 회사 기준에 맞게 바꿔 주세요.</p>
        </div>
      </div>
      <div class="grid grid-2">
        <div class="card"><h2>팀 평가 등급 구간</h2>${bandsOk('teamBands') ? '' : '<div class="callout">하한값이 높은 등급부터 내림차순이어야 합니다.</div>'}${bandTable('teamBands')}</div>
        <div class="card"><h2>개인 평가 등급 구간</h2>${bandsOk('personalBands') ? '' : '<div class="callout">하한값이 높은 등급부터 내림차순이어야 합니다.</div>'}${bandTable('personalBands')}</div>
      </div>
      <div class="card"><h2>팀/개인 종합 등급 매트릭스 <span class="sub">팀 50% · 개인 50%</span></h2>${combined}</div>
      <div class="card"><h2>연봉 수준별 인상률 <span class="sub">σ 대비 조정값 (%)</span></h2>
        <div class="form-row"><label>연봉 수준 구간 경계 (%)</label>${s.compaBands.map((v, i) => input('set', '', `compaBands.${i}`, v, 'pct')).join('')}</div>
        ${compaOk ? '' : '<div class="callout">구간 경계는 큰 값부터 내림차순이어야 합니다.</div>'}
        ${compa}
        <div class="form-row" style="margin-top:14px"><label>P등급 절대값 적용</label>${checkbox('set', '', 'pAbsolute', s.pAbsolute)}
          <span class="muted">켜져 있으면 P등급은 σ를 더하지 않고 표의 값을 그대로 최종 인상률로 씁니다. 시장 대비 고연봉자는 연봉이 삭감됩니다 (예: −3%).</span></div>
        <div class="form-row"><label>최저 인상률 (%)</label>${input('set', '', 'minRaise', s.minRaise, 'pct', '', 'placeholder="없음"')}<span class="muted">비워 두면 하한을 두지 않습니다.</span></div>
      </div>
      <div class="grid grid-2">
        <div class="card"><h2>동종업계 직급별 평균 연봉</h2><div class="table-wrap" style="max-height:420px;overflow-y:auto">${market}</div></div>
        <div class="card"><h2>데이터 관리</h2>
          <div class="toolbar">
            <button class="btn" data-action="backup">백업 파일 저장 (JSON)</button>
            <button class="btn" data-action="restore">백업 파일 불러오기</button>
          </div>
          <div class="toolbar">
            <button class="btn" data-action="reset-settings">평가 기준 초기화</button>
            <button class="btn" data-action="load-demo">데모 데이터 불러오기</button>
            <button class="btn danger" data-action="clear-all">전체 데이터 삭제</button>
          </div>
          <p class="hint">데이터는 이 브라우저의 저장소에만 있습니다. 브라우저 데이터를 지우면 함께 사라지므로 주기적으로 백업해 주세요.</p>
        </div>
      </div>`;
  }

  // ---------- 개인 평가표 (엑셀 '이름' 시트와 같은 순서) ----------
  function evalSheet(id) {
    const e = state.employees.find((x) => x.id === id);
    if (!e) return '';
    const s = state.settings;
    const t = teamById(e.teamId) || {};
    const tr = results.teams[e.teamId] || {};
    const r = results.employees[e.id];
    const ev = e.evaluation || {};
    const d = e.deductions || {};
    const u = s.deductionUnits;
    const hl = (on) => (on ? ' class="hl"' : '');

    const bandRow = (bands, grade) => {
      const labels = bandLabels(bands);
      return `<table><thead><tr><th>구분</th>${bands.map((b) => `<th>${b.grade}</th>`).join('')}</tr></thead>
        <tbody><tr><td>점수 구간</td>${bands.map((b, i) => `<td${hl(b.grade === grade)}>${labels[i]}</td>`).join('')}</tr></tbody></table>`;
    };
    const labels = compaLabels();

    return `<div class="modal-backdrop" data-action="close-modal-bg"><div class="modal sheet">
      <div class="modal-head"><h2>${esc(s.cycleName)} 평가표 · ${esc(e.name)}</h2>
        <div class="no-print"><button class="btn" data-action="print">인쇄</button> <button class="btn" data-action="close-modal">닫기</button></div></div>
      <table><thead><tr><th>구분</th><th>팀</th><th>이름</th><th>직급</th><th>팀 목표</th><th>개인 KPI</th></tr></thead>
        <tbody><tr><td>${esc(e.division || t.division)}</td><td>${esc(t.name)}</td><td>${esc(e.name)}</td><td>${esc(e.level)}</td><td>${esc(t.goal)}</td><td>${esc(e.kpi)}</td></tr></tbody></table>

      <div class="sheet-section"><h3>실적마감 <span class="muted">(억원)</span></h3>
        <table><thead><tr><th>구분</th><th>목표</th><th>누적 달성액</th><th>달성률</th></tr></thead><tbody>
          <tr><td>전사 매출</td><td>${eok(s.company.revenueTarget)}</td><td>${eok(s.company.revenueActual)}</td><td>${pct(E.rate(s.company.revenueActual, s.company.revenueTarget))}</td></tr>
          <tr><td>전사 영업이익</td><td>${eok(s.company.profitTarget)}</td><td>${eok(s.company.profitActual)}</td><td>${pct(E.rate(s.company.profitActual, s.company.profitTarget))}</td></tr>
        </tbody></table></div>

      <div class="sheet-section"><h3>팀/매출 평가</h3>
        <table><thead><tr><th>구분</th><th>팀 목표</th><th>실적</th><th>달성률</th></tr></thead><tbody>
          <tr><td>팀 매출</td><td>${eok(t.revenueTarget)}</td><td>${eok(t.revenueActual)}</td><td>${pct(tr.revenueRate)}</td></tr>
          <tr><td>팀 영업이익</td><td>${eok(t.profitTarget)}</td><td>${eok(t.profitActual)}</td><td>${pct(tr.profitRate)}</td></tr>
          <tr><td colspan="3"><b>팀 평가 (평균)</b></td><td><b>${pct(tr.rate)}</b> ${gradeChip(tr.grade)}</td></tr>
        </tbody></table>
        ${bandRow(s.teamBands, tr.grade)}</div>

      <div class="sheet-section"><h3>감점요인 확인</h3>
        <table><thead><tr><th>지각</th><th>초과연차사용</th><th>교육미이수</th><th>합계</th></tr></thead><tbody>
          <tr><td>${d.late || 0}회 × ${trim(u.late * 100)}%p</td><td>${d.leave || 0}회 × ${trim(u.leave * 100)}%p</td><td>${d.training || 0}회 × ${trim(u.training * 100)}%p</td><td><b>${deductLabel(r.deduction)}</b></td></tr>
        </tbody></table></div>

      <div class="sheet-section"><h3>개인평가</h3>
        <table><thead><tr><th>본인평가</th><th>부문장평가</th><th>최종평가</th><th>감점</th><th>개인 점수</th><th>개인 등급</th></tr></thead><tbody>
          <tr><td${hl(r.kpiSource === 'self')}>${pct(ev.selfRate)}</td><td${hl(r.kpiSource === 'head')}>${pct(ev.headRate)}</td><td${hl(r.kpiSource === 'final')}>${pct(ev.finalRate)}</td>
            <td>${deductLabel(r.deduction)}</td><td><b>${pct(r.personalScore)}</b></td><td>${gradeChip(r.personalGrade)}</td></tr>
        </tbody></table>
        ${bandRow(s.personalBands, r.personalGrade)}
        <p class="formula">개인 점수 = 강조된 평가 점수 − 감점 합계</p></div>

      <div class="sheet-section"><h3>팀/개인 종합등급</h3>
        <table><thead><tr><th>개인 / 팀</th>${GRADES.map((g) => `<th>${g}</th>`).join('')}</tr></thead><tbody>
          ${GRADES.map((p) => `<tr><th>${p}</th>${GRADES.map((tg) => `<td${hl(p === r.personalGrade && tg === r.teamGrade)}>${s.combinedMatrix[p][tg]}</td>`).join('')}</tr>`).join('')}
        </tbody></table>
        <p class="formula">팀 ${r.teamGrade || '—'} × 개인 ${r.personalGrade || '—'} → 종합 <b>${r.finalGrade || '—'}</b></p></div>

      <div class="sheet-section"><h3>연봉 수준별 인상률</h3>
        <table><thead><tr><th>직급</th><th>종합등급</th><th>현재 연봉</th><th>업계평균</th><th>연봉 수준</th></tr></thead><tbody>
          <tr><td>${esc(e.level)}</td><td>${gradeChip(r.finalGrade)}</td><td>${won(e.salary)}</td><td>${won(r.market)}</td><td><b>${pct(r.compa)}</b></td></tr>
        </tbody></table>
        <table style="margin-top:8px"><thead><tr><th>연봉 수준</th>${GRADES.map((g) => `<th>${g} (σ ${pctShort(s.baseRates[g])})</th>`).join('')}</tr></thead><tbody>
          ${s.compaMatrix.map((row, i) => `<tr><td>${labels[i]}</td>${GRADES.map((g) => `<td${hl(i === r.compaBand && g === r.finalGrade)}>${g === 'P' && s.pAbsolute ? signed(row[g]) : sigmaLabel(row[g])}</td>`).join('')}</tr>`).join('')}
        </tbody></table></div>

      <div class="sheet-section"><h3>${esc(s.cycleName)} 연봉 인상률</h3>
        <table><thead><tr><th>현재 연봉</th><th>산출 인상률</th><th>부사장 검토</th><th>최종 인상률</th><th>인상 연봉</th></tr></thead><tbody>
          <tr><td>${won(e.salary)}</td><td>${pct(r.calcRaise)}</td><td>${r.overridden ? pct(e.vpOverride) : '—'}</td><td><b>${pct(r.finalRaise)}</b></td><td><b>${won(r.newSalary)}</b></td></tr>
        </tbody></table>
        ${r.overridden ? `<p class="formula">조정 사유: ${esc(e.vpReason) || '<span class="neg">사유 없음</span>'}</p>` : ''}
        ${r.complete ? '' : `<p class="formula neg">산출 불가: ${r.missing.join(', ')} 없음</p>`}</div>
    </div></div>`;
  }

  // ---------- 엑셀 ----------
  function loadXlsx() {
    if (window.XLSX) return Promise.resolve(window.XLSX);
    return new Promise((resolve, reject) => {
      const el = document.createElement('script');
      el.src = XLSX_URL;
      el.onload = () => resolve(window.XLSX);
      el.onerror = () => reject(new Error('엑셀 라이브러리를 불러오지 못했습니다. 인터넷 연결을 확인해 주세요.'));
      document.head.appendChild(el);
    });
  }

  function excelDate(v) {
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === 'number') return new Date(Date.UTC(1899, 11, 30) + Math.round(v) * 86400000).toISOString().slice(0, 10);
    if (typeof v === 'string' && /^\d{4}-\d{2}-\d{2}/.test(v)) return v.slice(0, 10);
    return '';
  }

  function pickFile(accept) {
    return new Promise((resolve) => {
      const el = document.createElement('input');
      el.type = 'file';
      el.accept = accept;
      el.onchange = () => resolve(el.files[0] || null);
      el.click();
    });
  }

  // 엑셀 평가지의 '대상자' 시트(머리글: 성명, 부서, 팀, 직급, 현재연봉, …)를 대상자·팀 목록으로 바꾼다
  function parseRosterWorkbook(XLSX, wb, existingTeams) {
    const sheetName = wb.SheetNames.find((n) => n.trim() === '대상자');
    if (!sheetName) throw new Error("'대상자' 시트를 찾지 못했습니다.");
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, raw: true, defval: null });
    const hi = rows.findIndex((r) => r && r.some((c) => c != null && String(c).trim() === '성명'));
    if (hi < 0) throw new Error("'성명' 머리글을 찾지 못했습니다.");
    const header = rows[hi].map((c) => (c == null ? '' : String(c).replace(/\s/g, '')));
    const col = (name) => header.indexOf(name);
    const c = { name: col('성명'), division: col('부서'), team: col('팀'), level: col('직급'), salary: col('현재연봉'),
      hire: col('입사일'), lastRaise: col('기존인상월'), evalDate: col('평가기준일'), eligible: col('대상자') };

    const data = rows.slice(hi + 1).filter((r) => r && r[c.name] != null && String(r[c.name]).trim());
    if (!data.length) throw new Error('가져올 대상자가 없습니다.');

    const teamsList = existingTeams.slice();
    const teamFor = (division, name) => {
      let t = teamsList.find((x) => x.name === name && x.division === division);
      if (!t) {
        t = { id: 't' + Date.now().toString(36) + teamsList.length, division, name, goal: '',
          revenueTarget: null, revenueActual: null, profitTarget: null, profitActual: null };
        teamsList.push(t);
      }
      return t;
    };

    const at = (r, i) => (i >= 0 ? r[i] : null);
    const employees = data.map((r, i) => {
      const division = String(at(r, c.division) || '').trim();
      const team = teamFor(division, String(at(r, c.team) || '(미지정)').trim());
      const salary = Number(at(r, c.salary));
      const flag = String(at(r, c.eligible) || 'O').trim().toUpperCase();
      return {
        id: 'e' + Date.now().toString(36) + i,
        name: String(at(r, c.name)).trim(),
        division,
        teamId: team.id,
        level: String(at(r, c.level) || '').trim(),
        salary: isFinite(salary) && salary > 0 ? salary : null,
        hireDate: excelDate(at(r, c.hire)),
        lastRaiseDate: excelDate(at(r, c.lastRaise)),
        eligible: flag !== 'X',
        kpi: '',
        evaluation: { selfRate: null, headRate: null, finalRate: null },
        deductions: { late: 0, leave: 0, training: 0 },
        reputation: '',
        vpOverride: null,
        vpReason: ''
      };
    });
    return { employees, teams: teamsList, evalDate: excelDate(at(data[0], c.evalDate)) };
  }

  async function importXlsx() {
    const file = await pickFile('.xlsx,.xls');
    if (!file) return;
    try {
      const XLSX = await loadXlsx();
      const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
      const parsed = parseRosterWorkbook(XLSX, wb, state.demo ? [] : state.teams);
      const employees = parsed.employees;
      if (!confirm(`'${file.name}'에서 ${employees.length}명을 가져옵니다. 현재 대상자 ${state.employees.length}명은 교체됩니다. 계속할까요?`)) return;

      if (parsed.evalDate) state.settings.evalDate = parsed.evalDate;
      state.teams = parsed.teams;
      state.employees = employees;
      state.demo = false;
      save();
      render();
      const unknown = employees.filter((e) => !E.marketSalary(e.level, state.settings)).length;
      toast(`${employees.length}명을 가져왔습니다.` + (unknown ? ` 업계평균이 없는 직급 ${unknown}명` : ''));
    } catch (err) {
      alert(err.message);
    }
  }

  function resultRows() {
    const labels = compaLabels();
    return eligibleEmployees().map((e) => {
      const r = results.employees[e.id];
      const t = teamById(e.teamId) || {};
      return [e.name, e.division || t.division, t.name, e.level, e.kpi, r.kpiScore, r.deduction, r.personalScore, r.personalGrade,
        r.teamRate, r.teamGrade, r.finalGrade, e.salary, r.market, r.compa, r.compaBand != null ? labels[r.compaBand] : '',
        r.sigma, r.adjust, r.calcRaise, e.vpOverride, e.vpReason, r.finalRaise, r.newSalary, r.increase, r.missing.join(', ')];
    });
  }

  const RESULT_HEAD = ['성명', '사업부', '팀', '직급', '개인 KPI', '개인 KPI 점수', '감점', '개인 점수', '개인 등급',
    '팀 달성률', '팀 등급', '종합 등급', '현재연봉', '업계평균', '연봉 수준', '연봉 수준 구간',
    '기본 인상률(σ)', '조정', '산출 인상률', '부사장 조정 인상률', '조정 사유', '최종 인상률', '인상 연봉', '증가액', '미입력 항목'];
  const RESULT_PCT_COLS = [5, 6, 7, 9, 14, 16, 17, 18, 19, 21];

  async function exportXlsx() {
    const fileBase = `${state.settings.cycleName} 평가결과`;
    const rows = resultRows();
    try {
      const XLSX = await loadXlsx();
      const ws = XLSX.utils.aoa_to_sheet([RESULT_HEAD].concat(rows));
      rows.forEach((row, ri) => {
        RESULT_PCT_COLS.forEach((ci) => {
          const cell = ws[XLSX.utils.encode_cell({ r: ri + 1, c: ci })];
          if (cell && typeof cell.v === 'number') cell.z = '0.0%';
        });
        [12, 13, 22, 23].forEach((ci) => {
          const cell = ws[XLSX.utils.encode_cell({ r: ri + 1, c: ci })];
          if (cell && typeof cell.v === 'number') cell.z = '#,##0';
        });
      });
      ws['!cols'] = RESULT_HEAD.map((h) => ({ wch: Math.max(10, h.length * 2) }));

      const teamHead = ['사업부', '팀', '팀 목표', '매출 목표(억)', '매출 실적(억)', '매출 달성률', '영업이익 목표(억)', '영업이익 실적(억)', '영업이익 달성률', '팀 달성률', '팀 등급'];
      const teamRows = state.teams.map((t) => {
        const r = results.teams[t.id];
        return [t.division, t.name, t.goal, t.revenueTarget, t.revenueActual, r.revenueRate, t.profitTarget, t.profitActual, r.profitRate, r.rate, r.grade];
      });
      const ws2 = XLSX.utils.aoa_to_sheet([teamHead].concat(teamRows));
      teamRows.forEach((row, ri) => [5, 8, 9].forEach((ci) => {
        const cell = ws2[XLSX.utils.encode_cell({ r: ri + 1, c: ci })];
        if (cell && typeof cell.v === 'number') cell.z = '0.0%';
      }));

      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, '개인 평가결과');
      XLSX.utils.book_append_sheet(wb, ws2, '팀 평가');
      XLSX.writeFile(wb, fileBase + '.xlsx');
    } catch (err) {
      // 오프라인이면 CSV로 대신 내보낸다
      const csv = [RESULT_HEAD].concat(rows).map((r) => r.map((v) => {
        const s = v == null ? '' : String(v);
        return /[",\n]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s;
      }).join(',')).join('\n');
      download(fileBase + '.csv', '﻿' + csv, 'text/csv;charset=utf-8');
      toast('엑셀 라이브러리를 불러오지 못해 CSV로 저장했습니다.');
    }
  }

  function download(name, content, type) {
    const url = URL.createObjectURL(new Blob([content], { type }));
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    a.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  // ---------- 동작 ----------
  function toast(msg) {
    const el = document.createElement('div');
    el.className = 'toast';
    el.textContent = msg;
    document.body.appendChild(el);
    setTimeout(() => el.remove(), 2600);
  }

  const actions = {
    'open-sheet': (id) => { ui.modal = id; render(); },
    'close-modal': () => { ui.modal = null; render(); },
    'print': () => window.print(),
    'add-team': () => {
      const division = state.teams.length ? state.teams[state.teams.length - 1].division : '';
      state.teams.push({ id: 't' + Date.now().toString(36), division, name: '새 팀', goal: '',
        revenueTarget: null, revenueActual: null, profitTarget: null, profitActual: null });
      save(); render();
    },
    'del-team': (id) => {
      const n = state.employees.filter((e) => e.teamId === id).length;
      if (n) { alert(`소속 대상자 ${n}명이 있어 삭제할 수 없습니다. 먼저 대상자의 팀을 바꿔 주세요.`); return; }
      state.teams = state.teams.filter((t) => t.id !== id);
      save(); render();
    },
    'add-emp': () => {
      const t = state.teams[0];
      state.employees.unshift({ id: 'e' + Date.now().toString(36), name: '새 대상자', division: t ? t.division : '', teamId: t ? t.id : '',
        level: '사원1', salary: null, hireDate: '', lastRaiseDate: '', eligible: true, kpi: '',
        evaluation: { selfRate: null, headRate: null, finalRate: null }, deductions: { late: 0, leave: 0, training: 0 },
        reputation: '', vpOverride: null, vpReason: '' });
      save(); render();
    },
    'del-emp': (id) => {
      const e = state.employees.find((x) => x.id === id);
      if (!e || !confirm(`${e.name} 님을 대상자에서 삭제할까요?`)) return;
      state.employees = state.employees.filter((x) => x.id !== id);
      save(); render();
    },
    'import-xlsx': importXlsx,
    'export-xlsx': exportXlsx,
    'backup': () => download(`${state.settings.cycleName} 백업.json`, JSON.stringify(state, null, 2), 'application/json'),
    'restore': async () => {
      const file = await pickFile('.json');
      if (!file) return;
      try {
        const data = normalize(JSON.parse(await file.text()));
        if (!Array.isArray(data.employees) || !Array.isArray(data.teams)) throw new Error();
        if (!confirm('현재 데이터를 백업 파일 내용으로 바꿉니다. 계속할까요?')) return;
        state = data; save(); render(); toast('백업을 불러왔습니다.');
      } catch (e) { alert('올바른 백업 파일이 아닙니다.'); }
    },
    'reset-settings': () => {
      if (!confirm('평가 기준을 엑셀 평가지 기본값으로 되돌립니다. 회차 이름, 평가기준일, 전사 실적은 유지됩니다.')) return;
      const keep = { cycleName: state.settings.cycleName, evalDate: state.settings.evalDate, company: state.settings.company };
      state.settings = Object.assign(HR.data.defaultSettings(), keep);
      save(); render();
    },
    'load-demo': () => {
      if (!confirm('현재 데이터를 지우고 데모 데이터(가상 인물 110명)를 불러옵니다. 계속할까요?')) return;
      state = HR.data.demoState(); save(); render();
    },
    'clear-all': () => {
      if (!confirm('대상자와 팀을 모두 삭제합니다. 되돌릴 수 없습니다. 계속할까요?')) return;
      state = { version: 2, demo: false, settings: HR.data.defaultSettings(), teams: [], employees: [] };
      save(); render();
    }
  };

  document.addEventListener('click', (ev) => {
    const nav = ev.target.closest('.nav-btn');
    if (nav) { ui.view = nav.dataset.view; saveUi(); render(); window.scrollTo(0, 0); return; }
    const btn = ev.target.closest('[data-action]');
    if (!btn) return;
    if (btn.dataset.action === 'close-modal-bg') {
      if (ev.target === btn) actions['close-modal']();
      return;
    }
    const fn = actions[btn.dataset.action];
    if (fn) fn(btn.dataset.id);
  });

  document.addEventListener('change', (ev) => {
    const el = ev.target;
    if (el.dataset.ui && el.tagName === 'SELECT') { ui[el.dataset.ui] = el.value; render(); return; }
    if (!el.dataset.path) return;
    const target = targetOf(el.dataset.kind, el.dataset.id);
    if (!target) return;
    const value = parseValue(el.dataset.type, el);
    if (value === undefined) { toast('숫자를 입력해 주세요.'); scheduleRender(); return; }
    setPath(target, el.dataset.path, value);
    // 팀을 바꾸면 사업부도 따라 바꾼다
    if (el.dataset.kind === 'emp' && el.dataset.path === 'teamId') {
      const t = teamById(value);
      if (t) target.division = t.division;
    }
    if (el.dataset.kind === 'team' && el.dataset.path === 'division') {
      state.employees.forEach((e) => { if (e.teamId === target.id) e.division = value; });
    }
    save();
    scheduleRender();
  });

  document.addEventListener('input', (ev) => {
    if (ev.target.dataset.ui === 'q') { ui.q = ev.target.value; render(); }
  });

  document.addEventListener('keydown', (ev) => {
    if (ev.key === 'Escape' && ui.modal) actions['close-modal']();
    if (ev.key === 'Enter' && ev.target.classList.contains('cell-input')) ev.target.blur();
  });

  // 테스트 페이지에서 가져오기 로직을 검증할 수 있도록 노출
  HR.app = { parseRosterWorkbook, loadXlsx };

  render();
})();
