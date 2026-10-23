/*
 * 평가 계산 엔진 — 화면과 분리된 순수 함수 모음.
 * 모든 비율은 소수로 다룬다 (115% → 1.15, 3% → 0.03).
 * 엑셀 평가지(2025년 4분기 평가지.xlsx)의 로직을 그대로 옮겼다. 근거는 docs/DESIGN.md 참고.
 */
(function (global) {
  'use strict';

  var GRADES = ['E', 'V', 'G', 'N', 'P'];

  function isNum(v) {
    return typeof v === 'number' && isFinite(v);
  }

  // 부동소수 오차로 1.21이 1.2099999…가 되어 등급이 내려가는 것을 막는다
  function round4(v) {
    return Math.round(v * 10000) / 10000;
  }

  // 실적 / 목표. 목표가 없거나 0이면 계산하지 않는다 (엑셀의 #DIV/0! 대응)
  function rate(actual, target) {
    if (!isNum(actual) || !isNum(target) || target === 0) return null;
    return actual / target;
  }

  // AVERAGE와 같이 빈 값은 건너뛴다
  function average(values) {
    var nums = values.filter(isNum);
    if (!nums.length) return null;
    return nums.reduce(function (a, b) { return a + b; }, 0) / nums.length;
  }

  // bands: [{ grade: 'E', min: 1.30 }, …] 높은 순. min 이상이면 해당 등급, 모두 미달이면 마지막 등급
  function gradeByScore(score, bands) {
    if (!isNum(score)) return null;
    var s = round4(score);
    for (var i = 0; i < bands.length; i++) {
      if (s >= round4(bands[i].min)) return bands[i].grade;
    }
    return bands[bands.length - 1].grade;
  }

  function teamResult(team, settings) {
    var revenueRate = rate(team.revenueActual, team.revenueTarget);
    var profitRate = rate(team.profitActual, team.profitTarget);
    var teamRate = average([revenueRate, profitRate]);
    return {
      revenueRate: revenueRate,
      profitRate: profitRate,
      rate: teamRate,
      grade: gradeByScore(teamRate, settings.teamBands)
    };
  }

  function deductionTotal(ded, settings) {
    var d = ded || {};
    var u = settings.deductionUnits;
    return (d.late || 0) * u.late + (d.leave || 0) * u.leave + (d.training || 0) * u.training;
  }

  // 최종평가 → 없으면 부문장평가 → 없으면 본인평가 순으로 개인 KPI 점수를 정한다
  function kpiScore(evaluation) {
    var e = evaluation || {};
    if (isNum(e.finalRate)) return { value: e.finalRate, source: 'final' };
    if (isNum(e.headRate)) return { value: e.headRate, source: 'head' };
    if (isNum(e.selfRate)) return { value: e.selfRate, source: 'self' };
    return { value: null, source: null };
  }

  // 연봉 수준 구간: compaBands = [1.40, 1.25, 0.95, 0.85] → 초과 기준, 모두 이하면 마지막 구간
  function compaBandIndex(compa, compaBands) {
    if (!isNum(compa)) return null;
    var c = round4(compa);
    for (var i = 0; i < compaBands.length; i++) {
      if (c > round4(compaBands[i])) return i;
    }
    return compaBands.length;
  }

  function marketSalary(level, settings) {
    var row = settings.market.filter(function (m) { return m.level === level; })[0];
    return row ? row.salary : null;
  }

  function stayDays(lastRaiseDate, evalDate) {
    if (!lastRaiseDate || !evalDate) return null;
    var a = Date.parse(lastRaiseDate);
    var b = Date.parse(evalDate);
    if (isNaN(a) || isNaN(b)) return null;
    return Math.round((b - a) / 86400000);
  }

  /*
   * 직원 1명의 전체 평가 결과.
   * team: teamResult()의 결과, emp: 직원 레코드, settings: 평가 설정
   */
  function employeeResult(emp, team, settings) {
    var kpi = kpiScore(emp.evaluation);
    var deduction = deductionTotal(emp.deductions, settings);
    var personalScore = isNum(kpi.value) ? kpi.value - deduction : null;
    var personalGrade = gradeByScore(personalScore, settings.personalBands);

    var teamGrade = team ? team.grade : null;
    var finalGrade = personalGrade && teamGrade ? settings.combinedMatrix[personalGrade][teamGrade] : null;

    var market = marketSalary(emp.level, settings);
    var compa = isNum(emp.salary) && isNum(market) && market > 0 ? emp.salary / market : null;
    var bandIdx = compaBandIndex(compa, settings.compaBands);

    var sigma = finalGrade ? settings.baseRates[finalGrade] : null;
    var adjust = finalGrade && bandIdx !== null ? settings.compaMatrix[bandIdx][finalGrade] : null;

    var calcRaise = null;
    if (isNum(sigma) && isNum(adjust)) {
      // P등급 조정값을 σ 대비가 아닌 절대 인상률로 볼지 여부 (DESIGN.md 5장 #3)
      calcRaise = finalGrade === 'P' && settings.pAbsolute ? adjust : sigma + adjust;
      if (isNum(settings.minRaise)) calcRaise = Math.max(calcRaise, settings.minRaise);
      calcRaise = round4(calcRaise);
    }

    var override = emp.vpOverride;
    var finalRaise = isNum(override) ? override : calcRaise;
    var newSalary = isNum(finalRaise) && isNum(emp.salary) ? Math.round(emp.salary * (1 + finalRaise)) : null;

    var missing = [];
    if (!team || !isNum(team.rate)) missing.push('팀 실적');
    if (!isNum(kpi.value)) missing.push('개인 평가');
    if (!isNum(market)) missing.push('업계평균');
    if (!isNum(emp.salary)) missing.push('현재연봉');

    return {
      kpiScore: kpi.value,
      kpiSource: kpi.source,
      deduction: deduction,
      personalScore: personalScore,
      personalGrade: personalGrade,
      teamRate: team ? team.rate : null,
      teamGrade: teamGrade,
      finalGrade: finalGrade,
      market: market,
      compa: compa,
      compaBand: bandIdx,
      sigma: sigma,
      adjust: adjust,
      calcRaise: calcRaise,
      finalRaise: finalRaise,
      overridden: isNum(override),
      newSalary: newSalary,
      increase: isNum(newSalary) ? newSalary - emp.salary : null,
      stayDays: stayDays(emp.lastRaiseDate, settings.evalDate),
      complete: missing.length === 0,
      missing: missing
    };
  }

  // 상태 전체를 받아 팀별·직원별 결과를 한 번에 계산한다
  function evaluateAll(state) {
    var s = state.settings;
    var teams = {};
    state.teams.forEach(function (t) { teams[t.id] = teamResult(t, s); });
    var results = {};
    state.employees.forEach(function (e) {
      results[e.id] = employeeResult(e, teams[e.teamId], s);
    });
    return { teams: teams, employees: results };
  }

  var api = {
    GRADES: GRADES,
    rate: rate,
    average: average,
    gradeByScore: gradeByScore,
    teamResult: teamResult,
    deductionTotal: deductionTotal,
    kpiScore: kpiScore,
    compaBandIndex: compaBandIndex,
    marketSalary: marketSalary,
    stayDays: stayDays,
    employeeResult: employeeResult,
    evaluateAll: evaluateAll
  };

  global.HR = global.HR || {};
  global.HR.engine = api;
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
})(typeof window !== 'undefined' ? window : globalThis);
