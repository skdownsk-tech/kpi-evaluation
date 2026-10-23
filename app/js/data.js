/*
 * 기본 설정값(엑셀 평가지 기준)과 가상 인물로 만든 데모 데이터.
 * 데모 데이터는 시드 고정 난수로 만들어 매번 같은 결과가 나온다.
 */
(function (global) {
  'use strict';

  function defaultSettings() {
    return {
      cycleName: '2025년 4분기',
      evalDate: '2025-12-31',
      // 전사 실적마감 (단위: 억원) — 평가표 ② 참고용
      company: { revenueTarget: 1000, revenueActual: 0, profitTarget: 300, profitActual: 0 },
      // 팀 평가 등급 구간 (이상 기준)
      teamBands: [
        { grade: 'E', min: 1.30 },
        { grade: 'V', min: 1.21 },
        { grade: 'G', min: 1.00 },
        { grade: 'N', min: 0.90 },
        { grade: 'P', min: 0 }
      ],
      // 개인 평가 등급 구간 (이상 기준)
      personalBands: [
        { grade: 'E', min: 1.30 },
        { grade: 'V', min: 1.21 },
        { grade: 'G', min: 1.00 },
        { grade: 'N', min: 0.80 },
        { grade: 'P', min: 0 }
      ],
      // 종합 등급: combinedMatrix[개인등급][팀등급]
      combinedMatrix: {
        E: { E: 'E', V: 'E', G: 'V', N: 'G', P: 'G' },
        V: { E: 'E', V: 'V', G: 'G', N: 'G', P: 'N' },
        G: { E: 'V', V: 'G', G: 'G', N: 'N', P: 'N' },
        N: { E: 'G', V: 'G', G: 'N', N: 'N', P: 'P' },
        P: { E: 'G', V: 'N', G: 'N', N: 'P', P: 'P' }
      },
      // 등급별 기본 인상률 σ
      baseRates: { E: 0.15, V: 0.10, G: 0.07, N: 0.05, P: 0.03 },
      // 연봉 수준 구간 경계 (초과 기준): 140% 초과 / 125% 초과 / 95% 초과 / 85% 초과 / 85% 이하
      compaBands: [1.40, 1.25, 0.95, 0.85],
      // 연봉 수준별 σ 조정값: compaMatrix[구간][등급]
      compaMatrix: [
        { E: 0.025, V: 0.01, G: -0.01, N: -0.02, P: -0.03 },
        { E: 0.03, V: 0.01, G: -0.01, N: -0.02, P: -0.03 },
        { E: 0.04, V: 0.02, G: 0, N: -0.01, P: -0.02 },
        { E: 0.05, V: 0.03, G: 0.01, N: 0, P: 0 },
        { E: 0.05, V: 0.05, G: 0.03, N: 0, P: 0 }
      ],
      // P등급 조정값은 σ 대비가 아니라 최종 인상률 그대로 사용한다 (시장 대비 고연봉 P등급은 연봉 삭감)
      pAbsolute: true,
      minRaise: null,
      // 감점 1회당 차감 점수 (%p를 소수로)
      deductionUnits: { late: 0.005, leave: 0.01, training: 0.02 },
      // 동종업계 직급별 평균 연봉 (원) — 업계평균 시트
      market: [
        ['사원1', 30474000], ['사원2', 31968000], ['사원3', 34452000],
        ['대리1', 36603000], ['대리2', 38835000], ['대리3', 41274000], ['대리4', 43749000],
        ['과장1', 45999000], ['과장2', 48249000], ['과장3', 50472000], ['과장4', 52857000],
        ['차장1', 55170000], ['차장2', 57744000], ['차장3', 60588000], ['차장4', 63675000], ['차장5', 66870000],
        ['부장1', 66177000], ['부장2', 68040000], ['부장3', 70200000], ['부장4', 73710000],
        ['실장1', 77395500], ['실장2', 81265275], ['실장3', 85328539], ['실장4', 89594966], ['실장5', 94074714],
        ['실장6', 98778450], ['실장7', 103717372], ['실장8', 108903241], ['실장9', 114348403]
      ].map(function (r) { return { level: r[0], salary: r[1] }; })
    };
  }

  // 시드 고정 난수 (mulberry32)
  function rng(seed) {
    var a = seed >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      var t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  function demoState() {
    var settings = defaultSettings();
    var r = rng(20251231);
    var pick = function (arr) { return arr[Math.floor(r() * arr.length)]; };
    var between = function (lo, hi) { return lo + r() * (hi - lo); };
    var round2 = function (v) { return Math.round(v * 100) / 100; };
    var pad = function (n) { return (n < 10 ? '0' : '') + n; };

    var structure = [
      ['브랜드 사업', ['브랜드마케팅팀', '국내사업팀', '해외사업팀', '온라인사업팀', '상품기획팀']],
      ['브랜드 운영', ['사업지원팀', '디자인팀', '물류팀', '경영지원팀', '인사총무팀']]
    ];
    var kpiByDivision = {
      '브랜드 사업': ['담당 브랜드 매출 목표 달성', '신규 거래처 확보', '채널별 매출 성장률', '캠페인 ROI 개선'],
      '브랜드 운영': ['운영비 절감 목표 달성', '업무 처리 리드타임 단축', '재고 회전율 개선', '내부 고객 만족도']
    };
    var surnames = '김이박최정강조윤장임한오서신권황안송류홍'.split('');
    var given = '민서지현수예도하준우진영은채유성소윤아연호재다희'.split('');
    var memberLevels = ['사원1', '사원2', '사원3', '대리1', '대리2', '대리3', '대리4',
      '과장1', '과장2', '과장3', '과장4', '차장1', '차장2', '차장3'];

    var teams = [];
    var employees = [];
    var used = {};
    var empNo = 0;

    structure.forEach(function (div, di) {
      div[1].forEach(function (teamName, ti) {
        var teamId = 't' + (di * 5 + ti + 1);
        // 팀별 성과 수준: 일부 팀은 초과 달성, 일부는 미달
        var perf = between(0.82, 1.38);
        var revenueTarget = Math.round(between(30, 120));
        var profitTarget = Math.round(revenueTarget * between(0.12, 0.25) * 10) / 10;
        teams.push({
          id: teamId,
          division: div[0],
          name: teamName,
          goal: teamName.replace('팀', '') + ' 연간 매출 ' + revenueTarget + '억 / 영업이익 ' + profitTarget + '억',
          revenueTarget: revenueTarget,
          revenueActual: round2(revenueTarget * perf * between(0.95, 1.05)),
          profitTarget: profitTarget,
          profitActual: round2(profitTarget * perf * between(0.9, 1.1))
        });

        for (var i = 0; i < 11; i++) {
          var name;
          do { name = pick(surnames) + pick(given) + pick(given); } while (used[name]);
          used[name] = true;
          empNo++;

          var level = i === 0 ? pick(['차장4', '차장5', '부장1', '부장2', '부장3']) : pick(memberLevels);
          var market = settings.market.filter(function (m) { return m.level === level; })[0].salary;
          // 연봉 수준이 80% ~ 145%에 고르게 퍼지도록
          var salary = Math.round(market * between(0.8, 1.45) / 10000) * 10000;

          var hireYear = 2010 + Math.floor(r() * 15);
          var raiseMonth = 1 + Math.floor(r() * 12);
          var self = round2(perf * between(0.95, 1.2));
          var head = round2(perf * between(0.82, 1.12));

          employees.push({
            id: 'e' + empNo,
            name: name,
            division: div[0],
            teamId: teamId,
            level: level,
            salary: salary,
            hireDate: hireYear + '-' + pad(1 + Math.floor(r() * 12)) + '-' + pad(1 + Math.floor(r() * 28)),
            lastRaiseDate: (raiseMonth > 9 ? 2023 : 2024) + '-' + pad(raiseMonth) + '-01',
            eligible: true,
            kpi: pick(kpiByDivision[div[0]]),
            evaluation: { selfRate: self, headRate: head, finalRate: null },
            deductions: {
              late: r() < 0.15 ? 1 + Math.floor(r() * 3) : 0,
              leave: r() < 0.08 ? 1 : 0,
              training: r() < 0.06 ? 1 : 0
            },
            reputation: '',
            vpOverride: null,
            vpReason: ''
          });
        }
      });
    });

    // 부사장 조정 예시 1건
    employees[3].vpOverride = 0.08;
    employees[3].vpReason = '핵심 인재 리텐션 (예시)';

    settings.company.revenueActual = round2(teams.reduce(function (s, t) { return s + t.revenueActual; }, 0));
    settings.company.revenueTarget = teams.reduce(function (s, t) { return s + t.revenueTarget; }, 0);
    settings.company.profitActual = round2(teams.reduce(function (s, t) { return s + t.profitActual; }, 0));
    settings.company.profitTarget = round2(teams.reduce(function (s, t) { return s + t.profitTarget; }, 0));

    return { version: 2, demo: true, settings: settings, teams: teams, employees: employees };
  }

  global.HR = global.HR || {};
  global.HR.data = { defaultSettings: defaultSettings, demoState: demoState };
})(typeof window !== 'undefined' ? window : globalThis);
