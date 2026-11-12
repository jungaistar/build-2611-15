/* 계산 규칙 고정 — node test/logic.test.js
 * 화면이 아니라 숫자와 판정이 틀리면 안 되는 도구다. 손으로 셀 수 있는 값과 대조한다. */
'use strict';
var C = require('../js/costidx.js'), L = require('../js/linkage.js');
var passed = 0, failed = 0;
function ok(c, label, detail) { if (c) passed++; else { failed++; console.log('  X ' + label + (detail != null ? '\n      ' + detail : '')); } }
function eq(g, w, label) { ok(JSON.stringify(g) === JSON.stringify(w), label, '기대: ' + JSON.stringify(w) + '  실제: ' + JSON.stringify(g)); }
function near(g, w, label, tol) { ok(typeof g === 'number' && Math.abs(g - w) <= (tol || 1e-6), label, '기대: ' + w + '  실제: ' + g); }

/* ── 손으로 만든 작은 Index ── */
function ser(name, unit, sheet, start, vals, extra) {
  var i0 = C.mIdx(start);
  return Object.assign({ name: name, role: 'index', family: C.indexFamily(name), grades: C.gradeTokens(name), points: vals.map(function (v, i) { return v == null ? null : { m: C.mStr(i0 + i), v: v }; }).filter(Boolean), source: 'test / ' + sheet, meta: { sheet: sheet, unit: unit, kind: '' } }, extra || {});
}
var steel = ser('철강 후판 SS400', 'KRW/kg', '테스트정밀', '2026-01', [1000, 1000, 1000, 1100, 1100, 1100]);     // +10%
var cu = ser('구리 전기동', 'KRW/ton', '비철금속', '2026-01', [10000000, 10000000, 10000000, 10500000, 10500000, 10500000]);   // +5%, 원/ton
var cuUsd = ser('구리 전기동 LME', 'USD/ton', '비철금속', '2026-01', [9000, 9000, 9000, 9900, 9900, 9900]);
var fx = ser('USD · KRW/USD', 'KRW/USD', '환율', '2026-01', [1400, 1400, 1400, 1372, 1372, 1372], { role: 'skip', meta: { sheet: '환율', unit: 'KRW/USD', kind: 'fx' } });   // −2%
var rubber = ser('고무 천연고무', 'KRW/kg', '기타', '2026-01', [2000, 2000, 2000, 2020, null, null]);            // +1% (변경 기간 1개월만)
var series = [steel, cu, cuUsd, fx, rubber];
var vendor = { id: 'v1', name: '테스트정밀(주)', sheetHint: '' };
function item(name, o) { return Object.assign(L.blankItem(name), o || {}); }
function contract(items, o) { return Object.assign(L.blankContract('v1'), { from: '2026-01', to: '2026-12', cycle: 'quarter', items: items }, o || {}); }
function part(no, base, usage, o) { return Object.assign(L.blankPart('v1'), { partNo: no, basePrice: base, usage: usage || {} }, o || {}); }
var P = { base: { from: '2026-01', to: '2026-03' }, next: { from: '2026-04', to: '2026-06' } };
function run(c, parts, o) { return L.calcRun(Object.assign({ vendor: vendor, contract: c, parts: parts, series: series, period: P, applyMonth: '2026-07' }, o || {})); }

/* 1. 숫자 읽기 — 빈 칸과 0 구분 */
eq([L.cellNum(''), L.cellNum(null), L.cellNum(0), L.cellNum('0'), L.cellNum('1,250'), L.cellNum('-')], [null, null, 0, 0, 1250, null], '빈 칸은 null, 0 은 0');
ok(isNaN(L.cellNum('abc')) && isNaN(L.cellNum('12kg')), '숫자가 아니면 NaN');
ok(L.isFxName('환율') && L.isFxName('USD 환율') && !L.isFxName('철강') && !L.isFxName('구리'), '환율 이름 알아보기');

/* 2. 품번 엑셀 — 머리글 · 열 순서 · 동적 원소재 열 */
var aoa = [['비고'], ['구리(kg)', '품번', '품명', '철강(kg)', '기준 파트단가', '현재 구매단가', '고무 (kg)', '환율 적용', '환율 대상 금액'],
  [3, 'A-1', '브래킷', 25, '100,000', 101000, 2, '', ''], [0, 'A-2', '', 30, 150000, '', '', 'Y', 20000], ['x', 'A-3', '', 5, 60000, '', 1, '', ''], [1, 'A-4', '', -2, 60000, '', 1, '', ''],
  [1, 'A-1', '', 2, 60000, '', 1, '', ''], [1, '', '', 2, 60000, '', 1, '', ''], [1, 'A-7', '', 2, '', '', 1, '', ''], ['', '', '', '', '', '', '', '', '']];
var insp = L.inspectParts(aoa);
ok(insp.ok && insp.headerRow === 1, '머리줄을 2행에서 찾는다');
eq(insp.columns.map(function (c) { return c.role + (c.material ? ':' + c.material : ''); }), ['usage:구리', 'partNo', 'name', 'usage:철강', 'basePrice', 'curPrice', 'usage:고무', 'fx', 'fxAmount'], '열 순서가 달라도 역할을 알아본다');
var pr = L.parseParts(aoa, insp, ['a-2']);
eq(pr.materials, ['구리', '철강', '고무'], '원소재 열은 파일에 있는 대로');
eq(pr.counts, { 'new': 1, update: 1, error: 5 }, '신규 1 · 기존 1 · 오류 5(빈 행은 건너뜀)');
eq([pr.rows[0].basePrice, pr.rows[0].curPrice, pr.rows[0].usage], [100000, 101000, { '구리': 3, '철강': 25, '고무': 2 }], '값 읽기(쉼표 숫자)');
eq([pr.rows[1].status, pr.rows[1].usage['구리'], pr.rows[1].usage['고무'], pr.rows[1].fx, pr.rows[1].fxAmount, pr.rows[1].curPrice], ['update', 0, null, true, 20000, null], '기존 품번 · 0 과 빈 칸 구분 · 환율 적용');
ok(/숫자가 아닙니다\(x\)/.test(pr.rows[2].errors[0].msg) && pr.rows[2].errors[0].col === '구리(kg)', '숫자가 아닌 칸을 열 이름과 함께 짚는다');
ok(/음수/.test(pr.rows[3].errors[0].msg), '음수 중량');
ok(/겹칩니다\(3행\)/.test(pr.rows[4].errors[0].msg), '파일 안 중복 품번');
ok(/품번이 비어/.test(pr.rows[5].errors[0].msg) && /기준 파트단가가 비어/.test(pr.rows[6].errors[0].msg), '품번 · 기준 파트단가 누락');
ok(!L.inspectParts([['a', 'b'], [1, 2]]).ok, '품번 칸이 없으면 머리줄을 못 찾았다고 알린다');
var bad = L.inspectParts([['품번', '기준 파트단가', '환율(kg)'], ['A', 1, 1]]);
ok(!L.parseParts([['품번', '기준 파트단가', '환율(kg)'], ['A', 1, 1]], bad, []).ok, '환율 사용량 열은 받지 않는다');
ok(L.partsTemplate(['철강', '환율'])[0].join().indexOf('환율(kg)') < 0, '양식에 환율(kg) 열을 만들지 않는다');
ok(L.checkPart(part('A', 100, { '환율': 3 })).length === 1 && L.checkPart(part('A', 100, { '철강': 3 })).length === 0, '환율에 kg 을 넣으면 오류');

/* 3. 수식 계산기 */
near(L.evalExpr('KG * (I1 - I0)', { KG: 25, I0: 1000, I1: 1100 }), 2500, '사용자 정의 수식');
near(L.evalExpr('P0 × S × R ÷ 2', { P0: 100000, S: 0.4, R: 0.1 }), 2000, '× ÷ 기호');
near(L.evalExpr('-(2+3)*4', {}), -20, '괄호 · 부호');
['alert(1)', 'KG;1', 'X*2', '(1+2', '1/0', ''].forEach(function (e) { var t = false; try { L.evalExpr(e, { KG: 1 }); } catch (x) { t = true; } ok(t, '쓸 수 없는 수식은 멈춘다 — ' + e); });
var thrown = ''; try { L.evalExpr('KG*2', {}); } catch (x) { thrown = x.message; } ok(/KG.*값이 없습니다/.test(thrown), '값이 없는 이름');

/* 4. 기간 평균 · 단위 */
var a = L.avgOver(rubber, '2026-04', '2026-06');
ok(a.ok && a.v === 2020 && a.used.join() === '2026-04' && a.missing.join() === '2026-05,2026-06', '있는 달만 평균하고 빈 달을 알려 준다');
ok(!L.avgOver(rubber, '2027-01', '2027-03').ok, '기간에 값이 없으면 다른 달 값을 끌어오지 않는다');
eq([L.perKgFactor('KRW/kg'), L.perKgFactor('KRW/ton'), L.perKgFactor('KRW/kg(환산)'), L.perKgFactor('원/kg'), L.perKgFactor('USD/ton'), L.perKgFactor('')], [1, 0.001, 1, 1, null, null], '원/kg 환산 곱수');
eq(L.suggestPeriod('2026-07', 'quarter'), { next: { from: '2026-04', to: '2026-06' }, base: { from: '2026-01', to: '2026-03' } }, '분기 기간 제안');
eq(L.suggestPeriod('2026-07', 'half'), { next: { from: '2026-01', to: '2026-06' }, base: { from: '2025-07', to: '2025-12' } }, '반기 기간 제안');

/* 5. 계약 조건 → 적용 변동률 */
eq(L.applyTerms(4, item('철강', { minChange: 5 })).rate, 0, '최소 변동률 미만은 0');
ok(L.applyTerms(4, item('철강', { minChange: 5 })).excluded, '연동 제외 표시');
eq(L.applyTerms(4, item('철강'), { trigger: 5 }).rate, 0, '항목 값이 없으면 계약의 조정요건(±%)');
eq(L.applyTerms(4, item('철강', { minChange: 3 }), { trigger: 5 }).rate, 4, '항목별 최소 변동률이 조정요건보다 먼저');
eq(L.applyTerms(-6, item('철강'), { trigger: 5 }).rate, -6, '하락도 절댓값으로 조정요건 판단');
eq(L.applyTerms(12, item('철강', { cap: 10, floor: -10 })).rate, 10, '상한');
eq(L.applyTerms(-15, item('철강', { cap: 10, floor: -10 })).rate, -10, '하한');
eq(L.applyTerms(10, item('철강', { reflect: 50 })).rate, 5, '반영률');
eq(L.applyTerms(8, item('철강', { tiers: [[5, 50], [10, 100]] })).rate, 4, '구간별 차등 — 5% 이상 50%');
eq(L.applyTerms(12, item('철강', { tiers: [[10, 100], [5, 50]] })).rate, 12, '구간별 차등 — 10% 이상 100%');
eq(L.applyTerms(-2, Object.assign(item('환율'), { reflect: 30 })).rate, -2, '환율은 반영률을 무시하고 100%');

/* 6. Index 맞추기 */
var span = { from: '2026-01', to: '2026-06' };
var m1 = L.matchItem(item('철강', { spec: 'SS400' }), vendor, series, span);
eq([m1.status, m1.chosen, m1.where], ['exact', '철강 후판 SS400', 'vendor'], '업체 시트에서 정확 일치');
var m2 = L.matchItem(item('구리'), vendor, series, span);
eq([m2.status, m2.candidates.length, m2.needApproval], ['multi', 2, true], '정확 일치가 여럿이면 고르지 않는다');
ok(/자동으로 쓰지 않았습니다/.test(m2.notes.join()), '업체 시트에 소재가 없으면 그 시트를 자동 적용하지 않는다고 알린다');
var m3 = L.matchItem(item('구리', { indexName: '구리 전기동', indexStatus: 'exact' }), vendor, series, span);
eq([m3.status, m3.chosen], ['approved', '구리 전기동'], '담당자 승인 Index');
var m4 = L.matchItem(item('환율', { spec: 'USD' }), vendor, series, span);
eq([m4.status, m4.chosen], ['exact', 'USD · KRW/USD'], '환율은 환율 계열에서만');
var m5 = L.matchItem(item('티타늄'), vendor, series, span);
eq([m5.status, m5.needApproval], ['none', true], '모르는 소재는 후보를 지어내지 않는다');
var sb = L.sample(), sbk = C.buildBook(sb.book, C.inspectBook(sb.book), 'x.xlsx').series;
var m6 = L.matchItem(sb.contracts[1].items[1], sb.vendors[1], sbk, { from: '2023-11', to: '2026-10' });
ok(m6.status === 'similar' && m6.candidates.length >= 2 && m6.candidates.length <= 3 && m6.candidates[0].score >= m6.candidates[1].score && m6.candidates[0].corr != null && m6.candidates[0].complete > 0.9, '정확 일치 없음 → 유사 후보 최대 3개(점수 · 상관 · 완전성)');
ok(L.sheetMatches({ meta: { sheet: '가상정밀' } }, { name: '가상정밀(주)' }) && L.sheetMatches({ meta: { sheet: 'ABC' } }, { name: '다른이름', sheetHint: 'abc' }) && !L.sheetMatches({ meta: { sheet: '철판' } }, { name: '가상정밀(주)' }), '업체 시트 알아보기((주) 무시 · 직접 지정)');

/* 7. 계산 — 중량 연동식 */
var cW = contract([item('철강', { spec: 'SS400', formula: 'weight' }), item('구리', { formula: 'weight', indexName: '구리 전기동', indexStatus: 'exact' })]);
var r = run(cW, [part('A-1', 100000, { '철강': 25, '구리': 3 }), part('A-2', 50000, { '철강': 10, '구리': 0 })]);
ok(r.ok && r.done === 2, '두 품번 계산');
near(r.items[0].rawRate, 10, '철강 +10%'); near(r.items[1].rawRate, 5, '구리 +5%');
near(r.items[1].i0kg, 10000, '원/ton → 원/kg');
near(r.rows[0].lines[0].amount, 2500, '철강 25kg × 1000원 × 10% = 2,500');
near(r.rows[0].lines[1].amount, 1500, '구리 3kg × 10,000원 × 5% = 1,500');
near(r.rows[0].total, 4000, '전체 연동액'); near(r.rows[0].newPrice, 104000, '변경 파트단가'); near(r.rows[0].rate, 4, '변동률');
near(r.rows[1].total, 1000, '품번마다 사용량이 달라 연동액도 다르다'); near(r.rows[1].newPrice, 51000, 'A-2 변경단가');
eq(r.rows[1].lines[1].note, '사용량 0', '사용량 0 은 0원(오류 아님)');

/* 8. 계산 — 비중 · 사용자 정의 · 환율 · 끝수 */
var cR = contract([item('철강', { spec: 'SS400', formula: 'ratio', share: 40, cap: 8 })], { round: '10' });
r = run(cR, [part('B-1', 123456, {})]);
near(r.rows[0].total, 123456 * 0.4 * 0.08, '비중 연동식 = 기준단가 × 40% × 상한 8%');
eq(r.rows[0].newPrice, Math.round((123456 + 123456 * 0.032) / 10) * 10, '10원 단위 반올림');
var cC = contract([item('철강', { spec: 'SS400', formula: 'custom', expr: 'KG * (I1 - I0) * 0.5' })]);
r = run(cC, [part('C-1', 100000, { '철강': 20 })]);
near(r.rows[0].total, 20 * 100 * 0.5, '사용자 정의 수식');
var cF = contract([item('철강', { spec: 'SS400', formula: 'weight' }), item('환율', { spec: 'USD', fxBase: 'amount' })]);
r = run(cF, [part('D-1', 100000, { '철강': 10 }, { fx: true, fxAmount: 20000 }), part('D-2', 100000, { '철강': 10 }), part('D-3', 100000, { '철강': 10 }, { fx: true })]);
near(r.rows[0].fxAmountDelta, -400, '환율 대상 금액 20,000 × −2% = −400');
near(r.rows[0].matAmount, 1000, '원소재 연동액과 환율 연동액을 따로 센다'); near(r.rows[0].total, 600, '합계');
eq([r.rows[1].lines[1].amount, r.rows[1].lines[1].note], [0, '환율 미적용 품번'], '환율을 적용하지 않는 품번');
ok(r.rows[2].newPrice == null && /환율 대상 금액이 없습니다/.test(r.rows[2].blocks.join()), '환율 대상 금액이 없으면 멈춘다');
eq(r.rows[0].lines[1].kg, '-', '환율 사용량은 「-」');
cF.items[1].fxBase = 'price'; r = run(cF, [part('D-1', 100000, { '철강': 0 }, { fx: true })]);
near(r.rows[0].total, -2000, '환율 × 기준 파트단가 전체');

/* 9. 멈추는 조건 — 지어내지 않는다 */
ok(/수식이 정해지지 않았습니다/.test(run(contract([item('철강', { spec: 'SS400' })]), [part('E', 1000, { '철강': 1 })]).blocks[0].msg), '수식 미정이면 전체 중단');
ok(/두 번 등록/.test(L.contractProblems(contract([item('철강', { formula: 'weight' }), item('철 강', { formula: 'weight' })])).join()), '같은 원소재 중복');
ok(/비중\(%\)이 없습니다/.test(L.contractProblems(contract([item('철강', { formula: 'ratio' })])).join()), '비중 누락');
ok(/무엇에 곱하는지/.test(L.contractProblems(contract([item('환율')])).join()), '환율 기준 미정');
ok(/100% 입니다/.test(L.contractProblems(contract([Object.assign(item('환율', { fxBase: 'price' }), { reflect: 80 })])).join()), '환율 비중이 100% 가 아니면 중단');
r = run(cW, [part('F-1', 100000, { '철강': 25 }), part('F-2', null, { '철강': 1, '구리': 1 }), part('F-1', 100000, { '철강': 1, '구리': 1 })]);
ok(/구리 사용량\(kg\)이 비어/.test(r.rows[0].blocks.join()) && r.rows[0].newPrice == null, '사용량이 빈 칸이면 그 품번은 계산하지 않는다');
ok(/기준 파트단가가 비어/.test(r.rows[1].blocks.join()), '기준 파트단가 누락');
ok(/중복 등록/.test(r.rows[2].blocks.join()) && !r.ok && r.stopped === 3, '품번 중복');
var cM = contract([item('구리', { formula: 'weight' })]);
r = run(cM, [part('G', 100000, { '구리': 1 })]);
ok(r.rows[0].newPrice == null && /여럿/.test(r.rows[0].blocks.join()), '일치 Index 가 여럿이면 승인 전 계산 안 함');
var cU = contract([item('구리', { formula: 'weight', indexName: '구리 전기동 LME', indexStatus: 'exact' })]);
r = run(cU, [part('G', 100000, { '구리': 1 })]);
ok(/USD\/ton[\s\S]*중량 연동식에 쓸 수 없습니다/.test(r.rows[0].blocks.join()), 'Index 단위 불일치(USD/ton 을 중량식에)');
cU.items[0].formula = 'ratio'; cU.items[0].share = 30; r = run(cU, [part('G', 100000, {})]);
near(r.rows[0].total, 100000 * 0.3 * 0.1, '비중 연동식은 단위와 무관(변동률만 사용)');
var cS = contract([item('주철', { spec: 'FC250', formula: 'ratio', share: 20, indexName: '철강 후판 SS400', indexStatus: 'similar' })]);
r = run(cS, [part('H', 100000, {})]);
ok(/대체 Index 사용 불허/.test(r.rows[0].blocks.join()), '유사 Index 를 승인해도 계약이 대체를 허용하지 않으면 멈춘다');
cS.allowSubstitute = true; r = run(cS, [part('H', 100000, {})]);
ok(r.ok && r.items[0].match.similar === true, '대체 허용 + 승인 → 계산, 유사 Index 표시');
r = run(contract([item('고무', { formula: 'weight' })]), [part('I', 100000, { '고무': 2 })]);
ok(r.ok && /값이 없는 달: 2026-05, 2026-06/.test(r.items[0].notes.join()), '빈 달이 있으면 알리고 있는 달만 평균');
r = run(contract([item('고무', { formula: 'weight' })]), [part('I', 100000, { '고무': 2 })], { period: { base: P.base, next: { from: '2026-08', to: '2026-09' } } });
ok(!r.ok && /변경 Index/.test(r.rows[0].blocks.join()), '변경 기간에 값이 없으면 멈춘다');
ok(/적용 기간.*밖/.test(run(cW, [part('A', 1, {})], { applyMonth: '2027-03' }).blocks[0].msg), '적용 월이 계약 버전 기간 밖이면 중단');
ok(/Index 엑셀/.test(run(cW, [part('A', 1, {})], { series: [] }).blocks[0].msg), 'Index 가 없으면 중단');
r = run(contract([item('철강', { spec: 'SS400', formula: 'weight', baseValue: 800 })]), [part('J', 100000, { '철강': 10 })], { period: { base: {}, next: P.next } });
near(r.rows[0].total, 10 * 800 * 0.375, '계약에 적힌 고정 기준 Index(800 → 1100, +37.5%)');
r = run(contract([item('철강', { spec: 'SS400', formula: 'weight' })]), [part('K', 10000, { '철강': 500 })]);
ok(/기준 파트단가보다 큽니다/.test(r.rows[0].warnings.join()), '원소재 금액이 단가보다 크면 중량 확인 알림');
r = run(contract([item('철강', { spec: 'SS400', formula: 'weight' })], { trigger: 15 }), [part('L', 100000, { '철강': 25 })]);
eq([r.rows[0].total, r.items[0].excluded], [0, true], '조정요건 미만이면 연동액 0(계산은 완료)');

/* 10. 업체 요청단가 비교 */
r = run(cW, [part('A-1', 100000, { '철강': 25, '구리': 3 }, { curPrice: 101000 })], { requests: { 'A-1': 106000 } });
var q = r.rows[0].request;
near(q.gap, 2000, '요청 − 계약 계산 = 106,000 − 104,000'); near(q.reqRate, (106000 / 101000 - 1) * 100, '요청 인상률(현재 구매단가 기준)'); near(q.calcRate, (104000 / 101000 - 1) * 100, '계약 계산 인상률');
ok(/높음/.test(q.verdict) && Math.abs(q.explained - 4000) < 1e-6, '원소재로 설명되는 금액과 요청을 구분');

/* 11. 계약서 글 읽기 */
var std = L.extractContract(['납품대금 연동 특별약정서(별첨)', '1. 물품등의 명칭 마스트 브래킷', '2. 주요 원재료 또는 주요 에너지 철강 / 구리, 환율', '3. 원재료 가격 또는 에너지 비용의 기준지표 구매기획팀 Index', '4. 원재료 기준가격 기준시점 : 1월 평균', '비교시점 : 직전 3개월 평균', '5. 조정요건 ±5% 이상', '6. 조정주기 반기', '7. 조정일 매 반기 첫 영업일', '8. 조정대금 반영시점 다음 달 납품분', '9. 납품대금 연동산식 A = B + C', '10. 기타사항 상한 99%'].join('\n'));
ok(std.standard && std.fields.trigger.value === 5 && std.fields.cycle.value === 'half' && std.fields.product.value === '마스트 브래킷', '표준 약정서 번호 항목');
eq(std.materials.map(function (x) { return x.name; }), ['철강', '구리', '환율'], '2번 항목의 원재료');
ok(/기준시점 : 1월 평균 비교시점/.test(std.fields.baseRule.value) && std.fields.formulaText.value === 'A = B + C' && !std.fields.cap, '4번은 이어지는 줄까지 · 10번 뒤 숫자는 조건으로 읽지 않는다');
var free = L.extractContract('공급자: 가상정밀(주)\n적용 기간: 2026.1.1 ~ 2026.12.31\n매 분기마다 조정한다\n원소재는 철강 60%, 구리 10% 비중으로 연동한다\n변동률이 3% 미만이면 연동하지 않는다. 상한 10%, 하한 -10%\n변경 단가 = 기준 단가 × (1 + 비중 × 변동률)');
eq([free.fields.from.value, free.fields.to.value, free.fields.cycle.value, free.fields.minChange.value, free.fields.cap.value, free.fields.floor.value], ['2026-01', '2026-12', 'quarter', 3, 10, -10], '자유 형식 계약서의 기간 · 주기 · 조건');
eq(free.materials.map(function (x) { return x.name + ':' + x.share; }), ['철강:60', '구리:10'], '원소재와 비중');
ok(free.formulaLines.length === 1 && L.extractContract('아무 내용').notes.length === 1, '수식 줄 · 못 찾으면 직접 입력 안내');

/* 12. 보고서 값 · 엑셀 */
var cRp = contract([item('철강', { spec: 'SS400', formula: 'weight' }), item('구리', { formula: 'weight' })], { product: '브래킷', trigger: 3 });
var res = run(cRp, [part('A-1', 100000, { '철강': 25, '구리': 3 }), part('A-2', 50000, { '철강': 1 })], { requests: { 'A-1': 105000 } });
var runObj = { at: '2026-11-12T00:00:00Z', vendorName: vendor.name, contract: cRp, period: P, applyMonth: '2026-07', indexFile: 'x.xlsx', result: res, opinion: '' };
var d = L.reportData(runObj);
eq([d['업체명'], d['기준기간'], d['변경기간'], d['조정요건'], d['검토의견'], d['반영시점']], ['테스트정밀(주)', '2026-01 ~ 2026-03', '2026-04 ~ 2026-06', '±3%', '확인 필요', '확인 필요'], '보고서 값 — 없는 값은 확인 필요');
ok(d.rows.length === 2 && d.rows[0]['변경단가'] === '확인 필요' && /확인 필요 —/.test(d.rows[0]['의견']), '계산하지 못한 품번은 확인 필요로');
ok(d.items.length === 2 && d.items[0]['변동률'] === '+10.00%' && d.details.length === 4, 'Index 표 · 내역 표');
var sh = L.exportSheets(runObj);
eq(Object.keys(sh), ['품번별결과', '계산내역', 'Index근거', '검증'], '엑셀 시트 4장');
ok(sh['품번별결과'].length === 7 && sh['계산내역'].length === 5, '엑셀 행 수');

/* 12-1. 가로형 Index(원소재가 행 · 날짜가 열) */
var wideBook = { names: ['Index', '메모'], sheets: { 'Index': [['원소재 이름', '규격', '기준', '단위', '출처', '원본 시트', 46023, '2026-02', '2026.3', '2026년 4월', '26.5', '비고'],
  ['철강 후판', 'SS400', '유통가', 'KRW/kg', '가상', '테스트정밀', 1000, 1000, 1000, '1,100', 1100, '메모'], ['구리', '전기동', 'LME', 'USD/ton', '가상', '', 9000, '', 'x', 9900, 9900, ''], ['환율 USD', '', '', 'KRW/USD', '', '', 1400, 1400, 1400, 1372, 1372, ''], ['', '', '', '', '', '', 1, 1, 1, 1, 1, '']], '메모': [['아무 글']] } };
eq([L.headMonth(46023), L.headMonth('2026-02'), L.headMonth('2026.3'), L.headMonth('2026년 4월'), L.headMonth('26.5'), L.headMonth('비고'), L.headMonth(1000), L.headMonth('202612')], ['2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '', '', '2026-12'], '날짜 머리글 읽기(엑셀 날짜 숫자 포함)');
var wb = L.buildIndex(wideBook, 'w.xlsx', true), ws = wb.series;
eq(wb.sheets.map(function (x) { return x.sheet + ':' + x.kind + ':' + x.ok; }), ['Index:wide:true', '메모:undefined:false'], '가로형 시트 판별, 못 읽는 시트는 따로');
eq(ws.filter(function (x) { return !x.converted; }).map(function (x) { return x.name + '|' + x.role + '|' + x.points.length; }), ['철강 후판 · SS400 · 유통가 · KRW/kg|index|5', '환율 USD · KRW/USD|skip|5'], '행마다 계열 — 이름 · 규격 · 기준 · 단위');
eq(L.buildIndex(wideBook, 'w.xlsx', false).series.map(function (x) { return x.points.length; }), [5, 3, 5], '빈 칸 · 글자 칸은 빼고 읽는다(구리 3개월)');
ok(ws[0].meta.sheet === '테스트정밀' && L.sheetMatches(ws[0], vendor), '「원본 시트」 칸이 업체 시트 판정에 쓰인다');
ok(ws.some(function (x) { return x.converted && /구리/.test(x.name) && /KRW\/ton/.test(x.meta.unit); }), '같은 표의 환율 행으로 원화 환산 계열을 만든다');
ok(/숫자가 아닌 값 1칸/.test(wb.issues.map(function (i) { return i.msg; }).join()), '숫자가 아닌 칸을 알린다');
r = L.calcRun({ vendor: vendor, contract: contract([item('철강', { spec: 'SS400', formula: 'weight' }), item('환율', { spec: 'USD', fxBase: 'price' })]), parts: [part('W', 100000, { '철강': 10 }, { fx: true })], series: ws, period: { base: { from: '2026-01', to: '2026-03' }, next: { from: '2026-04', to: '2026-05' } }, applyMonth: '2026-06' });
ok(r.ok && Math.abs(r.rows[0].matAmount - 1000) < 1e-6 && Math.abs(r.rows[0].fxAmountDelta + 2000) < 1e-6, '가로형 Index 로 끝까지 계산');
ok(L.wideInfo([['날짜', '철강'], ['2026-01', 1]]) === null, '세로형 표는 가로형으로 보지 않는다');

/* 13. 예제 — 끝까지 */
var s = L.sample(), bk = C.buildBook(s.book, C.inspectBook(s.book), s.indexFile).series;
var r1 = L.calcRun({ vendor: s.vendors[0], contract: s.contracts[0], parts: s.parts.filter(function (p) { return p.vendorId === s.vendors[0].id; }), series: bk, period: L.suggestPeriod(s.applyMonth, 'quarter'), requests: s.requests, applyMonth: s.applyMonth });
eq([r1.done, r1.stopped, r1.items.map(function (o) { return o.match.status; }).join()], [3, 1, 'exact,exact,exact,exact'], '예제 가상정밀: 3개 계산 · 1개 확인 필요 · 전 항목 정확 일치');
ok(r1.items[0].match.where === 'vendor' && r1.items[1].match.where === 'other', '예제: 철강은 업체 시트, 구리는 다른 시트');
ok(r1.rows.every(function (x) { return x.newPrice == null || Math.abs(x.newPrice - Math.round(x.newPrice)) < 1e-9; }), '예제 계약은 원 단위 반올림');
ok(s.contracts.every(function (c) { return L.contractProblems(c).length === 0; }), '예제 계약은 계산 가능 상태');
var before = JSON.stringify(s.contracts[0]); L.calcRun({ vendor: s.vendors[0], contract: s.contracts[0], parts: s.parts, series: bk, period: L.suggestPeriod('2026-07', 'quarter') });
ok(JSON.stringify(s.contracts[0]) === before, '계산이 계약 · 품번 값을 바꾸지 않는다');

console.log('계산 규칙 테스트: 통과 ' + passed + ' · 실패 ' + failed);
process.exit(failed ? 1 : 0);
