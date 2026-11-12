/*
 * build-2611-15 · OEM 원소재 연동제 — 계산 엔진
 * 계산 · 판정 · 검증은 전부 이 파일에 둔다(test/logic.test.js 로 고정). js/app.js 는 그리기와 저장만 한다.
 *
 * 가장 중요한 원칙: 업체별 계약(수식 · 조건)과 품번별 실제 데이터(기준 파트단가 · 원소재 사용량)를 나눠 둔다.
 * 모르는 값은 지어내지 않는다 — 수식이 불명확하거나 값이 비면 그 품번 계산을 멈추고 이유를 돌려준다.
 */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory(require('./costidx.js'));
  else root.Linkage = factory(root.CostIdx);
})(typeof self !== 'undefined' ? self : this, function (CostIdx) {
  'use strict';

  var FORMULA = { weight: '중량 연동식', ratio: '비중 연동식', custom: '사용자 정의 수식' };
  var CYCLE = { quarter: '분기', half: '반기', year: '연', etc: '기타' };
  var FX_BASE = { price: '기준 파트단가 전체', amount: '품번별 환율 대상 금액' };

  function uid(p) { return (p || 'id') + '-' + Date.now().toString(36) + Math.random().toString(36).slice(2, 7); }
  function txt(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
  function key(s) { return txt(s).toUpperCase().replace(/[\s\-_.·,()（）\[\]/]/g, ''); }
  function isNum(v) { return typeof v === 'number' && isFinite(v); }
  /** 빈 칸은 null, 숫자가 아니면 NaN — 빈 셀과 0 을 구분한다 */
  function cellNum(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : NaN;
    var s = String(v).replace(/[,\s원₩]/g, '');
    if (s === '' || s === '-') return null;
    if (!/^[-+]?(\d+\.?\d*|\.\d+)$/.test(s)) return NaN;
    return parseFloat(s);
  }
  function r2(v, d) { if (!isNum(v)) return null; var k = Math.pow(10, d == null ? 2 : d); return Math.round(v * k) / k; }
  function isFxName(name) { return /환율|FX|EXCHANGE|USD|달러|EUR|유로|JPY|엔화|CNY|RMB|위안/i.test(txt(name)); }
  function clone(o) { return JSON.parse(JSON.stringify(o)); }

  /* ───────────── 빈 틀 ───────────── */
  function blankVendor() { return { id: uid('v'), name: '', sheetHint: '', memo: '' }; }
  function blankItem(name) {
    var fx = isFxName(name);
    return {
      id: uid('m'), name: name || '', kind: fx ? 'fx' : 'material', spec: '',
      formula: fx ? 'fx' : '', share: null, expr: '',
      indexName: '', indexStatus: '', baseValue: null,
      reflect: 100, minChange: null, cap: null, floor: null, tiers: [],
      fxBase: ''
    };
  }
  function blankContract(vendorId) {
    return {
      id: uid('c'), vendorId: vendorId || '', version: 1, from: '', to: '', cycle: '', baseRule: '',
      firstDate: '', product: '', indexText: '', trigger: null, adjustDay: '', reflectDay: '',
      formulaText: '', allowSubstitute: false, round: 'none', exclude: '', notes: '',
      items: [], file: '', savedAt: ''
    };
  }
  function blankPart(vendorId) { return { vendorId: vendorId || '', partNo: '', name: '', basePrice: null, curPrice: null, usage: {}, fx: false, fxAmount: null, memo: '' }; }

  /* ───────────── 품번 엑셀 대량 등록 ───────────── */
  var PART_HEAD = [
    ['partNo', /^(품번|PARTNO|PARTNUMBER|P\/?N|자재번호|품목코드|부품번호)$/],
    ['name', /^(품명|PARTNAME|DESCRIPTION|자재명|품목명)$/],
    ['basePrice', /기준.*단가|^BASEPRICE$|계약.*단가/],
    ['curPrice', /현재.*단가|현.*구매.*단가|^구매단가$|^CURRENTPRICE$|현단가/],
    ['fxAmount', /환율.*(대상|금액)/],
    ['fx', /환율.*(적용|여부)|^환율$|^FX$/]
  ];
  function headRole(h) {
    var k = key(h);
    if (!k) return { role: 'skip' };
    for (var i = 0; i < PART_HEAD.length; i++) if (PART_HEAD[i][1].test(k)) return { role: PART_HEAD[i][0] };
    var m = txt(h).match(/^(.*?)\s*[\(\[（]\s*kg\s*[\)\]）]\s*$/i) || txt(h).match(/^(.*?)\s*(사용량|중량|kg)$/i);
    if (m && txt(m[1])) return { role: 'usage', material: txt(m[1]).replace(/\s*(사용)?\s*$/, '') };
    return { role: 'skip' };
  }
  /** 머리줄 찾기 — 「품번」 칸이 있는 첫 줄. 열 순서가 바뀌어도 이름으로 알아본다 */
  function inspectParts(aoa) {
    var hr = -1;
    for (var r = 0; r < Math.min(aoa.length, 20); r++) {
      if ((aoa[r] || []).some(function (c) { return headRole(c).role === 'partNo'; })) { hr = r; break; }
    }
    if (hr < 0) return { ok: false, error: '「품번」 칸이 있는 머리줄을 찾지 못했습니다 — 양식을 내려받아 맞춰 주세요. 열을 직접 지정할 수도 있습니다.', headerRow: 0, columns: (aoa[0] || []).map(function (h, c) { return { col: c, head: txt(h), role: 'skip', material: '' }; }) };
    var seen = {};
    var cols = (aoa[hr] || []).map(function (h, c) {
      var x = headRole(h), o = { col: c, head: txt(h), role: x.role, material: x.material || '' };
      if (o.role !== 'usage' && o.role !== 'skip') { if (seen[o.role]) o.role = 'skip'; seen[o.role] = 1; }
      return o;
    });
    return { ok: true, headerRow: hr, columns: cols };
  }
  /**
   * 엑셀 표 → 품번 행. 오류는 행 · 열을 짚어 돌려준다(임의로 고치지 않는다).
   * existing: 이미 등록된 품번 목록(같은 업체) — 신규 / 기존(덮어쓰기)을 구분한다.
   */
  function parseParts(aoa, insp, existing) {
    var have = {}; (existing || []).forEach(function (p) { have[key(p)] = 1; });
    var cols = insp.columns, by = function (role) { return cols.filter(function (c) { return c.role === role; })[0]; };
    var cNo = by('partNo'), cName = by('name'), cBase = by('basePrice'), cCur = by('curPrice'), cFx = by('fx'), cFxAmt = by('fxAmount');
    var uCols = cols.filter(function (c) { return c.role === 'usage' && c.material; });
    var problems = [];
    if (!cNo) problems.push('품번 열이 지정되지 않았습니다');
    if (!cBase) problems.push('기준 파트단가 열이 지정되지 않았습니다');
    var dupM = {}; uCols.forEach(function (c) { var k = key(c.material); if (dupM[k]) problems.push('원소재 열 이름이 겹칩니다 — ' + c.material); dupM[k] = 1; });
    uCols.forEach(function (c) { if (isFxName(c.material)) problems.push('환율은 사용량(kg)을 넣지 않습니다 — 「' + c.head + '」 열을 빼거나 「환율 적용」 열로 바꿔 주세요'); });
    var rows = [], seen = {};
    if (problems.length) return { ok: false, problems: problems, rows: [], materials: [] };
    for (var r = insp.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      if (!row.some(function (v) { return txt(v) !== ''; })) continue;
      var errs = [], o = { row: r + 1, partNo: txt(row[cNo.col]), name: cName ? txt(row[cName.col]) : '', basePrice: null, curPrice: null, usage: {}, fx: false, fxAmount: null };
      if (!o.partNo) errs.push({ col: cNo.head, msg: '품번이 비어 있습니다' });
      else if (seen[key(o.partNo)]) errs.push({ col: cNo.head, msg: '파일 안에서 품번이 겹칩니다(' + seen[key(o.partNo)] + '행)' });
      else seen[key(o.partNo)] = r + 1;
      var b = cellNum(row[cBase.col]);
      if (b == null) errs.push({ col: cBase.head, msg: '기준 파트단가가 비어 있습니다' });
      else if (isNaN(b)) errs.push({ col: cBase.head, msg: '숫자가 아닙니다(' + txt(row[cBase.col]) + ')' });
      else if (b <= 0) errs.push({ col: cBase.head, msg: '0 이하입니다' });
      else o.basePrice = b;
      if (cCur) {
        var c = cellNum(row[cCur.col]);
        if (c != null && isNaN(c)) errs.push({ col: cCur.head, msg: '숫자가 아닙니다(' + txt(row[cCur.col]) + ')' });
        else if (c != null && c <= 0) errs.push({ col: cCur.head, msg: '0 이하입니다' });
        else o.curPrice = c;
      }
      uCols.forEach(function (uc) {
        var v = cellNum(row[uc.col]);
        if (v == null) { o.usage[uc.material] = null; return; }            // 빈 칸 = 미입력(0 과 다르다)
        if (isNaN(v)) { errs.push({ col: uc.head, msg: '숫자가 아닙니다(' + txt(row[uc.col]) + ')' }); return; }
        if (v < 0) { errs.push({ col: uc.head, msg: '중량이 음수입니다(' + v + ')' }); return; }
        o.usage[uc.material] = v;
      });
      if (cFx) o.fx = /^(Y|YES|O|예|적용|TRUE|1|100%?)$/i.test(txt(row[cFx.col]));
      if (cFxAmt) {
        var fa = cellNum(row[cFxAmt.col]);
        if (fa != null && isNaN(fa)) errs.push({ col: cFxAmt.head, msg: '숫자가 아닙니다(' + txt(row[cFxAmt.col]) + ')' });
        else if (fa != null && fa < 0) errs.push({ col: cFxAmt.head, msg: '음수입니다' });
        else { o.fxAmount = fa; if (fa != null && fa > 0) o.fx = true; }
      }
      o.errors = errs;
      o.status = errs.length ? 'error' : (have[key(o.partNo)] ? 'update' : 'new');
      rows.push(o);
    }
    return {
      ok: true, problems: [], rows: rows, materials: uCols.map(function (c) { return c.material; }),
      counts: { 'new': rows.filter(function (x) { return x.status === 'new'; }).length, update: rows.filter(function (x) { return x.status === 'update'; }).length, error: rows.filter(function (x) { return x.status === 'error'; }).length }
    };
  }
  function partsTemplate(materials) {
    var m = (materials && materials.length ? materials : ['철강', '구리', '고무', '알루미늄']).filter(function (x) { return !isFxName(x); });
    var head = ['품번', '품명', '기준 파트단가', '현재 구매단가'].concat(m.map(function (x) { return x + '(kg)'; })).concat(['환율 적용', '환율 대상 금액']);
    return [head,
      ['A-1001', '예) 브래킷', 100000, 102000].concat(m.map(function (_, i) { return [25, 3, 2, 0][i] != null ? [25, 3, 2, 0][i] : ''; })).concat(['', '']),
      ['A-1002', '예) 하우징', 150000, 150000].concat(m.map(function (_, i) { return [30, 5, 0, 4][i] != null ? [30, 5, 0, 4][i] : ''; })).concat(['Y', 20000])];
  }
  /** 화면 표(수기 입력)의 한 행 검사 — 엑셀과 같은 규칙 */
  function checkPart(p, materials) {
    var e = [];
    if (!txt(p.partNo)) e.push('품번이 비어 있습니다');
    if (!isNum(p.basePrice)) e.push('기준 파트단가가 비어 있습니다'); else if (p.basePrice <= 0) e.push('기준 파트단가가 0 이하입니다');
    if (p.curPrice != null && (!isNum(p.curPrice) || p.curPrice <= 0)) e.push('현재 구매단가가 올바르지 않습니다');
    Object.keys(p.usage || {}).forEach(function (m) {
      var v = p.usage[m];
      if (isFxName(m) && v != null) e.push('환율에는 사용량(kg)을 넣을 수 없습니다');
      else if (v != null && (!isNum(v) || v < 0)) e.push(m + ' 사용량이 올바르지 않습니다(' + v + ')');
    });
    return e;
  }

  /* ───────────── 계약서 글에서 후보 뽑기(담당자 확인용 — 자동 확정하지 않는다) ───────────── */
  var MAT_WORDS = ['철강', '강판', '후판', '열연', '냉연', '특수강', '합금강', '탄소강', '스테인리스', '생철', '고철', '철스크랩', '선철', '주물', '주철', '구리', '전기동', '황동', '청동', '알루미늄', '니켈', '아연', '납', '주석', '고무', '천연고무', '합성고무', '수지', '플라스틱', 'PP', 'ABS', '환율'];
  /** 표준 「납품대금 연동 특별약정서(별첨)」의 1~10번 항목을 번호로 나눠 읽는다(10번 기타사항 뒤는 읽지 않는다) */
  var STD = { 1: /^물품\s*등의\s*명칭\s*/, 2: /^주요\s*원재료\s*(또는\s*주요\s*에너지)?\s*/, 3: /^원재료\s*가격.*?기준\s*지표\s*/, 4: /^원재료\s*기준\s*가격[^:：]*?(?=기준\s*시점|비교\s*시점|$)/, 5: /^조정\s*요건\s*/, 6: /^조정\s*주기\s*/, 7: /^조정일\s*/, 8: /^조정\s*대금\s*반영\s*시점\s*/, 9: /^납품\s*대금\s*연동\s*산식\s*/, 10: /^기타\s*사항\s*/ };
  function standardSections(lines) {
    var i0 = -1;
    for (var i = 0; i < lines.length; i++) if (/^1\s*[.．]\s*물품\s*등의\s*명칭/.test(lines[i])) { i0 = i; break; }
    if (i0 < 0) return null;
    var out = {}, cur = 0;
    for (var j = i0; j < Math.min(lines.length, i0 + 200); j++) {
      var t = lines[j], m = t.match(/^(1\s*0|[1-9])\s*[.．]\s*(.*)$/);
      if (m && STD[+m[1].replace(/\s/g, '')].test(m[2])) {
        cur = +m[1].replace(/\s/g, ''); if (cur === 10) break;
        out[cur] = m[2].replace(STD[cur], '').trim();
      } else if (cur) {
        if (/^(구분|기재\s*사항)$/.test(t)) continue;
        if (/【\s*첨부|납품단가\s*변동표|^위탁기업\s+수탁기업/.test(t)) break;
        out[cur] += (out[cur] ? '\n' : '') + t;
      }
    }
    return out;
  }
  function splitMaterials(s) { return String(s || '').split(/[\/,、\n·]| 및 /).map(function (x) { return txt(x).replace(/^[-•◦]\s*/, ''); }).filter(function (x) { return x && x.length <= 30; }); }
  function extractContract(text) {
    var lines = String(text || '').split(/\r?\n/).map(txt).filter(Boolean), out = { fields: {}, materials: [], formulaLines: [], notes: [], standard: false };
    var put = function (k, v, line) { if (out.fields[k] == null && v) { out.fields[k] = { value: v, line: line }; } };
    var sec = standardSections(lines);
    if (sec) {
      out.standard = true;
      var one = function (n) { return txt(sec[n]); };
      if (one(1)) put('product', one(1), '1. 물품등의 명칭');
      splitMaterials(sec[2]).forEach(function (n) { out.materials.push({ name: n, share: null, line: '2. 주요 원재료 또는 주요 에너지' }); });
      if (one(3)) put('indexText', one(3), '3. 원재료 가격의 기준지표');
      if (one(4)) put('baseRule', one(4), '4. 원재료 기준가격(기준시점 · 비교시점)');
      if (one(5)) { var tr = one(5).match(/[±+]?\s*(\d+(?:\.\d+)?)\s*%/); if (tr) put('trigger', parseFloat(tr[1]), '5. 조정요건 — ' + one(5)); }
      if (one(6)) put('cycle', /반기|6\s*개월/.test(one(6)) ? 'half' : /분기|3\s*개월/.test(one(6)) ? 'quarter' : /1\s*년|매년|연\s*1/.test(one(6)) ? 'year' : 'etc', '6. 조정주기 — ' + one(6));
      if (one(7)) put('adjustDay', one(7), '7. 조정일');
      if (one(8)) put('reflectDay', one(8), '8. 조정대금 반영시점');
      if (one(9)) { put('formulaText', sec[9].trim(), '9. 납품대금 연동산식'); out.formulaLines.push(one(9)); }
    }
    lines.forEach(function (ln) {
      var m;
      if ((m = ln.match(/(?:공급자|협력(?:업체|사)|수급사업자|"?을"?)\s*[:：]\s*([^,;()]{2,40})/))) put('vendor', txt(m[1]).replace(/\s*\(이하.*$/, ''), ln);
      if ((m = ln.match(/(20\d{2})[.\-년\s]+(\d{1,2})[.\-월\s]+(?:\d{1,2}[일.]?)?\s*(?:~|부터|-|–)\s*(20\d{2})[.\-년\s]+(\d{1,2})/)) && /기간|적용|계약/.test(ln)) {
        put('from', m[1] + '-' + ('0' + m[2]).slice(-2), ln); put('to', m[3] + '-' + ('0' + m[4]).slice(-2), ln);
      }
      if (/주기|마다|매\s*(분기|반기)/.test(ln)) { if (/반기/.test(ln)) put('cycle', 'half', ln); else if (/분기/.test(ln)) put('cycle', 'quarter', ln); else if (/매년|연\s*1회/.test(ln)) put('cycle', 'year', ln); }
      if (/기준(일|월|시점|가격|\s*Index)/i.test(ln)) put('baseRule', ln, ln);
      if ((m = ln.match(/(?:계약\s*체결일|체결일|계약일)\s*[:：]?\s*(20\d{2})[.\-년\s]+(\d{1,2})[.\-월\s]+(\d{1,2})/))) put('firstDate', m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2), ln);
      if (sec) return;                                 // 표준 약정서는 번호 항목만 믿는다(기타사항의 숫자를 조건으로 읽지 않는다)
      if ((m = ln.match(/상한[^0-9±+\-]{0,12}[±+]?\s*(\d+(?:\.\d+)?)\s*%/))) put('cap', parseFloat(m[1]), ln);
      if ((m = ln.match(/하한[^0-9±+\-]{0,12}[-−±]?\s*(\d+(?:\.\d+)?)\s*%/))) put('floor', -parseFloat(m[1]), ln);
      if ((m = ln.match(/(?:최소\s*변동|변동(?:률|폭)이?)[^0-9±]{0,14}[±]?\s*(\d+(?:\.\d+)?)\s*%\s*(?:미만|이내|이하)?/)) && /미만|이내|이하|최소/.test(ln)) put('minChange', parseFloat(m[1]), ln);
      if (/대체|유사\s*(Index|지수|지표)/i.test(ln)) put('substitute', ln, ln);
      if (/제외|적용하지\s*아니/.test(ln) && /연동/.test(ln)) put('exclude', ln, ln);
      if (/[=＝]/.test(ln) && /(단가|연동|조정|변동|Index|지수|가격)/i.test(ln)) out.formulaLines.push(ln);
      MAT_WORDS.forEach(function (w) {
        if (ln.indexOf(w) < 0) return;
        if (!/(원소재|원자재|연동|Index|지수|시세|가격|비중|환율)/i.test(ln)) return;
        if (out.materials.filter(function (x) { return x.name === w; }).length) return;
        var sh = ln.match(new RegExp(w + '[^0-9%]{0,12}(\\d+(?:\\.\\d+)?)\\s*%'));
        out.materials.push({ name: w, share: sh ? parseFloat(sh[1]) : null, line: ln });
      });
    });
    if (!out.formulaLines.length) out.notes.push('수식으로 보이는 줄을 찾지 못했습니다 — 계약서의 수식을 직접 입력해 주세요(임의로 해석하지 않습니다).');
    return out;
  }

  /* ───────────── 사용자 정의 수식 — 사칙연산과 괄호만(코드를 실행하지 않는다) ───────────── */
  var EXPR_VARS = { P0: '기준 파트단가', KG: '사용량(kg)', I0: '기준 Index(원/kg 환산 가능 시 원/kg)', I1: '변경 Index', R: '적용 변동률(소수)', S: '비중(소수)' };
  function evalExpr(expr, vars) {
    var s = String(expr || ''), i = 0, toks = [], m;
    while (i < s.length) {
      var ch = s[i];
      if (/\s/.test(ch)) { i++; continue; }
      if ((m = s.slice(i).match(/^\d+\.?\d*|^\.\d+/))) { toks.push({ t: 'n', v: parseFloat(m[0]) }); i += m[0].length; continue; }
      if ((m = s.slice(i).match(/^[A-Za-z][A-Za-z0-9]*/))) {
        var nm = m[0].toUpperCase();
        if (!(nm in EXPR_VARS)) throw new Error('모르는 이름 「' + m[0] + '」 — 쓸 수 있는 이름: ' + Object.keys(EXPR_VARS).join(', '));
        if (!isNum(vars[nm])) throw new Error(nm + '(' + EXPR_VARS[nm] + ') 값이 없습니다');
        toks.push({ t: 'n', v: vars[nm] }); i += m[0].length; continue;
      }
      if ('+-*/()×÷'.indexOf(ch) >= 0) { toks.push({ t: ch === '×' ? '*' : ch === '÷' ? '/' : ch }); i++; continue; }
      throw new Error('쓸 수 없는 글자 「' + ch + '」');
    }
    var p = 0;
    function peek() { return toks[p] && toks[p].t; }
    function prim() {
      var t = toks[p++];
      if (!t) throw new Error('수식이 중간에 끝났습니다');
      if (t.t === 'n') return t.v;
      if (t.t === '-') return -prim();
      if (t.t === '+') return prim();
      if (t.t === '(') { var v = sum(); if (peek() !== ')') throw new Error('괄호가 닫히지 않았습니다'); p++; return v; }
      throw new Error('수식을 읽지 못했습니다');
    }
    function term() { var v = prim(); while (peek() === '*' || peek() === '/') { var o = toks[p++].t, b = prim(); if (o === '/' && b === 0) throw new Error('0 으로 나눕니다'); v = o === '*' ? v * b : v / b; } return v; }
    function sum() { var v = term(); while (peek() === '+' || peek() === '-') { var o = toks[p++].t, b = term(); v = o === '+' ? v + b : v - b; } return v; }
    if (!toks.length) throw new Error('수식이 비어 있습니다');
    var out = sum();
    if (p < toks.length) throw new Error('수식을 끝까지 읽지 못했습니다');
    if (!isFinite(out)) throw new Error('계산 결과가 숫자가 아닙니다');
    return out;
  }

  /* ───────────── Index — 기간 평균 · 단위 ───────────── */
  function unitOf(s) {
    var u = (s.meta && s.meta.unit) || '';
    if (!u) { var m = txt(s.name).match(/(KRW|원|USD|RMB|CNY|US\$)\s*\/\s*(kg|ton|t|mt|usd|cny)\b/i); if (m) u = m[0]; }
    return txt(u);
  }
  /** 원/kg 으로 바꾸는 곱수. 원화 중량 단가가 아니면 null(중량 연동식에 쓸 수 없다) */
  function perKgFactor(unit) {
    var u = txt(unit).toUpperCase().replace(/\s/g, '').replace(/\(환산\)/, '').replace(/^원/, 'KRW');
    if (/^KRW\/KG$/.test(u)) return 1;
    if (/^KRW\/(TON|T|MT)$/.test(u)) return 1 / 1000;
    return null;
  }
  function monthsBetween(a, b) { var out = [], i = CostIdx.mIdx(a), j = CostIdx.mIdx(b); for (; i <= j; i++) out.push(CostIdx.mStr(i)); return out; }
  /** 기간 평균 — 그 기간에 있는 달만 쓴다. 다른 달 값을 끌어오지 않는다(없으면 없다고 한다) */
  function avgOver(series, from, to) {
    if (!series) return { ok: false, v: null, used: [], missing: [], msg: 'Index 가 없습니다' };
    if (!from || !to || CostIdx.mIdx(from) > CostIdx.mIdx(to)) return { ok: false, v: null, used: [], missing: [], msg: '기간이 올바르지 않습니다' };
    var by = {}; series.points.forEach(function (p) { by[p.m] = p.v; });
    var used = [], missing = [], sum = 0;
    monthsBetween(from, to).forEach(function (m) { if (by[m] == null) missing.push(m); else { used.push(m); sum += by[m]; } });
    if (!used.length) return { ok: false, v: null, used: [], missing: missing, msg: '기간(' + from + ' ~ ' + to + ')에 값이 하나도 없습니다' };
    return { ok: true, v: sum / used.length, used: used, missing: missing };
  }

  /* ───────────── Index 통합문서 → 계열 (가로형: 원소재가 행 · 날짜가 열 포함) ───────────── */
  var WIDE_NAME = /^(원소재이름|원소재명|원재료명|원재료이름|소재명|INDEX명|INDEX|지표명|품목)$/;
  var WIDE_META = { basis: /^기준$/, source: /^출처$/, region: /^(지역|지역\/거래소|거래소)$/, unit: /^단위$/, spec: /^규격$/, detail: /^상세출처$/, origin: /^(원본시트|시트)$/ };
  function headMonth(v) {
    if (v instanceof Date) return v.getFullYear() + '-' + ('0' + (v.getMonth() + 1)).slice(-2);
    if (typeof v === 'number' && v > 20000 && v < 80000) { var d = new Date(Math.round((v - 25569) * 86400000)); return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2); }
    var t = txt(v), m = t.match(/^(20\d{2})\s*[-./년]\s*(\d{1,2})(?:\s*[-./월]\s*\d{0,2}\s*일?)?$/) || t.match(/^(20\d{2})(\d{2})(?:\d{2})?$/);
    if (m && +m[2] >= 1 && +m[2] <= 12) return m[1] + '-' + ('0' + (+m[2])).slice(-2);
    m = t.match(/^(\d{2})\s*[.']\s*(\d{1,2})월?$/);
    if (m && +m[2] >= 1 && +m[2] <= 12) return '20' + m[1] + '-' + ('0' + (+m[2])).slice(-2);
    return '';
  }
  /** 머리줄에 「원소재 이름」 칸과 날짜 칸이 있으면 가로형 */
  function wideInfo(aoa) {
    for (var r = 0; r < Math.min((aoa || []).length, 12); r++) {
      var row = aoa[r] || [], cName = -1, meta = {}, dates = [];
      row.forEach(function (h, c) {
        var k = txt(h).toUpperCase().replace(/\s/g, '');
        if (cName < 0 && WIDE_NAME.test(k)) { cName = c; return; }
        for (var f in WIDE_META) if (meta[f] == null && WIDE_META[f].test(k)) { meta[f] = c; return; }
        var m = headMonth(h); if (m) dates.push({ col: c, m: m });
      });
      if (cName >= 0 && dates.length >= 2) return { headerRow: r, cName: cName, meta: meta, dates: dates };
    }
    return null;
  }
  function wideSeries(aoa, w, sheet, file) {
    var out = [], issues = [];
    for (var r = w.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [], nm = txt(row[w.cName]); if (!nm) continue;
      var g = function (f) { return w.meta[f] == null ? '' : txt(row[w.meta[f]]); };
      var acc = {}, badN = 0;
      w.dates.forEach(function (d) { var v = cellNum(row[d.col]); if (v == null) return; if (isNaN(v)) { badN++; return; } (acc[d.m] = acc[d.m] || []).push(v); });
      if (badN) issues.push({ msg: sheet + ' · ' + nm + ' — 숫자가 아닌 값 ' + badN + '칸을 뺐습니다' });
      var pts = Object.keys(acc).sort().map(function (m) { return { m: m, v: acc[m].reduce(function (a, b) { return a + b; }, 0) / acc[m].length }; });
      if (!pts.length) { issues.push({ msg: sheet + ' · ' + nm + ' — 날짜 열에 값이 없어 뺐습니다' }); continue; }
      var unit = g('unit'), fx = isFxName(nm) || /^KRW\s*\/\s*(USD|CNY|EUR|JPY)$/i.test(unit);
      var name = [nm, g('spec'), g('basis'), g('region'), unit].filter(Boolean).join(' · ');
      out.push({
        name: name, role: fx ? 'skip' : 'index', family: CostIdx.indexFamily(nm + ' ' + g('spec')) || '', grades: CostIdx.gradeTokens(g('spec') + ' ' + nm), points: pts, col: r,
        source: (file ? file + ' / ' : '') + '시트 ' + sheet + ' / ' + (r + 1) + '행',
        meta: { sheet: g('origin') || sheet, fileSheet: sheet, kind: fx ? 'fx' : '', item: nm, spec: g('spec'), basis: g('basis'), region: g('region'), unit: unit, source: g('source'), detail: g('detail'), wide: true }
      });
    }
    return { series: out, issues: issues };
  }
  /** 통합문서 전체 → 계열. 가로형 시트는 여기서, 나머지(머리 블록 · 단순 표)는 CostIdx 로 읽는다 */
  function buildIndex(book, file, fx) {
    var rest = { names: [], sheets: {}, merges: book.merges }, series = [], issues = [], sheets = [], seen = {}, notes = [];
    (book.names || Object.keys(book.sheets)).forEach(function (n) {
      var w = wideInfo(book.sheets[n]);
      if (!w) { rest.names.push(n); rest.sheets[n] = book.sheets[n]; return; }
      var b = wideSeries(book.sheets[n], w, n, file);
      sheets.push({ sheet: n, ok: true, kind: 'wide' }); series = series.concat(b.series); issues = issues.concat(b.issues);
    });
    if (rest.names.length) {
      var insp = CostIdx.inspectBook(rest), b2 = CostIdx.buildBook(rest, insp, file);
      insp.sheets.forEach(function (sh) { sheets.push({ sheet: sh.sheet, ok: !!sh.ok, kind: sh.kind, error: sh.error }); });
      series = series.concat(b2.series); issues = issues.concat(b2.issues);
    }
    series.forEach(function (x) { if (seen[x.name]) x.name = x.name + ' [' + ((x.meta && (x.meta.fileSheet || x.meta.sheet)) || '') + ' ' + ((x.col || 0) + 1) + ']'; seen[x.name] = 1; });
    if (fx) { var f = CostIdx.fxConvert(series); series = f.series; notes = f.notes; }
    return { series: series, issues: issues, sheets: sheets, notes: notes };
  }

  /* ───────────── 계약 원소재 ↔ Index ───────────── */
  /** 회사 꼴 표기((주) · ㈜ · 주식회사 …)를 떼고 비교한다 */
  function corpKey(v) { return key(txt(v).replace(/\(\s*주\s*\)|㈜|주식회사|\(\s*유\s*\)|유한회사|CO\.?\s*,?\s*LTD\.?|INC\.?|CORP\.?/gi, '')); }
  function sheetMatches(s, vendor) {
    var sh = corpKey(s.meta && s.meta.sheet), hints = [vendor && vendor.sheetHint, vendor && vendor.name].map(corpKey).filter(Boolean);
    return !!sh && hints.some(function (h) { return sh === h || (h.length >= 2 && sh.indexOf(h) >= 0); });
  }
  function usable(series, fx) {
    return series.filter(function (s) { return fx ? (s.meta && s.meta.kind === 'fx') || /환율|KRW\/(USD|CNY|EUR|JPY)/i.test(s.name + ' ' + unitOf(s)) : s.role !== 'skip' && !(s.meta && s.meta.kind === 'fx'); });
  }
  /** 이름 · 규격으로 「정확히 일치」하는가 — 낱말이 Index 이름에 그대로 들어 있거나, 규격의 재질 기호가 같을 때만 */
  function exactHit(item, s) {
    var nm = key(item.name), hay = key(s.name + ' ' + ((s.meta && s.meta.item) || ''));
    if (item.kind === 'fx') {
      var cur = (txt(item.spec) + ' ' + txt(item.name)).toUpperCase().match(/USD|EUR|JPY|CNY|RMB|달러|유로|엔|위안/);
      var c = cur ? { '달러': 'USD', '유로': 'EUR', '엔': 'JPY', '위안': 'CNY', 'RMB': 'CNY' }[cur[0]] || cur[0] : 'USD';
      return hay.indexOf(c) >= 0 || (c === 'CNY' && /RMB|위안/.test(hay));
    }
    var g = item.spec ? CostIdx.gradeOf(item.spec) : '';
    if (g && (s.grades || []).indexOf(g) >= 0) return true;
    if (item.spec && key(item.spec).length >= 3 && hay.indexOf(key(item.spec)) >= 0) return true;
    if (item.spec) return false;                      // 규격을 적었는데 규격이 안 맞으면 이름만으로 정확 일치라 하지 않는다
    return nm.length >= 2 && hay.indexOf(nm) >= 0;
  }
  function familyOf(item) {
    var a = item.spec ? CostIdx.normMaterial(item.spec).family : '';
    return a || CostIdx.normMaterial(item.name).family || CostIdx.indexFamily(item.name) || '';
  }
  /**
   * approved : 담당자가 이미 승인한 Index
   * exact    : 정확 일치 하나(업체 시트 우선) — 그대로 쓴다
   * multi    : 정확 일치가 여럿 — 담당자가 고른다
   * similar  : 정확 일치 없음 — 유사 후보 최대 3개(승인 전에는 계산에 쓰지 않는다)
   * none     : 후보 없음
   */
  function matchItem(item, vendor, series, period) {
    var res = { status: '', chosen: '', where: '', candidates: [], reason: '', needApproval: false, notes: [] };
    var pool = usable(series, item.kind === 'fx');
    if (item.indexName) {
      var a = series.filter(function (s) { return s.name === item.indexName; })[0];
      if (a) { res.status = 'approved'; res.chosen = a.name; res.where = sheetMatches(a, vendor) ? 'vendor' : 'other'; res.similar = item.indexStatus === 'similar'; res.reason = '담당자가 승인한 Index' + (res.similar ? '(유사 Index — 대체 지표)' : ''); return res; }
      res.notes.push('승인해 둔 Index 「' + item.indexName + '」가 이번 파일에 없습니다 — 다시 맞춰 주세요');
    }
    var vs = pool.filter(function (s) { return sheetMatches(s, vendor); });
    var hitV = vs.filter(function (s) { return exactHit(item, s); });
    var hitO = pool.filter(function (s) { return !sheetMatches(s, vendor) && exactHit(item, s); });
    if (vs.length && !hitV.length) res.notes.push('업체 시트는 있지만 「' + item.name + (item.spec ? ' ' + item.spec : '') + '」와 맞는 소재가 없어 그 시트의 Index 를 자동으로 쓰지 않았습니다');
    var hits = hitV.length ? hitV : hitO;
    res.where = hitV.length ? 'vendor' : (hitO.length ? 'other' : '');
    if (hits.length === 1) { res.status = 'exact'; res.chosen = hits[0].name; res.reason = (hitV.length ? '업체 시트에서 ' : '다른 시트에서 ') + '계약 원소재와 정확히 일치하는 Index'; res.candidates = [cand(hits[0], '정확 일치')]; return res; }
    if (hits.length > 1) {
      var ranked = hits.slice().sort(function (x, y) { return CostIdx.prefRank(x) - CostIdx.prefRank(y); });
      res.status = 'multi'; res.needApproval = true; res.candidates = ranked.map(function (s) { return cand(s, '정확 일치'); });
      res.reason = '정확히 일치하는 Index 가 ' + hits.length + '개입니다 — 임의로 고르지 않습니다. 계약서 기준에 맞는 것을 골라 주세요'; return res;
    }
    if (item.kind === 'fx') { res.status = 'none'; res.needApproval = true; res.reason = '환율 계열을 찾지 못했습니다 — 통합문서의 환율 시트를 확인하거나 직접 골라 주세요'; res.candidates = pool.slice(0, 3).map(function (s) { return cand(s, '환율 계열'); }); return res; }
    var fam = familyOf(item);
    if (!fam) { res.status = 'none'; res.needApproval = true; res.reason = '「' + item.name + '」와 일치하는 Index 가 없고, 재질 계열도 알 수 없어 유사 후보를 추리지 못했습니다 — 직접 골라 주세요'; return res; }
    var m = CostIdx.matchIndex(fam, pool, { start: period.from, end: period.to }, { refHint: false });
    var list = (m.candidates || []).slice(0, 3);
    if (!list.length && m.chosen) list = [{ name: m.chosen, matSim: 1, why: m.reason }];
    if (!list.length) { res.status = 'none'; res.needApproval = true; res.reason = m.reason || '유사한 Index 가 없습니다'; return res; }
    res.status = 'similar'; res.needApproval = true;
    res.candidates = list.map(function (c) {
      var s = pool.filter(function (x) { return x.name === c.name; })[0];
      var o = cand(s, c.why || '같은 재질 계열');
      var total = monthsBetween(period.from, period.to).length, st = avgOver(s, period.from, period.to);
      o.matSim = c.matSim; o.corr = c.corr == null ? null : c.corr; o.dirAgree = c.dirAgree == null ? null : c.dirAgree; o.rate = c.rate == null ? null : c.rate;
      o.complete = total ? st.used.length / total : null;
      o.score = c.score != null ? c.score : 0.4 * (c.matSim || 0);
      return o;
    });
    res.reason = '정확히 일치하는 Index 가 없어 유사 Index 후보를 추렸습니다 — 계약상 대체가 허용되는지 확인하고 승인해야 계산에 씁니다';
    res.notes = res.notes.concat(m.notes || []);
    if (m.basis) res.notes.push(m.basis);
    return res;
  }
  function cand(s, why) { return { name: s.name, sheet: (s.meta && s.meta.sheet) || '', unit: unitOf(s), source: s.source || '', why: why }; }

  /* ───────────── 계산 ───────────── */
  function contractProblems(contract) {
    var e = [], seen = {};
    if (!contract) return ['계약 정보가 없습니다'];
    if (!contract.items.length) e.push('계약에 연동 항목(원소재 · 환율)이 없습니다');
    contract.items.forEach(function (it) {
      var k = key(it.name);
      if (!k) e.push('이름이 빈 연동 항목이 있습니다');
      else if (seen[k]) e.push('같은 원소재가 두 번 등록돼 있습니다 — ' + it.name);
      seen[k] = 1;
      if (it.kind === 'fx') {
        if (!it.fxBase) e.push(it.name + ' — 환율을 무엇에 곱하는지(기준 파트단가 전체 / 품번별 환율 대상 금액) 계약 조건을 정해 주세요');
        if (it.reflect != null && it.reflect !== 100) e.push(it.name + ' — 환율 적용 비중은 계약상 100% 입니다(' + it.reflect + '% 로 저장돼 있습니다)');
      } else {
        if (!FORMULA[it.formula]) e.push(it.name + ' — 계약 수식이 정해지지 않았습니다(임의로 해석하지 않습니다)');
        if (it.formula === 'ratio' && !isNum(it.share)) e.push(it.name + ' — 비중 연동식인데 비중(%)이 없습니다');
        if (it.formula === 'custom' && !txt(it.expr)) e.push(it.name + ' — 사용자 정의 수식이 비어 있습니다');
      }
      if (isNum(it.cap) && isNum(it.floor) && it.floor > it.cap) e.push(it.name + ' — 하한이 상한보다 큽니다');
    });
    return e;
  }
  /** 변동률에 계약 조건을 적용: 최소 변동률 → 상한 · 하한 → 구간별 반영률 → 반영률 */
  function applyTerms(rate, it, contract) {
    var steps = [], r = rate;
    var own = isNum(it.minChange), min = own ? it.minChange : (contract && isNum(contract.trigger) ? contract.trigger : null);
    if (min != null && Math.abs(rate) < min) { return { rate: 0, steps: ['변동률 ' + r2(rate) + '% 가 ' + (own ? '최소 변동률 ' : '조정요건 ±') + min + '% 미만 — 이번에는 조정하지 않음(기준시점 유지)'], excluded: true }; }
    if (min != null) steps.push((own ? '최소 변동률 ' : '조정요건 ±') + min + '% 충족');
    if (isNum(it.cap) && r > it.cap) { steps.push('상한 ' + it.cap + '% 적용(' + r2(r) + '% → ' + it.cap + '%)'); r = it.cap; }
    if (isNum(it.floor) && r < it.floor) { steps.push('하한 ' + it.floor + '% 적용(' + r2(r) + '% → ' + it.floor + '%)'); r = it.floor; }
    var ref = it.kind === 'fx' ? 100 : (isNum(it.reflect) ? it.reflect : 100);
    var tiers = (it.tiers || []).filter(function (t) { return isNum(t[0]) && isNum(t[1]); }).sort(function (a, b) { return a[0] - b[0]; });
    if (tiers.length && it.kind !== 'fx') {
      var hit = null; tiers.forEach(function (t) { if (Math.abs(r) >= t[0]) hit = t; });
      if (hit) { ref = hit[1]; steps.push('구간별 차등 — 변동 ' + hit[0] + '% 이상 구간 반영률 ' + hit[1] + '%'); }
    }
    if (ref !== 100) steps.push('반영률 ' + ref + '%');
    return { rate: r * ref / 100, steps: steps, excluded: false };
  }

  /**
   * input: { vendor, contract, parts[], series[], period:{base:{from,to}, next:{from,to}}, requests:{품번: 요청단가}, applyMonth }
   * 계산을 멈추는 사유는 blocks(전체) · 행의 blocks(품번)에, 계속하되 알릴 것은 warnings 에 담는다.
   */
  function calcRun(input) {
    var out = { ok: false, blocks: [], warnings: [], items: [], rows: [], period: input.period, vendor: input.vendor && input.vendor.name };
    var c = input.contract, P = input.period || {}, series = input.series || [];
    contractProblems(c).forEach(function (m) { out.blocks.push({ code: 'contract', msg: m }); });
    if (!P.base || !P.next || !P.next.from || !P.next.to) out.blocks.push({ code: 'period', msg: '변경 Index 기간을 정해 주세요' });
    else if (CostIdx.mIdx(P.next.from) > CostIdx.mIdx(P.next.to)) out.blocks.push({ code: 'period', msg: '변경 Index 기간의 시작이 끝보다 늦습니다' });
    if (P.base && P.base.from && P.base.to && P.next && P.next.from && CostIdx.mIdx(P.base.to) >= CostIdx.mIdx(P.next.from)) out.warnings.push({ code: 'period', msg: '기준 기간(' + P.base.from + ' ~ ' + P.base.to + ')이 변경 기간(' + P.next.from + ' ~ ' + P.next.to + ')과 겹치거나 더 늦습니다 — 기준일을 확인해 주세요' });
    if (c && input.applyMonth && ((c.from && CostIdx.mIdx(input.applyMonth) < CostIdx.mIdx(c.from)) || (c.to && CostIdx.mIdx(input.applyMonth) > CostIdx.mIdx(c.to)))) out.blocks.push({ code: 'version', msg: '적용 월(' + input.applyMonth + ')이 이 계약 버전의 적용 기간(' + (c.from || '?') + ' ~ ' + (c.to || '?') + ') 밖입니다 — 계약 버전을 확인해 주세요' });
    if (!series.length) out.blocks.push({ code: 'index', msg: 'Index 엑셀을 먼저 올려 주세요' });
    if (out.blocks.length) return out;

    var matchSpan = { from: (P.base.from || P.next.from), to: P.next.to };
    // 1) 연동 항목마다 Index · 변동률
    out.items = c.items.map(function (it) {
      var o = { id: it.id, name: it.name, kind: it.kind, spec: it.spec, formula: it.kind === 'fx' ? 'fx' : it.formula, blocks: [], notes: [], steps: [] };
      var m = matchItem(it, input.vendor, series, matchSpan);
      o.match = m; o.notes = o.notes.concat(m.notes);
      var chosen = '';
      if (m.status === 'approved') {
        chosen = m.chosen;
        if (m.similar && !c.allowSubstitute) o.blocks.push('유사 Index 를 승인했지만 계약 설정이 「대체 Index 사용 불허」입니다 — 계약 조건을 확인해 주세요');
      } else if (m.status === 'exact') chosen = m.chosen;
      else if (m.status === 'multi') o.blocks.push('일치하는 Index 가 여럿입니다 — 「Index 데이터 관리」에서 골라 승인해 주세요');
      else if (m.status === 'similar') o.blocks.push(c.allowSubstitute ? '정확히 일치하는 Index 가 없습니다 — 유사 Index 후보를 검토해 승인해야 계산합니다' : '정확히 일치하는 Index 가 없고, 계약상 대체 Index 사용이 허용돼 있지 않습니다(자동 대체하지 않습니다)');
      else o.blocks.push(m.reason || '맞는 Index 가 없습니다');
      var s = chosen ? series.filter(function (x) { return x.name === chosen; })[0] : null;
      o.index = chosen; o.unit = s ? unitOf(s) : ''; o.source = s ? s.source : '';
      if (s) {
        var b;
        if (isNum(it.baseValue)) { b = { ok: true, v: it.baseValue, used: [], missing: [], fixed: true }; o.notes.push('기준 Index 는 계약에 적힌 고정값 ' + it.baseValue + ' 을 썼습니다'); }
        else if (!P.base.from || !P.base.to) b = { ok: false, msg: '기준 Index 기간이 없습니다(계약에 고정 기준값도 없습니다)', used: [], missing: [] };
        else b = avgOver(s, P.base.from, P.base.to);
        var n = avgOver(s, P.next.from, P.next.to);
        o.base = b; o.next = n;
        if (!b.ok) o.blocks.push('기준 Index — ' + b.msg);
        if (!n.ok) o.blocks.push('변경 Index — ' + n.msg);
        if (b.ok && b.missing.length) o.notes.push('기준 기간 중 값이 없는 달: ' + b.missing.join(', ') + ' — 있는 달만 평균했습니다');
        if (n.ok && n.missing.length) o.notes.push('변경 기간 중 값이 없는 달: ' + n.missing.join(', ') + ' — 있는 달만 평균했습니다');
        if (b.ok && n.ok) {
          if (b.v <= 0) o.blocks.push('기준 Index 가 0 이하입니다');
          else {
            o.i0 = b.v; o.i1 = n.v; o.rawRate = (n.v / b.v - 1) * 100;
            var t = applyTerms(o.rawRate, it, c); o.rate = t.rate; o.steps = t.steps; o.excluded = t.excluded;
            if (it.kind !== 'fx' && it.formula === 'weight') {
              var f = perKgFactor(o.unit);
              if (f == null) o.blocks.push('Index 단위(' + (o.unit || '없음') + ')가 원/kg · 원/ton 이 아니라 중량 연동식에 쓸 수 없습니다 — 원화 환산 계열을 고르거나 수식을 확인해 주세요');
              else { o.kgFactor = f; o.i0kg = b.v * f; o.i1kg = n.v * f; }
            }
          }
        }
      }
      return o;
    });

    // 2) 품번마다
    var seen = {};
    out.rows = (input.parts || []).map(function (p) {
      var row = { partNo: p.partNo, name: p.name, basePrice: p.basePrice, curPrice: p.curPrice, lines: [], blocks: [], warnings: [], total: null, newPrice: null };
      var k = key(p.partNo);
      if (seen[k]) row.blocks.push('품번이 중복 등록돼 있습니다'); seen[k] = 1;
      checkPart(p).forEach(function (m) { row.blocks.push(m); });
      var sum = 0, matCost = 0;
      out.items.forEach(function (o, idx) {
        var it = c.items[idx], ln = { item: o.name, kind: o.kind, formula: o.formula, index: o.index, i0: o.i0, i1: o.i1, rawRate: o.rawRate, rate: o.rate, kg: null, amount: null, note: '', calc: '' };
        if (o.kind === 'fx') {
          if (p.usage && p.usage[o.name] != null) { row.blocks.push('환율 항목에 사용량(kg)이 들어 있습니다'); }
          if (!p.fx) { ln.note = '환율 미적용 품번'; ln.amount = 0; ln.skip = true; row.lines.push(ln); return; }
          var base = it.fxBase === 'price' ? p.basePrice : p.fxAmount;
          if (it.fxBase === 'amount' && !isNum(p.fxAmount)) { row.blocks.push('환율 대상 금액이 없습니다(' + o.name + ')'); row.lines.push(ln); return; }
          if (o.blocks.length) { row.blocks.push(o.name + ' — ' + o.blocks[0]); row.lines.push(ln); return; }
          ln.amount = base * o.rate / 100; ln.kg = '-';
          ln.calc = (it.fxBase === 'price' ? '기준 파트단가 ' : '환율 대상 금액 ') + fmt(base) + ' × 환율 변동률 ' + fmtPct(o.rate) + ' (100% 적용)';
          sum += ln.amount; row.lines.push(ln); return;
        }
        var kg = p.usage ? p.usage[o.name] : undefined;
        ln.kg = kg == null ? null : kg;
        if (o.formula === 'weight' || o.formula === 'custom') {
          if (kg == null) {
            if (o.formula === 'weight') { row.blocks.push(o.name + ' 사용량(kg)이 비어 있습니다(쓰지 않는 소재면 0 을 넣어 주세요)'); row.lines.push(ln); return; }
          } else if (kg === 0 && o.formula === 'weight') { ln.amount = 0; ln.note = '사용량 0'; row.lines.push(ln); return; }
        }
        if (o.blocks.length) { row.blocks.push(o.name + ' — ' + o.blocks[0]); row.lines.push(ln); return; }
        if (o.formula === 'weight') {
          ln.amount = kg * o.i0kg * o.rate / 100;
          ln.calc = fmt(kg, 3) + ' kg × 기준 ' + fmt(o.i0kg, 2) + ' 원/kg × ' + fmtPct(o.rate);
          matCost += kg * o.i0kg;
        } else if (o.formula === 'ratio') {
          ln.amount = p.basePrice * it.share / 100 * o.rate / 100;
          ln.calc = '기준 파트단가 ' + fmt(p.basePrice) + ' × 비중 ' + it.share + '% × ' + fmtPct(o.rate);
        } else {
          try {
            ln.amount = evalExpr(it.expr, { P0: p.basePrice, KG: kg, I0: o.i0kg != null ? o.i0kg : o.i0, I1: o.i1kg != null ? o.i1kg : o.i1, R: o.rate / 100, S: isNum(it.share) ? it.share / 100 : undefined });
            ln.calc = '수식 ' + it.expr;
          } catch (e) { row.blocks.push(o.name + ' 수식 — ' + e.message); row.lines.push(ln); return; }
        }
        sum += ln.amount; row.lines.push(ln);
      });
      if (isNum(p.basePrice) && matCost > p.basePrice) row.warnings.push('원소재 금액(사용량 × 기준 Index = ' + fmt(matCost) + '원)이 기준 파트단가보다 큽니다 — 중량 입력을 확인해 주세요');
      if (!row.blocks.length) {
        row.total = sum;
        var np = p.basePrice + sum;
        row.newPriceRaw = np;
        row.newPrice = c.round === '1' ? Math.round(np) : c.round === '10' ? Math.round(np / 10) * 10 : np;
        row.rate = sum / p.basePrice * 100;
        row.matAmount = row.lines.filter(function (l) { return l.kind !== 'fx' && isNum(l.amount); }).reduce(function (a, l) { return a + l.amount; }, 0);
        row.fxAmountDelta = row.lines.filter(function (l) { return l.kind === 'fx' && isNum(l.amount); }).reduce(function (a, l) { return a + l.amount; }, 0);
      }
      // 업체 요청단가 비교
      var req = input.requests ? input.requests[p.partNo] : null;
      if (isNum(req)) {
        var cur = isNum(p.curPrice) ? p.curPrice : p.basePrice;
        row.request = { price: req, curBase: isNum(p.curPrice) ? '현재 구매단가' : '기준 파트단가(현재 구매단가 없음)', reqDelta: req - cur, reqRate: (req / cur - 1) * 100 };
        if (row.newPrice != null) {
          row.request.calcRate = (row.newPrice / cur - 1) * 100;
          row.request.gap = req - row.newPrice;
          row.request.explained = row.total;
          row.request.verdict = Math.abs(row.request.gap) < 0.5 ? '계약 계산과 같음' : row.request.gap > 0 ? '요청이 계약 계산보다 높음 — 가공비 · 인건비 · 경비 · 물류비 등 추가 근거 검토' : '요청이 계약 계산보다 낮음';
        }
      }
      return row;
    });
    out.items.forEach(function (o) { o.blocks.forEach(function (m) { out.warnings.push({ code: 'item', msg: o.name + ' — ' + m }); }); });
    out.done = out.rows.filter(function (r) { return r.newPrice != null; }).length;
    out.stopped = out.rows.length - out.done;
    out.ok = out.rows.length > 0 && out.stopped === 0;
    out.sum = { base: tot(out.rows, 'basePrice'), delta: tot(out.rows, 'total'), next: tot(out.rows, 'newPrice') };
    return out;
  }
  function tot(rows, k) { return rows.filter(function (r) { return r.newPrice != null; }).reduce(function (a, r) { return a + (r[k] || 0); }, 0); }
  function fmt(v, d) { if (v === '-') return '-'; if (!isNum(v)) return ''; return v.toLocaleString('ko-KR', { maximumFractionDigits: d == null ? 0 : d, minimumFractionDigits: 0 }); }
  function fmtPct(v) { return isNum(v) ? (v > 0 ? '+' : '') + v.toFixed(2) + '%' : ''; }

  /* ───────────── 보고서(Word 틀에 넣을 값) · 엑셀 ───────────── */
  var NA = '확인 필요';
  function termText(it) {
    var a = [];
    if (it.kind === 'fx') a.push('환율 100% 적용(' + (FX_BASE[it.fxBase] || NA) + ')');
    else a.push(FORMULA[it.formula] || NA);
    if (it.formula === 'ratio' && isNum(it.share)) a.push('비중 ' + it.share + '%');
    if (it.formula === 'custom') a.push(it.expr);
    if (isNum(it.minChange)) a.push('최소 변동률 ' + it.minChange + '%');
    if (isNum(it.cap)) a.push('상한 ' + it.cap + '%');
    if (isNum(it.floor)) a.push('하한 ' + it.floor + '%');
    if (it.kind !== 'fx' && isNum(it.reflect) && it.reflect !== 100) a.push('반영률 ' + it.reflect + '%');
    return a.join(' · ');
  }
  /** docxtemplater 꼬리표 이름 = 아래 키. 값이 없으면 「확인 필요」 */
  function reportData(run) {
    var res = run.result, c = run.contract, P = run.period, v = function (x) { return x == null || x === '' ? NA : x; };
    var per = function (p) { return p && p.from ? (p.from === p.to ? p.from : p.from + ' ~ ' + p.to) : NA; };
    var simil = res.items.filter(function (o) { return o.match && o.match.similar; });
    return {
      '업체명': v(run.vendorName), '작성일': run.at ? run.at.slice(0, 10) : NA, '적용월': v(run.applyMonth),
      '계약버전': 'v' + c.version, '계약기간': (c.from || NA) + ' ~ ' + (c.to || NA), '적용주기': CYCLE[c.cycle] || NA,
      '기준기간': res.items.some(function (o) { return o.base && o.base.fixed; }) && !(P.base && P.base.from) ? '계약 고정 기준값' : per(P.base), '변경기간': per(P.next),
      '적용품목': v(c.product), '조정요건': isNum(c.trigger) ? '±' + c.trigger + '%' : NA, '조정일': v(c.adjustDay), '반영시점': v(c.reflectDay), '기준지표': v(c.indexText), '기준시점': v(c.baseRule),
      '계약수식': v(c.formulaText), '계약조건': c.items.map(function (it) { return it.name + (it.spec ? '(' + it.spec + ')' : '') + ': ' + termText(it); }).join('\n') || NA,
      'Index파일': v(run.indexFile),
      '품번수': String(res.rows.length), '계산완료': String(res.done), '확인필요건수': String(res.stopped),
      '기준합계': fmt(res.sum.base), '연동액합계': fmt(res.sum.delta, 2), '변경합계': fmt(res.sum.next, 2),
      '유사Index': simil.length ? simil.map(function (o) { return o.name + ' → ' + o.index + ' (유사 Index · 담당자 승인)'; }).join('\n') : '해당 없음(전 항목 정확 일치 Index)',
      '검토의견': v(run.opinion),
      '경고': (res.warnings.map(function (w) { return w.msg; }).concat(res.rows.filter(function (r) { return r.blocks.length; }).map(function (r) { return r.partNo + ' — ' + r.blocks.join(' / '); }))).join('\n') || '없음',
      'items': res.items.map(function (o) {
        return { '원소재': o.name + (o.spec ? ' (' + o.spec + ')' : ''), 'Index명': v(o.index), '단위': v(o.unit), '기준Index': o.i0 != null ? fmt(o.i0, 2) : NA, '변경Index': o.i1 != null ? fmt(o.i1, 2) : NA, '변동률': o.rawRate != null ? fmtPct(o.rawRate) : NA, '적용변동률': o.rate != null ? fmtPct(o.rate) : NA, '비고': (o.steps || []).concat(o.match && o.match.similar ? ['유사 Index'] : []).join(' / ') || '-' };
      }),
      'rows': res.rows.map(function (r) {
        var q = r.request || {};
        return {
          '품번': r.partNo, '품명': r.name || '-', '기준단가': fmt(r.basePrice), '현재단가': isNum(r.curPrice) ? fmt(r.curPrice) : NA,
          '사용량': r.lines.filter(function (l) { return l.kind !== 'fx'; }).map(function (l) { return l.item + ' ' + (l.kg == null ? NA : fmt(l.kg, 3) + 'kg'); }).join(', ') || '-',
          '원소재연동액': r.newPrice != null ? fmt(r.matAmount, 2) : NA, '환율연동액': r.newPrice != null ? fmt(r.fxAmountDelta, 2) : NA,
          '전체연동액': r.newPrice != null ? fmt(r.total, 2) : NA, '변경단가': r.newPrice != null ? fmt(r.newPrice, 2) : NA, '변동률': r.newPrice != null ? fmtPct(r.rate) : NA,
          '요청단가': isNum(q.price) ? fmt(q.price) : '-', '차액': isNum(q.gap) ? fmt(q.gap, 2) : '-', '의견': r.blocks.length ? '확인 필요 — ' + r.blocks.join(' / ') : (q.verdict || '-')
        };
      }),
      'details': [].concat.apply([], res.rows.map(function (r) {
        return r.lines.filter(function (l) { return !l.skip; }).map(function (l) {
          return { '품번': r.partNo, '원소재': l.item, '사용량': l.kg === '-' ? '-' : (l.kg == null ? NA : fmt(l.kg, 3)), '기준Index': l.i0 != null ? fmt(l.i0, 2) : NA, '변경Index': l.i1 != null ? fmt(l.i1, 2) : NA, '변동률': l.rate != null ? fmtPct(l.rate) : NA, '연동액': isNum(l.amount) ? fmt(l.amount, 2) : NA };
        });
      }))
    };
  }
  function exportSheets(run) {
    var d = reportData(run), res = run.result;
    var head = [['OEM 원소재 연동제 계산 결과'], ['업체', d['업체명'], '계약 버전', d['계약버전'], '적용 월', d['적용월']], ['기준 기간', d['기준기간'], '변경 기간', d['변경기간'], 'Index 파일', d['Index파일']], []];
    var s1 = head.concat([['품번', '품명', '기준 파트단가', '현재 구매단가', '원소재 연동액', '환율 연동액', '전체 연동액', '계약 계산 변경단가', '변동률(%)', '업체 요청단가', '요청 − 계산', '요청 인상률(%)', '계약 계산 인상률(%)', '검토']]);
    res.rows.forEach(function (r) { var q = r.request || {}; s1.push([r.partNo, r.name, r.basePrice, r.curPrice, r2(r.matAmount), r2(r.fxAmountDelta), r2(r.total), r2(r.newPrice), r2(r.rate), q.price == null ? '' : q.price, r2(q.gap), r2(q.reqRate), r2(q.calcRate), r.blocks.length ? '확인 필요: ' + r.blocks.join(' / ') : (q.verdict || '')]); });
    var s2 = [['품번', '연동 항목', '수식', '사용량(kg)', 'Index', '기준 Index', '변경 Index', '변동률(%)', '적용 변동률(%)', '연동액(원)', '계산 과정', '비고']];
    res.rows.forEach(function (r) { r.lines.forEach(function (l) { s2.push([r.partNo, l.item, l.kind === 'fx' ? '환율 100%' : (FORMULA[l.formula] || ''), l.kg, l.index, r2(l.i0, 4), r2(l.i1, 4), r2(l.rawRate), r2(l.rate), r2(l.amount), l.calc, l.note]); }); });
    var s3 = [['연동 항목', '규격', 'Index', '단위', '출처', '매칭', '기준 Index', '기준에 쓴 달', '변경 Index', '변경에 쓴 달', '변동률(%)', '적용 변동률(%)', '계약 조건 적용', '알림']];
    res.items.forEach(function (o) { s3.push([o.name, o.spec, o.index, o.unit, o.source, o.match ? (o.match.similar ? '유사 Index(승인)' : { approved: '담당자 승인', exact: '정확 일치', multi: '여러 개 — 선택 필요', similar: '유사 후보 — 승인 전', none: '없음' }[o.match.status]) : '', r2(o.i0, 4), o.base ? (o.base.fixed ? '계약 고정값' : o.base.used.join(' ')) : '', r2(o.i1, 4), o.next ? o.next.used.join(' ') : '', r2(o.rawRate), r2(o.rate), (o.steps || []).join(' / '), o.blocks.concat(o.notes).join(' / ')]); });
    var s4 = [['구분', '내용']].concat(res.blocks.map(function (b) { return ['계산 중단', b.msg]; })).concat(res.warnings.map(function (w) { return ['알림', w.msg]; })).concat(res.rows.filter(function (r) { return r.blocks.length || r.warnings.length; }).map(function (r) { return [r.partNo, r.blocks.concat(r.warnings).join(' / ')]; }));
    return { '품번별결과': s1, '계산내역': s2, 'Index근거': s3, '검증': s4 };
  }

  /* ───────────── 분기 · 반기 기간 도우미 ───────────── */
  /** 적용 월과 주기로 「직전 기간」과 「그 앞 기간」을 제안한다(제안일 뿐 — 계약서 기준일이 우선) */
  function suggestPeriod(applyMonth, cycle) {
    var n = cycle === 'half' ? 6 : cycle === 'year' ? 12 : 3, a = CostIdx.mIdx(applyMonth);
    return { next: { from: CostIdx.mStr(a - n), to: CostIdx.mStr(a - 1) }, base: { from: CostIdx.mStr(a - 2 * n), to: CostIdx.mStr(a - n - 1) } };
  }

  /* ───────────── 예제(전부 가짜) ───────────── */
  function sampleBook() {
    var b = CostIdx.sampleBook();
    // 업체 이름과 같은 시트 — 「업체 시트 우선 검색」을 보여 주려는 가짜 계약 지표
    var hot = CostIdx.buildBook(b, CostIdx.inspectBook(b), '').series.filter(function (s) { return /SS400/.test(s.name); })[0];
    var aoa = [['날짜', '철강 후판 SS400 계약지표 KRW/kg', '고무 천연고무 RSS3 KRW/kg']];
    hot.points.forEach(function (p, i) { aoa.push([p.m, Math.round(p.v * 1.03), Math.round(2100 + 260 * Math.sin(i / 7) + i * 4)]); });
    b.names = b.names.concat(['가상정밀']); b.sheets['가상정밀'] = aoa;
    return b;
  }
  function sample() {
    var v1 = { id: 'v-sample-1', name: '가상정밀(주)', sheetHint: '가상정밀', memo: '예제 — 실제 업체가 아닙니다' };
    var v2 = { id: 'v-sample-2', name: '예제주물(주)', sheetHint: '', memo: '예제 — 실제 업체가 아닙니다' };
    var mk = function (name, o) { var it = blankItem(name); return Object.assign(it, o || {}); };
    var c1 = Object.assign(blankContract(v1.id), {
      id: 'c-sample-1', version: 1, from: '2026-01', to: '2026-12', cycle: 'quarter', baseRule: '기준시점: 직전 조정 분기 3개월 평균 / 비교시점: 조정일 직전 분기 3개월 평균',
      firstDate: '2026-01-02', product: '지게차 마스트 브래킷 외 3종', indexText: '구매기획팀 배포 「주요 원자재 가격 동향」(가상)', trigger: 0.3, adjustDay: '매 분기 첫 영업일', reflectDay: '01-01, 04-01, 07-01, 10-01',
      formulaText: '변경 파트단가 = 기준 파트단가 + Σ(원소재 사용량 kg × (변경 Index − 기준 Index)) + 환율 대상 금액 × 환율 변동률',
      allowSubstitute: false, round: '1', file: '(예제) 가상정밀_연동계약_v1.docx',
      items: [mk('철강', { id: 'm-s1', spec: 'SS400', formula: 'weight' }), mk('구리', { id: 'm-s2', formula: 'weight' }), mk('고무', { id: 'm-s3', formula: 'weight' }), mk('환율', { id: 'm-s4', spec: 'USD', fxBase: 'amount' })]
    });
    var c2 = Object.assign(blankContract(v2.id), {
      id: 'c-sample-2', version: 1, from: '2026-01', to: '2026-12', cycle: 'half', baseRule: '기준시점: 직전 조정 반기 평균 / 비교시점: 조정일 직전 반기 평균',
      firstDate: '2026-01-02', product: '기어 케이스 외 1종', indexText: '구매기획팀 배포 「주요 원자재 가격 동향」(가상)', trigger: 0.5, adjustDay: '매 반기 첫 영업일', reflectDay: '01-01, 07-01',
      formulaText: '변경 파트단가 = 기준 파트단가 × (1 + Σ(원소재 비중 × 변동률)), 변동률 ±0.5% 미만은 연동하지 않고 상한 +10% · 하한 −10%',
      allowSubstitute: true, round: '10', file: '(예제) 예제주물_연동계약_v1.docx',
      items: [mk('알루미늄', { id: 'm-s5', formula: 'ratio', share: 45, cap: 10, floor: -10 }), mk('주철', { id: 'm-s6', spec: 'FC250', formula: 'ratio', share: 20, cap: 10, floor: -10 })]
    });
    var P = function (vid, no, nm, base, cur, usage, fx, fxAmt) { return { vendorId: vid, partNo: no, name: nm, basePrice: base, curPrice: cur, usage: usage, fx: !!fx, fxAmount: fxAmt == null ? null : fxAmt, memo: '' }; };
    var parts = [
      P(v1.id, 'A-1001', 'BRACKET, MAST', 100000, 101500, { '철강': 25, '구리': 3, '고무': 2 }, false),
      P(v1.id, 'A-1002', 'HOUSING, PUMP', 150000, 150000, { '철강': 30, '구리': 5, '고무': 0 }, true, 20000),
      P(v1.id, 'A-1003', 'COVER, SIDE', 80000, 80500, { '철강': 12, '구리': 0, '고무': 4 }, false),
      P(v1.id, 'A-1004', 'PLATE, BASE', 64000, null, { '철강': 18, '구리': null, '고무': 0 }, false),
      P(v2.id, 'C-2001', 'CASE, GEAR', 210000, 212000, {}, false),
      P(v2.id, 'C-2002', 'COVER, AXLE', 95000, 95000, {}, false)
    ];
    return {
      vendors: [v1, v2], contracts: [c1, c2], parts: parts, book: sampleBook(), indexFile: '원자재가격동향_예제(가상).xlsx',
      applyMonth: '2026-07', requests: { 'A-1001': 104000, 'A-1002': 151000, 'A-1003': 83000, 'C-2001': 221000 }
    };
  }

  return {
    FORMULA: FORMULA, CYCLE: CYCLE, FX_BASE: FX_BASE, EXPR_VARS: EXPR_VARS, NA: NA,
    uid: uid, key: key, cellNum: cellNum, isFxName: isFxName, clone: clone, r2: r2, fmt: fmt, fmtPct: fmtPct,
    blankVendor: blankVendor, blankItem: blankItem, blankContract: blankContract, blankPart: blankPart,
    inspectParts: inspectParts, parseParts: parseParts, partsTemplate: partsTemplate, checkPart: checkPart, headRole: headRole,
    extractContract: extractContract, standardSections: standardSections, evalExpr: evalExpr,
    wideInfo: wideInfo, buildIndex: buildIndex, headMonth: headMonth,
    unitOf: unitOf, perKgFactor: perKgFactor, avgOver: avgOver, monthsBetween: monthsBetween,
    sheetMatches: sheetMatches, matchItem: matchItem, familyOf: familyOf,
    contractProblems: contractProblems, applyTerms: applyTerms, calcRun: calcRun, termText: termText,
    reportData: reportData, exportSheets: exportSheets, suggestPeriod: suggestPeriod,
    sample: sample, sampleBook: sampleBook
  };
});
