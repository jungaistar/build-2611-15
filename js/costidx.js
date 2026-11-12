/* costidx.js — 도면 기반 원소재 가격 변동 · 제품 단가 영향 계산 엔진
 * 브라우저(window.CostIdx)와 node(require) 겸용. 화면(app.js)은 여기서 낸 값을 그리기만 한다.
 *
 * 값마다 출처를 붙인다 — doc(도면에서 읽음) · input(직접 입력) · calc(계산) · est(추정) · assume(가정 기본값).
 * 같은 숫자라도 신뢰도가 다르다. 출처가 사라지면 사람이 구분하지 못한 채 그대로 믿게 된다.
 *
 * Index 는 사용자가 올린 엑셀에 있는 것만 쓴다. 엑셀에 없는 Index 를 만들어 내지 않는다.
 */
(function (root, factory) {
  var api = factory();
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else root.CostIdx = api;
})(this, function () {
  'use strict';

  /* ───────────── 재질 계열 · 비중 ───────────── */

  // 비중(g/cm³) — 계열 대표값. 같은 계열 안의 작은 차이(SUS304 7.93 · SUS316 7.98)는 무시한다.
  var FAM = {
    hot:    { label: '열연강판(일반 구조용 탄소강)', density: 7.85 },
    cold:   { label: '냉연강판', density: 7.85 },
    coated: { label: '도금강판', density: 7.85 },
    carbon: { label: '기계구조용 탄소강', density: 7.85 },
    alloy:  { label: '합금강(특수강)', density: 7.85 },
    sts:    { label: '스테인리스강', density: 7.93 },
    al:     { label: '알루미늄 합금', density: 2.70 },
    cu:     { label: '구리', density: 8.96 },
    brass:  { label: '황동', density: 8.50 },
    bronze: { label: '청동', density: 8.80 },
    cast:   { label: '주철', density: 7.20 }
  };
  var FAM_KEYS = Object.keys(FAM);

  // 재질 기호(공백·하이픈·점 뺀 대문자) → 계열. 위에서부터 먼저 맞는 것.
  var MAT_RULES = [
    [/^(SCM|SCR|SNCM|SNC|SACM|SMN)\d{3}H?$|^41\d{2}H?$|^42CRMO4$|^SCM\d{3}H?QT$/, 'alloy'],
    // 규격 칸의 묶음 표기(SMxxC · SCRxxx · STSxxx)는 「계열」이다 — 개별 재질과 같다고 보지 않는다
    [/^(SCM|SCR|SNCM|SNC)X{3}H?$/, 'alloy'],
    [/^SM?X{2}C$/, 'carbon'],
    [/^(SUS|STS)X{3}L?$/, 'sts'],
    [/^Q(195|215|235|275|345|355)[A-E]?B?$/, 'hot'],
    [/^SM?\d{2}C$|^S\d{2}CK$|^C\d{2}E?$|^10\d{2}$/, 'carbon'],
    [/^(SS|SM|SMA|SHN|SN)\d{3}[A-C]?$|^SPH[CDE]$|^SAPH\d{3}$/, 'hot'],
    [/^SPC[CDE]N?$|^SPFC\d{3}$/, 'cold'],
    [/^S[EG]C[CDE]$|^SGH[CD]$|^SGCC$/, 'coated'],
    [/^(SUS|STS)\d{3}L?$|^3\d{2}L?$/, 'sts'],
    [/^AL?[1-7]\d{3}(T\d+|H\d+|O)?$|^[2567]\d{3}T\d+$/, 'al'],
    [/^C1[01]\d{2}$|^CU$/, 'cu'],
    [/^C2\d{3}$|^C3[5-7]\d{2}$/, 'brass'],
    [/^C[56]\d{3}$|^CAC\d{3}$|^BC\d$|^PBC\d$/, 'bronze'],
    [/^FCD?\d{3}$|^GC\d{3}$/, 'cast']
  ];
  // 낱말(한글·영문) → 계열. Index 열 이름에도 쓴다. 순서가 중요하다(「특수강(합금강)」은 합금강).
  var KW = [
    ['sts', /스테인리스|스텐|STAINLESS|\bSTS|\bSUS/],
    ['al', /알루미늄|ALUMIN/],
    ['brass', /황동|BRASS/],
    ['bronze', /청동|BRONZE/],
    ['cu', /구리|전기동|COPPER|\bCU\b|^동$/],
    ['alloy', /합금강|특수강|ALLOY|크롬몰리/],
    ['carbon', /기계구조용|탄소강|CARBON/],
    ['coated', /도금|아연|GALVA/],
    ['cold', /냉연|COLD|\bCR\b/],
    ['hot', /열연|HOT|\bHR\b|후판|일반구조용/],
    ['cast', /주철|주물|CAST/]
  ];

  // 정확히 일치하는 Index 가 없을 때 같은 계열로 볼 후보와 재질 유사도(0~1) · 근거.
  // 이 표만으로 고르지 않는다 — 가격 움직임(상관)을 함께 본다.
  var RELATED = {
    carbon: [['alloy', 0.8, '같은 기계구조용 봉강(특수강) 시장 — 합금 원소만 다름'], ['hot', 0.6, '같은 탄소강 — 철광석·철스크랩 원료 공유'], ['cold', 0.5, '같은 탄소강이나 판재 시장']],
    alloy: [['carbon', 0.8, '같은 기계구조용 봉강 시장'], ['hot', 0.55, '철강 원료 공유'], ['cold', 0.45, '철강 원료 공유, 판재 시장']],
    hot: [['cold', 0.8, '냉연은 열연을 다시 압연한 판재'], ['coated', 0.7, '같은 판재 계열'], ['carbon', 0.5, '같은 탄소강, 봉강 시장']],
    cold: [['hot', 0.8, '냉연의 소재가 열연'], ['coated', 0.8, '도금강판의 소재가 냉연'], ['carbon', 0.45, '같은 탄소강, 봉강 시장']],
    coated: [['cold', 0.85, '도금강판의 소재가 냉연강판'], ['hot', 0.65, '같은 판재 계열']],
    sts: [['alloy', 0.4, '합금 원소(크롬·니켈) 가격 영향'], ['hot', 0.3, '철강 원료 일부 공유']],
    al: [['cu', 0.35, '같은 비철금속 시장']],
    cu: [['brass', 0.7, '황동의 주성분이 구리'], ['bronze', 0.7, '청동의 주성분이 구리'], ['al', 0.35, '같은 비철금속 시장']],
    brass: [['cu', 0.85, '황동 성분의 60~70% 가 구리'], ['bronze', 0.7, '같은 구리 합금'], ['al', 0.3, '같은 비철금속 시장']],
    bronze: [['cu', 0.85, '청동 성분의 대부분이 구리'], ['brass', 0.75, '같은 구리 합금']],
    cast: [['hot', 0.5, '철 원료(선철·스크랩) 공유'], ['carbon', 0.45, '철 원료 공유']]
  };

  var DEF = { yield: 0.85, recovery: 0.3, cutMm: 3, tolPct: 1, minCorrN: 6, spreadPct: 10, minMatSim: 0.5 };

  function compact(s) { return String(s == null ? '' : s).toUpperCase().replace(/[\s\-_.·]/g, ''); }

  /** 재질 이름 → 계열. { raw, key, family, label, density, by } — 못 알아보면 family '' */
  function normMaterial(s) {
    var raw = String(s == null ? '' : s).trim();
    var up = raw.toUpperCase().replace(/[（]/g, '(').replace(/[）]/g, ')');
    var tokens = up.split(/[\s,()\/]+/).filter(Boolean);
    var tries = [compact(up)].concat(tokens.map(compact));
    for (var i = 0; i + 1 < tokens.length; i++) tries.push(compact(tokens[i] + tokens[i + 1]));
    for (var t = 0; t < tries.length; t++) {
      for (var r = 0; r < MAT_RULES.length; r++) {
        if (MAT_RULES[r][0].test(tries[t])) return mk(MAT_RULES[r][1], tries[t], '재질 기호');
      }
    }
    var fk = familyByWords(up);
    if (fk) return mk(fk, compact(up), '재질 낱말');
    return { raw: raw, key: compact(up), family: '', label: '', density: null, by: '' };
    function mk(f, key, by) { return { raw: raw, key: key, family: f, label: FAM[f].label, density: FAM[f].density, by: by }; }
  }
  function familyByWords(up) {
    up = String(up || '').toUpperCase();
    for (var k = 0; k < KW.length; k++) if (KW[k][1].test(up)) return KW[k][0];
    return '';
  }
  /** 재질 기호를 비교용 표준 표기로 — SCM440H → SCM440, S45C → SM45C, SUS304 · 304 → STS304 */
  function canonGrade(k, fam) {
    k = compact(k);
    if (fam === 'alloy') k = k.replace(/QT$/, '').replace(/(\d{3}|X{3})H$/, '$1');
    if (fam === 'carbon') { var m = /^SM?(\d{2}|XX)C$/.exec(k); if (m) k = 'SM' + m[1] + 'C'; }
    if (fam === 'sts') { var m2 = /^(?:SUS|STS)?(\d{3}|XXX)(L?)$/.exec(k); if (m2) k = 'STS' + m2[1] + m2[2]; }
    return k;
  }
  /** 원소재 이름의 대표 재질 기호(비교용). 낱말로만 알아본 재질(「특수강봉강」)은 기호가 없다 */
  function gradeOf(s) {
    var m = normMaterial(s);
    return m.by === '재질 기호' ? canonGrade(m.key, m.family) : '';
  }
  /** 글자(Index 규격 칸) 안의 재질 기호 전부 — 「SS275 마환봉 Φ31~80」 → [SS275], 「STS 304 CR」 → [STS304] */
  function gradeTokens(text) {
    var up = String(text == null ? '' : text).toUpperCase().replace(/[（]/g, '(').replace(/[）]/g, ')');
    var tokens = up.split(/[\s,()\/·]+/).filter(Boolean), tries = tokens.map(compact), out = [];
    for (var i = 0; i + 1 < tokens.length; i++) tries.push(compact(tokens[i] + tokens[i + 1]));
    tries.forEach(function (t) {
      if (/^\d+(\.\d+)?(MM|T)?$/.test(t)) return;          // 치수(2.0mm · 100)는 재질이 아니다
      for (var r = 0; r < MAT_RULES.length; r++) if (MAT_RULES[r][0].test(t)) {
        var g = canonGrade(t, MAT_RULES[r][1]);
        if (out.indexOf(g) < 0) out.push(g);
        break;
      }
    });
    // 「STS 304」 처럼 짝으로 읽힌 것이 있으면 숫자만 읽힌 304 → STS304 는 이미 같은 값이다
    return out;
  }
  /** Index 열 이름 → 계열. 재질 기호가 있으면 그것, 없으면 낱말. 모르면 '' */
  function indexFamily(name) {
    var m = normMaterial(name);
    return m.family;
  }

  /* ───────────── 치수 · 중량 ───────────── */

  function num(s) { var v = parseFloat(String(s).replace(/,/g, '')); return isFinite(v) ? v : null; }

  /** 「Ø45 x 210 L」「PL t12 x 150 x 300」「OD60 x t8 x 80」 → { shape, ... } 또는 null */
  function parseDims(text) {
    if (text == null) return null;
    var s = String(text).toUpperCase().replace(/[ΦφØø∅]/g, 'Ø').replace(/[×*＊]/g, 'X').replace(/(\d)\s*X\s*/g, '$1X').replace(/\s*X\s*(?=[\dTØ])/g, 'X');
    var m;
    if ((m = /(?:OD|PIPE\s*Ø?|Ø)\s*([\d.]+)\s*X\s*T\s*([\d.]+)\s*X\s*([\d.]+)/.exec(s)) && (/OD|PIPE/.test(s) || /XT/.test(s))) {
      var od = num(m[1]), t = num(m[2]), L = num(m[3]);
      if (od && t && L && 2 * t < od) return { shape: 'pipe', od: od, t: t, L: L };
    }
    if ((m = /(?:^|[^A-Z])(?:PL|PLATE)?\s*T\s*([\d.]+)\s*X\s*([\d.]+)\s*X\s*([\d.]+)/.exec(s))) {
      return { shape: 'plate', t: num(m[1]), w: num(m[2]), L: num(m[3]) };
    }
    if ((m = /(?:Ø|DIA\.?\s*|(?:^|[^A-Z])D)\s*([\d.]+)\s*X\s*([\d.]+)/.exec(s))) {
      return { shape: 'bar', d: num(m[1]), L: num(m[2]) };
    }
    return null;
  }
  function dimsText(d) {
    if (!d) return '';
    if (d.shape === 'bar') return 'Ø' + d.d + ' × ' + d.L + ' L';
    if (d.shape === 'plate') return 't' + d.t + ' × ' + d.w + ' × ' + d.L;
    if (d.shape === 'pipe') return 'OD' + d.od + ' × t' + d.t + ' × ' + d.L + ' L';
    return '';
  }
  var SHAPE_LABEL = { bar: '봉재', plate: '판재', pipe: '파이프' };

  /** 부피(mm³) × 비중(g/cm³) / 10⁶ = kg. extraL 은 길이 방향 절단 여유(mm) */
  function volumeMm3(d, extraL) {
    var L = d.L + (extraL || 0);
    if (d.shape === 'bar') return Math.PI / 4 * d.d * d.d * L;
    if (d.shape === 'plate') return d.t * d.w * L;
    if (d.shape === 'pipe') { var id = d.od - 2 * d.t; return Math.PI / 4 * (d.od * d.od - id * id) * L; }
    return null;
  }
  function weightKg(d, density, extraL) {
    if (!d || !density) return null;
    var v = volumeMm3(d, extraL);
    return v == null ? null : v * density / 1e6;
  }

  /* ───────────── 도면 글자에서 후보 뽑기 ───────────── */

  var FIELD_RX = {
    name: /(?:품\s*명|부\s*품\s*명|PART\s*NAME|DESCRIPTION|TITLE)\s*[:：]?\s*(.+)$/i,
    dwg: /(?:도면\s*번호|도\s*번|DWG\.?\s*NO\.?|DRAWING\s*NO\.?)\s*[:：]?\s*([A-Z0-9][A-Z0-9\-_.\/]*)/i,
    material: /(?:재\s*질|재\s*료|MATERIAL|MAT'?L)\s*[:：]?\s*(.+)$/i,
    weight: /(?:중\s*량|무\s*게|WEIGHT|MASS|\bWT\.?)\s*[:：]?\s*(?:약\s*|APPROX\.?\s*)?([\d.,]+)\s*(KG|㎏|G\b|그램)?/i,
    qty: /(?:수\s*량|Q'?TY|QUANTITY)\s*[:：]?\s*(\d+)/i,
    stock: /(?:소재\s*규격|소재\s*치수|STOCK(?:\s*SIZE)?|BLANK\s*SIZE|RAW\s*SIZE)\s*[:：]?\s*(.+)$/i,
    finished: /(?:완성\s*치수|규\s*격|치\s*수|FINISHED\s*SIZE|SIZE|DIMENSIONS?)\s*[:：]?\s*(.+)$/i
  };
  function firstSeg(v) { return String(v).split(/\s{2,}|\t/)[0].trim(); }

  /** 도면 PDF 의 글자(줄) → 필드별 후보. 후보마다 몇째 줄에서 읽었는지 남긴다. */
  function extractDrawing(text, fileName) {
    var lines = String(text || '').split(/\r?\n/);
    var out = { file: fileName || '', empty: !String(text || '').trim(), fields: {}, notes: [] };
    Object.keys(FIELD_RX).forEach(function (k) { out.fields[k] = []; });
    lines.forEach(function (line, i) {
      var L = line.trim(); if (!L) return;
      var used = {};
      ['dwg', 'material', 'weight', 'qty', 'stock'].forEach(function (k) {
        var m = FIELD_RX[k].exec(L); if (!m) return;
        var c = { line: i + 1, text: L };
        if (k === 'weight') {
          var v = num(m[1]); if (v == null) return;
          var unit = (m[2] || '').toUpperCase();
          if (unit === 'G' || unit === '그램') v = v / 1000;
          c.value = v; c.unit = unit ? '' : '단위 없음 — kg 로 봄';
        } else if (k === 'qty') c.value = parseInt(m[1], 10);
        else if (k === 'stock') { c.value = firstSeg(m[1]); c.dims = parseDims(c.value); if (!c.dims) return; }
        else if (k === 'material') { c.value = firstSeg(m[1]); c.norm = normMaterial(c.value); }
        else c.value = m[1];
        out.fields[k].push(c); used[k] = 1;
      });
      if (!used.stock) {
        var f = FIELD_RX.finished.exec(L);
        if (f) { var d = parseDims(f[1]); if (d) out.fields.finished.push({ line: i + 1, text: L, value: firstSeg(f[1]), dims: d }); }
      }
      var n = FIELD_RX.name.exec(L);
      if (n && !used.dwg) out.fields.name.push({ line: i + 1, text: L, value: firstSeg(n[1]) });
    });
    if (out.empty) out.notes.push('글자가 없는 PDF(스캔 도면)입니다 — 재질과 중량을 직접 입력하세요.');
    else {
      if (!out.fields.material.length) out.notes.push('재질을 찾지 못했습니다 — 직접 입력하세요.');
      if (!out.fields.weight.length && !out.fields.finished.length && !out.fields.stock.length) out.notes.push('중량·치수를 찾지 못했습니다 — 직접 입력하세요.');
      if (out.fields.material.length > 1) out.notes.push('재질 후보가 ' + out.fields.material.length + '개입니다 — 맞는 것을 확인하세요.');
    }
    return out;
  }

  /** pdf.js getTextContent().items → 줄 글자. y 가 비슷한 것끼리 한 줄, x 순서. 칸 사이가 벌어지면 두 칸 띄움 */
  function pdfLines(items) {
    var lines = [];
    (items || []).forEach(function (it) {
      if (!it.str || !it.str.trim()) return;
      var y = it.transform[5], line = null;
      for (var i = 0; i < lines.length; i++) if (Math.abs(lines[i].y - y) <= 2) { line = lines[i]; break; }
      if (!line) { line = { y: y, items: [] }; lines.push(line); }
      line.items.push({ x: it.transform[4], w: it.width || 0, s: it.str });
    });
    lines.sort(function (a, b) { return b.y - a.y; });
    return lines.map(function (l) {
      l.items.sort(function (a, b) { return a.x - b.x; });
      var out = '', endX = null;
      l.items.forEach(function (it) {
        if (endX == null) out = it.s.trim();
        else out += (it.x - endX > 6 ? '  ' : (/\s$/.test(out) || /^\s/.test(it.s) ? '' : ' ')) + it.s.trim();
        endX = it.x + it.w;
      });
      return out.trim();
    });
  }

  /** 추출 결과 → 품목 한 개(첫 후보를 기본으로, 출처 doc). 못 읽은 칸은 비워 둔다. */
  function itemFromDrawing(ex, id) {
    var f = ex.fields, g = function (k) { return f[k][0]; };
    var line = blankLine();
    if (g('material')) { line.material = g('material').value; line.src.material = 'doc'; }
    if (g('weight')) { line.finishedKg = g('weight').value; line.src.finishedKg = 'doc'; }
    if (g('finished')) { line.finishedDims = dimsText(g('finished').dims); line.src.finishedDims = 'doc'; }
    if (g('stock')) { line.stockDims = dimsText(g('stock').dims); line.src.stockDims = 'doc'; }
    var it = blankItem(id);
    it.file = ex.file;
    it.name = g('name') ? g('name').value : (ex.file || '').replace(/\.pdf$/i, ''); it.src.name = g('name') ? 'doc' : 'input';
    if (g('dwg')) { it.dwg = g('dwg').value; it.src.dwg = 'doc'; }
    if (g('qty')) { it.qty = g('qty').value; it.src.qty = 'doc'; }
    it.lines = [line];
    it.extract = ex;
    return it;
  }
  function blankLine() {
    return { material: '', family: '', finishedKg: null, finishedDims: '', stockDims: '', cutMm: null, yield: null, recovery: null, unitPrice: null, indexPick: '', refName: '', src: {} };
  }
  function blankItem(id) {
    return { id: id || 'item', name: '', dwg: '', qty: null, price0: null, priceNew: null, compMode: 'ratio', comp: { material: null, process: null, labor: null, overhead: null }, lines: [blankLine()], src: {} };
  }

  /* ───────────── Index 엑셀 읽기 ───────────── */

  function pad2(n) { return (n < 10 ? '0' : '') + n; }
  function mOf(y, m) { return y + '-' + pad2(m); }
  function mIdx(s) { var p = String(s).split('-'); return (+p[0]) * 12 + (+p[1] - 1); }
  function mStr(i) { return mOf(Math.floor(i / 12), (i % 12) + 1); }

  /** 날짜 칸 → 'YYYY-MM'. 엑셀 일련번호 · Date · '2022-03' · '2022.03.15' · '2022년 3월' */
  function parseMonth(v) {
    if (v == null || v === '') return null;
    if (v instanceof Date && !isNaN(v)) return mOf(v.getFullYear(), v.getMonth() + 1);
    if (typeof v === 'number') {
      if (v > 20000 && v < 80000) { var d = new Date(Math.round((v - 25569) * 86400000)); return mOf(d.getUTCFullYear(), d.getUTCMonth() + 1); }
      return null;
    }
    var s = String(v).trim(), m;
    if ((m = /^(\d{4})\s*[-./년]\s*(\d{1,2})(?:\s*[-./월]\s*(\d{1,2})?\s*일?)?\s*월?\.?$/.exec(s))) {
      var mo = +m[2]; if (mo >= 1 && mo <= 12) return mOf(+m[1], mo);
    }
    if ((m = /^(\d{4})(\d{2})(\d{2})?$/.exec(s))) { var mo2 = +m[2]; if (mo2 >= 1 && mo2 <= 12 && +m[1] > 1900) return mOf(+m[1], mo2); }
    return null;
  }
  function parseVal(v) {
    if (v == null) return null;
    if (typeof v === 'number') return isFinite(v) ? v : null;
    var s = String(v).trim().replace(/,/g, '');
    if (!s || !/^-?\d+(\.\d+)?$/.test(s)) return null;
    return parseFloat(s);
  }
  var DATE_HEAD = /날짜|일자|기준월|년월|연월|월|기간|DATE|MONTH|PERIOD/i;

  /** 시트(행 배열) 구조 확인 — 머리줄 · 날짜 열 · 값 열 후보와 계열 추정. 사람이 고칠 수 있게 제안만 한다. */
  function inspectSheet(aoa) {
    aoa = aoa || [];
    var width = aoa.reduce(function (w, r) { return Math.max(w, (r || []).length); }, 0);
    var headerRow = -1;
    for (var r = 0; r < Math.min(aoa.length, 20) && headerRow < 0; r++) {
      var row = aoa[r] || [];
      var texts = row.filter(function (c) { return typeof c === 'string' && c.trim() && parseVal(c) == null && !parseMonth(c); }).length;
      if (texts < 2) continue;
      for (var r2 = r + 1; r2 < Math.min(aoa.length, r + 4); r2++) {
        var nx = aoa[r2] || [];
        if (nx.some(function (c) { return parseMonth(c); }) && nx.some(function (c) { return parseVal(c) != null; })) { headerRow = r; break; }
      }
    }
    if (headerRow < 0) return { ok: false, error: '날짜 열과 값 열이 있는 머리줄을 찾지 못했습니다 — 첫 줄에 「날짜 · Index 이름」이 있어야 합니다.' };
    var head = aoa[headerRow] || [], body = aoa.slice(headerRow + 1);
    var best = -1, bestScore = -1, cols = [];
    for (var c = 0; c < width; c++) {
      var dates = 0, nums = 0, filled = 0;
      body.forEach(function (row) { var v = (row || [])[c]; if (v === '' || v == null) return; filled++; if (parseMonth(v)) dates++; if (parseVal(v) != null) nums++; });
      var name = String(head[c] == null ? '' : head[c]).trim();
      var score = dates + (DATE_HEAD.test(name) ? body.length * 0.5 : 0);
      if (dates > 0 && score > bestScore) { bestScore = score; best = c; }
      cols.push({ col: c, name: name || ('열 ' + (c + 1)), dates: dates, nums: nums, filled: filled });
    }
    if (best < 0) return { ok: false, error: '날짜로 읽히는 열이 없습니다.' };
    cols.forEach(function (o) {
      o.family = indexFamily(o.name);
      o.role = o.col === best ? 'date' : (o.nums >= Math.max(3, o.filled * 0.6) ? 'index' : 'skip');
    });
    return { ok: true, headerRow: headerRow, dateCol: best, columns: cols, rows: body.length };
  }

  /**
   * 시트 + 매핑 → 월 단위 계열. map = { headerRow, dateCol, columns:[{col,name,role:'index'|'ref'|'skip',family}] }
   * 같은 달 값이 여러 개면 월평균. 빈 칸·글자는 그 달을 비우고 issues 에 남긴다.
   */
  function buildSeries(aoa, map, source) {
    var issues = [], series = [];
    var use = map.columns.filter(function (c) { return c.role === 'index' || c.role === 'ref'; });
    var acc = use.map(function () { return {}; });
    for (var r = map.headerRow + 1; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var any = row.some(function (v) { return v !== '' && v != null; });
      if (!any) continue;
      var m = parseMonth(row[map.dateCol]);
      if (!m) { issues.push({ row: r + 1, msg: (r + 1) + '행 날짜를 읽지 못해 뺐습니다(' + String(row[map.dateCol] == null ? '빈 칸' : row[map.dateCol] || '빈 칸') + ')' }); continue; }
      use.forEach(function (c, k) {
        var v = parseVal(row[c.col]);
        if (v == null) { issues.push({ row: r + 1, col: c.name, msg: c.name + ' ' + m + ' 값이 비었거나 숫자가 아님 — 그 달은 비워 둡니다' }); return; }
        (acc[k][m] = acc[k][m] || []).push(v);
      });
    }
    use.forEach(function (c, k) {
      var pts = Object.keys(acc[k]).sort().map(function (m) {
        var a = acc[k][m];
        if (a.length > 1) issues.push({ col: c.name, msg: c.name + ' ' + m + ' 값 ' + a.length + '개 → 월평균' });
        return { m: m, v: a.reduce(function (s, x) { return s + x; }, 0) / a.length };
      });
      series.push({ name: c.name, family: c.family || '', role: c.role, col: c.col, points: pts, source: source || '' });
    });
    return { series: series, issues: issues };
  }

  /* ───────────── 통합문서(머리 블록) 형식 ─────────────
   * 실무 원자재 가격 통합문서: 1행 제목, A열 낱말(Item · 지역 · 기준 · 단위 · 규격 · 출처 · 상세출처 · 주소)이 붙은 머리 블록,
   * 그 아래 최신 월이 위인 데이터(A=연 · B=월 · C=분기, 또는 A=「2026-5」). 값 열 사이에 전월 대비 증감률 열이 끼어 있다.
   * 품목명은 가로 병합이라 그룹 첫 열에만 있다 — 오른쪽 열은 물려받는다. 단위가 적힌 열만 Index 후보다. */

  var LABELS = { 'ITEM': 'item', '품목': 'item', '품명': 'item', '지역': 'region', '국가': 'region', '기준': 'basis', '단위': 'unit', '규격': 'spec', '출처': 'source', '상세출처': 'detail', '주소': 'url', 'URL': 'url', 'LINK': 'url' };
  function cellText(v) { return String(v == null ? '' : v).replace(/\s+/g, ' ').trim(); }
  function labelOf(v) { return LABELS[cellText(v).replace(/\s/g, '').toUpperCase()] || ''; }
  function yearOf(v) {
    if (typeof v === 'number') return v > 1900 && v < 2200 && v === Math.floor(v) ? v : null;
    var m = /^(\d{4})\s*년?$/.exec(cellText(v)); return m ? +m[1] : null;
  }
  function monthOf(v) {
    if (typeof v === 'number') return v >= 1 && v <= 12 && v === Math.floor(v) ? v : null;
    var m = /^(\d{1,2})\s*월$/.exec(cellText(v)); return m && +m[1] >= 1 && +m[1] <= 12 ? +m[1] : null;
  }
  var QUARTER = /^\d{2}\.\s*[1-4]\s*Q$/i;
  function colLetter(c) { var s = ''; c++; while (c > 0) { var r = (c - 1) % 26; s = String.fromCharCode(65 + r) + s; c = Math.floor((c - 1) / 26); } return s; }
  function sheetKind(name, title) {
    var t = name + ' ' + title;
    if (/환율|EXCHANGE/i.test(t)) return 'fx';
    if (/PMI|PPI|운임|BDI|물가|경기/i.test(t)) return 'macro';
    return 'material';
  }

  /** 데이터 행의 달 — { m, width } . 연이 위 행에만 있으면(세로 병합) carry 로 물려받는다 */
  function rowMonth(row, carry) {
    var a = row[0], b = row[1], y = yearOf(a), mo = monthOf(b);
    if (mo && (y || (carry && cellText(a) === ''))) return { m: mOf(y || carry, mo), y: y || carry, width: 2 };
    if (typeof a === 'string' && !y) { var pm = parseMonth(a); if (pm) return { m: pm, width: 1 }; }
    return null;
  }

  /** 머리 블록 형식인지 보고 열마다 이름 · 단위 · 출처 · 역할을 정한다. 형식이 아니면 { ok:false } */
  function inspectBlock(aoa, sheetName) {
    aoa = aoa || [];
    var first = -1, labels = {}, width = 0, carry = null;
    for (var r = 0; r < Math.min(aoa.length, 40); r++) {
      var row = aoa[r] || [];
      width = Math.max(width, row.length);
      var lb = labelOf(row[0]);
      if (lb && labels[lb] == null && first < 0) { labels[lb] = r; continue; }
      if (rowMonth(row, null)) { first = r; break; }
    }
    if (first < 0) return { ok: false, error: '데이터 행(연·월 또는 날짜)을 찾지 못했습니다' };
    var first2 = rowMonth(aoa[first], null);
    if (!Object.keys(labels).length && first2.width === 1) return { ok: false, error: '머리 블록이 없는 단순 표' };
    for (r = first; r < aoa.length; r++) width = Math.max(width, (aoa[r] || []).length);
    var dw = first2.width, title = '';
    for (r = 0; r < first; r++) if (!labelOf((aoa[r] || [])[0]) && cellText((aoa[r] || [])[0])) { title = cellText(aoa[r][0]); break; }
    if (labels.item == null && labels.region == null) {   // Item · 지역 줄이 모두 없으면 글자가 있는 마지막 머리 줄을 이름 줄로
      for (r = first - 1; r >= 0; r--) { var hr = aoa[r] || []; if (hr.slice(dw).some(function (c) { return cellText(c); })) { labels.item = r; break; } }
    }
    var kind = sheetKind(sheetName || '', title);
    var get = function (key, c) { return labels[key] == null ? '' : cellText((aoa[labels[key]] || [])[c]); };
    var data = aoa.slice(first);
    var cols = [], skipped = { rate: 0, quarter: 0, empty: 0 }, groupStart = dw, lastItem = '';
    for (var c = dw; c < width; c++) {
      var own = get('item', c);
      if (own) { groupStart = c; lastItem = own; }
      var vals = data.map(function (row) { return (row || [])[c]; }).filter(function (v) { return v !== '' && v != null; });
      var q = vals.filter(function (v) { return QUARTER.test(cellText(v)); }).length;
      var nums = vals.filter(function (v) { return parseVal(v) != null; }).length;
      if (vals.length && q >= vals.length * 0.6) { skipped.quarter++; continue; }
      if (!nums) { skipped.empty++; continue; }
      var unit = get('unit', c);
      var name0 = own || lastItem;
      var named = own || get('region', c);
      if (labels.unit != null ? !unit : !named) { skipped.rate++; continue; }   // 단위(단위 줄이 없으면 이름)가 빈 숫자 열 = 증감률
      var inh = function (key) { var v = get(key, c); for (var k = c - 1; !v && k >= groupStart; k--) v = get(key, k); return v; };
      var o = { sheet: sheetName || '', col: c, letter: colLetter(c), item: name0, region: inh('region'), basis: inh('basis'), unit: unit, spec: inh('spec'), source: inh('source'), detail: inh('detail'), url: inh('url'), nums: nums, filled: vals.length };
      o.name = [o.item, o.spec, o.basis, o.region, o.unit].filter(Boolean).join(' · ');
      o.grades = gradeTokens(o.spec);
      var fm = normMaterial(o.spec);
      o.family = fm.family || indexFamily(o.item) || '';
      o.inhouse = /구매팀|사내|구매가|자사/.test(o.source + ' ' + o.detail);
      o.kind = kind;
      o.role = kind !== 'material' ? 'skip' : (o.inhouse ? 'ref' : 'index');
      o.why = kind === 'fx' ? '환율 — 원화 환산에만 씀' : kind === 'macro' ? '재료 가격이 아님 — 참고 계열로 바꿀 수 있음' : (o.inhouse ? '사내 구매가 — 참고 계열(매칭 비교만)' : '');
      cols.push(o);
    }
    // 같은 품목 · 규격 · 지역에 원화 열이 있으면 외화 열은 기본 「쓰지 않음」(비철의 USD/ton ↔ KRW/kg 환산 짝)
    cols.forEach(function (o) {
      if (o.role !== 'index' || /^KRW/i.test(o.unit)) return;
      var twin = cols.filter(function (x) { return x !== o && x.item === o.item && x.spec === o.spec && x.region === o.region && /^KRW/i.test(x.unit); })[0];
      if (twin) { o.role = 'skip'; o.why = '같은 품목의 원화 열(' + twin.unit + ')을 우선 — 바꿀 수 있음'; }
    });
    var months = [];
    carry = null;
    data.forEach(function (row) { var rm = rowMonth(row || [], carry); if (rm) { if (rm.y) carry = rm.y; months.push(rm.m); } });
    return {
      ok: true, kind: 'block', sheet: sheetName || '', sheetKind: kind, title: title, firstRow: first, dateWidth: dw, labels: labels, columns: cols, skipped: skipped,
      newestFirst: months.length > 1 && months[0] > months[months.length - 1],
      firstMonth: months.slice().sort()[0] || '', lastMonth: months.slice().sort().pop() || '', rows: months.length
    };
  }

  /** 머리 블록 시트 → 월 단위 계열(meta 에 품목 · 규격 · 기준 · 지역 · 단위 · 출처를 담는다). 역할이 skip 인 열도 돌려준다(환율 환산용) */
  function buildBlockSeries(aoa, blk, file) {
    var issues = [], acc = blk.columns.map(function () { return {}; }), carry = null;
    for (var r = blk.firstRow; r < aoa.length; r++) {
      var row = aoa[r] || [];
      var rm = rowMonth(row, carry);
      if (!rm) {
        if (row.slice(blk.dateWidth).some(function (v) { return parseVal(v) != null; })) issues.push({ row: r + 1, msg: blk.sheet + ' ' + (r + 1) + '행 날짜를 읽지 못해 뺐습니다' });
        continue;
      }
      if (rm.y) carry = rm.y;
      blk.columns.forEach(function (c, k) {
        var raw = row[c.col], v = parseVal(raw);
        if (v == null) { if (c.role !== 'skip' && cellText(raw) !== '') issues.push({ row: r + 1, col: c.name, msg: blk.sheet + ' · ' + c.name + ' ' + rm.m + ' 값이 숫자가 아님(' + cellText(raw) + ')' }); return; }
        (acc[k][rm.m] = acc[k][rm.m] || []).push(v);
      });
    }
    var series = blk.columns.map(function (c, k) {
      var pts = Object.keys(acc[k]).sort().map(function (m) { var a = acc[k][m]; return { m: m, v: a.reduce(function (s, x) { return s + x; }, 0) / a.length }; });
      return {
        name: c.name, family: c.family, role: c.role, col: c.col, points: pts, grades: c.grades,
        source: (file ? file + ' / ' : '') + '시트 ' + blk.sheet + ' / ' + c.letter + '열',
        meta: { sheet: blk.sheet, kind: c.kind, item: c.item, spec: c.spec, basis: c.basis, region: c.region, unit: c.unit, source: c.source, detail: c.detail, url: c.url, inhouse: c.inhouse, letter: c.letter }
      };
    });
    return { series: series, issues: issues };
  }

  /** 통합문서 전체 확인 — 시트마다 머리 블록 형식이면 inspectBlock, 아니면 단순 표(inspectSheet). 이름이 겹치면 시트 이름을 붙인다 */
  function inspectBook(book) {
    var sheets = [], seen = {};
    (book.names || Object.keys(book.sheets)).forEach(function (n) {
      var aoa = book.sheets[n], b = inspectBlock(aoa, n);
      if (!b.ok) {
        var s = inspectSheet(aoa);
        if (!s.ok) { sheets.push({ ok: false, sheet: n, error: s.error }); return; }
        s.kind = 'simple'; s.sheet = n; s.sheetKind = sheetKind(n, '');
        s.columns.forEach(function (c) { c.sheet = n; c.letter = colLetter(c.col); c.kind = s.sheetKind; c.grades = gradeTokens(c.name); });
        b = s;
      }
      sheets.push(b);
    });
    sheets.forEach(function (sh) { (sh.columns || []).forEach(function (c) { if (c.role === 'date') return; if (seen[c.name]) c.name = c.name + ' [' + sh.sheet + ']'; seen[c.name] = 1; }); });
    return { sheets: sheets };
  }
  /** 확인 결과(사람이 고친 역할 · 계열 포함) → 계열 전부 */
  function buildBook(book, insp, file) {
    var series = [], issues = [];
    insp.sheets.forEach(function (sh) {
      if (!sh.ok) { issues.push({ msg: '시트 「' + sh.sheet + '」 — ' + sh.error }); return; }
      var r;
      if (sh.kind === 'block') r = buildBlockSeries(book.sheets[sh.sheet], sh, file);
      else {
        r = buildSeries(book.sheets[sh.sheet], sh, (file ? file + ' / ' : '') + '시트 ' + sh.sheet);
        r.series.forEach(function (s) { var c = sh.columns.filter(function (x) { return x.col === s.col; })[0]; s.grades = c.grades; s.meta = { sheet: sh.sheet, kind: sh.sheetKind, item: c.name, unit: '', letter: c.letter }; s.source += ' / 열 「' + s.name + '」'; });
      }
      series = series.concat(r.series); issues = issues.concat(r.issues);
    });
    return { series: series, issues: issues };
  }

  /** USD · RMB 단위 Index 를 같은 통합문서 환율 시트로 원화 환산(달마다 그 달 환율). 바꾼 값은 계산 값이다 */
  function fxConvert(series) {
    var fx = series.filter(function (s) { return s.meta && s.meta.kind === 'fx'; });
    var find = function (rx) { return fx.filter(function (s) { return rx.test(s.name); })[0]; };
    var usd = find(/USD|달러/i), cny = find(/CNY|RMB|위안/i), notes = [];
    var out = series.map(function (s) {
      var u = s.meta && s.meta.unit || '';
      var cur = /^USD/i.test(u) ? 'USD' : /^(RMB|CNY)/i.test(u) ? 'CNY' : '';
      if (!cur || s.meta.kind === 'fx' || s.role === 'skip') return s;
      var per = /\/\s*(TON|T|MT)\b/i.test(u) ? 'ton' : /\/\s*KG/i.test(u) ? 'kg' : '';
      var rate = cur === 'USD' ? usd : cny;
      if (!rate || !per) { notes.push(s.name + ' — ' + (!rate ? cur + ' 환율 열이 없어' : '단위(' + u + ')를 알 수 없어') + ' 환산하지 않았습니다'); return s; }
      var byM = {}; rate.points.forEach(function (p) { byM[p.m] = p.v; });
      var miss = 0, pts = [];
      s.points.forEach(function (p) { if (byM[p.m] == null) { miss++; return; } pts.push({ m: p.m, v: p.v * byM[p.m] }); });
      if (miss) notes.push(s.name + ' — 환율이 없는 달 ' + miss + '개는 환산 계열에서 뺐습니다');
      var meta = Object.assign({}, s.meta, { unit: 'KRW/' + per + '(환산)', fx: rate.name, origUnit: u });
      return Object.assign({}, s, { name: s.name + ' → 원화 환산', points: pts, meta: meta, converted: true, source: s.source + ' × ' + rate.source });
    });
    return { series: out, notes: notes };
  }

  /* ───────────── 기간 값 · 변동률 ───────────── */

  /** 그 달 값. 없으면 가장 가까운 이전 관측월, 이전이 없으면 가장 가까운 다음 관측월(표시) */
  function valueAt(s, month) {
    var pts = s.points, t = mIdx(month), prev = null, next = null;
    for (var i = 0; i < pts.length; i++) {
      var k = mIdx(pts[i].m);
      if (k === t) return { m: pts[i].m, v: pts[i].v, kind: 'exact' };
      if (k < t) prev = pts[i];
      else if (!next) next = pts[i];
    }
    if (prev) return { m: prev.m, v: prev.v, kind: 'prev' };
    if (next) return { m: next.m, v: next.v, kind: 'next' };
    return null;
  }
  function inPeriod(s, start, end) {
    var a = mIdx(start), b = mIdx(end);
    return s.points.filter(function (p) { var k = mIdx(p.m); return k >= a && k <= b; });
  }
  /** 기간 시작·종료 값, 변동액·률, 기간 최고·최저, 오른 달·내린 달 수, 월별 변동 */
  function periodStats(s, start, end) {
    var a = valueAt(s, start), b = valueAt(s, end), notes = [];
    if (!a || !b) return { ok: false, notes: ['값이 없습니다'] };
    if (a.kind !== 'exact') notes.push('시작 ' + start + ' 값이 없어 ' + a.m + ' 값 사용' + (a.kind === 'next' ? ' (기간 시작 전 데이터 없음 — 확인 필요)' : ''));
    if (b.kind !== 'exact') notes.push('종료 ' + end + ' 값이 없어 ' + b.m + ' 값 사용' + (b.kind === 'next' ? ' (확인 필요)' : ''));
    var pts = inPeriod(s, start, end), months = [], up = 0, down = 0, hi = null, lo = null;
    pts.forEach(function (p, i) {
      var prev = i > 0 && mIdx(pts[i - 1].m) === mIdx(p.m) - 1 ? pts[i - 1] : null;
      var chg = prev ? (p.v / prev.v - 1) * 100 : null;
      if (chg != null) { if (chg > 0) up++; else if (chg < 0) down++; }
      if (!hi || p.v > hi.v) hi = p;
      if (!lo || p.v < lo.v) lo = p;
      months.push({ m: p.m, v: p.v, chg: chg });
    });
    var missing = mIdx(end) - mIdx(start) + 1 - pts.length;
    if (missing > 0) notes.push('기간 안에 값이 빈 달 ' + missing + '개');
    return {
      ok: true, start: a, end: b, diff: b.v - a.v, rate: (b.v / a.v - 1) * 100,
      max: hi, min: lo, up: up, down: down, months: months, notes: notes
    };
  }

  function pearson(x, y) {
    var n = Math.min(x.length, y.length);
    if (n < 3) return null;
    var mx = 0, my = 0, i;
    for (i = 0; i < n; i++) { mx += x[i]; my += y[i]; }
    mx /= n; my /= n;
    var sxy = 0, sxx = 0, syy = 0;
    for (i = 0; i < n; i++) { var dx = x[i] - mx, dy = y[i] - my; sxy += dx * dy; sxx += dx * dx; syy += dy * dy; }
    if (sxx === 0 || syy === 0) return null;
    return sxy / Math.sqrt(sxx * syy);
  }
  /** 기간 안 월간 변동률(%) — 앞 달 관측이 있는 달만. {m: chg} */
  function monthlyChanges(s, start, end) {
    var out = {};
    var a = mIdx(start), b = mIdx(end), byM = {};
    s.points.forEach(function (p) { byM[mIdx(p.m)] = p.v; });
    for (var k = a + 1; k <= b; k++) if (byM[k] != null && byM[k - 1] != null) out[mStr(k)] = (byM[k] / byM[k - 1] - 1) * 100;
    return out;
  }
  /** 두 월간 변동 계열의 상관 · 방향 일치율 · 비교한 달 수 */
  function comparePattern(ca, cb, minN) {
    var ks = Object.keys(ca).filter(function (k) { return cb[k] != null; }).sort();
    var x = ks.map(function (k) { return ca[k]; }), y = ks.map(function (k) { return cb[k]; });
    var same = 0;
    ks.forEach(function (k) { if ((ca[k] > 0 && cb[k] > 0) || (ca[k] < 0 && cb[k] < 0) || (ca[k] === 0 && cb[k] === 0)) same++; });
    var n = ks.length;
    return { n: n, corr: n >= (minN || DEF.minCorrN) ? pearson(x, y) : null, dirAgree: n ? same / n : null };
  }
  function avgChanges(list) {
    var out = {}, cnt = {};
    list.forEach(function (c) { Object.keys(c).forEach(function (k) { out[k] = (out[k] || 0) + c[k]; cnt[k] = (cnt[k] || 0) + 1; }); });
    Object.keys(out).forEach(function (k) { out[k] /= cnt[k]; });
    return out;
  }

  /* ───────────── Index 매칭 ───────────── */

  /** 지역(한국 우선) · 단위(KRW/kg 우선) 순위 — 작을수록 먼저 */
  function prefRank(s) {
    var m = s.meta || {};
    var r = /한국|국내|KOREA|KR\b/i.test(m.region || '') ? 0 : (m.region ? 2 : 1);
    var u = m.unit || '';
    var uu = /^KRW\s*\/\s*KG/i.test(u) ? 0 : /^KRW/i.test(u) ? 1 : (u ? 3 : 2);
    return r * 10 + uu;
  }
  function wildRe(p) { return new RegExp('^' + p.replace(/X/g, '\\d') + '$'); }

  /**
   * 원소재 계열 → 적용 Index.
   *  exact    : 계열이 같은 Index 하나 → 그것만 쓴다(다른 후보를 임의로 고르지 않는다)
   *  multi    : 계열이 같은 Index 가 여럿 → 담당자가 고를 때까지 계산하지 않는다
   *  estimate : 정확 일치 없음 → 같은 계열 후보를 재질 유사도 + 가격 움직임(월간 변동 상관)으로 순위
   *  none     : 후보 없음 → 분석 불가(임의로 고르지 않는다)
   *  unknown  : 재질 계열을 모름
   *  manual   : 담당자가 직접 고른 Index
   */
  function matchIndex(family, series, period, opt) {
    opt = opt || {};
    var idx = series.filter(function (s) { return s.role !== 'ref' && s.role !== 'skip'; });
    var res = { family: family, status: '', chosen: null, candidates: [], confidence: '', review: false, reason: '', basis: '', notes: [] };
    if (opt.pick) {
      var p = idx.filter(function (s) { return s.name === opt.pick; })[0];
      if (p) {
        var exactPick = family && p.family === family;
        res.status = exactPick ? 'exact' : 'manual'; res.chosen = p.name;
        res.confidence = exactPick ? '높음' : '담당자 선택';
        res.reason = exactPick ? '재질 계열과 같은 Index(담당자 확인)' : '담당자가 직접 고른 Index' + (family ? ' — 재질 계열(' + FAM[family].label + ')과 다르면 대체 지표입니다' : '');
        res.review = !exactPick;
      }
    }
    if (!family) {
      if (!res.chosen) { res.status = 'unknown'; res.confidence = '분석 불가'; res.review = true; res.reason = '재질 계열을 알 수 없습니다 — 재질 계열을 직접 고르세요'; }
      return res;
    }
    // 1) 규격 재질 일치(SCM440 = SCM440) → 2) 계열 일치(SMxxC ∋ SM45C) → 3) 재질 계열 일치(합금강 = 합금강)
    var g = opt.grade || '';
    var gx = g ? idx.filter(function (s) { return (s.grades || []).indexOf(g) >= 0; }) : [];
    var gw = g && !gx.length ? idx.filter(function (s) { return (s.grades || []).some(function (p) { return /X/.test(p) && wildRe(p).test(g); }); }) : [];
    var same = idx.filter(function (s) { return s.family === family; });
    var tier = gx.length ? ['grade', gx] : gw.length ? ['series', gw] : ['exact', same];
    if (opt.refHint !== false) res.refHint = series.filter(function (s) { return s.role === 'ref' && g && (s.grades || []).indexOf(g) >= 0; }).map(function (s) { return s.name; });
    if (!res.chosen && tier[1].length) {
      var list = tier[1].slice().sort(function (a, b) { return prefRank(a) - prefRank(b); });
      var why = { grade: '규격 칸의 재질(' + g + ')이 원소재와 같은 Index', series: '규격 칸의 계열 표기(' + (list[0].grades || []).filter(function (p) { return /X/.test(p); }).join(', ') + ')에 원소재 ' + g + ' 가 들어가는 Index', exact: '재질 계열(' + FAM[family].label + ')이 같은 Index' }[tier[0]];
      if (list.length > 1 && prefRank(list[0]) === prefRank(list[1])) {
        res.status = 'multi'; res.confidence = '확인 필요'; res.review = true;
        res.reason = why + ' 가 ' + list.length + '개(' + list.map(function (s) { return s.name; }).join(' · ') + ') — 지역·단위로도 가려지지 않아 임의로 고르지 않습니다. 담당자가 고르세요';
        res.candidates = list.map(function (s) { return { name: s.name, family: s.family, matSim: 1, why: why, meta: s.meta }; });
        return res;
      }
      res.status = tier[0]; res.chosen = list[0].name; res.confidence = '높음';
      res.reason = why + ' 를 그대로 적용';
      if (list.length > 1) {
        res.notes.push('같은 조건의 Index 가 ' + list.length + '개라 지역(한국 우선) · 단위(KRW/kg 우선)로 골랐습니다 — 나머지: ' + list.slice(1).map(function (s) { return s.name; }).join(' · '));
        res.alts = list.slice(1).map(function (s) { return s.name; });
      }
      return res;
    }
    if (tier[1].length && res.chosen) return res;

    // 정확 일치 없음 — 같은 계열 후보만(이름만 비슷한 것은 넣지 않는다)
    var rel = (RELATED[family] || []);
    var cands = [];
    rel.forEach(function (r) {
      if (r[1] < DEF.minMatSim) return;   // 계열이 먼 것(알루미늄 ↔ 구리 등)은 후보로도 두지 않는다
      idx.filter(function (s) { return s.family === r[0]; }).forEach(function (s) { cands.push({ s: s, matSim: r[1], why: r[2] }); });
    });
    if (!cands.length) {
      if (!res.chosen) { res.status = 'none'; res.confidence = '분석 불가'; res.review = true; res.reason = '엑셀에 ' + FAM[family].label + ' 와 같은 계열 Index 가 없습니다 — 임의로 고르지 않습니다'; }
      return res;
    }
    var start = period.start, end = period.end;
    var ch = cands.map(function (c) { return monthlyChanges(c.s, start, end); });
    var ref = opt.refName ? series.filter(function (s) { return s.name === opt.refName; })[0] : null;
    var refCh = ref ? monthlyChanges(ref, start, end) : null;
    res.basis = ref
      ? '참고 계열 「' + ref.name + '」(계산에는 쓰지 않음)과 각 후보의 월간 변동률을 비교했습니다.'
      : FAM[family].label + ' 자체의 가격 계열이 엑셀에 없어 직접 비교할 수 없습니다. 같은 계열 후보끼리 월간 변동률을 비교해 나머지 후보 평균과 얼마나 같이 움직이는지(상관)를 봤습니다. 실제 구매가 이력이 있으면 「참고 계열」로 지정해 그것과 비교하세요.';
    cands.forEach(function (c, i) {
      var other = refCh || avgChanges(ch.filter(function (_, j) { return j !== i; }));
      var pat = cands.length > 1 || refCh ? comparePattern(ch[i], other, opt.minN) : { n: 0, corr: null, dirAgree: null };
      var st = periodStats(c.s, start, end);
      c.corr = pat.corr; c.dirAgree = pat.dirAgree; c.n = pat.n; c.rate = st.ok ? st.rate : null;
      c.score = 0.4 * c.matSim + 0.6 * Math.max(0, c.corr == null ? 0 : c.corr);
    });
    cands.sort(function (a, b) { return b.score - a.score || b.matSim - a.matSim || (a.s.name < b.s.name ? -1 : 1); });
    res.candidates = cands.slice(0, 3).map(function (c) {
      return { name: c.s.name, family: c.s.family, matSim: c.matSim, why: c.why, corr: c.corr, dirAgree: c.dirAgree, n: c.n, rate: c.rate, score: c.score };
    });
    var rates = res.candidates.map(function (c) { return c.rate; }).filter(function (v) { return v != null; });
    var spread = rates.length > 1 ? Math.max.apply(null, rates) - Math.min.apply(null, rates) : 0;
    if (spread > (opt.spreadPct || DEF.spreadPct)) res.notes.push('후보 Index 간 기간 변동률 차이가 ' + spread.toFixed(1) + '%p — 결과가 어떤 Index 를 고르느냐에 크게 달라집니다');
    res.spread = spread;
    if (res.chosen) return res;   // 담당자 선택이 있으면 후보 표만 덧붙인다
    var b = res.candidates[0];
    res.status = 'estimate'; res.chosen = b.name; res.review = true;
    res.confidence = b.corr == null ? '낮음' : (b.score >= 0.75 && b.corr >= 0.8 ? '중간' : (b.score >= 0.55 ? '낮음' : '매우 낮음'));
    if (b.corr == null) res.notes.push('비교할 달이 부족해 가격 움직임을 비교하지 못했습니다 — 재질 유사도만으로 정한 순위입니다');
    if (res.confidence === '매우 낮음') res.notes.push('유사성이 낮아 결과를 확정하지 않습니다');
    res.reason = '정확히 일치하는 Index 가 없어 같은 계열 후보 중 점수(재질 유사도 40% + 가격 움직임 상관 60%)가 가장 높은 Index 를 대체 지표로 적용 — 추정, 담당자 확인 필요';
    return res;
  }

  /* ───────────── 품목 계산 ───────────── */

  function nz(v) { return v != null && v !== '' && isFinite(+v) ? +v : null; }

  /** 소재 한 줄: 완성 중량 · 투입 중량 · 스크랩 · 유효 중량 · 소재 원가(중량 × 단가) */
  function lineMaterial(line) {
    var src = line.src || {}, notes = [];
    var mat = normMaterial(line.material);
    var family = line.family || mat.family;
    var grade = mat.by === '재질 기호' ? canonGrade(mat.key, mat.family) : '';
    var famSrc = line.family ? 'input' : (mat.family ? 'calc' : '');
    var density = family ? FAM[family].density : null;
    var fd = parseDims(line.finishedDims), sd = parseDims(line.stockDims);
    if (line.finishedDims && !fd) notes.push('완성 치수를 읽지 못했습니다(' + line.finishedDims + ')');
    if (line.stockDims && !sd) notes.push('소재 규격을 읽지 못했습니다(' + line.stockDims + ')');
    var fin = { kg: null, src: '' };
    if (nz(line.finishedKg) != null) fin = { kg: +line.finishedKg, src: src.finishedKg || 'input' };
    else if (fd && density) fin = { kg: weightKg(fd, density), src: 'calc', how: dimsText(fd) + ' × 비중 ' + density };
    var cut = nz(line.cutMm) != null ? { v: +line.cutMm, src: 'input' } : { v: DEF.cutMm, src: 'assume' };
    var yl = nz(line.yield) != null ? { v: +line.yield, src: 'input' } : { v: DEF.yield, src: 'assume' };
    var rc = nz(line.recovery) != null ? { v: +line.recovery, src: 'input' } : { v: DEF.recovery, src: 'assume' };
    var inp = { kg: null, src: '' };
    if (sd && density) { inp = { kg: weightKg(sd, density, cut.v), src: 'calc', how: dimsText(sd) + ' + 절단 여유 ' + cut.v + 'mm' }; yl = fin.kg != null ? { v: fin.kg / inp.kg, src: 'calc' } : yl; }
    else if (fin.kg != null && yl.v > 0) inp = { kg: fin.kg / yl.v, src: 'calc', how: '완성 중량 ÷ 수율 ' + yl.v };
    if (!family) notes.push('재질 계열을 모릅니다');
    if (fin.kg == null) notes.push('완성 중량을 정할 수 없습니다(중량 또는 치수 필요)');
    if (inp.kg != null && fin.kg != null && inp.kg < fin.kg) notes.push('소재 투입 중량이 완성 중량보다 작습니다 — 소재 규격 확인');
    var scrap = inp.kg != null && fin.kg != null ? Math.max(0, inp.kg - fin.kg) : null;
    var eff = inp.kg != null ? inp.kg - (scrap || 0) * rc.v : null;
    var up = nz(line.unitPrice);
    return {
      material: line.material, mat: mat, grade: grade, family: family, famLabel: family ? FAM[family].label : '', famSrc: famSrc, density: density,
      finished: fin, finishedDims: fd, stockDims: sd, cut: cut, yield: yl, recovery: rc, input: inp, scrapKg: scrap, effKg: eff,
      unitPrice: up, unitPriceSrc: up != null ? (src.unitPrice || 'input') : '',
      cost: eff != null && up != null ? eff * up : null, notes: notes
    };
  }

  /** 원가 구성 → 금액(원). ratio 면 기준 단가 × %, amount 면 그대로 */
  function compAmounts(item) {
    var p0 = nz(item.price0), c = item.comp || {}, out = {};
    ['material', 'process', 'labor', 'overhead'].forEach(function (k) {
      var v = nz(c[k]);
      out[k] = v == null ? null : (item.compMode === 'amount' ? v : (p0 == null ? null : p0 * v / 100));
    });
    return out;
  }

  /**
   * 품목 하나 분석. ctx = { series, period, tolPct }
   * 소재 원가 기준: 모든 줄에 소재 단가가 있으면 중량 × 단가(계산), 없으면 원가 구성의 소재비(직접 입력).
   * 영향 = 소재 원가 × Index 변동률. 가공비·인건비·경비는 Index 영향을 받지 않는 별도 항목으로 둔다.
   */
  function analyzeItem(item, ctx) {
    var period = ctx.period, notes = [], review = [];
    var p0 = nz(item.price0), pn = nz(item.priceNew);
    var comp = compAmounts(item);
    var lines = (item.lines || []).map(function (ln) {
      var L = lineMaterial(ln);
      L.match = matchIndex(L.family, ctx.series, period, { pick: ln.indexPick, refName: ln.refName, grade: L.grade });
      var s = L.match.chosen && L.match.status !== 'multi' ? ctx.series.filter(function (x) { return x.name === L.match.chosen; })[0] : null;
      L.stats = s ? periodStats(s, period.start, period.end) : null;
      return L;
    });
    if (p0 == null) review.push('기준 단가가 없습니다');
    var allCost = lines.length && lines.every(function (L) { return L.cost != null; });
    var M0 = null, mSrc = '';
    if (allCost) { M0 = lines.reduce(function (s, L) { return s + L.cost; }, 0); mSrc = 'calc'; lines.forEach(function (L) { L.base = L.cost; L.baseSrc = 'calc'; }); }
    else if (comp.material != null) {
      M0 = comp.material; mSrc = 'input';
      var known = lines.reduce(function (s, L) { return s + (L.cost || 0); }, 0);
      var rest = lines.filter(function (L) { return L.cost == null; });
      var restAmt = M0 - known;
      var wsum = rest.reduce(function (s, L) { return s + (L.effKg || 0); }, 0);
      lines.forEach(function (L) {
        if (L.cost != null) { L.base = L.cost; L.baseSrc = 'calc'; return; }
        L.base = rest.length === 1 ? restAmt : (wsum > 0 ? restAmt * (L.effKg || 0) / wsum : restAmt / rest.length);
        L.baseSrc = rest.length === 1 && !known ? 'input' : 'assume';
      });
      if (rest.length > 1) notes.push('소재 단가가 없는 소재 ' + rest.length + '줄은 원가 구성의 소재비를 유효 중량 비율로 나눴습니다(가정)');
    } else review.push('소재 원가를 정할 수 없습니다 — 소재 단가(원/kg) 또는 원가 구성의 소재비를 넣으세요');
    if (allCost && comp.material != null && comp.material > 0 && Math.abs(M0 / comp.material - 1) > 0.1) {
      review.push('중량 × 단가로 계산한 소재 원가(' + Math.round(M0).toLocaleString('ko-KR') + '원)와 원가 구성의 소재비(' + Math.round(comp.material).toLocaleString('ko-KR') + '원)가 10% 넘게 다릅니다');
    }
    var dM = 0, incomplete = false;
    lines.forEach(function (L, i) {
      var tag = (item.lines.length > 1 ? '소재 ' + (i + 1) + ' ' : '') + (L.material || '재질 미입력');
      L.notes.forEach(function (n) { review.push(tag + ': ' + n); });
      if (L.match.review) review.push(tag + ': ' + matchLabel(L.match) + ' — ' + L.match.reason);
      L.match.notes.forEach(function (n) { review.push(tag + ': ' + n); });
      if (L.stats && L.stats.ok) L.stats.notes.forEach(function (n) { review.push(tag + ' · ' + L.match.chosen + ': ' + n); });
      if (L.base != null && L.stats && L.stats.ok) { L.rate = L.stats.rate; L.dM = L.base * L.stats.rate / 100; dM += L.dM; }
      else { L.dM = null; incomplete = true; }
    });
    var res = {
      id: item.id, name: item.name, dwg: item.dwg, qty: item.qty, file: item.file || '', src: item.src || {},
      price0: p0, priceNew: pn, comp: comp, lines: lines, M0: M0, M0src: mSrc,
      share: M0 != null && p0 ? M0 / p0 * 100 : null,
      dM: incomplete ? null : dM, notes: notes, review: review
    };
    res.rateOnPrice = res.dM != null && p0 ? res.dM / p0 * 100 : null;
    res.matRate = res.dM != null && M0 ? res.dM / M0 * 100 : null;
    res.expected = res.dM != null && p0 != null ? p0 + res.dM : null;
    // 소재 외 항목 — Index 영향 없음
    if (p0 != null && M0 != null) {
      var other = p0 - M0, named = 0;
      ['process', 'labor', 'overhead'].forEach(function (k) { if (comp[k] != null) named += comp[k]; });
      res.others = { process: comp.process, labor: comp.labor, overhead: comp.overhead, rest: other - named, total: other };
      if (other - named < -0.5) review.push('소재 원가 + 가공비 + 인건비 + 경비가 기준 단가보다 큽니다(' + Math.round(other - named).toLocaleString('ko-KR') + '원) — 원가 구성 확인');
    }
    // 검증 — 요청(현재) 단가와 비교
    if (pn != null && p0 != null) {
      var actual = pn - p0, tol = p0 * (ctx.tolPct != null ? ctx.tolPct : DEF.tolPct) / 100;
      var v = { actual: actual, actualRate: actual / p0 * 100, tol: tol };
      if (res.dM == null) { v.verdict = 'na'; v.text = '소재 영향을 계산하지 못해 비교할 수 없습니다'; }
      else {
        v.unexplained = actual - res.dM;
        if (res.dM < 0 && actual > 0) { v.verdict = 'contradict'; v.text = '원소재 가격은 내렸는데 단가는 올랐습니다 — 원소재 외 인상 사유(가공비·인건비 등) 근거 확인 필요'; }
        else if (Math.abs(v.unexplained) <= tol + 1e-6) { v.verdict = 'explained'; v.text = '단가 변동이 원소재 가격 변동으로 설명됩니다(차이 ' + fmtPct(v.unexplained / p0 * 100) + ' 이내 허용)'; }
        else { v.verdict = 'review'; v.text = '원소재 변동으로 설명되지 않는 차이 ' + Math.round(v.unexplained).toLocaleString('ko-KR') + '원(' + fmtPct(v.unexplained / p0 * 100) + ') — 추가 검토 필요'; }
      }
      res.verify = v;
      if (v.verdict === 'contradict' || v.verdict === 'review') review.push(v.text);
    }
    var conf = lines.map(function (L) { return L.match.confidence; });
    res.confidence = conf.indexOf('분석 불가') >= 0 ? '분석 불가' : (conf.indexOf('확인 필요') >= 0 ? '확인 필요' : (conf.indexOf('매우 낮음') >= 0 ? '매우 낮음' : (conf.indexOf('낮음') >= 0 ? '낮음' : (conf.indexOf('중간') >= 0 ? '중간' : (conf.indexOf('담당자 선택') >= 0 ? '담당자 선택' : (conf.length ? '높음' : ''))))));
    if (lines.some(function (L) { return L.yield.src === 'assume' && !L.stockDims || L.recovery.src === 'assume'; })) notes.push('수율·스크랩 회수율 중 가정 기본값을 쓴 칸이 있습니다');
    return res;
  }
  function matchLabel(m) {
    return { grade: '규격 재질 일치', series: '계열 일치', exact: '재질 계열 일치', multi: '여러 Index 일치', estimate: '추정', none: '분석 불가', unknown: '재질 불명', manual: '담당자 선택' }[m.status] || m.status;
  }
  function fmtPct(v) { return v == null ? '' : (v > 0 ? '+' : '') + v.toFixed(2) + '%'; }

  function checkPeriod(p) {
    var errs = [];
    if (!p || !/^\d{4}-\d{2}$/.test(p.start || '')) errs.push('분석 시작월을 넣으세요(YYYY-MM)');
    if (!p || !/^\d{4}-\d{2}$/.test(p.end || '')) errs.push('분석 종료월을 넣으세요(YYYY-MM)');
    if (!errs.length && mIdx(p.start) >= mIdx(p.end)) errs.push('시작월이 종료월보다 앞이어야 합니다');
    return errs;
  }

  /** 여러 품목 일괄 분석 + 종합 */
  function analyze(input) {
    var period = input.period, series = (input.index && input.index.series) || [];
    var errs = checkPeriod(period);
    if (!series.filter(function (s) { return s.role === 'index'; }).length) errs.push('Index 엑셀을 올리세요(Index 로 쓸 열이 하나 이상)');
    if (errs.length) return { ok: false, errors: errs };
    var ctx = { series: series, period: period, tolPct: input.tolPct };
    var items = (input.items || []).map(function (it) { return analyzeItem(it, ctx); });
    var used = {};
    items.forEach(function (r) { r.lines.forEach(function (L) { if (L.stats && L.stats.ok) used[L.match.chosen] = L.stats; }); });
    var indexes = Object.keys(used).map(function (n) {
      var s = series.filter(function (x) { return x.name === n; })[0];
      return { name: n, family: s.family, source: s.source, meta: s.meta || {}, stats: used[n] };
    });
    var review = [];
    items.forEach(function (r) { r.review.forEach(function (t) { review.push({ item: r.name || r.id, text: t }); }); });
    var done = items.filter(function (r) { return r.dM != null; });
    return {
      ok: true, period: period, items: items, indexes: indexes, review: review, allSeries: series,
      summary: {
        count: items.length, done: done.length,
        estimated: items.filter(function (r) { return r.lines.some(function (L) { return L.match.status === 'estimate' || L.match.status === 'manual'; }); }).length,
        base: done.reduce(function (s, r) { return s + r.price0; }, 0),
        dM: done.reduce(function (s, r) { return s + r.dM; }, 0)
      },
      issues: (input.index && input.index.issues) || [], source: (input.index && input.index.source) || ''
    };
  }

  /* ───────────── 엑셀 시트(행 배열) ───────────── */

  var SRC_LABEL = { doc: '도면에서 읽음', input: '직접 입력', calc: '계산', est: '추정', assume: '가정' };
  function r2(v, d) { if (v == null || !isFinite(v)) return ''; var k = Math.pow(10, d == null ? 2 : d); return Math.round(v * k) / k; }

  function exportSheets(res) {
    var P = res.period;
    var draw = [['제품명', '도면번호', '도면 파일', '원소재(재질)', '재질 계열', '완성 치수', '소재 규격', '완성 중량(kg)', '완성 중량 출처', '투입 중량(kg)', '수율', '수율 출처', '스크랩 회수율', '유효 중량(kg)', '소재 단가(원/kg)', '적용 Index', '매칭']];
    var match = [['제품명', '원소재(재질)', '재질 계열', '매칭', '적용 Index', '적용 Index 단위', '적용 Index 출처', '상세출처 · 주소', '선정 이유', '신뢰도', '비교 기준', '후보1', '후보1 재질유사도', '후보1 상관', '후보1 변동률(%)', '후보2', '후보2 재질유사도', '후보2 상관', '후보2 변동률(%)', '후보3', '후보3 재질유사도', '후보3 상관', '후보3 변동률(%)', '주의']];
    var impact = [['제품명', '도면번호', '기준 단가(원)', '제품당 소재 원가(원)', '소재 원가 출처', '소재 비중(%)', 'Index 변동률(%)', '소재 원가 증감액(원)', '단가 대비 영향률(%)', '예상 단가(원)', '가공비(원)', '인건비(원)', '경비(원)', '그 밖(이윤 등, 원)', '요청·현재 단가(원)', '실제 변동(원)', '설명 안 되는 차이(원)', '판정', '신뢰도']];
    res.items.forEach(function (r) {
      r.lines.forEach(function (L) {
        draw.push([r.name, r.dwg, r.file, L.material, L.famLabel, L.finishedDims ? dimsText(L.finishedDims) : '', L.stockDims ? dimsText(L.stockDims) : '', r2(L.finished.kg, 3), SRC_LABEL[L.finished.src] || '', r2(L.input.kg, 3), r2(L.yield.v, 3), SRC_LABEL[L.yield.src] || '', r2(L.recovery.v, 2), r2(L.effKg, 3), L.unitPrice == null ? '' : L.unitPrice, L.match.chosen || '', matchLabel(L.match)]);
        var cs = L.match.status === 'multi' ? null : (res.allSeries || []).filter(function (x) { return x.name === L.match.chosen; })[0], cm = cs && cs.meta || {};
        var row = [r.name, L.material, L.famLabel, matchLabel(L.match), L.match.status === 'multi' ? '' : (L.match.chosen || ''), cm.unit || '', cs ? [cm.source, cm.basis].filter(Boolean).join(' · ') + ' (' + cs.source + ')' : '', [cm.detail, cm.url].filter(Boolean).join(' · '), L.match.reason, L.match.confidence, L.match.basis || ''];
        for (var k = 0; k < 3; k++) { var c = L.match.candidates[k]; row.push(c ? c.name : '', c ? c.matSim : '', c && c.corr != null ? r2(c.corr, 3) : '', c && c.rate != null ? r2(c.rate) : ''); }
        row.push(L.match.status === 'estimate' || L.match.status === 'manual' ? '유사 Index 는 실제 원소재 가격이 아니라 가격 변동을 추정하는 대체 지표입니다. ' + L.match.notes.join(' / ') : L.match.notes.join(' / '));
        match.push(row);
      });
      var o = r.others || {};
      impact.push([r.name, r.dwg, r.price0, r2(r.M0, 0), SRC_LABEL[r.M0src] || '', r2(r.share), r2(r.matRate), r2(r.dM, 0), r2(r.rateOnPrice), r2(r.expected, 0), r2(o.process, 0), r2(o.labor, 0), r2(o.overhead, 0), r2(o.rest, 0), r.priceNew == null ? '' : r.priceNew, r.verify ? r2(r.verify.actual, 0) : '', r.verify && r.verify.unexplained != null ? r2(r.verify.unexplained, 0) : '', r.verify ? r.verify.text : '', r.confidence]);
    });
    // 단위가 다른 Index 를 한 표에 섞지 않는다 — 단위별로 묶는다
    var units = [];
    res.indexes.forEach(function (x) { var u = (x.meta && x.meta.unit) || '(단위 없음)'; if (units.indexOf(u) < 0) units.push(u); });
    var price = [['단위', 'Index', '기준', '지역', '출처', '상세출처', '파일 위치', '시작(요청)', '시작(사용 월)', '시작 값', '종료(요청)', '종료(사용 월)', '종료 값', '변동액', '변동률(%)', '기간 최고(월)', '최고 값', '기간 최저(월)', '최저 값', '오른 달', '내린 달', '비고']];
    units.forEach(function (u) {
      res.indexes.filter(function (x) { return ((x.meta && x.meta.unit) || '(단위 없음)') === u; }).forEach(function (x) {
        var s = x.stats, m = x.meta || {};
        price.push([u, x.name, m.basis || '', m.region || '', m.source || '', m.detail || '', x.source, P.start, s.start.m, s.start.v, P.end, s.end.m, s.end.v, r2(s.diff), r2(s.rate), s.max.m, s.max.v, s.min.m, s.min.v, s.up, s.down, s.notes.join(' / ') + (m.fx ? ' / 원화 환산(' + m.origUnit + ' × ' + m.fx + ') — 계산 값' : '')]);
      });
    });
    units.forEach(function (u) {
      var xs = res.indexes.filter(function (x) { return ((x.meta && x.meta.unit) || '(단위 없음)') === u; });
      price.push([]); price.push(['월별 값(분석 기간 · 사용한 Index 원자료) — 단위 ' + u]);
      price.push(['월'].concat(xs.map(function (x) { return x.name; })));
      for (var k = mIdx(P.start); k <= mIdx(P.end); k++) {
        var mm = mStr(k);
        price.push([mm].concat(xs.map(function (x) { var p = x.stats.months.filter(function (q) { return q.m === mm; })[0]; return p ? p.v : ''; })));
      }
    });
    var S = res.summary;
    var total = [['항목', '값'], ['분석 기간', P.start + ' ~ ' + P.end], ['품목 수', S.count], ['영향 계산 완료', S.done], ['유사 Index(추정·담당자 선택) 적용 품목', S.estimated], ['계산 완료 품목 기준 단가 합(원)', r2(S.base, 0)], ['원소재 영향 합(원)', r2(S.dM, 0)], []];
    total.push(['Index', '기간 변동률(%)']);
    res.indexes.forEach(function (x) { total.push([x.name, r2(x.stats.rate)]); });
    total.push([]); total.push(['제품', '소재 원가 증감액(원)', '단가 대비 영향률(%)', '신뢰도', '유사 Index 적용']);
    res.items.forEach(function (r) { total.push([r.name, r2(r.dM, 0), r2(r.rateOnPrice), r.confidence, r.lines.some(function (L) { return L.match.status === 'estimate' || L.match.status === 'manual'; }) ? '예 — 추정' : '아니오']); });
    total.push([]); total.push(['추가 검토가 필요한 항목']);
    res.review.forEach(function (v) { total.push([v.item, v.text]); });
    var basis = [['근거 · 수식 · 가정'],
      ['Index 출처', res.source || ''], ['분석에 쓴 Index', '올린 엑셀에 있는 열만 씁니다. 외부 시세·인터넷 검색 값은 쓰지 않습니다.'],
      ['기간 값', '시작·종료월 값이 없으면 가장 가까운 이전 관측월 값을 씁니다(사용 월을 표시). 이전 값이 없으면 다음 관측월 값을 쓰고 확인 필요로 표시합니다.'],
      ['Index 변동률', '(종료 값 ÷ 시작 값 − 1) × 100'],
      ['투입 중량', '소재 규격이 있으면 소재 규격 치수(길이 + 절단 여유) × 비중, 없으면 완성 중량 ÷ 수율'],
      ['유효 중량', '투입 중량 − 스크랩(투입 − 완성) × 스크랩 회수율'],
      ['소재 원가', '유효 중량 × 소재 단가(원/kg). 소재 단가가 없으면 원가 구성의 소재비'],
      ['원가 증감액', '소재 원가 × Index 변동률. 가공비·인건비·경비는 Index 영향을 받지 않는 별도 항목'],
      ['단가 대비 영향률', '원가 증감액 ÷ 기준 단가 × 100'],
      ['검증', '요청·현재 단가 − 기준 단가 와 원가 증감액의 차이가 기준 단가의 허용 % 를 넘으면 추가 검토'],
      ['유사 Index 점수', '재질 유사도 × 0.4 + 월간 변동률 상관 × 0.6 (상관은 음수면 0)'],
      ['가정 기본값', '수율 ' + DEF.yield + ' · 스크랩 회수율 ' + DEF.recovery + ' · 절단 여유 ' + DEF.cutMm + 'mm — 입력하지 않은 칸에만 쓰고 「가정」으로 표시'],
      ['비중(g/cm³)', FAM_KEYS.map(function (k) { return FAM[k].label + ' ' + FAM[k].density; }).join(' · ')],
      ['주의', '유사 Index 는 실제 원소재 가격이 아니라 가격 변동을 추정하는 대체 지표입니다.']];
    (res.issues || []).forEach(function (i) { basis.push(['엑셀 읽기', i.msg]); });
    return { '도면별분석': draw, 'Index매칭': match, '가격변동': price, '단가영향': impact, '종합': total, '근거가정': basis };
  }

  /* ───────────── AI(선택) — 도면은 보내지 않는다 ───────────── */

  function aiPrompt(material, indexNames) {
    return '구매 원가 분석을 돕고 있습니다. 아래 원소재와 가장 가까운 원자재 Index 를 목록 안에서만 골라 이유를 3문장 이내로 적어 주세요. 목록에 없는 Index 를 만들지 말고, 적절한 것이 없으면 「없음」이라고 답하세요. 가격 수치는 추측하지 마세요.\n\n원소재(담당자 확인): ' + material + '\nIndex 목록: ' + indexNames.join(', ');
  }

  /* ───────────── 예제(가짜 데이터) ───────────── */

  function rng(seed) { return function () { seed |= 0; seed = seed + 0x6D2B79F5 | 0; var t = Math.imul(seed ^ seed >>> 15, 1 | seed); t = t + Math.imul(t ^ t >>> 7, 61 | t) ^ t; return ((t ^ t >>> 14) >>> 0) / 4294967296; }; }
  function curve(anchors, n) {
    var out = [];
    for (var i = 0; i < n; i++) {
      var a = anchors[0], b = anchors[anchors.length - 1];
      for (var k = 0; k + 1 < anchors.length; k++) if (i >= anchors[k][0] && i <= anchors[k + 1][0]) { a = anchors[k]; b = anchors[k + 1]; break; }
      out.push(b[0] === a[0] ? a[1] : a[1] + (b[1] - a[1]) * (i - a[0]) / (b[0] - a[0]));
    }
    return out;
  }
  var SAMPLE_COLS = [
    ['열연강판(원/톤)', [[0, 1080000], [3, 1250000], [5, 1180000], [11, 960000], [14, 1050000], [23, 920000], [29, 880000], [35, 840000], [40, 870000], [47, 900000]], 1000, 1],
    ['냉연강판(원/톤)', [[0, 1200000], [4, 1400000], [6, 1330000], [12, 1090000], [15, 1170000], [23, 1040000], [29, 1000000], [35, 960000], [40, 985000], [47, 1010000]], 1000, 1],
    ['특수강(합금강)(원/톤)', [[0, 1350000], [8, 1560000], [14, 1520000], [23, 1420000], [35, 1330000], [40, 1340000], [47, 1350000]], 1000, 1],
    ['스테인리스 STS304(원/톤)', [[0, 3600000], [2, 4300000], [5, 3900000], [14, 3550000], [23, 3100000], [35, 2950000], [40, 2900000], [47, 2950000]], 1000, 0],
    ['알루미늄 LME(USD/톤)', [[0, 3000], [2, 3500], [8, 2250], [17, 2200], [28, 2600], [35, 2500], [40, 2450], [47, 2650]], 1, 0],
    ['구리 LME(USD/톤)', [[0, 9800], [2, 10200], [6, 7600], [12, 9000], [21, 8000], [28, 10100], [35, 9000], [40, 9500], [47, 10000]], 1, 0]
  ];
  /** 가짜 Index 시트(2022-01~2025-12 월별 6열). 특수강 2025-05 · 알루미늄 2023-07 은 일부러 비워 둔다. */
  function sampleIndexAoa() {
    var R = rng(2611), n = 48, common = [];
    for (var i = 0; i < n; i++) common.push(R() - 0.5);
    var cols = SAMPLE_COLS.map(function (c) {
      var base = curve(c[1], n);
      return base.map(function (v, i) {
        var nz = c[3] ? 0.6 * common[i] + 0.4 * (R() - 0.5) : (R() - 0.5);
        return Math.round(v * (1 + 0.016 * nz) / c[2]) * c[2];
      });
    });
    var aoa = [['원자재 Index — 가짜 예제 데이터(실제 시세 아님)'], [], ['기준월'].concat(SAMPLE_COLS.map(function (c) { return c[0]; }))];
    for (var k = 0; k < n; k++) {
      var m = mStr(2022 * 12 + k);
      aoa.push([m].concat(cols.map(function (c, j) {
        if (j === 2 && m === '2025-05') return '';
        if (j === 4 && m === '2023-07') return '';
        return c[k];
      })));
    }
    return aoa;
  }
  /**
   * 가짜 통합문서 — 실무 「원자재 가격 동향」 형식을 흉내 낸다(값 · 출처 · 주소는 모두 지어낸 것).
   * 시트 5장: 철판 · 특수강,환봉 · 비철금속 · 환율 · 제조업 PMI. 머리 블록 · 가로 병합 품목명 · 증감률 열 · 분기 열 · 최신이 위 · 시트마다 다른 최신 월.
   * 특수강 SCRxxx 2025-05 는 일부러 비워 둔다. merges 는 엑셀로 쓸 때 병합 범위.
   */
  function sampleBook() {
    var K0 = 2020 * 12, R = rng(2611);
    var common = []; for (var i = 0; i < 82; i++) common.push(R() - 0.5);
    function ser(anchors, steel, round) { return curve(anchors, 82).map(function (v, i) { var nz = steel ? 0.6 * common[i] + 0.4 * (R() - 0.5) : R() - 0.5; return +(v * (1 + 0.016 * nz)).toFixed(round == null ? 0 : round); }); }
    var fxUsd = ser([[0, 1160], [14, 1120], [30, 1300], [40, 1380], [55, 1320], [64, 1470], [72, 1420], [81, 1390]], false, 2);
    var fxCny = fxUsd.map(function (v, i) { return +(v / (7.0 + 0.2 * Math.sin(i / 9))).toFixed(2); });
    var URL = 'https://example.invalid/sample';
    var merges = {}, sheets = {}, names = [];
    // spec: [item, region, basis, unit, spec, source, detail, values]  — values null 이면 증감률 열
    function build(name, title, labelOrder, cols, opt) {
      var newest = K0 + opt.last, rows = [], m = [];
      var head = [[title]];
      labelOrder.forEach(function (lb) { head.push([lb === 'Item' ? 'Item' : lb]); });
      var key = { 'Item': 0, '지역': 1, '기준': 2, '단위': 3, '규격': 4, '출처': 5, '상세출처': 6, '주소': 7 };
      var lead = opt.ym ? (opt.quarter ? 3 : 2) : 1;
      cols.forEach(function (c, j) {
        var col = lead + j;
        labelOrder.forEach(function (lb, li) { var v = c.head ? c.head[key[lb]] : ''; head[li + 1][col] = v == null ? '' : v; });
      });
      // 품목명 가로 병합 범위
      var itemRow = labelOrder.indexOf('Item') + 1, start = null;
      cols.forEach(function (c, j) {
        if (c.head && c.head[0]) { if (start != null && lead + j - 1 > start) m.push({ s: { r: itemRow, c: start }, e: { r: itemRow, c: lead + j - 1 } }); start = lead + j; }
      });
      if (start != null && lead + cols.length - 1 > start) m.push({ s: { r: itemRow, c: start }, e: { r: itemRow, c: lead + cols.length - 1 } });
      head.forEach(function (h) { for (var x = 0; x < lead + cols.length; x++) if (h[x] == null) h[x] = ''; });
      var yearTop = null;
      for (var k = newest; k >= K0; k--) {
        var i = k - K0, y = Math.floor(k / 12), mo = k % 12 + 1, row = [];
        if (opt.ym) {
          var showY = !opt.mergeYear || y !== yearTop;
          if (opt.mergeYear && y !== yearTop) { if (yearTop != null) m.push({ s: { r: head.length + rows.length - (rows._run || 0), c: 0 }, e: { r: head.length + rows.length - 1, c: 0 } }); rows._run = 0; yearTop = y; }
          rows._run = (rows._run || 0) + 1;
          row.push(showY ? y + '년' : '', mo + '월');
          if (opt.quarter) row.push(String(y).slice(2) + '.' + Math.ceil(mo / 3) + 'Q');
        } else row.push(y + '-' + mo);
        var lastVal = null;
        cols.forEach(function (c) {
          if (c.vals) { lastVal = c; var v = c.vals[i]; row.push(c.blank && c.blank === mStr(k) ? '' : v); }
          else { var prev = lastVal.vals[i - 1]; row.push(i > 0 && !(lastVal.blank === mStr(k) || lastVal.blank === mStr(k - 1)) ? +(lastVal.vals[i] / prev - 1).toFixed(4) : ''); }
        });
        rows.push(row);
      }
      if (opt.mergeYear && rows._run > 1) m.push({ s: { r: head.length + rows.length - rows._run, c: 0 }, e: { r: head.length + rows.length - 1, c: 0 } });
      delete rows._run;
      sheets[name] = head.concat(rows); merges[name] = m; names.push(name);
    }
    var L8 = ['Item', '지역', '기준', '단위', '규격', '출처', '상세출처', '주소'];
    var SRC = '가상 자료';
    var cr = ser([[0, 820], [15, 900], [27, 1350], [30, 1300], [36, 1150], [48, 1080], [60, 1020], [72, 1000], [81, 1040]], true);
    var crCn = ser([[0, 4300], [15, 4700], [27, 6200], [36, 4900], [48, 4600], [60, 4200], [81, 4000]], true);
    var hr = ser([[0, 700], [15, 800], [27, 1250], [30, 1180], [36, 960], [39, 1050], [48, 920], [54, 880], [60, 840], [65, 870], [81, 900]], true);
    var hrCn = ser([[0, 3700], [15, 4300], [27, 5600], [36, 4200], [48, 3900], [60, 3600], [81, 3400]], true);
    var pl = ser([[0, 760], [15, 850], [28, 1320], [36, 1100], [48, 1000], [60, 950], [65, 960], [81, 980]], true);
    build('철판', '■ 원자재 가격 정보_철강(판재) — 가상 예제', L8, [
      { head: ['냉연강판(CR)', '한국', '유통가(가상)', 'KRW/kg', 'SPCC 1.0t', SRC, '예제용 지어낸 값', URL], vals: cr }, {},
      { head: ['', '중국', '중국 국내가(가상)', 'RMB/ton', 'DC01 1.0mm', SRC, '예제용 지어낸 값', URL], vals: crCn }, {},
      { head: ['열연강판(HR)', '한국', '유통가(가상)', 'KRW/kg', 'SPHC 2.0t', SRC, '예제용 지어낸 값', URL], vals: hr }, {},
      { head: ['', '중국', '중국 국내가(가상)', 'RMB/ton', 'Q235 2.0mm', SRC, '예제용 지어낸 값', URL], vals: hrCn }, {},
      { head: ['후판', '한국', '유통가(가상)', 'KRW/kg', 'SS400 20t', SRC, '예제용 지어낸 값', URL], vals: pl }, {}
    ], { ym: true, quarter: true, mergeYear: true, last: 81 });
    var rb = ser([[0, 780], [15, 880], [28, 1300], [36, 1150], [60, 1000], [81, 1020]], true);
    var smc = ser([[0, 1100], [24, 1250], [32, 1500], [44, 1460], [60, 1360], [72, 1330], [81, 1340]], true);
    var scr = ser([[0, 1250], [24, 1380], [33, 1620], [44, 1580], [60, 1470], [72, 1400], [81, 1410]], true);
    var scm = scr.map(function (v, i) { return Math.round(v * 1.06 + 8 * Math.sin(i / 4)); });
    var sts = ser([[0, 3400], [26, 4400], [30, 4000], [40, 3600], [52, 3200], [64, 3050], [81, 3100]], false);
    var ST = function (item, unit, spec, src) { return [item, '한국', '도매가(가상)', unit, spec, src || SRC, '예제용 지어낸 값', URL]; };
    build('특수강,환봉', '■ 원자재 가격 정보_특수강(봉강) — 가상 예제', L8, [
      { head: ST('일반 환봉', 'KRW/kg', 'SS275 마환봉 Φ31~80'), vals: rb }, {},
      { head: ST('기계구조용 탄소강 환봉', 'KRW/kg', 'SMxxC Φ260~350'), vals: smc }, {},
      { head: ST('합금강 환봉', 'KRW/kg', 'SCRxxx Φ25~160'), vals: scr, blank: '2025-05' }, {},
      { head: ST('합금강 환봉(사내)', 'KRW/kg', 'SCM440 Φ50', '구매팀 단가(가상)'), vals: scm }, {},
      { head: ST('스테인리스 환봉', 'KRW/kg', 'STS304 환봉 Φ100'), vals: sts }, {}
    ], { ym: true, quarter: false, mergeYear: false, last: 81 });
    var cu = ser([[0, 6000], [6, 5200], [16, 9000], [27, 10200], [31, 7600], [37, 9000], [46, 8000], [53, 10100], [60, 9000], [65, 9500], [80, 9900]], false);
    var al = ser([[0, 1800], [5, 1500], [22, 2900], [26, 3500], [32, 2250], [41, 2200], [52, 2600], [59, 2500], [64, 2450], [80, 2650]], false);
    var ni = ser([[0, 13500], [5, 12000], [26, 30000], [28, 24000], [40, 21000], [52, 17500], [64, 15500], [80, 15800]], false);
    var krw = function (a) { return a.map(function (v, i) { return +(v * fxUsd[i] / 1000).toFixed(1); }); };
    var NF = function (item, unit, spec) { return [item, '국제', 'LME 현물(가상)', unit, spec, SRC, '예제용 지어낸 값', URL]; };
    build('비철금속', '■ 원자재 가격 정보_비철금속 — 가상 예제', ['Item', '지역', '기준', '단위', '출처', '규격', '상세출처', '주소'], [
      { head: NF('구리', 'USD/ton', '전기동'), vals: cu }, { head: ['', '', '', 'KRW/kg(환산)', '', '', '', ''], vals: krw(cu) }, {},
      { head: NF('알루미늄', 'USD/ton', '알루미늄 지금'), vals: al }, { head: ['', '', '', 'KRW/kg(환산)', '', '', '', ''], vals: krw(al) }, {},
      { head: NF('니켈', 'USD/ton', '니켈 지금'), vals: ni }, { head: ['', '', '', 'KRW/kg(환산)', '', '', '', ''], vals: krw(ni) }, {}
    ], { ym: true, quarter: true, mergeYear: false, last: 80 });
    build('환율', '■ 환율 — 가상 예제', ['Item', '단위', '출처'], [
      { head: ['USD', '', '', 'KRW/USD', '', '가상 자료(매매기준율 월평균 흉내)'], vals: fxUsd }, {},
      { head: ['CNY', '', '', 'KRW/CNY', '', '가상 자료(매매기준율 월평균 흉내)'], vals: fxCny }, {}
    ], { ym: true, quarter: true, mergeYear: false, last: 76 });
    var pmi = function (b) { return ser([[0, b], [3, b - 8], [12, b + 3], [30, b - 2], [50, b], [76, b + 1]], false, 1); };
    // PMI: 2행 지역 · 데이터 A=「2026-5」 — Item · 단위 줄이 없다
    build('제조업 PMI', '■ 제조업 PMI — 가상 예제', ['지역', '출처'], [
      { head: ['', 'China', '', '', '', SRC], vals: pmi(50.5) }, { head: ['', 'US', '', '', '', SRC], vals: pmi(52) }, { head: ['', 'Korea', '', '', '', SRC], vals: pmi(49.5) }
    ], { ym: false, last: 76 });
    return { names: names, sheets: sheets, merges: merges };
  }
  /** 가짜 도면 글자(samples/ 의 PDF 와 같은 내용). 세 번째는 스캔 도면이라 글자가 없다. */
  var SAMPLE_DRAWINGS = [
    { file: 'BRK-2611-A_bracket.pdf', lines: [
      'DREAMIT SAMPLE DRAWING - NOT A REAL COMPANY DRAWING',
      'TITLE: BRACKET, MOUNTING',
      'DWG NO: BRK-2611-A   REV: B',
      'MATERIAL: SS400',
      'FINISHED SIZE: t12 x 150 x 300',
      'STOCK SIZE: PL t12 x 160 x 310',
      'WEIGHT: 3.6 kg',
      'QTY: 1 EA',
      'SCALE 1:2   UNIT: mm',
      'NOTE: REMOVE ALL BURRS AND SHARP EDGES'] },
    { file: 'PIN-2611-B_pin.pdf', lines: [
      'DREAMIT SAMPLE DRAWING - NOT A REAL COMPANY DRAWING',
      'TITLE: PIN, PIVOT',
      'DWG NO: PIN-2611-B   REV: A',
      'MATERIAL: SCM440H (Q/T)',
      'SIZE: Ø45 x 210 L',
      'STOCK: Ø50 x 215 L',
      'QTY: 2 EA',
      'HARDNESS: HRC 28-32'] },
    { file: 'BSH-2611-C_scan.pdf', lines: [] }
  ];
  function sample() {
    var aoa = sampleIndexAoa();
    var ex = SAMPLE_DRAWINGS.map(function (d) { return extractDrawing(d.lines.join('\n'), d.file); });
    var a = itemFromDrawing(ex[0], 'i1'), b = itemFromDrawing(ex[1], 'i2'), c = itemFromDrawing(ex[2], 'i3');
    a.price0 = 18500; a.compMode = 'ratio'; a.comp = { material: null, process: 38, labor: 22, overhead: 12 };
    a.lines[0].unitPrice = 1050;
    b.price0 = 32000; b.priceNew = 34500; b.compMode = 'amount'; b.comp = { material: null, process: 15000, labor: 6500, overhead: 3200 };
    b.lines[0].unitPrice = 1850;
    c.name = 'BUSHING, GUIDE'; c.dwg = 'BSH-2611-C'; c.src.name = 'input'; c.src.dwg = 'input';
    c.price0 = 9800; c.priceNew = 9500; c.compMode = 'amount'; c.comp = { material: 1950, process: 4300, labor: 2000, overhead: 1050 };
    c.lines[0] = Object.assign(blankLine(), { material: 'S45C', finishedDims: 'OD60 x t8 x 80', stockDims: 'OD65 x t12 x 85', src: { material: 'input', finishedDims: 'input', stockDims: 'input' } });
    // 4번째: 맞는 Index 가 없는 도금강판(SECC) — 추정 흐름
    var d = blankItem('i4');
    d.name = 'COVER, SIDE'; d.dwg = 'CVR-2611-D'; d.src = { name: 'input', dwg: 'input' };
    d.price0 = 6800; d.compMode = 'ratio'; d.comp = { material: null, process: 25, labor: 15, overhead: 8 };
    d.lines[0] = Object.assign(blankLine(), { material: 'SECC', finishedDims: 't1.2 x 400 x 600', stockDims: 't1.2 x 420 x 620', cutMm: 0, unitPrice: 1250, src: { material: 'input', finishedDims: 'input', stockDims: 'input' } });
    var book = sampleBook();
    // indexAoa(v1.0 단순 표)는 단순 형식 테스트용으로 남긴다. 예제 화면은 통합문서(book)를 쓴다.
    return { indexAoa: aoa, indexBook: book, indexFile: '원자재가격동향_예제(가상).xlsx', drawings: ex, items: [a, b, c, d], period: { start: '2022-03', end: '2025-05' } };
  }

  return {
    FAM: FAM, RELATED: RELATED, DEF: DEF, SRC_LABEL: SRC_LABEL, SHAPE_LABEL: SHAPE_LABEL, SAMPLE_DRAWINGS: SAMPLE_DRAWINGS,
    normMaterial: normMaterial, indexFamily: indexFamily, parseDims: parseDims, dimsText: dimsText, weightKg: weightKg,
    extractDrawing: extractDrawing, pdfLines: pdfLines, itemFromDrawing: itemFromDrawing, blankItem: blankItem, blankLine: blankLine,
    parseMonth: parseMonth, parseVal: parseVal, inspectSheet: inspectSheet, buildSeries: buildSeries,
    inspectBlock: inspectBlock, buildBlockSeries: buildBlockSeries, inspectBook: inspectBook, buildBook: buildBook, fxConvert: fxConvert,
    canonGrade: canonGrade, gradeOf: gradeOf, gradeTokens: gradeTokens, prefRank: prefRank, colLetter: colLetter,
    valueAt: valueAt, periodStats: periodStats, pearson: pearson, monthlyChanges: monthlyChanges, comparePattern: comparePattern,
    matchIndex: matchIndex, lineMaterial: lineMaterial, compAmounts: compAmounts, analyzeItem: analyzeItem, analyze: analyze,
    checkPeriod: checkPeriod, matchLabel: matchLabel, exportSheets: exportSheets, aiPrompt: aiPrompt,
    sample: sample, sampleIndexAoa: sampleIndexAoa, sampleBook: sampleBook, mIdx: mIdx, mStr: mStr, fmtPct: fmtPct
  };
});
