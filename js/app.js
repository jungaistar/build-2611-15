/*
 * build-2611-15 · 화면 — 그리기와 저장만 한다. 계산 · 판정은 js/linkage.js(Linkage) 에 있다.
 * 자료는 이 브라우저의 IndexedDB 에만 둔다(서버로 보내지 않는다).
 */
(function () {
  'use strict';
  var L = window.Linkage, C = window.CostIdx;
  var S = null;                       // 저장되는 상태
  var ui = { view: 'dash', vendorId: '', cDraft: null, extract: null, pDraft: null, pCols: null, px: null, runId: '', calc: null, result: null, confirm: '', cmp: '' };
  var seriesCache = null;
  var $ = function (q, el) { return (el || document).querySelector(q); };
  function h(s) { return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; }); }
  function num(v, d) { return L.fmt(v, d); }
  function blank() { return { vendors: [], contracts: [], parts: [], runs: [], approvals: [], index: null, template: null, fx: true }; }

  /* ───────────── 저장(IndexedDB, 안 되면 메모리) ───────────── */
  var idb = null;
  function openDb() {
    return new Promise(function (res) {
      try {
        var rq = indexedDB.open('build-2611-15', 1);
        rq.onupgradeneeded = function () { rq.result.createObjectStore('kv'); };
        rq.onsuccess = function () { idb = rq.result; res(); };
        rq.onerror = function () { res(); };
      } catch (e) { res(); }
    });
  }
  function kvGet(k) { return new Promise(function (res) { if (!idb) return res(null); var r = idb.transaction('kv').objectStore('kv').get(k); r.onsuccess = function () { res(r.result == null ? null : r.result); }; r.onerror = function () { res(null); }; }); }
  function kvSet(k, v) { return new Promise(function (res) { if (!idb) return res(); var t = idb.transaction('kv', 'readwrite'); t.objectStore('kv').put(v, k); t.oncomplete = function () { res(); }; t.onerror = function () { res(); }; }); }
  function save() { return kvSet('state', JSON.stringify(S)); }
  function toast(msg, err) { var t = $('#toast'); t.textContent = msg; t.className = 'toast' + (err ? ' err' : ''); t.hidden = false; clearTimeout(toast.t); toast.t = setTimeout(function () { t.hidden = true; }, err ? 6000 : 3200); }

  /* ───────────── 공통 ───────────── */
  function vendor(id) { return S.vendors.filter(function (v) { return v.id === id; })[0]; }
  function contractsOf(vid) { return S.contracts.filter(function (c) { return c.vendorId === vid; }).sort(function (a, b) { return b.version - a.version; }); }
  function partsOf(vid) { return S.parts.filter(function (p) { return p.vendorId === vid; }); }
  function contractUsed(cid) { return S.runs.some(function (r) { return r.contract.id === cid; }); }
  function series() {
    if (!S.index) return [];
    if (!seriesCache) {
      seriesCache = L.buildIndex(S.index.book, S.index.file, S.fx);
    }
    return seriesCache.series;
  }
  function pickVendor() { if (!vendor(ui.vendorId)) ui.vendorId = S.vendors[0] ? S.vendors[0].id : ''; return vendor(ui.vendorId); }
  function vendorSelect(act) {
    return '<label>업체 <select data-ch="' + act + '">' + S.vendors.map(function (v) { return '<option value="' + v.id + '"' + (v.id === ui.vendorId ? ' selected' : '') + '>' + h(v.name || '(이름 없음)') + '</option>'; }).join('') + '</select></label>';
  }
  function materialsOf(vid) {
    var out = [], add = function (m) { if (m && !L.isFxName(m) && out.indexOf(m) < 0) out.push(m); };
    var c = contractsOf(vid)[0]; if (c) c.items.forEach(function (it) { if (it.kind !== 'fx') add(it.name); });
    partsOf(vid).forEach(function (p) { Object.keys(p.usage || {}).forEach(add); });
    return out;
  }
  function download(blob, name) { var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1500); }
  function pickFile(accept, cb) { var f = $('#fileany'); f.value = ''; f.accept = accept; f.onchange = function () { if (f.files[0]) cb(f.files[0]); }; f.click(); }
  function readAb(file) { return new Promise(function (res, rej) { var r = new FileReader(); r.onload = function () { res(r.result); }; r.onerror = rej; r.readAsArrayBuffer(file); }); }
  function abToB64(ab) { var u = new Uint8Array(ab), s = ''; for (var i = 0; i < u.length; i += 0x8000) s += String.fromCharCode.apply(null, u.subarray(i, i + 0x8000)); return btoa(s); }
  function b64ToAb(b) { var s = atob(b), u = new Uint8Array(s.length); for (var i = 0; i < s.length; i++) u[i] = s.charCodeAt(i); return u.buffer; }
  function readBook(file) {
    return readAb(file).then(function (ab) {
      var wb = /\.csv$/i.test(file.name) ? XLSX.read(new TextDecoder('utf-8').decode(ab), { type: 'string', raw: true }) : XLSX.read(ab, { type: 'array' });
      var sheets = {}; wb.SheetNames.forEach(function (n) { sheets[n] = XLSX.utils.sheet_to_json(wb.Sheets[n], { header: 1, raw: true, defval: '' }); });
      return { names: wb.SheetNames, sheets: sheets };
    });
  }
  var pdfjsP = null;
  function pdfText(ab) {
    if (!pdfjsP) pdfjsP = import(new URL('lib/pdf.min.mjs', document.baseURI).href).then(function (m) { m.GlobalWorkerOptions.workerSrc = new URL('lib/pdf.worker.min.mjs', document.baseURI).href; return m; });
    return pdfjsP.then(function (m) { return m.getDocument({ data: ab, cMapUrl: new URL('lib/cmaps/', document.baseURI).href, cMapPacked: true }).promise; }).then(function (doc) {
      var ps = []; for (var i = 1; i <= doc.numPages; i++) ps.push(doc.getPage(i).then(function (pg) { return pg.getTextContent(); }));
      return Promise.all(ps).then(function (pages) { return pages.map(function (tc) { return C.pdfLines(tc.items).join('\n'); }).join('\n'); });
    });
  }
  function confirmBtn(key, label, act, extra) {
    return ui.confirm === key
      ? '<button class="btn sm danger" data-act="' + act + '" ' + (extra || '') + '>정말 ' + label + '</button> <button class="btn sm" data-act="cancel-confirm">취소</button>'
      : '<button class="btn sm danger" data-act="ask-confirm" data-key="' + key + '">' + label + '</button>';
  }

  /* ───────────── 대시보드 ───────────── */
  function quarterOf(d) { return d.getFullYear() + '년 ' + (Math.floor(d.getMonth() / 3) + 1) + '분기'; }
  function halfOf(d) { return d.getFullYear() + '년 ' + (d.getMonth() < 6 ? '상' : '하') + '반기'; }
  function inThis(cycle, month) {
    if (!month) return false;
    var d = new Date(), y = +month.slice(0, 4), m = +month.slice(5) - 1;
    if (y !== d.getFullYear()) return false;
    return cycle === 'half' ? (m < 6) === (d.getMonth() < 6) : cycle === 'year' ? true : Math.floor(m / 3) === Math.floor(d.getMonth() / 3);
  }
  function vendorStatus(v) {
    var c = contractsOf(v.id)[0], ps = partsOf(v.id), runs = S.runs.filter(function (r) { return r.vendorId === v.id; });
    var wait = 0, sim = 0;
    if (c && S.index) c.items.forEach(function (it) { var m = L.matchItem(it, v, series(), { from: '2000-01', to: '2099-12' }); if (m.needApproval) wait++; if (m.similar) sim++; });
    var last = runs[0];
    return { c: c, parts: ps.length, wait: wait, last: last, doneNow: c && runs.some(function (r) { return inThis(c.cycle, r.applyMonth); }), probs: c ? L.contractProblems(c) : ['계약 미등록'] };
  }
  function views_dash() {
    var now = new Date(), st = S.vendors.map(function (v) { return { v: v, s: vendorStatus(v) }; });
    var target = st.filter(function (x) { return x.s.c && (x.s.c.cycle === 'quarter' || x.s.c.cycle === 'half'); });
    var todo = [];
    st.forEach(function (x) {
      if (!x.s.c) todo.push([x.v.name, '계약 정보가 없습니다', 'vendors']);
      else if (x.s.probs.length) todo.push([x.v.name, '계약 확인 — ' + x.s.probs[0], 'vendors']);
      if (!x.s.parts) todo.push([x.v.name, '등록된 품번이 없습니다', 'parts']);
      if (x.s.wait) todo.push([x.v.name, 'Index 승인 대기 ' + x.s.wait + '건', 'index']);
      if (x.s.last && x.s.last.result.stopped) todo.push([x.v.name, '최근 계산에서 확인 필요 품번 ' + x.s.last.result.stopped + '건', 'reports']);
    });
    var steps = [
      [S.vendors.length > 0 && S.contracts.length > 0, '업체와 계약 수식 등록 <a href="#/vendors">업체 · 계약</a>'],
      [S.parts.length > 0, '품번별 기준 파트단가 · 원소재 사용량 등록(엑셀 또는 직접 입력) <a href="#/parts">품번 · 사용량</a>'],
      [!!S.index, '원소재 Index 엑셀 올리고 Index 맞추기 <a href="#/index">Index</a>'],
      [S.runs.length > 0, '품번별 연동액 계산 · 업체 요청단가 비교 <a href="#/calc">연동제 계산</a>'],
      [S.runs.some(function (r) { return (r.reports || []).length; }), '팀장 보고용 Word 문서 만들기 <a href="#/reports">보고서 · 이력</a>']
    ];
    return '<section class="card"><h2>대시보드 <small>' + now.getFullYear() + '-' + ('0' + (now.getMonth() + 1)).slice(-2) + ' · ' + quarterOf(now) + ' · ' + halfOf(now) + '</small></h2>' +
      '<div class="stats">' +
      '<div class="stat"><b id="st-vendors">' + S.vendors.length + '</b>등록 업체</div>' +
      '<div class="stat"><b id="st-parts">' + S.parts.length + '</b>등록 품번</div>' +
      '<div class="stat"><b>' + target.length + '</b>이번 분기 · 반기 대상 업체</div>' +
      '<div class="stat"><b>' + target.filter(function (x) { return x.s.doneNow; }).length + ' / ' + target.length + '</b>이번 기간 계산 완료</div>' +
      '<div class="stat"><b>' + todo.length + '</b>검토 대기</div>' +
      '<div class="stat"><b>' + S.runs.length + '</b>저장된 계산 이력</div></div>' +
      (S.vendors.length ? '' : '<div class="note"><p>처음이시면 <b>예제 불러오기</b>로 가짜 업체 2곳 · 품번 6개 · 가짜 Index 통합문서를 넣고 흐름을 먼저 살펴보세요.</p></div>') +
      '<div class="btnrow"><button class="btn primary" data-act="load-sample">예제 불러오기</button><button class="btn" data-act="backup">백업 내려받기(JSON)</button><button class="btn" data-act="restore">백업 불러오기</button>' + confirmBtn('wipe', '전체 지우기', 'wipe') + '</div></section>' +
      '<div class="cols2"><section class="card"><h2>진행 순서</h2><ol class="steps">' + steps.map(function (s) { return '<li class="' + (s[0] ? 'done' : '') + '">' + s[1] + '</li>'; }).join('') + '</ol></section>' +
      '<section class="card"><h2>검토 대기</h2>' + (todo.length ? '<div class="tablewrap free"><table><tr><th>업체</th><th>내용</th><th></th></tr>' + todo.map(function (t) { return '<tr><td>' + h(t[0]) + '</td><td class="wrap">' + h(t[1]) + '</td><td><a href="#/' + t[2] + '" data-act="go-vendor" data-name="' + h(t[0]) + '">열기</a></td></tr>'; }).join('') + '</table></div>' : '<p class="sub">검토를 기다리는 항목이 없습니다.</p>') + '</section></div>' +
      '<section class="card"><h2>업체별 진행 상태</h2>' + (st.length ? '<div class="tablewrap free"><table><tr><th>업체</th><th>계약</th><th>적용 주기</th><th>품번</th><th>Index 승인 대기</th><th>이번 기간</th><th>최근 계산</th></tr>' + st.map(function (x) {
        var s = x.s;
        return '<tr><td>' + h(x.v.name) + '</td><td>' + (s.c ? 'v' + s.c.version + ' · ' + h(s.c.from || '?') + ' ~ ' + h(s.c.to || '?') : '<span class="tag t-bad">미등록</span>') + '</td><td>' + (s.c ? (L.CYCLE[s.c.cycle] || '<span class="bad">확인 필요</span>') : '') + '</td><td class="num">' + s.parts + '</td><td class="num">' + (S.index ? s.wait : '-') + '</td><td>' + (s.c ? (s.doneNow ? '<span class="tag t-ok">계산 완료</span>' : '<span class="tag t-warn">대기</span>') : '') + '</td><td>' + (s.last ? h(s.last.at.slice(0, 10)) + ' · 적용 ' + h(s.last.applyMonth || '') : '<span class="muted">없음</span>') + '</td></tr>';
      }).join('') + '</table></div>' : '<p class="sub">등록된 업체가 없습니다.</p>') + '</section>' +
      '<section class="card"><h2>최근 보고서 및 이력</h2>' + runTable(S.runs.slice(0, 6)) + '</section>';
  }
  function runTable(runs) {
    if (!runs.length) return '<p class="sub">저장된 계산 이력이 없습니다.</p>';
    return '<div class="tablewrap free"><table><tr><th>저장 일시</th><th>업체</th><th>계약</th><th>적용 월</th><th>변경 Index 기간</th><th>품번</th><th>상태</th><th>보고서</th><th></th></tr>' + runs.map(function (r) {
      return '<tr><td>' + h(r.at.replace('T', ' ').slice(0, 16)) + '</td><td>' + h(r.vendorName) + '</td><td>v' + r.contract.version + '</td><td>' + h(r.applyMonth || '') + '</td><td>' + h(r.period.next.from + ' ~ ' + r.period.next.to) + '</td><td class="num">' + r.result.done + ' / ' + r.result.rows.length + '</td><td>' + (r.result.stopped ? '<span class="tag t-warn">확인 필요 ' + r.result.stopped + '</span>' : '<span class="tag t-ok">승인 저장</span>') + '</td><td class="num">' + (r.reports || []).length + '</td><td><a href="#/reports" data-act="open-run" data-id="' + r.id + '">다시 보기</a></td></tr>';
    }).join('') + '</table></div>';
  }

  /* ───────────── 업체 · 계약 ───────────── */
  function loadDraft(v, cid) {
    var list = v ? contractsOf(v.id) : [], c = list.filter(function (x) { return x.id === cid; })[0] || list[0];
    ui.cDraft = c ? L.clone(c) : (v ? L.blankContract(v.id) : null);
    if (ui.cDraft && !c) ui.cDraft.isNew = true;
    ui.extract = null;
  }
  function views_vendors() {
    var v = pickVendor();
    if (v && (!ui.cDraft || ui.cDraft.vendorId !== v.id)) loadDraft(v);
    var left = '<ul class="list">' + S.vendors.map(function (x) { var c = contractsOf(x.id)[0]; return '<li class="' + (x.id === ui.vendorId ? 'on' : '') + '"><button data-act="sel-vendor" data-id="' + x.id + '">' + h(x.name || '(이름 없음)') + '<small>' + (c ? '계약 v' + c.version + ' · 품번 ' + partsOf(x.id).length : '계약 미등록') + '</small></button></li>'; }).join('') + '</ul><div class="btnrow"><button class="btn sm primary" data-act="add-vendor">업체 추가</button></div>';
    if (!v) return '<section class="card"><h2>업체 및 계약 관리</h2><p class="sub">업체를 먼저 추가하세요. 계약 수식은 업체마다 한 번 등록하고, 같은 업체의 품번이 함께 씁니다.</p><div class="btnrow"><button class="btn primary" data-act="add-vendor">업체 추가</button></div></section>';
    var c = ui.cDraft, list = contractsOf(v.id), used = !c.isNew && contractUsed(c.id), probs = L.contractProblems(c);
    var ver = list.length ? '<p class="sub">계약 버전: ' + list.map(function (x) { return '<button class="x" data-act="sel-contract" data-id="' + x.id + '">' + (x.id === c.id ? '<b>v' + x.version + '</b>' : 'v' + x.version) + '</button> <span class="muted">(' + h(x.from || '?') + ' ~ ' + h(x.to || '?') + ')</span>'; }).join(' · ') + '</p>' : '';
    var itemsRows = c.items.map(function (it, i) {
      var fx = it.kind === 'fx', b = function (f) { return 'data-bind="ci:' + i + ':' + f + '"'; };
      return '<tr><td><input type="text" class="w-m" ' + b('name') + ' value="' + h(it.name) + '" placeholder="철강 · 구리 · 환율…"></td>' +
        '<td><select class="w-s" ' + b('kind') + ' data-rr="1"><option value="material"' + (!fx ? ' selected' : '') + '>원소재</option><option value="fx"' + (fx ? ' selected' : '') + '>환율</option></select></td>' +
        '<td><input type="text" class="w-s" ' + b('spec') + ' value="' + h(it.spec) + '" placeholder="' + (fx ? 'USD' : 'SS400') + '"></td>' +
        (fx ? '<td colspan="2"><select class="w-l" ' + b('fxBase') + '><option value="">환율을 곱할 금액 — 계약 확인</option>' + Object.keys(L.FX_BASE).map(function (k) { return '<option value="' + k + '"' + (it.fxBase === k ? ' selected' : '') + '>' + L.FX_BASE[k] + '</option>'; }).join('') + '</select></td><td class="fxna" title="환율은 계약상 100% 적용 — 고칠 수 없습니다">100%</td>'
          : '<td><select class="w-m" ' + b('formula') + ' data-rr="1"><option value="">수식 미정(계산 안 함)</option>' + Object.keys(L.FORMULA).map(function (k) { return '<option value="' + k + '"' + (it.formula === k ? ' selected' : '') + '>' + L.FORMULA[k] + '</option>'; }).join('') + '</select></td>' +
            '<td>' + (it.formula === 'ratio' ? '<input type="number" class="w-s" step="any" data-t="num" ' + b('share') + ' value="' + h(it.share) + '" placeholder="비중 %">' : it.formula === 'custom' ? '<input type="text" class="w-l" ' + b('expr') + ' value="' + h(it.expr) + '" placeholder="KG * (I1 - I0)">' : '<span class="muted">사용량(kg) × Index 차이</span>') + '</td>' +
            '<td><input type="number" class="w-xs" step="any" data-t="num" ' + b('reflect') + ' value="' + h(it.reflect) + '"></td>') +
        '<td><input type="number" class="w-xs" step="any" data-t="num" ' + b('minChange') + ' value="' + h(it.minChange) + '"></td>' +
        '<td><input type="number" class="w-xs" step="any" data-t="num" ' + b('cap') + ' value="' + h(it.cap) + '"></td>' +
        '<td><input type="number" class="w-xs" step="any" data-t="num" ' + b('floor') + ' value="' + h(it.floor) + '"></td>' +
        '<td>' + (fx ? '<span class="muted">-</span>' : '<input type="text" class="w-m" data-t="tiers" ' + b('tiers') + ' value="' + h((it.tiers || []).map(function (t) { return t[0] + ':' + t[1]; }).join(', ')) + '" placeholder="5:50, 10:100">') + '</td>' +
        '<td><input type="number" class="w-s" step="any" data-t="num" ' + b('baseValue') + ' value="' + h(it.baseValue) + '" placeholder="없으면 비움"></td>' +
        '<td>' + (it.indexName ? '<span class="tag ' + (it.indexStatus === 'similar' ? 't-warn' : 't-ok') + '">' + (it.indexStatus === 'similar' ? '유사 승인' : '승인') + '</span>' : '<span class="muted">자동</span>') + '</td>' +
        '<td><button class="x del" data-act="del-item" data-i="' + i + '">삭제</button></td></tr>';
    }).join('');
    var ex = ui.extract ? extractPanel(ui.extract) : '';
    var right = '<section class="card"><h2>업체 정보</h2><div class="form">' +
      '<label class="wide">업체명 <input type="text" data-bind="v:name" value="' + h(v.name) + '"></label>' +
      '<label>Index 엑셀의 업체 시트 이름(다르면) <input type="text" data-bind="v:sheetHint" value="' + h(v.sheetHint) + '" placeholder="비우면 업체명으로 찾음"></label>' +
      '<label class="wide">메모 <input type="text" data-bind="v:memo" value="' + h(v.memo) + '"></label></div>' +
      '<div class="btnrow"><button class="btn sm primary" data-act="save-vendor">업체 정보 저장</button>' + confirmBtn('delv', '업체 삭제', 'del-vendor') + '</div></section>' +
      '<section class="card"><h2>계약 정보 <small>' + (c.isNew ? '새 계약' : 'v' + c.version + (used ? ' · 계산에 쓰인 버전' : '')) + '</small></h2>' + ver +
      '<p class="sub">계약서에 적힌 수식이 가장 우선입니다. 수식이 분명하지 않으면 「수식 미정」으로 두세요 — 그 항목은 계산하지 않고 확인을 요청합니다.</p>' +
      '<div class="btnrow"><button class="btn sm" data-act="up-contract">계약서(Word · PDF) 올려 항목 찾기</button>' + (c.file ? '<span class="muted" style="align-self:center">올린 파일: ' + h(c.file) + '</span>' : '') + '</div>' + ex +
      '<div class="form">' +
      '<label>최초 계약일 <input type="text" data-bind="c:firstDate" value="' + h(c.firstDate) + '" placeholder="2026-01-02"></label>' +
      '<label class="wide">1. 물품등의 명칭(연동제 적용 품목) <input type="text" data-bind="c:product" value="' + h(c.product) + '"></label>' +
      '<label>적용 시작월 <input type="month" data-bind="c:from" value="' + h(c.from) + '"></label>' +
      '<label>적용 종료월 <input type="month" data-bind="c:to" value="' + h(c.to) + '"></label>' +
      '<label>5. 조정요건 — 변동률 ±% <input type="number" step="any" data-t="num" data-bind="c:trigger" value="' + h(c.trigger) + '" placeholder="예) 5"></label>' +
      '<label>6. 조정주기 <select data-bind="c:cycle"><option value="">확인 필요</option>' + Object.keys(L.CYCLE).map(function (k) { return '<option value="' + k + '"' + (c.cycle === k ? ' selected' : '') + '>' + L.CYCLE[k] + '</option>'; }).join('') + '</select></label>' +
      '<label>변경 단가 끝수 <select data-bind="c:round"><option value="none"' + (c.round === 'none' ? ' selected' : '') + '>그대로(반올림 안 함)</option><option value="1"' + (c.round === '1' ? ' selected' : '') + '>원 단위 반올림</option><option value="10"' + (c.round === '10' ? ' selected' : '') + '>10원 단위 반올림</option></select></label>' +
      '<label class="wide">7. 조정일 <input type="text" data-bind="c:adjustDay" value="' + h(c.adjustDay) + '"></label>' +
      '<label class="wide">8. 조정대금 반영시점(갱신 시점) <input type="text" data-bind="c:reflectDay" value="' + h(c.reflectDay) + '" placeholder="예) 01-01, 04-01, 07-01, 10-01"></label>' +
      '<label class="full">3. 원재료 가격의 기준지표(계약서 문구) <input type="text" data-bind="c:indexText" value="' + h(c.indexText) + '"></label>' +
      '<label class="full">4. 기준시점 · 비교시점(계약서 문구) <input type="text" data-bind="c:baseRule" value="' + h(c.baseRule) + '" placeholder="예) 기준시점: 직전 조정 분기 평균 / 비교시점: 조정일 직전 분기 평균"></label>' +
      '<label class="full">9. 납품대금 연동산식(계약서 문구 그대로 — 자동으로 해석하지 않습니다) <textarea data-bind="c:formulaText">' + h(c.formulaText) + '</textarea></label>' +
      '<label class="full">연동 제외 조건 · 기타 예외 <input type="text" data-bind="c:exclude" value="' + h(c.exclude) + '"></label></div>' +
      '<label class="chk"><input type="checkbox" data-t="bool" data-bind="c:allowSubstitute"' + (c.allowSubstitute ? ' checked' : '') + '> 계약상 대체(유사) Index 사용이 허용됨 — 끄면 정확히 일치하는 Index 가 없을 때 자동 대체하지 않고 계산을 멈춥니다</label>' +
      '<h3>원소재 · 환율 연동 항목</h3>' +
      '<div class="tablewrap free"><table class="grid" id="items"><tr><th>이름</th><th>구분</th><th>규격 · 통화</th><th>수식</th><th>비중 · 수식</th><th>반영률 %</th><th title="비우면 계약의 조정요건(±%)을 씁니다">최소 변동 %(항목별)</th><th>상한 %</th><th>하한 %</th><th>구간별 반영(변동%:반영%)</th><th>기준 Index 고정값</th><th>Index</th><th></th></tr>' + (itemsRows || '<tr><td colspan="13" class="muted">연동 항목이 없습니다 — 아래 단추로 추가하세요.</td></tr>') + '</table></div>' +
      '<div class="btnrow"><button class="btn sm" data-act="add-item">원소재 추가</button><button class="btn sm" data-act="add-fx">환율 추가</button></div>' +
      '<details class="box"><summary>수식 종류 설명</summary><ul style="margin:8px 0 0 18px;font-size:13.5px">' +
      '<li><b>중량 연동식</b> — 연동액 = 품번의 사용량(kg) × 기준 Index(원/kg) × 적용 변동률. Index 단위가 원/kg · 원/ton 이어야 합니다.</li>' +
      '<li><b>비중 연동식</b> — 연동액 = 기준 파트단가 × 비중(%) × 적용 변동률.</li>' +
      '<li><b>사용자 정의 수식</b> — 사칙연산과 괄호. 쓸 수 있는 이름: ' + Object.keys(L.EXPR_VARS).map(function (k) { return '<code>' + k + '</code> ' + L.EXPR_VARS[k]; }).join(' · ') + '</li>' +
      '<li><b>환율</b> — 사용량(kg)이 없습니다. 연동액 = (기준 파트단가 전체 또는 품번별 환율 대상 금액) × 환율 변동률. 적용 비중은 100% 로 고정됩니다.</li>' +
      '<li><b>조정요건</b> — 기준시점 대비 변동률이 ±조정요건(%) 미만이면 그 항목은 이번에 조정하지 않습니다(연동액 0). 항목별 「최소 변동 %」를 넣으면 그 항목은 그 값을 씁니다.</li>' +
      '<li>적용 변동률 = 조정요건(최소 변동률) 미만이면 0 → 상한 · 하한 → 구간별 반영률(또는 반영률) 순서로 적용합니다.</li></ul></details>' +
      (probs.length ? '<div class="warn" id="c-probs"><b>이대로는 계산하지 않는 항목</b><ul>' + probs.map(function (m) { return '<li>' + h(m) + '</li>'; }).join('') + '</ul></div>' : '<div class="ok" id="c-probs">계약 정보가 계산할 수 있는 상태입니다.</div>') +
      '<div class="btnrow">' + (c.isNew ? '<button class="btn primary" data-act="save-contract">계약 저장(v' + ((list[0] ? list[0].version : 0) + 1) + ')</button>'
        : (used ? '' : '<button class="btn primary" data-act="save-contract">이 버전 고쳐 저장</button>') + '<button class="btn' + (used ? ' primary' : '') + '" data-act="save-contract-new">새 계약 버전으로 저장(v' + (list[0].version + 1) + ')</button>') + '</div>' +
      (used ? '<p class="sub" style="margin-top:8px">이 버전은 이미 계산 이력에 쓰였습니다. 과거 결과가 바뀌지 않도록 고친 내용은 새 버전으로만 저장합니다.</p>' : '') + '</section>';
    return '<section class="card"><h2>업체 및 계약 관리</h2><div class="split"><div>' + left + '</div><div style="min-width:0">' + right + '</div></div></section>';
  }
  function extractPanel(ex) {
    var F = { vendor: '업체명', firstDate: '최초 계약일', product: '1. 물품등의 명칭(적용 품목)', indexText: '3. 가격 기준지표', trigger: '5. 조정요건(±%)', adjustDay: '7. 조정일', reflectDay: '8. 조정대금 반영시점', formulaText: '9. 납품대금 연동산식', from: '적용 시작월', to: '적용 종료월', cycle: '적용 주기', baseRule: '기준일 문구', cap: '상한 %', floor: '하한 %', minChange: '최소 변동률 %', substitute: '대체 Index 문구', exclude: '연동 제외 문구' };
    var rows = Object.keys(F).filter(function (k) { return ex.fields[k]; }).map(function (k) {
      var f = ex.fields[k];
      return '<tr><td>' + F[k] + '</td><td>' + h(k === 'cycle' ? L.CYCLE[f.value] : f.value) + '</td><td class="wrap muted">' + h(f.line) + '</td><td>' + (k === 'vendor' || k === 'substitute' ? '' : '<button class="x" data-act="use-ex" data-k="' + k + '">넣기</button>') + '</td></tr>';
    }).join('');
    return '<div class="note" id="extract">' + (ex.standard ? '<p><span class="tag t-info">표준 약정서</span> 「납품대금 연동 특별약정서(별첨)」의 1~9번 항목을 번호로 읽었습니다(10번 기타사항 뒤는 읽지 않습니다).</p>' : '') + '<b>계약서에서 찾은 후보 — 자동으로 넣지 않습니다. 확인하고 「넣기」를 누르세요.</b>' +
      (rows ? '<div class="tablewrap free" style="margin-top:8px"><table><tr><th>항목</th><th>찾은 값</th><th>근거 문장</th><th></th></tr>' + rows + '</table></div>' : '<p>기간 · 주기 · 상한 · 하한으로 보이는 문장을 찾지 못했습니다.</p>') +
      (ex.materials.length ? '<p style="margin-top:8px">원소재 · 환율로 보이는 낱말: ' + ex.materials.map(function (m, i) { return '<button class="x" data-act="use-mat" data-i="' + i + '">' + h(m.name) + (m.share != null ? '(' + m.share + '%)' : '') + ' 추가</button>'; }).join(' · ') + '</p>' : '') +
      (ex.formulaLines.length ? '<p style="margin-top:8px">수식으로 보이는 줄:</p><ul>' + ex.formulaLines.map(function (f, i) { return '<li>' + h(f) + ' <button class="x" data-act="use-formula" data-i="' + i + '">계약 수식 칸에 넣기</button></li>'; }).join('') + '</ul>' : '') +
      ex.notes.map(function (n) { return '<p class="bad">' + h(n) + '</p>'; }).join('') + '</div>';
  }

  /* ───────────── 품번 · 사용량 ───────────── */
  function loadParts(v) {
    ui.pDraft = partsOf(v.id).map(L.clone); ui.pCols = materialsOf(v.id); ui.pVendor = v.id; ui.px = null;
  }
  function views_parts() {
    var v = pickVendor();
    if (!v) return '<section class="card"><h2>품번 및 원소재 사용량 관리</h2><p class="sub">업체를 먼저 등록하세요. <a href="#/vendors">업체 · 계약</a></p></section>';
    if (ui.pVendor !== v.id || !ui.pDraft) loadParts(v);
    var c = contractsOf(v.id)[0], hasFx = c && c.items.some(function (it) { return it.kind === 'fx'; });
    var head = '<section class="card"><h2>품번 및 원소재 사용량 관리</h2><p class="sub">기준 파트단가와 원소재 사용량은 <b>품번마다 따로</b> 관리합니다. 빈 칸은 「미입력」, 0 은 「쓰지 않음」으로 구분합니다. 환율은 사용량(kg)이 없습니다.</p>' +
      '<div class="form">' + vendorSelect('parts-vendor') + '</div>' +
      '<div class="btnrow"><button class="btn sm" data-act="up-parts">엑셀로 한 번에 올리기</button><button class="btn sm" data-act="dl-parts-template">엑셀 양식 내려받기</button><button class="btn sm" data-act="dl-parts">등록된 품번 엑셀로 받기</button></div></section>';
    if (ui.px) return head + partsExcel(v);
    var cols = ui.pCols;
    var rows = ui.pDraft.map(function (p, i) {
      var errs = L.checkPart(p), b = function (f) { return 'data-bind="p:' + i + ':' + f + '"'; };
      return '<tr class="' + (errs.length ? 'err' : '') + '"><td><input type="text" class="w-m" ' + b('partNo') + ' value="' + h(p.partNo) + '"></td><td><input type="text" class="w-l" ' + b('name') + ' value="' + h(p.name) + '"></td>' +
        '<td><input type="number" class="w-m" step="any" data-t="num" ' + b('basePrice') + ' value="' + h(p.basePrice) + '"></td><td><input type="number" class="w-m" step="any" data-t="num" ' + b('curPrice') + ' value="' + h(p.curPrice) + '"></td>' +
        cols.map(function (m) { return '<td><input type="number" class="w-s" step="any" min="0" data-t="num" data-bind="pu:' + i + ':' + h(m) + '" value="' + h(p.usage[m]) + '"></td>'; }).join('') +
        '<td class="fxna">-</td><td style="text-align:center"><input type="checkbox" data-t="bool" ' + b('fx') + (p.fx ? ' checked' : '') + '></td><td><input type="number" class="w-m" step="any" data-t="num" ' + b('fxAmount') + ' value="' + h(p.fxAmount) + '"' + (p.fx ? '' : ' disabled') + '></td>' +
        '<td class="wrap bad">' + h(errs.join(' / ')) + '</td><td><button class="x del" data-act="del-part" data-i="' + i + '">행 삭제</button></td></tr>';
    }).join('');
    var bad = ui.pDraft.filter(function (p) { return L.checkPart(p).length; }).length, dup = dupParts(ui.pDraft);
    return head + '<section class="card"><h2>직접 입력 <small>' + h(v.name) + ' · ' + ui.pDraft.length + '개 품번</small></h2>' +
      '<div class="tablewrap free"><table class="grid" id="parts"><tr><th>품번</th><th>품명(선택)</th><th>기준 파트단가(원/EA)</th><th>현재 구매단가(원/EA)</th>' +
      cols.map(function (m, k) { return '<th>' + h(m) + '(kg) <button class="x" data-act="ren-col" data-k="' + k + '">이름</button><button class="x del" data-act="del-col" data-k="' + k + '">열 삭제</button></th>'; }).join('') +
      '<th title="환율은 사용량이 없습니다">환율(kg)</th><th>환율 적용</th><th>환율 대상 금액(원/EA)</th><th>확인</th><th></th></tr>' + (rows || '<tr><td colspan="' + (9 + cols.length) + '" class="muted">품번이 없습니다 — 「행 추가」로 넣거나 엑셀로 올리세요.</td></tr>') + '</table></div>' +
      '<div class="btnrow"><button class="btn sm" data-act="add-part">행 추가(품번)</button><span style="display:inline-flex;gap:6px;align-items:center"><input type="text" id="newcol" class="w-m" placeholder="새 원소재 이름"><button class="btn sm" data-act="add-col">열 추가(원소재)</button></span></div>' +
      (hasFx ? '<p class="sub" style="margin-top:8px">이 업체 계약에는 환율 항목이 있습니다. 환율을 적용하는 품번에 「환율 적용」을 켜세요. 계약이 「품번별 환율 대상 금액」 기준이면 금액도 넣어야 합니다.</p>' : '') +
      (dup.length ? '<div class="warn">품번이 겹칩니다: ' + h(dup.join(', ')) + '</div>' : '') +
      '<div class="' + (bad || dup.length ? 'warn' : 'ok') + '" id="parts-check">저장 전 확인 — 품번 ' + ui.pDraft.length + '개 · 원소재 열 ' + cols.length + '개 · 오류 ' + bad + '행' + (bad ? ' (오류가 있는 행을 고쳐야 저장할 수 있습니다)' : '') + '</div>' +
      '<div class="btnrow"><button class="btn primary" data-act="save-parts"' + (bad || dup.length ? ' disabled' : '') + '>저장</button><button class="btn" data-act="reload-parts">되돌리기</button></div></section>';
  }
  function dupParts(list) { var s = {}, d = []; list.forEach(function (p) { var k = L.key(p.partNo); if (!k) return; if (s[k] && d.indexOf(p.partNo) < 0) d.push(p.partNo); s[k] = 1; }); return d; }
  function partsExcel(v) {
    var px = ui.px, roles = { partNo: '품번', name: '품명', basePrice: '기준 파트단가', curPrice: '현재 구매단가', usage: '원소재 사용량(kg)', fx: '환율 적용', fxAmount: '환율 대상 금액', skip: '쓰지 않음' };
    var map = '<div class="tablewrap free"><table class="grid" id="pxmap"><tr><th>엑셀 열</th><th>머리글</th><th>역할</th><th>원소재 이름</th></tr>' + px.insp.columns.map(function (c, i) {
      return '<tr><td>' + C.colLetter(c.col) + '</td><td>' + h(c.head) + '</td><td><select class="w-l" data-ch="px-role" data-i="' + i + '">' + Object.keys(roles).map(function (r) { return '<option value="' + r + '"' + (c.role === r ? ' selected' : '') + '>' + roles[r] + '</option>'; }).join('') + '</select></td><td>' + (c.role === 'usage' ? '<input type="text" class="w-m" data-ch="px-mat" data-i="' + i + '" value="' + h(c.material) + '">' : '') + '</td></tr>';
    }).join('') + '</table></div>';
    var r = px.res, body = '';
    if (!r.ok) body = '<div class="warn"><ul>' + r.problems.map(function (m) { return '<li>' + h(m) + '</li>'; }).join('') + '</ul></div>';
    else {
      body = '<div class="stats" style="margin:12px 0"><div class="stat"><b id="px-new">' + r.counts['new'] + '</b>신규 품번</div><div class="stat"><b id="px-upd">' + r.counts.update + '</b>기존 품번(덮어쓰기)</div><div class="stat"><b id="px-err">' + r.counts.error + '</b>오류 행(저장 안 함)</div><div class="stat"><b>' + r.materials.length + '</b>원소재 열</div></div>' +
        '<div class="tablewrap"><table id="pxprev"><tr><th>행</th><th>상태</th><th>품번</th><th>품명</th><th>기준 파트단가</th><th>현재 구매단가</th>' + r.materials.map(function (m) { return '<th>' + h(m) + '(kg)</th>'; }).join('') + '<th>환율</th><th>오류</th></tr>' + r.rows.map(function (x) {
          return '<tr class="' + (x.status === 'error' ? 'err' : x.status === 'update' ? 'upd' : '') + '"><td>' + x.row + '</td><td>' + (x.status === 'error' ? '<span class="tag t-bad">오류</span>' : x.status === 'update' ? '<span class="tag t-warn">기존 — 덮어씀</span>' : '<span class="tag t-ok">신규</span>') + '</td><td>' + h(x.partNo) + '</td><td>' + h(x.name) + '</td><td class="num">' + num(x.basePrice) + '</td><td class="num">' + num(x.curPrice) + '</td>' + r.materials.map(function (m) { var u = x.usage[m]; return '<td class="num">' + (u == null ? '<span class="muted">빈 칸</span>' : num(u, 3)) + '</td>'; }).join('') + '<td>' + (x.fx ? '적용' + (x.fxAmount != null ? ' ' + num(x.fxAmount) : '') : '') + '</td><td class="wrap bad">' + h(x.errors.map(function (e) { return '[' + e.col + '] ' + e.msg; }).join(' / ')) + '</td></tr>';
        }).join('') + '</table></div>' +
        (r.counts.update ? '<label class="chk"><input type="checkbox" id="px-ok"> 이미 등록된 품번 ' + r.counts.update + '개를 이 파일의 값으로 덮어쓰는 데 동의합니다</label>' : '');
    }
    return '<section class="card"><h2>엑셀 대량 등록 — 미리보기 <small>' + h(px.file) + ' · 시트 ' + h(px.sheet) + '</small></h2>' +
      (px.sheets.length > 1 ? '<div class="form"><label>시트 <select data-ch="px-sheet">' + px.sheets.map(function (n) { return '<option' + (n === px.sheet ? ' selected' : '') + '>' + h(n) + '</option>'; }).join('') + '</select></label></div>' : '') +
      (px.insp.ok ? '' : '<div class="warn">' + h(px.insp.error) + '</div>') +
      '<details class="box"' + (px.insp.ok && r.ok ? '' : ' open') + '><summary>열 지정 — 자동으로 알아본 결과입니다. 다르면 고치세요</summary>' + map + '</details>' + body +
      '<div class="btnrow"><button class="btn primary" data-act="px-approve"' + (r.ok && (r.counts['new'] + r.counts.update) ? '' : ' disabled') + '>승인하고 저장(오류 행 제외)</button><button class="btn" data-act="px-cancel">취소</button></div></section>';
  }
  function pxParse() { var px = ui.px, v = pickVendor(); px.res = L.parseParts(px.book.sheets[px.sheet], px.insp, partsOf(v.id).map(function (p) { return p.partNo; })); }

  /* ───────────── Index ───────────── */
  function views_index() {
    var v = pickVendor(), out = '<section class="card"><h2>Index 데이터 관리</h2><p class="sub">구매기획팀이 배포한 원소재 Index 엑셀을 올리세요. <b>올린 파일의 값만</b> 씁니다 — 인터넷 가격이나 만들어 낸 Index 는 쓰지 않습니다.</p>' +
      '<div class="btnrow"><button class="btn primary" data-act="up-index">Index 엑셀 올리기(.xlsx · .csv)</button><button class="btn" data-act="dl-sample-index">예제 Index 통합문서 내려받기(가짜)</button>' + (S.index ? confirmBtn('delidx', 'Index 파일 지우기', 'del-index') : '') + '</div>';
    if (!S.index) return out + '<p class="sub" style="margin-top:10px">아직 올린 파일이 없습니다.</p></section>';
    var all = series(), sc = seriesCache;
    out += '<dl class="kv"><dt>파일</dt><dd id="index-file">' + h(S.index.file) + '</dd><dt>올린 날</dt><dd>' + h(S.index.at.slice(0, 10)) + '</dd><dt>읽은 계열</dt><dd>' + all.length + '개 (시트 ' + S.index.book.names.length + '장)</dd></dl>' +
      '<label class="chk"><input type="checkbox" data-ch="fx-toggle"' + (S.fx ? ' checked' : '') + '> USD · RMB 단위 Index 를 같은 파일의 환율 시트로 원화 환산한 계열도 만들기(중량 연동식에 필요)</label>' +
      (sc.notes.length ? '<p class="sub">' + sc.notes.map(h).join('<br>') + '</p>' : '') + '</section>';
    // 시트 구조
    out += '<section class="card"><h2>Sheet 구조 분석</h2><div class="tablewrap free"><table id="sheets"><tr><th>시트</th><th>형식</th><th>업체 시트</th><th>Index 열</th><th>데이터 기간</th><th>빈 달</th><th>알림</th></tr>' + sc.sheets.map(function (sh) {
      var ss = all.filter(function (s) { return s.meta && (s.meta.fileSheet || s.meta.sheet) === sh.sheet && !s.converted; });
      var ms = [].concat.apply([], ss.map(function (s) { return s.points.map(function (p) { return p.m; }); })).sort();
      var gaps = 0; ss.forEach(function (s) { if (s.points.length) gaps += L.monthsBetween(s.points[0].m, s.points[s.points.length - 1].m).length - s.points.length; });
      var owner = S.vendors.filter(function (x) { return ss[0] && L.sheetMatches(ss[0], x); }).map(function (x) { return x.name; });
      return '<tr><td>' + h(sh.sheet) + '</td><td>' + (!sh.ok ? '<span class="tag t-bad">못 읽음</span>' : sh.kind === 'wide' ? '가로형(원소재가 행 · 날짜가 열)' : sh.kind === 'block' ? '머리 블록(품목이 열)' : '세로형(날짜 열 + Index 열)') + '</td><td>' + (owner.length ? '<span class="tag t-info">' + h(owner.join(', ')) + '</span>' : '') + '</td><td class="num">' + ss.length + '</td><td>' + (ms.length ? ms[0] + ' ~ ' + ms[ms.length - 1] : '') + '</td><td class="num">' + (ss.length ? gaps : '') + '</td><td class="wrap">' + h(sh.ok ? '' : sh.error) + '</td></tr>';
    }).join('') + '</table></div>' + (sc.issues.length ? '<details class="box"><summary>읽으면서 뺀 값 ' + sc.issues.length + '건</summary><ul style="margin:8px 0 0 18px;font-size:13px">' + sc.issues.slice(0, 60).map(function (i) { return '<li>' + h(i.msg) + '</li>'; }).join('') + '</ul></details>' : '') + '</section>';
    if (!v) return out;
    var c = contractsOf(v.id)[0];
    out += '<section class="card"><h2>계약 원소재 ↔ Index 맞추기</h2><div class="form">' + vendorSelect('index-vendor') + '</div>';
    if (!c) return out + '<p class="sub">이 업체는 계약 정보가 없습니다. <a href="#/vendors">업체 · 계약</a></p></section>';
    var span = matchSpan(c);
    out += '<p class="sub">업체명과 같은 시트를 먼저 찾고, 그 시트에 계약 원소재와 맞는 Index 가 없으면 다른 시트에서 찾습니다. 유사 Index 는 승인해야 계산에 씁니다' + (c.allowSubstitute ? '' : ' — <b class="bad">이 계약은 대체 Index 사용이 허용돼 있지 않습니다</b>') + '. 유사성 비교 기간: ' + span.from + ' ~ ' + span.to + '</p>';
    out += c.items.map(function (it) {
      var m = L.matchItem(it, v, all, span);
      var tag = { approved: m.similar ? ['t-warn', '유사 Index 승인됨'] : ['t-ok', '승인됨'], exact: ['t-ok', '정확 일치 — 자동 적용'], multi: ['t-warn', '여러 개 — 선택 필요'], similar: ['t-warn', '유사 후보 — 승인 필요'], none: ['t-bad', '없음'] }[m.status];
      var simi = m.status === 'similar';
      var tbl = m.candidates.length ? '<div class="tablewrap free"><table class="cands"><tr><th>후보 Index</th><th>시트</th><th>단위</th>' + (simi ? '<th>재질 유사도</th><th>가격 변동 상관</th><th>방향 일치</th><th>기간 변동률</th><th>데이터 완전성</th><th>유사성 점수</th>' : '') + '<th>근거</th><th></th></tr>' + m.candidates.map(function (cd) {
        var pc = function (x) { return x == null ? '<span class="muted">비교 불가</span>' : Math.round(x * 100) + '%'; };
        return '<tr><td>' + h(cd.name) + '</td><td>' + h(cd.sheet) + '</td><td>' + h(cd.unit) + '</td>' + (simi ? '<td class="num">' + pc(cd.matSim) + '</td><td class="num">' + (cd.corr == null ? '<span class="muted">비교 불가</span>' : cd.corr.toFixed(2)) + '</td><td class="num">' + pc(cd.dirAgree) + '</td><td class="num">' + L.fmtPct(cd.rate) + '</td><td class="num">' + pc(cd.complete) + '</td><td class="num"><b>' + (cd.score == null ? '' : cd.score.toFixed(2)) + '</b></td>' : '') + '<td class="wrap">' + h(cd.why) + '</td><td>' + (m.status === 'exact' || m.status === 'approved' ? '' : '<button class="btn sm" data-act="approve-index" data-item="' + it.id + '" data-name="' + h(cd.name) + '" data-st="' + (simi ? 'similar' : 'exact') + '"' + (simi && !c.allowSubstitute ? ' disabled title="계약상 대체 Index 사용 불허"' : '') + '>승인</button>') + '</td></tr>';
      }).join('') + '</table></div>' : '';
      var pool = all.filter(function (s) { return it.kind === 'fx' ? (s.meta && s.meta.kind === 'fx') : s.role !== 'skip' && !(s.meta && s.meta.kind === 'fx'); });
      return '<div class="mcard st-' + m.status + '" data-item="' + h(it.name) + '"><div class="mhead"><b>' + h(it.name) + (it.spec ? ' <span class="muted">(' + h(it.spec) + ')</span>' : '') + '</b><span class="tag ' + tag[0] + '">' + tag[1] + '</span>' + (m.where === 'vendor' ? '<span class="tag t-info">업체 시트</span>' : '') + '</div>' +
        (m.chosen ? '<p>적용 Index: <b>' + h(m.chosen) + '</b></p>' : '') + '<p class="muted">' + h(m.reason) + '</p>' + tbl +
        m.notes.map(function (n) { return '<p class="muted">· ' + h(n) + '</p>'; }).join('') +
        '<div class="mctl"><select data-pick="' + it.id + '"><option value="">직접 고르기…</option>' + pool.map(function (s) { return '<option value="' + h(s.name) + '">' + h('[' + ((s.meta && s.meta.sheet) || '') + '] ' + s.name) + '</option>'; }).join('') + '</select><button class="btn sm" data-act="approve-pick" data-item="' + it.id + '">고른 Index 승인(대체 지표)</button>' + (it.indexName ? '<button class="btn sm danger" data-act="unapprove" data-item="' + it.id + '">승인 취소</button>' : '') + '</div></div>';
    }).join('');
    var log = S.approvals.filter(function (a) { return a.vendorId === v.id; }).slice(0, 8);
    if (log.length) out += '<details class="box"><summary>Index 연결 · 승인 이력</summary><ul style="margin:8px 0 0 18px;font-size:13px">' + log.map(function (a) { return '<li>' + h(a.at.replace('T', ' ').slice(0, 16)) + ' · ' + h(a.item) + ' → ' + h(a.index || '(승인 취소)') + ' · ' + h(a.type) + ' · ' + h(a.file) + '</li>'; }).join('') + '</ul></details>';
    return out + '</section>';
  }
  function matchSpan(c) {
    var ms = [].concat.apply([], series().map(function (s) { return s.points.length ? [s.points[s.points.length - 1].m] : []; })).sort(), last = ms[ms.length - 1] || '2026-12';
    return { from: C.mStr(C.mIdx(last) - 35), to: last };
  }

  /* ───────────── 연동제 계산 ───────────── */
  function initCalc(v) {
    var c = contractsOf(v.id)[0], d = new Date(), am = d.getFullYear() + '-' + ('0' + (d.getMonth() + 1)).slice(-2);
    var p = L.suggestPeriod(am, c ? c.cycle : 'quarter');
    ui.calc = { vendorId: v.id, contractId: c ? c.id : '', applyMonth: am, base: p.base, next: p.next, requests: {}, opinion: '' };
    ui.result = null;
  }
  function views_calc() {
    var v = pickVendor();
    if (!v) return '<section class="card"><h2>연동제 계산</h2><p class="sub">업체를 먼저 등록하세요. <a href="#/vendors">업체 · 계약</a></p></section>';
    if (!ui.calc || ui.calc.vendorId !== v.id) initCalc(v);
    var k = ui.calc, list = contractsOf(v.id), c = list.filter(function (x) { return x.id === k.contractId; })[0], ps = partsOf(v.id);
    var out = '<section class="card"><h2>연동제 계산</h2><div class="form">' + vendorSelect('calc-vendor') +
      '<label>계약 버전 <select data-bind="k:contractId" data-rr="1">' + list.map(function (x) { return '<option value="' + x.id + '"' + (x.id === k.contractId ? ' selected' : '') + '>v' + x.version + ' (' + h(x.from || '?') + ' ~ ' + h(x.to || '?') + ')</option>'; }).join('') + '</select></label>' +
      '<label>적용 월 <input type="month" data-bind="k:applyMonth" value="' + h(k.applyMonth) + '"></label>' +
      '<label>기준 Index 기간 — 시작 <input type="month" data-bind="k:base.from" value="' + h(k.base.from) + '"></label><label>끝 <input type="month" data-bind="k:base.to" value="' + h(k.base.to) + '"></label>' +
      '<label>변경 Index 기간 — 시작 <input type="month" data-bind="k:next.from" value="' + h(k.next.from) + '"></label><label>끝 <input type="month" data-bind="k:next.to" value="' + h(k.next.to) + '"></label></div>' +
      (c ? '<p class="sub">계약서의 기준 문구: ' + (c.baseRule ? '「' + h(c.baseRule) + '」' : '<span class="bad">없음 — 계약 정보에 적어 두세요</span>') + ' · 기간 값은 그 기간에 있는 달의 평균입니다. 한 달만 쓰려면 시작과 끝을 같은 달로 두세요.</p>' : '<div class="warn">이 업체는 계약 정보가 없습니다. <a href="#/vendors">업체 · 계약</a></div>') +
      '<div class="btnrow"><button class="btn sm" data-act="suggest-period">적용 월 · 주기로 기간 제안</button></div>' +
      '<h3>업체 요청단가 <small>업체가 요구한 변경 단가 — 비워 두면 비교하지 않습니다</small></h3>' +
      (ps.length ? '<div class="tablewrap" style="max-height:320px"><table class="grid" id="reqs"><tr><th>품번</th><th>품명</th><th>기준 파트단가</th><th>현재 구매단가</th><th>업체 요청단가(원/EA)</th></tr>' + ps.map(function (p) { return '<tr><td>' + h(p.partNo) + '</td><td>' + h(p.name) + '</td><td class="num">' + num(p.basePrice) + '</td><td class="num">' + num(p.curPrice) + '</td><td><input type="number" class="w-m" step="any" data-t="num" data-bind="rq:' + h(p.partNo) + '" value="' + h(k.requests[p.partNo]) + '"></td></tr>'; }).join('') + '</table></div>' : '<div class="warn">등록된 품번이 없습니다. <a href="#/parts">품번 · 사용량</a></div>') +
      '<div class="btnrow"><button class="btn primary" data-act="run-calc"' + (c && ps.length ? '' : ' disabled') + '>계산하기</button></div></section>';
    if (ui.result) out += '<section class="card" id="result"><h2>계산 결과 <small>' + h(v.name) + ' · 계약 v' + c.version + ' · 적용 ' + h(k.applyMonth) + '</small></h2>' + renderResult(ui.result, c) +
      '<h3>검토 의견 <small>보고서 「검토 의견」에 들어갑니다</small></h3><textarea data-bind="k:opinion" placeholder="예) 구리 Index 상승분만 반영, 요청단가와의 차액은 가공비 인상 근거 자료를 받아 재검토">' + h(k.opinion) + '</textarea>' +
      (ui.result.blocks.length ? '' : '<div class="btnrow"><button class="btn green" data-act="save-run">결과 승인 · 이력에 저장</button></div><p class="sub" style="margin-top:8px">저장하면 그때의 계약 버전 · 품번 값 · Index 값을 함께 묶어 둡니다. 나중에 계약이나 Index 파일이 바뀌어도 이 결과는 바뀌지 않습니다.</p>') + '</section>';
    return out;
  }
  function renderResult(res, c) {
    var out = '';
    if (res.blocks.length) return '<div class="warn" id="blocks"><b>계산을 멈췄습니다</b><ul>' + res.blocks.map(function (b) { return '<li>' + h(b.msg) + '</li>'; }).join('') + '</ul></div>';
    out += '<div class="stats"><div class="stat"><b id="r-done">' + res.done + '</b>계산 완료 품번</div><div class="stat"><b id="r-stop">' + res.stopped + '</b>확인 필요 품번(계산 안 함)</div><div class="stat"><b>' + num(res.sum.delta, 0) + '</b>연동액 합(품번당, 원)</div></div>';
    if (res.warnings.length) out += '<div class="warn" id="warns"><b>확인이 필요한 항목</b><ul>' + res.warnings.map(function (w) { return '<li>' + h(w.msg) + '</li>'; }).join('') + '</ul></div>';
    out += '<h3>원소재 · 환율 Index 변동</h3><div class="tablewrap free"><table id="r-items"><tr><th>연동 항목</th><th>적용 Index</th><th>단위</th><th>기준 Index</th><th>변경 Index</th><th>변동률</th><th>적용 변동률</th><th>계약 조건 · 알림</th></tr>' + res.items.map(function (o) {
      return '<tr class="' + (o.blocks.length ? 'stop' : '') + '"><td>' + h(o.name) + (o.spec ? ' <span class="muted">(' + h(o.spec) + ')</span>' : '') + '</td><td class="wrap">' + (o.index ? h(o.index) + (o.match.similar ? ' <span class="tag t-warn">유사 Index</span>' : '') + '<br><span class="muted" style="font-size:12px">' + h(o.source) + '</span>' : '<span class="bad">없음</span>') + '</td><td>' + h(o.unit) + '</td><td class="num">' + num(o.i0, 2) + (o.base && !o.base.fixed && o.base.used ? '<br><span class="muted" style="font-size:11.5px">' + o.base.used.length + '개월 평균</span>' : o.base && o.base.fixed ? '<br><span class="muted" style="font-size:11.5px">계약 고정값</span>' : '') + '</td><td class="num">' + num(o.i1, 2) + (o.next && o.next.used ? '<br><span class="muted" style="font-size:11.5px">' + o.next.used.length + '개월 평균</span>' : '') + '</td><td class="num ' + (o.rawRate > 0 ? 'up' : o.rawRate < 0 ? 'down' : '') + '">' + L.fmtPct(o.rawRate) + '</td><td class="num"><b>' + L.fmtPct(o.rate) + '</b></td><td class="wrap">' + h((o.steps || []).concat(o.notes).join(' / ')) + (o.blocks.length ? '<div class="bad">' + h(o.blocks.join(' / ')) + '</div>' : '') + '</td></tr>';
    }).join('') + '</table></div>';
    out += '<h3>품번별 결과 · 업체 요청단가 비교</h3><div class="tablewrap free"><table id="r-rows"><tr><th>품번</th><th>기준 파트단가</th><th>현재 구매단가</th><th>원소재 연동액</th><th>환율 연동액</th><th>전체 연동액</th><th>계약 계산 변경단가</th><th>변동률</th><th>업체 요청단가</th><th>요청 − 계산</th><th>요청 인상률</th><th>계약 계산 인상률</th><th>검토</th></tr>' + res.rows.map(function (r) {
      var q = r.request || {}, ok = r.newPrice != null;
      return '<tr class="' + (ok ? '' : 'stop') + '"><td><b>' + h(r.partNo) + '</b><br><span class="muted" style="font-size:12px">' + h(r.name) + '</span></td><td class="num">' + num(r.basePrice) + '</td><td class="num">' + num(r.curPrice) + '</td><td class="num">' + (ok ? num(r.matAmount, 2) : '') + '</td><td class="num">' + (ok ? num(r.fxAmountDelta, 2) : '') + '</td><td class="num ' + (r.total > 0 ? 'up' : r.total < 0 ? 'down' : '') + '">' + (ok ? num(r.total, 2) : '') + '</td><td class="num"><b>' + (ok ? num(r.newPrice, 2) : '<span class="bad">확인 필요</span>') + '</b></td><td class="num">' + (ok ? L.fmtPct(r.rate) : '') + '</td><td class="num">' + num(q.price) + '</td><td class="num ' + (q.gap > 0 ? 'up' : q.gap < 0 ? 'down' : '') + '">' + num(q.gap, 2) + '</td><td class="num">' + L.fmtPct(q.reqRate) + '</td><td class="num">' + L.fmtPct(q.calcRate) + '</td><td class="wrap">' + (r.blocks.length ? '<span class="bad">' + h(r.blocks.join(' / ')) + '</span>' : h(q.verdict || '')) + (r.warnings.length ? '<div class="bad">' + h(r.warnings.join(' / ')) + '</div>' : '') + '</td></tr>';
    }).join('') + '</table></div>';
    out += '<details class="box"><summary>중간 계산 과정 — 품번 · 원소재별</summary><div class="tablewrap free" style="margin-top:8px"><table id="r-lines"><tr><th>품번</th><th>연동 항목</th><th>사용량(kg)</th><th>기준 Index</th><th>변경 Index</th><th>적용 변동률</th><th>연동액(원)</th><th>계산</th></tr>' + [].concat.apply([], res.rows.map(function (r) {
      return r.lines.map(function (l) { return '<tr><td>' + h(r.partNo) + '</td><td>' + h(l.item) + '</td><td class="num">' + (l.kg === '-' ? '-' : l.kg == null ? '<span class="bad">빈 칸</span>' : num(l.kg, 3)) + '</td><td class="num">' + num(l.i0, 2) + '</td><td class="num">' + num(l.i1, 2) + '</td><td class="num">' + L.fmtPct(l.rate) + '</td><td class="num">' + num(l.amount, 2) + '</td><td class="calcline">' + h(l.calc || l.note) + '</td></tr>'; });
    })).join('') + '</table></div></details>';
    out += '<p class="sub" style="margin-top:10px">원소재 연동으로 설명되는 금액은 「전체 연동액」까지입니다. 요청단가와의 차액은 가공비 · 인건비 · 제조경비 · 물류비 · 계약상 추가 조건 · 업체 제출 원가분석 자료로 따로 검토하세요.</p>';
    return out;
  }

  /* ───────────── 보고서 · 이력 ───────────── */
  function views_reports() {
    var run = S.runs.filter(function (r) { return r.id === ui.runId; })[0];
    var tags = '{업체명} {작성일} {적용월} {계약버전} {계약기간} {적용주기} {기준기간} {변경기간} {계약수식} {계약조건} {Index파일} {품번수} {계산완료} {확인필요건수} {기준합계} {연동액합계} {변경합계} {유사Index} {검토의견} {경고}\n' +
      '표 반복(한 행의 첫 칸에 여는 꼬리표, 끝 칸에 닫는 꼬리표):\n{#items}{원소재} {Index명} {단위} {기준Index} {변경Index} {변동률} {적용변동률} {비고}{/items}\n{#rows}{품번} {품명} {기준단가} {현재단가} {사용량} {원소재연동액} {환율연동액} {전체연동액} {변경단가} {변동률} {요청단가} {차액} {의견}{/rows}\n{#details}{품번} {원소재} {사용량} {기준Index} {변경Index} {변동률} {연동액}{/details}';
    var out = '<section class="card"><h2>보고서 및 이력 관리</h2><h3 style="margin-top:6px">Word 템플릿</h3>' +
      '<p class="sub">팀장 보고용 Word 양식(.docx)에 꼬리표를 적어 올리면 표 · 글꼴 · 배치를 그대로 두고 값만 채웁니다. 올리지 않으면 기본 틀을 씁니다.</p>' +
      '<dl class="kv"><dt>지금 틀</dt><dd id="tpl-name">' + (S.template ? h(S.template.name) + ' (올린 날 ' + h(S.template.at.slice(0, 10)) + ')' : '기본 틀') + '</dd></dl>' +
      '<div class="btnrow"><button class="btn sm" data-act="up-template">Word 템플릿 등록</button><button class="btn sm" data-act="dl-template">기본 틀 내려받기</button>' + (S.template ? '<button class="btn sm danger" data-act="reset-template">기본 틀로 되돌리기</button>' : '') + '</div>' +
      '<details class="box"><summary>양식에 쓸 수 있는 꼬리표</summary><pre class="tags">' + h(tags) + '</pre><p class="sub">값이 없는 항목은 「확인 필요」로 들어갑니다.</p></details></section>';
    out += '<section class="card"><h2>계산 이력</h2>' + runTable(S.runs) + '</section>';
    if (run) {
      var others = S.runs.filter(function (r) { return r.vendorId === run.vendorId && r.id !== run.id; });
      out += '<section class="card" id="runview"><h2>' + h(run.vendorName) + ' · 계약 v' + run.contract.version + ' · 적용 ' + h(run.applyMonth) + ' <small>저장 ' + h(run.at.replace('T', ' ').slice(0, 16)) + '</small></h2>' +
        '<dl class="kv"><dt>Index 파일</dt><dd>' + h(run.indexFile) + '</dd><dt>기준 · 변경 기간</dt><dd>' + h((run.period.base.from || '계약 고정값') + ' ~ ' + (run.period.base.to || '')) + ' → ' + h(run.period.next.from + ' ~ ' + run.period.next.to) + '</dd><dt>검토 의견</dt><dd>' + (run.opinion ? h(run.opinion) : '<span class="bad">확인 필요</span>') + '</dd><dt>보고서 생성</dt><dd>' + ((run.reports || []).map(function (x) { return h(x.at.replace('T', ' ').slice(0, 16)) + ' ' + h(x.file); }).join('<br>') || '아직 없음') + '</dd></dl>' +
        '<div class="btnrow"><button class="btn green" data-act="make-docx">Word 보고서 만들기</button><button class="btn" data-act="make-xlsx">엑셀로 내려받기</button>' + confirmBtn('delrun', '이 이력 삭제', 'del-run') + '</div>' +
        '<p class="sub" style="margin-top:8px">아래가 보고서에 들어갈 내용입니다. 저장할 때의 계약 버전 · 품번 값 · Index 값 그대로이며 다시 계산하지 않습니다.</p>' + renderResult(run.result, run.contract);
      if (others.length) {
        out += '<h3>다른 이력과 비교 <small>계약 버전 · 기간이 다른 결과를 품번별로 나란히</small></h3><div class="form"><label class="wide">비교할 이력 <select data-ch="cmp"><option value="">고르기…</option>' + others.map(function (r) { return '<option value="' + r.id + '"' + (ui.cmp === r.id ? ' selected' : '') + '>' + h(r.at.slice(0, 10) + ' · v' + r.contract.version + ' · 적용 ' + r.applyMonth) + '</option>'; }).join('') + '</select></label></div>';
        var o = others.filter(function (r) { return r.id === ui.cmp; })[0];
        if (o) {
          var by = {}; o.result.rows.forEach(function (r) { by[r.partNo] = r; });
          out += '<div class="tablewrap free"><table id="cmp"><tr><th>품번</th><th>이 이력(v' + run.contract.version + ' · ' + h(run.applyMonth) + ') 변경단가</th><th>비교 이력(v' + o.contract.version + ' · ' + h(o.applyMonth) + ') 변경단가</th><th>차이</th></tr>' + run.result.rows.map(function (r) { var x = by[r.partNo], d = x && r.newPrice != null && x.newPrice != null ? r.newPrice - x.newPrice : null; return '<tr><td>' + h(r.partNo) + '</td><td class="num">' + num(r.newPrice, 2) + '</td><td class="num">' + (x ? num(x.newPrice, 2) : '<span class="muted">없음</span>') + '</td><td class="num ' + (d > 0 ? 'up' : d < 0 ? 'down' : '') + '">' + num(d, 2) + '</td></tr>'; }).join('') + '</table></div>';
        }
      }
      out += '</section>';
    }
    return out;
  }

  /* ───────────── 그리기 · 길 찾기 ───────────── */
  var views = { dash: views_dash, vendors: views_vendors, parts: views_parts, index: views_index, calc: views_calc, reports: views_reports };
  function render() {
    var y = window.scrollY;
    $('#view').innerHTML = views[ui.view]();
    [].forEach.call(document.querySelectorAll('#nav a[data-v]'), function (a) { a.classList.toggle('active', a.dataset.v === ui.view); });
    document.body.dataset.view = ui.view;
    if (render.keep) { window.scrollTo(0, y); render.keep = false; }
  }
  function rr() { render.keep = true; render(); }
  function route() { var v = (location.hash.match(/^#\/(\w+)/) || [])[1] || 'dash'; ui.view = views[v] ? v : 'dash'; ui.confirm = ''; render(); window.scrollTo(0, 0); }

  /* ───────────── 값 묶기 ───────────── */
  function conv(el) {
    var t = el.dataset.t;
    if (t === 'bool') return el.checked;
    if (t === 'num') { var s = el.value.trim(); return s === '' ? null : parseFloat(s); }
    if (t === 'tiers') return el.value.split(/[,;]/).map(function (x) { var p = x.split(':').map(function (y) { return parseFloat(y); }); return p; }).filter(function (p) { return p.length === 2 && isFinite(p[0]) && isFinite(p[1]); });
    return el.value;
  }
  function bind(el) {
    var p = el.dataset.bind.split(':'), val = conv(el), v = vendor(ui.vendorId);
    if (p[0] === 'v') v[p[1]] = val;
    else if (p[0] === 'c') ui.cDraft[p[1]] = val;
    else if (p[0] === 'ci') {
      var it = ui.cDraft.items[+p[1]]; it[p[2]] = val;
      if (p[2] === 'kind') { if (val === 'fx') { it.formula = 'fx'; it.reflect = 100; it.tiers = []; it.share = null; it.expr = ''; } else { it.formula = ''; it.fxBase = ''; } }
      if (p[2] === 'name' && !it.spec && it.kind !== 'fx' && L.isFxName(val)) { it.kind = 'fx'; it.formula = 'fx'; it.reflect = 100; return true; }
    }
    else if (p[0] === 'p') { ui.pDraft[+p[1]][p[2]] = val; if (p[2] === 'fx') { if (!val) ui.pDraft[+p[1]].fxAmount = null; return true; } }
    else if (p[0] === 'pu') ui.pDraft[+p[1]].usage[p.slice(2).join(':')] = val;
    else if (p[0] === 'rq') { if (val == null) delete ui.calc.requests[p.slice(1).join(':')]; else ui.calc.requests[p.slice(1).join(':')] = val; }
    else if (p[0] === 'k') { var path = p[1].split('.'); if (path.length === 2) ui.calc[path[0]][path[1]] = val; else ui.calc[path[0]] = val; if (p[1] !== 'opinion') ui.result = null; }
    return !!el.dataset.rr;
  }

  /* ───────────── 동작 ───────────── */
  var acts = {
    'ask-confirm': function (el) { ui.confirm = el.dataset.key; rr(); },
    'cancel-confirm': function () { ui.confirm = ''; rr(); },
    'go-vendor': function (el) { var v = S.vendors.filter(function (x) { return x.name === el.dataset.name; })[0]; if (v) ui.vendorId = v.id; },
    'load-sample': function () {
      var s = L.sample();
      s.vendors.forEach(function (v) { if (!vendor(v.id)) S.vendors.push(v); });
      s.contracts.forEach(function (c) { if (!S.contracts.some(function (x) { return x.id === c.id; })) { c.savedAt = new Date().toISOString(); S.contracts.push(c); } });
      s.parts.forEach(function (p) { if (!S.parts.some(function (x) { return x.vendorId === p.vendorId && L.key(x.partNo) === L.key(p.partNo); })) S.parts.push(p); });
      S.index = { file: s.indexFile, at: new Date().toISOString(), book: { names: s.book.names, sheets: s.book.sheets } }; seriesCache = null;
      ui.vendorId = s.vendors[0].id; ui.cDraft = null; ui.pDraft = null; ui.calc = null; ui.result = null;
      ui.sampleHint = { applyMonth: s.applyMonth, requests: s.requests };
      save(); toast('예제(가짜 업체 2곳 · 품번 6개 · 가짜 Index 통합문서)를 불러왔습니다'); render();
    },
    backup: function () { download(new Blob([JSON.stringify(S)], { type: 'application/json' }), '원소재연동제_백업_' + new Date().toISOString().slice(0, 10) + '.json'); },
    restore: function () { pickFile('.json,application/json', function (f) { f.text().then(function (t) { try { var o = JSON.parse(t); if (!o || !Array.isArray(o.vendors) || !Array.isArray(o.parts)) throw new Error('백업 파일 모양이 아닙니다'); S = Object.assign(blank(), o); seriesCache = null; ui.cDraft = ui.pDraft = ui.calc = ui.result = null; save(); toast('백업을 불러왔습니다'); render(); } catch (e) { toast('불러오지 못했습니다 — ' + e.message, true); } }); }); },
    wipe: function () { S = blank(); seriesCache = null; ui = Object.assign(ui, { vendorId: '', cDraft: null, pDraft: null, px: null, calc: null, result: null, runId: '', confirm: '' }); save(); toast('이 브라우저에 저장된 자료를 모두 지웠습니다'); render(); },

    'add-vendor': function () { var v = L.blankVendor(); v.name = '새 업체 ' + (S.vendors.length + 1); S.vendors.push(v); ui.vendorId = v.id; ui.cDraft = null; save(); render(); },
    'sel-vendor': function (el) { ui.vendorId = el.dataset.id; ui.cDraft = null; ui.confirm = ''; render(); },
    'save-vendor': function () { var v = vendor(ui.vendorId); if (!String(v.name).trim()) return toast('업체명을 넣어 주세요', true); if (S.vendors.some(function (x) { return x.id !== v.id && L.key(x.name) === L.key(v.name); })) return toast('같은 이름의 업체가 이미 있습니다', true); save(); seriesCache = null; toast('업체 정보를 저장했습니다'); rr(); },
    'del-vendor': function () { var id = ui.vendorId; if (S.runs.some(function (r) { return r.vendorId === id; })) { ui.confirm = ''; rr(); return toast('계산 이력이 있는 업체는 지울 수 없습니다(이력을 먼저 지우세요)', true); } S.vendors = S.vendors.filter(function (v) { return v.id !== id; }); S.contracts = S.contracts.filter(function (c) { return c.vendorId !== id; }); S.parts = S.parts.filter(function (p) { return p.vendorId !== id; }); ui.vendorId = ''; ui.cDraft = null; ui.confirm = ''; save(); render(); },
    'sel-contract': function (el) { loadDraft(vendor(ui.vendorId), el.dataset.id); rr(); },
    'add-item': function () { ui.cDraft.items.push(L.blankItem('')); rr(); },
    'add-fx': function () { var it = L.blankItem('환율'); it.spec = 'USD'; ui.cDraft.items.push(it); rr(); },
    'del-item': function (el) { ui.cDraft.items.splice(+el.dataset.i, 1); rr(); },
    'up-contract': function () {
      pickFile('.docx,.pdf', function (f) {
        var pdf = /\.pdf$/i.test(f.name);
        readAb(f).then(function (ab) { return pdf ? pdfText(ab) : mammoth.extractRawText({ arrayBuffer: ab }).then(function (r) { return r.value; }); }).then(function (text) {
          if (pdf && text.replace(/\s/g, '').length < 30) return toast('글자를 읽을 수 없는 PDF(스캔본)입니다 — 글자로 된 PDF · Word 를 올리거나 직접 입력해 주세요', true);
          ui.extract = L.extractContract(text); ui.cDraft.file = f.name; toast('계약서에서 후보를 찾았습니다 — 확인하고 넣어 주세요'); rr();
        }).catch(function (e) { toast('계약서 파일을 읽지 못했습니다 — ' + e.message, true); });
      });
    },
    'use-ex': function (el) { var k = el.dataset.k, f = ui.extract.fields[k]; if (k === 'cap' || k === 'floor' || k === 'minChange') { if (!ui.cDraft.items.length) return toast('연동 항목을 먼저 추가하세요', true); ui.cDraft.items.forEach(function (it) { if (it.kind !== 'fx') it[k] = f.value; }); toast('원소재 항목 전체에 넣었습니다 — 항목마다 다르면 고쳐 주세요'); } else ui.cDraft[k] = f.value; rr(); },
    'use-mat': function (el) { var m = ui.extract.materials[+el.dataset.i]; if (ui.cDraft.items.some(function (it) { return L.key(it.name) === L.key(m.name); })) return toast('이미 있는 항목입니다', true); var it = L.blankItem(m.name); if (m.share != null && it.kind !== 'fx') it.share = m.share; ui.cDraft.items.push(it); rr(); },
    'use-formula': function (el) { ui.cDraft.formulaText = ui.extract.formulaLines[+el.dataset.i]; rr(); },
    'save-contract': function () { saveContract(false); },
    'save-contract-new': function () { saveContract(true); },

    'add-part': function () { var p = L.blankPart(ui.vendorId); ui.pCols.forEach(function (m) { p.usage[m] = null; }); ui.pDraft.push(p); rr(); },
    'del-part': function (el) { ui.pDraft.splice(+el.dataset.i, 1); rr(); },
    'add-col': function () { var n = $('#newcol').value.trim(); if (!n) return toast('원소재 이름을 넣어 주세요', true); if (L.isFxName(n)) return toast('환율은 사용량(kg) 열로 만들 수 없습니다 — 「환율 적용」 칸을 쓰세요', true); if (ui.pCols.some(function (m) { return L.key(m) === L.key(n); })) return toast('이미 있는 열입니다', true); ui.pCols.push(n); rr(); },
    'del-col': function (el) { var m = ui.pCols[+el.dataset.k]; ui.pCols.splice(+el.dataset.k, 1); ui.pDraft.forEach(function (p) { delete p.usage[m]; }); rr(); },
    'ren-col': function (el) { ui.renCol = +el.dataset.k; var th = el.parentNode, old = ui.pCols[ui.renCol]; th.innerHTML = '<input type="text" class="w-m" id="rencol" value="' + h(old) + '"> <button class="x" data-act="ren-col-ok">확인</button>'; $('#rencol').focus(); },
    'ren-col-ok': function () { var n = $('#rencol').value.trim(), old = ui.pCols[ui.renCol]; if (!n || L.isFxName(n) || ui.pCols.some(function (m, i) { return i !== ui.renCol && L.key(m) === L.key(n); })) return toast('쓸 수 없는 이름입니다', true); ui.pCols[ui.renCol] = n; ui.pDraft.forEach(function (p) { if (old in p.usage) { p.usage[n] = p.usage[old]; if (n !== old) delete p.usage[old]; } }); rr(); },
    'reload-parts': function () { ui.pDraft = null; rr(); },
    'save-parts': function () {
      var vid = ui.vendorId;
      ui.pDraft.forEach(function (p) { p.vendorId = vid; p.partNo = String(p.partNo).trim(); ui.pCols.forEach(function (m) { if (!(m in p.usage)) p.usage[m] = null; }); });
      S.parts = S.parts.filter(function (p) { return p.vendorId !== vid; }).concat(ui.pDraft.map(L.clone));
      save(); toast('품번 ' + ui.pDraft.length + '개를 저장했습니다'); ui.pDraft = null; rr();
    },
    'up-parts': function () { pickFile('.xlsx,.xls,.csv', function (f) { readBook(f).then(function (b) { var sh = b.names[0]; ui.px = { file: f.name, book: b, sheets: b.names, sheet: sh, insp: L.inspectParts(b.sheets[sh]) }; pxParse(); rr(); }).catch(function (e) { toast('엑셀을 읽지 못했습니다 — ' + e.message, true); }); }); },
    'px-cancel': function () { ui.px = null; rr(); },
    'px-approve': function () {
      var px = ui.px, r = px.res, vid = ui.vendorId, okBox = $('#px-ok');
      if (r.counts.update && !(okBox && okBox.checked)) return toast('기존 품번을 덮어쓰는 데 동의해 주세요', true);
      var n = 0;
      r.rows.forEach(function (x) {
        if (x.status === 'error') return;
        var old = S.parts.filter(function (p) { return p.vendorId === vid && L.key(p.partNo) === L.key(x.partNo); })[0];
        var p = old || L.blankPart(vid);
        p.partNo = x.partNo; if (x.name || !old) p.name = x.name; p.basePrice = x.basePrice;
        if (px.insp.columns.some(function (c) { return c.role === 'curPrice'; })) p.curPrice = x.curPrice;
        Object.keys(x.usage).forEach(function (m) { p.usage[m] = x.usage[m]; });
        if (px.insp.columns.some(function (c) { return c.role === 'fx' || c.role === 'fxAmount'; })) { p.fx = x.fx; p.fxAmount = x.fxAmount; }
        p.src = { file: px.file, at: new Date().toISOString() };
        if (!old) S.parts.push(p); n++;
      });
      save(); ui.px = null; ui.pDraft = null; toast(n + '개 품번을 저장했습니다(' + px.file + ')'); rr();
    },
    'dl-parts-template': function () { var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(L.partsTemplate(materialsOf(ui.vendorId))), '품번별원소재'); XLSX.writeFile(wb, '품번별_원소재_등록양식.xlsx'); },
    'dl-parts': function () {
      var v = vendor(ui.vendorId), cols = materialsOf(v.id), aoa = [['품번', '품명', '기준 파트단가', '현재 구매단가'].concat(cols.map(function (m) { return m + '(kg)'; })).concat(['환율 적용', '환율 대상 금액'])];
      partsOf(v.id).forEach(function (p) { aoa.push([p.partNo, p.name, p.basePrice, p.curPrice == null ? '' : p.curPrice].concat(cols.map(function (m) { return p.usage[m] == null ? '' : p.usage[m]; })).concat([p.fx ? 'Y' : '', p.fxAmount == null ? '' : p.fxAmount])); });
      var wb = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), '품번별원소재'); XLSX.writeFile(wb, v.name + '_품번별원소재.xlsx');
    },

    'up-index': function () { pickFile('.xlsx,.xls,.csv', function (f) { readBook(f).then(function (b) { S.index = { file: f.name, at: new Date().toISOString(), book: b }; seriesCache = null; ui.result = null; save(); toast('Index 엑셀을 읽었습니다 — ' + series().length + '개 계열'); rr(); }).catch(function (e) { toast('엑셀을 읽지 못했습니다 — ' + e.message, true); }); }); },
    'dl-sample-index': function () { var B = L.sampleBook(), wb = XLSX.utils.book_new(); B.names.forEach(function (n) { var ws = XLSX.utils.aoa_to_sheet(B.sheets[n]); if (B.merges && B.merges[n]) ws['!merges'] = B.merges[n]; XLSX.utils.book_append_sheet(wb, ws, n); }); XLSX.writeFile(wb, '원자재가격동향_예제(가상).xlsx'); },
    'del-index': function () {
      var n = 0; S.contracts.forEach(function (c) { c.items.forEach(function (it) { if (it.indexName) { it.indexName = ''; it.indexStatus = ''; n++; } }); });
      S.approvals.unshift({ at: new Date().toISOString(), vendorId: '', item: '(전체)', index: '', type: 'Index 파일 삭제 — 승인 연결 ' + n + '건 해제', file: S.index.file });
      S.index = null; seriesCache = null; ui.cDraft = null; ui.result = null; ui.confirm = ''; save(); toast('Index 파일을 지웠습니다. 업체 · 품번 · 사용량 · 계산 이력은 그대로입니다'); render();
    },
    'approve-index': function (el) { approve(el.dataset.item, el.dataset.name, el.dataset.st); },
    'approve-pick': function (el) { var sel = document.querySelector('select[data-pick="' + el.dataset.item + '"]'); if (!sel.value) return toast('Index 를 먼저 고르세요', true); approve(el.dataset.item, sel.value, 'similar'); },
    unapprove: function (el) { approve(el.dataset.item, '', ''); },

    'suggest-period': function () { var k = ui.calc, c = S.contracts.filter(function (x) { return x.id === k.contractId; })[0]; if (!k.applyMonth) return toast('적용 월을 먼저 넣어 주세요', true); var p = L.suggestPeriod(k.applyMonth, c ? c.cycle : 'quarter'); k.base = p.base; k.next = p.next; ui.result = null; toast('직전 ' + (c && c.cycle === 'half' ? '반기' : c && c.cycle === 'year' ? '1년' : '분기') + '와 그 앞 기간으로 제안했습니다 — 계약서 기준일과 맞는지 확인하세요'); rr(); },
    'run-calc': function () {
      var k = ui.calc, v = vendor(k.vendorId), c = S.contracts.filter(function (x) { return x.id === k.contractId; })[0];
      ui.result = L.calcRun({ vendor: v, contract: c, parts: partsOf(v.id), series: series(), period: { base: k.base, next: k.next }, requests: k.requests, applyMonth: k.applyMonth });
      rr(); setTimeout(function () { var r = $('#result'); if (r) r.scrollIntoView({ block: 'start' }); }, 0);
    },
    'save-run': function () {
      var k = ui.calc, v = vendor(k.vendorId), c = S.contracts.filter(function (x) { return x.id === k.contractId; })[0];
      var run = { id: L.uid('r'), at: new Date().toISOString(), vendorId: v.id, vendorName: v.name, contract: L.clone(c), parts: partsOf(v.id).map(L.clone), period: L.clone({ base: k.base, next: k.next }), applyMonth: k.applyMonth, requests: L.clone(k.requests), indexFile: S.index ? S.index.file : '', result: L.clone(ui.result), opinion: k.opinion, reports: [] };
      S.runs.unshift(run); save(); ui.runId = run.id; ui.result = null; toast('결과를 이력에 저장했습니다'); location.hash = '#/reports';
    },
    'open-run': function (el) { ui.runId = el.dataset.id; ui.cmp = ''; if (ui.view === 'reports') { render(); setTimeout(function () { var r = $('#runview'); if (r) r.scrollIntoView(); }, 0); } },
    'del-run': function () { S.runs = S.runs.filter(function (r) { return r.id !== ui.runId; }); ui.runId = ''; ui.confirm = ''; save(); render(); },
    'make-xlsx': function () { var run = curRun(), sh = L.exportSheets(run), wb = XLSX.utils.book_new(); Object.keys(sh).forEach(function (n) { XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sh[n]), n); }); XLSX.writeFile(wb, '원소재연동제_' + run.vendorName + '_' + run.applyMonth + '.xlsx'); },
    'make-docx': function () {
      var run = curRun();
      try {
        var zip = new PizZip(b64ToAb(S.template ? S.template.b64 : window.REPORT_TEMPLATE_B64));
        var doc = new window.docxtemplater(zip, { paragraphLoop: true, linebreaks: true, nullGetter: function () { return L.NA; } });
        doc.render(L.reportData(run));
        var name = '원소재연동제_검토보고_' + run.vendorName + '_' + run.applyMonth + '.docx';
        var blob = doc.getZip().generate({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
        window.__lastDocx = { name: name, size: blob.size };
        download(blob, name);
        run.reports = run.reports || []; run.reports.push({ at: new Date().toISOString(), file: name, template: S.template ? S.template.name : '기본 틀' }); save(); toast('Word 보고서를 만들었습니다 — ' + name); rr();
      } catch (e) {
        var why = e.properties && e.properties.errors ? e.properties.errors.map(function (x) { return x.properties && x.properties.explanation || x.message; }).join(' / ') : e.message;
        toast('보고서를 만들지 못했습니다 — 템플릿의 꼬리표를 확인해 주세요: ' + why, true);
      }
    },
    'up-template': function () { pickFile('.docx', function (f) { readAb(f).then(function (ab) { try { new PizZip(ab); } catch (e) { return toast('Word(.docx) 파일이 아닙니다', true); } S.template = { name: f.name, at: new Date().toISOString(), b64: abToB64(ab) }; save(); toast('템플릿을 등록했습니다'); rr(); }); }); },
    'dl-template': function () { download(new Blob([b64ToAb(window.REPORT_TEMPLATE_B64)], { type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' }), '원소재연동제_보고서_기본틀.docx'); },
    'reset-template': function () { S.template = null; save(); rr(); }
  };
  function curRun() { return S.runs.filter(function (r) { return r.id === ui.runId; })[0]; }
  function saveContract(asNew) {
    var c = ui.cDraft, v = vendor(ui.vendorId), list = contractsOf(v.id);
    c.items.forEach(function (it) { it.name = String(it.name).trim(); if (it.kind === 'fx') { it.reflect = 100; it.formula = 'fx'; } });
    if (c.items.some(function (it) { return !it.name; })) return toast('이름이 빈 연동 항목이 있습니다', true);
    if (c.from && c.to && c.from > c.to) return toast('적용 시작월이 종료월보다 늦습니다', true);
    var o = L.clone(c); delete o.isNew; o.savedAt = new Date().toISOString();
    if (c.isNew || asNew) { o.id = L.uid('c'); o.version = (list[0] ? list[0].version : 0) + 1; S.contracts.push(o); }
    else { var i = S.contracts.findIndex(function (x) { return x.id === c.id; }); S.contracts[i] = o; }
    save(); loadDraft(v, o.id); ui.calc = null; ui.result = null;
    toast('계약 v' + o.version + ' 을 저장했습니다' + (L.contractProblems(o).length ? ' — 계산 전에 확인할 항목이 있습니다' : '')); rr();
  }
  function approve(itemId, name, st) {
    var v = vendor(ui.vendorId), c = contractsOf(v.id)[0], it = c.items.filter(function (x) { return x.id === itemId; })[0];
    if (st === 'similar' && !c.allowSubstitute) return toast('이 계약은 대체 Index 사용이 허용돼 있지 않습니다 — 계약 정보에서 먼저 확인해 주세요', true);
    it.indexName = name; it.indexStatus = name ? st : '';
    S.approvals.unshift({ at: new Date().toISOString(), vendorId: v.id, item: it.name, index: name, type: !name ? '승인 취소' : st === 'similar' ? '유사 Index 승인' : '정확 일치 선택', file: S.index ? S.index.file : '' });
    if (ui.cDraft && ui.cDraft.id === c.id) ui.cDraft = null;
    ui.result = null; save(); toast(name ? '승인했습니다 — ' + it.name + ' → ' + name : '승인을 취소했습니다'); rr();
  }

  var changes = {
    'parts-vendor': function (el) { ui.vendorId = el.value; ui.pDraft = null; ui.px = null; rr(); },
    'index-vendor': function (el) { ui.vendorId = el.value; rr(); },
    'calc-vendor': function (el) { ui.vendorId = el.value; ui.calc = null; rr(); },
    'fx-toggle': function (el) { S.fx = el.checked; seriesCache = null; save(); rr(); },
    'px-role': function (el) { var c = ui.px.insp.columns[+el.dataset.i]; c.role = el.value; if (c.role === 'usage' && !c.material) c.material = c.head.replace(/\s*[\(\[（].*$/, ''); ui.px.insp.ok = true; pxParse(); rr(); },
    'px-mat': function (el) { ui.px.insp.columns[+el.dataset.i].material = el.value.trim(); pxParse(); rr(); },
    'px-sheet': function (el) { ui.px.sheet = el.value; ui.px.insp = L.inspectParts(ui.px.book.sheets[el.value]); pxParse(); rr(); },
    cmp: function (el) { ui.cmp = el.value; rr(); }
  };

  document.addEventListener('click', function (e) {
    var el = e.target.closest('[data-act]'); if (!el || el.disabled) return;
    var a = acts[el.dataset.act]; if (!a) return;
    if (el.tagName !== 'A') e.preventDefault();
    a(el, e);
  });
  document.addEventListener('change', function (e) {
    var el = e.target;
    if (el.dataset.bind) { var need = bind(el); if (ui.view === 'parts' || ui.view === 'vendors' && /^ci:/.test(el.dataset.bind) || need) rr(); return; }
    if (el.dataset.ch && changes[el.dataset.ch]) changes[el.dataset.ch](el);
  });
  window.addEventListener('hashchange', route);

  openDb().then(function () { return kvGet('state'); }).then(function (raw) {
    try { S = raw ? Object.assign(blank(), JSON.parse(raw)) : blank(); } catch (e) { S = blank(); }
    window.__app = { state: function () { return S; }, ui: ui };
    route();
    // 예제를 불러온 직후 계산 화면에 적용 월 · 요청단가를 채워 준다
    var orig = initCalc;
    initCalc = function (v) { orig(v); if (/^v-sample/.test(v.id)) { if (!ui.sampleHint) { var sm = L.sample(); ui.sampleHint = { applyMonth: sm.applyMonth, requests: sm.requests }; } var k = ui.calc, c = contractsOf(v.id)[0]; k.applyMonth = ui.sampleHint.applyMonth; var p = L.suggestPeriod(k.applyMonth, c.cycle); k.base = p.base; k.next = p.next; partsOf(v.id).forEach(function (pp) { if (ui.sampleHint.requests[pp.partNo] != null) k.requests[pp.partNo] = ui.sampleHint.requests[pp.partNo]; }); } };
  });
})();
