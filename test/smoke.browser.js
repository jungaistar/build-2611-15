/* 브라우저에 실제로 띄워 처음부터 끝까지 한 번 돌린다 — node test/smoke.browser.js
 * 예제 불러오기 → 계약 확인 → 품번 직접 입력 · 엑셀 대량 등록 → Index 승인 → 계산 → 이력 저장 → Word · 엑셀 → 새로고침 뒤에도 남는지.
 * playwright 가 없으면 건너뛴다. */
'use strict';
var http = require('http'), fs = require('fs'), path = require('path'), os = require('os');
var chromium;
try { chromium = require('playwright').chromium; }
catch (e) { try { chromium = require('playwright-core').chromium; } catch (e2) { console.log('playwright 가 없어 화면 연기 테스트를 건너뜁니다.'); process.exit(0); } }
var XLSX = require('../lib/xlsx.full.min.js');
var ROOT = path.join(__dirname, '..');
var MIME = { '.html': 'text/html; charset=utf-8', '.js': 'text/javascript', '.css': 'text/css', '.svg': 'image/svg+xml', '.png': 'image/png', '.docx': 'application/octet-stream' };
var srv = http.createServer(function (q, s) {
  var p = path.join(ROOT, decodeURIComponent(q.url.split('?')[0]).replace(/\/$/, '/index.html'));
  fs.readFile(p, function (e, b) { if (e) { s.writeHead(404); return s.end(); } s.writeHead(200, { 'Content-Type': MIME[path.extname(p)] || 'application/octet-stream' }); s.end(b); });
});
var passed = 0, failed = 0;
function ok(c, label, d) { if (process.env.SMOKE_V) console.log((c ? '  o ' : '  x ') + label); if (c) passed++; else { failed++; console.log('  X ' + label + (d != null ? '  ' + d : '')); } }
var count = function (pg, sel) { return pg.$$eval(sel, function (t) { return t.length; }); };
srv.listen(0, async function () {
  var base = 'http://127.0.0.1:' + srv.address().port + '/';
  var br = await chromium.launch(process.env.PW_CHANNEL ? { channel: process.env.PW_CHANNEL } : {});
  var ctx = await br.newContext({ viewport: { width: 1360, height: 900 }, acceptDownloads: true });
  var pg = await ctx.newPage();
  var errs = []; pg.on('pageerror', function (e) { errs.push(e.message); }); pg.on('console', function (m) { if (m.type() === 'error') errs.push('console: ' + m.text()); });
  var tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'b15-smoke-'));
  async function go(v) { await pg.evaluate(function (x) { location.hash = '#/' + x; }, v); await pg.waitForFunction(function (x) { return document.body.dataset.view === x; }, v); }
  async function upload(act, file) { var fc = await Promise.all([pg.waitForEvent('filechooser'), pg.click('[data-act=' + act + ']')]); await fc[0].setFiles(file); }
  try {
    await pg.goto(base);
    await pg.waitForSelector('#st-vendors');
    ok(await pg.textContent('#st-vendors') === '0', '처음엔 업체 0');
    await pg.click('[data-act=load-sample]');
    ok(await pg.textContent('#st-vendors') === '2' && await pg.textContent('#st-parts') === '6', '예제: 업체 2 · 품번 6');

    // 업체 · 계약
    await go('vendors');
    ok(await count(pg, '#items tr') === 5, '가상정밀 연동 항목 4개(철강 · 구리 · 고무 · 환율)');
    ok(/100%/.test(await pg.textContent('#items td.fxna')), '환율 적용 비중은 100% 로 고정(입력 칸 없음)');
    ok(/계산할 수 있는 상태/.test(await pg.textContent('#c-probs')), '예제 계약은 계산 가능 상태');
    await pg.selectOption('#items tr:nth-child(2) select[data-bind="ci:0:formula"]', '');
    ok(/수식이 정해지지 않았습니다/.test(await pg.textContent('#c-probs')), '수식 미정이면 계산하지 않는다고 알린다');
    await pg.selectOption('#items tr:nth-child(2) select[data-bind="ci:0:formula"]', 'weight');
    // 계약서 Word 올려 후보 찾기
    var docx = path.join(tmp, 'contract.docx');
    fs.writeFileSync(docx, makeDocx(['납품대금 연동 특별약정서(별첨)', '적용 기간은 2027년 1월 1일부터 2027년 12월 31일까지로 한다.', '구분', '기재사항', '1. 물품등의 명칭 지게차 마스트 브래킷', '2. 주요 원재료 또는 주요 에너지 철강 / 구리 / 환율', '3. 원재료 가격 또는 에너지 비용의 기준지표 구매기획팀 배포 Index', '4. 원재료 기준가격 기준시점 : 직전 조정 분기 평균', '비교시점 : 조정일 직전 분기 평균', '5. 조정요건 기준시점 대비 ±5% 이상 변동', '6. 조정주기 분기', '7. 조정일 매 분기 첫 영업일', '8. 조정대금 반영시점 조정일이 속한 달 납품분부터', '9. 납품대금 연동산식 변경 단가 = 기준 단가 + 사용량 × (비교시점 가격 − 기준시점 가격)', '10. 기타사항 상한 99% 는 읽지 않는다']));
    await upload('up-contract', docx);
    await pg.waitForSelector('#extract');
    var ext = await pg.textContent('#extract');
    ok(/표준 약정서/.test(ext) && /2027-01/.test(ext) && /분기/.test(ext) && /변경 단가 = 기준 단가/.test(ext) && /조정요건[\s\S]*5/.test(ext) && /철강 추가/.test(ext) && /환율 추가/.test(ext) && !/99/.test(ext), '표준 약정서 1~9번 항목 · 기간 · 원재료 후보를 찾는다(10번 뒤는 읽지 않음)', ext.slice(0, 200));
    await pg.click('[data-act=use-ex][data-k=trigger]');
    ok(await pg.inputValue('input[data-bind="c:trigger"]') === '5', '「넣기」를 눌러야 들어간다');
    ok(await pg.inputValue('input[data-bind="c:from"]') === '2026-01', '후보는 자동으로 넣지 않는다(담당자 확인)');
    await go('dash'); await go('vendors');

    // 품번 — 직접 입력
    await go('parts');
    ok(await count(pg, '#parts tr') === 5, '가상정밀 품번 4개');
    ok(/구리\(kg\)/.test(await pg.textContent('#parts tr:first-child')) && await count(pg, '#parts tr:nth-child(2) td.fxna') === 1, '원소재 열은 계약 · 등록값에서, 환율은 사용량 칸이 없다');
    await pg.click('[data-act=add-part]');
    ok(/품번이 비어/.test(await pg.textContent('#parts tr:last-child')) && await pg.isDisabled('[data-act=save-parts]'), '빈 행은 오류로 짚고 저장을 막는다');
    await pg.fill('#parts tr:last-child input[data-bind="p:4:partNo"]', 'A-1005'); await pg.dispatchEvent('#parts tr:last-child input[data-bind="p:4:partNo"]', 'change');
    await pg.fill('#parts tr:last-child input[data-bind="p:4:basePrice"]', '50000'); await pg.dispatchEvent('#parts tr:last-child input[data-bind="p:4:basePrice"]', 'change');
    await pg.fill('#newcol', '환율'); await pg.click('[data-act=add-col]');
    ok(!/환율\(kg\) <button/.test(await pg.innerHTML('#parts tr:first-child')) && await count(pg, '#parts tr:first-child th') === 12, '환율은 사용량 열로 추가할 수 없다');
    await pg.fill('#newcol', '알루미늄'); await pg.click('[data-act=add-col]');
    ok(/알루미늄\(kg\)/.test(await pg.textContent('#parts tr:first-child')), '원소재 열 추가');
    await pg.click('[data-act=save-parts]');
    ok((await pg.evaluate(function () { return window.__app.state().parts.length; })) === 7, '직접 입력 저장(7개)');

    // 품번 — 엑셀 대량 등록(열 순서 바꿈 · 오류 행 · 기존 품번)
    var px = path.join(tmp, 'parts.xlsx'), wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([['구리(kg)', '품번', '철강(kg)', '기준 파트단가', '고무(kg)'], [1, 'A-1001', 26, 101000, 2], [0, 'B-9001', 10, 70000, ''], ['x', 'B-9002', 5, 60000, 1], [2, 'B-9003', -3, 60000, 1], [2, '', 3, 60000, 1]]), '품번');
    fs.writeFileSync(px, XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' }));
    await upload('up-parts', px);
    await pg.waitForSelector('#pxprev');
    ok(await pg.textContent('#px-new') === '1' && await pg.textContent('#px-upd') === '1' && await pg.textContent('#px-err') === '3', '엑셀 미리보기: 신규 1 · 기존 1 · 오류 3');
    var prev = await pg.textContent('#pxprev');
    ok(/숫자가 아닙니다\(x\)/.test(prev) && /음수/.test(prev) && /품번이 비어/.test(prev) && /빈 칸/.test(prev), '오류 행 · 열을 짚고 빈 칸을 0 과 구분');
    await pg.click('[data-act=px-approve]');
    ok(await count(pg, '#pxprev') === 1, '덮어쓰기 동의 없이는 저장하지 않는다');
    await pg.check('#px-ok'); await pg.click('[data-act=px-approve]');
    var st = await pg.evaluate(function () { var s = window.__app.state(); var a = s.parts.filter(function (p) { return p.partNo === 'A-1001'; })[0], b = s.parts.filter(function (p) { return p.partNo === 'B-9001'; })[0]; return [s.parts.length, a.basePrice, a.usage['철강'], a.curPrice, b.usage['고무'], b.usage['구리']]; });
    ok(st[0] === 8 && st[1] === 101000 && st[2] === 26 && st[3] === 101500 && st[4] === null && st[5] === 0, '승인 저장: 기존은 덮어쓰고(현재 단가 유지) 신규는 추가, 오류 행은 제외', JSON.stringify(st));

    // Index
    await go('index');
    ok(/가상/.test(await pg.textContent('#index-file')) && await count(pg, '#sheets tr') === 7, 'Sheet 구조 분석(6장)');
    ok(/가상정밀\(주\)/.test(await pg.textContent('#sheets')), '업체 이름과 같은 시트를 알아본다');
    ok(await count(pg, '.mcard') === 4 && await count(pg, '.mcard.st-exact') === 4, '가상정밀: 4개 항목 모두 정확 일치');
    ok(await count(pg, '.mcard[data-item="철강"] .mhead .t-info') === 1 && await count(pg, '.mcard[data-item="구리"] .mhead .t-info') === 0 && /자동으로 쓰지 않았습니다/.test(await pg.textContent('.mcard[data-item="구리"]')), '철강은 업체 시트에서, 구리는 업체 시트에 없어 다른 시트에서');
    await pg.selectOption('select[data-ch=index-vendor]', 'v-sample-2');
    ok(await count(pg, '.mcard.st-similar') === 1 && await count(pg, '.mcard.st-similar table.cands tr') >= 3, '예제주물 주철: 유사 Index 후보와 점수');

    // 계산 — 승인 전에는 멈춘다
    await go('calc');
    await pg.click('[data-act=run-calc]');
    await pg.waitForSelector('#r-rows');
    ok(await pg.textContent('#r-done') === '0' && await pg.textContent('#r-stop') === '2' && /승인해야 계산/.test(await pg.textContent('#result')), '유사 Index 승인 전에는 계산하지 않는다');
    await go('index');
    await pg.click('.mcard.st-similar [data-act=approve-index] >> nth=0');
    ok(await count(pg, '.mcard.st-approved') === 1 && /유사 Index 승인됨/.test(await pg.textContent('.mcard.st-approved')), '유사 Index 승인');
    await go('calc');
    await pg.click('[data-act=run-calc]'); await pg.waitForSelector('#r-rows');
    ok(await pg.textContent('#r-done') === '2' && /유사 Index/.test(await pg.textContent('#r-items')), '승인 뒤 계산 · 유사 Index 표시');
    ok(/비중 45%/.test(await pg.textContent('#r-lines')), '중간 계산 과정(비중 연동식)');
    await pg.click('[data-act=save-run]');
    await pg.waitForFunction(function () { return document.body.dataset.view === 'reports'; });

    // 가상정밀 계산
    await go('calc');
    await pg.selectOption('select[data-ch=calc-vendor]', 'v-sample-1');
    ok(await pg.inputValue('input[data-bind="rq:A-1001"]') === '104000', '예제 요청단가');
    await pg.click('[data-act=run-calc]'); await pg.waitForSelector('#r-rows');
    var rows = await pg.textContent('#r-rows');
    ok(/A-1004[\s\S]*구리 사용량\(kg\)이 비어/.test(rows), '사용량이 빈 품번은 계산하지 않고 이유를 보여 준다');
    ok(/요청이 계약 계산보다 높음/.test(rows) && /환율 대상 금액/.test(await pg.textContent('#r-lines')), '요청단가 비교 · 환율 별도 계산');
    await pg.fill('textarea[data-bind="k:opinion"]', '구리 상승분만 반영'); await pg.dispatchEvent('textarea[data-bind="k:opinion"]', 'change');
    await pg.click('[data-act=save-run]');
    await pg.waitForSelector('#runview');
    ok(/구리 상승분만 반영/.test(await pg.textContent('#runview')), '이력에 검토 의견과 함께 저장');

    // 보고서
    var dl = await Promise.all([pg.waitForEvent('download'), pg.click('[data-act=make-docx]')]);
    var f = path.join(tmp, 'r.docx'); await dl[0].saveAs(f);
    var info = await pg.evaluate(function () { return window.__lastDocx; });
    ok(/\.docx$/.test(dl[0].suggestedFilename()) && fs.statSync(f).size > 20000 && info.size === fs.statSync(f).size, 'Word 보고서 내려받기', fs.statSync(f).size);
    var xml = unzipText(f);
    ok(/가상정밀\(주\)/.test(xml) && /A-1003/.test(xml) && /구리 상승분만 반영/.test(xml) && !/\{[#/]?[가-힣A-Za-z]+\}/.test(xml), '보고서에 값이 채워지고 꼬리표가 남지 않는다');
    var dx = await Promise.all([pg.waitForEvent('download'), pg.click('[data-act=make-xlsx]')]);
    var fx = path.join(tmp, 'r.xlsx'); await dx[0].saveAs(fx);
    ok(XLSX.read(fs.readFileSync(fx), { type: 'buffer' }).SheetNames.join() === '품번별결과,계산내역,Index근거,검증', '엑셀 4시트');

    // 가로형 Index(원소재가 행 · 날짜가 열)도 읽는다
    var wx = path.join(tmp, 'wide.xlsx'), wb2 = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb2, XLSX.utils.aoa_to_sheet([['원소재 이름', '규격', '기준', '단위', '출처', '2026-01', '2026-02', '2026-03', '2026-04', '2026-05', '2026-06'], ['철강 후판', 'SS400', '유통가', 'KRW/kg', '가상', 1000, 1000, 1000, 1100, 1100, 1100], ['구리', '전기동', '', 'KRW/kg', '가상', 13000, 13000, 13000, 13650, 13650, 13650], ['고무', '', '', 'KRW/kg', '가상', 2000, 2000, 2000, 2000, 2000, 2000], ['환율 USD', '', '', 'KRW/USD', '가상', 1400, 1400, 1400, 1372, 1372, 1372]]), 'Index');
    fs.writeFileSync(wx, XLSX.write(wb2, { type: 'buffer', bookType: 'xlsx' }));
    await go('index'); await upload('up-index', wx);
    await pg.waitForFunction(function () { return /wide\.xlsx/.test(document.querySelector('#index-file').textContent); });
    ok(/가로형\(원소재가 행/.test(await pg.textContent('#sheets')), '가로형 Index 엑셀을 알아본다');
    await pg.selectOption('select[data-ch=index-vendor]', 'v-sample-1');
    ok(await count(pg, '.mcard.st-exact') === 4, '가로형 Index 에서도 4개 항목 정확 일치');
    await go('calc'); await pg.selectOption('select[data-ch=calc-vendor]', 'v-sample-1');
    await pg.click('[data-act=run-calc]'); await pg.waitForSelector('#r-rows');
    ok(/104,(?:150|15\d)/.test((await pg.textContent('#r-rows')).replace(/\s/g, '')) || /\+10\.00%/.test(await pg.textContent('#r-items')), '가로형 Index 로 계산(철강 +10%)');
    await go('index'); await pg.click('[data-act=ask-confirm][data-key=delidx]'); await pg.click('[data-act=del-index]');
    ok(await count(pg, '#index-file') === 0 && (await pg.evaluate(function () { var s = window.__app.state(); return [s.parts.length, s.runs.length].join(); })) === '8,2', 'Index 파일을 지워도 업체 · 품번 · 이력은 남는다');
    await go('dash'); await pg.click('[data-act=load-sample]');

    // 계약을 고치면 새 버전 — 과거 이력은 그대로
    await go('vendors');
    ok(await count(pg, '[data-act=save-contract]') === 0 && await count(pg, '[data-act=save-contract-new]') === 1, '계산에 쓰인 계약 버전은 새 버전으로만 저장');
    await pg.click('[data-act=save-contract-new]');
    ok((await pg.evaluate(function () { var s = window.__app.state(); return [s.contracts.filter(function (c) { return c.vendorId === 'v-sample-1'; }).length, s.runs[0].contract.version]; })).join() === '2,1', '새 계약 버전 v2, 이력은 v1 그대로');

    // 새로고침 뒤에도 남는다
    await pg.reload(); await pg.waitForSelector('body[data-view]'); await go('dash');
    ok(await pg.textContent('#st-vendors') === '2' && await pg.textContent('#st-parts') === '8', '새로고침 뒤에도 자료가 남는다(IndexedDB)');
    await go('reports');
    ok(await count(pg, '#view table tr') >= 3, '이력 2건');

    // 좁은 화면에서 가로로 넘치지 않는다
    await pg.setViewportSize({ width: 390, height: 800 });
    for (var v of ['dash', 'vendors', 'parts', 'index', 'calc', 'reports']) {
      await go(v);
      var over = await pg.evaluate(function () { return document.documentElement.scrollWidth - document.documentElement.clientWidth; });
      ok(over <= 1, '390px 에서 가로 넘침 없음 — ' + v, over);
    }
    for (var page of ['guide.html', 'plan.html']) { var r = await pg.goto(base + page); ok(r.status() === 200, page + ' 열림'); }
    ok(errs.length === 0, '페이지 오류 없음', errs.join(' | '));
  } catch (e) { failed++; console.log('  X 중단: ' + e.message.split('\n')[0]); if (errs.length) console.log('    ' + errs.join(' | ')); }
  await br.close(); srv.close();
  console.log('화면 연기 테스트: 통과 ' + passed + ' · 실패 ' + failed);
  process.exit(failed ? 1 : 0);
});

/* 시험용 최소 docx(문단만) — zip 은 lib/xlsx 에 든 CFB 로 만든다 */
function makeDocx(paras) {
  var CFB = XLSX.CFB, z = CFB.utils.cfb_new({ type: 'zip' });
  var esc = function (s) { return s.replace(/&/g, '&amp;').replace(/</g, '&lt;'); };
  CFB.utils.cfb_add(z, '[Content_Types].xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>'));
  CFB.utils.cfb_add(z, '_rels/.rels', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'));
  CFB.utils.cfb_add(z, 'word/document.xml', Buffer.from('<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' + paras.map(function (p) { return '<w:p><w:r><w:t xml:space="preserve">' + esc(p) + '</w:t></w:r></w:p>'; }).join('') + '</w:body></w:document>'));
  return CFB.write(z, { fileType: 'zip', type: 'buffer' });
}
function unzipText(file) {
  var z = XLSX.CFB.read(fs.readFileSync(file), { type: 'buffer' });
  var i = z.FullPaths.findIndex(function (p) { return /word\/document\.xml$/.test(p); });
  return Buffer.from(z.FileIndex[i].content).toString('utf8').replace(/<[^>]+>/g, '');
}
