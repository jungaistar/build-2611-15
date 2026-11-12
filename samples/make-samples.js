/* 예제 파일을 다시 만든다 — node samples/make-samples.js  (전부 가짜 자료)
 *  index-workbook_sample.xlsx : 원자재 가격 통합문서(머리 블록 5장 + 업체 시트 1장)
 *  index-wide_sample.xlsx     : 가로형(원소재가 행 · 날짜가 열) Index
 *  parts_sample.xlsx          : 품번별 원소재 대량 등록 예 */
'use strict';
var fs = require('fs'), path = require('path');
var XLSX = require('../lib/xlsx.full.min.js'), C = require('../js/costidx.js'), L = require('../js/linkage.js');
function write(name, wb) { fs.writeFileSync(path.join(__dirname, name), XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' })); console.log(name); }
var s = L.sample(), B = s.book, wb = XLSX.utils.book_new();
B.names.forEach(function (n) { var ws = XLSX.utils.aoa_to_sheet(B.sheets[n]); if (B.merges && B.merges[n]) ws['!merges'] = B.merges[n]; XLSX.utils.book_append_sheet(wb, ws, n); });
write('index-workbook_sample.xlsx', wb);
var all = L.buildIndex(B, '', true).series.filter(function (x) { return x.role !== 'skip' || (x.meta && x.meta.kind === 'fx'); });
var months = L.monthsBetween('2025-01', '2026-06');
var aoa = [['원소재 이름', '규격', '기준', '지역/거래소', '단위', '출처', '원본 시트'].concat(months)];
all.forEach(function (x) { var by = {}; x.points.forEach(function (p) { by[p.m] = Math.round(p.v * 100) / 100; }); aoa.push([x.meta.item || x.name, x.meta.spec || '', x.meta.basis || '', x.meta.region || '', L.unitOf(x), '가상 자료', x.meta.sheet || ''].concat(months.map(function (m) { return by[m] == null ? '' : by[m]; }))); });
var w2 = XLSX.utils.book_new(); XLSX.utils.book_append_sheet(w2, XLSX.utils.aoa_to_sheet(aoa), 'Index'); write('index-wide_sample.xlsx', w2);
var w3 = XLSX.utils.book_new(), cols = ['철강', '구리', '고무'];
var rows = [['품번', '품명', '기준 파트단가', '현재 구매단가'].concat(cols.map(function (m) { return m + '(kg)'; })).concat(['환율 적용', '환율 대상 금액'])];
s.parts.filter(function (p) { return p.vendorId === s.vendors[0].id; }).forEach(function (p) { rows.push([p.partNo, p.name, p.basePrice, p.curPrice == null ? '' : p.curPrice].concat(cols.map(function (m) { return p.usage[m] == null ? '' : p.usage[m]; })).concat([p.fx ? 'Y' : '', p.fxAmount == null ? '' : p.fxAmount])); });
XLSX.utils.book_append_sheet(w3, XLSX.utils.aoa_to_sheet(rows), '품번별원소재'); write('parts_sample.xlsx', w3);
