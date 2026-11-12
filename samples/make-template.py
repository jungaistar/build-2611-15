# 팀장 보고용 Word 기본 틀(꼬리표가 들어간 docx)을 만든다 — python-docx 필요.
# 꼬리표는 js/linkage.js 의 reportData() 키와 같다. 틀을 바꾸면 js/template.js(base64)도 다시 만든다.
import base64, os
from docx import Document
from docx.shared import Pt, Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn

here = os.path.dirname(os.path.abspath(__file__))
d = Document()
sec = d.sections[0]
sec.orientation = WD_ORIENT.LANDSCAPE
sec.page_width, sec.page_height = Mm(297), Mm(210)
for m in ('left_margin', 'right_margin'): setattr(sec, m, Mm(15))
sec.top_margin = sec.bottom_margin = Mm(15)
st = d.styles['Normal']; st.font.name = '맑은 고딕'; st.font.size = Pt(9.5)
st.element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

def para(text, size=9.5, bold=False, align=None):
    p = d.add_paragraph(); r = p.add_run(text); r.bold = bold; r.font.size = Pt(size)
    r.font.name = '맑은 고딕'; r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
    if align: p.alignment = align
    p.paragraph_format.space_after = Pt(3)
    return p

def table(heads, tags, widths=None):
    t = d.add_table(rows=2, cols=len(heads)); t.style = 'Table Grid'
    for i, h in enumerate(heads):
        c = t.rows[0].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(h); r.bold = True; r.font.size = Pt(8.5)
    for i, g in enumerate(tags):
        c = t.rows[1].cells[i]; c.text = ''; r = c.paragraphs[0].add_run(g); r.font.size = Pt(8.5)
    return t

para('OEM 원소재 연동제 검토 보고', 16, True, WD_ALIGN_PARAGRAPH.CENTER)
para('작성일 {작성일}   ·   적용 월 {적용월}', 9, False, WD_ALIGN_PARAGRAPH.RIGHT)

para('1. 개요', 11.5, True)
t = d.add_table(rows=5, cols=4); t.style = 'Table Grid'
cells = [('업체명', '{업체명}', '계약 버전 · 기간', '{계약버전} · {계약기간}'),
         ('적용 품목', '{적용품목}', '조정요건 · 조정주기', '{조정요건} · {적용주기}'),
         ('조정대금 반영시점', '{반영시점}', 'Index 파일', '{Index파일}'),
         ('기준 Index 기간', '{기준기간}', '변경 Index 기간', '{변경기간}'),
         ('대상 품번', '{품번수}건 (계산 완료 {계산완료} · 확인 필요 {확인필요건수})', '연동액 합계(품번당 합)', '{연동액합계} 원')]
for r, row in enumerate(cells):
    for c, v in enumerate(row):
        cell = t.rows[r].cells[c]; cell.text = ''; run = cell.paragraphs[0].add_run(v); run.font.size = Pt(9); run.bold = (c % 2 == 0)

para('')
para('2. 계약 조건', 11.5, True)
para('기준지표: {기준지표}')
para('기준시점 · 비교시점: {기준시점}')
para('연동산식: {계약수식}')
para('{계약조건}')

para('3. 원소재 · 환율 Index 변동', 11.5, True)
table(['원소재(규격)', '적용 Index', '단위', '기준 Index', '변경 Index', '변동률', '적용 변동률', '비고'],
      ['{#items}{원소재}', '{Index명}', '{단위}', '{기준Index}', '{변경Index}', '{변동률}', '{적용변동률}', '{비고}{/items}'])
para('유사 Index 적용 여부 및 근거: {유사Index}')

para('4. 품번별 계산 결과', 11.5, True)
table(['품번', '품명', '기준 파트단가', '원소재별 사용량', '원소재 연동액', '환율 연동액', '전체 연동액', '계약 계산 변경단가', '변동률', '업체 요청단가', '차액(요청 − 계산)', '검토'],
      ['{#rows}{품번}', '{품명}', '{기준단가}', '{사용량}', '{원소재연동액}', '{환율연동액}', '{전체연동액}', '{변경단가}', '{변동률}', '{요청단가}', '{차액}', '{의견}{/rows}'])

para('')
para('5. 품번 · 원소재별 내역', 11.5, True)
table(['품번', '원소재', '사용량(kg)', '기준 Index', '변경 Index', '적용 변동률', '연동액(원)'],
      ['{#details}{품번}', '{원소재}', '{사용량}', '{기준Index}', '{변경Index}', '{변동률}', '{연동액}{/details}'])

para('')
para('6. 검토 의견', 11.5, True)
para('{검토의견}')
para('확인이 필요한 항목', 10, True)
para('{경고}')
para('※ 계산은 올린 Index 엑셀과 등록한 계약 조건만으로 했습니다. 값이 없는 항목은 「확인 필요」로 표시했습니다.', 8.5)

out = os.path.join(here, 'report-template.docx')
d.save(out)
b64 = base64.b64encode(open(out, 'rb').read()).decode()
open(os.path.join(here, '..', 'js', 'template.js'), 'w').write(
    '/* samples/make-template.py 가 만든 파일 — 직접 고치지 않는다. 기본 Word 보고서 틀(base64) */\n'
    '(function (root) { var b = "' + b64 + '"; if (typeof module === "object" && module.exports) module.exports = b; else root.REPORT_TEMPLATE_B64 = b; })(typeof self !== "undefined" ? self : this);\n')
print('ok', len(b64))
