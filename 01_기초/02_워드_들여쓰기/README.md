# 기초 2강 — 워드 들여쓰기 (첫 줄 들여쓰기 · 내어쓰기(둘째 줄 이하) · 눈금자 · Tab 과 다른 점)

- 강의: https://www.coding-now.com/blog/word/word-basic-02-indent
- English: https://www.coding-now.com/en/blog/word/word-basic-02-indent

## 파일

| 파일 | 내용 |
|---|---|
| `실습.docx` | 따라 할 파일 — 가상 '11월 도서관 이용 안내', 들여쓰기 없음(둘째 본문 단락은 Tab 으로 민 상태) |
| `정답.docx` | 본문 첫 줄 2글자 + 번호 단락 둘째 줄 이하 2글자 |
| `practice.docx` · `answer.docx` | 영어판 강의용(첫 줄 0.5in · 내어쓰기 0.25in) |
| `measure.py` → `measure.json` · `measure_en.json` | 줄마다 첫 글자 가로 위치(pt)·단축키 명령·홈 탭 단추가 바꾸는 값. `python measure.py en` 은 영어판 |
| `measure_text.py` | 한국어 연습 글(measure·screens·make_files 공통) |
| `make_files.py` | 실습·정답 파일을 만드는 스크립트 |
| `screens.py` → `crop_web.py` | 강의 화면(문서·눈금자·단락 창)을 찍고 강의 페이지용으로 자른 스크립트(`shots/` 는 올리지 않음) |

스크립트는 Windows + 워드 + `pip install pywin32 pillow` 에서 돈다. 워드를 따로 띄워 쓰고(가로 위치는 보이는 창에서 Selection.Information 으로만 맞게 나왔다), 시험 문서는 저장하지 않고 닫는다.

## 핵심

- 첫 줄 2글자 = 22pt(맑은 고딕 11pt) → 첫 줄 72 → 93.75pt, 둘째 줄 72 그대로.
- 내어쓰기는 단락 창 첫 줄 목록의 **'둘째 줄 이하'** — 번호 뒤 Tab 과 둘째 줄이 같은 93.75pt 에 맞는다.
- Tab 은 기본 탭 간격 40pt 까지(111.75pt), 단락 서식은 0 그대로.
- 홈 탭 ⇤ 내어쓰기 · ⇥ 들여쓰기 단추는 왼쪽 들여쓰기를 1글자(11pt)씩 — 단락 창의 '둘째 줄 이하'와 다르다.
- Ctrl+M = Indent(다음 탭 40pt) · Ctrl+T = HangingIndent · Ctrl+Q = ResetPara.
