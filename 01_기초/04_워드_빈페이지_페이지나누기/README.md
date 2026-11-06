# 기초 4강 — 워드 빈 페이지 삭제 · 페이지 나누기 (표 뒤 빈 쪽, 엔터 대신 Ctrl+Enter)

- 강의: https://www.coding-now.com/blog/word/word-basic-04-blank-page
- English: https://www.coding-now.com/en/blog/word/word-basic-04-blank-page

## 파일

| 파일 | 내용 |
|---|---|
| `실습.docx` | 따라 할 파일 — 가상 '독서 동아리 11월 회비 정산서', '2. 쓴 곳' 을 Enter 20번으로 2쪽에 민 상태 + 2쪽을 꽉 채운 표 뒤 단락만 3쪽(빈 쪽) — 3쪽 |
| `정답.docx` | Enter 대신 페이지 나누기 + 표 뒤 단락 글꼴 1pt · 단락 간격 0 · 줄 간격 고정 1pt — 2쪽 |
| `practice.docx` · `answer.docx` | 영어판 강의용 같은 연습 파일(영어 글, 같은 구조) |
| `measure.py` → `measure.json` | 강의에 적은 값을 워드로 다시 재는 스크립트 — 단축키(FindKey), 페이지 나누기 뒤 위에 다섯 줄을 더 쓴 제목 위치, Enter 로 민 제목 위치, 빈 페이지 네 가지(빈 단락 60개 · 끝 페이지 나누기 · 끝 구역 나누기 · 표 뒤 단락)의 쪽수·PDF 쪽수와 지운 뒤 쪽수 |
| `measure.py en` → `measure_en.json` | 영어판 글로 같은 측정 + 영어판 연습 파일 |
| `measure_text.py` | 정산서 글(KR·EN)과 문서 만들기 공통 |
| `screens.py` → `crop_web.py` | 편집 기호를 켠 화면(`shots/`, 저장소에는 올리지 않음)과 PDF 쪽 비교 그림을 만든 스크립트 |

스크립트는 Windows + 워드 + `pip install pywin32 pymupdf pillow` 에서 돈다. 워드를 따로 띄워(숨김) 쓰고, 시험 문서는 저장하지 않고 닫는다.

## 핵심

- 단축키(이 PC 워드 키 연결표): Ctrl+Enter = InsertPageBreak(페이지 나누기), Ctrl+Shift+8 = ShowAll(편집 기호 '모두 표시'), Ctrl+Shift+Enter = InsertColumnBreak(단 나누기).
- '2. 쓴 곳' 앞 페이지 나누기 → 2쪽 84.8pt, 위에 다섯 줄을 더 써도 84.8pt. Enter 20번으로 민 경우 → 84.8pt 였다가 다섯 줄 뒤 193.5pt.
- 빈 단락 60개 → 3쪽(PDF 마지막 쪽 글자 0) → 지우면 1쪽. 끝 페이지 나누기 → 2쪽 → ^m 모두 바꾸기 → 1쪽. 끝 구역 나누기 → 2쪽·구역 2 → ^b → 1쪽·구역 1.
- 표 뒤 단락만 다음 쪽(KR 표 15줄 + 간격): 그 단락 Delete → 글자 수 변화 0, 2쪽 그대로. 글꼴 1pt → 1쪽. 간격 0 · 고정 1pt 까지 → PDF 1쪽. 숨김 글자로 → 2쪽 그대로(PDF 2쪽).
- 찾기 및 바꾸기는 이름 붙인 인수(FindText=…)로 부르면 바뀌지 않아 위치 인수로 부른다. 같은 실행 안의 ComputeStatistics 쪽수는 끝 단락을 줄인 뒤에도 3 으로 남은 적이 있어 파일 쪽수는 PDF 로 센다(다시 열면 2).
