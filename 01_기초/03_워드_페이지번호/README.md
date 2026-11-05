# 기초 3강 — 워드 페이지 번호 (아래쪽 가운데 · 표지 빼고 2쪽부터 1 · 'X / Y')

- 강의: https://www.coding-now.com/blog/word/word-basic-03-page-numbers
- English: https://www.coding-now.com/en/blog/word/word-basic-03-page-numbers

## 파일

| 파일 | 내용 |
|---|---|
| `실습.docx` | 따라 할 파일 — 가상 '독서 동아리 11월 활동 보고서'(표지 1쪽 + 본문 3쪽, 페이지 번호 없음) |
| `정답.docx` | 아래쪽 가운데 번호 + 첫 페이지를 다르게 지정 + 시작 번호 0 (표지 없음, 2쪽부터 1 · 2 · 3) |
| `practice.docx` · `answer.docx` | 영어판 강의용 같은 연습 파일(영어 글) |
| `measure.py` → `measure.json` | 번호 없음 · 아래쪽 가운데 · 첫 페이지 다르게 · 시작 번호 0 · 시작 번호 5 · 직접 친 1 · PAGE / NUMPAGES · 표지 빼고 X / Y · `{ = { NUMPAGES } - 1 }` 을 PDF 로 저장해 쪽마다 바닥글 글자와 위치(mm)를 읽는다 |
| `measure.py en` → `measure_en.json` | 영어판 자료로 같은 측정 + 영어판 연습 파일 |
| `measure_text.py` | 보고서 글(한·영)과 4쪽 만들기 공통 |
| `screens.py` → `crop_web.py` | 강의 화면을 찍고(`shots/`, 저장소에는 올리지 않음) 강의 페이지용 그림을 만든 스크립트(쪽 아래 그림은 정답 PDF 에서) |

스크립트는 Windows + 워드 + `pip install pywin32 pillow pymupdf` 에서 돈다. 워드를 따로 띄워 쓰고, 시험 문서는 저장하지 않고 닫는다(실습·정답만 저장).

## 핵심

- 아래쪽 가운데 번호 → 1 · 2 · 3 · 4(번호 가운데 105.0mm = A4 폭의 한가운데, 아래 끝에서 21.2mm).
- 첫 페이지를 다르게 지정 → 표지 번호만 사라지고 2쪽은 2. 페이지 번호 서식 → 시작 번호 0 → 2쪽부터 1 · 2 · 3.
- 바닥글에 1 을 직접 치면 4쪽 모두 1.
- PAGE / NUMPAGES 에서 표지를 빼면 끝 쪽이 3 / 4 — `{ = { NUMPAGES } - 1 }` 로 3 / 3.
- 메뉴 이름(이 PC 워드): 페이지 번호 서식... · 페이지 번호 제거 · 첫 페이지를 다르게 지정 · 머리글/바닥글 닫기. 서식 파일의 아래쪽 번호 갤러리는 '단순형 — 일반 번호 1·2·3', 'X/Y 페이지 — 굵게 표시된 번호 1~3'.
