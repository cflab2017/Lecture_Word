# 기초 6강 — 워드 PDF 변환 (다른 이름으로 저장 PDF, 메모가 같이 찍힐 때, 최소 크기 · 책갈피 · 쪽 범위)

- 강의: https://www.coding-now.com/blog/word/word-basic-06-save-as-pdf
- English: https://www.coding-now.com/en/blog/word/word-basic-06-save-as-pdf

## 파일

| 파일 | 내용 |
|---|---|
| `실습.docx` | 따라 할 파일 — 가상 동아리 11월 활동 보고서 3쪽(제목 1 · 2, 링크, 큰 그림, 표, 제목에 메모 하나) |
| `정답.pdf` | 메모 빼고(게시 대상 '문서') · 제목으로 책갈피 · 표준으로 내보낸 PDF |
| `practice.docx` · `answer.pdf` | 영어판 강의용 같은 연습 파일 |
| `measure.py` → `measure.json` | 같은 문서를 일곱 가지로 PDF 저장해 PyMuPDF 로 쪽 수 · 크기 · 링크 · 책갈피 · 글꼴 · 그림 픽셀 · 메모 글자를 잰 값 |
| `measure.py en` → `measure_en.json` | 영어판 자료로 같은 측정 |
| `make_files.py` | 실습 · 정답 파일 만들기(먼저 measure.py) |
| `screens.py` · `crop_web.py` | 강의 화면(`shots/`, 저장소에는 올리지 않음) — 메모 창, 다른 이름으로 저장(PDF), PDF 옵션, PDF 첫 쪽 비교 |

스크립트는 Windows + 워드 + `pip install pywin32 pymupdf pillow` 에서 돈다. 워드를 따로 띄워 쓰고 문서는 %TEMP% 에 저장한다. 메모 작성자 · 문서 작성자는 CODINGNOW.

## 핵심

- 메모가 있는 문서를 다른 이름으로 저장 → PDF(그대로): 메모 풍선이 PDF 에 찍히고 본문이 줄어듦, 책갈피 5개, 그림 944×629, 179,761바이트. PDF 옵션 기본값이 '변경 내용을 표시하는 문서' · '책갈피 만들기: 제목'.
- 내보내기 '문서'만 · 표준: 메모 없음, 그림 1379×919, 224,876바이트 · 최소 크기: 그림 601×401, 102,870바이트(절반 아래).
- 제목으로 책갈피: 목차 5개 · 2쪽만(From 2 To 2): 1쪽, 링크 0(링크는 1쪽에 있음) · PDF/A: 355,385바이트.
- 링크는 어느 방법이든 살아 있음(1개) · 글꼴은 늘 포함(맑은 고딕 일부만 — 'BCDEEE+MalgunGothic' 꼴).
- 원본 사진 2400×1600 PNG 1,522,373바이트 → docx 416,878바이트.
