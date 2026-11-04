# 기초 1강 — 워드 줄간격 (1.0 · 1.15 · 2.0 · 단락 사이 간격 · 줄 간격이 안 바뀔 때)

- 강의: https://www.coding-now.com/blog/word/word-basic-01-line-spacing
- English: https://www.coding-now.com/en/blog/word/word-basic-01-line-spacing

## 파일

| 파일 | 내용 |
|---|---|
| `실습.docx` | 따라 할 파일 — 회의 안내문, 단락마다 줄 간격이 제각각이고 빈 줄로 띄워져 있다 |
| `정답.docx` | 본문 단락을 배수 1.15 · 단락 뒤 10pt 로 맞추고 빈 줄을 지운 결과 |
| `practice.docx` · `answer.docx` | 영어판 강의용 — 한 줄 간격 에세이 → 두 줄 간격(단락 뒤 0) |
| `measure.py` → `measure.json` | 줄 간격 설정마다 줄 사이(pt), 단락 뒤 간격, 글꼴별 한 줄 높이, Shift+Enter, 줄 격자, 단축키 명령을 재는 스크립트 |
| `measure_en.py` → `measure_en.json` | 영어 글로 같은 측정(Aptos · Times New Roman · Calibri · Arial) |
| `menu_names.py` → `menu_names.json` | '줄 및 단락 간격' 메뉴 항목 이름을 화면의 접근성 정보로 읽는 스크립트 |
| `screens.py` · `dialog_rules.py` · `crop_web.py` · `screens_util.py` | 강의 화면을 찍고 자른 스크립트(`shots/` 는 저장소에 올리지 않음) |
| `make_files.py` | 실습·정답 파일을 만드는 스크립트 |

스크립트는 Windows + 워드 + `pip install pywin32 pillow` 에서 돈다. 워드를 따로 띄워 새 문서에서 하고 저장하지 않으며, Normal.dotm(기본 서식 파일)은 바꾸지 않는다.

## 핵심

- 줄 간격: 홈 → 단락 → 줄 및 단락 간격(1.0 · 1.15 · 1.5 · 2.0 · 2.5 · 3.0), 단축키 Ctrl+1 · Ctrl+5 · Ctrl+2.
- 단락 사이는 빈 줄이 아니라 단락 뒤 간격으로. Shift+Enter 는 단락을 나누지 않아 단락 뒤 간격이 붙지 않는다.
- 줄 간격 1.0 인데 두 배로 벌어지면 줄 격자 — 단락 대화상자의 '[페이지 설정]에서 지정된 문자 수에 맞춰 문자 간격 조정'을 끈다.
