# 기초 8강 — 워드 단축키 (글자 크기 · 서식 지우기 · 들여쓰기 · Shift+Enter, Ctrl+Shift+V 는 텍스트만 붙여넣기)

- 강의: https://www.coding-now.com/blog/word/word-basic-08-shortcuts
- English: https://www.coding-now.com/en/blog/word/word-basic-08-shortcuts

## 파일

| 파일 | 내용 |
|---|---|
| `실습.docx` | 따라 할 파일 — 가상 사진 동아리 11월 정기 모임 안내, 서식 없는 6단락 |
| `정답.docx` | 제목 굵게 · 가운데 · 14pt, 일시~문의 단락 내어쓰기(왼쪽 40pt · 첫 줄 -40pt), 일시 단락에 Shift+Enter 로 '비가 오면 …' 줄 |
| `practice.docx` · `answer.docx` | 영어판 강의용 같은 연습 파일 |
| `measure.py` → `measure.json` | 워드 키 지정표(`Application.FindKey`)에서 단축키별 명령 이름을 읽고, 그 명령을 실습 파일 사본에서 실행해 결과를 읽음 |
| `measure.py en` → `measure_en.json` | 영어판 파일로 같은 측정 |
| `make_files.py` | 실습 · 정답 파일 만들기(`en` 이면 영어판) |
| `crop_web.py` | 실습 · 정답을 PDF 로 내보내 나란히 놓은 강의 그림 |

스크립트는 Windows + 워드 + `pip install pywin32 pymupdf pillow` 에서 돈다. 워드 창은 띄우지 않는다(숨긴 워드).
`BuildKeyCode` 에는 수식키를 따로따로 넘겨야 한다(`BuildKeyCode(Ctrl, Shift, 키)` — 합쳐 넘기면 '매개 변수가 잘못'). WordBasic 명령 중 숨긴 창에서 반영되지 않는 것(Bold · CenterPara · SpacePara · ResetChar · ResetPara · ChangeCase · ShowAll)은 같은 동작의 개체 모델 메서드로 실행했다.
처음엔 보이는 워드 창에 키를 보내 봤는데 Shift 가 섞인 키가 자주 빠져 결과가 들쑥날쑥해 이 방식으로 바꿨다.

## 핵심(이 PC 워드, Microsoft 365)

- 키 지정표: Ctrl+] GrowFontOnePoint · Ctrl+Shift+> GrowFont · Ctrl+Space ResetChar · Ctrl+Q ResetPara · Ctrl+M Indent · Ctrl+Shift+M UnIndent · Ctrl+T HangingIndent · Ctrl+Shift+T UnHang · Shift+Enter 문자 11 · Ctrl+Enter InsertPageBreak · Ctrl+Shift+8 ShowAll · F4 · Ctrl+Y EditRedoOrRepeat · Shift+F3 ChangeCase · **Ctrl+Shift+V PasteTextOnly** · Ctrl+Shift+C 명령 없음.
- 제목 11pt: Ctrl+] ×3 → 12 · 13 · 14, Ctrl+[ → 13, Ctrl+Shift+> ×4 → 14 · 16 · 18 · 20, Ctrl+Shift+< → 18. Ctrl+Space → 굵게 꺼짐 · 11pt(가운데 그대로), Ctrl+Q → 왼쪽.
- 맞춤 Ctrl+R 2 · J 3 · E 1 · L 0, 줄 간격 Ctrl+2 24 · Ctrl+5 18 · Ctrl+1 12.
- 들여쓰기(기본 탭 40pt): Ctrl+M 40.05 → 79.95, Ctrl+Shift+M 40.05 → 0, Ctrl+T 40 / -40 → 80 / -80, Ctrl+Shift+T 40 / -40, Ctrl+Q 0 / 0.
- Shift+Enter: 단락 6 → 6(문자 11), Enter: 6 → 7(단락 뒤 8pt), Ctrl+Enter: 1쪽 → 2쪽. Shift+F3: photo walk → Photo Walk → PHOTO WALK → photo walk.
