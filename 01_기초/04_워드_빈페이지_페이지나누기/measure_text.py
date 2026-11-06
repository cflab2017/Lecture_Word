"""워드 기초 4강 공통 글 — 가상 '독서 동아리 11월 회비 정산서'(본문 두 장 + 마지막 장 표). 사람·단체 이름은 모두 가상이다."""

KO = {
    "title": "독서 동아리 11월 회비 정산서",
    "p1": ["1. 걷은 회비", "11월 회비는 한 사람에 1만 원씩 12명에게 걷어 모두 12만 원입니다. 지난달에서 넘어온 2만 4천 원을 더해 이번 달에 쓸 수 있는 돈은 14만 4천 원이었습니다. "
           "회비는 동아리 통장으로만 받고, 현금으로 낸 두 사람의 돈은 그날 바로 통장에 넣었습니다."],
    "p2": ["2. 쓴 곳", "모임 장소는 시립 도서관 세미나실이라 대여료가 들지 않았습니다. 이번 달에는 함께 읽을 책 네 권과 모임 때 마실 차, 연말 모임 장소 예약금을 냈습니다. "
           "쓴 내역은 아래 표에 날짜순으로 적었고, 영수증은 모두 사진으로 찍어 단체 대화방에 올렸습니다."],
    "table_head": ["날짜", "쓴 곳", "금액(원)"],
    "table_row": ["11월 {d}일", "책·차·예약금", "5,000"],
}
EN = {
    "title": "Book Club Dues — November Statement",
    "p1": ["1. Dues collected", "November dues were 10,000 won from each of 12 members, 120,000 won in total. With 24,000 won carried over from October, "
           "the club had 144,000 won to spend this month. Dues go only into the club bank account; the two members who paid in cash had it deposited the same day."],
    "p2": ["2. Spending", "We meet in the city library seminar room, so there was no room fee. This month we bought four books to read together, tea for the meetings, "
           "and paid a deposit for the year-end venue. Each expense is listed by date in the table below, and every receipt was posted to the group chat."],
    "table_head": ["Date", "Item", "Amount (won)"],
    "table_row": ["Nov {d}", "Books, tea, deposit", "5,000"],
}


def build(doc, T, table_rows=0):
    """제목 + 두 단락 묶음. table_rows > 0 이면 문서 끝에 표(머리글 + table_rows 줄)."""
    r = doc.Content
    r.Text = ""
    lines = [T["title"]] + T["p1"] + T["p2"]
    doc.Content.InsertAfter("\r".join(lines))
    doc.Paragraphs(1).Range.Font.Size = 18
    doc.Paragraphs(1).Range.Font.Bold = True
    for k in (2, 4):
        doc.Paragraphs(k).Range.Font.Bold = True
    if table_rows:
        doc.Content.InsertParagraphAfter()
        end = doc.Paragraphs(doc.Paragraphs.Count).Range
        t = doc.Tables.Add(end, table_rows + 1, 3)
        t.Borders.Enable = True
        for c, h in enumerate(T["table_head"], 1):
            t.Cell(1, c).Range.Text = h
        for i in range(table_rows):
            vals = [T["table_row"][0].format(d=1 + i % 30), T["table_row"][1], T["table_row"][2]]
            for c, v in enumerate(vals, 1):
                t.Cell(i + 2, c).Range.Text = v
        return t
    return None
