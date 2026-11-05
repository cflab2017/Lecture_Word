"""워드 기초 3강 공통 글 — 가상 '독서 동아리 11월 활동 보고서'(표지 1쪽 + 본문 3쪽). 사람·단체 이름은 모두 가상이다."""

KO = {
    "cover": ["독서 동아리 11월 활동 보고서", "2026년 11월 · 새솔 독서 동아리(가상)"],
    "pages": [
        ["1. 모임 개요", "11월에는 매주 토요일 오전 10시에 시립 도서관 2층 세미나실에서 네 번 모였습니다. 회원 12명 가운데 평균 9명이 나왔고, "
         "모임마다 한 사람이 발제를 맡아 30분 동안 책 내용을 정리해 들려준 뒤 나머지 시간에는 자유롭게 이야기를 나눴습니다."],
        ["2. 함께 읽은 책", "이번 달에는 소설 두 권과 에세이 두 권을 읽었습니다. 소설은 분량이 많아 두 주에 나눠 읽었고, 에세이는 한 번에 읽고 "
         "마음에 남은 문장을 한 사람에 하나씩 골라 와서 돌아가며 소리 내어 읽었습니다."],
        ["3. 다음 달 계획", "12월에는 연말 모임을 겸해 한 해 동안 읽은 책 가운데 가장 좋았던 한 권을 소개하는 시간을 갖습니다. "
         "새 회원 모집 공고는 12월 첫 주에 도서관 게시판과 동아리 단체 대화방에 올립니다."],
    ],
}
EN = {
    "cover": ["Book Club Report — November", "November 2026 · Pinewood Book Club (fictional)"],
    "pages": [
        ["1. Meetings", "In November the club met four times, every Saturday at 10 a.m. in the second-floor seminar room of the city library. "
         "On average 9 of our 12 members came. Each week one member led a 30-minute summary of the book before an open discussion."],
        ["2. What we read", "We read two novels and two essay collections this month. The novels were long, so each one was split across two weeks; "
         "for the essays, everyone brought one favorite passage and we took turns reading them aloud."],
        ["3. Plans for December", "December doubles as our year-end meeting: each member will introduce the best book they read this year. "
         "The call for new members goes up in the first week of December on the library notice board and in the club group chat."],
    ],
}


def build(doc, t, filler=10):
    """표지 1쪽 + 본문 3쪽 — 쪽마다 쪽 나누기. filler: 본문 쪽마다 같은 문단을 몇 번 더 넣을지(쪽이 꽉 차 보이게, 그래도 한 쪽 안)."""
    r = doc.Content
    r.Text = ""
    parts = []
    parts.append("\r".join(t["cover"]))
    for head, body in t["pages"]:
        parts.append("\r".join([head] + [body] * 3))
    doc.Content.InsertAfter(parts[0])
    for p in parts[1:]:
        end = doc.Content
        end.Collapse(0)
        end.InsertBreak(7)  # wdPageBreak
        doc.Content.InsertAfter(p)
    return doc
