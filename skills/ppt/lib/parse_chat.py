#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
parse_chat.py — 카톡 채팅 export(.txt) → (시간, 발신자, 제목, 사진수) 후보 목록(JSON)

카톡 규칙:
    YYYY. M. D. HH:MM, 발신자 : 사진          (사진 1장)
    YYYY. M. D. HH:MM, 발신자 : 사진 N장        (사진 N장)
    YYYY. M. D. HH:MM, 발신자 : <제목>          (바로 다음 같은/다른 발신자 메시지 = 제목)

즉, "사진" / "사진 N장" 메시지를 만나면 그 직후의 일반 텍스트 메시지를 제목 후보로 잡는다.
메뉴/링크/무제 등은 maybe 플래그로 표시.

사용법:
    python3 parse_chat.py <chat.txt> --date YYMMDD [--since YYYY.M.D]
    --date  : 대상 주일(YYMMDD). --since 생략 시 이 날짜의 7일 전(직전 일요일)부터 필터한다.
              대상 날짜가 아니라 "실행하는 날"을 기준으로 잡으면 토요일 사진이 잘려 나갈 수 있으므로 항상 넘긴다.
    --since : 필터 시작일 직접 지정. --date 도 --since 도 없으면 오늘 기준 직전 일요일.
출력: JSON  { "since": "YYYY.M.D", "candidates": [ {time,sender,title,photos,maybe,reason}, ... ] }
"""
import sys, re, json, argparse, datetime


def last_sunday(today=None):
    """직전 일요일. 오늘이 일요일이면 7일 전(지난주 일요일)."""
    today = today or datetime.date.today()
    back = (today.weekday() + 1) % 7  # Mon=0..Sun=6 → Sun이면 0
    return today - datetime.timedelta(days=back or 7)

LINE = re.compile(r'^(\d{4})\.\s*(\d{1,2})\.\s*(\d{1,2})\.\s*(\d{1,2}:\d{2}),\s*(.+?)\s*:\s*(.*)$')
PHOTO = re.compile(r'^사진(?:\s*(\d+)\s*장)?$')
SKIP_HINTS = ("http://", "https://", "파일:", "님이", "삭제되었습니다", "초대", "이모티콘",
              "baemin", "s.baemin", "menu", "메뉴", "주문")


def parse(path, since=None):
    raw = open(path, encoding="utf-8-sig").read().splitlines()
    msgs = []  # (date_tuple, time, sender, text)
    for ln in raw:
        m = LINE.match(ln)
        if not m:
            # continuation line of previous message body — append
            if msgs and ln.strip():
                d, t, s, txt = msgs[-1]
                msgs[-1] = (d, t, s, (txt + "\n" + ln).strip())
            continue
        y, mo, da, tm, sender, text = m.groups()
        msgs.append(((int(y), int(mo), int(da)), tm, sender.strip(), text.strip()))

    since_t = None
    if since:
        p = [int(x) for x in re.split(r'[.\s]+', since.strip()) if x]
        since_t = (p[0], p[1], p[2])

    cands = []
    i = 0
    while i < len(msgs):
        d, t, s, txt = msgs[i]
        pm = PHOTO.match(txt)
        if pm and (since_t is None or d >= since_t):
            n = int(pm.group(1)) if pm.group(1) else 1
            # look ahead for the title (next non-photo text, prefer same sender)
            title, maybe, reason = None, False, ""
            for j in range(i+1, min(i+4, len(msgs))):
                _, _, s2, txt2 = msgs[j]
                if PHOTO.match(txt2):
                    break
                low = txt2.lower()
                if any(h in low for h in SKIP_HINTS):
                    continue
                if not txt2:
                    continue
                title = txt2.splitlines()[0][:60]
                if s2 != s:
                    maybe, reason = True, "제목이 다른 발신자(%s)" % s2
                break
            if title is None:
                title, maybe, reason = "(무제)", True, "다음 줄에 제목 없음"
            cands.append({
                "time": "%04d.%02d.%02d %s" % (d[0], d[1], d[2], t),
                "sender": s, "title": title, "photos": n,
                "maybe": maybe, "reason": reason,
            })
        i += 1

    return {"since": since, "candidates": cands, "total_titles": len(cands),
            "total_photos": sum(c["photos"] for c in cands)}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("chat")
    ap.add_argument("--date", default=None, help="대상 주일 YYMMDD (since 기본값 = 이 날짜-7일)")
    ap.add_argument("--since", default=None,
                    help="YYYY.M.D 이후만. 생략 시 직전 일요일부터")
    args = ap.parse_args()
    if args.since:
        since = args.since
    elif args.date:
        d = datetime.datetime.strptime(args.date, "%y%m%d").date()
        since = (d - datetime.timedelta(days=7)).strftime("%Y.%m.%d")
    else:
        since = last_sunday().strftime("%Y.%m.%d")
    print(json.dumps(parse(args.chat, since), ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
