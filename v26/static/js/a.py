#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pathlib import Path
import string

MAX_PASS = 10

DELIMS = set(
    " \t\r\n"
    "\"'`()[]{}<>"
    ",;:+-*/=%&|^!?~"
)

BAD_HINTS = (
    "Ã", "Â", "Ä", "Å", "Æ",
    "â", "€", "™",
    "áº", "á»", "Ä‘", "Ä",
    "Æ°", "Æ¡", "Ă"
)

VIET = set(
    "ăâđêôơư"
    "ĂÂĐÊÔƠƯ"
    "áàảãạ"
    "ắằẳẵặ"
    "ấầẩẫậ"
    "éèẻẽẹ"
    "ếềểễệ"
    "íìỉĩị"
    "óòỏõọ"
    "ốồổỗộ"
    "ớờởỡợ"
    "úùủũụ"
    "ứừửữự"
    "ýỳỷỹỵ"
)


def score(s: str):
    score = 0

    for c in s:
        if c in VIET:
            score += 10

    for b in BAD_HINTS:
        score -= s.count(b) * 5

    return score


def transforms(token):

    yield token

    for enc in ("latin1", "cp1252"):

        try:
            yield token.encode(enc).decode("utf8")
        except:
            pass

        try:
            x = token

            for _ in range(5):
                x2 = x.encode(enc).decode("utf8")

                if x2 == x:
                    break

                yield x2
                x = x2

        except:
            pass


def repair_token(token):

    best = token
    best_score = score(token)

    seen = set()

    stack = [token]

    while stack:

        cur = stack.pop()

        if cur in seen:
            continue

        seen.add(cur)

        sc = score(cur)

        if sc > best_score:
            best = cur
            best_score = sc

        for nxt in transforms(cur):
            if nxt not in seen:
                stack.append(nxt)

    return best


def tokenize(text):

    start = 0

    for i, c in enumerate(text):

        if c in DELIMS:

            if start != i:
                yield True, text[start:i]

            yield False, c

            start = i + 1

    if start != len(text):
        yield True, text[start:]


def repair(text):

    for _ in range(MAX_PASS):

        changed = False

        out = []

        for is_token, part in tokenize(text):

            if not is_token:
                out.append(part)
                continue

            if not any(x in part for x in BAD_HINTS):
                out.append(part)
                continue

            fixed = repair_token(part)

            if fixed != part:
                changed = True

            out.append(fixed)

        text2 = "".join(out)

        if not changed:
            break

        text = text2

    return text


def main():

    inp = "bplan_report.js"
    out = "output.js"

    text = Path(inp).read_text(
        encoding="utf8",
        errors="surrogateescape"
    )

    text = repair(text)

    Path(out).write_text(
        text,
        encoding="utf8",
        errors="surrogateescape"
    )

    print("Done")


if __name__ == "__main__":
    main()