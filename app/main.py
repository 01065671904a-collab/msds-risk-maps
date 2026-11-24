def parse_boiling_point(combined_text: str) -> str:
    """
    '초기 끓는점과 끓는점 범위', '초기 끓는점', '끓는점 범위', '초기끓는점', '끓는점'
    바로 뒤부터 '℃'가 나오기 직전까지만 잘라서 숫자를 추출합니다.
    범위(예: 337~340 ℃)면 최소값(337)만 반환합니다.
    """
    if not combined_text:
        return ""

    patterns = [
        r"초기\s*끓는점과\s*끓는점\s*범위",
        r"초기\s*끓는점",
        r"끓는점\s*범위",
        r"초기끓는점",
        r"끓는점",
    ]

    for pat in patterns:
        for m in re.finditer(pat, combined_text, re.I):
            after = combined_text[m.end() :]
            deg_m = re.search(r"℃", after)
            if not deg_m:
                continue
            segment = after[: deg_m.start()]

            range_m = re.search(
                r"(-?\d+(?:\.\d+)?)\s*(?:~|-|–|—|to)\s*(-?\d+(?:\.\d+)?)",
                segment,
                re.I,
            )
            if range_m:
                v1 = float(range_m.group(1))
                v2 = float(range_m.group(2))
                low = min(v1, v2)
                return str(int(low)) if low.is_integer() else str(low)

            single_m = re.search(r"(-?\d+(?:\.\d+)?)", segment)
            if single_m:
                val = float(single_m.group(1))
                return str(int(val)) if val.is_integer() else str(val)

    return ""
