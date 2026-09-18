"""
비고컴퍼니 매출현황 - 공통 유틸리티

지점 설정(BRANCHES)은 이 파일 한 곳에서만 관리한다.
"""

import re

# ------------------------------------------------------------------
# 지점 설정 (단일 소스)
# ------------------------------------------------------------------
BRANCHES = {
    1: "역삼",
    2: "선정릉",
    3: "대치",
    4: "송파",
    5: "장안",
    6: "가락",
    7: "삼전",
}

BRANCH_NAME_TO_NO = {name: no for no, name in BRANCHES.items()}

# 파일명 패턴 예: 2609_b_1_역삼점_총매출.xlsx
# YYMM _ b _ 지점번호 _ 지점명(+"점" 선택) _ 총매출 ...
FILENAME_PATTERN = re.compile(
    r"^(?P<yymm>\d{4})_b_(?P<branch_no>\d)_(?P<branch_name>[^_]+?)점?_?총매출"
)


class ParsedFileName:
    def __init__(self, yymm: str, branch_no: int, branch_name: str, path: str):
        self.yymm = yymm
        self.branch_no = branch_no
        self.branch_name = branch_name
        self.path = path

    def __repr__(self):
        return f"ParsedFileName(yymm={self.yymm}, branch_no={self.branch_no}, branch_name={self.branch_name}, path={self.path})"


def parse_filename(filename: str):
    """
    파일명에서 YYMM / 지점번호 / 지점명을 정규식으로 추출한다.
    매칭되지 않으면 None을 반환한다 (특정 월을 코드에 고정하지 않는다).
    """
    match = FILENAME_PATTERN.match(filename)
    if not match:
        return None

    branch_no = int(match.group("branch_no"))
    if branch_no not in BRANCHES:
        return None

    return {
        "yymm": match.group("yymm"),
        "branch_no": branch_no,
        "branch_name_in_file": match.group("branch_name"),
    }


def format_won(amount) -> str:
    """숫자를 '123,456,789원' 형태로 표시한다. 원본 값은 변형하지 않는다."""
    if amount is None:
        return "-"
    try:
        return f"{round(amount):,}원"
    except (TypeError, ValueError):
        return "-"


def format_kr_unit(amount) -> str:
    """숫자를 '1억 2,346만원' 같은 축약 표현으로 변환한다 (KPI 카드 보조 표시용)."""
    if amount is None:
        return "-"
    try:
        amount = int(round(amount))
    except (TypeError, ValueError):
        return "-"

    sign = "-" if amount < 0 else ""
    amount = abs(amount)

    eok = amount // 100_000_000
    man = (amount % 100_000_000) // 10_000

    parts = []
    if eok:
        parts.append(f"{eok}억")
    if man or not parts:
        parts.append(f"{man:,}만원")
    else:
        parts[-1] = parts[-1]  # eok만 있는 경우
        parts.append("원")

    return sign + " ".join(parts)


def safe_div_rate(numerator, denominator) -> float:
    """0으로 나누기를 방지하면서 퍼센트 달성률을 계산한다."""
    if not denominator:
        return 0.0
    return round((numerator / denominator) * 100, 1)
