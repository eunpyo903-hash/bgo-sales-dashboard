"""
비고컴퍼니 매출현황 - 4분기 프로모션 집계

프로모션 기간(10/1~12/31)에 해당하는 월별 총매출 Excel을 모두 읽어
트레이너별 PT 개인매출을 누적하고 두 프로모션의 순위를 만든다.

  ① PT·PT시니어 : 기간 내 누적 개인 매출 1등 100만 / 2등 50만 / 3등 20만
                  (대상: 시니어·팀원 직급, PT부서장 제외 / 워크인 매출 제외)
  ② PT부서장    : 기간 내 누적 개인 매출 1등 50만
                  (부서장 공석 시 해당 지점장으로 대체)

워크인 매출 구분:
  [매출] 시트 상세 거래 내역의 비고(V열) 또는 기타(F열)에 '워크인'이라고
  적힌 행의 PT 매출(O,P,Q,T열)을 해당 담당자 매출에서 뺀다.
"""

from .parser import extract_pt_ranking, extract_walkin_pt_sales, select_files_for_month
from .utils import BRANCHES

PROMO_TITLE = "2026 4분기 프로모션"
PROMO_PERIOD_LABEL = "2026.10.01 ~ 2026.12.31"
PROMO_MONTHS = ["2610", "2611", "2612"]

# PT부서장 명단 (지점번호: [성명, ...]) - 인사 변경 시 이 목록만 수정한다.
PT_HEADS = {
    1: ["황하늘", "이초은"],   # 역삼
    2: ["장광희", "변재용"],   # 선정릉
    3: ["이상훈", "배재민"],   # 대치
    4: ["김원준"],             # 송파
    5: ["김진웅"],             # 장안
    6: ["송석민"],             # 가락
    7: ["윤휘용"],             # 삼전
}

# ① 프로모션 대상 직급 ([매출순위] 시트 C열 기준)
STAFF_TITLES = {"시니어", "팀원"}

STAFF_PRIZES = [1_000_000, 500_000, 200_000]
HEAD_PRIZES = [500_000]


def _rank(rows: list, prizes: list) -> list:
    """매출 내림차순으로 순위를 매긴다. 동액은 같은 순위(공동)로 처리한다."""
    ordered = sorted(rows, key=lambda r: (-r["sales"], r["branch_no"], r["name"]))
    ranked = []
    prev_sales, prev_rank = None, 0
    for i, row in enumerate(ordered, start=1):
        rank = prev_rank if row["sales"] == prev_sales else i
        prev_sales, prev_rank = row["sales"], rank
        ranked.append({
            **row,
            "rank": rank,
            "prize": prizes[rank - 1] if rank <= len(prizes) and row["sales"] > 0 else 0,
            "gap_to_prev": (ordered[i - 2]["sales"] - row["sales"]) if i > 1 else 0,
        })
    return ranked


def build_promotion(data_dir: str, warnings: list) -> dict:
    """
    data_dir에 있는 프로모션 기간 월 파일을 모두 읽어 누적 순위를 만든다.
    기간에 해당하는 파일이 하나도 없으면 None을 반환한다.
    """
    totals = {}   # (branch_no, name) -> row
    months_used = []

    for yymm in PROMO_MONTHS:
        files = select_files_for_month(data_dir, yymm, warnings)
        if not files:
            continue
        months_used.append(yymm)

        for branch_no, file_info in sorted(files.items()):
            walkin = extract_walkin_pt_sales(file_info["path"])
            for row in extract_pt_ranking(file_info["path"], branch_no):
                key = (branch_no, row["name"])
                entry = totals.setdefault(key, {
                    "name": row["name"],
                    "branch_no": branch_no,
                    "branch_name": row["branch_name"],
                    "title": row.get("title"),
                    "gross_sales": 0,
                    "walkin_sales": 0,
                    "monthly": {},
                })
                entry["title"] = row.get("title") or entry["title"]  # 최신 월 직급 사용
                w = walkin.get(row["name"], 0)
                entry["gross_sales"] += row["sales"]
                entry["walkin_sales"] += w
                entry["monthly"][yymm] = row["sales"] - w

    if not months_used:
        return None

    for entry in totals.values():
        entry["sales"] = entry["gross_sales"] - entry["walkin_sales"]

    head_keys = {(no, name) for no, names in PT_HEADS.items() for name in names}

    heads = []
    for no, names in PT_HEADS.items():
        for name in names:
            entry = totals.get((no, name))
            if entry is None:
                warnings.append(
                    f"WARNING: 4분기 프로모션 - PT부서장 '{name}'({no}번 지점)을 [매출순위] 시트에서 찾지 못해 0원으로 표시합니다."
                )
                entry = {"name": name, "branch_no": no, "branch_name": BRANCHES[no],
                         "title": None, "gross_sales": 0, "walkin_sales": 0,
                         "monthly": {}, "sales": 0}
            heads.append(entry)

    staff = [
        e for key, e in totals.items()
        if key not in head_keys and e["title"] in STAFF_TITLES
    ]

    return {
        "title": PROMO_TITLE,
        "period_label": PROMO_PERIOD_LABEL,
        "months": [f"20{m[:2]}-{m[2:]}" for m in months_used],
        "walkin_rule": "상세 거래 내역 비고/기타 칸에 '워크인' 표기된 PT 매출 제외",
        "walkin_total": sum(e["walkin_sales"] for e in totals.values()),
        "staff": {
            "label": "PT · PT시니어",
            "prizes": STAFF_PRIZES,
            "ranking": _rank(staff, STAFF_PRIZES),
        },
        "heads": {
            "label": "PT부서장",
            "prizes": HEAD_PRIZES,
            "note": "부서장 공석 시 해당 지점장으로 대체",
            "ranking": _rank(heads, HEAD_PRIZES),
        },
    }
