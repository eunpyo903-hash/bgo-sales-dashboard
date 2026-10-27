"""
비고컴퍼니 매출현황 - 수기 반영 및 표기 제외 설정

① MANUAL_ADJUSTMENTS: 지점 Excel에 아직 입력되지 않았지만 사용자가 확인해 대시보드에 먼저
   반영하라고 지시한 매출. 해당 월 파일의 상세 내역(PT 매출 O~Q열)에 같은 금액이 이미
   입력되어 있으면 중복 집계를 막기 위해 자동으로 반영하지 않는다.
   파일에 정식 입력되면 이 목록에서 해당 항목을 삭제하면 된다.
② EXCLUDED_STAFF: 순위(PT 개인매출 순위, 4분기 프로모션)에서 표기하지 않는 담당자.
"""

from .utils import BRANCHES, safe_div_rate

EXCLUDED_STAFF = {"이은표"}

MANUAL_ADJUSTMENTS = [
    {"period": "2026-10", "branch_no": 3, "field": "pt", "amount": 3_850_000,
     "note": "지점 파일 미입력분 (사용자 확인 후 수기 반영)"},
    {"period": "2026-10", "branch_no": 3, "field": "pt", "amount": 4_802_000,
     "note": "지점 파일 미입력분 (사용자 확인 후 수기 반영)"},
]


def apply_manual_adjustments(period: str, branch_data: dict, selected_files: dict,
                             amount_in_detail) -> tuple:
    """
    branch_data에 수기 반영액을 더하고 달성률을 다시 계산한다.
    amount_in_detail(file_path, amount) -> 원본 상세내역에 같은 PT 금액이 이미 있는지 여부.
    반환: (반영된 항목 리스트, 안내/경고 메시지 리스트)
    """
    applied, messages = [], []
    for adj in MANUAL_ADJUSTMENTS:
        if adj["period"] != period or adj["branch_no"] not in branch_data:
            continue
        no, amount = adj["branch_no"], adj["amount"]
        name = BRANCHES[no]

        if amount_in_detail(selected_files[no]["path"], amount):
            messages.append(
                f"WARNING: {name}점 수기 반영 {amount:,}원은 원본 파일에 이미 입력되어 있어 반영하지 않았습니다. "
                "(src/adjustments.py에서 해당 항목을 삭제해 주세요)"
            )
            continue

        d = branch_data[no]
        d["actual_pt"] += amount
        d["actual_total"] += amount
        d["rate_pt"] = safe_div_rate(d["actual_pt"], d["target_pt"])
        d["rate_total"] = safe_div_rate(d["actual_total"], d["target_total"])
        applied.append({
            "branch_no": no, "branch_name": name, "field": "PT",
            "amount": amount, "note": adj["note"],
        })
        messages.append(f"ADJUST: {name}점 PT +{amount:,}원 수기 반영 ({adj['note']})")
    return applied, messages
