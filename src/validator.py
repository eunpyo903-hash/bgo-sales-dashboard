"""
비고컴퍼니 매출현황 - Validator

Dashboard 생성 전에 반드시 실행한다.
빈 값을 임의로 0원으로 처리하지 않는다.
"""

from .utils import BRANCHES

REQUIRED_FIELDS = [
    "target_fc", "target_pt", "target_total",
    "actual_fc", "actual_pt", "actual_total",
    "rate_fc", "rate_pt", "rate_total",
    "fc_new_count", "fc_new_sales",
    "fc_renew_count", "fc_renew_sales",
    "pt_new_count", "pt_new_sales",
    "pt_renew_count", "pt_renew_sales",
]

TOLERANCE = 0.5  # 반올림으로 인한 미세한 차이 허용 범위 (퍼센트 포인트 / 원 단위 비율)


class ValidationError(Exception):
    """검증 실패 시 Dashboard를 업데이트하지 않기 위해 사용하는 예외"""
    pass


def validate_files_present(selected_files: dict):
    """7개 지점 파일이 모두 있는지 확인한다. 하나라도 없으면 ValidationError."""
    missing = [no for no in BRANCHES if no not in selected_files]
    if missing:
        missing_names = ", ".join(BRANCHES[no] for no in missing)
        raise ValidationError(f"{missing_names}점 총매출 파일이 없습니다.")


def validate_branch_fields(branch_no: int, data: dict):
    """지점 데이터에 필수 항목이 모두 존재하는지 확인한다 (None 금지)."""
    missing = [f for f in REQUIRED_FIELDS if data.get(f) is None]
    if missing:
        raise ValidationError(
            f"{BRANCHES[branch_no]}점 데이터에서 다음 항목을 찾을 수 없습니다: {', '.join(missing)}"
        )


def cross_check_branch(branch_no: int, data: dict, warnings: list):
    """
    숫자 교차 검증:
      target_fc + target_pt vs target_total
      actual_fc + actual_pt vs actual_total
      계산된 달성률 vs Excel 달성률
    차이가 있으면 warnings 리스트에 WARNING을 추가한다 (실행은 계속한다).
    """
    name = BRANCHES[branch_no]

    target_sum = data["target_fc"] + data["target_pt"]
    if abs(target_sum - data["target_total"]) > TOLERANCE:
        warnings.append(
            f"WARNING: {name}점 목표 합계 불일치 (FC+PT={target_sum:,} / TOTAL={data['target_total']:,})"
        )

    actual_sum = data["actual_fc"] + data["actual_pt"]
    if abs(actual_sum - data["actual_total"]) > TOLERANCE:
        warnings.append(
            f"WARNING: {name}점 실적 합계 불일치 (FC+PT={actual_sum:,} / TOTAL={data['actual_total']:,})"
        )

    for label, actual_key, target_key, rate_key in [
        ("총매출", "actual_total", "target_total", "rate_total"),
        ("FC", "actual_fc", "target_fc", "rate_fc"),
        ("PT", "actual_pt", "target_pt", "rate_pt"),
    ]:
        target_val = data[target_key]
        if not target_val:
            continue
        calculated_rate = round((data[actual_key] / target_val) * 100, 1)
        excel_rate = data[rate_key]
        if abs(calculated_rate - excel_rate) > TOLERANCE:
            warnings.append(
                f"WARNING: {name}점 {label} 달성률 불일치 "
                f"(계산값={calculated_rate}% / Excel값={excel_rate}%)"
            )


def run_all_validations(selected_files: dict, branch_data: dict):
    """
    1) 파일 존재 검증
    2) 지점별 필수 항목 검증
    3) 숫자 교차 검증 (WARNING만, 실패로 처리하지 않음)

    반환: warnings 리스트
    """
    warnings = []

    validate_files_present(selected_files)

    for branch_no, data in branch_data.items():
        validate_branch_fields(branch_no, data)
        cross_check_branch(branch_no, data, warnings)

    return warnings
