"""
비고컴퍼니 매출현황 - Aggregator

7개 지점 데이터를 합산하여 회사 전체 집계, 지점별 순위,
PT 개인매출 통합 순위, 표준 JSON 구조를 만든다.
"""

from datetime import datetime

from .utils import BRANCHES, safe_div_rate


def aggregate_company(branch_data: dict) -> dict:
    """
    회사 전체 집계.
    주의: 지점별 달성률의 평균이 아니라, 반드시
    (전체 실제매출 합계) / (전체 목표 합계) 로 계산한다.
    """
    target_total = sum(d["target_total"] for d in branch_data.values())
    actual_total = sum(d["actual_total"] for d in branch_data.values())

    target_fc = sum(d["target_fc"] for d in branch_data.values())
    actual_fc = sum(d["actual_fc"] for d in branch_data.values())

    target_pt = sum(d["target_pt"] for d in branch_data.values())
    actual_pt = sum(d["actual_pt"] for d in branch_data.values())

    return {
        "target_total": target_total,
        "actual_total": actual_total,
        "rate_total": safe_div_rate(actual_total, target_total),

        "target_fc": target_fc,
        "actual_fc": actual_fc,
        "rate_fc": safe_div_rate(actual_fc, target_fc),

        "target_pt": target_pt,
        "actual_pt": actual_pt,
        "rate_pt": safe_div_rate(actual_pt, target_pt),
    }


def aggregate_new_renew_totals(branch_data: dict) -> dict:
    """FC/PT 신규·재등록 전체 집계."""
    def total_of(key):
        return sum(d[key] for d in branch_data.values())

    return {
        "fc_new_count": total_of("fc_new_count"),
        "fc_new_sales": total_of("fc_new_sales"),
        "fc_renew_count": total_of("fc_renew_count"),
        "fc_renew_sales": total_of("fc_renew_sales"),
        "pt_new_count": total_of("pt_new_count"),
        "pt_new_sales": total_of("pt_new_sales"),
        "pt_renew_count": total_of("pt_renew_count"),
        "pt_renew_sales": total_of("pt_renew_sales"),
    }


def rank_branches(branch_data: dict) -> list:
    """총매출 목표 달성률(rate_total) 내림차순으로 지점 순위를 매긴다."""
    ordered = sorted(
        branch_data.items(),
        key=lambda item: item[1]["rate_total"],
        reverse=True,
    )

    ranking = []
    for rank, (branch_no, data) in enumerate(ordered, start=1):
        ranking.append({
            "rank": rank,
            "branch_no": branch_no,
            "branch_name": BRANCHES[branch_no],
            "target_total": data["target_total"],
            "actual_total": data["actual_total"],
            "rate_total": data["rate_total"],
            "actual_fc": data["actual_fc"],
            "rate_fc": data["rate_fc"],
            "actual_pt": data["actual_pt"],
            "rate_pt": data["rate_pt"],
        })
    return ranking


def build_pt_ranking(trainer_rows: list) -> list:
    """
    trainer_rows: [{"branch_no": int, "branch_name": str, "name": str, "sales": number}, ...]
    branch + name 을 고유 식별 기준으로 매출 내림차순 순위를 만든다.
    동명이인은 지점이 다르면 별도 인물로 취급한다.
    """
    dedup = {}
    for row in trainer_rows:
        key = (row["branch_no"], row["name"])
        # 동일 키가 중복 등장하면 매출이 더 큰 값을 채택 (데이터 중복 방지)
        if key not in dedup or row["sales"] > dedup[key]["sales"]:
            dedup[key] = row

    ordered = sorted(dedup.values(), key=lambda r: r["sales"], reverse=True)

    ranking = []
    for rank, row in enumerate(ordered, start=1):
        ranking.append({
            "rank": rank,
            "name": row["name"],
            "branch_no": row["branch_no"],
            "branch_name": row["branch_name"],
            "sales": row["sales"],
        })
    return ranking


def build_dashboard_json(period: str, branch_data: dict, trainer_rows: list, reference_date: str = None) -> dict:
    """표준 JSON 데이터 구조 전체를 조립한다."""
    company = aggregate_company(branch_data)
    company.update(aggregate_new_renew_totals(branch_data))

    branches = []
    for branch_no in sorted(branch_data.keys()):
        d = branch_data[branch_no]
        branches.append({
            "branch_no": branch_no,
            "branch_name": BRANCHES[branch_no],

            "target_fc": d["target_fc"],
            "target_pt": d["target_pt"],
            "target_total": d["target_total"],

            "actual_fc": d["actual_fc"],
            "actual_pt": d["actual_pt"],
            "actual_total": d["actual_total"],

            "rate_fc": d["rate_fc"],
            "rate_pt": d["rate_pt"],
            "rate_total": d["rate_total"],

            "fc_new_count": d["fc_new_count"],
            "fc_new_sales": d["fc_new_sales"],
            "fc_renew_count": d["fc_renew_count"],
            "fc_renew_sales": d["fc_renew_sales"],

            "pt_new_count": d["pt_new_count"],
            "pt_new_sales": d["pt_new_sales"],
            "pt_renew_count": d["pt_renew_count"],
            "pt_renew_sales": d["pt_renew_sales"],
        })

    return {
        "period": period,
        "reference_date": reference_date,
        "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "company": company,
        "branches": branches,
        "branch_ranking": rank_branches(branch_data),
        "pt_ranking": build_pt_ranking(trainer_rows),
    }
