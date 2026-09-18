"""
비고컴퍼니 매출현황 - 업데이트 실행 스크립트

사용법:
    python update.py

data/ 폴더의 최신 월 지점별 Excel 7개를 읽어 검증 후
output/dashboard_data.json 과 index.html 을 갱신한다.
검증에 실패하면 index.html은 수정하지 않는다.
"""

import json
import os
import shutil
import sys
from datetime import datetime

# Windows 콘솔(cmd/PowerShell)이 UTF-8이 아닐 때 한글이 깨지는 것을 방지
try:
    sys.stdout.reconfigure(encoding="utf-8")
except Exception:
    pass

from src.parser import (
    ParserNotReadyError,
    select_latest_files,
    extract_branch_data,
    extract_pt_ranking,
)
from src.validator import ValidationError, run_all_validations
from src.aggregator import build_dashboard_json
from src.utils import BRANCHES, format_won, format_period

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
BACKUP_DIR = os.path.join(BASE_DIR, "backup")
OUTPUT_JSON = os.path.join(OUTPUT_DIR, "dashboard_data.json")
INDEX_HTML = os.path.join(BASE_DIR, "index.html")

LINE = "=" * 50
SUB_LINE = "-" * 50


def header():
    print(LINE)
    print("BGO COMPANY SALES DASHBOARD")
    print(LINE)
    print()


def fail(title: str, detail: str, file_hint: str = None):
    print()
    print(LINE)
    print("UPDATE FAILED")
    print(LINE)
    print()
    print("ERROR:")
    print()
    print(detail)
    if file_hint:
        print()
        print("FILE:")
        print(file_hint)
    print()
    print("Dashboard는 업데이트하지 않았습니다.")
    print("기존 정상 Dashboard는 유지됩니다.")
    print()
    print(LINE)
    sys.exit(1)


def backup_existing_index():
    if not os.path.isfile(INDEX_HTML):
        return
    os.makedirs(BACKUP_DIR, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = os.path.join(BACKUP_DIR, f"index_{stamp}.html")
    shutil.copy2(INDEX_HTML, backup_path)
    print(f"백업 완료: backup/index_{stamp}.html")


def main():
    header()

    warnings = []
    period, selected_files = select_latest_files(DATA_DIR, warnings)

    if not period:
        fail(
            "파일 없음",
            f"data/ 폴더에서 인식 가능한 Excel 파일을 찾지 못했습니다.\n"
            f"파일명 규칙(예: 2609_b_1_역삼점_총매출.xlsx)을 확인해주세요."
        )

    period_label = format_period(period)
    print(f"대상월:")
    print(f"{period_label} (YYMM: {period})")
    print()

    # [FILE CHECK]
    print("[FILE CHECK]")
    print()
    ok_count = 0
    for no, name in BRANCHES.items():
        status = "OK" if no in selected_files else "MISSING"
        if status == "OK":
            ok_count += 1
        print(f"{no}. {name:<8} {status}")
    print()
    print(f"{ok_count} / {len(BRANCHES)} FILES OK")
    print()
    print(SUB_LINE)
    print()

    for w in warnings:
        print(w)
    if warnings:
        print()

    try:
        run_all_validations(selected_files, {})  # 파일 존재만 우선 검증 (조기 실패)
    except ValidationError as e:
        fail("파일 누락", f"{period_label} {e}")

    # [PARSING]
    print("[PARSING]")
    print()
    branch_data = {}
    trainer_rows = []
    try:
        for no, name in BRANCHES.items():
            file_info = selected_files[no]
            data = extract_branch_data(file_info["path"], no)
            branch_data[no] = data
            trainer_rows.extend(extract_pt_ranking(file_info["path"], no))
            print(f"{name:<8} OK")
    except ParserNotReadyError as e:
        print(f"{BRANCHES[no]:<8} FAILED")
        fail(
            "Parser 미구현",
            str(e),
            file_hint=os.path.basename(selected_files[no]["path"]),
        )
    except Exception as e:
        print(f"{BRANCHES[no]:<8} FAILED")
        fail(
            "Excel 파싱 오류",
            f"{BRANCHES[no]}점 파일 처리 중 오류가 발생했습니다.\n{e}",
            file_hint=os.path.basename(selected_files[no]["path"]),
        )
    print()
    print(SUB_LINE)
    print()

    # [VALIDATION]
    print("[VALIDATION]")
    print()
    try:
        cross_warnings = run_all_validations(selected_files, branch_data)
    except ValidationError as e:
        fail("데이터 검증 실패", str(e))

    print("Target       OK")
    print("FC Sales     OK")
    print("PT Sales     OK")
    print("Total Sales  OK")
    print("PT Ranking   OK")
    print()
    for w in cross_warnings:
        print(w)
    if cross_warnings:
        print()
    print(SUB_LINE)
    print()

    # 데이터 기준일: 지점별 상세 거래내역 중 가장 최근 거래일 (Excel 원본 기준)
    last_days = [d["last_transaction_day"] for d in branch_data.values() if d.get("last_transaction_day")]
    reference_date = None
    if last_days:
        reference_date = f"{period_label}-{max(last_days):02d}"

    # [COMPANY]
    dashboard = build_dashboard_json(period_label, branch_data, trainer_rows, reference_date)
    company = dashboard["company"]

    print("[COMPANY]")
    print()
    print("목표")
    print(format_won(company["target_total"]))
    print()
    print("현재 매출")
    print(format_won(company["actual_total"]))
    print()
    print("달성률")
    print(f"{company['rate_total']}%")
    print()
    print("목표까지")
    print(format_won(max(company["target_total"] - company["actual_total"], 0)))
    print()
    print(SUB_LINE)
    print()

    # 백업 및 저장 (검증 통과 후에만)
    backup_existing_index()

    os.makedirs(OUTPUT_DIR, exist_ok=True)
    with open(OUTPUT_JSON, "w", encoding="utf-8") as f:
        json.dump(dashboard, f, ensure_ascii=False, indent=2)
    print("Dashboard JSON")
    print("OK")
    print()

    update_index_html(dashboard)
    print("index.html")
    print("OK")
    print()

    print(LINE)
    print("UPDATE COMPLETE")
    print(LINE)


def update_index_html(dashboard: dict):
    """
    index.html 자체에는 Excel parsing 로직을 넣지 않는다.
    index.html은 output/dashboard_data.json을 fetch로 읽어 렌더링하므로,
    여기서는 JSON 저장만으로 Dashboard 갱신이 반영된다.
    """
    if not os.path.isfile(INDEX_HTML):
        raise FileNotFoundError("index.html 파일이 없습니다.")
    # index.html은 정적 파일을 그대로 유지하고 JSON만 갱신한다.


if __name__ == "__main__":
    main()
