"""
비고컴퍼니 매출현황 - Excel Parser

책임:
  1) data/ 폴더에서 최신 월(YYMM)의 지점별 Excel 파일을 자동으로 찾는다.
  2) 각 지점 Excel 파일에서 표준 데이터 항목을 추출한다.

중요:
  extract_branch_data() / extract_pt_ranking() 의 실제 셀 매핑은
  아직 실제 Excel 원본 파일을 분석하기 전이므로 구현되어 있지 않다.
  실제 파일 7개가 data/ 폴더에 들어오면, 구조를 먼저 분석한 뒤
  이 함수들을 실제 셀 주소 기준으로 완성한다. (셀 주소 추측 금지)
"""

import os
import re
from collections import defaultdict

from openpyxl import load_workbook

from .utils import BRANCHES, parse_filename


class DuplicateFileWarning(Exception):
    """중복 파일 경고를 나타내기 위한 용도 (raise 하지 않고 메시지 수집용으로도 사용 가능)"""
    pass


class ParserNotReadyError(Exception):
    """실제 Excel 구조 분석 전이라 파싱 로직이 아직 완성되지 않았음을 알리는 예외"""
    pass


def scan_data_dir(data_dir: str):
    """
    data_dir 안의 모든 xlsx 파일명을 정규식으로 분석하여
    {yymm: {branch_no: [file_info, ...]}} 형태로 반환한다.
    """
    found = defaultdict(lambda: defaultdict(list))

    if not os.path.isdir(data_dir):
        return found

    for fname in os.listdir(data_dir):
        if not fname.lower().endswith((".xlsx", ".xlsm")):
            continue
        if fname.startswith("~$"):  # Excel 임시 잠금 파일 제외
            continue

        parsed = parse_filename(fname)
        if not parsed:
            continue

        full_path = os.path.join(data_dir, fname)
        found[parsed["yymm"]][parsed["branch_no"]].append({
            "filename": fname,
            "path": full_path,
            "branch_name_in_file": parsed["branch_name_in_file"],
            "mtime": os.path.getmtime(full_path),
        })

    return found


def select_latest_files(data_dir: str, warnings: list = None):
    """
    가장 최신 YYMM을 자동 선택하고, 지점별로 파일 하나씩 확정한다.

    같은 월/같은 지점에 파일이 여러 개면:
      1차: 파일명 정확도 (표준 패턴에 더 정확히 맞는 파일 우선)
      2차: modified time (최신 수정 파일 우선)
    을 기준으로 선택하고, 중복이 있었다는 WARNING을 warnings 리스트에 남긴다.

    반환: (period_yymm, {branch_no: file_info}) 또는 (None, {}) - 파일이 전혀 없을 때
    """
    if warnings is None:
        warnings = []

    scanned = scan_data_dir(data_dir)
    if not scanned:
        return None, {}

    latest_yymm = sorted(scanned.keys())[-1]
    branch_files = scanned[latest_yymm]

    selected = {}
    for branch_no, candidates in branch_files.items():
        if len(candidates) == 1:
            selected[branch_no] = candidates[0]
            continue

        # 1차: 표준 파일명 패턴과의 정확도 우선 정렬, 2차: mtime 최신
        def sort_key(c):
            standard_name = f"{latest_yymm}_b_{branch_no}_{BRANCHES[branch_no]}점_총매출"
            exact_match = 1 if c["filename"].startswith(standard_name) else 0
            return (exact_match, c["mtime"])

        candidates_sorted = sorted(candidates, key=sort_key, reverse=True)
        selected[branch_no] = candidates_sorted[0]

        other_names = ", ".join(c["filename"] for c in candidates_sorted[1:])
        warnings.append(
            f"WARNING: {latest_yymm} {BRANCHES[branch_no]}점 파일이 {len(candidates)}개 발견되어 "
            f"'{candidates_sorted[0]['filename']}' 파일을 선택했습니다. (제외: {other_names})"
        )

    return latest_yymm, selected


def extract_branch_data(file_path: str, branch_no: int) -> dict:
    """
    지점 Excel 1개에서 표준 데이터 항목을 추출한다.

    TODO (실제 Excel 파일 분석 후 구현):
      - Sheet 이름 / 개수 확인
      - 목표(target_fc/pt/total) 셀 위치
      - 실적(actual_fc/pt/total) 셀 위치
      - FC/PT 신규·재등록 건수/매출 위치
      - 결제 데이터(card/cash/account/deduction/real_cash) 위치
      - 병합 셀, 수식 결과(load_workbook(..., data_only=True)) 처리

    실제 파일이 제공되기 전까지는 셀 주소를 추측하지 않기 위해
    의도적으로 미구현 상태로 남겨둔다.
    """
    raise ParserNotReadyError(
        f"'{os.path.basename(file_path)}' 파싱 로직이 아직 구현되지 않았습니다. "
        "실제 Excel 파일 구조를 분석한 뒤 extract_branch_data()를 완성해야 합니다."
    )


def extract_pt_ranking(file_path: str, branch_no: int) -> list:
    """
    "매출순위" (또는 해당하는) Sheet에서 트레이너별 [name, branch, sales]를 추출한다.

    TODO (실제 Excel 파일 분석 후 구현):
      - 매출순위 Sheet명 확인
      - 트레이너명 컬럼 위치
      - 개인 PT 매출 컬럼 위치
    """
    raise ParserNotReadyError(
        f"'{os.path.basename(file_path)}' PT 개인매출 순위 파싱 로직이 아직 구현되지 않았습니다."
    )


def open_workbook(file_path: str, data_only: bool = True):
    """수식 결과값을 읽을 때는 data_only=True를 기본으로 사용한다."""
    return load_workbook(filename=file_path, data_only=data_only)
