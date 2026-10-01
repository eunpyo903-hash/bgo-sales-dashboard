"""
Google Drive 연동 설정 (2026-10부터 적용)

"1. 매출보고서" 폴더 아래 지점별 하위 폴더에 각 지점 매니저가
"YYMM_b_N_지점명_총매출.xlsx" (지속적으로 덮어써지는 누적 매출 파일)과
"YYMMDD_b_N_지점명_일일매출.xlsx" (매일 추가되는 일일 파일)을 올린다.
우리는 "총매출" 파일만 사용한다.

주의: 이 ID들은 Python 코드가 아니라, Google Drive 커넥터를 가진
Claude 에이전트(예약 작업)가 검색·다운로드할 때 참조하는 값이다.
일반 python 스크립트는 Drive API를 직접 호출하지 않는다.

폴더 구조가 바뀌면(예: 지점 폴더 재생성) 이 파일의 ID만 갱신하면 된다.
"""

DRIVE_REPORT_ROOT_FOLDER_ID = "11sTOVtgfiKoh8ELMeykNQVgb6dvGxBPj"  # "1. 매출보고서"

BRANCH_DRIVE_FOLDER_IDS = {
    1: "1-1i0DxaVEh327mQMw0hHk5F8FHygpfpL",  # 1.역삼점
    2: "1Etylf--mhF5KzCVuq6MlNJX2SbBwhQWw",  # 2.선정릉점
    3: "1h-i28w_lu3WzZWeZ1a8M4THLZp0RSoLD",  # 3.대치점
    4: "19IlLtflEfr5_NXePUfMiOZcjBroj7wJ4",  # 4.송파점
    5: "1NeKMLymBzuM8GCgXKV9xZY9THy6DSIVE",  # 5.장안점
    6: "1RCD1xxrAGKT-O_w5A6HrLCQIvFCBi5Vo",  # 6.가락점
    7: "1jvmMfPBlL_hftmgJE7AiOEbU5fYI2At2",  # 7.삼전점
}
