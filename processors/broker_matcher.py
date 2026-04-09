"""
증권사 코드 매칭 유틸리티.
code.xlsx를 읽어 증권사명 → 코드 매핑 딕셔너리 생성.
"""
import openpyxl
import os

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
CODE_XLSX_PATH = os.path.join(BASE_DIR, 'code.xlsx')

# 캐싱용 전역 변수
_BROKER_CODE_MAP = None


def load_broker_codes():
    """code.xlsx에서 {증권사명: 코드} 딕셔너리 로드 (캐싱)."""
    global _BROKER_CODE_MAP
    if _BROKER_CODE_MAP is not None:
        return _BROKER_CODE_MAP

    if not os.path.isfile(CODE_XLSX_PATH):
        print(f"[Warning] code.xlsx not found at {CODE_XLSX_PATH}")
        _BROKER_CODE_MAP = {}
        return _BROKER_CODE_MAP

    wb = openpyxl.load_workbook(CODE_XLSX_PATH)
    ws = wb.active
    broker_map = {}

    for row in ws.iter_rows(min_row=2, values_only=True):  # 첫 행은 헤더
        code, name = row[0], row[1]
        if code and name:
            # code가 숫자면 int로 변환, 문자면 그대로
            if isinstance(code, (int, float)):
                code = int(code)
            broker_map[str(name).strip()] = str(code)

    _BROKER_CODE_MAP = broker_map
    return _BROKER_CODE_MAP


def match_broker_code(broker_name):
    """
    증권사명으로 코드 검색.
    완전일치 시도 → 부분일치 시도 → None 반환.
    """
    if not broker_name:
        return None

    broker_map = load_broker_codes()
    broker_name = str(broker_name).strip()

    # 1. 완전일치
    if broker_name in broker_map:
        return broker_map[broker_name]

    # 2. 부분일치 (증권사명에 키워드 포함)
    for name, code in broker_map.items():
        if broker_name in name or name in broker_name:
            return code

    return None
