"""
주주총회의사록 및 조정산식 매칭 유틸리티

신청자들의 (발행가액, 부여일) 조합으로 필요한 파일들을 찾음
"""
import os
import glob
from collections import defaultdict


def get_project_root():
    """프로젝트 루트 경로 반환"""
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def get_shareholder_meeting_folder():
    """주주총회의사록 및 조정산식 폴더 경로"""
    return os.path.join(get_project_root(), '주주총회의사록 및 조정산식')


def group_applicants_by_price_and_grant_date(applicants):
    """
    신청자들을 (발행가액, 부여일) 조합으로 그룹핑

    Returns:
        dict: {
            (price, grant_date): [applicant1, applicant2, ...],
            ...
        }
    """
    groups = defaultdict(list)

    for ap in applicants:
        price = ap.get('exercise_price')
        grant_date = ap.get('grant_date', '').strip()

        if not price or not grant_date:
            continue

        # 부여일 정규화: YYYY-MM-DD → YYYYMMDD
        grant_date_normalized = grant_date.replace('-', '').replace(' ', '')

        groups[(price, grant_date_normalized)].append(ap)

    return dict(groups)


def find_meeting_files(price, grant_date):
    """
    발행가액과 부여일에 해당하는 주주총회의사록 + 조정산식 파일 찾기

    Args:
        price: 발행가액 (int)
        grant_date: 부여일 YYYYMMDD (str)

    Returns:
        list: 찾은 PDF 파일 경로 리스트 (보통 2개: 의사록 + 조정산식)
    """
    base_folder = get_shareholder_meeting_folder()
    price_folder = os.path.join(base_folder, str(price))

    if not os.path.exists(price_folder):
        print(f"  ⚠️ 발행가액 {price}원 폴더를 찾을 수 없습니다: {price_folder}")
        return []

    # 부여일로 시작하는 파일 찾기
    pattern = os.path.join(price_folder, f'{grant_date}*.pdf')
    files = glob.glob(pattern)

    if not files:
        print(f"  ⚠️ {price}원 / {grant_date} 파일을 찾을 수 없습니다")
        return []

    # 세트 검증: 의사록 + 조정산식 모두 있는지 확인
    has_minutes = False  # 주주총회의사록
    has_adjustment = False  # 조정산식

    for f in files:
        filename = os.path.basename(f)
        if '의사록' in filename:
            has_minutes = True
        if '조정산식' in filename:
            has_adjustment = True

    # 검증 및 경고
    if not has_minutes or not has_adjustment:
        print(f"  ⚠️ {price}원 / {grant_date}: 불완전한 세트!")
        if not has_minutes:
            print(f"      ✗ 주주총회의사록 없음 - 추가 필요!")
        if not has_adjustment:
            print(f"      ✗ 조정산식 없음 - 추가 필요!")
        # 있는 파일만 표시
        for f in files:
            print(f"    - {os.path.basename(f)}")
        print(f"      → 파일을 추가하거나 부여일 정보를 확인하세요")
    else:
        print(f"  ✓ {price}원 / {grant_date}: 세트 완전")
        for f in files:
            filename = os.path.basename(f)
            if '의사록' in filename:
                print(f"    1. {filename}")
            elif '조정산식' in filename:
                print(f"    2. {filename}")

    # 정렬: 의사록 먼저, 조정산식 나중에
    minutes = [f for f in files if '의사록' in os.path.basename(f)]
    adjustments = [f for f in files if '조정산식' in os.path.basename(f)]
    return minutes + adjustments


def get_all_required_meeting_files(applicants):
    """
    신청자들에게 필요한 모든 주주총회의사록/조정산식 파일 찾기

    Returns:
        dict: {
            (price, grant_date): [file1.pdf, file2.pdf],
            ...
        }
    """
    groups = group_applicants_by_price_and_grant_date(applicants)

    print(f"\n[주주총회의사록/조정산식 매칭]")
    print(f"  총 {len(groups)}개 (발행가액, 부여일) 조합")

    result = {}

    for (price, grant_date), group_applicants in groups.items():
        print(f"\n  [{price}원 / {grant_date}] - {len(group_applicants)}명")
        files = find_meeting_files(price, grant_date)
        if files:
            result[(price, grant_date)] = files

    return result
