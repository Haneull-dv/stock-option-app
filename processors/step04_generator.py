"""
Step 04 — 등기신청 서류 생성
"""
import os
import shutil
import glob
from pypdf import PdfWriter, PdfReader

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
COMMON_TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates_common')
ARTICLES_PDF = os.path.join(COMMON_TEMPLATES_DIR, '정관_원본 20260326 일부개정.pdf')


def generate_step04_documents(round_id, applicants, output_dir, templates_dir):
    """
    등기신청 서류 생성

    Args:
        round_id: 회차 ID
        applicants: 신청자 리스트 (sort_order 순)
        output_dir: 출력 폴더
        templates_dir: 템플릿 폴더 (templates_step04)

    Returns:
        dict: 생성된 파일 정보
    """
    print(f"\n=== Step04 서류 생성 시작 ===")
    print(f"출력 폴더: {output_dir}")
    print(f"템플릿 폴더: {templates_dir}")
    print(f"신청자 수: {len(applicants)}명")

    os.makedirs(output_dir, exist_ok=True)

    results = {
        'success': True,
        'files': [],
        'manual_tasks': []
    }

    # ── 1. 수동 준비 서류 안내 ──
    results['manual_tasks'].extend([
        {'name': '스톡옵션 행사 신청서 (원본)'},
        {'name': '주금납입보관증명서 (원본)'},
        {'name': '법인인감증명서'}
    ])

    # ── 2. 주주총회의사록 + 조정내역서 통합 (발행가액 + 부여일 기반) ──
    meeting_pdf_path = os.path.join(output_dir, '주주총회의사록_및_조정내역서_전체.pdf')
    copy_all_meeting_minutes(applicants, meeting_pdf_path)
    print(f"주주총회의사록 파일 존재 여부: {os.path.exists(meeting_pdf_path)}, 경로: {meeting_pdf_path}")
    if os.path.exists(meeting_pdf_path):
        results['files'].append({
            'name': '주주총회의사록 및 조정내역서 (전체)',
            'path': meeting_pdf_path,
            'note': '📍 원본대조필 후 법인인감 날인 및 접어서 각장 간인 필요\n📍 조정내역서는 회사명판과 법인인감 날인 필요'
        })
        print(f"  → files 배열에 추가됨")

    # ── 3. 정관 ──
    articles_dst = os.path.join(output_dir, '정관.pdf')
    print(f"정관 원본 존재 여부: {os.path.exists(ARTICLES_PDF)}, 경로: {ARTICLES_PDF}")
    if os.path.exists(ARTICLES_PDF):
        shutil.copy2(ARTICLES_PDF, articles_dst)
        results['files'].append({
            'name': '정관',
            'path': articles_dst,
            'note': '📍 원본대조필 후 법인인감 날인 및 접어서 각장 간인 필요'
        })
        print(f"  → files 배열에 추가됨")

    # ── 4. 등기위임장 ──
    delegation_src = os.path.join(templates_dir, '등기위임장.hwpx')
    delegation_dst = os.path.join(output_dir, '등기위임장.hwpx')
    print(f"등기위임장 원본 존재 여부: {os.path.exists(delegation_src)}, 경로: {delegation_src}")
    if os.path.exists(delegation_src):
        shutil.copy2(delegation_src, delegation_dst)
        results['files'].append({
            'name': '등기위임장',
            'path': delegation_dst,
            'note': '📍 법인인감 날인 필요'
        })
        print(f"  → files 배열에 추가됨")

    # ── 5. 주식매수선택권 부여계약서 (신청자별 매칭 후 합본) ──
    contract_pdf_path = os.path.join(output_dir, '주식매수선택권_부여계약서_합본.pdf')
    matched_count = merge_grant_contracts(templates_dir, applicants, contract_pdf_path)
    print(f"부여계약서 합본 파일 존재 여부: {os.path.exists(contract_pdf_path)}, 경로: {contract_pdf_path}")
    if os.path.exists(contract_pdf_path):
        results['files'].append({
            'name': f'주식매수선택권 부여계약서 ({matched_count}명)',
            'path': contract_pdf_path,
            'note': '📍 원본대조필 후 법인인감 날인 및 접어서 각장 간인 필요'
        })
        print(f"  → files 배열에 추가됨")

    print(f"\n=== 생성 완료 ===")
    print(f"생성된 파일 수: {len(results['files'])}개")
    print(f"수동 준비 서류: {len(results['manual_tasks'])}개")
    for i, f in enumerate(results['files'], 1):
        print(f"  {i}. {f['name']}")

    return results


def copy_all_meeting_minutes(applicants, output_path):
    """
    신청자들의 (발행가액, 부여일) 기반 주주총회의사록 + 조정산식 합본

    Args:
        applicants: 신청자 리스트 (exercise_price, grant_date 포함)
        output_path: 출력 PDF 경로
    """
    from processors.shareholder_meeting_matcher import get_all_required_meeting_files

    # 필요한 파일들 찾기
    required_files = get_all_required_meeting_files(applicants)

    if not required_files:
        print("  ⚠️ 주주총회의사록/조정산식 파일을 찾을 수 없습니다.")
        return

    # 모든 파일을 하나의 리스트로
    all_pdfs = []
    for (price, grant_date), files in sorted(required_files.items()):
        all_pdfs.extend(files)

    print(f"\n[PDF 합본 시작] 총 {len(all_pdfs)}개 파일")

    # PDF 합본
    writer = PdfWriter()
    for pdf_path in all_pdfs:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)
            print(f"  ✓ {os.path.basename(pdf_path)}")
        except Exception as e:
            print(f"  ✗ PDF 추가 실패: {os.path.basename(pdf_path)}, {e}")

    with open(output_path, 'wb') as f:
        writer.write(f)

    print(f"\n✓ 주주총회의사록 합본 완료: {len(all_pdfs)}개 파일 → {output_path}")


def merge_grant_contracts(templates_dir, applicants, output_path):
    """
    신청자별 부여계약서 매칭 후 PDF 합본 (부여일 기반)

    Args:
        templates_dir: templates_step04 폴더
        applicants: 신청자 리스트 (sort_order 순, grant_date 포함)
        output_path: 출력 PDF 경로

    Returns:
        int: 매칭된 계약서 개수
    """
    # 부여계약서 폴더는 루트에 위치
    project_root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    contract_base_dir = os.path.join(project_root, '부여계약서')

    if not os.path.exists(contract_base_dir):
        print(f"부여계약서 폴더를 찾을 수 없습니다: {contract_base_dir}")
        return 0

    # 신청자별 계약서 매칭 (부여일 + 신청자명)
    matched_contracts = []
    unmatched_applicants = []

    for applicant in applicants:
        name = applicant['name']
        grant_date = applicant.get('grant_date', '').strip()
        matched = False

        if not grant_date:
            print(f"  ⚠️ {name}: 부여일 정보 없음")
            unmatched_applicants.append(f"{name} (부여일 없음)")
            continue

        # 부여일 폴더에서 파일 찾기
        # grant_date 형식: YYYY-MM-DD 또는 YYMMDD 등 다양할 수 있으므로 여러 패턴 시도
        possible_folders = [
            grant_date,  # 2020-03-27
            grant_date.replace('-', ' '),  # 2020 03 27 (공백 구분)
            grant_date.replace('-', ''),  # 20200327
            grant_date[:6] if len(grant_date) >= 6 else grant_date,  # 200327
        ]

        for folder_pattern in possible_folders:
            grant_folder = os.path.join(contract_base_dir, folder_pattern)
            if os.path.exists(grant_folder):
                # 해당 폴더에서 신청자명이 포함된 PDF 찾기
                contract_files = glob.glob(os.path.join(grant_folder, '*.pdf'))
                for contract_path in contract_files:
                    filename = os.path.basename(contract_path)
                    if name in filename:
                        matched_contracts.append((applicant, contract_path))
                        matched = True
                        print(f"  ✓ {name}: {folder_pattern}/{filename}")
                        break
                if matched:
                    break

        if not matched:
            unmatched_applicants.append(f"{name} (부여일: {grant_date})")
            print(f"  ✗ {name}: 부여일 {grant_date} 폴더에서 파일을 찾을 수 없음")

    if unmatched_applicants:
        print(f"\n⚠️ 부여계약서를 찾을 수 없는 신청자 ({len(unmatched_applicants)}명):")
        for name in unmatched_applicants:
            print(f"  - {name}")

    if not matched_contracts:
        print("매칭된 부여계약서가 없습니다.")
        return 0

    # PDF 합본 (신청자 순서대로)
    writer = PdfWriter()
    for applicant, contract_path in matched_contracts:
        try:
            reader = PdfReader(contract_path)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"PDF 추가 실패: {applicant['name']}, {contract_path}, {e}")

    with open(output_path, 'wb') as f:
        writer.write(f)

    print(f"부여계약서 합본 완료: {len(matched_contracts)}명")
    return len(matched_contracts)
