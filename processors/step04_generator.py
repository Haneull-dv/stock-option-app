"""
Step 04 — 등기신청 서류 생성
"""
import os
import shutil
import glob
from pypdf import PdfWriter, PdfReader


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

    # ── 2. 주주총회의사록 + 조정내역서 통합 (모든 발행가액) ──
    meeting_pdf_path = os.path.join(output_dir, '주주총회의사록_및_조정내역서_전체.pdf')
    copy_all_meeting_minutes(templates_dir, meeting_pdf_path)
    print(f"주주총회의사록 파일 존재 여부: {os.path.exists(meeting_pdf_path)}, 경로: {meeting_pdf_path}")
    if os.path.exists(meeting_pdf_path):
        results['files'].append({
            'name': '주주총회의사록 및 조정내역서 (전체)',
            'path': meeting_pdf_path,
            'note': '📍 원본대조필 수 법인인감 날인 및 접어서 각장 간인 필요\n📍 조정내역서는 회사명판과 법인인감 날인 필요'
        })
        print(f"  → files 배열에 추가됨")

    # ── 3. 정관 ──
    articles_src = os.path.join(templates_dir, '정관.pdf')
    articles_dst = os.path.join(output_dir, '정관.pdf')
    print(f"정관 원본 존재 여부: {os.path.exists(articles_src)}, 경로: {articles_src}")
    if os.path.exists(articles_src):
        shutil.copy2(articles_src, articles_dst)
        results['files'].append({
            'name': '정관',
            'path': articles_dst,
            'note': '📍 원본대조필 수 법인인감 날인 및 접어서 각장 간인 필요'
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
            'note': '📍 원본대조필 수 법인인감 날인 및 접어서 각장 간인 필요'
        })
        print(f"  → files 배열에 추가됨")

    print(f"\n=== 생성 완료 ===")
    print(f"생성된 파일 수: {len(results['files'])}개")
    print(f"수동 준비 서류: {len(results['manual_tasks'])}개")
    for i, f in enumerate(results['files'], 1):
        print(f"  {i}. {f['name']}")

    return results


def copy_all_meeting_minutes(templates_dir, output_path):
    """
    모든 발행가액의 주주총회의사록 + 조정내역서를 하나의 PDF로 합침

    templates_step05/price_specific/{가액}/(붙임4) 주주총회의사록 {가액}원/ 폴더들에서
    모든 PDF 파일을 가져와서 합본
    """
    price_dirs = glob.glob(
        os.path.join(templates_dir, '..', 'templates_step05', 'price_specific', '*', '(붙임4) 주주총회의사록*')
    )

    all_pdfs = []
    for price_dir in sorted(price_dirs):
        pdf_files = glob.glob(os.path.join(price_dir, '*.pdf'))
        all_pdfs.extend(sorted(pdf_files))

    if not all_pdfs:
        print("주주총회의사록 PDF 파일을 찾을 수 없습니다.")
        return

    # PDF 합본
    writer = PdfWriter()
    for pdf_path in all_pdfs:
        try:
            reader = PdfReader(pdf_path)
            for page in reader.pages:
                writer.add_page(page)
        except Exception as e:
            print(f"PDF 추가 실패: {pdf_path}, {e}")

    with open(output_path, 'wb') as f:
        writer.write(f)

    print(f"주주총회의사록 합본 완료: {len(all_pdfs)}개 파일")


def merge_grant_contracts(templates_dir, applicants, output_path):
    """
    신청자별 부여계약서 매칭 후 PDF 합본

    Args:
        templates_dir: templates_step04 폴더
        applicants: 신청자 리스트 (sort_order 순)
        output_path: 출력 PDF 경로

    Returns:
        int: 매칭된 계약서 개수
    """
    contract_dir = os.path.join(templates_dir, '부여계약서')
    if not os.path.exists(contract_dir):
        print("부여계약서 폴더를 찾을 수 없습니다.")
        return 0

    # 모든 부여계약서 파일 목록
    contract_files = glob.glob(os.path.join(contract_dir, '*.pdf'))

    # 신청자별 계약서 매칭
    matched_contracts = []
    unmatched_applicants = []

    for applicant in applicants:
        name = applicant['name']
        matched = False

        for contract_path in contract_files:
            filename = os.path.basename(contract_path)
            if name in filename:
                matched_contracts.append((applicant, contract_path))
                matched = True
                break

        if not matched:
            unmatched_applicants.append(name)

    if unmatched_applicants:
        print(f"⚠️ 부여계약서를 찾을 수 없는 신청자: {', '.join(unmatched_applicants)}")

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
