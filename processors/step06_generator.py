"""
Step06 KIND 상장신청 문서 자동 생성.

결과물: ZIP 파일
- 정관
- 상장수수료 납부영수증
- 법인등기부등본
- 주주총회의사록 및 조정산식/ (발행가액별 폴더)
- 의무보유증명서 및 의무보유청구내역/
- 발행등록사실확인서/ (발행가액별)
- 주금납입금 보관증명서/ (발행가액별)
- 금융거래정보제공동의서
- 주식매수선택권 행사신청서 (마스킹)
- 의무보유확약서/ (대상자별)
- 스톡옵션 행사현황표
- 기타비공개첨부서류/
    - 재직증명서/
    - 주식매수선택권부여계약서/
    - 법인인감증명서.pdf
"""
import os
import shutil
import sqlite3
from datetime import datetime
from PyPDF2 import PdfReader, PdfWriter, PdfMerger
from processors.zip_utils import create_zip_from_folder
from processors.docx_to_pdf import convert_docx_to_pdf
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates_step06')
COMMON_TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates_common')
STEP04_TEMPLATES_DIR = os.path.join(BASE_DIR, 'templates_step04')
DB_PATH = os.path.join(BASE_DIR, 'data', 'stockops.db')


def generate_step06_zip(round_obj, config, output_base_dir):
    """
    KIND 상장신청 서류 ZIP 생성.

    Args:
        round_obj: 회차 정보 dict
        config: step06_config dict
        output_base_dir: 출력 기본 디렉토리

    Returns:
        dict: {
            'success': bool,
            'zip_path': str,
            'message': str,
            'warnings': list
        }
    """
    round_id = round_obj['id']
    round_name = round_obj['name']
    warnings = []

    # 출력 폴더 생성
    step06_dir = os.path.join(output_base_dir, f"step06")
    os.makedirs(step06_dir, exist_ok=True)

    work_folder = os.path.join(step06_dir, f"KIND_상장신청_{round_name}")
    if os.path.exists(work_folder):
        shutil.rmtree(work_folder)
    os.makedirs(work_folder)

    print(f"\n[Step06] KIND 상장신청 서류 생성 중...")
    print(f"출력 폴더: {work_folder}\n")

    try:
        # 1. 정관
        print("[1/13] 정관 복사 중...")
        _copy_articles(work_folder)

        # 2. 상장수수료 납부영수증
        print("[2/13] 상장수수료 납부영수증 복사 중...")
        if config.get('listing_fee_receipt'):
            _copy_uploaded_file(config['listing_fee_receipt'], work_folder, '상장수수료 납부영수증')
        else:
            warnings.append("상장수수료 납부영수증이 업로드되지 않았습니다.")

        # 3. 법인등기부등본
        print("[3/13] 법인등기부등본 복사 중...")
        _copy_corporate_registry(work_folder)

        # 4. 주주총회의사록 및 조정산식
        print("[4/13] 주주총회의사록 및 조정산식 복사 중...")
        _copy_shareholder_meeting_minutes(round_id, work_folder, warnings)

        # 5. 의무보유증명서 및 의무보유청구내역
        print("[5/13] 의무보유증명서 및 의무보유청구내역 복사 중...")
        if config.get('holding_proof_folder'):
            _copy_uploaded_folder(config['holding_proof_folder'], work_folder, '의무보유증명서 및 의무보유청구내역')
        else:
            warnings.append("의무보유증명서 및 의무보유청구내역이 업로드되지 않았습니다.")

        # 6. 발행등록사실확인서
        print("[6/13] 발행등록사실확인서 복사 중...")
        _copy_issuance_confirmations(round_id, work_folder, warnings)

        # 7. 주금납입금 보관증명서
        print("[7/13] 주금납입금 보관증명서 복사 중...")
        _copy_deposit_certificates(round_id, work_folder, warnings)

        # 8. 금융거래정보제공동의서
        print("[8/13] 금융거래정보제공동의서 생성 중...")
        _generate_financial_consent(config, work_folder)

        # 9. 주식매수선택권 행사신청서 (마스킹)
        print("[9/13] 주식매수선택권 행사신청서 (마스킹) 생성 중...")
        _generate_masked_applications(round_id, work_folder, warnings)

        # 10. 의무보유확약서
        print("[10/13] 의무보유확약서 복사 중...")
        _copy_holding_commitments(round_id, work_folder, warnings)

        # 11. 스톡옵션 행사현황표
        print("[11/13] 스톡옵션 행사현황표 생성 중...")
        _generate_exercise_summary(round_id, config, work_folder, warnings)

        # 12. 기타비공개첨부서류
        print("[12/13] 기타비공개첨부서류 생성 중...")
        _generate_other_documents(round_id, config, work_folder, warnings)

        # 13. ZIP 파일 생성
        print("[13/13] ZIP 파일 생성 중...")
        zip_filename = f"KIND_상장신청_{round_name}.zip"
        zip_path = create_zip_from_folder(work_folder, os.path.join(step06_dir, zip_filename))

        # 작업 폴더 삭제
        shutil.rmtree(work_folder)

        print(f"\n✅ Step06 생성 완료: {zip_filename}")
        if warnings:
            print(f"\n⚠️ 경고 {len(warnings)}개:")
            for w in warnings:
                print(f"  - {w}")

        return {
            'success': True,
            'zip_path': zip_path,
            'message': f'ZIP 파일 생성 완료: {zip_filename}',
            'warnings': warnings
        }

    except Exception as e:
        print(f"\n❌ Step06 생성 실패: {e}")
        import traceback
        traceback.print_exc()
        return {
            'success': False,
            'message': f'오류 발생: {str(e)}',
            'warnings': warnings
        }


def _copy_articles(work_folder):
    """정관 복사."""
    src = os.path.join(COMMON_TEMPLATES_DIR, '정관_원본 20260326 일부개정.pdf')
    dest = os.path.join(work_folder, '정관.pdf')
    shutil.copy(src, dest)
    print(f"  ✓ 정관 복사 완료")


def _copy_uploaded_file(file_path, work_folder, display_name):
    """업로드된 파일 복사."""
    if os.path.exists(file_path):
        ext = os.path.splitext(file_path)[1]
        dest = os.path.join(work_folder, f"{display_name}{ext}")
        shutil.copy(file_path, dest)
        print(f"  ✓ {display_name} 복사 완료")
    else:
        print(f"  ⚠️ {display_name} 파일을 찾을 수 없습니다: {file_path}")


def _copy_uploaded_folder(folder_path, work_folder, display_name):
    """업로드된 폴더 복사."""
    if os.path.exists(folder_path):
        dest = os.path.join(work_folder, display_name)
        shutil.copytree(folder_path, dest)
        print(f"  ✓ {display_name} 폴더 복사 완료")
    else:
        print(f"  ⚠️ {display_name} 폴더를 찾을 수 없습니다: {folder_path}")


def _copy_corporate_registry(work_folder):
    """법인등기부등본 복사 (260408 최신)."""
    src = os.path.join(COMMON_TEMPLATES_DIR, '(붙임9) 법인등기부등본_에스투더블유 260408.pdf')
    dest = os.path.join(work_folder, '법인등기부등본.pdf')
    shutil.copy(src, dest)
    print(f"  ✓ 법인등기부등본 복사 완료")


def _copy_shareholder_meeting_minutes(round_id, work_folder, warnings):
    """주주총회의사록 및 조정산식 복사 (Step05 붙임4 폴더)."""
    # Step05 출력 폴더에서 각 발행가액별 붙임4 폴더 찾기
    step05_output = os.path.join(BASE_DIR, 'data', 'outputs', str(round_id), 'step05')

    if not os.path.exists(step05_output):
        warnings.append("Step05 출력 폴더가 없습니다. 먼저 Step05를 완료해주세요.")
        return

    dest_folder = os.path.join(work_folder, '주주총회의사록 및 조정산식')
    os.makedirs(dest_folder, exist_ok=True)

    # Step05에서 생성된 ZIP 파일들을 풀어서 붙임4 폴더 추출
    # 또는 templates_step05/price_specific/에서 직접 복사
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('SELECT DISTINCT exercise_price FROM applicants WHERE round_id = ? ORDER BY exercise_price', (round_id,))
    prices = [row[0] for row in c.fetchall()]
    conn.close()

    found_count = 0
    for price in prices:
        # templates_step05/price_specific/{price}/(붙임4) 폴더 찾기
        price_folder = os.path.join(BASE_DIR, 'templates_step05', 'price_specific', str(price))
        attachment4_folders = [
            d for d in os.listdir(price_folder)
            if os.path.isdir(os.path.join(price_folder, d)) and '붙임4' in d
        ]

        if attachment4_folders:
            src_folder = os.path.join(price_folder, attachment4_folders[0])
            dest_price_folder = os.path.join(dest_folder, f"주주총회의사록_{price}원")
            shutil.copytree(src_folder, dest_price_folder)
            found_count += 1
            print(f"  ✓ {price}원 주주총회의사록 복사 완료")

    if found_count == 0:
        warnings.append("주주총회의사록 폴더를 찾을 수 없습니다.")
    else:
        print(f"  ✓ 총 {found_count}개 발행가액 주주총회의사록 복사 완료")


def _copy_issuance_confirmations(round_id, work_folder, warnings):
    """발행등록사실확인서 복사 (발행가액별)."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT exercise_price, file_path, original_name
        FROM step06_issuance_confirmations
        WHERE round_id = ?
        ORDER BY exercise_price
    ''', (round_id,))
    rows = c.fetchall()
    conn.close()

    if not rows:
        warnings.append("발행등록사실확인서가 업로드되지 않았습니다.")
        return

    dest_folder = os.path.join(work_folder, '발행등록사실확인서')
    os.makedirs(dest_folder, exist_ok=True)

    for price, file_path, original_name in rows:
        if os.path.exists(file_path):
            ext = os.path.splitext(original_name)[1]
            dest = os.path.join(dest_folder, f"발행등록사실확인서_{price}원{ext}")
            shutil.copy(file_path, dest)
            print(f"  ✓ {price}원 발행등록사실확인서 복사 완료")
        else:
            warnings.append(f"{price}원 발행등록사실확인서 파일을 찾을 수 없습니다.")


def _copy_deposit_certificates(round_id, work_folder, warnings):
    """주금납입금 보관증명서 복사 (Step05 attachment8)."""
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT exercise_price, file_path, original_name
        FROM attachment8
        WHERE round_id = ?
        ORDER BY exercise_price
    ''', (round_id,))
    rows = c.fetchall()
    conn.close()

    if not rows:
        warnings.append("주금납입금 보관증명서가 업로드되지 않았습니다. (Step05에서 붙임8 업로드 필요)")
        return

    dest_folder = os.path.join(work_folder, '주금납입금 보관증명서')
    os.makedirs(dest_folder, exist_ok=True)

    for price, file_path, original_name in rows:
        if os.path.exists(file_path):
            ext = os.path.splitext(original_name)[1]
            dest = os.path.join(dest_folder, f"주금납입금보관증명서_{price}원{ext}")
            shutil.copy(file_path, dest)
            print(f"  ✓ {price}원 주금납입금 보관증명서 복사 완료")
        else:
            warnings.append(f"{price}원 주금납입금 보관증명서 파일을 찾을 수 없습니다.")


def _generate_financial_consent(config, work_folder):
    """금융거래정보제공동의서 생성 (날짜 치환)."""
    # 간단한 복사만 (실제 날짜 치환은 추후 구현)
    src = os.path.join(TEMPLATES_DIR, '금융거래정보제공동의서.pdf')
    dest = os.path.join(work_folder, '금융거래정보제공동의서.pdf')

    if os.path.exists(src):
        shutil.copy(src, dest)
        print(f"  ✓ 금융거래정보제공동의서 복사 완료")
        print(f"  ⚠️ 날짜(년 월 일) 직접 기입 후 법인인감 날인 필요!")
    else:
        print(f"  ⚠️ 금융거래정보제공동의서 템플릿을 찾을 수 없습니다: {src}")


def _generate_masked_applications(round_id, work_folder, warnings):
    """주식매수선택권 행사신청서 합본 + 마스킹."""
    # Step01에서 업로드된 신청서들 합치기
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''
        SELECT d.file_path, a.name
        FROM documents d
        JOIN applicants a ON d.applicant_id = a.id
        WHERE a.round_id = ? AND d.doc_type = 'application'
        ORDER BY a.id
    ''', (round_id,))
    rows = c.fetchall()
    conn.close()

    if not rows:
        warnings.append("행사신청서가 업로드되지 않았습니다.")
        return

    # PDF 합본
    merger = PdfMerger()
    for file_path, name in rows:
        if os.path.exists(file_path):
            merger.append(file_path)

    # 임시 합본 파일
    temp_merged = os.path.join(work_folder, '_temp_merged.pdf')
    merger.write(temp_merged)
    merger.close()

    # 마스킹 처리 (현재는 간단히 복사만)
    # TODO: 실제 마스킹은 OCR + 좌표 검출 필요
    dest = os.path.join(work_folder, '주식매수선택권 행사신청서 (마스킹).pdf')
    shutil.copy(temp_merged, dest)
    os.remove(temp_merged)

    print(f"  ✓ 행사신청서 {len(rows)}개 합본 완료")
    print(f"  ⚠️ 주민번호 뒷자리, 연락처, 주소 마스킹이 필요합니다.")
    print(f"  ⚠️ 현재 원본 그대로 포함되어 있으므로, 생성 후 수동으로 마스킹 처리해주세요.")


def _copy_holding_commitments(round_id, work_folder, warnings):
    """의무보유확약서 복사 (Step03-3 결과물 1,2페이지만)."""
    # Step03-3 출력 폴더
    step033_output = os.path.join(BASE_DIR, 'data', 'outputs', str(round_id), 'step033')

    if not os.path.exists(step033_output):
        warnings.append("Step03-3이 완료되지 않았습니다. 의무보유확약서를 생성할 수 없습니다.")
        return

    # 의무보유확약서 파일 찾기
    docx_files = [f for f in os.listdir(step033_output) if '의무보유확약서' in f and f.endswith('.docx')]

    if not docx_files:
        warnings.append("의무보유확약서 파일을 찾을 수 없습니다.")
        return

    dest_folder = os.path.join(work_folder, '의무보유확약서')
    os.makedirs(dest_folder, exist_ok=True)

    # DOCX → PDF 변환 후 1,2페이지만 추출
    for docx_file in docx_files:
        src_docx = os.path.join(step033_output, docx_file)
        temp_pdf = os.path.join(dest_folder, '_temp.pdf')

        # DOCX → PDF 변환
        if convert_docx_to_pdf(src_docx, temp_pdf):
            # 1,2페이지만 추출
            reader = PdfReader(temp_pdf)
            writer = PdfWriter()

            # 최대 2페이지까지만
            for i in range(min(2, len(reader.pages))):
                writer.add_page(reader.pages[i])

            # 저장
            base_name = os.path.splitext(docx_file)[0]
            dest_pdf = os.path.join(dest_folder, f"{base_name}_1-2페이지.pdf")
            with open(dest_pdf, 'wb') as f:
                writer.write(f)

            os.remove(temp_pdf)
            print(f"  ✓ {base_name} 1,2페이지 추출 완료")

    print(f"  ⚠️ 개인 날인 페이지와 개인 인감증명서는 별도 스캔하여 추가해주세요.")


def _generate_exercise_summary(round_id, config, work_folder, warnings):
    """스톡옵션 행사현황표 생성 (엑셀 → PDF)."""
    # config에 엑셀 파일 경로가 있어야 함
    excel_path = config.get('exercise_summary_excel')

    if not excel_path or not os.path.exists(excel_path):
        warnings.append("스톡옵션 행사현황표 엑셀 파일이 업로드되지 않았습니다.")
        return

    # TODO: 엑셀 → PDF 변환 (openpyxl + reportlab)
    # 현재는 간단히 복사만
    dest = os.path.join(work_folder, '스톡옵션 행사현황표.xlsx')
    shutil.copy(excel_path, dest)

    print(f"  ✓ 스톡옵션 행사현황표 복사 완료")
    print(f"  ⚠️ 회사명판 및 법인인감 날인 필요!")
    print(f"  ⚠️ 전체현황 시트와 행사내역_누적 시트를 PDF로 변환해주세요.")


def _generate_other_documents(round_id, config, work_folder, warnings):
    """기타비공개첨부서류 폴더 생성."""
    dest_folder = os.path.join(work_folder, '기타비공개첨부서류')
    os.makedirs(dest_folder, exist_ok=True)

    # 1. 재직증명서 폴더
    if config.get('employment_cert_folder'):
        _copy_uploaded_folder(
            config['employment_cert_folder'],
            dest_folder,
            '재직증명서'
        )
    else:
        warnings.append("재직증명서 폴더가 업로드되지 않았습니다.")

    # 2. 주식매수선택권 부여계약서 폴더
    contracts_src = os.path.join(STEP04_TEMPLATES_DIR, '부여계약서')
    if os.path.exists(contracts_src):
        contracts_dest = os.path.join(dest_folder, '주식매수선택권부여계약서')
        shutil.copytree(contracts_src, contracts_dest)
        print(f"  ✓ 주식매수선택권부여계약서 폴더 복사 완료")
    else:
        warnings.append("주식매수선택권부여계약서 폴더를 찾을 수 없습니다.")

    # 3. 법인인감증명서
    cert_src = os.path.join(COMMON_TEMPLATES_DIR, '법인인감증명서.pdf')
    cert_dest = os.path.join(dest_folder, '법인인감증명서.pdf')
    if os.path.exists(cert_src):
        shutil.copy(cert_src, cert_dest)
        print(f"  ✓ 법인인감증명서 복사 완료")
    else:
        warnings.append("법인인감증명서를 찾을 수 없습니다.")

    print(f"  ✓ 기타비공개첨부서류 폴더 생성 완료")
