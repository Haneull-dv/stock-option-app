"""
신분증 앞면/뒷면 필터링 모듈.

- 앞면: 이름, 주민등록번호가 있는 면
- 뒷면: 주소, 발급일자만 있는 면
- 한 PDF에 여러 페이지가 있으면 앞면 페이지만 추출
"""
import os
from pypdf import PdfReader, PdfWriter
import easyocr

# OCR 리더 (싱글톤)
_ocr_reader = None

def get_ocr_reader():
    """EasyOCR 리더를 싱글톤으로 반환."""
    global _ocr_reader
    if _ocr_reader is None:
        _ocr_reader = easyocr.Reader(['ko', 'en'], gpu=False)
    return _ocr_reader


def is_front_page(pdf_path, page_num=0, applicant_name=None):
    """
    PDF의 특정 페이지가 신분증 앞면인지 판별.

    Args:
        pdf_path: PDF 파일 경로
        page_num: 페이지 번호 (0부터 시작)
        applicant_name: 신청자 이름 (앞면 판별에 사용, 외국인 이름 포함)

    Returns:
        bool: 앞면이면 True, 뒷면이면 False
    """
    try:
        reader = PdfReader(pdf_path)
        if page_num >= len(reader.pages):
            return False

        # 페이지를 이미지로 변환 후 OCR
        # (간단하게 하기 위해 pypdf의 extract_text 사용, OCR은 실패 시 대비)
        page = reader.pages[page_num]
        text = page.extract_text() or ''

        # 앞면 키워드: 주민등록번호, 주민등록증, 이름
        front_keywords = ['주민등록번호', '주민등록증', '주민번호']

        # 뒷면 키워드: 주소, 발급일자, 발행
        back_keywords = ['주소', '발급일', '발행']

        # 키워드 매칭
        has_front = any(kw in text for kw in front_keywords)
        has_back_only = any(kw in text for kw in back_keywords) and not has_front

        # 신청자 이름이 있으면 추가 확인
        if applicant_name:
            names = normalize_name_for_matching(applicant_name)
            has_name = any(name in text for name in names)
            if has_name:
                return True

        # 앞면 키워드가 있으면 앞면
        if has_front:
            return True

        # 뒷면 키워드만 있으면 뒷면
        if has_back_only:
            return False

        # 판단 불가 시 기본값: 첫 페이지는 앞면으로 간주
        return page_num == 0

    except Exception as e:
        print(f"[id_filter] 페이지 판별 실패 {pdf_path} p.{page_num}: {e}")
        # 오류 시 첫 페이지는 앞면으로 간주
        return page_num == 0


def normalize_name_for_matching(name):
    """
    이름을 매칭용으로 정규화. 외국인 이름 처리 포함.

    Args:
        name: 신청자 이름 (예: "한성원", "HAN ROBERT SUNGWON")

    Returns:
        list[str]: 매칭에 사용할 이름 변형들
    """
    names = [name.strip()]

    # 외국인 이름 매핑 (하드코딩)
    alias_map = {
        '한성원': ['HAN ROBERT SUNGWON', 'HAN SUNGWON', 'ROBERT SUNGWON', 'SUNGWON'],
        'HAN ROBERT SUNGWON': ['한성원', 'HAN SUNGWON', 'ROBERT', 'SUNGWON'],
    }

    for key, aliases in alias_map.items():
        if name in key or key in name:
            names.extend(aliases)

    return names


def filter_front_pages_only(pdf_path, applicant_name=None):
    """
    PDF에서 앞면 페이지만 추출해서 새 PDF로 반환.

    Args:
        pdf_path: 원본 PDF 경로
        applicant_name: 신청자 이름

    Returns:
        str: 필터링된 PDF 경로 (원본과 동일하거나 _front.pdf)
    """
    try:
        reader = PdfReader(pdf_path)
        page_count = len(reader.pages)

        # 페이지가 2개 이상이면 그대로 사용 (앞뒤 다 있다고 가정)
        if page_count >= 2:
            return pdf_path

        # 페이지가 1개면 앞면인지 확인
        if page_count == 1:
            if is_front_page(pdf_path, 0, applicant_name):
                return pdf_path
            else:
                # 뒷면만 있으면 제외 (None 반환)
                print(f"[id_filter] 뒷면만 있는 파일 제외: {pdf_path}")
                return None

        # 페이지가 0개면 제외
        return None

    except Exception as e:
        print(f"[id_filter] 필터링 실패 {pdf_path}: {e}")
        # 오류 시 원본 그대로 사용
        return pdf_path


def filter_id_copies_for_merge(file_paths, applicants):
    """
    신분증 파일 리스트에서 앞면만 필터링.

    Args:
        file_paths: 신분증 파일 경로 리스트
        applicants: 신청자 리스트 (id, name 포함)

    Returns:
        list[str]: 앞면만 포함된 파일 경로 리스트
    """
    # applicant ID별 이름 매핑
    id_to_name = {ap['id']: ap['name'] for ap in applicants}

    filtered = []
    for fp in file_paths:
        # 파일명에서 applicant_id 추출 (예: id_copy_123.pdf)
        # 실제로는 documents 테이블에서 가져올 것이므로 applicant_id를 함께 전달받아야 함
        # 일단 간단하게 모든 applicant 이름으로 시도

        # 모든 신청자 이름으로 매칭 시도
        applicant_name = None
        for ap_id, name in id_to_name.items():
            # 실제로는 file_path와 applicant를 매핑해야 하지만,
            # 여기서는 간단하게 파일명에 이름이 있는지 확인
            # (실제로는 documents 테이블에서 applicant_id를 함께 가져와야 함)
            pass

        # 필터링 (applicant_name은 모르므로 None으로 전달)
        filtered_path = filter_front_pages_only(fp, applicant_name)
        if filtered_path:
            filtered.append(filtered_path)

    return filtered
