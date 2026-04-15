"""
PDF 신청서에서 신청자 이름 추출
"""
import re
try:
    import fitz  # PyMuPDF
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False


def extract_name_from_pdf(file_path):
    """
    PDF 신청서에서 성명 추출

    Returns:
        str: 추출된 이름 또는 None
    """
    if not HAS_PYMUPDF:
        print("PyMuPDF가 설치되지 않았습니다. pip install pymupdf")
        return None

    try:
        doc = fitz.open(file_path)
        text = ""

        # 첫 페이지만 읽기 (신청서는 보통 1페이지)
        if len(doc) > 0:
            page = doc[0]
            text = page.get_text()

        doc.close()

        if not text:
            return None

        # 패턴 1: "성명 : 홍길동" 또는 "성명: 홍길동"
        match = re.search(r'성명\s*[:\s]\s*([가-힣]{2,4})', text)
        if match:
            return match.group(1).strip()

        # 패턴 2: "이름 : 홍길동"
        match = re.search(r'이름\s*[:\s]\s*([가-힣]{2,4})', text)
        if match:
            return match.group(1).strip()

        # 패턴 3: "신청자 : 홍길동"
        match = re.search(r'신청자\s*[:\s]\s*([가-힣]{2,4})', text)
        if match:
            return match.group(1).strip()

        # 패턴 4: "성 명" (띄어쓰기) 이후 한글 2-4자
        match = re.search(r'성\s*명\s*[:\s]\s*([가-힣]{2,4})', text)
        if match:
            return match.group(1).strip()

        # 패턴 5: 폼 필드 방식 - "성명홍길동" (붙어있는 경우)
        match = re.search(r'성명([가-힣]{2,4})', text)
        if match:
            name = match.group(1).strip()
            # "성명" 바로 뒤에 특정 키워드가 오는 경우 제외
            if name not in ['작성', '기재', '입력']:
                return name

        return None

    except Exception as e:
        print(f"PDF 이름 추출 실패: {file_path}, {e}")
        return None


def match_name_to_applicants(extracted_name, applicant_names):
    """
    추출한 이름을 신청자 명단과 매칭

    Args:
        extracted_name: PDF에서 추출한 이름
        applicant_names: {id: name} 딕셔너리

    Returns:
        (applicant_id, applicant_name) 또는 (None, None)
    """
    if not extracted_name:
        return None, None

    # 완전 일치 우선
    for aid, name in applicant_names.items():
        if name == extracted_name:
            return aid, name

    # 부분 일치 (성만 같거나, 이름만 같은 경우)
    for aid, name in applicant_names.items():
        if extracted_name in name or name in extracted_name:
            return aid, name

    return None, None
