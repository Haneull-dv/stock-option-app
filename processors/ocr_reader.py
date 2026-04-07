"""신분증 OCR로 주민등록번호 추출."""
import re
import os
import io
from PIL import Image

_reader = None


def _get_reader():
    """EasyOCR reader 싱글턴 (첫 호출 시 모델 로드)."""
    global _reader
    if _reader is None:
        import easyocr
        _reader = easyocr.Reader(['ko', 'en'], gpu=False, verbose=False)
    return _reader


# 주민등록번호 패턴: 6자리-7자리
_RRN_PATTERN = re.compile(r'(\d{6})\s*[-–—]\s*(\d{7})')
# 느슨한 패턴 (OCR 오류 대비: 구분자가 . 이나 공백인 경우)
_RRN_LOOSE = re.compile(r'(\d{6})\s*[-–—.,\s]\s*(\d{6,7})')


def _images_from_pdf(pdf_path):
    """pypdf로 PDF에서 모든 이미지를 추출."""
    images = []
    try:
        import pypdf
        reader = pypdf.PdfReader(pdf_path)
        for page in reader.pages:
            if '/XObject' not in (page.get('/Resources') or {}):
                continue
            x_objects = page['/Resources']['/XObject'].get_object()
            for obj_name in x_objects:
                obj = x_objects[obj_name].get_object()
                if obj.get('/Subtype') != '/Image':
                    continue
                data = obj.get_data()
                width = obj.get('/Width', 0)
                height = obj.get('/Height', 0)
                filt = obj.get('/Filter')
                # 배열 형태 필터 처리 (예: ['/FlateDecode', '/DCTDecode'])
                if isinstance(filt, list):
                    filt = filt[-1] if filt else None
                elif hasattr(filt, 'get_object'):
                    filt = str(filt.get_object())
                else:
                    filt = str(filt) if filt else None

                try:
                    if filt in ('/DCTDecode', '/JPXDecode'):
                        images.append(Image.open(io.BytesIO(data)))
                    elif width and height:
                        cs = str(obj.get('/ColorSpace', ''))
                        if 'RGB' in cs:
                            images.append(Image.frombytes('RGB', (width, height), data))
                        elif 'Gray' in cs:
                            images.append(Image.frombytes('L', (width, height), data))
                except Exception:
                    pass
    except Exception:
        pass
    return images


def _load_images(file_path):
    """파일 경로에서 이미지 목록 로드."""
    ext = os.path.splitext(file_path)[1].lower()
    if ext in ('.jpg', '.jpeg', '.png', '.bmp', '.tiff', '.tif'):
        try:
            return [Image.open(file_path)]
        except Exception:
            return []
    elif ext == '.pdf':
        return _images_from_pdf(file_path)
    return []


def _find_rrn_in_text(text):
    """텍스트에서 주민번호 패턴 찾기."""
    # 정확한 패턴
    m = _RRN_PATTERN.search(text)
    if m:
        return f'{m.group(1)}-{m.group(2)}'

    # 느슨한 패턴
    m = _RRN_LOOSE.search(text)
    if m:
        part2 = m.group(2)
        if len(part2) == 7:
            return f'{m.group(1)}-{part2}'
        elif len(part2) == 6:
            return f'{m.group(1)}-{part2}*'

    # 공백 제거 후 재시도
    text_no_space = text.replace(' ', '')
    m = _RRN_PATTERN.search(text_no_space)
    if m:
        return f'{m.group(1)}-{m.group(2)}'

    return ''


def extract_rrn(file_path):
    """
    신분증 이미지/PDF에서 주민등록번호를 추출.
    PDF는 모든 이미지를 추출하여 각각 OCR 시도.
    Returns: 'XXXXXX-XXXXXXX' 형태 문자열 or ''
    """
    if not os.path.isfile(file_path):
        return ''

    images = _load_images(file_path)
    if not images:
        return ''

    import numpy as np
    reader = _get_reader()

    # 크기가 큰 이미지부터 시도 (신분증 본문일 가능성 높음)
    images.sort(key=lambda img: img.size[0] * img.size[1], reverse=True)

    for img in images:
        try:
            w, h = img.size
            if w < 100 or h < 100:
                continue  # 너무 작은 이미지 스킵

            # 작은 이미지는 확대
            if w < 500:
                ratio = 1000 / w
                img = img.resize((int(w * ratio), int(h * ratio)), Image.LANCZOS)

            img_array = np.array(img.convert('RGB'))
            results = reader.readtext(img_array, detail=0)
            full_text = ' '.join(results)

            rrn = _find_rrn_in_text(full_text)
            if rrn:
                return rrn
        except Exception:
            continue

    return ''


def extract_rrn_batch(applicant_docs):
    """
    여러 신청자의 신분증에서 일괄 주민번호 추출.
    applicant_docs: list of {'applicant_id': int, 'name': str, 'file_path': str}
    Returns: {applicant_id: 'XXXXXX-XXXXXXX' or ''}
    """
    results = {}
    for doc in applicant_docs:
        aid = doc['applicant_id']
        rrn = extract_rrn(doc['file_path'])
        results[aid] = rrn
    return results
