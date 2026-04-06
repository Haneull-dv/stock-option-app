import os
import io
from pypdf import PdfWriter, PdfReader

try:
    from PIL import Image
    PILLOW_AVAILABLE = True
except ImportError:
    PILLOW_AVAILABLE = False


IMAGE_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp', '.JPG', '.JPEG', '.PNG'}
PDF_EXTENSION = '.pdf'

# Magic bytes for file type detection (확장자 없는 파일 처리용)
_MAGIC = {
    b'%PDF':         'pdf',
    b'\xff\xd8\xff': 'jpg',
    b'\x89PNG':      'png',
    b'GIF8':         'gif',
    b'BM':           'bmp',
}

def detect_type_by_content(file_path: str) -> str:
    """Read first 8 bytes and guess file type. Returns 'pdf', 'jpg', 'png', or ''."""
    try:
        with open(file_path, 'rb') as f:
            header = f.read(8)
        for magic, ftype in _MAGIC.items():
            if header.startswith(magic):
                return ftype
    except Exception:
        pass
    return ''


def image_to_pdf_bytes(image_path: str) -> bytes:
    """Convert an image file to PDF bytes using Pillow."""
    if not PILLOW_AVAILABLE:
        raise RuntimeError("Pillow is not installed. Cannot convert images to PDF.")
    with Image.open(image_path) as img:
        # Convert to RGB if necessary (e.g. RGBA, P)
        if img.mode not in ('RGB', 'L'):
            img = img.convert('RGB')
        buf = io.BytesIO()
        img.save(buf, format='PDF', resolution=150)
        return buf.getvalue()


def merge_pdfs_in_order(file_paths: list, output_path: str) -> str:
    """
    Merge multiple PDF/image files into a single PDF.

    Handles:
    - .pdf  files directly via pypdf
    - .jpg/.jpeg/.png (and other image types) via Pillow conversion

    Returns the output_path on success.
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    writer = PdfWriter()

    for fp in file_paths:
        if not os.path.isfile(fp):
            print(f"[pdf_merger] Skipping missing file: {fp}")
            continue

        ext = os.path.splitext(fp)[1].lower()

        # 확장자가 없거나 알 수 없으면 파일 내용으로 타입 감지
        if not ext or ext not in (IMAGE_EXTENSIONS | {'.pdf'}):
            detected = detect_type_by_content(fp)
            if detected == 'pdf':
                ext = '.pdf'
            elif detected in ('jpg', 'png', 'gif', 'bmp'):
                ext = f'.{detected}'
            else:
                print(f"[pdf_merger] Cannot detect type: {fp} (skipped)")
                continue

        if ext.lower() in IMAGE_EXTENSIONS:
            try:
                pdf_bytes = image_to_pdf_bytes(fp)
                reader = PdfReader(io.BytesIO(pdf_bytes))
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                print(f"[pdf_merger] Image conversion error {fp}: {e}")
                continue

        elif ext.lower() == PDF_EXTENSION:
            try:
                reader = PdfReader(fp)
                for page in reader.pages:
                    writer.add_page(page)
            except Exception as e:
                print(f"[pdf_merger] Error reading PDF {fp}: {e}")
                continue
        else:
            print(f"[pdf_merger] Unsupported file type: {fp} (skipped)")
            continue

    with open(output_path, 'wb') as f:
        writer.write(f)

    return output_path


def merge_docs_by_type(round_id, doc_type, applicants, upload_dir, output_dir) -> str:
    """
    Collect documents of `doc_type` in sort_order and merge into a single PDF.

    중복 처리 규칙:
      - application  : 행 수만큼 전부 포함 (동일인 여러 행사 → 모두 포함)
      - id_copy      : 동일 이름은 첫 번째 파일 하나만 포함
      - account_copy : 동일 이름+동일 계좌번호는 하나만; 동일인이라도 계좌가 다르면 모두 포함

    Returns: absolute path to the merged PDF
    """
    from database import get_documents_by_type

    # ASCII 파일명으로 저장 (URL 인코딩 문제 방지)
    output_filenames = {
        'application': 'application_merged.pdf',
        'id_copy':     'id_copy_merged.pdf',
        'account_copy':'account_copy_merged.pdf',
    }
    output_filename = output_filenames.get(doc_type, f'{doc_type}_merged.pdf')

    docs = get_documents_by_type(round_id, doc_type)

    # 중복 제거 필터링
    file_paths = []
    seen_names    = set()   # id_copy 중복 제거용
    seen_accounts = set()   # account_copy 중복 제거용

    for d in docs:
        fp = d['file_path']
        if not os.path.isfile(fp):
            continue

        name    = (d.get('name') or '').strip()
        account = (d.get('account_number') or '').strip()

        if doc_type == 'id_copy':
            if name in seen_names:
                continue          # 동일인 → 첫 번째 신분증만 사용
            seen_names.add(name)

        elif doc_type == 'account_copy':
            key = (name, account) if account else (name, fp)
            if key in seen_accounts:
                continue          # 동일인+동일계좌 → 한 번만
            seen_accounts.add(key)

        file_paths.append(fp)

    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, output_filename)

    if not file_paths:
        writer = PdfWriter()
        with open(output_path, 'wb') as f:
            writer.write(f)
        return output_path

    merge_pdfs_in_order(file_paths, output_path)
    return output_path


def get_pdf_page_count(pdf_path: str) -> int:
    """Return number of pages in a PDF, or 0 on error."""
    try:
        reader = PdfReader(pdf_path)
        return len(reader.pages)
    except Exception:
        return 0
