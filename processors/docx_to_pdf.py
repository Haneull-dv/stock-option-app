"""
DOCX → PDF 변환 유틸리티.
Windows: docx2pdf 사용
"""
import os
import shutil


def convert_docx_to_pdf(docx_path, pdf_path):
    """
    DOCX 파일을 PDF로 변환.

    Windows에서는 docx2pdf 라이브러리 사용.
    실패 시 예외 발생.
    """
    if not os.path.isfile(docx_path):
        raise FileNotFoundError(f"DOCX file not found: {docx_path}")

    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    try:
        from docx2pdf import convert
        convert(docx_path, pdf_path)
        return pdf_path
    except Exception as e:
        raise RuntimeError(f"Failed to convert DOCX to PDF: {e}")


def convert_image_to_pdf(image_path, pdf_path):
    """
    이미지 파일(JPG, PNG)을 PDF로 변환.
    PIL(Pillow) 사용.
    """
    if not os.path.isfile(image_path):
        raise FileNotFoundError(f"Image file not found: {image_path}")

    os.makedirs(os.path.dirname(pdf_path), exist_ok=True)

    try:
        from PIL import Image
        img = Image.open(image_path)
        if img.mode == 'RGBA':
            img = img.convert('RGB')
        img.save(pdf_path, 'PDF')
        return pdf_path
    except Exception as e:
        raise RuntimeError(f"Failed to convert image to PDF: {e}")
