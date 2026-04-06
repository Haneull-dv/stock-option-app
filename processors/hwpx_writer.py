"""
HWPX 파일 텍스트 치환 유틸리티.
HWPX는 ZIP + XML 구조이므로 section0.xml 내 <hp:t> 텍스트를 직접 교체.
"""
import zipfile
import shutil
import os
import re
import io


def _replace_in_xml(xml_bytes: bytes, replacements: dict) -> bytes:
    """XML 바이트 문자열에서 텍스트 치환 수행."""
    text = xml_bytes.decode('utf-8')
    for old, new in replacements.items():
        text = text.replace(old, new)
    return text.encode('utf-8')


def generate_hwpx(template_path: str, output_path: str, replacements: dict) -> str:
    """
    template_path HWPX를 읽어 replacements 적용 후 output_path로 저장.

    replacements: {old_text: new_text, ...}
    """
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    buf = io.BytesIO()
    with zipfile.ZipFile(template_path, 'r') as zin, \
         zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zout:

        for item in zin.infolist():
            data = zin.read(item.filename)

            if item.filename in ('Contents/section0.xml', 'Preview/PrvText.txt'):
                data = _replace_in_xml(data, replacements)

            zout.writestr(item, data)

    with open(output_path, 'wb') as f:
        f.write(buf.getvalue())

    return output_path
