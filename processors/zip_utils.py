"""
ZIP 파일 생성 유틸리티.
"""
import zipfile
import os
import shutil


def create_zip_from_folder(folder_path, zip_path):
    """
    폴더를 ZIP 파일로 압축.

    folder_path: 압축할 폴더 경로
    zip_path: 생성할 ZIP 파일 경로
    """
    if not os.path.isdir(folder_path):
        raise ValueError(f"Folder not found: {folder_path}")

    os.makedirs(os.path.dirname(zip_path), exist_ok=True)

    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for root, dirs, files in os.walk(folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                arcname = os.path.relpath(file_path, folder_path)
                zipf.write(file_path, arcname)

    return zip_path


def add_files_to_zip(zip_path, files_dict):
    """
    기존 ZIP에 파일 추가 (또는 새로 생성).

    files_dict: {arcname: file_path, ...}
    """
    os.makedirs(os.path.dirname(zip_path), exist_ok=True)

    mode = 'a' if os.path.isfile(zip_path) else 'w'
    with zipfile.ZipFile(zip_path, mode, zipfile.ZIP_DEFLATED) as zipf:
        for arcname, file_path in files_dict.items():
            if os.path.isfile(file_path):
                zipf.write(file_path, arcname)

    return zip_path


def copy_folder_contents(src_folder, dest_folder):
    """폴더 내용을 복사 (shutil.copytree 대신 개별 복사)."""
    if not os.path.isdir(src_folder):
        raise ValueError(f"Source folder not found: {src_folder}")

    os.makedirs(dest_folder, exist_ok=True)

    for item in os.listdir(src_folder):
        src_item = os.path.join(src_folder, item)
        dest_item = os.path.join(dest_folder, item)

        if os.path.isdir(src_item):
            shutil.copytree(src_item, dest_item, dirs_exist_ok=True)
        else:
            shutil.copy2(src_item, dest_item)
