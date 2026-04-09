"""
DB 마이그레이션: applicants 테이블에 OCR 관련 컬럼 추가
실행 방법: python migrate_add_ocr_columns.py
"""
import sqlite3
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'data', 'stockops.db')

def migrate():
    if not os.path.exists(DB_PATH):
        print(f"DB 파일이 없습니다: {DB_PATH}")
        print("database.py의 init_db()가 자동으로 새 테이블을 생성합니다.")
        return

    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # 기존 컬럼 확인
    c.execute("PRAGMA table_info(applicants)")
    columns = [row[1] for row in c.fetchall()]

    print(f"현재 applicants 테이블 컬럼: {columns}")

    # 컬럼 추가
    if 'rrn' not in columns:
        print("  > rrn column adding...")
        c.execute("ALTER TABLE applicants ADD COLUMN rrn TEXT")
        print("  + rrn column added")
    else:
        print("  + rrn column already exists")

    if 'ocr_account' not in columns:
        print("  > ocr_account column adding...")
        c.execute("ALTER TABLE applicants ADD COLUMN ocr_account TEXT")
        print("  + ocr_account column added")
    else:
        print("  + ocr_account column already exists")

    if 'ocr_extracted_at' not in columns:
        print("  > ocr_extracted_at column adding...")
        c.execute("ALTER TABLE applicants ADD COLUMN ocr_extracted_at TEXT")
        print("  + ocr_extracted_at column added")
    else:
        print("  + ocr_extracted_at column already exists")

    conn.commit()
    conn.close()

    print("\n마이그레이션 완료!")

if __name__ == '__main__':
    migrate()
