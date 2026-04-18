"""
Step06 KIND 상장신청을 위한 DB 스키마 추가.
"""
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'stockops.db')

def migrate():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()

    # step06_config 테이블
    c.execute('''
        CREATE TABLE IF NOT EXISTS step06_config (
            round_id INTEGER PRIMARY KEY,
            submission_date TEXT,
            listing_fee_receipt TEXT,
            holding_proof_folder TEXT,
            employment_cert_folder TEXT,
            exercise_summary_excel TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY(round_id) REFERENCES rounds(id) ON DELETE CASCADE
        )
    ''')

    # step06_issuance_confirmations 테이블 (발행가액별 발행등록사실확인서)
    c.execute('''
        CREATE TABLE IF NOT EXISTS step06_issuance_confirmations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            exercise_price INTEGER NOT NULL,
            file_path TEXT,
            original_name TEXT,
            uploaded_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(round_id, exercise_price),
            FOREIGN KEY(round_id) REFERENCES rounds(id) ON DELETE CASCADE
        )
    ''')

    conn.commit()
    conn.close()
    print("Step06 tables created successfully")

if __name__ == '__main__':
    migrate()
