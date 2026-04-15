#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Step05 업데이트 마이그레이션
- issuance_config에 shareholders_count, shares_count 추가
- attachment8 테이블 생성 (발행가액별 주식납입금보관증명서 업로드용)
"""
import sqlite3

def migrate():
    conn = sqlite3.connect('data/stockops.db')
    cur = conn.cursor()

    # 1. issuance_config에 주주수/주식수 컬럼 추가
    try:
        cur.execute("ALTER TABLE issuance_config ADD COLUMN shareholders_count INTEGER DEFAULT NULL")
        print("[OK] issuance_config.shareholders_count added")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("[SKIP] issuance_config.shareholders_count already exists")
        else:
            raise

    try:
        cur.execute("ALTER TABLE issuance_config ADD COLUMN shares_count INTEGER DEFAULT NULL")
        print("[OK] issuance_config.shares_count added")
    except sqlite3.OperationalError as e:
        if "duplicate column name" in str(e).lower():
            print("[SKIP] issuance_config.shares_count already exists")
        else:
            raise

    # 2. attachment8 테이블 생성 (발행가액별 붙임8 업로드)
    cur.execute("""
        CREATE TABLE IF NOT EXISTS attachment8 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            exercise_price INTEGER NOT NULL,
            file_name TEXT NOT NULL,
            original_name TEXT NOT NULL,
            file_path TEXT NOT NULL,
            uploaded_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(round_id, exercise_price)
        )
    """)
    print("[OK] attachment8 table created")

    conn.commit()
    conn.close()
    print("\nMigration completed!")

if __name__ == '__main__':
    migrate()
