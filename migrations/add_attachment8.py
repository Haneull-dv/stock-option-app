"""Add attachment8_uploads table for price-specific attachments."""
import sqlite3
import os

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'data', 'stock_option.db')

def migrate():
    conn = sqlite3.connect(DB_PATH)
    cursor = conn.cursor()

    # Create attachment8_uploads table
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS attachment8_uploads (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            price INTEGER NOT NULL,
            file_path TEXT NOT NULL,
            original_filename TEXT,
            uploaded_at TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE,
            UNIQUE(round_id, price)
        )
    """)

    conn.commit()
    conn.close()
    print("Migration completed: attachment8_uploads table created")

if __name__ == '__main__':
    migrate()
