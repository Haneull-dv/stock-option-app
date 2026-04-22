import sqlite3
import os
import secrets
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'data', 'stockops.db')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("PRAGMA foreign_keys=ON")
    return conn


def init_db():
    os.makedirs(os.path.dirname(DB_PATH), exist_ok=True)
    conn = get_db()
    c = conn.cursor()
    c.executescript("""
        CREATE TABLE IF NOT EXISTS rounds (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT NOT NULL,
            exercise_date TEXT,
            notes TEXT,
            status TEXT DEFAULT '진행중',
            created_at TEXT DEFAULT (datetime('now', 'localtime'))
        );

        CREATE TABLE IF NOT EXISTS exercise_prices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            price INTEGER NOT NULL,
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS applicants (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            sort_order INTEGER DEFAULT 0,
            name TEXT NOT NULL,
            exercise_price INTEGER,
            quantity INTEGER,
            broker TEXT,
            account_number TEXT,
            submit_token TEXT UNIQUE,
            doc_submitted INTEGER DEFAULT 0,
            rrn TEXT,
            grant_date TEXT,
            ocr_account TEXT,
            ocr_extracted_at TEXT,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS documents (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            applicant_id INTEGER NOT NULL,
            doc_type TEXT NOT NULL,
            filename TEXT NOT NULL,
            original_filename TEXT,
            file_path TEXT NOT NULL,
            uploaded_at TEXT DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (applicant_id) REFERENCES applicants(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS step_outputs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            step TEXT NOT NULL,
            output_filename TEXT NOT NULL,
            output_path TEXT NOT NULL,
            created_at TEXT DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS holding_config (
            round_id INTEGER PRIMARY KEY,
            holding_start TEXT,
            holding_end TEXT,
            doc_date TEXT,
            processing_date TEXT,
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS holding_subjects (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            sort_order INTEGER DEFAULT 0,
            name TEXT NOT NULL,
            relationship TEXT DEFAULT '미등기임원',
            quantity INTEGER DEFAULT 0,
            branch TEXT DEFAULT '도곡',
            account_number TEXT,
            note TEXT,
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS reg_config (
            round_id INTEGER PRIMARY KEY,
            reg_date TEXT,
            issue_date TEXT,
            par_value INTEGER DEFAULT 500,
            capital_before INTEGER,
            shares_before INTEGER,
            company_name TEXT DEFAULT 'S2W Inc.',
            company_reg_num TEXT,
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS issuance_config (
            round_id INTEGER PRIMARY KEY,
            payment_date TEXT,
            dividend_base_date TEXT,
            listing_date TEXT,
            contact_name TEXT DEFAULT '정민우',
            contact_phone TEXT DEFAULT '010-3615-4909',
            stock_code TEXT DEFAULT '488280',
            agent_name TEXT,
            agent_phone TEXT,
            agent_rrn TEXT,
            agent_address TEXT,
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS attachment8 (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            exercise_price INTEGER NOT NULL,
            file_name TEXT,
            original_name TEXT,
            file_path TEXT,
            uploaded_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(round_id, exercise_price),
            FOREIGN KEY (round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );

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
        );

        CREATE TABLE IF NOT EXISTS step06_issuance_confirmations (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            round_id INTEGER NOT NULL,
            exercise_price INTEGER NOT NULL,
            file_path TEXT,
            original_name TEXT,
            uploaded_at TEXT DEFAULT (datetime('now', 'localtime')),
            UNIQUE(round_id, exercise_price),
            FOREIGN KEY(round_id) REFERENCES rounds(id) ON DELETE CASCADE
        );
    """)

    # Migration: grant_date 칼럼 추가 (이미 있으면 무시)
    try:
        c.execute("ALTER TABLE applicants ADD COLUMN grant_date TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # 이미 칼럼이 있음

    conn.commit()
    conn.close()


# ── Round helpers ──────────────────────────────────────────────────────────────

def get_all_rounds():
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM rounds ORDER BY created_at DESC"
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_round(round_id):
    conn = get_db()
    row = conn.execute("SELECT * FROM rounds WHERE id=?", (round_id,)).fetchone()
    conn.close()
    return dict(row) if row else None


def get_exercise_prices(round_id):
    """회차의 행사가액 목록 조회."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM exercise_prices WHERE round_id=? ORDER BY price",
        (round_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def create_round(name, exercise_date, notes, prices):
    conn = get_db()
    c = conn.cursor()
    c.execute(
        "INSERT INTO rounds (name, exercise_date, notes) VALUES (?,?,?)",
        (name, exercise_date, notes)
    )
    round_id = c.lastrowid
    for p in prices:
        if p:
            c.execute(
                "INSERT INTO exercise_prices (round_id, price) VALUES (?,?)",
                (round_id, int(p))
            )
    conn.commit()
    conn.close()
    return round_id


def update_round(round_id, name, exercise_date, notes, prices):
    conn = get_db()
    conn.execute(
        "UPDATE rounds SET name=?, exercise_date=?, notes=? WHERE id=?",
        (name, exercise_date, notes, round_id)
    )
    conn.execute("DELETE FROM exercise_prices WHERE round_id=?", (round_id,))
    for p in prices:
        conn.execute(
            "INSERT INTO exercise_prices (round_id, price) VALUES (?,?)",
            (round_id, int(p))
        )
    conn.commit()
    conn.close()


def update_round_status(round_id, status):
    conn = get_db()
    conn.execute("UPDATE rounds SET status=? WHERE id=?", (status, round_id))
    conn.commit()
    conn.close()


def get_prices_for_round(round_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT price FROM exercise_prices WHERE round_id=? ORDER BY price",
        (round_id,)
    ).fetchall()
    conn.close()
    return [r['price'] for r in rows]


# ── Applicant helpers ──────────────────────────────────────────────────────────

def get_applicants(round_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM applicants WHERE round_id=? ORDER BY sort_order, id",
        (round_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_applicant(applicant_id):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM applicants WHERE id=?", (applicant_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def get_applicant_by_token(token):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM applicants WHERE submit_token=?", (token,)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def add_applicant(round_id, name, exercise_price, quantity, broker, account_number, grant_date=None):
    conn = get_db()
    c = conn.cursor()
    max_order = c.execute(
        "SELECT COALESCE(MAX(sort_order),0) FROM applicants WHERE round_id=?",
        (round_id,)
    ).fetchone()[0]
    token = secrets.token_urlsafe(16)
    c.execute(
        """INSERT INTO applicants
           (round_id, sort_order, name, exercise_price, quantity, broker, account_number, submit_token, grant_date)
           VALUES (?,?,?,?,?,?,?,?,?)""",
        (round_id, max_order + 1, name, exercise_price, quantity, broker, account_number, token, grant_date)
    )
    applicant_id = c.lastrowid
    conn.commit()
    conn.close()
    return applicant_id


def delete_applicant(applicant_id):
    conn = get_db()
    conn.execute("DELETE FROM applicants WHERE id=?", (applicant_id,))
    conn.commit()
    conn.close()


def delete_all_applicants(round_id):
    conn = get_db()
    conn.execute("DELETE FROM applicants WHERE round_id=?", (round_id,))
    conn.commit()
    conn.close()


def reorder_applicants(round_id, id_list):
    conn = get_db()
    for idx, aid in enumerate(id_list):
        conn.execute(
            "UPDATE applicants SET sort_order=? WHERE id=? AND round_id=?",
            (idx + 1, aid, round_id)
        )
    conn.commit()
    conn.close()


# ── Document helpers ───────────────────────────────────────────────────────────

def add_document(applicant_id, doc_type, filename, original_filename, file_path):
    conn = get_db()
    c = conn.cursor()
    # Remove previous doc of same type for this applicant
    c.execute(
        "DELETE FROM documents WHERE applicant_id=? AND doc_type=?",
        (applicant_id, doc_type)
    )
    c.execute(
        """INSERT INTO documents (applicant_id, doc_type, filename, original_filename, file_path)
           VALUES (?,?,?,?,?)""",
        (applicant_id, doc_type, filename, original_filename, file_path)
    )
    doc_id = c.lastrowid
    # update doc_submitted flag
    _update_doc_submitted(conn, applicant_id)
    conn.commit()
    conn.close()
    return doc_id


def _update_doc_submitted(conn, applicant_id):
    count = conn.execute(
        "SELECT COUNT(DISTINCT doc_type) FROM documents WHERE applicant_id=?",
        (applicant_id,)
    ).fetchone()[0]
    submitted = 1 if count >= 3 else 0
    conn.execute(
        "UPDATE applicants SET doc_submitted=? WHERE id=?",
        (submitted, applicant_id)
    )


def get_documents(applicant_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM documents WHERE applicant_id=?", (applicant_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_documents_by_type(round_id, doc_type):
    """Return documents of a given type for all applicants in sort_order."""
    conn = get_db()
    rows = conn.execute(
        """SELECT d.*, a.name, a.sort_order, a.account_number
           FROM documents d
           JOIN applicants a ON d.applicant_id = a.id
           WHERE a.round_id=? AND d.doc_type=?
           ORDER BY a.sort_order, a.id""",
        (round_id, doc_type)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_all_documents_for_round(round_id):
    conn = get_db()
    rows = conn.execute(
        """SELECT d.*, a.name, a.sort_order
           FROM documents d
           JOIN applicants a ON d.applicant_id = a.id
           WHERE a.round_id=?
           ORDER BY a.sort_order, a.id, d.doc_type""",
        (round_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_document(doc_id):
    conn = get_db()
    row = conn.execute("SELECT applicant_id FROM documents WHERE id=?", (doc_id,)).fetchone()
    conn.execute("DELETE FROM documents WHERE id=?", (doc_id,))
    if row:
        _update_doc_submitted(conn, row['applicant_id'])
    conn.commit()
    conn.close()


# ── Step output helpers ────────────────────────────────────────────────────────

def save_step_output(round_id, step, output_filename, output_path):
    conn = get_db()
    conn.execute(
        """INSERT INTO step_outputs (round_id, step, output_filename, output_path)
           VALUES (?,?,?,?)""",
        (round_id, step, output_filename, output_path)
    )
    conn.commit()
    conn.close()


def get_step_outputs(round_id, step=None):
    conn = get_db()
    if step:
        rows = conn.execute(
            "SELECT * FROM step_outputs WHERE round_id=? AND step=? ORDER BY created_at DESC",
            (round_id, step)
        ).fetchall()
    else:
        rows = conn.execute(
            "SELECT * FROM step_outputs WHERE round_id=? ORDER BY created_at DESC",
            (round_id,)
        ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def delete_step_output(output_id):
    """결과물 삭제 (DB에서만, 파일 삭제는 app.py에서)"""
    conn = get_db()
    conn.execute("DELETE FROM step_outputs WHERE id=?", (output_id,))
    conn.commit()
    conn.close()


# ── Submission status helper ───────────────────────────────────────────────────

def get_submission_status(round_id):
    """Return per-applicant doc submission status."""
    applicants = get_applicants(round_id)
    result = []
    for ap in applicants:
        docs = get_documents(ap['id'])
        doc_map = {d['doc_type']: d for d in docs}
        result.append({
            'applicant_id': ap['id'],
            'name': ap['name'],
            'sort_order': ap['sort_order'],
            'application': 'application' in doc_map,
            'id_copy': 'id_copy' in doc_map,
            'account_copy': 'account_copy' in doc_map,
            'all_submitted': len(doc_map) >= 3,
        })
    return result


# ── Holding (Step 03-3) helpers ────────────────────────────────────────────────

def get_holding_config(round_id):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM holding_config WHERE round_id=?", (round_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else {}


def save_holding_config(round_id, holding_start, holding_end, doc_date, processing_date):
    conn = get_db()
    conn.execute(
        """INSERT INTO holding_config (round_id, holding_start, holding_end, doc_date, processing_date)
           VALUES (?,?,?,?,?)
           ON CONFLICT(round_id) DO UPDATE SET
             holding_start=excluded.holding_start,
             holding_end=excluded.holding_end,
             doc_date=excluded.doc_date,
             processing_date=excluded.processing_date""",
        (round_id, holding_start, holding_end, doc_date, processing_date)
    )
    conn.commit()
    conn.close()


def get_holding_subjects(round_id):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM holding_subjects WHERE round_id=? ORDER BY sort_order, id",
        (round_id,)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_holding_subject(round_id, name, relationship, quantity, branch, account_number, note):
    conn = get_db()
    max_order = conn.execute(
        "SELECT COALESCE(MAX(sort_order),0) FROM holding_subjects WHERE round_id=?",
        (round_id,)
    ).fetchone()[0]
    conn.execute(
        """INSERT INTO holding_subjects
           (round_id, sort_order, name, relationship, quantity, branch, account_number, note)
           VALUES (?,?,?,?,?,?,?,?)""",
        (round_id, max_order + 1, name, relationship, quantity, branch, account_number, note)
    )
    conn.commit()
    conn.close()


def update_holding_subject(subject_id, name, relationship, quantity, branch, account_number, note):
    conn = get_db()
    conn.execute(
        """UPDATE holding_subjects SET name=?, relationship=?, quantity=?,
           branch=?, account_number=?, note=? WHERE id=?""",
        (name, relationship, quantity, branch, account_number, note, subject_id)
    )
    conn.commit()
    conn.close()


def delete_holding_subject(subject_id):
    conn = get_db()
    conn.execute("DELETE FROM holding_subjects WHERE id=?", (subject_id,))
    conn.commit()
    conn.close()


def delete_all_holding_subjects(round_id):
    conn = get_db()
    conn.execute("DELETE FROM holding_subjects WHERE round_id=?", (round_id,))
    conn.commit()
    conn.close()


# ── Registration (Step 04) helpers ────────────────────────────────────────────

def get_reg_config(round_id):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM reg_config WHERE round_id=?", (round_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else {}


def save_reg_config(round_id, reg_date, issue_date, par_value,
                    capital_before, shares_before, company_name, company_reg_num):
    conn = get_db()
    conn.execute(
        """INSERT INTO reg_config
           (round_id, reg_date, issue_date, par_value, capital_before, shares_before,
            company_name, company_reg_num)
           VALUES (?,?,?,?,?,?,?,?)
           ON CONFLICT(round_id) DO UPDATE SET
             reg_date=excluded.reg_date,
             issue_date=excluded.issue_date,
             par_value=excluded.par_value,
             capital_before=excluded.capital_before,
             shares_before=excluded.shares_before,
             company_name=excluded.company_name,
             company_reg_num=excluded.company_reg_num""",
        (round_id, reg_date, issue_date,
         int(par_value) if par_value else 500,
         int(capital_before) if capital_before else None,
         int(shares_before) if shares_before else None,
         company_name or 'S2W Inc.',
         company_reg_num or '')
    )
    conn.commit()
    conn.close()


# ── Issuance (Step 05) helpers ────────────────────────────────────────────────

def get_issuance_config(round_id):
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM issuance_config WHERE round_id=?", (round_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else {}


def save_issuance_config(round_id, payment_date, dividend_base_date,
                         listing_date, contact_name, contact_phone, stock_code,
                         agent_name=None, agent_phone=None, agent_rrn=None, agent_address=None):
    conn = get_db()
    conn.execute(
        """INSERT INTO issuance_config
           (round_id, payment_date, dividend_base_date, listing_date,
            contact_name, contact_phone, stock_code,
            agent_name, agent_phone, agent_rrn, agent_address)
           VALUES (?,?,?,?,?,?,?,?,?,?,?)
           ON CONFLICT(round_id) DO UPDATE SET
             payment_date=excluded.payment_date,
             dividend_base_date=excluded.dividend_base_date,
             listing_date=excluded.listing_date,
             contact_name=excluded.contact_name,
             contact_phone=excluded.contact_phone,
             stock_code=excluded.stock_code,
             agent_name=excluded.agent_name,
             agent_phone=excluded.agent_phone,
             agent_rrn=excluded.agent_rrn,
             agent_address=excluded.agent_address""",
        (round_id, payment_date or '', dividend_base_date or '',
         listing_date or '', contact_name or '정민우',
         contact_phone or '010-3615-4909', stock_code or '488280',
         agent_name or '', agent_phone or '', agent_rrn or '', agent_address or '')
    )
    conn.commit()
    conn.close()


def get_applicants_by_price(round_id, price):
    conn = get_db()
    rows = conn.execute(
        """SELECT * FROM applicants
           WHERE round_id=? AND exercise_price=?
           ORDER BY sort_order, id""",
        (round_id, price)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def get_documents_for_applicant_ids(applicant_ids, doc_type):
    """Return documents of a given type for specific applicant IDs."""
    if not applicant_ids:
        return []
    conn = get_db()
    placeholders = ','.join('?' * len(applicant_ids))
    rows = conn.execute(
        f"""SELECT d.*, a.name, a.sort_order, a.exercise_price
            FROM documents d
            JOIN applicants a ON d.applicant_id = a.id
            WHERE d.applicant_id IN ({placeholders}) AND d.doc_type=?
            ORDER BY a.sort_order, a.id""",
        (*applicant_ids, doc_type)
    ).fetchall()
    conn.close()
    return [dict(r) for r in rows]


# ── Round statistics ───────────────────────────────────────────────────────────

def get_round_stats(round_id):
    conn = get_db()
    total = conn.execute(
        "SELECT COUNT(*) FROM applicants WHERE round_id=?", (round_id,)
    ).fetchone()[0]
    submitted = conn.execute(
        "SELECT COUNT(*) FROM applicants WHERE round_id=? AND doc_submitted=1", (round_id,)
    ).fetchone()[0]
    conn.close()
    return {'total': total, 'submitted': submitted}


# ── OCR helpers ────────────────────────────────────────────────────────────────

def update_applicant_ocr(applicant_id, rrn=None, ocr_account=None, broker=None, force_update=False):
    """
    신청자의 OCR 추출 데이터 업데이트.

    force_update=True이면 값이 없어도 ocr_extracted_at를 업데이트 (재시도 방지용)
    """
    conn = get_db()
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    updates = []
    params = []

    if rrn is not None:
        updates.append("rrn=?")
        params.append(rrn)

    if ocr_account is not None:
        updates.append("ocr_account=?")
        params.append(ocr_account)

    if broker is not None:
        updates.append("broker=?")
        params.append(broker)

    # force_update이거나 실제 업데이트할 값이 있으면 타임스탬프 업데이트
    if updates or force_update:
        updates.append("ocr_extracted_at=?")
        params.append(now)
        params.append(applicant_id)

        sql = f"UPDATE applicants SET {', '.join(updates)} WHERE id=?"
        conn.execute(sql, params)
        conn.commit()

    conn.close()


def get_applicant_ocr(applicant_id):
    """신청자의 OCR 데이터 조회."""
    conn = get_db()
    row = conn.execute(
        "SELECT rrn, ocr_account, ocr_extracted_at FROM applicants WHERE id=?",
        (applicant_id,)
    ).fetchone()
    conn.close()
    if row:
        return {'rrn': row[0], 'ocr_account': row[1], 'ocr_extracted_at': row[2]}
    return {}


# ── Attachment8 (주식납입금 보관증명서) helpers ────────────────────────────────

def save_attachment8(round_id, exercise_price, file_name, original_name, file_path):
    """발행가액별 붙임8 파일 저장."""
    conn = get_db()
    conn.execute("""
        INSERT INTO attachment8
        (round_id, exercise_price, file_name, original_name, file_path)
        VALUES (?, ?, ?, ?, ?)
        ON CONFLICT(round_id, exercise_price) DO UPDATE SET
          file_name=excluded.file_name,
          original_name=excluded.original_name,
          file_path=excluded.file_path,
          uploaded_at=datetime('now', 'localtime')
    """, (round_id, exercise_price, file_name, original_name, file_path))
    conn.commit()
    conn.close()


def get_attachment8(round_id, exercise_price):
    """특정 발행가액의 붙임8 파일 조회."""
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM attachment8 WHERE round_id=? AND exercise_price=?",
        (round_id, exercise_price)
    ).fetchone()
    conn.close()
    return dict(row) if row else None


def get_all_attachment8(round_id):
    """해당 회차의 모든 붙임8 파일 조회 (exercise_price를 키로 하는 딕셔너리)."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM attachment8 WHERE round_id=? ORDER BY exercise_price",
        (round_id,)
    ).fetchall()
    conn.close()
    return {r['exercise_price']: dict(r) for r in rows}


def delete_attachment8(round_id, exercise_price):
    """특정 발행가액의 붙임8 파일 삭제."""
    conn = get_db()
    conn.execute("DELETE FROM attachment8 WHERE round_id=? AND exercise_price=?", (round_id, exercise_price))
    conn.commit()
    conn.close()


# ── Step03 (행사대금 납입) helpers ─────────────────────────────────────────────

def save_step03_config(round_id, payment_date=None, total_capital=None):
    """Step03 기본 설정 저장."""
    conn = get_db()
    conn.execute(
        """INSERT INTO step03_config (round_id, payment_date, total_capital)
           VALUES (?, ?, ?)
           ON CONFLICT(round_id) DO UPDATE SET
             payment_date=excluded.payment_date,
             total_capital=excluded.total_capital,
             updated_at=datetime('now', 'localtime')""",
        (round_id, payment_date, total_capital)
    )
    conn.commit()
    conn.close()


def get_step03_config(round_id):
    """Step03 설정 조회."""
    conn = get_db()
    row = conn.execute(
        "SELECT * FROM step03_config WHERE round_id=?", (round_id,)
    ).fetchone()
    conn.close()
    return dict(row) if row else {}


# ── Step06 (KIND 상장신청) helpers ─────────────────────────────────────────────

def save_step06_config(round_id, submission_date=None, listing_fee_receipt=None,
                       holding_proof_folder=None, employment_cert_folder=None,
                       exercise_summary_excel=None):
    """Step06 기본 설정 저장."""
    conn = get_db()
    updates = []
    params = []

    if submission_date is not None:
        updates.append("submission_date=?")
        params.append(submission_date)
    if listing_fee_receipt is not None:
        updates.append("listing_fee_receipt=?")
        params.append(listing_fee_receipt)
    if holding_proof_folder is not None:
        updates.append("holding_proof_folder=?")
        params.append(holding_proof_folder)
    if employment_cert_folder is not None:
        updates.append("employment_cert_folder=?")
        params.append(employment_cert_folder)
    if exercise_summary_excel is not None:
        updates.append("exercise_summary_excel=?")
        params.append(exercise_summary_excel)

    if updates:
        updates.append("updated_at=datetime('now', 'localtime')")
        params.append(round_id)

        # UPSERT
        conn.execute(f"""
            INSERT INTO step06_config (round_id, {', '.join(u.split('=')[0] for u in updates[:-1])})
            VALUES (?, {', '.join('?' * (len(updates) - 1))})
            ON CONFLICT(round_id) DO UPDATE SET {', '.join(updates)}
        """, [round_id] + params[:-1] + [round_id])
        conn.commit()

    conn.close()


def get_step06_config(round_id):
    """Step06 설정 조회."""
    conn = get_db()
    row = conn.execute("SELECT * FROM step06_config WHERE round_id=?", (round_id,)).fetchone()
    conn.close()
    return dict(row) if row else {}


def save_step06_issuance_confirmation(round_id, exercise_price, file_path, original_name):
    """발행가액별 발행등록사실확인서 저장."""
    conn = get_db()
    conn.execute("""
        INSERT INTO step06_issuance_confirmations
        (round_id, exercise_price, file_path, original_name)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(round_id, exercise_price) DO UPDATE SET
          file_path=excluded.file_path,
          original_name=excluded.original_name,
          uploaded_at=datetime('now', 'localtime')
    """, (round_id, exercise_price, file_path, original_name))
    conn.commit()
    conn.close()


def get_step06_issuance_confirmations(round_id):
    """해당 회차의 모든 발행등록사실확인서 조회."""
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM step06_issuance_confirmations WHERE round_id=? ORDER BY exercise_price",
        (round_id,)
    ).fetchall()
    conn.close()
    return {r['exercise_price']: dict(r) for r in rows}


def delete_step06_issuance_confirmation(round_id, exercise_price):
    """특정 발행가액의 발행등록사실확인서 삭제."""
    conn = get_db()
    conn.execute(
        "DELETE FROM step06_issuance_confirmations WHERE round_id=? AND exercise_price=?",
        (round_id, exercise_price)
    )
    conn.commit()
    conn.close()
