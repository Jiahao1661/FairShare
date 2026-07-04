import sqlite3
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "fairshare.db"


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def column_exists(conn, table, column):
    columns = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return any(row[1] == column for row in columns)


def init_db():
    conn = get_db()
    conn.executescript(
        """
        CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            email TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL
        );

        CREATE TABLE IF NOT EXISTS houses (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            house_name TEXT NOT NULL,
            house_emoji TEXT NOT NULL DEFAULT '🏠',
            house_color TEXT NOT NULL DEFAULT 'cyan',
            house_image TEXT,
            owner_id INTEGER NOT NULL,
            FOREIGN KEY (owner_id) REFERENCES users(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS housemates (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            house_id INTEGER NOT NULL,
            name TEXT NOT NULL,
            email TEXT,
            phone TEXT,
            FOREIGN KEY (house_id) REFERENCES houses(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS bills (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            house_id INTEGER NOT NULL,
            month TEXT NOT NULL,
            rent REAL NOT NULL CHECK (rent >= 0),
            utilities REAL NOT NULL CHECK (utilities >= 0),
            electricity REAL NOT NULL DEFAULT 0 CHECK (electricity >= 0),
            water REAL NOT NULL DEFAULT 0 CHECK (water >= 0),
            indah_water REAL NOT NULL DEFAULT 0 CHECK (indah_water >= 0),
            internet REAL NOT NULL DEFAULT 0 CHECK (internet >= 0),
            groceries REAL NOT NULL DEFAULT 0 CHECK (groceries >= 0),
            other_expenses REAL NOT NULL DEFAULT 0 CHECK (other_expenses >= 0),
            split_type TEXT NOT NULL CHECK (split_type IN ('equal', 'custom')),
            due_date TEXT,
            receipt_filename TEXT,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (house_id) REFERENCES houses(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS bill_splits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bill_id INTEGER NOT NULL,
            housemate_id INTEGER NOT NULL,
            rent_amount REAL NOT NULL,
            utility_amount REAL NOT NULL,
            total_amount REAL NOT NULL,
            paid INTEGER NOT NULL DEFAULT 0 CHECK (paid IN (0, 1)),
            FOREIGN KEY (bill_id) REFERENCES bills(id) ON DELETE CASCADE,
            FOREIGN KEY (housemate_id) REFERENCES housemates(id) ON DELETE CASCADE
        );

        CREATE TABLE IF NOT EXISTS reminder_logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            split_id INTEGER NOT NULL,
            sent_on TEXT NOT NULL,
            sent_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (split_id) REFERENCES bill_splits(id) ON DELETE CASCADE,
            UNIQUE (split_id, sent_on)
        );
        """
    )

    migrations = [
        ("houses", "house_emoji", "ALTER TABLE houses ADD COLUMN house_emoji TEXT NOT NULL DEFAULT '🏠'"),
        ("houses", "house_color", "ALTER TABLE houses ADD COLUMN house_color TEXT NOT NULL DEFAULT 'cyan'"),
        ("houses", "house_image", "ALTER TABLE houses ADD COLUMN house_image TEXT"),
        ("housemates", "email", "ALTER TABLE housemates ADD COLUMN email TEXT"),
        ("housemates", "phone", "ALTER TABLE housemates ADD COLUMN phone TEXT"),
        ("bills", "due_date", "ALTER TABLE bills ADD COLUMN due_date TEXT"),
        ("bills", "receipt_filename", "ALTER TABLE bills ADD COLUMN receipt_filename TEXT"),
        ("bills", "electricity", "ALTER TABLE bills ADD COLUMN electricity REAL NOT NULL DEFAULT 0"),
        ("bills", "water", "ALTER TABLE bills ADD COLUMN water REAL NOT NULL DEFAULT 0"),
        ("bills", "indah_water", "ALTER TABLE bills ADD COLUMN indah_water REAL NOT NULL DEFAULT 0"),
        ("bills", "internet", "ALTER TABLE bills ADD COLUMN internet REAL NOT NULL DEFAULT 0"),
        ("bills", "groceries", "ALTER TABLE bills ADD COLUMN groceries REAL NOT NULL DEFAULT 0"),
        ("bills", "other_expenses", "ALTER TABLE bills ADD COLUMN other_expenses REAL NOT NULL DEFAULT 0"),
        ("bill_splits", "paid", "ALTER TABLE bill_splits ADD COLUMN paid INTEGER NOT NULL DEFAULT 0"),
    ]
    for table, column, sql in migrations:
        if not column_exists(conn, table, column):
            conn.execute(sql)

    conn.commit()
    conn.close()
