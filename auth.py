#!/usr/bin/env python3
"""
AuthManager — zarządzanie użytkownikami w SQLite.

Tabela `users` w ceny_tge.db. Hasła: sha256(salt + password).
Role: admin, handlowiec, guest.
"""

import hashlib
import os
import sqlite3
from datetime import datetime
from typing import Optional


DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ceny_tge.db')


class AuthManager:
    """CRUD użytkowników z hashowaniem haseł."""

    def __init__(self, db_path: str = DB_PATH):
        self._db = db_path
        self._init_db()

    def _conn(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self._db)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        with self._conn() as conn:
            conn.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    username TEXT UNIQUE NOT NULL,
                    password_hash TEXT NOT NULL,
                    salt TEXT NOT NULL,
                    rola TEXT NOT NULL DEFAULT 'handlowiec',
                    aktywny INTEGER NOT NULL DEFAULT 1,
                    utworzony TEXT NOT NULL,
                    ostatnie_logowanie TEXT
                )
            ''')
            # Seed default admin if no users exist
            count = conn.execute('SELECT COUNT(*) FROM users').fetchone()[0]
            if count == 0:
                self.create_user('admin', 'admin', 'admin', _conn=conn)

    @staticmethod
    def _hash_password(password: str, salt: str) -> str:
        return hashlib.sha256((salt + password).encode()).hexdigest()

    def authenticate(self, username: str, password: str) -> Optional[dict]:
        """Zwraca dict z danymi usera lub None jeśli błąd."""
        with self._conn() as conn:
            row = conn.execute(
                'SELECT * FROM users WHERE username = ? AND aktywny = 1',
                (username,),
            ).fetchone()
            if row is None:
                return None
            if self._hash_password(password, row['salt']) != row['password_hash']:
                return None
            conn.execute(
                'UPDATE users SET ostatnie_logowanie = ? WHERE id = ?',
                (datetime.now().isoformat(), row['id']),
            )
            return dict(row)

    def list_users(self) -> list[dict]:
        with self._conn() as conn:
            rows = conn.execute(
                'SELECT id, username, rola, aktywny, utworzony, ostatnie_logowanie '
                'FROM users ORDER BY id'
            ).fetchall()
        return [dict(r) for r in rows]

    def create_user(self, username: str, password: str, rola: str = 'handlowiec',
                    _conn=None) -> bool:
        salt = os.urandom(16).hex()
        pw_hash = self._hash_password(password, salt)
        conn = _conn or self._conn()
        try:
            conn.execute(
                'INSERT INTO users (username, password_hash, salt, rola, aktywny, utworzony) '
                'VALUES (?, ?, ?, ?, 1, ?)',
                (username, pw_hash, salt, rola, datetime.now().isoformat()),
            )
            if _conn is None:
                conn.commit()
            return True
        except sqlite3.IntegrityError:
            return False
        finally:
            if _conn is None:
                conn.close()

    def update_user(self, user_id: int, rola: str = None, aktywny: int = None):
        updates = []
        params = []
        if rola is not None:
            updates.append('rola = ?')
            params.append(rola)
        if aktywny is not None:
            updates.append('aktywny = ?')
            params.append(aktywny)
        if not updates:
            return
        params.append(user_id)
        with self._conn() as conn:
            conn.execute(
                f'UPDATE users SET {", ".join(updates)} WHERE id = ?', params
            )

    def change_password(self, user_id: int, new_password: str):
        salt = os.urandom(16).hex()
        pw_hash = self._hash_password(new_password, salt)
        with self._conn() as conn:
            conn.execute(
                'UPDATE users SET password_hash = ?, salt = ? WHERE id = ?',
                (pw_hash, salt, user_id),
            )

    def delete_user(self, user_id: int):
        """Soft-delete: ustawia aktywny = 0."""
        with self._conn() as conn:
            conn.execute('UPDATE users SET aktywny = 0 WHERE id = ?', (user_id,))
