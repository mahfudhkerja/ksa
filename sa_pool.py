"""Pool service account Google Sheets.
Tetap di bot aktif sampai pemakaiannya >= SA_SWITCH_RATIO x limit/menit,
lalu pindah ke bot berikutnya. Hitungan disimpan di SQLite supaya dibagi
ke semua worker gunicorn dan script subprocess."""
import json
import os
import sqlite3
import threading
import time
from pathlib import Path

import gspread
from google.auth.transport.requests import AuthorizedSession
from google.oauth2.service_account import Credentials

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]
BASE_DIR = Path(__file__).resolve().parent
LIMIT_PER_MIN = int(os.environ.get("SA_LIMIT_PER_MIN", "60"))
SWITCH_RATIO = float(os.environ.get("SA_SWITCH_RATIO", "0.7"))
THRESHOLD = LIMIT_PER_MIN * SWITCH_RATIO
WINDOW = 60
DB_PATH = os.environ.get("SA_POOL_DB", str(BASE_DIR / "sa_pool_usage.db"))


def _cred_files():
    raw = os.environ.get("GOOGLE_CREDENTIALS_FILES") or os.environ.get("GOOGLE_CREDENTIALS_FILE", "credentials.json")
    files = [f.strip() for f in raw.split(",") if f.strip()]
    return [f if os.path.isabs(f) else str(BASE_DIR / f) for f in files]


class SAPool:
    def __init__(self, files):
        self.files = files
        self.n = len(files)
        self._sessions = [None] * self.n
        self._lock = threading.Lock()
        con = self._db()
        con.execute("CREATE TABLE IF NOT EXISTS hits (bot INTEGER, ts REAL, w INTEGER)")
        con.execute("CREATE INDEX IF NOT EXISTS ix_hits ON hits(bot, ts)")
        con.execute("CREATE TABLE IF NOT EXISTS state (k TEXT PRIMARY KEY, v INTEGER)")
        con.execute("INSERT OR IGNORE INTO state VALUES ('cur', 0)")
        con.close()

    def _db(self):
        return sqlite3.connect(DB_PATH, timeout=10, isolation_level=None)

    def session(self, i):
        with self._lock:
            if self._sessions[i] is None:
                creds = Credentials.from_service_account_file(self.files[i], scopes=SCOPES)
                self._sessions[i] = AuthorizedSession(creds)
            return self._sessions[i]

    def _usage(self, con, now):
        con.execute("DELETE FROM hits WHERE ts < ?", (now - WINDOW,))
        usage = [0] * self.n
        for bot, s in con.execute("SELECT bot, SUM(w) FROM hits GROUP BY bot"):
            if 0 <= bot < self.n:
                usage[bot] = s or 0
        return usage

    def pick(self, exclude=()):
        cands = [i for i in range(self.n) if i not in exclude] or list(range(self.n))
        now = time.time()
        con = self._db()
        try:
            con.execute("BEGIN IMMEDIATE")
            usage = self._usage(con, now)
            cur = con.execute("SELECT v FROM state WHERE k='cur'").fetchone()[0] % self.n
            if cur in cands and usage[cur] < THRESHOLD:
                chosen = cur
            else:
                chosen = None
                for step in range(1, self.n + 1):
                    j = (cur + step) % self.n
                    if j in cands and usage[j] < THRESHOLD:
                        chosen = j
                        break
                if chosen is None:
                    chosen = min(cands, key=lambda k: usage[k])
            con.execute("UPDATE state SET v=? WHERE k='cur'", (chosen,))
            con.execute("INSERT INTO hits VALUES (?,?,1)", (chosen, now))
            con.execute("COMMIT")
            return chosen
        except Exception:
            try:
                con.execute("ROLLBACK")
            except Exception:
                pass
            return cands[0]
        finally:
            con.close()

    def penalize(self, i):
        con = self._db()
        try:
            con.execute("INSERT INTO hits VALUES (?,?,?)", (i, time.time(), int(THRESHOLD) + 1))
        finally:
            con.close()

    def stats(self):
        con = self._db()
        try:
            usage = self._usage(con, time.time())
            cur = con.execute("SELECT v FROM state WHERE k='cur'").fetchone()[0] % self.n
        finally:
            con.close()
        bots = []
        for i, f in enumerate(self.files):
            try:
                with open(f, encoding="utf-8") as fh:
                    email = json.load(fh).get("client_email", "")
            except Exception:
                email = ""
            bots.append({
                "bot": i + 1,
                "email": email,
                "request_1_menit": usage[i],
                "persen_dari_limit": round(usage[i] / LIMIT_PER_MIN * 100, 1),
                "aktif": i == cur,
            })
        return {"batas_pindah_persen": SWITCH_RATIO * 100, "bots": bots}


class PooledSession:
    """Session gspread: tiap request memilih bot sendiri; kena 429 -> coba bot lain."""
    def __init__(self, pool):
        self.pool = pool

    def request(self, method, url, **kwargs):
        tried = []
        while True:
            i = self.pool.pick(exclude=tried)
            resp = self.pool.session(i).request(method, url, **kwargs)
            if resp.status_code != 429 or len(tried) + 1 >= self.pool.n:
                return resp
            self.pool.penalize(i)
            tried.append(i)

    def __getattr__(self, name):
        return getattr(self.pool.session(0), name)


_pool = None
_client = None
_lock = threading.Lock()


def get_pool():
    global _pool
    with _lock:
        if _pool is None:
            _pool = SAPool(_cred_files())
        return _pool


def get_client():
    global _client
    pool = get_pool()
    with _lock:
        if _client is None:
            creds = Credentials.from_service_account_file(pool.files[0], scopes=SCOPES)
            _client = gspread.Client(auth=creds, session=PooledSession(pool))
            try:
                _client.http_client.timeout = 20
            except Exception as e:
                print(f"   ⚠️ Tidak bisa set timeout: {e}")
            print(f"   🤖 Pool service account aktif: {pool.n} bot, pindah di >= {SWITCH_RATIO * 100:.0f}%")
        return _client
