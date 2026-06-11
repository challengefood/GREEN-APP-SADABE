from __future__ import annotations

import hashlib
import os
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import date, datetime
from pathlib import Path
from typing import Any, Iterable

import pandas as pd

APP_DIR = Path(__file__).resolve().parent.parent
DATA_DIR = APP_DIR / "data"
DATA_DIR.mkdir(exist_ok=True)
DB_PATH = os.environ.get("GREENAPP_DB_PATH", str(DATA_DIR / "greenapp.db"))

DEFAULT_PROJECTS = [
    ("SOS Lemurs", "Projet de conservation et suivi lié aux lémuriens."),
    ("Darwin Initiatives", "Projet Darwin Initiatives."),
    ("Seacology", "Projet Seacology."),
    ("Rainforest Trust", "Projet Rainforest Trust."),
]

DEFAULT_PARTNERS = [
    ("TGBS (MBG)", "Partenaire technique / Missouri Botanical Garden", "Partenaire"),
    ("MfM", "Partenaire", "Partenaire"),
    ("UWE", "Partenaire académique / technique", "Partenaire"),
    ("Regen", "Partenaire", "Partenaire"),
    ("UNI", "Université / partenaire académique", "Partenaire"),
    ("ENS", "École Normale Supérieure / partenaire académique", "Partenaire"),
]

DEFAULT_TEAMS = [
    ("Direction Programme", "Coordination globale des projets et validation."),
    ("Terrain", "Missions, descentes et activités communautaires."),
    ("MEL", "Suivi-évaluation, collecte de données, reporting."),
    ("Finance & Administration", "Budget, justification, RH et logistique."),
]

STATUSES = ["À planifier", "Planifiée", "En cours", "Terminée", "Reportée", "Annulée"]
PRIORITIES = ["Faible", "Normale", "Haute", "Critique"]
ACTIVITY_TYPES = ["Activité projet", "Mission commune", "Réunion", "Urgence", "Demande bailleur", "Autre"]
USER_ROLES = ["admin", "manager", "staff", "lecture"]


def hash_password(password: str, salt: str | None = None) -> str:
    salt = salt or secrets.token_hex(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt.encode("utf-8"), 120_000)
    return f"pbkdf2_sha256${salt}${dk.hex()}"


def verify_password(password: str, password_hash: str) -> bool:
    try:
        algo, salt, expected = password_hash.split("$", 2)
    except ValueError:
        return False
    return hash_password(password, salt).split("$", 2)[2] == expected


@contextmanager
def get_conn():
    conn = sqlite3.connect(DB_PATH, detect_types=sqlite3.PARSE_DECLTYPES | sqlite3.PARSE_COLNAMES)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    try:
        yield conn
        conn.commit()
    finally:
        conn.close()


def init_db() -> None:
    with get_conn() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                email TEXT UNIQUE NOT NULL,
                poste TEXT,
                role TEXT NOT NULL DEFAULT 'staff',
                status TEXT NOT NULL DEFAULT 'pending',
                password_hash TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                last_login TEXT
            );

            CREATE TABLE IF NOT EXISTS members (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL UNIQUE,
                poste TEXT,
                email TEXT,
                phone TEXT,
                team_default TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                notes TEXT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS projects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS partners (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT,
                type TEXT DEFAULT 'Partenaire',
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS teams (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                description TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );

            CREATE TABLE IF NOT EXISTS team_members (
                team_id INTEGER NOT NULL,
                member_id INTEGER NOT NULL,
                role_in_team TEXT,
                PRIMARY KEY (team_id, member_id),
                FOREIGN KEY (team_id) REFERENCES teams(id) ON DELETE CASCADE,
                FOREIGN KEY (member_id) REFERENCES members(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS activities (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT NOT NULL,
                description TEXT,
                activity_type TEXT DEFAULT 'Activité projet',
                project_id INTEGER,
                partner_id INTEGER,
                location TEXT,
                start_date TEXT,
                end_date TEXT,
                planned_date TEXT,
                status TEXT DEFAULT 'À planifier',
                priority TEXT DEFAULT 'Normale',
                responsible_member_id INTEGER,
                budget_estimated REAL DEFAULT 0,
                source_file TEXT,
                source_sheet TEXT,
                source_row TEXT,
                smart_score REAL DEFAULT 0,
                smart_notes TEXT,
                created_by INTEGER,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (project_id) REFERENCES projects(id),
                FOREIGN KEY (partner_id) REFERENCES partners(id),
                FOREIGN KEY (responsible_member_id) REFERENCES members(id),
                FOREIGN KEY (created_by) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS activity_projects (
                activity_id INTEGER NOT NULL,
                project_id INTEGER NOT NULL,
                PRIMARY KEY(activity_id, project_id),
                FOREIGN KEY(activity_id) REFERENCES activities(id) ON DELETE CASCADE,
                FOREIGN KEY(project_id) REFERENCES projects(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS activity_partners (
                activity_id INTEGER NOT NULL,
                partner_id INTEGER NOT NULL,
                role TEXT,
                PRIMARY KEY(activity_id, partner_id),
                FOREIGN KEY(activity_id) REFERENCES activities(id) ON DELETE CASCADE,
                FOREIGN KEY(partner_id) REFERENCES partners(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS activity_assignments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                activity_id INTEGER NOT NULL,
                member_id INTEGER NOT NULL,
                role_in_activity TEXT,
                is_lead INTEGER DEFAULT 0,
                notes TEXT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(activity_id, member_id),
                FOREIGN KEY(activity_id) REFERENCES activities(id) ON DELETE CASCADE,
                FOREIGN KEY(member_id) REFERENCES members(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS activity_teams (
                activity_id INTEGER NOT NULL,
                team_id INTEGER NOT NULL,
                role_in_activity TEXT,
                PRIMARY KEY(activity_id, team_id),
                FOREIGN KEY(activity_id) REFERENCES activities(id) ON DELETE CASCADE,
                FOREIGN KEY(team_id) REFERENCES teams(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS activity_tasks (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                activity_id INTEGER NOT NULL,
                task_title TEXT NOT NULL,
                task_description TEXT,
                task_date TEXT,
                responsible_member_id INTEGER,
                status TEXT DEFAULT 'À faire',
                notes TEXT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(activity_id) REFERENCES activities(id) ON DELETE CASCADE,
                FOREIGN KEY(responsible_member_id) REFERENCES members(id)
            );

            CREATE TABLE IF NOT EXISTS import_batches (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT,
                imported_by INTEGER,
                imported_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                total_rows INTEGER DEFAULT 0,
                saved_rows INTEGER DEFAULT 0,
                analysis_summary TEXT,
                FOREIGN KEY(imported_by) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS budget_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                activity_id INTEGER,
                title TEXT NOT NULL,
                beneficiary_member_id INTEGER,
                responsible_member_id INTEGER,
                project_id INTEGER,
                partner_id INTEGER,
                location TEXT,
                date_depart TEXT,
                date_retour TEXT,
                nb_days REAL DEFAULT 0,
                total_needs REAL DEFAULT 0,
                received_before REAL DEFAULT 0,
                received_during REAL DEFAULT 0,
                remaining REAL DEFAULT 0,
                status TEXT DEFAULT 'Brouillon',
                justification TEXT,
                created_by INTEGER,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(activity_id) REFERENCES activities(id),
                FOREIGN KEY(beneficiary_member_id) REFERENCES members(id),
                FOREIGN KEY(responsible_member_id) REFERENCES members(id),
                FOREIGN KEY(project_id) REFERENCES projects(id),
                FOREIGN KEY(partner_id) REFERENCES partners(id),
                FOREIGN KEY(created_by) REFERENCES users(id)
            );

            CREATE TABLE IF NOT EXISTS budget_lines (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                request_id INTEGER NOT NULL,
                rubrique TEXT NOT NULL,
                base_calcul TEXT,
                qty REAL DEFAULT 0,
                unit_rate REAL DEFAULT 0,
                total REAL DEFAULT 0,
                received_before REAL DEFAULT 0,
                received_during REAL DEFAULT 0,
                remaining REAL DEFAULT 0,
                comment TEXT,
                FOREIGN KEY(request_id) REFERENCES budget_requests(id) ON DELETE CASCADE
            );

            CREATE TABLE IF NOT EXISTS leave_requests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                member_id INTEGER,
                full_name TEXT NOT NULL,
                poste TEXT,
                leave_type TEXT,
                motif TEXT,
                start_date TEXT,
                end_date TEXT,
                nb_days REAL DEFAULT 0,
                interim TEXT,
                contact TEXT,
                manager TEXT,
                status TEXT DEFAULT 'Brouillon',
                created_by INTEGER,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(member_id) REFERENCES members(id),
                FOREIGN KEY(created_by) REFERENCES users(id)
            );
            """
        )
        seed_defaults(conn)


def seed_defaults(conn: sqlite3.Connection) -> None:
    for name, desc in DEFAULT_PROJECTS:
        conn.execute("INSERT OR IGNORE INTO projects(name, description) VALUES (?, ?)", (name, desc))
    for name, desc, typ in DEFAULT_PARTNERS:
        conn.execute("INSERT OR IGNORE INTO partners(name, description, type) VALUES (?, ?, ?)", (name, desc, typ))
    for name, desc in DEFAULT_TEAMS:
        conn.execute("INSERT OR IGNORE INTO teams(name, description) VALUES (?, ?)", (name, desc))

    admin_count = conn.execute("SELECT COUNT(*) FROM users WHERE role='admin'").fetchone()[0]
    if admin_count == 0:
        conn.execute(
            "INSERT INTO users(full_name, email, poste, role, status, password_hash) VALUES (?, ?, ?, ?, ?, ?)",
            ("Administrateur GREEN'APP", "admin@sadabe.org", "Administrateur", "admin", "approved", hash_password("admin123")),
        )


def query_df(sql: str, params: Iterable[Any] | dict[str, Any] = ()) -> pd.DataFrame:
    with get_conn() as conn:
        return pd.read_sql_query(sql, conn, params=params)


def fetch_one(sql: str, params: Iterable[Any] | dict[str, Any] = ()) -> sqlite3.Row | None:
    with get_conn() as conn:
        return conn.execute(sql, params).fetchone()


def execute(sql: str, params: Iterable[Any] | dict[str, Any] = ()) -> int:
    with get_conn() as conn:
        cur = conn.execute(sql, params)
        return int(cur.lastrowid or 0)


def get_user_by_email(email: str):
    return fetch_one("SELECT * FROM users WHERE LOWER(email)=LOWER(?)", (email.strip(),))


def authenticate(email: str, password: str):
    user = get_user_by_email(email)
    if not user:
        return None, "Compte introuvable."
    if user["status"] != "approved":
        return None, "Compte en attente de validation par un administrateur."
    if not verify_password(password, user["password_hash"]):
        return None, "Mot de passe incorrect."
    with get_conn() as conn:
        conn.execute("UPDATE users SET last_login=CURRENT_TIMESTAMP WHERE id=?", (user["id"],))
    return dict(user), None


def register_user(full_name: str, email: str, poste: str, password: str) -> tuple[bool, str]:
    try:
        with get_conn() as conn:
            conn.execute(
                "INSERT INTO users(full_name, email, poste, role, status, password_hash) VALUES (?, ?, ?, 'staff', 'pending', ?)",
                (full_name.strip(), email.strip().lower(), poste.strip(), hash_password(password)),
            )
        return True, "Compte créé. Il doit être validé par l’administrateur avant l’accès."
    except sqlite3.IntegrityError:
        return False, "Cet email existe déjà."


def option_map(table: str, label_col: str = "name", where: str = "active=1") -> dict[str, int]:
    df = query_df(f"SELECT id, {label_col} AS label FROM {table} WHERE {where} ORDER BY {label_col}")
    return {row["label"]: int(row["id"]) for _, row in df.iterrows()}


def members_map(active_only: bool = True) -> dict[str, int]:
    where = "active=1" if active_only else "1=1"
    df = query_df(f"SELECT id, full_name, poste FROM members WHERE {where} ORDER BY full_name")
    return {f"{r['full_name']} — {r['poste'] or 'Sans poste'}": int(r["id"]) for _, r in df.iterrows()}


def get_activity_full(activity_id: int) -> dict[str, Any] | None:
    df = query_df(
        """
        SELECT a.*, p.name AS project_name, pa.name AS partner_name,
               m.full_name AS responsible_name, m.poste AS responsible_poste
        FROM activities a
        LEFT JOIN projects p ON p.id=a.project_id
        LEFT JOIN partners pa ON pa.id=a.partner_id
        LEFT JOIN members m ON m.id=a.responsible_member_id
        WHERE a.id=?
        """,
        (activity_id,),
    )
    if df.empty:
        return None
    return df.iloc[0].to_dict()


def get_month_bounds(year: int, month: int) -> tuple[str, str]:
    first = date(year, month, 1)
    if month == 12:
        last_exclusive = date(year + 1, 1, 1)
    else:
        last_exclusive = date(year, month + 1, 1)
    return first.isoformat(), last_exclusive.isoformat()


def activity_query_base() -> str:
    return """
        SELECT a.id, a.title, a.description, a.activity_type, a.location, a.start_date, a.end_date, a.planned_date,
               a.status, a.priority, a.budget_estimated, a.smart_score, a.smart_notes, a.source_file,
               p.name AS project, pa.name AS partner,
               m.full_name AS responsible, m.poste AS responsible_poste,
               GROUP_CONCAT(DISTINCT ass_m.full_name || COALESCE(' (' || ass.role_in_activity || ')', '')) AS membres,
               GROUP_CONCAT(DISTINCT t.name || COALESCE(' (' || at.role_in_activity || ')', '')) AS equipes,
               GROUP_CONCAT(DISTINCT apn.name || COALESCE(' - ' || ap.role, '')) AS partenaires_associes,
               GROUP_CONCAT(DISTINCT prj.name) AS projets_associes
        FROM activities a
        LEFT JOIN projects p ON p.id=a.project_id
        LEFT JOIN partners pa ON pa.id=a.partner_id
        LEFT JOIN members m ON m.id=a.responsible_member_id
        LEFT JOIN activity_assignments ass ON ass.activity_id=a.id
        LEFT JOIN members ass_m ON ass_m.id=ass.member_id
        LEFT JOIN activity_teams at ON at.activity_id=a.id
        LEFT JOIN teams t ON t.id=at.team_id
        LEFT JOIN activity_partners ap ON ap.activity_id=a.id
        LEFT JOIN partners apn ON apn.id=ap.partner_id
        LEFT JOIN activity_projects apr ON apr.activity_id=a.id
        LEFT JOIN projects prj ON prj.id=apr.project_id
    """


def get_activities_filtered(
    month_start: str | None = None,
    month_end: str | None = None,
    project_id: int | None = None,
    partner_id: int | None = None,
    member_id: int | None = None,
    status: str | None = None,
) -> pd.DataFrame:
    conditions = []
    params: list[Any] = []
    if month_start and month_end:
        conditions.append(
            """
            (
              (a.planned_date >= ? AND a.planned_date < ?)
              OR (a.start_date IS NOT NULL AND a.end_date IS NOT NULL AND a.start_date < ? AND a.end_date >= ?)
              OR (a.start_date >= ? AND a.start_date < ?)
              OR (a.end_date >= ? AND a.end_date < ?)
            )
            """
        )
        params += [month_start, month_end, month_end, month_start, month_start, month_end, month_start, month_end]
    if project_id:
        conditions.append("(a.project_id=? OR EXISTS(SELECT 1 FROM activity_projects x WHERE x.activity_id=a.id AND x.project_id=?))")
        params += [project_id, project_id]
    if partner_id:
        conditions.append("(a.partner_id=? OR EXISTS(SELECT 1 FROM activity_partners x WHERE x.activity_id=a.id AND x.partner_id=?))")
        params += [partner_id, partner_id]
    if member_id:
        conditions.append("(a.responsible_member_id=? OR EXISTS(SELECT 1 FROM activity_assignments x WHERE x.activity_id=a.id AND x.member_id=?))")
        params += [member_id, member_id]
    if status and status != "Tous":
        conditions.append("a.status=?")
        params.append(status)
    sql = activity_query_base()
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " GROUP BY a.id ORDER BY COALESCE(a.planned_date, a.end_date, a.start_date, '9999-12-31'), a.priority DESC, a.id DESC"
    return query_df(sql, params)


def upsert_activity_links(activity_id: int, project_ids: list[int] | None = None, partner_ids: list[int] | None = None) -> None:
    with get_conn() as conn:
        if project_ids is not None:
            conn.execute("DELETE FROM activity_projects WHERE activity_id=?", (activity_id,))
            for pid in project_ids:
                conn.execute("INSERT OR IGNORE INTO activity_projects(activity_id, project_id) VALUES (?, ?)", (activity_id, pid))
        if partner_ids is not None:
            conn.execute("DELETE FROM activity_partners WHERE activity_id=?", (activity_id,))
            for pid in partner_ids:
                conn.execute("INSERT OR IGNORE INTO activity_partners(activity_id, partner_id, role) VALUES (?, ?, 'Partenaire/Bailleur')", (activity_id, pid))


def insert_activity(data: dict[str, Any], project_ids: list[int] | None = None, partner_ids: list[int] | None = None) -> int:
    allowed = {
        "title", "description", "activity_type", "project_id", "partner_id", "location", "start_date", "end_date", "planned_date",
        "status", "priority", "responsible_member_id", "budget_estimated", "source_file", "source_sheet", "source_row",
        "smart_score", "smart_notes", "created_by"
    }
    fields = [k for k in data if k in allowed and data[k] not in ("", [])]
    values = [data[k] for k in fields]
    placeholders = ",".join(["?"] * len(fields))
    sql = f"INSERT INTO activities({','.join(fields)}) VALUES ({placeholders})"
    with get_conn() as conn:
        cur = conn.execute(sql, values)
        activity_id = int(cur.lastrowid)
        if project_ids:
            for pid in project_ids:
                conn.execute("INSERT OR IGNORE INTO activity_projects(activity_id, project_id) VALUES (?, ?)", (activity_id, pid))
        if partner_ids:
            for pid in partner_ids:
                conn.execute("INSERT OR IGNORE INTO activity_partners(activity_id, partner_id, role) VALUES (?, ?, 'Partenaire/Bailleur')", (activity_id, pid))
        return activity_id
