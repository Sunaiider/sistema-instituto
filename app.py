import io
import math
import os
import secrets
import sqlite3
import ssl
import unicodedata
from datetime import date
from functools import wraps
from pathlib import Path
from urllib.parse import urlsplit
from xml.etree.ElementTree import ParseError
from zipfile import BadZipFile

import click
from flask import (
    Flask,
    abort,
    current_app,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from flask_wtf.csrf import CSRFProtect
from openpyxl import Workbook, load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

try:
    import pymysql
except ImportError:  # pragma: no cover
    pymysql = None

try:
    import psycopg
    from psycopg.rows import tuple_row
except ImportError:  # pragma: no cover
    psycopg = None
    tuple_row = None

try:
    import pg8000.dbapi as pg8000_dbapi
except ImportError:  # pragma: no cover
    pg8000_dbapi = None

MYSQL_INTEGRITY_ERROR = getattr(getattr(pymysql, "err", None), "IntegrityError", None)
POSTGRES_INTEGRITY_ERROR = getattr(psycopg, "IntegrityError", None)
PG8000_INTEGRITY_ERROR = getattr(pg8000_dbapi, "IntegrityError", None)
DATABASE_INTEGRITY_ERRORS = (sqlite3.IntegrityError,)
if MYSQL_INTEGRITY_ERROR is not None:
    DATABASE_INTEGRITY_ERRORS += (MYSQL_INTEGRITY_ERROR,)
if POSTGRES_INTEGRITY_ERROR is not None:
    DATABASE_INTEGRITY_ERRORS += (POSTGRES_INTEGRITY_ERROR,)
if PG8000_INTEGRITY_ERROR is not None:
    DATABASE_INTEGRITY_ERRORS += (PG8000_INTEGRITY_ERROR,)


class StudentWorkbookError(ValueError):
    pass


def parse_student_workbook(upload):
    if not upload.filename or not upload.filename.lower().endswith(".xlsx"):
        raise StudentWorkbookError("Selecciona un archivo de Excel con formato .xlsx.")

    workbook = None
    try:
        workbook = load_workbook(upload, read_only=True, data_only=True)
        rows = workbook.active.iter_rows(values_only=True)
        aliases = {
            "nombre": "name",
            "nombreyapellido": "name",
            "nombresyapellidos": "name",
            "name": "name",
            "sexo": "sex",
            "sex": "sex",
            "codigounicodepersona": "person_code",
            "codigounico": "person_code",
            "codigodepersona": "person_code",
            "personcode": "person_code",
            "modalidad": "modality",
            "nivel": "grade",
            "grado": "grade",
            "grade": "grade",
            "sec": "classroom",
            "seccion": "classroom",
            "aula": "classroom",
            "section": "classroom",
        }
        columns = None
        header_row_number = None
        for row_number, row in enumerate(rows, start=1):
            candidate = {}
            for index, value in enumerate(row):
                normalized = "".join(
                    character
                    for character in unicodedata.normalize(
                        "NFKD", str(value or "").casefold()
                    )
                    if character.isalnum()
                )
                field = aliases.get(normalized)
                if field and field not in candidate:
                    candidate[field] = index
            if {"name", "sex"}.issubset(candidate):
                columns = candidate
                header_row_number = row_number
                break
            if row_number >= 15:
                break

        if columns is None:
            raise StudentWorkbookError(
                "El archivo debe incluir encabezados para Nombre y Sexo."
            )

        group_fields = {"modality", "grade", "classroom"} & columns.keys()
        group_location_fields = {"grade", "classroom"} & group_fields
        grouped_import = group_location_fields == {"grade", "classroom"}
        if group_location_fields and not grouped_import:
            raise StudentWorkbookError(
                "Para asignar grupos desde Excel incluye Nivel y Seccion (o Sec./Aula)."
            )

        students = []
        previous_group_values = {"modality": "", "grade": "", "classroom": ""}
        for row_number, row in enumerate(rows, start=header_row_number + 1):
            values = {}
            for field in ("name", "sex", "person_code"):
                index = columns.get(field)
                value = row[index] if index is not None and index < len(row) else None
                if isinstance(value, float) and value.is_integer():
                    value = int(value)
                values[field] = str(value).strip() if value is not None else ""

            if not any(values.values()):
                continue

            if group_fields:
                for field in group_fields:
                    index = columns[field]
                    value = row[index] if index < len(row) else None
                    value = str(value).strip() if value is not None else ""
                    if value:
                        previous_group_values[field] = value
                    values[field] = value or previous_group_values[field]

            students.append({"row_number": row_number, **values})
        return students, grouped_import
    except StudentWorkbookError:
        raise
    except (BadZipFile, InvalidFileException, OSError, ParseError, ValueError) as exc:
        raise StudentWorkbookError(
            "No se pudo leer el archivo de Excel. Verifica que sea un .xlsx válido."
        ) from exc
    finally:
        if workbook is not None:
            workbook.close()


EVALUATION_PERIODS = (
    "Primer corte evaluativo",
    "Segundo corte evaluativo",
    "Tercer corte evaluativo",
    "Cuarto corte evaluativo",
)
INSTITUTION_NAME = "INSTITUTO DR. MIGUEL JARQUIN VALLEJOS"


class Row(dict):
    def __getitem__(self, key):
        if isinstance(key, int):
            return list(self.values())[key]
        return super().__getitem__(key)

    def __getattr__(self, name):
        try:
            return self[name]
        except KeyError as exc:
            raise AttributeError(name) from exc


class DatabaseCursorWrapper:
    def __init__(self, cursor):
        self._cursor = cursor

    @property
    def rowcount(self):
        return self._cursor.rowcount

    def __getattr__(self, name):
        return getattr(self._cursor, name)

    def fetchone(self):
        row = self._cursor.fetchone()
        if row is None:
            return None
        columns = [
            column.name if hasattr(column, "name") else column[0]
            for column in self._cursor.description
        ]
        return Row(dict(zip(columns, row)))

    def fetchall(self):
        rows = self._cursor.fetchall()
        columns = [
            column.name if hasattr(column, "name") else column[0]
            for column in self._cursor.description
        ]
        return [Row(dict(zip(columns, row))) for row in rows]


class MySQLConnectionWrapper:
    db_type = "mysql"

    def __init__(self, database_url):
        if pymysql is None:
            raise RuntimeError(
                "Para usar MySQL debes instalar PyMySQL: pip install PyMySQL"
            )
        parsed = urlsplit(database_url)
        if parsed.scheme not in {"mysql", "mysql+pymysql"}:
            raise ValueError(
                "La URL de la base debe empezar por mysql:// o mysql+pymysql://"
            )
        db_name = parsed.path.lstrip("/")
        if not db_name:
            raise ValueError(
                "La URL de MySQL debe incluir el nombre de la base de datos."
            )
        self._connection = pymysql.connect(
            host=parsed.hostname or "localhost",
            port=parsed.port or 3306,
            user=parsed.username or "root",
            password=parsed.password or "",
            database=db_name,
            charset="utf8mb4",
            autocommit=False,
            cursorclass=pymysql.cursors.Cursor,
        )
        self.IntegrityError = MYSQL_INTEGRITY_ERROR or sqlite3.IntegrityError

    def close(self):
        self._connection.close()

    def commit(self):
        self._connection.commit()

    def rollback(self):
        self._connection.rollback()

    def executescript(self, script):
        for statement in _split_sql_statements(script):
            if statement.strip():
                self.execute(statement)

    def execute(self, sql, params=()):
        normalized_sql, normalized_params = _normalize_mysql_sql(sql, params)
        cursor = self._connection.cursor()
        cursor.execute(normalized_sql, normalized_params)
        return DatabaseCursorWrapper(cursor)

    def executemany(self, sql, seq_of_params):
        normalized_sql, normalized_params = _normalize_mysql_sql(sql, seq_of_params)
        cursor = self._connection.cursor()
        cursor.executemany(normalized_sql, normalized_params)
        return DatabaseCursorWrapper(cursor)


class PostgreSQLConnectionWrapper:
    db_type = "postgresql"

    def __init__(self, database_url):
        if database_url == "postgresql+pg8000://":
            if pg8000_dbapi is None:
                raise RuntimeError(
                    "Para usar la base PostgreSQL administrada de Wasmer, "
                    "instala pg8000."
                )
            ssl_context = ssl.SSLContext(ssl.PROTOCOL_TLS_CLIENT)
            ssl_context.check_hostname = False
            ssl_context.verify_mode = ssl.CERT_NONE
            self._connection = pg8000_dbapi.connect(
                user=os.environ["DB_USERNAME"],
                password=os.environ["DB_PASSWORD"],
                host=os.environ["DB_HOST"],
                port=int(os.environ["DB_PORT"]),
                database=os.environ["DB_NAME"],
                ssl_context=ssl_context,
                timeout=10,
            )
            self.IntegrityError = PG8000_INTEGRITY_ERROR or sqlite3.IntegrityError
            return

        if psycopg is None:
            raise RuntimeError(
                "Para usar PostgreSQL debes instalar psycopg: "
                "pip install 'psycopg[binary]'"
            )
        if database_url.startswith("postgres://"):
            database_url = "postgresql://" + database_url[len("postgres://") :]
        elif database_url.startswith("postgresql+psycopg://"):
            database_url = "postgresql://" + database_url[
                len("postgresql+psycopg://") :
            ]
        self._connection = psycopg.connect(
            database_url, autocommit=False, row_factory=tuple_row
        )
        self.IntegrityError = POSTGRES_INTEGRITY_ERROR or sqlite3.IntegrityError

    def close(self):
        self._connection.close()

    def commit(self):
        self._connection.commit()

    def rollback(self):
        self._connection.rollback()

    def executescript(self, script):
        for statement in _split_sql_statements(script):
            if statement.strip():
                self.execute(statement)

    def execute(self, sql, params=()):
        normalized_sql, normalized_params = _normalize_postgresql_sql(sql, params)
        cursor = self._connection.cursor()
        cursor.execute(normalized_sql, normalized_params)
        return DatabaseCursorWrapper(cursor)

    def executemany(self, sql, seq_of_params):
        normalized_sql, normalized_params = _normalize_postgresql_sql(sql, seq_of_params)
        cursor = self._connection.cursor()
        cursor.executemany(normalized_sql, normalized_params)
        return DatabaseCursorWrapper(cursor)


def _split_sql_statements(script):
    statements = []
    current = []
    in_single = False
    in_double = False
    for char in script:
        if char == "'" and not in_double:
            in_single = not in_single
        elif char == '"' and not in_single:
            in_double = not in_double
        elif char == ";" and not in_single and not in_double:
            statement = "".join(current).strip()
            if statement:
                statements.append(statement)
            current = []
            continue
        current.append(char)
    tail = "".join(current).strip()
    if tail:
        statements.append(tail)
    return statements


def _normalize_mysql_sql(sql, params=()):
    normalized_sql = str(sql).strip()
    normalized_sql = normalized_sql.replace("BEGIN IMMEDIATE", "START TRANSACTION")
    normalized_sql = normalized_sql.replace("BEGIN", "START TRANSACTION")
    normalized_sql = normalized_sql.replace(" COLLATE NOCASE", "")
    if "?" in normalized_sql:
        if isinstance(params, dict):
            return normalized_sql, params
        if not isinstance(params, (tuple, list)):
            params = (params,)
        normalized_sql = normalized_sql.replace("?", "%s")
    return normalized_sql, params


def _normalize_postgresql_sql(sql, params=()):
    normalized_sql = str(sql).strip()
    normalized_sql = normalized_sql.replace("BEGIN IMMEDIATE", "BEGIN")
    normalized_sql = normalized_sql.replace(" COLLATE NOCASE", "")
    if "?" in normalized_sql:
        if isinstance(params, dict):
            return normalized_sql, params
        if not isinstance(params, (tuple, list)):
            params = (params,)
        normalized_sql = normalized_sql.replace("?", "%s")
    return normalized_sql, params


def _database_type(database_url):
    scheme = urlsplit(str(database_url)).scheme.lower()
    if scheme in {"mysql", "mysql+pymysql"}:
        return "mysql"
    if scheme in {
        "postgres",
        "postgresql",
        "postgresql+pg8000",
        "postgresql+psycopg",
    }:
        return "postgresql"
    return "sqlite"


def _database_url_from_environment(default_database):
    database_url = os.environ.get("DATABASE_URL")
    if database_url:
        return database_url

    managed_database_keys = (
        "DB_HOST",
        "DB_PORT",
        "DB_NAME",
        "DB_USERNAME",
        "DB_PASSWORD",
    )
    use_wasmer_database = (
        os.environ.get("APP_DATABASE_ENGINE", "").lower() == "postgresql"
        or any(os.environ.get(key) for key in managed_database_keys)
    )
    if not use_wasmer_database:
        return default_database

    missing_keys = [key for key in managed_database_keys if not os.environ.get(key)]
    if missing_keys:
        raise RuntimeError(
            "Faltan variables de conexión de la base administrada: "
            + ", ".join(missing_keys)
        )
    return "postgresql+pg8000://"


SCHEMA = """
CREATE TABLE IF NOT EXISTS users (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('admin', 'teacher')),
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS classrooms (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    grade TEXT NOT NULL,
    name TEXT NOT NULL,
    UNIQUE (grade, name)
);

CREATE TABLE IF NOT EXISTS students (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    classroom_id INTEGER NOT NULL REFERENCES classrooms(id),
    name TEXT NOT NULL,
    modality TEXT NOT NULL DEFAULT 'Regular',
    sex TEXT,
    person_code TEXT
);

CREATE TABLE IF NOT EXISTS student_enrollments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id),
    classroom_id INTEGER NOT NULL REFERENCES classrooms(id),
    school_year TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (student_id, school_year)
);

CREATE TABLE IF NOT EXISTS modalities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS subjects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS grades (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id),
    subject_id INTEGER NOT NULL REFERENCES subjects(id),
    school_year TEXT NOT NULL,
    teacher_id INTEGER REFERENCES users(id),
    teacher_name TEXT NOT NULL DEFAULT '',
    score REAL NOT NULL,
    term TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (student_id, subject_id, term, school_year)
);

CREATE TABLE IF NOT EXISTS teacher_assignments (
    teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    classroom_id INTEGER NOT NULL REFERENCES classrooms(id) ON DELETE CASCADE,
    PRIMARY KEY (teacher_id, subject_id, classroom_id)
);
"""

SCHEMA_MYSQL = """
CREATE TABLE IF NOT EXISTS users (
    id INT PRIMARY KEY AUTO_INCREMENT,
    username VARCHAR(255) NOT NULL UNIQUE,
    password_hash VARCHAR(255) NOT NULL,
    role VARCHAR(20) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS classrooms (
    id INT PRIMARY KEY AUTO_INCREMENT,
    grade VARCHAR(255) NOT NULL,
    name VARCHAR(255) NOT NULL,
    UNIQUE KEY unique_classroom (grade, name)
);

CREATE TABLE IF NOT EXISTS students (
    id INT PRIMARY KEY AUTO_INCREMENT,
    classroom_id INT NOT NULL,
    name VARCHAR(255) NOT NULL,
    modality VARCHAR(100) NOT NULL DEFAULT 'Regular',
    sex VARCHAR(1) NULL,
    person_code VARCHAR(100) NULL,
    CONSTRAINT fk_students_classroom FOREIGN KEY (classroom_id) REFERENCES classrooms(id)
);

CREATE TABLE IF NOT EXISTS student_enrollments (
    id INT PRIMARY KEY AUTO_INCREMENT,
    student_id INT NOT NULL,
    classroom_id INT NOT NULL,
    school_year VARCHAR(4) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY unique_student_enrollment (student_id, school_year),
    CONSTRAINT fk_student_enrollments_student FOREIGN KEY (student_id) REFERENCES students(id),
    CONSTRAINT fk_student_enrollments_classroom FOREIGN KEY (classroom_id) REFERENCES classrooms(id)
);

CREATE TABLE IF NOT EXISTS modalities (
    id INT PRIMARY KEY AUTO_INCREMENT,
    name VARCHAR(100) NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS subjects (
    id INT PRIMARY KEY AUTO_INCREMENT,
    name VARCHAR(255) NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS grades (
    id INT PRIMARY KEY AUTO_INCREMENT,
    student_id INT NOT NULL,
    subject_id INT NOT NULL,
    school_year VARCHAR(4) NOT NULL,
    teacher_id INT NULL,
    teacher_name VARCHAR(255) NOT NULL DEFAULT '',
    score VARCHAR(30) NOT NULL,
    term VARCHAR(255) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE KEY unique_grade (student_id, subject_id, term, school_year),
    CONSTRAINT fk_grades_student FOREIGN KEY (student_id) REFERENCES students(id),
    CONSTRAINT fk_grades_subject FOREIGN KEY (subject_id) REFERENCES subjects(id),
    CONSTRAINT fk_grades_teacher FOREIGN KEY (teacher_id) REFERENCES users(id)
);

CREATE TABLE IF NOT EXISTS teacher_assignments (
    teacher_id INT NOT NULL,
    subject_id INT NOT NULL,
    classroom_id INT NOT NULL,
    PRIMARY KEY (teacher_id, subject_id, classroom_id),
    CONSTRAINT fk_teacher_assignments_teacher FOREIGN KEY (teacher_id) REFERENCES users(id) ON DELETE CASCADE,
    CONSTRAINT fk_teacher_assignments_subject FOREIGN KEY (subject_id) REFERENCES subjects(id) ON DELETE CASCADE,
    CONSTRAINT fk_teacher_assignments_classroom FOREIGN KEY (classroom_id) REFERENCES classrooms(id) ON DELETE CASCADE
);
"""


SCHEMA_POSTGRESQL = """
CREATE TABLE IF NOT EXISTS users (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL CHECK (role IN ('admin', 'teacher')),
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP
);

CREATE TABLE IF NOT EXISTS classrooms (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    grade TEXT NOT NULL,
    name TEXT NOT NULL,
    UNIQUE (grade, name)
);

CREATE TABLE IF NOT EXISTS students (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    classroom_id BIGINT NOT NULL REFERENCES classrooms(id),
    name TEXT NOT NULL,
    modality TEXT NOT NULL DEFAULT 'Regular',
    sex TEXT,
    person_code TEXT
);

CREATE TABLE IF NOT EXISTS student_enrollments (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id),
    classroom_id BIGINT NOT NULL REFERENCES classrooms(id),
    school_year VARCHAR(4) NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (student_id, school_year)
);

CREATE TABLE IF NOT EXISTS modalities (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS subjects (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS grades (
    id BIGINT GENERATED BY DEFAULT AS IDENTITY PRIMARY KEY,
    student_id BIGINT NOT NULL REFERENCES students(id),
    subject_id BIGINT NOT NULL REFERENCES subjects(id),
    school_year VARCHAR(4) NOT NULL,
    teacher_id BIGINT REFERENCES users(id),
    teacher_name TEXT NOT NULL DEFAULT '',
    score TEXT NOT NULL,
    term TEXT NOT NULL,
    created_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TIMESTAMP NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (student_id, subject_id, term, school_year)
);

CREATE TABLE IF NOT EXISTS teacher_assignments (
    teacher_id BIGINT NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    subject_id BIGINT NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    classroom_id BIGINT NOT NULL REFERENCES classrooms(id) ON DELETE CASCADE,
    PRIMARY KEY (teacher_id, subject_id, classroom_id)
);
"""


def create_app(test_config=None):
    app = Flask(
        __name__,
        instance_path=os.environ.get("FLASK_INSTANCE_PATH"),
        instance_relative_config=True,
    )
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
        DATABASE=_database_url_from_environment(
            os.path.join(app.instance_path, "instituto.sqlite")
        ),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "").lower() == "true",
    )
    if test_config:
        app.config.update(test_config)
    app.config["DB_TYPE"] = _database_type(app.config["DATABASE"])
    app.jinja_env.filters["grade_score"] = format_grade_score
    app.jinja_env.globals["institution_name"] = INSTITUTION_NAME

    Path(app.instance_path).mkdir(parents=True, exist_ok=True)
    CSRFProtect(app)

    @app.teardown_appcontext
    def close_db(error=None):
        db = g.pop("db", None)
        if db is not None:
            db.close()

    with app.app_context():
        db = get_db()
        schema = {
            "mysql": SCHEMA_MYSQL,
            "postgresql": SCHEMA_POSTGRESQL,
        }.get(app.config["DB_TYPE"], SCHEMA)
        db.executescript(schema)
        migrate_student_modality(db)
        migrate_grade_teacher_snapshot(db)
        migrate_modalities(db)
        migrate_student_identity(db)
        ensure_current_enrollments(db)
        migrate_grade_school_year(db)
        db.commit()

    @app.before_request
    def load_logged_in_user():
        user_id = session.get("user_id")
        g.user = (
            get_db().execute(
                "SELECT id, username, role FROM users WHERE id = ?", (user_id,)
            ).fetchone()
            if user_id
            else None
        )

    @app.cli.command("init-admin")
    @click.option("--username", prompt=True, help="Nombre de usuario del administrador.")
    def init_admin(username):
        """Crea la primera cuenta de administrador."""
        if get_db().execute("SELECT 1 FROM users WHERE role = 'admin' LIMIT 1").fetchone():
            raise click.ClickException("Ya existe una cuenta de administrador.")
        username = username.strip()
        if not username:
            raise click.ClickException("El nombre de usuario no puede quedar vacio.")
        password = click.prompt("Contrasena (minimo 6 caracteres)", hide_input=True)
        if len(password) < 6:
            raise click.ClickException("La contrasena debe tener al menos 6 caracteres.")
        try:
            get_db().execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'admin')",
                (username.strip(), generate_password_hash(password)),
            )
            get_db().commit()
        except DATABASE_INTEGRITY_ERRORS as exc:
            raise click.ClickException("Ese nombre de usuario ya existe.") from exc
        click.echo("Administrador creado.")

    def login_required(view):
        @wraps(view)
        def wrapped_view(**kwargs):
            if g.user is None:
                return redirect(url_for("login", next=request.path))
            return view(**kwargs)

        return wrapped_view

    def role_required(role):
        def decorator(view):
            @wraps(view)
            def wrapped_view(**kwargs):
                if g.user is None:
                    return redirect(url_for("login", next=request.path))
                if g.user["role"] != role:
                    abort(403)
                return view(**kwargs)

            return wrapped_view

        return decorator

    @app.route("/login", methods=("GET", "POST"))
    def login():
        if g.user:
            return redirect(url_for("dashboard"))
        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            user = get_db().execute(
                "SELECT id, username, password_hash, role FROM users WHERE username = ?",
                (username,),
            ).fetchone()
            if user is None or not check_password_hash(user["password_hash"], password):
                flash("Usuario o contrasena incorrectos.", "error")
            else:
                session.clear()
                session["user_id"] = user["id"]
                destination = request.args.get("next", "")
                parsed_destination = urlsplit(destination)
                if (
                    not destination.startswith("/")
                    or destination.startswith("//")
                    or "\\" in destination
                    or parsed_destination.netloc
                ):
                    destination = url_for("dashboard")
                return redirect(destination)
        return render_template("login.html")

    @app.post("/logout")
    @login_required
    def logout():
        session.clear()
        return redirect(url_for("login"))

    @app.get("/")
    @login_required
    def dashboard():
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        if g.user["role"] == "admin":
            grade_count = db.execute("SELECT COUNT(*) FROM grades").fetchone()[0]
            student_count = db.execute("SELECT COUNT(*) FROM students").fetchone()[0]
            classroom_count = db.execute("SELECT COUNT(*) FROM classrooms").fetchone()[0]
        else:
            grade_count = db.execute(
                """SELECT COUNT(*) FROM grades g
                   JOIN student_enrollments e
                     ON e.student_id = g.student_id AND e.school_year = g.school_year
                   JOIN teacher_assignments ta
                     ON ta.teacher_id = ? AND ta.subject_id = g.subject_id
                    AND ta.classroom_id = e.classroom_id""",
                (g.user["id"],),
            ).fetchone()[0]
            student_count = db.execute(
                """SELECT COUNT(DISTINCT s.id) FROM students s
                   JOIN teacher_assignments ta ON ta.classroom_id = s.classroom_id
                   WHERE ta.teacher_id = ?""",
                (g.user["id"],),
            ).fetchone()[0]
            classroom_count = db.execute(
                "SELECT COUNT(DISTINCT classroom_id) FROM teacher_assignments WHERE teacher_id = ?",
                (g.user["id"],),
            ).fetchone()[0]
        return render_template(
            "dashboard.html",
            grade_count=grade_count,
            student_count=student_count,
            classroom_count=classroom_count,
        )

    @app.route("/grades/new", methods=("GET", "POST"))
    @role_required("teacher")
    def new_grades():
        db = get_db()
        school_year = str(date.today().year)
        ensure_current_enrollments(db, school_year)
        db.commit()
        classrooms = db.execute(
            """SELECT DISTINCT c.id, c.grade, c.name
               FROM teacher_assignments ta
               JOIN classrooms c ON c.id = ta.classroom_id
               WHERE ta.teacher_id = ?
               ORDER BY c.grade, c.name""",
            (g.user["id"],),
        ).fetchall()
        selected_classroom = request.values.get("classroom_id", type=int)
        subject_id = request.values.get("subject_id", type=int)
        term = request.values.get("term", "").strip()
        subjects = []
        if selected_classroom:
            subjects = db.execute(
                """SELECT s.id, s.name
                   FROM teacher_assignments ta
                   JOIN subjects s ON s.id = ta.subject_id
                   WHERE ta.teacher_id = ? AND ta.classroom_id = ?
                   ORDER BY s.name COLLATE NOCASE""",
                (g.user["id"], selected_classroom),
            ).fetchall()
            if not subjects:
                abort(403)
        if selected_classroom and subject_id and not db.execute(
            """SELECT 1 FROM teacher_assignments
               WHERE teacher_id = ? AND classroom_id = ? AND subject_id = ?""",
            (g.user["id"], selected_classroom, subject_id),
        ).fetchone():
            abort(403)

        students = []
        existing_ids = set()

        if selected_classroom:
            students = db.execute(
                "SELECT id, name FROM students WHERE classroom_id = ? ORDER BY name",
                (selected_classroom,),
            ).fetchall()
            if subject_id and term:
                existing_ids = {
                    row["student_id"]
                    for row in db.execute(
                        """SELECT student_id FROM grades
                           WHERE student_id IN (
                               SELECT id FROM students WHERE classroom_id = ?
                           ) AND subject_id = ? AND term = ? AND school_year = ?""",
                        (selected_classroom, subject_id, term, school_year),
                    ).fetchall()
                }

        if request.method == "POST":
            if not subject_id or not db.execute(
                "SELECT 1 FROM subjects WHERE id = ?", (subject_id,)
            ).fetchone():
                flash("Selecciona una materia valida.", "error")
            elif term not in EVALUATION_PERIODS:
                flash("Selecciona un corte evaluativo valido.", "error")
            elif not students:
                flash("Este grado y aula aun no tienen alumnos.", "error")
            else:
                entries = []
                invalid_score = False
                for student in students:
                    raw_score = request.form.get(f"score_{student['id']}", "").strip()
                    if not raw_score or student["id"] in existing_ids:
                        continue
                    try:
                        score = parse_grade_score(raw_score)
                    except ValueError:
                        invalid_score = True
                        break
                    entries.append(
                        (
                            student["id"],
                            subject_id,
                            school_year,
                            g.user["id"],
                            score,
                            term,
                        )
                    )

                if invalid_score:
                    flash(
                        "Cada nota debe ser un numero entre 0 y 100 o texto de hasta 3 caracteres.",
                        "error",
                    )
                elif not entries:
                    flash("No hay notas nuevas para guardar.", "error")
                else:
                    try:
                        db.executemany(
                            """INSERT INTO grades
                               (student_id, subject_id, school_year, teacher_id,
                                teacher_name, score, term)
                               VALUES (?, ?, ?, ?, ?, ?, ?)""",
                            [
                                (
                                    *entry[:4],
                                    g.user["username"],
                                    score_for_database(entry[4]),
                                    entry[5],
                                )
                                for entry in entries
                            ],
                        )
                        db.commit()
                    except DATABASE_INTEGRITY_ERRORS:
                        db.rollback()
                        flash(
                            "Una o mas notas ya fueron registradas para ese corte evaluativo. "
                            "Las notas guardadas no se pueden cambiar desde la cuenta del maestro.",
                            "error",
                        )
                    else:
                        flash(f"Se guardaron {len(entries)} notas. Ya quedaron bloqueadas para el maestro.", "success")
                        return redirect(url_for("grade_list"))

        return render_template(
            "grades_new.html",
            classrooms=classrooms,
            subjects=subjects,
            students=students,
            selected_classroom=selected_classroom,
            subject_id=subject_id,
            term=term,
            school_year=school_year,
            periods=EVALUATION_PERIODS,
            existing_ids=existing_ids,
        )

    @app.get("/grades")
    @login_required
    def grade_list():
        summaries = query_grade_summaries(
            teacher_id=g.user["id"] if g.user["role"] == "teacher" else None
        )
        return render_template("grades_list.html", summaries=summaries)

    @app.get("/grades/classroom/<int:classroom_id>")
    @login_required
    def classroom_grades(classroom_id):
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        subject_id = request.args.get("subject_id", type=int)
        if not subject_id:
            abort(404)
        classroom = db.execute(
            "SELECT id, grade, name FROM classrooms WHERE id = ?", (classroom_id,)
        ).fetchone()
        subject = db.execute(
            "SELECT id, name FROM subjects WHERE id = ?", (subject_id,)
        ).fetchone()
        if classroom is None or subject is None:
            abort(404)
        if g.user["role"] == "teacher" and not teacher_has_assignment(
            classroom_id, subject_id, g.user["id"]
        ):
            abort(403)

        sql = """SELECT g.id, g.score, g.term, g.created_at, g.updated_at,
                        g.school_year, s.name AS student_name, c.grade, c.name AS classroom,
                        sub.name AS subject,
                        COALESCE(NULLIF(g.teacher_name, ''), u.username, 'Cuenta eliminada') AS teacher
                 FROM grades g
                 JOIN students s ON s.id = g.student_id
                 JOIN student_enrollments e
                   ON e.student_id = g.student_id AND e.school_year = g.school_year
                 JOIN classrooms c ON c.id = e.classroom_id
                 JOIN subjects sub ON sub.id = g.subject_id
                 LEFT JOIN users u ON u.id = g.teacher_id
                 WHERE c.id = ? AND sub.id = ?"""
        params = [classroom_id, subject_id]
        sql += " ORDER BY g.school_year DESC, g.term COLLATE NOCASE, s.name COLLATE NOCASE"
        grades = db.execute(sql, params).fetchall()
        return render_template(
            "grades_detail.html",
            classroom=classroom,
            subject=subject,
            grades=grades,
        )

    @app.get("/grades/export.xlsx")
    @login_required
    def export_grades():
        db = get_db()
        if g.user["role"] == "teacher":
            subject_rows = db.execute(
                """SELECT DISTINCT sub.name
                   FROM subjects sub
                   JOIN teacher_assignments ta ON ta.subject_id = sub.id
                   WHERE ta.teacher_id = ?
                   ORDER BY sub.name""",
                (g.user["id"],),
            ).fetchall()
        else:
            subject_rows = db.execute(
                "SELECT name FROM subjects ORDER BY name"
            ).fetchall()
        rows = query_grade_rows(
            assigned_teacher_id=g.user["id"] if g.user["role"] == "teacher" else None
        )
        return create_grade_export(
            rows,
            f"notas_{date.today().isoformat()}.xlsx",
            [subject["name"] for subject in subject_rows],
        )

    @app.get("/grades/export/<int:classroom_id>/<int:subject_id>.xlsx")
    @login_required
    def export_classroom_subject(classroom_id, subject_id):
        db = get_db()
        classroom = db.execute(
            "SELECT grade, name FROM classrooms WHERE id = ?", (classroom_id,)
        ).fetchone()
        subject = db.execute(
            "SELECT name FROM subjects WHERE id = ?", (subject_id,)
        ).fetchone()
        if classroom is None or subject is None:
            abort(404)
        if g.user["role"] == "teacher" and not teacher_has_assignment(
            classroom_id, subject_id, g.user["id"]
        ):
            abort(403)
        rows = query_grade_rows(
            assigned_teacher_id=g.user["id"] if g.user["role"] == "teacher" else None,
            classroom_id=classroom_id,
            subject_id=subject_id,
        )
        filename = secure_filename(
            f"notas_{subject['name']}_{classroom['grade']}_{classroom['name']}.xlsx"
        )
        return create_grade_export(rows, filename, [subject["name"]])

    @app.get("/admin")
    @role_required("admin")
    def admin():
        db = get_db()
        users = db.execute(
            "SELECT id, username, role FROM users ORDER BY role, username"
        ).fetchall()
        classrooms = db.execute(
            """SELECT c.id, c.grade, c.name, COUNT(s.id) AS student_count
               FROM classrooms c LEFT JOIN students s ON s.classroom_id = c.id
               GROUP BY c.id ORDER BY c.grade, c.name"""
        ).fetchall()
        subjects = db.execute("SELECT id, name FROM subjects ORDER BY name").fetchall()
        assignment_options = [
            {
                "subject_id": subject["id"],
                "subject": subject["name"],
                "classroom_id": classroom["id"],
                "grade": classroom["grade"],
                "classroom": classroom["name"],
            }
            for subject in subjects
            for classroom in classrooms
        ]
        assigned_by_teacher = {}
        for row in db.execute(
            """SELECT teacher_id, subject_id, classroom_id
               FROM teacher_assignments"""
        ).fetchall():
            assigned_by_teacher.setdefault(row["teacher_id"], set()).add(
                f"{row['subject_id']}:{row['classroom_id']}"
            )
        return render_template(
            "admin.html",
            users=users,
            classrooms=classrooms,
            subjects=subjects,
            assignment_options=assignment_options,
            assigned_by_teacher=assigned_by_teacher,
            current_user_id=g.user["id"],
        )

    @app.get("/admin/students")
    @role_required("admin")
    def admin_students():
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        classrooms = db.execute(
            "SELECT id, grade, name FROM classrooms ORDER BY grade, name"
        ).fetchall()
        students = db.execute(
            """SELECT s.id, s.classroom_id, s.name, s.modality, s.sex,
                      s.person_code, c.grade, c.name AS classroom
               FROM students s
               JOIN classrooms c ON c.id = s.classroom_id
               ORDER BY c.grade, c.name, s.name"""
        ).fetchall()
        modalities = db.execute(
            "SELECT name FROM modalities ORDER BY name"
        ).fetchall()
        return render_template(
            "admin_students.html",
            classrooms=classrooms,
            students=students,
            modalities=modalities,
            current_school_year=str(date.today().year),
        )

    @app.route("/admin/students/<int:student_id>/record", methods=("GET", "POST"))
    @role_required("admin")
    def student_record(student_id):
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        student = db.execute(
            """SELECT s.id, s.classroom_id, s.name, s.modality, s.sex, s.person_code,
                      c.grade, c.name AS classroom
               FROM students s
               JOIN classrooms c ON c.id = s.classroom_id
               WHERE s.id = ?""",
            (student_id,),
        ).fetchone()
        if student is None:
            abort(404)

        if request.method == "POST":
            classroom_id = request.form.get("classroom_id", type=int)
            school_year = request.form.get("school_year", "").strip()
            existing_enrollment = db.execute(
                """SELECT classroom_id FROM student_enrollments
                   WHERE student_id = ? AND school_year = ?""",
                (student_id, school_year),
            ).fetchone()
            valid_year = (
                len(school_year) == 4
                and school_year.isdigit()
                and 1900 <= int(school_year) <= 2100
            )
            if not valid_year:
                flash("Indica un año escolar válido de cuatro cifras.", "error")
            elif not classroom_id or not db.execute(
                "SELECT 1 FROM classrooms WHERE id = ?", (classroom_id,)
            ).fetchone():
                flash("Selecciona un grado y aula válidos.", "error")
            elif existing_enrollment and existing_enrollment["classroom_id"] == classroom_id:
                flash("Ya existe una matrícula para ese año escolar.", "error")
            else:
                try:
                    db.execute("BEGIN IMMEDIATE")
                    if existing_enrollment:
                        db.execute(
                            """UPDATE student_enrollments SET classroom_id = ?
                               WHERE student_id = ? AND school_year = ?""",
                            (classroom_id, student_id, school_year),
                        )
                    else:
                        db.execute(
                            """INSERT INTO student_enrollments
                               (student_id, classroom_id, school_year)
                               VALUES (?, ?, ?)""",
                            (student_id, classroom_id, school_year),
                        )
                    if school_year == str(date.today().year):
                        db.execute(
                            "UPDATE students SET classroom_id = ? WHERE id = ?",
                            (classroom_id, student_id),
                        )
                    db.commit()
                except DATABASE_INTEGRITY_ERRORS:
                    db.rollback()
                    flash("No se pudo guardar la matrícula del estudiante.", "error")
                else:
                    flash(
                        "Matrícula actualizada en el expediente."
                        if existing_enrollment
                        else "Matrícula agregada al expediente.",
                        "success",
                    )
                    return redirect(
                        url_for("student_record", student_id=student_id)
                    )

        enrollments = db.execute(
            """SELECT e.school_year, c.grade, c.name AS classroom
               FROM student_enrollments e
               JOIN classrooms c ON c.id = e.classroom_id
               WHERE e.student_id = ?
               ORDER BY e.school_year DESC""",
            (student_id,),
        ).fetchall()
        grades = db.execute(
            """SELECT g.school_year, g.term, g.score, g.created_at,
                      sub.name AS subject,
                      COALESCE(NULLIF(g.teacher_name, ''), u.username, 'Cuenta eliminada') AS teacher,
                      c.grade, c.name AS classroom
               FROM grades g
               JOIN student_enrollments e
                 ON e.student_id = g.student_id AND e.school_year = g.school_year
               JOIN classrooms c ON c.id = e.classroom_id
               JOIN subjects sub ON sub.id = g.subject_id
               LEFT JOIN users u ON u.id = g.teacher_id
               WHERE g.student_id = ?
               ORDER BY g.school_year DESC, g.term, sub.name""",
            (student_id,),
        ).fetchall()
        classrooms = db.execute(
            "SELECT id, grade, name FROM classrooms ORDER BY grade, name"
        ).fetchall()
        return render_template(
            "student_record.html",
            student=student,
            enrollments=enrollments,
            grades=grades,
            classrooms=classrooms,
            school_year=str(date.today().year),
        )

    @app.post("/admin/students/enrollments/bulk")
    @role_required("admin")
    def bulk_student_enrollment():
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        source_classroom_id = request.form.get("source_classroom_id", type=int)
        target_classroom_id = request.form.get("target_classroom_id", type=int)
        school_year = request.form.get("school_year", "").strip()
        valid_year = (
            len(school_year) == 4
            and school_year.isdigit()
            and 1900 <= int(school_year) <= 2100
        )
        if not valid_year:
            flash("Indica un año escolar válido de cuatro cifras.", "error")
            return redirect(url_for("admin_students"))
        if not source_classroom_id or not target_classroom_id:
            flash("Selecciona el grado/aula de origen y el grupo destino.", "error")
            return redirect(url_for("admin_students"))

        classrooms = db.execute(
            "SELECT id FROM classrooms WHERE id IN (?, ?)",
            (source_classroom_id, target_classroom_id),
        ).fetchall()
        if len(classrooms) != len({source_classroom_id, target_classroom_id}):
            flash("El grado o aula seleccionado ya no existe.", "error")
            return redirect(url_for("admin_students"))

        students = db.execute(
            "SELECT id FROM students WHERE classroom_id = ? ORDER BY id",
            (source_classroom_id,),
        ).fetchall()
        if not students:
            flash("No hay estudiantes en el grado/aula de origen.", "error")
            return redirect(url_for("admin_students"))

        created_count = 0
        updated_count = 0
        unchanged_count = 0
        try:
            db.execute("BEGIN IMMEDIATE")
            for student in students:
                existing = db.execute(
                    """SELECT classroom_id FROM student_enrollments
                       WHERE student_id = ? AND school_year = ?""",
                    (student["id"], school_year),
                ).fetchone()
                if existing and existing["classroom_id"] == target_classroom_id:
                    unchanged_count += 1
                elif existing:
                    db.execute(
                        """UPDATE student_enrollments SET classroom_id = ?
                           WHERE student_id = ? AND school_year = ?""",
                        (target_classroom_id, student["id"], school_year),
                    )
                    updated_count += 1
                else:
                    db.execute(
                        """INSERT INTO student_enrollments
                           (student_id, classroom_id, school_year)
                           VALUES (?, ?, ?)""",
                        (student["id"], target_classroom_id, school_year),
                    )
                    created_count += 1
                if school_year == str(date.today().year):
                    db.execute(
                        "UPDATE students SET classroom_id = ? WHERE id = ?",
                        (target_classroom_id, student["id"]),
                    )
            db.commit()
        except DATABASE_INTEGRITY_ERRORS:
            db.rollback()
            flash("No se pudo registrar la matrícula para todo el grupo.", "error")
        else:
            flash(
                f"Matrículas del grupo procesadas para {school_year}: "
                f"{created_count} nuevas, {updated_count} actualizadas y "
                f"{unchanged_count} sin cambios.",
                "success",
            )
        return redirect(url_for("admin_students"))

    @app.get("/admin/students/template.xlsx")
    @role_required("admin")
    def student_import_template():
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "Estudiantes"
        sheet.append(["Nombre", "Sexo", "Código de persona"])
        sheet.column_dimensions["A"].width = 32
        sheet.column_dimensions["B"].width = 12
        sheet.column_dimensions["C"].width = 28

        output = io.BytesIO()
        workbook.save(output)
        output.seek(0)
        return send_file(
            output,
            as_attachment=True,
            download_name="plantilla_estudiantes.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    @app.post("/admin/users")
    @role_required("admin")
    def create_user():
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        role = request.form.get("role", "")
        assignments, valid_assignments = parse_assignment_values(
            request.form.getlist("assignments")
        )
        if not username or len(password) < 6 or role not in ("admin", "teacher"):
            flash("Completa el usuario, una contrasena de al menos 6 caracteres y un rol valido.", "error")
        elif role == "teacher" and (not valid_assignments or not assignments):
            flash("Selecciona al menos una combinacion de materia, grado y aula valida para el maestro.", "error")
        else:
            try:
                db = get_db()
                db.execute("BEGIN IMMEDIATE")
                if role == "teacher" and not assignment_values_exist(db, assignments):
                    db.rollback()
                    flash("Una o mas asignaciones no existen. Actualiza la pagina y vuelve a intentarlo.", "error")
                    return redirect(url_for("admin"))
                cursor = db.execute(
                    """INSERT INTO users (username, password_hash, role)
                       VALUES (?, ?, ?)""" + (
                           " RETURNING id"
                           if current_app.config["DB_TYPE"] == "postgresql"
                           else ""
                       ),
                    (username, generate_password_hash(password), role),
                )
                if role == "teacher":
                    teacher_id = (
                        cursor.fetchone()["id"]
                        if current_app.config["DB_TYPE"] == "postgresql"
                        else cursor.lastrowid
                    )
                    db.executemany(
                        """INSERT INTO teacher_assignments
                           (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                        [
                            (teacher_id, subject_id, classroom_id)
                            for subject_id, classroom_id in assignments
                        ],
                    )
                db.commit()
            except DATABASE_INTEGRITY_ERRORS:
                get_db().rollback()
                flash("Ese nombre de usuario ya esta en uso.", "error")
            else:
                flash(
                    "Cuenta creada con sus asignaciones."
                    if role == "teacher"
                    else "Cuenta creada.",
                    "success",
                )
        return redirect(url_for("admin"))

    @app.route("/admin/users/<int:user_id>/edit", methods=("GET", "POST"))
    @role_required("admin")
    def edit_user(user_id):
        db = get_db()
        account = db.execute(
            "SELECT id, username, role FROM users WHERE id = ?", (user_id,)
        ).fetchone()
        if account is None:
            abort(404)

        classrooms = db.execute(
            "SELECT id, grade, name FROM classrooms ORDER BY grade, name"
        ).fetchall()
        subjects = db.execute("SELECT id, name FROM subjects ORDER BY name").fetchall()
        assignment_options = [
            {
                "subject_id": subject["id"],
                "subject": subject["name"],
                "classroom_id": classroom["id"],
                "grade": classroom["grade"],
                "classroom": classroom["name"],
            }
            for subject in subjects
            for classroom in classrooms
        ]
        current_assignments = {
            f"{row['subject_id']}:{row['classroom_id']}"
            for row in db.execute(
                """SELECT subject_id, classroom_id FROM teacher_assignments
                   WHERE teacher_id = ?""",
                (user_id,),
            ).fetchall()
        }
        account_values = dict(account)

        if request.method == "POST":
            username = request.form.get("username", "").strip()
            password = request.form.get("password", "")
            role = request.form.get("role", "")
            assignments, valid_assignments = parse_assignment_values(
                request.form.getlist("assignments")
            )
            account_values.update(username=username, role=role)
            current_assignments = {
                f"{subject_id}:{classroom_id}"
                for subject_id, classroom_id in assignments
            }

            error = None
            if not username or len(username) > 80:
                error = "Indica un nombre de usuario de hasta 80 caracteres."
            elif password and len(password) < 6:
                error = "La nueva contraseña debe tener al menos 6 caracteres."
            elif role not in ("admin", "teacher"):
                error = "Selecciona un rol valido."
            elif user_id == g.user["id"] and role != account["role"]:
                error = "No puedes cambiar el rol de tu propia cuenta."
            elif role == "teacher" and (not valid_assignments or not assignments):
                error = "Selecciona al menos una combinacion de materia, grado y aula valida para el maestro."

            if error:
                flash(error, "error")
            else:
                try:
                    db.execute("BEGIN IMMEDIATE")
                    if role == "teacher" and not assignment_values_exist(
                        db, assignments
                    ):
                        db.rollback()
                        flash(
                            "Una o mas asignaciones ya no existen. Actualiza la pagina y vuelve a intentarlo.",
                            "error",
                        )
                    elif (
                        account["role"] == "admin"
                        and role == "teacher"
                        and db.execute(
                            "SELECT COUNT(*) FROM users WHERE role = 'admin'"
                        ).fetchone()[0]
                        <= 1
                    ):
                        db.rollback()
                        flash(
                            "No se puede cambiar el rol del ultimo administrador.",
                            "error",
                        )
                    else:
                        if password:
                            db.execute(
                                """UPDATE users
                                   SET username = ?, password_hash = ?, role = ?
                                   WHERE id = ?""",
                                (
                                    username,
                                    generate_password_hash(password),
                                    role,
                                    user_id,
                                ),
                            )
                        else:
                            db.execute(
                                "UPDATE users SET username = ?, role = ? WHERE id = ?",
                                (username, role, user_id),
                            )

                        db.execute(
                            "DELETE FROM teacher_assignments WHERE teacher_id = ?",
                            (user_id,),
                        )
                        if role == "teacher":
                            db.executemany(
                                """INSERT INTO teacher_assignments
                                   (teacher_id, subject_id, classroom_id)
                                   VALUES (?, ?, ?)""",
                                [
                                    (user_id, subject_id, classroom_id)
                                    for subject_id, classroom_id in assignments
                                ],
                            )
                        db.commit()
                        flash("Cuenta actualizada.", "success")
                        return redirect(url_for("admin"))
                except DATABASE_INTEGRITY_ERRORS:
                    db.rollback()
                    flash("Ese nombre de usuario ya esta en uso.", "error")

        return render_template(
            "edit_user.html",
            account=account_values,
            assignment_options=assignment_options,
            current_assignments=current_assignments,
            current_user_id=g.user["id"],
        )

    @app.route("/admin/users/<int:user_id>/assignments", methods=("GET", "POST"))
    @role_required("admin")
    def edit_teacher_assignments(user_id):
        db = get_db()
        teacher = db.execute(
            "SELECT id, username FROM users WHERE id = ? AND role = 'teacher'",
            (user_id,),
        ).fetchone()
        if teacher is None:
            abort(404)
        classrooms = db.execute(
            "SELECT id, grade, name FROM classrooms ORDER BY grade, name"
        ).fetchall()
        subjects = db.execute("SELECT id, name FROM subjects ORDER BY name").fetchall()
        assignment_options = [
            {
                "subject_id": subject["id"],
                "subject": subject["name"],
                "classroom_id": classroom["id"],
                "grade": classroom["grade"],
                "classroom": classroom["name"],
            }
            for subject in subjects
            for classroom in classrooms
        ]
        current_assignments = {
            f"{row['subject_id']}:{row['classroom_id']}"
            for row in db.execute(
                """SELECT subject_id, classroom_id FROM teacher_assignments
                   WHERE teacher_id = ?""",
                (user_id,),
            ).fetchall()
        }

        if request.method == "POST":
            assignments, valid_assignments = parse_assignment_values(
                request.form.getlist("assignments")
            )
            if not valid_assignments:
                flash("Una o mas asignaciones no son validas.", "error")
            elif not assignment_values_exist(db, assignments):
                flash("Una o mas asignaciones ya no existen. Actualiza la pagina y vuelve a intentarlo.", "error")
            else:
                try:
                    db.execute("BEGIN IMMEDIATE")
                    db.execute(
                        "DELETE FROM teacher_assignments WHERE teacher_id = ?",
                        (user_id,),
                    )
                    db.executemany(
                        """INSERT INTO teacher_assignments
                           (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                        [
                            (user_id, subject_id, classroom_id)
                            for subject_id, classroom_id in assignments
                        ],
                    )
                    db.commit()
                except DATABASE_INTEGRITY_ERRORS:
                    db.rollback()
                    flash("No se pudieron guardar las asignaciones. Vuelve a intentarlo.", "error")
                else:
                    flash(f"Asignaciones de {teacher['username']} actualizadas.", "success")
                    return redirect(url_for("admin"))
        return render_template(
            "teacher_assignments.html",
            teacher=teacher,
            assignment_options=assignment_options,
            current_assignments=current_assignments,
        )

    @app.post("/admin/users/<int:user_id>/delete")
    @role_required("admin")
    def delete_user(user_id):
        db = get_db()
        db.execute("BEGIN IMMEDIATE")
        user = db.execute(
            "SELECT id, username, role FROM users WHERE id = ?", (user_id,)
        ).fetchone()
        if user is None:
            db.rollback()
            abort(404)
        if user["id"] == g.user["id"]:
            db.rollback()
            flash("No puedes eliminar la cuenta con la que iniciaste sesion.", "error")
        elif user["role"] == "admin" and db.execute(
            "SELECT COUNT(*) FROM users WHERE role = 'admin'"
        ).fetchone()[0] <= 1:
            db.rollback()
            flash("No se puede eliminar el ultimo administrador.", "error")
        else:
            try:
                db.execute(
                    """UPDATE grades SET teacher_name = ?, teacher_id = NULL
                       WHERE teacher_id = ?""",
                    (user["username"], user_id),
                )
                db.execute("DELETE FROM users WHERE id = ?", (user_id,))
                db.commit()
            except DATABASE_INTEGRITY_ERRORS:
                db.rollback()
                flash("No se pudo eliminar la cuenta porque tiene datos relacionados.", "error")
            else:
                flash(
                    f"Se elimino la cuenta de {user['username']}; se conservaron sus notas e historial.",
                    "success",
                )
        return redirect(url_for("admin"))

    @app.post("/admin/classrooms")
    @role_required("admin")
    def create_classroom():
        grade = request.form.get("grade", "").strip()
        name = request.form.get("name", "").strip()
        if not grade or not name:
            flash("Indica el grado y el nombre del aula.", "error")
        else:
            try:
                get_db().execute(
                    "INSERT INTO classrooms (grade, name) VALUES (?, ?)", (grade, name)
                )
                get_db().commit()
            except DATABASE_INTEGRITY_ERRORS:
                flash("Ese grado y aula ya existen.", "error")
            else:
                flash("Grado y aula creados.", "success")
        return redirect(url_for("admin"))

    @app.post("/admin/classrooms/<int:classroom_id>/delete")
    @role_required("admin")
    def delete_classroom(classroom_id):
        db = get_db()
        db.execute("BEGIN IMMEDIATE")
        classroom = db.execute(
            "SELECT grade, name FROM classrooms WHERE id = ?", (classroom_id,)
        ).fetchone()
        if classroom is None:
            db.rollback()
            abort(404)
        try:
            db.execute(
                """DELETE FROM grades WHERE student_id IN
                   (SELECT id FROM students WHERE classroom_id = ?)""",
                (classroom_id,),
            )
            db.execute(
                """DELETE FROM student_enrollments WHERE student_id IN
                   (SELECT id FROM students WHERE classroom_id = ?)""",
                (classroom_id,),
            )
            db.execute("DELETE FROM students WHERE classroom_id = ?", (classroom_id,))
            db.execute("DELETE FROM classrooms WHERE id = ?", (classroom_id,))
            db.commit()
        except DATABASE_INTEGRITY_ERRORS:
            db.rollback()
            flash("No se pudo eliminar el grado y aula porque tiene datos relacionados.", "error")
        else:
            flash(
                f"Se elimino {classroom['grade']} · {classroom['name']} junto con sus alumnos y notas.",
                "success",
            )
        return redirect(url_for("admin"))

    @app.route("/admin/students/<int:student_id>/edit", methods=("GET", "POST"))
    @role_required("admin")
    def edit_student(student_id):
        db = get_db()
        student = db.execute(
            """SELECT id, name, modality, sex, person_code
               FROM students WHERE id = ?""",
            (student_id,),
        ).fetchone()
        if student is None:
            abort(404)

        student_values = dict(student)
        if request.method == "POST":
            student_values.update(
                name=request.form.get("name", "").strip(),
                modality=request.form.get("modality", "").strip(),
                sex=request.form.get("sex", "").strip(),
                person_code=request.form.get("person_code", "").strip(),
            )
            name = student_values["name"]
            modality = student_values["modality"]
            sex = student_values["sex"]
            person_code = student_values["person_code"]

            if not name or not modality or not sex or not person_code:
                flash("Completa todos los datos del estudiante.", "error")
            elif len(name) > 255:
                flash("El nombre debe tener como máximo 255 caracteres.", "error")
            elif sex not in {"F", "M"}:
                flash("Selecciona un sexo valido: F o M.", "error")
            elif len(person_code) > 100:
                flash("El código único debe tener como máximo 100 caracteres.", "error")
            elif not db.execute(
                "SELECT 1 FROM modalities WHERE name = ?", (modality,)
            ).fetchone():
                flash("Selecciona una modalidad valida.", "error")
            else:
                try:
                    db.execute(
                        """UPDATE students
                           SET name = ?, modality = ?, sex = ?, person_code = ?
                           WHERE id = ?""",
                        (
                            name,
                            modality,
                            sex,
                            person_code,
                            student_id,
                        ),
                    )
                    db.commit()
                except DATABASE_INTEGRITY_ERRORS:
                    db.rollback()
                    flash("No se pudieron actualizar los datos.", "error")
                else:
                    flash("Datos del estudiante actualizados.", "success")
                    return redirect(url_for("admin_students"))

        modalities = db.execute(
            "SELECT name FROM modalities ORDER BY name"
        ).fetchall()
        return render_template(
            "edit_student.html",
            student=student_values,
            modalities=modalities,
        )

    @app.post("/admin/students/bulk")
    @role_required("admin")
    def create_students_bulk():
        db = get_db()
        classroom_id = request.form.get("classroom_id", type=int)
        modality = request.form.get("modality", "").strip()
        uploaded_file = request.files.get("student_file")
        grouped_import = False
        if uploaded_file is not None:
            try:
                students, grouped_import = parse_student_workbook(uploaded_file)
            except StudentWorkbookError as exc:
                flash(str(exc), "error")
                return redirect(url_for("admin_students"))
        else:
            names = request.form.getlist("student_name")
            sexes = request.form.getlist("student_sex")
            person_codes = request.form.getlist("student_person_code")
            if len(names) != len(sexes) or len(names) != len(person_codes):
                flash("La tabla de estudiantes no es válida.", "error")
                return redirect(url_for("admin_students"))
            students = [
                {
                    "row_number": row_number,
                    "name": name,
                    "sex": sex,
                    "person_code": code,
                }
                for row_number, (name, sex, code) in enumerate(
                    zip(names, sexes, person_codes), start=1
                )
            ]

        if not students:
            flash("Ingresa al menos un estudiante.", "error")
            return redirect(url_for("admin_students"))
        if not grouped_import:
            if not classroom_id or (uploaded_file is None and not modality):
                flash("Selecciona el grado, aula y modalidad.", "error")
                return redirect(url_for("admin_students"))
            if not db.execute(
                "SELECT 1 FROM classrooms WHERE id = ?", (classroom_id,)
            ).fetchone():
                flash("El grado y aula seleccionados no existen.", "error")
                return redirect(url_for("admin_students"))
            if uploaded_file is None and not db.execute(
                "SELECT 1 FROM modalities WHERE name = ?", (modality,)
            ).fetchone():
                flash("Selecciona una modalidad valida.", "error")
                return redirect(url_for("admin_students"))

        parsed_students = []
        for student in students:
            row_number = student["row_number"]
            name = student["name"].strip()
            sex = student["sex"].strip().upper()
            person_code = student["person_code"].strip() or None
            if not name and not sex and not person_code:
                continue
            if not name or not sex or (uploaded_file is None and not person_code):
                flash(
                    f"Completa nombre y sexo"
                    f"{' y código de persona' if uploaded_file is None else ''}"
                    f" en la fila {row_number}.",
                    "error",
                )
                return redirect(url_for("admin_students"))
            if len(name) > 255:
                flash(f"El nombre de la fila {row_number} supera los 255 caracteres.", "error")
                return redirect(url_for("admin_students"))
            if sex not in {"F", "M"}:
                flash(f"El sexo de la fila {row_number} debe ser F o M.", "error")
                return redirect(url_for("admin_students"))
            if person_code and len(person_code) > 100:
                flash(
                    f"El código de la fila {row_number} supera los 100 caracteres.",
                    "error",
                )
                return redirect(url_for("admin_students"))
            row_modality = student.get("modality", "").strip() or modality
            if grouped_import:
                grade = student.get("grade", "").strip()
                classroom = student.get("classroom", "").strip()
                if not grade or not classroom or not row_modality:
                    flash(
                        f"Completa la modalidad, nivel y sección en la fila {row_number}.",
                        "error",
                    )
                    return redirect(url_for("admin_students"))
                if len(grade) > 255 or len(classroom) > 255 or len(row_modality) > 100:
                    flash(
                        f"Los datos del grupo en la fila {row_number} superan el límite permitido.",
                        "error",
                    )
                    return redirect(url_for("admin_students"))
                parsed_students.append(
                    (row_number, name, sex, person_code, grade, classroom, row_modality)
                )
            else:
                if not row_modality:
                    flash(
                        f"Completa la modalidad en la fila {row_number}.",
                        "error",
                    )
                    return redirect(url_for("admin_students"))
                parsed_students.append(
                    (row_number, name, sex, person_code, classroom_id, row_modality)
                )

        if not parsed_students:
            flash("Ingresa al menos un estudiante completo.", "error")
            return redirect(url_for("admin_students"))

        try:
            student_rows = []
            classroom_cache = {}
            modality_cache = set()
            for student in parsed_students:
                if grouped_import:
                    row_number, name, sex, person_code, grade, classroom, row_modality = student
                    if row_modality not in modality_cache:
                        if not db.execute(
                            "SELECT 1 FROM modalities WHERE name = ?", (row_modality,)
                        ).fetchone():
                            db.execute(
                                "INSERT INTO modalities (name) VALUES (?)",
                                (row_modality,),
                            )
                        modality_cache.add(row_modality)
                    classroom_key = (grade, classroom)
                    if classroom_key not in classroom_cache:
                        existing_classroom = db.execute(
                            "SELECT id FROM classrooms WHERE grade = ? AND name = ?",
                            classroom_key,
                        ).fetchone()
                        if existing_classroom:
                            classroom_cache[classroom_key] = existing_classroom["id"]
                        else:
                            db.execute(
                                "INSERT INTO classrooms (grade, name) VALUES (?, ?)",
                                classroom_key,
                            )
                            classroom_cache[classroom_key] = db.execute(
                                "SELECT id FROM classrooms WHERE grade = ? AND name = ?",
                                classroom_key,
                            ).fetchone()["id"]
                    student_rows.append(
                        (
                            classroom_cache[classroom_key],
                            name,
                            row_modality,
                            sex,
                            person_code,
                        )
                    )
                else:
                    _, name, sex, person_code, classroom_id, row_modality = student
                    if row_modality not in modality_cache:
                        if not db.execute(
                            "SELECT 1 FROM modalities WHERE name = ?", (row_modality,)
                        ).fetchone():
                            db.execute(
                                "INSERT INTO modalities (name) VALUES (?)",
                                (row_modality,),
                            )
                        modality_cache.add(row_modality)
                    student_rows.append(
                        (classroom_id, name, row_modality, sex, person_code)
                    )
            db.executemany(
                """INSERT INTO students
                   (classroom_id, name, modality, sex, person_code)
                   VALUES (?, ?, ?, ?, ?)""",
                student_rows,
            )
            ensure_current_enrollments(db)
            db.commit()
        except DATABASE_INTEGRITY_ERRORS:
            db.rollback()
            flash("No se pudieron importar los estudiantes y sus grupos.", "error")
        else:
            flash(f"Se agregaron {len(student_rows)} estudiantes.", "success")
        return redirect(url_for("admin_students"))

    @app.post("/admin/modalities")
    @role_required("admin")
    def create_modality():
        name = request.form.get("name", "").strip()
        if not name or len(name) > 100:
            flash("Indica una modalidad de hasta 100 caracteres.", "error")
        else:
            try:
                get_db().execute(
                    "INSERT INTO modalities (name) VALUES (?)", (name,)
                )
                get_db().commit()
            except DATABASE_INTEGRITY_ERRORS:
                flash("Esa modalidad ya existe.", "error")
            else:
                flash("Modalidad agregada.", "success")
        return redirect(url_for("admin_students"))

    @app.post("/admin/students/<int:student_id>/delete")
    @role_required("admin")
    def delete_student(student_id):
        db = get_db()
        db.execute("BEGIN IMMEDIATE")
        student = db.execute(
            "SELECT name FROM students WHERE id = ?", (student_id,)
        ).fetchone()
        if student is None:
            db.rollback()
            abort(404)
        try:
            db.execute("DELETE FROM grades WHERE student_id = ?", (student_id,))
            db.execute(
                "DELETE FROM student_enrollments WHERE student_id = ?",
                (student_id,),
            )
            db.execute("DELETE FROM students WHERE id = ?", (student_id,))
            db.commit()
        except DATABASE_INTEGRITY_ERRORS:
            db.rollback()
            flash("No se pudo eliminar el alumno porque tiene datos relacionados.", "error")
        else:
            flash(f"Se elimino a {student['name']} junto con sus notas.", "success")
        return redirect(url_for("admin_students"))

    @app.post("/admin/subjects")
    @role_required("admin")
    def create_subject():
        name = request.form.get("name", "").strip()
        if not name:
            flash("Indica el nombre de la materia.", "error")
        else:
            try:
                get_db().execute("INSERT INTO subjects (name) VALUES (?)", (name,))
                get_db().commit()
            except DATABASE_INTEGRITY_ERRORS:
                flash("Esa materia ya existe.", "error")
            else:
                flash("Materia creada.", "success")
        return redirect(url_for("admin"))

    @app.post("/admin/subjects/<int:subject_id>/delete")
    @role_required("admin")
    def delete_subject(subject_id):
        db = get_db()
        db.execute("BEGIN IMMEDIATE")
        subject = db.execute(
            "SELECT name FROM subjects WHERE id = ?", (subject_id,)
        ).fetchone()
        if subject is None:
            db.rollback()
            abort(404)
        try:
            db.execute("DELETE FROM grades WHERE subject_id = ?", (subject_id,))
            db.execute("DELETE FROM subjects WHERE id = ?", (subject_id,))
            db.commit()
        except DATABASE_INTEGRITY_ERRORS:
            db.rollback()
            flash("No se pudo eliminar la materia porque tiene datos relacionados.", "error")
        else:
            flash(f"Se elimino la materia {subject['name']} junto con sus notas.", "success")
        return redirect(url_for("admin"))

    @app.route("/admin/grades/manage", methods=("GET", "POST"))
    @role_required("admin")
    def manage_grades():
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        current_school_year = str(date.today().year)
        classrooms = db.execute(
            "SELECT id, grade, name FROM classrooms ORDER BY grade, name"
        ).fetchall()
        school_years = db.execute(
            """SELECT DISTINCT school_year FROM student_enrollments
               ORDER BY school_year DESC"""
        ).fetchall()
        subjects = db.execute("SELECT id, name FROM subjects ORDER BY name").fetchall()
        classroom_id = request.values.get("classroom_id", type=int)
        subject_id = request.values.get("subject_id", type=int)
        term = request.values.get("term", "").strip()
        school_year = request.values.get(
            "school_year", current_school_year
        ).strip()
        students = []
        existing_grades = {}

        if classroom_id:
            if not db.execute(
                "SELECT 1 FROM classrooms WHERE id = ?", (classroom_id,)
            ).fetchone():
                abort(404)
            students = db.execute(
                """SELECT s.id, s.name
                   FROM student_enrollments e
                   JOIN students s ON s.id = e.student_id
                   WHERE e.classroom_id = ? AND e.school_year = ?
                   ORDER BY s.name""",
                (classroom_id, school_year),
            ).fetchall()
        if subject_id and not db.execute(
            "SELECT 1 FROM subjects WHERE id = ?", (subject_id,)
        ).fetchone():
            abort(404)
        if classroom_id and subject_id and term and school_year:
            existing_grades = {
                row["student_id"]: row
                for row in db.execute(
                    """SELECT id, student_id, score FROM grades
                       WHERE subject_id = ? AND term = ? AND school_year = ?
                         AND student_id IN (
                             SELECT student_id FROM student_enrollments
                             WHERE classroom_id = ? AND school_year = ?
                         )""",
                    (subject_id, term, school_year, classroom_id, school_year),
                ).fetchall()
            }

        if request.method == "POST":
            if (
                not classroom_id
                or not subject_id
                or term not in EVALUATION_PERIODS
                or len(school_year) != 4
                or not school_year.isdigit()
                or not 1900 <= int(school_year) <= 2100
            ):
                flash("Selecciona el grado, aula, materia y corte evaluativo.", "error")
            elif not students:
                flash("Este grado y aula aun no tienen alumnos.", "error")
            else:
                entries = []
                invalid_score = False
                for student in students:
                    raw_score = request.form.get(f"score_{student['id']}", "").strip()
                    if not raw_score:
                        continue
                    try:
                        score = parse_grade_score(raw_score)
                    except ValueError:
                        invalid_score = True
                        break
                    entries.append((student["id"], score))

                if invalid_score:
                    flash(
                        "Cada nota debe ser un numero entre 0 y 100 o texto de hasta 3 caracteres.",
                        "error",
                    )
                elif not entries:
                    flash("No hay notas para guardar o actualizar.", "error")
                else:
                    try:
                        db.execute("BEGIN")
                        for student_id, score in entries:
                            existing = existing_grades.get(student_id)
                            if existing:
                                if normalize_grade_score(existing["score"]) != score:
                                    db.execute(
                                        """UPDATE grades
                                           SET score = ?, updated_at = CURRENT_TIMESTAMP
                                           WHERE id = ?""",
                                        (score_for_database(score), existing["id"]),
                                    )
                            else:
                                db.execute(
                                    """INSERT INTO grades
                                       (student_id, subject_id, school_year,
                                        teacher_id, teacher_name, score, term)
                                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                                    (
                                        student_id,
                                        subject_id,
                                        school_year,
                                        g.user["id"],
                                        g.user["username"],
                                        score_for_database(score),
                                        term,
                                    ),
                                )
                        db.commit()
                    except DATABASE_INTEGRITY_ERRORS:
                        db.rollback()
                        flash("No se pudieron guardar las notas. Revisa si ya existe una nota para ese corte evaluativo.", "error")
                    else:
                        count = len(entries)
                        count_label = "calificacion" if count == 1 else "calificaciones"
                        flash(
                            f"Se guardaron o actualizaron {count} {count_label}.",
                            "success",
                        )
                        return redirect(
                            url_for(
                                "manage_grades",
                                classroom_id=classroom_id,
                                subject_id=subject_id,
                                term=term,
                                school_year=school_year,
                            )
                        )

        return render_template(
            "admin_grades.html",
            classrooms=classrooms,
            school_years=school_years,
            subjects=subjects,
            students=students,
            classroom_id=classroom_id,
            subject_id=subject_id,
            term=term,
            school_year=school_year,
            periods=EVALUATION_PERIODS,
            existing_grades=existing_grades,
        )

    @app.get("/admin/grades/<int:grade_id>/edit")
    @role_required("admin")
    def edit_grade(grade_id):
        db = get_db()
        ensure_current_enrollments(db)
        db.commit()
        grade = db.execute(
            """SELECT e.classroom_id, g.subject_id, g.term, g.school_year
               FROM grades g
               JOIN student_enrollments e
                 ON e.student_id = g.student_id AND e.school_year = g.school_year
               WHERE g.id = ?""",
            (grade_id,),
        ).fetchone()
        if grade is None:
            abort(404)
        return redirect(
            url_for(
                "manage_grades",
                classroom_id=grade["classroom_id"],
                subject_id=grade["subject_id"],
                term=grade["term"],
                school_year=grade["school_year"],
            )
        )

    @app.post("/admin/grades/<int:grade_id>/delete")
    @role_required("admin")
    def delete_grade(grade_id):
        cursor = get_db().execute("DELETE FROM grades WHERE id = ?", (grade_id,))
        get_db().commit()
        if cursor.rowcount == 0:
            abort(404)
        flash("Nota eliminada por el administrador.", "success")
        return redirect(url_for("grade_list"))

    return app


def create_grade_export(rows, filename, export_subjects):
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter

    workbook = Workbook()
    include_teacher = request.path.startswith("/grades/export/") and request.path != "/grades/export.xlsx"
    subjects = sorted(
        set(export_subjects) | {row["subject"] for row in rows},
        key=str.casefold,
    )
    headers = [
        "N.º",
        "Modalidad",
        "Nivel",
        "Sección",
        "Nombre y apellido",
        "Sexo",
        "Código de persona",
        *subjects,
    ]
    if include_teacher and len(subjects) == 1:
        headers.append("Maestro")
    headers.append("Año escolar")
    header_fill = PatternFill("solid", fgColor="00B7F0")
    header_font = Font(bold=True, color="000000")
    failing_grade_font = Font(color="FFFF0000")
    header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    institution_font = Font(bold=True, size=16)
    period_font = Font(bold=True, size=14)
    thin_border = Border(
        left=Side(style="thin", color="000000"),
        right=Side(style="thin", color="000000"),
        top=Side(style="thin", color="000000"),
        bottom=Side(style="thin", color="000000"),
    )

    grades_by_student_period = {}
    for row in rows:
        key = (
            row["classroom_id"],
            row["student_id"],
            row["term"],
            row["school_year"],
        )
        values = grades_by_student_period.setdefault(
            key,
            {
                "student_id": row["student_id"],
                "classroom_id": row["classroom_id"],
                "student_modality": row["student_modality"],
                "student_name": row["student_name"],
                "student_sex": row["student_sex"],
                "student_person_code": row["student_person_code"],
                "grade": row["grade"],
                "classroom": row["classroom"],
                "term": row["term"],
                "school_year": row["school_year"],
                "scores": {},
                "teacher": row["teacher"],
            },
        )
        values["scores"][row["subject"]] = normalize_grade_score(row["score"])
        if include_teacher and len(subjects) == 1:
            values["teacher"] = row["teacher"]

    periods = sorted(
        {grade["term"] for grade in grades_by_student_period.values()},
        key=lambda period: (
            EVALUATION_PERIODS.index(period)
            if period in EVALUATION_PERIODS
            else len(EVALUATION_PERIODS),
            period.casefold(),
        ),
    ) or [""]
    used_sheet_titles = set()

    for period_index, period in enumerate(periods):
        sheet = workbook.active if period_index == 0 else workbook.create_sheet()
        title_base = "".join(
            "-" if char in "[]:*?/\\" else char for char in (period or "Notas")
        )[:31] or "Notas"
        sheet_title = title_base
        suffix_number = 2
        while sheet_title in used_sheet_titles:
            suffix = f" ({suffix_number})"
            sheet_title = f"{title_base[:31 - len(suffix)]}{suffix}"
            suffix_number += 1
        used_sheet_titles.add(sheet_title)
        sheet.title = sheet_title
        sheet.sheet_view.showGridLines = False

        institution_cell = sheet.cell(
            row=1,
            column=1,
            value=INSTITUTION_NAME,
        )
        institution_cell.font = institution_font
        institution_cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(headers))
        period_cell = sheet.cell(
            row=2,
            column=1,
            value=f"Corte evaluativo: {period}" if period else "Sin notas registradas",
        )
        period_cell.font = period_font
        period_cell.alignment = Alignment(horizontal="center", vertical="center")
        sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=len(headers))
        sheet.row_dimensions[1].height = 28
        sheet.row_dimensions[2].height = 24

        for column_index, header in enumerate(headers, start=1):
            cell = sheet.cell(row=3, column=column_index, value=header)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        period_grades = [
            grade
            for grade in grades_by_student_period.values()
            if grade["term"] == period
        ]
        students_by_classroom = {}
        for grade in period_grades:
            classroom_year = (grade["classroom_id"], grade["school_year"])
            students_by_classroom.setdefault(classroom_year, {})[
                grade["student_id"]
            ] = grade
        student_numbers = {}
        for (classroom_id, school_year), students in students_by_classroom.items():
            ordered_students = sorted(
                students.values(),
                key=lambda student: (
                    {"F": 0, "M": 1}.get(student["student_sex"], 2),
                    student["student_name"].casefold(),
                ),
            )
            for number, student in enumerate(ordered_students, start=1):
                student_numbers[
                    (classroom_id, school_year, student["student_id"])
                ] = number

        sorted_grades = sorted(
            period_grades,
            key=lambda grade: (
                grade["school_year"],
                grade["grade"].casefold(),
                grade["classroom"].casefold(),
                {"F": 0, "M": 1}.get(grade["student_sex"], 2),
                grade["student_name"].casefold(),
            ),
        )
        for grade in sorted_grades:
            values = [
                student_numbers[
                    (
                        grade["classroom_id"],
                        grade["school_year"],
                        grade["student_id"],
                    )
                ],
                grade["student_modality"],
                grade["grade"],
                grade["classroom"],
                grade["student_name"],
                grade["student_sex"],
                grade["student_person_code"],
                *(
                    format_export_grade_score(grade["scores"].get(subject))
                    for subject in subjects
                ),
            ]
            if include_teacher and len(subjects) == 1:
                values.append(grade["teacher"])
            values.append(grade["school_year"])
            sheet.append(values)
            for column_index, value in enumerate(values, start=1):
                cell = sheet.cell(row=sheet.max_row, column=column_index)
                cell.border = thin_border
                cell.alignment = Alignment(
                    horizontal="center"
                    if column_index in {1, 3, 4, 6} or column_index >= 8
                    else "left"
                )
                if 8 <= column_index < 8 + len(subjects):
                    score = normalize_grade_score(
                        grade["scores"].get(subjects[column_index - 8])
                    )
                    if isinstance(score, (int, float)) and math.isfinite(score) and score < 60:
                        cell.font = failing_grade_font
                if isinstance(value, str):
                    cell.data_type = "s"

        roster_last_row = sheet.max_row
        summary_title_row = roster_last_row + 2
        summary_header_row = summary_title_row + 1
        summary_start_column = len(headers) + 2
        summary_title = sheet.cell(
            row=summary_title_row,
            column=summary_start_column,
            value="Resumen de aplazados (<60 o AI)",
        )
        summary_title.font = period_font
        sheet.merge_cells(
            start_row=summary_title_row,
            start_column=summary_start_column,
            end_row=summary_title_row,
            end_column=summary_start_column + 5,
        )
        summary_headers = (
            "Año escolar",
            "Nivel",
            "Sección",
            "Materia",
            "Promedio (<60)",
            "Cantidad (<60 o AI)",
        )
        for column_offset, header in enumerate(summary_headers):
            cell = sheet.cell(
                row=summary_header_row,
                column=summary_start_column + column_offset,
                value=header,
            )
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = header_alignment
            cell.border = thin_border

        failing_scores = {}
        failing_counts = {}
        for grade in period_grades:
            for subject, value in grade["scores"].items():
                key = (
                    grade["school_year"],
                    grade["grade"],
                    grade["classroom"],
                    subject,
                )
                failing_scores.setdefault(key, [])
                failing_counts.setdefault(key, 0)
                score = normalize_grade_score(value)
                if (
                    isinstance(score, (int, float))
                    and math.isfinite(score)
                    and score < 60
                ):
                    failing_scores[key].append(score)
                    failing_counts[key] += 1
                elif isinstance(score, str) and score.strip().upper() == "AI":
                    failing_counts[key] += 1

        summary_keys = sorted(
            failing_scores,
            key=lambda key: tuple(value.casefold() for value in key),
        )
        for row_offset, (school_year, grade_name, classroom_name, subject) in enumerate(
            summary_keys, start=1
        ):
            key = (school_year, grade_name, classroom_name, subject)
            scores = failing_scores[key]
            summary_values = (
                school_year,
                grade_name,
                classroom_name,
                subject,
                sum(scores) / len(scores) if scores else None,
                failing_counts[key],
            )
            for column_offset, value in enumerate(summary_values):
                cell = sheet.cell(
                    row=summary_header_row + row_offset,
                    column=summary_start_column + column_offset,
                    value=value,
                )
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="left" if column_offset in {0, 1, 3} else "center")
                if column_offset == 4 and value is not None:
                    cell.number_format = "0.00"

        sheet.freeze_panes = "A4"
        sheet.auto_filter.ref = (
            f"A3:{get_column_letter(len(headers))}{roster_last_row}"
        )
        for column_index in range(1, sheet.max_column + 1):
            cells = [
                sheet.cell(row=row_index, column=column_index)
                for row_index in range(3, sheet.max_row + 1)
                if sheet.cell(row=row_index, column=column_index).value is not None
            ]
            if not cells:
                continue
            width = min(max(len(str(cell.value)) for cell in cells) + 2, 24)
            if column_index == 1:
                width = 6
            elif 8 <= column_index < 8 + len(subjects):
                width = max(width, 10)
            sheet.column_dimensions[get_column_letter(column_index)].width = width

    output = io.BytesIO()
    workbook.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def get_db():
    if "db" not in g:
        database_url = current_app.config["DATABASE"]
        db_type = current_app.config["DB_TYPE"]
        if db_type == "mysql":
            db = MySQLConnectionWrapper(database_url)
        elif db_type == "postgresql":
            db = PostgreSQLConnectionWrapper(database_url)
        else:
            db = sqlite3.connect(database_url)
            db.row_factory = sqlite3.Row
            db.execute("PRAGMA foreign_keys = ON")
        g.db = db
    return g.db


def _postgres_table_columns(db, table_name):
    return {
        row["column_name"]: row
        for row in db.execute(
            """SELECT column_name, data_type, is_nullable
               FROM information_schema.columns
               WHERE table_schema = current_schema() AND table_name = ?""",
            (table_name,),
        ).fetchall()
    }


def _postgres_unique_constraints(db, table_name):
    return db.execute(
        """SELECT tc.constraint_name,
                  array_agg(kcu.column_name ORDER BY kcu.ordinal_position) AS columns
           FROM information_schema.table_constraints AS tc
           JOIN information_schema.key_column_usage AS kcu
             ON kcu.constraint_catalog = tc.constraint_catalog
            AND kcu.constraint_schema = tc.constraint_schema
            AND kcu.constraint_name = tc.constraint_name
            AND kcu.table_name = tc.table_name
           WHERE tc.table_schema = current_schema()
             AND tc.table_name = ?
             AND tc.constraint_type = 'UNIQUE'
           GROUP BY tc.constraint_name""",
        (table_name,),
    ).fetchall()


def migrate_student_modality(db):
    if getattr(db, "db_type", "sqlite") == "mysql":
        columns = {
            row["Field"] for row in db.execute("SHOW COLUMNS FROM students").fetchall()
        }
        if "modality" not in columns:
            db.execute(
                "ALTER TABLE students ADD COLUMN modality VARCHAR(100) NOT NULL DEFAULT 'Regular'"
            )
        return

    if getattr(db, "db_type", "sqlite") == "postgresql":
        if "modality" not in _postgres_table_columns(db, "students"):
            db.execute(
                "ALTER TABLE students ADD COLUMN modality TEXT NOT NULL DEFAULT 'Regular'"
            )
        return

    columns = {
        row["name"] for row in db.execute("PRAGMA table_info(students)").fetchall()
    }
    if "modality" not in columns:
        db.execute(
            "ALTER TABLE students ADD COLUMN modality TEXT NOT NULL DEFAULT 'Regular'"
        )


def migrate_modalities(db):
    modality_names = ["Regular"] + [
        row["modality"]
        for row in db.execute("SELECT DISTINCT modality FROM students").fetchall()
    ]
    for name in modality_names:
        if not name or not name.strip():
            continue
        if not db.execute(
            "SELECT 1 FROM modalities WHERE name = ?", (name,)
        ).fetchone():
            db.execute("INSERT INTO modalities (name) VALUES (?)", (name,))


def migrate_student_identity(db):
    if getattr(db, "db_type", "sqlite") == "mysql":
        columns = {
            row["Field"] for row in db.execute("SHOW COLUMNS FROM students").fetchall()
        }
        if "sex" not in columns:
            db.execute("ALTER TABLE students ADD COLUMN sex VARCHAR(1) NULL")
        if "person_code" not in columns:
            db.execute("ALTER TABLE students ADD COLUMN person_code VARCHAR(100) NULL")
        indexes = {
            row["Key_name"] for row in db.execute("SHOW INDEX FROM students").fetchall()
        }
        if "unique_student_person_code" in indexes:
            db.execute("DROP INDEX unique_student_person_code ON students")
        return

    if getattr(db, "db_type", "sqlite") == "postgresql":
        columns = _postgres_table_columns(db, "students")
        if "sex" not in columns:
            db.execute("ALTER TABLE students ADD COLUMN sex TEXT")
        if "person_code" not in columns:
            db.execute("ALTER TABLE students ADD COLUMN person_code TEXT")
        for constraint in _postgres_unique_constraints(db, "students"):
            if constraint["columns"] == ["person_code"]:
                escaped_name = constraint["constraint_name"].replace('"', '""')
                db.execute(f'ALTER TABLE students DROP CONSTRAINT "{escaped_name}"')
        db.execute("DROP INDEX IF EXISTS unique_student_person_code")
        return

    columns = {
        row["name"] for row in db.execute("PRAGMA table_info(students)").fetchall()
    }
    if "sex" not in columns:
        db.execute("ALTER TABLE students ADD COLUMN sex TEXT")
    if "person_code" not in columns:
        db.execute("ALTER TABLE students ADD COLUMN person_code TEXT")
    db.execute("DROP INDEX IF EXISTS unique_student_person_code")


def ensure_current_enrollments(db, school_year=None):
    school_year = str(school_year or date.today().year)
    db.execute(
        """UPDATE students
           SET classroom_id = (
               SELECT e.classroom_id FROM student_enrollments e
               WHERE e.student_id = students.id AND e.school_year = ?
           )
           WHERE id IN (
               SELECT student_id FROM student_enrollments WHERE school_year = ?
           )""",
        (school_year, school_year),
    )

    students_without_enrollment = db.execute(
        """SELECT s.id, s.classroom_id
           FROM students s
           LEFT JOIN student_enrollments e
             ON e.student_id = s.id AND e.school_year = ?
           WHERE e.id IS NULL""",
        (school_year,),
    ).fetchall()
    db.executemany(
        """INSERT INTO student_enrollments (student_id, classroom_id, school_year)
           VALUES (?, ?, ?)""",
        [
            (student["id"], student["classroom_id"], school_year)
            for student in students_without_enrollment
        ],
    )


def migrate_grade_school_year(db):
    school_year = str(date.today().year)
    if getattr(db, "db_type", "sqlite") == "mysql":
        columns = {
            row["Field"] for row in db.execute("SHOW COLUMNS FROM grades").fetchall()
        }
        if "school_year" not in columns:
            db.execute("ALTER TABLE grades ADD COLUMN school_year VARCHAR(4) NULL")
        db.execute(
            "UPDATE grades SET school_year = ? WHERE school_year IS NULL OR school_year = ''",
            (school_year,),
        )

        indexes = {}
        for row in db.execute("SHOW INDEX FROM grades").fetchall():
            indexes.setdefault(row["Key_name"], []).append(row)
        for index_name, rows in indexes.items():
            columns_in_index = [
                row["Column_name"]
                for row in sorted(rows, key=lambda item: item["Seq_in_index"])
            ]
            if (
                index_name != "PRIMARY"
                and rows[0]["Non_unique"] == 0
                and columns_in_index == ["student_id", "subject_id", "term"]
            ):
                escaped_name = index_name.replace("`", "``")
                db.execute(f"ALTER TABLE grades DROP INDEX `{escaped_name}`")

        indexes = {}
        for row in db.execute("SHOW INDEX FROM grades").fetchall():
            indexes.setdefault(row["Key_name"], []).append(row)
        has_school_year_unique = any(
            rows[0]["Non_unique"] == 0
            and [
                row["Column_name"]
                for row in sorted(rows, key=lambda item: item["Seq_in_index"])
            ]
            == ["student_id", "subject_id", "term", "school_year"]
            for rows in indexes.values()
        )
        if not has_school_year_unique:
            db.execute(
                """ALTER TABLE grades
                   ADD UNIQUE KEY unique_grade_school_year
                   (student_id, subject_id, term, school_year)"""
            )
        db.execute("ALTER TABLE grades MODIFY school_year VARCHAR(4) NOT NULL")
        return

    if getattr(db, "db_type", "sqlite") == "postgresql":
        columns = _postgres_table_columns(db, "grades")
        if "school_year" not in columns:
            db.execute("ALTER TABLE grades ADD COLUMN school_year VARCHAR(4)")
        db.execute(
            "UPDATE grades SET school_year = ? WHERE school_year IS NULL OR school_year = ''",
            (school_year,),
        )

        for constraint in _postgres_unique_constraints(db, "grades"):
            if constraint["columns"] == ["student_id", "subject_id", "term"]:
                escaped_name = constraint["constraint_name"].replace('"', '""')
                db.execute(f'ALTER TABLE grades DROP CONSTRAINT "{escaped_name}"')

        unique_constraints = _postgres_unique_constraints(db, "grades")
        has_school_year_unique = any(
            constraint["columns"]
            == ["student_id", "subject_id", "term", "school_year"]
            for constraint in unique_constraints
        )
        if not has_school_year_unique:
            db.execute(
                """ALTER TABLE grades
                   ADD CONSTRAINT unique_grade_school_year
                   UNIQUE (student_id, subject_id, term, school_year)"""
            )
        db.execute("ALTER TABLE grades ALTER COLUMN school_year SET NOT NULL")
        return

    columns = {
        row["name"] for row in db.execute("PRAGMA table_info(grades)").fetchall()
    }
    if "school_year" in columns:
        return

    db.execute("ALTER TABLE grades RENAME TO grades_legacy")
    db.execute(
        """CREATE TABLE grades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL REFERENCES students(id),
            subject_id INTEGER NOT NULL REFERENCES subjects(id),
            school_year TEXT NOT NULL,
            teacher_id INTEGER REFERENCES users(id),
            teacher_name TEXT NOT NULL DEFAULT '',
            score REAL NOT NULL,
            term TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (student_id, subject_id, term, school_year)
        )"""
    )
    db.execute(
        """INSERT INTO grades
           (id, student_id, subject_id, school_year, teacher_id, teacher_name,
            score, term, created_at, updated_at)
           SELECT id, student_id, subject_id, ?, teacher_id, teacher_name,
                  score, term, created_at, updated_at
           FROM grades_legacy""",
        (school_year,),
    )
    db.execute("DROP TABLE grades_legacy")


def migrate_grade_teacher_snapshot(db):
    if getattr(db, "db_type", "sqlite") == "mysql":
        columns = {row["Field"]: row for row in db.execute("SHOW COLUMNS FROM grades").fetchall()}
        if "teacher_name" not in columns:
            db.execute("ALTER TABLE grades ADD COLUMN teacher_name VARCHAR(255) NOT NULL DEFAULT ''")
            db.execute(
                """UPDATE grades g
                   LEFT JOIN users u ON u.id = g.teacher_id
                   SET g.teacher_name = COALESCE(u.username, '')"""
            )
        if not columns["score"]["Type"].lower().startswith(("varchar", "char", "text")):
            db.execute("ALTER TABLE grades MODIFY score VARCHAR(30) NOT NULL")
        return

    if getattr(db, "db_type", "sqlite") == "postgresql":
        columns = _postgres_table_columns(db, "grades")
        if "teacher_id" not in columns:
            db.execute(
                "ALTER TABLE grades ADD COLUMN teacher_id BIGINT REFERENCES users(id)"
            )
            columns = _postgres_table_columns(db, "grades")
        if "teacher_name" not in columns:
            db.execute(
                "ALTER TABLE grades ADD COLUMN teacher_name TEXT NOT NULL DEFAULT ''"
            )
            db.execute(
                """UPDATE grades AS g
                   SET teacher_name = COALESCE(u.username, '')
                   FROM users AS u
                   WHERE u.id = g.teacher_id"""
            )
        if columns["teacher_id"]["is_nullable"] == "NO":
            db.execute("ALTER TABLE grades ALTER COLUMN teacher_id DROP NOT NULL")

        for constraint in db.execute(
            """SELECT tc.constraint_name, cc.check_clause
               FROM information_schema.table_constraints AS tc
               JOIN information_schema.check_constraints AS cc
                 ON cc.constraint_catalog = tc.constraint_catalog
                AND cc.constraint_schema = tc.constraint_schema
                AND cc.constraint_name = tc.constraint_name
               WHERE tc.table_schema = current_schema()
                 AND tc.table_name = 'grades'
                 AND tc.constraint_type = 'CHECK'"""
        ).fetchall():
            if "score" in constraint["check_clause"].lower():
                escaped_name = constraint["constraint_name"].replace('"', '""')
                db.execute(f'ALTER TABLE grades DROP CONSTRAINT "{escaped_name}"')

        columns = _postgres_table_columns(db, "grades")
        if columns["score"]["data_type"] != "text":
            db.execute(
                """ALTER TABLE grades ALTER COLUMN score TYPE TEXT
                   USING score::text"""
            )
        return

    columns = {
        row["name"]: row for row in db.execute("PRAGMA table_info(grades)").fetchall()
    }
    teacher_id_column = columns.get("teacher_id")
    table_sql = db.execute(
        "SELECT sql FROM sqlite_master WHERE type = 'table' AND name = 'grades'"
    ).fetchone()["sql"]
    needs_score_migration = "CHECK (SCORE" in table_sql.upper()
    if "teacher_name" in columns and (
        teacher_id_column is None or not teacher_id_column["notnull"]
    ) and not needs_score_migration:
        return

    db.execute("BEGIN")
    db.execute("ALTER TABLE grades RENAME TO grades_legacy")
    db.execute(
        """CREATE TABLE grades (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            student_id INTEGER NOT NULL REFERENCES students(id),
            subject_id INTEGER NOT NULL REFERENCES subjects(id),
            teacher_id INTEGER REFERENCES users(id),
            teacher_name TEXT NOT NULL DEFAULT '',
            score REAL NOT NULL,
            term TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (student_id, subject_id, term)
        )"""
    )
    teacher_name = (
        "COALESCE(NULLIF(old.teacher_name, ''), users.username, '')"
        if "teacher_name" in columns
        else "COALESCE(users.username, '')"
    )
    db.execute(
        """INSERT INTO grades
           (id, student_id, subject_id, teacher_id, teacher_name, score, term,
            created_at, updated_at)
           SELECT old.id, old.student_id, old.subject_id, old.teacher_id,
                  {teacher_name}, old.score, old.term,
                  old.created_at, old.updated_at
           FROM grades_legacy AS old
           LEFT JOIN users ON users.id = old.teacher_id""".format(
            teacher_name=teacher_name
        )
    )
    db.execute("DROP TABLE grades_legacy")
    db.commit()


def parse_grade_score(value):
    value = value.strip()
    try:
        score = float(value.replace(",", "."))
    except ValueError:
        if not value or len(value) > 3:
            raise ValueError("Text grades must contain at most 3 characters.")
        return value
    if not math.isfinite(score) or not 0 <= score <= 100:
        raise ValueError("Numeric grades must be between 0 and 100.")
    return score


def score_for_database(score):
    if current_app.config["DB_TYPE"] == "postgresql":
        return str(score)
    return score


def normalize_grade_score(value):
    if isinstance(value, (int, float)):
        return value
    try:
        score = float(str(value).replace(",", "."))
    except ValueError:
        return value
    if math.isfinite(score) and 0 <= score <= 100:
        return score
    return value


def format_export_grade_score(value):
    score = normalize_grade_score(value)
    if (
        not isinstance(score, (int, float))
        or not math.isfinite(score)
        or not 0 <= score <= 100
    ):
        return score
    if score < 60:
        letter_grade = "AI"
    elif score < 76:
        letter_grade = "AF"
    elif score < 90:
        letter_grade = "AS"
    else:
        letter_grade = "AA"
    return f"{format_grade_score(score):g}/{letter_grade}"


def format_grade_score(value):
    score = normalize_grade_score(value)
    return round(score, 2) if isinstance(score, (int, float)) else score


def query_grade_rows(assigned_teacher_id=None, classroom_id=None, subject_id=None):
    db = get_db()
    ensure_current_enrollments(db)
    db.commit()
    sql = """SELECT g.id, g.score, g.term, g.school_year, g.created_at, g.updated_at,
                    s.name AS student_name, s.sex AS student_sex,
                    s.person_code AS student_person_code,
                    s.modality AS student_modality,
                    c.grade, c.name AS classroom,
                    sub.name AS subject,
                    COALESCE(NULLIF(g.teacher_name, ''), u.username, 'Cuenta eliminada') AS teacher,
                    s.id AS student_id, e.classroom_id, g.subject_id
             FROM grades g
             JOIN students s ON s.id = g.student_id
             JOIN student_enrollments e
               ON e.student_id = g.student_id AND e.school_year = g.school_year
             JOIN classrooms c ON c.id = e.classroom_id
             JOIN subjects sub ON sub.id = g.subject_id
             LEFT JOIN users u ON u.id = g.teacher_id"""
    params = []
    conditions = []
    if assigned_teacher_id is not None:
        conditions.append(
            """EXISTS (
                SELECT 1 FROM teacher_assignments ta
                WHERE ta.teacher_id = ? AND ta.subject_id = g.subject_id
                  AND ta.classroom_id = e.classroom_id
            )"""
        )
        params.append(assigned_teacher_id)
    if classroom_id is not None:
        conditions.append("c.id = ?")
        params.append(classroom_id)
    if subject_id is not None:
        conditions.append("sub.id = ?")
        params.append(subject_id)
    if conditions:
        sql += " WHERE " + " AND ".join(conditions)
    sql += " ORDER BY sub.name COLLATE NOCASE, c.grade, c.name, g.school_year DESC, g.term, s.name"
    return db.execute(sql, params).fetchall()


def query_grade_summaries(teacher_id=None):
    db = get_db()
    ensure_current_enrollments(db)
    db.commit()
    if teacher_id is not None:
        sql = """SELECT c.id AS classroom_id, c.grade, c.name AS classroom,
                        sub.id AS subject_id, sub.name AS subject,
                        COUNT(DISTINCT g.id) AS grade_count,
                        COUNT(DISTINCT g.term) AS period_count
                 FROM teacher_assignments ta
                 JOIN classrooms c ON c.id = ta.classroom_id
                 JOIN subjects sub ON sub.id = ta.subject_id
                 LEFT JOIN student_enrollments e ON e.classroom_id = c.id
                 LEFT JOIN students s ON s.id = e.student_id
                 LEFT JOIN grades g ON g.student_id = s.id
                                  AND g.school_year = e.school_year
                                  AND g.subject_id = sub.id
                 WHERE ta.teacher_id = ?
                 GROUP BY sub.id, sub.name, c.id, c.grade, c.name"""
        params = (teacher_id,)
    else:
        sql = """SELECT c.id AS classroom_id, c.grade, c.name AS classroom,
                        sub.id AS subject_id, sub.name AS subject,
                        COUNT(g.id) AS grade_count,
                        COUNT(DISTINCT g.term) AS period_count
                 FROM grades g
                 JOIN students s ON s.id = g.student_id
                 JOIN student_enrollments e
                   ON e.student_id = g.student_id AND e.school_year = g.school_year
                 JOIN classrooms c ON c.id = e.classroom_id
                 JOIN subjects sub ON sub.id = g.subject_id
                 GROUP BY sub.id, sub.name, c.id, c.grade, c.name"""
        params = ()
    sql += " ORDER BY sub.name COLLATE NOCASE, c.grade COLLATE NOCASE, c.name COLLATE NOCASE"
    return get_db().execute(sql, params).fetchall()


def teacher_has_assignment(classroom_id, subject_id, teacher_id):
    return get_db().execute(
        """SELECT 1 FROM teacher_assignments
           WHERE teacher_id = ? AND classroom_id = ? AND subject_id = ?""",
        (teacher_id, classroom_id, subject_id),
    ).fetchone() is not None


def parse_assignment_values(values):
    assignments = set()
    valid = True
    for value in values:
        try:
            subject_id, classroom_id = (int(part) for part in value.split(":"))
            if subject_id < 1 or classroom_id < 1:
                raise ValueError
        except (ValueError, TypeError):
            valid = False
            continue
        assignments.add((subject_id, classroom_id))
    return sorted(assignments), valid


def assignment_values_exist(db, assignments):
    return all(
        db.execute("SELECT 1 FROM subjects WHERE id = ?", (subject_id,)).fetchone()
        and db.execute(
            "SELECT 1 FROM classrooms WHERE id = ?", (classroom_id,)
        ).fetchone()
        for subject_id, classroom_id in assignments
    )


app = create_app()


if __name__ == "__main__":
    app.run(debug=os.environ.get("FLASK_DEBUG", "").lower() == "true")
