import io
import os
import secrets
import sqlite3
from datetime import date, datetime, timezone
from functools import wraps
from pathlib import Path
from urllib.parse import urlsplit

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
from openpyxl import Workbook
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename


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
    name TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS subjects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS grades (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id INTEGER NOT NULL REFERENCES students(id),
    subject_id INTEGER NOT NULL REFERENCES subjects(id),
    teacher_id INTEGER REFERENCES users(id),
    teacher_name TEXT NOT NULL DEFAULT '',
    score REAL NOT NULL CHECK (score >= 0 AND score <= 100),
    term TEXT NOT NULL,
    created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
    UNIQUE (student_id, subject_id, term)
);

CREATE TABLE IF NOT EXISTS teacher_assignments (
    teacher_id INTEGER NOT NULL REFERENCES users(id) ON DELETE CASCADE,
    subject_id INTEGER NOT NULL REFERENCES subjects(id) ON DELETE CASCADE,
    classroom_id INTEGER NOT NULL REFERENCES classrooms(id) ON DELETE CASCADE,
    PRIMARY KEY (teacher_id, subject_id, classroom_id)
);
"""


def create_app(test_config=None):
    app = Flask(__name__, instance_relative_config=True)
    app.config.from_mapping(
        SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
        DATABASE=os.path.join(app.instance_path, "instituto.sqlite"),
        SESSION_COOKIE_HTTPONLY=True,
        SESSION_COOKIE_SAMESITE="Lax",
        SESSION_COOKIE_SECURE=os.environ.get("COOKIE_SECURE", "").lower() == "true",
    )
    if test_config:
        app.config.update(test_config)

    Path(app.instance_path).mkdir(parents=True, exist_ok=True)
    CSRFProtect(app)

    @app.teardown_appcontext
    def close_db(error=None):
        db = g.pop("db", None)
        if db is not None:
            db.close()

    with app.app_context():
        db = get_db()
        db.executescript(SCHEMA)
        migrate_grade_teacher_snapshot(db)
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
        password = click.prompt("Contrasena (minimo 12 caracteres)", hide_input=True)
        if len(password) < 12:
            raise click.ClickException("La contrasena debe tener al menos 12 caracteres.")
        try:
            get_db().execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, 'admin')",
                (username.strip(), generate_password_hash(password)),
            )
            get_db().commit()
        except sqlite3.IntegrityError as exc:
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
        if g.user["role"] == "admin":
            grade_count = db.execute("SELECT COUNT(*) FROM grades").fetchone()[0]
            student_count = db.execute("SELECT COUNT(*) FROM students").fetchone()[0]
            classroom_count = db.execute("SELECT COUNT(*) FROM classrooms").fetchone()[0]
        else:
            grade_count = db.execute(
                """SELECT COUNT(*) FROM grades g
                   JOIN students s ON s.id = g.student_id
                   JOIN teacher_assignments ta
                     ON ta.teacher_id = ? AND ta.subject_id = g.subject_id
                    AND ta.classroom_id = s.classroom_id""",
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
                           ) AND subject_id = ? AND term = ?""",
                        (selected_classroom, subject_id, term),
                    ).fetchall()
                }

        if request.method == "POST":
            if not subject_id or not db.execute(
                "SELECT 1 FROM subjects WHERE id = ?", (subject_id,)
            ).fetchone():
                flash("Selecciona una materia valida.", "error")
            elif not term:
                flash("Indica el periodo o trimestre.", "error")
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
                        score = float(raw_score.replace(",", "."))
                    except ValueError:
                        invalid_score = True
                        break
                    if not 0 <= score <= 100:
                        invalid_score = True
                        break
                    entries.append((student["id"], subject_id, g.user["id"], score, term))

                if invalid_score:
                    flash("Cada nota debe ser un numero entre 0 y 100.", "error")
                elif not entries:
                    flash("No hay notas nuevas para guardar.", "error")
                else:
                    try:
                        db.executemany(
                            """INSERT INTO grades
                               (student_id, subject_id, teacher_id, teacher_name, score, term)
                               VALUES (?, ?, ?, ?, ?, ?)""",
                            [
                                (*entry[:3], g.user["username"], *entry[3:])
                                for entry in entries
                            ],
                        )
                        db.commit()
                    except sqlite3.IntegrityError:
                        db.rollback()
                        flash(
                            "Una o mas notas ya fueron registradas para ese periodo. "
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
                        s.name AS student_name, c.grade, c.name AS classroom,
                        sub.name AS subject,
                        COALESCE(NULLIF(g.teacher_name, ''), u.username, 'Cuenta eliminada') AS teacher
                 FROM grades g
                 JOIN students s ON s.id = g.student_id
                 JOIN classrooms c ON c.id = s.classroom_id
                 JOIN subjects sub ON sub.id = g.subject_id
                 LEFT JOIN users u ON u.id = g.teacher_id
                 WHERE c.id = ? AND sub.id = ?"""
        params = [classroom_id, subject_id]
        sql += " ORDER BY g.term COLLATE NOCASE, s.name COLLATE NOCASE"
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
        rows = query_grade_rows(
            assigned_teacher_id=g.user["id"] if g.user["role"] == "teacher" else None
        )
        return create_grade_export(
            rows, f"notas_{date.today().isoformat()}.xlsx"
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
        return create_grade_export(rows, filename)

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
        students = db.execute(
            """SELECT s.id, s.name, c.grade, c.name AS classroom
               FROM students s
               JOIN classrooms c ON c.id = s.classroom_id
               ORDER BY c.grade, c.name, s.name"""
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
            students=students,
            subjects=subjects,
            assignment_options=assignment_options,
            assigned_by_teacher=assigned_by_teacher,
            current_user_id=g.user["id"],
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
        if not username or len(password) < 12 or role not in ("admin", "teacher"):
            flash("Completa el usuario, una contrasena de al menos 12 caracteres y un rol valido.", "error")
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
                    "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                    (username, generate_password_hash(password), role),
                )
                if role == "teacher":
                    db.executemany(
                        """INSERT INTO teacher_assignments
                           (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                        [
                            (cursor.lastrowid, subject_id, classroom_id)
                            for subject_id, classroom_id in assignments
                        ],
                    )
                db.commit()
            except sqlite3.IntegrityError:
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
                except sqlite3.IntegrityError:
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
            except sqlite3.IntegrityError:
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
            except sqlite3.IntegrityError:
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
            db.execute("DELETE FROM students WHERE classroom_id = ?", (classroom_id,))
            db.execute("DELETE FROM classrooms WHERE id = ?", (classroom_id,))
            db.commit()
        except sqlite3.IntegrityError:
            db.rollback()
            flash("No se pudo eliminar el grado y aula porque tiene datos relacionados.", "error")
        else:
            flash(
                f"Se elimino {classroom['grade']} · {classroom['name']} junto con sus alumnos y notas.",
                "success",
            )
        return redirect(url_for("admin"))

    @app.post("/admin/students")
    @role_required("admin")
    def create_student():
        classroom_id = request.form.get("classroom_id", type=int)
        name = request.form.get("name", "").strip()
        if not classroom_id or not name:
            flash("Selecciona el grado y aula e indica el nombre del alumno.", "error")
        elif not get_db().execute(
            "SELECT 1 FROM classrooms WHERE id = ?", (classroom_id,)
        ).fetchone():
            flash("El grado y aula seleccionados no existen.", "error")
        else:
            get_db().execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (classroom_id, name),
            )
            get_db().commit()
            flash("Alumno agregado.", "success")
        return redirect(url_for("admin"))

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
            db.execute("DELETE FROM students WHERE id = ?", (student_id,))
            db.commit()
        except sqlite3.IntegrityError:
            db.rollback()
            flash("No se pudo eliminar el alumno porque tiene datos relacionados.", "error")
        else:
            flash(f"Se elimino a {student['name']} junto con sus notas.", "success")
        return redirect(url_for("admin"))

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
            except sqlite3.IntegrityError:
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
        except sqlite3.IntegrityError:
            db.rollback()
            flash("No se pudo eliminar la materia porque tiene datos relacionados.", "error")
        else:
            flash(f"Se elimino la materia {subject['name']} junto con sus notas.", "success")
        return redirect(url_for("admin"))

    @app.route("/admin/grades/manage", methods=("GET", "POST"))
    @role_required("admin")
    def manage_grades():
        db = get_db()
        classrooms = db.execute(
            "SELECT id, grade, name FROM classrooms ORDER BY grade, name"
        ).fetchall()
        subjects = db.execute("SELECT id, name FROM subjects ORDER BY name").fetchall()
        classroom_id = request.values.get("classroom_id", type=int)
        subject_id = request.values.get("subject_id", type=int)
        term = request.values.get("term", "").strip()
        students = []
        existing_grades = {}

        if classroom_id:
            if not db.execute(
                "SELECT 1 FROM classrooms WHERE id = ?", (classroom_id,)
            ).fetchone():
                abort(404)
            students = db.execute(
                "SELECT id, name FROM students WHERE classroom_id = ? ORDER BY name",
                (classroom_id,),
            ).fetchall()
        if subject_id and not db.execute(
            "SELECT 1 FROM subjects WHERE id = ?", (subject_id,)
        ).fetchone():
            abort(404)
        if classroom_id and subject_id and term:
            existing_grades = {
                row["student_id"]: row
                for row in db.execute(
                    """SELECT id, student_id, score FROM grades
                       WHERE subject_id = ? AND term = ?
                         AND student_id IN (
                             SELECT id FROM students WHERE classroom_id = ?
                         )""",
                    (subject_id, term, classroom_id),
                ).fetchall()
            }

        if request.method == "POST":
            if not classroom_id or not subject_id or not term:
                flash("Selecciona el grado, aula, materia y periodo.", "error")
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
                        score = float(raw_score.replace(",", "."))
                    except ValueError:
                        invalid_score = True
                        break
                    if not 0 <= score <= 100:
                        invalid_score = True
                        break
                    entries.append((student["id"], score))

                if invalid_score:
                    flash("Cada nota debe ser un numero entre 0 y 100.", "error")
                elif not entries:
                    flash("No hay notas para guardar o actualizar.", "error")
                else:
                    timestamp = datetime.now(timezone.utc).isoformat()
                    try:
                        db.execute("BEGIN")
                        for student_id, score in entries:
                            existing = existing_grades.get(student_id)
                            if existing:
                                if existing["score"] != score:
                                    db.execute(
                                        """UPDATE grades SET score = ?, updated_at = ?
                                           WHERE id = ?""",
                                        (score, timestamp, existing["id"]),
                                    )
                            else:
                                db.execute(
                                    """INSERT INTO grades
                                       (student_id, subject_id, teacher_id, teacher_name,
                                        score, term)
                                       VALUES (?, ?, ?, ?, ?, ?)""",
                                    (
                                        student_id,
                                        subject_id,
                                        g.user["id"],
                                        g.user["username"],
                                        score,
                                        term,
                                    ),
                                )
                        db.commit()
                    except sqlite3.IntegrityError:
                        db.rollback()
                        flash("No se pudieron guardar las notas. Revisa si ya existe una nota para ese periodo.", "error")
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
                            )
                        )

        return render_template(
            "admin_grades.html",
            classrooms=classrooms,
            subjects=subjects,
            students=students,
            classroom_id=classroom_id,
            subject_id=subject_id,
            term=term,
            existing_grades=existing_grades,
        )

    @app.get("/admin/grades/<int:grade_id>/edit")
    @role_required("admin")
    def edit_grade(grade_id):
        grade = get_db().execute(
            """SELECT s.classroom_id, g.subject_id, g.term
               FROM grades g JOIN students s ON s.id = g.student_id
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


def create_grade_export(rows, filename):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Notas"
    sheet.append(
        ["Alumno", "Grado", "Aula", "Materia", "Periodo", "Nota", "Maestro", "Fecha de registro"]
    )
    for row in rows:
        values = [
            row["student_name"],
            row["grade"],
            row["classroom"],
            row["subject"],
            row["term"],
            row["score"],
            row["teacher"],
            row["created_at"],
        ]
        sheet.append(values)
        for column, value in enumerate(values, start=1):
            if isinstance(value, str):
                sheet.cell(row=sheet.max_row, column=column).data_type = "s"
    sheet.freeze_panes = "A2"
    sheet.auto_filter.ref = sheet.dimensions
    for column in sheet.columns:
        width = min(max(len(str(cell.value or "")) for cell in column) + 2, 40)
        sheet.column_dimensions[column[0].column_letter].width = width

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
        db = sqlite3.connect(current_app.config["DATABASE"])
        db.row_factory = sqlite3.Row
        db.execute("PRAGMA foreign_keys = ON")
        g.db = db
    return g.db


def migrate_grade_teacher_snapshot(db):
    columns = {
        row["name"]: row for row in db.execute("PRAGMA table_info(grades)").fetchall()
    }
    teacher_id_column = columns.get("teacher_id")
    if "teacher_name" in columns and (
        teacher_id_column is None or not teacher_id_column["notnull"]
    ):
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
            score REAL NOT NULL CHECK (score >= 0 AND score <= 100),
            term TEXT NOT NULL,
            created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
            UNIQUE (student_id, subject_id, term)
        )"""
    )
    db.execute(
        """INSERT INTO grades
           (id, student_id, subject_id, teacher_id, teacher_name, score, term,
            created_at, updated_at)
           SELECT old.id, old.student_id, old.subject_id, old.teacher_id,
                  COALESCE(users.username, ''), old.score, old.term,
                  old.created_at, old.updated_at
           FROM grades_legacy AS old
           LEFT JOIN users ON users.id = old.teacher_id"""
    )
    db.execute("DROP TABLE grades_legacy")
    db.commit()


def query_grade_rows(assigned_teacher_id=None, classroom_id=None, subject_id=None):
    db = get_db()
    sql = """SELECT g.id, g.score, g.term, g.created_at, g.updated_at,
                    s.name AS student_name, c.grade, c.name AS classroom,
                    sub.name AS subject,
                    COALESCE(NULLIF(g.teacher_name, ''), u.username, 'Cuenta eliminada') AS teacher,
                    s.classroom_id, g.subject_id
             FROM grades g
             JOIN students s ON s.id = g.student_id
             JOIN classrooms c ON c.id = s.classroom_id
             JOIN subjects sub ON sub.id = g.subject_id
             LEFT JOIN users u ON u.id = g.teacher_id"""
    params = []
    conditions = []
    if assigned_teacher_id is not None:
        conditions.append(
            """EXISTS (
                SELECT 1 FROM teacher_assignments ta
                WHERE ta.teacher_id = ? AND ta.subject_id = g.subject_id
                  AND ta.classroom_id = s.classroom_id
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
    sql += " ORDER BY sub.name COLLATE NOCASE, c.grade, c.name, g.term, s.name"
    return db.execute(sql, params).fetchall()


def query_grade_summaries(teacher_id=None):
    if teacher_id is not None:
        sql = """SELECT c.id AS classroom_id, c.grade, c.name AS classroom,
                        sub.id AS subject_id, sub.name AS subject,
                        COUNT(DISTINCT g.id) AS grade_count,
                        COUNT(DISTINCT g.term) AS period_count
                 FROM teacher_assignments ta
                 JOIN classrooms c ON c.id = ta.classroom_id
                 JOIN subjects sub ON sub.id = ta.subject_id
                 LEFT JOIN students s ON s.classroom_id = c.id
                 LEFT JOIN grades g ON g.student_id = s.id AND g.subject_id = sub.id
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
                 JOIN classrooms c ON c.id = s.classroom_id
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
