import io
import os
import sqlite3
import tempfile
import unittest
from datetime import date
from io import BytesIO
from unittest.mock import patch
from openpyxl import Workbook, load_workbook
from werkzeug.security import check_password_hash, generate_password_hash

from app import (
    _database_type,
    _database_url_from_environment,
    _normalize_postgresql_sql,
    create_app,
)


class GradeSystemTests(unittest.TestCase):
    def test_postgresql_database_urls_are_recognized(self):
        for database_url in (
            "postgres://user:pass@localhost/instituto",
            "postgresql://user:pass@localhost/instituto",
            "postgresql+pg8000://",
            "postgresql+psycopg://user:pass@localhost/instituto",
        ):
            with self.subTest(database_url=database_url):
                self.assertEqual(_database_type(database_url), "postgresql")

    def test_wasmer_database_environment_selects_managed_postgresql(self):
        wasmer_database_env = {
            "APP_DATABASE_ENGINE": "postgresql",
            "DB_HOST": "psql.fr-roub1.wasmer.app",
            "DB_PORT": "12345",
            "DB_NAME": "instituto",
            "DB_USERNAME": "instituto_user",
            "DB_PASSWORD": "db-password",
        }
        with patch.dict(os.environ, wasmer_database_env, clear=True):
            self.assertEqual(
                _database_url_from_environment("instance/instituto.sqlite"),
                "postgresql+pg8000://",
            )

    def test_incomplete_wasmer_database_environment_fails_explicitly(self):
        with patch.dict(os.environ, {"APP_DATABASE_ENGINE": "postgresql"}, clear=True):
            with self.assertRaisesRegex(RuntimeError, "DB_HOST"):
                _database_url_from_environment("instance/instituto.sqlite")

    def test_postgresql_sql_is_normalized_for_psycopg(self):
        sql, params = _normalize_postgresql_sql(
            "BEGIN IMMEDIATE; SELECT name FROM subjects "
            "WHERE id = ? ORDER BY name COLLATE NOCASE",
            (7,),
        )
        self.assertEqual(
            sql,
            "BEGIN; SELECT name FROM subjects WHERE id = %s ORDER BY name",
        )
        self.assertEqual(params, (7,))

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.database = os.path.join(self.temp_dir.name, "test.sqlite")
        self.app = create_app(
            {
                "TESTING": True,
                "WTF_CSRF_ENABLED": False,
                "DATABASE": self.database,
                "SECRET_KEY": "test-secret",
            }
        )
        self.client = self.app.test_client()
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                ("admin", generate_password_hash("admin-password"), "admin"),
            )
            db.execute(
                "INSERT INTO users (username, password_hash, role) VALUES (?, ?, ?)",
                ("maestra", generate_password_hash("teacher-password"), "teacher"),
            )
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("1ro", "A")
            )
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Matematicas",))
            db.execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (1, "Ana Perez"),
            )
            db.execute(
                """INSERT INTO teacher_assignments
                   (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                (2, 1, 1),
            )
            db.commit()

    def tearDown(self):
        self.temp_dir.cleanup()

    def login(self, username, password):
        return self.client.post(
            "/login", data={"username": username, "password": password}
        )

    def add_grade(self, score="90"):
        return self.client.post(
            "/grades/new",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": score,
            },
            follow_redirects=True,
        )

    def test_institution_name_appears_on_login_and_dashboard(self):
        response = self.client.get("/login")
        self.assertIn(b"INSTITUTO DR. MIGUEL JARQUIN VALLEJOS", response.data)

        self.login("maestra", "teacher-password")
        response = self.client.get("/")
        self.assertIn(b"INSTITUTO DR. MIGUEL JARQUIN VALLEJOS", response.data)

    def test_theme_toggle_is_available_before_and_after_login(self):
        response = self.client.get("/login")
        self.assertIn(b'id="theme-toggle"', response.data)
        self.assertIn(b"Modo oscuro", response.data)
        self.assertIn(b'role="switch"', response.data)
        self.assertIn(b'aria-checked="false"', response.data)

        self.login("maestra", "teacher-password")
        response = self.client.get("/")
        self.assertIn(b'id="theme-toggle"', response.data)
        self.assertIn(b"localStorage.setItem", response.data)

    def test_grade_forms_use_evaluation_period_options(self):
        periods = (
            "Primer corte evaluativo",
            "Segundo corte evaluativo",
            "Tercer corte evaluativo",
            "Cuarto corte evaluativo",
        )
        self.login("maestra", "teacher-password")
        teacher_page = self.client.get(
            "/grades/new?classroom_id=1&subject_id=1"
        )
        for period in periods:
            self.assertIn(f'value="{period}"'.encode(), teacher_page.data)
        invalid_teacher_grade = self.client.post(
            "/grades/new",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Periodo libre",
                "score_1": "90",
            },
            follow_redirects=True,
        )
        self.assertIn(
            b"Selecciona un corte evaluativo valido.",
            invalid_teacher_grade.data,
        )
        self.client.post("/logout")

        self.login("admin", "admin-password")
        admin_page = self.client.get(
            "/admin/grades/manage?classroom_id=1&subject_id=1"
        )
        for period in periods:
            self.assertIn(f'value="{period}"'.encode(), admin_page.data)
        invalid_admin_grade = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Periodo libre",
                "score_1": "90",
            },
            follow_redirects=True,
        )
        self.assertIn(
            b"Selecciona el grado, aula, materia y corte evaluativo.",
            invalid_admin_grade.data,
        )
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0],
                0,
            )

    def test_grades_are_saved_independently_for_each_evaluation_period(self):
        self.login("maestra", "teacher-password")
        for period, score in (
            ("Primer corte evaluativo", "81"),
            ("Segundo corte evaluativo", "92"),
        ):
            response = self.client.post(
                "/grades/new",
                data={
                    "classroom_id": "1",
                    "subject_id": "1",
                    "term": period,
                    "score_1": score,
                },
                follow_redirects=True,
            )
            self.assertIn(b"Se guardaron 1 notas", response.data)

        with self.app.app_context():
            from app import get_db

            grades = get_db().execute(
                """SELECT term, score FROM grades
                   WHERE student_id = 1 AND subject_id = 1 ORDER BY term"""
            ).fetchall()
            self.assertEqual(
                [(grade["term"], grade["score"]) for grade in grades],
                [
                    ("Primer corte evaluativo", 81),
                    ("Segundo corte evaluativo", 92),
                ],
            )

        self.client.post("/logout")
        self.login("admin", "admin-password")
        response = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": "70",
            },
            follow_redirects=True,
        )
        self.assertIn(b"1 calificacion.", response.data)
        with self.app.app_context():
            from app import get_db

            grades = get_db().execute(
                """SELECT term, score FROM grades
                   WHERE student_id = 1 AND subject_id = 1 ORDER BY term"""
            ).fetchall()
            self.assertEqual(
                [(grade["term"], grade["score"]) for grade in grades],
                [
                    ("Primer corte evaluativo", 70),
                    ("Segundo corte evaluativo", 92),
                ],
            )

    def test_teacher_can_record_grade_but_cannot_change_or_delete_it(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")

        response = self.client.post(
            "/grades/new",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": "40",
            },
            follow_redirects=True,
        )
        self.assertIn(b"bloqueadas", response.data)
        self.assertEqual(self.client.get("/admin").status_code, 403)
        self.assertEqual(
            self.client.post("/admin/grades/1/delete").status_code, 403
        )
        with self.app.app_context():
            from app import get_db

            row = get_db().execute("SELECT score FROM grades WHERE id = 1").fetchone()
            self.assertEqual(row["score"], 85)

    def test_invalid_grade_is_not_saved(self):
        self.login("maestra", "teacher-password")
        response = self.add_grade("101")
        self.assertIn(b"entre 0 y 100", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0], 0
            )

    def test_teacher_can_record_text_grade_and_export_it(self):
        self.login("maestra", "teacher-password")
        response = self.add_grade("A+")
        self.assertIn(b"Se guardaron 1 notas", response.data)
        detail = self.client.get("/grades/classroom/1?subject_id=1")
        self.assertIn(b"<strong>A+</strong>", detail.data)
        export = self.client.get("/grades/export.xlsx")
        workbook = load_workbook(io.BytesIO(export.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(rows[3][7], "A+")
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT score FROM grades").fetchone()["score"],
                "A+",
            )

    def test_text_grade_is_limited_to_three_characters(self):
        self.login("maestra", "teacher-password")
        response = self.add_grade("N/A!")
        self.assertIn(b"texto de hasta 3 caracteres", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0],
                0,
            )

    def test_admin_can_save_and_update_text_grade(self):
        self.login("admin", "admin-password")
        response = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": "A+",
            },
            follow_redirects=True,
        )
        self.assertIn(b"1 calificacion.", response.data)
        response = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": "N/A",
            },
            follow_redirects=True,
        )
        self.assertIn(b"1 calificacion.", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT score FROM grades").fetchone()["score"],
                "N/A",
            )

    def test_admin_can_edit_and_delete_grade(self):
        self.login("maestra", "teacher-password")
        self.add_grade()
        self.client.post("/logout")
        self.login("admin", "admin-password")

        response = self.client.get("/admin/grades/1/edit")
        self.assertEqual(response.status_code, 302)
        self.assertIn("classroom_id=1", response.headers["Location"])
        response = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": "95",
            },
            follow_redirects=True,
        )
        self.assertIn(b"1 calificacion.", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT score FROM grades WHERE id = 1").fetchone()["score"],
                95,
            )

        response = self.client.post(
            "/admin/grades/1/delete", follow_redirects=True
        )
        self.assertIn(b"Nota eliminada", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0], 0
            )

    def test_admin_can_add_notes_for_students_from_class_list(self):
        with self.app.app_context():
            from app import get_db

            get_db().execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (1, "Luis Gomez"),
            )
            get_db().commit()
        self.login("admin", "admin-password")

        response = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Segundo corte evaluativo",
                "score_1": "92",
                "score_2": "87",
            },
            follow_redirects=True,
        )
        self.assertIn(b"2 calificaciones.", response.data)
        with self.app.app_context():
            from app import get_db

            rows = get_db().execute(
                "SELECT score, teacher_name FROM grades ORDER BY student_id"
            ).fetchall()
            self.assertEqual([row["score"] for row in rows], [92, 87])
            self.assertEqual({row["teacher_name"] for row in rows}, {"admin"})

    def test_admin_can_delete_teacher_account_and_keep_grade_history(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")
        self.client.post("/logout")
        self.login("admin", "admin-password")

        response = self.client.post(
            "/admin/users/2/delete", follow_redirects=True
        )
        self.assertIn(b"se conservaron sus notas", response.data)
        with self.app.app_context():
            from app import get_db

            grade = get_db().execute(
                "SELECT teacher_id, teacher_name FROM grades WHERE id = 1"
            ).fetchone()
            self.assertIsNone(grade["teacher_id"])
            self.assertEqual(grade["teacher_name"], "maestra")
            self.assertEqual(
                get_db().execute("SELECT COUNT(*) FROM users WHERE id = 2").fetchone()[0],
                0,
            )
        response = self.client.get(
            "/grades/classroom/1?subject_id=1"
        )
        self.assertIn(b"maestra", response.data)

    def test_admin_created_grades_are_visible_and_downloadable_to_assigned_teacher(self):
        self.login("admin", "admin-password")
        response = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "score_1": "93",
            },
            follow_redirects=True,
        )
        self.assertIn(b"1 calificacion.", response.data)
        self.client.post("/logout")
        self.login("maestra", "teacher-password")

        summary = self.client.get("/grades")
        self.assertIn(b"Matematicas", summary.data)
        self.assertIn(b"1ro", summary.data)
        detail = self.client.get("/grades/classroom/1?subject_id=1")
        self.assertIn(b"Ana Perez", detail.data)
        self.assertIn(b"<strong>93.0</strong>", detail.data)
        workbook_response = self.client.get("/grades/export/1/1.xlsx")
        workbook = load_workbook(io.BytesIO(workbook_response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(rows[3][7], "93/AA")
        self.assertEqual(rows[3][8], "admin")

    def test_teacher_cannot_access_or_export_unassigned_combination(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (1, 2, str(date.today().year), 1, "admin", 81, "Primer corte evaluativo"),
            )
            db.commit()
        self.login("maestra", "teacher-password")

        self.assertEqual(
            self.client.get("/grades/classroom/1?subject_id=2").status_code, 403
        )
        self.assertEqual(self.client.get("/grades/export/1/2.xlsx").status_code, 403)
        self.assertEqual(
            self.client.get(
                "/grades/new?classroom_id=1&subject_id=2&term=Primer"
            ).status_code,
            403,
        )
        self.assertNotIn(b"81", self.client.get("/grades").data)

    def test_admin_can_create_teacher_with_multiple_exact_assignments(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B")
            )
            db.commit()
        self.login("admin", "admin-password")

        response = self.client.post(
            "/admin/users",
            data={
                "username": "profe2",
                "password": "teacher-password-2",
                "role": "teacher",
                "assignments": ["1:1", "2:1", "1:2"],
            },
            follow_redirects=True,
        )
        self.assertIn(b"Cuenta creada con sus asignaciones", response.data)
        with self.app.app_context():
            from app import get_db

            teacher_id = get_db().execute(
                "SELECT id FROM users WHERE username = 'profe2'"
            ).fetchone()["id"]
            assignments = get_db().execute(
                """SELECT subject_id, classroom_id FROM teacher_assignments
                   WHERE teacher_id = ? ORDER BY subject_id, classroom_id""",
                (teacher_id,),
            ).fetchall()
            self.assertEqual(
                [(row["subject_id"], row["classroom_id"]) for row in assignments],
                [(1, 1), (1, 2), (2, 1)],
            )
        self.client.get(f"/admin/users/{teacher_id}/assignments")
        self.client.post("/logout")
        self.login("profe2", "teacher-password-2")
        response = self.client.get("/grades/new?classroom_id=2")
        self.assertEqual(response.status_code, 200)
        self.assertIn(b'name="classroom_id" required onchange=', response.data)
        self.assertIn(
            b"this.form.elements.subject_id.value = ''; this.form.submit()",
            response.data,
        )
        self.assertIn(b">Matematicas</option>", response.data)
        self.assertEqual(
            self.client.get("/grades/new?classroom_id=2&subject_id=2").status_code,
            403,
        )

    def test_account_passwords_require_six_characters(self):
        self.login("admin", "admin-password")
        admin_page = self.client.get("/admin")
        self.assertIn(b'minlength="6"', admin_page.data)

        rejected_creation = self.client.post(
            "/admin/users",
            data={"username": "short-pass", "password": "12345", "role": "admin"},
            follow_redirects=True,
        )
        self.assertIn(b"al menos 6 caracteres", rejected_creation.data)
        with self.app.app_context():
            from app import get_db

            self.assertIsNone(
                get_db().execute(
                    "SELECT id FROM users WHERE username = 'short-pass'"
                ).fetchone()
            )

        accepted_creation = self.client.post(
            "/admin/users",
            data={"username": "six-pass", "password": "123456", "role": "admin"},
            follow_redirects=True,
        )
        self.assertIn(b"Cuenta creada.", accepted_creation.data)
        with self.app.app_context():
            from app import get_db

            user_id = get_db().execute(
                "SELECT id FROM users WHERE username = 'six-pass'"
            ).fetchone()["id"]

        edit_page = self.client.get(f"/admin/users/{user_id}/edit")
        self.assertIn(b'minlength="6"', edit_page.data)
        rejected_password_change = self.client.post(
            f"/admin/users/{user_id}/edit",
            data={"username": "six-pass", "password": "12345", "role": "admin"},
            follow_redirects=True,
        )
        self.assertIn(b"al menos 6 caracteres", rejected_password_change.data)
        accepted_password_change = self.client.post(
            f"/admin/users/{user_id}/edit",
            data={"username": "six-pass", "password": "abcdef", "role": "admin"},
            follow_redirects=True,
        )
        self.assertIn(b"Cuenta actualizada.", accepted_password_change.data)

    def test_init_admin_accepts_six_character_password(self):
        with self.app.app_context():
            from app import get_db

            get_db().execute("DELETE FROM users WHERE role = 'admin'")
            get_db().commit()

        runner = self.app.test_cli_runner()
        rejected_password = runner.invoke(
            args=["init-admin", "--username", "cli-admin"],
            input="12345\n",
        )
        self.assertEqual(rejected_password.exit_code, 1)
        self.assertIn("al menos 6 caracteres", rejected_password.output)

        accepted_password = runner.invoke(
            args=["init-admin", "--username", "cli-admin"],
            input="123456\n",
        )
        self.assertEqual(accepted_password.exit_code, 0)
        self.assertIn("Administrador creado.", accepted_password.output)

    def test_admin_can_edit_accounts_roles_passwords_and_assignments(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B")
            )
            db.commit()
        self.login("admin", "admin-password")

        admin_page = self.client.get("/admin")
        self.assertIn(b'href="/admin/users/1/edit">Editar</a>', admin_page.data)
        self.assertIn(b'href="/admin/users/2/edit">Editar</a>', admin_page.data)

        self_edit_page = self.client.get("/admin/users/1/edit")
        self.assertEqual(self_edit_page.status_code, 200)
        self.assertIn(b'name="password"', self_edit_page.data)
        self.assertIn(b'name="role" value="admin"', self_edit_page.data)
        protected_self_edit = self.client.post(
            "/admin/users/1/edit",
            data={
                "username": "admin",
                "role": "teacher",
                "assignments": ["1:1"],
            },
            follow_redirects=True,
        )
        self.assertIn(b"No puedes cambiar el rol de tu propia cuenta.", protected_self_edit.data)

        admin_update = self.client.post(
            "/admin/users/1/edit",
            data={
                "username": "admin-nuevo",
                "password": "new-admin-password",
                "role": "admin",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Cuenta actualizada.", admin_update.data)

        teacher_page = self.client.get("/admin/users/2/edit")
        self.assertIn(b'value="1:1" checked', teacher_page.data)
        teacher_update = self.client.post(
            "/admin/users/2/edit",
            data={
                "username": "maestra-actualizada",
                "password": "new-teacher-password",
                "role": "teacher",
                "assignments": ["2:2"],
            },
            follow_redirects=True,
        )
        self.assertIn(b"Cuenta actualizada.", teacher_update.data)

        with self.app.app_context():
            from app import get_db

            db = get_db()
            admin = db.execute(
                "SELECT username, role, password_hash FROM users WHERE id = 1"
            ).fetchone()
            teacher = db.execute(
                "SELECT username, role, password_hash FROM users WHERE id = 2"
            ).fetchone()
            assignments = db.execute(
                """SELECT subject_id, classroom_id FROM teacher_assignments
                   WHERE teacher_id = 2"""
            ).fetchall()
            self.assertEqual((admin["username"], admin["role"]), ("admin-nuevo", "admin"))
            self.assertTrue(check_password_hash(admin["password_hash"], "new-admin-password"))
            self.assertEqual(
                (teacher["username"], teacher["role"]),
                ("maestra-actualizada", "teacher"),
            )
            self.assertTrue(
                check_password_hash(teacher["password_hash"], "new-teacher-password")
            )
            self.assertEqual(
                [(row["subject_id"], row["classroom_id"]) for row in assignments],
                [(2, 2)],
            )

        role_change = self.client.post(
            "/admin/users/2/edit",
            data={"username": "maestra-actualizada", "password": "", "role": "admin"},
            follow_redirects=True,
        )
        self.assertIn(b"Cuenta actualizada.", role_change.data)
        with self.app.app_context():
            from app import get_db

            db = get_db()
            teacher = db.execute(
                "SELECT username, role, password_hash FROM users WHERE id = 2"
            ).fetchone()
            self.assertEqual((teacher["username"], teacher["role"]), ("maestra-actualizada", "admin"))
            self.assertTrue(
                check_password_hash(teacher["password_hash"], "new-teacher-password")
            )
            self.assertEqual(
                db.execute(
                    "SELECT COUNT(*) FROM teacher_assignments WHERE teacher_id = 2"
                ).fetchone()[0],
                0,
            )

    def test_admin_can_replace_existing_teacher_assignments(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.commit()
        self.login("admin", "admin-password")

        response = self.client.get("/admin/users/2/assignments")
        self.assertIn(b"Asignaciones de maestra", response.data)
        response = self.client.post(
            "/admin/users/2/assignments",
            data={"assignments": ["2:1"]},
            follow_redirects=True,
        )
        self.assertIn(b"Asignaciones de maestra actualizadas", response.data)
        with self.app.app_context():
            from app import get_db

            rows = get_db().execute(
                "SELECT subject_id, classroom_id FROM teacher_assignments WHERE teacher_id = 2"
            ).fetchall()
            self.assertEqual(
                [(row["subject_id"], row["classroom_id"]) for row in rows],
                [(2, 1)],
            )

    def test_grade_summary_hides_student_and_score_until_opening_classroom(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")

        summary = self.client.get("/grades")
        self.assertIn(b"Matematicas", summary.data)
        self.assertIn(b"1ro", summary.data)
        self.assertIn(b"Ver notas", summary.data)
        self.assertNotIn(b"Ana Perez", summary.data)
        self.assertNotIn(b">85<", summary.data)

        detail = self.client.get("/grades/classroom/1?subject_id=1")
        self.assertIn(b"Ana Perez", detail.data)
        self.assertIn(b"<strong>85.0</strong>", detail.data)
        self.assertNotIn(b"Editar", detail.data)

    def test_classroom_subject_export_contains_only_selected_group_and_subject(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (1, 2, str(date.today().year), 2, "maestra", 96, "Primer corte evaluativo"),
            )
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B")
            )
            db.execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (2, "Luis Gomez"),
            )
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (2, 1, str(date.today().year), 2, "maestra", 78, "Primer corte evaluativo"),
            )
            db.commit()

        response = self.client.get("/grades/export/1/1.xlsx")
        self.assertEqual(response.status_code, 200)
        self.assertIn("notas_Matematicas_1ro_A.xlsx", response.headers["Content-Disposition"])
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(
            rows[0][0],
            "INSTITUTO DR. MIGUEL JARQUIN VALLEJOS",
        )
        self.assertEqual(
            rows[1][0],
            "Corte evaluativo: Primer corte evaluativo",
        )
        self.assertEqual(
            rows[2][:9],
            (
                "N.º",
                "Modalidad",
                "Nivel",
                "Sección",
                "Nombre y apellido",
                "Sexo",
                "Código de persona",
                "Matematicas",
                "Maestro",
            ),
        )
        self.assertEqual(
            rows[3][:9],
            (
                1,
                "Regular",
                "1ro",
                "A",
                "Ana Perez",
                None,
                None,
                "85/AS",
                "maestra",
            ),
        )

        summary = self.client.get("/grades")
        self.assertIn(b"/grades/export/1/1.xlsx", summary.data)
        detail = self.client.get("/grades/classroom/1?subject_id=1")
        self.assertIn(b"/grades/export/1/1.xlsx", detail.data)
        self.assertEqual(self.client.get("/grades/export/999/1.xlsx").status_code, 404)

    def test_teacher_individual_export_includes_assigned_admin_and_teacher_grades(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (1, 1, str(date.today().year), 1, "admin", 70, "Segundo corte evaluativo"),
            )
            db.commit()
        self.login("maestra", "teacher-password")
        self.add_grade("85")

        response = self.client.get("/grades/export/1/1.xlsx")
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        self.assertEqual(
            workbook.sheetnames,
            ["Primer corte evaluativo", "Segundo corte evaluativo"],
        )
        exported_period_scores = set()
        for sheet in workbook.worksheets:
            rows = list(sheet.values)
            self.assertEqual(rows[0][0], "INSTITUTO DR. MIGUEL JARQUIN VALLEJOS")
            self.assertEqual(rows[1][0], f"Corte evaluativo: {sheet.title}")
            exported_period_scores.add((sheet.title, rows[3][7]))
        self.assertEqual(
            exported_period_scores,
            {("Primer corte evaluativo", "85/AS"), ("Segundo corte evaluativo", "70/AF")},
        )

    def test_admin_detail_has_individual_edit_action_and_full_class_roster_link(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")
        self.client.post("/logout")
        self.login("admin", "admin-password")

        response = self.client.get("/grades/classroom/1?subject_id=1")
        self.assertIn(b"Ana Perez", response.data)
        self.assertIn(b"Editar", response.data)
        self.assertIn(b"/admin/grades/1/edit", response.data)

        edit_response = self.client.get("/admin/grades/1/edit")
        self.assertIn("classroom_id=1", edit_response.headers["Location"])
        self.assertIn("subject_id=1", edit_response.headers["Location"])
        self.assertIn(
            "term=Primer+corte+evaluativo", edit_response.headers["Location"]
        )

    def test_admin_can_delete_classroom_and_related_students_and_grades(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")
        self.client.post("/logout")
        self.login("admin", "admin-password")

        response = self.client.post(
            "/admin/classrooms/1/delete", follow_redirects=True
        )
        self.assertIn(b"junto con sus alumnos y notas", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM classrooms").fetchone()[0], 0)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM students").fetchone()[0], 0)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0], 0)

    def test_admin_can_delete_student_and_related_grades(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")
        self.client.post("/logout")
        self.login("admin", "admin-password")

        response = self.client.post(
            "/admin/students/1/delete", follow_redirects=True
        )
        self.assertIn(b"junto con sus notas", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM students").fetchone()[0], 0)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0], 0)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM classrooms").fetchone()[0], 1)

    def test_admin_student_page_is_separate_sorted_and_has_bulk_add_form(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "A")
            )
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("1ro", "B")
            )
            db.execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (2, "Luis Gomez"),
            )
            db.execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (3, "Beatriz Ruiz"),
            )
            db.commit()

        self.login("admin", "admin-password")
        admin_response = self.client.get("/admin")
        self.assertIn(b' href="/admin/students">Gestionar estudiantes', admin_response.data)
        self.assertNotIn(b"Luis Gomez", admin_response.data)

        response = self.client.get("/admin/students")
        self.assertEqual(response.status_code, 200)
        content = response.data.decode()
        self.assertLess(content.index("1ro · Aula A"), content.index("1ro · Aula B"))
        self.assertLess(content.index("1ro · Aula B"), content.index("2do · Aula A"))
        self.assertIn(b'action="/admin/students/bulk"', response.data)
        self.assertNotIn(b"<h2>Agregar estudiante</h2>", response.data)
        self.assertIn(b'<option value="Regular">Regular</option>', response.data)
        self.assertIn(b'id="student-rows"', response.data)
        self.assertIn(b'id="student-row-count"', response.data)
        self.assertIn(b'id="add-student-rows"', response.data)
        self.assertIn(b'max="100"', response.data)
        self.assertIn(b"<th scope=\"col\">Nombre</th>", response.data)
        self.assertIn(b'href="/admin/students/template.xlsx"', response.data)
        self.assertIn(b'name="student_file"', response.data)

        modality_response = self.client.post(
            "/admin/modalities",
            data={"name": "Nocturna"},
            follow_redirects=True,
        )
        self.assertIn(b"Modalidad agregada.", modality_response.data)
        self.assertIn(b'<option value="Nocturna">Nocturna</option>', modality_response.data)
        invalid_student_response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "student_name": ["Modalidad invalida"],
                "student_sex": ["F"],
                "student_person_code": ["PERSON-INVALID"],
                "modality": "No existe",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Selecciona una modalidad valida.", invalid_student_response.data)
        add_response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "student_name": ["Ana Lopez"],
                "student_sex": ["F"],
                "student_person_code": ["PERSON-001"],
                "modality": "Nocturna",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Se agregaron 1 estudiantes.", add_response.data)
        self.assertIn(b"Ana Lopez", add_response.data)
        self.assertIn(b"Modalidad: Nocturna", add_response.data)
        self.assertIn(b"Sexo: F", add_response.data)
        self.assertIn("Código: PERSON-001".encode("utf-8"), add_response.data)
        duplicate_code_response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "student_name": ["Luis Lopez"],
                "student_sex": ["M"],
                "student_person_code": ["PERSON-001"],
                "modality": "Regular",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Se agregaron 1 estudiantes.", duplicate_code_response.data)
        with self.app.app_context():
            from app import get_db

            student = get_db().execute(
                "SELECT modality, sex, person_code FROM students WHERE name = ?",
                ("Ana Lopez",),
            ).fetchone()
            self.assertEqual(student["modality"], "Nocturna")
            self.assertEqual(student["sex"], "F")
            self.assertEqual(student["person_code"], "PERSON-001")
            repeated_code_student = get_db().execute(
                "SELECT person_code FROM students WHERE name = 'Luis Lopez'"
            ).fetchone()
            self.assertEqual(repeated_code_student["person_code"], "PERSON-001")
            self.assertIsNone(
                get_db().execute(
                    "SELECT id FROM students WHERE name = 'Modalidad invalida'"
                ).fetchone()
            )

    def test_admin_can_add_multiple_students_to_one_class_and_modality(self):
        self.login("admin", "admin-password")
        response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "modality": "Regular",
                "student_name": ["Ana Lopez", "Luis Gomez"],
                "student_sex": ["F", "M"],
                "student_person_code": ["PERSON-101", "PERSON-102"],
            },
            follow_redirects=True,
        )

        self.assertIn(b"Se agregaron 2 estudiantes.", response.data)
        with self.app.app_context():
            from app import get_db

            students = get_db().execute(
                """SELECT classroom_id, modality, sex, person_code
                   FROM students WHERE person_code LIKE 'PERSON-10%'"""
            ).fetchall()
            self.assertEqual(len(students), 2)
            self.assertTrue(all(student["classroom_id"] == 1 for student in students))
            self.assertTrue(all(student["modality"] == "Regular" for student in students))
            self.assertEqual(
                {student["sex"] for student in students}, {"F", "M"}
            )

    def test_admin_can_import_students_from_excel_and_download_template(self):
        self.login("admin", "admin-password")
        template_response = self.client.get("/admin/students/template.xlsx")
        self.assertEqual(template_response.status_code, 200)
        self.assertEqual(
            template_response.mimetype,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        template = load_workbook(BytesIO(template_response.data), read_only=True)
        self.assertEqual(
            next(template.active.iter_rows(values_only=True)),
            ("Nombre", "Sexo", "Código de persona"),
        )
        template.close()

        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Nombre y apellido", "Sexo", "Código único de persona"])
        sheet.append(["Ana Lopez", "F", "PERSON-201"])
        sheet.append(["Luis Gomez", "M", "PERSON-202"])
        excel_file = BytesIO()
        workbook.save(excel_file)
        excel_file.seek(0)

        response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "modality": "Regular",
                "student_file": (excel_file, "estudiantes.xlsx"),
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )

        self.assertIn(b"Se agregaron 2 estudiantes.", response.data)
        with self.app.app_context():
            from app import get_db

            students = get_db().execute(
                """SELECT classroom_id, modality, name, sex, person_code
                   FROM students WHERE person_code LIKE 'PERSON-20%' ORDER BY name"""
            ).fetchall()
            self.assertEqual(len(students), 2)
            self.assertEqual(
                [(student["name"], student["sex"], student["person_code"]) for student in students],
                [
                    ("Ana Lopez", "F", "PERSON-201"),
                    ("Luis Gomez", "M", "PERSON-202"),
                ],
            )
            self.assertTrue(
                all(
                    student["classroom_id"] == 1 and student["modality"] == "Regular"
                    for student in students
                )
            )

    def test_excel_import_uses_nivel_and_section_headers_without_modality(self):
        self.login("admin", "admin-password")
        imports = (
            ("Seccion", "UNDÉCIMO", "C", "Ana Grupo"),
            ("Sec.", "DUODÉCIMO", "D", "Luis Grupo"),
        )

        for section_header, grade, section, student_name in imports:
            workbook = Workbook()
            sheet = workbook.active
            sheet.append(["Nivel", section_header, "Nombre", "Sexo"])
            sheet.append([grade, section, student_name, "F"])
            excel_file = BytesIO()
            workbook.save(excel_file)
            excel_file.seek(0)

            response = self.client.post(
                "/admin/students/bulk",
                data={
                    "modality": "Regular",
                    "student_file": (excel_file, "estudiantes.xlsx"),
                },
                content_type="multipart/form-data",
                follow_redirects=True,
            )

            self.assertIn(b"Se agregaron 1 estudiantes.", response.data)

        with self.app.app_context():
            from app import get_db

            students = get_db().execute(
                """SELECT s.name, s.modality, c.grade, c.name AS section
                   FROM students s
                   JOIN classrooms c ON c.id = s.classroom_id
                   WHERE s.name IN ('Ana Grupo', 'Luis Grupo')
                   ORDER BY s.name"""
            ).fetchall()
            self.assertEqual(
                [tuple(student) for student in students],
                [
                    ("Ana Grupo", "Regular", "UNDÉCIMO", "C"),
                    ("Luis Grupo", "Regular", "DUODÉCIMO", "D"),
                ],
            )

    def test_excel_import_creates_missing_modalities_for_selected_classroom(self):
        self.login("admin", "admin-password")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Modalidad", "Nombre", "Sexo"])
        sheet.append(["Nocturna", "Ana Nocturna", "F"])
        sheet.append(["Sabatina", "Luis Sabatino", "M"])
        excel_file = BytesIO()
        workbook.save(excel_file)
        excel_file.seek(0)

        response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "student_file": (excel_file, "estudiantes.xlsx"),
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )

        self.assertIn(b"Se agregaron 2 estudiantes.", response.data)
        with self.app.app_context():
            from app import get_db

            rows = get_db().execute(
                """SELECT s.name, s.modality, c.grade, c.name AS section
                   FROM students s
                   JOIN classrooms c ON c.id = s.classroom_id
                   WHERE s.name IN ('Ana Nocturna', 'Luis Sabatino')
                   ORDER BY s.name"""
            ).fetchall()
            self.assertEqual(
                [tuple(row) for row in rows],
                [
                    ("Ana Nocturna", "Nocturna", "1ro", "A"),
                    ("Luis Sabatino", "Sabatina", "1ro", "A"),
                ],
            )
            self.assertEqual(
                get_db().execute(
                    """SELECT COUNT(*) FROM modalities
                       WHERE name IN ('Nocturna', 'Sabatina')"""
                ).fetchone()[0],
                2,
            )

    def test_excel_import_uses_group_columns_and_allows_missing_or_repeated_codes(self):
        self.login("admin", "admin-password")
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Sábana de estudiantes"])
        sheet.append(
            [
                "N.º",
                "Modalidad",
                "Nivel",
                "Sec.",
                "NOMBRES Y APELLIDOS",
                "Sexo",
                "Código de persona",
            ]
        )
        sheet.append([1, "SECUN.REGULAR", "UNDÉCIMO", "C", "Ana Lopez", "F", "PERSON-301"])
        sheet.append([2, "SECUN.REGULAR", "UNDÉCIMO", "C", "Luis Gomez", "M", "PERSON-301"])
        sheet.append([3, "SECUN.REGULAR", "UNDÉCIMO", "D", "Eva Ruiz", "F", None])
        sheet.append([4, None, None, None, "Mario Diaz", "M", None])
        excel_file = BytesIO()
        workbook.save(excel_file)
        excel_file.seek(0)

        response = self.client.post(
            "/admin/students/bulk",
            data={
                "student_file": (excel_file, "estudiantes.xlsx"),
            },
            content_type="multipart/form-data",
            follow_redirects=True,
        )

        self.assertIn(b"Se agregaron 4 estudiantes.", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM students WHERE person_code = 'PERSON-301'"
                ).fetchone()[0],
                2,
            )
            self.assertEqual(
                get_db().execute(
                    """SELECT COUNT(*) FROM students
                       WHERE person_code IS NULL AND name IN ('Eva Ruiz', 'Mario Diaz')"""
                ).fetchone()[0],
                2,
            )
            self.assertEqual(
                get_db().execute(
                    """SELECT COUNT(*) FROM students s
                       JOIN classrooms c ON c.id = s.classroom_id
                       WHERE c.grade = 'UNDÉCIMO' AND c.name = 'C'"""
                ).fetchone()[0],
                2,
            )
            self.assertEqual(
                get_db().execute(
                    """SELECT COUNT(*) FROM students s
                       JOIN classrooms c ON c.id = s.classroom_id
                       WHERE c.grade = 'UNDÉCIMO' AND c.name = 'D'"""
                ).fetchone()[0],
                2,
            )
            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM modalities WHERE name = 'SECUN.REGULAR'"
                ).fetchone()[0],
                1,
            )

    def test_admin_can_edit_student_to_complete_missing_details(self):
        self.login("admin", "admin-password")
        student_page = self.client.get("/admin/students")
        self.assertIn(b'href="/admin/students/1/edit">Editar</a>', student_page.data)

        edit_page = self.client.get("/admin/students/1/edit")
        self.assertEqual(edit_page.status_code, 200)
        self.assertIn(b'name="person_code"', edit_page.data)
        response = self.client.post(
            "/admin/students/1/edit",
            data={
                "classroom_id": "1",
                "name": "Ana Perez",
                "modality": "Regular",
                "sex": "F",
                "person_code": "PERSON-EDITED",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Datos del estudiante actualizados.", response.data)
        self.assertIn(b"Sexo: F", response.data)
        self.assertIn(b"C\xc3\xb3digo: PERSON-EDITED", response.data)

        with self.app.app_context():
            from app import get_db

            get_db().execute(
                """INSERT INTO students
                   (classroom_id, name, modality, sex, person_code)
                   VALUES (?, ?, ?, ?, ?)""",
                (1, "Otro alumno", "Regular", "M", "PERSON-TAKEN"),
            )
            get_db().commit()

        duplicate_code_response = self.client.post(
            "/admin/students/1/edit",
            data={
                "classroom_id": "1",
                "name": "Ana Perez",
                "modality": "Regular",
                "sex": "F",
                "person_code": "PERSON-TAKEN",
            },
            follow_redirects=True,
        )
        self.assertIn(
            b"Datos del estudiante actualizados.", duplicate_code_response.data
        )
        with self.app.app_context():
            from app import get_db

            student = get_db().execute(
                "SELECT person_code FROM students WHERE id = 1"
            ).fetchone()
            self.assertEqual(student["person_code"], "PERSON-TAKEN")

    def test_student_record_tracks_enrollments_and_grades_by_school_year(self):
        self.login("admin", "admin-password")
        current_year = str(date.today().year)
        previous_year = str(date.today().year - 1)
        with self.app.app_context():
            from app import get_db

            get_db().execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B")
            )
            get_db().commit()

        record = self.client.get("/admin/students/1/record")
        self.assertEqual(record.status_code, 200)
        self.assertIn(b"Datos del estudiante", record.data)
        self.assertIn(b"Historial de grados cursados", record.data)

        enrollment = self.client.post(
            "/admin/students/1/record",
            data={"school_year": previous_year, "classroom_id": "2"},
            follow_redirects=True,
        )
        self.assertIn(b"Matr\xc3\xadcula agregada al expediente.", enrollment.data)

        previous_grade = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "2",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "school_year": previous_year,
                "score_1": "88",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Se guardaron o actualizaron 1", previous_grade.data)

        current_grade = self.client.post(
            "/admin/grades/manage",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer corte evaluativo",
                "school_year": current_year,
                "score_1": "91",
            },
            follow_redirects=True,
        )
        self.assertIn(b"Se guardaron o actualizaron 1", current_grade.data)

        record = self.client.get("/admin/students/1/record")
        self.assertIn(f"<td>{previous_year}</td>".encode(), record.data)
        self.assertIn(b"<td>2do</td>", record.data)
        self.assertIn(b"<td>88.0</td>", record.data)
        self.assertIn(f"<td>{current_year}</td>".encode(), record.data)
        self.assertIn(b"<td>91.0</td>", record.data)
        with self.app.app_context():
            from app import get_db

            student = get_db().execute(
                "SELECT classroom_id FROM students WHERE id = 1"
            ).fetchone()
            self.assertEqual(student["classroom_id"], 1)

        duplicate_enrollment = self.client.post(
            "/admin/students/1/record",
            data={"school_year": previous_year, "classroom_id": "2"},
            follow_redirects=True,
        )
        self.assertIn(b"Ya existe una matr\xc3\xadcula para ese a\xc3\xb1o escolar.", duplicate_enrollment.data)

        updated_current_enrollment = self.client.post(
            "/admin/students/1/record",
            data={"school_year": current_year, "classroom_id": "2"},
            follow_redirects=True,
        )
        self.assertIn(b"Matr\xc3\xadcula actualizada en el expediente.", updated_current_enrollment.data)
        with self.app.app_context():
            from app import get_db

            student = get_db().execute(
                "SELECT classroom_id FROM students WHERE id = 1"
            ).fetchone()
            self.assertEqual(student["classroom_id"], 2)

    def test_admin_can_enroll_a_whole_classroom_in_one_action(self):
        current_year = str(date.today().year)
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (1, "Luis Gomez"),
            )
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B")
            )
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("3ro", "A")
            )
            db.execute(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                (3, "Beatriz Ruiz"),
            )
            db.commit()

        self.login("admin", "admin-password")
        page = self.client.get("/admin/students")
        self.assertIn(b"Matricular o promover un grupo completo", page.data)
        response = self.client.post(
            "/admin/students/enrollments/bulk",
            data={
                "source_classroom_id": "1",
                "target_classroom_id": "2",
                "school_year": current_year,
            },
            follow_redirects=True,
        )
        self.assertIn(b"2 actualizadas", response.data)

        with self.app.app_context():
            from app import get_db

            db = get_db()
            moved_students = db.execute(
                "SELECT id, classroom_id FROM students WHERE name IN (?, ?)",
                ("Ana Perez", "Luis Gomez"),
            ).fetchall()
            self.assertEqual(len(moved_students), 2)
            self.assertTrue(all(student["classroom_id"] == 2 for student in moved_students))
            enrollments = db.execute(
                """SELECT classroom_id FROM student_enrollments
                   WHERE school_year = ? AND student_id IN (?, ?)
                   ORDER BY student_id""",
                (current_year, moved_students[0]["id"], moved_students[1]["id"]),
            ).fetchall()
            self.assertEqual(len(enrollments), 2)
            self.assertTrue(all(row["classroom_id"] == 2 for row in enrollments))
            unaffected_student = db.execute(
                "SELECT classroom_id FROM students WHERE name = ?",
                ("Beatriz Ruiz",),
            ).fetchone()
            self.assertEqual(unaffected_student["classroom_id"], 3)

    def test_bulk_student_import_allows_repeated_person_codes(self):
        self.login("admin", "admin-password")
        with self.app.app_context():
            from app import get_db

            get_db().execute(
                """INSERT INTO students
                   (classroom_id, name, modality, sex, person_code)
                   VALUES (?, ?, ?, ?, ?)""",
                (1, "Existente", "Regular", "F", "PERSON-USED"),
            )
            get_db().commit()

        response = self.client.post(
            "/admin/students/bulk",
            data={
                "classroom_id": "1",
                "modality": "Regular",
                "student_name": ["Nuevo alumno", "Repetido"],
                "student_sex": ["M", "F"],
                "student_person_code": ["PERSON-USED", "PERSON-USED"],
            },
            follow_redirects=True,
        )

        self.assertIn(b"Se agregaron 2 estudiantes.", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM students WHERE person_code = ?",
                    ("PERSON-USED",),
                ).fetchone()[0],
                3,
            )

    def test_admin_can_delete_subject_and_related_grades(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")
        self.client.post("/logout")
        self.login("admin", "admin-password")

        response = self.client.post(
            "/admin/subjects/1/delete", follow_redirects=True
        )
        self.assertIn(b"junto con sus notas", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM subjects").fetchone()[0], 0)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM grades").fetchone()[0], 0)

    def test_teacher_cannot_delete_classroom_student_or_subject(self):
        self.login("maestra", "teacher-password")

        self.assertEqual(self.client.post("/admin/classrooms/1/delete").status_code, 403)
        self.assertEqual(self.client.post("/admin/students/1/delete").status_code, 403)
        self.assertEqual(self.client.post("/admin/subjects/1/delete").status_code, 403)

        with self.app.app_context():
            from app import get_db

            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM classrooms").fetchone()[0], 1)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM students").fetchone()[0], 1)
            self.assertEqual(get_db().execute("SELECT COUNT(*) FROM subjects").fetchone()[0], 1)

    def test_admin_cannot_delete_own_account_and_another_admin_can_be_removed(self):
        self.login("admin", "admin-password")
        response = self.client.post(
            "/admin/users/1/delete", follow_redirects=True
        )
        self.assertIn(b"No puedes eliminar la cuenta", response.data)
        response = self.client.post(
            "/admin/users",
            data={
                "username": "admin2",
                "password": "another-admin-password",
                "role": "admin",
            },
        )
        self.assertEqual(response.status_code, 302)
        with self.app.app_context():
            from app import get_db

            second_admin_id = get_db().execute(
                "SELECT id FROM users WHERE username = 'admin2'"
            ).fetchone()["id"]
        response = self.client.post(
            f"/admin/users/{second_admin_id}/delete", follow_redirects=True
        )
        self.assertIn(b"Se elimino la cuenta de admin2", response.data)
        with self.app.app_context():
            from app import get_db

            self.assertEqual(
                get_db().execute(
                    "SELECT COUNT(*) FROM users WHERE role = 'admin'"
                ).fetchone()[0],
                1,
            )

    def test_grade_list_groups_and_sorts_by_subject(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, score, term)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (1, 1, str(date.today().year), 2, 90, "Primer corte evaluativo"),
            )
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, score, term)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (1, 2, str(date.today().year), 2, 95, "Primer corte evaluativo"),
            )
            db.commit()
        self.login("admin", "admin-password")

        response = self.client.get("/grades")
        biology_position = response.data.index(b"<h2>Biologia</h2>")
        mathematics_position = response.data.index(b"<h2>Matematicas</h2>")
        self.assertLess(biology_position, mathematics_position)

    def test_existing_database_migration_preserves_teacher_name(self):
        legacy_database = os.path.join(self.temp_dir.name, "legacy.sqlite")
        db = sqlite3.connect(legacy_database)
        db.executescript(
            """
            CREATE TABLE users (
                id INTEGER PRIMARY KEY, username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL, role TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP
            );
            CREATE TABLE classrooms (
                id INTEGER PRIMARY KEY, grade TEXT NOT NULL, name TEXT NOT NULL,
                UNIQUE (grade, name)
            );
            CREATE TABLE students (
                id INTEGER PRIMARY KEY, classroom_id INTEGER NOT NULL, name TEXT NOT NULL
            );
            CREATE TABLE subjects (id INTEGER PRIMARY KEY, name TEXT NOT NULL UNIQUE);
            CREATE TABLE grades (
                id INTEGER PRIMARY KEY, student_id INTEGER NOT NULL,
                subject_id INTEGER NOT NULL, teacher_id INTEGER NOT NULL,
                score REAL NOT NULL CHECK (score >= 0 AND score <= 100),
                term TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (student_id, subject_id, term)
            );
            INSERT INTO users VALUES (1, 'maestra', 'hash', 'teacher', CURRENT_TIMESTAMP);
            INSERT INTO classrooms VALUES (1, '1ro', 'A');
            INSERT INTO students VALUES (1, 1, 'Ana Perez');
            INSERT INTO subjects VALUES (1, 'Matematicas');
            INSERT INTO grades VALUES
                (1, 1, 1, 1, 90, 'Primer corte evaluativo', CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
            """
        )
        db.close()
        migrated_app = create_app(
            {
                "TESTING": True,
                "DATABASE": legacy_database,
                "SECRET_KEY": "migration-test",
            }
        )
        with migrated_app.app_context():
            from app import get_db

            student = get_db().execute(
                "SELECT modality, sex, person_code FROM students WHERE id = 1"
            ).fetchone()
            self.assertEqual(student["modality"], "Regular")
            self.assertIsNone(student["sex"])
            self.assertIsNone(student["person_code"])
            self.assertIsNotNone(
                get_db().execute(
                    "SELECT id FROM modalities WHERE name = 'Regular'"
                ).fetchone()
            )
            row = get_db().execute(
                "SELECT teacher_id, teacher_name, score, school_year FROM grades WHERE id = 1"
            ).fetchone()
            self.assertEqual(row["teacher_id"], 1)
            self.assertEqual(row["teacher_name"], "maestra")
            self.assertEqual(row["score"], 90)
            self.assertEqual(row["school_year"], str(date.today().year))
            enrollment = get_db().execute(
                """SELECT e.school_year, c.grade
                   FROM student_enrollments e
                   JOIN classrooms c ON c.id = e.classroom_id
                   WHERE e.student_id = 1"""
            ).fetchone()
            self.assertEqual(enrollment["school_year"], str(date.today().year))
            self.assertEqual(enrollment["grade"], "1ro")
            get_db().execute("UPDATE grades SET score = ? WHERE id = 1", ("A+",))
            get_db().commit()
            self.assertEqual(
                get_db().execute(
                    "SELECT score FROM grades WHERE id = 1"
                ).fetchone()["score"],
                "A+",
            )

    def test_excel_export_numbers_students_by_class_with_females_first(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            students = [
                (1, "Zoe Mujer", "F", "PERSON-ZOE"),
                (1, "Ana Mujer", "F", "PERSON-ANA"),
                (1, "Aaron Hombre", "M", "PERSON-AARON"),
            ]
            for classroom_id, name, sex, person_code in students:
                db.execute(
                    """INSERT INTO students
                       (classroom_id, name, modality, sex, person_code)
                       VALUES (?, ?, ?, ?, ?)""",
                    (classroom_id, name, "Regular", sex, person_code),
                )
                student_id = db.execute(
                    "SELECT id FROM students WHERE person_code = ?", (person_code,)
                ).fetchone()["id"]
                db.execute(
                    """INSERT INTO grades
                       (student_id, subject_id, school_year, teacher_id, teacher_name,
                        score, term)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (
                        student_id,
                        1,
                        str(date.today().year),
                        2,
                        "maestra",
                        90,
                        "Primer corte evaluativo",
                    ),
                )

            db.execute("INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B"))
            second_classroom_id = db.execute(
                "SELECT id FROM classrooms WHERE grade = ? AND name = ?",
                ("2do", "B"),
            ).fetchone()["id"]
            db.execute(
                """INSERT INTO students
                   (classroom_id, name, modality, sex, person_code)
                   VALUES (?, ?, ?, ?, ?)""",
                (second_classroom_id, "Bruno Hombre", "Regular", "M", "PERSON-BRUNO"),
            )
            second_student_id = db.execute(
                "SELECT id FROM students WHERE person_code = ?", ("PERSON-BRUNO",)
            ).fetchone()["id"]
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (
                    second_student_id,
                    1,
                    str(date.today().year),
                    2,
                    "maestra",
                    90,
                    "Primer corte evaluativo",
                ),
            )
            db.commit()

        self.login("admin", "admin-password")
        response = self.client.get("/grades/export.xlsx")
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(
            [(row[0], row[4]) for row in rows[3:7]],
            [
                (1, "Ana Mujer"),
                (2, "Zoe Mujer"),
                (3, "Aaron Hombre"),
                (1, "Bruno Hombre"),
            ],
        )
        self.assertEqual(rows[3][5:7], ("F", "PERSON-ANA"))

    def test_excel_export_contains_recorded_grade(self):
        self.login("maestra", "teacher-password")
        self.add_grade("88.5")

        response = self.client.get("/grades/export.xlsx")
        self.assertEqual(response.status_code, 200)
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(
            rows[0][0],
            "INSTITUTO DR. MIGUEL JARQUIN VALLEJOS",
        )
        self.assertEqual(
            rows[1][0],
            "Corte evaluativo: Primer corte evaluativo",
        )
        self.assertEqual(
            rows[2][:8],
            (
                "N.º",
                "Modalidad",
                "Nivel",
                "Sección",
                "Nombre y apellido",
                "Sexo",
                "Código de persona",
                "Matematicas",
            ),
        )
        self.assertEqual(
            rows[3][:8],
            (
                1,
                "Regular",
                "1ro",
                "A",
                "Ana Perez",
                None,
                None,
                "88.5/AS",
            ),
        )

    def test_excel_export_summarizes_failing_grade_average_and_count(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.executemany(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                [
                    (1, "Aplazado uno"),
                    (1, "Aplazado dos"),
                    (1, "En el limite"),
                    (1, "Nota de texto"),
                    (1, "Nota AI"),
                ],
            )
            db.execute(
                "INSERT INTO classrooms (grade, name) VALUES (?, ?)", ("2do", "B")
            )
            second_classroom_id = db.execute(
                "SELECT id FROM classrooms WHERE grade = ? AND name = ?",
                ("2do", "B"),
            ).fetchone()["id"]
            db.executemany(
                "INSERT INTO students (classroom_id, name) VALUES (?, ?)",
                [
                    (second_classroom_id, "Otro grupo aplazado"),
                    (second_classroom_id, "Otro grupo AI"),
                ],
            )
            student_ids = {
                row["name"]: row["id"]
                for row in db.execute(
                    """SELECT id, name FROM students
                       WHERE name IN ('Ana Perez', 'Aplazado uno', 'Aplazado dos',
                                      'En el limite', 'Nota de texto', 'Nota AI',
                                      'Otro grupo aplazado', 'Otro grupo AI')"""
                ).fetchall()
            }
            scores = [
                (student_ids["Ana Perez"], 50),
                (student_ids["Aplazado uno"], 59),
                (student_ids["Aplazado dos"], 60),
                (student_ids["En el limite"], 60),
                (student_ids["Nota de texto"], "N/A"),
                (student_ids["Nota AI"], "AI"),
                (student_ids["Otro grupo aplazado"], 40),
                (student_ids["Otro grupo AI"], "AI"),
            ]
            db.execute(
                """INSERT INTO teacher_assignments
                   (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                (2, 1, second_classroom_id),
            )
            db.executemany(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                [
                    (
                        student_id,
                        1,
                        str(date.today().year),
                        2,
                        "maestra",
                        score,
                        "Primer corte evaluativo",
                    )
                    for student_id, score in scores
                ],
            )
            db.commit()

        self.login("maestra", "teacher-password")
        response = self.client.get("/grades/export.xlsx")
        workbook = load_workbook(io.BytesIO(response.data))
        sheet = workbook.active
        rows = list(sheet.values)
        summary_start_column = len(rows[2]) - 6
        summary_index = next(
            index
            for index, row in enumerate(rows)
            if row[summary_start_column] == "Resumen de aplazados (<60 o AI)"
        )
        student_rows = {
            sheet.cell(row=row_number, column=5).value: row_number
            for row_number in range(4, summary_index + 1)
        }
        failing_score_color = sheet.cell(
            row=student_rows["Aplazado uno"], column=8
        ).font.color
        self.assertIsNotNone(failing_score_color)
        self.assertEqual(failing_score_color.type, "rgb")
        self.assertEqual(failing_score_color.rgb, "FFFF0000")
        passing_score_color = sheet.cell(
            row=student_rows["Aplazado dos"], column=8
        ).font.color
        self.assertFalse(
            passing_score_color
            and passing_score_color.type == "rgb"
            and passing_score_color.rgb[-6:].upper() == "FF0000"
        )
        self.assertEqual(
            rows[summary_index + 1][summary_start_column:summary_start_column + 6],
            (
                "Año escolar",
                "Nivel",
                "Sección",
                "Materia",
                "Promedio (<60)",
                "Cantidad (<60 o AI)",
            ),
        )
        self.assertEqual(
            rows[summary_index + 2][summary_start_column:summary_start_column + 6],
            ("2026", "1ro", "A", "Matematicas", 54.5, 3),
        )
        self.assertEqual(
            rows[summary_index + 3][summary_start_column:summary_start_column + 6],
            ("2026", "2do", "B", "Matematicas", 40, 2),
        )

    def test_excel_export_uses_a_column_for_each_subject(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Ciencias",))
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Fis",))
            db.execute(
                """INSERT INTO teacher_assignments
                   (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                (2, 2, 1),
            )
            db.execute(
                """INSERT INTO teacher_assignments
                   (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                (2, 3, 1),
            )
            db.execute(
                """INSERT INTO teacher_assignments
                   (teacher_id, subject_id, classroom_id) VALUES (?, ?, ?)""",
                (2, 4, 1),
            )
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, school_year, teacher_id, teacher_name,
                    score, term)
                   VALUES (?, ?, ?, ?, ?, ?, ?)""",
                (1, 2, str(date.today().year), 2, "maestra", 92, "Primer corte evaluativo"),
            )
            db.commit()
        self.login("maestra", "teacher-password")
        self.add_grade("88.5")

        response = self.client.get("/grades/export.xlsx")
        workbook = load_workbook(io.BytesIO(response.data))
        sheet = workbook.active
        rows = list(sheet.values)
        self.assertEqual(
            rows[2][:11],
            (
                "N.º",
                "Modalidad",
                "Nivel",
                "Sección",
                "Nombre y apellido",
                "Sexo",
                "Código de persona",
                "Biologia",
                "Ciencias",
                "Fis",
                "Matematicas",
            ),
        )
        self.assertEqual(
            rows[3][:11],
            (
                1,
                "Regular",
                "1ro",
                "A",
                "Ana Perez",
                None,
                None,
                "92/AA",
                None,
                None,
                "88.5/AS",
            ),
        )
        self.assertEqual(sheet.column_dimensions["A"].width, 6)
        self.assertGreaterEqual(sheet.column_dimensions["J"].width, 10)

    def test_excel_export_assigns_letter_grade_by_numeric_range(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            for term, score in (
                ("A", 59),
                ("B", 60),
                ("C", 75),
                ("D", 76),
                ("E", 89),
                ("F", 90),
                ("G", 100),
            ):
                db.execute(
                    """INSERT INTO grades
                       (student_id, subject_id, school_year, teacher_id, teacher_name,
                        score, term)
                       VALUES (?, ?, ?, ?, ?, ?, ?)""",
                    (1, 1, str(date.today().year), 2, "maestra", score, term),
                )
            db.commit()

        self.login("maestra", "teacher-password")
        response = self.client.get("/grades/export.xlsx")
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        scores = [list(sheet.values)[3][7] for sheet in workbook.worksheets]
        self.assertEqual(
            scores,
            ["59/AI", "60/AF", "75/AF", "76/AS", "89/AS", "90/AA", "100/AA"],
        )


if __name__ == "__main__":
    unittest.main()
