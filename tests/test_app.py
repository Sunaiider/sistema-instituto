import io
import os
import sqlite3
import tempfile
import unittest
from openpyxl import load_workbook
from werkzeug.security import generate_password_hash

from app import create_app


class GradeSystemTests(unittest.TestCase):
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
                "term": "Primer trimestre",
                "score_1": score,
            },
            follow_redirects=True,
        )

    def test_teacher_can_record_grade_but_cannot_change_or_delete_it(self):
        self.login("maestra", "teacher-password")
        self.add_grade("85")

        response = self.client.post(
            "/grades/new",
            data={
                "classroom_id": "1",
                "subject_id": "1",
                "term": "Primer trimestre",
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
                "term": "Primer trimestre",
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
                "term": "Segundo trimestre",
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
                "term": "Primer trimestre",
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
        self.assertEqual(rows[1][5], 93)
        self.assertEqual(rows[1][6], "admin")

    def test_teacher_cannot_access_or_export_unassigned_combination(self):
        with self.app.app_context():
            from app import get_db

            db = get_db()
            db.execute("INSERT INTO subjects (name) VALUES (?)", ("Biologia",))
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, teacher_id, teacher_name, score, term)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (1, 2, 1, "admin", 81, "Primer trimestre"),
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
        self.assertEqual(self.client.get("/grades/new?classroom_id=2").status_code, 200)
        self.assertEqual(
            self.client.get("/grades/new?classroom_id=2&subject_id=2").status_code,
            403,
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
                   (student_id, subject_id, teacher_id, teacher_name, score, term)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (1, 2, 2, "maestra", 96, "Primer trimestre"),
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
                   (student_id, subject_id, teacher_id, teacher_name, score, term)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (2, 1, 2, "maestra", 78, "Primer trimestre"),
            )
            db.commit()

        response = self.client.get("/grades/export/1/1.xlsx")
        self.assertEqual(response.status_code, 200)
        self.assertIn("notas_Matematicas_1ro_A.xlsx", response.headers["Content-Disposition"])
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1][0:6], ("Ana Perez", "1ro", "A", "Matematicas", "Primer trimestre", 85))

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
                   (student_id, subject_id, teacher_id, teacher_name, score, term)
                   VALUES (?, ?, ?, ?, ?, ?)""",
                (1, 1, 1, "admin", 70, "Segundo trimestre"),
            )
            db.commit()
        self.login("maestra", "teacher-password")
        self.add_grade("85")

        response = self.client.get("/grades/export/1/1.xlsx")
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(len(rows), 3)
        exported_period_scores = {(row[4], row[5]) for row in rows[1:]}
        self.assertEqual(
            exported_period_scores,
            {("Primer trimestre", 85), ("Segundo trimestre", 70)},
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
        self.assertIn("term=Primer+trimestre", edit_response.headers["Location"])

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
                   (student_id, subject_id, teacher_id, score, term)
                   VALUES (?, ?, ?, ?, ?)""",
                (1, 1, 2, 90, "Primer trimestre"),
            )
            db.execute(
                """INSERT INTO grades
                   (student_id, subject_id, teacher_id, score, term)
                   VALUES (?, ?, ?, ?, ?)""",
                (1, 2, 2, 95, "Primer trimestre"),
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
                score REAL NOT NULL, term TEXT NOT NULL,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (student_id, subject_id, term)
            );
            INSERT INTO users VALUES (1, 'maestra', 'hash', 'teacher', CURRENT_TIMESTAMP);
            INSERT INTO classrooms VALUES (1, '1ro', 'A');
            INSERT INTO students VALUES (1, 1, 'Ana Perez');
            INSERT INTO subjects VALUES (1, 'Matematicas');
            INSERT INTO grades VALUES
                (1, 1, 1, 1, 90, 'Primer trimestre', CURRENT_TIMESTAMP, CURRENT_TIMESTAMP);
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

            row = get_db().execute(
                "SELECT teacher_id, teacher_name, score FROM grades WHERE id = 1"
            ).fetchone()
            self.assertEqual(row["teacher_id"], 1)
            self.assertEqual(row["teacher_name"], "maestra")
            self.assertEqual(row["score"], 90)

    def test_excel_export_contains_recorded_grade(self):
        self.login("maestra", "teacher-password")
        self.add_grade("88.5")

        response = self.client.get("/grades/export.xlsx")
        self.assertEqual(response.status_code, 200)
        workbook = load_workbook(io.BytesIO(response.data), read_only=True)
        rows = list(workbook.active.values)
        self.assertEqual(rows[0][0:6], ("Alumno", "Grado", "Aula", "Materia", "Periodo", "Nota"))
        self.assertEqual(rows[1][0:6], ("Ana Perez", "1ro", "A", "Matematicas", "Primer trimestre", 88.5))


if __name__ == "__main__":
    unittest.main()
