# Sistema de notas para instituto

Aplicacion web en espanol para que los maestros registren calificaciones por grado, aula, materia y periodo. Una vez guardada, una nota no puede ser editada ni eliminada por el maestro. Solo un administrador puede hacerlo. Los registros se pueden descargar en formato Excel (`.xlsx`).

## Funciones

- Inicio de sesion con cuentas de administrador y maestro.
- Administrador: crear cuentas, grados/aulas, alumnos y materias.
- Al crear una cuenta de maestro, el administrador selecciona una o varias combinaciones de materia y grado/aula; puede cambiarlas despues desde Administracion.
- Maestro: elegir grado, aula, materia y periodo, y capturar notas de 0 a 100.
- El maestro solo puede capturar, consultar y exportar notas de las combinaciones que tiene asignadas. Puede consultar y descargar tambien las notas de esas combinaciones ingresadas por el administrador.
- Una combinacion de alumno, materia y periodo solo se puede registrar una vez.
- Administrador: agregar y editar notas desde la lista completa de alumnos de un grado y aula; puede eliminar notas individuales.
- El listado de notas primero muestra solo materia, grado/aula y cantidades; al abrir "Ver notas" aparecen los alumnos, calificaciones y acciones individuales.
- El resumen se agrupa por materia y se ordena por grado y aula.
- Se puede descargar un archivo Excel independiente para cada materia y grado/aula, desde el resumen o su detalle.
- Administrador: eliminar grados/aulas, alumnos y materias. Al eliminar cualquiera de estos elementos, tambien se borran las notas relacionadas.
- Administrador: crear y eliminar cuentas. Al eliminar una cuenta, se conserva el historial y el nombre del maestro en sus notas.
- El maestro no puede editar ni eliminar calificaciones; su consulta y exportacion se limita a las combinaciones que tiene asignadas.
- Descarga de las notas visibles como archivo Excel.
- Base de datos SQLite local en `instance/instituto.sqlite`.

## Ejecutar en Windows (PowerShell)

Requiere Python 3.10 o posterior.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:SECRET_KEY = (python -c "import secrets; print(secrets.token_hex(32))")
python -m flask --app app init-admin
python -m flask --app app run
```

Abre `http://127.0.0.1:5000`. El comando `init-admin` pedira el usuario y una contrasena de al menos 12 caracteres. Desde la cuenta administradora se crean las cuentas para los maestros.

Para un despliegue, configura `SECRET_KEY` con un secreto persistente y usa un servidor WSGI de produccion. Si la aplicacion se sirve exclusivamente mediante HTTPS, configura tambien `COOKIE_SECURE=true`. No expongas el servidor de desarrollo de Flask a Internet.

## Pruebas

```powershell
python -m unittest discover -s tests -v
```


Van por cortes y no periodos/trimestre