# Sistema de notas para instituto

Aplicacion web en espanol para que los maestros registren calificaciones por grado, aula, materia y corte evaluativo. El ciclo usa cuatro cortes, no trimestres. Una vez guardada, una nota no puede ser editada ni eliminada por el maestro. Solo un administrador puede hacerlo. Los registros se pueden descargar en formato Excel (`.xlsx`).

## Funciones

- Interruptor animado para alternar entre modo claro y oscuro; la preferencia se conserva en el navegador.
- Inicio de sesion con cuentas de administrador y maestro.
- Administrador: editar cuentas de maestros y administradores, cambiar usuario y rol, restablecer contraseña y actualizar las asignaciones docentes. La contraseña puede dejarse vacía para conservarla.
- Administrador: crear cuentas, grados/aulas y materias; agregar modalidades y asignarlas a estudiantes desde una lista desplegable. Los estudiantes se pueden registrar desde una tabla, incluso uno solo, o importar desde Excel. Los archivos con columnas Modalidad, Nivel y Sec. asignan cada estudiante a su grupo y crean los que falten; los archivos sin esos campos usan el grado, aula y modalidad seleccionados.
- Cada estudiante tiene un expediente con sus datos, notas e historial de matrículas por año escolar y grado. El administrador registra manualmente cada promoción o cambio de grupo; el año escolar actual se propone por defecto. Las notas guardadas antes de esta actualización se asignan al año actual, porque no tenían un año escolar asociado.
- El administrador también puede registrar o promover en una sola acción a todos los estudiantes de un grado/aula, indicando el grupo destino y el año escolar.
- Al crear una cuenta de maestro, el administrador selecciona una o varias combinaciones de materia y grado/aula; puede cambiarlas despues desde Administracion.
- Maestro: elegir grado, aula, materia y uno de los cuatro cortes evaluativos (primero a cuarto), y capturar notas numericas de 0 a 100 o notas de texto de hasta 3 caracteres (por ejemplo, `A+` o `N/A`).
- El maestro solo puede capturar, consultar y exportar notas de las combinaciones que tiene asignadas. Puede consultar y descargar tambien las notas de esas combinaciones ingresadas por el administrador.
- Una combinación de alumno, materia, año escolar y corte evaluativo solo se puede registrar una vez.
- Administrador: agregar y editar notas desde la lista completa de alumnos de un grado y aula; puede eliminar notas individuales.
- El listado de notas primero muestra solo materia, grado/aula y cantidades; al abrir "Ver notas" aparecen los alumnos, calificaciones y acciones individuales.
- El resumen se agrupa por materia y se ordena por grado y aula.
- Se puede descargar un archivo Excel independiente para cada materia y grado/aula, desde el resumen o su detalle.
- Administrador: eliminar grados/aulas, alumnos y materias. Al eliminar cualquiera de estos elementos, tambien se borran las notas relacionadas.
- Administrador: crear y eliminar cuentas. Al eliminar una cuenta, se conserva el historial y el nombre del maestro en sus notas.
- El maestro no puede editar ni eliminar calificaciones; su consulta y exportacion se limita a las combinaciones que tiene asignadas.
- En el Excel general aparecen todas las materias del instituto para administradores y todas las materias asignadas al maestro, aunque todavía no tengan notas; las celdas sin calificación quedan vacías.
- Descarga de las notas visibles como archivo Excel, con una pestaña por corte, el nombre del instituto en mayúsculas y el corte encima de la tabla; incluye numeración, modalidad, nivel, sección, nombre, sexo y código de persona, y ordena primero a las mujeres y luego a los varones. Al final de cada pestaña incluye por grado, aula y materia el promedio de notas numéricas menores de 60 y la cantidad de notas menores de 60 o con `AI`; `AI` se suma a la cantidad, pero no al promedio.
- Base de datos SQLite local por defecto en `instance/instituto.sqlite`. Tambien puede usarse MySQL si defines `DATABASE_URL` con una cadena como `mysql+pymysql://usuario:password@localhost:3306/instituto`.

## Inicio rapido

Requiere Python 3.10 o posterior.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
$env:SECRET_KEY = (python -c "import secrets; print(secrets.token_hex(32))")
python -m flask --app app init-admin
python -m flask --app app run
```

Abre `http://127.0.0.1:5000`. El comando `init-admin` pedira el usuario y una contrasena de al menos 6 caracteres. Desde la cuenta administradora se crean las cuentas para los maestros.

## MySQL

Define `DATABASE_URL` antes de iniciar la app, por ejemplo:

```powershell
$env:DATABASE_URL = "mysql+pymysql://usuario:password@localhost:3306/instituto"
python -m flask --app app init-admin
python -m flask --app app run
```

La base de datos MySQL debe existir antes de ejecutar la app. Tambien puedes usar el ejemplo de entorno en `.env.example` y el archivo `docker-compose.yml` para levantar MySQL y la aplicacion.

## PostgreSQL

La aplicacion tambien admite PostgreSQL. Define `DATABASE_URL` con la URL que proporciona tu servicio de base de datos:

```powershell
$env:DATABASE_URL = "postgresql://usuario:password@host:5432/instituto?sslmode=require"
python -m flask --app app init-admin
```

En un servicio de alojamiento, configura `DATABASE_URL` y `SECRET_KEY` como variables de entorno. La aplicacion crea las tablas necesarias al iniciar; crea el administrador una sola vez con `init-admin`. La base PostgreSQL debe estar vacia: cambiar desde SQLite o MySQL no transfiere los datos existentes. Para desplegar el contenedor se usa Gunicorn; el servidor de desarrollo de Flask queda reservado para uso local.

### Despliegue en Wasmer

Los archivos `app.yaml` y `wasmer.toml` configuran la aplicacion Flask y una base PostgreSQL administrada por Wasmer. Wasmer solo ofrece bases de datos en la region `fr-roub1` (Roubaix, Francia), por lo que la aplicacion se fija en esa misma region. Los procesos de la app son efimeros; las tablas y los datos se guardan en PostgreSQL, no en el disco local.

1. Reemplaza `tu-namespace` por tu namespace de Wasmer en `app.yaml` y `wasmer.toml`.
2. Instala el CLI de Wasmer e inicia sesion con `wasmer login`. Prepara las dependencias WASI desde PowerShell y despliega desde la raiz del proyecto:

   ```powershell
   python -m pip install --target .wasmer/site-packages -r requirements-wasmer.txt
   wasmer deploy
   ```

   La primera publicacion crea la base PostgreSQL. Revisa el plan y sus limites en Wasmer antes de usarla.
3. En el panel de la app, agrega `SECRET_KEY` como variable secreta y redepliega. Generala localmente con `python -c "import secrets; print(secrets.token_hex(32))"`; no la guardes en el repositorio.
4. Obtén las credenciales de **Databases** en el panel de Wasmer y ejecuta una sola vez el comando de inicializacion desde PowerShell, sustituyendo los valores por los de tu base:

   ```powershell
   $env:APP_DATABASE_ENGINE = "postgresql"
   $env:DB_HOST = "<host>"
   $env:DB_PORT = "<puerto>"
   $env:DB_NAME = "<base>"
   $env:DB_USERNAME = "<usuario>"
   $env:DB_PASSWORD = "<contrasena>"
   python -m flask --app app init-admin
   ```

   No compartas ni guardes esas credenciales en el repositorio.

Las bases administradas de Wasmer tienen respaldos automaticos con al menos 14 dias de retencion. La region disponible para la base es Francia. El plan gratuito Hobby esta destinado a proyectos personales no comerciales; verifica las condiciones del plan Pro para uso institucional. Revisa tambien con el instituto si almacenar datos estudiantiles en Francia cumple sus requisitos de privacidad.

## Docker Compose

```powershell
docker compose up --build
```

Luego crea el administrador:

```powershell
docker compose exec app python -m flask --app app init-admin
```

Para un despliegue, configura `SECRET_KEY` con un secreto persistente y usa un servidor WSGI de produccion. Si la aplicacion se sirve exclusivamente mediante HTTPS, configura tambien `COOKIE_SECURE=true`. No expongas el servidor de desarrollo de Flask a Internet.

## Comandos rapidos

```powershell
make venv
make install
make run
make init-admin
make mysql-up
make app-up
make all-up
make down
make clean
```

## Pruebas

```powershell
python -m unittest discover -s tests -v
```

Para importar estudiantes, descarga la plantilla `.xlsx` desde la sección de estudiantes y completa las columnas `Nombre`, `Sexo` (`F` o `M`) y `Código de persona` (opcional, puede repetirse). Si el archivo incluye `Modalidad`, las modalidades nuevas se crearán durante la importación; si no, selecciona una en el formulario. Para asignar cada fila a su grupo, incluye `Nivel` y `Seccion` (o `Sec.`); los grados y aulas que no existan se crearán. Si el archivo no incluye esas columnas de grupo, selecciona un grado y aula antes de importar.
