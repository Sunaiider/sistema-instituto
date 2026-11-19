# Changelog

Todos los cambios relevantes del proyecto se documentan aquí.

## [Unreleased]

### Agregado
- Soporte para base de datos MySQL mediante `DATABASE_URL`, manteniendo SQLite como opción por defecto.
- Compatibilidad de SQL para ejecutar la misma aplicación sobre SQLite y MySQL con ajustes mínimos.
- Ejemplo de configuración `.env.example` y preparación para entorno local con Docker.
- Archivo `docker-compose.yml` para levantar MySQL y la aplicación localmente.
- Script `mysql_init.sql` con estructura inicial orientada a la gestión académica del instituto.
- Comandos rápidos en `makefile` para entorno virtual, instalación, arranque y tareas de Docker.
- Exportación de Excel con formato más cercano al diseño solicitado por la institución.

### Corregido
- Ajuste de la lógica de exportación para evitar errores al calcular anchos de columnas con celdas fusionadas.
- Corrección de la exportación de notas para mantener el formato y la estructura esperada por las pruebas.
- Validación del flujo de notas y permisos para maestros y administradores.

### Documentado
- README actualizado con instrucciones de instalación, uso con MySQL, Docker Compose y pruebas.
- Se documentaron ejemplos de conexión y variables de entorno para facilitar la puesta en marcha.

## [2026-11-16]

### Nuevo
- Expediente académico por estudiante con sus datos, notas e historial de grados/aulas por año escolar.
- Registro manual de matrículas y promociones individuales o para todo un grado/aula, con selección de grupo destino y año escolar.
- Las calificaciones se asocian al año escolar, permitiendo guardar los mismos cortes evaluativos en años diferentes.
- Las notas numéricas menores de 60 se resaltan en rojo en el Excel exportado.
- El resumen de aplazados aparece a la derecha de las notas en el Excel exportado.
- La importación de estudiantes admite `Nivel` y `Seccion` o `Sec.` para asignarlos a los grupos correspondientes.
- Las modalidades que aún no existen se crean al importar estudiantes desde Excel.
- El formulario de registro múltiple permite agregar un solo estudiante, sin un formulario separado.
- El mínimo de contraseña para las cuentas se redujo de 12 a 6 caracteres.

### Corregido
- Las notas existentes se asignan al año escolar actual durante la migración, ya que las versiones anteriores no guardaban ese dato.

## [2026-11-15]

### Agregado
- Interruptor accesible de modo oscuro con iconos de sol y luna, deslizador animado, preferencia guardada en el navegador y respeto a la configuración de movimiento reducido.
- Edición de cuentas de maestros y administradores, con cambio opcional de contraseña, rol y asignaciones docentes.
- Importación masiva de estudiantes desde Excel, con detección de encabezados, asignación por modalidad, nivel y sección, y creación automática de grupos faltantes.
- Plantilla descargable para importar estudiantes; el código de persona puede repetirse o quedar vacío.
- Resumen de aplazados por grado, aula y materia al final de cada pestaña del Excel; cuenta notas menores de 60 y `AI`, y promedia solo las notas numéricas menores de 60.
- Exportación general de notas con todas las materias del instituto para administradores y todas las materias asignadas para maestros, incluso sin calificaciones registradas.

### Corregido
- Eliminada la restricción de unicidad del código de persona para admitir los datos existentes en los archivos de matrícula.

## [2026-11-11]

### Agregado
- Las notas aceptan valores numéricos de 0 a 100 o texto de hasta 3 caracteres, como `A+` y `N/A`.
- El Excel general organiza las materias en columnas y agrupa las notas por alumno, año escolar y corte evaluativo.

### Corregido
- Al seleccionar el grado y aula para registrar notas, se cargan las materias asignadas al maestro.
- La base de datos y la visualización conservan notas de texto, además de las numéricas.

## [2026-10-28]

### Inicial
- Se inicia la versión de desarrollo del sistema de gestión de notas para instituto con autenticación, administración, asignaciones docentes, registro de calificaciones y exportación a Excel.
