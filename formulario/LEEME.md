# Formulario del monitor, ERM 2026

Crea el Google Forms con la ruta del monitor y una hoja de respuestas con seguimiento. Plan: `../PLAN_FORMULARIO_ERM.md`.

## Qué crea
- **Formulario** (un solo enlace para los monitores): monitor, departamento, distrito (25 secciones, una por departamento), «¿qué va a registrar?» y dos rutas: comunicación de un orientador (momento, local, horas, 5 consultas acumuladas, incidente en texto libre) o canal informal (canal, quién, fecha y hora, incidente, enlace).
- **Hoja de cálculo** vinculada, con estas pestañas:

| Pestaña | Para qué |
|---|---|
| `RESPUESTAS` | Lo que envía el Forms. A la derecha, 4 columnas amarillas para Datos: `bandera`, `menciona_reniec`, `resumen_publicable`, `validado` (SI/NO) |
| `ORIENTADORES` | Catálogo de orientadores. Vacía hasta que llegue la lista |
| `SEGUIMIENTO` | Quién ya reportó en inicio, mediodía y cierre; «pendiente» en ámbar |
| `CALCULADO` | Una fila por respuesta con los campos derivados (ubigeo, provincia, distrito...). No editar |
| `PUBLICO` | Solo filas con `validado = SI` y sin datos internos. Es la fuente del dashboard |

## Pasos (unos 5 minutos, con la cuenta que será dueña del formulario)
1. Entre a https://script.google.com y cree un proyecto nuevo (nombre sugerido: «Monitoreo ERM 2026»).
2. Pegue el contenido de `crear_formulario.gs` en el archivo `Código.gs`.
3. Cree un segundo archivo de script llamado `distritos` y pegue el contenido de `distritos.gs` (65 KB).
4. En **Configuración del proyecto**, marque «Mostrar el archivo de manifiesto appsscript.json» y compruebe que `timeZone` sea `America/Lima`.
5. Elija la función `crearFormulario` y pulse **Ejecutar**. Autorice los permisos de Formularios, Hojas y Propiedades.
6. En **Registro de ejecución** aparecen los enlaces: el del formulario (para los monitores), el de edición y el de la hoja.
7. Abra el enlace del formulario y envíe 3 respuestas de prueba: un orientador al inicio, un orientador al cierre con consultas y un canal informal. Compruebe `RESPUESTAS`, `SEGUIMIENTO` y que `CALCULADO` rellene ubigeo, provincia y distrito.
8. Borre las respuestas de prueba (en `RESPUESTAS`, filas 2 en adelante) antes del domingo.

## Cuando llegue la lista de orientadores
Pegue un orientador por fila en la columna A de `ORIENTADORES` (por ejemplo `APELLIDO NOMBRE | Local de votación`) y ejecute `actualizarOrientadores`. El desplegable del formulario se actualiza sin crear otro formulario ni perder respuestas.

## Reglas para Datos el domingo
- No ordene ni borre filas de `RESPUESTAS`: `CALCULADO` y `PUBLICO` se leen por posición de fila.
- Para publicar una fila: complete `resumen_publicable` (máximo 280 caracteres, sin nombres ni DNI), `bandera` y `menciona_reniec` (SI/NO) y ponga `validado = SI`.
- El texto original del incidente, el monitor y el orientador nunca aparecen en `PUBLICO`.

## Límites sin probar
Google Forms permite muchas secciones, pero no pude probar aquí el límite de 29 secciones con listas de hasta 171 opciones (Lima) ni los permisos de su cuenta. Si el script falla por tamaño, avíseme: el plan B es departamento en lista y provincia y distrito como texto.

## Regenerar los distritos
`distritos.gs` se genera con `python scripts/generar_distritos_apps_script.py` desde el catálogo INEI del proyecto REA. No editar a mano.
