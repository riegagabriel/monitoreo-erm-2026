# Plan rápido: Google Forms de monitoreo ERM 2026

Estado: borrador para revisión (1/10/2026). Nada construido. Base: `Formato de reporte_ ERM.docx` (papel del orientador) más lo que necesita el dashboard v3 (`bocetos/dashboard_erm_v3.html`) y el contrato de datos del plan principal.

## 1. Quién llena y qué registra
Lo llena el equipo de monitoreo (no el orientador), a partir del formato en papel, WhatsApp, correo o llamada. Un envío es **un reporte**. Hay dos tipos y el primero decide qué se pregunta después:

| Tipo de registro | Cuándo se usa | Alimenta |
|---|---|---|
| **A. Reporte de orientador** | Lo que reporta un orientador desde un local, en uno de los 3 cortes | Cifras por corte, consultas, cobertura, banderas |
| **B. Denuncia o alerta por otro canal** | Denuncia REA entrante, entidad (ONPE, Defensoría, Fiscalía), prensa y redes, llamada ciudadana | Tabla de otros canales, banderas |

Las alertas del JNE no pasan por este formulario por ahora: vendrán de una base aparte.

## 2. Qué trae el formato en papel y dónde va

| Campo del papel | Va al formulario | Pública |
|---|---|---|
| Nombre y apellidos del orientador | Sí, como «orientador o monitor que reporta» | No |
| DNI del orientador | **No se pide**: no lo usa el dashboard ni el equipo | |
| Local de votación | Sí, texto | Sí |
| Coordinador o responsable del local (ONPE) | Sí, texto opcional | No |
| Distrito, provincia, departamento | Sí, lista encadenada con ubigeo INEI | Sí |
| Hora de llegada y de término | Sí, dos horas | Hora de llegada sí; término no |
| Consultas: DNI vencido, restitución de domicilio, «cara de niño», cierre de padrón, otros (y total) | Sí, cinco números; el total lo calcula la hoja | Sí |
| Incidencia (texto libre) | Sí, como «situación» interna | No |

## 3. Lo que agrega el dashboard (no está en el papel)
Decisión de Gabriel (1/10): **el monitor no categoriza.** El monitor solo transcribe lo que le reporta el orientador, con un campo de texto libre para el incidente. Del formulario salen únicamente `corte`, `fuente de ingreso` y ese texto.

Lo que el dashboard necesita y el monitor ya no llena se completa **después, en la hoja, por el equipo de datos** (con apoyo de IA, ver sección 5):
- `bandera` (sí/no): incidencia de conflictividad.
- `menciona_reniec` (sí/no).
- `resumen_publicable`: máximo 280 caracteres, sin nombres.
- `validado` (sí/no).

No hay gravedad en el formulario. Si el dashboard la necesita, se deriva de `bandera`.

## 4. Estructura del formulario

**Sección 0, inicio**
1. Tipo de registro: A o B.
2. Quién registra (nombre del monitor): texto, interno.

**Territorio (común a A y B)**
3. Departamento: lista de 25. Cada opción salta a su propia sección con una lista de distritos «PROVINCIA · DISTRITO (ubigeo INEI)». Son 25 secciones con 40 a 130 opciones, en vez de una lista de 1 891 que Forms podría rechazar.

**Sección A, reporte de orientador**
4. Corte: 1 inicio de votación, 2 mediodía, 3 16:30 cierre del local.
5. Local de votación (texto) y coordinador (texto, opcional).
6. Fuente por la que llegó: formato en papel, WhatsApp, correo, llamada.
7. Hora de llegada y hora de término (esta, solo en el corte 3).
8. Consultas atendidas: cinco números enteros (≥ 0).
9. **Incidente (texto libre, opcional):** «Describa lo que reportó el orientador. Si no hubo incidencias, déjelo en blanco o escriba "Sin incidencias".» Sin listas ni categorías.

**Sección B, otros canales**
4. Canal: denuncia REA, entidad del sistema, prensa y redes, llamada o mesa de ayuda.
5. Entidad o medio y número de documento o enlace (texto, opcional).
6. Fecha y hora en que se recibió.
7. **Incidente o denuncia (texto libre).**

Reglas: nada de subir archivos; la ayuda del campo dice «Transcriba lo que le reportaron; no hace falta categorizar»; el mensaje de confirmación dice que el reporte entra al dashboard tras la validación del equipo de datos.

## 5. Del texto libre al dato del dashboard (en la hoja, no en el formulario)
Flujo: el monitor envía el texto → el equipo de datos revisa en la hoja y completa `bandera`, `menciona_reniec`, `resumen_publicable` y `validado`.
- **Apoyo de IA (opcional, mismo patrón del pipeline de PDFs):** un script con `agy -p` lee las filas nuevas con incidente y propone `bandera` (sí/no con motivo), `menciona_reniec` y un `resumen_publicable` de máximo 280 caracteres sin nombres. Son **propuestas**: el equipo decide y marca `validado`. Hasta que llegue esa automatización, se llenan a mano.
- Una fila sin incidente (campo vacío o «Sin incidencias») cuenta como reporte recibido y no lleva bandera.

## 6. Contrato hacia el dashboard
La hoja de respuestas es privada. Una pestaña `PUBLICO` (o el Apps Script) expone **solo** filas con `validado = SI` y estas columnas: `tipo`, `corte`, `hora`, `ubigeo_inei`, `departamento`, `provincia`, `distrito`, `local`, `bandera`, `menciona_reniec`, `resumen_publicable`, las cinco consultas y `canal`. Nunca salen: monitor, orientador, coordinador, texto original del incidente, documento, contacto. El dashboard lee ese JSON cada 60 s.

## 7. Control de calidad en la hoja
- Duplicados: aviso si hay otra fila con el mismo local, corte y tipo.
- Un resumen con 7 o más dígitos seguidos o con palabras de nombres conocidos queda sin validar (mismo detector que el REA).
- El ubigeo se toma del final de la opción de distrito, nunca se escribe a mano.

## 8. Orden de trabajo
1. Aprobar este plan.
2. Script de Apps Script que crea el formulario, sus 25 secciones y la hoja (generado desde `distritos_lookup.parquet`).
3. Ensayo con respuestas simuladas de ambos tipos y los 3 cortes.
4. Conectar el endpoint al dashboard. El dashboard v3 deja de mostrar «graves»: usa reportes, reportes con incidente, banderas y menciones a RENIEC.

## 9. Pendientes que este plan no resuelve
- Cuándo se hace la validación el domingo (quién del equipo de datos revisa las filas, cada cuánto) y si se automatiza con `agy`.
- Cuenta de Google dueña del formulario y la hoja.
- Si hay lista de locales de votación para convertir el campo «local» en una lista.
- Quién valida filas el domingo y si el formulario se cierra de madrugada.
- Si el orientador debe identificarse con nombre o con un código de puesto.
