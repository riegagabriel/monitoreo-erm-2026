# Plan: de la denuncia en PDF a la fila del Excel y al dashboard

Estado: borrador para revisión (1/10/2026). Nada de esto está construido todavía. Los datos nominales de los PDFs (nombres, DNI) no se transcriben en ningún archivo del proyecto.

## 1. Qué contienen los dos PDFs (lectura hecha el 1/10/2026)

| | Villa Kintiarina | Cheto |
|---|---|---|
| Archivo | `DENUNCIA VOTOS GOLONDRINOS - VILLA KINTIARINA.pdf` | `RIESGO ELECTORAL_DISTRIO CHETO_PROVINCIA CHACHAPOYAS JNE.pdf` |
| Páginas | 8: 2 con texto digital (7 y 8), 6 escaneadas (1 a 6) | 53: 5 con texto digital (1 a 5), 48 escaneadas (6 a 53) |
| Qué es | Denuncia penal anónima ante la Fiscalía Provincial Mixta de Pichari, por voto golondrino (art. 359, inc. 8, Código Penal). Sello de recibido del 30/09/2026 (hora ilegible) | Escrito de una ciudadana al Jurado Electoral Especial de Chachapoyas (25/09/2026), con anexos: solicitud de constatación a la municipalidad (15/09), oficio 297-2026-MDCH/A (23/09) y actas de constatación (16/09) |
| Territorio | Villa Kintiarina, La Convención, Cusco | Cheto, Chachapoyas, Amazonas |
| Ciudadanos | 281 (anexo nominal, páginas 1 a 6) | 43 (anexo nominal en texto digital) |
| Qué pide | Que se oficie a RENIEC con domicilios anteriores y actuales, fechas de cambio, sustento y verificaciones | Fiscalización y verificación de domicilios; poner el hecho en conocimiento de ONPE, JNE y RENIEC |
| Lectura de las actas | no aplica | Las páginas escaneadas que revisé son actas de «no fue posible constatar la presencia» |

Clasificación (corregida el 1/10 por Gabriel):
- **Kintiarina:** categoría 4 (denuncia o alerta de golondrinaje), canal «Ministerio Público», bandera de conflictividad. El texto no menciona violencia, solo «concertación ilícita».
- **Cheto:** es una denuncia que ingresó a RENIEC (no una alerta del JNE), categoría 2 (verificación o fiscalización). Por defecto la trato como bandera, igual que Kintiarina, porque ambas se presentaron como las denuncias para banderas rojas. **A confirmar.**
- **Alertas del JNE:** vendrán de una base de datos que aún no se ha compartido. Hasta entonces el marcador JNE solo se muestra con datos simulados.
- **Hora de recepción:** si el sello no es legible, el campo queda vacío y no se menciona.

## 2. Pipeline repetible (los 4 pasos que describiste)

```
1. Llega el PDF            → carpeta denuncias/ (el original no se modifica)
2. Extraer por página      → capa de texto si existe; si no, OCR con IA
3. Sistematizar con IA     → un JSON por denuncia, con esquema fijo
4. Fila en el Excel        → borrador; un humano valida; recién entonces alimenta el dashboard
```

### Paso 2: extracción mixta por página
- Regla: una página con ≥ 40 caracteres de texto digital se extrae con PyMuPDF (ya instalado, gratis, local). Una página sin texto va a OCR con IA.
- **No hace falta leer todo.** El dashboard necesita el cuerpo de la denuncia, no las tablas de ciudadanos ni las actas. Se OCR-ean solo las primeras páginas hasta que aparece la primera tabla; del anexo se toma únicamente el **conteo de filas** (última numeración visible, o DNI únicos). Con Cheto esto baja de 53 a unas 4 páginas, y evita enviar nombres y DNI a un servicio externo.
- Motor OCR con IA: Claude lee como imagen cuando son pocas páginas (Kintiarina, hoy). Para volumen, **Gemini con tu plan de Antigravity, sin API key, desde la terminal con la CLI de Antigravity (`agy`)**.
  - `agy` tiene modo no interactivo: `agy -p "<instrucción>" --model <slug> --output-format json --json-schema <archivo.json> --print-timeout 10m`. Usa las credenciales ya guardadas de tu sesión de Antigravity; si no hay sesión, falla con «authentication required» (código de salida 1).
  - El slug del modelo que muestra la documentación de `agy` es `gemini-3.8-flash-high`; va en la variable `AGY_MODEL` para poder cambiarlo. `agy models` lista los disponibles una vez iniciada la sesión.
  - En esta PC están el IDE y la app de escritorio de Antigravity (`%LOCALAPPDATA%\Programs\Antigravity`), pero **no el comando `agy`**: la CLI es un producto aparte y la app de escritorio no la instala. La instalación vive en `~/.gemini/antigravity`, que trae un `agentapi.bat` interno sin documentación; no lo uso. El instalador oficial de `agy` en Windows es un script de PowerShell de `antigravity.google` que deja un binario en `~/.local/bin`. **Requiere tu visto bueno para descargarlo y ejecutarlo.**
  - Alternativa semiautomática, sin instalar nada: una carpeta `denuncias/entrada/` y un archivo de instrucción fijo que pegas en la app de escritorio de Antigravity; el agente procesa lo que haya en la carpeta y deja el JSON. Funciona, pero es un paso manual por lote.
  - Cómo se automatiza: `scripts/ocr_agy.py` renderiza cada página escaneada a PNG con PyMuPDF en una carpeta de trabajo temporal, lanza `agy -p` con `--add-dir` solo sobre esa carpeta, pide el texto de las páginas (de 8 en 8) con un esquema JSON fijo (`pagina`, `texto`) y guarda el resultado. Como el agente necesita leer los archivos, se usa `--dangerously-skip-permissions` limitado a esa carpeta temporal, que no contiene nada más.
  - Privacidad: decidiste que no importa, porque son denuncias públicas, así que se envían todas las páginas escaneadas. El script trae el modo `--solo-cuerpo` por si algún día un documento sí es sensible.
- Salida: `texto/<archivo>__pNN.txt` por página (local, fuera de Git) y un `manifiesto` con el método usado (`digital` u `ocr`), caracteres y hash SHA-256 del PDF para no procesar dos veces el mismo archivo.

### Paso 3: sistematización con IA
Entrada: el texto del cuerpo. Salida: un JSON con este esquema (validado por el script; si falta un campo obligatorio, la fila queda como borrador incompleto):

| Campo | Qué es |
|---|---|
| `fecha_recepcion` | Fecha del sello o del documento, `dd/mm/aaaa`, más `hora` si se lee |
| `entidad_emisora` | Ministerio Público, JEE, ONPE, Defensoría, ciudadano, prensa, etc. |
| `tipo_documento` | Denuncia penal, escrito ciudadano, oficio, alerta, nota de prensa |
| `departamento`, `provincia`, `distrito` | Tal como figura, para cruzar con `distritos_lookup.parquet` |
| `n_ciudadanos` | Entero o vacío |
| `categoria_rea` | 1 a 5 de las categorías REA vigentes |
| `fuente_dashboard` | `conflictividad` (bandera), `jne` (marcador) u `otro_canal` |
| `conflictividad` | Sí/no con motivo en una frase |
| `resumen_publicable` | Máximo 280 caracteres, sin nombres de personas, redactado con la guía REA («presunto golondrinaje») |
| `pide_a_reniec` | Qué solicita, en una frase |
| `confianza` | Alta, media o baja, con las páginas dudosas |

Reglas del prompt: no copiar nombres ni DNI al JSON; el resumen no puede traer nombres de ciudadanos ni de funcionarios; si no está en el texto, el campo queda vacío (no inventar); citar la página de donde sale cada dato.

### Paso 4: Excel base
`BASE_DENUNCIAS_ERM_Q<fecha>.xlsx`, hoja `DENUNCIAS` (una fila por documento):
- Columnas públicas: `id`, `fecha_recepcion`, `hora`, `departamento`, `provincia`, `distrito`, `ubigeo_inei`, `fuente_dashboard`, `categoria_rea`, `conflictividad`, `resumen_publicable`, `n_ciudadanos`, `entidad_emisora`.
- Columnas internas (nunca van al dashboard): `archivo_origen`, `sha256`, `paginas_ocr`, `confianza`, `pide_a_reniec`, `revisado_por`, `notas`.
- Control: `estado` (`BORRADOR` / `VALIDADO` / `DESCARTADO`) y `simulado` (no/si). Solo `VALIDADO` y `simulado = no` cuentan como dato real.
- Ubigeo: se resuelve por la tripleta departamento-provincia-distrito contra el catálogo INEI; lo que no cuadre usa `equivalencias_territoriales.csv`. Nunca el ubigeo RENIEC.
- Antes de validar, el script corre el detector de nombres del REA (`depurar_nombres` y la lista local de frases prohibidas).

## 3. Datos simulados del bosquejo (la mezcla propuesta)

| Capa | Origen | Cantidad | Marca |
|---|---|---|---|
| Riesgo previo REA | Real: 8 denuncias con bandera, 67 distritos verificados, restituidos | 8, 67, 165 | `simulado = no`, atenuado |
| Denuncias nuevas del 1/10 | Real: Kintiarina y Cheto (ambas con bandera, a confirmar) | 2 | `simulado = no`, bandera del día |
| Alertas JNE | Simuladas, hasta que llegue la base del JNE | 3 a 4 | `simulado = si` |
| Banderas de conflictividad del día | Simuladas, ubicadas en distritos con riesgo previo y al azar | unas 10 | `simulado = si` |
| Reportes de orientadores y otros canales | Simulados, 3 cortes | como en el v3 | `simulado = si` |

Contrato de datos: un `data/eventos.json` con el mismo formato que luego servirá el endpoint de la hoja de Google; el dashboard solo cambia la URL de lectura. El banner del dashboard dirá cuántos eventos son reales y cuántos simulados.

## 4. Orden de trabajo propuesto
1. Confirmar la clasificación de las dos denuncias (sección 1).
2. Extraer las filas de Kintiarina y Cheto (Claude lee el cuerpo como imagen) y crear el Excel con 2 filas `BORRADOR`.
3. Escribir `scripts/denuncias_pipeline.py` (extracción mixta, manifiesto, validaciones) con pruebas.
4. Armar `data/eventos.json` mezclando reales y simulados.
5. Proyecto real `dashboard/` leyendo ese JSON.
6. Instalar `agy`, iniciar sesión y probar `scripts/ocr_agy.py` con las páginas escaneadas de Cheto. Es el único paso que depende de tu aprobación para descargar el instalador.

## 5. Pendientes que este plan no resuelve
- Aprobar la instalación de `agy` y confirmar que tu sesión de Antigravity sirve para la CLI.
- Confirmar si Cheto va como bandera y si Kintiarina es «conflictividad» aunque el texto no mencione violencia.
- Formato de la base de alertas del JNE cuando se comparta.
- Quién valida las filas el domingo.
