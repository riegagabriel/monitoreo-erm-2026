# Plan — Dashboard de monitoreo ERM 2026 (jornada del domingo 4/10/2026)

## Context
El domingo 4/10 los orientadores de RENIEC estarán en locales de votación y reportarán (papel, WhatsApp, correo, llamada) a un equipo de monitoreo, que los digita en un Google Forms en tres cortes (inicio, mediodía, 16:30). Hace falta un dashboard en vivo con mapa para equipo de datos, jefatura (Lucía, Manuel, Milagros) e instituciones externas: **un solo enlace abierto, `noindex`, solo datos sin información personal** (decidido).
Se reutiliza el sitio REA (`DENUNCIAS_REA/sitio_web`, Vite + React 19 + MapLibre) **sin modificarlo**: se copia a un proyecto nuevo.

## Decisiones tomadas
- Audiencia: equipo + jefatura + instituciones externas → todo lo publicado es público-seguro.
- Acceso: enlace abierto, noindex (como `rea-denuncias.vercel.app`).
- Mapa a **escala distrito** (no hay lista de locales de votación ni asignación de orientadores en el proyecto; el local queda como texto).
- **Precarga desde el REA, solo:** geometría, contexto REA (solo las denuncias con bandera = 8, distritos con verificación domiciliaria = 67, restituidos = 38 597 en 165 distritos) e identidad (logo, tokens).
- **Alertas del JNE:** marker propio en el mapa (no bandera). **Cómo llegan está sin definir** → el modelo de datos las acepta como `fuente = JNE` y el adaptador de ingreso se decide después.

## Arquitectura (enfoque 1: Hoja de Google → Apps Script → sitio estático)
```
Google Forms (monitoreo, 3 cortes) → Hoja privada (REPORTES)
   └─ Apps Script (doGet) → JSON solo con filas validado=SI y campos públicos
        └─ Sitio estático en Vercel, consulta cada 60 s → mapa + cifras + feed
Precarga (una vez, Python): reutiliza scripts REA → public/data/precarga/*.json
```
Fallback si Apps Script falla: pestaña PUBLICO (fórmula QUERY, solo columnas públicas) publicada como CSV; el sitio lee CSV (cache de Google ≈ 5 min).

## Esquema de la hoja REPORTES (una fila por reporte)
Públicos: `corte` (1/2/3), `hora`, `departamento`, `provincia`, `distrito` (opción «NOMBRE · ubigeo INEI»), `local_votacion`, `fuente` (Orientador · WhatsApp · Correo · Llamada · JNE · Otro), `gravedad` (0 sin incidencia · 1 atención · 2 grave), `menciona_reniec` (SI/NO), `resumen_publicable`, consultas `dni_vencido`, `restitucion`, `cara_nino`, `cierre_padron`, `otros`.
Internos (nunca salen en el JSON): `registra` (monitor), `situacion` (texto libre), `observaciones`, `contacto`.
Control: `validado` (SI/NO). Solo `SI` se publica. Un nombre de persona en `resumen_publicable` bloquea la publicación (reutilizar `depurar_nombres` y la lista local de frases prohibidas de REA en la precarga y como regla en Apps Script).
Territorio: listas encadenadas dep > prov > dist con ubigeo INEI (generadas desde `data/geodata/distritos_lookup.parquet`); plan B del spec REA: modo texto si Forms no soporta 1 891 opciones. Ubigeo INEI, nunca RENIEC.

## Estética (un canal visual por pregunta; paleta propia, distinta de las 5 categorías REA)
| Pregunta | Canal |
|---|---|
| Gravedad del reporte | Relleno del distrito: normal / atención / grave |
| Menciona a RENIEC | Anillo o icono sobre el distrito |
| Alerta del JNE | **Marker JNE** (insignia con rótulo «JNE»; patrón de `mapas/banderas.ts`) |
| Distrito sin reporte en el corte | Hueco (sin relleno) |
| Reporte reciente (<30 min) | Pulso suave |
| Riesgo previo (REA) | Bandera roja (8 denuncias con alerta) + contorno rayado en distritos verificados + gradiente tenue de restituidos; capa apagable |

Pantalla: cabecera con hora del último corte → cifras (locales reportando, incidencias, graves, mencionan RENIEC, consultas por tipo) → mapa con panel lateral + selector de corte (con barra recibidos/esperados) → feed cronológico de incidencias. Móvil primero, modo claro/oscuro heredado.

## Archivos
Nuevos, en `D:\RENIEC_09_09_2026\MONITOREO_ERM_2026_7-10\`:
- `SPEC_DASHBOARD_ERM.md` (este plan consolidado)
- `formulario/crear_formulario.gs` + `LEEME.md` (crea el Form y la hoja; encadenado de territorio)
- `formulario/endpoint.gs` (doGet con solo columnas públicas y validado=SI)
- `scripts/precarga.py` (+ `tests/`): lee `DENUNCIAS_REA/sitio_web/public/data/{casos,distritos,*.geojson}` y escribe `dashboard/public/data/precarga/`: `riesgo_previo.json` (solo `alerta=true`, verificación, restituidos por distrito), geometría y logo. Reutiliza `DENUNCIAS_REA/scripts/sitio_datos.py` (`unir_restituidos`, guardas) en modo lectura.
- `scripts/simulador.py`: genera reportes simulados para el ensayo (incluye filas `fuente=JNE`).
- `dashboard/` (copia de `sitio_web`): se reutilizan `estilos.css`, `mapas/{Mapas,controlador,capas,estiloBase,Leyendas}`, `SelectorFondo`, `Cifras`, `Encabezado`, `Pie`, `hooks/useDatos.ts`, `agregados.ts`. Nuevos: `SelectorCorte`, `FeedIncidencias`, `mapas/markerJNE.ts`, `datos.ts` con polling a 60 s, `tipos.ts` (Reporte, FuenteJNE).
- `bocetos/dashboard_erm_v1.html`: boceto con datos reales precargados + simulados, antes de codificar.
- `GUIA_JORNADA_4_OCT.md`: roles, qué es gravedad 1/2, qué hacer si algo falla.
No se toca `DENUNCIAS_REA/` (sitio, Streamlit ni repo público).

## Fases
1. **Jue 1/10:** boceto → aprobación estética; formulario + hoja + endpoint; `precarga.py`.
2. **Vie 2/10:** `dashboard/` v1 (mapa con capas, cortes, feed, marker JNE), repo privado + proyecto Vercel nuevo (commits con `riega.gabriel@pucp.edu.pe`).
3. **Sáb 3/10:** ensayo de punta a punta con el simulador y 3 cortes; prueba en celular; congelar.
4. **Dom 4/10:** operación; cada cambio posterior solo en datos, no en código.

## Verificación
- Pruebas `unittest` de `precarga.py` (solo 8 alertas, 67 verificados, 165 distritos restituidos, sin PII).
- `vitest` de agregados por corte/gravedad; build (`tsc -b && vite build`) limpio.
- Ensayo: cargar reportes simulados al Form real y confirmar que aparecen en ≤ 90 s con territorio, gravedad y marker JNE correctos, que una fila `validado=NO` no sale, y que un nombre en `resumen_publicable` se bloquea.
- Navegador (preview): 375×812 y escritorio, claro/oscuro, mapa con las 4 capas y selector de corte.
- Inspección del JSON público: ninguna columna interna.

## Pendientes abiertos (no bloquean el boceto)
1. **Cómo llegan las alertas del JNE** (mismo Form / pestaña aparte / archivo) y su esquema final.
2. Cuenta de Google dueña del Form y la hoja (Workspace institucional o personal).
3. Quién valida filas (`validado`) el domingo y si el Form se cierra de madrugada.
4. Lista de locales de votación (ONPE), si se consigue → puntos por local.
5. Definición oficial de gravedad 1/2 y confirmar «esperados» por corte (cuántos locales/orientadores).
