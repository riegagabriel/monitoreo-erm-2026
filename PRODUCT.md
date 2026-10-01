# Product

<!-- impeccable:product-schema 1 -->

## Platform

web

## Stack
Boceto en HTML estático (bocetos/). Producto final: Vite + React 19 + MapLibre copiado de `DENUNCIAS_REA/sitio_web`, desplegado en Vercel como sitio estático que consulta cada 60 s un endpoint de Google Apps Script (decidido en el plan aprobado, SPEC_DASHBOARD_ERM.md).

## Users
Equipo de datos y de monitoreo de la SDPEG/DRE (RENIEC), la jefatura (Lucía, Manuel Chuquillanqui, Milagros Suito) e instituciones externas. Todos usan el mismo enlace abierto (noindex). Se mira en tres situaciones confirmadas: pantalla grande en la sala de monitoreo, laptops del equipo y celulares de la jefatura. Trabajo: seguir en vivo, durante el domingo 4 de octubre de 2026 (ERM 2026), qué ocurre en los locales de votación y qué le llega a RENIEC, para decidir a quién llamar o escalar.

## Product Purpose
Dashboard de monitoreo de la jornada electoral. Ordena en tres cortes (inicio de votación, mediodía, 16:30) y en un total del día los reportes que el equipo de monitoreo digita en un Google Forms a partir de los orientadores de RENIEC en los locales de votación (papel, WhatsApp, correo, llamada). Prioriza incidencias violentas o que mencionan a RENIEC. Éxito: un reporte validado aparece en el mapa en menos de 90 segundos, con territorio, gravedad y fuente correctos, y sin datos personales.

## Positioning
Es el único tablero que cruza en un mismo mapa lo que RENIEC ve en campo con lo que ya sabía antes del día (alertas REA, verificaciones domiciliarias, restituidos) y lo que otros canales y el JNE reportan ese mismo día, sin exponer datos personales.

## Operating Context
- Fecha: domingo 4/10/2026. Orientadores en el local de 08:15 a 17:00, con un monitor asignado cada uno.
- El orientador reporta con el Formato de reporte ERM (papel) u otros medios; el equipo de monitoreo lo digita en Google Forms y marca `validado`. Solo las filas validadas se publican.
- Tres fuentes que deben distinguirse: (1) monitoreo de orientadores, (2) otros canales del mismo día, (3) alertas del JNE (marker propio, con el logo del JNE).
- Otros canales confirmados: denuncias REA entrantes, entidades del sistema (ONPE, Defensoría, Fiscalía), prensa y redes sociales, llamadas o mesa de ayuda.
- Mapa a escala distrito, ubigeo INEI (nunca RENIEC). Precarga REA: solo las 8 denuncias con bandera, los 67 distritos con verificación domiciliaria y los restituidos (38 597 en 165 distritos).
- Consultas del formato en papel: DNI vencido, restitución de domicilio, «cara de niño», cierre de padrón, otros.

## Capabilities and Constraints
- Vistas: «Todo el día» (acumulado) y filtro por corte 1, 2 y 3. KPIs de totales y gravedad confirmados para «Todo el día»; consultas y cobertura quedan como lectura por corte (a decidir).
- Gravedad: 0 sin incidencia, 1 atención, 2 grave. Definición oficial pendiente.
- «Esperados» por corte provisionales; no hay lista de locales de votación ni asignación de orientadores.
- Cómo llegan las alertas del JNE: sin definir. El modelo las acepta como `fuente = JNE`.
- Sin datos personales: no se publican DNI, nombres de ciudadanos, de monitores ni de orientadores, ni el texto libre de la situación. Solo un «resumen publicable» revisado.
- Marca JNE: rojo y gris, similar al rojo de gravedad; la señal de JNE debe distinguirse por forma y logo, no por color.

## Brand Commitments
Identidad RENIEC (logo, teal y rojo institucionales) heredada del sitio REA. Logo JNE entregado en `MONITOREO_ERM_2026_7-10/logo/JNE.png` (PNG con fondo transparente, 1222x1024).

## Evidence on Hand
Datos reales precargados en `dashboard/public/data/precarga/` (REA, corte 30/09). Reportes de campo y alertas JNE aún no existen: solo datos simulados rotulados como tales. No inventar cifras reales del día.

## Product Principles
1. Lectura de un vistazo: lo grave y lo que menciona a RENIEC se ve primero, a varios metros y en el celular.
2. Cada fuente tiene identidad propia: monitoreo, otros canales y JNE nunca se confunden.
3. Lo público es seguro por construcción: nada con datos personales sale del Google Sheet privado.
4. Lo precargado es contexto, no ruido: se puede apagar sin perder lo del día.
5. Lo simulado se rotula siempre.

## Accessibility & Inclusion
Lectura a distancia en pantalla grande; el color nunca es el único canal (forma, ícono o rótulo acompañan); contraste AA; modo claro y oscuro; funcional en celular.
