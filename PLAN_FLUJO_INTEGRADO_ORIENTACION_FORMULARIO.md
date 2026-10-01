# Plan: flujo integrado de orientación + formulario del monitor

Estado: borrador para revisión (1/10/2026). Nada construido. Fuentes: `FLUJO_ELECCIONES REGIONALES Y MUNICIPALES 2026.docx` (flujo de orientación), `PLAN_FORMULARIO_ERM.md` y `formulario/crear_formulario.gs` (formulario actual), `Guia operativa del orientador ERM 2026-v2.docx`.

## 1. Qué dice el flujo de orientación (resumen fiel)

| Cuándo | Qué ocurre | Responsable |
|---|---|---|
| **3/10** | Comunicación con orientadores: recordatorio y pautas finales | Equipo de monitoreo |
| 3/10 | Remite la confirmación de orientadores electorales con OS (orden de servicio) | Área administrativa |
| **4/10, 06:50** | Inicio de actividades del equipo de monitoreo | Monitoreo |
| 07:00 | Llegada de orientadores a locales de votación (28 distritos); presentación ante el coordinador local de ONPE | Orientadores |
| 07:05 a 07:20 | **Registro en el formulario** | Orientadores y monitoreo |
| 08:00 a 08:30 | Consolidación del primer reporte | Monitoreo y Base de datos |
| 08:00 a 16:45 | Reporte de consultas e incidencias | Monitoreo |
| durante el día | Consulta que requiere apoyo legal (p. ej. restitución o verificación de domicilio): el orientador informa a su monitor y el monitor operativo la traslada al Área Legal | Monitoreo y Área Legal |
| durante el día | Incidencia reportada por orientadores, ONPE o JNE: se registra y consolida para el **dashboard de monitoreo de conflictos** | Monitoreo, Supervisión, Base de datos |
| 16:30 a 16:45 | Los orientadores comunican a los coordinadores de ONPE que se retiran | Orientadores |
| **16:45** | **Ingresar al formulario el reporte final**: total y tipo de consultas, incidencias | Orientadores |
| 17:00 | Cierre de locales y consolidación del reporte final | Monitoreo y Base de datos |

## 2. Dónde el flujo choca con lo que ya construimos

1. **Quién llena el formulario.** El plan actual dice que solo el monitor registra lo que le reportan. El flujo dice que a las 07:05 y a las 16:45 **el orientador entra al formulario** (llegada y reporte final); entre ambos, el monitor registra consultas e incidencias.
2. **Cortes.** Nuestro formulario tiene tres cortes (inicio, mediodía, 16:30). El flujo tiene **dos momentos fijos de los orientadores** (llegada 07:05 a 07:20 y final 16:45) y un reporte continuo del monitor entre 08:00 y 16:45. No hay un corte de mediodía en el flujo.
3. **Consultas derivadas al Área Legal.** El flujo las contempla y el formulario no las registra.
4. **Fuentes de incidencia.** El flujo nombra orientadores, ONPE y JNE; el formulario tiene orientador y «otros canales», y el JNE entra aparte (32 alertas ya cargadas).
5. **Cobertura.** El flujo habla de **28 distritos**. El dashboard usa 380 «locales esperados» (valor provisional).

## 3. Propuesta de flujo integrado

```
3/10  Monitoreo: recordatorio y pautas ──► Administrativa: confirma orientadores con OS ──► lista de orientadores a ORIENTADORES
4/10  06:50 Monitoreo inicia
      07:05-07:20  ORIENTADOR llega y se registra   ── formulario, ruta «Llegada»
      08:00-08:30  Monitoreo + Base de datos consolidan el primer reporte
      08:00-16:45  MONITOR registra consultas e incidencias   ── ruta «Reporte del monitor»
                   ├─ consulta que exige apoyo legal ──► marcar «derivada al Área Legal»
                   ├─ incidencia (orientador, ONPE, JNE) ──► texto libre ──► Datos valida bandera ──► dashboard
                   └─ canal informal ──► ruta «Canal informal»
      16:30-16:45  Orientadores avisan a ONPE y se retiran
      16:45        ORIENTADOR envía reporte final       ── ruta «Cierre»
      17:00        Cierre y consolidación final
```

### Cambios al formulario (si se aprueba)
- **Primera pregunta nueva:** «¿Quién registra?» (orientador o monitor). El orientador ve solo Llegada y Cierre; el monitor ve el reporte en curso y el canal informal.
- **Ruta Llegada (orientador):** orientador (lista), local, hora de llegada, ¿coordinador de ONPE ubicado?, ¿espacio asignado? (opcionales). Cuenta como corte 1.
- **Ruta Cierre (orientador):** orientador, hora de término, consultas totales por tipo (los 5 de siempre) e incidente de texto libre. Cuenta como corte final.
- **Ruta Reporte del monitor:** consultas e incidencias acumuladas durante el día, con el campo nuevo **«Derivada al Área Legal»** (sí/no) y una nota corta de qué se derivó. Reemplaza al corte de mediodía como seguimiento intermedio.
- **Ruta Canal informal:** sin cambios.
- **Campos no categorizados:** se mantienen: incidente en texto libre; Datos valida bandera y resumen.

### Cambios al dashboard
- Fila de seguimiento: **llegada a tiempo** (registrada entre 07:05 y 07:20) y **reporte final enviado** por orientador.
- Cifra nueva: **consultas derivadas al Área Legal**, con conteo y total de consultas.
- Cobertura sobre **28 distritos** (lista de la administrativa) en vez de 380 locales; resaltar esos 28 distritos en el mapa.
- Las vistas «Todo el día» y por momento se redefinen: Llegada, Durante el día, Cierre.

## 4. Tareas

| # | Tarea | Cuándo | Responsable propuesto |
|---|---|---|---|
| 1 | Confirmar la lista de orientadores con OS y los 28 distritos | 3/10 | Área administrativa → Datos |
| 2 | Aprobar los cambios al formulario (sección 3) | 2/10 | Gabriel y Lucía |
| 3 | Actualizar `crear_formulario.gs` y la hoja (rutas, campo de Legal, `SEGUIMIENTO`) | 2/10 | Datos |
| 4 | Ensayo con respuestas de prueba de los dos perfiles | 3/10 | Datos y monitoreo |
| 5 | Ajustar el dashboard (28 distritos, Legal, llegada y cierre) | 3/10 | Datos |
| 6 | Mensaje del 3/10 a orientadores con el enlace y la hora de registro | 3/10 | Monitoreo |

## 5. Pendientes que este plan no resuelve
- ¿El orientador llena el formulario desde su celular (con qué cuenta o sin cuenta)? Un enlace abierto expone el formulario a envíos ajenos; la validación humana lo cubre, pero hay que decidir.
- ¿Se mantiene un corte de mediodía o lo cubre el reporte continuo del monitor?
- Lista de los 28 distritos y de orientadores con su local.
- Quién del Área Legal recibe lo derivado y por qué canal.
