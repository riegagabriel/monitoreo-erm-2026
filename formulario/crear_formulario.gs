/**
 * Monitoreo ERM 2026: crea el Google Forms del monitor y su hoja de respuestas.
 *
 * Pegar este archivo y `distritos.gs` en un proyecto de script.google.com, ejecutar `crearFormulario`
 * y autorizar. Ver LEEME.md. Plan: PLAN_FORMULARIO_ERM.md y el plan del 1/10/2026.
 *
 * Flujo del monitor (Google Forms solo ramifica «ir a la sección según la respuesta», por eso el
 * territorio va primero y la ruta después):
 *   Sección 1   monitor + departamento
 *   Secciones   una por departamento, con solo sus distritos «PROVINCIA · DISTRITO (ubigeo INEI)»
 *   Sección 27  ¿qué va a registrar?  → comunicación de un orientador | canal informal
 *   Ruta A      orientador, momento, local, horas, consultas acumuladas, incidente (texto libre)
 *   Ruta B      canal, quién envía, fecha y hora, incidente (texto libre), enlace
 *
 * El monitor NO categoriza. La bandera, «menciona a RENIEC», el resumen publicable y la validación
 * los completa el equipo de Datos en las últimas columnas de RESPUESTAS.
 */

const CONFIG = {
  TITULO: 'Monitoreo ERM 2026 · Registro del monitor',
  ZONA: 'America/Lima',
  // Lista de orientadores tal como aparecerá en el desplegable, p. ej. "APELLIDO NOMBRE | Local de votación".
  // Vacía por ahora: se completa en la pestaña ORIENTADORES y se aplica con actualizarOrientadores().
  ORIENTADORES: [],
  TIPOS: ['Comunicación de un orientador', 'Canal informal'],
  MOMENTOS: ['Inicio de votación', 'Mediodía', 'Cierre 16:30', 'Incidente fuera de corte'],
  CANALES: ['WhatsApp', 'Facebook o redes', 'Correo', 'Llamada', 'Denuncia presentada', 'Prensa', 'Otro'],
  CONSULTAS: ['DNI vencido', 'Restitución de domicilio', '«Cara de niño»', 'Cierre de padrón electoral', 'Otros'],
  PLACEHOLDER_ORIENTADOR: '(lista de orientadores pendiente)',
  FILAS_SEGUIMIENTO: 200,
};

// Títulos de las preguntas: son los encabezados de RESPUESTAS y de ellos salen las columnas.
const T = {
  monitor: 'Nombre del monitor que registra',
  dep: 'Departamento',
  tipo: '¿Qué va a registrar?',
  orientador: 'Orientador',
  momento: 'Momento del reporte',
  local: 'Local de votación',
  llegada: 'Hora de llegada al local',
  termino: 'Hora de término de la orientación',
  incidenteOri: 'Incidente reportado por el orientador',
  canal: 'Canal por el que llegó',
  quien: 'Quién lo envía o medio',
  fechaCanal: 'Fecha y hora en que llegó',
  incidenteCanal: 'Incidente o denuncia',
  enlace: 'Enlace o número de documento',
};
const tituloDistrito = (dep) => 'Distrito (' + dep + ')';
const tituloConsulta = (c) => 'Consultas acumuladas: ' + c;
const COLS_DATOS = ['bandera', 'menciona_reniec', 'resumen_publicable', 'validado'];

function crearFormulario() {
  const ss = SpreadsheetApp.create('Monitoreo ERM 2026 · Respuestas');
  ss.setSpreadsheetTimeZone(CONFIG.ZONA);
  const form = FormApp.create(CONFIG.TITULO);
  form.setDescription(
    'Registre aquí lo que le comunican los orientadores de RENIEC y lo que llega por otros canales. ' +
    'Transcriba lo reportado; no hace falta categorizar. No escriba DNI ni nombres de ciudadanos.'
  );
  form.setConfirmationMessage('Registro recibido. Entra al dashboard cuando el equipo de Datos lo valide.');
  form.setAllowResponseEdits(false).setCollectEmail(false).setProgressBar(true).setShowLinkToRespondAgain(true);
  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // ---- Sección 1: monitor y departamento
  form.addTextItem().setTitle(T.monitor).setRequired(true);
  const depItem = form.addListItem().setTitle(T.dep).setRequired(true);

  // ---- Una sección por departamento, con su lista de distritos
  const deps = Object.keys(DISTRITOS).sort();
  const pbDep = {};
  deps.forEach((dep) => {
    pbDep[dep] = form.addPageBreakItem().setTitle('Distrito de ' + dep);
    const li = form.addListItem().setTitle(tituloDistrito(dep)).setRequired(true);
    li.setChoiceValues(DISTRITOS[dep].map((d) => d[1] + ' · ' + d[2] + ' (' + d[0] + ')'));
  });

  // ---- Qué va a registrar
  const pbTipo = form.addPageBreakItem().setTitle('¿Qué va a registrar?');
  const tipoItem = form.addMultipleChoiceItem().setTitle(T.tipo).setRequired(true);

  // ---- Ruta A: comunicación de un orientador
  const pbOri = form.addPageBreakItem().setTitle('Comunicación de un orientador');
  const orientadores = CONFIG.ORIENTADORES.length ? CONFIG.ORIENTADORES : [CONFIG.PLACEHOLDER_ORIENTADOR];
  form.addListItem().setTitle(T.orientador).setRequired(true).setChoiceValues(orientadores);
  form.addMultipleChoiceItem().setTitle(T.momento).setRequired(true).setChoiceValues(CONFIG.MOMENTOS);
  form.addTextItem().setTitle(T.local).setRequired(true);
  form.addTimeItem().setTitle(T.llegada).setHelpText('Opcional. Normalmente en el reporte de inicio.');
  form.addTimeItem().setTitle(T.termino).setHelpText('Opcional. Normalmente en el reporte de cierre.');
  const validNum = FormApp.createTextValidation()
    .setHelpText('Escriba un número entero, 0 o mayor.')
    .requireWholeNumber().build();
  CONFIG.CONSULTAS.forEach((c) =>
    form.addTextItem().setTitle(tituloConsulta(c))
      .setHelpText('Opcional. Total acumulado hasta este momento, no solo lo nuevo.')
      .setValidation(validNum)
  );
  form.addParagraphTextItem().setTitle(T.incidenteOri)
    .setHelpText('Opcional. Transcriba lo que reportó el orientador. Sin incidencias: déjelo en blanco.');

  // ---- Ruta B: canal informal
  const pbCan = form.addPageBreakItem().setTitle('Canal informal');
  form.addListItem().setTitle(T.canal).setRequired(true).setChoiceValues(CONFIG.CANALES);
  form.addTextItem().setTitle(T.quien).setHelpText('Opcional. Entidad, medio o cuenta que lo envía.');
  form.addDateTimeItem().setTitle(T.fechaCanal).setRequired(true);
  form.addParagraphTextItem().setTitle(T.incidenteCanal).setRequired(true)
    .setHelpText('Transcriba lo recibido. No hace falta categorizar.');
  form.addTextItem().setTitle(T.enlace).setHelpText('Opcional.');

  // ---- Ramificación (se conecta al final, cuando ya existen todas las secciones)
  depItem.setChoices(deps.map((d) => depItem.createChoice(d, pbDep[d])));
  deps.forEach((d) => pbDep[d].setGoToPage(pbTipo));
  tipoItem.setChoices([
    tipoItem.createChoice(CONFIG.TIPOS[0], pbOri),
    tipoItem.createChoice(CONFIG.TIPOS[1], pbCan),
  ]);
  pbOri.setGoToPage(FormApp.PageNavigationType.SUBMIT);

  const props = PropertiesService.getScriptProperties();
  props.setProperty('FORM_ID', form.getId());
  props.setProperty('SS_ID', ss.getId());
  SpreadsheetApp.flush();
  configurarHoja();
  Logger.log('Formulario para el monitor (enlace para compartir): ' + form.getPublishedUrl());
  Logger.log('Editar el formulario: ' + form.getEditUrl());
  Logger.log('Hoja de respuestas: ' + ss.getUrl());
}

/** Orden exacto de las columnas que Google Forms crea en RESPUESTAS. */
function columnasRespuestas() {
  const deps = Object.keys(DISTRITOS).sort();
  return ['Marca temporal', T.monitor, T.dep]
    .concat(deps.map(tituloDistrito))
    .concat([T.tipo, T.orientador, T.momento, T.local, T.llegada, T.termino])
    .concat(CONFIG.CONSULTAS.map(tituloConsulta))
    .concat([T.incidenteOri, T.canal, T.quien, T.fechaCanal, T.incidenteCanal, T.enlace]);
}

function letra(n) { // 1 -> A
  let s = '';
  while (n > 0) { const r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

/** Crea o rehace las pestañas ORIENTADORES, SEGUIMIENTO, CALCULADO y PUBLICO y las columnas de Datos. Se puede repetir. */
function configurarHoja() {
  const ss = SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SS_ID'));
  const resp = ss.getSheets().find((h) => /^(Respuestas de formulario|Form Responses)/.test(h.getName()));
  if (!resp) throw new Error('Aún no existe la pestaña de respuestas del Forms. Espere unos segundos y ejecute configurarHoja().');
  resp.setName('RESPUESTAS');

  const cols = columnasRespuestas();
  const C = (nombre) => {
    const i = cols.indexOf(nombre);
    if (i < 0) throw new Error('Columna desconocida: ' + nombre);
    return 'RESPUESTAS!$' + letra(i + 1) + '$2:$' + letra(i + 1);
  };
  const colDatos = (k) => 'RESPUESTAS!$' + letra(cols.length + 1 + COLS_DATOS.indexOf(k)) + '$2:$' + letra(cols.length + 1 + COLS_DATOS.indexOf(k));

  // Columnas de Datos al final de RESPUESTAS (no las toca el Forms)
  resp.getRange(1, cols.length + 1, 1, COLS_DATOS.length).setValues([COLS_DATOS]).setFontWeight('bold').setBackground('#fff2cc');
  const sino = SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true).setAllowInvalid(false).build();
  resp.getRange(2, cols.length + 1, 998, 1).setDataValidation(sino);
  resp.getRange(2, cols.length + 2, 998, 1).setDataValidation(sino);
  resp.getRange(2, cols.length + 4, 998, 1).setDataValidation(sino);
  resp.setFrozenRows(1);

  const hoja = (nombre) => ss.getSheetByName(nombre) || ss.insertSheet(nombre);

  // ---- ORIENTADORES (catálogo)
  const ori = hoja('ORIENTADORES');
  if (ori.getLastRow() < 1) {
    ori.getRange(1, 1, 1, 3).setValues([['orientador (igual que en el desplegable, p. ej. APELLIDO NOMBRE | Local)', 'local', 'ubigeo_inei']]).setFontWeight('bold');
    ori.getRange(1, 5).setValue('Pegue aquí la lista y ejecute actualizarOrientadores() para llenar el desplegable del formulario.');
    ori.setColumnWidth(1, 420);
  }

  // ---- SEGUIMIENTO (quién ya reportó en cada momento)
  const seg = hoja('SEGUIMIENTO');
  seg.clear();
  const momentos = CONFIG.MOMENTOS.slice(0, 3);
  seg.getRange(1, 1, 1, 5).setValues([['Orientador'].concat(momentos).concat(['Último reporte'])]).setFontWeight('bold');
  const orient = C(T.orientador), mom = C(T.momento), marca = 'RESPUESTAS!$A$2:$A';
  const filas = [];
  for (let r = 2; r <= CONFIG.FILAS_SEGUIMIENTO + 1; r++) {
    const fila = ['=IF(ORIENTADORES!A' + r + '="","",ORIENTADORES!A' + r + ')'];
    momentos.forEach((m) => {
      fila.push('=IF($A' + r + '="","",IF(COUNTIFS(' + orient + ',$A' + r + ',' + mom + ',"' + m + '")>0,"✓ "&TEXT(MAXIFS(' + marca + ',' + orient + ',$A' + r + ',' + mom + ',"' + m + '"),"HH:mm"),"pendiente"))');
    });
    fila.push('=IF($A' + r + '="","",IFERROR(TEXT(MAXIFS(' + marca + ',' + orient + ',$A' + r + '),"HH:mm"),""))');
    filas.push(fila);
  }
  seg.getRange(2, 1, filas.length, 5).setFormulas(filas);
  const regla = SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo('pendiente').setBackground('#fce8b2').setRanges([seg.getRange(2, 2, filas.length, 3)]).build();
  seg.setConditionalFormatRules([regla]);
  seg.setFrozenRows(1);
  seg.setColumnWidth(1, 380);

  // ---- CALCULADO (una fila por respuesta, con los campos derivados)
  const deps = Object.keys(DISTRITOS).sort();
  const concatDistritos = deps.map((d) => C(tituloDistrito(d)));
  const af = (expr) => '=ARRAYFORMULA(IF(RESPUESTAS!$A$2:$A="","",' + expr + '))';
  const etq = '$A$2:$A'; // la etiqueta «PROV · DIST (ubigeo)» es la columna A de CALCULADO (campos[0])
  const esOri = C(T.tipo) + '="' + CONFIG.TIPOS[0] + '"';
  const campos = [
    ['etiqueta_distrito', af(concatDistritos.join('&'))],
    ['tipo', af(C(T.tipo))],
    ['momento_o_canal', af('IF(' + esOri + ',' + C(T.momento) + ',' + C(T.canal) + ')')],
    ['fecha_hora', af('IF(' + esOri + ',RESPUESTAS!$A$2:$A,' + C(T.fechaCanal) + ')')],
    ['departamento', af(C(T.dep))],
    ['ubigeo_inei', af('REGEXEXTRACT(' + etq + ',"\\((\\d{6})\\)\\s*$")')],
    ['provincia', af('REGEXEXTRACT(' + etq + ',"^(.*?) · ")')],
    ['distrito', af('REGEXEXTRACT(' + etq + ',"· (.*) \\(\\d{6}\\)\\s*$")')],
    ['local', af('IF(' + esOri + ',' + C(T.local) + ',"")')],
  ].concat(CONFIG.CONSULTAS.map((c, i) => ['consulta_' + (i + 1), af('IF(' + esOri + ',' + C(tituloConsulta(c)) + ',"")')]))
   .concat([
    ['bandera', af(colDatos('bandera'))],
    ['menciona_reniec', af(colDatos('menciona_reniec'))],
    ['resumen_publicable', af(colDatos('resumen_publicable'))],
    ['validado', af(colDatos('validado'))],
  ]);
  const calc = hoja('CALCULADO');
  calc.clear();
  calc.getRange(1, 1, 1, campos.length).setValues([campos.map((c) => c[0])]).setFontWeight('bold');
  calc.getRange(2, 1, 1, campos.length).setFormulas([campos.map((c) => c[1])]);
  calc.setFrozenRows(1);

  // ---- PUBLICO: solo validado = SI y sin etiqueta_distrito ni validado
  const pub = hoja('PUBLICO');
  pub.clear();
  const nombres = campos.map((c) => c[0]);
  const publicos = nombres.filter((n) => n !== 'etiqueta_distrito' && n !== 'validado');
  pub.getRange(1, 1, 1, publicos.length).setValues([publicos]).setFontWeight('bold');
  const L = (n) => letra(nombres.indexOf(n) + 1);
  const ult = letra(nombres.length);
  const rango = (n) => 'CALCULADO!' + L(n) + '2:' + L(n);
  const partes = publicos.map((n) => rango(n)).join(',');
  pub.getRange(2, 1).setFormula('=IFERROR(FILTER({' + partes + '},CALCULADO!' + L('validado') + '2:' + L('validado') + '="SI"),"")');
  pub.setFrozenRows(1);
}

/** Aplica la lista de la pestaña ORIENTADORES al desplegable del formulario. */
function actualizarOrientadores() {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  const ss = SpreadsheetApp.openById(props.getProperty('SS_ID'));
  const lista = ss.getSheetByName('ORIENTADORES').getRange('A2:A').getValues()
    .map((f) => String(f[0]).trim()).filter((v) => v);
  if (!lista.length) throw new Error('La pestaña ORIENTADORES está vacía.');
  const item = form.getItems(FormApp.ItemType.LIST).find((i) => i.getTitle() === T.orientador).asListItem();
  item.setChoiceValues(lista);
  Logger.log('Desplegable actualizado con ' + lista.length + ' orientadores.');
}

/** Muestra los enlaces del formulario ya creado. */
function enlaces() {
  const props = PropertiesService.getScriptProperties();
  const form = FormApp.openById(props.getProperty('FORM_ID'));
  Logger.log('Enlace para los monitores: ' + form.getPublishedUrl());
  Logger.log('Hoja: ' + SpreadsheetApp.openById(props.getProperty('SS_ID')).getUrl());
}
