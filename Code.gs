function doGet(e) {
  var email = Session.getActiveUser().getEmail();
  if (!email) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;">Debes iniciar sesión con tu cuenta de Farmatodo para continuar.</p>');
  }
  var role = getRole_(email);
  var page = (e.parameter.page || 'form').toLowerCase();

  if (page === 'admin') {
    if (role !== 'admin') return accessDenied_();
    var t = HtmlService.createTemplateFromFile('Admin');
    t.email = email; t.role = role;
    return t.evaluate().setTitle('Panel de administración').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'instructor') {
    if (role !== 'admin' && role !== 'instructor') return accessDenied_();
    var t2 = HtmlService.createTemplateFromFile('Instructor');
    t2.email = email; t2.role = role;
    return t2.evaluate().setTitle('Editor de preguntas').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  var t3 = HtmlService.createTemplateFromFile('Form');
  t3.email = email; t3.role = role;
  return t3.evaluate().setTitle('Encuesta de adopción de IA').addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function accessDenied_() {
  return HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;">No tienes permiso para ver esta página. Si crees que es un error, contacta al administrador.</p>');
}

function getSheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function getRole_(email) {
  var sheet = getSheet_('Accesos');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase().trim()) {
      return String(data[i][1]).toLowerCase().trim();
    }
  }
  return 'usuario';
}

function requireRole_(allowed) {
  var email = Session.getActiveUser().getEmail();
  var role = getRole_(email);
  if (allowed.indexOf(role) === -1) throw new Error('No tienes permiso para esta acción.');
  return { email: email, role: role };
}

/* ---------- Preguntas (público autenticado lee, instructor/admin escribe) ---------- */

function getPreguntasPublicas() {
  var sheet = getSheet_('Preguntas');
  var data = sheet.getDataRange().getValues();
  var rows = data.slice(1).filter(function (r) { return r[0] !== ''; });
  rows.sort(function (a, b) { return a[2] - b[2]; });
  return rows.map(function (r) {
    return {
      id: r[0], seccion: r[1], orden: r[2], texto: r[3], tipo: r[4],
      opciones: r[5] ? String(r[5]).split('|') : [],
      requerido: String(r[6]).toUpperCase() === 'SI',
      otro: String(r[7]).toUpperCase() === 'SI'
    };
  });
}

function guardarPregunta(p) {
  requireRole_(['admin', 'instructor']);
  var sheet = getSheet_('Preguntas');
  var data = sheet.getDataRange().getValues();
  var opciones = (p.opciones || []).join('|');
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(p.id)) {
      sheet.getRange(i + 1, 1, 1, 8).setValues([[p.id, p.seccion, p.orden, p.texto, p.tipo, opciones, p.requerido ? 'SI' : 'NO', p.otro ? 'SI' : 'NO']]);
      return { ok: true };
    }
  }
  sheet.appendRow([p.id, p.seccion, p.orden, p.texto, p.tipo, opciones, p.requerido ? 'SI' : 'NO', p.otro ? 'SI' : 'NO']);
  return { ok: true };
}

function eliminarPregunta(id) {
  requireRole_(['admin', 'instructor']);
  var sheet = getSheet_('Preguntas');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(id)) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

/* ---------- Respuestas (cualquiera autenticado guarda, solo admin lee) ---------- */

function guardarRespuesta(payload) {
  var email = Session.getActiveUser().getEmail();
  if (!email) throw new Error('Debes iniciar sesión.');
  var sheet = getSheet_('Respuestas');
  sheet.appendRow([
    new Date(), email, payload.nombre, payload.apellido, payload.area, payload.pais, payload.cedula,
    JSON.stringify(payload.respuestas || {})
  ]);
  return { ok: true };
}

function getRespuestasAdmin() {
  requireRole_(['admin']);
  var sheet = getSheet_('Respuestas');
  var data = sheet.getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[1] !== ''; }).map(function (r) {
    return {
      timestamp: r[0], email: r[1], nombre: r[2], apellido: r[3], area: r[4], pais: r[5], cedula: r[6],
      respuestas: JSON.parse(r[7] || '{}')
    };
  });
}

function getAnaliticaAdmin() {
  requireRole_(['admin']);
  var rows = getRespuestasAdmin();
  var preguntas = getPreguntasPublicas();
  var analitica = {};
  preguntas.forEach(function (p) {
    if (p.tipo === 'radio' || p.tipo === 'select') {
      var counts = {};
      p.opciones.forEach(function (o) { counts[o] = 0; });
      rows.forEach(function (r) {
        var v = r.respuestas[p.id];
        if (v && counts.hasOwnProperty(v)) counts[v]++;
      });
      analitica[p.id] = { texto: p.texto, counts: counts };
    } else if (p.tipo === 'checkbox') {
      var counts2 = {};
      p.opciones.forEach(function (o) { counts2[o] = 0; });
      rows.forEach(function (r) {
        var v = r.respuestas[p.id];
        if (Array.isArray(v)) v.forEach(function (x) { if (counts2.hasOwnProperty(x)) counts2[x]++; });
      });
      analitica[p.id] = { texto: p.texto, counts: counts2 };
    }
  });
  return { total: rows.length, preguntas: analitica };
}

/* ---------- Accesos (solo admin) ---------- */

function getAccesos() {
  requireRole_(['admin']);
  var sheet = getSheet_('Accesos');
  var data = sheet.getDataRange().getValues();
  return data.slice(1).filter(function (r) { return r[0] !== ''; }).map(function (r) { return { email: r[0], rol: r[1] }; });
}

function guardarAcceso(email, rol) {
  requireRole_(['admin']);
  var sheet = getSheet_('Accesos');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase().trim()) {
      sheet.getRange(i + 1, 2).setValue(rol);
      return { ok: true };
    }
  }
  sheet.appendRow([email, rol]);
  return { ok: true };
}

function eliminarAcceso(email) {
  requireRole_(['admin']);
  var sheet = getSheet_('Accesos');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).toLowerCase().trim() === email.toLowerCase().trim()) {
      sheet.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false };
}

/* ---------- Areas (público autenticado lee, instructor/admin escribe) ---------- */

function getAreas() {
  var sheet = getSheet_('Areas');
  var data = sheet.getDataRange().getValues();
  return data.slice(1).map(function (r) { return r[0]; }).filter(Boolean);
}

function guardarArea(nombre) {
  requireRole_(['admin', 'instructor']);
  var sheet = getSheet_('Areas');
  if (getAreas().indexOf(nombre) !== -1) return { ok: false, error: 'Esa área ya existe.' };
  sheet.appendRow([nombre]);
  return { ok: true };
}

function eliminarArea(nombre) {
  requireRole_(['admin', 'instructor']);
  var sheet = getSheet_('Areas');
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(nombre)) { sheet.deleteRow(i + 1); return { ok: true }; }
  }
  return { ok: false };
}

/* ---------- Insights con IA (solo admin) ---------- */
/* Requiere una API key gratuita de Google AI Studio (aistudio.google.com/app/apikey)
   guardada en: Configuración del proyecto -> Propiedades del script -> GEMINI_API_KEY */

function getInsightsIA() {
  requireRole_(['admin']);
  var apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!apiKey) return { ok: false, error: 'Falta configurar GEMINI_API_KEY en Propiedades del script.' };

  var analitica = getAnaliticaAdmin();
  if (analitica.total === 0) return { ok: false, error: 'Todavía no hay respuestas suficientes para analizar.' };

  var resumenTexto = 'Total de respuestas: ' + analitica.total + '\n';
  Object.keys(analitica.preguntas).forEach(function (id) {
    var q = analitica.preguntas[id];
    resumenTexto += '\nPregunta: ' + q.texto + '\n';
    Object.keys(q.counts).forEach(function (op) { resumenTexto += '- ' + op + ': ' + q.counts[op] + '\n'; });
  });

  var prompt = 'Eres un analista de RRHH. A partir de estos resultados de una encuesta interna sobre adopción de IA en el equipo, escribe en español un resumen ejecutivo breve (máximo 6 líneas) con los hallazgos más relevantes y una recomendación concreta para el equipo de Capacitación (CENCAP) de Farmatodo:\n\n' + resumenTexto;

  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=' + apiKey;
  var options = {
    method: 'post', contentType: 'application/json',
    payload: JSON.stringify({ contents: [{ parts: [{ text: prompt }] }] }),
    muteHttpExceptions: true
  };
  try {
    var resp = UrlFetchApp.fetch(url, options);
    var json = JSON.parse(resp.getContentText());
    var texto = json.candidates && json.candidates[0] && json.candidates[0].content.parts[0].text;
    return { ok: true, texto: texto || 'No se pudo generar el resumen.' };
  } catch (e) {
    return { ok: false, error: 'Error al llamar la IA: ' + e.message };
  }
}

/* ---------- Instalación (ejecutar UNA sola vez) ---------- */

function inicializarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  crearHojaSiNoExiste_(ss, 'Accesos', ['Email', 'Rol']);
  crearHojaSiNoExiste_(ss, 'Preguntas', ['ID', 'Seccion', 'Orden', 'Texto', 'Tipo', 'Opciones', 'Requerido', 'Otro']);
  crearHojaSiNoExiste_(ss, 'Respuestas', ['Timestamp', 'Email', 'Nombre', 'Apellido', 'Area', 'Pais', 'Cedula', 'RespuestasJSON']);
  crearHojaSiNoExiste_(ss, 'Areas', ['Area']);
  var areasSheet = getSheet_('Areas');
  if (areasSheet.getLastRow() < 2) {
    ['CENCAP / Capacitación', 'Operaciones', 'Logística y Distribución', 'Tecnología',
     'Comercial', 'Finanzas', 'Call Center', 'Recursos Humanos', 'Farmacia (Tienda)', 'Otra']
      .forEach(function (a) { areasSheet.appendRow([a]); });
  }


  var accesos = getSheet_('Accesos');
  if (accesos.getLastRow() < 2) {
    accesos.appendRow([Session.getActiveUser().getEmail(), 'admin']);
  }

  var preguntas = getSheet_('Preguntas');
  if (preguntas.getLastRow() < 2) {
    var seed = [
      ['q1', 'Sección 1: Nivel de familiaridad y conocimiento general', 1, '¿Cómo calificarías tu nivel general de conocimiento sobre inteligencia artificial?', 'radio', 'Básico: conozco los conceptos generales.|Intermedio: he usado herramientas o APIs de IA.|Avanzado: he desarrollado soluciones con modelos de IA.|Experto: diseño y entreno modelos de IA.', 'SI', 'NO'],
      ['q2', 'Sección 1: Nivel de familiaridad y conocimiento general', 2, '¿Has recibido alguna formación o capacitación reciente relacionada con IA?', 'radio', 'Sí, interna (dada por la empresa).|Sí, externa (cursos, bootcamps, etc.).|No.', 'SI', 'NO'],
      ['q3', 'Sección 2: Herramientas y uso actual', 3, '¿Con qué frecuencia usas herramientas basadas en IA en tu trabajo diario?', 'radio', 'Nunca|Ocasionalmente|Semanalmente|A diario', 'SI', 'NO'],
      ['q4', 'Sección 2: Herramientas y uso actual', 4, '¿Qué herramientas basadas en IA o automatización usas actualmente?', 'checkbox', 'GitHub Copilot|ChatGPT / Gemini / Claude / Grok / Deepseek|n8n|Replit|Lovable|Appsmith|V0.dev|Cursor|Windsurf|Claude Code|Codex|Openclaw|Nano Banana|Sora|Eleven Labs|Ninguna', 'SI', 'SI'],
      ['q5', 'Sección 2: Herramientas y uso actual', 5, '¿Para qué tareas sueles usar IA en tu trabajo?', 'checkbox', 'Generación de código|Depuración o revisión de código|Generación de documentación|Automatización de procesos|Soporte en decisiones técnicas|Análisis de datos o dashboards|Optimización de consultas SQL|Diseño de soluciones técnicas|Diseño UI/UX', 'SI', 'SI'],
      ['q6', 'Sección 3: Cultura y barreras', 6, '¿Sientes que la empresa promueve activamente el uso de IA?', 'radio', 'Sí, mucho|Algo|Poco|Nada', 'SI', 'NO'],
      ['q7', 'Sección 3: Cultura y barreras', 7, '¿Qué barreras enfrentas para adoptar más herramientas de IA?', 'checkbox', 'Falta de conocimiento|Falta de tiempo|Falta de lineamientos o permisos|Desconfianza en los resultados|No sé por dónde empezar|Ninguna', 'SI', 'SI'],
      ['q8', 'Sección 3: Cultura y barreras', 8, '¿Te gustaría recibir capacitaciones o talleres prácticos sobre IA aplicada a tu trabajo?', 'radio', 'Sí|Tal vez|No', 'SI', 'NO'],
      ['q9', 'Sección 4: Aplicación práctica reciente', 9, 'En el último mes, ¿has creado alguno de los siguientes utilizando inteligencia artificial o automatización?', 'checkbox', 'Fragmentos de código o scripts individuales|Un feature completo (funcionalidad terminada y desplegada)|Un producto completo (MVP o solución lista para usuarios internos o externos)|No, pero estoy explorando ideas|Infra como código|No', 'SI', 'SI'],
      ['q10', 'Sección 5: Interés y próximos pasos', 10, '¿Qué temas específicos te gustaría aprender o profundizar?', 'checkbox', 'Low code y automatización (n8n, Cursor, Windsurf, Replit, Lovable)|APIs y modelos como GPT, Claude, Gemini|Integración de IA en flujos de trabajo|LLMOps / Fine-tuning de modelos|Ética, seguridad y gobernanza en IA', 'SI', 'SI']
    ];
    seed.forEach(function (r) { preguntas.appendRow(r); });
  }
  Logger.log('Listo. Hojas creadas y preguntas sembradas.');
}

function crearHojaSiNoExiste_(ss, nombre, headers) {
  var sheet = ss.getSheetByName(nombre);
  if (!sheet) {
    sheet = ss.insertSheet(nombre);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }
  return sheet;
}
