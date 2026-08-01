/**
 * Sala de Edición · YOD — el conector con el Sheet.
 *
 * Qué hace: guarda cada revisión de Alejandro en el Sheet (nada se pisa: además del
 * estado hay bitácora), le sirve el día al portal, y manda el correo de las 7:00 am
 * con la liga y el resumen.
 *
 * CÓMO SE INSTALA (una sola vez, ~3 minutos):
 *   1. Abre un Google Sheet nuevo y ponle "Sala de Edición · YOD".
 *   2. Extensiones → Apps Script. Borra lo que haya y pega ESTE archivo completo.
 *   3. Arriba, corre la función  instalar  (▶). Autoriza cuando lo pida.
 *      — Eso crea las pestañas, pone una clave nueva en CONFIG y deja el correo
 *        diario programado a las 7:00 (hora de Hermosillo).
 *   4. Implementar → Nueva implementación → Aplicación web:
 *        · Ejecutar como: TÚ  ·  Acceso: Cualquier persona
 *      Copia la URL que termina en /exec.
 *   5. En el portal, pícale a «⚙︎ conectar Sheet» y pega la URL y la clave
 *      (la clave está en la pestaña CONFIG del Sheet).
 *
 * La clave es la "seguridad suave" de la casa: sin ella no se lee ni se escribe.
 */

var TZ = 'America/Hermosillo';
var PORTAL = 'https://alexpueblag.github.io/sala-edicion/';  // al migrar: yodesarrollo.github.io/sala-edicion
var CORREO = 'direccion@aurumarquitectos.com';

var PESTANAS = {
  CONFIG:     ['clave', 'valor'],
  PROPUESTAS: ['fecha', 'prop_id', 'titulo', 'tipo', 'laminas_json', 'opciones_json', 'video', 'estado'],
  DECISIONES: ['guardado', 'fecha', 'prop_id', 'lamina', 'marca', 'nota_propuesta', 'nota_dia'],
  PARRILLA:   ['fecha', 'pieza', 'gate', 'desde'],
  CONTROL:    ['pieza', 'desde', 'formula', 'alcance', 'clics', 'leads', 'nota'],
  PRODUCCION: ['fecha', 'pieza', 'estado', 'detalle', 'enlace'],
  BITACORA:   ['fecha_hora', 'evento', 'detalle']
};

function instalar() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(PESTANAS).forEach(function (n) {
    var h = ss.getSheetByName(n) || ss.insertSheet(n);
    if (h.getLastRow() === 0) {
      h.appendRow(PESTANAS[n]);
      h.setFrozenRows(1);
    }
  });
  if (!leerConfig('clave')) {
    var clave = Utilities.getUuid().slice(0, 8);
    ss.getSheetByName('CONFIG').appendRow(['clave', clave]);
  }
  // correo diario 7:00 am — se borra el trigger previo para no duplicar
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'correoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('correoDiario').timeBased().everyDays(1).atHour(7)
    .inTimezone(TZ).create();
  bitacora('Sala instalada; correo diario programado a las 7:00');
}

/* ------------------------------------------------ utilería */
function hoja(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function hoy() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function leerConfig(k) {
  var v = hoja('CONFIG'); if (!v) return '';
  var datos = v.getDataRange().getValues();
  for (var i = 1; i < datos.length; i++) if (String(datos[i][0]) === k) return String(datos[i][1]);
  return '';
}
function filas(n) {
  var h = hoja(n); if (!h || h.getLastRow() < 2) return [];
  var cab = PESTANAS[n];
  return h.getRange(2, 1, h.getLastRow() - 1, cab.length).getValues().map(function (r) {
    var o = {}; cab.forEach(function (c, i) { o[c] = r[i]; }); return o;
  });
}
function bitacora(evento, detalle) {
  hoja('BITACORA').appendRow([ahora(), evento, detalle || '']);
}
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
    .setMimeType(ContentService.MimeType.JSON);
}
function fechaDe(v) {
  return (v instanceof Date) ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v).slice(0, 10);
}

/* ------------------------------------------------ lectura: el día del portal */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.clave !== leerConfig('clave')) return json({ error: 'clave incorrecta' });
  if (p.recurso !== 'dia') return json({ error: 'recurso desconocido' });

  var f = /^\d{4}-\d{2}-\d{2}$/.test(p.f || '') ? p.f : hoy();
  var props = filas('PROPUESTAS').filter(function (x) { return fechaDe(x.fecha) === f; })
    .map(function (x) {
      var lam, opc;
      try { lam = JSON.parse(x.laminas_json); } catch (err) { lam = []; }
      try { opc = JSON.parse(x.opciones_json); } catch (err) { opc = []; }
      return { id: String(x.prop_id), titulo: String(x.titulo),
               tipo: String(x.tipo || 'laminas'), laminas: lam, opciones: opc,
               video: String(x.video || '') || null };
    });

  // la última revisión guardada de ese día, reconstruida por propuesta
  var dec = { propuestas: {} }, ult = '';
  filas('DECISIONES').forEach(function (d) {
    if (fechaDe(d.fecha) !== f) return;
    var pid = String(d.prop_id);
    var ficha = dec.propuestas[pid] = dec.propuestas[pid] || { laminas: [], nota: '' };
    if (d.lamina !== '' && d.lamina !== null) ficha.laminas[Number(d.lamina)] = String(d.marca || '') || null;
    if (d.nota_propuesta) ficha.nota = String(d.nota_propuesta);
    if (d.nota_dia) dec.nota_general = String(d.nota_dia);
    if (String(d.guardado) > ult) ult = String(d.guardado);
  });
  if (ult) dec.guardado = ult;

  // la retro de ayer, contada desde los registros
  var ayer = Utilities.formatDate(new Date(new Date(f + 'T12:00:00').getTime() - 864e5), TZ, 'yyyy-MM-dd');
  var r = { fecha: ayer, aprobadas: 0, tiradas: 0, producidas: 0, rehechas: 0, nota: '' };
  filas('DECISIONES').forEach(function (d) {
    if (fechaDe(d.fecha) !== ayer) return;
    if (String(d.marca) === 'si') r.aprobadas++;
    if (String(d.marca) === 'no') r.tiradas++;
    if (d.nota_dia) r.nota = String(d.nota_dia);
  });
  filas('PRODUCCION').forEach(function (x) {
    if (fechaDe(x.fecha) !== ayer) return;
    if (String(x.estado) === 'video' || String(x.estado) === 'publicada') r.producidas++;
    if (String(x.estado) === 'rehacer') r.rehechas++;
  });

  var dias = {}; filas('PROPUESTAS').forEach(function (x) { dias[fechaDe(x.fecha)] = 1; });
  var bit = filas('BITACORA').filter(function (b) {
    return String(b.fecha_hora).slice(0, 10) === f;
  }).map(function (b) {
    return { hora: String(b.fecha_hora).slice(11, 16), evento: String(b.evento) };
  });

  // parrilla: solo de hoy en adelante (5 semanas) · control: la ultima fila coronada
  var lim = Utilities.formatDate(new Date(new Date(f + 'T12:00:00').getTime() + 35 * 864e5), TZ, 'yyyy-MM-dd');
  var parr = filas('PARRILLA').filter(function (x) {
    var d = fechaDe(x.fecha); return d >= f && d <= lim;
  }).map(function (x) {
    return { fecha: fechaDe(x.fecha), pieza: String(x.pieza || ''),
             gate: String(x.gate || ''), desde: x.desde ? fechaDe(x.desde) : '' };
  }).sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
  var ctrl = filas('CONTROL').pop() || null;

  return json({
    fecha: f, dias: Object.keys(dias).sort(), propuestas: props, decisiones: dec,
    retro: r, parrilla: parr, control: ctrl,
    produccion: filas('PRODUCCION').reverse().slice(0, 30).map(function (x) {
      x.fecha = fechaDe(x.fecha); return x;
    }),
    bitacora: bit
  });
}

/* ------------------------------------------------ escritura */
function doPost(e) {
  var d; try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ error: 'cuerpo ilegible' }); }
  if (d.clave !== leerConfig('clave')) return json({ error: 'clave incorrecta' });

  if (d.accion === 'decidir') {
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.fecha || '')) return json({ error: 'falta la fecha; no se guardó nada' });
    var sello = ahora(), h = hoja('DECISIONES'), fs = [];
    Object.keys(d.propuestas || {}).forEach(function (pid) {
      var p = d.propuestas[pid] || {};
      (p.laminas || []).forEach(function (m, i) {
        fs.push([sello, d.fecha, pid, i, m || '', '', '']);
      });
      if (p.nota) fs.push([sello, d.fecha, pid, '', '', p.nota, '']);
    });
    if (d.nota_general) fs.push([sello, d.fecha, '', '', '', '', d.nota_general]);
    if (fs.length) h.getRange(h.getLastRow() + 1, 1, fs.length, 7).setValues(fs);
    bitacora('Revisión del editor recibida', fs.length + ' registro(s)');
    return json({ ok: true, guardado: sello });
  }

  if (d.accion === 'proponer') {          // la Mac sube las propuestas del día
    var hp = hoja('PROPUESTAS');
    (d.propuestas || []).forEach(function (p) {
      hp.appendRow([d.fecha || hoy(), p.id, p.titulo, p.tipo || 'laminas',
                    JSON.stringify(p.laminas || []), JSON.stringify(p.opciones || []),
                    p.video || '', 'en revisión']);
    });
    bitacora('Propuestas del día montadas', (d.propuestas || []).length + ' propuesta(s)');
    return json({ ok: true });
  }

  if (d.accion === 'parrilla') {          // la Mac actualiza el hueco de una fecha
    var hpa = hoja('PARRILLA'), datos = hpa.getDataRange().getValues(), fila = -1;
    for (var i = 1; i < datos.length; i++)
      if (fechaDe(datos[i][0]) === d.dia) { fila = i + 1; break; }
    var v = [d.dia, d.pieza || '', d.gate || '', d.desde || hoy()];
    if (fila > 0) hpa.getRange(fila, 1, 1, 4).setValues([v]); else hpa.appendRow(v);
    bitacora('Parrilla: ' + d.dia + ' → ' + (d.pieza || 'hueco') + ' (' + (d.gate || '—') + ')');
    return json({ ok: true });
  }

  if (d.accion === 'produccion') {        // la Mac reporta avances de producción
    hoja('PRODUCCION').appendRow([d.fecha || hoy(), d.pieza || '', d.estado || '',
                                  d.detalle || '', d.enlace || '']);
    bitacora('Producción: ' + (d.pieza || '') + ' → ' + (d.estado || ''), d.detalle || '');
    return json({ ok: true });
  }

  return json({ error: 'acción desconocida' });
}

/* ------------------------------------------------ el correo de las 7:00 */
function correoDiario() {
  var f = hoy();
  var n = filas('PROPUESTAS').filter(function (x) { return fechaDe(x.fecha) === f; }).length;
  var pend = filas('PRODUCCION').filter(function (x) {
    return String(x.estado) === 'video' || String(x.estado) === 'animando';
  }).length;
  var asunto = n > 0
    ? 'Sala de Edición · ' + n + ' propuesta(s) esperan tu revisión'
    : 'Sala de Edición · hoy no hay pendientes de revisión';
  var html =
    '<div style="background:#0a0a0c;padding:34px 22px;font-family:Georgia,serif;color:#f4f1ec">' +
    '<p style="color:#debc7e;font-size:12px;letter-spacing:3px;margin:0 0 6px">YO DESARROLLO</p>' +
    '<h2 style="font-weight:400;margin:0 0 16px">Sala de <em style="color:#debc7e">Edición</em></h2>' +
    '<p style="font-size:15px;line-height:1.6;margin:0 0 22px">' +
    (n > 0 ? 'Hay <b>' + n + ' propuesta(s)</b> del día esperándote.' : 'Hoy no hay propuestas nuevas.') +
    (pend > 0 ? ' En producción: ' + pend + ' pieza(s) en camino.' : '') + '</p>' +
    '<a href="' + PORTAL + '" style="background:#c2a06b;color:#17130c;text-decoration:none;' +
    'padding:13px 26px;border-radius:8px;font-family:-apple-system,sans-serif;font-weight:600">' +
    'Abrir la Sala</a>' +
    '<p style="color:#8a867e;font-size:12px;margin:26px 0 0">Palomea lo que sirve, tacha lo que no, ' +
    'y escribe lo que quieres distinto. Cinco minutos.</p></div>';
  MailApp.sendEmail({ to: CORREO, subject: asunto, htmlBody: html });
  bitacora('Correo de las 7:00 enviado', asunto);
}
