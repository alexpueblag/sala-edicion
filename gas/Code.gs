/**
 * Sala de Edición · YOD — el conector con el Sheet.  (v3, espec de la mesa 1-ago-2026)
 *
 * Qué hace: recibe la revisión diaria del editor en UN envío consolidado e idempotente,
 * la guarda sin pisar nada, sirve el día al portal, registra las decisiones de parrilla
 * (andón), y manda el correo de las 7:00 con la liga que trae la sesión resuelta.
 *
 * CÓMO SE INSTALA (una sola vez, ~3 minutos):
 *   1. Google Sheet nuevo → "Sala de Edición · YOD".
 *   2. Extensiones → Apps Script → pegar ESTE archivo completo.
 *   3. Correr  instalar  (▶) y autorizar. Crea pestañas, 4 claves por rol y el correo 7:00.
 *   4. Implementar → Aplicación web (Ejecutar como TÚ · Acceso: cualquier persona) → copiar /exec.
 *   5. La liga del correo diario ya lleva la sesión; para la primera vez, en el portal:
 *      «conexión: … toca para conectar» y pegar /exec + la clave de CONFIG (fila clave).
 *
 * ROLES (patrón de la casa: clave suave por rol, en CONFIG):
 *   clave          → editor  (Alejandro: lee y decide)
 *   clave_editor2  → editor2 (Sayri: lee y decide; queda su autoría en 'quien')
 *   clave_lector   → lector  (clientes/socios: solo lectura)
 *   clave_agente   → agente  (la Mac: propone, reporta producción, acusa cosecha)
 */

var TZ = 'America/Hermosillo';
var PORTAL = 'https://alexpueblag.github.io/sala-edicion/';  // al migrar: yodesarrollo.github.io/sala-edicion
var CORREO = 'direccion@aurumarquitectos.com';

var PESTANAS = {
  CONFIG:     ['clave', 'valor'],
  PROPUESTAS: ['fecha', 'prop_id', 'titulo', 'tipo', 'laminas_json', 'opciones_json', 'video', 'estado', 'origen'],
  DECISIONES: ['envio_id', 'quien', 'guardado', 'fecha', 'prop_id', 'lamina', 'marca', 'nota_propuesta', 'nota_dia', 'nota_estrategia'],
  PARRILLA:   ['fecha', 'pieza', 'gate', 'desde', 'decision_editor', 'decidido'],
  CONTROL:    ['pieza', 'desde', 'formula', 'alcance', 'clics', 'leads', 'nota'],
  PRODUCCION: ['fecha', 'pieza', 'estado', 'detalle', 'enlace'],
  BITACORA:   ['fecha_hora', 'evento', 'detalle']
};

function instalar() {
  var ss = SpreadsheetApp.getActive();
  Object.keys(PESTANAS).forEach(function (n) {
    var h = ss.getSheetByName(n) || ss.insertSheet(n);
    if (h.getLastRow() === 0) { h.appendRow(PESTANAS[n]); h.setFrozenRows(1); }
  });
  // 4 claves por rol; si ya existe la del editor se conserva (migración sin dolor)
  var cfg = ss.getSheetByName('CONFIG');
  ['clave', 'clave_editor2', 'clave_lector', 'clave_agente'].forEach(function (k) {
    if (!leerConfig(k)) cfg.appendRow([k, Utilities.getUuid().slice(0, 8)]);
  });
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'correoDiario') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('correoDiario').timeBased().everyDays(1).atHour(7).inTimezone(TZ).create();
  bitacora('Sala instalada; roles creados; correo diario 7:00 programado');
}

/* ------------------------------------------------ utilería */
function hoja(n) { return SpreadsheetApp.getActive().getSheetByName(n); }
function hoy() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function ahora() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm:ss"); }
function leerConfig(k) {
  var v = hoja('CONFIG'); if (!v || v.getLastRow() < 1) return '';
  var datos = v.getDataRange().getValues();
  for (var i = 0; i < datos.length; i++) if (String(datos[i][0]) === k) return String(datos[i][1]);
  return '';
}
function rolDe(clave) {
  if (!clave) return null;
  if (clave === leerConfig('clave')) return 'editor';
  if (clave === leerConfig('clave_editor2')) return 'editor2';
  if (clave === leerConfig('clave_lector')) return 'lector';
  if (clave === leerConfig('clave_agente')) return 'agente';
  return null;
}
function filas(n) {
  var h = hoja(n); if (!h || h.getLastRow() < 2) return [];
  var cab = PESTANAS[n];
  return h.getRange(2, 1, h.getLastRow() - 1, cab.length).getValues().map(function (r) {
    var o = {}; cab.forEach(function (c, i) { o[c] = r[i]; }); return o;
  });
}
function bitacora(evento, detalle) { hoja('BITACORA').appendRow([ahora(), evento, detalle || '']); }
function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function fechaDe(v) {
  // a prueba de zonas: medianoche de CUALQUIER zona +12 h cae en la fecha correcta en UTC
  return (v instanceof Date)
    ? Utilities.formatDate(new Date(v.getTime() + 43200000), 'Etc/UTC', 'yyyy-MM-dd')
    : String(v).slice(0, 10);
}

/* El envio vigente de un dia = el del sello 'guardado' mas alto QUE PUSO EL GAS al
   recibir (nunca el reloj del telefono — objecion firmada del ingeniero). */
function envioVigente(decRows, f) {
  var mejor = '';
  decRows.forEach(function (d) {
    if (fechaDe(d.fecha) === f && String(d.guardado) > mejor) mejor = String(d.guardado);
  });
  var id = '';
  decRows.forEach(function (d) {
    if (fechaDe(d.fecha) === f && String(d.guardado) === mejor) id = String(d.envio_id);
  });
  return id;
}

/* ------------------------------------------------ lectura */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (!rolDe(p.clave)) return json({ error: 'clave incorrecta' });
  if (p.recurso !== 'dia') return json({ error: 'recurso desconocido' });
  var f = /^\d{4}-\d{2}-\d{2}$/.test(p.f || '') ? p.f : hoy();

  // una lectura por pestaña (presupuesto <3 s firmado por el ingeniero)
  var PR = filas('PROPUESTAS'), DE = filas('DECISIONES'), PA = filas('PARRILLA'),
      CO = filas('CONTROL'), PD = filas('PRODUCCION'), BI = filas('BITACORA');

  var props = PR.filter(function (x) { return fechaDe(x.fecha) === f; }).map(function (x) {
    var lam, opc;
    try { lam = JSON.parse(x.laminas_json); } catch (err) { lam = []; }
    try { opc = JSON.parse(x.opciones_json); } catch (err) { opc = []; }
    return { id: String(x.prop_id), titulo: String(x.titulo), tipo: String(x.tipo || 'laminas'),
             laminas: lam, opciones: opc, video: String(x.video || '') || null,
             origen: String(x.origen || '') || null };
  });

  // decisiones: SOLO el envio vigente del dia (reenviar deja numeros identicos)
  var vig = envioVigente(DE, f);
  var dec = { propuestas: {} }, ult = '';
  DE.forEach(function (d) {
    if (fechaDe(d.fecha) !== f || String(d.envio_id) !== vig) return;
    var pid = String(d.prop_id);
    if (pid) {
      var fi = dec.propuestas[pid] = dec.propuestas[pid] || { laminas: [], nota: '' };
      if (d.lamina !== '' && d.lamina !== null) {
        var m = String(d.marca || '');
        fi.laminas[Number(d.lamina)] = (m === 'si' || m === 'no') ? m : null;
        // la nota POR LAMINA es la instruccion de edicion del editor: no se pierde
        if (d.nota_propuesta) { fi.notas = fi.notas || []; fi.notas[Number(d.lamina)] = String(d.nota_propuesta); }
      } else if (d.nota_propuesta) fi.nota = String(d.nota_propuesta);
    }
    if (d.nota_dia) dec.nota_general = String(d.nota_dia);
    if (d.nota_estrategia) dec.nota_estrategia = String(d.nota_estrategia);
    if (String(d.guardado) > ult) ult = String(d.guardado);
  });
  if (ult) dec.guardado = ult;

  // retro de ayer, contando SOLO su envio vigente
  var ayer = Utilities.formatDate(new Date(new Date(f + 'T12:00:00').getTime() - 864e5), TZ, 'yyyy-MM-dd');
  var vigA = envioVigente(DE, ayer);
  var r = { fecha: ayer, aprobadas: 0, tiradas: 0, producidas: 0, rehechas: 0, nota: '' };
  DE.forEach(function (d) {
    if (fechaDe(d.fecha) !== ayer || String(d.envio_id) !== vigA) return;
    if (String(d.marca) === 'si') r.aprobadas++;
    if (String(d.marca) === 'no') r.tiradas++;
    if (d.nota_dia) r.nota = String(d.nota_dia);
  });
  PD.forEach(function (x) {
    if (fechaDe(x.fecha) !== ayer) return;
    if (String(x.estado) === 'video' || String(x.estado) === 'publicada') r.producidas++;
  });
  r.rehechas = props.filter(function (x) { return x.origen; }).length;

  var lim = Utilities.formatDate(new Date(new Date(f + 'T12:00:00').getTime() + 35 * 864e5), TZ, 'yyyy-MM-dd');
  var parr = PA.filter(function (x) { var d = fechaDe(x.fecha); return d >= f && d <= lim; })
    .map(function (x) { return { fecha: fechaDe(x.fecha), pieza: String(x.pieza || ''),
      gate: String(x.gate || ''), desde: x.desde ? fechaDe(x.desde) : '',
      decision_editor: String(x.decision_editor || '') }; })
    .sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });

  var dias = {}; PR.forEach(function (x) { dias[fechaDe(x.fecha)] = 1; });
  var bit = BI.filter(function (b) { return String(b.fecha_hora).slice(0, 10) === f; })
    .map(function (b) { return { hora: String(b.fecha_hora).slice(11, 16), evento: String(b.evento) }; });

  // ultima revision enviada (para el modo «Mientras no estabas» del Umbral)
  var ultRev = '';
  DE.forEach(function (d) { if (String(d.guardado) > ultRev) ultRev = String(d.guardado); });

  return json({
    fecha: f, dias: Object.keys(dias).sort(), propuestas: props, decisiones: dec, retro: r,
    parrilla: parr, control: CO.length ? CO[CO.length - 1] : null,
    produccion: PD.slice().reverse().slice(0, 30).map(function (x) { x.fecha = fechaDe(x.fecha); return x; }),
    bitacora: bit, ultima_revision: ultRev, rol: rolDe(p.clave)
  });
}

/* ------------------------------------------------ escritura */
function doPost(e) {
  var d; try { d = JSON.parse(e.postData.contents); } catch (err) { return json({ error: 'cuerpo ilegible' }); }
  var rol = rolDe(d.clave);
  if (!rol) return json({ error: 'clave incorrecta' });

  if (d.accion === 'decidir') {
    if (rol !== 'editor' && rol !== 'editor2') return json({ error: 'tu rol solo lee' });
    if (!/^\d{4}-\d{2}-\d{2}$/.test(d.fecha || '')) return json({ error: 'falta la fecha; no se guardó nada' });
    if (!d.envio_id) return json({ error: 'falta envio_id' });
    if (typeof d.propuestas !== 'object') return json({ error: 'revisión mal formada' });

    var DE = filas('DECISIONES');
    for (var i = 0; i < DE.length; i++)                     // idempotencia: mismo envio_id = ya esta
      if (String(DE[i].envio_id) === String(d.envio_id))
        return json({ ok: true, guardado: String(DE[i].guardado), repetido: true });

    var sello = ahora(), h = hoja('DECISIONES'), fs = [];
    Object.keys(d.propuestas).forEach(function (pid) {
      var p = d.propuestas[pid] || {};
      (p.laminas || []).forEach(function (m, i2) {
        // la nota POR LAMINA y las mejoras palomeadas viajan en la misma fila de
        // la marca: el agente lee exactamente que le cambio a cada foto tachada
        var notaLam = (p.notas && p.notas[i2]) || '';
        var mej = (p.mejoras && p.mejoras[i2] && p.mejoras[i2].length)
                  ? ' [mejoras elegidas: ' + p.mejoras[i2].join(',') + ']' : '';
        fs.push([d.envio_id, rol, sello, d.fecha, pid, i2, m || 'pendiente', notaLam + mej, '', '']);
      });
      fs.push([d.envio_id, rol, sello, d.fecha, pid, '', '', p.nota || '', '', '']);
    });
    fs.push([d.envio_id, rol, sello, d.fecha, '', '', '', '', d.nota_general || '', d.nota_estrategia || '']);
    h.getRange(h.getLastRow() + 1, 1, fs.length, 10).setValues(fs);
    bitacora('Revisión recibida de ' + rol, fs.length + ' registro(s) · envío ' + String(d.envio_id).slice(0, 8));
    return json({ ok: true, guardado: sello });
  }

  if (d.accion === 'parrilla_decision') {                   // el andon jalado desde la Mesa
    if (rol !== 'editor' && rol !== 'editor2') return json({ error: 'tu rol solo lee' });
    if (d.decision !== 'empujar' && d.decision !== 'matar') return json({ error: 'decisión desconocida' });
    var hp = hoja('PARRILLA'), datos = hp.getDataRange().getValues(), fila = -1;
    for (var j = 1; j < datos.length; j++)
      if (fechaDe(datos[j][0]) === d.dia) { fila = j + 1; break; }
    if (fila < 0) return json({ error: 'no encontré ese día en la parrilla' });
    hp.getRange(fila, 5, 1, 2).setValues([[d.decision, ahora()]]);
    bitacora('Andón: ' + (d.pieza || '') + ' → ' + d.decision + ' (' + rol + ')');
    return json({ ok: true });
  }

  if (d.accion === 'proponer') {                            // la Mac monta el dia
    if (rol !== 'agente' && rol !== 'editor') return json({ error: 'solo el agente propone' });
    var hpr = hoja('PROPUESTAS');
    (d.propuestas || []).forEach(function (p) {
      hpr.appendRow([d.fecha || hoy(), p.id, p.titulo, p.tipo || 'laminas',
                     JSON.stringify(p.laminas || []), JSON.stringify(p.opciones || []),
                     p.video || '', 'en revisión', p.origen || '']);
    });
    bitacora('Propuestas del día montadas', (d.propuestas || []).length + ' propuesta(s)');
    return json({ ok: true });
  }

  if (d.accion === 'parrilla') {                            // la Mac actualiza un hueco
    if (rol !== 'agente' && rol !== 'editor') return json({ error: 'solo el agente' });
    var hpa = hoja('PARRILLA'), dt = hpa.getDataRange().getValues(), f2 = -1;
    for (var k = 1; k < dt.length; k++) if (fechaDe(dt[k][0]) === d.dia) { f2 = k + 1; break; }
    var v = [d.dia, d.pieza || '', d.gate || '', d.desde || hoy(), '', ''];
    if (f2 > 0) hpa.getRange(f2, 1, 1, 6).setValues([v]); else hpa.appendRow(v);
    bitacora('Parrilla: ' + d.dia + ' → ' + (d.pieza || 'hueco'));
    return json({ ok: true });
  }

  if (d.accion === 'produccion') {                          // la Mac reporta avances
    if (rol !== 'agente' && rol !== 'editor') return json({ error: 'solo el agente' });
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
  var asunto = n > 0 ? 'Sala de Edición · ' + n + ' propuesta(s) esperan tu revisión'
                     : 'Sala de Edición · hoy no hay pendientes';
  // la liga trae la sesion en el fragmento #: no viaja al servidor y el portal la siembra solo
  var exec = ScriptApp.getService().getUrl();
  var liga = PORTAL + '#gas=' + encodeURIComponent(exec) +
             '&clave=' + encodeURIComponent(leerConfig('clave')) + '&rol=editor';
  var html =
    '<div style="background:#0a0a0c;padding:34px 22px;font-family:Georgia,serif;color:#f4f1ec">' +
    '<p style="color:#debc7e;font-size:12px;letter-spacing:3px;margin:0 0 6px">YO DESARROLLO</p>' +
    '<h2 style="font-weight:400;margin:0 0 16px">Sala de <em style="color:#debc7e">Edición</em></h2>' +
    '<p style="font-size:15px;line-height:1.6;margin:0 0 22px">' +
    (n > 0 ? 'Hay <b>' + n + ' propuesta(s)</b> esperándote. Cinco minutos.' : 'Hoy no hay propuestas nuevas.') + '</p>' +
    '<a href="' + liga + '" style="background:#c2a06b;color:#17130c;text-decoration:none;' +
    'padding:13px 26px;border-radius:8px;font-family:-apple-system,sans-serif;font-weight:600">Abrir mi turno</a>' +
    '<p style="color:#8a867e;font-size:12px;margin:26px 0 0">Palomea lo que sirve, tacha lo que no, ' +
    'y escribe lo que quieres distinto.</p></div>';
  MailApp.sendEmail({ to: CORREO, subject: asunto, htmlBody: html });
  bitacora('Correo de las 7:00 enviado', asunto);
}
