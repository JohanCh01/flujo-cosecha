/* Vista PPT del Reporte diario espárrago.
   Arma las diapositivas con los MISMOS cálculos y filtros de reporte.html (window.RPT), las muestra
   a la izquierda, deja editar los textos (se guardan en este navegador) y descarga en PowerPoint y PDF.

   Cada diapositiva es una lista de elementos con posición en pulgadas (lienzo 13.333 × 7.5, como el PPT):
     { t: 'rect', x, y, w, h, fill, line, r }                 rectángulo (r = esquinas redondeadas)
     { t: 'text', k, x, y, w, h, txt, pt, b, i, u, c, al, va, fill } texto (k = clave para editarlo)
     { t: 'table', k, x, y, w, colW, rowH, pt, rows, head }    tabla (cada celda se puede editar)
   El mismo modelo se dibuja en HTML (vista previa y PDF) y en PowerPoint (pptxgenjs), así salen iguales. */
(function () {
  'use strict';
  var R = window.RPT;
  if (!R) return;
  var $ = function (id) { return document.getElementById(id); };
  var W = 13.333, H = 7.5, PX = 96;                 // 1 pulgada = 96 px en la vista HTML (1280 × 720)
  var VERDE = '1E7B34', GUINDA = '8B1E3F', GRIS = '595959', TINTA = '262626', VALOR = '3BA55C', ROJO = 'C0392B';
  var BANDA_A = 'C9EFEF', BANDA_B = 'D9EAD3', BANDA_C = 'E2EFDA';
  var abierto = false, actual = 0, editando = false, slides = [];

  // ---------- ediciones guardadas en este navegador (por año, semana y fecha)
  function claveGuardado() { var s = R.state; return 'rpt.ppt.' + s.anio + '.' + s.sem + '.' + (s.fecha || 'sem'); }
  function leerEd() { try { return JSON.parse(localStorage.getItem(claveGuardado()) || '{}') || {}; } catch (e) { return {}; } }
  function guardarEd(ed) { try { localStorage.setItem(claveGuardado(), JSON.stringify(ed)); return true; } catch (e) { return false; } }
  var ed = {};

  // ---------- formatos
  function p1(x) { return x == null || isNaN(x) ? '—' : (x * 100).toFixed(1) + ' %'; }
  function mill(x) { return x >= 1e6 ? (x / 1e6).toFixed(1) + ' mill' : x >= 1e3 ? (x / 1e3).toFixed(1) + ' mil' : Math.round(x).toLocaleString('es-PE'); }
  function n2(x) { return x.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function hoyYmd() { var d = new Date(Date.now() - 5 * 3600000); return d.getUTCFullYear() * 10000 + (d.getUTCMonth() + 1) * 100 + d.getUTCDate(); }
  // fecha del reporte: la elegida; si no, el día siguiente al último día de la semana (como tu PPT del 12/10 para la SEM 41), sin pasar de hoy
  function fechaReporte() {
    var s = R.state; if (s.fecha) return s.fecha;
    var dias = []; ['rows', 'cal', 'lon', 'rec'].forEach(function (k) { (s[k] || []).forEach(function (r) { if (r.anio === s.anio && r.sem === s.sem) dias.push(r.ymd); }); });
    if (!dias.length) return hoyYmd();
    var ult = Math.max.apply(null, dias), d = new Date(R.utc(ult)), lunes = new Date(d.getTime() + ((7 - (d.getUTCDay() + 6) % 7) % 7 || 7) * 86400000);
    var sig = lunes.getUTCFullYear() * 10000 + (lunes.getUTCMonth() + 1) * 100 + lunes.getUTCDate();
    return Math.min(sig, hoyYmd());
  }
  function fechaLarga(ymd) { var s = String(ymd); return s.slice(6, 8) + '/' + s.slice(4, 6) + '/' + s.slice(0, 4); }

  // ---------- piezas comunes
  function titulo(txt) {
    return [{ t: 'text', k: 'tit', x: 0.45, y: 0.28, w: 12.4, h: 0.55, txt: txt, pt: 20, b: true, i: true, c: VERDE },
      { t: 'rect', x: 0.45, y: 0.86, w: 12.45, h: 0.025, fill: GUINDA }];
  }
  function pie(n) {
    return [{ t: 'text', x: 0.45, y: 7.08, w: 6, h: 0.3, txt: 'Aseguramiento de la Calidad Fresco', pt: 9, c: GRIS, i: true },
      { t: 'text', x: 11.9, y: 7.08, w: 1, h: 0.3, txt: String(n), pt: 9, c: GRIS, al: 'right' }];
  }
  function banda(k, x, y, w, txt, color) { return { t: 'text', k: k, x: x, y: y, w: w, h: 0.36, txt: txt, pt: 12, b: true, c: TINTA, al: 'center', va: 'middle', fill: color || BANDA_A }; }
  function nota(k, x, y, w, txt) { return { t: 'text', k: k, x: x, y: y, w: w, h: 0.26, txt: txt, pt: 8.5, c: GRIS, i: true }; }
  // caja con etiqueta arriba, valor grande y texto chico abajo
  function caja(k, x, y, w, h, etiqueta, valor, abajo, color) {
    return [{ t: 'rect', x: x, y: y, w: w, h: h, fill: 'FFFFFF', line: '7F7F7F', r: true },
      { t: 'text', k: k + '.l', x: x + 0.08, y: y + 0.08, w: w - 0.16, h: 0.32, txt: etiqueta, pt: 11, u: true, c: TINTA, al: 'center' },
      { t: 'text', k: k + '.v', x: x + 0.08, y: y + 0.42, w: w - 0.16, h: h - 0.82, txt: valor, pt: 30, c: color || VALOR, al: 'center', va: 'middle' },
      { t: 'text', k: k + '.s', x: x + 0.08, y: y + h - 0.4, w: w - 0.16, h: 0.32, txt: abajo, pt: 11, c: TINTA, al: 'center' }];
  }
  // barra apilada al 100 % (como en tu Power BI), dibujada con rectángulos
  function barra(x, y, w, h, segs) {
    var out = [{ t: 'rect', x: x, y: y, w: w, h: h, fill: 'EEEEEE' }], tot = 0, acc = 0;
    segs.forEach(function (s) { tot += s.v; });
    if (!tot) return out;
    segs.forEach(function (s) {
      if (!s.v) return;
      var hh = h * s.v / tot, yy = y + h - acc - hh;
      out.push({ t: 'rect', x: x, y: yy, w: w, h: hh, fill: s.color });
      if (hh >= 0.24) out.push({ t: 'text', x: x, y: yy, w: w, h: hh, txt: (s.v / tot * 100).toFixed(1) + '%', pt: 8, b: true, c: TINTA, al: 'center', va: 'middle' });
      acc += hh;
    });
    return out;
  }
  var CSS_A_HEX = { 'var(--c-azul)': '6D7FD1', 'var(--c-verde)': '9FD77C', 'var(--c-amarillo)': 'F3E675', 'var(--c-rojo)': 'F07B73', 'var(--c-rojo2)': 'D33A33', 'var(--c-naranja)': 'EAA73C', 'var(--c-otro)': 'B8C0B9' };

  // ---------- modelo de diapositivas con los datos filtrados
  function modelo() {
    var s = R.state, sem = s.sem, fechaRep = fechaReporte(), cortes = s.fecha ? R.ddmm(s.fecha) : 'SEM. ' + sem;
    var semTxt = s.fecha ? 'del ' + R.ddmm(s.fecha) : 'semana ' + sem;
    var out = [], n = 0;
    function nueva(id, nombre) { var sl = { id: id, nombre: nombre, el: [] }; out.push(sl); n++; return sl; }

    // 1. Portada
    var p = nueva('portada', 'Portada');
    p.el.push({ t: 'rect', x: 0, y: 0, w: 0.35, h: H, fill: GUINDA },
      { t: 'rect', x: 1.2, y: 2.35, w: 0.08, h: 1.9, fill: GUINDA },
      { t: 'text', k: 't1', x: 1.45, y: 2.3, w: 10.5, h: 0.75, txt: 'REPORTE DIARIO ESPÁRRAGO (' + fechaLarga(fechaRep) + ')', pt: 30, b: true, c: TINTA },
      { t: 'text', k: 't2', x: 1.45, y: 3.05, w: 10, h: 0.6, txt: 'AÑO ' + s.anio, pt: 26, b: true, c: TINTA },
      { t: 'text', k: 't3', x: 1.45, y: 3.75, w: 10, h: 0.7, txt: 'Aseguramiento de la Calidad\nFresco', pt: 15, b: true, c: GRIS });

    // 2. Separador avance de campaña
    var sp = nueva('sep1', 'Avance de campaña');
    sp.el.push({ t: 'rect', x: 0, y: 0, w: W, h: H, fill: 'F4F6F2' }, { t: 'rect', x: 0, y: 3.2, w: W, h: 0.05, fill: GUINDA },
      { t: 'text', k: 't1', x: 0.8, y: 2.2, w: 11.7, h: 0.8, txt: 'AVANCE DE CAMPAÑA ESPÁRRAGO - ' + s.anio, pt: 32, b: true, c: VERDE, al: 'center' },
      { t: 'text', k: 't2', x: 0.8, y: 3.45, w: 11.7, h: 0.6, txt: 'CORTE HASTA ' + cortes, pt: 22, c: TINTA, al: 'center' },
      { t: 'text', k: 't3', x: 0.8, y: 6.6, w: 11.7, h: 0.4, txt: 'Aseguramiento de la Calidad Fresco', pt: 12, i: true, c: GRIS, al: 'center' });

    // 3. Kilos recepcionados y drenados (diapositiva 4 de tu PPT)
    var rc = R.rec(), k = nueva('kilos', 'Kilos recepcionados');
    var proy = R.PROYECTADO.kg, hasta = s.fecha ? R.ddmm(s.fecha) : 'S.' + sem;
    k.el = k.el.concat(titulo('2. Kilos de Espárrago Recepcionado y Drenado'),
      [banda('b1', 0.45, 1.1, 12.45, 'KILOGRAMOS RECEPCIONADOS', BANDA_A)],
      caja('c1', 0.9, 1.65, 2.6, 1.5, 'PROYECTADOS', mill(proy), 'kilos', '1F3864'),
      caja('c2', 3.85, 1.65, 2.6, 1.5, 'EJECUTADOS HASTA ' + hasta, rc.hay ? mill(rc.eje) : '—', 'kilos'),
      caja('c3', 6.8, 1.65, 2.6, 1.5, 'AVANCE DE CAMPAÑA', rc.hay && proy ? p1(rc.eje / proy) : '—', ''),
      caja('c4', 9.75, 1.65, 2.6, 1.5, s.fecha ? 'SOLO EL ' + R.ddmm(s.fecha) : 'SOLO SEMANA ' + sem, rc.hay ? mill(rc.per) : '—', 'kilos'),
      [nota('n1', 0.6, 3.22, 8, '*' + R.PROYECTADO.rev + '   *Información de avance de campaña hasta corte ' + (s.fecha ? 'del ' + R.ddmm(s.fecha) : 'semana ' + sem)),
        banda('b2', 0.45, 3.65, 12.45, 'KILOGRAMOS DRENADOS', BANDA_B)],
      caja('d1', 0.9, 4.2, 2.6, 1.5, 'PROYECTADOS', '—', 'kilos', '1F3864'),
      caja('d2', 3.85, 4.2, 2.6, 1.5, 'EJECUTADOS HASTA ' + hasta, '—', 'kilos'),
      caja('d3', 6.8, 4.2, 2.6, 1.5, 'AVANCE DE CAMPAÑA', '—', ''),
      caja('d4', 9.75, 4.2, 2.6, 1.5, s.fecha ? 'SOLO EL ' + R.ddmm(s.fecha) : 'SOLO SEMANA ' + sem, '—', 'kilos'),
      [nota('n2', 0.6, 5.78, 11, 'Drenados: todavía sin base conectada; escribe los valores con «Editar».')],
      pie(n));

    // 4. Separador indicadores de la semana
    var s2 = nueva('sep2', 'Indicadores de la semana');
    s2.el.push({ t: 'rect', x: 0, y: 0, w: W, h: H, fill: 'F4F6F2' }, { t: 'rect', x: 0, y: 3.2, w: W, h: 0.05, fill: GUINDA },
      { t: 'text', k: 't1', x: 0.8, y: 2.2, w: 11.7, h: 0.8, txt: 'INDICADORES DE LA ' + (s.fecha ? 'FECHA ' + fechaLarga(s.fecha) : 'SEMANA ' + sem), pt: 32, b: true, c: VERDE, al: 'center' },
      { t: 'text', k: 't3', x: 0.8, y: 6.6, w: 11.7, h: 0.4, txt: 'Aseguramiento de la Calidad Fresco', pt: 12, i: true, c: GRIS, al: 'center' });

    // 5. Recepción de espárrago (diapositiva 8)
    var re = nueva('recepcion', 'Recepción de espárrago');
    var tot = rc.zonas.reduce(function (a, z) { return a + z.kg; }, 0);
    re.el = re.el.concat(titulo('5. Recepción de Espárrago'),
      [banda('b1', 0.45, 1.15, 6.0, 'KILOS RECEPCIONADOS – ' + (s.fecha ? R.ddmm(s.fecha) : 'Sem. ' + sem), BANDA_C),
        banda('b2', 6.85, 1.15, 6.0, 'KILOS EN STOCK – ' + R.ddmm(fechaRep), BANDA_C)],
      caja('c1', 1.6, 1.75, 3.7, 1.45, 'KILOS RECEPCIONADOS', rc.hay ? mill(rc.per) : '—', '', TINTA),
      caja('c2', 8.0, 1.75, 3.7, 1.45, 'KILOS EN STOCK', '0', '', TINTA),
      [{ t: 'table', k: 'tz', x: 1.3, y: 3.5, w: 4.3, colW: [1.5, 1.6, 1.2], rowH: 0.3, pt: 10, head: true,
        rows: [['Zona', 'Kg Recepcionados', '% Kg Recep']].concat(rc.zonas.map(function (z) { return [z.z, n2(z.kg), tot ? (z.kg / tot * 100).toFixed(2) + '%' : '—']; }),
          [['Total', n2(tot), tot ? '100.00%' : '—']]) }],
      pie(n));

    // 6–8. Indicadores de calidad MMPP por zona (diapositivas 9, 10 y 11)
    R.ZONAS.forEach(function (z) {
      var nom = R.NOM_ZONA[z].toUpperCase(), sl = nueva('mmpp_' + z.toLowerCase(), 'Calidad MMPP · ' + R.NOM_ZONA[z]);
      sl.el = sl.el.concat(titulo('2. Indicadores de Calidad MMPP  - ' + nom), [banda('b1', 0.45, 1.05, 12.45, 'FITOSANITARIOS', BANDA_A)]);
      var nombres = { elasmo: 'SEMAFORIZACIÓN ELASMOPALPUS', posturas: 'POSTURAS', larva: 'LEPIDOPTERA', prodi: 'PRODIPLOSIS' };
      R.fito(z).forEach(function (f, i) {
        var x = 0.45 + i * 3.15, y = 1.55, w = 3.0, hh = 2.35, segs = [], lineas = [];
        f.cats.forEach(function (c) { var v = f.res.peso[c.k] || 0; segs.push({ v: v, color: CSS_A_HEX[c.css] || 'B8C0B9' }); });
        f.cats.slice().reverse().forEach(function (c) { var v = f.res.peso[c.k] || 0; if (f.res.total && v) lineas.push(c.k + ': ' + (v / f.res.total * 100).toFixed(1) + ' %'); });
        sl.el.push({ t: 'rect', x: x, y: y, w: w, h: hh, fill: 'FFFFFF', line: '7F7F7F', r: true },
          { t: 'text', k: 'f' + i + '.t', x: x + 0.1, y: y + 0.06, w: w - 0.2, h: 0.3, txt: nombres[f.ind.id], pt: 9.5, b: true, u: true, c: TINTA, al: 'center' });
        sl.el = sl.el.concat(barra(x + 0.15, y + 0.45, 0.55, hh - 0.6, segs));
        sl.el.push({ t: 'text', k: 'f' + i + '.c', x: x + 0.8, y: y + 0.45, w: w - 0.9, h: hh - 1.15, txt: f.falta ? 'Sin columna en la base' : (lineas.join('\n') || 'Sin datos'), pt: 10, c: TINTA, va: 'top' },
          { t: 'text', k: 'f' + i + '.p', x: x + 0.8, y: y + hh - 0.68, w: w - 0.9, h: 0.55, txt: 'Prom: ' + p1(f.res.prom) + '\nMax: ' + p1(f.res.max), pt: 10, b: true, c: TINTA, va: 'top' });
      });
      sl.el.push(nota('n1', 0.45, 3.95, 9, '*Información de avance de campaña: ' + semTxt), banda('b2', 0.45, 4.35, 12.45, 'CALIDADES Y CALIBRES', BANDA_B));
      var cal = R.cal(z), KP = [['SMALL', 'sm'], ['Long >30', 'lo'], ['C- Y Ramas', 'cr']];
      KP.forEach(function (q, i) {
        var x = 1.6 + i * 3.6, y = 4.85, w = 3.0, hh = 1.75;
        sl.el.push({ t: 'rect', x: x, y: y, w: w, h: hh, fill: 'FFFFFF', line: '7F7F7F', r: true },
          { t: 'text', k: 'k' + i + '.t', x: x + 0.1, y: y + 0.08, w: w - 0.2, h: 0.32, txt: q[0], pt: 11, b: true, u: true, c: TINTA, al: 'center' },
          { t: 'text', k: 'k' + i + '.a', x: x + 0.1, y: y + 0.5, w: w - 0.2, h: 0.5, txt: 'Acumulado:  ' + p1(cal.acu[q[1]]), pt: 13, c: GRIS, al: 'center', va: 'middle' },
          { t: 'text', k: 'k' + i + '.e', x: x + 0.1, y: y + 1.05, w: w - 0.2, h: 0.55, txt: (s.fecha ? 'Ejec. ' + R.ddmm(s.fecha) : 'Ejec. Sem' + sem) + ':  ' + p1(cal.per[q[1]]), pt: 15, b: true, c: VALOR, al: 'center', va: 'middle' });
      });
      sl.el = sl.el.concat([nota('n2', 0.45, 6.7, 9, '*Información de avance de campaña hasta corte ' + (s.fecha ? 'del ' + R.ddmm(s.fecha) : 'semana de ' + sem))], pie(n));
    });
    return out;
  }

  // ---------- dibujo HTML (vista previa y PDF)
  function hex(c) { return '#' + c; }
  function textoDe(sl, e) { var k = sl.id + '.' + e.k; return e.k && ed[k] != null ? ed[k] : e.txt; }
  function htmlSlide(sl, editable) {
    var h = '<div class="slide' + (editable ? ' edit' : '') + '" data-slide="' + sl.id + '">';
    sl.el.forEach(function (e) {
      var st = 'left:' + e.x * PX + 'px;top:' + e.y * PX + 'px;width:' + e.w * PX + 'px;';
      if (e.t === 'rect') {
        h += '<div class="sh" style="' + st + 'height:' + e.h * PX + 'px;background:' + (e.fill ? hex(e.fill) : 'transparent') + ';' + (e.line ? 'border:1.3px solid ' + hex(e.line) + ';' : '') + (e.r ? 'border-radius:12px;' : '') + '"></div>';
      } else if (e.t === 'text') {
        st += 'height:' + e.h * PX + 'px;font-size:' + e.pt * PX / 72 + 'px;color:' + hex(e.c || TINTA) + ';text-align:' + (e.al || 'left') + ';justify-content:' +
          ({ middle: 'center', bottom: 'flex-end' }[e.va] || 'flex-start') + ';' + (e.b ? 'font-weight:700;' : '') + (e.i ? 'font-style:italic;' : '') + (e.u ? 'text-decoration:underline;' : '') + (e.fill ? 'background:' + hex(e.fill) + ';' : '');
        var t = textoDe(sl, e);
        h += '<div class="tx" style="' + st + '"' + (e.k ? ' data-k="' + sl.id + '.' + e.k + '"' + (editable ? ' contenteditable="true" spellcheck="false"' : '') : '') + '>' + esc(t) + '</div>';
      } else if (e.t === 'table') {
        h += '<table class="tb" style="' + st + 'font-size:' + e.pt * PX / 72 + 'px">';
        e.rows.forEach(function (row, r) {
          h += '<tr style="height:' + e.rowH * PX + 'px">';
          row.forEach(function (cel, c) {
            var k = sl.id + '.' + e.k + '.' + r + '.' + c, tag = e.head && r === 0 ? 'th' : 'td', ult = r === e.rows.length - 1;
            h += '<' + tag + ' style="width:' + e.colW[c] * PX + 'px;' + (ult ? 'font-weight:700;' : '') + (c ? 'text-align:right;' : '') + '" data-k="' + k + '"' + (editable ? ' contenteditable="true" spellcheck="false"' : '') + '>' + esc(ed[k] != null ? ed[k] : cel) + '</' + tag + '>';
          });
          h += '</tr>';
        });
        h += '</table>';
      }
    });
    return h + '</div>';
  }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // ---------- panel
  var panel = $('pptPanel'), split = $('split'), btn = $('btnPpt');
  panel.innerHTML =
    '<div class="ppt-bar">' +
      '<strong>Vista PPT</strong><span class="ppt-n" id="pptN"></span>' +
      '<span class="ppt-sp"></span>' +
      '<button class="btn sm" id="pptPrev" type="button" aria-label="Diapositiva anterior">‹</button>' +
      '<button class="btn sm" id="pptNext" type="button" aria-label="Diapositiva siguiente">›</button>' +
      '<button class="btn sm" id="pptCerrar" type="button" aria-label="Cerrar la vista PPT">✕</button>' +
    '</div>' +
    '<div class="ppt-view" id="pptView"></div>' +
    '<div class="ppt-bar">' +
      '<button class="btn sm" id="pptEdit" type="button" aria-pressed="false">Editar</button>' +
      '<button class="btn sm" id="pptGuardar" type="button" hidden>Guardar cambios</button>' +
      '<button class="btn sm" id="pptReset" type="button" title="Vuelve a los valores calculados de los datos">Descartar cambios</button>' +
      '<span class="ppt-sp"></span>' +
      '<button class="btn sm" id="pptPdf" type="button">Descargar PDF</button>' +
      '<button class="btn sm primary" id="pptPptx" type="button">Descargar PowerPoint</button>' +
    '</div>' +
    '<p class="ppt-msg" id="pptMsg" aria-live="polite"></p>' +
    '<div class="ppt-thumbs" id="pptThumbs"></div>';

  var css = document.createElement('style');
  css.textContent =
    '.ppt{position:sticky;top:0;height:100vh;overflow:auto;background:var(--surface);border-right:1px solid var(--line);padding:12px 14px;display:flex;flex-direction:column;gap:10px;z-index:5}' +
    '.ppt-bar{display:flex;align-items:center;gap:6px;flex-wrap:wrap}.ppt-sp{flex:1}.ppt-n{color:var(--muted);font-size:12.5px;margin-left:6px}' +
    '.btn.sm{height:30px;padding:0 10px;font-size:12.5px}.btn[aria-pressed="true"]{background:var(--accent-soft);border-color:var(--accent);color:var(--accent)}' +
    '.ppt-view{position:relative;width:100%;aspect-ratio:16/9;overflow:hidden;border:1px solid var(--line);border-radius:6px;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.08)}' +
    '.slide{position:absolute;left:0;top:0;width:1280px;height:720px;transform-origin:0 0;background:#fff;color:#262626;font-family:Calibri,"Segoe UI",Arial,sans-serif;overflow:hidden}' +
    '.slide .sh{position:absolute}.slide .tx{position:absolute;display:flex;flex-direction:column;white-space:pre-line;line-height:1.18;padding:2px 4px;box-sizing:border-box;overflow:hidden}' +
    '.slide .tb{position:absolute;border-collapse:collapse;table-layout:fixed}.slide .tb th,.slide .tb td{border:1px solid #BFBFBF;padding:2px 6px;white-space:nowrap;overflow:hidden}.slide .tb th{background:#F2F2F2;font-weight:700}' +
    '.slide.edit [data-k]{outline:1px dashed #5b8def;outline-offset:-1px;cursor:text}.slide.edit [data-k]:focus{outline:2px solid #2f6fe0;background:rgba(47,111,224,.06)}' +
    '.ppt-msg{margin:0;font-size:12.5px;color:var(--muted);min-height:16px}' +
    '.ppt-thumbs{display:grid;grid-template-columns:repeat(auto-fill,minmax(118px,1fr));gap:8px}' +
    '.ppt-th{position:relative;aspect-ratio:16/9;border:1px solid var(--line);border-radius:4px;overflow:hidden;background:#fff;cursor:pointer;padding:0}' +
    '.ppt-th[aria-current="true"]{outline:2px solid var(--accent);outline-offset:1px}.ppt-th span{position:absolute;left:3px;bottom:2px;font-size:10px;background:rgba(255,255,255,.85);color:#333;padding:0 4px;border-radius:3px}';
  document.head.appendChild(css);

  function escalar() {
    var v = $('pptView'); if (!v) return;
    var k = v.clientWidth / 1280, sl = v.querySelector('.slide'); if (sl) sl.style.transform = 'scale(' + k + ')';
    Array.prototype.forEach.call(document.querySelectorAll('.ppt-th'), function (th) { var s2 = th.querySelector('.slide'); if (s2) s2.style.transform = 'scale(' + th.clientWidth / 1280 + ')'; });
  }
  function pintar() {
    if (!abierto) return;
    if (actual >= slides.length) actual = slides.length - 1;
    var sl = slides[actual];
    $('pptView').innerHTML = htmlSlide(sl, editando);
    $('pptN').textContent = (actual + 1) + ' / ' + slides.length + ' · ' + sl.nombre;
    $('pptThumbs').innerHTML = slides.map(function (x, i) { return '<button class="ppt-th" type="button" data-i="' + i + '"' + (i === actual ? ' aria-current="true"' : '') + ' aria-label="' + esc(x.nombre) + '">' + htmlSlide(x, false) + '<span>' + (i + 1) + '</span></button>'; }).join('');
    $('pptPrev').disabled = actual === 0; $('pptNext').disabled = actual === slides.length - 1;
    escalar();
  }
  function refrescar() { if (!abierto) return; ed = leerEd(); slides = modelo(); pintar(); }
  function msg(t) { $('pptMsg').textContent = t || ''; }

  function abrir(on) {
    abierto = on; panel.hidden = !on; split.classList.toggle('on', on); btn.setAttribute('aria-pressed', String(on));
    if (on) { refrescar(); msg('Las diapositivas usan los filtros de la derecha (año, semana, fecha y status).'); }
    setTimeout(escalar, 30);
  }
  btn.addEventListener('click', function () { abrir(!abierto); });
  $('pptCerrar').onclick = function () { abrir(false); };
  $('pptPrev').onclick = function () { if (actual > 0) { actual--; pintar(); } };
  $('pptNext').onclick = function () { if (actual < slides.length - 1) { actual++; pintar(); } };
  $('pptThumbs').addEventListener('click', function (e) { var b = e.target.closest('.ppt-th'); if (b) { actual = Number(b.getAttribute('data-i')); pintar(); } });
  document.addEventListener('keydown', function (e) {
    if (!abierto || editando || /INPUT|SELECT|TEXTAREA/.test((e.target.tagName || ''))) return;
    if (e.key === 'ArrowRight' && actual < slides.length - 1) { actual++; pintar(); }
    if (e.key === 'ArrowLeft' && actual > 0) { actual--; pintar(); }
  });
  window.addEventListener('resize', escalar);

  // ---------- edición
  $('pptEdit').onclick = function () {
    editando = !editando; this.setAttribute('aria-pressed', String(editando)); this.textContent = editando ? 'Terminar edición' : 'Editar';
    $('pptGuardar').hidden = !editando; pintar();
    msg(editando ? 'Haz clic en cualquier texto, número o celda para cambiarlo. Se guarda en este navegador para esta semana/fecha.' : '');
  };
  $('pptView').addEventListener('input', function (e) {
    var el = e.target.closest('[data-k]'); if (!el) return;
    ed[el.getAttribute('data-k')] = el.innerText.replace(/\n$/, '');
    guardarEd(ed); msg('Cambios guardados en este navegador.');
  });
  $('pptView').addEventListener('focusout', function () { $('pptThumbs').innerHTML = slides.map(function (x, i) { return '<button class="ppt-th" type="button" data-i="' + i + '"' + (i === actual ? ' aria-current="true"' : '') + '>' + htmlSlide(x, false) + '<span>' + (i + 1) + '</span></button>'; }).join(''); escalar(); });
  $('pptGuardar').onclick = function () { msg(guardarEd(ed) ? 'Cambios guardados en este navegador.' : 'Este navegador no deja guardar (modo privado): los cambios se pierden al cerrar.'); };
  $('pptReset').onclick = function () {
    if (!Object.keys(ed).length) { msg('No hay cambios que descartar.'); return; }
    if (!confirm('¿Descartar todos los cambios de esta semana/fecha y volver a los valores calculados?')) return;
    ed = {}; try { localStorage.removeItem(claveGuardado()); } catch (e) { /* sin almacenamiento */ }
    pintar(); msg('Se descartaron los cambios.');
  };

  // ---------- descargas (las librerías se cargan solo al usarlas)
  function cargar(src) {
    return new Promise(function (ok, mal) {
      if (document.querySelector('script[data-lib="' + src + '"]')) return ok();
      var s = document.createElement('script'); s.src = src; s.setAttribute('data-lib', src); s.onload = ok; s.onerror = function () { mal(new Error('No se pudo cargar ' + src)); };
      document.head.appendChild(s);
    });
  }
  function nombreArchivo(ext) { var s = R.state; return 'Reporte_ESPARRAGO_' + (s.fecha ? fechaLarga(s.fecha).replace(/\//g, '-') : s.anio + '_SEM' + s.sem) + '.' + ext; }
  function ocupado(b, on, txt) { b.disabled = on; if (txt) b.textContent = txt; }

  $('pptPptx').onclick = function () {
    var b = this; ocupado(b, true, 'Generando…');
    cargar('lib/pptxgen.bundle.js').then(function () {
      var pp = new window.PptxGenJS(); pp.layout = 'LAYOUT_WIDE'; pp.title = 'Reporte diario espárrago';
      slides.forEach(function (sl) {
        var d = pp.addSlide();
        sl.el.forEach(function (e) {
          if (e.t === 'rect') {
            d.addShape(e.r ? pp.ShapeType.roundRect : pp.ShapeType.rect, { x: e.x, y: e.y, w: e.w, h: e.h, rectRadius: e.r ? 0.12 : 0,
              fill: e.fill ? { color: e.fill } : { color: 'FFFFFF', transparency: 100 }, line: e.line ? { color: e.line, width: 1 } : { type: 'none' } });
          } else if (e.t === 'text') {
            d.addText(textoDe(sl, e), { x: e.x, y: e.y, w: e.w, h: e.h, fontFace: 'Calibri', fontSize: e.pt, bold: !!e.b, italic: !!e.i, underline: e.u ? { style: 'sng' } : undefined,
              color: e.c || TINTA, align: e.al || 'left', valign: e.va || 'top', margin: 2, fill: e.fill ? { color: e.fill } : undefined, fit: 'shrink' });
          } else if (e.t === 'table') {
            var filas = e.rows.map(function (row, r) {
              return row.map(function (cel, c) {
                var k = sl.id + '.' + e.k + '.' + r + '.' + c;
                return { text: ed[k] != null ? ed[k] : cel, options: { bold: (e.head && r === 0) || r === e.rows.length - 1, align: c ? 'right' : 'left', fill: e.head && r === 0 ? { color: 'F2F2F2' } : undefined } };
              });
            });
            d.addTable(filas, { x: e.x, y: e.y, w: e.w, colW: e.colW, rowH: e.rowH, fontFace: 'Calibri', fontSize: e.pt, color: TINTA, border: { type: 'solid', color: 'BFBFBF', pt: 0.75 } });
          }
        });
      });
      return pp.writeFile({ fileName: nombreArchivo('pptx') });
    }).then(function () { msg('PowerPoint descargado (editable).'); })
      .catch(function (e) { msg('No se pudo generar el PowerPoint: ' + e.message); })
      .then(function () { ocupado(b, false, 'Descargar PowerPoint'); });
  };

  $('pptPdf').onclick = function () {
    var b = this; ocupado(b, true, 'Generando…');
    var caja = document.createElement('div');
    caja.style.cssText = 'position:fixed;left:-20000px;top:0;width:1280px;height:720px;';
    document.body.appendChild(caja);
    Promise.all([cargar('lib/html2canvas.min.js'), cargar('lib/jspdf.umd.min.js')]).then(function () {
      var pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'in', format: [W, H] }), cadena = Promise.resolve();
      slides.forEach(function (sl, i) {
        cadena = cadena.then(function () {
          caja.innerHTML = htmlSlide(sl, false);
          msg('Generando PDF… ' + (i + 1) + ' de ' + slides.length);
          return window.html2canvas(caja.firstChild, { scale: 2, backgroundColor: '#ffffff', logging: false }).then(function (cv) {
            if (i) pdf.addPage([W, H], 'landscape');
            pdf.addImage(cv.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, W, H);
          });
        });
      });
      return cadena.then(function () { pdf.save(nombreArchivo('pdf')); });
    }).then(function () { msg('PDF descargado.'); })
      .catch(function (e) { msg('No se pudo generar el PDF: ' + e.message); })
      .then(function () { caja.remove(); ocupado(b, false, 'Descargar PDF'); });
  };

  window.RPT_PPT = { refrescar: refrescar };
})();
