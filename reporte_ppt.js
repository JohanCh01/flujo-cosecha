/* Vista PPT del Reporte diario espárrago — usa TU PowerPoint como plantilla.

   lib/plantilla_reporte.pptx  tu PPT, con cada valor cambiado por una marca {{#Fnn}}
   lib/plantilla_campos.json   dónde va cada valor (posición, letra, color) y qué dato lleva, p. ej. "Prom: {{f.olmos.larva.prom}}"
   lib/fondos/sNN.jpg          cada diapositiva sin esos valores (fondo de la vista previa y del PDF)
   (se regeneran con herramientas/preparar_plantilla.py cuando cambie el diseño del PPT)

   Los datos salen de los MISMOS cálculos y filtros de reporte.html (window.RPT). Lo que no tiene fuente queda en blanco.
   La fecha del reporte es HOY; el corte es la semana elegida (por defecto, la semana anterior). */
(function () {
  'use strict';
  var R = window.RPT;
  if (!R) return;
  var $ = function (id) { return document.getElementById(id); };
  var PX = 96;                                     // 1 pulgada = 96 px → diapositiva de 1280 × 720
  var plantilla = null, abierto = false, actual = 0, editando = false, valores = {}, ed = {};

  // ---------- ediciones guardadas en este navegador (por semana de corte y fecha)
  function claveGuardado() { var s = R.state; return 'rpt.ppt2.' + s.anio + '.' + s.sem + '.' + (s.fecha || 'sem') + '.' + hoyYmd(); }
  function leerEd() { try { return JSON.parse(localStorage.getItem(claveGuardado()) || '{}') || {}; } catch (e) { return {}; } }
  function guardarEd() { try { localStorage.setItem(claveGuardado(), JSON.stringify(ed)); return true; } catch (e) { return false; } }

  // ---------- formatos (como en tu PPT)
  function pc(x, esp) { return x == null || isNaN(x) ? '' : (x * 100).toFixed(1) + (esp ? ' %' : '%'); }
  function mill(x) { return x >= 1e6 ? (x / 1e6).toFixed(1) + ' mill' : x >= 1e3 ? (x / 1e3).toFixed(1) + ' mil' : x ? Math.round(x).toLocaleString('es-PE') : '0'; }
  function n2(x) { return x.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function hoyYmd() { var d = new Date(Date.now() - 5 * 3600000); return d.getUTCFullYear() * 10000 + (d.getUTCMonth() + 1) * 100 + d.getUTCDate(); }
  function ddmm(ymd) { var s = String(ymd); return s.slice(6, 8) + '/' + s.slice(4, 6); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // etiquetas de cada caja fitosanitaria, en el orden de tu PPT
  var ETQ = {
    elasmo: [['ROJO2', 'Rojo2'], ['ROJO 2', 'Rojo2'], ['ROJO', 'Rojo'], ['AMARILLO', 'Amarillo'], ['VERDE', 'Verde'], ['AZUL', 'Azul']],
    posturas: [['[0.00 - 1.0]', '0 - 1'], ['[1.01 - 2.0]', '1.1 - 2'], ['[>2]', '> 2 ']],
    larva: [['0% - 5%', '0 – 5% '], ['5.1% - 10%', '5.1 – 10%'], ['> 10%', '> 10%']],
    prodi: [['0% - 5%', '0 – 5% '], ['5.1% - 10%', '5.1 – 10%'], ['> 10%', '> 10%']]
  };

  // ---------- valores de cada marca, con los filtros actuales
  function calcularValores() {
    var s = R.state, hoy = hoyYmd(), v = {};
    v.fecha = ddmm(hoy) + '/' + String(hoy).slice(0, 4);
    v.ddmm = ddmm(hoy);
    v.dia = R.DIAS[new Date(R.utc(hoy)).getUTCDay()] + '.';
    v.anio = String(s.anio);
    v.sem = String(s.sem);
    v.sem_ant = String(s.sem - 1);
    // kilos recepcionados (SQL)
    var rc = R.rec(), proy = R.PROYECTADO.kg;
    if (rc.hay) {
      v['rec.proy'] = proy ? mill(proy) : '';
      v['rec.eje'] = mill(rc.eje);
      v['rec.sem'] = mill(rc.per);
      v['rec.avance'] = proy ? (rc.eje / proy * 100).toFixed(1) + ' %' : '';
      var tot = rc.zonas.reduce(function (a, z) { return a + z.kg; }, 0);
      rc.zonas.forEach(function (z, i) {
        v['tz.' + (i + 1) + '.0'] = z.z; v['tz.' + (i + 1) + '.1'] = n2(z.kg); v['tz.' + (i + 1) + '.2'] = tot ? (z.kg / tot * 100).toFixed(2) + '%' : '';
      });
      v['tz.4.0'] = 'Total'; v['tz.4.1'] = n2(tot); v['tz.4.2'] = tot ? '100.00%' : '';
    }
    // fitosanitarios y calidades por zona
    R.ZONAS.forEach(function (Z) {
      var z = Z.toLowerCase();
      R.fito(Z).forEach(function (f) {
        var id = f.ind.id, res = f.res, lineas = [];
        v['f.' + z + '.' + id + '.prom'] = pc(res.prom);
        v['f.' + z + '.' + id + '.max'] = pc(res.max);
        if (res.total && !f.falta) {
          var usadas = {};
          ETQ[id].forEach(function (e) {
            var k = e[0], w = res.peso[k] || 0;
            if (usadas[e[1]] || !w) return;
            usadas[e[1]] = 1;
            lineas.push(e[1] + ': ' + (w / res.total * 100).toFixed(1) + (id === 'elasmo' ? ' %' : '%'));
          });
          Object.keys(res.peso).forEach(function (k) {            // categorías que no estén en la lista (otro color)
            if (!ETQ[id].some(function (e) { return e[0] === k; }) && res.peso[k]) lineas.push(k + ': ' + (res.peso[k] / res.total * 100).toFixed(1) + ' %');
          });
        }
        v['f.' + z + '.' + id + '.lineas'] = lineas.join('\n');
      });
      var c = R.cal(Z);
      [['sm', 'sm'], ['lo', 'lo'], ['cr', 'cr']].forEach(function (q) {
        v['c.' + z + '.' + q[0] + '.acu'] = pc(c.acu[q[1]], true);
        v['c.' + z + '.' + q[0] + '.eje'] = c.per[q[1]] == null ? '' : ' ' + pc(c.per[q[1]]);
      });
    });
    return v;
  }
  function textoCampo(p) {
    if (ed[p.id] != null) return ed[p.id];
    return p.t.replace(/\{\{([^}#]+)\}\}/g, function (_, k) { return valores[k] != null ? valores[k] : ''; });
  }

  // ---------- dibujo de una diapositiva (fondo de tu PPT + valores encima)
  function htmlSlide(n, editable) {
    var h = '<div class="slide' + (editable ? ' edit' : '') + '"><img class="bg" alt="" src="lib/fondos/s' + (n < 10 ? '0' : '') + n + '.jpg">';
    plantilla.campos.forEach(function (c) {
      if (c.s !== n) return;
      var b = c.box, st = 'left:' + b[0] * PX + 'px;top:' + b[1] * PX + 'px;width:' + b[2] * PX + 'px;height:' + b[3] * PX + 'px;' +
        'padding:' + c.ins[1] * PX + 'px ' + c.ins[2] * PX + 'px ' + c.ins[3] * PX + 'px ' + c.ins[0] * PX + 'px;justify-content:' + ({ middle: 'center', bottom: 'flex-end' }[c.anchor] || 'flex-start');
      h += '<div class="cp" style="' + st + '">';
      c.p.forEach(function (p) {
        h += '<div class="pr" data-f="' + p.id + '"' + (editable ? ' contenteditable="true" spellcheck="false"' : '') + ' style="font-size:' + p.sz * PX / 72 + 'px;color:#' + p.c +
          ';text-align:' + p.al + ';font-family:\'' + p.font + '\',Calibri,Carlito,Arial,sans-serif;' + (p.b ? 'font-weight:700;' : '') + (p.i ? 'font-style:italic;' : '') + (p.u ? 'text-decoration:underline;' : '') + '">' + esc(textoCampo(p)) + '</div>';
      });
      h += '</div>';
    });
    return h + '</div>';
  }

  // ---------- panel
  var panel = $('pptPanel'), split = $('split'), btn = $('btnPpt');
  panel.innerHTML =
    '<div class="ppt-bar"><strong>Vista PPT</strong><span class="ppt-n" id="pptN"></span><span class="ppt-sp"></span>' +
      '<button class="btn sm" id="pptPrev" type="button" aria-label="Diapositiva anterior">‹</button>' +
      '<button class="btn sm" id="pptNext" type="button" aria-label="Diapositiva siguiente">›</button>' +
      '<button class="btn sm" id="pptCerrar" type="button" aria-label="Cerrar la vista PPT">✕</button></div>' +
    '<div class="ppt-view" id="pptView"></div>' +
    '<div class="ppt-bar">' +
      '<button class="btn sm" id="pptEdit" type="button" aria-pressed="false">Editar</button>' +
      '<button class="btn sm" id="pptGuardar" type="button" hidden>Guardar cambios</button>' +
      '<button class="btn sm" id="pptReset" type="button" title="Vuelve a los valores calculados de los datos">Descartar cambios</button>' +
      '<span class="ppt-sp"></span>' +
      '<button class="btn sm" id="pptPdf" type="button">Descargar PDF</button>' +
      '<button class="btn sm primary" id="pptPptx" type="button">Descargar PowerPoint</button></div>' +
    '<p class="ppt-msg" id="pptMsg" aria-live="polite"></p>' +
    '<div class="ppt-thumbs" id="pptThumbs"></div>';
  var css = document.createElement('style');
  css.textContent =
    '.ppt{position:sticky;top:0;height:100vh;overflow:auto;background:var(--surface);border-right:1px solid var(--line);padding:12px 14px;display:flex;flex-direction:column;gap:10px;z-index:5}' +
    '.ppt-bar{display:flex;align-items:center;gap:6px;flex-wrap:wrap}.ppt-sp{flex:1}.ppt-n{color:var(--muted);font-size:12.5px;margin-left:6px}' +
    '.btn.sm{height:30px;padding:0 10px;font-size:12.5px}.btn[aria-pressed="true"]{background:var(--accent-soft);border-color:var(--accent);color:var(--accent)}' +
    '.ppt-view{position:relative;width:100%;aspect-ratio:16/9;overflow:hidden;border:1px solid var(--line);border-radius:6px;background:#fff;box-shadow:0 2px 10px rgba(0,0,0,.08);flex:none}' +
    '.slide{position:absolute;left:0;top:0;width:1280px;height:720px;transform-origin:0 0;background:#fff;overflow:hidden}' +
    '.slide .bg{position:absolute;left:0;top:0;width:1280px;height:720px;pointer-events:none}' +
    '.slide .cp{position:absolute;display:flex;flex-direction:column;box-sizing:border-box;overflow:visible}' +
    '.slide .pr{white-space:pre;line-height:1.2;min-height:1em}' +
    '.slide.edit .pr{outline:1px dashed #5b8def;outline-offset:1px;cursor:text;min-width:20px}.slide.edit .pr:focus{outline:2px solid #2f6fe0;background:rgba(47,111,224,.08)}' +
    '.ppt-msg{margin:0;font-size:12.5px;color:var(--muted);min-height:16px}' +
    '.ppt-thumbs{display:grid;grid-template-columns:repeat(auto-fill,minmax(118px,1fr));gap:8px}' +
    '.ppt-th{position:relative;aspect-ratio:16/9;border:1px solid var(--line);border-radius:4px;overflow:hidden;background:#fff;cursor:pointer;padding:0}' +
    '.ppt-th[aria-current="true"]{outline:2px solid var(--accent);outline-offset:1px}.ppt-th span{position:absolute;left:3px;bottom:2px;font-size:10px;background:rgba(255,255,255,.85);color:#333;padding:0 4px;border-radius:3px}';
  document.head.appendChild(css);

  function msg(t) { $('pptMsg').textContent = t || ''; }
  function escalar() {
    var v = $('pptView'), sl = v && v.querySelector('.slide');
    if (sl) sl.style.transform = 'scale(' + v.clientWidth / 1280 + ')';
    Array.prototype.forEach.call(document.querySelectorAll('.ppt-th'), function (th) { var s2 = th.querySelector('.slide'); if (s2) s2.style.transform = 'scale(' + th.clientWidth / 1280 + ')'; });
  }
  function miniaturas() {
    var h = '';
    for (var i = 1; i <= plantilla.n; i++) h += '<button class="ppt-th" type="button" data-i="' + i + '"' + (i === actual ? ' aria-current="true"' : '') + ' aria-label="Diapositiva ' + i + '">' + htmlSlide(i, false) + '<span>' + i + '</span></button>';
    $('pptThumbs').innerHTML = h; escalar();
  }
  function pintar() {
    if (!abierto || !plantilla) return;
    actual = Math.max(1, Math.min(actual || 1, plantilla.n));
    $('pptView').innerHTML = htmlSlide(actual, editando);
    $('pptN').textContent = actual + ' / ' + plantilla.n + ' · corte SEM ' + R.state.sem + ' · reporte del ' + valores.fecha;
    $('pptPrev').disabled = actual === 1; $('pptNext').disabled = actual === plantilla.n;
    miniaturas();
  }
  function refrescar() { if (!abierto || !plantilla) return; ed = leerEd(); valores = calcularValores(); pintar(); }

  function cargarPlantilla() {
    if (plantilla) return Promise.resolve();
    return fetch('lib/plantilla_campos.json', { cache: 'no-cache' }).then(function (r) { if (!r.ok) throw new Error('sin plantilla'); return r.json(); }).then(function (j) { plantilla = j; });
  }
  function abrir(on) {
    abierto = on; panel.hidden = !on; split.classList.toggle('on', on); btn.setAttribute('aria-pressed', String(on));
    if (on) {
      msg('Cargando la plantilla…');
      cargarPlantilla().then(function () { refrescar(); msg('Fecha del reporte: hoy. Corte: la semana elegida a la derecha (por defecto, la semana anterior).'); })
        .catch(function () { msg('No se pudo cargar la plantilla (lib/plantilla_campos.json).'); });
    }
    setTimeout(escalar, 30);
  }
  btn.addEventListener('click', function () { abrir(!abierto); });
  $('pptCerrar').onclick = function () { abrir(false); };
  $('pptPrev').onclick = function () { if (actual > 1) { actual--; pintar(); } };
  $('pptNext').onclick = function () { if (actual < plantilla.n) { actual++; pintar(); } };
  $('pptThumbs').addEventListener('click', function (e) { var b = e.target.closest('.ppt-th'); if (b) { actual = Number(b.getAttribute('data-i')); pintar(); } });
  document.addEventListener('keydown', function (e) {
    if (!abierto || editando || !plantilla || /INPUT|SELECT|TEXTAREA/.test(e.target.tagName || '')) return;
    if (e.key === 'ArrowRight' && actual < plantilla.n) { actual++; pintar(); }
    if (e.key === 'ArrowLeft' && actual > 1) { actual--; pintar(); }
  });
  window.addEventListener('resize', escalar);

  // ---------- edición
  $('pptEdit').onclick = function () {
    editando = !editando; this.setAttribute('aria-pressed', String(editando)); this.textContent = editando ? 'Terminar edición' : 'Editar';
    $('pptGuardar').hidden = !editando; pintar();
    msg(editando ? 'Haz clic en cualquier valor o texto marcado para cambiarlo (también los que están en blanco). Se guarda en este navegador.' : '');
  };
  $('pptView').addEventListener('input', function (e) {
    var el = e.target.closest('[data-f]'); if (!el) return;
    ed[el.getAttribute('data-f')] = el.innerText.replace(/\n$/, '');
    guardarEd(); msg('Cambios guardados en este navegador.');
  });
  $('pptView').addEventListener('focusout', miniaturas);
  $('pptGuardar').onclick = function () { msg(guardarEd() ? 'Cambios guardados en este navegador.' : 'Este navegador no deja guardar (modo privado): los cambios se pierden al cerrar.'); };
  $('pptReset').onclick = function () {
    if (!Object.keys(ed).length) { msg('No hay cambios que descartar.'); return; }
    if (!confirm('¿Descartar todos los cambios de este reporte y volver a los valores calculados?')) return;
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
  function nombreArchivo(ext) { return 'Reporte_ESPARRAGO_' + valores.fecha.replace(/\//g, '-') + '.' + ext; }
  function xmlEsc(s) { return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
  function llenarXml(xml, textos) {
    // primero los valores de varias líneas: se repite el párrafo una vez por línea
    xml = xml.replace(/<a:p>(?:(?!<\/a:p>)[\s\S])*?\{\{#(F\d+)\}\}[\s\S]*?<\/a:p>/g, function (par, id) {
      var t = textos[id] != null ? textos[id] : '';
      return t.split('\n').map(function (linea) { return par.replace('{{#' + id + '}}', xmlEsc(linea)); }).join('');
    });
    return xml.replace(/\{\{#(F\d+)\}\}/g, function (_, id) { return xmlEsc(textos[id] != null ? textos[id] : ''); });
  }
  $('pptPptx').onclick = function () {
    var b = this; b.disabled = true; b.textContent = 'Generando…';
    var textos = {};
    plantilla.campos.forEach(function (c) { c.p.forEach(function (p) { textos[p.id] = textoCampo(p); }); });
    Promise.all([cargar('lib/jszip.min.js'), fetch('lib/plantilla_reporte.pptx').then(function (r) { if (!r.ok) throw new Error('sin plantilla PPT'); return r.arrayBuffer(); })])
      .then(function (res) { return window.JSZip.loadAsync(res[1]); })
      .then(function (zip) {
        var nombres = Object.keys(zip.files).filter(function (f) { return /^ppt\/slides\/slide\d+\.xml$/.test(f); });
        return Promise.all(nombres.map(function (f) { return zip.file(f).async('string').then(function (x) { zip.file(f, llenarXml(x, textos)); }); }))
          .then(function () { return zip.generateAsync({ type: 'blob', mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', compression: 'DEFLATE' }); });
      })
      .then(function (blob) {
        var a = document.createElement('a'); a.href = URL.createObjectURL(blob); a.download = nombreArchivo('pptx');
        document.body.appendChild(a); a.click(); setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
        msg('PowerPoint descargado: es tu misma plantilla con los valores puestos (editable).');
      })
      .catch(function (e) { msg('No se pudo generar el PowerPoint: ' + e.message); })
      .then(function () { b.disabled = false; b.textContent = 'Descargar PowerPoint'; });
  };
  $('pptPdf').onclick = function () {
    var b = this; b.disabled = true; b.textContent = 'Generando…';
    var caja = document.createElement('div');
    caja.style.cssText = 'position:fixed;left:-20000px;top:0;width:1280px;height:720px;';
    document.body.appendChild(caja);
    Promise.all([cargar('lib/html2canvas.min.js'), cargar('lib/jspdf.umd.min.js')]).then(function () {
      var W = plantilla.w, H = plantilla.h, pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'in', format: [W, H] }), cadena = Promise.resolve();
      for (var i = 1; i <= plantilla.n; i++) (function (i) {
        cadena = cadena.then(function () {
          caja.innerHTML = htmlSlide(i, false);
          msg('Generando PDF… ' + i + ' de ' + plantilla.n);
          var img = caja.querySelector('img');
          return (img.complete ? Promise.resolve() : new Promise(function (ok) { img.onload = img.onerror = ok; }))
            .then(function () { return window.html2canvas(caja.firstChild, { scale: 1.5, backgroundColor: '#ffffff', logging: false, useCORS: true }); })
            .then(function (cv) { if (i > 1) pdf.addPage([W, H], 'landscape'); pdf.addImage(cv.toDataURL('image/jpeg', 0.9), 'JPEG', 0, 0, W, H); });
        });
      })(i);
      return cadena.then(function () { pdf.save(nombreArchivo('pdf')); });
    }).then(function () { msg('PDF descargado.'); })
      .catch(function (e) { msg('No se pudo generar el PDF: ' + e.message); })
      .then(function () { caja.remove(); b.disabled = false; b.textContent = 'Descargar PDF'; });
  };

  window.RPT_PPT = { refrescar: refrescar };
})();
