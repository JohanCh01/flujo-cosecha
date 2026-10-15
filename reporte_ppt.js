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
  var V = '202610151241';                                   // versión: obliga al navegador a bajar los archivos nuevos
  var PX = 96;                                     // 1 pulgada = 96 px → diapositiva de 1280 × 720
  var plantilla = null, abierto = false, actual = 0, editando = false, valores = {}, ed = { t: {}, f: {}, m: {} };

  // ---------- ediciones guardadas en este navegador (por semana de corte y fecha)
  function claveGuardado() { var s = R.state; return 'rpt.ppt3.' + s.anio + '.' + s.sem + '.' + (s.fecha || 'sem') + '.' + hoyYmd(); }
  // ed.t = textos cambiados, ed.f = formato por párrafo {sz, b, i, font}, ed.m = cuánto se movió cada cuadro [dx, dy] en pulgadas
  function leerEd() {
    var o = {}; try { o = JSON.parse(localStorage.getItem(claveGuardado()) || '{}') || {}; } catch (e) { o = {}; }
    return { t: o.t || {}, f: o.f || {}, m: o.m || {} };
  }
  function guardarEd() { try { localStorage.setItem(claveGuardado(), JSON.stringify(ed)); return true; } catch (e) { return false; } }

  // ---------- formatos (como en tu PPT)
  function pc(x, esp) { return x == null || isNaN(x) ? '' : (x * 100).toFixed(1) + (esp ? ' %' : '%'); }
  function mill(x) { return x >= 1e6 ? (x / 1e6).toFixed(1) + ' mill' : x >= 1e3 ? (x / 1e3).toFixed(1) + ' mil' : x ? Math.round(x).toLocaleString('es-PE') : '0'; }
  function n2(x) { return x.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }
  function hoyYmd() { var d = new Date(Date.now() - 5 * 3600000); return d.getUTCFullYear() * 10000 + (d.getUTCMonth() + 1) * 100 + d.getUTCDate(); }
  function ddmm(ymd) { var s = String(ymd); return s.slice(6, 8) + '/' + s.slice(4, 6); }
  function esc(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]; }); }

  // etiquetas de cada caja fitosanitaria (como en tu PPT) y colores de la barra (como en tu Power BI)
  var ETQ = {
    posturas: { '[0.00 - 1.0]': '0 - 1', '[1.01 - 2.0]': '1.1 - 2', '[>2]': '> 2 ' },
    larva: { '0% - 5%': '0 – 5% ', '5.1% - 10%': '5.1 – 10%', '> 10%': '> 10%' },
    prodi: { '0% - 5%': '0 – 5% ', '5.1% - 10%': '5.1 – 10%', '> 10%': '> 10%' }
  };
  var HEX = { 'var(--c-azul)': '6D7FD1', 'var(--c-verde)': '9FD77C', 'var(--c-amarillo)': 'F3E675', 'var(--c-rojo)': 'F07B73', 'var(--c-rojo2)': 'D33A33', 'var(--c-naranja)': 'EAA73C', 'var(--c-otro)': 'B8C0B9' };
  function etiqueta(id, k) {
    if (id === 'elasmo') { var t = k.replace(/\s+/g, '').toLowerCase(); return t.charAt(0).toUpperCase() + t.slice(1); }
    return ETQ[id][k] || k;
  }
  var barras = {};                                  // 'z.ind' → segmentos de abajo hacia arriba [{v, c}]

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
        var id = f.ind.id, res = f.res, lineas = [], colores = [], segs = [];
        v['f.' + z + '.' + id + '.prom'] = pc(res.prom);
        v['f.' + z + '.' + id + '.max'] = pc(res.max);
        if (res.total && !f.falta) {
          // barra: de abajo hacia arriba en el orden de las categorías; las líneas van de arriba hacia abajo (igual que la barra)
          f.cats.forEach(function (c) { var w = res.peso[c.k] || 0; if (w) segs.push({ v: w / res.total, c: HEX[c.css] || 'B8C0B9' }); });
          f.cats.slice().reverse().forEach(function (c) {
            var w = res.peso[c.k] || 0; if (!w) return;
            lineas.push(etiqueta(id, c.k) + ': ' + (w / res.total * 100).toFixed(1) + (id === 'elasmo' ? ' %' : '%'));
            colores.push(HEX[c.css] || 'B8C0B9');
          });
        }
        v['f.' + z + '.' + id + '.lineas'] = lineas.join('\n');
        v['f.' + z + '.' + id + '.colores'] = colores;
        barras[z + '.' + id] = segs;
      });
      var c = R.cal(Z), kgRec = (rc.zonas || []).filter(function (x) { return x.z === Z; }).reduce(function (a, x) { return a + x.kg; }, 0);
      v['hay.' + z] = kgRec > 0 || (c.per && c.per.kg > 0) || R.fito(Z).some(function (f) { return f.res && f.res.n > 0; });
      [['sm', 'sm'], ['lo', 'lo'], ['cr', 'cr']].forEach(function (q) {
        v['c.' + z + '.' + q[0] + '.acu'] = pc(c.acu[q[1]], true);
        v['c.' + z + '.' + q[0] + '.eje'] = c.per[q[1]] == null ? '' : ' ' + pc(c.per[q[1]]);
      });
    });
    return v;
  }
  // diapositivas que van: si una zona (Compositan, Olmos, Terceros) no tiene datos en el periodo, su diapositiva no se incluye
  function zonaDeSlide(n) { var b = (plantilla.barras || []).filter(function (x) { return x.s === n; })[0]; return b ? b.z : null; }
  function activas() {
    var out = [];
    for (var i = 1; i <= plantilla.n; i++) { var z = zonaDeSlide(i); if (!z || valores['hay.' + z]) out.push(i); }
    return out;
  }
  function textoCampo(p) {
    if (ed.t[p.id] != null) return ed.t[p.id];
    return p.t.replace(/\{\{([^}#]+)\}\}/g, function (_, k) { return valores[k] != null ? valores[k] : ''; });
  }

  var SUSTITUTA = { 'Abadi ExtraLight': 'AbadiWeb', 'Century Gothic': 'GothicWeb', 'Calibri': 'Carlito', 'Candara': 'Carlito', 'Aptos Light': 'AbadiWeb', 'Aptos': 'Carlito' };
  var FUENTES = ['Abadi ExtraLight', 'Century Gothic', 'Calibri', 'Candara', 'Arial', 'Aptos Light', 'Times New Roman', 'Verdana', 'Tahoma'];
  function familia(f) { return "'" + f + "'" + (SUSTITUTA[f] ? ",'" + SUSTITUTA[f] + "'" : '') + ",'Carlito',Arial,sans-serif"; }
  function fmt(p) { var o = ed.f[p.id] || {}; return { sz: o.sz || p.sz, b: o.b != null ? o.b : p.b, i: o.i != null ? o.i : p.i, font: o.font || p.font }; }

  // ---------- dibujo de una diapositiva (fondo de tu PPT + todos los textos encima)
  function htmlSlide(n, editable) {
    var h = '<div class="slide' + (editable ? ' edit' : '') + '"><img class="bg" alt="" src="lib/fondos/s' + (n < 10 ? '0' : '') + n + '.jpg?v=' + V + '">';
    plantilla.campos.forEach(function (c, ci) {
      if (c.s !== n) return;
      var b = c.box, m = ed.m['C' + ci] || [0, 0];
      var st = 'left:' + (b[0] + m[0]) * PX + 'px;top:' + (b[1] + m[1]) * PX + 'px;width:' + b[2] * PX + 'px;height:' + b[3] * PX + 'px;' +
        'padding:' + c.ins[1] * PX + 'px ' + c.ins[2] * PX + 'px ' + c.ins[3] * PX + 'px ' + c.ins[0] * PX + 'px;justify-content:' + ({ middle: 'center', bottom: 'flex-end' }[c.anchor] || 'flex-start');
      h += '<div class="cp' + (c.wrap === 'none' ? ' nw' : '') + '" data-c="' + ci + '" style="' + st + '">';
      c.p.forEach(function (p) {
        var f = fmt(p), txt = textoCampo(p), corto = txt.length <= 28 && txt.indexOf('\n') < 0;   // los textos cortos no se parten en dos líneas
        var cuerpo = esc(txt);
        if (p.lineas) {                            // cuadrito de color delante de cada línea (mismo color que la barra)
          var cols = valores[p.lineas + '.colores'] || [];
          cuerpo = txt.split('\n').map(function (l, k) { return (cols[k] ? '<span class="sw" contenteditable="false" style="background:#' + cols[k] + '"></span>' : '') + esc(l); }).join('\n');
        }
        h += '<div class="pr' + (corto ? ' nw1' : '') + '" data-f="' + p.id + '"' + (editable ? ' contenteditable="true" spellcheck="false"' : '') + ' style="font-size:' + f.sz * PX / 72 + 'px;color:#' + p.c +
          ';text-align:' + p.al + ';font-family:' + familia(f.font) + ';font-weight:' + (f.b ? 700 : 400) + ';font-style:' + (f.i ? 'italic' : 'normal') + ';' + (p.u ? 'text-decoration:underline;' : '') + '">' + cuerpo + '</div>';
      });
      h += '</div>';
    });
    (plantilla.barras || []).forEach(function (bb) {
      if (bb.s !== n) return;
      var segs = barras[bb.z + '.' + bb.ind] || [], b = bb.box, acc = 0;
      h += '<div class="pbar" style="left:' + b[0] * PX + 'px;top:' + b[1] * PX + 'px;width:' + b[2] * PX + 'px;height:' + b[3] * PX + 'px">';
      segs.forEach(function (sg) {
        var hh = sg.v * b[3] * PX;
        h += '<div style="bottom:' + acc + 'px;height:' + hh + 'px;background:#' + sg.c + '">' + (hh >= 15 ? (sg.v * 100).toFixed(1) + '%' : '') + '</div>';
        acc += hh;
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
      '<span class="ppt-fmt" id="pptFmt" hidden>' +
        '<select id="fmtFont" class="sel sm" title="Fuente"></select>' +
        '<button class="btn sm" id="fmtMenos" type="button" title="Letra más chica">A−</button><span id="fmtSz" class="ppt-n"></span>' +
        '<button class="btn sm" id="fmtMas" type="button" title="Letra más grande">A+</button>' +
        '<button class="btn sm" id="fmtB" type="button" title="Negrita"><b>N</b></button>' +
        '<button class="btn sm" id="fmtI" type="button" title="Cursiva"><i>K</i></button>' +
        '<button class="btn sm" id="fmtOrig" type="button" title="Vuelve este cuadro a su formato y lugar original">Original</button>' +
      '</span>' +
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
    '.slide .pr{white-space:pre-wrap;line-height:1.2;min-height:1em}.slide .cp.nw .pr,.slide .pr.nw1{white-space:pre}' +
    '.slide.edit .cp.sel{outline:2px solid #e8912d;outline-offset:2px}' +
    '.slide .sw{display:inline-block;width:.62em;height:.62em;margin-right:.35em;vertical-align:.02em;border-radius:1px}' +
    '.slide .pbar{position:absolute}.slide .pbar>div{position:absolute;left:0;right:0;border-top:1px solid #fff;box-sizing:border-box;display:flex;align-items:center;justify-content:center;font:700 10.5px/1 Calibri,Carlito,Arial,sans-serif;color:#262626}' +
    '.slide .asa{position:absolute;width:26px;height:26px;margin:-30px 0 0 -30px;border-radius:6px;background:#e8912d;color:#fff;font:700 16px/26px Arial;text-align:center;cursor:move;user-select:none;touch-action:none;box-shadow:0 1px 4px rgba(0,0,0,.3)}' +
    '.ppt-fmt{display:inline-flex;gap:4px;align-items:center;flex-wrap:wrap}.sel.sm{height:30px;min-width:0;padding:0 24px 0 8px;font-size:12.5px}' +
"@font-face{font-family:AbadiWeb;font-weight:400;src:url(lib/fuentes/encode-sans-semi-condensed-latin-200-normal.woff2) format('woff2')}" +
    "@font-face{font-family:AbadiWeb;font-weight:700;src:url(lib/fuentes/encode-sans-semi-condensed-latin-500-normal.woff2) format('woff2')}" +
    "@font-face{font-family:GothicWeb;src:url(lib/fuentes/didact-gothic-latin-400-normal.woff2) format('woff2')}" +
    "@font-face{font-family:Carlito;font-weight:400;src:url(lib/fuentes/carlito-latin-400-normal.woff2) format('woff2')}" +
    "@font-face{font-family:Carlito;font-weight:700;src:url(lib/fuentes/carlito-latin-700-normal.woff2) format('woff2')}" +
    "@font-face{font-family:Carlito;font-weight:400;font-style:italic;src:url(lib/fuentes/carlito-latin-400-italic.woff2) format('woff2')}" +
    "@font-face{font-family:Carlito;font-weight:700;font-style:italic;src:url(lib/fuentes/carlito-latin-700-italic.woff2) format('woff2')}" +
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
    activas().forEach(function (i, k) { h += '<button class="ppt-th" type="button" data-i="' + i + '"' + (i === actual ? ' aria-current="true"' : '') + ' aria-label="Diapositiva ' + (k + 1) + '">' + htmlSlide(i, false) + '<span>' + (k + 1) + '</span></button>'; });
    $('pptThumbs').innerHTML = h; escalar();
  }
  function pintar() {
    if (!abierto || !plantilla) return;
    var lista = activas();
    if (lista.indexOf(actual) < 0) actual = lista.filter(function (x) { return x <= (actual || 1); }).pop() || lista[0];
    var pos = lista.indexOf(actual);
    $('pptView').innerHTML = htmlSlide(actual, editando);
    if (selC != null && plantilla.campos[selC] && plantilla.campos[selC].s !== actual) { selP = null; selC = null; }
    setTimeout(marcarSel, 0);
    var omitidas = plantilla.n - lista.length;
    $('pptN').textContent = (pos + 1) + ' / ' + lista.length + ' · corte SEM ' + R.state.sem + ' · reporte del ' + valores.fecha + (omitidas ? ' · ' + omitidas + ' zona(s) sin datos no se incluyen' : '');
    $('pptPrev').disabled = pos === 0; $('pptNext').disabled = pos === lista.length - 1;
    miniaturas();
  }
  function refrescar() { if (!abierto || !plantilla) return; ed = leerEd(); valores = calcularValores(); pintar(); }

  function cargarPlantilla() {
    if (plantilla) return Promise.resolve();
    return fetch('lib/plantilla_campos.json?v=' + V, { cache: 'no-cache' }).then(function (r) { if (!r.ok) throw new Error('sin plantilla'); return r.json(); }).then(function (j) { plantilla = j; });
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
  function paso(d) { var l = activas(), k = l.indexOf(actual) + d; if (k >= 0 && k < l.length) { actual = l[k]; pintar(); } }
  $('pptPrev').onclick = function () { paso(-1); };
  $('pptNext').onclick = function () { paso(1); };
  $('pptThumbs').addEventListener('click', function (e) { var b = e.target.closest('.ppt-th'); if (b) { actual = Number(b.getAttribute('data-i')); pintar(); } });
  document.addEventListener('keydown', function (e) {
    if (!abierto || editando || !plantilla || /INPUT|SELECT|TEXTAREA/.test(e.target.tagName || '')) return;
    if (e.key === 'ArrowRight') paso(1);
    if (e.key === 'ArrowLeft') paso(-1);
  });
  window.addEventListener('resize', escalar);

  // ---------- edición
  $('pptEdit').onclick = function () {
    editando = !editando; this.setAttribute('aria-pressed', String(editando)); this.textContent = editando ? 'Terminar edición' : 'Editar';
    $('pptGuardar').hidden = !editando; selP = null; selC = null; pintar(); marcarSel();
    msg(editando ? 'Haz clic en un texto para cambiarlo; arriba eliges fuente, tamaño, negrita y cursiva; arrastra el botón ✥ para moverlo. Se guarda en este navegador.' : '');
  };
  var selP = null, selC = null;                    // párrafo y cuadro elegidos
  function campoDe(id) { for (var i = 0; i < plantilla.campos.length; i++) for (var j = 0; j < plantilla.campos[i].p.length; j++) if (plantilla.campos[i].p[j].id === id) return plantilla.campos[i].p[j]; return null; }
  function marcarSel() {
    Array.prototype.forEach.call(document.querySelectorAll('#pptView .cp.sel'), function (x) { x.classList.remove('sel'); });
    var asa = document.querySelector('#pptView .asa'); if (asa) asa.remove();
    $('pptFmt').hidden = !editando || !selP;
    if (!editando || selC == null) return;
    var cp = document.querySelector('#pptView .cp[data-c="' + selC + '"]'); if (!cp) return;
    cp.classList.add('sel');
    if (!plantilla.campos[selC].tabla) {
      var a = document.createElement('div'); a.className = 'asa'; a.title = 'Arrastra para mover el cuadro'; a.textContent = '✥';
      a.style.left = cp.style.left; a.style.top = cp.style.top; cp.parentNode.appendChild(a);
    }
    var p = campoDe(selP); if (!p) return;
    var f = fmt(p);
    $('fmtSz').textContent = (Math.round(f.sz * 2) / 2) + ' pt';
    $('fmtB').setAttribute('aria-pressed', String(!!f.b)); $('fmtI').setAttribute('aria-pressed', String(!!f.i));
    var opciones = FUENTES.indexOf(f.font) < 0 ? [f.font].concat(FUENTES) : FUENTES;
    $('fmtFont').innerHTML = opciones.map(function (x) { return '<option' + (x === f.font ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join('');
  }
  $('pptView').addEventListener('focusin', function (e) {
    var el = e.target.closest('[data-f]'); if (!el) return;
    selP = el.getAttribute('data-f'); selC = Number(el.closest('.cp').getAttribute('data-c')); marcarSel();
  });
  $('pptView').addEventListener('input', function (e) {
    var el = e.target.closest('[data-f]'); if (!el) return;
    ed.t[el.getAttribute('data-f')] = el.innerText.replace(/\n$/, '');
    guardarEd(); msg('Cambios guardados en este navegador.');
  });
  function cambiarFmt(cambio) {
    if (!selP) return;
    var p = campoDe(selP), o = ed.f[selP] || (ed.f[selP] = {}), f = fmt(p);
    cambio(o, f); guardarEd();
    var foco = selP; pintar(); selP = foco; marcarSel();
    var el = document.querySelector('#pptView [data-f="' + foco + '"]'); if (el) { el.focus(); }
  }
  $('fmtMenos').onclick = function () { cambiarFmt(function (o, f) { o.sz = Math.max(6, f.sz - 1); }); };
  $('fmtMas').onclick = function () { cambiarFmt(function (o, f) { o.sz = Math.min(96, f.sz + 1); }); };
  $('fmtB').onclick = function () { cambiarFmt(function (o, f) { o.b = !f.b; }); };
  $('fmtI').onclick = function () { cambiarFmt(function (o, f) { o.i = !f.i; }); };
  $('fmtFont').onchange = function () { var v = this.value; cambiarFmt(function (o) { o.font = v; }); };
  $('fmtOrig').onclick = function () { if (!selP) return; var c = plantilla.campos[selC]; c.p.forEach(function (p) { delete ed.f[p.id]; }); delete ed.m['C' + selC]; cambiarFmt(function () {}); };
  // mover: se arrastra el botón ✥ que aparece junto al cuadro elegido
  $('pptView').addEventListener('pointerdown', function (e) {
    var a = e.target.closest('.asa'); if (!a || selC == null) return;
    e.preventDefault(); a.setPointerCapture(e.pointerId);
    var k = $('pptView').clientWidth / 1280, x0 = e.clientX, y0 = e.clientY, m0 = (ed.m['C' + selC] || [0, 0]).slice();
    var cp = document.querySelector('#pptView .cp[data-c="' + selC + '"]'), b = plantilla.campos[selC].box;
    function mover(ev) {
      var dx = m0[0] + (ev.clientX - x0) / k / PX, dy = m0[1] + (ev.clientY - y0) / k / PX;
      cp.style.left = (b[0] + dx) * PX + 'px'; cp.style.top = (b[1] + dy) * PX + 'px'; a.style.left = cp.style.left; a.style.top = cp.style.top;
      ed.m['C' + selC] = [Math.round(dx * 1000) / 1000, Math.round(dy * 1000) / 1000];
    }
    function soltar() { a.removeEventListener('pointermove', mover); a.removeEventListener('pointerup', soltar); guardarEd(); miniaturas(); msg('Cuadro movido (guardado en este navegador).'); }
    a.addEventListener('pointermove', mover); a.addEventListener('pointerup', soltar);
  });
  $('pptView').addEventListener('focusout', function () { setTimeout(miniaturas, 0); });
  $('pptGuardar').onclick = function () { msg(guardarEd() ? 'Cambios guardados en este navegador.' : 'Este navegador no deja guardar (modo privado): los cambios se pierden al cerrar.'); };
  $('pptReset').onclick = function () {
    if (!Object.keys(ed.t).length && !Object.keys(ed.f).length && !Object.keys(ed.m).length) { msg('No hay cambios que descartar.'); return; }
    if (!confirm('¿Descartar todos los cambios de este reporte y volver a los valores calculados?')) return;
    ed = { t: {}, f: {}, m: {} }; try { localStorage.removeItem(claveGuardado()); } catch (e) { /* sin almacenamiento */ }
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
  function ajustarRpr(rpr, f) {
    // rpr = etiqueta <a:rPr ...> o <a:rPr .../> completa con su contenido
    var abre = rpr.match(/^<a:rPr\b[^>]*?\/?>/)[0], cierra = /\/>$/.test(abre), resto = cierra ? '' : rpr.slice(abre.length, rpr.length - '</a:rPr>'.length);
    var at = abre.replace(/^<a:rPr\b/, '').replace(/\/?>$/, '');
    function poner(nom, val) { at = new RegExp('\\s' + nom + '="[^"]*"').test(at) ? at.replace(new RegExp('(\\s' + nom + '=")[^"]*"'), '$1' + val + '"') : at + ' ' + nom + '="' + val + '"'; }
    if (f.sz) poner('sz', Math.round(f.sz * 100)); if (f.b != null) poner('b', f.b ? 1 : 0); if (f.i != null) poner('i', f.i ? 1 : 0);
    if (f.font) {
      var lat = '<a:latin typeface="' + xmlEsc(f.font) + '"/>';
      if (/<a:latin\b[^>]*\/>/.test(resto)) resto = resto.replace(/<a:latin\b[^>]*\/>/, lat);
      else { var pos = resto.search(/<a:(ea|cs|sym|hlinkClick|hlinkMouseOver|rtl|extLst)\b/); resto = pos < 0 ? resto + lat : resto.slice(0, pos) + lat + resto.slice(pos); }
    }
    return '<a:rPr' + at + '>' + resto + '</a:rPr>';
  }
  function aplicarFormatos(xml) {
    Object.keys(ed.f).forEach(function (id) {
      var tok = '{{#' + id + '}}', i = xml.indexOf(tok); if (i < 0) return;
      var r0 = xml.lastIndexOf('<a:r>', i); if (r0 < 0) return;
      var seg = xml.slice(r0, i), m = seg.match(/<a:rPr\b[^>]*?\/>|<a:rPr\b[\s\S]*?<\/a:rPr>/), nuevo;
      if (m) { nuevo = seg.replace(m[0], ajustarRpr(m[0], ed.f[id])); }
      else { nuevo = seg.replace('<a:r>', '<a:r>' + ajustarRpr('<a:rPr lang="es-PE"/>', ed.f[id])); }
      xml = xml.slice(0, r0) + nuevo + xml.slice(i);
    });
    Object.keys(ed.m).forEach(function (cid) {
      var c = plantilla.campos[Number(cid.slice(1))]; if (!c || c.tabla) return;
      var tok = '{{#' + c.p[0].id + '}}', i = xml.indexOf(tok); if (i < 0) return;
      var s0 = Math.max(xml.lastIndexOf('<p:sp>', i), xml.lastIndexOf('<p:sp ', i)); if (s0 < 0) return;
      var trozo = xml.slice(s0, i), mm = trozo.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>/); if (!mm) return;
      var d = ed.m[cid], k = c.k || [1, 1];
      var nx = Math.round(Number(mm[1]) + d[0] * 914400 / k[0]), ny = Math.round(Number(mm[2]) + d[1] * 914400 / k[1]);
      xml = xml.slice(0, s0) + trozo.replace(mm[0], '<a:off x="' + nx + '" y="' + ny + '"/>') + xml.slice(i);
    });
    return xml;
  }
  function conColor(rpr, hex) {
    var abre = rpr.match(/^<a:rPr\b[^>]*?\/?>/)[0], cierra = /\/>$/.test(abre), resto = cierra ? '' : rpr.slice(abre.length, rpr.length - '</a:rPr>'.length);
    resto = resto.replace(/<a:solidFill>[\s\S]*?<\/a:solidFill>/, '');
    return abre.replace(/\/>$/, '>') + '<a:solidFill><a:srgbClr val="' + hex + '"/></a:solidFill>' + resto + '</a:rPr>';
  }
  function barrasXml(n) {
    var out = '', id = 5000;
    (plantilla.barras || []).forEach(function (bb) {
      if (bb.s !== n) return;
      var segs = barras[bb.z + '.' + bb.ind] || [], b = bb.box, acc = 0, E = 914400;
      segs.forEach(function (sg) {
        var hh = sg.v * b[3], y = b[1] + b[3] - acc - hh;
        out += '<p:sp><p:nvSpPr><p:cNvPr id="' + (id++) + '" name="Barra ' + bb.ind + '"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="' + Math.round(b[0] * E) + '" y="' + Math.round(y * E) + '"/><a:ext cx="' + Math.round(b[2] * E) + '" cy="' + Math.max(1, Math.round(hh * E)) + '"/></a:xfrm>' +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:srgbClr val="' + sg.c + '"/></a:solidFill><a:ln w="9525"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></p:spPr>' +
          '<p:txBody><a:bodyPr lIns="0" tIns="0" rIns="0" bIns="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/>' +
          (hh >= 0.16 ? '<a:r><a:rPr lang="es-PE" sz="800" b="1"><a:solidFill><a:srgbClr val="262626"/></a:solidFill><a:latin typeface="Calibri"/></a:rPr><a:t>' + (sg.v * 100).toFixed(1) + '%</a:t></a:r>' : '') +
          '<a:endParaRPr lang="es-PE" sz="800"/></a:p></p:txBody></p:sp>';
        acc += hh;
      });
    });
    return out;
  }
  function llenarXml(xml, textos, n) {
    xml = aplicarFormatos(xml);
    // líneas con cuadrito de color: un párrafo por línea, con un primer run "■ " del color de la barra
    plantilla.campos.forEach(function (c) {
      c.p.forEach(function (p) {
        if (!p.lineas || c.s !== n) return;
        var tok = '{{#' + p.id + '}}', re = new RegExp('<a:p>(?:(?!<\\/a:p>)[\\s\\S])*?' + tok.replace(/[{}#]/g, '\\$&') + '[\\s\\S]*?<\\/a:p>');
        var m = xml.match(re); if (!m) return;
        var par = m[0], cols = valores[p.lineas + '.colores'] || [], lineas = (textos[p.id] || '').split('\n');
        var nuevo = lineas.map(function (l, k) {
          var q = par.replace(tok, xmlEsc(l));
          if (!cols[k]) return q;
          return q.replace(/<a:r>(<a:rPr\b[^>]*?\/>|<a:rPr\b[\s\S]*?<\/a:rPr>)?<a:t>/, function (all, rpr) {
            return '<a:r>' + conColor(rpr || '<a:rPr lang="es-PE"/>', cols[k]) + '<a:t>■ </a:t></a:r><a:r>' + (rpr || '') + '<a:t>';
          });
        }).join('');
        xml = xml.replace(par, nuevo);
      });
    });
    var b = barrasXml(n); if (b) xml = xml.replace('</p:spTree>', b + '</p:spTree>');
    // primero los valores de varias líneas: se repite el párrafo una vez por línea
    xml = xml.replace(/<a:p>(?:(?!<\/a:p>)[\s\S])*?\{\{#(F\d+)\}\}[\s\S]*?<\/a:p>/g, function (par, id) {
      var t = textos[id] != null ? textos[id] : '';
      return t.split('\n').map(function (linea) { return par.replace('{{#' + id + '}}', xmlEsc(linea)); }).join('');
    });
    return xml.replace(/\{\{#(F\d+)\}\}/g, function (_, id) { return xmlEsc(textos[id] != null ? textos[id] : ''); });
  }
  // quita del PowerPoint las diapositivas que no van (zona sin datos): de la lista, de las relaciones y sus archivos
  function quitarSlides(zip, lista) {
    var arch = plantilla.archivos || {}, fuera = [];
    for (var i = 1; i <= plantilla.n; i++) if (lista.indexOf(i) < 0 && arch[i]) fuera.push(arch[i]);
    if (!fuera.length) return Promise.resolve();
    return Promise.all([zip.file('ppt/_rels/presentation.xml.rels').async('string'), zip.file('ppt/presentation.xml').async('string'), zip.file('[Content_Types].xml').async('string')])
      .then(function (r) {
        var rels = r[0], pres = r[1], tipos = r[2];
        fuera.forEach(function (f) {
          var dest = f.replace(/^ppt\//, ''), m = rels.match(new RegExp('<Relationship[^>]*Id="([^"]+)"[^>]*Target="' + dest.replace(/[.]/g, '\\.') + '"[^>]*/>'))
            || rels.match(new RegExp('<Relationship[^>]*Target="' + dest.replace(/[.]/g, '\\.') + '"[^>]*Id="([^"]+)"[^>]*/>'));
          if (m) { rels = rels.replace(m[0], ''); pres = pres.replace(new RegExp('<p:sldId\\b[^>]*r:id="' + m[1] + '"[^>]*/>'), ''); }
          tipos = tipos.replace(new RegExp('<Override[^>]*PartName="/' + f.replace(/[.]/g, '\\.') + '"[^>]*/>'), '');
          zip.remove(f); zip.remove(f.replace(/slides\/(slide\d+\.xml)$/, 'slides/_rels/$1.rels'));
        });
        zip.file('ppt/_rels/presentation.xml.rels', rels); zip.file('ppt/presentation.xml', pres); zip.file('[Content_Types].xml', tipos);
      });
  }
  $('pptPptx').onclick = function () {
    var b = this; b.disabled = true; b.textContent = 'Generando…';
    var textos = {};
    plantilla.campos.forEach(function (c) { c.p.forEach(function (p) { textos[p.id] = textoCampo(p); }); });
    Promise.all([cargar('lib/jszip.min.js'), fetch('lib/plantilla_reporte.pptx?v=' + V).then(function (r) { if (!r.ok) throw new Error('sin plantilla PPT'); return r.arrayBuffer(); })])
      .then(function (res) { return window.JSZip.loadAsync(res[1]); })
      .then(function (zip) {
        var arch = plantilla.archivos || {}, nombres = Object.keys(zip.files).filter(function (f) { return /^ppt\/slides\/slide\d+\.xml$/.test(f); });
        var num = {}; Object.keys(arch).forEach(function (k) { num[arch[k]] = Number(k); });
        return Promise.all(nombres.map(function (f) { return zip.file(f).async('string').then(function (x) { zip.file(f, llenarXml(x, textos, num[f] || Number(f.match(/slide(\d+)/)[1]))); }); }))
          .then(function () { return quitarSlides(zip, activas()); })
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
      var W = plantilla.w, H = plantilla.h, pdf = new window.jspdf.jsPDF({ orientation: 'landscape', unit: 'in', format: [W, H] }), cadena = Promise.resolve(), lista = activas();
      lista.forEach(function (i, k) {
        cadena = cadena.then(function () {
          caja.innerHTML = htmlSlide(i, false);
          msg('Generando PDF… ' + (k + 1) + ' de ' + lista.length);
          var img = caja.querySelector('img');
          return (img.complete ? Promise.resolve() : new Promise(function (ok) { img.onload = img.onerror = ok; }))
            .then(function () { return window.html2canvas(caja.firstChild, { scale: 1.5, backgroundColor: '#ffffff', logging: false, useCORS: true }); })
            .then(function (cv) { if (k > 0) pdf.addPage([W, H], 'landscape'); pdf.addImage(cv.toDataURL('image/jpeg', 0.9), 'JPEG', 0, 0, W, H); });
        });
      });
      return cadena.then(function () { pdf.save(nombreArchivo('pdf')); });
    }).then(function () { msg('PDF descargado.'); })
      .catch(function (e) { msg('No se pudo generar el PDF: ' + e.message); })
      .then(function () { caja.remove(); b.disabled = false; b.textContent = 'Descargar PDF'; });
  };

  window.RPT_PPT = { refrescar: refrescar };
})();
