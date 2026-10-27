/*
 * Stock de paletas — lo usa reporte.html (diapositiva 8, «Kilos en stock»): botones «Generar tabla» y «Pegar la última».
 *   StockPaletas.dibujar(datos)          → Promise<canvas> con la tabla al estilo de la tabla dinámica de Excel
 *   StockPaletas.leer(tok)                → Promise<{cab, datos}> con lo último guardado en paletas.json (descifrado)
 *   StockPaletas.pedir(tok, quien)        → Promise<id>: guarda solicitud_paletas.json, eso arranca la tarea en GitHub
 *   StockPaletas.esperar(tok, id, avance) → Promise<{cab, datos}> cuando la tarea termina (o error)
 * El descifrado usa la misma sesión (DNI) del Flujo de cosecha, guardada en el navegador como fc.ses.
 */
(function () {
  'use strict';
  var REPO = 'JohanCh01/flujo-cosecha';
  var te = new TextEncoder(), td = new TextDecoder();

  // ---------- grupos: Unidad Agrícola con sus categorías (la tabla ya viene ordenada)
  function grupos(datos) {
    var g = [], porUa = {};
    (datos ? datos.tabla : []).forEach(function (x) {
      if (!porUa[x.ua]) { porUa[x.ua] = { ua: x.ua, kg: 0, n: 0, cats: [] }; g.push(porUa[x.ua]); }
      porUa[x.ua].kg += x.kg; porUa[x.ua].n += x.n; porUa[x.ua].cats.push(x);
    });
    return g;
  }

  // ---------- imagen con el aspecto de la tabla dinámica de Excel (solo la tabla)
  var FUENTE = 'Carlito, Calibri, Arial, sans-serif';
  var fuentes = null;
  function cargarFuentes() {
    if (fuentes) return fuentes;
    fuentes = Promise.all([['400', 'lib/fuentes/carlito-latin-400-normal.woff2'], ['700', 'lib/fuentes/carlito-latin-700-normal.woff2']].map(function (f) {
      try { var ff = new FontFace('Carlito', 'url(' + f[1] + ')', { weight: f[0] }); return ff.load().then(function (x) { document.fonts.add(x); }).catch(function () {}); } catch (e) { return Promise.resolve(); }
    }));
    return fuentes;
  }
  function numXls(v) { var r = Math.round(v * 100) / 100; return String(r); }   // como «General» en Excel: 8493.4, 943
  function pintar(d) {
    var g = grupos(d);
    var C = { cab: '#8064A2', cabTxt: '#FFFFFF', proc: '#B1A0C7', ua: '#E4DFEC', subT: '#D9D9D9', totF: '#8064A2', totFTxt: '#FFFFFF',
      blanco: '#FFFFFF', linea: '#9C88B8', lineaSuave: '#CFC6DD', txt: '#000000' };
    var FILA = 20, PAD = 4, ESC = 2;
    var cv = document.createElement('canvas'), cx = cv.getContext('2d');
    function f(peso) { return (peso || 400) + ' 14.5px ' + FUENTE; }
    // anchos de columna según el texto más largo (como «Autoajustar» en Excel)
    cx.font = f(700);
    var w = [cx.measureText('⊟ FRESCO').width, cx.measureText('Unidad Agricola').width, cx.measureText('Categoria').width, cx.measureText('Suma de Peso Neto Paleta').width];
    g.forEach(function (x) {
      w[1] = Math.max(w[1], cx.measureText('⊟ ' + x.ua).width, cx.measureText('Total ' + x.ua).width);
      x.cats.forEach(function (k) { w[2] = Math.max(w[2], cx.measureText(k.cat).width); });
    });
    w[0] = Math.max(w[0], cx.measureText('Total FRESCO').width); w[3] = Math.max(w[3], cx.measureText(numXls(d.total)).width);
    w = w.map(function (v, i) { return Math.ceil(v + PAD * 2 + (i < 3 ? 22 : 8)); });
    var X = [0, w[0], w[0] + w[1], w[0] + w[1] + w[2], w[0] + w[1] + w[2] + w[3]], ANCHO = X[4];
    var nFilas = 1 + g.reduce(function (s, x) { return s + x.cats.length + 1; }, 0) + 2;
    var M = 14, ALTO = nFilas * FILA;
    cv.width = (ANCHO + M * 2) * ESC; cv.height = (ALTO + M * 2) * ESC;
    cx.scale(ESC, ESC); cx.fillStyle = C.blanco; cx.fillRect(0, 0, ANCHO + M * 2, ALTO + M * 2);
    cx.translate(M, M); cx.textBaseline = 'middle';

    function celda(x0, x1, y, fondo) { cx.fillStyle = fondo; cx.fillRect(X[x0], y, X[x1] - X[x0], FILA); }
    function texto(t, col, y, peso, color, der) {
      cx.font = f(peso); cx.fillStyle = color || C.txt; cx.textAlign = der ? 'right' : 'left';
      cx.fillText(t, der ? X[col + 1] - PAD : X[col] + PAD, y + FILA / 2 + 1);
    }
    function caja(x, y, color) {   // el «⊟» de Excel para contraer
      cx.strokeStyle = color; cx.lineWidth = 1; cx.strokeRect(x + .5, y + 5.5, 9, 9);
      cx.beginPath(); cx.moveTo(x + 2.5, y + 10); cx.lineTo(x + 7.5, y + 10); cx.stroke();
    }
    function hLinea(x0, x1, y, color) { cx.strokeStyle = color; cx.lineWidth = 1; cx.beginPath(); cx.moveTo(X[x0], y + .5); cx.lineTo(X[x1], y + .5); cx.stroke(); }
    function vLinea(x, y0, y1, color) { cx.strokeStyle = color; cx.lineWidth = 1; cx.beginPath(); cx.moveTo(X[x] - .5, y0); cx.lineTo(X[x] - .5, y1); cx.stroke(); }

    // encabezado
    var y = 0;
    celda(0, 4, y, C.cab);
    ['Proceso', 'Unidad Agricola', 'Categoria', 'Suma de Peso Neto Paleta'].forEach(function (t, i) { texto(t, i, y, 400, C.cabTxt); if (i) vLinea(i, y, y + FILA, C.blanco); });
    y += FILA;

    // cuerpo: Proceso (FRESCO) a la izquierda, zonas con sus categorías y el total de cada zona
    var yIni = y, yFin = y + g.reduce(function (s, x) { return s + x.cats.length + 1; }, 0) * FILA;
    cx.fillStyle = C.proc; cx.fillRect(X[0], yIni, w[0], yFin - yIni);
    caja(X[0] + PAD, yIni, '#FFFFFF'); cx.font = f(400); cx.fillStyle = '#FFFFFF'; cx.textAlign = 'left'; cx.fillText('FRESCO', X[0] + PAD + 14, yIni + FILA / 2 + 1);
    g.forEach(function (x) {
      var y0 = y;
      cx.fillStyle = C.ua; cx.fillRect(X[1], y0, w[1], x.cats.length * FILA);
      caja(X[1] + PAD, y0, '#555555'); cx.font = f(400); cx.fillStyle = C.txt; cx.textAlign = 'left'; cx.fillText(x.ua, X[1] + PAD + 14, y0 + FILA / 2 + 1);
      x.cats.forEach(function (k, i) {
        celda(2, 4, y, C.blanco);
        texto(k.cat, 2, y, 400); texto(numXls(k.kg), 3, y, 400, null, true);
        if (i) hLinea(2, 4, y, C.lineaSuave);
        y += FILA;
      });
      celda(1, 4, y, C.subT); hLinea(1, 4, y, C.linea); hLinea(1, 4, y + FILA, C.linea);
      texto('Total ' + x.ua, 1, y, 700); texto(numXls(x.kg), 3, y, 700, null, true);
      y += FILA;
      vLinea(2, y0, y - FILA, C.linea); vLinea(3, y0, y, C.linea);
    });
    vLinea(1, yIni, yFin, C.linea);

    // Total FRESCO y Total general
    celda(0, 4, y, C.totF); texto('Total FRESCO', 0, y, 700, C.totFTxt); texto(numXls(d.total), 3, y, 700, C.totFTxt, true);
    y += FILA;
    celda(0, 4, y, C.blanco); hLinea(0, 4, y, C.cab); hLinea(0, 4, y + FILA - 1, C.cab);
    texto('Total general', 0, y, 700); texto(numXls(d.total), 3, y, 700, null, true);
    vLinea(3, y, y + FILA, C.lineaSuave);
    return cv;
  }
  function dibujar(datos) { return cargarFuentes().then(function () { return pintar(datos); }); }

  // ---------- descifrado con la sesión del Flujo de cosecha
  function b64(u8) { var s = '', i; u8 = new Uint8Array(u8); for (i = 0; i < u8.length; i += 32768) s += String.fromCharCode.apply(null, u8.subarray(i, i + 32768)); return btoa(s); }
  function unb64(str) { var t = atob(str), u = new Uint8Array(t.length); for (var i = 0; i < t.length; i++) u[i] = t.charCodeAt(i); return u; }
  function read(st, k) { try { return st.getItem(k); } catch (e) { return null; } }
  function dec(key, iv, data) { return crypto.subtle.decrypt({ name: 'AES-GCM', iv: unb64(iv) }, key, unb64(data)); }
  async function llave(pk) {
    var ses = null;
    try { ses = JSON.parse(read(sessionStorage, 'fc.ses') || read(localStorage, 'fc.ses') || 'null'); } catch (e) { ses = null; }
    if (ses && ses.k && pk.keys) {
      var dk = await crypto.subtle.importKey('raw', unb64(ses.k), 'AES-GCM', false, ['decrypt']);
      for (var i = 0; i < pk.keys.length; i++) { try { return new Uint8Array(await dec(dk, pk.keys[i].i, pk.keys[i].c)); } catch (e) { /* otra cerradura */ } }
    }
    var adm = read(localStorage, 'fc.adm');
    if (adm && pk.adm) {
      try {
        var base = await crypto.subtle.importKey('raw', te.encode('adm:' + adm), 'PBKDF2', false, ['deriveKey']);
        var ak = await crypto.subtle.deriveKey({ name: 'PBKDF2', salt: unb64(pk.adm.s), iterations: pk.it, hash: 'SHA-256' }, base, { name: 'AES-GCM', length: 256 }, false, ['decrypt']);
        return unb64(JSON.parse(td.decode(await dec(ak, pk.adm.i, pk.adm.c))).k);
      } catch (e) { /* clave vieja */ }
    }
    return null;
  }
  async function abrir(pk) {
    if (!pk || !pk.data) return null;
    var raw = await llave(pk); if (!raw) return null;
    var k = await crypto.subtle.importKey('raw', raw, 'AES-GCM', false, ['decrypt']);
    var plano = new Uint8Array(await dec(k, pk.iv, pk.data));
    if (pk.z) plano = new Uint8Array(await new Response(new Blob([plano]).stream().pipeThrough(new DecompressionStream('gzip'))).arrayBuffer());
    return JSON.parse(td.decode(plano));
  }

  // ---------- archivos del repositorio (directo de GitHub, sin esperar la publicación del sitio)
  function api(tok, ruta, op) {
    op = op || {}; op.cache = 'no-store';
    op.headers = Object.assign({ Authorization: 'Bearer ' + tok, Accept: 'application/vnd.github+json' }, op.headers || {});
    return fetch('https://api.github.com/repos/' + REPO + '/contents/' + ruta, op);
  }
  async function leerPk(tok) {
    if (tok) {
      try {
        var r = await api(tok, 'paletas.json?t=' + Date.now(), { headers: { Accept: 'application/vnd.github.raw' } });
        if (r.status === 404) return null;
        if (r.ok) return JSON.parse(await r.text());
      } catch (e) { /* se usa la copia del sitio */ }
    }
    try { var r2 = await fetch('paletas.json?t=' + Date.now(), { cache: 'no-store' }); return r2.ok ? await r2.json() : null; } catch (e) { return null; }
  }
  async function leer(tok) {
    var pk = await leerPk(tok);
    var datos = null; try { datos = await abrir(pk); } catch (e) { datos = null; }
    return { cab: pk, datos: datos };
  }
  async function pedir(tok, quien) {
    var id = Date.now().toString(36) + Math.random().toString(36).slice(2, 7), sha = null;
    var r0 = await api(tok, 'solicitud_paletas.json?t=' + Date.now());
    if (r0.ok) sha = (await r0.json()).sha; else if (r0.status !== 404) throw new Error('GitHub respondió ' + r0.status);
    var cuerpo = { message: 'Pedido de stock de paletas', content: b64(te.encode(JSON.stringify({ id: id, t: new Date().toISOString(), w: quien || '' }) + '\n')) };
    if (sha) cuerpo.sha = sha;
    var r = await api(tok, 'solicitud_paletas.json', { method: 'PUT', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(cuerpo) });
    if (!r.ok) throw new Error(r.status === 401 ? 'el permiso de guardado venció' : 'GitHub respondió ' + r.status);
    try { localStorage.setItem('pal.sol', JSON.stringify({ id: id, t: Date.now() })); } catch (e) { /* nada */ }
    return id;
  }
  function esperar(tok, id, avance, desde) {
    var t0 = desde || Date.now();
    return new Promise(function (ok, ko) {
      var reloj = setInterval(function () { if (avance) avance(Date.now() - t0); }, 1000);
      if (avance) avance(Date.now() - t0);
      (async function revisar() {
        var pk = null;
        try { pk = await leerPk(tok); } catch (e) { pk = null; }
        if (pk && pk.sol === id) {
          clearInterval(reloj);
          try { localStorage.removeItem('pal.sol'); } catch (e) { /* nada */ }
          var datos = null; try { datos = await abrir(pk); } catch (e) { datos = null; }
          return ok({ cab: pk, datos: datos });
        }
        if (Date.now() - t0 > 20 * 60000) {
          clearInterval(reloj);
          try { localStorage.removeItem('pal.sol'); } catch (e) { /* nada */ }
          return ko(new Error('la tarea tardó más de 20 minutos; revisa GitHub → Actions → «Stock de paletas»'));
        }
        setTimeout(revisar, 10000);
      })();
    });
  }
  function pendiente() {   // un pedido que seguía en curso (por ejemplo, si se recargó la página)
    var p = null; try { p = JSON.parse(read(localStorage, 'pal.sol') || 'null'); } catch (e) { p = null; }
    return p && p.id && Date.now() - p.t < 20 * 60000 ? p : null;
  }

  window.StockPaletas = { grupos: grupos, dibujar: dibujar, abrir: abrir, leer: leer, pedir: pedir, esperar: esperar, pendiente: pendiente };
})();
