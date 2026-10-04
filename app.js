/*
 * Rastro · demo: la página (hoja editable, menú, barra lateral y ventanas).
 */
(function (G) {
  'use strict';
  var DEMO = G.__demo, ss, hojaVista = null, sel = { fila: 2, col: 3 }, editor = null;
  var ALTO = 24, COLS = 14, ANCHOS = [96, 96, 118, 84, 70, 150, 100, 110, 170, 210, 76, 170, 100, 160];
  var LISTAS = { 'Entradas|5': ['SI', 'NO'], 'Entradas|12': ['OK', 'RECLAMAR', 'RECLAMAR/PENDIENTE'], 'Entradas|4': ['SI', ''], 'Retornos|4': ['SI', 'NO'] };
  function $(id) { return document.getElementById(id); }
  function esc(s) { return String(s === null || s === undefined ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); }
  function letra(c) { return String.fromCharCode(64 + c); }

  // ------------------------------------------------------------ avisos (toast)
  var tToast = null;
  function toast(t, m, s) {
    var el = $('toast');
    el.innerHTML = (t ? '<b>' + esc(t) + '</b>' : '') + '<span>' + esc(m) + '</span><button aria-label="Cerrar" onclick="this.parentNode.classList.remove(\'ver\')">✕</button>';
    el.classList.add('ver'); clearTimeout(tToast);
    if (s !== -1) tToast = setTimeout(function () { el.classList.remove('ver'); }, Math.max(2, s || 5) * 1000);
  }

  // ------------------------------------------------------------ menú del sistema
  var menus = {};
  function pintarMenus() {
    var barra = $('menus'), h = '';
    Object.keys(menus).forEach(function (n) {
      h += '<div class="menu"><button class="menu-b" aria-haspopup="true">' + esc(n) + '</button><div class="menu-l" role="menu">' +
        menus[n].map(function (it) {
          if (it.sep) return '<hr>';
          return '<button role="menuitem" data-fn="' + esc(it.fn) + '">' + esc(it.t) + '</button>';
        }).join('') + '</div></div>';
    });
    barra.innerHTML = h;
  }
  document.addEventListener('click', function (e) {
    var b = e.target.closest && e.target.closest('.menu-b');
    document.querySelectorAll('.menu.abierto').forEach(function (m) { if (!b || m !== b.parentNode) m.classList.remove('abierto'); });
    if (b) { b.parentNode.classList.toggle('abierto'); return; }
    var it = e.target.closest && e.target.closest('[data-fn]');
    if (it) {
      document.querySelectorAll('.menu.abierto').forEach(function (m) { m.classList.remove('abierto'); });
      ejecutar(it.getAttribute('data-fn'));
    }
  });
  function ejecutar(fn) {
    setTimeout(function () {
      try { if (typeof G[fn] !== 'function') throw new Error('No existe ' + fn); G[fn](); }
      catch (e) { console.error(e); toast('⚠️ Error', e.message, 8); }
    }, 30);
  }

  // ------------------------------------------------------------ pestañas
  function pintarPestanas() {
    var h = '';
    ss.getSheets().forEach(function (s) {
      if (s.isSheetHidden()) return;
      h += '<button class="pest' + (s === hojaVista ? ' on' : '') + '" data-hoja="' + esc(s.getName()) + '">' + esc(s.getName()) + '</button>';
    });
    var ocultas = ss.getSheets().filter(function (s) { return s.isSheetHidden(); }).length;
    if (ocultas) h += '<span class="pest-ocultas" title="Historiales internos, ocultos como en el documento real">+' + ocultas + ' ocultas</span>';
    $('pestanas').innerHTML = h;
  }
  $('pestanas').addEventListener('click', function (e) { var b = e.target.closest('[data-hoja]'); if (b) verHoja(ss.getSheetByName(b.getAttribute('data-hoja'))); });

  // ------------------------------------------------------------ hoja (virtualizada)
  var pintarPendiente = false;
  function pedirPintar() { if (pintarPendiente) return; pintarPendiente = true; requestAnimationFrame(function () { pintarPendiente = false; pintarHoja(); }); }
  function verHoja(h, fila, col) {
    if (!h) return;
    cerrarEditor(false);
    var cambia = h !== hojaVista;
    hojaVista = h;
    if (fila) { sel = { fila: fila, col: col || 1 }; }
    else if (cambia) { sel = { fila: Math.max(2, h.getLastRow()), col: 3 }; }
    var cont = $('rejilla');
    $('lienzo').style.height = (Math.max(h.getLastRow() + 40, 60) * ALTO + ALTO) + 'px';
    if (fila || cambia) cont.scrollTop = Math.max(0, (sel.fila - 1) * ALTO - cont.clientHeight / 2);
    pintarPestanas(); pintarHoja(); actualizarBarraCelda();
  }
  function pintarHoja() {
    var h = hojaVista; if (!h || editor) return;
    var cont = $('rejilla'), alto = cont.clientHeight || 600;
    $('lienzo').style.height = (Math.max(h.getLastRow() + 40, 60) * ALTO + ALTO) + 'px';
    var desde = Math.max(1, Math.floor(cont.scrollTop / ALTO) - 10), hasta = desde + Math.ceil(alto / ALTO) + 22;
    var ancho = 46 + ANCHOS.slice(0, COLS).reduce(function (a, b) { return a + b; }, 0);
    var html = '<div class="fila cab" style="top:' + cont.scrollTop + 'px;width:' + ancho + 'px"><div class="n"></div>';
    for (var c = 1; c <= COLS; c++) html += '<div class="c" style="width:' + ANCHOS[c - 1] + 'px">' + letra(c) + '</div>';
    html += '</div>';
    for (var r = desde; r <= hasta; r++) {
      var f = h.datos[r - 1] || [];
      html += '<div class="fila' + (r === 1 ? ' tit' : '') + '" style="top:' + (r * ALTO) + 'px;width:' + ancho + 'px"><div class="n' + (r === sel.fila ? ' on' : '') + '">' + r + '</div>';
      for (var c2 = 1; c2 <= COLS; c2++) {
        var v = f[c2 - 1], fondo = h.fondos[r + ',' + c2], txt = h.getRange(r, c2).getDisplayValue();
        var cls = 'c' + (r === sel.fila && c2 === sel.col ? ' sel' : '') + (typeof v === 'number' ? ' num' : '');
        html += '<div class="' + cls + '" data-r="' + r + '" data-c="' + c2 + '" style="width:' + ANCHOS[c2 - 1] + 'px' + (fondo ? ';background:' + fondo : '') + '" title="' + esc(txt.length > 18 ? txt : '') + '">' + esc(txt) + '</div>';
      }
      html += '</div>';
    }
    $('lienzo').innerHTML = html;
  }
  $('rejilla').addEventListener('scroll', pedirPintar);
  G.addEventListener('resize', pedirPintar);
  function actualizarBarraCelda() {
    if (!hojaVista) return;
    $('ref-celda').textContent = letra(sel.col) + sel.fila;
    $('valor-celda').textContent = hojaVista.getRange(sel.fila, sel.col).getDisplayValue();
    DEMO.celdaActiva = { fila: sel.fila, col: sel.col };
  }
  $('lienzo').addEventListener('mousedown', function (e) {
    var c = e.target.closest('.c[data-r]'); if (!c) return;
    if (editor && c.contains(editor)) return;
    cerrarEditor(true);
    var f = +c.getAttribute('data-r'), co = +c.getAttribute('data-c'), ahora = Date.now();
    var mismo = f === sel.fila && co === sel.col;
    sel = { fila: f, col: co };
    if (mismo && ahora - ultimoClic < 450) { ultimoClic = 0; e.preventDefault(); abrirEditor(); return; }
    ultimoClic = ahora;
    pintarHoja(); actualizarBarraCelda();
    $('rejilla').focus();
  });
  var ultimoClic = 0;

  function abrirEditor(textoInicial) {
    if (!hojaVista || sel.fila < 1) return;
    cerrarEditor(false);
    var celda = $('lienzo').querySelector('.c[data-r="' + sel.fila + '"][data-c="' + sel.col + '"]'); if (!celda) return;
    var lista = LISTAS[hojaVista.getName() + '|' + sel.col];
    var actual = hojaVista.getRange(sel.fila, sel.col).getDisplayValue();
    if (lista) {
      editor = document.createElement('select');
      editor.innerHTML = '<option value="">(vacío)</option>' + lista.filter(Boolean).map(function (x) { return '<option' + (x === actual ? ' selected' : '') + '>' + esc(x) + '</option>'; }).join('');
      editor.addEventListener('change', function () { cerrarEditor(true); });
    } else {
      editor = document.createElement('input');
      editor.value = textoInicial !== undefined ? textoInicial : actual;
    }
    editor.className = 'editor';
    editor.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); cerrarEditor(true); mover(1, 0); }
      else if (e.key === 'Escape') { cerrarEditor(false); }
      else if (e.key === 'Tab') { e.preventDefault(); cerrarEditor(true); mover(0, e.shiftKey ? -1 : 1); }
    });
    editor.addEventListener('blur', function () { setTimeout(function () { cerrarEditor(true); }, 0); });
    celda.textContent = ''; celda.appendChild(editor); editor.focus();
    if (editor.select && textoInicial === undefined) editor.select();
  }
  function cerrarEditor(guardar) {
    if (!editor) return;
    var ed = editor; editor = null;
    if (guardar) escribir(sel.fila, sel.col, ed.value);
    if (ed.parentNode) ed.parentNode.removeChild(ed);
    pintarHoja(); $('rejilla').focus();
  }
  function mover(dr, dc) { sel.fila = Math.max(1, sel.fila + dr); sel.col = Math.min(COLS, Math.max(1, sel.col + dc)); asegurarVisible(); pintarHoja(); actualizarBarraCelda(); }
  function asegurarVisible() { var cont = $('rejilla'), y = sel.fila * ALTO; if (y < cont.scrollTop + ALTO) cont.scrollTop = y - ALTO; else if (y + ALTO > cont.scrollTop + cont.clientHeight) cont.scrollTop = y + ALTO - cont.clientHeight; }
  $('rejilla').addEventListener('keydown', function (e) {
    if (editor) return;
    var k = e.key;
    if (k === 'ArrowDown') { e.preventDefault(); mover(1, 0); }
    else if (k === 'ArrowUp') { e.preventDefault(); mover(-1, 0); }
    else if (k === 'ArrowLeft') { e.preventDefault(); mover(0, -1); }
    else if (k === 'ArrowRight' || k === 'Tab') { e.preventDefault(); mover(0, 1); }
    else if (k === 'Enter' || k === 'F2') { e.preventDefault(); abrirEditor(); }
    else if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); escribir(sel.fila, sel.col, ''); }
    else if (k.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); abrirEditor(k); }
  });

  // Escribir «a mano»: como en Google, se guarda la celda y luego salta onEdit.
  function escribir(fila, col, texto) {
    var h = hojaVista, rg = h.getRange(fila, col), antes = rg.getDisplayValue();
    if (String(texto) === antes) return;
    rg.setValue(texto);
    var ev = { range: rg, source: ss, value: texto === '' ? undefined : String(texto), oldValue: antes === '' ? undefined : antes, user: { getEmail: function () { return DEMO.usuario; } }, authMode: 'LIMITED' };
    setTimeout(function () { try { if (typeof G.onEdit === 'function') G.onEdit(ev); } catch (e) { console.error('onEdit', e); } }, 40);
  }

  // ------------------------------------------------------------ barra lateral y ventanas
  DEMO.cerrarMarco = function (id) { if (id === 'barra') cerrarBarra(); else cerrarVentana(); };
  // Demo: las contraseñas se enseñan (todas son «demo») y en el móvil el panel
  // se dibuja un poco más pequeño para que se vea más hoja.
  var AYUDA_CLAVE = '<script>(function(){function marcar(){' +
    'var w=document.createTreeWalker(document.body,NodeFilter.SHOW_TEXT),n;while((n=w.nextNode())){if(n.nodeValue.indexOf("pide contraseña")>=0)n.nodeValue=n.nodeValue.replace(/pide contraseña/g,"contraseña: demo");}' +
    'var ps=document.querySelectorAll("input[type=password]:not([data-demo])");' +
    'for(var i=0;i<ps.length;i++){var p=ps[i];p.setAttribute("data-demo","1");p.removeAttribute("inputmode");p.placeholder="demo";' +
    'var a=document.createElement("div");a.className="demo-clave";a.textContent="🔑 En la demo, la contraseña es «demo»";' +
    'a.style.cssText="margin:6px 0 2px;font-size:12px;font-weight:600;color:#92400e;background:#fffbeb;border:1px solid #fcd34d;border-radius:8px;padding:5px 8px;text-align:center";' +
    'if(p.parentNode)p.parentNode.insertBefore(a,p.nextSibling);}}' +
    'var t=0;document.addEventListener("DOMContentLoaded",function(){marcar();new MutationObserver(function(){if(t)return;t=setTimeout(function(){t=0;marcar();},250);}).observe(document.body,{childList:true,subtree:true});});})();<\/script>';
  function preparar(html, id) {
    var puente = DEMO.puenteParaPanel(id) + AYUDA_CLAVE;
    if (id === 'barra' && movil.matches) puente += '<style>html{zoom:.86}</style>';
    html = html.replace(/pide contraseña/g, 'contraseña: demo');
    return /<head[^>]*>/i.test(html) ? html.replace(/<head[^>]*>/i, function (m) { return m + puente; }) : puente + html;
  }
  // Móvil: Rufo en un cajón a la derecha.
  var movil = G.matchMedia ? G.matchMedia('(max-width: 599px)') : { matches: false };
  function cajon(abrir) { document.body.classList.toggle('cajon-abierto', !!abrir); pedirPintar(); }
  function medirArriba() { var r = $('menus').getBoundingClientRect(); document.documentElement.style.setProperty('--arriba', Math.max(40, Math.round(r.bottom)) + 'px'); }
  G.addEventListener('resize', medirArriba);
  $('asa').addEventListener('click', function () { cajon(true); });
  $('barra-plegar').addEventListener('click', function () { cajon(false); });
  // En el móvil, al cargar la demo Rufo espera plegado (se ve la hoja entera);
  // tocar la hoja lo vuelve a apartar.
  var primeraBarra = true;
  $('rejilla').addEventListener('pointerdown', function () { if (movil.matches && document.body.classList.contains('cajon-abierto')) cajon(false); });
  function abrirBarra(html, titulo) {
    if (movil.matches && primeraBarra) { cajon(false); $('asa').classList.add('llama'); setTimeout(function () { $('asa').classList.remove('llama'); }, 6000); }
    else cajon(true);
    primeraBarra = false; medirArriba();
    $('barra-titulo').textContent = titulo || '';
    $('barra-marco').srcdoc = preparar(html, 'barra');
    document.body.classList.add('con-barra');
    pedirPintar();
  }
  function cerrarBarra() { cajon(false); document.body.classList.remove('con-barra'); $('barra-marco').srcdoc = ''; pedirPintar(); }
  $('barra-cerrar').addEventListener('click', cerrarBarra);
  function abrirVentana(d) {
    $('ventana-titulo').textContent = d.titulo || '';
    var c = $('ventana-caja');
    c.style.width = Math.min(d.ancho, G.innerWidth - 24) + 'px';
    c.style.height = Math.min(d.alto + 46, G.innerHeight - 24) + 'px';
    $('ventana-marco').srcdoc = preparar(d.html, 'ventana');
    $('ventana').classList.add('ver');
  }
  function cerrarVentana() { $('ventana').classList.remove('ver'); $('ventana-marco').srcdoc = ''; }
  $('ventana-cerrar').addEventListener('click', cerrarVentana);

  // ------------------------------------------------------------ eventos del motor
  DEMO.oyentes.push(function (tipo, d) {
    if (tipo === 'celdas') { if (!d || d.hoja === hojaVista) pedirPintar(); }
    else if (tipo === 'hojas') pintarPestanas();
    else if (tipo === 'menu') { menus[d.nombre] = d.items.map(function (x) { return x.sep ? { sep: true } : { t: x.t, fn: x.fn }; }); pintarMenus(); }
    else if (tipo === 'toast') toast(d.titulo, d.msg, d.seg);
    else if (tipo === 'barra') abrirBarra(d.html, d.titulo);
    else if (tipo === 'ventana') abrirVentana(d);
    else if (tipo === 'activar') { if (movil.matches) cajon(false); verHoja(d.hoja, d.fila, d.col); destello(); }
  });
  function destello() { var el = $('lienzo').querySelector('.c.sel'); if (el) { el.classList.remove('flash'); void el.offsetWidth; el.classList.add('flash'); } }

  // ------------------------------------------------------------ usuario de la demo
  $('como').addEventListener('change', function () {
    DEMO.usuario = this.value;
    menus = {}; pintarMenus(); cerrarBarra(); cerrarVentana();
    try { PropertiesService.getUserProperties().deleteAllProperties(); } catch (e) {}
    try { onOpen(); } catch (e) { console.error(e); }
    if (DEMO.usuario !== 'invitado@ejemplo.com') { try { mRufo(); } catch (e) { console.error(e); } }
    else {
      // Como en el documento real: el disparador instalable corre como su dueño
      // (administración) y el panel comprueba el acceso como quien lo está viendo.
      var quien = DEMO.usuario; DEMO.usuario = 'admin@ejemplo.com';
      try { disparadorAbrirRufo(); } catch (e) { console.error(e); } finally { DEMO.usuario = quien; }
    }
    toast('👤 Ahora ves la hoja como', this.options[this.selectedIndex].text, 5);
  });
  $('reiniciar').addEventListener('click', function () { try { sessionStorage.clear(); } catch (e) {} location.reload(); });
  $('ayuda').addEventListener('click', function () { $('guia').classList.toggle('ver'); });
  $('guia-cerrar').addEventListener('click', function () { $('guia').classList.remove('ver'); });

  // ------------------------------------------------------------ arranque
  function arrancar() {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    var t0 = performance.now(), n = DEMO.generarDatos();
    console.log('Datos de la demo:', n, Math.round(performance.now() - t0) + ' ms');
    verHoja(ss.getSheetByName('Entradas'));
    // El dibujo de Rufo para la pestañita del cajón (el mismo del sistema).
    try {
      var linea = (obtenerHTMLBase().split('\n').filter(function (l) { return l.indexOf('<g id="rufo-svg">') > -1; })[0] || '').trim();
      if (linea) document.body.insertAdjacentHTML('beforeend', '<svg width="0" height="0" style="position:absolute" aria-hidden="true"><defs>' + linea.replace('id="rufo-svg"', 'id="asa-rufo"') + '</defs></svg>');
    } catch (e) { console.warn('asa', e); }
    medirArriba();
    try { onOpen(); } catch (e) { console.error('onOpen', e); }
    // Lo que en el documento real hace el disparador instalable: abrir a Rufo.
    setTimeout(function () { try { mRufo(); } catch (e) { console.error(e); } }, 200);
    document.body.classList.add('listo');
  }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', arrancar); else arrancar();
})(window);
