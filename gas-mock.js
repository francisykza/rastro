/*
 * Rastro · demo en el navegador
 * Simulación de los servicios de Google Apps Script que usa el sistema:
 * hoja de cálculo en memoria, propiedades, caché, bloqueos, menús,
 * plantillas HTML, ventanas, barra lateral y google.script.run.
 * El código del sistema (Rastro.gs) corre tal cual encima de esto.
 */
(function (G) {
  'use strict';

  var DEMO = G.__demo = G.__demo || {};
  DEMO.usuario = 'admin@ejemplo.com';
  DEMO.oyentes = [];
  DEMO.avisar = function (tipo, datos) { DEMO.oyentes.forEach(function (f) { try { f(tipo, datos); } catch (e) { console.error(e); } }); };

  // ---------------------------------------------------------------- utilidades
  function letra(c) { var s = ''; while (c > 0) { var m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; }
  function numCol(l) { var n = 0; for (var i = 0; i < l.length; i++) n = n * 26 + (l.charCodeAt(i) - 64); return n; }
  function dos(n) { return ('0' + n).slice(-2); }
  function esFecha(v) { return Object.prototype.toString.call(v) === '[object Date]'; }
  function textoVisible(v, formato) {
    if (v === null || v === undefined) return '';
    if (esFecha(v)) {
      if (isNaN(v.getTime())) return '';
      var f = dos(v.getDate()) + '/' + dos(v.getMonth() + 1) + '/' + v.getFullYear();
      if (formato && /h/i.test(formato) || v.getHours() || v.getMinutes() || v.getSeconds()) f += ' ' + v.getHours() + ':' + dos(v.getMinutes()) + ':' + dos(v.getSeconds());
      return f;
    }
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return String(v).replace('.', ',');
    return String(v);
  }
  function copiar(v) { return esFecha(v) ? new Date(v.getTime()) : v; }
  function convertirEntrada(v, formato) {
    // Lo que hace Sheets al escribir: con formato texto (@) se queda tal cual;
    // si no, los números «limpios» pasan a número (y pierden los ceros).
    if (typeof v !== 'string' || formato === '@') return v;
    if (/^-?\d{1,15}$/.test(v) && !/^-?0\d/.test(v)) return Number(v);
    return v;
  }

  // ---------------------------------------------------------------- hoja
  var ID = 1;
  function Hoja(libro, nombre) {
    this.libro = libro; this.nombre = nombre; this.id = ID++;
    this.datos = []; this.fondos = {}; this.formatos = {}; this.oculta = false;
    this.filasMax = 1000; this.colsMax = 26; this.reglas = []; this.protecciones = [];
  }
  Hoja.prototype._fila = function (r) { while (this.datos.length < r) this.datos.push([]); return this.datos[r - 1]; };
  Hoja.prototype._v = function (r, c) { var f = this.datos[r - 1]; var v = f ? f[c - 1] : undefined; return v === undefined || v === null ? '' : v; };
  Hoja.prototype._set = function (r, c, v) {
    var f = this._fila(r); while (f.length < c - 1) f.push('');
    f[c - 1] = convertirEntrada(v, this.formatos[r + ',' + c]);
    if (r > this.filasMax) this.filasMax = r;
  };
  Hoja.prototype.getName = function () { return this.nombre; };
  Hoja.prototype.setName = function (n) { this.nombre = n; DEMO.avisar('hojas'); return this; };
  Hoja.prototype.getSheetId = function () { return this.id; };
  Hoja.prototype.getIndex = function () { return this.libro.hojas.indexOf(this) + 1; };
  Hoja.prototype.getParent = function () { return this.libro; };
  Hoja.prototype.getLastRow = function () {
    for (var r = this.datos.length; r >= 1; r--) { var f = this.datos[r - 1]; if (f && f.some(function (x) { return x !== '' && x !== null && x !== undefined; })) return r; }
    return 0;
  };
  Hoja.prototype.getLastColumn = function () {
    var m = 0; this.datos.forEach(function (f) { for (var c = f.length; c > m; c--) { if (f[c - 1] !== '' && f[c - 1] !== null && f[c - 1] !== undefined) { m = c; break; } } });
    return m;
  };
  Hoja.prototype.getMaxRows = function () { return Math.max(this.filasMax, this.datos.length); };
  Hoja.prototype.getMaxColumns = function () { return this.colsMax; };
  Hoja.prototype.getFrozenRows = function () { return 1; };
  Hoja.prototype.getRange = function (a, b, c, d) {
    if (typeof a === 'string') {
      var m = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/i.exec(a.replace(/^.*!/, '').replace(/\$/g, ''));
      if (!m) throw new Error('Rango no válido: ' + a);
      var c1 = numCol(m[1].toUpperCase()), r1 = m[2] ? +m[2] : 1, c2 = m[3] ? numCol(m[3].toUpperCase()) : c1, r2 = m[4] ? +m[4] : (m[2] ? r1 : this.getMaxRows());
      return new Rango(this, r1, c1, r2 - r1 + 1, c2 - c1 + 1);
    }
    return new Rango(this, a, b, c || 1, d || 1);
  };
  Hoja.prototype.getDataRange = function () { return new Rango(this, 1, 1, Math.max(1, this.getLastRow()), Math.max(1, this.getLastColumn())); };
  Hoja.prototype.appendRow = function (fila) { var r = this.getLastRow() + 1; for (var c = 0; c < fila.length; c++) this._set(r, c + 1, fila[c]); DEMO.avisar('celdas', { hoja: this }); return this; };
  Hoja.prototype.deleteRows = function (r, n) { this.datos.splice(r - 1, n || 1); this._moverMapas(r, -(n || 1)); this.filasMax = Math.max(1, this.filasMax - (n || 1)); DEMO.avisar('celdas', { hoja: this }); };
  Hoja.prototype.deleteRow = function (r) { this.deleteRows(r, 1); };
  Hoja.prototype._moverMapas = function (desde, delta) {
    ['fondos', 'formatos'].forEach(function (k) {
      var nuevo = {}, viejo = this[k];
      Object.keys(viejo).forEach(function (clave) { var p = clave.split(','), r = +p[0]; if (r >= desde) { if (delta < 0 && r < desde - delta) return; r += delta; } nuevo[r + ',' + p[1]] = viejo[clave]; });
      this[k] = nuevo;
    }, this);
  };
  Hoja.prototype.insertRowsAfter = function (r, n) { this.filasMax += n; return this; };
  Hoja.prototype.insertRowAfter = function (r) { return this.insertRowsAfter(r, 1); };
  Hoja.prototype.insertColumnsAfter = function (c, n) { this.colsMax += n; return this; };
  Hoja.prototype.hideSheet = function () { this.oculta = true; DEMO.avisar('hojas'); return this; };
  Hoja.prototype.showSheet = function () { this.oculta = false; DEMO.avisar('hojas'); return this; };
  Hoja.prototype.isSheetHidden = function () { return this.oculta; };
  Hoja.prototype.activate = function () { this.libro.activa = this; DEMO.avisar('activar', { hoja: this }); return this; };
  Hoja.prototype.getConditionalFormatRules = function () { return this.reglas.slice(); };
  Hoja.prototype.setConditionalFormatRules = function (r) { this.reglas = (r || []).slice(); return this; };
  Hoja.prototype.getProtections = function () { return this.protecciones.slice(); };
  Hoja.prototype.protect = function () { var p = new Proteccion(this); this.protecciones.push(p); return p; };
  Hoja.prototype.createTextFinder = function (t) { return new Buscador([this], null, t); };
  Hoja.prototype.clear = function () { this.datos = []; this.fondos = {}; DEMO.avisar('celdas', { hoja: this }); return this; };
  Hoja.prototype.setFrozenRows = function () { return this; };
  Hoja.prototype.setColumnWidth = function () { return this; };
  Hoja.prototype.autoResizeColumns = function () { return this; };

  function Proteccion(h) { this.h = h; this.desc = ''; this.aviso = false; }
  Proteccion.prototype.setWarningOnly = function (v) { this.aviso = v; return this; };
  Proteccion.prototype.isWarningOnly = function () { return this.aviso; };
  Proteccion.prototype.setDescription = function (d) { this.desc = d; return this; };
  Proteccion.prototype.getDescription = function () { return this.desc; };
  Proteccion.prototype.remove = function () { var l = this.h.protecciones; l.splice(l.indexOf(this), 1); };
  Proteccion.prototype.getRange = function () { return this.h.getDataRange(); };
  Proteccion.prototype.setRanges = function () { return this; };

  // ---------------------------------------------------------------- rango
  function Rango(h, r, c, nr, nc) { this.h = h; this.r = r; this.c = c; this.nr = nr; this.nc = nc; }
  Rango.prototype._mapa = function (f) { var out = []; for (var i = 0; i < this.nr; i++) { var fila = []; for (var j = 0; j < this.nc; j++) fila.push(f(this.r + i, this.c + j)); out.push(fila); } return out; };
  Rango.prototype.getRow = function () { return this.r; };
  Rango.prototype.getColumn = function () { return this.c; };
  Rango.prototype.getLastRow = function () { return this.r + this.nr - 1; };
  Rango.prototype.getLastColumn = function () { return this.c + this.nc - 1; };
  Rango.prototype.getNumRows = function () { return this.nr; };
  Rango.prototype.getNumColumns = function () { return this.nc; };
  Rango.prototype.getSheet = function () { return this.h; };
  Rango.prototype.getA1Notation = function () { var a = letra(this.c) + this.r; return (this.nr === 1 && this.nc === 1) ? a : a + ':' + letra(this.c + this.nc - 1) + (this.r + this.nr - 1); };
  Rango.prototype.getValues = function () { var h = this.h; return this._mapa(function (r, c) { return copiar(h._v(r, c)); }); };
  Rango.prototype.getValue = function () { return copiar(this.h._v(this.r, this.c)); };
  Rango.prototype.getDisplayValues = function () { var h = this.h; return this._mapa(function (r, c) { return textoVisible(h._v(r, c), h.formatos[r + ',' + c]); }); };
  Rango.prototype.getDisplayValue = function () { return textoVisible(this.h._v(this.r, this.c), this.h.formatos[this.r + ',' + this.c]); };
  Rango.prototype.getFormula = function () { return ''; };
  Rango.prototype.getFormulas = function () { return this._mapa(function () { return ''; }); };
  Rango.prototype.setValues = function (vals) {
    if (!vals || vals.length !== this.nr || (vals[0] && vals[0].length !== this.nc)) throw new Error('Las dimensiones de los datos (' + (vals ? vals.length : 0) + 'x' + (vals && vals[0] ? vals[0].length : 0) + ') no coinciden con las del rango (' + this.nr + 'x' + this.nc + ').');
    for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this.h._set(this.r + i, this.c + j, vals[i][j]);
    DEMO.avisar('celdas', { hoja: this.h }); return this;
  };
  Rango.prototype.setValue = function (v) { for (var i = 0; i < this.nr; i++) for (var j = 0; j < this.nc; j++) this.h._set(this.r + i, this.c + j, v); DEMO.avisar('celdas', { hoja: this.h }); return this; };
  Rango.prototype.clearContent = function () { return this.setValue(''); };
  Rango.prototype.clearContents = Rango.prototype.clearContent;
  Rango.prototype.clear = function () { this.setValue(''); var h = this.h; this._mapa(function (r, c) { delete h.fondos[r + ',' + c]; }); return this; };
  Rango.prototype.getBackgrounds = function () { var h = this.h; return this._mapa(function (r, c) { return h.fondos[r + ',' + c] || '#ffffff'; }); };
  Rango.prototype.getBackground = function () { return this.h.fondos[this.r + ',' + this.c] || '#ffffff'; };
  Rango.prototype.getBackgroundObject = function () { var b = this.getBackground(); return { asRgbColor: function () { return { asHexString: function () { return b; } }; } }; };
  Rango.prototype.setBackground = function (col) { var h = this.h; this._mapa(function (r, c) { if (col === null || /^#?fff(fff)?$/i.test(col || '')) delete h.fondos[r + ',' + c]; else h.fondos[r + ',' + c] = String(col).toLowerCase(); }); DEMO.avisar('celdas', { hoja: h }); return this; };
  Rango.prototype.setBackgrounds = function (m) { var h = this.h, self = this; this._mapa(function (r, c) { var col = m[r - self.r][c - self.c]; if (!col || /^#?fff(fff)?$/i.test(col)) delete h.fondos[r + ',' + c]; else h.fondos[r + ',' + c] = String(col).toLowerCase(); }); DEMO.avisar('celdas', { hoja: h }); return this; };
  Rango.prototype.getNumberFormats = function () { var h = this.h; return this._mapa(function (r, c) { return h.formatos[r + ',' + c] || 'General'; }); };
  Rango.prototype.getNumberFormat = function () { return this.h.formatos[this.r + ',' + this.c] || 'General'; };
  Rango.prototype.setNumberFormat = function (f) {
    var h = this.h;
    // Una columna entera de golpe: se guarda solo hasta los datos (más un margen).
    var hasta = Math.min(this.nr, 20000);
    for (var i = 0; i < hasta; i++) for (var j = 0; j < this.nc; j++) h.formatos[(this.r + i) + ',' + (this.c + j)] = f;
    return this;
  };
  Rango.prototype.setNumberFormats = function (m) { var h = this.h, self = this; this._mapa(function (r, c) { h.formatos[r + ',' + c] = m[r - self.r][c - self.c]; }); return this; };
  ['setFontWeight', 'setFontColor', 'setFontColors', 'setBold', 'setItalic', 'setUnderline', 'setStrikethrough', 'setHorizontalAlignment', 'setVerticalAlignment', 'setWrap', 'setBorder', 'setFontSize', 'setNote', 'setDataValidation', 'setDataValidations', 'setFontLine', 'setFontStyle'].forEach(function (m) { Rango.prototype[m] = function () { return this; }; });
  Rango.prototype.getFontColor = function () { return '#000000'; };
  Rango.prototype.getFontColorObject = function () { return { asRgbColor: function () { return { asHexString: function () { return '#000000'; } }; } }; };
  Rango.prototype.getDataValidation = function () { return null; };
  Rango.prototype.getDataValidations = function () { return this._mapa(function () { return null; }); };
  Rango.prototype.activate = function () { this.h.libro.activa = this.h; DEMO.avisar('activar', { hoja: this.h, fila: this.r, col: this.c }); return this; };
  Rango.prototype.createTextFinder = function (t) { return new Buscador([this.h], this, t); };
  Rango.prototype.offset = function (dr, dc, nr, nc) { return new Rango(this.h, this.r + dr, this.c + dc, nr || this.nr, nc || this.nc); };
  Rango.prototype.getCell = function (r, c) { return new Rango(this.h, this.r + r - 1, this.c + c - 1, 1, 1); };

  // ---------------------------------------------------------------- buscador de texto (Ctrl+F)
  function Buscador(hojas, rango, texto) { this.hojas = hojas; this.rango = rango; this.texto = String(texto); this.entera = false; this.mayus = false; this.regex = false; this.i = 0; this.res = null; }
  Buscador.prototype.matchEntireCell = function (v) { this.entera = !!v; return this; };
  Buscador.prototype.matchCase = function (v) { this.mayus = !!v; return this; };
  Buscador.prototype.useRegularExpression = function (v) { this.regex = !!v; return this; };
  Buscador.prototype.matchFormulaText = function () { return this; };
  Buscador.prototype.ignoreDiacritics = function () { return this; };
  Buscador.prototype.findAll = function () {
    var self = this, out = [], re;
    if (this.regex) re = new RegExp(this.entera ? '^(?:' + this.texto + ')$' : this.texto, this.mayus ? '' : 'i');
    var aguja = this.mayus ? this.texto : this.texto.toLowerCase();
    this.hojas.forEach(function (h) {
      var r0 = self.rango ? self.rango.r : 1, c0 = self.rango ? self.rango.c : 1;
      var r1 = self.rango ? self.rango.r + self.rango.nr - 1 : h.datos.length, c1 = self.rango ? self.rango.c + self.rango.nc - 1 : 999;
      for (var r = r0; r <= Math.min(r1, h.datos.length); r++) {
        var f = h.datos[r - 1] || [];
        for (var c = c0; c <= Math.min(c1, f.length); c++) {
          var t = textoVisible(f[c - 1]); if (t === '') continue;
          var ok = re ? re.test(t) : (self.entera ? (self.mayus ? t : t.toLowerCase()) === aguja : (self.mayus ? t : t.toLowerCase()).indexOf(aguja) > -1);
          if (ok) out.push(new Rango(h, r, c, 1, 1));
        }
      }
    });
    return out;
  };
  Buscador.prototype.findNext = function () { if (!this.res) this.res = this.findAll(); return this.res[this.i++] || null; };

  // ---------------------------------------------------------------- libro
  function Libro() { this.hojas = []; this.activa = null; }
  Libro.prototype.getSheetByName = function (n) { for (var i = 0; i < this.hojas.length; i++) if (this.hojas[i].nombre === n) return this.hojas[i]; return null; };
  Libro.prototype.getSheets = function () { return this.hojas.slice(); };
  Libro.prototype.insertSheet = function (n) { if (this.getSheetByName(n)) throw new Error('Ya existe una hoja con el nombre «' + n + '».'); var h = new Hoja(this, n || ('Hoja ' + (this.hojas.length + 1))); this.hojas.push(h); DEMO.avisar('hojas'); return h; };
  Libro.prototype.getActiveSheet = function () { return this.activa || this.hojas[0]; };
  Libro.prototype.getSpreadsheetTimeZone = function () { return 'Europe/Madrid'; };
  Libro.prototype.getId = function () { return 'demo-rastro'; };
  Libro.prototype.getName = function () { return 'Devoluciones (demo)'; };
  Libro.prototype.getUrl = function () { return location.href; };
  Libro.prototype.toast = function (msg, titulo, seg) { DEMO.avisar('toast', { msg: String(msg), titulo: titulo ? String(titulo) : '', seg: seg === undefined ? 5 : seg }); };
  Libro.prototype.createTextFinder = function (t) { return new Buscador(this.hojas, null, t); };
  Libro.prototype.getRange = function (a1) { var p = String(a1).split('!'); var h = p.length > 1 ? this.getSheetByName(p[0].replace(/'/g, '')) : this.getActiveSheet(); return h.getRange(p[p.length - 1]); };
  Libro.prototype.getActiveRange = function () { return this.getActiveSheet().getRange(DEMO.celdaActiva ? DEMO.celdaActiva.fila : 2, DEMO.celdaActiva ? DEMO.celdaActiva.col : 1); };
  var LIBRO = DEMO.libro = new Libro();

  // ---------------------------------------------------------------- reglas de formato condicional (solo se guardan)
  function ReglaBuilder() { this.d = { rangos: [], cond: null }; }
  ['whenFormulaSatisfied', 'whenTextContains', 'whenTextEqualTo', 'whenNumberGreaterThan', 'whenCellEmpty', 'whenCellNotEmpty', 'setBackground', 'setFontColor', 'setBold', 'setItalic', 'setStrikethrough', 'setUnderline', 'setGradientMaxpoint', 'setGradientMinpoint', 'setGradientMidpointWithValue', 'withCriteria'].forEach(function (m) { ReglaBuilder.prototype[m] = function () { this.d[m] = [].slice.call(arguments); return this; }; });
  ReglaBuilder.prototype.setRanges = function (r) { this.d.rangos = r; return this; };
  ReglaBuilder.prototype.build = function () { var d = this.d; return { getRanges: function () { return d.rangos; }, getBooleanCondition: function () { return null; }, getGradientCondition: function () { return null; }, copy: function () { var b = new ReglaBuilder(); b.d = d; return b; } }; };

  function enumeracion(nombres) { var o = {}; nombres.forEach(function (n) { o[n] = n; }); return o; }

  // ---------------------------------------------------------------- interfaz (menús, alertas, ventanas)
  var Boton = enumeracion(['OK', 'CANCEL', 'YES', 'NO', 'CLOSE']);
  var Ui = {
    ButtonSet: enumeracion(['OK', 'OK_CANCEL', 'YES_NO', 'YES_NO_CANCEL']),
    Button: Boton,
    createMenu: function (nombre) {
      var items = [];
      var m = {
        addItem: function (t, fn) { items.push({ t: t, fn: fn }); return m; },
        addSeparator: function () { items.push({ sep: true }); return m; },
        addSubMenu: function (sub) { items.push({ sub: sub }); return m; },
        addToUi: function () { DEMO.avisar('menu', { nombre: nombre, items: items }); }
      };
      return m;
    },
    createAddonMenu: function () { return Ui.createMenu('Complementos'); },
    alert: function (a, b, c) {
      var titulo = '', msg = a, botones = b;
      if (typeof b === 'string') { titulo = a; msg = b; botones = c; }
      var texto = (titulo ? titulo + '\n\n' : '') + msg;
      DEMO.avisar('toast', { titulo: titulo || 'Aviso', msg: String(msg), seg: 9 }); // también dentro de la página (donde no se ven las alertas del navegador)
      if (botones === 'YES_NO' || botones === 'OK_CANCEL' || botones === 'YES_NO_CANCEL') return G.confirm(texto) ? (botones === 'OK_CANCEL' ? Boton.OK : Boton.YES) : (botones === 'OK_CANCEL' ? Boton.CANCEL : Boton.NO);
      G.alert(texto); return Boton.OK;
    },
    prompt: function (a, b, c) {
      var titulo = typeof b === 'string' ? a : '', msg = typeof b === 'string' ? b : a;
      var r = G.prompt((titulo ? titulo + '\n\n' : '') + msg);
      return { getResponseText: function () { return r || ''; }, getSelectedButton: function () { return r === null ? Boton.CANCEL : Boton.OK; } };
    },
    showSidebar: function (salida) { DEMO.avisar('barra', { html: salida.getContent(), titulo: salida._titulo || '' }); },
    showModalDialog: function (salida, titulo) { DEMO.avisar('ventana', { html: salida.getContent(), titulo: titulo || '', ancho: salida._ancho || 600, alto: salida._alto || 400 }); },
    showModelessDialog: function (salida, titulo) { Ui.showModalDialog(salida, titulo); }
  };

  // ---------------------------------------------------------------- plantillas HTML (<? ?>, <?= ?>, <?!= ?>)
  function escHtml(s) { return String(s === null || s === undefined ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;'); }
  function compilar(fuente) {
    var cod = 'var __o=[];with(__v){', re = /<\?(!=|=)?([\s\S]*?)\?>/g, ult = 0, m;
    while ((m = re.exec(fuente))) {
      if (m.index > ult) cod += '__o.push(' + JSON.stringify(fuente.slice(ult, m.index)) + ');';
      if (m[1] === '=') cod += '__o.push(__esc(' + m[2] + '));';
      else if (m[1] === '!=') cod += '__o.push(String(' + m[2] + '));';
      else cod += m[2] + '\n';
      ult = re.lastIndex;
    }
    if (ult < fuente.length) cod += '__o.push(' + JSON.stringify(fuente.slice(ult)) + ');';
    cod += '}return __o.join("");';
    return new Function('__v', '__esc', cod);
  }
  function Salida(html) { this._html = html; this._ancho = 600; this._alto = 400; this._titulo = ''; }
  Salida.prototype.getContent = function () { return this._html; };
  Salida.prototype.setWidth = function (w) { this._ancho = w; return this; };
  Salida.prototype.setHeight = function (h) { this._alto = h; return this; };
  Salida.prototype.setTitle = function (t) { this._titulo = t; return this; };
  Salida.prototype.setSandboxMode = function () { return this; };
  Salida.prototype.setXFrameOptionsMode = function () { return this; };
  Salida.prototype.addMetaTag = function () { return this; };
  Salida.prototype.setFaviconUrl = function () { return this; };
  Salida.prototype.append = function (s) { this._html += s; return this; };
  function Plantilla(fuente) { this._f = fuente; }
  Plantilla.prototype.evaluate = function () {
    var vars = {}, self = this; Object.keys(this).forEach(function (k) { if (k !== '_f') vars[k] = self[k]; });
    return new Salida(compilar(this._f)(vars, escHtml));
  };
  Plantilla.prototype.getCode = function () { return this._f; };
  var HtmlService = {
    createTemplate: function (s) { return new Plantilla(String(s)); },
    createTemplateFromFile: function (n) { var f = (G.__ARCHIVOS_DEMO || {})[n]; if (f === undefined) throw new Error('No existe el archivo HTML «' + n + '».'); return new Plantilla(f); },
    createHtmlOutput: function (s) { return new Salida(String(s || '')); },
    createHtmlOutputFromFile: function (n) { return new Salida((G.__ARCHIVOS_DEMO || {})[n] || ''); },
    SandboxMode: enumeracion(['IFRAME', 'NATIVE', 'EMULATED']),
    XFrameOptionsMode: enumeracion(['ALLOWALL', 'DEFAULT'])
  };

  // ---------------------------------------------------------------- propiedades, caché, bloqueos, sesión
  function Almacen() { this.d = {}; }
  Almacen.prototype.getProperty = function (k) { return Object.prototype.hasOwnProperty.call(this.d, k) ? this.d[k] : null; };
  Almacen.prototype.setProperty = function (k, v) { this.d[k] = String(v); return this; };
  Almacen.prototype.setProperties = function (o, borrarResto) { if (borrarResto) this.d = {}; for (var k in o) this.d[k] = String(o[k]); return this; };
  Almacen.prototype.getProperties = function () { var o = {}; for (var k in this.d) o[k] = this.d[k]; return o; };
  Almacen.prototype.getKeys = function () { return Object.keys(this.d); };
  Almacen.prototype.deleteProperty = function (k) { delete this.d[k]; return this; };
  Almacen.prototype.deleteAllProperties = function () { this.d = {}; return this; };
  var PROPS = { doc: new Almacen(), user: new Almacen(), script: new Almacen() };
  DEMO.props = PROPS;
  var PropertiesService = { getDocumentProperties: function () { return PROPS.doc; }, getUserProperties: function () { return PROPS.user; }, getScriptProperties: function () { return PROPS.script; } };

  function Cache() { this.d = {}; }
  Cache.prototype.get = function (k) { var e = this.d[k]; if (!e) return null; if (e.f < Date.now()) { delete this.d[k]; return null; } return e.v; };
  Cache.prototype.put = function (k, v, s) { this.d[k] = { v: String(v), f: Date.now() + 1000 * (s || 600) }; };
  Cache.prototype.getAll = function (ks) { var o = {}, self = this; ks.forEach(function (k) { var v = self.get(k); if (v !== null) o[k] = v; }); return o; };
  Cache.prototype.putAll = function (o, s) { for (var k in o) this.put(k, o[k], s); };
  Cache.prototype.remove = function (k) { delete this.d[k]; };
  Cache.prototype.removeAll = function (ks) { var self = this; ks.forEach(function (k) { delete self.d[k]; }); };
  var CACHES = { doc: new Cache(), user: new Cache(), script: new Cache() };
  var CacheService = { getDocumentCache: function () { return CACHES.doc; }, getUserCache: function () { return CACHES.user; }, getScriptCache: function () { return CACHES.script; } };

  function Cerrojo() { this.tiene = false; }
  Cerrojo.prototype.tryLock = function () { this.tiene = true; return true; };
  Cerrojo.prototype.waitLock = function () { this.tiene = true; };
  Cerrojo.prototype.releaseLock = function () { this.tiene = false; };
  Cerrojo.prototype.hasLock = function () { return this.tiene; };
  var LockService = { getDocumentLock: function () { return new Cerrojo(); }, getScriptLock: function () { return new Cerrojo(); }, getUserLock: function () { return new Cerrojo(); } };

  var Session = {
    getActiveUser: function () { return { getEmail: function () { return DEMO.usuario; } }; },
    getEffectiveUser: function () { return { getEmail: function () { return DEMO.usuario; } }; },
    getScriptTimeZone: function () { return 'Europe/Madrid'; },
    getActiveUserLocale: function () { return 'es'; },
    getTemporaryActiveUserKey: function () { return 'demo'; }
  };

  // ---------------------------------------------------------------- disparadores
  var DISPARADORES = [];
  function Disparador(fn, tipo) { this.fn = fn; this.tipo = tipo; this.id = 'd' + Math.random().toString(36).slice(2); }
  Disparador.prototype.getHandlerFunction = function () { return this.fn; };
  Disparador.prototype.getUniqueId = function () { return this.id; };
  Disparador.prototype.getEventType = function () { return this.tipo; };
  Disparador.prototype.getTriggerSource = function () { return 'SPREADSHEETS'; };
  function ConstructorDisparador(fn) {
    var tipo = 'CLOCK', b = {
      timeBased: function () { return b; }, everyMinutes: function () { return b; }, everyHours: function () { return b; }, everyDays: function () { return b; },
      everyWeeks: function () { return b; }, atHour: function () { return b; }, nearMinute: function () { return b; }, onWeekDay: function () { return b; },
      inTimezone: function () { return b; }, at: function () { return b; }, after: function () { return b; },
      forSpreadsheet: function () { return b; }, onOpen: function () { tipo = 'ON_OPEN'; return b; }, onEdit: function () { tipo = 'ON_EDIT'; return b; },
      onChange: function () { tipo = 'ON_CHANGE'; return b; },
      create: function () { var d = new Disparador(fn, tipo); DISPARADORES.push(d); return d; }
    };
    return b;
  }
  var ScriptApp = {
    getProjectTriggers: function () { return DISPARADORES.slice(); },
    getUserTriggers: function () { return DISPARADORES.slice(); },
    newTrigger: function (fn) { return ConstructorDisparador(fn); },
    deleteTrigger: function (d) { var i = DISPARADORES.indexOf(d); if (i > -1) DISPARADORES.splice(i, 1); },
    EventType: enumeracion(['CLOCK', 'ON_OPEN', 'ON_EDIT', 'ON_CHANGE']),
    WeekDay: enumeracion(['MONDAY', 'TUESDAY', 'WEDNESDAY', 'THURSDAY', 'FRIDAY', 'SATURDAY', 'SUNDAY']),
    getScriptId: function () { return 'demo'; },
    getService: function () { return { getUrl: function () { return location.href; } }; }
  };

  // ---------------------------------------------------------------- utilidades de Apps Script
  var DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
  var Utilities = {
    formatDate: function (d, tz, p) {
      d = new Date(d);
      return String(p).replace(/yyyy|yy|MM|dd|HH|mm|ss|EEEE|EEE|E|d|M|H/g, function (t) {
        switch (t) {
          case 'yyyy': return d.getFullYear(); case 'yy': return String(d.getFullYear()).slice(-2);
          case 'MM': return dos(d.getMonth() + 1); case 'dd': return dos(d.getDate()); case 'HH': return dos(d.getHours());
          case 'mm': return dos(d.getMinutes()); case 'ss': return dos(d.getSeconds()); case 'd': return d.getDate(); case 'M': return d.getMonth() + 1; case 'H': return d.getHours();
          default: return DIAS[d.getDay()];
        }
      });
    },
    sleep: function () {},
    getUuid: function () { return 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function (c) { var r = Math.random() * 16 | 0; return (c === 'x' ? r : (r & 3 | 8)).toString(16); }); },
    base64Decode: function (s) { var b = atob(String(s)), out = []; for (var i = 0; i < b.length; i++) out.push(b.charCodeAt(i)); return out; },
    base64Encode: function (x) { return btoa(typeof x === 'string' ? unescape(encodeURIComponent(x)) : String.fromCharCode.apply(null, x)); },
    newBlob: function (bytes, tipo, nombre) { return { getBytes: function () { return bytes; }, getContentType: function () { return tipo; }, getName: function () { return nombre; }, setName: function (n) { nombre = n; return this; } }; },
    computeDigest: function () { return []; },
    DigestAlgorithm: enumeracion(['MD5', 'SHA_256']),
    Charset: enumeracion(['UTF_8'])
  };
  var MailApp = {
    sendEmail: function () { DEMO.avisar('toast', { titulo: '✉️ Correo (demo)', msg: 'En la demo no se envía ningún correo de verdad.', seg: 6 }); },
    getRemainingDailyQuota: function () { return 100; }
  };
  var Logger = { log: function () { if (DEMO.depurar) console.log.apply(console, arguments); }, clear: function () {}, getLog: function () { return ''; } };

  var SpreadsheetApp = {
    getActiveSpreadsheet: function () { return LIBRO; },
    getActive: function () { return LIBRO; },
    getActiveSheet: function () { return LIBRO.getActiveSheet(); },
    getUi: function () { return Ui; },
    flush: function () {},
    newConditionalFormatRule: function () { return new ReglaBuilder(); },
    newDataValidation: function () { var b = { requireValueInList: function () { return b; }, setAllowInvalid: function () { return b; }, build: function () { return null; } }; return b; },
    BooleanCriteria: enumeracion(['CUSTOM_FORMULA', 'TEXT_CONTAINS', 'TEXT_EQ', 'CELL_EMPTY', 'CELL_NOT_EMPTY', 'NUMBER_GREATER_THAN', 'DATE_EQ']),
    DataValidationCriteria: enumeracion(['VALUE_IN_LIST', 'VALUE_IN_RANGE', 'CHECKBOX', 'CUSTOM_FORMULA']),
    InterpolationType: enumeracion(['MIN', 'MAX', 'NUMBER', 'PERCENT', 'PERCENTILE']),
    ProtectionType: enumeracion(['SHEET', 'RANGE']),
    RelativeDate: enumeracion(['TODAY', 'YESTERDAY', 'TOMORROW', 'PAST_WEEK', 'PAST_MONTH', 'PAST_YEAR']),
    WrapStrategy: enumeracion(['WRAP', 'CLIP', 'OVERFLOW'])
  };

  G.SpreadsheetApp = SpreadsheetApp; G.PropertiesService = PropertiesService; G.CacheService = CacheService;
  G.LockService = LockService; G.Session = Session; G.HtmlService = HtmlService; G.ScriptApp = ScriptApp;
  G.Utilities = Utilities; G.MailApp = MailApp; G.GmailApp = MailApp; G.Logger = Logger;

  // ---------------------------------------------------------------- google.script.run (desde los paneles)
  // Cada panel es un iframe de la misma página: sus llamadas llegan aquí,
  // se ejecutan en el motor con un pequeño retraso (como el servidor de
  // verdad) y los datos se copian (como hace Google al serializar).
  DEMO.llamar = function (nombre, args, ok, mal) {
    setTimeout(function () {
      var fn = G[nombre];
      if (/_$/.test(nombre) || typeof fn !== 'function') { mal(new Error('Script function not found: ' + nombre)); return; }
      var res;
      try { res = fn.apply(null, JSON.parse(JSON.stringify(args || []))); }
      catch (e) { console.error('[' + nombre + ']', e); mal({ name: e.name, message: e.message }); return; }
      ok(res === undefined ? undefined : JSON.parse(JSON.stringify(res, function (k, v) { return v === undefined ? null : v; })));
    }, 120 + Math.random() * 180);
  };
  DEMO.puenteParaPanel = function (idMarco) {
    return '<script>(function(){var P=window.parent.__demo;window.addEventListener("error",function(e){window.parent.console.error("[panel] "+e.message+" @"+e.lineno);});' +
      'function Runner(ok,mal,u){this._ok=ok;this._mal=mal;this._u=u;}' +
      'Runner.prototype.withSuccessHandler=function(f){return new Runner(f,this._mal,this._u);};' +
      'Runner.prototype.withFailureHandler=function(f){return new Runner(this._ok,f,this._u);};' +
      'Runner.prototype.withUserObject=function(u){return new Runner(this._ok,this._mal,u);};' +
      'var proxy=function(r){return new Proxy(r,{get:function(t,k){if(k in t)return t[k];return function(){var a=[].slice.call(arguments);' +
      'P.llamar(k,a,function(v){if(t._ok)t._ok(v,t._u);},function(e){var err=new Error(e&&e.message?e.message:String(e));if(t._mal)t._mal(err,t._u);else console.error(err);});};}});};' +
      'var base=new Runner(null,null,undefined);var orig=Runner.prototype;' +
      '["withSuccessHandler","withFailureHandler","withUserObject"].forEach(function(m){var f=orig[m];orig[m]=function(x){return proxy(f.call(this,x));};});' +
      'window.google={script:{run:proxy(base),host:{close:function(){P.cerrarMarco(' + JSON.stringify(idMarco) + ');},setHeight:function(){},setWidth:function(){},editor:{focus:function(){}}},' +
      'url:{getLocation:function(cb){cb({hash:"",parameter:{},parameters:{}});}},history:{push:function(){},replace:function(){},setChangeHandler:function(){}}}};})();<\/script>';
  };
})(window);
