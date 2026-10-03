/*
 * Rastro · demo: datos ficticios.
 * Todo inventado: pedidos, clientes, correos y agencias. Se genera al abrir
 * la página, con fechas relativas a hoy, así que la demo siempre está «al día».
 * Usa las mismas funciones del sistema para guardar los historiales
 * (codificarRef_, getOrCreateSheet), así que el Recuento, el Calendario y la
 * revisión automática los leen exactamente como en el documento real.
 */
(function (G) {
  'use strict';
  var semilla = 20261003;
  function azar() { semilla = (semilla * 1103515245 + 12345) & 0x7fffffff; return semilla / 0x7fffffff; }
  function elige(l) { return l[Math.floor(azar() * l.length)]; }
  function ponderado(pares) { var t = 0, i; for (i = 0; i < pares.length; i++) t += pares[i][1]; var x = azar() * t; for (i = 0; i < pares.length; i++) { x -= pares[i][1]; if (x <= 0) return pares[i][0]; } return pares[0][0]; }
  function quitarTildes(s) { return s.normalize('NFD').replace(/[̀-ͯ]/g, ''); }

  var AGENCIAS = [['VELOX', 30], ['PAQNORTE', 20], ['CORREOMAX', 14], ['RUTASUR', 12], ['ATLAS', 9], ['PRONTO', 7], ['BOLIDO', 5], ['FARO', 3]];
  var NOMBRES = ['Lucía', 'Marta', 'Elena', 'Carmen', 'Laura', 'Sofía', 'Irene', 'Nuria', 'Noelia', 'Andrea', 'Raquel', 'Beatriz', 'Alba', 'Clara', 'Julia', 'Sara', 'Diego', 'Javier', 'Sergio', 'Álvaro', 'Rubén', 'Hugo', 'Adrián', 'Iván', 'Óscar', 'Mario', 'Víctor', 'Pablo', 'Daniel', 'Tomás'];
  var APELLIDOS = ['Ferrer', 'Navarro', 'Gil', 'Ortega', 'Molina', 'Romero', 'Castro', 'Vidal', 'Serrano', 'Rubio', 'Iglesias', 'Medina', 'Garrido', 'Cortés', 'Lozano', 'Santos', 'Prieto', 'Herrera', 'Peña', 'Marín', 'Cano', 'Vega', 'Campos', 'León', 'Montero', 'Pastor', 'Soto', 'Fuentes', 'Calvo', 'Nieto'];
  var ARTICULOS = ['Crema hidratante 50 ml', 'Champú reparador 400 ml', 'Perfume floral 100 ml', 'Sérum vitamina C 30 ml', 'Paleta de sombras nude', 'Máscara de pestañas negra', 'Protector solar SPF50 200 ml', 'Gel de baño avena 750 ml', 'Desodorante roll-on 50 ml', 'Colonia infantil 200 ml', 'Contorno de ojos 15 ml', 'Base de maquillaje 30 ml', 'Aceite corporal 100 ml', 'Mascarilla capilar 300 ml', 'Laca de uñas rojo', 'Agua micelar 400 ml', 'Exfoliante facial 75 ml', 'Bálsamo labial', 'Eau de toilette cítrica 50 ml', 'Cepillo alisador'];
  var PAGOS = [['TARJETA', 46], ['CONTRA', 24], ['PAYPAL', 15], ['BIZUM', 10], ['TRANSFERENCIA', 5]];
  var EQUIPO = ['ana.soler@ejemplo.com', 'pablo.mena@ejemplo.com', 'admin@ejemplo.com'];
  var DOMINIOS = ['correo-demo.es', 'buzon-demo.com', 'mail-demo.es'];

  var clientes = [];
  for (var i = 0; i < 240; i++) {
    var n = elige(NOMBRES), a = elige(APELLIDOS), b = elige(APELLIDOS);
    clientes.push({ n: n + ' ' + a + ' ' + b, e: quitarTildes((n + '.' + a + (i % 7 === 0 ? i : '')).toLowerCase()) + '@' + elige(DOMINIOS) });
  }
  // Unos cuantos que devuelven mucho contra reembolso (para el ranking de Hucho).
  var abusones = clientes.slice(0, 7);

  function dia(offset) { var d = new Date(); d.setHours(0, 0, 0, 0); d.setDate(d.getDate() + offset); return d; }
  function conHora(d, minutos) { var x = new Date(d.getTime()); x.setMinutes(minutos); return x; }
  var pedidoBase = 7002400000;
  function nuevoPedido() { pedidoBase += 1 + Math.floor(azar() * 40); var s = String(pedidoBase); return azar() < 0.12 ? '0' + s.slice(1) : s; }

  G.__demo.generarDatos = function () {
    var ss = SpreadsheetApp.getActiveSpreadsheet(), ahora = new Date();
    var minutosAhora = ahora.getHours() * 60 + ahora.getMinutes();

    // ------------------------------------------------------------ Entradas
    var ent = ss.insertSheet('Entradas');
    ent.appendRow(['FECHA', 'AGENCIA', 'Nº PEDIDO', 'MULTIBULTO', 'ROTURA', 'OBSERVACIONES', 'FECHA PEDIDO', 'MÉTODO PAGO', 'CLIENTE', 'EMAIL', 'IMPORTE', 'VERIFICACIÓN']);
    var filas = [], eventos = [], pedidosVistos = [];
    var DIAS_ATRAS = 34, refHist = [];
    // Pedidos contra reembolso de este mes de quienes más devuelven, repartidos
    // entre los últimos días (para que el ranking de Hucho tenga contenido
    // también a principios de mes).
    var colaHucho = [];
    [4, 3, 3, 2, 2].forEach(function (c, ia) { for (var x = 0; x < c; x++) colaHucho.push(abusones[ia]); });
    colaHucho.sort(function () { return azar() - 0.5; });
    var hoyH = dia(0), inicioMes = new Date(hoyH.getFullYear(), hoyH.getMonth(), 1);
    for (var off = -DIAS_ATRAS; off <= 0; off++) {
      var d = dia(off), ds = d.getDay();
      if (ds === 0) continue;
      var total = ds === 6 ? 14 + Math.floor(azar() * 8) : 55 + Math.floor(azar() * 35);
      var mInicio = ds === 6 ? 15 * 60 : 6 * 60 + 40, mFin = ds === 6 ? 19 * 60 : 20 * 60;
      if (off === 0) { mFin = Math.min(mFin, Math.max(mInicio + 20, minutosAhora)); total = Math.max(4, Math.round(total * Math.min(1, (mFin - mInicio) / (20 * 60 - mInicio)))); }
      var k = 0;
      while (k < total) {
        var ag = ds === 6 ? ponderado([['RUTASUR', 6], ['VELOX', 2]]) : ponderado(AGENCIAS);
        var pedido, repetido = false;
        if (azar() < 0.02 && pedidosVistos.length > 50) { pedido = elige(pedidosVistos); repetido = true; } else pedido = nuevoPedido();
        var bultos = azar() < 0.08 ? 2 + (azar() < 0.3 ? 1 : 0) : 1;
        var pago = ponderado(PAGOS), cli;
        if (pago === 'CONTRA' && azar() < 0.22) cli = elige(abusones); else cli = elige(clientes);
        var fPed = new Date(d.getTime()); fPed.setDate(d.getDate() - (3 + Math.floor(azar() * 18)));
        var diasQuedan = Math.round((hoyH - d) / 86400000) + 1;
        if (colaHucho.length && d >= inicioMes && azar() < Math.min(1, colaHucho.length / Math.max(1, diasQuedan * 6))) {
          cli = colaHucho.pop(); pago = 'CONTRA'; bultos = 1;
          fPed = new Date(inicioMes.getTime() + Math.floor(azar() * Math.max(1, Math.round((d - inicioMes) / 86400000) + 1)) * 86400000);
        }
        var importe = (12 + Math.floor(azar() * 90)) + ',' + elige(['00', '50', '90', '95']);
        for (var bu = 0; bu < bultos && k < total; bu++, k++) {
          var minuto = Math.round(mInicio + (mFin - mInicio) * (k + azar() * 0.6) / total);
          var llegada = conHora(d, minuto);
          var rotura = azar() < 0.06 ? 'SI' : 'NO';
          var estado = '';
          var abierto = off === 0 ? (minutosAhora - minuto > 25 && azar() < 0.72) : azar() > 0.015;
          if (abierto) estado = ponderado([['OK', 80], ['RECLAMAR', 15], ['RECLAMAR/PENDIENTE', 5]]);
          if (rotura === 'SI' && estado === 'OK') estado = 'RECLAMAR';
          filas.push([d, ag, pedido, bultos > 1 ? 'SI' : '', rotura, repetido ? 'Ya vino otro día' : '', fPed, pago, cli.n, cli.e, importe, estado]);
          eventos.push({ fila: filas.length + 1, llegada: llegada, ag: ag, pedido: pedido, rotura: rotura, estado: estado, u: elige(EQUIPO) });
        }
        pedidosVistos.push(pedido);
      }
    }
    // De golpe: columna C en texto (como deja el sistema la columna de referencias).
    ent.getRange(2, 3, filas.length + 300, 1).setNumberFormat('@');
    ent.getRange(2, 1, filas.length, 12).setValues(filas);
    // Algunas filas pintadas a mano de otro color (aviso).
    [filas.length - 9, filas.length - 31, filas.length - 160].forEach(function (f) { if (f > 2) ent.getRange(f, 1, 1, 12).setBackground('#fce8b2'); });

    // ------------------------------------------------------------ Retornos
    var ret = ss.insertSheet('Retornos');
    ret.appendRow(['FECHA', 'AGENCIA', 'Nº PEDIDO', 'ABIERTO', 'TIPO', 'BULTOS', 'UBICACIÓN', 'NOMBRE CLIENTE', 'EMAIL', 'ARTC. DEVUELTO', 'ESTADO', 'REVISADO POR', 'REEMBOLSO', 'OBSERVACIONES']);
    var filasRet = [], fondosRet = [], eventosRet = [];
    for (var r = 0; r < 520; r++) {
      var offR = -Math.floor((520 - r) / 8.5);
      var dR = dia(offR); if (dR.getDay() === 0) dR = dia(offR - 1);
      var cliR = elige(clientes), procesado = offR < -2 ? azar() > 0.04 : azar() > 0.55;
      var pedR = azar() < 0.4 && pedidosVistos.length ? elige(pedidosVistos) : nuevoPedido();
      var agR = ponderado(AGENCIAS);
      var art = elige(ARTICULOS), estadoArt = ponderado([['PERFECTO', 70], ['ABIERTO', 18], ['ROTO', 7], ['USADO', 5]]);
      filasRet.push([procesado ? dR : '', agR, pedR, procesado ? 'SI' : 'NO', ponderado([['DEVOLUCIÓN', 80], ['CAMBIO', 20]]), 1 + (azar() < 0.1 ? 1 : 0), elige(['E-1', 'E-2', 'E-3', 'M-1', 'M-2']),
        azar() < 0.05 ? cliR.n.toUpperCase() : cliR.n, cliR.e, art, procesado ? estadoArt : '', procesado ? elige(['Ana', 'Pablo', 'Admin']) : '', procesado ? (estadoArt === 'USADO' ? 'NO' : 'SI') : '', estadoArt === 'ROTO' && procesado ? 'Caja aplastada' : '']);
      fondosRet.push(procesado ? '#b7e1cd' : '#ffffff');
      if (procesado) eventosRet.push({ fila: filasRet.length + 1, llegada: conHora(dR, 9 * 60 + Math.floor(azar() * 600)), ag: agR, pedido: pedR, u: elige(EQUIPO) });
    }
    ret.getRange(2, 3, filasRet.length + 300, 1).setNumberFormat('@');
    ret.getRange(2, 1, filasRet.length, 14).setValues(filasRet);
    ret.getRange(2, 1, filasRet.length, 1).setBackgrounds(fondosRet.map(function (c) { return [c]; }));
    for (var fr = 0; fr < fondosRet.length; fr++) if (fondosRet[fr] !== '#ffffff') ret.getRange(fr + 2, 2, 1, 13).setBackground(fondosRet[fr]);

    // ------------------------------------------------------------ historiales
    var limite = dia(-DIAS_CADUCIDAD_HISTORIALES);
    function hist(nombre, nombreArchivo, fecha, fila) { getOrCreateSheet(ss, fecha < limite ? nombreArchivo : nombre).appendRow(fila); }
    eventos.forEach(function (ev) {
      var cod = codificarRef_(ev.pedido);
      hist(HOJA_AGENCIAS, HOJA_AGENCIAS_ARCHIVO, ev.llegada, [ev.llegada, ev.u, 'Entradas', ev.ag, ev.fila, cod]);
      if (ev.rotura === 'SI') hist(HOJA_ROTURAS, HOJA_ROTURAS_ARCHIVO, ev.llegada, [new Date(ev.llegada.getTime() + 4 * 60000), ev.u, 'Entradas', ev.ag, ev.fila, cod]);
      if (ev.estado) {
        var t = new Date(Math.min(ev.llegada.getTime() + (5 + Math.floor(azar() * 50)) * 60000, Date.now() - 60000));
        var esOK = ev.estado === 'OK';
        hist(esOK ? HOJA_OK : HOJA_RECLAMAR, esOK ? HOJA_OK_ARCHIVO : HOJA_RECLAMAR_ARCHIVO, t, [t, ev.u, 'Entradas', ev.estado, 12, 'L', ev.fila, ev.ag, cod]);
      }
    });
    eventosRet.forEach(function (ev) {
      if (ev.llegada < limite) return;
      getOrCreateSheet(ss, HOJA_AGENCIAS).appendRow([ev.llegada, ev.u, 'Retornos', ev.ag, ev.fila, codificarRef_(ev.pedido)]);
    });
    // Ordenados por fecha, como quedan al ir registrando.
    [HOJA_AGENCIAS, HOJA_ROTURAS, HOJA_OK, HOJA_RECLAMAR, HOJA_AGENCIAS_ARCHIVO, HOJA_ROTURAS_ARCHIVO, HOJA_OK_ARCHIVO, HOJA_RECLAMAR_ARCHIVO].forEach(function (n) {
      var h = ss.getSheetByName(n); if (!h || h.getLastRow() < 3) return;
      var cab = h.datos[0], resto = h.datos.slice(1).sort(function (x, y) { return new Date(x[0]) - new Date(y[0]); });
      h.datos = [cab].concat(resto);
    });

    // ------------------------------------------------------------ notas entre turnos
    var notas = getOrCreateSheet(ss, HOJA_NOTAS);
    notas.appendRow([conHora(dia(-1), 20 * 60 + 12), 'pablo.mena@ejemplo.com', 'Han quedado 3 cajas de PAQNORTE sin abrir en la estantería E-2. Mañana a primera hora.', 'LEÍDA']);
    notas.appendRow([conHora(dia(0), 7 * 60 + 5), 'ana.soler@ejemplo.com', 'Ojo con los pedidos de VELOX de hoy: vienen varios multibulto. Revisad que estén todos los bultos antes de dar OK.', 'NUEVA']);

    // ------------------------------------------------------------ checklists enviados (con un hueco)
    var envios = { desde: claveFecha_(dia(-21)), e: {} };
    for (var o = -21; o < 0; o++) {
      var dd = dia(o); if (dd.getDay() === 0 || o === -3) continue;
      envios.e['d' + claveFecha_(dd)] = { t: conHora(dd, 20 * 60 + 6 + Math.floor(azar() * 9)).getTime(), u: elige(['ana.soler', 'pablo.mena']), o: 1 };
      if (dd.getDay() === 1 && o < -6) { var lunesAnt = dia(o - 7); envios.e['s' + claveFecha_(lunesAnt)] = { t: conHora(dd, 20 * 60 + 2).getTime(), u: 'ana.soler', o: 1 }; }
    }
    PropertiesService.getDocumentProperties().setProperty(PROP_ENVIOS_CHECKLIST_, JSON.stringify(envios));

    // La foto de la revisión automática parte de lo que hay ahora (sin falsos «faltaba»).
    try { if (typeof revisarRegistros_ === 'function') revisarRegistros_(60000); } catch (e) { console.warn('revisión inicial', e); }
    ent.activate();
    return { entradas: filas.length, retornos: filasRet.length };
  };
})(window);
