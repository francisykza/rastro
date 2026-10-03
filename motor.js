/* Rastro · demo. © Francisco Muñoz Dorado. Código ofuscado: no se permite su copia ni reutilización. */
var CONTRASEÑA_NOTAS = "demo";
var CORREOS_AUTORIZADOS = [
    "luis.prado@ejemplo.com",
    "ana.soler@ejemplo.com",
    "pablo.mena@ejemplo.com",
    "admin@ejemplo.com"
];
var CORREOS_ADMIN_ = ["admin@ejemplo.com"];
function esAdmin_() {
    var _a = "";
    try {
        _a = (Session.getActiveUser().getEmail() || "").toString().trim().toLowerCase();
    }
    catch (_b) { }
    if (_a) {
        for (var _c = 0; _c < CORREOS_ADMIN_.length; _c++) {
            if (CORREOS_ADMIN_[_c].toString().trim().toLowerCase() === _a)
                return true;
        }
        return false;
    }
    try {
        var _d = PropertiesService.getUserProperties().getProperty('devoluciones_admin');
        return _d === 'si' ? true : (_d === 'no' ? false : null);
    }
    catch (_e) {
        return null;
    }
}
function soloAdminOAvisar_() {
    if (esAdmin_() === true)
        return true;
    try {
        SpreadsheetApp.getUi().alert("🔒 Esta opción es solo para la administración del sistema (la administración).\n\nNo se ha cambiado nada.");
    }
    catch (_f) { }
    return false;
}
function usuarioAutorizado_() {
    var _g = "";
    try {
        _g = (Session.getActiveUser().getEmail() || "").toString().trim().toLowerCase();
    }
    catch (_h) { }
    if (_g) {
        var _i = false;
        for (var _j = 0; _j < CORREOS_AUTORIZADOS.length; _j++) {
            if (CORREOS_AUTORIZADOS[_j].toString().trim().toLowerCase() === _g) {
                _i = true;
                break;
            }
        }
        var _k = CORREOS_ADMIN_.some(function (_l) { return _l.toString().trim().toLowerCase() === _g; });
        try {
            PropertiesService.getUserProperties().setProperties({ devoluciones_autorizado: _i ? 'si' : 'no', devoluciones_admin: _k ? 'si' : 'no', devoluciones_email: _g });
        }
        catch (_m) { }
        return _i;
    }
    try {
        return PropertiesService.getUserProperties().getProperty('devoluciones_autorizado') === 'si';
    }
    catch (_n) {
        return false;
    }
}
function diagnosticoAccesoDevoluciones() {
    var _o = [];
    var _p = "";
    try {
        _p = (Session.getActiveUser().getEmail() || "").toString();
        _o.push("1) Email leído ahora mismo: " + (_p ? _p : "(vacío)"));
    }
    catch (_q) {
        _o.push("1) ERROR al leer el email: " + _q.message);
    }
    try {
        var _s = PropertiesService.getUserProperties().getProperty('devoluciones_autorizado');
        _o.push("2) Valor guardado para esta cuenta: " + (_s === null ? "(no hay nada guardado todavía)" : _s));
    }
    catch (_t) {
        _o.push("2) ERROR al leer/usar PropertiesService: " + _t.message);
    }
    var _u = "?";
    try {
        _u = usuarioAutorizado_() ? "SÍ, autorizada" : "NO autorizada";
    }
    catch (_v) {
        _u = "ERROR: " + _v.message;
    }
    _o.push("3) ¿usuarioAutorizado_() la reconoce?: " + _u);
    var _w = _o.join("\n\n");
    try {
        SpreadsheetApp.getUi().alert("🔎 Diagnóstico de acceso (Corrección 90)", _w, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    catch (_x) {
        try {
            Logger.log(_w);
        }
        catch (_y) { }
    }
}
function diagnosticoRendimientoDevoluciones() {
    var _z = new Date().getTime();
    var _aa = [];
    var _ab = SpreadsheetApp.getActiveSpreadsheet();
    _aa.push("1) FILAS POR HOJA (para saber si el Auto-Archivado ha purgado de verdad el backlog acumulado):");
    var _ac = [
        "Entradas", HOJA_RETORNOS,
        HOJA_AGENCIAS, HOJA_ROTURAS, HOJA_RECLAMAR, HOJA_OK,
        HOJA_AGENCIAS_ARCHIVO, HOJA_ROTURAS_ARCHIVO, HOJA_RECLAMAR_ARCHIVO, HOJA_OK_ARCHIVO,
        HOJA_NOTAS, HOJA_NOTAS_ARCHIVO, HOJA_AVISOS, HOJA_AVISOS_EVENTOS
    ];
    for (var _ad = 0; _ad < _ac.length; _ad++) {
        try {
            var _ae = _ab.getSheetByName(_ac[_ad]);
            _aa.push("   - " + _ac[_ad] + ": " + (_ae ? (_ae.getLastRow() + " filas") : "(no existe)"));
        }
        catch (_af) {
            _aa.push("   - " + _ac[_ad] + ": ERROR al leerla (" + _af.message + ")");
        }
    }
    _aa.push("");
    _aa.push("2) DISPARADORES INSTALABLES ACTIVOS (un mismo nombre repetido = esa tarea pesada se está ejecutando varias veces a la vez sin que se note a simple vista):");
    try {
        var _ag = ScriptApp.getProjectTriggers();
        var _ah = {};
        for (var _ai = 0; _ai < _ag.length; _ai++) {
            var _aj = _ag[_ai].getHandlerFunction();
            _ah[_aj] = (_ah[_aj] || 0) + 1;
        }
        if (_ag.length === 0) {
            _aa.push("   (no hay ningún disparador instalable activo)");
        }
        else {
            for (var _ak in _ah) {
                var _al = _ah[_ak] > 1 ? " ⚠️ DUPLICADO (" + _ah[_ak] + " veces)" : "";
                _aa.push("   - " + _ak + ": " + _ah[_ak] + _al);
            }
        }
    }
    catch (_am) {
        _aa.push("   ERROR al leer los disparadores: " + _am.message);
    }
    _aa.push("");
    _aa.push("3) VELOCIDAD REAL AHORA MISMO (lectura pequeña de verdad contra el documento):");
    try {
        var _an = new Date().getTime();
        var _ao = getOrCreateSheet(_ab, HOJA_AGENCIAS);
        var _ap = _ao.getLastRow();
        _ao.getRange(Math.max(_ap, 1), 1, 1, 1).getValues();
        var _aq = new Date().getTime() - _an;
        _aa.push("   - Leer 1 celda de " + HOJA_AGENCIAS + ": " + _aq + " ms");
    }
    catch (_ar) {
        _aa.push("   ERROR al medir la velocidad: " + _ar.message);
    }
    _aa.push("");
    _aa.push("4) Duración total de este diagnóstico: " + (new Date().getTime() - _z) + " ms (una medida directa de lo lenta que está la conexión con el documento ahora mismo).");
    var _as = _aa.join("\n");
    try {
        SpreadsheetApp.getUi().alert("🔎 Diagnóstico de rendimiento (Corrección 96)", _as, SpreadsheetApp.getUi().ButtonSet.OK);
    }
    catch (_at) {
        try {
            Logger.log(_as);
        }
        catch (_au) { }
    }
}
var DIAS_CADUCIDAD_NOTAS = 7;
var DIAS_CADUCIDAD_HISTORIALES = 21;
var HOJA_RECLAMAR = "Historial_Reclamar";
var HOJA_OK = "Historial_OK";
var HOJA_AGENCIAS = "Historial_Agencias";
var HOJA_NOTAS = "Historial_Notas";
var HOJA_NOTAS_ARCHIVO = "Historial_Notas_Archivo";
var HOJA_ROTURAS = "Historial_Roturas";
var HOJA_AGENCIAS_ARCHIVO = "Historial_Agencias_Archivo";
var HOJA_ROTURAS_ARCHIVO = "Historial_Roturas_Archivo";
var HOJA_RECLAMAR_ARCHIVO = "Historial_Reclamar_Archivo";
var HOJA_OK_ARCHIVO = "Historial_OK_Archivo";
var HOJA_RETORNOS = "Retornos";
var HOJA_AVISOS = "Avisos_Pedidos";
var HOJA_AVISOS_EVENTOS = "Avisos_Pedidos_Eventos";
var AVISOS_PEDIDOS_ACTIVOS_ = false;
var HOJA_CONTROL_REVISION = "Control_Revision";
var HOJA_CAMBIOS_RET_ = "Cambios_Retornos";
var HOJA_COPIA_RET_ = "Retornos_Copia";
var HOJAS_TRABAJO_PERMITIDAS_ = ["Entradas", HOJA_RETORNOS];
var COLUMNA_AGENCIAS = 2;
var COLUMNA_REFERENCIAS = 3;
var COLUMNA_ROTURA = 5;
var COLUMNA_NOTA_L = 12;
var COLUMNA_RETORNOS_FECHA = 1;
var AG_LOGOS_B64_ = {};
function onEdit(_av) {
    var _aw = new Date().getTime();
    PLAZO_REGISTRO_ = _aw + 15000;
    try {
        if (!_av)
            return;
        var _ax = _av.range.getSheet();
        var _ay = _ax.getName();
        if (HOJAS_TRABAJO_PERMITIDAS_.indexOf(_ay) === -1)
            return;
        var _az = _av.range.getRow(), _ba = _av.range.getNumRows(), _bb = _av.range.getColumn(), _bc = _av.range.getNumColumns();
        var _bd;
        var _be = (_ba === 1 && _bc === 1 && _av.oldValue === undefined && _av.value !== undefined);
        try {
            _bd = registrarCambioCelda_(_av.source, _ax, _az, _bb, _ba, _bc, { anteriorVacio: _be });
        }
        catch (_bf) {
            var _bg = (_bf && _bf.message) ? _bf.message : String(_bf);
            try {
                Logger.log('registrarCambioCelda_ ha fallado: ' + _bg);
            }
            catch (_bh) { }
            try {
                _av.source.toast("⚠️ No se pudo registrar el cambio: " + _bg, "Error al guardar", 8);
            }
            catch (_bi) { }
            sumarContador_("ERRORES", 1);
            return;
        }
        var _bj = new Date().getTime() - _aw, _bk = 0;
        if (_ay === HOJA_RETORNOS) {
            if (_bj > 12000) {
                try {
                    PropertiesService.getDocumentProperties().setProperty('RET_COPIA_DESCUADRE', '1');
                }
                catch (_bl) { }
                sumarContador_("HISTRETSALTADO", 1);
            }
            else {
                var _bm = new Date().getTime();
                try {
                    registrarCambiosRetornos_(_ax, _az, _bb, _ba, _bc, { oldValue: (_ba === 1 && _bc === 1) ? _av.oldValue : undefined });
                }
                catch (_bn) {
                    try {
                        Logger.log('Historial de Retornos: ' + (_bn && _bn.message ? _bn.message : _bn));
                    }
                    catch (_bo) { }
                }
                _bk = new Date().getTime() - _bm;
                if (_bk > 5000)
                    sumarContador_("HISTRETLENTO", 1);
            }
        }
        if (new Date().getTime() - _aw > 20000) {
            sumarContador_("LENTAS", 1);
            if (_ay === HOJA_RETORNOS)
                sumarContador_("LENTASRET", 1);
            anotarEventoSalud_("LENTAS", _ay + ", fila " + _az + (_ba > 1 ? "-" + (_az + _ba - 1) : "") + " (" + (_ba * _bc) + " celda" + (_ba * _bc === 1 ? "" : "s") + "): " + Math.round((new Date().getTime() - _aw) / 1000) + " s" + (_ay === HOJA_RETORNOS ? " (historial de Retornos " + Math.round(_bk / 1000) + " s)" : ""));
            try {
                console.log('Edición lenta en ' + _ay + ' (fila ' + _az + ', ' + (_ba * _bc) + ' celda' + (_ba * _bc === 1 ? '' : 's') + '): registrar ' + _bj + ' ms' + (_ay === HOJA_RETORNOS ? ', historial de Retornos ' + _bk + ' ms' : ''));
            }
            catch (_bp) { }
        }
        var _bq = (typeof MENSAJE_AVISO_PENDIENTE_ !== 'undefined') ? MENSAJE_AVISO_PENDIENTE_ : null;
        var _br = (typeof MENSAJE_ERROR_AVISO_ !== 'undefined') ? MENSAJE_ERROR_AVISO_ : null;
        if (_br) {
            _av.source.toast("⚠️ Avisos de Pedidos no ha podido comprobar esta fila: " + _br, "Error en Avisos de Pedidos" + (_bd ? " (✅ registrado)" : ""), 15);
        }
        else if (_bq) {
            _av.source.toast(_bq, "🔔 AVISO DE PEDIDO" + (_bd ? " (✅ registrado)" : ""), 20);
        }
        else if (_bd) {
            _av.source.toast("✅ Registrado.", "Actualizado", 3);
        }
    }
    catch (_bs) { }
}
function registrarCambioCelda_(_bt, _bu, _bv, _bw, _bx, _by, _bz) {
    var _ca = _bu.getName();
    var _cb = !!(_bz && _bz.anteriorVacio);
    var _cc = null;
    function _cd(_ce) {
        if (_cc === null)
            _cc = _bu.getRange(_bv, COLUMNA_REFERENCIAS, _bx, 1).getDisplayValues();
        return (_cc[_ce][0] || "").toString().trim();
    }
    function _cf(_cg, _ch, _ci) {
        if (!_cg || _cg.ref === "")
            return false;
        if (_ch !== null && _cg.valor !== _ch)
            return false;
        var _cj = normalizarRef(_cd(_ci));
        return _cj !== "" && _cg.ref === _cj;
    }
    if (HOJAS_TRABAJO_PERMITIDAS_.indexOf(_ca) === -1)
        return false;
    var _ck = new Date(), _cl = "Usuario anónimo";
    try {
        _cl = Session.getActiveUser().getEmail() || "Usuario anónimo";
    }
    catch (_cm) { }
    var _cn = false;
    var _co = (_ca === HOJA_RETORNOS);
    if (_bw <= COLUMNA_REFERENCIAS) {
        try {
            asegurarColumnaReferenciaTexto_(_bu, _bv, _bx);
        }
        catch (_cp) { }
    }
    if (!_co && _bw <= 12 && (_bw + _by - 1) >= 12) {
        var _cq = _bu.getRange(_bv, 12, _bx, 1).getValues();
        var _cr = _bu.getRange(_bv, COLUMNA_AGENCIAS, _bx, 1).getDisplayValues();
        var _cs = [], _ct = [];
        var _cu = false;
        for (var _cv = 0; _cv < _bx; _cv++) {
            if (categoriaEstado_(_cq[_cv][0])) {
                _cu = true;
                break;
            }
        }
        var _cw = _cu ? ultimosRegistrosPorFila_(_bt, HOJA_OK, _ca, 3, 6, 8) : {};
        var _cx = _cu ? ultimosRegistrosPorFila_(_bt, HOJA_RECLAMAR, _ca, 3, 6, 8) : {};
        for (var _cy = 0; _cy < _bx; _cy++) {
            var _cz = categoriaEstado_(_cq[_cy][0]);
            var _da = _cz === "FAC" ? _cs : (_cz === "OK" ? _ct : null);
            if (_da) {
                var _db = ultimoEstadoDeFila_(_cw, _cx, _bv + _cy);
                if (_db && _db.cat === _cz && _cf({ valor: "", ref: _db.ref }, null, _cy)) {
                    _cn = true;
                    continue;
                }
                var _dc = (_cr[_cy][0] || "").toString().trim();
                _da.push([_ck, _cl, _ca, _cq[_cy][0], COLUMNA_NOTA_L, columnToLetter(COLUMNA_NOTA_L), _bv + _cy, _dc, refParaHistorial_(_cd(_cy))]);
            }
        }
        if (_cs.length > 0) {
            if (anexarFilasHistorial_(_bt, HOJA_RECLAMAR, _cs, 9) !== false)
                _cn = true;
        }
        if (_ct.length > 0) {
            if (anexarFilasHistorial_(_bt, HOJA_OK, _ct, 9) !== false)
                _cn = true;
        }
        var _dd = [];
        for (var _de = 0; _de < _bx; _de++) {
            if ((_cq[_de][0] || "").toString().trim() === "")
                _dd.push(_bv + _de);
        }
        if (_dd.length > 0 && !_cb) {
            if (quitarRegistrosDeFilas_(_bt, HOJA_OK, _ca, _dd, 2, 6) > 0)
                _cn = true;
            if (quitarRegistrosDeFilas_(_bt, HOJA_RECLAMAR, _ca, _dd, 2, 6) > 0)
                _cn = true;
        }
    }
    if (_bw <= COLUMNA_AGENCIAS && (_bw + _by - 1) >= COLUMNA_AGENCIAS) {
        var _df = _bu.getRange(_bv, COLUMNA_AGENCIAS, _bx, 1).getDisplayValues();
        var _dg = [];
        var _dh = _df.some(function (_di) { return (_di[0] || "").toString().trim() !== ""; });
        var _dj = _dh ? ultimosRegistrosPorFila_(_bt, HOJA_AGENCIAS, _ca, 3, 4, 5) : {};
        for (var _dk = 0; _dk < _bx; _dk++) {
            var _dl = (_df[_dk][0] || "").toString().trim();
            if (_dl === "")
                continue;
            if (_cf(_dj[_bv + _dk], _dl.toUpperCase(), _dk)) {
                _cn = true;
                continue;
            }
            _dg.push([_ck, _cl, _ca, _dl, _bv + _dk, refParaHistorial_(_cd(_dk))]);
        }
        if (_dg.length > 0) {
            if (anexarFilasHistorial_(_bt, HOJA_AGENCIAS, _dg, 6) !== false)
                _cn = true;
        }
        var _dm = [];
        for (var _dn = 0; _dn < _bx; _dn++) {
            if ((_df[_dn][0] || "").toString().trim() === "")
                _dm.push(_bv + _dn);
        }
        if (_dm.length > 0 && !_cb && quitarRegistrosDeFilas_(_bt, HOJA_AGENCIAS, _ca, _dm, 2, 4) > 0)
            _cn = true;
    }
    var _do = (_bw <= COLUMNA_AGENCIAS && (_bw + _by - 1) >= COLUMNA_AGENCIAS);
    var _dp = (_bw <= COLUMNA_REFERENCIAS && (_bw + _by - 1) >= COLUMNA_REFERENCIAS);
    if (AVISOS_PEDIDOS_ACTIVOS_ && (_do || _dp)) {
        try {
            procesarAvisosPorEdicion_(_bt, _ca, _bu, _bv, _bx);
        }
        catch (_dq) {
            var _dr = (_dq && _dq.message) ? _dq.message : String(_dq);
            try {
                Logger.log('procesarAvisosPorEdicion_ ha fallado: ' + _dr);
            }
            catch (_ds) { }
            try {
                _bt.toast('⚠️ Avisos de Pedidos no ha podido comprobar esta fila: ' + _dr, 'Error en Avisos de Pedidos', 8);
            }
            catch (_dt) { }
            MENSAJE_ERROR_AVISO_ = _dr;
        }
    }
    if (!_co && _bw <= COLUMNA_ROTURA && (_bw + _by - 1) >= COLUMNA_ROTURA) {
        var _du = _bu.getRange(_bv, COLUMNA_ROTURA, _bx, 1).getDisplayValues();
        var _dv = _bu.getRange(_bv, COLUMNA_AGENCIAS, _bx, 1).getDisplayValues();
        var _dw = [];
        var _dx = _du.some(function (_dy) { return (_dy[0] || "").toString().trim().toUpperCase() === "SI"; });
        var _dz = _dx ? ultimosRegistrosPorFila_(_bt, HOJA_ROTURAS, _ca, 3, 4, 5) : {};
        for (var _ea = 0; _ea < _bx; _ea++) {
            var _eb = (_du[_ea][0] || "").toString().trim().toUpperCase();
            if (_eb !== "SI")
                continue;
            if (_cf(_dz[_bv + _ea], null, _ea)) {
                _cn = true;
                continue;
            }
            _dw.push([_ck, _cl, _ca, (_dv[_ea][0] || "").toString().trim(), _bv + _ea, refParaHistorial_(_cd(_ea))]);
        }
        if (_dw.length > 0) {
            if (anexarFilasHistorial_(_bt, HOJA_ROTURAS, _dw, 6) !== false)
                _cn = true;
        }
        var _ec = [];
        for (var _ed = 0; _ed < _bx; _ed++) {
            if ((_du[_ed][0] || "").toString().trim().toUpperCase() !== "SI")
                _ec.push(_bv + _ed);
        }
        if (_ec.length > 0 && !_cb && quitarRegistrosDeFilas_(_bt, HOJA_ROTURAS, _ca, _ec, 2, 4) > 0)
            _cn = true;
    }
    return _cn;
}
function quitarRegistrosDeFilas_(_ee, _ef, _eg, _eh, _ei, _ej) {
    var _ek = _ee.getSheetByName(_ef);
    if (!_ek || _ek.getLastRow() < 2 || _eh.length === 0)
        return 0;
    var _el = {};
    for (var _em = 0; _em < _eh.length; _em++)
        _el[_eh[_em]] = true;
    function _en() {
        var _eo = _ek.getLastRow() - 1;
        if (_eo < 1)
            return [];
        var _ep = _ek.getRange(2, _ei + 1, _eo, 1).getValues();
        var _eq = _ek.getRange(2, _ej + 1, _eo, 1).getValues();
        var _er = [];
        for (var _es = 0; _es < _eo; _es++) {
            if (String(_ep[_es][0]) === _eg && _el[Number(_eq[_es][0])] === true)
                _er.push(_es + 2);
        }
        return _er;
    }
    if (_en().length === 0)
        return 0;
    var _et = esperaTurnoRegistro_();
    if (_et < 500) {
        sumarContador_("APLAZADOS", 1);
        anotarEventoSalud_("APLAZ", "quitar de " + _ef + " (de " + _eg + ", fila " + _eh[0] + ")");
        return 0;
    }
    var _eu = null;
    try {
        _eu = LockService.getDocumentLock();
        if (!_eu || !_eu.tryLock(_et))
            _eu = null;
    }
    catch (_ev) {
        _eu = null;
    }
    if (!_eu)
        return 0;
    try {
        var _ew = _en();
        if (_ew.length === 0)
            return 0;
        if (_ek.getMaxRows() <= _ek.getLastRow())
            _ek.insertRowsAfter(_ek.getMaxRows(), 1);
        var _ex = 0, _ey = _ew.length - 1;
        while (_ey >= 0) {
            var _ez = _ey;
            while (_ez > 0 && _ew[_ez - 1] === _ew[_ez] - 1)
                _ez--;
            _ek.deleteRows(_ew[_ez], _ey - _ez + 1);
            _ex += _ey - _ez + 1;
            _ey = _ez - 1;
        }
        SpreadsheetApp.flush();
        return _ex;
    }
    finally {
        try {
            _eu.releaseLock();
        }
        catch (_fa) { }
    }
}
var FILAS_COLA_HISTORIAL_ = 4000;
function ultimosRegistrosPorFila_(_fb, _fc, _fd, _fe, _ff, _fg) {
    var _fh = {};
    var _fi = _fb.getSheetByName(_fc);
    if (!_fi)
        return _fh;
    var _fj = _fi.getLastRow();
    if (_fj < 2)
        return _fh;
    var _fk = Math.max(2, _fj - FILAS_COLA_HISTORIAL_ + 1);
    var _fl = Math.min(Math.max(_fe, _ff, _fg) + 1, _fi.getMaxColumns());
    var _fm = _fi.getRange(_fk, 1, _fj - _fk + 1, _fl).getValues();
    for (var _fn = 0; _fn < _fm.length; _fn++) {
        if (String(_fm[_fn][2]) !== _fd)
            continue;
        var _fo = Number(_fm[_fn][_ff]);
        if (!(_fo >= 2))
            continue;
        var _fp = _fm[_fn][_fe];
        var _fq = Object.create(PROTO_REGISTRO_FILA_);
        _fq.valor = (_fp === undefined || _fp === null ? "" : String(_fp)).trim().toUpperCase();
        _fq.t = new Date(_fm[_fn][0]).getTime() || 0;
        _fq._r = _fm[_fn][_fg];
        _fq._ref = null;
        _fh[_fo] = _fq;
    }
    return _fh;
}
var PROTO_REGISTRO_FILA_ = Object.create(Object.prototype, {
    ref: { get: function () { if (this._ref === null)
            this._ref = decodificarRef_(this._r); return this._ref; } }
});
function refParaHistorial_(_fr) {
    return codificarRef_(_fr);
}
var MARCA_REF_CODIFICADA_ = "\uE000";
var BASE_REF_CODIFICADA_ = 0xE100;
function codificarRef_(_fs) {
    var _ft = normalizarRef(_fs === undefined || _fs === null ? "" : String(_fs));
    if (_ft === "")
        return "";
    var _fu = 0xE100, _fv = [0xE000];
    for (var _fw = 0; _fw < _ft.length; _fw++) {
        var _fx = _ft.charCodeAt(_fw);
        _fv.push(_fx < 0xF900 - _fu ? _fu + _fx : _fx);
    }
    return String.fromCharCode.apply(null, _fv);
}
function decodificarRef_(_fy) {
    if (_fy === "" || _fy === undefined || _fy === null)
        return "";
    if (typeof _fy !== "string")
        return (_fy instanceof Date) ? "" : normalizarRef(String(_fy));
    if (_fy.charCodeAt(0) !== 0xE000)
        return normalizarRef(_fy);
    var _fz = 0xE100, _ga = new Array(_fy.length - 1);
    for (var _gb = 1; _gb < _fy.length; _gb++) {
        var _gc = _fy.charCodeAt(_gb);
        _ga[_gb - 1] = (_gc >= _fz && _gc < 0xF900) ? _gc - _fz : _gc;
    }
    return String.fromCharCode.apply(null, _ga);
}
function codificarRefsExistentes_(_gd) {
    var _ge = PropertiesService.getDocumentProperties();
    if (_ge.getProperty('REFS_CODIFICADAS_V1') === '1')
        return true;
    var _gf = null;
    try {
        _gf = LockService.getDocumentLock();
        if (!_gf || !_gf.tryLock(20000))
            _gf = null;
    }
    catch (_gg) {
        _gf = null;
    }
    if (!_gf)
        return false;
    try {
        var _gh = [[HOJA_AGENCIAS, 6], [HOJA_AGENCIAS_ARCHIVO, 6], [HOJA_ROTURAS, 6], [HOJA_ROTURAS_ARCHIVO, 6],
            [HOJA_OK, 9], [HOJA_OK_ARCHIVO, 9], [HOJA_RECLAMAR, 9], [HOJA_RECLAMAR_ARCHIVO, 9], [HOJA_CONTROL_REVISION, 3]];
        _gh.forEach(function (_gi) {
            var _gj = _gd.getSheetByName(_gi[0]);
            if (!_gj || _gj.getLastRow() < 2 || _gj.getMaxColumns() < _gi[1])
                return;
            var _gk = _gj.getRange(2, _gi[1], _gj.getLastRow() - 1, 1);
            var _gl = _gk.getValues(), _gm = false;
            for (var _gn = 0; _gn < _gl.length; _gn++) {
                var _go = _gl[_gn][0];
                if (_go === "" || _go === null || _go instanceof Date)
                    continue;
                if (String(_go).charAt(0) === MARCA_REF_CODIFICADA_)
                    continue;
                _gl[_gn][0] = codificarRef_(_go);
                _gm = true;
            }
            if (_gm) {
                _gk.setNumberFormat('@');
                _gk.setValues(_gl);
            }
        });
        SpreadsheetApp.flush();
        _ge.setProperty('REFS_CODIFICADAS_V1', '1');
        return true;
    }
    finally {
        try {
            _gf.releaseLock();
        }
        catch (_gp) { }
    }
}
function categoriaEstado_(_gq) {
    var _gr = (_gq || "").toString().trim().toUpperCase();
    if (_gr === "OK")
        return "OK";
    if (_gr === "RECLAMAR" || _gr === "RECLAMAR/PENDIENTE")
        return "FAC";
    if (_gr.indexOf("RECLAMAR") === 0 && claveEstado_(_gr) === "RECLAMARPENDIENTE")
        return "FAC";
    return null;
}
function claveEstado_(_gs) {
    return (_gs || "").toString().toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '')
        .replace(/\bDE\b/g, '').replace(/[^A-Z0-9]/g, '');
}
function valorSegunDesplegable_(_gt, _gu) {
    try {
        var _gv = _gt.getDataValidation();
        if (!_gv)
            return _gu;
        var _gw = _gv.getCriteriaType(), _gx = _gv.getCriteriaValues(), _gy = [];
        if (_gw === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST)
            _gy = _gx[0] || [];
        else if (_gw === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && _gx[0]) {
            _gy = [].concat.apply([], _gx[0].getDisplayValues());
        }
        else
            return _gu;
        if (_gy.indexOf(_gu) > -1)
            return _gu;
        var _gz = claveEstado_(_gu);
        for (var _ha = 0; _ha < _gy.length; _ha++) {
            if (String(_gy[_ha]).trim() !== "" && claveEstado_(_gy[_ha]) === _gz)
                return String(_gy[_ha]);
        }
    }
    catch (_hb) { }
    return _gu;
}
function ultimoEstadoDeFila_(_hc, _hd, _he) {
    var _hf = _hc[_he], _hg = _hd[_he];
    if (_hf && (!_hg || _hf.t >= _hg.t))
        return { cat: "OK", ref: _hf.ref };
    if (_hg)
        return { cat: "FAC", ref: _hg.ref };
    return null;
}
function limpiarRegistrosDeFilasVacias() {
    if (!usuarioAutorizadoOAvisar_())
        return;
    if (!soloAdminOAvisar_())
        return;
    var _hh = SpreadsheetApp.getActiveSpreadsheet(), _hi = SpreadsheetApp.getUi();
    var _hj = planLimpiezaRegistros_(_hh);
    var _hk = _hj.reduce(function (_hl, _hm) { return _hl + _hm.registros; }, 0);
    if (_hk === 0) {
        _hi.alert("🧹 No hay registros de filas borradas: el historial está limpio.");
        return;
    }
    var _hn = _hj.filter(function (_ho) { return _ho.registros > 0; })
        .map(function (_hp) { return "• " + _hp.etiqueta + ": " + _hp.registros; }).join("\n");
    var _hq = _hi.alert("🧹 Quitar registros de filas borradas", "Hay " + _hk + " registros cuya fila ya está vacía:\n\n" + _hn + "\n\n¿Quitarlos? Dejarán de contar en el Recuento y el Calendario.", _hi.ButtonSet.YES_NO);
    if (_hq !== _hi.Button.YES)
        return;
    var _hr = 0;
    _hj.forEach(function (_hs) {
        var _ht = {};
        _hs.filas.forEach(function (_hu) { (_ht[_hu.origen] = _ht[_hu.origen] || []).push(_hu.fila); });
        for (var _hv in _ht)
            _hr += quitarRegistrosDeFilas_(_hh, _hs.hoja, _hv, _ht[_hv], 2, _hs.idxFila);
    });
    _hi.alert("✅ Hecho: se han quitado " + _hr + " registros.");
}
function planLimpiezaRegistros_(_hw) {
    var _hx = [
        { hoja: HOJA_AGENCIAS, etiqueta: "Agencias", idxFila: 4, col: COLUMNA_AGENCIAS, vacia: function (_hy) { return _hy === ""; }, soloTrabajo: false },
        { hoja: HOJA_OK, etiqueta: "OK", idxFila: 6, col: COLUMNA_NOTA_L, vacia: function (_hz) { return _hz === ""; }, soloTrabajo: true },
        { hoja: HOJA_RECLAMAR, etiqueta: "Reclamar", idxFila: 6, col: COLUMNA_NOTA_L, vacia: function (_ia) { return _ia === ""; }, soloTrabajo: true },
        { hoja: HOJA_ROTURAS, etiqueta: "Roturas", idxFila: 4, col: COLUMNA_ROTURA, vacia: function (_ib) { return _ib.toUpperCase() !== "SI"; }, soloTrabajo: true }
    ];
    var _ic = {};
    function _id(_ie, _if, _ig) {
        var _ih = _ie + "|" + _if;
        if (!(_ih in _ic)) {
            var _ii = _hw.getSheetByName(_ie);
            _ic[_ih] = (_ii && _ii.getLastRow() >= 2) ? _ii.getRange(2, _if, _ii.getLastRow() - 1, 1).getDisplayValues() : [];
        }
        var _ij = _ig - 2, _ik = _ic[_ih];
        return (_ij >= 0 && _ij < _ik.length) ? (_ik[_ij][0] || "").toString().trim() : "";
    }
    return _hx.map(function (_il) {
        var _im = { hoja: _il.hoja, etiqueta: _il.etiqueta, idxFila: _il.idxFila, filas: [], registros: 0 };
        var _in = _hw.getSheetByName(_il.hoja);
        if (!_in || _in.getLastRow() < 2)
            return _im;
        var _io = _in.getRange(2, 1, _in.getLastRow() - 1, _in.getLastColumn()).getValues();
        var _ip = {};
        _io.forEach(function (_iq) {
            var _ir = String(_iq[2]), _is = Number(_iq[_il.idxFila]);
            if (HOJAS_TRABAJO_PERMITIDAS_.indexOf(_ir) === -1 || !(_is >= 2))
                return;
            if (_il.soloTrabajo && _ir === HOJA_RETORNOS)
                return;
            var _it = _ir + "#" + _is;
            if (!(_it in _ip)) {
                _ip[_it] = _il.vacia(_id(_ir, _il.col, _is));
                if (_ip[_it])
                    _im.filas.push({ origen: _ir, fila: _is });
            }
            if (_ip[_it])
                _im.registros++;
        });
        return _im;
    });
}
var PLAZO_REGISTRO_ = 0;
function esperaTurnoRegistro_() {
    if (!PLAZO_REGISTRO_)
        return 10000;
    return Math.min(10000, PLAZO_REGISTRO_ - new Date().getTime());
}
function anexarFilasHistorial_(_iu, _iv, _iw, _ix) {
    var _iy = esperaTurnoRegistro_();
    if (_iy < 500) {
        sumarContador_("APLAZADOS", 1);
        anotarEventoSalud_("APLAZ", _iw.length + " apunte" + (_iw.length === 1 ? "" : "s") + " para " + _iv + (_iw[0] && _iw[0][2] ? " (de " + _iw[0][2] + ", fila " + (_iv === HOJA_OK || _iv === HOJA_RECLAMAR ? _iw[0][6] : _iw[0][4]) + ")" : ""));
        return false;
    }
    var _iz = null;
    try {
        _iz = LockService.getDocumentLock();
        if (!_iz || !_iz.tryLock(_iy))
            _iz = null;
    }
    catch (_ja) {
        _iz = null;
    }
    try {
        var _jb = getOrCreateSheet(_iu, _iv);
        try {
            if (_jb.getMaxColumns() < _ix)
                _jb.insertColumnsAfter(_jb.getMaxColumns(), _ix - _jb.getMaxColumns());
        }
        catch (_jc) { }
        if (_iz) {
            _jb.getRange(_jb.getLastRow() + 1, 1, _iw.length, _ix).setValues(_iw);
            SpreadsheetApp.flush();
        }
        else {
            for (var _jd = 0; _jd < _iw.length; _jd++)
                _jb.appendRow(_iw[_jd]);
        }
    }
    finally {
        if (_iz) {
            try {
                _iz.releaseLock();
            }
            catch (_je) { }
        }
    }
    return true;
}
var FILAS_FORMATO_AL_ABRIR_ = 3000;
function asegurarColumnaReferenciaTextoAlAbrir_(_jf) {
    var _jg = _jf.getMaxRows();
    var _jh = Math.max(2, _jf.getLastRow() - FILAS_FORMATO_AL_ABRIR_);
    if (_jg < _jh)
        return;
    asegurarColumnaReferenciaTexto_(_jf, _jh, _jg - _jh + 1);
}
function asegurarColumnaReferenciaTexto_(_ji, _jj, _jk) {
    var _jl = _jj || 2;
    var _jm = _jk || Math.max(_ji.getMaxRows() - 1, 1);
    var _jn = _ji.getRange(_jl, COLUMNA_REFERENCIAS, _jm, 1);
    if (_jj) {
        var _jo = _jn.getNumberFormats();
        var _jp = true;
        for (var _jq = 0; _jq < _jo.length; _jq++) {
            if (_jo[_jq][0] !== '@') {
                _jp = false;
                break;
            }
        }
        if (_jp)
            return;
    }
    _jn.setNumberFormat('@');
}
function procesarAvisosPorEdicion_(_jr, _js, _jt, _ju, _jv) {
    var _jw = _jr.getSheetByName(HOJA_AVISOS);
    if (!_jw || _jw.getLastRow() < 2)
        return;
    var _jx = _jw.getRange(2, 1, _jw.getLastRow() - 1, 5).getValues();
    var _jy = {};
    for (var _jz = 0; _jz < _jx.length; _jz++) {
        var _ka = normalizarRef(_jx[_jz][2]);
        if (_ka !== "")
            _jy[_ka] = _jz + 2;
    }
    if (Object.keys(_jy).length === 0)
        return;
    var _kb = _jt.getRange(_ju, COLUMNA_REFERENCIAS, _jv, 1).getDisplayValues();
    var _kc = new Date(), _kd = "Usuario anónimo";
    try {
        _kd = Session.getActiveUser().getEmail() || "Usuario anónimo";
    }
    catch (_ke) { }
    var _kf = getOrCreateSheet(_jr, HOJA_AVISOS_EVENTOS);
    var _kg = proximoIdEventoAviso_(_kf);
    var _kh = [];
    for (var _ki = 0; _ki < _jv; _ki++) {
        var _kj = (_kb[_ki][0] || "").toString().trim();
        if (_kj === "")
            continue;
        var _kk = normalizarRef(_kj);
        var _kl = _jy[_kk];
        if (!_kl)
            continue;
        var _km = _ju + _ki;
        var _kn = _js + ", fila " + _km;
        var _ko = _jw.getRange(_kl, 4).getValue();
        var _kp = _jw.getRange(_kl, 5).getValue();
        _kp = (typeof _kp === 'number' ? _kp : 0) + 1;
        _jw.getRange(_kl, 5, 1, 3).setValues([[_kp, _kc, _kn]]);
        _kh.push([_kg + _kh.length, _kc, _kk, _kj, _js, _km, _ko, "NUEVA"]);
    }
    if (_kh.length > 0) {
        var _kq = _kf.getLastRow() + 1, _kr = _kq + _kh.length - 1;
        if (_kr > _kf.getMaxRows())
            _kf.insertRowsAfter(_kf.getMaxRows(), _kr - _kf.getMaxRows());
        _kf.getRange(_kq, 3, _kh.length, 2).setNumberFormat('@');
        _kf.getRange(_kq, 1, _kh.length, 8).setValues(_kh);
        guardarUltimoIdAviso_(_kh[_kh.length - 1][0]);
        MENSAJE_AVISO_PENDIENTE_ = _kh.map(function (_ks) {
            var _kt = (_ks[6] || "").toString().replace(/\s+/g, " ").trim();
            if (_kt.length > 160)
                _kt = _kt.substring(0, 157) + "...";
            return "Pedido " + _ks[3] + ": " + (_kt || "(sin motivo)");
        }).join("  |  ");
        try {
            construirMenu();
        }
        catch (_ku) { }
    }
}
var MENSAJE_AVISO_PENDIENTE_ = null;
var MENSAJE_ERROR_AVISO_ = null;
function proximoIdEventoAviso_(_kv) {
    var _kw = _kv.getLastRow();
    if (_kw < 2)
        return 1;
    var _kx = _kv.getRange(_kw, 1).getValue();
    var _ky = (typeof _kx === 'number') ? _kx : parseInt(_kx, 10);
    return (isNaN(_ky) ? (_kw - 1) : _ky) + 1;
}
function onOpen() {
    try {
        ocultarHistorialesInterno();
    }
    catch (_kz) { }
    if (!usuarioAutorizado_())
        return;
    try {
        PropertiesService.getUserProperties().setProperty('carga_doc', String(new Date().getTime()));
    }
    catch (_la) { }
    try {
        var _lb = construirMenu();
        if (_lb)
            SpreadsheetApp.getActiveSpreadsheet().toast("⚠️ ATENCIÓN: Tienes notas pendientes de leer.", "🔔 NOTA PENDIENTE", 10);
    }
    catch (_lc) {
        try {
            SpreadsheetApp.getActiveSpreadsheet().toast("No se pudo crear el menú: " + _lc.message, "⚠️ Error", 8);
        }
        catch (_ld) { }
    }
    try {
        for (var _le = 0; _le < HOJAS_TRABAJO_PERMITIDAS_.length; _le++) {
            var _lf = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJAS_TRABAJO_PERMITIDAS_[_le]);
            if (!_lf)
                continue;
            try {
                asegurarColumnaReferenciaTextoAlAbrir_(_lf);
            }
            catch (_lg) { }
        }
    }
    catch (_lh) { }
}
function contarNotasNuevas_() {
    var _li = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_NOTAS);
    if (!_li || _li.getLastRow() < 2 || _li.getLastColumn() < 4)
        return 0;
    var _lj = _li.getRange(2, 4, _li.getLastRow() - 1, 1).getValues();
    var _lk = 0;
    for (var _ll = 0; _ll < _lj.length; _ll++) {
        if (_lj[_ll][0] === "NUEVA")
            _lk++;
    }
    return _lk;
}
function contarAvisosNuevos_() {
    var _lm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AVISOS_EVENTOS);
    if (!_lm || _lm.getLastRow() < 2)
        return 0;
    var _ln = _lm.getRange(2, 8, _lm.getLastRow() - 1, 1).getValues();
    var _lo = 0;
    for (var _lp = 0; _lp < _ln.length; _lp++) {
        if (_ln[_lp][0] === "NUEVA")
            _lo++;
    }
    return _lo;
}
function construirMenu() {
    var _lq = contarNotasNuevas_() > 0;
    var _lr = _lq ? '📌 Ver notas activas 🔴 (NUEVAS)' : '📌 Ver notas activas';
    var _ls = SpreadsheetApp.getUi().createMenu('📦 Devoluciones')
        .addItem('📊 Ver recuento por turnos y roturas', 'mRecuento')
        .addItem('📅 Calendario de Agencias', 'mCalendario')
        .addItem('🔍 Buscar Referencia', 'mBuscar')
        .addSeparator()
        .addItem('✍️ Dejar nota para turno', 'mEscribir')
        .addItem(_lr, 'mVerNotas')
        .addItem('⚙️ Gestionar notas', 'mGestionar')
        .addSeparator();
    if (AVISOS_PEDIDOS_ACTIVOS_) {
        var _lt = contarAvisosNuevos_();
        var _lu = _lt > 0 ? '🔔 Avisos de Pedidos 🔴 (' + _lt + ')' : '🔔 Avisos de Pedidos';
        _ls.addItem(_lu, 'mAvisos').addSeparator();
    }
    var _lv = esAdmin_() !== false;
    if (_lv) {
        _ls
            .addItem('🕰️ Auto-Archivado', 'activarDisparadorDiario')
            .addItem('✅ Comprobar que todo está contado', 'revisarRegistrosAhora')
            .addItem('🩺 Estado del sistema', 'estadoDelSistema')
            .addItem('🧹 Quitar registros de filas borradas', 'limpiarRegistrosDeFilasVacias')
            .addItem('📂 Mostrar / Ocultar Historiales', 'mHistoriales')
            .addSeparator();
    }
    _ls.addItem('🐾 Abrir a Rufo', 'mRufo');
    if (_lv)
        _ls.addItem('🐾 Activar apertura automática de Rufo', 'activarAperturaAutomaticaRufo');
    _ls.addToUi();
    return _lq;
}
function mRecuento() { abrirMotorApp("recuento", 850, 500, "📊 Recuento y Roturas"); }
function mCalendario() { abrirMotorApp("calendario", 950, 650, "📅 Calendario de Agencias"); }
function mBuscar() { abrirMotorLateral("buscar", "🔍 Buscar Referencia"); }
function mRufo() { abrirMotorLateral("rufo", "🐾 Rufo"); }
function mRufoRapido() { abrirMotorLateral("rufo_abierto", "🐾 Rufo"); }
function mEscribir() { abrirMotorApp("escribir", 600, 450, "📝 Escribir Nota"); }
function mVerNotas() { abrirMotorApp("vernotas", 650, 550, "📌 Tablón de Notas"); }
function mGestionar() { abrirMotorApp("gestionar", 750, 500, "⚙️ Gestionar Notas"); }
function mNotas() { abrirMotorApp("notas", 650, 600, "🗒️ Notas"); }
function mAvisos() { abrirMotorApp("avisos", 650, 600, "🔔 Avisos de Pedidos"); }
function mHistoriales() { if (!soloAdminOAvisar_())
    return; abrirMotorApp("historiales", 450, 300, "📂 Historiales"); }
function usuarioAutorizadoOAvisar_() {
    if (usuarioAutorizado_())
        return true;
    try {
        SpreadsheetApp.getUi().alert("🔒 No tienes permiso para usar este sistema.");
    }
    catch (_lw) { }
    return false;
}
function cargaDocUsuario_() {
    try {
        return String(PropertiesService.getUserProperties().getProperty('carga_doc') || "");
    }
    catch (_lx) {
        return "";
    }
}
function abrirMotorApp(_ly, _lz, _ma, _mb) {
    if (!usuarioAutorizadoOAvisar_())
        return;
    var _mc = HtmlService.createTemplate(obtenerHTMLBase());
    _mc.accion = _ly;
    _mc.htmlInicial = htmlInicialPanel_(_ly);
    _mc.cargaDoc = cargaDocUsuario_();
    var _md = _mc.evaluate().setWidth(_lz).setHeight(_ma);
    SpreadsheetApp.getUi().showModalDialog(_md, _mb);
}
function comprobarAccesoPanel() {
    return usuarioAutorizado_() === true;
}
var LATERALES_UNA_PIEZA_ = ["rufo", "rufo_abierto", "buscar", "retornos", "cambios", "pendientes"];
function obtenerSeccionesLaterales() {
    if (!usuarioAutorizado_())
        return null;
    return { v: VERSION_SISTEMA_, html: { rufo_abierto: genHtmlRufo(true), buscar: genHtmlBuscar(), retornos: genHtmlRetornos(), cambios: genHtmlCambiosRetornos(), pendientes: genHtmlPendientes() } };
}
var ACCIONES_PRECARGADAS_ = ["buscar", "rufo", "rufo_abierto", "avisos", "retornos", "cambios"];
function htmlInicialPanel_(_me) {
    if (_me === "calendario") {
        try {
            return genHtmlCalendario();
        }
        catch (_mf) {
            return "";
        }
    }
    if (ACCIONES_PRECARGADAS_.indexOf(_me) === -1)
        return "";
    try {
        var _mg = enrutadorApp('', _me);
        return (_mg && _mg.html) ? _mg.html : "";
    }
    catch (_mh) {
        try {
            Logger.log('htmlInicialPanel_(' + _me + ') ha fallado: ' + (_mh && _mh.message ? _mh.message : _mh));
        }
        catch (_mi) { }
        return "";
    }
}
function jsonSeguroParaScript_(_mj) {
    return JSON.stringify(_mj || "").replace(/</g, '\\u003c');
}
function abrirMotorLateral(_mk, _ml) {
    if (!usuarioAutorizadoOAvisar_())
        return;
    var _mm = HtmlService.createTemplate(obtenerHTMLBase());
    _mm.accion = _mk;
    _mm.htmlInicial = htmlInicialPanel_(_mk);
    _mm.cargaDoc = cargaDocUsuario_();
    var _mn = _mm.evaluate().setTitle(_ml);
    SpreadsheetApp.getUi().showSidebar(_mn);
}
function enrutadorApp(_mo, _mp) {
    var _mq = ["escribir", "vernotas", "gestionar", "notas"];
    if (_mq.indexOf(_mp) > -1 && _mo !== CONTRASEÑA_NOTAS)
        return { error: true };
    if (_mp === "recuento")
        return { html: genHtmlRecuento() };
    if (_mp === "calendario")
        return { html: genHtmlCalendario() };
    if (_mp === "buscar")
        return { html: genHtmlBuscar() };
    if (_mp === "pendientes")
        return { html: genHtmlPendientes() };
    if (_mp === "rufo")
        return { html: genHtmlRufo() };
    if (_mp === "rufo_abierto")
        return { html: genHtmlRufo(true) };
    if (_mp === "escribir")
        return { html: genHtmlEscribir() };
    if (_mp === "vernotas")
        return { html: genHtmlVerNotas() };
    if (_mp === "gestionar")
        return { html: genHtmlGestionarNotas() };
    if (_mp === "notas")
        return { html: genHtmlNotas(true) };
    if (_mp === "avisos")
        return { html: genHtmlAvisos(true) };
    if (_mp === "retornos")
        return { html: genHtmlRetornos() };
    if (_mp === "cambios")
        return { html: genHtmlCambiosRetornos() };
    if (_mp === "historiales") {
        if (esAdmin_() !== true)
            return { html: "<div class='result-box'><div class='result-box-icon'>🔒</div><h3>Solo para la administración del sistema.</h3><p>Puedes cerrar esta ventana.</p></div>" };
        var _mr = alternarHistoriales();
        return { html: "<div class='result-box'><div class='result-box-icon'>📂</div><h3>" + _mr + "</h3><p>Puedes cerrar esta ventana.</p></div>" };
    }
}
function contarTraidasPorAgenciaHoy_(_ms, _mt) {
    var _mu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AGENCIAS);
    if (!_mu)
        return contarTraidasEnFilas_(_ms, [], _mt);
    var _mv = _mu.getDataRange().getValues();
    return contarTraidasEnFilas_(_ms, _mv.slice(1), _mt);
}
function contarTraidasEnFilas_(_mw, _mx, _my) {
    var _mz = {};
    _mw.forEach(function (_na) { _mz[_na] = 0; });
    var _nb = {};
    for (var _nc = 0; _nc < _mx.length; _nc++) {
        if (esMismoDia(new Date(_mx[_nc][0]), _my)) {
            _nb[_mx[_nc][2] + "_" + _mx[_nc][4]] = { ag: (_mx[_nc][3] || "").toString().toUpperCase().trim(), origen: _mx[_nc][2] };
        }
    }
    for (var _nd in _nb) {
        if (_nb[_nd].ag !== "" && _nb[_nd].origen !== HOJA_RETORNOS) {
            _mw.forEach(function (_ne) { if (_nb[_nd].ag.indexOf(_ne) !== -1)
                _mz[_ne]++; });
        }
    }
    return _mz;
}
var COMPAÑIAS_RECUENTO_ = ["VELOX", "PAQNORTE", "CORREOMAX", "RUTASUR", "ATLAS", "PRONTO", "BOLIDO", "FARO"];
var DIAS_SEMANA_RECUENTO_ = ["domingo", "lunes", "martes", "miércoles", "jueves", "viernes", "sábado"];
var MESES_RECUENTO_ = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
function genHtmlRecuento() {
    var _nf = SpreadsheetApp.getActiveSpreadsheet(), _ng = new Date();
    var _nh = COMPAÑIAS_RECUENTO_;
    function _ni(_nj) {
        var _nk = _nf.getSheetByName(_nj);
        if (!_nk)
            return [];
        return _nk.getDataRange().getValues().slice(1);
    }
    var _nl = _ni(HOJA_RECLAMAR), _nm = _ni(HOJA_OK);
    var _nn = [];
    var _no = _nf.getSheetByName(HOJA_ROTURAS);
    if (_no && _no.getLastRow() >= 2) {
        _nn = _no.getRange(2, 1, _no.getLastRow() - 1, 5).getValues();
    }
    var _np = calcularRecuentoDia_(_nl, _nm, _nn, _ng);
    _np.traidas = contarTraidasPorAgenciaHoy_(_nh, _ng);
    return `
    ${genEncabezadoRastro_('📊', 'Recuento de hoy', 'OK, Reclamar y roturas', 'recuento')}
    <div class="rc-barra" id="rc-barra" data-hoy="${claveFecha_(_ng)}">
      <span class="rc-dia" id="rc-dia">${tituloDiaRecuento_(_ng, true)}</span>
      <button class="rc-btn rc-volver" id="rc-volver" hidden onclick="rcVolverHoy()">↩️ Volver a hoy</button>
      <button class="rc-btn" id="rc-abrir" onclick="rcCalendario()">📅 Otros días</button>
    </div>
    <div class="rc-cal" id="rc-cal" hidden></div>
    <div id="rc-cuerpo">${htmlCuerpoRecuento_(_np)}</div>`;
}
function tituloDiaRecuento_(_nq, _nr) {
    var _ns = ('0' + _nq.getDate()).slice(-2), _nt = ('0' + (_nq.getMonth() + 1)).slice(-2);
    return (_nr ? "Hoy · " : "") + DIAS_SEMANA_RECUENTO_[_nq.getDay()] + " " + _ns + "/" + _nt + (_nr ? "" : "/" + _nq.getFullYear());
}
function calcularRecuentoDia_(_nu, _nv, _nw, _nx) {
    var _ny = COMPAÑIAS_RECUENTO_;
    var _nz = 0, _oa = 0, _ob = 0, _oc = 0;
    var _od = [];
    function _oe(_of, _og, _oh) {
        for (var _oi = 0; _oi < _of.length; _oi++) {
            var _oj = new Date(_of[_oi][0]);
            if (!esMismoDia(_oj, _nx))
                continue;
            var _ok = _of[_oi][2] !== "" && _of[_oi][2] !== undefined && Number(_of[_oi][6]) >= 2;
            _od.push({ clave: _ok ? _of[_oi][2] + "#" + Number(_of[_oi][6]) : "suelto#" + _og + "#" + _oi, cat: _oh, t: _oj.getTime(), fr: _oj });
        }
    }
    _oe(_nu, HOJA_RECLAMAR, "FAC");
    _oe(_nv, HOJA_OK, "OK");
    _od.sort(function (_ol, _om) { return _ol.t - _om.t; });
    var _on = {};
    _od.forEach(function (_oo) {
        var _op = _on[_oo.clave];
        if (!_op || _op.cat !== _oo.cat)
            _on[_oo.clave] = { cat: _oo.cat, fr: _oo.fr };
    });
    for (var _oq in _on) {
        var _or = _on[_oq];
        var _os = _or.fr.getHours() + (_or.fr.getMinutes() / 60);
        if (_os >= 5 && _os < 13) {
            _or.cat === "FAC" ? _nz++ : _oa++;
        }
        else if (_os >= 13 && _os < 20.5) {
            _or.cat === "FAC" ? _ob++ : _oc++;
        }
    }
    var _ot = 0, _ou = 0, _ov = 0, _ow = {};
    _ny.forEach(function (_ox) { _ow[_ox] = 0; });
    var _oy = {};
    for (var _oz = 0; _oz < _nw.length; _oz++) {
        var _pa = new Date(_nw[_oz][0]);
        if (!esMismoDia(_pa, _nx))
            continue;
        if (_nw[_oz][2] !== "" && Number(_nw[_oz][4]) >= 2) {
            var _pb = _nw[_oz][2] + "#" + Number(_nw[_oz][4]);
            if (_oy[_pb])
                continue;
            _oy[_pb] = true;
        }
        _ot++;
        var _pc = _pa.getHours() + (_pa.getMinutes() / 60);
        if (_pc >= 5 && _pc < 13)
            _ou++;
        else if (_pc >= 13 && _pc < 20.5)
            _ov++;
        var _pd = (_nw[_oz][3] || "").toString().trim().toUpperCase();
        _ny.forEach(function (_pe) { if (_pd.indexOf(_pe) !== -1)
            _ow[_pe]++; });
    }
    return { mf: _nz, mok: _oa, tf: _ob, tok: _oc, totalRoturas: _ot, mRot: _ou, tRot: _ov, roturasPorAgencia: _ow, traidas: null };
}
function htmlCuerpoRecuento_(_pf) {
    var _pg = COMPAÑIAS_RECUENTO_;
    var _ph = _pf.mf, _pi = _pf.mok, _pj = _pf.tf, _pk = _pf.tok, _pl = _pf.totalRoturas, _pm = _pf.mRot, _pn = _pf.tRot;
    var _po = _pf.roturasPorAgencia, _pp = _pf.traidas || {};
    var _pq = { "VELOX": { c: "#2F6FE4", t: "VELOX" }, "PAQNORTE": { c: "#0F8A7E", t: "PAQN" }, "CORREOMAX": { c: "#C99700", t: "CMAX" }, "RUTASUR": { c: "#D85A30", t: "RSUR" }, "ATLAS": { c: "#6D4BC9", t: "ATLAS" }, "PRONTO": { c: "#2E9E44", t: "PRON" }, "BOLIDO": { c: "#C62E3B", t: "BOLI" }, "FARO": { c: "#26354A", t: "FARO" } };
    function _pr(_ps) {
        var _pt = "<span class='ag-badge-label'>" + _ps + "</span>";
        if (AG_LOGOS_B64_[_ps])
            return "<span class='ag-badge-card ag-badge-sm' title='" + _ps + "'><span class='ag-badge-logo'><img src='data:image/png;base64," + AG_LOGOS_B64_[_ps] + "' alt='" + _ps + "'></span>" + _pt + "</span>";
        var _pu = _pq[_ps] || { c: "#64748b", t: _ps.substring(0, 3) };
        return "<span class='ag-badge-card ag-badge-sm ag-badge-color' title='" + _ps + "'><span class='ag-badge-logo' style='background:" + _pu.c + "'>" + _pu.t + "</span>" + _pt + "</span>";
    }
    var _pv = _pg.filter(function (_pw) { return _po[_pw] > 0; }).map(function (_px) {
        var _py = _po[_px], _pz = _pp[_px] || 0;
        var _qa = "—", _qb = 0;
        if (_py > 0 && _pz > 0) {
            _qb = Math.round((_py / _pz) * 1000) / 10;
            _qa = (_qb % 1 === 0 ? _qb.toFixed(0) : _qb.toFixed(1)) + "%";
        }
        return `<div class="bar-row"><div class="bar-info"><span style="display:inline-flex;align-items:center;gap:6px;">${_pr(_px)} (${_py} rotos)</span><span class="bar-pct">${_qa}</span></div><div class="bar-track"><div class="bar-fill" style="width: ${_qb}%;"></div></div></div>`;
    }).join('') || "<div class='empty-inline'>✅ ¡Excelente! Cero roturas.</div>";
    return `
    <div class="grid">
      <div class="card ok"><div class="icon">✅</div><div class="title">Total OK</div><div class="value">${_pi + _pk}</div></div>
      <div class="card fact"><div class="icon">🧾</div><div class="title">Reclamar</div><div class="value">${_ph + _pj}</div></div>
      <div class="card rotura"><div class="icon">💥</div><div class="title">Rotos</div><div class="value">${_pl}</div></div>
    </div>
    <div class="details-container">
      <div class="details"><h3 class="dh3">🕒 Turnos</h3><div class="turn"><span>Mañana OK:</span> <strong>${_pi}</strong></div><div class="turn"><span>Mañana Recl.:</span> <strong>${_ph}</strong></div><div class="turn"><span>Mañana Rotos:</span> <strong style="color:var(--red-dark);">${_pm}</strong></div><div class="turn"><span>Tarde OK:</span> <strong>${_pk}</strong></div><div class="turn"><span>Tarde Recl.:</span> <strong>${_pj}</strong></div><div class="turn"><span>Tarde Rotos:</span> <strong style="color:var(--red-dark);">${_pn}</strong></div></div>
      <div class="details"><h3 class="dh3">💥 % Roturas (Agencias)</h3><div style="max-height:160px; overflow-y:auto;">${_pv}</div></div>
    </div>`;
}
function filasHistorialEnRango_(_qc, _qd, _qe, _qf, _qg) {
    var _qh = [];
    var _qi = new Date().getTime() - (DIAS_CADUCIDAD_HISTORIALES - 1) * 86400000;
    if (_qe && _qf.getTime() < _qi) {
        var _qj = _qc.getSheetByName(_qe);
        if (_qj)
            _qh = leerFilasEnRangoFechas_(_qj, _qf, _qg);
    }
    var _qk = _qc.getSheetByName(_qd);
    if (_qk)
        _qh = _qh.concat(leerFilasEnRangoFechas_(_qk, _qf, _qg));
    return _qh;
}
function leerCacheRecuento_(_ql) {
    try {
        var _qm = CacheService.getDocumentCache().get(_ql);
        return _qm ? JSON.parse(_qm) : null;
    }
    catch (_qn) {
        return null;
    }
}
function guardarCacheRecuento_(_qo, _qp, _qq) {
    try {
        CacheService.getDocumentCache().put(_qo, JSON.stringify(_qp), _qq);
    }
    catch (_qr) { }
}
function obtenerMesRecuento(_qs) {
    if (!usuarioAutorizado_())
        return { error: "Herramienta de uso exclusivo para Devoluciones." };
    var _qt = Math.round(Number(_qs) || 0);
    _qt = Math.max(-24, Math.min(0, _qt));
    var _qu = new Date();
    var _qv = new Date(_qu.getFullYear(), _qu.getMonth() + _qt, 1, 0, 0, 0, 0);
    var _qw = new Date(_qv.getFullYear(), _qv.getMonth() + 1, 0, 23, 59, 59, 999);
    var _qx = "REC_MES_1_" + claveFecha_(_qv) + "_" + claveFecha_(_qu);
    var _qy = leerCacheRecuento_(_qx);
    if (_qy)
        return _qy;
    var _qz = SpreadsheetApp.getActiveSpreadsheet();
    var _ra = {};
    function _rb(_rc, _rd) {
        for (var _re = 0; _re < _rc.length; _re++) {
            var _rf = new Date(_rc[_re][0]);
            if (isNaN(_rf.getTime()) || _rf < _qv || _rf > _qw)
                continue;
            var _rg = claveFecha_(_rf);
            if (!_ra[_rg])
                _ra[_rg] = { fac: [], ok: [], rot: [] };
            _ra[_rg][_rd].push(_rc[_re]);
        }
    }
    _rb(filasHistorialEnRango_(_qz, HOJA_RECLAMAR, HOJA_RECLAMAR_ARCHIVO, _qv, _qw), "fac");
    _rb(filasHistorialEnRango_(_qz, HOJA_OK, HOJA_OK_ARCHIVO, _qv, _qw), "ok");
    _rb(filasHistorialEnRango_(_qz, HOJA_ROTURAS, HOJA_ROTURAS_ARCHIVO, _qv, _qw), "rot");
    var _rh = {}, _ri = { ok: 0, fac: 0, rot: 0 };
    for (var _rj in _ra) {
        var _rk = new Date(Number(_rj.slice(0, 4)), Number(_rj.slice(4, 6)) - 1, Number(_rj.slice(6, 8)));
        var _rl = calcularRecuentoDia_(_ra[_rj].fac, _ra[_rj].ok, _ra[_rj].rot, _rk);
        var _rm = { ok: _rl.mok + _rl.tok, fac: _rl.mf + _rl.tf, rot: _rl.totalRoturas };
        if (_rm.ok || _rm.fac || _rm.rot) {
            _rh[_rj] = _rm;
            _ri.ok += _rm.ok;
            _ri.fac += _rm.fac;
            _ri.rot += _rm.rot;
        }
    }
    var _rn = {
        off: _qt, anio: _qv.getFullYear(), mes: _qv.getMonth(),
        titulo: MESES_RECUENTO_[_qv.getMonth()] + " " + _qv.getFullYear(),
        primerDia: (_qv.getDay() + 6) % 7,
        diasMes: new Date(_qv.getFullYear(), _qv.getMonth() + 1, 0).getDate(),
        hoy: claveFecha_(_qu), dias: _rh, total: _ri,
        anterior: _qt > -24, siguiente: _qt < 0
    };
    guardarCacheRecuento_(_qx, _rn, _qt < 0 ? 3600 : 60);
    return _rn;
}
function obtenerRecuentoDia(_ro) {
    if (!usuarioAutorizado_())
        return { error: "Herramienta de uso exclusivo para Devoluciones." };
    _ro = String(_ro || "");
    if (!/^\d{8}$/.test(_ro))
        return { error: "Fecha no válida." };
    var _rp = new Date(Number(_ro.slice(0, 4)), Number(_ro.slice(4, 6)) - 1, Number(_ro.slice(6, 8)));
    if (claveFecha_(_rp) !== _ro)
        return { error: "Fecha no válida." };
    var _rq = new Date(), _rr = claveFecha_(_rq) === _ro;
    if (!_rr && _rp > _rq)
        return { error: "Ese día todavía no ha llegado." };
    var _rs = "REC_DIA_1_" + _ro;
    var _rt = _rr ? null : leerCacheRecuento_(_rs);
    if (!_rt) {
        var _ru = SpreadsheetApp.getActiveSpreadsheet();
        var _rv = new Date(_rp.getFullYear(), _rp.getMonth(), _rp.getDate(), 0, 0, 0, 0);
        var _rw = new Date(_rp.getFullYear(), _rp.getMonth(), _rp.getDate(), 23, 59, 59, 999);
        _rt = calcularRecuentoDia_(filasHistorialEnRango_(_ru, HOJA_RECLAMAR, HOJA_RECLAMAR_ARCHIVO, _rv, _rw), filasHistorialEnRango_(_ru, HOJA_OK, HOJA_OK_ARCHIVO, _rv, _rw), filasHistorialEnRango_(_ru, HOJA_ROTURAS, HOJA_ROTURAS_ARCHIVO, _rv, _rw), _rp);
        _rt.traidas = contarTraidasEnFilas_(COMPAÑIAS_RECUENTO_, filasHistorialEnRango_(_ru, HOJA_AGENCIAS, HOJA_AGENCIAS_ARCHIVO, _rv, _rw), _rp);
        if (!_rr)
            guardarCacheRecuento_(_rs, _rt, 600);
    }
    return { clave: _ro, esHoy: _rr, titulo: tituloDiaRecuento_(_rp, _rr), html: htmlCuerpoRecuento_(_rt) };
}
function genHtmlCalendario() {
    return `
    <div class="no-print">${genEncabezadoRastro_('📅', 'Calendario de Agencias', 'Roturas y pendientes por día', 'calendario')}</div>
    <div class="header-cal no-print">
      <button class="btn-cal" onclick="cargarCalendario(-1)">◀️ Ant.</button>
      <h2 id="mes-cal">Cargando...</h2>
      <div class="header-cal-actions">
        <button class="btn-cal btn-cal-orange" onclick="abrirPreviewChecklist('dia')">📧 Checklist (Hoy)</button>
        <button class="btn-cal btn-cal-red" onclick="abrirChecklistSemanal()">📧 Checklist (Semana)</button>
      </div>
      <button class="btn-cal" onclick="cargarCalendario(1)">Sig. ▶️</button>
    </div>
    <style>
      .cal-envios { display:flex; align-items:center; gap:4px; flex-wrap:wrap; margin:-2px 2px 10px; font-size:11.5px; color:#64748b; }
      .cal-envios[hidden] { display:none; }
      .ce-t { font-weight:700; margin-right:3px; }
      .ce-d { display:inline-flex; align-items:center; justify-content:center; min-width:22px; height:20px; padding:0 4px; border-radius:6px; font-size:10.5px; font-weight:800; background:#f1f5f9; color:#94a3b8; cursor:default; }
      .ce-d.si { background:#dcfce7; color:#15803d; } .ce-d.no { background:#fef2f2; color:#dc2626; }
      .ce-d.hoy { background:#fff; color:#b45309; box-shadow:inset 0 0 0 1.5px #f59e0b; }
      .ce-sep { width:1px; height:14px; background:#e2e8f0; margin:0 4px; }
      .ce-btn { border:none; font-family:inherit; cursor:pointer; } .ce-btn:hover { filter:brightness(.95); text-decoration:underline; }
      .ce-enviar { border:none; background:none; color:#E07A1F; font-weight:800; font-size:11.5px; cursor:pointer; padding:0 2px; text-decoration:underline; font-family:inherit; }
    </style>
    <div id="cal-envios" class="cal-envios no-print" hidden></div>
    <div id="ac-linea" class="ac-linea no-print" hidden></div>
    <div id="load-cal" class="loading-row"><span class="spinner"></span>Procesando datos del calendario...</div>
    <div class="grid-cal" id="grid-cal" style="display:none;"></div>

    <div class="cal-legend" id="cal-legend" style="display:none;">
      <div class="legend-item"><span class="legend-icon legend-roto">💥</span><span>Roto</span></div>
      <div class="legend-item"><span class="legend-icon legend-pend">⏳</span><span>Pendiente por abrir</span></div>
      <div class="legend-item"><span class="legend-icon legend-retornos">↩️</span><span>Retornos</span></div>
    </div>

    <div class="no-print" id="pendientes-wrap" style="display:none;">
      <button class="btn-cal btn-cal-amber" id="btn-pendientes" style="width:100%; margin-top:14px;" onclick="abrirPendientesEnLateral()">📋 Ver casillas vacías por compañía</button>
    </div>

    
    <div class="checklist-overlay no-print" id="checklist-overlay" hidden>
      <div class="checklist-card">
        <div class="checklist-card-head">
          <h3 id="checklist-card-titulo">Vista previa del checklist</h3>
          <button class="checklist-card-close" onclick="cerrarPreviewChecklist()">✕</button>
        </div>
        <div class="checklist-card-body">
          <div id="checklist-step-preview">
            <p class="checklist-warn" id="checklist-warn" style="display:none;">Todavía no hay un correo real configurado para Luis Prado (falta rellenar CORREO_JEFATURA en el código). Esta vista previa es solo para revisar cómo quedaría el contenido; no se enviará nada hasta que se indique esa dirección.</p>
            <div class="checklist-doc-preview" id="checklist-doc-preview">Cargando…</div>
          </div>
          <div id="checklist-step-confirm" style="display:none;">
            <p style="margin:0 0 6px; font-size:13.5px;">¿Enviar este checklist por correo a <strong>Luis Prado</strong>?</p>
            
            <p style="margin:0 0 6px; font-size:11.5px; color:var(--muted);" id="checklist-confirm-aviso-prueba"></p>
            <p style="margin:0; font-size:12px; color:var(--muted);" id="checklist-confirm-asunto"></p>
          </div>
          <div id="checklist-step-result" style="display:none;">
            <p class="checklist-ok-msg" id="checklist-result-msg"></p>
            <div id="checklist-outlook-info" style="display:none;"></div>
          </div>
        </div>
        <div class="checklist-card-foot">
          <div class="checklist-foot-group" id="checklist-foot-preview">
            <button class="btn-checklist-cancel" onclick="cerrarPreviewChecklist()">Cancelar</button>
            <button class="btn-checklist-print" onclick="imprimirDesdeVistaPreviaChecklist()">🖨️ Imprimir / Guardar PDF</button>
            <button class="btn-checklist-confirm" id="btn-checklist-outlook" onclick="prepararCorreoOutlook()">📧 Preparar correo en Outlook</button>
          </div>
          <div class="checklist-foot-group" id="checklist-foot-confirm" style="display:none;">
            <button class="btn-checklist-cancel" onclick="volverAPreviewChecklist()">← Volver</button>
            <button class="btn-checklist-confirm" id="btn-checklist-confirmar-envio" onclick="confirmarEnvioChecklist()">Sí, enviar</button>
          </div>
          <div class="checklist-foot-group" id="checklist-foot-result" style="display:none;">
            <button class="btn-checklist-confirm" onclick="cerrarPreviewChecklist()">Cerrar</button>
          </div>
        </div>
      </div>
    </div>

    
    <div class="print-view" id="print-view">
      <div class="print-letterhead">
        <div class="print-letterhead-brand">
          <span class="print-letterhead-logo">RAS<span class="print-letterhead-logo-accent">TRO</span></span>
          <span class="print-letterhead-dept">Departamento de Devoluciones</span>
        </div>
        <div class="print-letterhead-title">
          <h1 id="print-titulo">Calendario de Agencias</h1>
          <p id="print-subtitulo">Roturas y pendientes por día</p>
        </div>
        <div class="print-letterhead-meta">
          <div><strong>Periodo:</strong> <span id="print-rango">—</span></div>
          <div><strong>Generado:</strong> <span id="print-fecha-gen">—</span></div>
          <div><strong>Por:</strong> <span id="print-usuario">—</span></div>
        </div>
      </div>
      <div class="print-tables" id="print-tables"></div>
      <div class="print-footer">Documento de uso interno · Sistema de Auditoría, Agencias y Tablón de Notas · Rastro</div>
    </div>
  `;
}
function genHtmlPendientes() {
    return `
    <button class="volver-rufo" onclick="irSeccion_('rufo_abierto', '🐾 Rufo')">🐾 Volver a Rufo</button>
    ${genEncabezadoRastro_('📋', 'Casillas Pendientes', 'Vacías por compañía', 'casillas')}
    <div class="header-cal">
      <button class="btn-cal" onclick="cargarPendientesLateral(-1)">◀️ Ant.</button>
      <h2 id="mes-pend" style="font-size:12px; text-align:center; line-height:1.35; flex:1; min-width:140px;">Cargando...</h2>
      <div class="header-cal-actions">
        <button class="btn-cal" onclick="cargarPendientesLateral(0)" title="Actualizar">🔄</button>
        <button class="btn-cal" onclick="cargarPendientesLateral(1)">Sig. ▶️</button>
      </div>
    </div>
    <div id="load-pend" class="loading-row"><span class="spinner"></span>Buscando casillas vacías...</div>
    <div id="panel-pend-lateral" style="display:none;"></div>
  `;
}
function genHtmlCambiosRetornos() {
    return `
    <button class="volver-rufo" onclick="irSeccion_('retornos', '↩️ Retornos')">↩️ Volver a Retornos</button>
    ${genEncabezadoRastro_('🕵️', 'Cambios en Retornos', 'Qué se ha cambiado o borrado, quién y cuándo', 'buscar')}
    <div id="cr-filtros" class="cr-filtros"></div>
    <div id="cr-lista"><div class="loading-row"><span class="spinner"></span>Buscando cambios...</div></div>
    <p class="cr-nota">Aquí salen los de los últimos 7 días (se guardan 30). Para ver los de un retorno concreto, búscalo con el Rufi y pulsa «🕵️ Cambios».</p>
  `;
}
function genHtmlBuscar() {
    return `
    <button class="volver-rufo" onclick="irSeccion_('rufo_abierto', '🐾 Rufo')">🐾 Volver a Rufo</button>
    ${genEncabezadoRastro_('🔍', 'Buscar Referencia', 'Escribe o escanea y pulsa Intro', 'buscar')}
    <div class="mp-zona">${cabezaMiniRufo_('retornos')}<div class="search-card">
      <div class="search-row">
        <input type="text" id="b-ref" class="field field-ref" placeholder="Ej: 011324532">
        <select id="b-age" class="field field-select">
          <option value="">Todas las Agencias</option>
          <option value="VELOX">VELOX</option>
          <option value="PAQNORTE">PAQNORTE</option>
          <option value="CORREOMAX">CORREOMAX</option>
          <option value="RUTASUR">RUTASUR</option>
          <option value="ATLAS">ATLAS</option>
          <option value="PRONTO">PRONTO</option>
          <option value="BOLIDO">BOLIDO</option>
          <option value="FARO">FARO</option>
        </select>
      </div>
      <label class="bn-toggle" id="bn-toggle" title="Mientras esté marcada, busca nombres en la columna I de Entradas"><input type="checkbox" id="b-nombre" onchange="cambiarModoNombre_()"><span>👤 Buscar por nombre del cliente (columna I)</span></label>
      <button class="btn btn-block" onclick="iniciarBusqueda()">Buscar Coincidencias</button>

      <div id="b-res" class="search-results">
        <div class="empty-inline">Tus resultados aparecerán aquí...</div>
      </div>
    </div></div>
    ${DEFS_BEBE_RUFO_}${genBebeBuscador_(true)}
  `;
}
function genHtmlRetornos() {
    return `
    <button class="volver-rufo" onclick="irSeccion_('rufo_abierto', '🐾 Rufo')">🐾 Volver a Rufo</button>
    ${DEFS_BEBE_RUFO_}${DEFS_ESCRITORIO_RETORNOS_}
    ${genEncabezadoRastro_('↩️', 'Retornos', 'Artículos y clientes', 'retornos')}
    <button class="cr-b sec cr-abrir" onclick="irSeccion_('cambios', '🕵️ Cambios en Retornos')">🕵️ Ver cambios y borrados de Retornos</button>
    <div class="mp-zona ret">${cabezaMiniRufo_('buscar')}${genBebeBuscador_(false)}</div>
  `;
}
function cabezaMiniRufo_(_rx) {
    var _ry = _rx === 'retornos', _rz = _ry, _sa = _ry ? 'Retornos' : 'Buscar Referencia';
    var _sb = '#F4F5F7';
    var _sc = false
        ? '<defs><linearGradient id="mpRubio" x1="0" y1="0" x2="0" y2="1"><stop offset="0" stop-color="#fff1a8"/><stop offset=".55" stop-color="#f7cf4d"/><stop offset="1" stop-color="#d9a21e"/></linearGradient></defs>' +
            '<path d="M83 12 C99 3 114 16 110 34 C108 45 101 53 99 63 C95 52 95 42 93 33 C91 25 88 18 83 12 Z" fill="url(#mpRubio)"/>' +
            '<path d="M92 14 C102 14 106 24 103 38" stroke="#fff3c4" stroke-width="1.8" fill="none" stroke-linecap="round" opacity=".8"/>'
        : '';
    var _sd = false
        ? '<g transform="translate(50 45) scale(1.59) translate(-70 -36)"><path d="M47 33 C46 17 60 9 72 11 C86 12 96 22 93 34 C88 27 83 25 79 29 C77 23 71 22 68 27 C64 22 57 23 55 29 C52 26 49 28 47 33 Z" fill="url(#mpRubio)"/>' +
            '<path d="M55 17 C60 13 68 12 74 13" stroke="#fff3c4" stroke-width="2" fill="none" stroke-linecap="round" opacity=".8"/></g>' +
            '<ellipse cx="85" cy="13" rx="5.5" ry="4" fill="#E07A1F" transform="rotate(-30 85 13)"/>' +
            '<path d="M40.6 61.4 Q44.4 57.4 47.6 59.2 Q50 60.5 52.4 59.2 Q55.6 57.4 59.4 61.4 Q50 63.4 40.6 61.4 Z" fill="#c9184a"/>' +
            '<path d="M40.6 61.4 Q50 63.6 59.4 61.4 Q56 68.2 50 68.4 Q44 68.2 40.6 61.4 Z" fill="#e0335a"/>' +
            '<path d="M45.6 65.2 Q50 66.8 54.4 65.2" stroke="#ffa3b8" stroke-width="1.2" fill="none" stroke-linecap="round" opacity=".9"/>'
        : '';
    return '<button type="button" class="mp-cabeza' + (_rz ? ' rubia' : '') + '" id="mp-cabeza" data-destino="' + _rx + '" title="Ir a ' + _sa + ' (con lo que hayas escrito)" aria-label="Ir a ' + _sa + '" onclick="mpIr()">' +
        '<svg viewBox="8 -8 ' + (_rz ? 106 : 92) + ' 90" aria-hidden="true">' + _sc + (_rz ? '<use href="#rufo-svg"/>' : '<g class="ot-osita">' + OSITA_SVG_.replace(/ot-pelo/g, 'ot-pelo-m') + '</g>') + _sd +
        (!_rz ? '<g class="mp-susto"><ellipse cx="50" cy="62" rx="12" ry="7" fill="#f3eeea"/><ellipse cx="50" cy="63" rx="4.2" ry="5.4" fill="#3a1c14"/><path d="M85 24 Q89 31 85 34 Q81 31 85 24 Z" fill="#7dd3fc"/></g>' : '') +
        (!_rz ? '' : '<g class="mp-susto"><ellipse cx="50" cy="62" rx="12" ry="7" fill="#f3eeea"/>' +
            '<circle cx="34" cy="38" r="10" fill="#fff" stroke="#1B2636" stroke-width="1.4"/><circle cx="66" cy="38" r="10" fill="#fff" stroke="#1B2636" stroke-width="1.4"/>' +
            '<circle cx="34" cy="39" r="2.6" fill="#1B2636"/><circle cx="66" cy="39" r="2.6" fill="#1B2636"/>' +
            '<path d="M24 24 Q33 17 42 22" stroke="' + _sb + '" stroke-width="2.6" fill="none" stroke-linecap="round"/><path d="M58 22 Q67 17 76 24" stroke="' + _sb + '" stroke-width="2.6" fill="none" stroke-linecap="round"/>' +
            '<ellipse cx="50" cy="63" rx="4.6" ry="6" fill="#3a1c14"/>' +
            (false ? '<ellipse cx="50" cy="63" rx="5.7" ry="7.1" fill="none" stroke="#d6204a" stroke-width="2.4"/>' : '') +
            '<path d="M83 22 Q87 29 83 32 Q79 29 83 22 Z" fill="#7dd3fc"/></g>') +
        '<g class="mp-exc"><text x="92" y="6" font-family="Arial Black, Arial" font-weight="900" font-size="22" fill="#E07A1F">!</text><text x="80" y="0" font-family="Arial Black, Arial" font-weight="900" font-size="15" fill="#E07A1F" transform="rotate(-15 80 0)">!</text></g>' +
        '</svg></button>' +
        '<button type="button" class="mp-bocadillo" id="mp-bocadillo" tabindex="-1" onclick="mpIr()"><span class="mp-b-ir">Ir a ' + _sa + ' →</span><span class="mp-b-aviso">¿Lo busco en ' + _sa + '?</span></button>';
}
function genBebeBuscador_(_se) {
    var _sf = '<svg viewBox="0 0 130 112"><use href="#bebe-rufo"/></svg>';
    return (_se ? `<div class="rb-card" id="rb-card">${_sf}<div><b>¿Buscas un artículo o un cliente?</b><span>Rufi lo encuentra en Retornos aunque esté escrito un poco distinto.</span><br><button class="rb-abrir" onclick="rbAbrir()">👶 Abrir su buscador</button></div></div>` : '') +
        `<div class="rb-box" id="rb-box" style="display:${_se ? 'none' : 'block'}">
      <div class="rb-cab">${_sf}<div><b>Buscador de Rufi</b><span>Artículos y clientes de Retornos</span></div>${_se ? '<button class="rb-cerrar" onclick="rbCerrar()" title="Cerrar">✕</button>' : ''}</div>
      <input type="text" id="rb-texto" class="field rb-campo" placeholder="Artículo, cliente o nº de pedido…">
      <div class="rb-modos"><button class="rb-modo on" data-modo="todo" onclick="rbModo(this)">Todo</button><button class="rb-modo" data-modo="articulos" onclick="rbModo(this)">Artículos</button><button class="rb-modo" data-modo="clientes" onclick="rbModo(this)">Clientes</button><button class="rb-modo" data-modo="pedidos" onclick="rbModo(this)">Nº pedido</button></div>
      <label class="rb-hist"><input type="checkbox" id="rb-hist" onchange="if (document.getElementById('rb-texto').value.trim()) rbBuscar()"> 🕵️ Incluir borrados y cambiados</label>
      <button class="btn btn-block rb-btn" onclick="rbBuscar()">Buscar en Retornos</button>
      <div id="rb-res"><div class="rb-consejo">👶 Lo encuentra aunque haya tildes, mayúsculas, guiones o espacios distintos, o alguna letra cambiada. Primero salen las exactas y luego las parecidas, con su %. Con «Nº pedido» (o un número largo en «Todo») busca también en la columna C. Cada resultado dice si el retorno ya está procesado, y desde ahí se puede rellenar.</div></div>
    </div>`;
}
function genMiniCajaRufo_(_sg) {
    return _sg ? `<span class="rufo-chip-right">${_sg}</span>` : '';
}
function genEncabezadoRastro_(_sh, _si, _sj, _sk) {
    var _sl = _sk ? escenaCabecera_(_sk) : "";
    return `
    <div class="rufo-header encabezado-panel">
      ${_sl || `      <div class="rufo-avatar-mini-wrap">
        <svg class="rufo-avatar-mini" viewBox="0 0 100 100"><use href="#rufo-svg"/></svg>
        <svg class="rufo-gesto" viewBox="0 0 40 40">
          <rect class="rufo-mini-paquete" x="6" y="18" width="20" height="15" rx="3" fill="#E07A1F"/>
          <text class="rufo-mini-paquete" x="16" y="29" font-family="Georgia, 'Times New Roman', serif" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">P</text>
          <g class="rufo-sello">
            <circle cx="16" cy="10" r="7.5" fill="none" stroke="#fff" stroke-width="2"/>
            <path d="M12 10 L15 13 L20.5 6.5" stroke="#fff" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
          </g>
        </svg>
      </div>`}
      <div>
        <h2 class="rufo-nombre">${_sh} ${_si}</h2>
        <p class="rufo-sub">${_sj}</p>
      </div>
      <svg class="rufo-swoosh" viewBox="0 0 340 14" preserveAspectRatio="none">
        <path d="M0 8 C 60 -2, 120 16, 180 6 C 240 -4, 290 14, 340 5" stroke="#E07A1F" stroke-width="3" fill="none" stroke-linecap="round"/>
      </svg>
    </div>
  `;
}
function genHtmlRufo(_sm) {
    var _sn = _sm ? ' style="display:none;"' : '';
    var _so = _sm ? ' style="display:block;"' : ' style="display:none;"';
    var _sp = _sm ? 'rufo-colapsado' : 'rufo-colapsado dormido';
    var _sq = contarNotasNuevas_();
    var _sr = _sq > 0 ? `<span class="rufo-badge">${_sq}</span>` : '';
    var _ss = AVISOS_PEDIDOS_ACTIVOS_ ? contarAvisosNuevos_() : 0;
    var _st = _ss > 0 ? `<span class="rufo-badge">${_ss}</span>` : '';
    return `
    <div id="rufo-colapsado" class="${_sp}"${_sn}>
      <div class="rufo-avatar-click" onclick="despertarRufo()">
        <svg class="rufo-avatar-grande" viewBox="0 0 100 100" style="overflow:visible">
          <use href="#rufo-svg"/>
          
          <g class="rufo-parpados-dormido">
            <ellipse cx="34" cy="39" rx="9" ry="9" fill="url(#rufo-grad)"/>
            <ellipse cx="66" cy="39" rx="9" ry="9" fill="url(#rufo-grad)"/>
            <path d="M27 39 Q34 34 41 39" stroke="#D9DEE4" stroke-width="2.4" fill="none" stroke-linecap="round"/>
            <path d="M59 39 Q66 34 73 39" stroke="#D9DEE4" stroke-width="2.4" fill="none" stroke-linecap="round"/>
          </g>
          <text class="rufo-zzz rufo-zzz-1" x="83" y="13" font-family="Georgia, 'Times New Roman', serif" font-size="15" font-weight="bold" fill="#7c8aa0">z</text>
          <text class="rufo-zzz rufo-zzz-2" x="83" y="13" font-family="Georgia, 'Times New Roman', serif" font-size="12" font-weight="bold" fill="#7c8aa0">z</text>
          <text class="rufo-zzz rufo-zzz-3" x="83" y="13" font-family="Georgia, 'Times New Roman', serif" font-size="9" font-weight="bold" fill="#7c8aa0">z</text>
          
          <g class="rufo-boca-sorpresa">
            <ellipse cx="50" cy="63" rx="11" ry="7" fill="#fdfdfd"/>
            <ellipse cx="50" cy="64" rx="6.5" ry="7.5" fill="#3a2020"/>
            <ellipse cx="50" cy="68" rx="4" ry="2.6" fill="#E07A1F" opacity=".6"/>
          </g>
        </svg>
      </div>
      <button class="btn rufo-btn-hablar" onclick="despertarRufo()">Hablar con Rufo</button>
      <div id="pcn-hueco-tarjeta"></div>
    </div>

    <div id="rufo-expandido" class="rufo-panel"${_so}>
      <div class="rufo-header">
        <div class="rufo-avatar-mini-wrap">
          <svg class="rufo-avatar-mini" viewBox="0 0 100 100"><use href="#rufo-svg"/></svg>
          
          <svg class="rufo-gesto" viewBox="0 0 40 40">
            <rect class="rufo-mini-paquete" x="6" y="18" width="20" height="15" rx="3" fill="#E07A1F"/>
            <text class="rufo-mini-paquete" x="16" y="29" font-family="Georgia, 'Times New Roman', serif" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">P</text>
            <g class="rufo-sello">
              <circle cx="16" cy="10" r="7.5" fill="none" stroke="#fff" stroke-width="2"/>
              <path d="M12 10 L15 13 L20.5 6.5" stroke="#fff" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/>
            </g>
          </svg>
        </div>
        <div>
          <h2 class="rufo-nombre">Rufo</h2>
          <p class="rufo-sub">Tu ayudante de Devoluciones</p>
        </div>
        <svg class="rufo-swoosh" viewBox="0 0 340 14" preserveAspectRatio="none">
          <path d="M0 8 C 60 -2, 120 16, 180 6 C 240 -4, 290 14, 340 5" stroke="#E07A1F" stroke-width="3" fill="none" stroke-linecap="round"/>
        </svg>
      </div>
      <div class="rufo-body">
        <div class="rufo-burbuja">¡Hola! 👋 ¿En qué te ayudo hoy?</div>
        <div id="nt-franja" class="nt-franja" hidden></div>

        <div class="rufo-chips">
          <button class="rufo-chip" onclick="animarSeleccionRufo(this); irSeccion_('buscar', '🔍 Buscar Referencia')"><span>🔍</span> Buscar una referencia${genMiniCajaRufo_()}</button>
          <button class="rufo-chip" onclick="animarSeleccionRufo(this); irSeccion_('retornos', '↩️ Retornos')"><span>↩️</span> Retornos</button>
          <button class="rufo-chip" onclick="animarSeleccionRufo(this); google.script.run.mCalendario()"><span>📅</span> Ver calendario de agencias${genMiniCajaRufo_()}</button>
          <button class="rufo-chip" onclick="animarSeleccionRufo(this); irSeccion_('pendientes', '📋 Casillas Pendientes')"><span>📌</span> Casillas pendientes${genMiniCajaRufo_()}</button>
          <button class="rufo-chip" onclick="animarSeleccionRufo(this); google.script.run.mNotas()"><span>🗒️</span> Notas${genMiniCajaRufo_(_sr)}</button>
          ${AVISOS_PEDIDOS_ACTIVOS_ ? `<button class="rufo-chip" onclick="animarSeleccionRufo(this); google.script.run.mAvisos()"><span>🔔</span> Avisos de Pedidos${genMiniCajaRufo_(_st)}</button>` : ''}
          <button class="rufo-chip" onclick="animarSeleccionRufo(this); google.script.run.mRecuento()"><span>📊</span> Ver recuento y roturas${genMiniCajaRufo_()}</button>
          <div id="pcn-hueco-chip"></div>
        </div>
      </div>
    </div>
  `;
}
function genHtmlEscribir() {
    return `
    ${genEncabezadoRastro_('✍️', 'Redactar Nota', 'Para el turno siguiente', 'notas')}
    <div class="search-card">
      <textarea id="texto-nota" class="field field-textarea" placeholder="Escribe incidencias, paquetes pendientes..."></textarea>
      <button id="btn-guardar-nota" class="btn btn-block" style="margin-top:15px;" onclick="guardarNota()">Guardar Nota</button>
    </div>
  `;
}
function genHtmlVerNotas() {
    var _su = genEncabezadoRastro_('📌', 'Notas Activas', 'Últimas 10 anotaciones', 'notas');
    try {
        var _sv = SpreadsheetApp.getActiveSpreadsheet(), _sw = _sv.getSheetByName(HOJA_NOTAS);
        if (!_sw || _sw.getLastRow() < 2)
            return _su + "<div class='empty-state'><span class='emoji'>📭</span>No hay notas activas.</div>";
        var _sx = _sw.getDataRange().getValues(), _sy = "";
        for (var _sz = _sx.length - 1; _sz >= Math.max(1, _sx.length - 10); _sz--) {
            var _ta = _sx[_sz][0];
            var _tb = _ta instanceof Date ? _ta : new Date(_ta);
            var _tc = _tb instanceof Date && !isNaN(_tb.getTime());
            var _td = _tc
                ? (('0' + _tb.getDate()).slice(-2) + '/' + ('0' + (_tb.getMonth() + 1)).slice(-2) + ' ' + ('0' + _tb.getHours()).slice(-2) + ':' + ('0' + _tb.getMinutes()).slice(-2))
                : '--/--';
            var _te = String(_sx[_sz][1] === undefined || _sx[_sz][1] === null ? '' : _sx[_sz][1]);
            var _tf = String(_sx[_sz][2] === undefined || _sx[_sz][2] === null ? '' : _sx[_sz][2]);
            var _tg = (_sx[_sz][3] === "NUEVA");
            var _th = _tg ? '<span class="badge-nueva">🔴 NUEVA</span>' : '';
            _sy += `<div class="note-card" style="border-left-color:${_tg ? 'var(--red)' : 'var(--border)'};"><div class="note-header">📅 ${_td} &nbsp;·&nbsp; 👤 ${_te} ${_th}</div><div class="note-body">${_tf.replace(/\n/g, "<br>")}</div></div>`;
        }
        if (_sw.getLastColumn() >= 4) {
            var _ti = _sw.getRange(2, 4, _sw.getLastRow() - 1, 1), _tj = _ti.getValues(), _tk = false;
            for (var _tl = 0; _tl < _tj.length; _tl++) {
                if (_tj[_tl][0] === "NUEVA") {
                    _tj[_tl][0] = "LEÍDA";
                    _tk = true;
                }
            }
            if (_tk)
                _ti.setValues(_tj);
        }
        return _su + _sy;
    }
    catch (_tm) {
        Logger.log('genHtmlVerNotas() ha fallado: ' + (_tm && _tm.message ? _tm.message : _tm));
        return _su + "<div class='empty-state'><span class='emoji'>⚠️</span>No se han podido cargar las notas (dato inesperado en " + HOJA_NOTAS + "). Se ha registrado el error.</div>";
    }
}
function genHtmlGestionarNotas() {
    var _tn = SpreadsheetApp.getActiveSpreadsheet(), _to = _tn.getSheetByName(HOJA_NOTAS);
    var _tp = genEncabezadoRastro_('⚙️', 'Gestionar Notas', 'Edita o elimina', 'notas');
    if (!_to || _to.getLastRow() < 2)
        return _tp + "<div class='empty-state'><span class='emoji'>🗒️</span>No hay notas para gestionar.</div>";
    var _tq = _to.getDataRange().getValues(), _tr = `<table class="tbl"><tr><th>Fecha</th><th>Nota</th><th>Acción</th></tr>`;
    for (var _ts = _tq.length - 1; _ts >= 1; _ts--) {
        var _tt = new Date(_tq[_ts][0]), _tu = ('0' + _tt.getDate()).slice(-2) + '/' + ('0' + (_tt.getMonth() + 1)).slice(-2);
        var _tv = _ts + 1;
        _tr += `<tr id="fila-${_tv}"><td>${_tu}</td><td style="font-size:13px; max-width:300px;">
      <span id="texto-g-${_tv}">${_tq[_ts][2]}</span>
      <textarea id="edit-g-${_tv}" class="field field-textarea" style="display:none; margin-top:6px; min-height:70px;">${_tq[_ts][2]}</textarea>
    </td><td style="text-align:center; white-space:nowrap;">
      <span id="acciones-g-${_tv}">
        <button class="btn-cal" onclick="iniciarEdicionGestionar(${_tv})" title="Editar">✏️</button>
        <button class="btn btn-danger-sm" onclick="borrarNota(${_tv})">🗑️</button>
      </span>
      <span id="acciones-edit-g-${_tv}" style="display:none;">
        <button class="btn-cal" onclick="guardarEdicionGestionar(${_tv})" title="Guardar">💾</button>
        <button class="btn-cal" onclick="cancelarEdicionGestionar(${_tv})" title="Cancelar">✖️</button>
      </span>
    </td></tr>`;
    }
    return _tp + _tr + "</table>";
}
function genHtmlNotas(_tw) {
    if (_tw === undefined)
        _tw = true;
    var _tx = genEncabezadoRastro_('🗒️', 'Notas', 'Turno y pedidos · pendientes y resueltas', 'notas');
    var _ty = "", _tz = "", _ua = 0, _ub = 0, _uc = "";
    try {
        var _ud = leerNotas_(), _ue = _ud.hoja;
        for (var _uf = _ud.filas.length - 1; _uf >= 0; _uf--) {
            var _ug = _ud.filas[_uf], _uh = notaTarjetaHtml_(_ug);
            if (_ug.resuelta) {
                _tz += _uh;
                _ub++;
            }
            else {
                _ty += _uh;
                _ua++;
            }
        }
        if (_tw && _ue && _ue.getLastRow() >= 2 && _ue.getLastColumn() >= 4) {
            var _ui = _ue.getRange(2, 4, _ue.getLastRow() - 1, 1), _uj = _ui.getValues(), _uk = false;
            for (var _ul = 0; _ul < _uj.length; _ul++) {
                if (_uj[_ul][0] === "NUEVA") {
                    _uj[_ul][0] = "LEÍDA";
                    _uk = true;
                }
            }
            if (_uk) {
                _ui.setValues(_uj);
                construirMenu();
            }
        }
    }
    catch (_um) {
        Logger.log('genHtmlNotas() ha fallado: ' + (_um && _um.message ? _um.message : _um));
        _uc = "<div class='empty-state'><span class='emoji'>⚠️</span>No se han podido cargar las notas (dato inesperado en " + HOJA_NOTAS + "). Se ha registrado el error.</div>";
    }
    if (!_ty)
        _ty = "<div class='empty-state nt-vacio'><span class='emoji'>🎉</span>No hay nada pendiente.</div>";
    if (!_tz)
        _tz = "<div class='empty-state nt-vacio'><span class='emoji'>📭</span>Todavía no hay notas resueltas.</div>";
    return _tx +
        '<div class="search-card nt-escribir">' +
        '<div class="nt-tipos">' +
        '<button type="button" class="nt-tipo on" data-tipo="turno" onclick="ntElegirTipo(this)">📋 Para el turno</button>' +
        '<button type="button" class="nt-tipo" data-tipo="pedido" onclick="ntElegirTipo(this)">📦 De un pedido</button>' +
        '</div>' +
        '<input id="nt-pedido" class="field" placeholder="Nº de pedido" style="display:none; margin-bottom:8px;">' +
        '<textarea id="texto-nota-unif" class="field field-textarea nt-texto" placeholder="Escribe la nota…"></textarea>' +
        '<button id="btn-guardar-nota-unif" class="btn btn-block" style="margin-top:10px;" onclick="guardarNotaUnificada()">Guardar nota</button>' +
        '</div>' +
        _uc +
        '<div class="nt-pestanas">' +
        '<button type="button" class="nt-pes on" data-p="pend" onclick="ntPestana(this)">Pendientes' + (_ua ? ' <i class="nt-pill">' + _ua + '</i>' : '') + '</button>' +
        '<button type="button" class="nt-pes" data-p="hechas" onclick="ntPestana(this)">Resueltas' + (_ub ? ' <i class="nt-pill gris">' + _ub + '</i>' : '') + '</button>' +
        '</div>' +
        '<div class="notas-lista nt-lista" id="nt-lista-pend">' + _ty + '</div>' +
        '<div class="notas-lista nt-lista" id="nt-lista-hechas" style="display:none;">' + _tz + '</div>';
}
function notaTarjetaHtml_(_un) {
    var _uo = !!_un.pedido, _up = !!_un.resuelta, _uq = _un.fila;
    var _ur = 'note-card nt-nota ' + (_up ? 'hecha' : (_uo ? 'ped' : 'turno'));
    var _us = _uo ? '<span class="nt-chip ped">📦 ' + notaEsc_(_un.pedido) + '</span>' : '<span class="nt-chip tur">📋 Turno</span>';
    if (_un.estado === "NUEVA" && !_up)
        _us += '<span class="nt-chip nueva">NUEVA</span>';
    if (_up)
        _us += '<span class="nt-chip ok">✅ Resuelta</span>';
    var _ut;
    if (_up) {
        var _uu = _un.resuelta.split('|');
        _ut = 'Resuelta por <b>' + notaEsc_(notaCorto_(_uu[1])) + '</b> el ' + notaEsc_(String(_uu[0]).replace(/\/\d{4} /, ' · '));
    }
    else if (_uo)
        _ut = 'Sale al buscar el pedido';
    else
        _ut = _un.enteradas.length ? 'Enterada: <b>✓ ' + _un.enteradas.map(function (_uv) { return notaEsc_(notaCorto_(_uv)); }).join(', ') + '</b>' : 'Nadie la ha marcado como enterada';
    var _uw = _up
        ? '<button class="nt-acc" title="Volver a pendiente" onclick="resolverNotaUnificada(' + _uq + ', ' + _un.id + ', false)">↩️</button>'
        : '<button class="nt-acc res" onclick="resolverNotaUnificada(' + _uq + ', ' + _un.id + ', true)">✅ Resolver</button>' +
            '<button class="nt-acc" title="Editar" onclick="iniciarEdicionNota(' + _uq + ')">✏️</button>' +
            '<button class="nt-acc" title="Borrar" onclick="borrarNotaUnificada(' + _uq + ')">🗑️</button>';
    return '<div class="' + _ur + '" id="fila-' + _uq + '">' +
        '<div class="nt-arriba">' + _us + '<span>' + notaFechaTxt_(_un.fecha) + ' · ' + notaEsc_(notaCorto_(_un.usuario)) + '</span></div>' +
        '<div class="note-body nt-cuerpo" id="texto-' + _uq + '">' + notaEsc_(_un.texto).replace(/\n/g, "<br>") + '</div>' +
        '<textarea class="field field-textarea" id="edit-' + _uq + '" style="display:none; margin-top:8px; min-height:80px;">' + notaEsc_(_un.texto) + '</textarea>' +
        '<div class="nt-pie" id="acciones-' + _uq + '"><span class="nt-quien">' + _ut + '</span><span class="nt-accs">' + _uw + '</span></div>' +
        '<div class="note-actions" id="acciones-edit-' + _uq + '" style="display:none;">' +
        '<button class="btn btn-block" style="width:auto; padding:9px 16px;" onclick="guardarEdicionNota(' + _uq + ')">💾 Guardar cambios</button>' +
        '<button class="btn-cal" onclick="cancelarEdicionNota(' + _uq + ')">✖️ Cancelar</button>' +
        '</div>' +
        '</div>';
}
function genHtmlAvisos(_ux) {
    if (_ux === undefined)
        _ux = true;
    var _uy = genEncabezadoRastro_('🔔', 'Avisos de Pedidos', 'Referencias a vigilar cuando lleguen', 'avisos');
    var _uz = "<div class='empty-state'><span class='emoji'>📭</span>Todavía no hay avisos guardados.</div>";
    try {
        var _va = SpreadsheetApp.getActiveSpreadsheet(), _vb = _va.getSheetByName(HOJA_AVISOS);
        if (_vb && _vb.getLastRow() >= 2) {
            var _vc = [[]].concat(_vb.getRange(2, 1, _vb.getLastRow() - 1, 5).getValues());
            var _vd = _va.getSheetByName(HOJA_AVISOS_EVENTOS);
            var _ve = {};
            var _vf = _vd ? _vd.getLastRow() : 0;
            if (_vd && _vf >= 2) {
                var _vg = [[]].concat(_vd.getRange(2, 1, _vf - 1, 6).getValues());
                for (var _vh = _vg.length - 1; _vh >= 1; _vh--) {
                    var _vi = (_vg[_vh][2] || "").toString();
                    if (!_ve[_vi])
                        _ve[_vi] = [];
                    if (_ve[_vi].length < 3)
                        _ve[_vi].push(_vg[_vh]);
                }
            }
            _uz = "";
            for (var _vj = _vc.length - 1; _vj >= 1; _vj--) {
                var _vk = _vc[_vj][0];
                var _vl = _vk instanceof Date ? _vk : new Date(_vk);
                var _vm = _vl instanceof Date && !isNaN(_vl.getTime());
                var _vn = _vm ? (('0' + _vl.getDate()).slice(-2) + '/' + ('0' + (_vl.getMonth() + 1)).slice(-2)) : '--/--';
                var _vo = String(_vc[_vj][1] === undefined || _vc[_vj][1] === null ? '' : _vc[_vj][1]);
                var _vp = String(_vc[_vj][2] === undefined || _vc[_vj][2] === null ? '' : _vc[_vj][2]);
                var _vq = String(_vc[_vj][3] === undefined || _vc[_vj][3] === null ? '' : _vc[_vj][3]);
                var _vr = (typeof _vc[_vj][4] === 'number') ? _vc[_vj][4] : 0;
                var _vs = _vj + 1;
                var _vt = normalizarRef(_vp);
                var _vu = _vr > 0 ? `<span class="aviso-contador">🔔 ${_vr} ${_vr === 1 ? 'vez' : 'veces'}</span>` : '';
                var _vv = _vr > 0 ? 'var(--amber)' : 'var(--border)';
                var _vw = '';
                var _vx = _ve[_vt] || [];
                if (_vx.length > 0) {
                    _vw = '<div class="aviso-hist">' + _vx.map(function (_vy) {
                        var _vz = _vy[1] instanceof Date ? _vy[1] : new Date(_vy[1]);
                        var _wa = (_vz instanceof Date && !isNaN(_vz.getTime()))
                            ? (('0' + _vz.getDate()).slice(-2) + '/' + ('0' + (_vz.getMonth() + 1)).slice(-2) + ' ' + ('0' + _vz.getHours()).slice(-2) + ':' + ('0' + _vz.getMinutes()).slice(-2))
                            : '--/--';
                        return `<div class="aviso-hist-item">🕓 <b>${_wa}</b> — ${_vy[4]}, fila ${_vy[5]}</div>`;
                    }).join('') + '</div>';
                }
                _uz += `
          <div class="aviso-card" style="border-left-color:${_vv};">
            <div class="aviso-card-top">
              <span class="aviso-ref">📦 ${_vp}</span>
              ${_vu}
              <span class="aviso-meta">Creado ${_vn} · ${_vo}</span>
            </div>
            <div class="aviso-motivo" id="motivo-${_vs}">${_vq.replace(/\n/g, "<br>")}</div>
            <textarea class="field field-textarea" id="edit-aviso-${_vs}" style="display:none; margin-top:8px; min-height:70px;">${_vq}</textarea>
            ${_vw}
            <div class="aviso-acciones" id="acciones-aviso-${_vs}">
              <button class="btn-cal" onclick="iniciarEdicionAviso(${_vs})">✏️ Editar motivo</button>
              <button class="btn btn-danger-sm" onclick="borrarAvisoUnificado(${_vs})">🗑️ Eliminar</button>
            </div>
            <div class="aviso-acciones" id="acciones-edit-aviso-${_vs}" style="display:none;">
              <button class="btn btn-block" style="width:auto; padding:9px 16px;" onclick="guardarEdicionAviso(${_vs})">💾 Guardar cambios</button>
              <button class="btn-cal" onclick="cancelarEdicionAviso(${_vs})">✖️ Cancelar</button>
            </div>
          </div>`;
            }
            if (_ux && _vd && _vf >= 2 && _vd.getLastColumn() >= 8) {
                var _wb = _vd.getRange(2, 8, _vf - 1, 1).getValues(), _wc = -1, _wd = -1;
                for (var _we = 0; _we < _wb.length; _we++) {
                    if (_wb[_we][0] === "NUEVA") {
                        _wb[_we][0] = "LEÍDA";
                        if (_wc === -1)
                            _wc = _we;
                        _wd = _we;
                    }
                }
                if (_wc > -1) {
                    _vd.getRange(2 + _wc, 8, _wd - _wc + 1, 1).setValues(_wb.slice(_wc, _wd + 1));
                    construirMenu();
                }
            }
        }
    }
    catch (_wf) {
        Logger.log('genHtmlAvisos() ha fallado: ' + (_wf && _wf.message ? _wf.message : _wf));
        _uz = "<div class='empty-state'><span class='emoji'>⚠️</span>No se han podido cargar los avisos (dato inesperado en " + HOJA_AVISOS + "). Se ha registrado el error.</div>";
    }
    return `
    ${_uy}
    <div class="search-card">
      <h3>➕ Nuevo aviso</h3>
      <input id="ref-aviso-nuevo" class="field" style="width:100%; margin-bottom:10px;" placeholder="Referencia del pedido (ej. 4521178)">
      <textarea id="motivo-aviso-nuevo" class="field field-textarea" style="height:90px;" placeholder="Motivo: por qué hay que estar pendiente cuando llegue este pedido..."></textarea>
      <button id="btn-guardar-aviso" class="btn btn-block" style="margin-top:12px;" onclick="guardarAvisoNuevo()">Añadir aviso</button>
    </div>
    <div class="avisos-lista">
      ${_uz}
    </div>
  `;
}
function obtenerHTMLBase() {
    return `<!DOCTYPE html><html><head><base target="_top"><style>
    #aviso-panel { position:fixed; left:8px; right:8px; top:8px; z-index:99999; background:#1e293b; color:#fff; border-radius:12px; padding:10px 34px 10px 12px; font-size:13px; font-weight:600; line-height:1.35; box-shadow:0 10px 24px -8px rgba(0,0,0,.45); transform:translateY(-160%); transition:transform .25s; cursor:pointer; word-break:break-word; }
    #aviso-panel.ver { transform:none; }
    #marca-cache { position:fixed; left:8px; bottom:6px; z-index:9998; background:rgba(30,41,59,.82); color:#fff; font-size:11px; border-radius:999px; padding:3px 9px; opacity:0; transition:opacity .25s; pointer-events:none; } #marca-cache.ver { opacity:1; } 
    #aviso-panel::after { content:'✕'; position:absolute; right:12px; top:9px; opacity:.7; }
    :root{
      --primary:#2563eb; --primary-2:#3b82f6; --primary-dark:#1d4ed8;
      --purple:#8b5cf6; --purple-dark:#7c3aed;
      --green:#10b981; --green-dark:#059669;
      --red:#ef4444; --red-dark:#dc2626;
      --amber:#f59e0b;
      --bg:#f1f5f9; --card:#ffffff; --text:#1e293b; --muted:#64748b; --border:#e2e8f0;
      --radius:16px; --radius-sm:10px;
      --shadow:0 1px 3px rgba(15,23,42,.06), 0 10px 26px -6px rgba(15,23,42,.12);
      --shadow-sm:0 1px 2px rgba(15,23,42,.05), 0 2px 8px rgba(15,23,42,.05);
      --ease:cubic-bezier(.4,0,.2,1);
    }
    *{ box-sizing:border-box; }
    body {
      font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;
      background:var(--bg); margin:0; padding:20px; color:var(--text);
      -webkit-font-smoothing:antialiased;
    }
    h1,h2,h3{ font-family:inherit; }
    ::-webkit-scrollbar{ width:8px; height:8px; }
    ::-webkit-scrollbar-track{ background:transparent; }
    ::-webkit-scrollbar-thumb{ background:#cbd5e1; border-radius:8px; }
    ::-webkit-scrollbar-thumb:hover{ background:#94a3b8; }

    @keyframes fadeInUp{ from{opacity:0; transform:translateY(8px);} to{opacity:1; transform:translateY(0);} }
    @keyframes shakeErr{ 10%,90%{transform:translateX(-1px);} 20%,80%{transform:translateX(2px);} 30%,50%,70%{transform:translateX(-4px);} 40%,60%{transform:translateX(4px);} }
    @keyframes spin{ to{ transform:rotate(360deg); } }
    @keyframes pulseIcon{ 0%,100%{ transform:scale(1); } 50%{ transform:scale(1.06); } }

    .lock-box { background: var(--card); max-width: 340px; margin: 36px auto; padding: 32px 28px; border-radius: 20px; box-shadow: var(--shadow); text-align: center; animation: fadeInUp .35s var(--ease); }
    .lock-icon-oso { position:relative; width:74px; height:74px; margin:0 auto 16px; animation: pulseIcon 2.6s ease-in-out infinite; }
    .lock-icon-oso svg { width:100%; height:100%; filter: drop-shadow(0 6px 14px rgba(0,0,0,.18)); }
    .lock-badge { position:absolute; right:-4px; bottom:-2px; width:26px; height:26px; background:#E07A1F; border-radius:50%; display:flex; align-items:center; justify-content:center; font-size:13px; box-shadow:0 2px 6px rgba(0,0,0,.25); border:2px solid #fff; }
    .lock-box h2 { margin: 0 0 4px; font-size: 19px; font-weight: 700; color:var(--text); }
    .lock-box .sub { color: var(--muted); font-size: 13px; margin: 0 0 22px; }
    .pwd-wrap { position: relative; width: 100%; margin-bottom: 14px; }
    .pwd-wrap input { width: 100%; padding: 13px 42px 13px 16px; font-size: 15px; border: 1.5px solid var(--border); border-radius: var(--radius-sm); box-sizing: border-box; transition: border-color .2s var(--ease), box-shadow .2s var(--ease), background .2s; background:#f8fafc; }
    .pwd-wrap input:focus { border-color: var(--primary-2); outline: none; box-shadow: 0 0 0 4px rgba(59,130,246,.14); background:#fff; }
    .pwd-wrap .eye { position: absolute; right: 13px; top: 50%; transform: translateY(-50%); cursor: pointer; font-size: 19px; user-select: none; opacity: 0.55; transition: opacity .2s; }
    .pwd-wrap .eye:hover { opacity: 1; }
    #err { color: var(--red-dark); font-size: 12.5px; font-weight: 700; margin-bottom: 14px; display: none; }
    #err.show { display: block; animation: shakeErr .4s var(--ease); }

    .btn { background: linear-gradient(135deg, var(--primary-2), var(--primary-dark)); color: white; border: none; padding: 13px 20px; font-size: 14.5px; font-weight: 700; border-radius: var(--radius-sm); cursor: pointer; transition: transform .15s var(--ease), box-shadow .15s var(--ease), opacity .15s; box-shadow: 0 4px 12px -2px rgba(37,99,235,.4); }
    .btn:hover { transform: translateY(-1px); box-shadow: 0 6px 18px -3px rgba(37,99,235,.5); }
    .btn:active { transform: translateY(0); }
    .btn:disabled { opacity:.6; cursor:default; transform:none; box-shadow:none; }
    .btn-block { width: 100%; }
    .btn-danger-sm { background: linear-gradient(135deg, #f87171, var(--red-dark)); padding: 8px 14px; font-size: 12.5px; box-shadow: 0 3px 10px -2px rgba(220,38,38,.4); }

    .field { padding: 13px 15px; border: 1.5px solid var(--border); border-radius: var(--radius-sm); font-size: 14.5px; background: #f8fafc; transition: border-color .2s var(--ease), box-shadow .2s var(--ease), background .2s; font-family: inherit; }
    .field:focus { border-color: var(--primary-2); outline: none; box-shadow: 0 0 0 4px rgba(59,130,246,.14); background: #fff; }
    .field-ref { flex: 2; }
    .field-select { flex: 1; }
    .field-textarea { width: 100%; height: 210px; resize: none; box-sizing: border-box; line-height: 1.5; }

    .panel-title { font-size: 15px; font-weight: 800; color: #334155; margin-bottom: 14px; display:flex; align-items:center; gap:8px; }

    .grid { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 14px; margin-bottom: 20px; }
    .card { background: var(--card); border-radius: var(--radius); padding: 18px 10px; text-align: center; border-top: 4px solid var(--green); box-shadow: var(--shadow-sm); border-left:1px solid var(--border); border-right:1px solid var(--border); border-bottom:1px solid var(--border); transition: transform .2s var(--ease), box-shadow .2s var(--ease); animation: fadeInUp .35s var(--ease); }
    .card:hover { transform: translateY(-2px); box-shadow: var(--shadow); }
    .card .icon { font-size: 19px; margin-bottom: 2px; }
    .card.fact { border-top-color: var(--purple); }
    .card.rotura { border-top-color: var(--red); }
    .title { font-size: 12px; font-weight: 700; color: var(--muted); text-transform: uppercase; letter-spacing:.03em; margin-bottom: 6px; }
    .value { font-size: 36px; font-weight: 900; line-height: 1; color: var(--green-dark); }
    .card.fact .value { color: var(--purple-dark); }
    .card.rotura .value { color: var(--red-dark); }
    .details-container { display: grid; grid-template-columns: 1fr 1.2fr; gap: 15px; }
    .details { background: var(--card); border-radius: var(--radius); padding: 18px; box-shadow: var(--shadow-sm); border:1px solid var(--border); }
    .dh3 { margin-top:0; font-size:14px; font-weight:700; color:#334155; border-bottom: 1px solid var(--border); padding-bottom: 9px; margin-bottom:10px; }
    .turn { display: flex; justify-content: space-between; padding: 8px 0; border-bottom: 1px dashed var(--border); font-size: 13.5px; }
    .turn:last-child { border-bottom:none; }
    .turn span:first-child { color: var(--muted); }
    .bar-row { margin-bottom: 12px; }
    .bar-info { display: flex; justify-content: space-between; align-items: center; font-size: 12.5px; margin-bottom: 5px; }
    .bar-pct { font-weight: 700; color: var(--red-dark); }
    .bar-track { background: #f1f5f9; border-radius: 6px; height: 9px; overflow: hidden; }
    .bar-fill { background: linear-gradient(90deg, #f87171, var(--red-dark)); height: 100%; border-radius: 6px; transition: width .5s var(--ease); }

    .search-card { background: var(--card); border-radius: var(--radius); padding: 16px; box-shadow: var(--shadow-sm); border:1px solid var(--border); }
    .search-card h3 { margin-top: 0; margin-bottom: 4px; color: #2c3e50; display:flex; align-items:center; gap:8px; font-size:16px; }
    .search-icon { font-size:17px; }
    .search-hint { margin-top: -2px; margin-bottom: 16px; font-size: 12px; color: var(--muted); line-height:1.5; }
    .search-row { display: flex; gap: 10px; margin-bottom: 15px; flex-wrap: wrap; }
    .search-results { max-height: 350px; overflow-y: auto; border-top: 1px dashed var(--border); padding-top: 15px; margin-top:4px; }
    .bi-estado { font-size:11px; color:#64748b; margin:-6px 0 8px; } .bi-estado.aviso { color:#b45309; } 
    .bi-sin { font-size:12.5px; color:var(--red-dark); background:#fef2f2; border:1px solid #fecaca; border-radius:10px; padding:8px 10px; margin:0 0 10px; text-align:center; } 
    .fila-movida { font-size:11px; color:#0369a1; background:#f0f9ff; border:1px solid #bae6fd; border-radius:8px; padding:4px 8px; margin-top:6px; } 
    .empty-inline { color: var(--muted); text-align: center; padding: 10px; font-size: 13.5px; }
    .loading-row { text-align:center; padding: 34px 10px; color: var(--muted); font-size: 14px; }
    .spinner { display:inline-block; width:15px; height:15px; border:2px solid rgba(37,99,235,.22); border-top-color:var(--primary-2); border-radius:50%; animation: spin .7s linear infinite; vertical-align:middle; margin-right:8px; position:relative; top:-1px; }

    .res-item { background: #f8fafc; border: 1px solid var(--border); border-radius: var(--radius-sm); padding: 13px 14px; margin-bottom: 10px; display: flex; flex-direction: column; align-items: stretch; gap:10px; max-width: 100%; overflow: hidden; transition: .2s var(--ease); animation: fadeInUp .25s var(--ease); }
    .res-item:hover { border-color: var(--primary-2); box-shadow: var(--shadow-sm); }
    .res-item.visitado { background: #eef2f6; opacity: 0.65; border-color: var(--border); }
    .res-info { font-size: 13.5px; min-width: 0; overflow-wrap: anywhere; }
    .res-hoja { font-size: 10.5px; font-weight: bold; color: var(--primary-dark); text-transform: uppercase; letter-spacing:.02em; display:flex; flex-wrap: wrap; align-items:center; gap:6px 8px; margin-bottom:5px; }
    .res-ref { font-size: 12px; color: var(--muted); margin-top: 4px; overflow-wrap: anywhere; }
    .badge-exacta { background:#dcfce7; color:#15803d; padding:2px 8px; border-radius:20px; font-weight:800; font-size:10px; white-space:nowrap; }
    .badge-similar { background:#fef3c7; color:#b45309; padding:2px 8px; border-radius:20px; font-weight:800; font-size:10px; white-space:nowrap; }
    .badge-pct { padding:2px 8px; border-radius:20px; font-weight:800; font-size:10px; white-space:nowrap; }
    .badge-pct-alta { background:#dbeafe; color:#1d4ed8; }
    .badge-pct-media { background:#fef3c7; color:#b45309; }
    .badge-pct-baja { background:#fee2e2; color:#b91c1c; }
    .badge-dia { background:#e2e8f0; color:#334155; padding:2px 8px; border-radius:20px; font-weight:800; font-size:10px; white-space:nowrap; }
    .res-notaL { margin-top: 4px; font-size: 12px; }
    .res-notaL.con-dato { color: var(--purple-dark); font-weight:600; }
    .res-notaL.sin-dato { color: #bdc3c7; }
    .btn-ir { background: var(--green); color: white; border: none; padding: 10px 14px; border-radius: 8px; cursor: pointer; font-weight: 700; font-size: 12.5px; width: 100%; transition:.15s; }
    .btn-ir:hover { background: var(--green-dark); }
    .visitado .btn-ir { background: #94a3b8; }

    .res-acciones-row { display: flex; gap: 8px; }
    .btn-rapido-toggle { background: #fff; color: var(--primary-dark); border: 1px dashed var(--primary-2); padding: 9px 12px; border-radius: 8px; cursor: pointer; font-weight: 700; font-size: 12px; flex-shrink: 0; white-space: nowrap; transition:.15s; }
    .btn-rapido-toggle:hover { background: #eff6ff; }
    .btn-rapido-toggle.abierto { background: var(--primary-2); color: #fff; border-style: solid; }
    .res-rapido { display: none; margin-top: 2px; padding-top: 10px; border-top: 1px dashed var(--border); animation: fadeInUp .2s var(--ease); }
    .res-rapido.abierto { display: block; }
    .res-rapido-grupo { margin-bottom: 8px; }
    .res-rapido-grupo:last-child { margin-bottom: 0; }
    .res-rapido-label { font-size: 10.5px; font-weight: 800; color: var(--muted); text-transform: uppercase; letter-spacing: .02em; margin-bottom: 5px; }
    .res-rapido-btns { display: flex; gap: 6px; flex-wrap: wrap; }
    .qbtn { border: 1px solid var(--border); background: #fff; padding: 7px 11px; border-radius: 20px; cursor: pointer; font-weight: 700; font-size: 11.5px; transition:.15s; color: #334155; }
    .qbtn:hover { border-color: var(--primary-2); }
    .qbtn.qbtn-si.activo { background: var(--red); border-color: var(--red-dark); color: #fff; }
    .qbtn.qbtn-no.activo { background: var(--green); border-color: var(--green-dark); color: #fff; }
    .qbtn.qbtn-ok.activo { background: var(--green); border-color: var(--green-dark); color: #fff; }
    .qbtn.qbtn-reclamar.activo { background: var(--amber); border-color: #92400e; color: #fff; }
    .qbtn.qbtn-reclamar-pend.activo { background: var(--purple); border-color: var(--purple-dark); color: #fff; }
    .qbtn:disabled { opacity: .6; cursor: default; }

    .caja-compositor { display:none; margin-top:12px; padding:16px 18px; background:#eef2ff; border:1px solid #c7d2fe; border-radius:var(--radius-sm); animation: fadeInUp .2s var(--ease); }
    .caja-compositor.abierta { display:block; }
    .caja-compositor-titulo { font-size:13px; font-weight:800; color:#4338ca; text-transform:uppercase; letter-spacing:.02em; margin-bottom:12px; }
    .caja-compositor-fila { display:flex; gap:20px; margin-bottom:18px; }
    .caja-compositor-campo { flex:1; }
    .caja-compositor-campo label { display:block; font-size:11px; font-weight:700; color:var(--muted); text-transform:uppercase; margin-bottom:5px; }
    .caja-lado { flex:1; min-width:0; }
    .caja-lado-titulo { font-size:13.5px; font-weight:800; color:#4338ca; text-transform:uppercase; letter-spacing:.02em; margin-bottom:9px; }
    .caja-ro { background:#f8fafc; border:1.5px solid var(--border); border-radius:9px; padding:10px 12px; font-size:14.5px; color:#334155; font-weight:600; }
    .caja-modo-select { width:100%; border:1.5px solid #c7d2fe; border-radius:10px; padding:12px 13px; font-size:14.5px; background:#fff; margin-bottom:12px; font-family:inherit; color:#3730a3; font-weight:600; }
    .caja-modo-select:focus { outline:none; border-color:#6366f1; }
    .caja-codigos-wrap { display:none; }
    .caja-codigos-wrap.visible { display:block; }
    .caja-codigos-wrap .caja-sublabel { font-size:12px; color:#6b5fb5; margin-bottom:7px; line-height:1.45; }
    .caja-codigos { width:100%; border:1.5px solid #c7d2fe; border-radius:10px; padding:13px 14px; font-size:15px; background:#fff; box-sizing:border-box; min-height:84px; resize:vertical; font-family:inherit; }
    .caja-codigos:focus { outline:none; border-color:#6366f1; box-shadow:0 0 0 3px rgba(99,102,241,.15); }
    .caja-codigos-mal { border-color:var(--red); background:#fef2f2; }
    .caja-codigos-mal:focus { border-color:var(--red-dark); box-shadow:0 0 0 3px rgba(239,68,68,.15); }
    .caja-codigos-bien { border-color:var(--green); background:#f0fdf4; }
    .caja-codigos-bien:focus { border-color:var(--green-dark); box-shadow:0 0 0 3px rgba(16,185,129,.15); }
    .caja-contador { display:inline-block; font-size:12.5px; color:#7c3aed; font-weight:700; margin-top:7px; }

    .caja-mayoria-aviso { display:none; align-items:center; gap:12px; border:2px solid var(--green); background:#f0fdf4; border-radius:10px; padding:13px 15px; margin-top:10px; }
    .caja-mayoria-aviso.visible { display:flex; }
    .caja-mayoria-aviso-icono { font-size:24px; line-height:1; flex-shrink:0; }
    .caja-mayoria-aviso-texto { font-size:13.5px; font-weight:800; color:var(--green-dark); line-height:1.4; word-break:break-word; }

    @media (max-width: 560px) {
      .caja-compositor-fila { flex-direction: column; gap: 18px; }
    }

    .caja-codigos-wrap.caja-todo-aviso.visible { display:flex; }
    .caja-todo-aviso { align-items:center; gap:13px; border:2px solid var(--red); background:#fef2f2; border-radius:10px; padding:15px 17px; }
    .caja-todo-icono { font-size:27px; line-height:1; flex-shrink:0; }
    .caja-todo-texto { font-size:14.5px; font-weight:800; color:var(--red-dark); line-height:1.4; }

    .caja-frase-preview { background:#fff; border:1.5px dashed #a5b4fc; border-radius:10px; padding:15px 17px; font-size:15px; color:#3730a3; margin-top:8px; min-height:24px; line-height:1.6; word-break:break-word; }
    .caja-frase-preview.actualizada { animation: cajaFraseFlash .5s ease-out; }
    @keyframes cajaFraseFlash {
      0%   { background:#ede9fe; border-color:#8b5cf6; }
      100% { background:#fff; border-color:#a5b4fc; }
    }
    .caja-frase-acciones { display:flex; align-items:center; gap:14px; margin-top:13px; flex-wrap:wrap; }
    .caja-btn-copiar { border:1px solid #a5b4fc; background:#fff; color:#4338ca; font-weight:700; font-size:14px; padding:10px 20px; border-radius:22px; cursor:pointer; font-family:inherit; }
    .caja-btn-copiar:hover { background:#eef2ff; }
    .caja-frase-estado { font-size:13px; font-weight:700; color:#94a3b8; }
    .caja-frase-estado.copiado { color:var(--green-dark); }

    .cm-overlay { position:fixed; inset:0; background:rgba(15,23,42,.62); display:flex; align-items:center; justify-content:center; z-index:700; padding:18px; }
    .cm-overlay[hidden] { display:none; }
    .cm-card { background:var(--card); border-radius:var(--radius); box-shadow:var(--shadow); max-width:700px; width:100%; max-height:94vh; display:flex; flex-direction:column; overflow:hidden; animation: fadeInUp .22s var(--ease); }
    .cm-head { background:#1B2636; padding:18px 28px 30px; color:#fff; display:flex; align-items:center; gap:18px; position:relative; overflow:visible; }
    .cm-head-avatar-wrap { width:100px; height:128px; flex-shrink:0; position:relative; margin-top:-6px; }
    .cm-head-avatar { width:100px; height:128px; display:block; overflow:visible; }
    .cm-head-info { min-width:0; }
    .cm-head-titulo { margin:0 0 4px; font-size:19px; font-weight:800; font-family: Georgia, "Times New Roman", Times, serif; }
    .cm-head-sub { margin:0; font-size:13.5px; opacity:.8; }
    .cm-head-ref { margin:7px 0 0; font-size:13px; opacity:.92; background:rgba(255,255,255,.12); display:inline-block; padding:4px 11px; border-radius:20px; font-weight:700; letter-spacing:.02em; }
    .cm-head-close { position:absolute; top:14px; right:16px; background:rgba(255,255,255,.14); border:none; color:#fff; width:30px; height:30px; border-radius:50%; cursor:pointer; font-size:16px; line-height:1; z-index:2; }
    .cm-head-close:hover { background:rgba(255,255,255,.26); }
    .cm-swoosh { position:absolute; left:0; right:0; bottom:0; width:100%; height:14px; }

    .version-tag { position:fixed; right:6px; bottom:4px; font-size:9px; color:#94a3b8; opacity:.45; z-index:2000; pointer-events:none; font-family:inherit; user-select:none; }
    .cm-body { padding:24px 28px 26px; overflow-y:auto; }

    .cm-body .caja-compositor { display:block; margin-top:0; padding:0; background:none; border:none; animation:none; }
    .cm-body .caja-compositor-titulo { display:none; }

    .cm-gafas-lente { fill:rgba(255,255,255,.14); stroke:#f5f5f5; stroke-width:2.6; }
    .cm-gafas-puente { stroke:#f5f5f5; stroke-width:2.4; fill:none; }
    .cm-gafas-brillo { fill:#fff; opacity:.55; }

    .cm-brazo, .cm-mano { fill:url(#rufo-grad); }
    .cm-libreta { fill:#fdfaf5; stroke:#d8d2c8; stroke-width:1.2; }
    .cm-lapiz-grande { transform-box: fill-box; transform-origin: 92% 88%; animation: cmLapizGrande 1s ease-in-out infinite; }
    @keyframes cmLapizGrande {
      0%, 100% { transform: rotate(-6deg) translate(0,0); }
      50%      { transform: rotate(5deg) translate(-1px,.5px); }
    }
    .cm-renglon { stroke:#E07A1F; stroke-width:2.2; fill:none; stroke-linecap:round; stroke-dasharray:16; stroke-dashoffset:16; opacity:0; }
    .cm-renglon-1 { animation: cmRenglon 1.9s ease-in-out infinite; }
    .cm-renglon-2 { animation: cmRenglon 1.9s ease-in-out infinite; animation-delay:.6s; }
    .cm-renglon-3 { animation: cmRenglon 1.9s ease-in-out infinite; animation-delay:1.2s; }
    @keyframes cmRenglon {
      0%   { stroke-dashoffset:16; opacity:0; }
      12%  { opacity:1; }
      45%  { stroke-dashoffset:0; opacity:1; }
      70%  { stroke-dashoffset:0; opacity:1; }
      92%  { opacity:0; }
      100% { opacity:0; }
    }

    .note-card { background: var(--card); border-radius: var(--radius-sm); margin-bottom: 14px; padding: 16px; box-shadow: var(--shadow-sm); border:1px solid var(--border); border-left: 4px solid var(--border); transition:.2s; animation: fadeInUp .3s var(--ease); }
    .note-card:hover { box-shadow: var(--shadow); }
    .note-header { font-size: 11.5px; color: var(--muted); margin-bottom: 10px; border-bottom: 1px dashed var(--border); padding-bottom: 8px; display:flex; align-items:center; flex-wrap:wrap; gap:4px; }
    .note-body { font-size: 15px; line-height: 1.55; }
    .badge-nueva { background: linear-gradient(135deg,#f87171,var(--red-dark)); color:white; padding:3px 9px; border-radius:20px; font-size:10.5px; font-weight:800; margin-left:6px; }
    .notas-lista { display:flex; flex-direction:column; }
    .note-actions { display:flex; gap:8px; margin-top:10px; }
    .empty-state { padding:44px 20px; text-align:center; color: var(--muted); font-size:14px; }
    .ac-linea { display:flex; align-items:center; gap:8px; flex-wrap:wrap; background:#fffbeb; border:1px solid #fcd34d; border-radius:10px; padding:7px 10px; margin:0 2px 10px; font-size:12px; color:#92400e; }
    .ac-linea[hidden], .ac-chip[hidden] { display:none; }
    .ac-linea b { font-weight:800; }
    .ac-ag { display:inline-flex; align-items:center; gap:4px; background:#fff; border:1px solid #fde68a; border-radius:7px; padding:2px 7px; font-weight:800; color:#78350f; }
    .ac-ag small { font-weight:600; color:#b45309; }
    .ac-chip { display:flex; align-items:center; gap:8px; background:#fffbeb; border:1px solid #fcd34d; border-radius:12px; padding:8px 11px; font-size:12.5px; color:#92400e; line-height:1.35; cursor:pointer; }
    .ac-chip b { color:#78350f; }
    .ac-badge { display:inline-flex; align-items:center; gap:3px; background:linear-gradient(135deg,#fbbf24,#d97706); color:#fff; font-weight:800; font-size:11.5px; border-radius:20px; padding:2px 8px; box-shadow:0 2px 6px -2px rgba(217,119,6,.6); animation:acLate 2.4s ease-in-out infinite; }
    @keyframes acLate { 0%,100% { transform:scale(1); } 50% { transform:scale(1.08); } }
    .bp-marcas { display:flex; flex-wrap:wrap; gap:5px; margin-top:6px; }
    .bp-marca { display:inline-flex; align-items:center; gap:4px; font-size:11px; font-weight:800; border-radius:7px; padding:2px 7px; line-height:1.35; }
    .bp-m { background:#fef3c7; color:#92400e; border:1px solid #fcd34d; }
    .bp-r { background:#fee2e2; color:#b91c1c; border:1px solid #fca5a5; }
    .bp-c { background:#f1f5f9; color:#334155; border:1px solid #cbd5e1; }
    .bp-c i { width:11px; height:11px; border-radius:3px; border:1px solid rgba(0,0,0,.25); display:inline-block; }
    .nt-tipos { display:flex; gap:6px; margin-bottom:8px; }
    .nt-tipo { flex:1; border:1.5px solid var(--border); border-radius:10px; padding:7px 4px; text-align:center; font-size:12px; font-weight:700; color:var(--muted); background:#fff; cursor:pointer; font-family:inherit; }
    .nt-tipo.on { border-color:#E07A1F; color:#E07A1F; background:#fff1f2; }
    .nt-escribir textarea.nt-texto { height:96px; min-height:96px; }
    .nt-pestanas { display:flex; background:#e2e8f0; border-radius:10px; padding:3px; margin:14px 0 10px; }
    .nt-pes { flex:1; border:none; background:none; font-family:inherit; font-size:12px; font-weight:700; padding:7px 0; border-radius:8px; color:var(--muted); cursor:pointer; }
    .nt-pes.on { background:#fff; color:var(--text); box-shadow:0 1px 3px rgba(0,0,0,.12); }
    .nt-pill { font-style:normal; display:inline-block; background:#E07A1F; color:#fff; border-radius:999px; font-size:10px; padding:0 6px; margin-left:3px; line-height:16px; }
    .nt-pill.gris { background:#94a3b8; }
    .note-card.nt-nota { position:relative; background:#fffdf5; border:1px solid #fde68a; border-left:1px solid #fde68a; border-radius:12px; padding:11px 11px 9px 14px; box-shadow:0 3px 8px -5px rgba(180,120,0,.45); }
    .note-card.nt-nota:before { content:""; position:absolute; left:0; top:9px; bottom:9px; width:4px; border-radius:0 4px 4px 0; background:#f59e0b; }
    .note-card.nt-nota.turno { background:#f8fbff; border-color:#bfdbfe; box-shadow:0 3px 8px -5px rgba(29,78,216,.35); }
    .note-card.nt-nota.turno:before { background:#3b82f6; }
    .note-card.nt-nota.hecha { background:#f8fafc; border-color:var(--border); box-shadow:none; opacity:.8; }
    .note-card.nt-nota.hecha:before { background:#94a3b8; }
    .nt-arriba { display:flex; align-items:center; gap:6px; flex-wrap:wrap; font-size:11px; color:var(--muted); }
    .nt-chip { display:inline-flex; align-items:center; gap:3px; font-size:11px; font-weight:800; border-radius:6px; padding:1px 6px; }
    .nt-chip.ped { background:#fef3c7; color:#92400e; } .nt-chip.tur { background:#dbeafe; color:#1d4ed8; }
    .nt-chip.nueva { background:#fee2e2; color:#b91c1c; } .nt-chip.ok { background:#dcfce7; color:#15803d; }
    .nt-cuerpo { font-size:14px; margin:7px 0 8px; line-height:1.45; }
    .nt-pie { display:flex; align-items:center; justify-content:space-between; gap:6px; flex-wrap:wrap; }
    .nt-quien { font-size:11px; color:var(--muted); } .nt-quien b { color:#15803d; }
    .nt-accs { display:flex; gap:4px; }
    .nt-acc { border:1px solid var(--border); background:#fff; border-radius:8px; font-size:11.5px; padding:4px 8px; font-weight:700; color:#334155; cursor:pointer; font-family:inherit; }
    .nt-acc.res { border-color:#86efac; color:#15803d; background:#f0fdf4; }
    .nt-acc:disabled { opacity:.6; cursor:default; }
    .nt-posit { margin-top:9px; background:linear-gradient(180deg,#fff7c2,#ffef9a); border-radius:4px 4px 10px 4px; padding:9px 9px 7px; font-size:12.5px; line-height:1.4; white-space:pre-line; box-shadow:0 4px 8px -4px rgba(160,110,0,.5); transform:rotate(-.6deg); position:relative; color:#1e293b; }
    .nt-posit:before { content:""; position:absolute; top:-6px; left:50%; width:36px; height:12px; margin-left:-18px; background:rgba(255,255,255,.65); border:1px solid rgba(0,0,0,.05); transform:rotate(2deg); }
    .nt-posit small { display:block; color:#92400e; font-size:10.5px; margin-top:3px; white-space:normal; }
    .nt-posit .nt-acc { margin-top:6px; }
    .nt-btn-nota { flex:0 0 40px !important; min-width:40px; border:1px solid #fcd34d; background:#fff7c2; color:#92400e; border-radius:10px; font-size:15px; cursor:pointer; }
    .nt-form { margin-top:8px; background:#fffbeb; border:1px dashed #fcd34d; border-radius:10px; padding:8px; }
    .nt-form textarea { width:100%; min-height:56px; border:1px solid #fde68a; border-radius:8px; padding:6px 8px; font-family:inherit; font-size:12.5px; resize:vertical; box-sizing:border-box; }
    .nt-form .nt-accs { margin-top:6px; justify-content:flex-end; }
    .nt-franja { background:linear-gradient(135deg,#1e293b,#0f172a); color:#fff; border-radius:14px; padding:10px 12px; box-shadow:0 8px 18px -10px rgba(0,0,0,.6); animation: fadeInUp .3s var(--ease); }
    .nt-franja[hidden] { display:none; }
    .nt-franja-t { font-size:12.5px; font-weight:800; display:flex; justify-content:space-between; align-items:center; }
    .nt-franja-t span { background:#E07A1F; border-radius:999px; font-size:10.5px; padding:1px 7px; }
    .nt-mini { background:rgba(255,255,255,.08); border-radius:10px; padding:8px 9px; margin-top:8px; font-size:12.5px; line-height:1.4; white-space:pre-line; }
    .nt-mini small { display:block; color:#94a3b8; font-size:10.5px; margin-top:2px; white-space:normal; }
    .nt-ent { margin-top:6px; background:#22c55e; color:#052e16; border:none; border-radius:7px; font-weight:800; font-size:11.5px; padding:4px 9px; cursor:pointer; font-family:inherit; }
    .nt-ent:disabled { opacity:.6; }
    .empty-state .emoji { display:block; font-size:30px; margin-bottom:10px; }

    .avisos-lista { display:flex; flex-direction:column; }
    .aviso-card { background: var(--card); border-radius: var(--radius-sm); margin-bottom: 14px; padding: 16px; box-shadow: var(--shadow-sm); border:1px solid var(--border); border-left: 4px solid var(--border); transition:.2s; animation: fadeInUp .3s var(--ease); }
    .aviso-card:hover { box-shadow: var(--shadow); }
    .aviso-card-top { display:flex; align-items:center; gap:8px; flex-wrap:wrap; margin-bottom:8px; }
    .aviso-ref { font-weight:800; font-size:14.5px; color:#1e293b; background:#f1f5f9; border-radius:7px; padding:3px 9px; font-family:'SFMono-Regular',Consolas,monospace; letter-spacing:.02em; }
    .aviso-contador { background:#E07A1F; color:#fff; font-size:11px; font-weight:800; border-radius:20px; padding:3px 10px; }
    .aviso-meta { margin-left:auto; font-size:11px; color:var(--muted); white-space:nowrap; }
    .aviso-motivo { font-size:14px; line-height:1.5; color:#334155; }
    .aviso-acciones { display:flex; gap:8px; margin-top:10px; }
    .aviso-hist { margin-top:10px; padding-top:8px; border-top:1px dashed var(--border); font-size:11.5px; color:var(--muted); }
    .aviso-hist-item { display:flex; gap:6px; padding:3px 0; }
    .aviso-hist-item b { color:#334155; }
    .aviso-ya-escrito { background:#fffbeb; border:1.5px solid #fcd34d; color:#92400e; border-radius:12px; padding:10px 12px; font-size:13px; line-height:1.45; margin:12px 0; }

    .aviso-toast { position:fixed; left:12px; right:12px; bottom:12px; z-index:600; background:#161616; color:#fff; border-radius:12px; box-shadow:0 10px 26px -4px rgba(0,0,0,.4); padding:12px 14px; display:flex; align-items:center; gap:10px; flex-wrap:wrap; animation: recordatorioEntra .35s var(--ease); }
    .aviso-toast[hidden] { display:none; }
    .aviso-toast-icono { font-size:22px; flex-shrink:0; }
    .aviso-toast-texto { flex:1; min-width:120px; font-size:12.5px; line-height:1.4; word-break:break-word; }
    .aviso-toast-texto strong { color:#fff; }
    .aviso-toast-donde { color:#a3a3a3; font-size:11px; margin-top:3px; display:block; }
    .aviso-toast-btn { background:#E07A1F; color:#fff; border:none; padding:8px 12px; border-radius:8px; cursor:pointer; font-weight:700; font-size:11.5px; white-space:nowrap; }
    .aviso-toast-btn:hover { background:#a80d26; }
    .aviso-toast-cerrar { background:none; border:none; color:#bbb; font-size:15px; line-height:1; cursor:pointer; padding:4px 6px; flex-shrink:0; }
    .aviso-toast-cerrar:hover { color:#fff; }
    @media (max-width: 340px) {
      .aviso-toast-btn { flex:1 1 100%; }
    }

    .tbl { width: 100%; border-collapse: collapse; background:white; border-radius:var(--radius-sm); overflow:hidden; box-shadow: var(--shadow-sm); }
    .tbl th, .tbl td { padding: 12px 14px; border-bottom: 1px solid var(--border); text-align: left; }
    .tbl th { background: #f8fafc; color: var(--muted); font-size:11.5px; text-transform:uppercase; letter-spacing:.02em; }
    .tbl tr:last-child td { border-bottom:none; }
    .tbl tr:hover td { background:#f8fafc; }

    .header-cal { display:flex; justify-content:space-between; align-items:center; margin-bottom:10px; background:white; padding:10px 14px; border-radius:var(--radius-sm); box-shadow: var(--shadow-sm); border:1px solid var(--border); flex-wrap:wrap; gap:10px; }
    .header-cal h2 { margin:0; font-size:16px; font-weight:800; }
    .header-cal-actions { display:flex; gap:8px; }
    .btn-cal { background:#334155; color:white; border:none; padding:9px 13px; border-radius:8px; cursor:pointer; font-weight:700; font-size:12.5px; transition:.15s; }
    .btn-cal:hover { background:#1e293b; transform:translateY(-1px); }
    .btn-cal-orange { background:#ea580c; } .btn-cal-orange:hover { background:#c2410c; }
    .btn-cal-red { background:#dc2626; } .btn-cal-red:hover { background:#b91c1c; }
    .btn-cal-amber { background:#b45309; } .btn-cal-amber:hover { background:#92400e; }
    .grid-cal { display:grid; grid-template-columns:repeat(7,1fr); gap:10px; }
    .day-card { background:white; border-radius:var(--radius-sm); box-shadow: var(--shadow-sm); border:1px solid var(--border); display:flex; flex-direction:column; overflow:hidden; transition:.2s; animation: fadeInUp .3s var(--ease); }
    .day-card:hover { box-shadow: var(--shadow); transform: translateY(-2px); }
    .day-header { background:#f1f5f9; padding:10px; text-align:center; font-weight:bold; font-size:12.5px; border-bottom:2px solid var(--border); }
    .day-header.today { background:linear-gradient(135deg, var(--primary-2), var(--primary-dark)); color:white; border-bottom-color:var(--primary-dark); }
    .day-total { text-align:center; padding:9px 0; font-weight:bold; background:#f8fafc; border-bottom:1px solid var(--border); font-size:12.5px; color:#334155; }
    .agencies { padding:10px; font-size:12px; }
    .row { margin-bottom:6px; border-bottom:1px dashed #f1f5f9; padding-bottom:4px; }
    .ag-badge-card { display:inline-flex; flex-direction:column; align-items:center; width:80px; box-sizing:border-box; border-radius:7px; vertical-align:middle; background:#fff; border:1px solid #e2e8f0; padding:4px 4px 3px; }
    .ag-badge-logo { display:flex; align-items:center; justify-content:center; width:100%; height:30px; overflow:hidden; border-radius:5px; }
    .ag-badge-logo img { max-width:100%; max-height:100%; object-fit:contain; display:block; }
    .ag-badge-color .ag-badge-logo { color:#fff; font-size:14px; font-weight:800; letter-spacing:.03em; }
    .ag-badge-label { font-size:9px; font-weight:700; color:#64748b; letter-spacing:.03em; margin-top:3px; max-width:100%; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .ag-badge-card.ag-badge-sm { width:56px; padding:3px 3px 2px; border-radius:6px; }
    .ag-badge-card.ag-badge-sm .ag-badge-logo { height:22px; border-radius:4px; }
    .ag-badge-card.ag-badge-sm.ag-badge-color .ag-badge-logo { font-size:11px; }
    .ag-badge-card.ag-badge-sm .ag-badge-label { font-size:7.5px; margin-top:2px; }
    .row-top { display:flex; justify-content:space-between; align-items:center; }
    .count { font-weight:bold; background:#f1f5f9; padding:2px 7px; border-radius:10px; color:#334155; }
    .count.z { background:transparent; color:#cbd5e1; }
    .row-sub { display:flex; justify-content:flex-end; gap:8px; font-size:10.5px; margin-top:2px; }
    .mini-rot { color: var(--red-dark); font-weight:700; white-space:nowrap; }
    .mini-pend-ok { color: var(--green-dark); font-weight:700; white-space:nowrap; }
    .mini-pend-alerta { color: var(--red-dark); font-weight:700; white-space:nowrap; }
    .mini-retornos { color: #7c3aed; font-weight:700; white-space:nowrap; }

    .cal-legend { display:flex; flex-wrap:wrap; justify-content:center; align-items:center; gap:8px 20px; background:var(--card); border:1px solid var(--border); border-radius:var(--radius-sm); box-shadow: var(--shadow-sm); padding:12px 18px; margin-top:14px; animation: fadeInUp .3s var(--ease); }
    .legend-item { display:flex; align-items:center; gap:8px; font-size:12px; font-weight:600; color:#334155; }
    .legend-icon { width:24px; height:24px; border-radius:50%; display:inline-flex; align-items:center; justify-content:center; font-size:13px; flex-shrink:0; }
    .legend-roto { background:#fee2e2; }
    .legend-pend { background:#fef3c7; }
    .legend-retornos { background:#ede9fe; }

    .result-box { text-align:center; padding: 44px 20px; animation: fadeInUp .3s var(--ease); }
    .result-box-icon { font-size:34px; margin-bottom:6px; }
    .result-box h3 { margin: 4px 0 6px; font-size:17px; }
    .result-box p { color: var(--muted); font-size:13px; margin:0; }

    .print-view { display:none; }
    .print-letterhead { display:flex; justify-content:space-between; align-items:flex-start; gap:14px; border-bottom:3px solid #E07A1F; padding-bottom:12px; margin-bottom:18px; }
    .print-letterhead-brand { display:flex; flex-direction:column; }
    .print-letterhead-logo { font-family: Georgia, 'Times New Roman', serif; font-size:32px; font-weight:bold; letter-spacing:1px; color:#111; }
    .print-letterhead-logo-accent { color:#E07A1F; }
    .print-letterhead-dept { font-size:11px; text-transform:uppercase; letter-spacing:1.5px; color:#666; margin-top:2px; }
    .print-letterhead-title { flex:1; text-align:center; padding-top:4px; }
    .print-letterhead-title h1 { font-size:20px; margin:0 0 4px; color:#111; }
    .print-letterhead-title p { font-size:12.5px; margin:0; color:#555; }
    .print-letterhead-meta { font-size:11px; color:#444; text-align:right; line-height:1.7; white-space:nowrap; }
    .print-tables { display:flex; flex-wrap:wrap; gap:12px; align-items:flex-start; }
    .print-tables.print-tables-dia .print-day-table { flex:1 1 100%; max-width:560px; margin:0 auto; }
    .print-day-table { flex:1 1 22%; min-width:190px; width:100%; table-layout:fixed; border-collapse:collapse; font-size:12.5px; margin-bottom:10px; }
    .print-day-table.print-hide { display:none; }
    .print-day-title th { background:#1a1a1a; color:#fff; padding:7px 8px; font-size:13px; text-align:left; }
    .print-day-table.print-today .print-day-title th { background:#E07A1F; }
    .print-badge-hoy { background:#fff; color:#E07A1F; font-size:9.5px; font-weight:bold; padding:1.5px 5px; border-radius:3px; margin-left:5px; }
    .print-col-headers th { background:#f1f1f1; border-bottom:1px solid #ccc; padding:5px 8px; font-size:10px; text-transform:uppercase; color:#444; text-align:left; line-height:1.25; }
    .print-day-table td { padding:5px 8px; border-bottom:1px solid #eee; vertical-align:middle; }
    .print-day-table tr.print-total-row td { font-weight:bold; border-top:1.5px solid #333; border-bottom:none; background:#f8f8f8; }
    .print-day-table .col-pct { text-align:right; }
    .print-day-table .col-pct.col-pct-rota { color:#E07A1F; font-weight:bold; }
    .print-week-table { flex:1 1 100%; max-width:560px; margin:0 auto; }
    .print-footer { margin-top:18px; padding-top:8px; border-top:1px solid #ccc; font-size:10px; color:#999; text-align:center; }

    @media print {
       @page { size: A4 landscape; margin: 10mm 9mm; }
       body { background: white; margin: 0; padding: 0; }
       .no-print, .grid-cal, .cal-legend, #pendientes-wrap, .checklist-overlay { display: none !important; }
       .print-view { display: block !important; }
       .print-day-table thead { display: table-header-group; }
       .print-day-table { break-inside: avoid; }
    }

    .checklist-overlay { position:fixed; inset:0; background:rgba(15,23,42,.55); display:flex; align-items:center; justify-content:center; z-index:500; padding:20px; }
    .checklist-overlay[hidden] { display:none; }
    .checklist-card { background:var(--card); border-radius:var(--radius); box-shadow:var(--shadow); max-width:860px; width:100%; max-height:86vh; display:flex; flex-direction:column; overflow:hidden; animation: fadeInUp .2s var(--ease); }
    .checklist-card-head { padding:16px 20px; border-bottom:1px solid var(--border); display:flex; justify-content:space-between; align-items:center; gap:10px; }
    .checklist-card-head h3 { margin:0; font-size:14.5px; font-weight:800; }
    .checklist-card-close { background:none; border:none; font-size:18px; line-height:1; cursor:pointer; color:var(--muted); padding:4px 7px; border-radius:6px; flex-shrink:0; }
    .checklist-card-close:hover { background:#f1f5f9; }
    .checklist-card-body { padding:16px 20px; overflow-y:auto; overflow-x:auto; }
    .checklist-doc-preview { background:#fff; border:1px solid var(--border); border-radius:10px; padding:20px 22px; min-width:480px; }
    .checklist-warn { margin:0 0 12px; font-size:11.5px; color:#92400e; background:#fffbeb; border:1px solid #fde68a; border-radius:8px; padding:9px 11px; }
    .checklist-ok-msg { margin:0; font-size:12.5px; color:#065f46; background:#ecfdf5; border:1px solid #a7f3d0; border-radius:8px; padding:12px 14px; font-weight:600; line-height:1.5; }
    .checklist-card-foot { padding:14px 20px; border-top:1px solid var(--border); background:#f8fafc; }
    .checklist-foot-group { display:flex; gap:10px; justify-content:flex-end; flex-wrap:wrap; }
    .btn-checklist-cancel { background:#e2e8f0; color:#334155; border:none; padding:9px 16px; border-radius:8px; cursor:pointer; font-weight:700; font-size:12.5px; }
    .btn-checklist-cancel:hover { background:#cbd5e1; }
    .btn-checklist-print { background:#fff; color:#334155; border:1.5px solid var(--border); padding:8px 16px; border-radius:8px; cursor:pointer; font-weight:700; font-size:12.5px; margin-right:auto; }
    .btn-checklist-print:hover { border-color:#94a3b8; background:#f8fafc; }
    .btn-checklist-confirm { background:var(--primary); color:#fff; border:none; padding:9px 16px; border-radius:8px; cursor:pointer; font-weight:700; font-size:12.5px; }
    .btn-checklist-confirm:hover { background:var(--primary-dark); }
    .btn-checklist-confirm:disabled { opacity:.6; cursor:default; }

    @media (max-width: 360px) {
       body { padding: 14px; }
       .search-card { padding: 16px; }
    }

    .rufo-colapsado { display:flex; flex-direction:column; align-items:center; justify-content:center; gap:16px; padding: 70px 20px; text-align:center; }
    .rufo-avatar-click { cursor:pointer; display:inline-block; }
    .rufo-avatar-grande { width: 96px; height: 96px; filter: drop-shadow(0 6px 14px rgba(0,0,0,.18)); }

    .rufo-colapsado.dormido .rufo-avatar-grande {
      animation: rufoRespirarDormido 4s ease-in-out infinite;
      transform-box: fill-box; transform-origin: 50% 100%;
    }
    @keyframes rufoRespirarDormido {
      0%, 100% { transform: translateY(0) rotate(0deg); }
      50%      { transform: translateY(-4px) rotate(-2deg); }
    }
    .rufo-parpados-dormido { opacity: 0; transition: opacity .12s ease; }
    .rufo-colapsado.dormido .rufo-parpados-dormido { opacity: 1; }

    .rufo-zzz { opacity: 0; }
    .rufo-colapsado.dormido .rufo-zzz-1 { animation: rufoZzz 3s ease-in-out infinite; }
    .rufo-colapsado.dormido .rufo-zzz-2 { animation: rufoZzz 3s ease-in-out infinite .7s; }
    .rufo-colapsado.dormido .rufo-zzz-3 { animation: rufoZzz 3s ease-in-out infinite 1.4s; }
    @keyframes rufoZzz {
      0%   { opacity: 0; transform: translate(0,0) scale(.6); }
      15%  { opacity: .85; }
      70%  { opacity: .85; }
      100% { opacity: 0; transform: translate(14px, -22px) scale(1.15); }
    }

    .rufo-colapsado.despertando .rufo-avatar-grande {
      animation: rufoDespertar .55s cubic-bezier(.36,1.65,.4,1);
    }
    @keyframes rufoDespertar {
      0%   { transform: translateY(-4px) rotate(-2deg) scale(1); }
      25%  { transform: translateY(-11px) rotate(5deg) scale(1.08); }
      50%  { transform: translateY(2px) rotate(-4deg) scale(.96); }
      75%  { transform: translateY(-3px) rotate(3deg) scale(1.03); }
      100% { transform: translateY(0) rotate(0deg) scale(1); }
    }
    .rufo-boca-sorpresa { opacity: 0; transform-box: fill-box; transform-origin: center; }
    .rufo-colapsado.despertando .rufo-boca-sorpresa {
      animation: rufoBocaSorpresa .55s ease-out;
    }
    @keyframes rufoBocaSorpresa {
      0%   { opacity: 0; transform: scale(.4); }
      20%  { opacity: 1; transform: scale(1.2); }
      55%  { opacity: 1; transform: scale(1); }
      100% { opacity: 0; transform: scale(.75); }
    }
    .rufo-btn-hablar { background:#1B2636; color:#fff; border:none; padding:12px 22px; border-radius: 24px; font-size:14px; font-weight:700; cursor:pointer; box-shadow: 0 4px 14px rgba(0,0,0,.22); transition: transform .15s; }
    .rufo-btn-hablar:hover { transform: scale(1.04); }

    .rufo-panel { background: var(--card); border-radius: var(--radius); overflow:hidden; border:1px solid var(--border); box-shadow: var(--shadow-sm); animation: fadeInUp .25s var(--ease); }
    .rufo-header { background:#1B2636; padding: 18px 18px 20px; color:#fff; display:flex; align-items:center; gap:12px; position:relative; overflow:hidden; }
    .encabezado-panel { border-radius: var(--radius); margin-bottom: 10px; box-shadow: var(--shadow-sm); padding: 11px 14px; }
    .encabezado-panel .rufo-avatar-mini-wrap { width:38px; height:38px; }
    .encabezado-panel .rufo-avatar-mini { width:38px; height:38px; animation: none; }
    .encabezado-panel .rufo-gesto { width:18px; height:18px; right:-4px; bottom:-3px; }
    .encabezado-panel .rufo-nombre { font-size:14px; }
    .encabezado-panel .rufo-sub { font-size:11px; }

    .rufo-avatar-mini-wrap { position:relative; width:52px; height:52px; flex-shrink:0; }
    .rufo-avatar-mini {
      width:52px; height:52px; display:block;
      animation: rufoRespirar 3.4s ease-in-out infinite;
      transform-box: fill-box; transform-origin: 50% 100%;
    }
    @keyframes rufoRespirar {
      0%, 100% { transform: translateY(0) rotate(0deg); }
      50% { transform: translateY(-2px) rotate(-1.5deg); }
    }
    .rufo-gesto { position:absolute; right:-6px; bottom:-4px; width:24px; height:24px; overflow:visible; }
    .rufo-mini-paquete {
      opacity:0; animation: rufoPaquete 5s ease-in-out infinite;
      transform-box: fill-box; transform-origin: center;
    }
    @keyframes rufoPaquete {
      0%, 55%  { opacity:0; transform: scale(.6) translateY(6px); }
      62%, 80% { opacity:1; transform: scale(1) translateY(0); }
      92%,100% { opacity:0; transform: scale(.8) translateY(-6px); }
    }
    .rufo-sello {
      opacity:0; animation: rufoSello 5s ease-in-out infinite;
      transform-box: fill-box; transform-origin: center;
    }
    @keyframes rufoSello {
      0%, 68%  { transform: translateY(-12px) rotate(-20deg) scale(.8); opacity:0; }
      74%      { transform: translateY(0) rotate(0deg) scale(1); opacity:1; }
      80%, 94% { transform: translateY(0) rotate(0deg) scale(1); opacity:1; }
      100%     { transform: translateY(0) rotate(0deg) scale(1); opacity:0; }
    }

    .rufo-nombre { margin:0; font-size:16px; font-family: Georgia, "Times New Roman", Times, serif; }
    .rufo-sub { margin:2px 0 0; font-size:12px; opacity:.75; }
    .rufo-swoosh { position:absolute; left:0; right:0; bottom:0; width:100%; height:12px; }
    .rufo-body { padding: 16px; display:flex; flex-direction:column; gap:12px; }
    .rufo-burbuja { background:#f2f2f2; border-radius:14px 14px 14px 4px; padding:12px 14px; font-size:13.5px; line-height:1.5; }
    .rufo-chips { display:flex; flex-direction:column; gap:8px; }
    .rufo-chip { display:flex; align-items:center; gap:10px; background:#fff; border:1px solid var(--border); border-radius:12px; padding:11px 13px; font-size:13.5px; font-weight:600; color:#1B2636; cursor:pointer; transition:.15s; text-align:left; width:100%; }
    .rufo-chip:hover { border-color:#E07A1F; background:#fdf1f1; transform: translateX(2px); }
    .rufo-chip span { font-size:16px; }

    .rufo-chip-right { margin-left:auto; display:flex; align-items:center; gap:6px; flex-shrink:0; }
    .rufo-badge {
      min-width:20px; height:20px; padding:0 5px; border-radius:10px;
      background:#E07A1F; color:#fff; font-size:11px; font-weight:800;
      display:flex; align-items:center; justify-content:center; line-height:1;
      flex-shrink:0;
    }

    .mini-caja-wrap {
      width:30px; height:26px; flex-shrink:0;
      display:flex; align-items:center; justify-content:center;
      opacity:0; transform:scale(.7);
      transition: opacity .15s ease, transform .15s ease;
      pointer-events:none;
    }
    .mini-caja-wrap.activo { opacity:1; transform:scale(1); }
    .mini-caja-svg { width:30px; height:26px; overflow:visible; }
    .mini-caja-tapa { transform-box: fill-box; transform-origin: 12% 100%; }
    .mini-caja-wrap.activo .mini-caja-tapa { animation: rufoTapaPop .5s cubic-bezier(.34,1.56,.64,1); }
    @keyframes rufoTapaPop {
      0%   { transform: translateY(0) rotate(0deg); }
      55%  { transform: translateY(-9px) translateX(2px) rotate(-46deg); }
      100% { transform: translateY(0) rotate(0deg); }
    }
    .mini-caja-confeti-pieza { animation: rufoConfetiVuelo .55s ease-out forwards; }
    @keyframes rufoConfetiVuelo {
      0%   { opacity:1; transform: translate(0,0) rotate(0deg); }
      100% { opacity:0; transform: translate(var(--dx), var(--dy)) rotate(var(--rot)); }
    }

    .volver-rufo { display:inline-flex; align-items:center; gap:6px; background:#1B2636; color:#fff; border:none; padding:6px 12px; border-radius:20px; font-size:12px; font-weight:700; cursor:pointer; margin-bottom:8px; }
    .volver-rufo:hover { opacity:.85; }

    .recordatorio-checklist { position:fixed; left:12px; right:12px; bottom:12px; z-index:600; background:#161616; color:#fff; border-radius:12px; box-shadow:0 10px 26px -4px rgba(0,0,0,.4); padding:12px 14px; display:flex; align-items:center; gap:10px; flex-wrap:wrap; animation: recordatorioEntra .35s var(--ease); }
    .recordatorio-checklist[hidden] { display:none; }
    .recordatorio-checklist-icono { font-size:22px; flex-shrink:0; }
    .recordatorio-checklist-texto { flex:1; min-width:120px; font-size:12.5px; line-height:1.4; }
    .recordatorio-checklist-texto strong { color:#fff; }
    .recordatorio-checklist-btn { background:#E07A1F; color:#fff; border:none; padding:8px 12px; border-radius:8px; cursor:pointer; font-weight:700; font-size:11.5px; white-space:nowrap; }
    .recordatorio-checklist-btn:hover { background:#a80d26; }
    .recordatorio-checklist-cerrar { background:none; border:none; color:#bbb; font-size:15px; line-height:1; cursor:pointer; padding:4px 6px; flex-shrink:0; }
    .recordatorio-checklist-cerrar:hover { color:#fff; }
    @keyframes recordatorioEntra { from{ opacity:0; transform:translateY(14px); } to{ opacity:1; transform:translateY(0); } }
    .pr-marcha { position:fixed; left:0; right:0; bottom:6px; height:160px; z-index:700; pointer-events:none; overflow:hidden; animation:prFuera 2.8s linear forwards; }
    .pr-marcha-oso { position:absolute; bottom:24px; left:-80px; width:60px; height:88px; animation:prCorre 2.4s linear forwards; will-change:left; }
    .pr-marcha-salto { position:absolute; inset:0; transform-origin:50% 100%; animation:prSalta .28s ease-in-out infinite alternate; will-change:transform; }
    .pr-marcha-salto svg { width:100%; height:100%; overflow:visible; display:block; }
    .pr-pila-2 { animation:prSeVa .01s linear .48s forwards; }
    .pr-pila-1 { animation:prSeVa .01s linear 1.92s forwards; }
    .pr-gota-frente { animation:prGotaFrente .7s ease-in infinite; }
    .pr-sudor { position:absolute; width:6px; height:8px; background:#60a5fa; border-radius:50% 50% 50% 50% / 62% 62% 38% 38%; opacity:0; box-shadow:inset 1px 1px 0 rgba(255,255,255,.7); }
    .pr-sudor-i { left:6px; top:30px; animation:prSudorI .55s ease-out infinite; }
    .pr-sudor-d { left:50px; top:28px; animation:prSudorD .55s ease-out infinite; }
    .pr-caida { position:absolute; bottom:101px; width:14px; height:9px; background:#E07A1F; border-radius:2px; box-shadow:inset 0 -2px 0 rgba(0,0,0,.18); opacity:0; }
    .pr-caida::after { content:"P"; position:absolute; inset:0; color:#fff; font:bold 7px/9px Georgia, serif; text-align:center; }
    .pr-caida-2 { left:calc(25% - 7px); animation:prCae .8s cubic-bezier(.35,0,.7,1) .48s both; }
    .pr-caida-1 { left:calc(75% - 10px); bottom:93px; width:19px; height:10px; animation:prCae .8s cubic-bezier(.35,0,.7,1) 1.92s both; }
    .pr-marcha-polvo { position:absolute; bottom:22px; width:10px; height:10px; border-radius:50%; background:#d6d3d1; opacity:0; animation:prPolvo .6s ease-out forwards; }
    .pr-marcha-bocadillo { position:absolute; bottom:0; left:50%; transform:translateX(-50%); max-width:calc(100% - 24px); overflow:hidden; text-overflow:ellipsis; background:#1B2636; color:#fff; font-size:12px; font-weight:700; padding:5px 12px; border-radius:14px; white-space:nowrap; box-shadow:0 6px 16px -4px rgba(0,0,0,.35); animation:prBocadillo 2.4s ease forwards; }
    .rufo-chip.pr-pulsado, .volver-rufo.pr-pulsado { animation:prPulsado .35s ease; }
    @keyframes prCorre { 0% { left:-80px; } 20% { left:calc(25% - 30px); animation-timing-function:ease-out; } 42% { left:calc(50% - 30px); } 62% { left:calc(50% - 30px); animation-timing-function:ease-in; } 80% { left:calc(75% - 30px); } 100% { left:calc(100% + 80px); } }
    @keyframes prSalta { from { transform:translateY(0) rotate(-6deg); } to { transform:translateY(-10px) rotate(6deg); } }
    @keyframes prSeVa { to { opacity:0; } }
    @keyframes prCae { 0% { opacity:0; transform:none; } 1% { opacity:1; } 25% { opacity:1; transform:translate(-5px, -12px) rotate(-35deg); } 75% { transform:translate(-16px, 80px) rotate(-165deg); } 86% { transform:translate(-18px, 73px) rotate(-178deg); } 100% { opacity:1; transform:translate(-20px, 80px) rotate(-180deg); } }
    @keyframes prSudorI { 0% { opacity:.95; transform:translate(0, 0) rotate(20deg); } 100% { opacity:0; transform:translate(-13px, 9px) rotate(40deg); } }
    @keyframes prSudorD { 0% { opacity:.95; transform:translate(0, 0) rotate(-20deg); } 100% { opacity:0; transform:translate(12px, -4px) rotate(-40deg); } }
    @keyframes prGotaFrente { 0% { transform:translateY(0); opacity:1; } 100% { transform:translateY(9px); opacity:0; } }
    @keyframes prPolvo { 0% { opacity:.8; transform:scale(.4); } 100% { opacity:0; transform:scale(1.8) translateY(-6px); } }
    @keyframes prBocadillo { 0% { opacity:0; transform:translate(-50%, 8px); } 12% { opacity:1; transform:translate(-50%, 0); } 85% { opacity:1; } 100% { opacity:0; } }
    @keyframes prFuera { 0%, 88% { opacity:1; } 100% { opacity:0; } }
    @keyframes prPulsado { 40% { transform:scale(.95); } 100% { transform:scale(1); } }
    @media (prefers-reduced-motion: reduce) { .pr-marcha-salto { animation:none; } .pr-marcha-polvo { display:none; } }
    @media (max-width: 340px) {
      .recordatorio-checklist-btn { flex:1 1 100%; }
    }
    .enc-escena { width:74px; height:53px; flex-shrink:0; background:#f7f3ef; border-radius:10px; overflow:visible; }
    .pr-cab{animation:prCab 3s ease-in-out infinite;transform-box:fill-box;transform-origin:50% 100%}@keyframes prCab{0%,100%{transform:rotate(2deg)}50%{transform:rotate(-3deg)}}/* buscar */.bs-pistola{animation:bsPist 2s ease-in-out infinite;transform-box:fill-box;transform-origin:0% 100%}@keyframes bsPist{0%,100%{transform:rotate(0)}50%{transform:rotate(-6deg)}}.bs-laser{animation:bsLaser 2s steps(1) infinite}@keyframes bsLaser{0%,40%{opacity:0}45%,70%{opacity:1}75%,100%{opacity:0}}.bs-ok{animation:pop 2s ease-in-out infinite;transform-box:fill-box;transform-origin:center}@keyframes pop{0%,70%{opacity:0;transform:scale(.3)}78%{opacity:1;transform:scale(1.2)}88%,100%{opacity:1;transform:scale(1)}}/* osita de Buscar (Corrección 178) */.ot-osita{cursor:pointer}.ot-ojo{transform-box:fill-box;transform-origin:50% 55%;animation:otParpadeo 4.6s ease-in-out infinite}@keyframes otParpadeo{0%,90%,100%{transform:scaleY(1)}94%{transform:scaleY(.08)}}.ot-osita.pestaneo .ot-ojo{animation:otRapido .12s ease-in-out 9}@keyframes otRapido{0%,100%{transform:scaleY(1)}50%{transform:scaleY(.06)}}.ot-osita.pestaneo .ot-pest{transform-box:fill-box;transform-origin:50% 100%;animation:otAleteo .12s ease-in-out 9}@keyframes otAleteo{0%,100%{transform:rotate(0)}50%{transform:rotate(-7deg) scaleX(1.06)}}/* calendario */.cl-hoja{animation:clHoja 3s ease-in-out infinite;transform-box:fill-box;transform-origin:50% 0%}@keyframes clHoja{0%,45%{transform:scaleY(1);opacity:1}65%{transform:scaleY(-.15);opacity:.7}70%,100%{transform:scaleY(0);opacity:0}}.cl-marca{animation:clMarca 3s ease-in-out infinite;transform-box:fill-box;transform-origin:center}@keyframes clMarca{0%,72%{opacity:0;transform:scale(.4)}82%{opacity:1;transform:scale(1.2)}90%,100%{opacity:1;transform:scale(1)}}.cl-pata{animation:clPata 3s ease-in-out infinite}@keyframes clPata{0%,100%{transform:translate(0,0)}35%,60%{transform:translate(-10px,-8px)}}/* casillas */.cs-t1,.cs-t2,.cs-t3{opacity:0;animation:csTick 3.6s ease-in-out infinite}.cs-t2{animation-delay:.6s}.cs-t3{animation-delay:1.2s}@keyframes csTick{0%,15%{opacity:0}22%,85%{opacity:1}95%,100%{opacity:0}}.cs-lapiz{animation:csLapiz 3.6s ease-in-out infinite}@keyframes csLapiz{0%,10%{transform:translate(0,0)}30%{transform:translate(0,8px)}50%{transform:translate(0,16px)}80%,100%{transform:translate(0,0)}}/* notas */.nt-l1,.nt-l2,.nt-l3{stroke-dasharray:22;stroke-dashoffset:22;animation:ntLinea 3.6s ease-in-out infinite}.nt-l2{animation-delay:.7s}.nt-l3{animation-delay:1.4s}@keyframes ntLinea{0%,10%{stroke-dashoffset:22}35%,85%{stroke-dashoffset:0}100%{stroke-dashoffset:22}}.nt-lapiz{animation:ntLapiz 3.6s ease-in-out infinite}@keyframes ntLapiz{0%{transform:translate(0,0)}25%{transform:translate(14px,0)}30%{transform:translate(0,6px)}50%{transform:translate(14px,6px)}55%{transform:translate(0,12px)}75%{transform:translate(12px,12px)}100%{transform:translate(0,0)}}/* avisos */.av-campana{animation:avCamp 1.6s ease-in-out infinite;transform-box:fill-box;transform-origin:50% 0%}@keyframes avCamp{0%,100%{transform:rotate(0)}15%{transform:rotate(16deg)}30%{transform:rotate(-14deg)}45%{transform:rotate(10deg)}60%{transform:rotate(-6deg)}75%{transform:rotate(0)}}.av-ondas{animation:avOndas 1.6s ease-out infinite}@keyframes avOndas{0%,10%{opacity:0}30%{opacity:1}70%,100%{opacity:0}}.av-excl{animation:pop 3.2s ease-in-out infinite;transform-box:fill-box;transform-origin:center}/* recuento */.rc-b1,.rc-b2,.rc-b3{transform-box:fill-box;transform-origin:50% 100%;animation:rcBarra 3s ease-in-out infinite}.rc-b2{animation-delay:.25s}.rc-b3{animation-delay:.5s}@keyframes rcBarra{0%,10%{transform:scaleY(.15)}45%,85%{transform:scaleY(1)}100%{transform:scaleY(.15)}}.rc-caja{animation:rcCaja 3s ease-in-out infinite;transform-box:fill-box;transform-origin:50% 100%}@keyframes rcCaja{0%,60%,100%{transform:rotate(0)}66%{transform:rotate(-7deg)}72%{transform:rotate(6deg)}78%{transform:rotate(-3deg)}84%{transform:rotate(0)}}@media (prefers-reduced-motion: reduce){.esc *{animation:none!important;opacity:1!important}}.rb-lupa{animation:rbLupa 2.4s ease-in-out infinite;transform-box:fill-box;transform-origin:20% 90%}@keyframes rbLupa{0%,100%{transform:translate(0,0) rotate(-6deg)}50%{transform:translate(-16px,0) rotate(6deg)}}.rb-cabeza{animation:rbCabeza 2.4s ease-in-out infinite;transform-box:fill-box;transform-origin:50% 100%}@keyframes rbCabeza{0%,100%{transform:rotate(3deg)}50%{transform:rotate(-4deg)}}.rb-papel{animation:rbPapel 3.2s ease-in-out infinite;transform-box:fill-box;transform-origin:center}@keyframes rbPapel{0%,15%{transform:translate(0,0) rotate(0);opacity:1}50%{transform:translate(43px,-50px) rotate(-18deg);opacity:1}85%{transform:translate(86px,5px) rotate(3deg);opacity:1}100%{transform:translate(86px,5px) rotate(3deg);opacity:0}}.rb-check{animation:rbCheck 3.2s ease-in-out infinite;transform-box:fill-box;transform-origin:center}@keyframes rbCheck{0%,80%{opacity:0;transform:scale(.4)}88%{opacity:1;transform:scale(1.15)}100%{opacity:1;transform:scale(1)}}
<?!= (LATERALES_UNA_PIEZA_.indexOf(accion) > -1 || accion === 'buscar' || accion === 'retornos') ? cssRetornos_() : '' ?>
<?!= accion === 'recuento' ? cssRecuento_() : '' ?>
<?!= (LATERALES_UNA_PIEZA_.indexOf(accion) > -1 || accion === 'buscar' || accion === 'pendientes') ? cssBultosCaja_() : '' ?>
<?!= (LATERALES_UNA_PIEZA_.indexOf(accion) > -1 || accion === 'buscar' || accion === 'retornos' || accion === 'cambios') ? cssCambiosRet_() : '' ?>
  </style></head><body>

  
  <svg width="0" height="0" style="position:absolute">
    <defs>
      <g id="rufo-svg"><defs><radialGradient id="rufo-grad" cx="50%" cy="35%" r="70%"><stop offset="0%" stop-color="#3A3F46"/><stop offset="100%" stop-color="#16181C"/></radialGradient><radialGradient id="rufo-pelo" cx="38%" cy="30%" r="75%"><stop offset="0%" stop-color="#CDD2D9"/><stop offset="60%" stop-color="#9AA1AA"/><stop offset="100%" stop-color="#737A84"/></radialGradient><linearGradient id="rufo-cuerpo" x1="0" y1="0" x2="1" y2="1"><stop offset="0%" stop-color="#A7ADB5"/><stop offset="100%" stop-color="#666D77"/></linearGradient><radialGradient id="rufo-gradOrejaInt" cx="40%" cy="35%" r="70%"><stop offset="0%" stop-color="#F1DEDB"/><stop offset="100%" stop-color="#D9BDB8"/></radialGradient><radialGradient id="rufo-gradMorro" cx="45%" cy="30%" r="75%"><stop offset="0%" stop-color="#FFFFFF"/><stop offset="100%" stop-color="#DDE1E6"/></radialGradient><radialGradient id="rufo-gradOjo" cx="40%" cy="38%" r="60%"><stop offset="0%" stop-color="#7A4E2E"/><stop offset="55%" stop-color="#2C1A10"/><stop offset="100%" stop-color="#0B0705"/></radialGradient><radialGradient id="rufo-gradMejilla" cx="50%" cy="35%" r="70%"><stop offset="0%" stop-color="#F29C9C" stop-opacity=".55"/><stop offset="100%" stop-color="#F29C9C" stop-opacity="0"/></radialGradient><linearGradient id="rufo-gradLazo" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#FFC04D"/><stop offset="100%" stop-color="#E2781E"/></linearGradient><linearGradient id="rufo-gorra" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#3B5373"/><stop offset="100%" stop-color="#1E2C40"/></linearGradient></defs><ellipse cx="50" cy="105" rx="30" ry="4.5" fill="#1B2636" opacity=".12"/><path d="M70 98 C88 98 99 84 97 66 C96 58 92 53 86 51" stroke="url(#rufo-cuerpo)" stroke-width="13" stroke-linecap="round" fill="none"/><path d="M70 98 C88 98 99 84 97 66 C96 58 92 53 86 51" stroke="#2E3238" stroke-width="13" stroke-dasharray="6 7" stroke-dashoffset="-3" fill="none"/><circle cx="86" cy="51" r="6.5" fill="#2A2E33"/><path d="M22 106 C17 88 22 72 50 70 C78 72 83 88 78 106 Z" fill="url(#rufo-cuerpo)"/><ellipse cx="50" cy="93" rx="15" ry="13" fill="url(#rufo-gradMorro)"/><path d="M24 105 C21 90 25 79 40 74 L42 105 Z" fill="url(#rufo-gradLazo)"/><path d="M76 105 C79 90 75 79 60 74 L58 105 Z" fill="url(#rufo-gradLazo)"/><path d="M23 93 L41 92 L41 97 L23 98 Z" fill="#EEF2F6"/><path d="M77 93 L59 92 L59 97 L77 98 Z" fill="#EEF2F6"/><ellipse cx="37" cy="105" rx="9" ry="4" fill="#2A2E33"/><ellipse cx="63" cy="105" rx="9" ry="4" fill="#2A2E33"/><path d="M14 34 C8 18 13 5 25 3 C35 6 38 17 35 27 Z" fill="#4A5059"/><path d="M18 28 C15 18 18 10 25 9 C31 11 32 18 30 24 Z" fill="url(#rufo-gradOrejaInt)"/><path d="M86 34 C92 18 87 5 75 3 C65 6 62 17 65 27 Z" fill="#4A5059"/><path d="M82 28 C85 18 82 10 75 9 C69 11 68 18 70 24 Z" fill="url(#rufo-gradOrejaInt)"/><path d="M14 34 C8 18 13 5 25 3" stroke="#F1F3F5" stroke-width="1.6" fill="none"/><path d="M86 34 C92 18 87 5 75 3" stroke="#F1F3F5" stroke-width="1.6" fill="none"/><path d="M17 50 C8 53 3 60 1 68 C9 64 14 65 20 63 Z" fill="url(#rufo-gradMorro)"/><path d="M83 50 C92 53 97 60 99 68 C91 64 86 65 80 63 Z" fill="url(#rufo-gradMorro)"/><path d="M15 47 C14 24 31 10 50 10 C69 10 86 24 85 47 C84 66 69 79 50 79 C31 79 16 66 15 47 Z" fill="url(#rufo-pelo)"/><path d="M50 17 C47 23 47 29 50 35 C53 29 53 23 50 17 Z" fill="#3E434A" opacity=".6"/><g><ellipse cx="34" cy="26" rx="10" ry="3.8" transform="rotate(-12 34 26)" fill="#F4F5F7"/><ellipse cx="66" cy="26" rx="10" ry="3.8" transform="rotate(12 66 26)" fill="#F4F5F7"/></g><path d="M17 42 C18 30 30 27 40 32 C45 34 47 36 50 36 C53 36 55 34 60 32 C70 27 82 30 83 42 C82 51 72 54 62 51 C57 49 54 47 50 47 C46 47 43 49 38 51 C28 54 18 51 17 42 Z" fill="url(#rufo-grad)"/><g><circle cx="34" cy="39" r="7.6" fill="url(#rufo-gradOjo)" stroke="#5B636E" stroke-width=".8"/><circle cx="66" cy="39" r="7.6" fill="url(#rufo-gradOjo)" stroke="#5B636E" stroke-width=".8"/><circle cx="31.4" cy="36.1" r="2.8" fill="#fff" opacity=".95"/><circle cx="63.4" cy="36.1" r="2.8" fill="#fff" opacity=".95"/><circle cx="36.6" cy="42" r="1.1" fill="#fff" opacity=".6"/><circle cx="68.6" cy="42" r="1.1" fill="#fff" opacity=".6"/></g><path d="M34 66 C35 54 42 49 50 49 C58 49 65 54 66 66 C63 74 57 77 50 77 C43 77 37 74 34 66 Z" fill="url(#rufo-gradMorro)"/><ellipse cx="50" cy="55" rx="6" ry="4.2" fill="#15171A"/><ellipse cx="48" cy="53.6" rx="1.8" ry="1" fill="#fff" opacity=".55"/><path d="M50 59 L50 61.5" stroke="#2B2E33" stroke-width="1.6" stroke-linecap="round"/><path id="rufo-boca-cerrada" d="M42 61 Q50 68 58 61" stroke="#2B2E33" stroke-width="2.6" fill="none" stroke-linecap="round"/><ellipse id="rufo-boca-abierta" cx="50" cy="64" rx="6.5" ry="5.8" fill="#3a1c14" style="display:none;"/><ellipse id="rufo-lengua" cx="50" cy="67" rx="3.2" ry="2.1" fill="#E07A1F" opacity=".8" style="display:none;"/><circle cx="24" cy="58" r="6" fill="url(#rufo-gradMejilla)"/><circle cx="76" cy="58" r="6" fill="url(#rufo-gradMejilla)"/><path d="M38 73 L62 73 L50 83 Z" fill="url(#rufo-gradLazo)"/><circle cx="50" cy="74.5" r="2.6" fill="#B85F12"/><path d="M28 23 C30 11 40 6 51 6 C62 7 71 13 72 23 Z" fill="url(#rufo-gorra)"/><path d="M24 23 Q50 17 76 24 Q51 28 24 23 Z" fill="#F5A524"/><circle cx="51" cy="13.5" r="4.6" fill="#F5A524"/><g fill="#1E2C40"><ellipse cx="51" cy="15" rx="1.9" ry="1.5"/><circle cx="48.6" cy="12.6" r=".8"/><circle cx="50.2" cy="11.5" r=".8"/><circle cx="51.9" cy="11.5" r=".8"/><circle cx="53.5" cy="12.6" r=".8"/></g></g>
      
      <linearGradient id="esc-mesa-g" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#4a4a4a"/><stop offset="100%" stop-color="#262626"/></linearGradient>
      <g id="esc-mesa"><rect x="8" y="70" width="124" height="9" rx="3" fill="url(#esc-mesa-g)"/><rect x="8" y="70" width="124" height="2" rx="1" fill="#E07A1F"/><rect x="16" y="79" width="6" height="18" rx="2" fill="#262626"/><rect x="118" y="79" width="6" height="18" rx="2" fill="#262626"/></g>
    </defs>
  </svg>

  
  <div class="cm-overlay" id="cm-overlay" hidden>
    <div class="cm-card">
      <div class="cm-head">
        <button class="cm-head-close" onclick="cerrarCajaManipulada()">✕</button>
        <div class="cm-head-avatar-wrap">
          <svg class="cm-head-avatar" viewBox="0 0 100 128" style="overflow:visible">
            <use href="#rufo-svg"/>
            <g class="cm-gafas">
              <line class="cm-gafas-puente" x1="41.5" y1="39" x2="58.5" y2="39"/>
              <circle class="cm-gafas-lente" cx="34" cy="39" r="9.5"/>
              <circle class="cm-gafas-lente" cx="66" cy="39" r="9.5"/>
              <path class="cm-gafas-brillo" d="M30 35 Q33 33 36.5 34.5 Q32.5 34 30 35 Z"/>
              <path class="cm-gafas-brillo" d="M62 35 Q65 33 68.5 34.5 Q64.5 34 62 35 Z"/>
            </g>
            <rect class="cm-libreta" x="30" y="98" width="40" height="27" rx="3"/>
            <line class="cm-renglon cm-renglon-1" x1="36" y1="106" x2="53" y2="106"/>
            <line class="cm-renglon cm-renglon-2" x1="36" y1="112.5" x2="58" y2="112.5"/>
            <line class="cm-renglon cm-renglon-3" x1="36" y1="119" x2="49" y2="119"/>
            <path class="cm-brazo" d="M22 80 Q10 90 14 103 Q16 111 26 110 Q22 100 25 88 Z"/>
            <ellipse class="cm-mano" cx="27" cy="107" rx="6.5" ry="6"/>
            <path class="cm-brazo" d="M78 80 Q90 89 85 101 Q82 109 73 107 Q78 98 75 87 Z"/>
            <g class="cm-lapiz-grande">
              <ellipse class="cm-mano" cx="73" cy="104" rx="6.5" ry="6"/>
              <rect x="66" y="88" width="5" height="24" rx="1.8" fill="#E07A1F" transform="rotate(24 68.5 100)"/>
              <path d="M64.3 110 L61 115.8 L67 114.2 Z" fill="#3a1c14" transform="rotate(24 68.5 100)"/>
            </g>
          </svg>
        </div>
        <div class="cm-head-info">
          <h3 class="cm-head-titulo" id="cm-head-titulo"></h3>
          <p class="cm-head-sub" id="cm-head-sub"></p>
          <div class="cm-head-ref" id="cm-head-ref">📦 Bulto manipulado</div>
        </div>
        <svg class="cm-swoosh" viewBox="0 0 340 14" preserveAspectRatio="none">
          <path d="M0 8 C 60 -2, 120 16, 180 6 C 240 -4, 290 14, 340 5" stroke="#E07A1F" stroke-width="3" fill="none" stroke-linecap="round"/>
        </svg>
      </div>
      <div class="cm-body" id="cm-body"></div>
    </div>
  </div>

  
  <div class="version-tag" aria-hidden="true">v0.9.164</div>

  
  <? if (accion === 'escribir' || accion === 'vernotas' || accion === 'gestionar' || accion === 'notas') { ?>
  <div id="lock-screen" class="lock-box">
    <div class="lock-icon-oso">
      <svg viewBox="0 0 100 100"><use href="#rufo-svg"/></svg>
      <span class="lock-badge">🔒</span>
    </div>
    <h2>Seguridad</h2>
    <p class="sub">Introduce tu clave para acceder</p>
    <div class="pwd-wrap">
      <input type="password" id="pwd" placeholder="Contraseña..." onkeypress="if(event.keyCode===13) validar()">
      <span class="eye" onclick="togglePwd()">👁️</span>
    </div>
    <div id="err" data-default="❌ Contraseña mal escrita, inténtelo de nuevo.">❌ Contraseña mal escrita, inténtelo de nuevo.</div>
    <button id="btn-val" class="btn btn-block" onclick="validar()">Desbloquear Panel</button>
  </div>
  <? } ?>

  <div id="app-content" style="<?= (accion === 'escribir' || accion === 'vernotas' || accion === 'gestionar' || accion === 'notas') ? 'display:none;' : '' ?>"></div>

  
  <div id="recordatorio-checklist" class="recordatorio-checklist" hidden>
    <span class="recordatorio-checklist-icono">🐻</span>
    <div class="recordatorio-checklist-texto" id="recordatorio-checklist-texto"><strong>Recordatorio:</strong> envía el checklist diario</div>
    <button class="recordatorio-checklist-btn" onclick="irAChecklistDesdeRecordatorio()">📅 Ir al Calendario</button>
    <button class="recordatorio-checklist-cerrar" onclick="cerrarRecordatorioChecklist()" title="Cerrar">✕</button>
  </div>

  
  <div id="aviso-toast" class="aviso-toast" hidden>
    <span class="aviso-toast-icono">🐻</span>
    <div class="aviso-toast-texto">
      <strong id="aviso-toast-titulo"></strong> <span id="aviso-toast-mensaje"></span>
      <span class="aviso-toast-donde" id="aviso-toast-donde"></span>
    </div>
    <button class="aviso-toast-btn" onclick="irAAvisosDesdeToast()">🔔 Ver avisos</button>
    <button class="aviso-toast-cerrar" onclick="cerrarAvisoToast()" title="Cerrar">✕</button>
  </div>

  <script>
    var accionActual = "<?= accion ?>";
    var htmlInicial = <?!= jsonSeguroParaScript_(htmlInicial) ?>;
    var CARGA_DOC_ = <?!= (typeof cargaDoc !== 'undefined' && cargaDoc) ? JSON.stringify(String(cargaDoc).replace(/[^0-9]/g, '')) : '""' ?>;

    function togglePwd() {
      var p = document.getElementById('pwd');
      p.type = p.type === 'password' ? 'text' : 'password';
    }

    function validar() {
      var p = document.getElementById('pwd').value; if(!p) return;
      var errEl = document.getElementById('err');
      errEl.classList.remove('show');
      document.getElementById('btn-val').innerText = 'Comprobando...';
      document.getElementById('btn-val').disabled = true;

      google.script.run.withSuccessHandler(function(res) {
        if(res.error) {
          errEl.innerHTML = errEl.getAttribute('data-default');
          errEl.classList.remove('show'); void errEl.offsetWidth; errEl.classList.add('show');
          document.getElementById('btn-val').innerText = 'Desbloquear Panel';
          document.getElementById('btn-val').disabled = false;
          document.getElementById('pwd').value = '';
          document.getElementById('pwd').focus();
        } else {
          document.getElementById('lock-screen').style.display = 'none';
          document.getElementById('app-content').style.display = 'block';
          document.getElementById('app-content').innerHTML = res.html;
        }
      }).withFailureHandler(function(err) {
        errEl.innerHTML = '⚠️ Error al abrir el panel: ' + (err && err.message ? err.message : 'inténtalo de nuevo.');
        errEl.classList.remove('show'); void errEl.offsetWidth; errEl.classList.add('show');
        document.getElementById('btn-val').innerText = 'Desbloquear Panel';
        document.getElementById('btn-val').disabled = false;
      }).enrutadorApp(p, accionActual);
    }

    function guardarNota() {
      var txt = document.getElementById('texto-nota').value;
      if (txt.trim()==='') return alert('La nota está vacía');
      document.getElementById('btn-guardar-nota').innerText = 'Guardando...';
      document.getElementById('btn-guardar-nota').disabled = true;
      google.script.run.withSuccessHandler(function(){ google.script.host.close(); }).procesarNuevaNota(txt);
    }

    function borrarNota(fila) {
      if (!confirm("¿Seguro que deseas eliminar definitivamente esta nota?")) return;
      document.getElementById('fila-'+fila).style.opacity = '0.4';
      google.script.run.withSuccessHandler(function(){ document.getElementById('fila-'+fila).style.display = 'none'; }).eliminarNota(fila);
    }

    function iniciarEdicionGestionar(fila) {
      document.getElementById('texto-g-'+fila).style.display = 'none';
      document.getElementById('edit-g-'+fila).style.display = 'block';
      document.getElementById('acciones-g-'+fila).style.display = 'none';
      document.getElementById('acciones-edit-g-'+fila).style.display = 'inline';
    }
    function cancelarEdicionGestionar(fila) {
      document.getElementById('texto-g-'+fila).style.display = 'inline';
      document.getElementById('edit-g-'+fila).style.display = 'none';
      document.getElementById('acciones-g-'+fila).style.display = 'inline';
      document.getElementById('acciones-edit-g-'+fila).style.display = 'none';
    }
    function guardarEdicionGestionar(fila) {
      var nuevo = document.getElementById('edit-g-'+fila).value;
      if (nuevo.trim()==='') return alert('La nota no puede quedar vacía');
      google.script.run.withSuccessHandler(function(){
        document.getElementById('texto-g-'+fila).innerText = nuevo;
        cancelarEdicionGestionar(fila);
      }).editarNota(fila, nuevo);
    }

    function guardarNotaUnificada() {
      var campo = document.getElementById('texto-nota-unif');
      var txt = campo.value;
      if (txt.trim()==='') return alert('La nota está vacía');
      var tipoEl = document.querySelector('.nt-tipo.on'), pedEl = document.getElementById('nt-pedido');
      var pedido = (tipoEl && tipoEl.getAttribute('data-tipo') === 'pedido' && pedEl) ? pedEl.value.trim() : '';
      if (tipoEl && tipoEl.getAttribute('data-tipo') === 'pedido' && !pedido) { if (pedEl) pedEl.focus(); return alert('Falta el nº de pedido'); }
      var btn = document.getElementById('btn-guardar-nota-unif');
      btn.innerText = 'Guardando...'; btn.disabled = true;
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(function(err){ btn.innerText = 'Guardar nota'; btn.disabled = false; alert('No se pudo guardar: ' + (err && err.message ? err.message : err)); }).guardarYRefrescarNotas(txt, pedido);
    }
    function ntElegirTipo(b) {
      document.querySelectorAll('.nt-tipo').forEach(function(x) { x.classList.toggle('on', x === b); });
      var ped = document.getElementById('nt-pedido');
      if (ped) { ped.style.display = b.getAttribute('data-tipo') === 'pedido' ? 'block' : 'none'; if (ped.style.display === 'block') ped.focus(); }
    }
    function ntPestana(b) {
      document.querySelectorAll('.nt-pes').forEach(function(x) { x.classList.toggle('on', x === b); });
      var p = b.getAttribute('data-p');
      document.getElementById('nt-lista-pend').style.display = p === 'pend' ? '' : 'none';
      document.getElementById('nt-lista-hechas').style.display = p === 'hechas' ? '' : 'none';
    }
    function resolverNotaUnificada(fila, id, resolver) {
      var t = document.getElementById('fila-' + fila); if (t) t.style.opacity = '0.5';
      var enResueltas = !resolver;
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
        if (enResueltas) { var b = document.querySelector('.nt-pes[data-p="hechas"]'); if (b) ntPestana(b); }
      }).withFailureHandler(function(err){ if (t) t.style.opacity = ''; alert('No se pudo guardar: ' + (err && err.message ? err.message : err)); }).resolverYRefrescarNotas(fila, id, resolver);
    }
    function borrarNotaUnificada(fila) {
      if (!confirm("¿Seguro que deseas eliminar definitivamente esta nota?")) return;
      document.getElementById('fila-'+fila).style.opacity = '0.4';
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
      }).borrarNotaYRefrescarNotas(fila);
    }
    function iniciarEdicionNota(fila) {
      document.getElementById('texto-'+fila).style.display = 'none';
      document.getElementById('edit-'+fila).style.display = 'block';
      document.getElementById('acciones-'+fila).style.display = 'none';
      document.getElementById('acciones-edit-'+fila).style.display = 'flex';
    }
    function cancelarEdicionNota(fila) {
      document.getElementById('texto-'+fila).style.display = 'block';
      document.getElementById('edit-'+fila).style.display = 'none';
      document.getElementById('acciones-'+fila).style.display = 'flex';
      document.getElementById('acciones-edit-'+fila).style.display = 'none';
    }
    function guardarEdicionNota(fila) {
      var nuevo = document.getElementById('edit-'+fila).value;
      if (nuevo.trim()==='') return alert('La nota no puede quedar vacía');
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
      }).editarNotaYRefrescarNotas(fila, nuevo);
    }

    function guardarAvisoNuevo() {
      var refCampo = document.getElementById('ref-aviso-nuevo');
      var motivoCampo = document.getElementById('motivo-aviso-nuevo');
      var ref = refCampo.value, motivo = motivoCampo.value;
      if (ref.trim()==='') return alert('Falta la referencia del pedido');
      if (motivo.trim()==='') return alert('El motivo no puede quedar vacío');
      var btn = document.getElementById('btn-guardar-aviso');
      btn.innerText = 'Guardando...'; btn.disabled = true;
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(function(err){
        btn.innerText = 'Añadir aviso'; btn.disabled = false;
        alert('No se pudo guardar el aviso: ' + (err && err.message ? err.message : err));
      }).guardarYRefrescarAvisos(ref, motivo);
    }
    function borrarAvisoUnificado(fila) {
      if (!confirm("¿Seguro que deseas eliminar definitivamente este aviso?")) return;
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(function(err){ alert('No se pudo eliminar: ' + (err && err.message ? err.message : err)); }).borrarAvisoYRefrescarAvisos(fila);
    }
    function iniciarEdicionAviso(fila) {
      document.getElementById('motivo-'+fila).style.display = 'none';
      document.getElementById('edit-aviso-'+fila).style.display = 'block';
      document.getElementById('acciones-aviso-'+fila).style.display = 'none';
      document.getElementById('acciones-edit-aviso-'+fila).style.display = 'flex';
    }
    function cancelarEdicionAviso(fila) {
      document.getElementById('motivo-'+fila).style.display = 'block';
      document.getElementById('edit-aviso-'+fila).style.display = 'none';
      document.getElementById('acciones-aviso-'+fila).style.display = 'flex';
      document.getElementById('acciones-edit-aviso-'+fila).style.display = 'none';
    }
    function guardarEdicionAviso(fila) {
      var nuevo = document.getElementById('edit-aviso-'+fila).value;
      if (nuevo.trim()==='') return alert('El motivo no puede quedar vacío');
      google.script.run.withSuccessHandler(function(res){
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(function(err){ alert('No se pudo guardar: ' + (err && err.message ? err.message : err)); }).editarAvisoYRefrescarAvisos(fila, nuevo);
    }

    var calOff = 0;
    function initCalendario() { cargarCalendario(0); }
    function cargarCalendario(dir) {
      calOff += dir;
      document.getElementById('load-cal').style.display = 'block'; document.getElementById('grid-cal').style.display = 'none'; document.getElementById('cal-legend').style.display = 'none';
      document.getElementById('pendientes-wrap').style.display = 'none';
      var offPedido = calOff, cCal = cacheLeer_('cal_' + calOff);
      if (cCal && cCal.d) { try { mostrarCalendario(cCal.d); marcaCache_(cCal.t); } catch (eCc) {} }
      google.script.run.withSuccessHandler(function(d) {
        cacheGuardar_('cal_' + offPedido, d);
        if (offPedido !== calOff) return;
        quitarMarcaCache_(); mostrarCalendario(d);
      }).withFailureHandler(function() { quitarMarcaCache_(); }).obtenerDatosCalendario(calOff);
    }
    var AG_LOGOS_B64_ = ${JSON.stringify(AG_LOGOS_B64_)};
    var agBadgeMapa_ = {"VELOX":{c:"#2F6FE4",t:"VELOX"},"PAQNORTE":{c:"#0F8A7E",t:"PAQN"},"CORREOMAX":{c:"#C99700",t:"CMAX"},"RUTASUR":{c:"#D85A30",t:"RSUR"},"ATLAS":{c:"#6D4BC9",t:"ATLAS"},"PRONTO":{c:"#2E9E44",t:"PRON"},"BOLIDO":{c:"#C62E3B",t:"BOLI"},"FARO":{c:"#26354A",t:"FARO"}};
    function agBadge_(comp, compacto) {
      var claseCard = "ag-badge-card" + (compacto ? " ag-badge-sm" : "");
      var etiqueta = "<span class='ag-badge-label'>" + comp + "</span>";
      if (AG_LOGOS_B64_[comp]) return "<span class='" + claseCard + "' title='" + comp + "'><span class='ag-badge-logo'><img src='data:image/png;base64," + AG_LOGOS_B64_[comp] + "' alt='" + comp + "'></span>" + etiqueta + "</span>";
      var b = agBadgeMapa_[comp] || {c:"#64748b", t:comp.substring(0,3)};
      return "<span class='" + claseCard + " ag-badge-color' title='" + comp + "'><span class='ag-badge-logo' style='background:" + b.c + "'>" + b.t + "</span>" + etiqueta + "</span>";
    }
    function pintarEnviosChecklist_(env) {
      var el = document.getElementById('cal-envios');
      if (!el) return;
      if (!env || !env.dias) { el.hidden = true; return; }
      var letras = ['L', 'M', 'X', 'J', 'V', 'S', 'D'], nombres = ['Lunes', 'Martes', 'Miércoles', 'Jueves', 'Viernes', 'Sábado', 'Domingo'];
      function hora(t) { var f = new Date(t); return ('0' + f.getHours()).slice(-2) + ':' + ('0' + f.getMinutes()).slice(-2); }
      function fecha(t) { var f = new Date(t); return ('0' + f.getDate()).slice(-2) + '/' + ('0' + (f.getMonth() + 1)).slice(-2); }
      function titulo(x, que) { return x.estado === 'si' ? que + ': ' + (x.o ? 'correo preparado en Outlook ' : 'enviado ') + (x.tarde ? 'con retraso el ' + fecha(x.t) + ' ' : '') + 'a las ' + hora(x.t) + (x.u ? ' por ' + x.u : '') + (x.a ? ' · actualizado el ' + fecha(x.a) + ' a las ' + hora(x.a) + (x.au ? ' por ' + x.au : '') : '') + ' · pulsa para volver a enviarlo (saldrá como ACTUALIZADO)' : (x.estado === 'no' ? que + ': no se envió · pulsa para enviarlo ahora' : (x.estado === 'hoy' ? que + ': todavía sin enviar · pulsa para enviarlo' : que)); }
      function pulsable(x, accion, clase, texto, tit) { return (x.estado === 'no' || x.estado === 'hoy' || x.estado === 'si') ? '<button class="' + clase + ' ce-btn" onclick="' + accion + '" title="' + escAttr(tit) + '">' + texto + '</button>' : '<span class="' + clase + '" title="' + escAttr(tit) + '">' + texto + '</span>'; }
      function marca(x) { return x.estado === 'si' ? '✓' : (x.estado === 'no' ? '✗' : (x.estado === 'hoy' ? '…' : '')); }
      var h = '<span class="ce-t">✉️ Checklist:</span>', hoy = false;
      env.dias.forEach(function(x, i) {
        if (x.estado === 'hoy') hoy = true;
        var accion = x.estado === 'hoy' ? 'abrirPreviewChecklist(&quot;dia&quot;)' : 'abrirPreviewChecklist(&quot;dia&quot;, &quot;' + x.k + '&quot;)';
        h += pulsable(x, accion, 'ce-d ' + x.estado, letras[i] + marca(x), titulo(x, nombres[i] + ' (diario)'));
      });
      var s = env.semana || { estado: 'nada' };
      ultimosEnviosCal_ = env;
      h += '<span class="ce-sep"></span>' + pulsable(s, 'abrirPreviewChecklist(&quot;semana&quot;)', 'ce-d ' + s.estado, 'Semanal ' + marca(s), titulo(s, 'Semanal') + (s.estado === 'pronto' ? ': se envía el lunes siguiente' : ''));
      if (hoy) h += '<button class="ce-enviar" onclick="abrirPreviewChecklist(&quot;dia&quot;)">Enviar el de hoy</button>';
      if (s.estado === 'hoy') h += '<button class="ce-enviar" onclick="abrirPreviewChecklist(&quot;semana&quot;)">Enviar el semanal</button>';
      if (semanalAnteriorPendiente_()) h += '<button class="ce-enviar" onclick="abrirPreviewChecklist(&quot;semana&quot;, null, calOff - 1)" title="' + escAttr(env.anterior.estado === 'no' ? 'Semanal de la semana pasada: no se envió el lunes · pulsa para enviarlo ahora' : 'Semanal de la semana pasada (lunes a domingo): toca enviarlo hoy') + '">Enviar el semanal de la semana pasada</button>';
      el.innerHTML = h; el.hidden = false;
    }
    var ultimosEnviosCal_ = null;
    function semanalAnteriorPendiente_() {
      var a = ultimosEnviosCal_ && ultimosEnviosCal_.anterior;
      return calOff === 0 && !!a && (a.estado === 'hoy' || a.estado === 'no');
    }
    function abrirChecklistSemanal() {
      if (semanalAnteriorPendiente_()) abrirPreviewChecklist('semana', null, calOff - 1);
      else abrirPreviewChecklist('semana');
    }
    function mostrarCalendario(d) {
      document.getElementById('mes-cal').innerText = d.mes;
      var html = "";
      d.dias.forEach(function(dia) {
        var hc = dia.esHoy ? 'today' : '', ht = dia.esHoy ? dia.nombre + ' <br><small>(HOY)</small>' : dia.nombre, agH = "";
        d.compañias.forEach(function(c) {
          var v = dia.agencias[c];
          var retornosPreview = (dia.agenciasRetornos && dia.agenciasRetornos[c]) || 0;
          var op = (v===0 && retornosPreview===0) ? 'opacity:0.4;' : '', cc = v===0?'count z':'count';
          var filaHtml = "<div class='row-top'>"+agBadge_(c, true)+"<span class='"+cc+"'>"+v+"</span></div>";

          var rot = (dia.agenciasRotos && dia.agenciasRotos[c]) || 0;
          var pend = (dia.agenciasPendiente && dia.agenciasPendiente[c]) || 0;
          var retornos = (dia.agenciasRetornos && dia.agenciasRetornos[c]) || 0;
          if (v > 0 || rot > 0 || retornos > 0) {
            var clasePend = pend > 0 ? 'mini-pend-alerta' : 'mini-pend-ok';
            filaHtml += "<div class='row-sub'>" +
              (rot > 0 ? "<span class='mini-rot'>💥 " + rot + "</span>" : "") +
              (v > 0 ? "<span class='" + clasePend + "'>⏳ " + pend + "</span>" : "") +
              (retornos > 0 ? "<span class='mini-retornos'>↩️ " + retornos + "</span>" : "") +
              "</div>";
          }
          agH += "<div class='row' style='"+op+"'>" + filaHtml + "</div>";
        });
        html += "<div class='day-card'><div class='day-header "+hc+"'>"+ht+"</div><div class='day-total'>📊 Total: "+dia.total+"</div><div class='agencies'>"+agH+"</div></div>";
      });
      document.getElementById('grid-cal').innerHTML = html;
      document.getElementById('load-cal').style.display = 'none'; document.getElementById('grid-cal').style.display = 'grid'; document.getElementById('cal-legend').style.display = 'flex';
      document.getElementById('pendientes-wrap').style.display = 'block';
      pintarEnviosChecklist_(d.envios);
      try { cargarAvisoCompanias_(); } catch (eAc) {}

      var printHtml = "";
      d.dias.forEach(function(dia) {
        var filasP = "", totRoto = 0, totPend = 0, totRetornos = 0;
        d.compañias.forEach(function(c) {
          var v = dia.agencias[c] || 0;
          var rot = (dia.agenciasRotos && dia.agenciasRotos[c]) || 0;
          var pend = (dia.agenciasPendiente && dia.agenciasPendiente[c]) || 0;
          var retornos = (dia.agenciasRetornos && dia.agenciasRetornos[c]) || 0;
          totRoto += rot; totPend += pend; totRetornos += retornos;
          var pctTexto = "—", claseRot = "";
          if (rot > 0 && v > 0) {
            var pctRot = Math.round((rot / v) * 1000) / 10;
            pctTexto = (pctRot % 1 === 0 ? pctRot.toFixed(0) : pctRot.toFixed(1)) + "%";
            claseRot = "col-pct-rota";
          }
          filasP += "<tr><td>" + agBadge_(c) + "</td><td>" + v + "</td><td>" + (v > 0 ? pend : "—") + "</td><td>" + (retornos > 0 ? retornos : "—") + "</td></tr>";
        });
        var pctTotalTexto = "—", claseTotal = "";
        if (totRoto > 0 && dia.total > 0) {
          var pctTotal = Math.round((totRoto / dia.total) * 1000) / 10;
          pctTotalTexto = (pctTotal % 1 === 0 ? pctTotal.toFixed(0) : pctTotal.toFixed(1)) + "%";
          claseTotal = "col-pct-rota";
        }
        filasP += "<tr class='print-total-row'><td>TOTAL</td><td>" + dia.total + "</td><td>" + totPend + "</td><td>" + totRetornos + "</td></tr>";
        var badgeHoy = dia.esHoy ? " <span class='print-badge-hoy'>HOY</span>" : "";
        printHtml += "<table class='print-day-table" + (dia.esHoy ? " print-today" : "") + (dia.elegido ? " print-elegido" : "") + "'>" +
          "<thead>" +
            "<tr class='print-day-title'><th colspan='4'>" + dia.nombre + badgeHoy + "</th></tr>" +
            "<tr class='print-col-headers'><th>Agencia</th><th>Traídas</th><th>Pendientes por abrir</th><th>Retornos</th></tr>" +
          "</thead>" +
          "<tbody>" + filasP + "</tbody></table>";
      });

      var resumenSem = {};
      d.compañias.forEach(function(c) { resumenSem[c] = { v: 0, rot: 0, pend: 0, retornos: 0 }; });
      d.dias.forEach(function(dia) {
        d.compañias.forEach(function(c) {
          resumenSem[c].v += dia.agencias[c] || 0;
          resumenSem[c].rot += (dia.agenciasRotos && dia.agenciasRotos[c]) || 0;
          resumenSem[c].pend += (dia.agenciasPendiente && dia.agenciasPendiente[c]) || 0;
          resumenSem[c].retornos += (dia.agenciasRetornos && dia.agenciasRetornos[c]) || 0;
        });
      });
      var filasResumenSem = "", granV = 0, granRot = 0, granPend = 0, granRetornos = 0;
      d.compañias.forEach(function(c) {
        var r = resumenSem[c];
        granV += r.v; granRot += r.rot; granPend += r.pend; granRetornos += r.retornos;
        var pctTextoR = "—", claseR = "";
        if (r.rot > 0 && r.v > 0) {
          var pctR = Math.round((r.rot / r.v) * 1000) / 10;
          pctTextoR = (pctR % 1 === 0 ? pctR.toFixed(0) : pctR.toFixed(1)) + "%";
          claseR = "col-pct-rota";
        }
        filasResumenSem += "<tr><td>" + agBadge_(c) + "</td><td>" + r.v + "</td><td>" + (r.rot > 0 ? r.rot : "—") + "</td><td class='col-pct " + claseR + "'>" + pctTextoR + "</td><td>" + (r.v > 0 ? r.pend : "—") + "</td><td>" + (r.retornos > 0 ? r.retornos : "—") + "</td></tr>";
      });
      var pctGranTextoR = "—", claseGranR = "";
      if (granRot > 0 && granV > 0) {
        var pctGranR = Math.round((granRot / granV) * 1000) / 10;
        pctGranTextoR = (pctGranR % 1 === 0 ? pctGranR.toFixed(0) : pctGranR.toFixed(1)) + "%";
        claseGranR = "col-pct-rota";
      }
      filasResumenSem += "<tr class='print-total-row'><td>TOTAL SEMANA</td><td>" + granV + "</td><td>" + granRot + "</td><td class='col-pct " + claseGranR + "'>" + pctGranTextoR + "</td><td>" + granPend + "</td><td>" + granRetornos + "</td></tr>";
      printHtml += "<table class='print-day-table print-week-table'>" +
        "<thead>" +
          "<tr class='print-day-title'><th colspan='6' style='background:#E07A1F;'>📊 Resumen de la semana</th></tr>" +
          "<tr class='print-col-headers'><th>Agencia</th><th>Traídas</th><th>Roto</th><th class='col-pct'>% Rotura</th><th>Pendientes por abrir</th><th>Retornos</th></tr>" +
        "</thead>" +
        "<tbody>" + filasResumenSem + "</tbody></table>";

      document.getElementById('print-tables').innerHTML = printHtml;
      document.getElementById('print-rango').innerText = d.rangoFechas || d.mes || '—';
      var ahoraGen = new Date();
      document.getElementById('print-fecha-gen').innerText = ahoraGen.toLocaleDateString('es-ES') + ' ' + ahoraGen.toLocaleTimeString('es-ES', {hour:'2-digit', minute:'2-digit'});
      document.getElementById('print-usuario').innerText = d.generadoPor || '—';
    }

    function abrirPendientesEnLateral() {
      var btn = document.getElementById('btn-pendientes');
      btn.innerText = '⏳ Abriendo...'; btn.disabled = true;
      google.script.run.withSuccessHandler(function() {
        google.script.host.close();
      }).abrirMotorLateral('pendientes', '📋 Casillas Pendientes');
    }

    var calOffPend = 0, pendTocado = false, pendPide = 0;
    function initPendientesLateral() { cargarPendientesLateral(0); }
    function pendPintar_(res) {
      var mes = document.getElementById('mes-pend'), load = document.getElementById('load-pend'), panel = document.getElementById('panel-pend-lateral');
      if (!mes || !panel) return;
      var filtro = document.getElementById('pend-filtro'), valFiltro = filtro ? filtro.value : '';
      mes.innerText = res.mes || '';
      if (load) load.style.display = 'none';
      panel.style.display = 'block';
      mostrarPendientes(res);
      var f2 = document.getElementById('pend-filtro');
      if (valFiltro && f2 && [].some.call(f2.options, function(o) { return o.value === valFiltro; })) { f2.value = valFiltro; filtrarPendientes(); }
      pendTocado = false;
      if (!panel.getAttribute('data-oye')) { panel.setAttribute('data-oye', '1'); panel.addEventListener('click', function() { pendTocado = true; }, true); }
    }
    function cargarPendientesLateral(dir) {
      calOffPend += dir;
      var miOff = calOffPend, ficha = ++pendPide, k = 'pend_' + miOff, cP = cacheLeer_(k);
      var avisoV = document.getElementById('pend-nuevo'); if (avisoV) avisoV.remove();
      if (cP && cP.d) { pendPintar_(cP.d); marcaCache_(cP.t); }
      else {
        quitarMarcaCache_();
        var ld = document.getElementById('load-pend'), pn = document.getElementById('panel-pend-lateral');
        if (ld) ld.style.display = 'block'; if (pn) pn.style.display = 'none';
      }
      google.script.run.withSuccessHandler(function(res) {
        if (ficha !== pendPide || !document.getElementById('panel-pend-lateral')) return;
        quitarMarcaCache_();
        if (!res) return;
        cacheGuardar_(k, res);
        if (cP && cP.d && JSON.stringify(cP.d) === JSON.stringify(res)) return; // igual que lo enseñado
        if (cP && cP.d && pendTocado) {
          var pn2 = document.getElementById('panel-pend-lateral');
          pn2.insertAdjacentHTML('beforebegin', '<button id="pend-nuevo" class="pend-nuevo" type="button">🔄 Hay cambios en la lista · pulsa para verla al día</button>');
          document.getElementById('pend-nuevo').onclick = function() { this.remove(); pendPintar_(res); };
          return;
        }
        pendPintar_(res);
      }).withFailureHandler(function(err) {
        if (ficha !== pendPide) return;
        quitarMarcaCache_();
        if (cP && cP.d) return; // se queda lo guardado
        var ld = document.getElementById('load-pend');
        if (ld) ld.innerHTML = '⚠️ No se pudo cargar: ' + escAttr(err && err.message ? err.message : err);
      }).obtenerPendientesCalendario(miOff);
    }

    function avisoPanel_(texto) {
      try {
        var el = document.getElementById('aviso-panel');
        if (!el) { el = document.createElement('div'); el.id = 'aviso-panel'; el.setAttribute('role', 'alert'); el.onclick = function() { el.classList.remove('ver'); }; document.body.appendChild(el); }
        el.textContent = String(texto === undefined || texto === null ? '' : texto);
        void el.offsetWidth; el.classList.add('ver');
        clearTimeout(el.avisoT); el.avisoT = setTimeout(function() { el.classList.remove('ver'); }, 8000);
      } catch (e) {}
    }
    try { window.alert = function(t) { avisoPanel_(t); }; } catch (eAl) {}
    function escAttr(s) {
      return (s || '').toString()
        .replace(/&/g, '&amp;')
        .replace(/"/g, '&quot;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');
    }

    function mostrarPendientes(res) {
      var panel = document.getElementById('panel-pend-lateral');
      if (!document.getElementById('pend-nuevo-css')) { var st = document.createElement('style'); st.id = 'pend-nuevo-css'; st.textContent = '.pend-nuevo{display:block;width:100%;margin:0 0 10px;padding:9px 10px;border:1.5px solid #f59e0b;background:#fffbeb;color:#92400e;border-radius:10px;font:inherit;font-size:12.5px;font-weight:700;cursor:pointer}'; document.head.appendChild(st); }
      var lista = res.pendientes || [];
      if (lista.length === 0) {
        panel.innerHTML = "<div class='empty-inline'>✅ No queda ninguna casilla vacía esta semana.</div>";
        return;
      }
      var visitados = JSON.parse(localStorage.getItem('clicsPendientes') || '[]');
      var selectHtml = '<div class="search-row" style="margin-bottom:12px;"><select id="pend-filtro" class="field field-select" onchange="filtrarPendientes()">' +
        '<option value="">Todas las compañías (' + res.total + ')</option>' +
        res.compañias.map(function(c) {
          var n = lista.filter(function(p){ return p.agencia.indexOf(c) !== -1; }).length;
          return n > 0 ? '<option value="' + c + '">' + c + ' (' + n + ')</option>' : '';
        }).join('') + '</select></div>';
      var itemsHtml = lista.map(function(p) {
        var idUnico = 'pend-' + p.hoja + '-' + p.fila;
        var esVisitado = visitados.indexOf(idUnico) > -1 ? 'visitado' : '';
        var txtBtn = esVisitado ? '✔️ Revisado' : '➡️ Ir a celda';
        var refTxt = p.referencia ? p.referencia : '(sin referencia en columna C)';
        var html = '<div id="' + idUnico + '" class="res-item pend-row ' + esVisitado + '" data-ag="' + p.agencia + '" data-ref="' + escAttr(p.referencia) + '">';
        html += '<div class="res-info"><div class="res-hoja">📄 ' + p.hoja + ' (Fila ' + p.fila + ') <span class="badge-dia">' + p.dia + '</span></div>';
        html += '<div><strong>' + p.agencia + '</strong></div>';
        html += '<div class="res-notaL sin-dato">📋 Columna L: (vacío)</div>';
        html += '<div class="res-ref">🔖 ' + refTxt + '</div></div>';
        html += '<div class="res-acciones-row">';
        html += '<button class="btn-ir" onclick="saltarAPendiente(\\'' + p.hoja + '\\', ' + p.fila + ', ' + p.col + ', \\'' + idUnico + '\\')">' + txtBtn + '</button>';
        html += '<button class="btn-rapido-toggle" id="toggle-' + idUnico + '" onclick="toggleAccionesRapidas(\\'' + idUnico + '\\')">🛠️ Rellenar</button>';
        html += '</div>';
        html += '<div class="res-rapido" id="rapido-' + idUnico + '">';
        html += '<div class="res-rapido-grupo"><div class="res-rapido-label">Rotura</div><div class="res-rapido-btns">';
        html += '<button class="qbtn qbtn-si" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + p.hoja + '\\', ' + p.fila + ', \\'rotura\\', \\'SI\\', this)">Sí</button>';
        html += '<button class="qbtn qbtn-no" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + p.hoja + '\\', ' + p.fila + ', \\'rotura\\', \\'NO\\', this)">No</button>';
        html += '</div></div>';
        html += '<div class="res-rapido-grupo"><div class="res-rapido-label">Verificación (Col. L)</div><div class="res-rapido-btns">';
        html += '<button class="qbtn qbtn-ok" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + p.hoja + '\\', ' + p.fila + ', \\'verificacion\\', \\'OK\\', this)">OK</button>';
        html += '<button class="qbtn qbtn-reclamar" data-cm-ref="' + escAttr(p.referencia) + '" data-cm-age="' + escAttr(p.agencia) + '" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + p.hoja + '\\', ' + p.fila + ', \\'verificacion\\', \\'RECLAMAR\\', this); abrirCajaManipulada(\\'' + idUnico + '\\', this)">RECLAMAR</button>';
        html += '<button class="qbtn qbtn-reclamar-pend" data-cm-ref="' + escAttr(p.referencia) + '" data-cm-age="' + escAttr(p.agencia) + '" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + p.hoja + '\\', ' + p.fila + ', \\'verificacion\\', \\'RECLAMAR/PENDIENTE\\', this); abrirCajaManipulada(\\'' + idUnico + '\\', this)">RECL./PEND.</button>';
        html += '</div></div>';
        html += '</div>';
        html += '</div>';
        return html;
      }).join('');
      var avisoTruncado = res.truncado ? "<div class='empty-inline'>Mostrando las primeras " + lista.length + " de " + res.total + ". Ve marcando estas y vuelve a pulsar el botón para ver el resto.</div>" : "";
      panel.innerHTML = selectHtml + itemsHtml + avisoTruncado;
    }

    function filtrarPendientes() {
      var val = document.getElementById('pend-filtro').value;
      document.querySelectorAll('.pend-row').forEach(function(row) {
        var ag = row.getAttribute('data-ag') || '';
        row.style.display = (val === '' || ag.indexOf(val) !== -1) ? '' : 'none';
      });
    }

    function saltarAPendiente(hoja, fila, col, idHtml) {
      var filaEl = document.getElementById(idHtml);
      var btn = filaEl.querySelector('.btn-ir');
      filaEl.classList.add('visitado');
      var visitados = JSON.parse(localStorage.getItem('clicsPendientes') || '[]');
      if (visitados.indexOf(idHtml) === -1) { visitados.push(idHtml); localStorage.setItem('clicsPendientes', JSON.stringify(visitados)); }
      var referencia = filaEl.getAttribute('data-ref') || '';
      copiarAlPortapapeles(referencia, function(ok) {
        btn.innerText = (ok && referencia) ? '📋 Ref. copiada ✔️' : '✔️ Revisado';
      });
      google.script.run.activarCeldaEnHoja(hoja, fila, col);
    }

    function copiarAlPortapapeles(texto, cb) {
      if (!texto) { cb(false); return; }
      if (navigator.clipboard && navigator.clipboard.writeText) {
        navigator.clipboard.writeText(texto).then(function() {
          cb(true);
        }).catch(function() {
          cb(copiarAlPortapapelesFallback(texto));
        });
      } else {
        cb(copiarAlPortapapelesFallback(texto));
      }
    }
    function copiarAlPortapapelesFallback(texto) {
      try {
        var ta = document.createElement('textarea');
        ta.value = texto;
        ta.style.position = 'fixed';
        ta.style.left = '-9999px';
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        var ok = document.execCommand('copy');
        document.body.removeChild(ta);
        return ok;
      } catch (err) {
        return false;
      }
    }

    var checklistState = { modo: null, offset: 0, asunto: '', correoEsPrueba: false, correoDestino: '', correoCC: '' };

    function mostrarPasoChecklist(paso) {
      ['preview', 'confirm', 'result'].forEach(function(p) {
        var esActivo = (p === paso);
        document.getElementById('checklist-step-' + p).style.display = esActivo ? 'block' : 'none';
        document.getElementById('checklist-foot-' + p).style.display = esActivo ? 'flex' : 'none';
      });
    }

    function abrirPreviewChecklist(modo, fechaDia, offsetSemana, fichaRevisada) {
      var ficha = fichaRevisada || ++chkFicha;
      if (!fichaRevisada) { checklistState.revisado = 0; chkBotonCorreo_(false); }
      checklistState.modo = modo;
      checklistState.fechaDia = (modo === 'dia' && typeof fechaDia === 'string') ? fechaDia : '';
      checklistState.offset = typeof offsetSemana === 'number' ? offsetSemana : calOff;
      checklistState.asunto = '';
      checklistState.htmlOutlook = '';
      var infoOutlook = document.getElementById('checklist-outlook-info'); if (infoOutlook) { infoOutlook.style.display = 'none'; infoOutlook.innerHTML = ''; }
      var msgRes = document.getElementById('checklist-result-msg'); if (msgRes) msgRes.style.display = '';
      document.getElementById('checklist-overlay').hidden = false;
      document.getElementById('checklist-card-titulo').innerText = 'Generando la vista previa…';
      document.getElementById('checklist-doc-preview').innerHTML = '<div class="loading-row"><span class="spinner"></span>Generando la vista previa del PDF…</div>';
      document.getElementById('checklist-warn').style.display = 'none';
      mostrarPasoChecklist('preview');

      google.script.run.withSuccessHandler(function(d) {
        if (ficha !== chkFicha) return;
        mostrarCalendario(d);
        document.getElementById('print-titulo').innerText = modo === 'dia'
          ? 'Checklist Diario de Agencias'
          : 'Resumen Semanal de Agencias';
        document.getElementById('print-subtitulo').innerText = modo === 'dia'
          ? (d.diaElegidoTexto ? 'Traídas y pendientes del ' + d.diaElegidoTexto.toLowerCase() + (d.reenvio ? ' (actualizado)' : ' (enviado con retraso)') : 'Traídas y pendientes de hoy')
          : ('Por agencia · Semana del ' + (d.rangoFechas || '—'));
        if (modo === 'dia' && d.fechaDiaTexto) document.getElementById('print-rango').innerText = d.fechaDiaTexto;
        document.querySelectorAll('.print-day-table').forEach(function(t) {
          t.classList.remove('print-hide');
          var esResumenSemana = t.classList.contains('print-week-table');
          if (modo === 'dia') {
            if (!t.classList.contains(d.diaElegidoTexto ? 'print-elegido' : 'print-today')) t.classList.add('print-hide');
          } else if (!esResumenSemana) {
            t.classList.add('print-hide');
          }
        });
        document.getElementById('print-tables').classList.toggle('print-tables-dia', modo === 'dia');

        checklistState.asunto = d.checklistAsunto;
        checklistState.correoEsPrueba = !!d.correoEsPrueba;
        checklistState.correoDestino = d.correoDestino || '';
        checklistState.correoCC = d.correoCC || '';
        document.getElementById('checklist-card-titulo').innerText = d.checklistAsunto;
        document.getElementById('checklist-warn').style.display = d.correoConfigurado ? 'none' : 'block';
        document.getElementById('checklist-doc-preview').innerHTML = d.checklistHtml;
        checklistState.htmlOutlook = d.checklistHtmlOutlook || ''; checklistState.textoOutlook = d.checklistTexto || ''; checklistState.para = d.correoDestino || '';
        checklistState.yo = String(d.generadoPor || '').indexOf('@') > -1 ? String(d.generadoPor) : '';
        avisoYaPreparadoChecklist_(d.envios, modo, checklistState.fechaDia);
        if (d.errorDia) { document.getElementById('checklist-card-titulo').innerText = '⚠️ ' + d.errorDia; checklistState.fechaDia = ''; }
        chkTrasVistaPrevia_(ficha);
      }).datosParaChecklist(checklistState.offset, modo, checklistState.fechaDia, true);
    }
    var chkFicha = 0;
    function chkBotonCorreo_(listo) {
      var b = document.getElementById('btn-checklist-outlook'); if (!b) return;
      b.disabled = !listo; b.innerText = listo ? '📧 Preparar correo en Outlook' : '⏳ Comprobando que no falte nada…';
    }
    function chkTrasVistaPrevia_(ficha) {
      if (ficha !== chkFicha) return;
      if (checklistState.revisado === ficha) { chkBotonCorreo_(true); return; } // ya revisada (vista previa rehecha)
      checklistState.revisado = ficha;
      google.script.run.withSuccessHandler(function(r) {
        if (ficha !== chkFicha) return;
        if (r && r.cambios) abrirPreviewChecklist(checklistState.modo, checklistState.fechaDia || undefined, checklistState.offset, ficha);
        else chkBotonCorreo_(true);
      }).withFailureHandler(function() { if (ficha === chkFicha) chkBotonCorreo_(true); }).revisarParaChecklist();
    }

    function imprimirDesdeVistaPreviaChecklist() {
      window.print();
    }

    function pasoConfirmarEnvio() {
      document.getElementById('checklist-confirm-asunto').innerText = checklistState.asunto || '';
      var avisoEl = document.getElementById('checklist-confirm-aviso-prueba');
      if (checklistState.correoEsPrueba) {
        avisoEl.innerText = '⚠️ Modo prueba: se enviará a ' + checklistState.correoDestino + (checklistState.correoCC ? ' (copia a ' + checklistState.correoCC + ')' : '') + ', no a Luis Prado todavía.';
        avisoEl.style.display = 'block';
      } else {
        avisoEl.innerText = '';
        avisoEl.style.display = 'none';
      }
      mostrarPasoChecklist('confirm');
    }

    function copiarChecklistParaOutlook_() {
      var html = checklistState.htmlOutlook, texto = checklistState.textoOutlook || '', hecho = false;
      if (!html) return false;
      function alCopiar(e) { try { e.clipboardData.setData('text/html', html); e.clipboardData.setData('text/plain', texto); e.preventDefault(); hecho = true; } catch (err) {} }
      document.addEventListener('copy', alCopiar);
      try { document.execCommand('copy'); } catch (e2) {}
      document.removeEventListener('copy', alCopiar);
      return hecho;
    }
    function cuerpoOutlookChecklist_() {
      var t = String(checklistState.textoOutlook || '').replace(/☐ /g, '- ').replace(/ — /g, ' - ').replace(/—/g, '-').replace(/ · /g, ' | ');
      return t ? t + '\\n\\nUn saludo.' : '';
    }
    function ccOutlookChecklist_() { return String(checklistState.correoCC || '').split(',').join(';'); }
    function enlaceOutlookChecklist_(conCuerpo) {
      var ccOutlook = ccOutlookChecklist_();
      var enlace = 'mailto:' + checklistState.para + '?' + (ccOutlook ? 'cc=' + ccOutlook + '&' : '') + 'subject=' + encodeURIComponent(checklistState.asunto || 'Checklist de Agencias');
      var cuerpo = conCuerpo ? cuerpoOutlookChecklist_() : '';
      var conTexto = cuerpo ? enlace + '&body=' + encodeURIComponent(cuerpo.replace(/\\r?\\n/g, '\\r\\n')) : enlace;
      return conTexto.length <= 1900 ? conTexto : enlace; // muy largo: sin texto (se pega)
    }
    function prepararCorreoOutlook() {
      var ccOutlook = ccOutlookChecklist_();
      if (!checklistState.htmlOutlook) { alert('Espera un momento a que termine de cargar la vista previa.'); return; }
      var copiado = copiarChecklistParaOutlook_();
      var enlace = enlaceOutlookChecklist_(false), llevaTexto = false;
      try { var a = document.createElement('a'); a.href = enlace; a.target = '_blank'; a.rel = 'noopener'; document.body.appendChild(a); a.click(); document.body.removeChild(a); } catch (eA) {}
      var info = document.getElementById('checklist-outlook-info');
      document.getElementById('checklist-result-msg').style.display = 'none';
      var paso = 'style="margin:0 0 8px; font-size:13.5px; line-height:1.5;"';
      info.innerHTML = '<div style="background:#ecfdf5; border:1px solid #a7f3d0; border-radius:10px; padding:12px 14px; color:#065f46;">' +
        '<p ' + paso + '><b>1.</b> Se ha abierto un correo nuevo en Outlook para <b>Luis Prado</b>' + (ccOutlook ? ' (con copia a ' + escAttr(ccOutlook) + ')' : '') + ' y con el asunto puesto.</p>' +
        (llevaTexto
          ? '<p ' + paso + '><b>2.</b> El correo ya lleva el checklist escrito dentro. Solo tienes que pulsar <b>Enviar</b> en Outlook.</p>' +
            (copiado ? '<p style="margin:0; font-size:12px; opacity:.85;">¿Lo prefieres con la tabla de colores? En el correo, borra el texto y pulsa <b>Ctrl+V</b> (ya está copiada).</p>' : '') + '</div>'
          : '<p ' + paso + '><b>2.</b> ' + (copiado ? 'El checklist ya está copiado: haz clic en el cuerpo del correo y pulsa <b>Ctrl+V</b>.' : 'Pulsa «📋 Copiar el checklist» (abajo), haz clic en el cuerpo del correo y pulsa <b>Ctrl+V</b>.') + '</p>' +
            '<p style="margin:0; font-size:13.5px;"><b>3.</b> Pulsa <b>Enviar</b> en Outlook.</p></div>') +
        '<div style="display:flex; gap:8px; flex-wrap:wrap; margin-top:10px;">' +
          '<button class="btn-checklist-print" onclick="var ok = copiarChecklistParaOutlook_(); this.innerText = ok ? &quot;✅ Copiado&quot; : &quot;⚠️ No se pudo copiar&quot;;">📋 ' + (copiado ? 'Volver a copiar' : 'Copiar') + ' el checklist</button>' +
          '<a class="btn-checklist-print" style="text-decoration:none; display:inline-block;" href="' + escAttr(enlace) + '" target="_blank" rel="noopener">✉️ Abrir otra vez el correo</a>' +
        '</div>' +
        (checklistState.yo ? '<p style="margin:10px 0 0; font-size:12px; line-height:1.45;">🧪 ¿Quieres ver antes cómo queda? <a href="' + escAttr('mailto:' + checklistState.yo + '?subject=' + encodeURIComponent('PRUEBA - ' + (checklistState.asunto || 'Checklist de Agencias'))) + '" target="_blank" rel="noopener" onclick="copiarChecklistParaOutlook_();">Enviármelo antes a mí</a> (se abre otro correo solo para ti; pega con Ctrl+V y envía).</p>' : '') +
        '<p style="margin:10px 0 0; font-size:11.5px; color:var(--muted); line-height:1.45;">Si no se abre Outlook (se abre otro programa o nada), crea tú el correo: para <b>' + escAttr(checklistState.para) + '</b>' + (ccOutlook ? ', CC <b>' + escAttr(ccOutlook) + '</b>' : '') + ', asunto <b>' + escAttr(checklistState.asunto) + '</b>, y pega el checklist.</p>';
      info.style.display = 'block';
      mostrarPasoChecklist('result');
      google.script.run.withSuccessHandler(function(r) {
        if (r && r.envios) repintarEnviosTrasEnvio_(r.envios);
      }).withFailureHandler(function() {}).marcarChecklistPreparado(checklistState.modo, checklistState.offset, checklistState.fechaDia || '');
    }

    function repintarEnviosTrasEnvio_(envios) {
      if (checklistState.offset === calOff) pintarEnviosChecklist_(envios);
      else if (checklistState.offset === calOff - 1 && ultimosEnviosCal_) { ultimosEnviosCal_.anterior = envios.semana; pintarEnviosChecklist_(ultimosEnviosCal_); }
    }
    function avisoYaPreparadoChecklist_(env, modo, fechaDia) {
      var el = document.getElementById('checklist-warn');
      if (!el || !env || !env.dias) return;
      var hoy = new Date(), kHoy = hoy.getFullYear() + ('0' + (hoy.getMonth() + 1)).slice(-2) + ('0' + hoy.getDate()).slice(-2), x = null, que = '';
      if (modo === 'semana') { x = env.semana; que = 'el checklist semanal de esa semana'; }
      else { var k = fechaDia || kHoy; env.dias.forEach(function(d) { if (d.k === k) x = d; }); que = fechaDia && fechaDia !== kHoy ? 'el checklist de ese día' : 'el checklist de hoy'; }
      if (!x || x.estado !== 'si') return;
      var f = new Date(x.t), cuando = ('0' + f.getDate()).slice(-2) + '/' + ('0' + (f.getMonth() + 1)).slice(-2) + ' a las ' + ('0' + f.getHours()).slice(-2) + ':' + ('0' + f.getMinutes()).slice(-2);
      el.innerText = '🔄 ' + que.charAt(0).toUpperCase() + que.slice(1) + ' ya se preparó el ' + cuando + (x.u ? ' (' + x.u + ')' : '') + '. Si lo vuelves a enviar, saldrá como ACTUALIZADO (con lo que hay ahora) y sustituye al anterior.';
      el.style.display = 'block';
    }

    function volverAPreviewChecklist() {
      mostrarPasoChecklist('preview');
    }

    function confirmarEnvioChecklist() {
      var btn = document.getElementById('btn-checklist-confirmar-envio');
      var textoOriginal = btn.innerText;
      btn.innerText = 'Enviando…'; btn.disabled = true;
      google.script.run
        .withSuccessHandler(function(r) {
          btn.innerText = textoOriginal; btn.disabled = false;
          var msgEl = document.getElementById('checklist-result-msg');
          if (r.enviado && r.envios) repintarEnviosTrasEnvio_(r.envios);
          if (r.enviado) {
            msgEl.className = 'checklist-ok-msg';
            var destTxt = r.destinatario || 'la dirección configurada';
            var ccTxt = (r.cc ? ' (copia a ' + r.cc + ')' : '') + (r.copiaOculta ? '. Te ha llegado una copia a ' + r.copiaOculta + ' como prueba: si no la ves, mira en Correo no deseado' : '');
            msgEl.innerText = r.esPrueba
              ? ('✅ Checklist enviado en modo prueba a ' + destTxt + ccTxt + '.')
              : ('✅ Checklist enviado por correo a Luis Prado' + ccTxt + '.');
          } else {
            msgEl.className = 'checklist-warn';
            msgEl.innerText = 'ℹ️ Todavía no se ha enviado ningún correo real: falta indicar la dirección de Luis Prado en el código (CORREO_JEFATURA). El documento de arriba es exactamente el que se enviaría en cuanto se configure esa dirección.';
          }
          mostrarPasoChecklist('result');
        })
        .withFailureHandler(function(err) {
          btn.innerText = textoOriginal; btn.disabled = false;
          var msgEl = document.getElementById('checklist-result-msg');
          msgEl.className = 'checklist-warn';
          msgEl.innerText = '⚠️ No se ha podido procesar el envío: ' + (err && err.message ? err.message : err);
          mostrarPasoChecklist('result');
        })
        .enviarChecklistPorCorreo(checklistState.modo, checklistState.offset, checklistState.fechaDia || '');
    }

    function cerrarPreviewChecklist() {
      document.getElementById('checklist-overlay').hidden = true;
    }

    var debounceBusqueda = null;
    var busquedaEnCurso = false;

    function initBuscar() {
      var inputRef = document.getElementById('b-ref');
      var selectAge = document.getElementById('b-age');

      inputRef.value = localStorage.getItem('lastRef') || '';
      selectAge.value = localStorage.getItem('lastAge') || '';

      inputRef.addEventListener('keydown', function(ev) {
        if (ev.key === 'Enter' || ev.keyCode === 13) { ev.preventDefault(); iniciarBusqueda(); }
      });
      selectAge.addEventListener('change', function() {
        if (document.getElementById('b-ref').value.trim() !== '') iniciarBusqueda();
      });

      try { inputRef.focus(); inputRef.select(); } catch (eFoco) {}
      try { rbInit(); } catch (eRb) {}
      try { biIniciar_(); } catch (eBi) {}
      try { if (typeof mpRecoger_ === 'function') mpRecoger_('buscar'); } catch (eMp) {}
    }

    function iniciarBusqueda(incluirSimilares) {
      try { if (typeof mpQuitarAviso_ === 'function') mpQuitarAviso_(); } catch (eMp) {}
      var conSimilares = (incluirSimilares === true);
      if (!conSimilares && typeof iniciarBusquedaNombre_ === 'function' && modoNombreActivo_()) { iniciarBusquedaNombre_(); return; }
      var ref = document.getElementById('b-ref').value.trim();
      var age = document.getElementById('b-age').value;

      if (!ref) {
        document.getElementById('b-res').innerHTML = '<div class="empty-inline">Tus resultados aparecerán aquí...</div>';
        return;
      }
      if (ref.length < 3) {
        document.getElementById('b-res').innerHTML = '<div class="empty-inline">Escribe al menos 3 dígitos...</div>';
        return;
      }

      localStorage.setItem('lastRef', ref);
      localStorage.setItem('lastAge', age);

      var refVisible = ref.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
      var instYa = (!conSimilares && biMostrado && biMostrado.ref === ref && biMostrado.age === age);
      var inst = (conSimilares || instYa) ? null : biBuscarLocal_(ref, age);
      var parLoc = (conSimilares || (!instYa && !(inst && inst.length > 0))) ? biParecidasLocal_(ref, age) : null;
      biParecidasVista = null;
      if (inst && inst.length > 0) {
        biMostrado = { ref: ref, age: age, res: inst };
        mostrarResultados(inst);
        document.getElementById('b-res').insertAdjacentHTML('afterbegin', '<div class="bi-estado" id="bi-estado">⚡ Comprobando en el documento…</div>');
      } else if (parLoc && parLoc.length > 0) {
        biMostrado = null;
        mostrarResultados(parLoc);
        biCompletarL_(parLoc);
        if (conSimilares) document.getElementById('b-res').insertAdjacentHTML('afterbegin', '<div class="bi-estado" id="bi-estado">⚡ Parecidas entre las últimas filas · buscando en todo el documento…</div>');
        else {
          biParecidasVista = { ref: ref, age: age };
          document.getElementById('b-res').insertAdjacentHTML('afterbegin', '<div class="bi-estado" id="bi-estado">⚡ No está entre las últimas filas. Parecidos de las últimas 3.000 · buscando el exacto en todo el documento…</div>');
          if (biDatos) { biEsperaDelta = { ref: ref, age: age }; biDeltaUltimo = 0; biDelta_('falta'); }
        }
      } else if (!instYa) {
        if (!conSimilares && biDatos) { biEsperaDelta = { ref: ref, age: age }; biDeltaUltimo = 0; biDelta_('falta'); }
        biMostrado = null;
        document.getElementById('b-res').innerHTML = '<div class="loading-row"><span class="spinner"></span>' + (conSimilares ? 'Buscando pedidos parecidos a ' + refVisible + ' en todo el documento (tarda un poco más)...' : 'Buscando ' + refVisible + '...') + '</div>';
      }
      if (busquedaEnCurso) return;
      busquedaEnCurso = true;
      var resuelta = false, pendientes = 0, reintentoT = null, pistaEnviada = pistaBusqueda;
      function alLlegar(res) {
          if (resuelta) return; resuelta = true; clearTimeout(reintentoT);
          if (res && res.length > 0 && res[0].exacta) { pistaBusqueda = { hoja: res[0].hoja, fila: res[0].fila }; try { localStorage.setItem('pistaRef', JSON.stringify(pistaBusqueda)); } catch (ePista) {} }
          biAprender_(res);
          terminarBusqueda_(ref, age, function() {
          if (!conSimilares && biConfirmar_(ref, age, res)) return;
          biMostrado = null;
          if (!conSimilares && res.length === 0 && biParecidasVista && biParecidasVista.ref === ref && biParecidasVista.age === age && document.getElementById('bi-estado')) {
            var est = document.getElementById('bi-estado');
            est.className = 'bi-sin'; est.removeAttribute('id');
            est.innerHTML = '❌ El pedido ' + refVisible + ' no está en el documento. Estos son los <b>parecidos de las últimas 3.000 filas</b>:';
            document.getElementById('b-res').insertAdjacentHTML('beforeend', '<div style="text-align:center;margin-top:10px;"><button class="btn" onclick="iniciarBusqueda(true)">🔍 Buscar parecidas en todo el documento</button></div>');
            biParecidasVista = null; return;
          }
          if (!conSimilares && res.length === 0) { mostrarSinExactas_(refVisible); return; }
          biParecidasVista = null;
          var tocadas = {};
          [].forEach.call(document.querySelectorAll('#b-res .res-item[data-tocado="1"]'), function(el) { tocadas[el.id] = el; });
          mostrarResultados(res); if (!conSimilares) ofrecerSimilares_(res);
          Object.keys(tocadas).forEach(function(id) { var nuevo = document.getElementById(id); if (nuevo && nuevo !== tocadas[id]) nuevo.replaceWith(tocadas[id]); }); }); }
      function alFallar(err) {
          pendientes--; if (resuelta || pendientes > 0) return; resuelta = true; clearTimeout(reintentoT);
          terminarBusqueda_(ref, age, function() {
          if (biMostrado && biMostrado.ref === ref && biMostrado.age === age) { var est = document.getElementById('bi-estado'); if (est) { est.className = 'bi-estado aviso'; est.innerText = '⚠️ No se ha podido comprobar en el documento (' + (err && err.message ? err.message : 'error') + '). Es lo que había hace un rato.'; } return; }
          mostrarErrorBusqueda(err); }); }
      function lanzar() { pendientes++; google.script.run.withSuccessHandler(alLlegar).withFailureHandler(alFallar).backendBuscarReferencia(ref, age, conSimilares, pistaEnviada); }
      lanzar();
      if (!conSimilares) reintentoT = setTimeout(function() { if (!resuelta) { busquedasReintentadas++; lanzar(); } }, BUSQUEDA_REINTENTO_MS_);
    }
    var BUSQUEDA_REINTENTO_MS_ = 15000, busquedasReintentadas = 0;
    var biDatos = null, biMostrado = null, biCargando = false, biUltimaCarga = 0;
    var BI_RENOVAR_MS_ = 4 * 60000, BI_CADUCA_MS_ = 15 * 60000;
    function biNorm_(v) { return String(v === null || v === undefined ? '' : v).toLowerCase().trim().replace(/^0+(?=\\d)/, ''); }
    function biCargar_(forzar) {
      if (biCargando) return;
      if (!forzar && Date.now() - biUltimaCarga < 60000) return;
      if (document.hidden && biDatos) return;
      biCargando = true; biUltimaCarga = Date.now();
      google.script.run.withSuccessHandler(function(d) {
        biCargando = false; biDuroMs = Date.now() - biUltimaCarga;
        if (d && d.hojas) { d.recibido = Date.now(); biExtraDescomprimir_(d); biDatos = d; biGuardarMemoria_(true);  if (d.notas) ntNotasPed = d.notas; bpRep = d.rep || bpRep; bpColor = d.color || bpColor; ntRepintarPosits_(); }
      }).withFailureHandler(function() { biCargando = false; biDuroMs = Date.now() - biUltimaCarga; }).obtenerRecientesBuscar();
    }
    var biDeltaCargando = false, biDeltaUltimo = 0, biDeltaPausa = 0;
    function biCompletarL_(lista) {
      var faltan = (lista || []).filter(function(r) { return r.notaL === '…'; }).map(function(r) { return { hoja: r.hoja, fila: r.fila }; });
      if (!faltan.length) return;
      google.script.run.withSuccessHandler(function(m) {
        if (!m) return;
        faltan.forEach(function(x) {
          var el = document.getElementById(x.hoja + '-' + x.fila); if (!el || el.getAttribute('data-tocado') === '1') return;
          var nl = el.querySelector('.res-notaL'); if (!nl) return;
          var v = m[x.hoja + '#' + x.fila]; if (v === undefined) return;
          nl.className = 'res-notaL ' + (v ? 'con-dato' : 'sin-dato');
          nl.innerHTML = v ? '📋 Columna L: <strong>' + escAttr(v) + '</strong>' : '📋 Columna L: (vacío)';
        });
      }).withFailureHandler(function() {}).obtenerLFilasBuscar(faltan);
    }
    function biExtraDescomprimir_(d) {
      (d.hojas || []).forEach(function(H) {
        var x = H.extra; if (!x || Array.isArray(x)) { if (!x) H.extra = []; return; }
        var ns = x.n ? x.n.split('|') : [], out = [];
        for (var i = 0; i < ns.length; i++) { if (!ns[i]) continue; var c = x.a.charAt(i); out.push([x.top - i, ns[i], c === '-' ? -1 : parseInt(c, 36)]); }
        H.extra = out;
      });
    }
    var biEsperaDelta = null, biParecidasVista = null;
    function biTrasDelta_() {
      var q = biEsperaDelta; biEsperaDelta = null;
      if (!q || biMostrado || !busquedaEnCurso) return;
      var inp = document.getElementById('b-ref'), ag = document.getElementById('b-age');
      if (!inp || inp.value.trim() !== q.ref || (ag && ag.value !== q.age)) return;
      var inst = biBuscarLocal_(q.ref, q.age);
      if (!inst || !inst.length) return;
      biMostrado = { ref: q.ref, age: q.age, res: inst }; biParecidasVista = null;
      mostrarResultados(inst);
      document.getElementById('b-res').insertAdjacentHTML('afterbegin', '<div class="bi-estado" id="bi-estado">⚡ Comprobando en el documento…</div>');
    }
    function biLev_(a, b) {
      var an = a.length, bn = b.length; if (!an) return bn; if (!bn) return an;
      var fila = []; for (var j = 0; j <= an; j++) fila[j] = j;
      for (var i = 1; i <= bn; i++) { var ant = fila[0]; fila[0] = i; for (var k = 1; k <= an; k++) { var t = fila[k]; fila[k] = a.charAt(k - 1) === b.charAt(i - 1) ? ant : Math.min(ant + 1, fila[k] + 1, fila[k - 1] + 1); ant = t; } }
      return fila[an];
    }
    function biParecidasLocal_(ref, age) {
      if (!biDatos || Date.now() - biDatos.recibido > BI_CADUCA_MS_) return null;
      var t = biNorm_(ref), ag = String(age || '').toUpperCase().trim(), out = [];
      if (!t) return null;
      var umbral = t.length <= 4 ? 1 : (t.length <= 8 ? 2 : 3);
      biDatos.hojas.forEach(function(H) {
        H.filas.forEach(function(f) {
          if (ag !== '' && f[3].indexOf(ag) === -1) return;
          var v = f[1]; if (!v) return;
          var exacta = v.indexOf(t) > -1, d = exacta ? 0 : (Math.abs(v.length - t.length) > umbral ? umbral + 1 : biLev_(t, v));
          if (!exacta && d > umbral) return;
          var base = Math.max(t.length, v.length, 1), pct = exacta ? 100 : Math.max(0, Math.min(99, Math.round((1 - d / base) * 100)));
          out.push({ hoja: H.hoja, fila: f[0], col: 3, agencia: f[3] || 'SIN AGENCIA', matchStr: f[2], exacta: exacta, distancia: d, porcentaje: pct, notaL: f[4], multi: f[5] === 'SI' });
        });
        (H.extra || []).forEach(function(e) {
          var agE = (H.ag && H.ag[e[2]]) || '';
          if (ag !== '' && agE.indexOf(ag) === -1) return;
          var v = e[1]; if (!v) return;
          var exacta = v.indexOf(t) > -1, d = exacta ? 0 : (Math.abs(v.length - t.length) > umbral ? umbral + 1 : biLev_(t, v));
          if (!exacta && d > umbral) return;
          var base = Math.max(t.length, v.length, 1), pct = exacta ? 100 : Math.max(0, Math.min(99, Math.round((1 - d / base) * 100)));
          out.push({ hoja: H.hoja, fila: e[0], col: 3, agencia: agE || 'SIN AGENCIA', matchStr: v, exacta: exacta, distancia: d, porcentaje: pct, notaL: '…', multi: false });
        });
      });
      out.sort(function(a, b) { return a.distancia - b.distancia; });
      return out.slice(0, 30);
    }
    function biDelta_(motivo) {
      if (!biDatos || biDeltaCargando || biCargando) return;
      if (Date.now() < biDeltaPausa) return;
      var conocidas = {};
      biDatos.hojas.forEach(function(H) { conocidas[H.hoja] = H.filas.length ? H.filas[0][0] : 0; });
      biDeltaCargando = true; biDeltaUltimo = Date.now();
      google.script.run.withSuccessHandler(function(d) {
        biDeltaCargando = false;
        var duro = Date.now() - biDeltaUltimo;
        if (duro > 10000) biDeltaPausa = Date.now() + 120000; // el documento va lento: se espera
        if (!d || !biDatos) return;
        if (d.recargar) { biCargar_(true); return; }
        (d.hojas || []).forEach(function(N) {
          var H = null; biDatos.hojas.forEach(function(x) { if (x.hoja === N.hoja) H = x; });
          if (!H) return;
          var viejas = H.filas.filter(function(f) { return f[0] < N.desde; });
          H.filas = N.filas.concat(viejas);
          if (H.extra) { var minF = H.filas.length ? H.filas[H.filas.length - 1][0] : N.desde; H.extra = H.extra.filter(function(e) { return e[0] < minF; }); }
        });
        biDatos.delta = Date.now();
        biGuardarMemoria_(false);
        try { biTrasDelta_(); } catch (eTD) {}
      }).withFailureHandler(function() { biDeltaCargando = false; biDeltaPausa = Date.now() + 60000; }).obtenerFilasNuevasBuscar(conocidas);
    }
    var biDuroMs = 0;
    function biProgramar_() {
      var espera = BI_RENOVAR_MS_ + Math.floor(Math.random() * 60000);
      if (biDuroMs > 20000) espera = espera * 2;
      setTimeout(function() { if (!document.hidden && accionActual === 'buscar') biCargar_(true); biProgramar_(); }, espera);
    }
    var biMemoriaT = 0, biIniciado = false;
    function biGuardarMemoria_(ya) {
      if (!biDatos) return;
      if (!ya && Date.now() - biMemoriaT < 20000) return;
      biMemoriaT = Date.now();
      cacheGuardar_('buscar', { d: biDatos, notas: (typeof ntNotasPed !== 'undefined' ? ntNotasPed : null), rep: bpRep, color: bpColor });
    }
    function biIniciar_() {
      if (biIniciado) { if (!biDatos || Date.now() - biDatos.recibido > 60000) biCargar_(false); return; }
      biIniciado = true;
      var mem = cacheLeer_('buscar', BI_CADUCA_MS_);
      if (mem && mem.d && mem.d.d && mem.d.d.hojas) {
        biDatos = mem.d.d; if (mem.d.notas) ntNotasPed = mem.d.notas; bpRep = mem.d.rep || bpRep; bpColor = mem.d.color || bpColor;
        try { ntRepintarPosits_(); } catch (eRp) {}
        if (Date.now() - biDatos.recibido > 60000) biCargar_(true); else biUltimaCarga = Date.now();
      } else biCargar_(true);
      biProgramar_();
      setInterval(function() { if (!document.hidden && accionActual === 'buscar') biDelta_('reloj'); }, 40000);
      var campoRef = document.getElementById('b-ref');
      function biDeltaSiToca_() { if (accionActual === 'buscar' && Date.now() - biDeltaUltimo > 15000) biDelta_('foco'); }
      window.addEventListener('focus', biDeltaSiToca_);
      if (campoRef) { campoRef.addEventListener('focus', biDeltaSiToca_); campoRef.addEventListener('pointerdown', biDeltaSiToca_); }
      document.addEventListener('visibilitychange', function() { if (!document.hidden && accionActual === 'buscar' && biDatos && Date.now() - biDatos.recibido > BI_RENOVAR_MS_) biCargar_(false); });
    }
    function biBuscarLocal_(ref, age) {
      if (!biDatos || Date.now() - biDatos.recibido > BI_CADUCA_MS_) return null;
      var t = biNorm_(ref), ag = String(age || '').toUpperCase().trim(), out = [];
      if (!t) return null;
      for (var h = 0; h < biDatos.hojas.length && out.length < 30; h++) {
        var H = biDatos.hojas[h];
        for (var i = 0; i < H.filas.length && out.length < 30; i++) {
          var f = H.filas[i];
          if (f[1].indexOf(t) === -1) continue;
          if (ag !== '' && f[3].indexOf(ag) === -1) continue;
          if (biNorm_(f[2]) === '' || biNorm_(f[2]).indexOf(t) === -1) continue;
          out.push({ hoja: H.hoja, fila: f[0], col: 3, agencia: f[3] || 'SIN AGENCIA', matchStr: f[2], exacta: true, distancia: 0, porcentaje: 100, notaL: f[4], multi: f[5] === 'SI' });
        }
      }
      return out;
    }
    function biAprender_(res) {
      if (!biDatos || !res) return;
      var falta = false;
      res.forEach(function(r) {
        if (!r.exacta) return;
        var H = null; biDatos.hojas.forEach(function(x) { if (x.hoja === r.hoja) H = x; });
        if (!H) return;
        var hallada = false;
        H.filas.forEach(function(f) { if (f[0] === r.fila) { hallada = true; f[4] = r.notaL || ''; if (!r.soloL) { f[3] = r.agencia === 'SIN AGENCIA' ? '' : r.agencia; f[2] = String(r.matchStr); f[1] = biNorm_(r.matchStr); } } });
        if (!hallada && H.filas.length && r.fila > H.filas[H.filas.length - 1][0]) falta = true;
      });
      if (falta) biDelta_('falta');
    }
    function biClave_(r) { return r.hoja + '#' + r.fila + '#' + r.agencia + '#' + biNorm_(r.matchStr) + '#' + (r.exacta ? 'E' : 'P') + '#' + (r.multi ? 'M' : ''); }
    function biConfirmar_(ref, age, res) {
      if (!biMostrado || biMostrado.ref !== ref || biMostrado.age !== age) return false;
      var a = biMostrado.res.map(biClave_).join('|'), b = (res || []).map(biClave_).join('|');
      if (a !== b) return false;
      res.forEach(function(r) {
        var el = document.getElementById(r.hoja + '-' + r.fila);
        if (!el || el.getAttribute('data-tocado') === '1') return;
        var nl = el.querySelector('.res-notaL');
        if (!nl) return;
        var ahora = r.notaL ? '📋 Columna L: <strong>' + r.notaL + '</strong>' : '📋 Columna L: (vacío)';
        if (nl.innerHTML !== ahora) { nl.className = 'res-notaL ' + (r.notaL ? 'con-dato' : 'sin-dato'); nl.innerHTML = ahora; }
      });
      var est = document.getElementById('bi-estado'); if (est) est.remove();
      biMostrado = null;
      try { if (res.length > 0) { pistaBusqueda = { hoja: res[0].hoja, fila: res[0].fila }; localStorage.setItem('pistaRef', JSON.stringify(pistaBusqueda)); } } catch (eP) {}
      return true;
    }

    var ntNotasPed = null;
    function ntPositHtml_(ref) {
      var lista = ntNotasPed && ntNotasPed[biNorm_(ref)];
      if (!lista || !lista.length) return '';
      return lista.map(function(n) {
        return '<div class="nt-posit">📝 ' + escAttr(n.t) + '<small>' + escAttr(n.u) + ' · ' + escAttr(n.f) + '</small>' +
          '<button class="nt-acc res" onclick="ntResolverNotaPedido(this, ' + Number(n.id) + ', ' + Number(n.fila) + ')">✅ Resuelta</button></div>';
      }).join('');
    }
    function ntRepintarPosits_() {
      document.querySelectorAll('#b-res .res-item').forEach(function(card) {
        card.querySelectorAll('.nt-posit, .bp-marcas').forEach(function(x) { x.remove(); });
        var info = card.querySelector('.res-info'); if (!info) return;
        var partes = (card.id || '').split('-'), fila = Number(partes.pop()), hoja = partes.join('-');
        var m = bpMarcasHtml_({ hoja: hoja, fila: fila, matchStr: card.getAttribute('data-ref'), multi: card.getAttribute('data-multi') === '1' });
        var notaL = info.querySelector('.res-notaL');
        if (m) { if (notaL) notaL.insertAdjacentHTML('afterend', m); else info.insertAdjacentHTML('beforeend', m); }
        var h = ntPositHtml_(card.getAttribute('data-ref'));
        if (h) info.insertAdjacentHTML('beforeend', h);
      });
    }
    var bpRep = null, bpColor = null;
    function bpMarcasHtml_(r) {
      if (!r || r.hoja !== 'Entradas') return '';
      var h = '', m = bpRep && bpRep[biNorm_(r.matchStr)], col = bpColor && bpColor[r.fila];
      if (r.multi) m = { t: 'M', n: (m && m.n) || 1 };
      if (m && m.t === 'M') h += '<span class="bp-marca bp-m" title="Columna D = SI: hay más de un bulto para abrir">📦 MULTIBULTO · hay más bultos' + (m.n > 1 ? ' (' + m.n + ' filas)' : '') + '</span>';
      if (m && m.t === 'R') h += '<span class="bp-marca bp-r" title="Columna D = NO y el pedido está más de una vez (en rojo): pudo venir en otro momento y tener parte facturada">🔴 REPETIDO · ' + m.d.map(function(x) { return escAttr(x[0]) + (x[1] > 1 ? ' (' + x[1] + ')' : ''); }).join(', ') + '</span>';
      if (col) h += '<span class="bp-marca bp-c" title="La fila entera está pintada de otro color"><i style="background:' + escAttr(col) + '"></i>Fila de color (aviso)</span>';
      return h ? '<div class="bp-marcas">' + h + '</div>' : '';
    }
    function ntAbrirNotaPedido(btn) {
      var card = btn.closest('.res-item'); if (!card) return;
      var f = card.querySelector('.nt-form');
      if (f) { f.remove(); return; }
      var fila = btn.parentElement;
      fila.insertAdjacentHTML('afterend', '<div class="nt-form"><textarea placeholder="Nota para este pedido (la verá quien lo busque)…"></textarea>' +
        '<div class="nt-accs"><button class="nt-acc" onclick="this.closest(&quot;.nt-form&quot;).remove()">Cancelar</button>' +
        '<button class="nt-acc res" onclick="ntGuardarNotaPedido(this)">💾 Guardar nota</button></div></div>');
      var ta = card.querySelector('.nt-form textarea'); if (ta) ta.focus();
    }
    function ntGuardarNotaPedido(btn) {
      var card = btn.closest('.res-item'), form = btn.closest('.nt-form'), ta = form.querySelector('textarea');
      var txt = ta.value.trim(), ped = card.getAttribute('data-ref') || '';
      if (!txt) { ta.focus(); return; }
      btn.disabled = true; btn.innerText = 'Guardando…';
      google.script.run.withSuccessHandler(function(r) {
        if (!r || r.error) { btn.disabled = false; btn.innerText = '💾 Guardar nota'; alert('No se pudo guardar: ' + (r && r.error ? r.error : 'error')); return; }
        form.remove(); ntNotasPed = r.notas || ntNotasPed; ntRepintarPosits_();
      }).withFailureHandler(function(err) { btn.disabled = false; btn.innerText = '💾 Guardar nota'; alert('No se pudo guardar: ' + (err && err.message ? err.message : err)); }).guardarNotaPedido(ped, txt);
    }
    function ntResolverNotaPedido(btn, id, fila) {
      btn.disabled = true; btn.innerText = '✅ …';
      google.script.run.withSuccessHandler(function(r) {
        if (!r || r.error) { btn.disabled = false; btn.innerText = '✅ Resuelta'; alert('No se pudo guardar: ' + (r && r.error ? r.error : 'error')); return; }
        ntNotasPed = r.notas || ntNotasPed; ntRepintarPosits_();
      }).withFailureHandler(function(err) { btn.disabled = false; btn.innerText = '✅ Resuelta'; alert('No se pudo guardar: ' + (err && err.message ? err.message : err)); }).resolverNotaPedido(id, fila);
    }

    var pistaBusqueda = null;
    try { pistaBusqueda = JSON.parse(localStorage.getItem('pistaRef') || 'null'); } catch (ePistaIni) { pistaBusqueda = null; }

    function mostrarSinExactas_(refVisible) {
      document.getElementById('b-res').innerHTML =
        '<div style="color:var(--red-dark); text-align:center;">❌ El pedido ' + refVisible + ' no está en el documento.' + (biDatos && Date.now() - biDatos.recibido <= BI_CADUCA_MS_ ? '<br><span style="font-size:12px;color:var(--muted);">Tampoco hay parecidos en las últimas 3.000 filas.</span>' : '') + '</div>' +
        '<div style="text-align:center;margin-top:10px;"><button class="btn" onclick="iniciarBusqueda(true)">🔍 Buscar parecidas (por si hay algún número mal escrito)</button></div>' +
        '<div class="empty-inline" style="margin-top:6px;">Esa búsqueda mira todo el documento y tarda más.</div>';
      try { if (typeof mpAvisarSinResultados_ === 'function') mpAvisarSinResultados_(); } catch (eMp) {}
    }

    function ofrecerSimilares_(res) {
      if (!res || res.length === 0) return;
      for (var i = 0; i < res.length; i++) { if (!res[i].exacta) return; }
      document.getElementById('b-res').insertAdjacentHTML('beforeend',
        '<div style="text-align:center;margin-top:10px;"><button class="btn" onclick="iniciarBusqueda(true)">🔍 Buscar en todo el documento (y parecidas)</button></div>');
    }

    function terminarBusqueda_(refEnviada, ageEnviada, mostrar) {
      busquedaEnCurso = false;
      if (!document.getElementById('b-ref') || !document.getElementById('b-res')) return;
      var refAhora = document.getElementById('b-ref').value.trim();
      var ageAhora = document.getElementById('b-age').value;
      if (refAhora === refEnviada && ageAhora === ageEnviada) mostrar();
      else iniciarBusqueda();
    }

    function mostrarErrorBusqueda(err) {
      document.getElementById('b-res').innerHTML = '<div style="color:var(--red-dark); text-align:center;">⚠️ Error: ' + err.message + '</div>';
    }

    function mostrarResultados(res) {
      var clicsGuardados = JSON.parse(localStorage.getItem('clicsRef') || '[]');

      if (res.length === 0) {
        document.getElementById('b-res').innerHTML = '<div style="color:var(--red-dark); text-align:center;">❌ No se encontraron coincidencias, ni exactas ni similares.</div>';
        try { if (typeof mpAvisarSinResultados_ === 'function') mpAvisarSinResultados_(); } catch (eMp) {}
        return;
      }

      var html = '';
      res.forEach(function(r) {
        var idUnico = r.hoja + '-' + r.fila;
        var esVisitado = clicsGuardados.indexOf(idUnico) > -1 ? 'visitado' : '';
        var txtBtn = esVisitado ? '✔️ Revisado' : '➡️ Ir a celda';
        var etiqueta = r.exacta
          ? '<span class="badge-exacta">🎯 Exacta</span>'
          : '<span class="badge-similar">🟡 Similar</span>';

        var pctClase = r.porcentaje >= 90 ? 'badge-pct-alta' : (r.porcentaje >= 75 ? 'badge-pct-media' : 'badge-pct-baja');
        var pctHtml = '<span class="badge-pct ' + pctClase + '">' + r.porcentaje + '% texto</span>';

        var notaLHtml = r.notaL
          ? '<div class="res-notaL con-dato">📋 Columna L: <strong>' + r.notaL + '</strong></div>'
          : '<div class="res-notaL sin-dato">📋 Columna L: (vacío)</div>';

        html += '<div id="' + idUnico + '" class="res-item ' + esVisitado + '" data-ref="' + escAttr(r.matchStr) + '"' + (r.multi ? ' data-multi="1"' : '') + '>';
        html += '<div class="res-info"><div class="res-hoja">📄 ' + r.hoja + ' (Fila ' + r.fila + ') ' + etiqueta + ' ' + pctHtml + '</div>';
        html += '<div><strong>' + r.agencia + '</strong> - Ref: ' + r.matchStr + '</div>';
        if (r.nombre) html += '<div class="res-nombre">👤 ' + escAttr(r.nombre) + '</div>';
        html += notaLHtml;
        html += bpMarcasHtml_(r);
        html += ntPositHtml_(r.matchStr);
        html += '</div>';
        html += '<div class="res-acciones-row">';
        html += '<button class="btn-ir" onclick="saltarACelda(\\'' + r.hoja + '\\', ' + r.fila + ', ' + r.col + ', \\'' + idUnico + '\\')">' + txtBtn + '</button>';
        html += '<button class="btn-rapido-toggle" id="toggle-' + idUnico + '" onclick="toggleAccionesRapidas(\\'' + idUnico + '\\')">🛠️ Rellenar</button>';
        html += '<button class="nt-btn-nota" title="Dejar una nota de este pedido" onclick="ntAbrirNotaPedido(this)">📝</button>';
        html += '</div>';
        html += '<div class="res-rapido" id="rapido-' + idUnico + '">';
        html += '<div class="res-rapido-grupo"><div class="res-rapido-label">Rotura</div><div class="res-rapido-btns">';
        html += '<button class="qbtn qbtn-si" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + r.hoja + '\\', ' + r.fila + ', \\'rotura\\', \\'SI\\', this)">Sí</button>';
        html += '<button class="qbtn qbtn-no" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + r.hoja + '\\', ' + r.fila + ', \\'rotura\\', \\'NO\\', this)">No</button>';
        html += '</div></div>';
        html += '<div class="res-rapido-grupo"><div class="res-rapido-label">Verificación (Col. L)</div><div class="res-rapido-btns">';
        html += '<button class="qbtn qbtn-ok" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + r.hoja + '\\', ' + r.fila + ', \\'verificacion\\', \\'OK\\', this)">OK</button>';
        html += '<button class="qbtn qbtn-reclamar" data-cm-ref="' + escAttr(r.matchStr) + '" data-cm-age="' + escAttr(r.agencia) + '" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + r.hoja + '\\', ' + r.fila + ', \\'verificacion\\', \\'RECLAMAR\\', this); abrirCajaManipulada(\\'' + idUnico + '\\', this)">RECLAMAR</button>';
        html += '<button class="qbtn qbtn-reclamar-pend" data-cm-ref="' + escAttr(r.matchStr) + '" data-cm-age="' + escAttr(r.agencia) + '" onclick="marcarRapido(\\'' + idUnico + '\\', \\'' + r.hoja + '\\', ' + r.fila + ', \\'verificacion\\', \\'RECLAMAR/PENDIENTE\\', this); abrirCajaManipulada(\\'' + idUnico + '\\', this)">RECL./PEND.</button>';
        html += '</div></div>';
        html += '</div>';
        html += '</div>';
      });
      document.getElementById('b-res').innerHTML = html;
    }

    function toggleAccionesRapidas(idHtml) {
      var panel = document.getElementById('rapido-' + idHtml);
      var toggleBtn = document.getElementById('toggle-' + idHtml);
      var abierto = panel.classList.toggle('abierto');
      toggleBtn.classList.toggle('abierto', abierto);
      toggleBtn.innerText = abierto ? '✖️ Cerrar' : '🛠️ Rellenar';
    }

    function marcarRapido(idHtml, hoja, fila, campo, valor, btnEl) {
      var grupoBtns = btnEl.parentElement.querySelectorAll('.qbtn');
      var activoAntes = null;
      grupoBtns.forEach(function(b) { if (b.classList.contains('activo')) activoAntes = b; b.classList.remove('activo'); });
      btnEl.classList.add('activo');
      grupoBtns.forEach(function(b) { b.disabled = true; });
      function desbloquear() { grupoBtns.forEach(function(b) { b.disabled = false; }); }
      var tarjetaEl = document.getElementById(idHtml), pedidoTarjeta = '';
      if (tarjetaEl && tarjetaEl.classList.contains('res-item') && document.getElementById('b-res') && document.getElementById('b-res').contains(tarjetaEl)) { pedidoTarjeta = tarjetaEl.getAttribute('data-ref') || ''; tarjetaEl.setAttribute('data-tocado', '1'); }

      var notaLEl = null, notaLClaseAntes = '', notaLHtmlAntes = '';
      if (campo === 'verificacion') {
        var filaEl = document.getElementById(idHtml);
        notaLEl = filaEl ? filaEl.querySelector('.res-notaL') : null;
        if (notaLEl) {
          notaLClaseAntes = notaLEl.className; notaLHtmlAntes = notaLEl.innerHTML;
          notaLEl.className = 'res-notaL con-dato';
          notaLEl.innerHTML = '📋 Columna L: <strong>' + valor + '</strong>';
        }
      }

      function deshacer(motivo) {
        desbloquear();
        btnEl.classList.remove('activo');
        if (activoAntes) activoAntes.classList.add('activo');
        if (notaLEl) { notaLEl.className = notaLClaseAntes; notaLEl.innerHTML = notaLHtmlAntes; }
        alert('No se pudo guardar: ' + motivo);
      }

      google.script.run
        .withSuccessHandler(function(res) {
          if (res && res.error) deshacer(res.mensaje || 'error desconocido');
          else {
            desbloquear();
            if (res && res.movida && tarjetaEl) { var nm = tarjetaEl.querySelector('.fila-movida'); if (!nm) { tarjetaEl.insertAdjacentHTML('beforeend', '<div class="fila-movida">↪️ La fila se había movido: guardado en la fila ' + res.fila + '.</div>'); } else { nm.innerText = '↪️ La fila se había movido: guardado en la fila ' + res.fila + '.'; } }
            if (campo === 'verificacion' && pedidoTarjeta && !(res && res.movida)) biAprender_([{ hoja: hoja, fila: fila, exacta: true, notaL: res && res.valor ? res.valor : valor, agencia: '', matchStr: pedidoTarjeta, soloL: true }]);
          }
        })
        .withFailureHandler(function(err) {
          deshacer(err && err.message ? err.message : err);
        })
        .rellenarCampoRapido(hoja, fila, campo, valor, pedidoTarjeta);
    }

    function generarCajaManipuladaHtml(idUnico, referencia, agencia, pend) {
      function ladoHtml(lado, titulo, placeholderNormal, opcionTodo, ayudaMayoria, textoTodoConfirmado) {
        return '<div class="caja-lado">' +
          '<div class="caja-lado-titulo">' + titulo + '</div>' +
          '<select class="caja-modo-select" id="caja-modo-' + lado + '-' + idUnico + '" onchange="cambiarModoCaja(\\'' + idUnico + '\\',\\'' + lado + '\\')">' +
            '<option value="normal">Indicar los códigos</option>' +
            '<option value="todo">' + opcionTodo + '</option>' +
            '<option value="mayoria">Indicar los buenos</option>' +
          '</select>' +
          '<div class="caja-codigos-wrap visible" id="caja-wrap-normal-' + lado + '-' + idUnico + '">' +
            '<textarea class="caja-codigos caja-codigos-mal" id="caja-normal-' + lado + '-' + idUnico + '" placeholder="' + placeholderNormal + '" oninput="actualizarFraseCaja(\\'' + idUnico + '\\')"></textarea>' +
            '<span class="caja-contador" id="caja-cont-normal-' + lado + '-' + idUnico + '"></span>' +
          '</div>' +
          '<div class="caja-codigos-wrap" id="caja-wrap-mayoria-' + lado + '-' + idUnico + '">' +
            '<div class="caja-sublabel">' + ayudaMayoria + '</div>' +
            '<textarea class="caja-codigos caja-codigos-bien" id="caja-mayoria-' + lado + '-' + idUnico + '" placeholder="Pega aquí los códigos (uno por línea o separados por comas)" oninput="actualizarFraseCaja(\\'' + idUnico + '\\')"></textarea>' +
            '<span class="caja-contador" id="caja-cont-mayoria-' + lado + '-' + idUnico + '"></span>' +
            '<div class="caja-mayoria-aviso" id="caja-mayoria-aviso-' + lado + '-' + idUnico + '">' +
              '<span class="caja-mayoria-aviso-icono">✅</span>' +
              '<span class="caja-mayoria-aviso-texto" id="caja-mayoria-aviso-texto-' + lado + '-' + idUnico + '"></span>' +
            '</div>' +
          '</div>' +
          '<div class="caja-codigos-wrap caja-todo-aviso" id="caja-wrap-todo-' + lado + '-' + idUnico + '">' +
            '<span class="caja-todo-icono">📦</span>' +
            '<span class="caja-todo-texto">Se indicará:<br>' + textoTodoConfirmado + '</span>' +
          '</div>' +
        '</div>';
      }
      var bultosHtml = '', ladosHtml;
      if (pend) {
        bultosHtml = '<div class="caja-bultos">' +
          '<div class="caja-lado-titulo">📦 ¿Qué bulto falta por llegar?</div>' +
          '<div class="caja-bultos-fila" id="caja-bultos-' + idUnico + '">' + botonesBultosCaja_(idUnico, 4, []) + '</div>' +
          '<div class="caja-bultos-ayuda">Pulsa el bulto (o los bultos) que no ha llegado. Abajo, lo que pasa con el bulto que sí ha llegado.</div>' +
        '</div>';
        ladosHtml = ladoHtml('rotos', 'Artículos rotos del bulto que llega', 'Pega aquí los códigos (uno por línea o separados por comas)', 'Todo el bulto roto', 'Solo si la MAYORÍA del bulto vino rota: códigos de los que SÍ están bien', 'TODO ROTO') +
          ladoHtml('faltan', 'Artículos que faltan del bulto que llega', 'Pega aquí los códigos (uno por línea o separados por comas)', 'Falta todo el bulto', 'Solo si FALTA la mayoría del bulto: códigos de los que SÍ llegaron', 'FALTA TODO');
      } else {
        ladosHtml = ladoHtml('rotos', 'Artículos rotos', 'Pega aquí los códigos (uno por línea o separados por comas)', 'Todo el pedido roto', 'Solo si la MAYORÍA vino rota: códigos de los que SÍ están bien', 'TODO EL PEDIDO ROTO') +
          ladoHtml('faltan', 'Artículos que faltan', 'Pega aquí los códigos (uno por línea o separados por comas)', 'Falta todo el pedido', 'Solo si FALTA la mayoría: códigos de los que SÍ llegaron', 'FALTA TODO EL PEDIDO');
      }
      return '<div class="caja-compositor" id="caja-' + idUnico + '" data-ref="' + escAttr(referencia) + '" data-age="' + escAttr(agencia) + '"' + (pend ? ' data-pend="1" data-bultos="" data-nb="4"' : '') + '>' +
        '<div class="caja-compositor-titulo">📦 Bulto manipulado' + (pend ? ' · pendiente de facturar' : '') + '</div>' +
        bultosHtml +
        '<div class="caja-compositor-fila">' +
          ladosHtml +
        '</div>' +
        (typeof restoCajaHtml_ === 'function' ? restoCajaHtml_(idUnico) : '') +
        '<div class="caja-frase-preview" id="caja-preview-' + idUnico + '"></div>' +
        '<div class="caja-frase-acciones">' +
          '<button type="button" class="caja-btn-copiar" onclick="copiarFraseCajaManual(\\'' + idUnico + '\\')">📋 Copiar frase</button>' +
          '<div class="caja-frase-estado" id="caja-estado-' + idUnico + '"></div>' +
        '</div>' +
      '</div>';
    }

    function claveBorradorCaja_(idUnico) {
      var cont = document.getElementById('caja-' + idUnico);
      return (cont && cont.getAttribute('data-pend') === '1') ? idUnico + '#pend' : idUnico;
    }

    function abrirCajaManipulada(idUnico, btnEl) {
      var referencia = btnEl ? (btnEl.getAttribute('data-cm-ref') || '') : '';
      var agencia = btnEl ? (btnEl.getAttribute('data-cm-age') || '') : '';
      document.getElementById('cm-head-titulo').textContent = referencia || 'Sin referencia';
      document.getElementById('cm-head-sub').textContent = agencia ? ('Agencia: ' + agencia) : 'Rellena para copiar la frase';
      var pend = !!(btnEl && btnEl.classList && btnEl.classList.contains('qbtn-reclamar-pend')) && typeof botonesBultosCaja_ === 'function';
      var refCab = document.getElementById('cm-head-ref');
      if (refCab) refCab.textContent = pend ? '📦 Bulto manipulado · pendiente de facturar' : '📦 Bulto manipulado';
      document.getElementById('cm-body').innerHTML = generarCajaManipuladaHtml(idUnico, referencia, agencia, pend);
      document.getElementById('cm-overlay').hidden = false;
      if (typeof bocadilloCajaRufo_ === 'function') bocadilloCajaRufo_();

      var borrador = cajaDraftPorFila[claveBorradorCaja_(idUnico)];
      if (borrador) {
        document.getElementById('caja-modo-rotos-' + idUnico).value = borrador.modoRotos;
        document.getElementById('caja-modo-faltan-' + idUnico).value = borrador.modoFaltan;
        document.getElementById('caja-normal-rotos-' + idUnico).value = borrador.normalRotos;
        document.getElementById('caja-mayoria-rotos-' + idUnico).value = borrador.mayoriaRotos;
        document.getElementById('caja-normal-faltan-' + idUnico).value = borrador.normalFaltan;
        document.getElementById('caja-mayoria-faltan-' + idUnico).value = borrador.mayoriaFaltan;
        document.getElementById('caja-modo-rotos-' + idUnico).querySelector('option[value="mayoria"]').disabled = (borrador.modoFaltan === 'mayoria');
        document.getElementById('caja-modo-faltan-' + idUnico).querySelector('option[value="mayoria"]').disabled = (borrador.modoRotos === 'mayoria');
        actualizarVisibilidadLadoCaja(idUnico, 'rotos');
        actualizarVisibilidadLadoCaja(idUnico, 'faltan');
        if (pend && borrador.bultos !== undefined) {
          var contB = document.getElementById('caja-' + idUnico);
          contB.setAttribute('data-bultos', borrador.bultos);
          contB.setAttribute('data-nb', borrador.nb);
          pintarBultosCaja_(idUnico);
        }
        if (borrador.resto) { var cbResto = document.getElementById('caja-resto-' + borrador.resto + '-' + idUnico); if (cbResto) cbResto.checked = true; }
      }

      actualizarFraseCaja(idUnico);
    }

    function cerrarCajaManipulada() {
      document.getElementById('cm-overlay').hidden = true;
      var bocadilloAbierto = document.querySelector('#cm-overlay .cm-bocadillo');
      if (bocadilloAbierto && bocadilloAbierto.parentNode) bocadilloAbierto.parentNode.removeChild(bocadilloAbierto);
    }

    function parsearCodigosCaja(texto) {
      return texto.split(/[\\n,;\\t]+/).map(function(s){ return s.trim(); }).filter(function(s){ return s !== ''; });
    }

    function actualizarVisibilidadLadoCaja(idUnico, lado) {
      var modo = document.getElementById('caja-modo-' + lado + '-' + idUnico).value;
      document.getElementById('caja-wrap-normal-' + lado + '-' + idUnico).classList.toggle('visible', modo === 'normal');
      document.getElementById('caja-wrap-mayoria-' + lado + '-' + idUnico).classList.toggle('visible', modo === 'mayoria');
      document.getElementById('caja-wrap-todo-' + lado + '-' + idUnico).classList.toggle('visible', modo === 'todo');
    }

    function cambiarModoCaja(idUnico, ladoQueCambio) {
      var otroLado = ladoQueCambio === 'rotos' ? 'faltan' : 'rotos';
      var selectEste = document.getElementById('caja-modo-' + ladoQueCambio + '-' + idUnico);
      var selectOtro = document.getElementById('caja-modo-' + otroLado + '-' + idUnico);

      if (selectEste.value === 'mayoria' && selectOtro.value === 'mayoria') {
        selectOtro.value = 'normal';
        actualizarVisibilidadLadoCaja(idUnico, otroLado);
      }
      var selectRotos = document.getElementById('caja-modo-rotos-' + idUnico);
      var selectFaltan = document.getElementById('caja-modo-faltan-' + idUnico);
      selectRotos.querySelector('option[value="mayoria"]').disabled = (selectFaltan.value === 'mayoria');
      selectFaltan.querySelector('option[value="mayoria"]').disabled = (selectRotos.value === 'mayoria');

      actualizarVisibilidadLadoCaja(idUnico, ladoQueCambio);
      actualizarFraseCaja(idUnico);
      flashFrasePreview(idUnico);
    }

    function flashFrasePreview(idUnico) {
      var el = document.getElementById('caja-preview-' + idUnico);
      if (!el) return;
      el.classList.remove('actualizada');
      void el.offsetWidth;
      el.classList.add('actualizada');
    }

    function formatearCodigosCaja_(codigos) {
      var orden = [], conteo = {};
      codigos.forEach(function(c) {
        if (!conteo[c]) { conteo[c] = 0; orden.push(c); }
        conteo[c]++;
      });
      return orden.map(function(c) {
        return conteo[c] > 1 ? (c + '* ' + conteo[c]) : c;
      }).join('-');
    }

    function claseLadoCaja(idUnico, lado, esRotos) {
      var modo = document.getElementById('caja-modo-' + lado + '-' + idUnico).value;
      var codigosNormal = parsearCodigosCaja(document.getElementById('caja-normal-' + lado + '-' + idUnico).value);
      document.getElementById('caja-cont-normal-' + lado + '-' + idUnico).textContent = codigosNormal.length ? (codigosNormal.length + ' número' + (codigosNormal.length === 1 ? '' : 's')) : '';

      var pendCaja = document.getElementById('caja-' + idUnico).getAttribute('data-pend') === '1';
      if (modo === 'todo') return pendCaja ? (esRotos ? 'TODO ROTO' : 'FALTA TODO') : (esRotos ? 'TODO EL PEDIDO ROTO' : 'FALTA TODO EL PEDIDO');
      if (modo === 'mayoria') {
        var codigosOK = parsearCodigosCaja(document.getElementById('caja-mayoria-' + lado + '-' + idUnico).value);
        document.getElementById('caja-cont-mayoria-' + lado + '-' + idUnico).textContent = codigosOK.length ? (codigosOK.length + ' número' + (codigosOK.length === 1 ? '' : 's')) : '';
        var avisoEl = document.getElementById('caja-mayoria-aviso-' + lado + '-' + idUnico);
        if (codigosOK.length === 0) {
          avisoEl.classList.remove('visible');
          return '';
        }
        var trozoMayoria = codigosOK.length + ' ART OK (' + formatearCodigosCaja_(codigosOK) + '), RESTO ' + (pendCaja ? '' : 'DEL PEDIDO ') + (esRotos ? 'ROTO' : 'FALTA');
        document.getElementById('caja-mayoria-aviso-texto-' + lado + '-' + idUnico).textContent = 'Se indicará: ' + trozoMayoria;
        avisoEl.classList.add('visible');
        return trozoMayoria;
      }
      if (codigosNormal.length === 0) return '';
      return esRotos
        ? (codigosNormal.length + ' ART ROTOS: ' + formatearCodigosCaja_(codigosNormal))
        : ('FALTAN ' + codigosNormal.length + ' ART: ' + formatearCodigosCaja_(codigosNormal));
    }

    var cajaCopiaTimers = {};
    var cajaCopiaTokens = {};

    var cajaDraftPorFila = {};

    function guardarBorradorCaja(idUnico) {
      var elModoRotos = document.getElementById('caja-modo-rotos-' + idUnico);
      if (!elModoRotos) return;
      var contBorr = document.getElementById('caja-' + idUnico);
      cajaDraftPorFila[claveBorradorCaja_(idUnico)] = {
        bultos: contBorr ? (contBorr.getAttribute('data-bultos') || '') : '',
        nb: contBorr ? (contBorr.getAttribute('data-nb') || '4') : '4',
        modoRotos: elModoRotos.value,
        modoFaltan: document.getElementById('caja-modo-faltan-' + idUnico).value,
        normalRotos: document.getElementById('caja-normal-rotos-' + idUnico).value,
        mayoriaRotos: document.getElementById('caja-mayoria-rotos-' + idUnico).value,
        normalFaltan: document.getElementById('caja-normal-faltan-' + idUnico).value,
        mayoriaFaltan: document.getElementById('caja-mayoria-faltan-' + idUnico).value,
        resto: typeof restoElegidoCaja_ === 'function' ? restoElegidoCaja_(idUnico) : ''
      };
    }

    function actualizarFraseCaja(idUnico) {
      var cont = document.getElementById('caja-' + idUnico);
      if (!cont) return;
      if (typeof sincronizarRestoCaja_ === 'function') sincronizarRestoCaja_(idUnico);
      guardarBorradorCaja(idUnico);
      var ref = cont.getAttribute('data-ref') || '';
      var age = cont.getAttribute('data-age') || '';
      var partes = [claseLadoCaja(idUnico, 'rotos', true), claseLadoCaja(idUnico, 'faltan', false)].filter(function(s){ return s !== ''; });
      var resto = typeof restoCajaTexto_ === 'function' ? restoCajaTexto_(idUnico) : '';
      var frase = ref + ',' + age + ', BULTO MANIPULADO' + (partes.length ? ', ' + partes.join(' ') : '') + (resto ? ', ' + resto : '');
      if (cont.getAttribute('data-pend') === '1') {
        var textoBultos = textoBultosCaja_(bultosElegidosCaja_(cont));
        frase = ref + ',' + age + ', BULTO MANIPULADO' + (textoBultos ? ', ' + textoBultos : '') + (partes.length ? ', ' + partes.join(' ') + ' DE BULTO QUE LLEGA' : '') + (resto ? ', ' + resto : '');
        if (textoBultos) partes = partes.concat([textoBultos]);
      }
      if (resto) partes = partes.concat([resto]);

      document.getElementById('caja-preview-' + idUnico).textContent = frase;

      if (cajaCopiaTimers[idUnico]) {
        clearTimeout(cajaCopiaTimers[idUnico]);
        cajaCopiaTimers[idUnico] = null;
      }

      if (partes.length > 0) {
        var estadoListo = document.getElementById('caja-estado-' + idUnico);
        estadoListo.className = 'caja-frase-estado';
        estadoListo.textContent = 'Cuando esté completa, pulsa «📋 Copiar frase».';
      } else {
        var estadoEl = document.getElementById('caja-estado-' + idUnico);
        estadoEl.className = 'caja-frase-estado';
        estadoEl.textContent = 'Indica al menos una incidencia para copiar...';
      }
    }

    function marcarEstadoCopiaCaja(idUnico, ok) {
      var estadoEl = document.getElementById('caja-estado-' + idUnico);
      estadoEl.className = 'caja-frase-estado ' + (ok ? 'copiado' : '');
      estadoEl.textContent = ok ? '✅ Copiado al portapapeles' : '⚠️ No se pudo copiar automáticamente';
    }

    function copiarFraseCajaManual(idUnico) {
      var texto = document.getElementById('caja-preview-' + idUnico).textContent;
      copiarAlPortapapeles(texto, function(ok) { marcarEstadoCopiaCaja(idUnico, ok); });
    }

<?!= (LATERALES_UNA_PIEZA_.indexOf(accion) > -1 || accion === 'buscar' || accion === 'retornos') ? jsRetornos_() : '' ?>
<?!= accion === 'recuento' ? jsRecuento_() : '' ?>
<?!= (LATERALES_UNA_PIEZA_.indexOf(accion) > -1 || accion === 'buscar' || accion === 'pendientes') ? jsBultosCaja_() : '' ?>
<?!= (LATERALES_UNA_PIEZA_.indexOf(accion) > -1 || accion === 'buscar' || accion === 'retornos' || accion === 'cambios') ? jsCambiosRet_() : '' ?>
    function saltarACelda(hoja, fila, col, idHtml) {
      var filaEl = document.getElementById(idHtml);
      var btn = filaEl.querySelector('.btn-ir');
      filaEl.classList.add('visitado');

      var clicsGuardados = JSON.parse(localStorage.getItem('clicsRef') || '[]');
      if (clicsGuardados.indexOf(idHtml) === -1) {
        clicsGuardados.push(idHtml);
        localStorage.setItem('clicsRef', JSON.stringify(clicsGuardados));
      }

      var referencia = filaEl.getAttribute('data-ref') || '';
      copiarAlPortapapeles(referencia, function(ok) {
        btn.innerText = (ok && referencia) ? '📋 Ref. copiada ✔️' : '✔️ Revisado';
      });

      google.script.run.activarCeldaEnHoja(hoja, fila, col);
    }

    function alternarRufo() {
      document.getElementById('rufo-colapsado').style.display = 'none';
      document.getElementById('rufo-expandido').style.display = 'block';
    }

    function despertarRufo() {
      var cont = document.getElementById('rufo-colapsado');
      if (cont.classList.contains('despertando')) return; // evita relanzarla si ya se está despertando
      cont.classList.remove('dormido');
      cont.classList.add('despertando');
      setTimeout(alternarRufo, 550);
    }

    var coloresConfetiRufo = ['#E07A1F', '#1B2636', '#e8354f', '#fdfdfd', '#d4af37'];
    var prMarchaUltimo = { btn: null, t: 0 };
    function textoOpcionRufo_(btn) {
      var t = '';
      [].forEach.call(btn.childNodes, function (n) {
        if (n.nodeType === 3) t += n.textContent;
        else if (n.classList && !n.classList.contains('rufo-chip-right') && !n.classList.contains('rufo-badge')) t += n.textContent;
      });
      return t.replace(/\\s+/g, ' ').trim();
    }
    function rufoEnMarcha_(btn) {
      try {
        btn.classList.remove('pr-pulsado'); void btn.offsetWidth; btn.classList.add('pr-pulsado');
        var viejo = document.getElementById('pr-marcha');
        if (viejo && viejo.parentNode) viejo.parentNode.removeChild(viejo);
        var el = document.createElement('div');
        el.id = 'pr-marcha'; el.className = 'pr-marcha'; el.setAttribute('aria-hidden', 'true');
        el.innerHTML = '<div class="pr-marcha-oso"><div class="pr-marcha-salto"><svg viewBox="0 -30 100 148"><use href="#rufo-svg"/>'
          + '<g class="pr-pila-1"><rect x="34" y="-5" width="32" height="16" rx="2.5" fill="#E07A1F"/><rect x="34" y="-5" width="32" height="4" rx="2" fill="#e8354f"/><text x="50" y="8.5" font-family="Georgia, serif" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">P</text></g>'
          + '<g class="pr-pila-2"><rect x="38" y="-20" width="24" height="14" rx="2.5" fill="#a80d26"/><rect x="38" y="-20" width="24" height="3.5" rx="2" fill="#E07A1F"/><text x="50" y="-9" font-family="Georgia, serif" font-size="8" font-weight="bold" fill="#fff" text-anchor="middle">P</text></g>'
          + '<rect x="35" y="82" width="30" height="22" rx="3" fill="#E07A1F"/><rect x="35" y="82" width="30" height="5" rx="2.5" fill="#e8354f"/><text x="50" y="99" font-family="Georgia, serif" font-size="11" font-weight="bold" fill="#fff" text-anchor="middle">P</text>'
          + '<path class="pr-gota-frente" d="M76 22 Q79 27 76 29 Q73 27 76 22 Z" fill="#7dd3fc"/>'
          + '</svg><span class="pr-sudor pr-sudor-i"></span><span class="pr-sudor pr-sudor-i" style="animation-delay:.27s;top:38px"></span>'
          + '<span class="pr-sudor pr-sudor-d" style="animation-delay:.14s"></span><span class="pr-sudor pr-sudor-d" style="animation-delay:.41s;top:36px"></span></div></div>'
          + '<span class="pr-caida pr-caida-2"></span><span class="pr-caida pr-caida-1"></span>'
          + '<span class="pr-marcha-polvo" style="left:14%;animation-delay:.2s"></span><span class="pr-marcha-polvo" style="left:30%;animation-delay:.45s"></span>'
          + '<span class="pr-marcha-polvo" style="left:66%;animation-delay:1.7s"></span><span class="pr-marcha-polvo" style="left:84%;animation-delay:2s"></span>'
          + '<div class="pr-marcha-bocadillo"></div>';
        el.querySelector('.pr-marcha-bocadillo').textContent = '¡Marchando! ' + textoOpcionRufo_(btn);
        document.body.appendChild(el);
        setTimeout(function () { if (el.parentNode) el.parentNode.removeChild(el); }, 2850);
        try { animarBocaRufoHablando(1300); } catch (eBoca) {} // resopla mientras corre
      } catch (e) {}
    }
    document.addEventListener('click', function (ev) {
      var btn = ev.target && ev.target.closest ? ev.target.closest('.rufo-chip, .volver-rufo') : null;
      if (!btn) return;
      var ahora = Date.now();
      if (prMarchaUltimo.btn === btn && ahora - prMarchaUltimo.t < 1500) { ev.stopPropagation(); ev.preventDefault(); return; }
      prMarchaUltimo = { btn: btn, t: ahora };
      rufoEnMarcha_(btn);
    }, true);

    function animarSeleccionRufo(btn) {
      var wrap = btn.querySelector('.mini-caja-wrap');
      var grupo = btn.querySelector('.mini-caja-confeti');
      if (!wrap || !grupo) return;
      grupo.innerHTML = '';

      for (var i = 0; i < 8; i++) {
        var ang = Math.random() * Math.PI - Math.PI / 2 - Math.PI / 4;
        var dist = 10 + Math.random() * 12;
        var dx = Math.cos(ang) * dist;
        var dy = Math.sin(ang) * dist - 8;
        var pieza = document.createElementNS('http://www.w3.org/2000/svg', 'rect');
        var w = 2 + Math.random() * 2;
        pieza.setAttribute('width', w);
        pieza.setAttribute('height', w);
        pieza.setAttribute('rx', .8);
        pieza.setAttribute('x', 17 - w / 2);
        pieza.setAttribute('y', 9 - w / 2);
        pieza.setAttribute('fill', coloresConfetiRufo[i % coloresConfetiRufo.length]);
        pieza.setAttribute('class', 'mini-caja-confeti-pieza');
        pieza.style.animationDelay = (Math.random() * .08) + 's';
        pieza.style.transformOrigin = '17px 9px';
        pieza.style.setProperty('--dx', dx + 'px');
        pieza.style.setProperty('--dy', dy + 'px');
        pieza.style.setProperty('--rot', (Math.random() * 360 - 180) + 'deg');
        grupo.appendChild(pieza);
      }

      wrap.classList.remove('activo');
      void wrap.offsetWidth; // reinicia la animación aunque se pulse varias veces seguidas
      wrap.classList.add('activo');

      clearTimeout(wrap._tOculta);
      wrap._tOculta = setTimeout(function () {
        wrap.classList.remove('activo');
      }, 900);
    }

    var accesoDenegado_ = false;
    function mostrarAccesoExclusivo_() {
      accesoDenegado_ = true;
      document.getElementById('app-content').innerHTML =
        '<div style="text-align:center;padding:34px 16px;">'
        + '<div style="font-size:34px;margin-bottom:10px;">🔒</div>'
        + '<p style="font-size:15px;font-weight:700;margin:0 0 8px;">Herramienta de uso exclusivo del equipo de Devoluciones.</p>'
        + '<p style="font-size:14px;color:var(--muted);margin:0 0 18px;">Por favor, cierre este panel. Gracias.</p>'
        + '<button class="btn" onclick="google.script.host.close()">Cerrar panel</button>'
        + '</div>';
    }
    function esErrorDePermisos_(err) {
      var motivo = (err && err.message) ? err.message : String(err);
      return /autoriz|permis|authoriz|permission|acceso/i.test(motivo);
    }
    function comprobarAccesoRufo_() {
      google.script.run
        .withSuccessHandler(function(ok) { if (ok !== true) mostrarAccesoExclusivo_(); })
        .withFailureHandler(function(err) { if (esErrorDePermisos_(err)) mostrarAccesoExclusivo_(); })
        .comprobarAccesoPanel();
    }

    var acPedido = false;
    function cargarAvisoCompanias_() {
      if (!${(typeof AVISO_COMPANIAS_ACTIVO_ !== 'undefined' && AVISO_COMPANIAS_ACTIVO_) ? 'true' : 'false'}) return;
      if (acPedido) return; acPedido = true;
      google.script.run.withSuccessHandler(pintarAvisoCompanias_).withFailureHandler(function() {}).avisoCompaniasPanel();
    }
    function acDias_(x) { var n = Number(x && x.dias); return (x.dias === null || x.dias === undefined || isNaN(n) || n < 0) ? 'más de 21 días' : n + ' días'; }
    function pintarAvisoCompanias_(lista) {
      var linea = document.getElementById('ac-linea'), chip = document.getElementById('ac-chip');
      var hay = lista && lista.length && !accesoDenegado_;
      if (linea) {
        linea.hidden = !hay;
        if (hay) linea.innerHTML = '⚠️ <b>Más de 7 días sin una entrega de 20 o más:</b>' + lista.map(function(x) {
          return '<span class="ac-ag">' + escAttr(x.c) + ' <small>' + acDias_(x) + (x.ultima ? ' · última ' + escAttr(x.ultima) + ' (' + Number(x.n) + ')' : '') + '</small></span>';
        }).join('');
      }
      var chipCal = document.querySelector('.rufo-chip[onclick*="mCalendario"]');
      if (chipCal) {
        var viejo = chipCal.querySelector('.ac-badge'); if (viejo) viejo.remove();
        if (hay) chipCal.insertAdjacentHTML('beforeend', '<span class="rufo-chip-right"><span class="ac-badge" title="' + escAttr(lista.map(function(x) { return x.c + ': ' + acDias_(x); }).join(' · ') + ' sin una entrega de 20 o más') + '">⚠️ ' + lista.length + '</span></span>');
      }
      if (chip) {
        chip.hidden = true;
        if (false) {
          var partes = lista.map(function(x) { return '<b>' + escAttr(x.c) + '</b> ' + acDias_(x); });
          partes[0] = partes[0].replace('</b> ', '</b> lleva ');
          var frase = partes.length === 1 ? partes[0] : partes.slice(0, -1).join(', ') + ' y ' + partes[partes.length - 1];
          chip.innerHTML = '⚠️ <span>' + frase + ' sin una entrega de 20 o más.</span>';
        }
      }
    }

    function cargarFranjaNotas_() {
      try { cargarAvisoCompanias_(); } catch (eAc) {}
      var cF = cacheLeer_('franja'); if (cF) { try { ntPintarFranja_(cF.d); } catch (eCF) {} }
      google.script.run.withSuccessHandler(function(l) { cacheGuardar_('franja', l); ntPintarFranja_(l); }).withFailureHandler(function() {}).notasFranjaTurno();
    }
    function ntPintarFranja_(lista) {
      var el = document.getElementById('nt-franja');
      if (!el) return;
      if (accesoDenegado_ || !lista || !lista.length) { el.hidden = true; el.innerHTML = ''; return; }
      el.innerHTML = '<div class="nt-franja-t">📋 Del turno anterior <span>' + lista.length + '</span></div>' + lista.map(function(n) {
        return '<div class="nt-mini">' + escAttr(n.t) + '<small>' + escAttr(n.u) + ' · ' + escAttr(n.f) + '</small>' +
          '<button class="nt-ent" onclick="ntEnterada(this, ' + Number(n.id) + ', ' + Number(n.fila) + ')">✓ Enterada</button></div>';
      }).join('');
      el.hidden = false;
    }
    function ntEnterada(btn, id, fila) {
      btn.disabled = true; btn.innerText = '✓ …';
      google.script.run.withSuccessHandler(ntPintarFranja_).withFailureHandler(function() { btn.disabled = false; btn.innerText = '✓ Enterada'; }).marcarNotaEnterada(id, fila);
    }

    function otPestanear(el) {
      if (!el || !el.classList) return;
      el.classList.remove('pestaneo'); void el.getBoundingClientRect(); el.classList.add('pestaneo');
      clearTimeout(el.otT); el.otT = setTimeout(function() { el.classList.remove('pestaneo'); }, 1150);
    }

    var CACHE_V_ = '${VERSION_SISTEMA_}';
    function cacheDia_() { var d = new Date(); return d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate(); }
    function cacheLeer_(k, maxMs) {
      try {
        var o = JSON.parse(localStorage.getItem('pc_' + k) || 'null');
        if (!o || o.dia !== cacheDia_() || o.v !== CACHE_V_) return null;
        if (maxMs && Date.now() - o.t > maxMs) return null;
        return o;
      } catch (e) { return null; }
    }
    function cacheGuardar_(k, d) {
      var txt = '';
      try { txt = JSON.stringify({ v: CACHE_V_, dia: cacheDia_(), t: Date.now(), d: d }); } catch (e) { return; }
      try { localStorage.setItem('pc_' + k, txt); }
      catch (eLleno) { try { cacheLimpiar_(true); localStorage.setItem('pc_' + k, txt); } catch (e2) {} }
    }
    function cacheLimpiar_(todo) {
      try {
        var quitar = [];
        for (var i = 0; i < localStorage.length; i++) {
          var k = localStorage.key(i); if (!k || k.indexOf('pc_') !== 0) continue;
          if (todo) { quitar.push(k); continue; }
          try { var o = JSON.parse(localStorage.getItem(k) || 'null'); if (!o || o.dia !== cacheDia_() || o.v !== CACHE_V_) quitar.push(k); } catch (eP) { quitar.push(k); }
        }
        quitar.forEach(function(k) { localStorage.removeItem(k); });
      } catch (e) {}
    }
    try { setTimeout(function() { cacheLimpiar_(false); }, 3000); } catch (eCl) {}
    function marcaCache_(t) {
      try {
        var el = document.getElementById('marca-cache');
        if (!el) { el = document.createElement('div'); el.id = 'marca-cache'; document.body.appendChild(el); }
        var m = Math.max(0, Math.round((Date.now() - t) / 60000));
        el.textContent = '⏱ ' + (m < 1 ? 'de hace un momento' : 'de hace ' + m + ' min') + ' · actualizando…';
        el.classList.add('ver');
      } catch (e) {}
    }
    function quitarMarcaCache_() { try { var el = document.getElementById('marca-cache'); if (el) el.classList.remove('ver'); } catch (e) {} }

    var SECCIONES_LATERALES_ = ['rufo_abierto', 'buscar', 'retornos', 'cambios', 'pendientes'];
    var secHtml = null, secPidiendo = false;
    function precargarSecciones_() {
      if (secHtml || secPidiendo || (typeof accesoDenegado_ !== 'undefined' && accesoDenegado_)) return;
      var c = cacheLeer_('secciones');
      if (c && c.d) { secHtml = c.d; return; }
      secPidiendo = true;
      google.script.run.withSuccessHandler(function(r) {
        secPidiendo = false;
        if (!r || !r.html) return;
        secHtml = r.html; cacheGuardar_('secciones', r.html);
      }).withFailureHandler(function() { secPidiendo = false; }).obtenerSeccionesLaterales();
    }
    function irSeccion_(accion, titulo) {
      var a = accion === 'rufo' ? 'rufo_abierto' : accion;
      var aqui = document.getElementById('app-content');
      var puede = SECCIONES_LATERALES_.indexOf(a) > -1 && secHtml && secHtml[a] && aqui && !(typeof accesoDenegado_ !== 'undefined' && accesoDenegado_)
        && SECCIONES_LATERALES_.indexOf(accionActual === 'rufo' ? 'rufo_abierto' : accionActual) > -1;
      if (!puede) { google.script.run.abrirMotorLateral(accion, titulo); return; }
      quitarMarcaCache_();
      accionActual = a;
      aqui.innerHTML = secHtml[a];
      try { window.scrollTo(0, 0); } catch (eS) {}
      arrancarSeccion_(a);
    }
    function arrancarSeccion_(a) {
      if (a === 'buscar') { try { initBuscar(); } catch (e1) {} }
      else if (a === 'retornos') { try { rbInit(); } catch (e2) {} }
      else if (a === 'cambios') { try { crInit(); } catch (e3) {} }
      else if (a === 'pendientes') { try { calOffPend = 0; initPendientesLateral(); } catch (e8) {} }
      else if (a === 'rufo_abierto') {
        try { cargarHucho_(); } catch (e4) {}
        try { cargarFranjaNotas_(); } catch (e5) {}
        google.script.run.withSuccessHandler(function(res) {
          if (accionActual !== 'rufo_abierto' || !res || !res.html) return;
          var app = document.getElementById('app-content'); if (!app) return;
          app.innerHTML = res.html; secHtml.rufo_abierto = res.html;
          try { cargarHucho_(); } catch (e6) {}
          try { cargarFranjaNotas_(); } catch (e7) {}
        }).withFailureHandler(function() {}).enrutadorApp('', 'rufo_abierto');
      }
    }

    function cargarHucho_() {
      google.script.run.withSuccessHandler(function(r) {
        if (accesoDenegado_ || !r) return;
        var t = document.getElementById('pcn-hueco-tarjeta'), c = document.getElementById('pcn-hueco-chip');
        if (t && r.tarjeta) t.innerHTML = r.tarjeta;
        if (c && r.chip) c.innerHTML = r.chip;
      }).withFailureHandler(function() {}).huchoTarjetaHtml();
    }

    function errorCargaPanel_(nombreAccion) {
      return function (err) {
        var motivo = (err && err.message) ? err.message : String(err);
        if (nombreAccion === 'Rufo' && esErrorDePermisos_(err)) { mostrarAccesoExclusivo_(); return; }
        document.getElementById('app-content').innerHTML =
          '<div class="loading-row" style="color:#E07A1F;text-align:left;white-space:normal;">'
          + '⚠️ No se pudo cargar ' + nombreAccion + '.<br><br>Motivo: ' + motivo
          + '</div>';
      };
    }

    if (accionActual === 'buscar' && htmlInicial) {
      document.getElementById('app-content').innerHTML = htmlInicial;
      initBuscar();
    } else if (accionActual === 'buscar') {
      document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando buscador...</div>';
      google.script.run.withSuccessHandler(function(res) {
        document.getElementById('app-content').innerHTML = res.html;
        initBuscar();
      }).withFailureHandler(errorCargaPanel_('el buscador')).enrutadorApp('', 'buscar');
    }

    if (accionActual === 'retornos' && htmlInicial) {
      document.getElementById('app-content').innerHTML = htmlInicial;
      rbInit();
    } else if (accionActual === 'retornos') {
      document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando Retornos...</div>';
      google.script.run.withSuccessHandler(function(res) {
        document.getElementById('app-content').innerHTML = res.html;
        rbInit();
      }).withFailureHandler(errorCargaPanel_('Retornos')).enrutadorApp('', 'retornos');
    }

    if (accionActual === 'cambios') {
      if (htmlInicial) { document.getElementById('app-content').innerHTML = htmlInicial; crInit(); }
      else {
        document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando...</div>';
        google.script.run.withSuccessHandler(function(res) {
          document.getElementById('app-content').innerHTML = res.html; crInit();
        }).withFailureHandler(errorCargaPanel_('los cambios de Retornos')).enrutadorApp('', 'cambios');
      }
    }

    if (accionActual === 'pendientes' && cacheLeer_('secciones') && cacheLeer_('secciones').d && cacheLeer_('secciones').d.pendientes) {
      document.getElementById('app-content').innerHTML = cacheLeer_('secciones').d.pendientes;
      initPendientesLateral();
    } else if (accionActual === 'pendientes') {
      document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando...</div>';
      google.script.run.withSuccessHandler(function(res) {
        document.getElementById('app-content').innerHTML = res.html;
        initPendientesLateral();
      }).withFailureHandler(errorCargaPanel_('las casillas pendientes')).enrutadorApp('', 'pendientes');
    }

    if (accionActual === 'rufo') comprobarAccesoRufo_();
    if ((accionActual === 'rufo' || accionActual === 'rufo_abierto') && htmlInicial) {
      document.getElementById('app-content').innerHTML = htmlInicial;
      cargarHucho_();
      cargarFranjaNotas_();
    } else if (accionActual === 'rufo' || accionActual === 'rufo_abierto') {
      var yaDespierto = (accionActual === 'rufo_abierto');
      document.getElementById('app-content').innerHTML = yaDespierto
        ? '<div class="loading-row"><span class="spinner"></span>Cargando Rufo...</div>'
        : '<div class="loading-row"><span class="spinner"></span>Despertando a Rufo...</div>';
      google.script.run.withSuccessHandler(function(res) {
        if (accesoDenegado_) return;
        document.getElementById('app-content').innerHTML = res.html;
        cargarHucho_();
      cargarFranjaNotas_();
      }).withFailureHandler(errorCargaPanel_('Rufo')).enrutadorApp('', accionActual);
    }

    if (accionActual === 'recuento' && htmlInicial) {
      document.getElementById('app-content').innerHTML = htmlInicial;
    } else if (accionActual === 'recuento') {
      var cRc = cacheLeer_('recuento');
      if (cRc && cRc.d) { document.getElementById('app-content').innerHTML = cRc.d; marcaCache_(cRc.t); }
      else document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando recuento...</div>';
      google.script.run.withSuccessHandler(function(res) {
        cacheGuardar_('recuento', res.html); quitarMarcaCache_();
        if (cRc && typeof rcEstado !== 'undefined' && rcEstado.hoyHtml !== null) return; // ya lo estaba usando: no se le cambia
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(cRc ? function() { quitarMarcaCache_(); } : errorCargaPanel_('el recuento')).enrutadorApp('', 'recuento');
    }

    if (accionActual === 'calendario' && htmlInicial) {
      document.getElementById('app-content').innerHTML = htmlInicial;
      var calIni = document.getElementById('cal-datos-ini'), calD = null;
      try { calD = calIni ? JSON.parse(calIni.textContent) : null; } catch (eCalIni) { calD = null; }
      if (calIni) calIni.remove();
      if (calD) { calOff = 0; mostrarCalendario(calD); } else initCalendario();
    } else if (accionActual === 'calendario') {
      document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando calendario...</div>';
      google.script.run.withSuccessHandler(function(res) {
        document.getElementById('app-content').innerHTML = res.html;
        initCalendario();
      }).withFailureHandler(errorCargaPanel_('el calendario')).enrutadorApp('', 'calendario');
    }

    if (SECCIONES_LATERALES_.indexOf(accionActual === 'rufo' ? 'rufo_abierto' : accionActual) > -1) setTimeout(precargarSecciones_, 1500);

    if (accionActual === 'historiales') {
      document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando...</div>';
      google.script.run.withSuccessHandler(function(res) {
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(errorCargaPanel_('los historiales')).enrutadorApp('', 'historiales');
    }

    if (accionActual === 'avisos' && htmlInicial) {
      document.getElementById('app-content').innerHTML = htmlInicial;
    } else if (accionActual === 'avisos') {
      document.getElementById('app-content').innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando avisos...</div>';
      google.script.run.withSuccessHandler(function(res) {
        document.getElementById('app-content').innerHTML = res.html;
      }).withFailureHandler(errorCargaPanel_('los avisos')).enrutadorApp('', 'avisos');
    }

    function proximoRecordatorio_() {
      var ahora = new Date();
      for (var d = 0; d < 8; d++) {
        var dia = new Date(ahora.getFullYear(), ahora.getMonth(), ahora.getDate() + d), horas = [[20, 5, false]];
        if (dia.getDay() === 1) horas.unshift([20, 0, true]);
        for (var i = 0; i < horas.length; i++) {
          var f = new Date(dia); f.setHours(horas[i][0], horas[i][1], 0, 0);
          if (f.getTime() > ahora.getTime()) return { ms: f.getTime() - ahora.getTime(), semanal: horas[i][2] };
        }
      }
      return { ms: 86400000, semanal: false };
    }
    function msHastaProximoRecordatorio_() { return proximoRecordatorio_().ms; }

    function programarRecordatorioChecklist() {
      var proximo = proximoRecordatorio_();
      setTimeout(function() {
        mostrarRecordatorioChecklist(proximo.semanal);
        setTimeout(programarRecordatorioChecklist, 1000);
      }, proximo.ms);
    }

    function mostrarRecordatorioChecklist(semanal) {
      var textoRec = document.getElementById('recordatorio-checklist-texto');
      if (textoRec) textoRec.innerHTML = '<strong>Recordatorio:</strong> envía el checklist ' + (semanal ? '<strong>semanal</strong> de la semana pasada' : 'diario');
      document.getElementById('recordatorio-checklist').hidden = false;
      animarBocaRufoHablando(1400);
    }

    function cerrarRecordatorioChecklist() {
      document.getElementById('recordatorio-checklist').hidden = true;
    }

    var ultimoIdAvisoVisto = null;
    var colaAvisoToast = [];
    var sondeoAvisosActivo = false;
    function iniciarSondeoAvisos() {
      if (!sondeoAvisosActivo) { sondeoAvisosActivo = true; setInterval(comprobarAvisosNuevos_, 20000); }
      google.script.run.withSuccessHandler(function(id) {
        ultimoIdAvisoVisto = id;
      }).withFailureHandler(function() {}).obtenerUltimoIdEventoAviso();
    }
    function comprobarAvisosNuevos_() {
      if (document.hidden) return;
      if (ultimoIdAvisoVisto === null) { iniciarSondeoAvisos(); return; }
      google.script.run.withFailureHandler(function() {}).withSuccessHandler(function(eventos) {
        if (!eventos || eventos.length === 0) return;
        eventos.forEach(function(ev) {
          colaAvisoToast.push(ev);
          if (ev.id > ultimoIdAvisoVisto) ultimoIdAvisoVisto = ev.id;
        });
        var toast = document.getElementById('aviso-toast');
        if (toast && toast.hidden) mostrarSiguienteAvisoToast_();
      }).obtenerNuevosEventosAviso(ultimoIdAvisoVisto);
    }
    function mostrarSiguienteAvisoToast_() {
      var ev = colaAvisoToast.shift();
      if (!ev) return;
      document.getElementById('aviso-toast-titulo').textContent = 'Aviso ' + ev.referencia + ':';
      document.getElementById('aviso-toast-mensaje').textContent = String(ev.motivo || '').trim() || '(sin mensaje)';
      document.getElementById('aviso-toast-donde').textContent = 'Escrito en ' + ev.hoja + ', fila ' + ev.fila + (colaAvisoToast.length ? ' · ' + colaAvisoToast.length + ' aviso' + (colaAvisoToast.length === 1 ? '' : 's') + ' más' : '');
      document.getElementById('aviso-toast').hidden = false;
      animarBocaRufoHablando(1400);
    }
    function cerrarAvisoToast() {
      document.getElementById('aviso-toast').hidden = true;
      if (colaAvisoToast.length > 0) setTimeout(mostrarSiguienteAvisoToast_, 400);
    }
    function irAAvisosDesdeToast() {
      cerrarAvisoToast();
      colaAvisoToast = [];
      google.script.run.mAvisos();
    }

    function irAChecklistDesdeRecordatorio() {
      cerrarRecordatorioChecklist();
      google.script.run.mCalendario();
    }

    function ponerBocaRufo_(abierta) {
      var cerrada = document.getElementById('rufo-boca-cerrada');
      var abiertaEl = document.getElementById('rufo-boca-abierta');
      var lengua = document.getElementById('rufo-lengua');
      if (!cerrada || !abiertaEl) return;
      cerrada.style.display = abierta ? 'none' : '';
      abiertaEl.style.display = abierta ? '' : 'none';
      if (lengua) lengua.style.display = abierta ? '' : 'none';
    }

    function animarBocaRufoHablando(duracionMs) {
      var vecesRestantes = Math.round(duracionMs / 220);
      var abierta = false;
      var intervalo = setInterval(function() {
        abierta = !abierta;
        ponerBocaRufo_(abierta);
        vecesRestantes--;
        if (vecesRestantes <= 0) {
          clearInterval(intervalo);
          ponerBocaRufo_(false);
        }
      }, 220);
    }

    programarRecordatorioChecklist();
    if (${(typeof AVISOS_PEDIDOS_ACTIVOS_ !== 'undefined' && AVISOS_PEDIDOS_ACTIVOS_) ? 'true' : 'false'}) iniciarSondeoAvisos();
  </script></body></html>`;
}
function procesarNuevaNota(_wg, _wh) {
    var _wi = SpreadsheetApp.getActiveSpreadsheet(), _wj = "Usuario anónimo";
    try {
        _wj = Session.getActiveUser().getEmail() || "Usuario";
    }
    catch (_wk) { }
    var _wl = getOrCreateSheet(_wi, HOJA_NOTAS);
    if (_wl.getLastColumn() < 4)
        _wl.getRange("D1").setValue("Estado").setFontWeight("bold");
    var _wm = String(_wh === undefined || _wh === null ? "" : _wh).trim();
    if (_wm) {
        notaAsegurarColumnas_(_wl);
        var _wn = _wl.getLastRow() + 1;
        if (_wn > _wl.getMaxRows())
            _wl.insertRowsAfter(_wl.getMaxRows(), 1);
        _wl.getRange(_wn, NOTAS_COL_PEDIDO_).setNumberFormat('@');
        _wl.getRange(_wn, 1, 1, 5).setValues([[new Date(), _wj, _wg, "NUEVA", _wm]]);
    }
    else
        _wl.appendRow([new Date(), _wj, _wg, "NUEVA"]);
    construirMenu();
}
var NOTAS_COL_PEDIDO_ = 5, NOTAS_COL_RESUELTA_ = 6, NOTAS_COL_ENTERADAS_ = 7;
var NOTAS_HORAS_FRANJA_ = 48, NOTAS_DIAS_PEDIDO_PENDIENTE_ = 60;
function notaQuien_() {
    var _wo = "";
    try {
        _wo = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    }
    catch (_wp) { }
    if (!_wo) {
        try {
            _wo = String(PropertiesService.getUserProperties().getProperty('devoluciones_email') || "").trim().toLowerCase();
        }
        catch (_wq) { }
    }
    return _wo || "usuario";
}
function notaCorto_(_wr) { _wr = String(_wr === null || _wr === undefined ? "" : _wr).trim(); return _wr.indexOf('@') > -1 ? _wr.split('@')[0] : (_wr || '—'); }
function notaEsc_(_ws) { return String(_ws === null || _ws === undefined ? "" : _ws).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
function notaFechaTxt_(_wt, _wu) {
    if (!(_wt instanceof Date) || isNaN(_wt.getTime()))
        return '--/--';
    return ('0' + _wt.getDate()).slice(-2) + '/' + ('0' + (_wt.getMonth() + 1)).slice(-2) + (_wu ? '/' + _wt.getFullYear() : '') + ' ' + ('0' + _wt.getHours()).slice(-2) + ':' + ('0' + _wt.getMinutes()).slice(-2);
}
function notaAsegurarColumnas_(_wv) {
    if (_wv.getMaxColumns() < NOTAS_COL_ENTERADAS_)
        _wv.insertColumnsAfter(_wv.getMaxColumns(), NOTAS_COL_ENTERADAS_ - _wv.getMaxColumns());
    var _ww = _wv.getRange(1, NOTAS_COL_PEDIDO_, 1, 3).getValues()[0];
    if (!_ww[0] && !_ww[1] && !_ww[2])
        _wv.getRange(1, NOTAS_COL_PEDIDO_, 1, 3).setValues([["Pedido", "Resuelta", "Enteradas"]]);
}
function leerNotas_() {
    var _wx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_NOTAS), _wy = { hoja: _wx, filas: [] };
    if (!_wx || _wx.getLastRow() < 2)
        return _wy;
    var _wz = Math.min(NOTAS_COL_ENTERADAS_, _wx.getMaxColumns());
    var _xa = _wx.getRange(2, 1, _wx.getLastRow() - 1, _wz).getValues();
    for (var _xb = 0; _xb < _xa.length; _xb++) {
        var _xc = _xa[_xb], _xd = _xc[0] instanceof Date ? _xc[0] : new Date(_xc[0]);
        if (String(_xc[2] === null || _xc[2] === undefined ? "" : _xc[2]).trim() === "" && !(_xd instanceof Date && !isNaN(_xd.getTime())))
            continue;
        var _xe = _wz >= 5 ? String(_xc[4] === null || _xc[4] === undefined ? "" : _xc[4]).trim() : "";
        var _xf = _wz >= 6 ? String(_xc[5] === null || _xc[5] === undefined ? "" : _xc[5]).trim() : "";
        var _xg = _wz >= 7 ? String(_xc[6] === null || _xc[6] === undefined ? "" : _xc[6]).split(';').map(function (_xh) { return _xh.trim().toLowerCase(); }).filter(function (_xi) { return _xi; }) : [];
        _wy.filas.push({ fila: _xb + 2, id: isNaN(_xd.getTime()) ? 0 : _xd.getTime(), fecha: _xd, usuario: String(_xc[1] === null || _xc[1] === undefined ? "" : _xc[1]), texto: String(_xc[2] === null || _xc[2] === undefined ? "" : _xc[2]), estado: _xc[3], pedido: _xe, pedidoN: _xe ? normalizarRef(_xe) : "", resuelta: _xf, enteradas: _xg });
    }
    return _wy;
}
function notaLocalizar_(_xj, _xk, _xl) {
    _xk = Number(_xk) || 0;
    _xl = Number(_xl) || 0;
    for (var _xm = 0; _xm < _xj.filas.length; _xm++)
        if (_xj.filas[_xm].fila === _xl && _xj.filas[_xm].id === _xk)
            return _xj.filas[_xm];
    for (var _xn = 0; _xn < _xj.filas.length; _xn++)
        if (_xk && _xj.filas[_xn].id === _xk)
            return _xj.filas[_xn];
    return null;
}
function mapaNotasPedido_(_xo) {
    var _xp = {};
    _xo.filas.forEach(function (_xq) {
        if (!_xq.pedidoN || _xq.resuelta)
            return;
        (_xp[_xq.pedidoN] = _xp[_xq.pedidoN] || []).push({ id: _xq.id, fila: _xq.fila, t: _xq.texto, u: notaCorto_(_xq.usuario), f: notaFechaTxt_(_xq.fecha) });
    });
    return _xp;
}
function notaMarcarResuelta_(_xr, _xs, _xt) {
    var _xu = LockService.getDocumentLock(), _xv = false;
    try {
        _xv = _xu.tryLock(8000);
    }
    catch (_xw) { }
    try {
        var _xx = leerNotas_(), _xy = notaLocalizar_(_xx, _xr, _xs);
        if (!_xy)
            return false;
        notaAsegurarColumnas_(_xx.hoja);
        _xx.hoja.getRange(_xy.fila, NOTAS_COL_RESUELTA_).setValue(_xt ? notaFechaTxt_(new Date(), true) + '|' + notaQuien_() : "");
        return true;
    }
    finally {
        if (_xv) {
            try {
                _xu.releaseLock();
            }
            catch (_xz) { }
        }
    }
}
function guardarNotaPedido(_ya, _yb) {
    if (!usuarioAutorizado_())
        return { error: "No tienes permiso para usar este sistema." };
    var _yc = String(_ya || "").trim(), _yd = String(_yb || "").trim();
    if (!_yc || !_yd)
        return { error: "Falta el pedido o la nota." };
    procesarNuevaNota(_yd.substring(0, 2000), _yc.substring(0, 60));
    return { ok: true, notas: mapaNotasPedido_(leerNotas_()) };
}
function resolverNotaPedido(_ye, _yf) {
    if (!usuarioAutorizado_())
        return { error: "No tienes permiso para usar este sistema." };
    var _yg = notaMarcarResuelta_(_ye, _yf, true);
    return { ok: _yg, notas: mapaNotasPedido_(leerNotas_()) };
}
function notasFranjaTurno() {
    if (!usuarioAutorizado_())
        return [];
    var _yh = notaQuien_(), _yi = new Date().getTime() - NOTAS_HORAS_FRANJA_ * 3600000, _yj = [];
    var _yk = leerNotas_();
    for (var _yl = _yk.filas.length - 1; _yl >= 0 && _yj.length < 6; _yl--) {
        var _ym = _yk.filas[_yl];
        if (_ym.pedido || _ym.resuelta || !_ym.id || _ym.id < _yi)
            continue;
        if (_ym.enteradas.indexOf(_yh) > -1)
            continue;
        _yj.push({ id: _ym.id, fila: _ym.fila, t: _ym.texto, u: notaCorto_(_ym.usuario), f: notaFechaTxt_(_ym.fecha) });
    }
    return _yj;
}
function marcarNotaEnterada(_yn, _yo) {
    if (!usuarioAutorizado_())
        return [];
    var _yp = LockService.getDocumentLock(), _yq = false;
    try {
        _yq = _yp.tryLock(8000);
    }
    catch (_yr) { }
    try {
        var _ys = leerNotas_(), _yt = notaLocalizar_(_ys, _yn, _yo), _yu = notaQuien_();
        if (_yt && _yt.enteradas.indexOf(_yu) === -1) {
            notaAsegurarColumnas_(_ys.hoja);
            _ys.hoja.getRange(_yt.fila, NOTAS_COL_ENTERADAS_).setValue(_yt.enteradas.concat([_yu]).join(';'));
        }
    }
    finally {
        if (_yq) {
            try {
                _yp.releaseLock();
            }
            catch (_yv) { }
        }
    }
    return notasFranjaTurno();
}
function eliminarNota(_yw) {
    var _yx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_NOTAS);
    if (_yx) {
        _yx.deleteRow(_yw);
        construirMenu();
    }
}
function editarNota(_yy, _yz) {
    var _za = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_NOTAS);
    if (_za) {
        _za.getRange(_yy, 3).setValue(_yz);
    }
    construirMenu();
}
function guardarYRefrescarNotas(_zb, _zc) {
    procesarNuevaNota(_zb, _zc);
    refrescarSidebarRufo_();
    return { html: genHtmlNotas(false) };
}
function borrarNotaYRefrescarNotas(_zd) {
    eliminarNota(_zd);
    refrescarSidebarRufo_();
    return { html: genHtmlNotas(false) };
}
function editarNotaYRefrescarNotas(_ze, _zf) {
    editarNota(_ze, _zf);
    refrescarSidebarRufo_();
    return { html: genHtmlNotas(false) };
}
function resolverYRefrescarNotas(_zg, _zh, _zi) {
    notaMarcarResuelta_(_zh, _zg, _zi !== false);
    return { html: genHtmlNotas(false) };
}
function refrescarSidebarRufo_() {
    try {
        abrirMotorLateral('rufo_abierto', '🐾 Rufo');
    }
    catch (_zj) { }
}
function procesarNuevoAviso(_zk, _zl) {
    var _zm = SpreadsheetApp.getActiveSpreadsheet(), _zn = "Usuario anónimo";
    try {
        _zn = Session.getActiveUser().getEmail() || "Usuario";
    }
    catch (_zo) { }
    var _zp = getOrCreateSheet(_zm, HOJA_AVISOS);
    var _zq = _zp.getLastRow() + 1;
    if (_zq > _zp.getMaxRows())
        _zp.insertRowsAfter(_zp.getMaxRows(), 1);
    _zp.getRange(_zq, 3).setNumberFormat('@');
    _zp.getRange(_zq, 1, 1, 7).setValues([[new Date(), _zn, (_zk || "").toString().trim(), (_zl || "").toString().trim(), 0, "", ""]]);
}
function eliminarAviso_(_zr) {
    var _zs = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AVISOS);
    if (_zs)
        _zs.deleteRow(_zr);
}
function editarMotivoAviso_(_zt, _zu) {
    var _zv = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AVISOS);
    if (_zv)
        _zv.getRange(_zt, 4).setValue(_zu);
}
function guardarYRefrescarAvisos(_zw, _zx) {
    var _zy = [];
    try {
        _zy = dondeEstaYaEscrita_(_zw);
    }
    catch (_zz) {
        _zy = [];
    }
    procesarNuevoAviso(_zw, _zx);
    var _aaa = genHtmlAvisos(false);
    if (_zy.length > 0) {
        var _aab = _zy.slice(0, 3).map(function (_aac) { return escHtmlAviso_(_aac.hoja) + ", fila " + _aac.fila; }).join(" · ");
        var _aad = '<div class="aviso-ya-escrito">⚠️ <b>' + escHtmlAviso_(String(_zw || "").trim()) + '</b> ya estaba escrito antes de crear el aviso (' + _aab + '). El aviso saltará la próxima vez que se escriba o escanee.</div>';
        _aaa = _aaa.replace('<div class="avisos-lista">', _aad + '<div class="avisos-lista">');
    }
    return { html: _aaa };
}
function borrarAvisoYRefrescarAvisos(_aae) {
    eliminarAviso_(_aae);
    return { html: genHtmlAvisos(false) };
}
function editarAvisoYRefrescarAvisos(_aaf, _aag) {
    editarMotivoAviso_(_aaf, _aag);
    return { html: genHtmlAvisos(false) };
}
function dondeEstaYaEscrita_(_aah) {
    var _aai = normalizarRef(_aah);
    if (_aai.length < 6)
        return [];
    var _aaj = buscarReferenciaExactaRapida_(_aah, "") || [];
    return _aaj.filter(function (_aak) {
        return HOJAS_TRABAJO_PERMITIDAS_.indexOf(_aak.hoja) > -1 && normalizarRef(_aak.matchStr) === _aai;
    });
}
function escHtmlAviso_(_aal) {
    return String(_aal === undefined || _aal === null ? "" : _aal).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function obtenerUltimoIdEventoAviso() { return obtenerUltimoIdEventoAviso_(); }
function obtenerNuevosEventosAviso(_aam) { return obtenerNuevosEventosAviso_(_aam); }
function ultimoIdAvisoGuardado_() {
    try {
        var _aan = PropertiesService.getDocumentProperties().getProperty('AVISOS_ULTIMO_ID');
        if (_aan === null || _aan === "")
            return null;
        var _aao = Number(_aan);
        return isNaN(_aao) ? null : _aao;
    }
    catch (_aap) {
        return null;
    }
}
function guardarUltimoIdAviso_(_aaq) {
    try {
        PropertiesService.getDocumentProperties().setProperty('AVISOS_ULTIMO_ID', String(_aaq));
    }
    catch (_aar) { }
}
function obtenerUltimoIdEventoAviso_() {
    var _aas = ultimoIdAvisoGuardado_();
    if (_aas !== null)
        return _aas;
    var _aat = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AVISOS_EVENTOS);
    if (!_aat || _aat.getLastRow() < 2) {
        guardarUltimoIdAviso_(0);
        return 0;
    }
    var _aau = _aat.getRange(_aat.getLastRow(), 1).getValue();
    var _aav = (typeof _aau === 'number') ? _aau : parseInt(_aau, 10);
    _aav = isNaN(_aav) ? 0 : _aav;
    guardarUltimoIdAviso_(_aav);
    return _aav;
}
function obtenerNuevosEventosAviso_(_aaw) {
    var _aax = ultimoIdAvisoGuardado_();
    if (_aax !== null && _aax <= _aaw)
        return [];
    var _aay = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AVISOS_EVENTOS);
    if (!_aay || _aay.getLastRow() < 2)
        return [];
    var _aaz = _aay.getRange(2, 1, _aay.getLastRow() - 1, 7).getValues();
    var _aba = [];
    for (var _abb = 0; _abb < _aaz.length; _abb++) {
        var _abc = (typeof _aaz[_abb][0] === 'number') ? _aaz[_abb][0] : parseInt(_aaz[_abb][0], 10);
        if (!isNaN(_abc) && _abc > _aaw) {
            _aba.push({ id: _abc, referencia: _aaz[_abb][3], hoja: _aaz[_abb][4], fila: _aaz[_abb][5], motivo: _aaz[_abb][6] });
        }
    }
    if (_aax === null) {
        var _abd = 0;
        for (var _abe = 0; _abe < _aaz.length; _abe++) {
            var _abf = Number(_aaz[_abe][0]);
            if (_abf > _abd)
                _abd = _abf;
        }
        guardarUltimoIdAviso_(_abd);
    }
    return _aba;
}
var TODAS_HOJAS_HISTORIAL_ = [HOJA_RECLAMAR, HOJA_OK, HOJA_AGENCIAS, HOJA_NOTAS, HOJA_NOTAS_ARCHIVO, HOJA_ROTURAS, HOJA_AGENCIAS_ARCHIVO, HOJA_ROTURAS_ARCHIVO, HOJA_RECLAMAR_ARCHIVO, HOJA_OK_ARCHIVO, HOJA_AVISOS, HOJA_AVISOS_EVENTOS, HOJA_CONTROL_REVISION, HOJA_CAMBIOS_RET_, HOJA_COPIA_RET_];
function alternarHistoriales() {
    var _abg = SpreadsheetApp.getActiveSpreadsheet(), _abh = _abg.getSheetByName(HOJA_OK);
    var _abi = (_abh && _abh.isSheetHidden());
    TODAS_HOJAS_HISTORIAL_.forEach(function (_abj) { var _abk = _abg.getSheetByName(_abj); if (_abk) {
        _abi ? _abk.showSheet() : _abk.hideSheet();
    } });
    return _abi ? "👁️ Historiales ahora son VISIBLES" : "🙈 Historiales OCULTADOS";
}
function ocultarHistorialesInterno() { TODAS_HOJAS_HISTORIAL_.forEach(function (_abl) { var _abm = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_abl); if (_abm)
    _abm.hideSheet(); }); }
function archivarNotasAntiguas() {
    var _abn = SpreadsheetApp.getActiveSpreadsheet(), _abo = _abn.getSheetByName(HOJA_NOTAS);
    if (!_abo || _abo.getLastRow() < 2)
        return;
    var _abp = _abo.getLastRow();
    var _abq = Math.max(_abo.getLastColumn(), 4);
    var _abr = _abo.getRange(2, 1, _abp - 1, _abq).getValues();
    var _abs = new Date(), _abt = DIAS_CADUCIDAD_NOTAS * 86400000, _abu = NOTAS_DIAS_PEDIDO_PENDIENTE_ * 86400000;
    var _abv = [], _abw = [];
    _abr.forEach(function (_abx) {
        var _aby = _abs.getTime() - new Date(_abx[0]).getTime();
        var _abz = _abq >= 5 && String(_abx[4] || "").trim() !== "" && !(_abq >= 6 && String(_abx[5] || "").trim() !== "");
        var _aca = _aby > (_abz ? _abu : _abt);
        if (_aca) {
            var _acb = _abx.slice();
            _acb[3] = _acb[3] || "LEÍDA";
            _abw.push(_acb);
        }
        else
            _abv.push(_abx);
    });
    if (_abw.length === 0)
        return;
    var _acc = getOrCreateSheet(_abn, HOJA_NOTAS_ARCHIVO);
    _acc.getRange("D1").setValue("Estado").setFontWeight("bold");
    if (_abq > 4 && _acc.getMaxColumns() < _abq)
        _acc.insertColumnsAfter(_acc.getMaxColumns(), _abq - _acc.getMaxColumns());
    _acc.getRange(_acc.getLastRow() + 1, 1, _abw.length, _abq).setValues(_abw);
    _abo.getRange(2, 1, _abp - 1, _abq).clearContent();
    if (_abv.length > 0)
        _abo.getRange(2, 1, _abv.length, _abq).setValues(_abv);
}
function archivarHistorialesAntiguos() {
    var _acd = SpreadsheetApp.getActiveSpreadsheet();
    var _ace = [
        [HOJA_AGENCIAS, HOJA_AGENCIAS_ARCHIVO],
        [HOJA_ROTURAS, HOJA_ROTURAS_ARCHIVO],
        [HOJA_RECLAMAR, HOJA_RECLAMAR_ARCHIVO],
        [HOJA_OK, HOJA_OK_ARCHIVO]
    ];
    var _acf = new Date(), _acg = DIAS_CADUCIDAD_HISTORIALES * 86400000;
    var _ach = null;
    try {
        _ach = LockService.getDocumentLock();
        if (!_ach || !_ach.tryLock(30000))
            _ach = null;
    }
    catch (_aci) {
        _ach = null;
    }
    try {
        _ace.forEach(function (_acj) {
            var _ack = _acj[0], _acl = _acj[1];
            var _acm = _acd.getSheetByName(_ack);
            if (!_acm)
                return;
            var _acn = _acm.getLastRow();
            if (_acn < 2)
                return;
            var _aco = _acm.getLastColumn();
            var _acp = _acm.getRange(2, 1, _acn - 1, _aco).getValues();
            var _acq = [], _acr = [];
            _acp.forEach(function (_acs) {
                var _act = (_acf.getTime() - new Date(_acs[0]).getTime()) > _acg;
                (_act ? _acr : _acq).push(_acs);
            });
            if (_acr.length === 0)
                return;
            var _acu = getOrCreateSheet(_acd, _acl);
            if (_acu.getMaxColumns() < _aco)
                _acu.insertColumnsAfter(_acu.getMaxColumns(), _aco - _acu.getMaxColumns());
            _acu.getRange(_acu.getLastRow() + 1, 1, _acr.length, _aco).setValues(_acr);
            var _acv = true;
            for (var _acw = 0; _acw < _acr.length; _acw++) {
                if ((_acf.getTime() - new Date(_acp[_acw][0]).getTime()) <= _acg) {
                    _acv = false;
                    break;
                }
            }
            if (_acv) {
                if (_acm.getMaxRows() <= _acn)
                    _acm.insertRowsAfter(_acm.getMaxRows(), 1);
                _acm.deleteRows(2, _acr.length);
            }
            else {
                _acm.getRange(2, 1, _acn - 1, _aco).clearContent();
                if (_acq.length > 0)
                    _acm.getRange(2, 1, _acq.length, _aco).setValues(_acq);
            }
        });
    }
    finally {
        if (_ach) {
            try {
                _ach.releaseLock();
            }
            catch (_acx) { }
        }
    }
    try {
        purgarCambiosRetornos_(SpreadsheetApp.getActiveSpreadsheet());
    }
    catch (_acy) { }
    try {
        limpiarReglasBasuraNoche_();
    }
    catch (_acz) {
        try {
            Logger.log('limpiarReglasBasuraNoche_ ha fallado: ' + (_acz && _acz.message ? _acz.message : _acz));
        }
        catch (_ada) { }
    }
    try {
        actualizarReglaRepetidos_();
    }
    catch (_adb) {
        try {
            Logger.log('actualizarReglaRepetidos_ ha fallado: ' + (_adb && _adb.message ? _adb.message : _adb));
        }
        catch (_adc) { }
    }
}
var CLAVE_COPIA_REGLAS_NOCHE_ = 'RN_COPIA_REGLAS';
function limpiarReglasBasuraNoche_() {
    var _add = SpreadsheetApp.getActiveSpreadsheet(), _ade = [], _adf = { hojas: {} };
    _add.getSheets().forEach(function (_adg) {
        var _adh = _adg.getConditionalFormatRules();
        if (!_adh.length)
            return;
        var _adi = rnReglasQueSobran_(_adg, _adh, HOJAS_TRABAJO_PERMITIDAS_.indexOf(_adg.getName()) > -1);
        if (!_adi.length)
            return;
        _adf.hojas[_adg.getName()] = _adh.map(rnSerializar_);
        _adf.hojas[_adg.getName()].forEach(function (_adj) { rnConstruir_(_adg, _adj); });
        _ade.push({ sh: _adg, nuevas: _adh.filter(function (_adk, _adl) { return _adi.indexOf(_adl) === -1; }), n: _adi.length });
    });
    if (!_ade.length)
        return 0;
    _adf.fecha = Utilities.formatDate(new Date(), _add.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm');
    var _adm = PropertiesService.getDocumentProperties(), _adn = JSON.stringify(_adf), _ado = 8000, _adp = Math.ceil(_adn.length / _ado);
    var _adq = Number(_adm.getProperty(CLAVE_COPIA_REGLAS_NOCHE_ + '_N') || 0);
    for (var _adr = 0; _adr < _adq; _adr++)
        _adm.deleteProperty(CLAVE_COPIA_REGLAS_NOCHE_ + '_' + _adr);
    var _ads = {};
    for (var _adt = 0; _adt < _adp; _adt++)
        _ads[CLAVE_COPIA_REGLAS_NOCHE_ + '_' + _adt] = _adn.substring(_adt * _ado, (_adt + 1) * _ado);
    _ads[CLAVE_COPIA_REGLAS_NOCHE_ + '_N'] = String(_adp);
    _adm.setProperties(_ads);
    var _adu = 0;
    _ade.forEach(function (_adv) { _adv.sh.setConditionalFormatRules(_adv.nuevas); _adu += _adv.n; });
    anotarEventoSalud_("REGLAS", _adu + " regla" + (_adu === 1 ? "" : "s") + " de color que sobraba" + (_adu === 1 ? "" : "n") + " (" + _ade.map(function (_adw) { return _adw.sh.getName() + " " + _adw.n; }).join(", ") + ")");
    return _adu;
}
function restaurarReglasNoche() {
    var _adx = PropertiesService.getDocumentProperties(), _ady = Number(_adx.getProperty(CLAVE_COPIA_REGLAS_NOCHE_ + '_N') || 0), _adz = '';
    if (!_ady) {
        Logger.log('No hay copia de reglas de ninguna limpieza nocturna.');
        return;
    }
    for (var _aea = 0; _aea < _ady; _aea++)
        _adz += _adx.getProperty(CLAVE_COPIA_REGLAS_NOCHE_ + '_' + _aea) || '';
    var _aeb = JSON.parse(_adz), _aec = SpreadsheetApp.getActiveSpreadsheet();
    Object.keys(_aeb.hojas).forEach(function (_aed) {
        var _aee = _aec.getSheetByName(_aed);
        if (!_aee)
            return;
        _aee.setConditionalFormatRules(_aeb.hojas[_aed].map(function (_aef) { return rnConstruir_(_aee, _aef); }));
        Logger.log('«' + _aed + '»: ' + _aeb.hojas[_aed].length + ' reglas restauradas (copia del ' + _aeb.fecha + ')');
    });
}
function rnReglasQueSobran_(_aeg, _aeh, _aei) {
    var _aej = COLUMNA_REFERENCIAS, _aek = 4, _ael = _aeh.map(function (_aem, _aen) {
        var _aeo = _aem.getBooleanCondition(), _aep = _aeo ? String(_aeo.getCriteriaType()) : '', _aeq = _aeo ? (_aeo.getCriteriaValues() || []).map(String).join(' | ') : '';
        var _aer = [], _aes = [], _aet = 0, _aeu = true, _aev = true;
        _aem.getRanges().forEach(function (_aew) {
            var _aex = _aew.getColumn(), _aey = _aex + _aew.getNumColumns() - 1, _aez = _aew.getRow(), _afa = _aez + _aew.getNumRows() - 1;
            _aet += _aew.getNumRows() * _aew.getNumColumns();
            if (_aex <= _aej && _aej <= _aey)
                _aer.push([_aez, _afa]);
            if (_aex <= _aek && _aek <= _aey)
                _aes.push([_aez, _afa]);
            if (!(_aex === _aej && _aey === _aej))
                _aeu = false;
            if (!(_aex === _aek && _aey === _aek))
                _aev = false;
        });
        return { i: _aen, tipo: _aep, f: _aeq, intC: _aer, intD: _aes, celdas: _aet, todoC: _aeu, soloD: _aev,
            rep: _aep === 'CUSTOM_FORMULA' && /(CONTAR\.SI|COUNTIF)\s*\(/i.test(_aeq) && />\s*1\s*\)?\s*$/.test(_aeq),
            miraC: /(CONTAR\.SI|COUNTIF)\s*\(\s*C:C\s*[,;]/i.test(_aeq), celdasC: _aer.reduce(function (_afb, _afc) { return _afb + _afc[1] - _afc[0] + 1; }, 0) };
    });
    var _afd = {};
    _ael.forEach(function (_afe) { if (_afe.f.indexOf('#REF!') > -1)
        _afd[_afe.i] = true; });
    if (_aei) {
        var _aff = _ael.filter(function (_afg) { return !_afd[_afg.i] && _afg.rep; });
        var _afh = _aff.filter(function (_afi) { return _afi.celdasC > 0 && _afi.miraC; }).sort(function (_afj, _afk) {
            if (_aeg.getName() === 'Entradas' && _afj.todoC !== _afk.todoC)
                return _afj.todoC ? -1 : 1;
            return _afk.celdasC - _afj.celdasC;
        });
        var _afl = _afh[0], _afm = _afl ? rnUnir_(_afl.intC) : [];
        _aff.forEach(function (_afn) {
            if (_afn === _afl)
                return;
            if (!_afn.celdasC) {
                _afd[_afn.i] = true;
                return;
            }
            if (!_afl)
                return;
            if (!rnRestar_(_afn.intC, _afm).length || _aeg.getName() === 'Entradas')
                _afd[_afn.i] = true;
        });
        var _afo = {};
        _ael.forEach(function (_afp) { if (!_afd[_afp.i] && _afp.tipo && _afp.tipo !== 'CUSTOM_FORMULA')
            (_afo[_afp.tipo + '|' + _afp.f] = _afo[_afp.tipo + '|' + _afp.f] || []).push(_afp); });
        Object.keys(_afo).forEach(function (_afq) {
            var _afr = _afo[_afq];
            if (_afr.length < 2)
                return;
            _afr.sort(function (_afs, _aft) { return _aft.celdas - _afs.celdas; });
            var _afu = rnUnir_(_afr[0].intD);
            _afr.slice(1).forEach(function (_afv) { if (_afv.soloD && !rnRestar_(_afv.intD, _afu).length)
                _afd[_afv.i] = true; });
        });
    }
    return Object.keys(_afd).map(Number);
}
function rnUnir_(_afw) {
    var _afx = _afw.slice().sort(function (_afy, _afz) { return _afy[0] - _afz[0]; }), _aga = [];
    _afx.forEach(function (_agb) { var _agc = _aga[_aga.length - 1]; if (_agc && _agb[0] <= _agc[1] + 1)
        _agc[1] = Math.max(_agc[1], _agb[1]);
    else
        _aga.push([_agb[0], _agb[1]]); });
    return _aga;
}
function rnRestar_(_agd, _age) {
    var _agf = [];
    rnUnir_(_agd).forEach(function (_agg) {
        var _agh = _agg[0];
        for (var _agi = 0; _agi < _age.length && _agh <= _agg[1]; _agi++) {
            var _agj = _age[_agi];
            if (_agj[1] < _agh)
                continue;
            if (_agj[0] > _agg[1])
                break;
            if (_agj[0] > _agh)
                _agf.push([_agh, _agj[0] - 1]);
            _agh = Math.max(_agh, _agj[1] + 1);
        }
        if (_agh <= _agg[1])
            _agf.push([_agh, _agg[1]]);
    });
    return _agf;
}
function rnColor_(_agk, _agl) {
    try {
        if (_agk && _agk.asRgbColor)
            return _agk.asRgbColor().asHexString();
    }
    catch (_agm) { }
    try {
        return _agl ? _agl() : null;
    }
    catch (_agn) {
        return null;
    }
}
function rnSerializar_(_ago) {
    var _agp = { g: _ago.getRanges().map(function (_agq) { return _agq.getA1Notation(); }) }, _agr = _ago.getBooleanCondition();
    if (_agr) {
        _agp.t = String(_agr.getCriteriaType());
        _agp.v = (_agr.getCriteriaValues() || []).map(function (_ags) {
            if (_ags === null || typeof _ags === 'string' || typeof _ags === 'number' || typeof _ags === 'boolean')
                return _ags;
            if (_ags instanceof Date)
                return { d: _ags.getTime() };
            return { e: String(_ags) };
        });
        _agp.bg = rnColor_(_agr.getBackgroundObject && _agr.getBackgroundObject(), _agr.getBackground && function () { return _agr.getBackground(); });
        _agp.fc = rnColor_(_agr.getFontColorObject && _agr.getFontColorObject(), _agr.getFontColor && function () { return _agr.getFontColor(); });
        _agp.b = _agr.getBold();
        _agp.i = _agr.getItalic();
        _agp.s = _agr.getStrikethrough();
        _agp.u = _agr.getUnderline();
        return _agp;
    }
    var _agt = _ago.getGradientCondition();
    if (!_agt)
        throw new Error('regla de un tipo desconocido en ' + _agp.g.join(','));
    _agp.t = 'GRADIENTE';
    _agp.min = [rnColor_(_agt.getMinColorObject && _agt.getMinColorObject(), _agt.getMinColor && function () { return _agt.getMinColor(); }), String(_agt.getMinType()), _agt.getMinValue()];
    _agp.max = [rnColor_(_agt.getMaxColorObject && _agt.getMaxColorObject(), _agt.getMaxColor && function () { return _agt.getMaxColor(); }), String(_agt.getMaxType()), _agt.getMaxValue()];
    if (_agt.getMidType && _agt.getMidType())
        _agp.mid = [rnColor_(_agt.getMidColorObject && _agt.getMidColorObject(), _agt.getMidColor && function () { return _agt.getMidColor(); }), String(_agt.getMidType()), _agt.getMidValue()];
    return _agp;
}
function rnConstruir_(_agu, _agv) {
    var _agw = SpreadsheetApp.newConditionalFormatRule().setRanges(_agv.g.map(function (_agx) { return _agu.getRange(_agx); }));
    if (_agv.t === 'GRADIENTE') {
        var _agy = function (_agz, _aha, _ahb) {
            if (!_agz)
                return;
            if (_agz[1] === 'MIN' || _agz[1] === 'MAX')
                _ahb.call(_agw, _agz[0]);
            else
                _aha.call(_agw, _agz[0], SpreadsheetApp.InterpolationType[_agz[1]], String(_agz[2]));
        };
        _agy(_agv.min, _agw.setGradientMinpointWithValue, _agw.setGradientMinpoint);
        _agy(_agv.mid, _agw.setGradientMidpointWithValue, function () { });
        _agy(_agv.max, _agw.setGradientMaxpointWithValue, _agw.setGradientMaxpoint);
    }
    else {
        _agw.withCriteria(SpreadsheetApp.BooleanCriteria[_agv.t], _agv.v.map(function (_ahc) {
            if (_ahc && typeof _ahc === 'object' && _ahc.d !== undefined)
                return new Date(_ahc.d);
            if (_ahc && typeof _ahc === 'object' && _ahc.e !== undefined)
                return SpreadsheetApp.RelativeDate[_ahc.e];
            return _ahc;
        }));
        if (_agv.bg)
            _agw.setBackground(_agv.bg);
        if (_agv.fc)
            _agw.setFontColor(_agv.fc);
        if (_agv.b !== null && _agv.b !== undefined)
            _agw.setBold(_agv.b);
        if (_agv.i !== null && _agv.i !== undefined)
            _agw.setItalic(_agv.i);
        if (_agv.s !== null && _agv.s !== undefined)
            _agw.setStrikethrough(_agv.s);
        if (_agv.u !== null && _agv.u !== undefined)
            _agw.setUnderline(_agv.u);
    }
    return _agw.build();
}
var FILAS_VENTANA_REPETIDOS_ = 1500;
function esReglaRepetidos_(_ahd) {
    return /(CONTAR\.SI|COUNTIF)\(\s*C:C\s*[,;]/i.test(String(_ahd || ""));
}
function actualizarReglaRepetidos_() {
    var _ahe = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Entradas");
    if (!_ahe)
        return;
    var _ahf = COLUMNA_REFERENCIAS, _ahg = columnToLetter(_ahf);
    var _ahh = Math.max(2, _ahe.getLastRow() - FILAS_VENTANA_REPETIDOS_);
    var _ahi = _ahe.getMaxRows();
    var _ahj = _ahg + _ahh;
    var _ahk = '=IF(' + _ahj + '="",FALSE,COUNTIF(' + _ahg + ':' + _ahg + ',' + _ahj + ')>1)';
    var _ahl = _ahe.getRange(_ahh, _ahf, _ahi - _ahh + 1, 1);
    var _ahm = _ahe.getConditionalFormatRules();
    var _ahn = [], _aho = -1, _ahp = null;
    for (var _ahq = 0; _ahq < _ahm.length; _ahq++) {
        var _ahr = _ahm[_ahq], _ahs = _ahr.getBooleanCondition();
        var _aht = _ahr.getRanges().every(function (_ahu) { return _ahu.getColumn() === _ahf && _ahu.getNumColumns() === 1; });
        if (_ahs && _aht && _ahs.getCriteriaType() === SpreadsheetApp.BooleanCriteria.CUSTOM_FORMULA &&
            esReglaRepetidos_((_ahs.getCriteriaValues() || [])[0])) {
            if (_aho === -1) {
                _aho = _ahn.length;
                _ahp = _ahr;
                _ahn.push(null);
            }
            continue;
        }
        _ahn.push(_ahr);
    }
    var _ahv = _ahp ? _ahp.copy() : SpreadsheetApp.newConditionalFormatRule().setBackground("#ff0000");
    var _ahw = _ahv.whenFormulaSatisfied(_ahk).setRanges([_ahl]).build();
    if (_aho === -1)
        _ahn.push(_ahw);
    else
        _ahn[_aho] = _ahw;
    _ahe.setConditionalFormatRules(_ahn);
}
var FILAS_REVISION_ = 2500;
var HORAS_MAX_FOTO_REVISION_ = 24;
var USUARIO_REVISION_ = "Revisión automática";
function revisarRegistrosAutomatico() {
    try {
        var _ahx = revisarRegistros_(20000);
        if (_ahx && !_ahx.ocupado)
            PropertiesService.getDocumentProperties().setProperty('REVISION_ULTIMA_AUTO', String(new Date().getTime()));
    }
    catch (_ahy) {
        try {
            Logger.log('Revisión automática: ' + (_ahy && _ahy.message ? _ahy.message : _ahy));
        }
        catch (_ahz) { }
    }
    try {
        sincronizarCopiaRetornos_();
    }
    catch (_aia) {
        try {
            Logger.log('Copia de Retornos: ' + (_aia && _aia.message ? _aia.message : _aia));
        }
        catch (_aib) { }
    }
}
function revisarRegistrosAhora() {
    if (!usuarioAutorizadoOAvisar_())
        return;
    if (!soloAdminOAvisar_())
        return;
    var _aic = SpreadsheetApp.getUi();
    var _aid;
    try {
        _aid = revisarRegistros_(30000);
    }
    catch (_aie) {
        _aic.alert("⚠️ No se ha podido revisar: " + (_aie && _aie.message ? _aie.message : _aie));
        return;
    }
    if (_aid.ocupado) {
        _aic.alert("⏳ Ahora mismo ya se está haciendo una revisión. Prueba de nuevo en un minuto.");
        return;
    }
    _aic.alert("✅ Comprobar que todo está contado", textoResultadoRevision_(_aid), _aic.ButtonSet.OK);
}
function textoResultadoRevision_(_aif) {
    var _aig = [];
    if (_aif.primeraVez) {
        _aig.push("📸 Primera revisión: se han guardado las últimas " + _aif.filas + " filas de Entradas y Retornos.");
        _aig.push("A partir de ahora, cualquier cambio que no llegue a registrarse se añadirá solo en la siguiente revisión.");
    }
    else if (_aif.total === 0) {
        _aig.push("✅ Todo está contado. Revisadas las últimas " + _aif.filas + " filas de Entradas y Retornos: no faltaba nada.");
    }
    else {
        _aig.push("🔧 Se ha corregido lo que faltaba (revisadas " + _aif.filas + " filas):");
        if (_aif.agencias)
            _aig.push("• Agencias añadidas: " + _aif.agencias);
        if (_aif.ok)
            _aig.push("• OK añadidos: " + _aif.ok);
        if (_aif.reclamar)
            _aig.push("• Reclamar añadidos: " + _aif.reclamar);
        if (_aif.roturas)
            _aig.push("• Roturas añadidas: " + _aif.roturas);
        if (_aif.quitados)
            _aig.push("• Registros quitados (la celda ya estaba vacía): " + _aif.quitados);
        _aig.push("Ya salen bien en el Calendario y en el Recuento.");
    }
    var _aih = 0;
    try {
        _aih = Number(PropertiesService.getDocumentProperties().getProperty('REVISION_ULTIMA_AUTO') || 0);
    }
    catch (_aii) { }
    var _aij = _aih ? Math.round((new Date().getTime() - _aih) / 60000) : null;
    _aig.push("");
    if (_aij !== null && _aij <= 15)
        _aig.push("🔄 La revisión automática está activa (la última fue hace " + _aij + " min).");
    else
        _aig.push("⚠️ La revisión automática cada 5 minutos no está activa. Para activarla, pulsa \"🕰️ Auto-Archivado\" en este mismo menú (solo una vez).");
    return _aig.join("\n");
}
function revisarRegistros_(_aik) {
    PLAZO_REGISTRO_ = 0;
    var _ail = null;
    try {
        _ail = LockService.getScriptLock();
    }
    catch (_aim) {
        _ail = null;
    }
    if (!_ail || !_ail.tryLock(_aik || 20000))
        return { ocupado: true };
    try {
        return revisarRegistrosSinBloqueo_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
    }
    finally {
        try {
            _ail.releaseLock();
        }
        catch (_ain) { }
    }
}
function leerFilasRecientesRevision_(_aio, _aip) {
    var _aiq = [];
    var _air = _aio.getLastRow();
    if (_air < 2)
        return _aiq;
    var _ais = Math.max(2, _air - FILAS_REVISION_ - 1000 + 1);
    var _ait = Math.min(COLUMNA_NOTA_L - COLUMNA_AGENCIAS + 1, _aio.getMaxColumns() - COLUMNA_AGENCIAS + 1);
    var _aiu = _aio.getRange(_ais, COLUMNA_AGENCIAS, _air - _ais + 1, _ait).getDisplayValues();
    var _aiv = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _aiw = COLUMNA_ROTURA - COLUMNA_AGENCIAS, _aix = COLUMNA_NOTA_L - COLUMNA_AGENCIAS;
    function _aiy(_aiz) { return (_aiz === undefined || _aiz === null) ? "" : String(_aiz).trim(); }
    var _aja = _aiu.length - 1;
    while (_aja >= 0 && !/\d/.test(_aiy(_aiu[_aja][_aiv])) && _aiy(_aiu[_aja][0]) === "")
        _aja--;
    if (_aja < 0)
        return _aiq;
    var _ajb = Math.max(0, _aja - FILAS_REVISION_ + 1);
    for (var _ajc = _ajb; _ajc <= _aja; _ajc++) {
        var _ajd = _aiy(_aiu[_ajc][_aiv]), _aje = _aiy(_aiu[_ajc][0]), _ajf = _aip ? "" : _aiy(_aiu[_ajc][_aix]);
        _aiq.push({
            fila: _ais + _ajc, refMostrada: _ajd, ref: normalizarRef(_ajd),
            bRaw: _aje, b: _aje.toUpperCase(),
            lRaw: _ajf, l: _ajf.toUpperCase(),
            e: _aip ? "" : _aiy(_aiu[_ajc][_aiw]).toUpperCase()
        });
    }
    return _aiq;
}
function emparejarConFoto_(_ajg, _ajh) {
    var _aji = {}, _ajj = {}, _ajk = 0;
    _ajh.forEach(function (_ajl) {
        _aji[_ajl.fila] = _ajl;
        if (_ajl.ref)
            (_ajj[_ajl.ref] = _ajj[_ajl.ref] || []).push(_ajl);
        if (_ajl.fila > _ajk)
            _ajk = _ajl.fila;
    });
    _ajg.forEach(function (_ajm) {
        var _ajn = _aji[_ajm.fila];
        if (_ajn && !_ajn.usado && _ajn.ref === _ajm.ref) {
            _ajm.prev = _ajn;
            _ajn.usado = true;
        }
    });
    _ajg.forEach(function (_ajo) {
        if (_ajo.prev || !_ajo.ref || !_ajj[_ajo.ref])
            return;
        var _ajp = null;
        _ajj[_ajo.ref].forEach(function (_ajq) {
            if (!_ajq.usado && (!_ajp || Math.abs(_ajq.fila - _ajo.fila) < Math.abs(_ajp.fila - _ajo.fila)))
                _ajp = _ajq;
        });
        if (_ajp) {
            _ajo.prev = _ajp;
            _ajp.usado = true;
        }
    });
    _ajg.forEach(function (_ajr) {
        if (_ajr.prev)
            return;
        if (_ajr.fila > _ajk) {
            _ajr.prev = { fila: null, ref: "", b: "", l: "", e: "" };
            return;
        }
        var _ajs = _aji[_ajr.fila];
        if (_ajs && !_ajs.usado) {
            _ajr.prev = _ajs;
            _ajs.usado = true;
        }
    });
}
function registroCoincide_(_ajt, _aju, _ajv, _ajw) {
    return _aju.some(function (_ajx) {
        var _ajy = _ajx ? _ajt[_ajx] : null;
        if (!_ajy)
            return false;
        if (_ajv !== null && _ajy.valor !== _ajv)
            return false;
        return !_ajy.ref || !_ajw || _ajy.ref === _ajw;
    });
}
function revisarRegistrosSinBloqueo_(_ajz, _aka) {
    try {
        codificarRefsExistentes_(_ajz);
    }
    catch (_akb) { }
    var _akc = PropertiesService.getDocumentProperties();
    var _akd = Number(_akc.getProperty('REVISION_FOTO_T') || 0);
    var _ake = getOrCreateSheet(_ajz, HOJA_CONTROL_REVISION);
    var _akf = HOJAS_TRABAJO_PERMITIDAS_.length * FILAS_REVISION_;
    var _akg = FILAS_REVISION_ + '|' + HOJAS_TRABAJO_PERMITIDAS_.join(',');
    var _akh = _akc.getProperty('REVISION_FOTO_ANILLO') === _akg;
    var _aki = Math.max(_ake.getLastRow() - 1, 0);
    var _akj = _aki > 0 ? _ake.getRange(2, 1, _aki, 6).getValues() : [];
    for (var _akk = 0; _akk < _akj.length; _akk++)
        _akj[_akk][2] = decodificarRef_(_akj[_akk][2]);
    var _akl = _akd > 0 && (_aka.getTime() - _akd) <= HORAS_MAX_FOTO_REVISION_ * 3600000;
    var _akm = { primeraVez: !_akl, filas: 0, agencias: 0, ok: 0, reclamar: 0, roturas: 0, quitados: 0, total: 0 };
    var _akn = [];
    var _ako = [];
    var _akp = _aka, _akq = USUARIO_REVISION_;
    HOJAS_TRABAJO_PERMITIDAS_.forEach(function (_akr) {
        var _aks = _ajz.getSheetByName(_akr);
        if (!_aks)
            return;
        var _akt = (_akr === HOJA_RETORNOS);
        var _aku = leerFilasRecientesRevision_(_aks, _akt);
        _akm.filas += _aku.length;
        _aku.forEach(function (_akv) { _ako.push([_akr, String(_akv.fila), _akv.ref, _akv.b, _akv.l, _akv.e]); });
        if (!_akl)
            return;
        var _akw = _akj.filter(function (_akx) { return String(_akx[0]) === _akr; }).map(function (_aky) {
            return { fila: Number(_aky[1]), ref: String(_aky[2]), b: String(_aky[3]), l: String(_aky[4]), e: String(_aky[5]), usado: false };
        });
        if (_akw.length === 0)
            return;
        emparejarConFoto_(_aku, _akw);
        var _akz = [], _ala = [], _alb = [], _alc = [], _ald = [], _ale = [];
        _aku.forEach(function (_alf) {
            var _alg = _alf.prev;
            if (!_alg)
                return;
            var _alh = (_alg.fila === _alf.fila);
            if (_alf.b !== "" && _alf.b !== _alg.b)
                _akz.push(_alf);
            if (_alf.b === "" && _alg.b !== "" && _alh)
                _ala.push(_alf.fila);
            if (_akt)
                return;
            var _ali = categoriaEstado_(_alf.l), _alj = categoriaEstado_(_alg.l);
            if (_ali && _ali !== _alj)
                _alb.push(_alf);
            if (_alf.l === "" && _alj && _alh)
                _alc.push(_alf.fila);
            if (_alf.e === "SI" && _alg.e !== "SI")
                _ald.push(_alf);
            if (_alf.e !== "SI" && _alg.e === "SI" && _alh)
                _ale.push(_alf.fila);
        });
        if (_akz.length > 0) {
            var _alk = ultimosRegistrosPorFila_(_ajz, HOJA_AGENCIAS, _akr, 3, 4, 5);
            var _all = [];
            _akz.forEach(function (_alm) {
                if (registroCoincide_(_alk, [_alm.fila, _alm.prev.fila], _alm.b, _alm.ref))
                    return;
                _all.push([_akp, _akq, _akr, _alm.bRaw, _alm.fila, refParaHistorial_(_alm.refMostrada)]);
            });
            if (_all.length > 0) {
                anexarFilasHistorial_(_ajz, HOJA_AGENCIAS, _all, 6);
                _akm.agencias += _all.length;
                _all.forEach(function (_aln) { _akn.push({ o: _akr, f: _aln[4] }); });
            }
        }
        if (_alb.length > 0) {
            var _alo = ultimosRegistrosPorFila_(_ajz, HOJA_OK, _akr, 3, 6, 8);
            var _alp = ultimosRegistrosPorFila_(_ajz, HOJA_RECLAMAR, _akr, 3, 6, 8);
            var _alq = [], _alr = [];
            _alb.forEach(function (_als) {
                var _alt = categoriaEstado_(_als.l);
                var _alu = [_als.fila, _als.prev.fila].some(function (_alv) {
                    if (!_alv)
                        return false;
                    var _alw = ultimoEstadoDeFila_(_alo, _alp, _alv);
                    return !!_alw && _alw.cat === _alt && (!_alw.ref || !_als.ref || _alw.ref === _als.ref);
                });
                if (_alu)
                    return;
                (_alt === "OK" ? _alq : _alr).push([_akp, _akq, _akr, _als.lRaw, COLUMNA_NOTA_L, columnToLetter(COLUMNA_NOTA_L), _als.fila, _als.bRaw, refParaHistorial_(_als.refMostrada)]);
            });
            if (_alq.length > 0) {
                anexarFilasHistorial_(_ajz, HOJA_OK, _alq, 9);
                _akm.ok += _alq.length;
            }
            if (_alr.length > 0) {
                anexarFilasHistorial_(_ajz, HOJA_RECLAMAR, _alr, 9);
                _akm.reclamar += _alr.length;
            }
            _alq.concat(_alr).forEach(function (_alx) { _akn.push({ o: _akr, f: _alx[6] }); });
        }
        if (_ald.length > 0) {
            var _aly = ultimosRegistrosPorFila_(_ajz, HOJA_ROTURAS, _akr, 3, 4, 5);
            var _alz = [];
            _ald.forEach(function (_ama) {
                if (registroCoincide_(_aly, [_ama.fila, _ama.prev.fila], null, _ama.ref))
                    return;
                _alz.push([_akp, _akq, _akr, _ama.bRaw, _ama.fila, refParaHistorial_(_ama.refMostrada)]);
            });
            if (_alz.length > 0) {
                anexarFilasHistorial_(_ajz, HOJA_ROTURAS, _alz, 6);
                _akm.roturas += _alz.length;
                _alz.forEach(function (_amb) { _akn.push({ o: _akr, f: _amb[4] }); });
            }
        }
        if (_ala.length > 0)
            _akm.quitados += quitarRegistrosDeFilas_(_ajz, HOJA_AGENCIAS, _akr, _ala, 2, 4);
        if (_alc.length > 0) {
            _akm.quitados += quitarRegistrosDeFilas_(_ajz, HOJA_OK, _akr, _alc, 2, 6);
            _akm.quitados += quitarRegistrosDeFilas_(_ajz, HOJA_RECLAMAR, _akr, _alc, 2, 6);
        }
        if (_ale.length > 0)
            _akm.quitados += quitarRegistrosDeFilas_(_ajz, HOJA_ROTURAS, _akr, _ale, 2, 4);
    });
    _akm.total = _akm.agencias + _akm.ok + _akm.reclamar + _akm.roturas + _akm.quitados;
    if (_akm.total > 0)
        sumarContador_("CORRECCIONES", _akm.total);
    if (_akm.total > 0) {
        var _amc = [];
        if (_akm.agencias)
            _amc.push(_akm.agencias + " agencia" + (_akm.agencias === 1 ? "" : "s"));
        if (_akm.ok)
            _amc.push(_akm.ok + " OK");
        if (_akm.reclamar)
            _amc.push(_akm.reclamar + " Reclamar");
        if (_akm.roturas)
            _amc.push(_akm.roturas + " rotura" + (_akm.roturas === 1 ? "" : "s"));
        if (_akm.quitados)
            _amc.push(_akm.quitados + " quitado" + (_akm.quitados === 1 ? "" : "s"));
        var _amd = {}, _ame = [];
        _akn.forEach(function (_amf) { if (!_amd[_amf.o]) {
            _amd[_amf.o] = [];
            _ame.push(_amf.o);
        } if (_amd[_amf.o].indexOf(_amf.f) === -1)
            _amd[_amf.o].push(_amf.f); });
        var _amg = _ame.map(function (_amh) { var _ami = _amd[_amh]; return _amh + ": fila" + (_ami.length === 1 ? " " : "s ") + _ami.slice(0, 8).join(", ") + (_ami.length > 8 ? "…" : ""); }).join(" · ");
        anotarEventoSalud_("CORR", _amc.join(", ") + (_amg ? " · " + _amg : ""));
    }
    var _amj = false;
    var _amk = [];
    for (var _aml = 0; _aml < _akf; _aml++)
        _amk.push(null);
    _ako.forEach(function (_amm) { var _amn = HOJAS_TRABAJO_PERMITIDAS_.indexOf(_amm[0]); if (_amn < 0)
        return; _amk[_amn * FILAS_REVISION_ + (Number(_amm[1]) % FILAS_REVISION_)] = _amm; });
    var _amo = ["", "", "", "", "", ""];
    function _amp(_amq) { return _amq ? [_amq[0], _amq[1], codificarRef_(_amq[2]), _amq[3], _amq[4], _amq[5]] : _amo; }
    if (!_akh) {
        if (_ake.getMaxRows() < _akf + 1)
            _ake.insertRowsAfter(_ake.getMaxRows(), _akf + 1 - _ake.getMaxRows());
        var _amr = _ake.getRange(2, 1, _akf, 6);
        _amr.setNumberFormat('@');
        _amr.setValues(_amk.map(_amp));
        if (_aki > _akf)
            _ake.getRange(_akf + 2, 1, _aki - _akf, 6).clearContent();
        _akc.setProperty('REVISION_FOTO_ANILLO', _akg);
        _amj = true;
    }
    else {
        var _ams = [];
        for (var _amt = 0; _amt < _akf; _amt++) {
            var _amu = _amk[_amt] || _amo, _amv = _akj[_amt] || _amo;
            for (var _amw = 0; _amw < 6; _amw++) {
                var _amx = _amv[_amw];
                if (String(_amx === undefined || _amx === null ? "" : _amx) !== String(_amu[_amw])) {
                    _ams.push(_amt);
                    break;
                }
            }
        }
        if (_ams.length > 0) {
            var _amy = [], _amz = _ams[0], _ana = _ams[0];
            for (var _anb = 1; _anb < _ams.length; _anb++) {
                if (_ams[_anb] - _ana <= 6)
                    _ana = _ams[_anb];
                else {
                    _amy.push([_amz, _ana]);
                    _amz = _ana = _ams[_anb];
                }
            }
            _amy.push([_amz, _ana]);
            if (_amy.length > 25)
                _amy = [[_ams[0], _ams[_ams.length - 1]]];
            _amy.forEach(function (_anc) {
                var _and = [];
                for (var _ane = _anc[0]; _ane <= _anc[1]; _ane++)
                    _and.push(_amp(_amk[_ane]));
                _ake.getRange(_anc[0] + 2, 1, _and.length, 6).setValues(_and);
            });
            _amj = true;
        }
    }
    _akc.setProperty('REVISION_FOTO_T', String(_aka.getTime()));
    if (_amj || _akm.total > 0)
        SpreadsheetApp.flush();
    return _akm;
}
var DESCRIPCION_PROTECCION_ = "Sistema Devoluciones (solo aviso)";
function aplicarProteccionesAviso_(_anf, _ang) {
    _anf = _anf || SpreadsheetApp.getActiveSpreadsheet();
    var _anh = [];
    try {
        _anh = _anf.getProtections(SpreadsheetApp.ProtectionType.SHEET).concat(_anf.getProtections(SpreadsheetApp.ProtectionType.RANGE));
    }
    catch (_ani) { }
    var _anj = _anh.map(function (_ank) { try {
        return _ank.getDescription();
    }
    catch (_anl) {
        return "";
    } });
    var _anm = 0;
    TODAS_HOJAS_HISTORIAL_.forEach(function (_ann) {
        var _ano = _anf.getSheetByName(_ann);
        var _anp = DESCRIPCION_PROTECCION_ + ": " + _ann;
        if (!_ano || _anj.indexOf(_anp) > -1)
            return;
        try {
            if (_ano.getProtections(SpreadsheetApp.ProtectionType.SHEET).length > 0)
                return;
        }
        catch (_anq) {
            return;
        }
        _ano.protect().setDescription(_anp).setWarningOnly(true);
        _anm++;
        if (_ang)
            _ang.push(_ann);
    });
    HOJAS_TRABAJO_PERMITIDAS_.forEach(function (_anr) {
        var _ans = _anf.getSheetByName(_anr);
        var _ant = DESCRIPCION_PROTECCION_ + ": encabezados de " + _anr;
        if (!_ans || _anj.indexOf(_ant) > -1)
            return;
        _ans.getRange(1, 1, 1, _ans.getMaxColumns()).protect().setDescription(_ant).setWarningOnly(true);
        _anm++;
        if (_ang)
            _ang.push("encabezados de " + _anr);
    });
    return _anm;
}
var VERSION_SISTEMA_ = "0.9.164";
var CORREO_AVISOS_SALUD_ = "admin@ejemplo.com";
var FORMULAS_LENTAS_ = ["IMPORTRANGE", "INDIRECT", "OFFSET", "NOW(", "TODAY(", "RAND", "QUERY(", "ARRAYFORMULA",
    "INDIRECTO", "DESREF", "AHORA(", "HOY(", "ALEATORIO", "CONSULTA("];
function claveFecha_(_anu) {
    return _anu.getFullYear() + ('0' + (_anu.getMonth() + 1)).slice(-2) + ('0' + _anu.getDate()).slice(-2);
}
function anotarEventoSalud_(_anv, _anw) {
    try {
        var _anx = PropertiesService.getDocumentProperties(), _any = "SALUD_EV_" + _anv, _anz = [];
        try {
            _anz = JSON.parse(_anx.getProperty(_any) || "[]");
        }
        catch (_aoa) {
            _anz = [];
        }
        if (!Array.isArray(_anz))
            _anz = [];
        _anz.unshift({ t: new Date().getTime(), x: String(_anw).substring(0, 160) });
        _anx.setProperty(_any, JSON.stringify(_anz.slice(0, 6)));
    }
    catch (_aob) { }
}
function leerEventosSalud_(_aoc, _aod, _aoe) {
    try {
        var _aof = JSON.parse(_aoc.getProperty("SALUD_EV_" + _aod) || "[]");
        return (Array.isArray(_aof) ? _aof : []).filter(function (_aog) { return _aog && _aog.t >= _aoe; });
    }
    catch (_aoh) {
        return [];
    }
}
function horaEventoSalud_(_aoi) {
    var _aoj = new Date(_aoi);
    return ('0' + _aoj.getDate()).slice(-2) + '/' + ('0' + (_aoj.getMonth() + 1)).slice(-2) + ' ' + ('0' + _aoj.getHours()).slice(-2) + ':' + ('0' + _aoj.getMinutes()).slice(-2);
}
function sumarContador_(_aok, _aol) {
    try {
        var _aom = PropertiesService.getDocumentProperties(), _aon = "CONT_" + _aok + "_" + claveFecha_(new Date());
        _aom.setProperty(_aon, String(Number(_aom.getProperty(_aon) || 0) + (_aol || 1)));
    }
    catch (_aoo) { }
}
function leerContador_(_aop, _aoq, _aor) {
    return Number(_aop.getProperty("CONT_" + _aoq + "_" + claveFecha_(_aor)) || 0);
}
function estadoDelSistema() {
    if (!usuarioAutorizadoOAvisar_())
        return;
    if (!soloAdminOAvisar_())
        return;
    var _aos = SpreadsheetApp.getActiveSpreadsheet(), _aot = SpreadsheetApp.getUi();
    try {
        _aos.toast("Revisando el estado del sistema (tarda unos segundos)...", "🩺 Estado del sistema", 30);
    }
    catch (_aou) { }
    var _aov = revisarSalud_(_aos, new Date());
    var _aow = _aov.lineas.join("\n");
    if (_aov.cambiosReferencia.length > 0) {
        var _aox = _aot.alert("🩺 Estado del sistema", _aow +
            "\n\n📐 = ha cambiado respecto a cuando todo iba bien. Si esos cambios los habéis hecho a propósito y el documento va bien, pulsa SÍ para tomarlos como lo normal a partir de ahora.", _aot.ButtonSet.YES_NO);
        if (_aox === _aot.Button.YES) {
            guardarReferenciaSalud_(_aov.actual);
            _aot.alert("✅ Guardado: a partir de ahora esto es lo normal.");
        }
    }
    else {
        _aot.alert("🩺 Estado del sistema", _aow, _aot.ButtonSet.OK);
    }
}
function revisarSaludDiaria() {
    try {
        var _aoy = revisarSalud_(SpreadsheetApp.getActiveSpreadsheet(), new Date());
        if (_aoy.avisos > 0 && CORREO_AVISOS_SALUD_) {
            MailApp.sendEmail(CORREO_AVISOS_SALUD_, "⚠️ DEVOLUCIONES (DEMO): el sistema necesita un vistazo", _aoy.lineas.join("\n") + "\n\nPara verlo en el documento: menú 📦 Devoluciones → 🩺 Estado del sistema.");
        }
    }
    catch (_aoz) {
        try {
            Logger.log("Revisión de salud: " + (_aoz && _aoz.message ? _aoz.message : _aoz));
        }
        catch (_apa) { }
    }
}
function leerReferenciaSalud_(_apb) {
    try {
        var _apc = _apb.getProperty('SALUD_REFERENCIA');
        return _apc ? JSON.parse(_apc) : null;
    }
    catch (_apd) {
        return null;
    }
}
function guardarReferenciaSalud_(_ape) {
    PropertiesService.getDocumentProperties().setProperty('SALUD_REFERENCIA', JSON.stringify(_ape));
}
var TIPOS_REGLA_SALUD_ = { CUSTOM_FORMULA: "fórmula", TEXT_CONTAINS: "el texto contiene", TEXT_NOT_CONTAINS: "el texto no contiene", TEXT_EQUAL_TO: "el texto es", TEXT_STARTS_WITH: "empieza por", TEXT_ENDS_WITH: "termina en", CELL_EMPTY: "celda vacía", CELL_NOT_EMPTY: "celda no vacía", NUMBER_EQUAL_TO: "igual a", NUMBER_GREATER_THAN: "mayor que", NUMBER_LESS_THAN: "menor que", NUMBER_BETWEEN: "entre", DATE_EQUAL_TO: "fecha igual a", DATE_BEFORE: "fecha anterior a", DATE_AFTER: "fecha posterior a" };
function descripcionReglaSalud_(_apf, _apg, _aph, _api) {
    var _apj = "";
    try {
        _apj = _apf.getRanges().map(function (_apk) { return _apk.getA1Notation(); }).join(", ");
    }
    catch (_apl) { }
    var _apm = "";
    try {
        _apm = String(_apg.getCriteriaType());
    }
    catch (_apn) { }
    var _apo = (TIPOS_REGLA_SALUD_[_apm] || _apm.toLowerCase().replace(/_/g, " ")) + " " + (_apg.getCriteriaValues() || []).map(function (_app) { return String(_app); }).join(" y ");
    var _apq = "";
    try {
        var _apr = _apg.getBackground();
        if (_apr)
            _apq = ", pinta el fondo de " + _apr;
    }
    catch (_aps) { }
    try {
        var _apt = _apg.getFontColor();
        if (_apt)
            _apq += (_apq ? " y" : ", pinta") + " la letra de " + _apt;
    }
    catch (_apu) { }
    return "Regla nº " + (_aph + 1) + " de " + _api + ": celdas " + (_apj || "?") + " · condición: " + _apo.trim().substring(0, 90) + _apq + ".";
}
function revisarSalud_(_apv, _apw) {
    var _apx = PropertiesService.getDocumentProperties();
    var _apy = leerReferenciaSalud_(_apx);
    var _apz = {}, _aqa = [], _aqb = 0, _aqc = [];
    function _aqd(_aqe) { _aqa.push("✅ " + _aqe); }
    function _aqf(_aqg) { _aqa.push("⚠️ " + _aqg); _aqb++; }
    function _aqh(_aqi) { _aqa.push("📐 " + _aqi); _aqc.push(_aqi); _aqb++; }
    function _aqj(_aqk) { _aqa.push("      ↳ " + _aqk); }
    function _aql(_aqm, _aqn) { try {
        _aqn();
    }
    catch (_aqo) {
        _aqf(_aqm + ": no se ha podido comprobar (" + (_aqo && _aqo.message ? _aqo.message : _aqo) + ").");
    } }
    _aqa.push("Versión del script que está funcionando: " + VERSION_SISTEMA_);
    _aqa.push("");
    _aql("Columnas", function () {
        var _aqp = true;
        HOJAS_TRABAJO_PERMITIDAS_.forEach(function (_aqq) {
            var _aqr = _apv.getSheetByName(_aqq);
            if (!_aqr) {
                _aqf("No existe la pestaña «" + _aqq + "» (¿se le ha cambiado el nombre?). Sin ella no se registra nada de esa pestaña.");
                _aqp = false;
                return;
            }
            var _aqs = _aqr.getRange(1, 1, 1, Math.min(_aqr.getMaxColumns(), 17)).getDisplayValues()[0].map(function (_aqt) { return String(_aqt).trim(); });
            _apz["cab_" + _aqq] = _aqs;
            if (_apy && _apy["cab_" + _aqq] && JSON.stringify(_apy["cab_" + _aqq]) !== JSON.stringify(_aqs)) {
                _aqh("Los encabezados de «" + _aqq + "» han cambiado (¿se ha insertado, borrado o movido alguna columna?). El sistema usa B = Agencia, C = Referencia, E = Rotura y L = Estado.");
                var _aqu = _apy["cab_" + _aqq], _aqv = 0;
                for (var _aqw = 0; _aqw < Math.max(_aqu.length, _aqs.length) && _aqv < 6; _aqw++) {
                    if ((_aqu[_aqw] || "") !== (_aqs[_aqw] || "")) {
                        _aqj("Columna " + columnToLetter(_aqw + 1) + ": antes «" + (_aqu[_aqw] || "vacía") + "», ahora «" + (_aqs[_aqw] || "vacía") + "»");
                        _aqv++;
                    }
                }
                _aqp = false;
            }
        });
        if (_aqp)
            _aqd("Entradas y Retornos, con las columnas en su sitio.");
    });
    _aql("Formato condicional", function () {
        var _aqx = true;
        HOJAS_TRABAJO_PERMITIDAS_.forEach(function (_aqy) {
            var _aqz = _apv.getSheetByName(_aqy);
            if (!_aqz)
                return;
            var _ara = _aqz.getConditionalFormatRules(), _arb = 0, _arc = [];
            _ara.forEach(function (_ard, _are) {
                var _arf = _ard.getBooleanCondition();
                if (_arf && (_arf.getCriteriaValues() || []).some(function (_arg) { return String(_arg).indexOf("#REF!") > -1; })) {
                    _arb++;
                    if (_arc.length < 5)
                        _arc.push(descripcionReglaSalud_(_ard, _arf, _are, _ara.length));
                }
            });
            _apz["reglas_" + _aqy] = _ara.length;
            if (_arb > 0) {
                _aqf("«" + _aqy + "» tiene " + (_arb === 1 ? "1 regla de formato condicional rota (#REF!). Conviene quitarla: no marca nada y ralentiza." : _arb + " reglas de formato condicional rotas (#REF!). Conviene quitarlas: no marcan nada y ralentizan."));
                _arc.forEach(_aqj);
                _aqj("Para quitarla: selecciona toda la hoja (el cuadradito de arriba a la izquierda) → Formato → Formato condicional → la regla con esas celdas → papelera.");
                _aqx = false;
            }
            if (_apy && _apy["reglas_" + _aqy] !== undefined && _ara.length > _apy["reglas_" + _aqy] + 5) {
                _aqh("«" + _aqy + "» tiene " + _ara.length + " reglas de formato condicional (lo normal eran " + _apy["reglas_" + _aqy] + "). Muchas reglas hacen que todo vaya lento.");
                _aqx = false;
            }
        });
        if (_aqx)
            _aqd("Formato condicional en orden.");
        try {
            var _arh = leerEventosSalud_(PropertiesService.getDocumentProperties(), "REGLAS", new Date().getTime() - 7 * 86400000);
            if (_arh.length)
                _aqj("🧹 Limpieza nocturna (" + Utilities.formatDate(new Date(_arh[0].t), Session.getScriptTimeZone(), "dd/MM") + "): " + _arh[0].x + ". Para deshacerla: «restaurarReglasNoche» desde el editor.");
        }
        catch (_ari) { }
    });
    _aql("Fórmulas", function () {
        var _arj = {}, _ark = 0;
        var _arl = "^=.*(" + FORMULAS_LENTAS_.map(function (_arm) { return _arm.replace(/[()]/g, function (_arn) { return "\\" + _arn; }); }).join("|") + ")";
        var _aro = {};
        _apv.createTextFinder(_arl).useRegularExpression(true).matchFormulaText(true).matchCase(false).findAll().forEach(function (_arp) {
            var _arq = _arp.getSheet().getName();
            _ark++;
            _arj[_arq] = (_arj[_arq] || 0) + 1;
            if (!_aro[_arq])
                _aro[_arq] = [];
            if (_aro[_arq].length < 3) {
                var _arr = "";
                try {
                    _arr = _arp.getFormula();
                }
                catch (_ars) { }
                _aro[_arq].push(_arp.getA1Notation() + (_arr ? " " + _arr.substring(0, 60) : ""));
            }
        });
        _apz.formulas = _arj;
        var _art = [];
        if (_apy && _apy.formulas) {
            Object.keys(_arj).forEach(function (_aru) {
                var _arv = _apy.formulas[_aru] || 0;
                if (_arj[_aru] > _arv)
                    _art.push("«" + _aru + "»: " + (_arj[_aru] - _arv));
            });
        }
        if (_art.length > 0) {
            _aqh("Hay fórmulas «lentas» nuevas (IMPORTRANGE, INDIRECTO, HOY, AHORA, QUERY...) en " + _art.join(", ") + ". Si el documento va más lento, empieza por ahí.");
            Object.keys(_aro).forEach(function (_arw) { _aqj("«" + _arw + "»: " + _aro[_arw].join(" · ")); });
        }
        else
            _aqd("Sin fórmulas lentas nuevas (" + _ark + " en total, las mismas que cuando iba bien).");
    });
    _aql("Tamaño", function () {
        var _arx = 0, _ary = [];
        _apv.getSheets().forEach(function (_arz) { var _asa = _arz.getMaxRows() * _arz.getMaxColumns(); _arx += _asa; _ary.push({ n: _arz.getName(), c: _asa }); });
        _ary.sort(function (_asb, _asc) { return _asc.c - _asb.c; });
        var _asd = _ary.slice(0, 4).map(function (_ase) { return "«" + _ase.n + "» " + (Math.round(_ase.c / 10000) / 100) + " M"; }).join(" · ");
        _apz.celdas = _arx;
        var _asf = (Math.round(_arx / 100000) / 10) + " millones de celdas (Google no deja pasar de 10)";
        if (_arx > 8000000) {
            _aqf("El documento tiene " + _asf + ". Hay que archivar o borrar filas/columnas vacías que sobren.");
            _aqj("Las más grandes: " + _asd);
        }
        else if (_apy && _apy.celdas && _arx > _apy.celdas * 1.3) {
            _aqh("El documento ha crecido mucho: " + _asf + ", antes " + (Math.round(_apy.celdas / 100000) / 10) + ". ¿Alguien ha añadido una pestaña grande o muchas filas vacías?");
            _aqj("Las más grandes: " + _asd);
        }
        else
            _aqd("Tamaño: " + _asf + ".");
    });
    _aql("Velocidad", function () {
        var _asg = _apv.getSheetByName(HOJAS_TRABAJO_PERMITIDAS_[0]);
        if (!_asg || _asg.getLastRow() < 2)
            return;
        var _ash = new Date().getTime();
        var _asi = _asg.getLastRow(), _asj = Math.max(2, _asi - 2500 + 1);
        _asg.getRange(_asj, COLUMNA_AGENCIAS, _asi - _asj + 1, COLUMNA_NOTA_L - COLUMNA_AGENCIAS + 1).getValues();
        var _ask = new Date().getTime() - _ash;
        _apz.velocidadMs = (_apy && _apy.velocidadMs) ? _apy.velocidadMs : _ask;
        if (_apy && _apy.velocidadMs && _ask > Math.max(3000, _apy.velocidadMs * 3))
            _aqf("El documento va más lento de lo normal: leer las últimas 2.500 filas ha tardado " + _ask + " ms (lo normal, unos " + _apy.velocidadMs + " ms).");
        else
            _aqd("Velocidad normal (" + _ask + " ms para leer las últimas 2.500 filas).");
    });
    _aql("Registro", function () {
        var _asl = Number(_apx.getProperty('REVISION_ULTIMA_AUTO') || 0);
        var _asm = _asl ? Math.round((_apw.getTime() - _asl) / 60000) : null;
        if (_asm === null || _asm > 20)
            _aqf("La revisión automática cada 5 minutos no está funcionando" + (_asm !== null ? " (la última fue hace " + _asm + " min)" : "") + ". Pulsa «🕰️ Auto-Archivado» una vez para activarla.");
        else
            _aqd("Revisión automática activa (la última, hace " + _asm + " min).");
        var _asn = new Date(_apw.getTime() - 86400000);
        var _aso = leerContador_(_apx, "CORRECCIONES", _apw), _asp = leerContador_(_apx, "CORRECCIONES", _asn);
        var _asq = leerContador_(_apx, "ERRORES", _apw), _asr = leerContador_(_apx, "ERRORES", _asn);
        var _ass = leerContador_(_apx, "LENTAS", _apw), _ast = leerContador_(_apx, "LENTAS", _asn);
        var _asu = new Date(_asn.getFullYear(), _asn.getMonth(), _asn.getDate()).getTime();
        function _asv(_asw) { leerEventosSalud_(_apx, _asw, _asu).slice(0, 5).forEach(function (_asx) { _aqj(horaEventoSalud_(_asx.t) + " · " + _asx.x); }); }
        if (_asp > 30 || _aso > 30) {
            _aqf("La revisión automática ha tenido que añadir o quitar muchos registros (hoy " + _aso + ", ayer " + _asp + "): el registro al momento está fallando más de lo normal (documento lento o sobrecargado). Ya está corregido, pero conviene mirar por qué.");
            _asv("CORR");
        }
        else
            _aqd("Registro al momento funcionando (correcciones de la revisión automática: hoy " + _aso + ", ayer " + _asp + ").");
        if (_asq + _asr > 0)
            _aqf("Ha habido " + (_asq + _asr) + " errores al registrar (hoy " + _asq + ", ayer " + _asr + "). La revisión automática los recupera, pero conviene mirar Extensiones → Apps Script → Ejecuciones.");
        var _asy = leerContador_(_apx, "LENTASRET", _apw), _asz = leerContador_(_apx, "LENTASRET", _asn);
        if (_ass + _ast > 5)
            _aqf("Ha habido " + (_ass + _ast) + " ediciones que tardaron más de 20 segundos en registrarse (hoy " + _ass + ", ayer " + _ast + "; de ellas en Retornos: hoy " + _asy + ", ayer " + _asz + "). Google corta a los 30: el documento va cargado. En Extensiones → Apps Script → Ejecuciones, cada una dice qué parte tardó.");
        if (_ass + _ast > 5)
            _asv("LENTAS");
        var _ata = leerContador_(_apx, "APLAZADOS", _apw), _atb = leerContador_(_apx, "APLAZADOS", _asn);
        if (_ata + _atb > 0)
            _aqf("Apuntes aplazados a la revisión de cada 5 minutos porque el documento iba cargado: hoy " + _ata + ", ayer " + _atb + ". No se pierde nada (la revisión los añade), pero indica atascos.");
        if (_ata + _atb > 0)
            _asv("APLAZ");
        var _atc = leerContador_(_apx, "HISTRETLENTO", _apw), _atd = leerContador_(_apx, "HISTRETSALTADO", _apw);
        if (_atc + _atd > 0)
            _aqf("Historial de cambios de Retornos hoy: tardó más de 5 s en " + _atc + " edición(es) y en " + _atd + " no se apuntó porque la edición ya iba lenta.");
        var _ate = leerContador_(_apx, "BUSQLENTAS", _apw), _atf = leerContador_(_apx, "BUSQLENTAS", _asn);
        if (_ate + _atf > 10)
            _aqf("Búsquedas de referencia lentas (más de 10 s): hoy " + _ate + ", ayer " + _atf + ". El documento va cargado a esas horas.");
        else
            _aqd("Búsquedas de referencia lentas (más de 10 s): hoy " + _ate + ", ayer " + _atf + ".");
    });
    _aql("Archivado", function () {
        var _atg = _apv.getSheetByName(HOJA_AGENCIAS);
        if (!_atg || _atg.getLastRow() < 2)
            return;
        var _ath = new Date(_atg.getRange(2, 1).getValue());
        var _ati = Math.floor((_apw.getTime() - _ath.getTime()) / 86400000);
        if (!isNaN(_ati) && _ati > DIAS_CADUCIDAD_HISTORIALES + 3)
            _aqf("El archivado nocturno no se está haciendo: el historial de agencias guarda registros de hace " + _ati + " días. Pulsa «🕰️ Auto-Archivado».");
        else
            _aqd("Archivado nocturno al día (historial de agencias: " + (_atg.getLastRow() - 1) + " registros).");
    });
    _aql("Protecciones", function () {
        var _atj = [];
        var _atk = aplicarProteccionesAviso_(_apv, _atj);
        _aqd(_atk > 0 ? "Protecciones «solo aviso» puestas de nuevo (" + _atk + "): " + _atj.join(", ") + "." : "Protecciones «solo aviso» en su sitio.");
    });
    _aql("Disparadores", function () {
        var _atl = {};
        ScriptApp.getProjectTriggers().forEach(function (_atm) { var _atn = _atm.getHandlerFunction(); _atl[_atn] = (_atl[_atn] || 0) + 1; });
        var _ato = Object.keys(_atl).filter(function (_atp) { return _atl[_atp] > 1; });
        if (_ato.length > 0)
            _aqf("Hay disparadores repetidos: " + _ato.join(", ") + ". Pulsa «🕰️ Auto-Archivado» para dejarlos bien.");
    });
    try {
        var _atq = claveFecha_(new Date(_apw.getTime() - 14 * 86400000));
        _apx.getKeys().forEach(function (_atr) { var _ats = /^CONT_[A-Z]+_(\d{8})$/.exec(_atr); if (_ats && _ats[1] < _atq)
            _apx.deleteProperty(_atr); });
    }
    catch (_att) { }
    if (!_apy) {
        guardarReferenciaSalud_(_apz);
        _aqa.push("");
        _aqa.push("📌 Primera revisión: se ha guardado el estado actual como «lo normal» para compararlo en adelante.");
    }
    _aqa.splice(2, 0, _aqb === 0 ? "🟢 Todo en orden." : "🟠 Hay " + _aqb + " cosa(s) que mirar:");
    return { lineas: _aqa, avisos: _aqb, cambiosReferencia: _aqc, actual: _apz };
}
var OSITA_SVG_ = '<defs><radialGradient id="ot-pelo" cx="38%" cy="30%" r="75%"><stop offset="0%" stop-color="#D3D8DE"/><stop offset="60%" stop-color="#A3AAB3"/><stop offset="100%" stop-color="#7A818B"/></radialGradient></defs><ellipse cx="50" cy="105" rx="30" ry="4.5" fill="#1B2636" opacity=".12"/><path d="M70 98 C88 98 99 84 97 66 C96 58 92 53 86 51" stroke="url(#rufo-cuerpo)" stroke-width="13" stroke-linecap="round" fill="none"/><path d="M70 98 C88 98 99 84 97 66 C96 58 92 53 86 51" stroke="#2E3238" stroke-width="13" stroke-dasharray="6 7" stroke-dashoffset="-3" fill="none"/><circle cx="86" cy="51" r="6.5" fill="#2A2E33"/><path d="M22 106 C17 88 22 72 50 70 C78 72 83 88 78 106 Z" fill="url(#rufo-cuerpo)"/><ellipse cx="50" cy="93" rx="15" ry="13" fill="url(#rufo-gradMorro)"/><ellipse cx="37" cy="105" rx="9" ry="4" fill="#2A2E33"/><ellipse cx="63" cy="105" rx="9" ry="4" fill="#2A2E33"/><path d="M14 34 C8 18 13 5 25 3 C35 6 38 17 35 27 Z" fill="#4A5059"/><path d="M18 28 C15 18 18 10 25 9 C31 11 32 18 30 24 Z" fill="url(#rufo-gradOrejaInt)"/><path d="M86 34 C92 18 87 5 75 3 C65 6 62 17 65 27 Z" fill="#4A5059"/><path d="M82 28 C85 18 82 10 75 9 C69 11 68 18 70 24 Z" fill="url(#rufo-gradOrejaInt)"/><path d="M14 34 C8 18 13 5 25 3" stroke="#F1F3F5" stroke-width="1.6" fill="none"/><path d="M86 34 C92 18 87 5 75 3" stroke="#F1F3F5" stroke-width="1.6" fill="none"/><path d="M17 50 C8 53 3 60 1 68 C9 64 14 65 20 63 Z" fill="url(#rufo-gradMorro)"/><path d="M83 50 C92 53 97 60 99 68 C91 64 86 65 80 63 Z" fill="url(#rufo-gradMorro)"/><path d="M15 47 C14 24 31 10 50 10 C69 10 86 24 85 47 C84 66 69 79 50 79 C31 79 16 66 15 47 Z" fill="url(#ot-pelo)"/><path d="M50 17 C47 23 47 29 50 35 C53 29 53 23 50 17 Z" fill="#3E434A" opacity=".6"/><g class="ot-pest"><ellipse cx="34" cy="26" rx="10" ry="3.8" transform="rotate(-12 34 26)" fill="#F4F5F7"/><ellipse cx="66" cy="26" rx="10" ry="3.8" transform="rotate(12 66 26)" fill="#F4F5F7"/></g><path d="M17 42 C18 30 30 27 40 32 C45 34 47 36 50 36 C53 36 55 34 60 32 C70 27 82 30 83 42 C82 51 72 54 62 51 C57 49 54 47 50 47 C46 47 43 49 38 51 C28 54 18 51 17 42 Z" fill="url(#rufo-grad)"/><g class="ot-ojo"><circle cx="34" cy="39" r="9" fill="url(#rufo-gradOjo)" stroke="#5B636E" stroke-width=".8"/><circle cx="66" cy="39" r="9" fill="url(#rufo-gradOjo)" stroke="#5B636E" stroke-width=".8"/><circle cx="31.4" cy="36.1" r="3.3" fill="#fff" opacity=".95"/><circle cx="63.4" cy="36.1" r="3.3" fill="#fff" opacity=".95"/><circle cx="36.6" cy="42" r="1.1" fill="#fff" opacity=".6"/><circle cx="68.6" cy="42" r="1.1" fill="#fff" opacity=".6"/></g><path d="M34 66 C35 54 42 49 50 49 C58 49 65 54 66 66 C63 74 57 77 50 77 C43 77 37 74 34 66 Z" fill="url(#rufo-gradMorro)"/><ellipse cx="50" cy="55" rx="6" ry="4.2" fill="#15171A"/><ellipse cx="48" cy="53.6" rx="1.8" ry="1" fill="#fff" opacity=".55"/><path d="M50 59 L50 61.5" stroke="#2B2E33" stroke-width="1.6" stroke-linecap="round"/><path d="M42 61 Q50 68 58 61" stroke="#2B2E33" stroke-width="2.6" fill="none" stroke-linecap="round"/><circle cx="24" cy="58" r="6" fill="url(#rufo-gradMejilla)"/><circle cx="76" cy="58" r="6" fill="url(#rufo-gradMejilla)"/><path d="M30 74 Q50 84 70 74 L67 81 Q50 90 33 81 Z" fill="url(#rufo-gradLazo)"/><path d="M60 80 L68 92 L57 88 Z" fill="#E2781E"/><path d="M45 13 C41 3 48 -2 51 6 C53 -2 61 1 56 13 Z" fill="#5A6069"/>';
var ESCENAS_CABECERA_ = {
    buscar: '<g class="pr-cab ot-osita" onclick="otPestanear(this)"><g transform="translate(40,4) scale(.6)">' + OSITA_SVG_ + '</g></g><use href="#esc-mesa"/><rect x="104" y="53" width="26" height="17" rx="2.5" fill="#E07A1F"/><rect x="104" y="53" width="26" height="4" fill="#9f0c24"/><path d="M109 60 v7 M111 60 v7 M114 60 v7 M115.5 60 v7 M118 60 v7 M121 60 v7 M122.5 60 v7 M125 60 v7" stroke="#fff" stroke-width="1"/><g class="bs-ok"><circle cx="128" cy="46" r="5.5" fill="#22c55e"/><path d="M125.4 46 l1.8 1.8 l3.4 -3.6" stroke="#fff" stroke-width="1.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/></g><circle cx="54" cy="69" r="5" fill="url(#rufo-grad)"/><g class="bs-pistola"><path d="M84 62 L98 57 L100 61 L90 65 L91 71 L86 72 Z" fill="#374151"/><rect x="97" y="56" width="4" height="6" rx="1" fill="#111827"/><circle cx="87" cy="68" r="5" fill="url(#rufo-grad)"/><path class="bs-laser" d="M101 59 L118 62" stroke="#ef4444" stroke-width="1.6" stroke-linecap="round"/></g>',
    calendario: '<g class="pr-cab"><g transform="translate(40,4) scale(.6)"><use href="#rufo-svg"/></g></g><use href="#esc-mesa"/><rect x="10" y="34" width="34" height="36" rx="3" fill="#fff" stroke="#d1d5db" stroke-width=".8"/><rect x="10" y="34" width="34" height="9" rx="3" fill="#E07A1F"/><circle cx="18" cy="34" r="1.8" fill="#6b1620"/><circle cx="36" cy="34" r="1.8" fill="#6b1620"/><g fill="#cbd5e1"><rect x="14" y="47" width="5" height="4" rx="1"/><rect x="21" y="47" width="5" height="4" rx="1"/><rect x="28" y="47" width="5" height="4" rx="1"/><rect x="35" y="47" width="5" height="4" rx="1"/><rect x="14" y="54" width="5" height="4" rx="1"/><rect x="21" y="54" width="5" height="4" rx="1"/><rect x="35" y="54" width="5" height="4" rx="1"/><rect x="14" y="61" width="5" height="4" rx="1"/><rect x="21" y="61" width="5" height="4" rx="1"/><rect x="28" y="61" width="5" height="4" rx="1"/></g><g class="cl-marca"><circle cx="30.5" cy="56" r="4.2" fill="none" stroke="#E07A1F" stroke-width="1.6"/></g><g class="cl-hoja"><rect x="10" y="43" width="34" height="27" fill="#f8fafc" stroke="#d1d5db" stroke-width=".8"/><text x="27" y="62" font-family="Georgia,serif" font-size="13" font-weight="bold" fill="#94a3b8" text-anchor="middle">24</text></g><g class="cl-pata"><circle cx="52" cy="66" r="5" fill="url(#rufo-grad)"/></g><circle cx="88" cy="69" r="5" fill="url(#rufo-grad)"/>',
    casillas: '<g class="pr-cab"><g transform="translate(40,4) scale(.6)"><use href="#rufo-svg"/></g></g><use href="#esc-mesa"/><rect x="100" y="34" width="32" height="37" rx="3" fill="#b45309"/><rect x="103" y="38" width="26" height="31" rx="1.5" fill="#fff"/><rect x="111" y="31" width="10" height="6" rx="2" fill="#6b7280"/><g stroke="#94a3b8" stroke-width="1" fill="none"><rect x="106" y="42" width="5" height="5" rx="1"/><rect x="106" y="51" width="5" height="5" rx="1"/><rect x="106" y="60" width="5" height="5" rx="1"/></g><path d="M114 44.5 h11 M114 53.5 h11 M114 62.5 h9" stroke="#cbd5e1" stroke-width="1.4" stroke-linecap="round"/><path class="cs-t1" d="M106.5 44.3 l1.6 1.6 l3 -3.4" stroke="#16a34a" stroke-width="1.6" fill="none" stroke-linecap="round"/><path class="cs-t2" d="M106.5 53.3 l1.6 1.6 l3 -3.4" stroke="#16a34a" stroke-width="1.6" fill="none" stroke-linecap="round"/><path class="cs-t3" d="M106.5 62.3 l1.6 1.6 l3 -3.4" stroke="#16a34a" stroke-width="1.6" fill="none" stroke-linecap="round"/><circle cx="54" cy="69" r="5" fill="url(#rufo-grad)"/><g class="cs-lapiz"><line x1="92" y1="52" x2="103" y2="42" stroke="#E07A1F" stroke-width="3" stroke-linecap="round"/><path d="M103 42 l2.6 -2.2 l-1 3.4 z" fill="#111"/><circle cx="91" cy="53" r="5" fill="url(#rufo-grad)"/></g>',
    notas: '<g class="pr-cab"><g transform="translate(40,4) scale(.6)"><use href="#rufo-svg"/></g></g><use href="#esc-mesa"/><g transform="rotate(-4 27 50)"><rect x="10" y="33" width="34" height="34" rx="2" fill="#fde68a"/><path d="M10 33 h34 v5 h-34 z" fill="#fcd34d"/><path class="nt-l1" d="M14 45 h22" stroke="#92400e" stroke-width="1.6" stroke-linecap="round"/><path class="nt-l2" d="M14 51 h22" stroke="#92400e" stroke-width="1.6" stroke-linecap="round"/><path class="nt-l3" d="M14 57 h16" stroke="#92400e" stroke-width="1.6" stroke-linecap="round"/><circle cx="27" cy="34" r="2.6" fill="#E07A1F"/><g class="nt-lapiz"><line x1="37" y1="36" x2="44" y2="30" stroke="#E07A1F" stroke-width="2.6" stroke-linecap="round"/><path d="M36.5 36.5 l1.5 -3 l1.5 1.5 z" fill="#111"/></g></g><circle cx="50" cy="66" r="5" fill="url(#rufo-grad)"/><circle cx="88" cy="69" r="5" fill="url(#rufo-grad)"/>',
    avisos: '<g class="pr-cab"><g transform="translate(40,4) scale(.6)"><use href="#rufo-svg"/></g></g><use href="#esc-mesa"/><rect x="104" y="56" width="22" height="14" rx="2.5" fill="#E07A1F"/><text x="115" y="66" font-family="Georgia,serif" font-size="8" font-weight="bold" fill="#fff" text-anchor="middle">P</text><g class="av-ondas" stroke="#f59e0b" stroke-width="1.6" fill="none" stroke-linecap="round"><path d="M103 22 q-4 6 0 12"/><path d="M99 20 q-6 8 0 16"/><path d="M127 22 q4 6 0 12"/><path d="M131 20 q6 8 0 16"/></g><g class="av-campana"><rect x="114" y="12" width="2" height="5" fill="#92400e"/><path d="M107 36 q0 -18 8 -19 q8 1 8 19 z" fill="#fbbf24"/><rect x="105" y="35" width="20" height="3" rx="1.5" fill="#f59e0b"/><circle cx="115" cy="40" r="2.2" fill="#92400e"/></g><g class="av-excl"><circle cx="96" cy="16" r="6" fill="#E07A1F"/><text x="96" y="19.5" font-family="Arial" font-size="9" font-weight="bold" fill="#fff" text-anchor="middle">!</text></g><circle cx="54" cy="69" r="5" fill="url(#rufo-grad)"/><circle cx="92" cy="67" r="5" fill="url(#rufo-grad)"/>',
    recuento: '<g class="pr-cab"><g transform="translate(40,4) scale(.6)"><use href="#rufo-svg"/></g></g><use href="#esc-mesa"/><line x1="100" y1="70" x2="134" y2="70" stroke="#94a3b8" stroke-width="1"/><rect class="rc-b1" x="102" y="50" width="8" height="20" rx="1.5" fill="#22c55e"/><rect class="rc-b2" x="113" y="40" width="8" height="30" rx="1.5" fill="#2563eb"/><rect class="rc-b3" x="124" y="56" width="8" height="14" rx="1.5" fill="#E07A1F"/><g class="rc-caja"><rect x="14" y="54" width="24" height="16" rx="2" fill="#d6a468"/><rect x="14" y="54" width="24" height="4" fill="#b7844a"/><path d="M22 58 l3 4 l-2 3 l3 5" stroke="#E07A1F" stroke-width="1.6" fill="none" stroke-linecap="round" stroke-linejoin="round"/></g><circle cx="52" cy="69" r="5" fill="url(#rufo-grad)"/><circle cx="88" cy="69" r="5" fill="url(#rufo-grad)"/>',
    retornos: '<g class="pr-cab ot-osita" onclick="otPestanear(this)"><g transform="translate(40,4) scale(.6)">' + OSITA_SVG_ + '</g></g><use href="#esc-mesa"/><rect x="98" y="44" width="30" height="26" rx="2" fill="#C48E57"/><rect x="110" y="44" width="6" height="26" fill="#EBCB98"/><text x="16" y="64" font-size="22">↩️</text>'
};
var DEFS_BEBE_RUFO_ = '<svg width="0" height="0" style="position:absolute" aria-hidden="true"><defs><radialGradient id="bb-g" cx="38%" cy="30%" r="75%"><stop offset="0%" stop-color="#D3D8DE"/><stop offset="60%" stop-color="#A3AAB3"/><stop offset="100%" stop-color="#7A818B"/></radialGradient><g id="bebe-rufo"><g transform="translate(14,2)"><ellipse cx="50" cy="105" rx="30" ry="4.5" fill="#1B2636" opacity=".12"/><path d="M70 98 C88 98 99 84 97 66 C96 58 92 53 86 51" stroke="url(#rufo-cuerpo)" stroke-width="13" stroke-linecap="round" fill="none"/><path d="M70 98 C88 98 99 84 97 66 C96 58 92 53 86 51" stroke="#2E3238" stroke-width="13" stroke-dasharray="6 7" stroke-dashoffset="-3" fill="none"/><circle cx="86" cy="51" r="6.5" fill="#2A2E33"/><path d="M22 106 C17 88 22 72 50 70 C78 72 83 88 78 106 Z" fill="url(#rufo-cuerpo)"/><ellipse cx="50" cy="93" rx="15" ry="13" fill="url(#rufo-gradMorro)"/><ellipse cx="37" cy="105" rx="9" ry="4" fill="#2A2E33"/><ellipse cx="63" cy="105" rx="9" ry="4" fill="#2A2E33"/><path d="M14 34 C8 18 13 5 25 3 C35 6 38 17 35 27 Z" fill="#4A5059"/><path d="M18 28 C15 18 18 10 25 9 C31 11 32 18 30 24 Z" fill="url(#rufo-gradOrejaInt)"/><path d="M86 34 C92 18 87 5 75 3 C65 6 62 17 65 27 Z" fill="#4A5059"/><path d="M82 28 C85 18 82 10 75 9 C69 11 68 18 70 24 Z" fill="url(#rufo-gradOrejaInt)"/><path d="M14 34 C8 18 13 5 25 3" stroke="#F1F3F5" stroke-width="1.6" fill="none"/><path d="M86 34 C92 18 87 5 75 3" stroke="#F1F3F5" stroke-width="1.6" fill="none"/><path d="M17 50 C8 53 3 60 1 68 C9 64 14 65 20 63 Z" fill="url(#rufo-gradMorro)"/><path d="M83 50 C92 53 97 60 99 68 C91 64 86 65 80 63 Z" fill="url(#rufo-gradMorro)"/><path d="M15 47 C14 24 31 10 50 10 C69 10 86 24 85 47 C84 66 69 79 50 79 C31 79 16 66 15 47 Z" fill="url(#bb-g)"/><path d="M50 17 C47 23 47 29 50 35 C53 29 53 23 50 17 Z" fill="#3E434A" opacity=".6"/><g><ellipse cx="34" cy="26" rx="10" ry="3.8" transform="rotate(-12 34 26)" fill="#F4F5F7"/><ellipse cx="66" cy="26" rx="10" ry="3.8" transform="rotate(12 66 26)" fill="#F4F5F7"/></g><path d="M17 42 C18 30 30 27 40 32 C45 34 47 36 50 36 C53 36 55 34 60 32 C70 27 82 30 83 42 C82 51 72 54 62 51 C57 49 54 47 50 47 C46 47 43 49 38 51 C28 54 18 51 17 42 Z" fill="url(#rufo-grad)"/><g><circle cx="34" cy="39" r="9" fill="url(#rufo-gradOjo)" stroke="#5B636E" stroke-width=".8"/><circle cx="66" cy="39" r="9" fill="url(#rufo-gradOjo)" stroke="#5B636E" stroke-width=".8"/><circle cx="31.4" cy="36.1" r="3.3" fill="#fff" opacity=".95"/><circle cx="63.4" cy="36.1" r="3.3" fill="#fff" opacity=".95"/><circle cx="36.6" cy="42" r="1.1" fill="#fff" opacity=".6"/><circle cx="68.6" cy="42" r="1.1" fill="#fff" opacity=".6"/></g><path d="M34 66 C35 54 42 49 50 49 C58 49 65 54 66 66 C63 74 57 77 50 77 C43 77 37 74 34 66 Z" fill="url(#rufo-gradMorro)"/><ellipse cx="50" cy="55" rx="6" ry="4.2" fill="#15171A"/><ellipse cx="48" cy="53.6" rx="1.8" ry="1" fill="#fff" opacity=".55"/><path d="M50 59 L50 61.5" stroke="#2B2E33" stroke-width="1.6" stroke-linecap="round"/><path d="M42 61 Q50 68 58 61" stroke="#2B2E33" stroke-width="2.6" fill="none" stroke-linecap="round"/><circle cx="24" cy="58" r="6" fill="url(#rufo-gradMejilla)"/><circle cx="76" cy="58" r="6" fill="url(#rufo-gradMejilla)"/><path d="M30 74 Q50 84 70 74 L67 81 Q50 90 33 81 Z" fill="url(#rufo-gradLazo)"/><path d="M60 80 L68 92 L57 88 Z" fill="#E2781E"/><path d="M45 13 C41 3 48 -2 51 6 C53 -2 61 1 56 13 Z" fill="#5A6069"/></g><g><circle cx="103" cy="80" r="15" fill="#DCEBFA" fill-opacity=".55" stroke="#26354A" stroke-width="4.5"/><circle cx="98" cy="75" r="4" fill="#fff" opacity=".8"/><path d="M113 91 L124 104" stroke="#7A4E26" stroke-width="6.5" stroke-linecap="round"/></g></g></defs></svg>';
var DEFS_ESCRITORIO_RETORNOS_ = '<svg width="0" height="0" style="position:absolute" aria-hidden="true"><defs><radialGradient id="rbg" cx="38%" cy="28%" r="80%"><stop offset="0%" stop-color="#454545"/><stop offset="55%" stop-color="#1e1e1e"/><stop offset="100%" stop-color="#070707"/></radialGradient><radialGradient id="rbm" cx="40%" cy="25%" r="85%"><stop offset="0%" stop-color="#fff"/><stop offset="100%" stop-color="#e7e0db"/></radialGradient><radialGradient id="rbo" cx="38%" cy="30%" r="75%"><stop offset="0%" stop-color="#a3303f"/><stop offset="100%" stop-color="#5c1420"/></radialGradient><linearGradient id="rbw" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#ffe083"/><stop offset="60%" stop-color="#f5c542"/><stop offset="100%" stop-color="#d9a21e"/></linearGradient><radialGradient id="rbl" cx="35%" cy="30%" r="70%"><stop offset="0%" stop-color="#fff" stop-opacity=".95"/><stop offset="100%" stop-color="#bfdbfe" stop-opacity=".4"/></radialGradient><linearGradient id="rbd" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#4a4a4a"/><stop offset="100%" stop-color="#262626"/></linearGradient></defs></svg>';
function escenaCabecera_(_atu) {
    var _atv = ESCENAS_CABECERA_[_atu];
    return _atv ? '<svg class="esc enc-escena" viewBox="0 0 140 100" aria-hidden="true">' + _atv + '</svg>' : '';
}
var COL_RET_CLIENTE_ = 8;
var COL_RET_ARTICULO_ = 10;
var COL_RET_ULTIMA_ = 14;
var MAX_RESULTADOS_RETORNOS_ = 30;
var UMBRAL_PARECIDO_RETORNOS_ = 70;
var COLOR_PROCESADO_POR_DEFECTO_ = "#b6d7a8";
function normTextoRetornos_(_atw) {
    var _atx = (_atw === null || _atw === undefined || _atw instanceof Date) ? "" : String(_atw);
    return _atx.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
}
function compactoRetornos_(_aty) {
    return normTextoRetornos_(_aty).replace(/ /g, '').replace(/^0+(?=\d)/, '');
}
function distanciaDentro_(_atz, _aua) {
    var _aub = _atz.length, _auc = _aua.length, _aud = new Array(_auc + 1), _aue = new Array(_auc + 1), _auf, _aug;
    for (_aug = 0; _aug <= _auc; _aug++)
        _aud[_aug] = 0;
    for (_auf = 1; _auf <= _aub; _auf++) {
        _aue[0] = _auf;
        var _auh = _atz.charCodeAt(_auf - 1);
        for (_aug = 1; _aug <= _auc; _aug++) {
            var _aui = _aud[_aug - 1] + (_auh === _aua.charCodeAt(_aug - 1) ? 0 : 1);
            if (_aud[_aug] + 1 < _aui)
                _aui = _aud[_aug] + 1;
            if (_aue[_aug - 1] + 1 < _aui)
                _aui = _aue[_aug - 1] + 1;
            _aue[_aug] = _aui;
        }
        var _auj = _aud;
        _aud = _aue;
        _aue = _auj;
    }
    var _auk = _aub;
    for (_aug = 0; _aug <= _auc; _aug++)
        if (_aud[_aug] < _auk)
            _auk = _aud[_aug];
    return _auk;
}
function similitudRetornos_(_aul, _aum) {
    var _aun = Math.max(_aul.length, _aum.length);
    return _aun ? 1 - levenshtein(_aul, _aum) / _aun : 0;
}
function prepararBusquedaRetornos_(_auo) {
    var _aup = normTextoRetornos_(_auo).split(' ').filter(function (_auq) { return _auq.length >= 2; });
    var _aur = /[a-z]/.test(normTextoRetornos_(_auo)), _aus = compactoRetornos_(_auo);
    return { compacto: _aus, palabras: _aup, conLetras: _aur, parecePedido: !_aur && _aus.length >= 6 };
}
function cuentaLetras_(_aut) { var _auu = {}; for (var _auv = 0; _auv < _aut.length; _auv++) {
    var _auw = _aut.charCodeAt(_auv);
    _auu[_auw] = (_auu[_auw] || 0) + 1;
} return _auu; }
function comunesLetras_(_aux, _auy) { var _auz = {}, _ava = 0; for (var _avb = 0; _avb < _auy.length; _avb++) {
    var _avc = _auy.charCodeAt(_avb), _avd = _auz[_avc] || 0;
    if (_avd < (_aux[_avc] || 0)) {
        _auz[_avc] = _avd + 1;
        _ava++;
    }
} return _ava; }
function prepararFiltroRetornos_(_ave) {
    var _avf = { L: _ave.compacto.length, cq: cuentaLetras_(_ave.compacto), pal: [] };
    if (_ave.conLetras)
        _ave.palabras.forEach(function (_avg) { _avf.pal.push({ n: _avg.length, c: cuentaLetras_(_avg) }); });
    return _avf;
}
function cotaParecidoRetornos_(_avh, _avi) {
    if (_avi === "" || _avi === null || _avi === undefined)
        return 0;
    var _avj = compactoRetornos_(_avi);
    if (!_avj || !_avh.L)
        return 0;
    var _avk = comunesLetras_(_avh.cq, _avj) / _avh.L;
    if (_avh.pal.length) {
        var _avl = normTextoRetornos_(_avi).split(' '), _avm = 0;
        _avh.pal.forEach(function (_avn) { var _avo = 0; for (var _avp = 0; _avp < _avl.length; _avp++) {
            var _avq = comunesLetras_(_avn.c, _avl[_avp]) / Math.max(_avn.n, _avl[_avp].length);
            if (_avq > _avo)
                _avo = _avq;
        } _avm += _avo; });
        if (_avm / _avh.pal.length > _avk)
            _avk = _avm / _avh.pal.length;
    }
    return _avk;
}
var COTA_MINIMA_RETORNOS_ = (UMBRAL_PARECIDO_RETORNOS_ - 0.5) / 100 - 1e-9;
function parecidoRetornos_(_avr, _avs) {
    if (_avs === "" || _avs === null || _avs === undefined)
        return { pct: 0 };
    var _avt = compactoRetornos_(_avs);
    if (!_avt || !_avr.compacto)
        return { pct: 0 };
    if (_avt.indexOf(_avr.compacto) > -1)
        return { pct: 100, exacta: true };
    var _avu = _avr.compacto.length;
    if (_avu < 4)
        return { pct: 0 };
    var _avv = similitudRetornos_(_avr.compacto, _avt), _avw = null;
    var _avx = distanciaDentro_(_avr.compacto, _avt);
    var _avy = _avu <= 6 ? 1 : Math.floor(_avu * 0.3);
    if (_avx <= _avy) {
        var _avz = 1 - _avx / _avu;
        if (_avz > _avv) {
            _avv = _avz;
            _avw = _avx;
        }
    }
    if (_avr.conLetras && _avr.palabras.length > 0) {
        var _awa = normTextoRetornos_(_avs).split(' '), _awb = 0;
        _avr.palabras.forEach(function (_awc) {
            var _awd = 0;
            for (var _awe = 0; _awe < _awa.length; _awe++) {
                var _awf = similitudRetornos_(_awc, _awa[_awe]);
                if (_awf > _awd)
                    _awd = _awf;
            }
            _awb += _awd;
        });
        var _awg = _awb / _avr.palabras.length;
        if (_awg > _avv)
            _avv = _awg;
    }
    var _awh = Math.min(99, Math.round(_avv * 100));
    return { pct: _awh, dist: _avw };
}
function porQueParecidoRetornos_(_awi, _awj, _awk) {
    if (_awi.exacta) {
        return String(_awj).indexOf(String(_awk).trim()) > -1 ? "Igual." : "Igual sin tildes, mayúsculas ni espacios.";
    }
    if (_awi.dist === 1)
        return "Una letra o cifra distinta.";
    if (_awi.dist === 2)
        return "Dos letras o cifras distintas.";
    return "Se parece un " + _awi.pct + "%.";
}
function esVerdeProcesado_(_awl) {
    var _awm = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(String(_awl || ""));
    if (!_awm)
        return false;
    var _awn = parseInt(_awm[1], 16), _awo = parseInt(_awm[2], 16), _awp = parseInt(_awm[3], 16);
    return _awo >= 120 && _awo > _awn + 12 && _awo > _awp + 12;
}
function cssBultosCaja_() {
    return `    
    .caja-bultos { background:#fff; border:1.5px solid #fcd34d; border-radius:12px; padding:12px 14px; margin-bottom:16px; }
    .caja-bultos .caja-lado-titulo { color:#b45309; }
    .caja-bultos-fila { display:flex; flex-wrap:wrap; gap:8px; }
    .caja-bulto { border:1.5px solid #e5e7eb; background:#f9fafb; color:#374151; border-radius:10px; padding:9px 13px; font-size:14px; font-weight:700; cursor:pointer; font-family:inherit; }
    .caja-bulto:hover { border-color:#f59e0b; }
    .caja-bulto.on { background:#f59e0b; border-color:#d97706; color:#fff; }
    .caja-bulto-mas { border:1.5px dashed #d1d5db; background:none; color:#6b7280; border-radius:10px; padding:9px 12px; font-size:14px; font-weight:700; cursor:pointer; font-family:inherit; }
    .caja-bultos-ayuda { font-size:12px; color:#92400e; margin-top:8px; }
    .cm-bocadillo { position:absolute; left:112px; top:30px; right:54px; max-width:300px; background:#fff; color:#1B2636; border-radius:14px; padding:10px 13px; font-size:13.5px; font-weight:700; line-height:1.35; box-shadow:0 8px 22px -6px rgba(0,0,0,.45); z-index:3; cursor:pointer; animation:cmBocadilloEntra .35s cubic-bezier(.2,1.4,.4,1) both; }
    .cm-bocadillo::before { content:""; position:absolute; left:-9px; top:22px; border-style:solid; border-width:7px 11px 7px 0; border-color:transparent #fff transparent transparent; }
    .cm-bocadillo b { color:#E07A1F; }
    .cm-bocadillo.sale { animation:cmBocadilloSale .25s ease-in both; }
    @keyframes cmBocadilloEntra { from { opacity:0; transform:translateX(-10px) scale(.7); } to { opacity:1; transform:none; } }
    @keyframes cmBocadilloSale { to { opacity:0; transform:translateX(-6px) scale(.9); } }
    @media (max-width: 420px) { .cm-bocadillo { left:96px; top:20px; right:46px; font-size:12.5px; padding:8px 10px; } }
    @media (prefers-reduced-motion: reduce) { .cm-bocadillo, .cm-bocadillo.sale { animation:none; } }
    .caja-resto { background:#fff; border:1.5px solid #e5e7eb; border-radius:12px; padding:10px 12px; margin-bottom:14px; }
    .caja-resto-titulo { font-size:12.5px; font-weight:800; color:#374151; margin-bottom:7px; }
    .caja-resto-ops { display:flex; flex-wrap:wrap; gap:8px; }
    .caja-resto-op { display:flex; align-items:center; gap:7px; border:1.5px solid #e5e7eb; background:#f9fafb; border-radius:10px; padding:8px 11px; font-size:13px; font-weight:800; color:#374151; cursor:pointer; user-select:none; }
    .caja-resto-op input { width:17px; height:17px; margin:0; accent-color:#E07A1F; cursor:pointer; }
    .caja-resto-op.on { border-color:#E07A1F; background:#fef2f2; color:#991b1b; }
    .caja-resto-op.off { opacity:.45; cursor:not-allowed; }
    .caja-resto-ayuda { font-size:11.5px; color:#6b7280; margin-top:6px; }
`.replace(/\n\s+/g, "\n");
}
function jsBultosCaja_() {
    return `
    function bocadilloCajaRufo_() {
      var hoy = new Date(), clave = (typeof CARGA_DOC_ !== 'undefined' && CARGA_DOC_) ? 'carga-' + CARGA_DOC_ : hoy.getFullYear() + '-' + (hoy.getMonth() + 1) + '-' + hoy.getDate();
      try { if (localStorage.getItem('bocadilloCaja') === clave) return; localStorage.setItem('bocadilloCaja', clave); } catch (e) {}
      var cab = document.querySelector('#cm-overlay .cm-head');
      if (!cab) return;
      var viejo = cab.querySelector('.cm-bocadillo');
      if (viejo) viejo.parentNode.removeChild(viejo);
      var b = document.createElement('div');
      b.className = 'cm-bocadillo'; b.setAttribute('role', 'status'); b.title = 'Toca para cerrar';
      b.innerHTML = 'Recuerda pulsar <b>Intro</b> después de cada código para que sume 😉';
      function quitar() { if (!b.parentNode) return; b.classList.add('sale'); setTimeout(function() { if (b.parentNode) b.parentNode.removeChild(b); }, 260); }
      b.onclick = quitar;
      cab.appendChild(b);
      try { animarBocaRufoHablando(1600); } catch (eBoca) {}
      setTimeout(quitar, 8000);
    }
    function restoCajaHtml_(idUnico) {
      function op(cual, texto) {
        return '<label class="caja-resto-op" id="caja-resto-op-' + cual + '-' + idUnico + '"><input type="checkbox" id="caja-resto-' + cual + '-' + idUnico + '" onchange="marcarRestoCaja(\\'' + idUnico + '\\', \\'' + cual + '\\')"><span>' + texto + '</span></label>';
      }
      return '<div class="caja-resto"><div class="caja-resto-titulo">¿Y el resto del pedido? (opcional)</div>' +
        '<div class="caja-resto-ops">' + op('roto', 'RESTO DEL PEDIDO ROTO') + op('falta', 'RESTO DEL PEDIDO FALTA') + '</div>' +
        '<div class="caja-resto-ayuda" id="caja-resto-ayuda-' + idUnico + '">Se añade al final de la frase.</div></div>';
    }
    function restoElegidoCaja_(idUnico) {
      var r = document.getElementById('caja-resto-roto-' + idUnico), f = document.getElementById('caja-resto-falta-' + idUnico);
      return (r && r.checked && !r.disabled) ? 'roto' : ((f && f.checked && !f.disabled) ? 'falta' : '');
    }
    function restoCajaTexto_(idUnico) {
      var e = restoElegidoCaja_(idUnico);
      return e === 'roto' ? 'RESTO DEL PEDIDO ROTO' : (e === 'falta' ? 'RESTO DEL PEDIDO FALTA' : '');
    }
    function sincronizarRestoCaja_(idUnico) {
      var r = document.getElementById('caja-resto-roto-' + idUnico);
      if (!r) return;
      var f = document.getElementById('caja-resto-falta-' + idUnico), bloqueo = '';
      ['rotos', 'faltan'].forEach(function(lado) {
        var sel = document.getElementById('caja-modo-' + lado + '-' + idUnico), m = sel ? sel.value : 'normal';
        if (m === 'mayoria') bloqueo = 'No hace falta: «Indicar los buenos» ya dice lo del resto.';
        else if (m === 'todo' && !bloqueo) bloqueo = 'No hace falta: ya está marcado «' + sel.options[sel.selectedIndex].text + '».';
      });
      [r, f].forEach(function(cb) {
        cb.disabled = !!bloqueo;
        if (bloqueo) cb.checked = false;
        var et = document.getElementById('caja-resto-op-' + (cb === r ? 'roto' : 'falta') + '-' + idUnico);
        if (et) { et.classList.toggle('off', !!bloqueo); et.classList.toggle('on', cb.checked && !bloqueo); }
      });
      var ay = document.getElementById('caja-resto-ayuda-' + idUnico);
      if (ay) ay.textContent = bloqueo || 'Se añade al final de la frase.';
    }
    function marcarRestoCaja(idUnico, cual) {
      var este = document.getElementById('caja-resto-' + cual + '-' + idUnico), otro = document.getElementById('caja-resto-' + (cual === 'roto' ? 'falta' : 'roto') + '-' + idUnico);
      if (este && este.checked && otro) otro.checked = false;
      actualizarFraseCaja(idUnico);
      flashFrasePreview(idUnico);
    }
    function botonesBultosCaja_(idUnico, nb, elegidos) {
      var h = '';
      for (var n = 1; n <= nb; n++) {
        var on = elegidos.indexOf(n) > -1;
        h += '<button type="button" class="caja-bulto' + (on ? ' on' : '') + '" aria-pressed="' + (on ? 'true' : 'false') + '" onclick="alternarBultoCaja(\\'' + idUnico + '\\', ' + n + ')">Bulto ' + n + '</button>';
      }
      if (nb < 12) h += '<button type="button" class="caja-bulto-mas" title="Más bultos" onclick="masBultosCaja(\\'' + idUnico + '\\')">＋</button>';
      return h;
    }
    function bultosElegidosCaja_(cont) {
      var t = cont.getAttribute('data-bultos') || '';
      return t ? t.split(',').map(Number) : [];
    }
    function pintarBultosCaja_(idUnico) {
      var cont = document.getElementById('caja-' + idUnico);
      if (!cont) return;
      document.getElementById('caja-bultos-' + idUnico).innerHTML = botonesBultosCaja_(idUnico, Number(cont.getAttribute('data-nb')) || 4, bultosElegidosCaja_(cont));
    }
    function alternarBultoCaja(idUnico, n) {
      var cont = document.getElementById('caja-' + idUnico);
      if (!cont) return;
      var el = bultosElegidosCaja_(cont), i = el.indexOf(n);
      if (i > -1) el.splice(i, 1); else el.push(n);
      el.sort(function(a, b) { return a - b; });
      cont.setAttribute('data-bultos', el.join(','));
      pintarBultosCaja_(idUnico);
      actualizarFraseCaja(idUnico);
      flashFrasePreview(idUnico);
    }
    function masBultosCaja(idUnico) {
      var cont = document.getElementById('caja-' + idUnico);
      if (!cont) return;
      cont.setAttribute('data-nb', Math.min(12, (Number(cont.getAttribute('data-nb')) || 4) + 2));
      pintarBultosCaja_(idUnico);
      guardarBorradorCaja(idUnico);
    }
    function textoBultosCaja_(el) {
      if (el.length === 0) return '';
      if (el.length === 1) return 'FALTA BULTO NUMERO ' + el[0];
      return 'FALTAN BULTOS NUMERO ' + el.slice(0, -1).join(', ') + ' Y ' + el[el.length - 1];
    }
`.replace(/\n\s+/g, "\n");
}
function cssCambiosRet_() {
    return `    
    .cr-filtros { display:flex; flex-wrap:wrap; gap:6px; margin:12px 0 8px; }
    .cr-f { border:1px solid var(--border); background:var(--card); color:#334155; border-radius:20px; padding:5px 11px; font-size:12px; font-weight:700; cursor:pointer; font-family:inherit; }
    .cr-f.on { background:#1B2636; color:#fff; border-color:#1B2636; }
    .cr-dia { font-size:11.5px; font-weight:800; color:var(--muted); text-transform:uppercase; letter-spacing:.03em; margin:14px 2px 6px; }
    .cr-ev { background:var(--card); border:1px solid var(--border); border-left:4px solid #94a3b8; border-radius:12px; padding:10px 12px; margin:0 0 8px; box-shadow:var(--shadow-sm); font-size:12.5px; }
    .cr-ev.borrado { border-left-color:#E07A1F; } .cr-ev.cambio { border-left-color:#f59e0b; } .cr-ev.fila { border-left-color:#7c3aed; } .cr-ev.restaurado, .cr-ev.hecho { border-left-color:#059669; }
    .cr-top { display:flex; justify-content:space-between; gap:8px; color:var(--muted); font-size:12px; } .cr-top b { color:#1e293b; font-size:12.5px; }
    .cr-ped { margin:5px 0 4px; } .cr-ped code { background:#f1f5f9; padding:1px 6px; border-radius:6px; font-size:12px; }
    .cr-dif { background:#f8fafc; border:1px solid var(--border); border-radius:8px; padding:6px 8px; word-break:break-word; }
    .cr-antes { color:#b91c1c; text-decoration:line-through; } .cr-despues { color:#047857; font-weight:700; } .cr-vacio { color:#94a3b8; font-style:italic; }
    .cr-det { margin-top:6px; } .cr-det table { width:100%; border-collapse:collapse; font-size:12px; } .cr-det td { border-bottom:1px dashed var(--border); padding:3px 4px; vertical-align:top; } .cr-det td:first-child { color:var(--muted); width:38%; }
    .cr-acc { display:flex; flex-wrap:wrap; gap:6px; margin-top:8px; }
    .cr-b { border:none; border-radius:8px; padding:6px 10px; font-size:12px; font-weight:700; cursor:pointer; background:#1B2636; color:#fff; font-family:inherit; }
    .cr-b.sec { background:#e2e8f0; color:#334155; } .cr-b:disabled { opacity:.5; }
    .cr-msg { font-size:12px; font-weight:700; color:#047857; margin-top:6px; } .cr-msg:empty { display:none; }
    .cr-ok { font-size:12px; color:#047857; font-weight:700; margin-top:6px; }
    .cr-quien { font-size:11.5px; color:var(--muted); margin-top:3px; }
    .cr-nota { font-size:11.5px; color:var(--muted); margin:12px 2px 0; }
    .cr-ev.borr { border-left-color:#f59e0b; background:#fffbeb; } .cr-b.seguro { background:#fee2e2; color:#b91c1c; } 
    .cr-abrir { display:block; width:100%; margin:10px 0 0; padding:9px 12px; font-size:12.5px; }
    .cr-hist-cab { margin:14px 0 8px; font-size:12.5px; font-weight:800; color:#7c3aed; }
    .rb-hist { display:flex; align-items:center; gap:7px; font-size:12.5px; color:#475569; margin:2px 0 10px; cursor:pointer; }
    .rb-hist input { width:16px; height:16px; accent-color:#7c3aed; }
    [id^="rb-camb-"] .cr-ev { margin-top:8px; }
`.replace(/\n\s+/g, "\n");
}
function jsCambiosRet_() {
    return `    // --- Corrección 144: HISTORIAL DE CAMBIOS DE RETORNOS ---
    var crEstado = { eventos: [], filtro: 'todo' }, crPorId = {};
    var CR_TIPOS = { borrado: ['🗑️', 'Celda borrada', 'borrado'], cambio: ['✏️', 'Cambiado', 'cambio'], borrado_varias: ['🧹', 'Varias celdas borradas a la vez', 'borrado'],
      cambio_varias: ['✏️', 'Varias celdas cambiadas a la vez', 'cambio'], rellenar: ['🛠️', 'Cambiado desde «Rellenar»', 'cambio'], deshacer: ['↩️', 'Deshecho en «Rellenar»', 'cambio'],
      restaurado: ['✅', 'Restaurado', 'restaurado'], fila_eliminada: ['❌', 'Fila eliminada entera', 'fila'] };
    function crEsc(s) { return escAttr(s === undefined || s === null ? '' : String(s)); }
    function crDos(n) { return ('0' + n).slice(-2); }
    function crHora(t) { var d = new Date(t); return crDos(d.getHours()) + ':' + crDos(d.getMinutes()); }
    function crDia(t) {
      var d = new Date(t), hoy = new Date(), dias = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
      var a = new Date(hoy.getFullYear(), hoy.getMonth(), hoy.getDate()).getTime(), b = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
      var txt = dias[d.getDay()] + ' ' + crDos(d.getDate()) + '/' + crDos(d.getMonth() + 1);
      return b === a ? 'Hoy · ' + txt : (a - b === 86400000 ? 'Ayer · ' + txt : txt);
    }
    function crGrupo(ev) { return (CR_TIPOS[ev.tipo] || CR_TIPOS.cambio)[2]; }
    function crValor(v, clase) { return v === '' ? '<span class="cr-vacio">vacío</span>' : '<span class="cr-' + clase + '">' + crEsc(v) + '</span>'; }
    function crTarjeta(ev, compacta) {
      crPorId[ev.id] = ev;
      var t = CR_TIPOS[ev.tipo] || CR_TIPOS.cambio;
      var quien = ev.tipo === 'fila_eliminada' ? ((ev.desde ? 'entre ' + crHora(ev.desde) + ' y ' : 'hacia las ') + crHora(ev.t)) : (crHora(ev.t) + ' · ' + (ev.usuario && ev.usuario !== 'Usuario' ? ev.usuario : 'alguien sin identificar'));
      if (compacta) quien = crDia(ev.t).replace('Hoy · ', 'hoy ').replace('Ayer · ', 'ayer ') + ' · ' + quien;
      var h = '<div class="cr-ev ' + t[2] + (ev.restaurado ? ' hecho' : '') + '" data-id="' + crEsc(ev.id) + '">';
      h += '<div class="cr-top"><b>' + t[0] + ' ' + t[1] + '</b><span>' + crEsc(quien) + '</span></div>';
      var filaTxt = ev.tipo === 'fila_eliminada' ? 'Era la fila ' + ev.filas : (ev.filas.indexOf('-') > -1 ? 'Filas ' + ev.filas.replace('-', ' a ') : 'Fila ' + ev.filas);
      h += '<div class="cr-ped">' + crEsc(filaTxt) + (ev.pedido ? ' · <code>' + crEsc(ev.pedido) + '</code>' : '') + (ev.cliente ? ' · ' + crEsc(ev.cliente) : '') + '</div>';
      if (ev.tipo === 'fila_eliminada') {
        h += '<div class="cr-dif">Se guardó la fila completa antes de desaparecer (Google no dice quién la eliminó).</div>';
        if (ev.editando && ev.editando.length) h += '<div class="cr-quien">👥 En ese rato editaban Retornos: ' + crEsc(ev.editando.join(', ')) + '</div>';
        h += '<div class="cr-det" hidden><table>' + (ev.fila || []).map(function (x) { return '<tr><td>' + crEsc(x.col + ' · ' + x.nombre) + '</td><td>' + crEsc(x.valor) + '</td></tr>'; }).join('') + '</table></div>';
      } else if (ev.cambios.length === 1) {
        var x = ev.cambios[0];
        h += '<div class="cr-dif">' + crEsc(x.nombre) + ' (' + x.col + '): ' + crValor(x.antes, 'antes') + ' → ' + crValor(x.despues, 'despues') + '</div>';
      } else {
        var borrado = ev.cambios.every(function (c) { return c.despues === ''; });
        h += '<div class="cr-dif">' + ev.total + ' celdas ' + (borrado ? 'vaciadas' : 'cambiadas') + '</div>';
        h += '<div class="cr-det" hidden><table>' + ev.cambios.map(function (c) { return '<tr><td>' + c.f + ' · ' + crEsc(c.nombre) + '</td><td>' + crValor(c.antes, 'antes') + ' → ' + crValor(c.despues, 'despues') + '</td></tr>'; }).join('') +
          (ev.total > ev.cambios.length ? '<tr><td colspan="2">… y ' + (ev.total - ev.cambios.length) + ' más</td></tr>' : '') + '</table></div>';
      }
      if (ev.restaurado) h += '<div class="cr-ok">✅ Restaurado: ' + crEsc(ev.restaurado) + '</div>';
      var acc = '';
      if (ev.tipo === 'fila_eliminada' || ev.cambios.length > 1) acc += '<button class="cr-b sec" data-cr="ver">👀 Ver lo que había</button>';
      if (!ev.restaurado && ev.tipo !== 'restaurado') acc += '<button class="cr-b" data-cr="restaurar">↩️ ' + (ev.tipo === 'fila_eliminada' ? 'Volver a añadirla al final' : 'Restaurar') + '</button>';
      if (ev.tipo !== 'fila_eliminada' && !compacta) acc += '<button class="cr-b sec" data-cr="ir" data-fila="' + parseInt(ev.filas, 10) + '" data-ped="' + crEsc(ev.pedido || '') + '">🔍 Ir a la fila</button>';
      if (acc) h += '<div class="cr-acc">' + acc + '</div>';
      return h + '<div class="cr-msg"></div></div>';
    }
    function crInit() {
      var lista = document.getElementById('cr-lista');
      if (!lista) return;
      google.script.run.withSuccessHandler(function (r) {
        if (!r || r.error) { lista.innerHTML = '<div class="rb-no">' + crEsc(r && r.error ? r.error : 'No se pudo cargar.') + '</div>'; return; }
        crEstado.eventos = r.eventos || []; crPintar();
      }).withFailureHandler(function (err) {
        lista.innerHTML = '<div class="rb-no">⚠️ No se pudo cargar: ' + crEsc(err && err.message ? err.message : err) + '</div>';
      }).obtenerCambiosRetornos();
    }
    function crBorradores() {
      var b = {}; try { b = JSON.parse(localStorage.getItem('rbBorradores') || '{}') || {}; } catch (e) { b = {}; }
      var lim = Date.now() - 7 * 86400000;
      return Object.keys(b).filter(function (k) { return b[k] && b[k].t > lim && b[k].c && Object.keys(b[k].c).length; }).map(function (k) { var d = b[k]; d.k = k; return d; }).sort(function (a, c) { return c.t - a.t; });
    }
    function crBorradoresHtml() {
      var l = crBorradores();
      if (!l.length) return '<div class="rb-consejo">📝 No hay nada sin guardar en este ordenador. Lo que se escribe en «Rellenar» y no se guarda aparece aquí (se guarda solo en este navegador).</div>';
      return '<div class="cr-nota" style="margin:0 2px 8px">Solo en este ordenador. Se borran solos a los 7 días.</div>' + l.map(function (d) {
        return '<div class="cr-ev borr" data-k="' + crEsc(d.k) + '"><div class="cr-top"><b>📝 Sin guardar</b><span>' + crEsc(crDia(d.t).replace('Hoy · ', 'hoy ').replace('Ayer · ', 'ayer ') + ' · ' + crHora(d.t)) + '</span></div>' +
          '<div class="cr-ped">Fila ' + crEsc(d.f) + (d.p ? ' · <code>' + crEsc(d.p) + '</code>' : '') + (d.ag ? ' · ' + crEsc(d.ag) : '') + '</div>' +
          Object.keys(d.c).map(function (c) { return '<div class="cr-dif">' + crEsc((d.cab && d.cab[c]) || String.fromCharCode(64 + Number(c))) + ': ' + crValor(d.antes && d.antes[c] !== undefined ? String(d.antes[c]) : '', 'antes') + ' → ' + crValor(String(d.c[c]), 'despues') + '</div>'; }).join('') +
          '<div class="cr-acc"><button class="cr-b" data-crb="abrir">✏️ Abrir en Retornos</button><button class="cr-b sec" data-crb="descartar">Descartar</button></div></div>';
      }).join('');
    }
    document.addEventListener('click', function (e) {
      var b = e.target && e.target.closest ? e.target.closest('[data-crb]') : null;
      if (!b) return;
      var card = b.closest('[data-k]'), k = card ? card.getAttribute('data-k') : '';
      if (!k) return;
      if (b.getAttribute('data-crb') === 'abrir') {
        try { localStorage.setItem('rbzIr', JSON.stringify({ k: k, t: Date.now() })); } catch (eI) {}
        b.innerText = 'Abriendo…'; b.disabled = true;
        irSeccion_('retornos', '↩️ Retornos'); return;
      }
      if (!b.classList.contains('seguro')) { b.classList.add('seguro'); b.innerText = '¿Seguro? Descartar'; return; }
      try { var bb = JSON.parse(localStorage.getItem('rbBorradores') || '{}') || {}; delete bb[k]; localStorage.setItem('rbBorradores', JSON.stringify(bb)); } catch (eD) {}
      crPintar();
    });
    function crPintar() {
      if (!document.getElementById('cr-filtros') || !document.getElementById('cr-lista')) return;
      var evs = crEstado.eventos, n = { todo: evs.length, borrado: 0, cambio: 0, fila: 0, rellenar: 0, borr: crBorradores().length };
      evs.forEach(function (ev) { var g = crGrupo(ev); if (n[g] !== undefined) n[g]++; if (ev.tipo === 'rellenar' || ev.tipo === 'deshacer') n.rellenar++; });
      var fs = [['todo', 'Todo'], ['rellenar', '🛠️ Desde Rellenar'], ['borr', '📝 Borradores'], ['borrado', '🗑️ Borrados'], ['cambio', '✏️ Cambiados'], ['fila', '❌ Filas eliminadas']];
      document.getElementById('cr-filtros').innerHTML = fs.map(function (f) {
        return '<button class="cr-f' + (crEstado.filtro === f[0] ? ' on' : '') + '" data-cr="filtro" data-f="' + f[0] + '">' + f[1] + ' (' + n[f[0]] + ')</button>';
      }).join('');
      var lista = document.getElementById('cr-lista');
      if (crEstado.filtro === 'borr') { lista.innerHTML = crBorradoresHtml(); return; }
      if (evs.length === 0) { lista.innerHTML = '<div class="rb-consejo">🕵️ Todavía no hay cambios ni borrados apuntados en los últimos 7 días. Se apuntan desde que se instaló esta versión.</div>'; return; }
      var h = '', diaAnt = '';
      evs.forEach(function (ev) {
        if (crEstado.filtro === 'rellenar') { if (ev.tipo !== 'rellenar' && ev.tipo !== 'deshacer') return; }
        else if (crEstado.filtro !== 'todo' && crGrupo(ev) !== crEstado.filtro) return;
        var dia = crDia(ev.t);
        if (dia !== diaAnt) { h += '<div class="cr-dia">' + crEsc(dia) + '</div>'; diaAnt = dia; }
        h += crTarjeta(ev, false);
      });
      lista.innerHTML = h || '<div class="rb-consejo">No hay nada de este tipo en los últimos 7 días.</div>';
    }
    function crRestaurar(card, btn) {
      var ev = crPorId[card.getAttribute('data-id')] || {}, msg = card.querySelector('.cr-msg');
      if (!confirm(ev.tipo === 'fila_eliminada' ? '¿Volver a añadir esta fila al final de Retornos?' : '¿Restaurar lo que había en Retornos?')) return;
      btn.disabled = true; msg.textContent = 'Restaurando…';
      google.script.run.withSuccessHandler(function (r) {
        if (!r || r.error) { msg.textContent = '⚠️ ' + (r && r.error ? r.error : 'No se pudo restaurar.'); btn.disabled = false; return; }
        var t = r.restauradas > 0 ? '✅ Restaurado' + (r.filaNueva ? ' en la fila ' + r.filaNueva : '') : '⚠️ No se ha restaurado nada';
        if (r.conflictos && r.conflictos.length) t += ' · No se tocó: ' + r.conflictos.slice(0, 3).map(function (c) { return (c.col ? c.col + c.f : 'Fila ' + c.f) + ' — ' + c.motivo; }).join(' | ');
        msg.textContent = t;
        if (r.restauradas > 0) { card.classList.add('hecho'); btn.parentNode.removeChild(btn); ev.restaurado = r.marca || 'ahora'; }
        else btn.disabled = false;
      }).withFailureHandler(function (err) {
        msg.textContent = '⚠️ ' + (err && err.message ? err.message : err); btn.disabled = false;
      }).restaurarCambioRetornos(card.getAttribute('data-id'));
    }
    function crCambiosDeFila(i, btn) {
      var cont = document.getElementById('rb-camb-' + i), x = rbX(i);
      if (cont.getAttribute('data-abierto') === '1') { cont.setAttribute('data-abierto', '0'); cont.innerHTML = ''; btn.classList.remove('oscuro'); return; }
      cont.setAttribute('data-abierto', '1'); btn.classList.add('oscuro');
      cont.innerHTML = '<div class="rb-cargando"><span class="spinner"></span>Mirando el historial de esta fila…</div>';
      google.script.run.withSuccessHandler(function (r) {
        if (!r || r.error) { cont.innerHTML = '<div class="rb-no">' + crEsc(r && r.error ? r.error : 'No se pudo cargar.') + '</div>'; return; }
        cont.innerHTML = (r.eventos && r.eventos.length) ? r.eventos.map(function (ev) { return crTarjeta(ev, true); }).join('')
          : '<div class="rb-consejo">🕵️ Sin cambios ni borrados apuntados en los últimos 30 días.</div>';
      }).withFailureHandler(function (err) {
        cont.innerHTML = '<div class="rb-no">⚠️ ' + crEsc(err && err.message ? err.message : err) + '</div>';
      }).obtenerCambiosFilaRetornos(x.fila, x.datos[2] || '');
    }
    function crHistorialHtml(lista) {
      if (!lista || !lista.length) return '';
      var h = '<div class="cr-hist-cab">🕵️ En el historial: ya no está así en la hoja</div>';
      lista.forEach(function (x) {
        var el = x.tipo === 'fila_eliminada';
        var ev = { id: x.id, tipo: x.tipo, t: x.t, usuario: x.usuario, filas: String(x.fila), pedido: x.pedido, cliente: x.cliente, restaurado: x.restaurado, cambios: [], total: 1 };
        crPorId[x.id] = ev;
        h += '<div class="cr-ev ' + (el ? 'fila' : 'cambio') + (x.restaurado ? ' hecho' : '') + '" data-id="' + crEsc(x.id) + '">' +
          '<div class="cr-top"><b>' + (el ? '❌ Fila eliminada' : '✏️ Antes ponía') + '</b><span class="rb-b-pct">' + x.pct + '%</span></div>' +
          '<div class="cr-dif">' + crEsc(x.campo) + ': <mark>' + crEsc(x.valor) + '</mark>' + (el ? '' : ' → ahora ' + (x.ahora ? '«' + crEsc(x.ahora) + '»' : '<span class="cr-vacio">vacío</span>')) + '</div>' +
          '<div class="cr-ped">' + (el ? 'Era la fila ' : 'Fila ') + x.fila + (x.pedido ? ' · <code>' + crEsc(x.pedido) + '</code>' : '') + (x.cliente && x.campo !== 'Cliente' ? ' · ' + crEsc(x.cliente) : '') + '</div>' +
          '<div class="cr-quien">' + crEsc(crDia(x.t) + ' ' + crHora(x.t)) + (x.usuario ? ' · ' + crEsc(x.usuario) : '') + '</div>' +
          (x.restaurado ? '<div class="cr-ok">✅ Restaurado: ' + crEsc(x.restaurado) + '</div>' : '<div class="cr-acc"><button class="cr-b" data-cr="restaurar">↩️ ' + (el ? 'Volver a añadirla al final' : 'Restaurar') + '</button></div>') +
          '<div class="cr-msg"></div></div>';
      });
      return h;
    }
    document.addEventListener('click', function (e) {
      var b = e.target && e.target.closest ? e.target.closest('[data-cr]') : null;
      if (!b) return;
      var accion = b.getAttribute('data-cr'), card = b.closest('.cr-ev');
      if (accion === 'filtro') { crEstado.filtro = b.getAttribute('data-f'); crPintar(); return; }
      if (accion === 'fila') { crCambiosDeFila(Number(b.getAttribute('data-i')), b); return; }
      if (!card) return;
      if (accion === 'ver') { var det = card.querySelector('.cr-det'); if (det) det.hidden = !det.hidden; return; }
      if (accion === 'ir') {
        var pedIr = b.getAttribute('data-ped') || '';
        if (pedIr) copiarAlPortapapeles(pedIr, function (ok) { if (ok) b.innerText = '📋 Ref. copiada ✔️'; });
        google.script.run.activarCeldaEnHoja('Retornos', Number(b.getAttribute('data-fila')), 3); return;
      }
      if (accion === 'restaurar') crRestaurar(card, b);
    });
`.replace(/\n\s+/g, "\n");
}
function cssRecuento_() {
    return `    
    .rc-barra { display:flex; align-items:center; gap:8px; margin:0 0 12px; flex-wrap:wrap; }
    .rc-dia { font-size:13.5px; font-weight:700; color:#334155; margin-right:auto; }
    .rc-btn { background:var(--card); border:1px solid var(--border); color:#334155; border-radius:20px; padding:6px 13px; font-size:12.5px; font-weight:700; cursor:pointer; box-shadow:var(--shadow-sm); font-family:inherit; }
    .rc-btn:hover { border-color:#E07A1F; color:#E07A1F; }
    .rc-btn[hidden] { display:none; }
    .rc-cal { background:var(--card); border:1px solid var(--border); border-radius:var(--radius); padding:12px 14px; margin:0 0 14px; box-shadow:var(--shadow-sm); animation:fadeInUp .25s var(--ease); }
    .rc-cal-cab { display:flex; align-items:center; justify-content:space-between; gap:8px; margin-bottom:6px; }
    .rc-cal-cab b { font-size:14px; }
    .rc-flecha { background:none; border:1px solid var(--border); border-radius:8px; width:30px; height:28px; cursor:pointer; font-size:13px; color:#334155; }
    .rc-flecha:disabled { opacity:.3; cursor:default; }
    .rc-total { font-size:12px; color:var(--muted); text-align:center; margin-bottom:8px; }
    .rc-sem, .rc-rejilla { display:grid; grid-template-columns:repeat(7, 1fr); gap:4px; }
    .rc-sem span { font-size:11px; font-weight:700; color:var(--muted); text-align:center; padding-bottom:2px; }
    .rc-d { min-height:44px; border:1px solid var(--border); border-radius:8px; background:var(--card); padding:3px 6px; cursor:pointer; font-family:inherit; font-size:11px; line-height:1.35; display:flex; flex-direction:column; align-items:flex-start; text-align:left; color:#334155; }
    .rc-d:hover:not(:disabled) { border-color:#E07A1F; }
    .rc-d:disabled { opacity:.35; cursor:default; }
    .rc-d .rc-n { font-weight:700; font-size:12px; }
    .rc-d.rc-sin .rc-n { color:var(--muted); font-weight:400; }
    .rc-d.rc-hoy { border:2px solid #E07A1F; }
    .rc-d.rc-sel { background:#fdf2f4; border-color:#E07A1F; }
    .rc-d .rc-r { color:var(--red-dark); }
    .rc-d .rc-l { display:flex; flex-wrap:wrap; column-gap:7px; }
    .rc-d .rc-l span, .rc-d .rc-r { white-space:nowrap; }
    .rc-nota { font-size:11px; color:var(--muted); margin-top:8px; text-align:center; }
`.replace(/\n\s+/g, "\n");
}
function jsRecuento_() {
    return `    var rcEstado = { off: 0, sel: null, hoyHtml: null, hoyTitulo: null, hoyDia: null, mes: {}, dia: {} };
    function rcGuardarHoy() {
      if (rcEstado.hoyHtml !== null) return;
      rcEstado.hoyHtml = document.getElementById('rc-cuerpo').innerHTML;
      var h = document.querySelector('.encabezado-panel .rufo-nombre');
      rcEstado.hoyTitulo = h ? h.textContent : null;
      rcEstado.hoyDia = document.getElementById('rc-dia').textContent;
    }
    function rcClaveHoy() { return document.getElementById('rc-barra').getAttribute('data-hoy'); }
    function rcCalendario() {
      var cal = document.getElementById('rc-cal'), btn = document.getElementById('rc-abrir');
      rcGuardarHoy();
      if (!cal.hidden) { cal.hidden = true; btn.textContent = '📅 Otros días'; return; }
      cal.hidden = false; btn.textContent = '✕ Cerrar calendario';
      rcCargarMes(rcEstado.off);
    }
    function rcCargarMes(off) {
      var cal = document.getElementById('rc-cal');
      rcEstado.off = off;
      if (rcEstado.mes[off]) { rcPintarMes(rcEstado.mes[off]); return; }
      cal.innerHTML = '<div class="loading-row"><span class="spinner"></span>Cargando el mes…</div>';
      google.script.run.withSuccessHandler(function(m) {
        if (!m || m.error) { cal.innerHTML = '<div class="empty-inline">' + escAttr(m && m.error ? m.error : 'No se pudo cargar el mes.') + '</div>'; return; }
        rcEstado.mes[off] = m;
        if (rcEstado.off === off && !cal.hidden) rcPintarMes(m);
      }).withFailureHandler(function(err) {
        cal.innerHTML = '<div class="empty-inline">⚠️ No se pudo cargar el mes: ' + escAttr(err && err.message ? err.message : err) + '</div>';
      }).obtenerMesRecuento(off);
    }
    function rcDos(n) { return ('0' + n).slice(-2); }
    function rcPintarMes(m) {
      var sel = rcEstado.sel || m.hoy;
      var h = '<div class="rc-cal-cab"><button class="rc-flecha" title="Mes anterior" ' + (m.anterior ? '' : 'disabled ') + 'onclick="rcCargarMes(' + (m.off - 1) + ')">◀</button>'
        + '<b>' + escAttr(m.titulo) + '</b>'
        + '<button class="rc-flecha" title="Mes siguiente" ' + (m.siguiente ? '' : 'disabled ') + 'onclick="rcCargarMes(' + (m.off + 1) + ')">▶</button></div>';
      h += '<div class="rc-total">En el mes: ✅ ' + Number(m.total.ok) + ' OK · 🧾 ' + Number(m.total.fac) + ' Reclamar · 💥 ' + Number(m.total.rot) + ' rotos</div>';
      h += '<div class="rc-sem"><span>L</span><span>M</span><span>X</span><span>J</span><span>V</span><span>S</span><span>D</span></div><div class="rc-rejilla">';
      for (var i = 0; i < m.primerDia; i++) h += '<span></span>';
      for (var d = 1; d <= m.diasMes; d++) {
        var k = String(m.anio) + rcDos(m.mes + 1) + rcDos(d), info = m.dias[k], futuro = k > m.hoy;
        var cls = 'rc-d' + (info ? '' : ' rc-sin') + (k === m.hoy ? ' rc-hoy' : '') + (k === sel ? ' rc-sel' : '');
        h += '<button class="' + cls + '" ' + (futuro ? 'disabled ' : '') + 'onclick="rcVerDia(\\'' + k + '\\')"><span class="rc-n">' + d + '</span>';
        if (info) {
          h += '<span class="rc-l"><span>✅ ' + Number(info.ok) + '</span>' + (info.fac ? '<span>🧾 ' + Number(info.fac) + '</span>' : '') + '</span>';
          if (info.rot) h += '<span class="rc-r">💥 ' + Number(info.rot) + '</span>';
        }
        h += '</button>';
      }
      h += '</div><div class="rc-nota">Pulsa un día para ver su recuento completo.</div>';
      document.getElementById('rc-cal').innerHTML = h;
    }
    function rcMostrarDia(r) {
      document.getElementById('rc-cuerpo').innerHTML = r.html;
      document.getElementById('rc-dia').textContent = r.titulo;
      var t = document.querySelector('.encabezado-panel .rufo-nombre');
      if (t) t.textContent = '📊 Recuento del ' + r.titulo;
      document.getElementById('rc-volver').hidden = false;
    }
    function rcVerDia(k) {
      rcGuardarHoy();
      if (k === rcClaveHoy()) { rcVolverHoy(); return; }
      rcEstado.sel = k;
      var cal = document.getElementById('rc-cal');
      cal.hidden = true; document.getElementById('rc-abrir').textContent = '📅 Otros días';
      if (rcEstado.dia[k]) { rcMostrarDia(rcEstado.dia[k]); return; }
      var cuerpo = document.getElementById('rc-cuerpo');
      cuerpo.innerHTML = '<div class="loading-row"><span class="spinner"></span>Contando ese día…</div>';
      google.script.run.withSuccessHandler(function(r) {
        if (rcEstado.sel !== k) return;
        if (!r || r.error) { cuerpo.innerHTML = '<div class="empty-inline">' + escAttr(r && r.error ? r.error : 'No se pudo cargar ese día.') + '</div>'; document.getElementById('rc-volver').hidden = false; return; }
        rcEstado.dia[k] = r; rcMostrarDia(r);
      }).withFailureHandler(function(err) {
        if (rcEstado.sel !== k) return;
        cuerpo.innerHTML = '<div class="empty-inline">⚠️ No se pudo cargar ese día: ' + escAttr(err && err.message ? err.message : err) + '</div>';
        document.getElementById('rc-volver').hidden = false;
      }).obtenerRecuentoDia(k);
    }
    function rcVolverHoy() {
      rcGuardarHoy();
      rcEstado.sel = null;
      document.getElementById('rc-cuerpo').innerHTML = rcEstado.hoyHtml;
      document.getElementById('rc-dia').textContent = rcEstado.hoyDia;
      var t = document.querySelector('.encabezado-panel .rufo-nombre');
      if (t && rcEstado.hoyTitulo !== null) t.textContent = rcEstado.hoyTitulo;
      document.getElementById('rc-volver').hidden = true;
      var cal = document.getElementById('rc-cal');
      if (!cal.hidden && rcEstado.mes[rcEstado.off]) rcPintarMes(rcEstado.mes[rcEstado.off]);
    }
`.replace(/\n\s+/g, "\n");
}
function cssRetornos_() {
    return `    
    .rb-card { margin-top:12px; background:var(--card); border:1.5px dashed #cbd5e1; border-radius:var(--radius); padding:10px 12px; display:flex; gap:8px; align-items:center; }
    .rb-card svg, .rb-cab svg { width:64px; height:55px; flex-shrink:0; animation:rbSalto 1.8s ease-in-out infinite; }
    .rb-card b, .rb-cab b { display:block; font-size:13.5px; }
    .rb-card span, .rb-cab span { font-size:12px; color:var(--muted); }
    .rb-abrir { margin-top:7px; background:#1B2636; color:#fff; border:none; border-radius:20px; padding:5px 12px; font-size:12px; font-weight:700; cursor:pointer; }
    @keyframes rbSalto { 0%,100% { transform:translateY(0) rotate(0); } 50% { transform:translateY(-3px) rotate(-3deg); } }
    @keyframes rbBusca { 0%,100% { transform:translateX(0) rotate(-4deg); } 50% { transform:translateX(6px) rotate(4deg); } }
    .rb-box { margin-top:12px; background:var(--card); border:1.5px solid #1B2636; border-radius:var(--radius); padding:12px; }
    .rb-cab { display:flex; gap:8px; align-items:center; margin-bottom:10px; }
    .rb-cerrar { margin-left:auto; background:none; border:none; color:var(--muted); font-size:16px; cursor:pointer; }
    .rb-campo { width:100%; margin-bottom:8px; }
    .rb-modos { display:flex; gap:6px; margin-bottom:8px; }
    .rb-modo { flex:1; border:1.5px solid var(--border); border-radius:20px; padding:5px 6px; font-size:12px; font-weight:700; color:var(--muted); background:#fff; cursor:pointer; }
    .rb-modo.on { border-color:#1B2636; background:#1B2636; color:#fff; }
    .rb-btn { background:linear-gradient(135deg,#2b2b2b,#0e0e0e); box-shadow:none; }
    .rb-consejo { font-size:11.5px; color:var(--muted); margin-top:8px; line-height:1.45; }
    .rb-buscando { text-align:center; padding:12px 6px 6px; }
    .rbl-comp { display:flex; align-items:center; gap:6px; font-size:11.5px; font-weight:700; color:#6b7280; margin:0 0 8px; }
    .rbl-comp i { width:9px; height:9px; border-radius:50%; border:2px solid #cbd5e1; border-top-color:#f59e0b; animation:rblGira .8s linear infinite; flex-shrink:0; }
    @keyframes rblGira { to { transform:rotate(360deg); } }
    .rbl-nuevo { display:block; width:100%; margin:0 0 10px; padding:9px 10px; border:1.5px solid #f59e0b; background:#fffbeb; color:#92400e; border-radius:10px; font:inherit; font-size:12.5px; font-weight:700; cursor:pointer; }
    .rb-buscando svg { width:110px; height:95px; animation:rbBusca 1s ease-in-out infinite; }
    .rb-buscando p { margin:4px 0 0; font-weight:700; font-size:13px; }
    .rb-no { margin-top:10px; text-align:center; font-size:13px; color:#b91c1c; background:#fef2f2; border:1px solid #fecaca; border-radius:10px; padding:10px; }
    .rb-resumen { font-size:12.5px; margin:10px 2px 6px; color:var(--muted); }
    .rb-resumen b { color:var(--text); }
    .rb-filtros { display:flex; gap:6px; flex-wrap:wrap; margin-bottom:8px; }
    .rb-filtro { border:1.5px solid var(--border); border-radius:20px; padding:4px 10px; font-size:12px; font-weight:700; background:#fff; color:var(--muted); cursor:pointer; }
    .rb-filtro.on { background:#1B2636; border-color:#1B2636; color:#fff; }
    .rb-filtro.v { border-color:#86efac; color:#15803d; } .rb-filtro.a { border-color:#fcd34d; color:#b45309; }
    .rb-filtro.v.on { background:#15803d; border-color:#15803d; color:#fff; } .rb-filtro.a.on { background:#d97706; border-color:#d97706; color:#fff; }
    .rb-item { border:1px solid var(--border); border-left:6px solid #f59e0b; border-radius:var(--radius-sm); background:#f8fafc; padding:10px 11px; margin-bottom:9px; }
    .rb-item.proc { border-left-color:#22c55e; }
    .rb-estado { display:inline-block; border-radius:6px; padding:2px 8px; font-size:11.5px; font-weight:800; margin-bottom:4px; background:#fef3c7; color:#92400e; }
    .rb-estado.proc { background:#dcfce7; color:#166534; }
    .rb-top { display:flex; justify-content:space-between; gap:6px; align-items:center; font-size:11.5px; color:var(--muted); }
    .rb-badges { display:flex; gap:4px; }
    .rb-b-exacta, .rb-b-similar, .rb-b-pct { border-radius:6px; padding:1px 7px; font-weight:700; font-size:11px; }
    .rb-b-exacta { background:#dcfce7; color:#166534; } .rb-b-similar { background:#fef9c3; color:#854d0e; } .rb-b-pct { background:#e0e7ff; color:#3730a3; }
    .rb-ref { font-size:17px; font-weight:800; margin:4px 0 2px; word-break:break-all; }
    .rb-ref small { display:block; font-size:11px; color:var(--muted); font-weight:700; }
    .rb-chip { display:inline-block; background:#1B2636; color:#fff; border-radius:6px; padding:1px 7px; font-size:11px; font-weight:700; vertical-align:middle; word-break:normal; }
    .rb-match { font-size:12.5px; margin:4px 0 2px; word-break:break-word; }
    .rb-match mark { background:#fef3c7; border-radius:3px; padding:0 3px; color:inherit; }
    .rb-porque { font-size:11.5px; color:var(--muted); margin-bottom:6px; }
    .rb-vista { display:grid; grid-template-columns:1fr 1.3fr 1.1fr; border:1px solid #475569; font-size:11px; font-weight:700; margin:6px 0 3px; border-radius:3px; overflow:hidden; }
    .rb-vista div { padding:4px 5px; border-right:1px solid #475569; text-align:center; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
    .rb-vista div:last-child { border-right:none; }
    .rb-vista-cap { font-size:10.5px; color:var(--muted); margin-bottom:6px; }
    .rb-datos dl { display:grid; grid-template-columns:auto 1fr; gap:3px 10px; margin:6px 0 8px; font-size:12px; }
    .rb-datos dt { color:var(--muted); } .rb-datos dd { margin:0; font-weight:600; word-break:break-word; }
    .rb-datos dd.resalte { background:#fef3c7; border-radius:4px; padding:0 4px; }
    .rb-ver { background:none; border:none; padding:0; font-size:12px; font-weight:700; color:var(--primary-dark); cursor:pointer; margin-bottom:6px; }
    .rb-acciones { display:flex; gap:6px; flex-wrap:wrap; }
    .rb-acc { flex:1 1 auto; border:1px solid var(--border); background:var(--card); border-radius:8px; padding:7px 8px; font-size:12px; font-weight:700; cursor:pointer; color:var(--text); }
    .rb-acc.azul { background:var(--primary-dark); border-color:var(--primary-dark); color:#fff; }
    .rb-acc.oscuro { background:#1B2636; border-color:#1B2636; color:#fff; }
    .rb-cargando { font-size:12.5px; color:var(--muted); padding:10px 2px; display:flex; gap:8px; align-items:center; }
    .rb-aviso-fila { margin-top:8px; background:#fffbeb; border:1px solid #fde68a; color:#92400e; border-radius:10px; padding:9px 10px; font-size:12.5px; }
    .rb-hecho { margin-top:8px; display:flex; justify-content:space-between; align-items:center; gap:8px; background:#ecfdf5; border:1px solid #a7f3d0; color:#065f46; border-radius:10px; padding:8px 10px; font-size:12.5px; font-weight:700; }
    .rb-deshacer { background:none; border:none; color:#065f46; font-weight:800; text-decoration:underline; cursor:pointer; white-space:nowrap; }
    .rb-rell { margin-top:8px; border:1.5px solid #1B2636; border-radius:12px; background:#fff; padding:10px; }
    .rb-proc { border:1.5px solid #16a34a; border-radius:12px; padding:10px; background:#f0fdf4; margin-bottom:10px; }
    .rb-proc.hecho { display:flex; justify-content:space-between; align-items:center; gap:8px; font-size:12.5px; color:#166534; }
    .rb-proc-btn { display:block; width:100%; border:none; border-radius:10px; padding:11px; font-weight:800; font-size:14px; background:#16a34a; color:#fff; cursor:pointer; }
    .rb-proc-sub { font-size:11.5px; color:#166534; margin:6px 0; line-height:1.4; }
    .rb-quitar { background:#fff; color:#16a34a; border:1.5px solid #16a34a; border-radius:8px; padding:5px 9px; font-size:12px; font-weight:700; cursor:pointer; white-space:nowrap; }
    .rb-ops { display:flex; gap:5px; flex-wrap:wrap; }
    .rb-op { border:1.5px solid var(--border); border-radius:14px; padding:3px 10px; font-size:12px; font-weight:700; background:#fff; color:#334155; cursor:pointer; }
    .rb-op.on { background:#1B2636; border-color:#1B2636; color:#fff; }
    .rb-op.on.cambiado { background:#d97706; border-color:#d97706; }
    .rb-toggle { display:flex; align-items:center; justify-content:space-between; gap:8px; font-size:12.5px; font-weight:700; margin:4px 2px 8px; }
    .rb-sw { width:34px; height:19px; border-radius:10px; background:#1B2636; position:relative; flex-shrink:0; border:none; cursor:pointer; }
    .rb-sw::after { content:""; position:absolute; right:2px; top:2px; width:15px; height:15px; border-radius:50%; background:#fff; }
    .rb-sw.off { background:#cbd5e1; } .rb-sw.off::after { right:auto; left:2px; }
    .rb-grupo { border:1px solid var(--border); border-radius:10px; margin-bottom:8px; overflow:hidden; background:#fff; }
    .rb-grupo-cab { background:#f8fafc; padding:6px 10px; font-size:11.5px; font-weight:800; color:#475569; text-transform:uppercase; letter-spacing:.04em; display:flex; justify-content:space-between; }
    .rb-grupo-cab span { font-weight:700; text-transform:none; letter-spacing:0; color:var(--muted); }
    .rb-cf { padding:7px 10px; border-top:1px solid #f1f5f9; }
    .rb-cf:first-of-type { border-top:none; }
    .rb-cf label { display:flex; justify-content:space-between; font-size:11px; font-weight:800; color:#64748b; margin-bottom:3px; }
    .rb-cf label i { font-style:normal; font-weight:700; color:#94a3b8; }
    .rb-fila-inp { display:flex; gap:6px; }
    .rb-inp { flex:1; width:100%; padding:7px 9px; border:1.5px solid var(--border); border-radius:8px; font-size:13px; background:#f8fafc; color:var(--text); font-family:inherit; }
    .rb-inp:focus { outline:none; border-color:var(--primary-dark); background:#fff; }
    .rb-inp.cambiado { border-color:#f59e0b; background:#fffbeb; }
    .rb-hoy { border:1px solid var(--border); border-radius:8px; padding:6px 9px; font-size:12px; font-weight:700; background:#fff; cursor:pointer; }
    .rb-barra { display:flex; gap:8px; align-items:center; margin-top:6px; }
    .rb-guardar { flex:1; border:none; border-radius:10px; padding:10px; font-weight:800; font-size:13.5px; background:linear-gradient(135deg,var(--primary-2),var(--primary-dark)); color:#fff; cursor:pointer; }
    .rb-guardar:disabled { opacity:.45; cursor:default; }
    .rb-descartar { border:1px solid var(--border); background:#fff; border-radius:10px; padding:9px 10px; font-size:12px; font-weight:700; cursor:pointer; }
    .rb-modos { gap:5px; } .rb-modo { white-space:nowrap; padding:5px 3px; } 
    .rb-rell-cab { display:flex; justify-content:space-between; align-items:baseline; gap:8px; margin:2px 2px 2px; }
    .rb-rell-cab b { font-size:13.5px; } .rb-rell-cab span { font-size:11.5px; font-weight:700; color:#b45309; }
    .rb-rell-cab span.lleno { color:#15803d; }
    .rb-rell-ayuda { font-size:11.5px; color:var(--muted); margin:0 2px 8px; line-height:1.4; }
    .rb-cf.vacio { background:#fffbeb; } .rb-cf.vacio label i { color:#b45309; }
    .rb-cf.cambiado { background:#fff7ed; } .rb-cf.cambiado label span::after { content:" · cambiado"; color:#d97706; }
    .rbz-aviso { margin:10px 0 12px; border:2px solid #f59e0b; background:linear-gradient(180deg,#fffbeb,#fef3c7); border-radius:14px; padding:10px 11px; box-shadow:0 6px 16px -8px rgba(217,119,6,.45); animation:rbzLatido 2.4s ease-in-out infinite; }
    @keyframes rbzLatido { 0%,100% { box-shadow:0 6px 16px -8px rgba(217,119,6,.45); } 50% { box-shadow:0 6px 22px -4px rgba(217,119,6,.7); } }
    .rbz-aviso .rbz-t { font-weight:800; color:#92400e; font-size:14px; display:flex; align-items:center; gap:6px; }
    .rbz-aviso .rbz-t em { font-style:normal; background:#d97706; color:#fff; border-radius:999px; padding:1px 8px; font-size:12px; }
    .rbz-aviso .rbz-s { font-size:11.5px; color:#a16207; margin:3px 0 4px; }
    .rbz-it { background:#fff; border:1px solid #fcd34d; border-radius:10px; padding:8px 9px; margin-top:7px; }
    .rbz-it .rbz-ref { font-weight:800; color:#1e293b; font-size:13.5px; word-break:break-all; } .rbz-it .rbz-ref small { font-weight:600; color:#64748b; font-size:11px; margin-left:4px; }
    .rbz-it .rbz-cam { font-size:12px; color:#334155; margin:4px 0 6px; line-height:1.35; word-break:break-word; } .rbz-it .rbz-cam b { color:#b45309; }
    .rbz-it .rbz-h { font-size:11px; color:#94a3b8; }
    .rbz-bt { display:flex; gap:6px; margin-top:6px; } .rbz-bt button { border:0; border-radius:9px; padding:7px 9px; font-weight:700; font-size:12.5px; cursor:pointer; font-family:inherit; }
    .rbz-g { background:#16a34a; color:#fff; flex:1; } .rbz-d { background:#f1f5f9; color:#475569; } .rbz-d.seguro { background:#fee2e2; color:#b91c1c; }
    .rbz-bt { flex-wrap:wrap; } .rbz-r { background:#fde68a; color:#78350f; } .rbz-bt button:disabled { opacity:.6; }
    .rbz-hecho { margin:8px 0; background:#ecfdf5; border:1.5px solid #34d399; color:#065f46; border-radius:10px; padding:8px 10px; font-size:12.5px; font-weight:800; }
    .rb-guardar.rbz-latir:not(:disabled) { animation:rbzLatirB 1.4s ease-in-out infinite; }
    @keyframes rbzLatirB { 50% { box-shadow:0 0 0 5px rgba(22,163,74,.35); } }
    .rbz-no { font-size:12px; color:#b91c1c; margin-top:6px; }
    .rbz-ult { margin:10px 0 4px; border:1px solid #e2e8f0; background:#f8fafc; border-radius:12px; padding:8px 10px; }
    .rbz-ult .rbz-ut { font-weight:800; font-size:12.5px; color:#334155; display:flex; justify-content:space-between; } .rbz-ult .rbz-ut span { font-weight:600; color:#64748b; }
    .rbz-ul { display:flex; align-items:center; gap:7px; padding:6px 0; border-top:1px dashed #e2e8f0; font-size:12px; color:#334155; } .rbz-ut + .rbz-ul { border-top:0; }
    .rbz-ul i { font-style:normal; color:#64748b; min-width:34px; } .rbz-ul b { flex:1; min-width:0; word-break:break-all; } .rbz-ul small { display:block; font-weight:500; color:#64748b; word-break:break-word; }
    .rbz-ul button { border:0; background:#1d4ed8; color:#fff; border-radius:8px; padding:5px 8px; font-size:11.5px; font-weight:700; cursor:pointer; font-family:inherit; }
    .rbz-pend { margin:10px 0 6px; background:#fff7ed; border:1.5px solid #fb923c; color:#9a3412; border-radius:10px; padding:7px 10px; font-size:12.5px; font-weight:700; display:flex; align-items:center; gap:7px; }
    .rbz-pend .rbz-pt { flex:none; width:9px; height:9px; border-radius:50%; background:#f97316; animation:rbzPunto 1s ease-in-out infinite; }
    @keyframes rbzPunto { 50% { opacity:.25; } }
    .rbz-pend small { display:block; font-weight:500; color:#c2410c; font-size:11px; }
    .rbz-rec { margin:0 0 8px; background:#fef3c7; border:1px solid #fcd34d; color:#92400e; border-radius:10px; padding:6px 9px; font-size:12px; font-weight:700; }
    .rbz-badge { display:inline-block; background:#fef3c7; color:#92400e; border:1px solid #fcd34d; border-radius:999px; padding:1px 8px; font-size:11px; font-weight:800; margin-left:6px; }
    .rbz-cerrar { border:2px solid #f97316; background:#fff7ed; border-radius:14px; padding:12px; }
    .rbz-cerrar .rbz-t { font-weight:800; color:#9a3412; font-size:14.5px; margin-bottom:6px; }
    .rbz-cerrar ul { margin:0 0 10px 0; padding-left:18px; font-size:12.5px; color:#334155; line-height:1.5; word-break:break-word; }
    .rbz-cerrar button { display:block; width:100%; border:0; border-radius:10px; padding:10px; font-weight:800; font-size:13.5px; margin-top:7px; cursor:pointer; font-family:inherit; }
    .rbz-cerrar .g { background:#16a34a; color:#fff; } .rbz-cerrar .b { background:#fde68a; color:#78350f; } .rbz-cerrar .x { background:#f1f5f9; color:#64748b; font-weight:700; }
    @media (prefers-reduced-motion: reduce) { .rbz-aviso, .rbz-pend .rbz-pt { animation:none !important; } }
    .rb-resumen-cambios { border:1.5px solid #f59e0b; background:#fffbeb; border-radius:10px; padding:8px 10px; margin:2px 0 4px; font-size:12px; }
    .rb-resumen-cambios > b { display:block; margin-bottom:4px; color:#92400e; }
    .rb-rc { display:flex; justify-content:space-between; align-items:center; gap:8px; padding:3px 0; border-top:1px dashed #fde68a; word-break:break-word; }
    .rb-rc s { color:#94a3b8; } .rb-rc i { color:#94a3b8; }
    .rb-rc-x { flex-shrink:0; border:1px solid #fcd34d; background:#fff; color:#92400e; border-radius:6px; padding:1px 7px; font-weight:800; cursor:pointer; }
    .bn-toggle { display:flex; align-items:center; gap:8px; margin:-4px 0 12px; padding:7px 10px; border:1.5px dashed #cbd5e1; border-radius:10px; font-size:12.5px; font-weight:700; color:#475569; cursor:pointer; user-select:none; }
    .bn-toggle input { width:16px; height:16px; margin:0; accent-color:#1B2636; cursor:pointer; }
    .bn-toggle.on { border-style:solid; border-color:#1B2636; background:#1B2636; color:#fff; }
    .bn-toggle.on input { accent-color:#fff; }
    .bn-resumen { font-size:12.5px; color:var(--muted); margin:2px 2px 8px; }
    .bn-resumen b { color:var(--text); }
    .res-nombre { font-size:12.5px; font-weight:700; color:#334155; margin-top:2px; word-break:break-word; }
    .mp-zona { position:relative; margin-top:70px; }
    .mp-zona > .search-card, .mp-zona > .rb-box { position:relative; z-index:2; }
    .mp-cabeza { position:absolute; top:-38px; right:28px; width:68px; height:66px; z-index:1; cursor:pointer; background:none; border:none; padding:0; margin:0; -webkit-tap-highlight-color:transparent; transform:translateY(-1px) rotate(-8deg); transition:transform .35s cubic-bezier(.3,1.5,.5,1); animation:mpAsoma 5s ease-in-out infinite; }
    .mp-cabeza svg { width:68px; height:66px; display:block; overflow:visible; }
    .mp-cabeza.rubia { width:78px; right:22px; }
    .mp-cabeza.rubia svg { width:78px; }
    .mp-cabeza:hover, .mp-zona.aviso .mp-cabeza { animation:none; transform:translateY(-9px) rotate(-3deg); }
    .mp-cabeza:focus-visible { animation:none; transform:translateY(-9px) rotate(-3deg); outline:none; }
    .mp-cabeza.rubia:hover { transform:translateY(-21px) rotate(-3deg); }
    .mp-zona.aviso .mp-cabeza { animation:mpAvisa 1.8s ease-in-out infinite; }
    .mp-cabeza .mp-susto, .mp-cabeza .mp-exc { opacity:0; }
    .mp-cabeza.susto, .mp-cabeza.susto:hover, .mp-zona.aviso .mp-cabeza.susto { animation:mpSalta .9s ease-out forwards; }
    .mp-cabeza.susto .mp-susto { opacity:1; }
    .mp-cabeza.susto .ot-ojo { animation:otRapido .12s ease-in-out 9; } 
    .mp-cabeza.susto .ot-pest { transform-box:fill-box; transform-origin:50% 100%; animation:otAleteo .12s ease-in-out 9; }
    .mp-cabeza.susto .mp-exc { animation:mpExc .5s ease-out .05s forwards; }
    .mp-bocadillo { position:absolute; top:-30px; right:104px; z-index:3; background:#1B2636; color:#fff; font:inherit; font-size:11.5px; font-weight:700; padding:5px 9px; border:none; border-radius:10px; white-space:nowrap; opacity:0; transform:translateY(4px); transition:opacity .2s, transform .2s; pointer-events:none; cursor:pointer; }
    .mp-bocadillo::after { content:''; position:absolute; right:-5px; top:50%; margin-top:-5px; border:5px solid transparent; border-right:0; border-left-color:#1B2636; }
    .mp-bocadillo .mp-b-aviso { display:none; }
    .mp-cabeza:hover + .mp-bocadillo { opacity:1; transform:none; }
    .mp-cabeza:focus-visible + .mp-bocadillo { opacity:1; transform:none; }
    .mp-zona.aviso .mp-bocadillo { opacity:1; transform:none; pointer-events:auto; background:#fde047; color:#1B2636; box-shadow:0 3px 10px -3px rgba(0,0,0,.35); animation:mpBoca .35s ease-out; }
    .mp-zona.aviso .mp-bocadillo::after { border-left-color:#fde047; }
    .mp-zona.aviso .mp-bocadillo .mp-b-ir { display:none; }
    .mp-zona.aviso .mp-bocadillo .mp-b-aviso { display:inline; }
    .mp-cabeza.susto + .mp-bocadillo { opacity:0 !important; pointer-events:none !important; }
    @keyframes mpAsoma { 0%,70%,100% { transform:translateY(-1px) rotate(-8deg); } 78% { transform:translateY(-6px) rotate(-12deg); } 86% { transform:translateY(-4px) rotate(-4deg); } }
    @keyframes mpAvisa { 0%,100% { transform:translateY(-9px) rotate(-3deg); } 50% { transform:translateY(-12px) rotate(-7deg); } }
    @keyframes mpBoca { 0% { opacity:0; transform:translateY(6px) scale(.85); } 100% { opacity:1; transform:none; } }
    @keyframes mpSalta { 0% { transform:translateY(-1px) rotate(-8deg) scale(1,1); } 12% { transform:translateY(4px) rotate(-6deg) scale(1.08,.9); } 32% { transform:translateY(-30px) rotate(0) scale(.94,1.1); } 46% { transform:translateY(-24px) rotate(0) scale(1.06,.96); } 56% { transform:translateY(-26px) rotate(-4deg) scale(1); } 64% { transform:translateY(-26px) rotate(4deg); } 72% { transform:translateY(-26px) rotate(-3deg); } 80% { transform:translateY(-26px) rotate(2deg); } 100% { transform:translateY(-26px) rotate(0); } }
    @keyframes mpExc { 0% { opacity:0; transform:translate(0,6px) scale(.4); } 60% { opacity:1; transform:translate(0,-3px) scale(1.2); } 100% { opacity:1; transform:none; } }
    @media (prefers-reduced-motion: reduce) { .rb-card svg, .rb-cab svg, .rb-buscando svg, .mp-cabeza, .mp-zona.aviso .mp-cabeza { animation:none !important; } }
`.replace(/\n\s+/g, "\n");
}
function jsRetornos_() {
    var _awq = codigoBusquedaRetornosCliente_();
    return `    // --- Corrección 131: BEBÉ RUFO · BUSCADOR Y RELLENAR DE RETORNOS ---
    var rbEstado = { modo: 'todo', res: null, filtro: 'todos', texto: '' };
    var RBL_ = null;
    ${_awq ? "try { RBL_ = (function() {\n" + _awq + "\n})(); } catch (eRbl) { RBL_ = null; }" : ""}
    var rblDatos = null, rblPidiendo = false, rblViejo = false, RBL_MAX_MS_ = 10 * 60000;
    function rblCargar_(forzar) {
      if (!RBL_ || rblPidiendo) return;
      if (!rblDatos) { var c = cacheLeer_('retdatos'); if (c && c.d && c.d.v === CACHE_V_) { rblDatos = c.d; rblDatos.recibido = c.t; } }
      if (!forzar && !rblViejo && rblDatos && Date.now() - rblDatos.recibido < RBL_MAX_MS_) return;
      rblPidiendo = true;
      google.script.run.withSuccessHandler(function(D) {
        rblPidiendo = false;
        if (!D || typeof D.n !== 'number') return;
        D.recibido = Date.now(); rblDatos = D; rblViejo = false;
        try { localStorage.setItem('pc_retdatos', JSON.stringify({ v: CACHE_V_, dia: cacheDia_(), t: D.recibido, d: D })); }
        catch (eLleno) { try { localStorage.removeItem('pc_retdatos'); } catch (e2) {} }
      }).withFailureHandler(function() { rblPidiendo = false; }).obtenerDatosRetornosPanel();
    }
    function rblFirma_(r) {
      if (!r || r.error || !r.resultados) return 'x';
      return JSON.stringify([r.total || 0, r.resultados.map(function(x) { return [x.fila, x.tipo, x.col, x.valor, x.pct, x.exacta, x.porque]; })]);
    }
    function rblEstado_(html) { var el = document.getElementById('rbl-estado'); if (el) el.innerHTML = html || ''; }
    function rblCompletar_(loc, r) {
      var cambioProc = false;
      loc.resultados.forEach(function(x, i) {
        var y = r.resultados[i]; if (x.procesado !== y.procesado) cambioProc = true;
        x.datos = y.datos; x.procesado = y.procesado; x.color = y.color; x.parcial = false;
      });
      loc.cabeceras = r.cabeceras; loc.historial = r.historial; loc.local = false;
      var abiertos = Object.keys(rbForms).some(function(i) { return rbForms[i]; }) || [].some.call(document.querySelectorAll('[id^="rb-camb-"], [id^="rb-form-"]'), function(el) { return el.innerHTML.trim() !== ''; });
      if (!abiertos && cambioProc && rbEstado.filtro !== 'todos') { rbPintar(); return; } // cambia qué se ve con el filtro
      loc.resultados.forEach(function(x, i) { rblRepintarItem_(i); });
      var cab = document.getElementById('rb-cabres'); if (cab) cab.innerHTML = rbCabHtml_(loc.resultados, loc.total);
      rblEstado_('');
    }
    function rblRepintarItem_(i) {
      var el = document.getElementById('rb-item-' + i), x = rbX(i); if (!el || !x) return;
      var tmp = document.createElement('div'); tmp.innerHTML = rbItemHtml(x, i);
      el.className = 'rb-item ' + (x.procesado ? 'proc' : 'pend');
      ['.rb-estado', '.rb-ref', '.rb-vista', '.rb-vista-cap', '.rb-datos'].forEach(function(sel) {
        var viejo = el.querySelector(sel), nuevo = tmp.querySelector(sel); if (!viejo || !nuevo) return;
        if (sel === '.rb-datos') nuevo.style.display = viejo.style.display;
        viejo.parentNode.replaceChild(nuevo, viejo);
      });
    }
    function rblAvisoNuevo_(r) {
      var el = document.getElementById('rbl-estado'); if (!el) return;
      el.innerHTML = '<button class="rbl-nuevo" type="button">🔄 En la hoja hay cambios desde que se cargó esta lista · pulsa para verla al día</button>';
      el.firstChild.onclick = function() { rbEstado.res = r; rbEstado.filtro = 'todos'; rbForms = {}; rbPintar(); try { rbzTrasPintar_(); } catch (e) {} };
    }
    var rbForms = {};
    var RB_GRUPOS = [['📦 Pedido y cliente', [3, 8, 10]], ['🔎 Revisión', [11, 12, 4, 14]], ['🚚 Llegada', [1, 2]]];
    var RB_CAMPOS = [3, 8, 10, 11, 12, 4, 14, 1, 2];
    function rbEsc(s) { return escAttr(s); }
    var RBZ_BORR = 'rbBorradores', RBZ_GUARD = 'rbGuardados', RBZ_DIAS = 7, rbzAbrirTras = null;
    function rbzLeer_(k) { try { var v = JSON.parse(localStorage.getItem(k) || 'null'); return v || null; } catch (e) { return null; } }
    function rbzEscribir_(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); } catch (e) {} }
    function rbzClave_(ped, fila) { var n = String(ped || '').trim().toLowerCase().replace(/^0+/, ''); return n ? 'p:' + n : 'f:' + fila; }
    function rbzClaveX_(x) { return rbzClave_(x && x.datos ? x.datos[2] : '', x ? x.fila : ''); }
    function rbzBorradores_() {
      var b = rbzLeer_(RBZ_BORR) || {}, lim = Date.now() - RBZ_DIAS * 86400000, cambio = false;
      Object.keys(b).forEach(function(k) { if (!b[k] || !(b[k].t > lim) || !b[k].c || !Object.keys(b[k].c).length) { delete b[k]; cambio = true; } });
      if (cambio) rbzEscribir_(RBZ_BORR, b);
      return b;
    }
    function rbzGuardarBorrador_(i) {
      var f = rbForms[i], x = rbX(i); if (!f || !x) return;
      var b = rbzBorradores_(), k = rbzClaveX_(x), n = Object.keys(f.cambios).length;
      if (n === 0) { if (b[k]) { delete b[k]; rbzEscribir_(RBZ_BORR, b); rbzPintarAviso_(); } return; }
      var cab = {}, antes = {}; Object.keys(f.cambios).forEach(function(c) { cab[c] = f.d.cabeceras[c - 1] || rbLetra(Number(c)); antes[c] = String(f.d.valores[c - 1]); });
      var nuevo = !b[k];
      b[k] = { p: String(x.datos[2] || ''), f: x.fila, ag: String(x.datos[1] || ''), c: JSON.parse(JSON.stringify(f.cambios)), cab: cab, antes: antes, t: Date.now() };
      rbzEscribir_(RBZ_BORR, b);
      if (nuevo) rbzPintarAviso_();
    }
    function rbzQuitarBorrador_(k) { var b = rbzBorradores_(); if (b[k]) { delete b[k]; rbzEscribir_(RBZ_BORR, b); } rbzPintarAviso_(); }
    function rbzHace_(t) {
      var m = Math.round((Date.now() - t) / 60000);
      if (m < 1) return 'ahora mismo'; if (m < 60) return 'hace ' + m + ' min';
      var h = Math.floor(m / 60); if (h < 24) return 'hace ' + h + ' h' + (m % 60 ? ' ' + (m % 60) + ' min' : '');
      var d = Math.floor(h / 24); return 'hace ' + d + ' día' + (d === 1 ? '' : 's');
    }
    function rbzHora_(t) { var d = new Date(t); return ('0' + d.getHours()).slice(-2) + ':' + ('0' + d.getMinutes()).slice(-2); }
    function rbzAbiertas_() { var o = {}; Object.keys(rbForms || {}).forEach(function(i) { if (rbForms[i] && rbForms[i].abierto) o[rbzClaveX_(rbX(Number(i)))] = true; }); return o; }
    function rbzPintarAviso_() {
      var cont = document.getElementById('rbz-aviso'); if (!cont) return;
      var b = rbzBorradores_(), abiertas = rbzAbiertas_();
      var ks = Object.keys(b).filter(function(k) { return !abiertas[k]; }).sort(function(a, c) { return b[c].t - b[a].t; });
      if (!ks.length) { cont.innerHTML = ''; return; }
      var h = '<div class="rbz-aviso"><div class="rbz-t">📝 Retornos sin guardar <em>' + ks.length + '</em></div><div class="rbz-s">No se ha perdido nada: siguen aquí hasta que los guardes o los descartes.</div>';
      ks.forEach(function(k) {
        var d = b[k];
        h += '<div class="rbz-it" data-k="' + rbEsc(k) + '"><div class="rbz-ref">' + rbEsc(d.p || '(sin nº de pedido)') + '<small>fila ' + rbEsc(d.f) + (d.ag ? ' · ' + rbEsc(d.ag) : '') + '</small></div><div class="rbz-cam">' +
          Object.keys(d.c).map(function(c) { return '<b>' + rbEsc(d.cab[c] || rbLetra(Number(c))) + ':</b> ' + (String(d.c[c]).trim() ? rbEsc(d.c[c]) : '<i>vacío</i>'); }).join('<br>') +
          '</div><div class="rbz-h">Escrito ' + rbzHace_(d.t) + '</div><div class="rbz-bt"><button class="rbz-g" onclick="rbzGuardarYa_(this)">💾 Guardar ya</button><button class="rbz-r" onclick="rbzAbrir_(this)">✏️ Revisar</button><button class="rbz-d" onclick="rbzDescartar_(this)">Descartar</button></div><div class="rbz-msg"></div></div>';
      });
      cont.innerHTML = h + '</div>';
    }
    function rbzGuardarYa_(btn) {
      var k = rbzK_(btn), d = rbzBorradores_()[k], it = btn.closest('.rbz-it'), msg = it ? it.querySelector('.rbz-msg') : null;
      if (!d) { rbzPintarAviso_(); return; }
      var botones = it ? it.querySelectorAll('button') : [];
      [].forEach.call(botones, function(b) { b.disabled = true; });
      btn.innerText = 'Guardando…';
      function fallo(t) { [].forEach.call(botones, function(b) { b.disabled = false; }); btn.innerText = '💾 Guardar ya'; if (msg) msg.innerHTML = '<div class="rbz-no">' + t + '</div>'; }
      google.script.run.withSuccessHandler(function(o) {
        if (!o || o.error) { fallo(o && o.error === 'fila_cambiada' ? 'No encuentro ese pedido en Retornos (¿lo han borrado?). Lo escrito sigue aquí.' : rbEsc(o && o.error ? o.error : 'No se pudo abrir la fila.')); return; }
        var pisaria = Object.keys(d.c).filter(function(c) {
          var ahora = String(o.valores[c - 1]), antes = d.antes && d.antes[c] !== undefined ? String(d.antes[c]) : ahora;
          return ahora !== antes && ahora !== String(d.c[c]);
        });
        if (pisaria.length) { fallo('Alguien ha cambiado ' + pisaria.map(function(c) { return rbEsc(d.cab[c] || rbLetra(Number(c))); }).join(', ') + ' desde que lo escribiste. Pulsa «✏️ Revisar» para verlo antes de guardar.'); return; }
        google.script.run.withSuccessHandler(function(r) {
          if (!r || r.error) { fallo(rbEsc(r && r.error === 'fila_cambiada' ? 'La fila ha cambiado; pulsa «✏️ Revisar».' : (r && r.error ? r.error : 'No se pudo guardar.'))); return; }
          rbzQuitarBorrador_(k);
          try { rbzAnotarGuardado_({ datos: ['', d.ag || '', d.p], fila: r.fila || d.f }, Object.keys(d.c).map(function(c) { return d.cab[c] || rbLetra(Number(c)); }).join(', ')); } catch (eG) {}
          var av = document.getElementById('rbz-aviso');
          if (av) { av.insertAdjacentHTML('afterbegin', '<div class="rbz-hecho">✅ Guardado en la fila ' + rbEsc(r.fila || d.f) + ' (' + rbEsc(d.p) + ') · ' + rbzHora_(Date.now()) + '</div>'); setTimeout(function() { var x = av.querySelector('.rbz-hecho'); if (x) x.remove(); }, 8000); }
        }).withFailureHandler(function(err) { fallo('No se pudo guardar: ' + rbEsc(err && err.message ? err.message : err)); }).guardarFilaRetornos(o.fila, d.p, d.c, {});
      }).withFailureHandler(function(err) { fallo('No se pudo abrir la fila: ' + rbEsc(err && err.message ? err.message : err)); }).obtenerFilaRetornos(d.f, d.p);
    }
    function rbzK_(btn) { var it = btn; while (it && !it.getAttribute('data-k')) it = it.parentElement; return it ? it.getAttribute('data-k') : ''; }
    function rbzDescartar_(btn) {
      if (!btn.classList.contains('seguro')) { btn.classList.add('seguro'); btn.innerText = '¿Seguro? Descartar'; setTimeout(function() { if (btn.isConnected) { btn.classList.remove('seguro'); btn.innerText = 'Descartar'; } }, 4000); return; }
      rbzQuitarBorrador_(rbzK_(btn));
    }
    function rbzAbrir_(btn) {
      var k = rbzK_(btn), d = rbzBorradores_()[k]; if (!d) { rbzPintarAviso_(); return; }
      rbzAbrirTras = { k: k, f: d.f };
      var inp = document.getElementById('rb-texto'); inp.value = d.p || '';
      rbBuscar();
    }
    function rbzTrasPintar_() {
      rbzPintarAviso_();
      var q = rbzAbrirTras; rbzAbrirTras = null;
      if (!q || !rbEstado.res || !rbEstado.res.resultados) return;
      var lista = rbEstado.res.resultados, idx = -1;
      lista.forEach(function(x, i) { if (rbzClaveX_(x) === q.k && (idx === -1 || x.fila === q.f)) idx = i; });
      if (idx === -1) { var av = document.querySelector('.rbz-it[data-k="' + q.k + '"]'); if (av) av.insertAdjacentHTML('beforeend', '<div class="rbz-no">No encuentro ese pedido en Retornos (¿lo han borrado?). Lo escrito sigue aquí.</div>'); return; }
      if (!document.getElementById('rb-item-' + idx)) { rbEstado.filtro = 'todos'; rbPintar(); }
      rbRellenar(idx);
      try { document.getElementById('rb-item-' + idx).scrollIntoView({ block: 'start', behavior: 'smooth' }); } catch (e) {}
    }
    function rbzAnotarGuardado_(x, que) {
      var l = rbzLeer_(RBZ_GUARD) || [], hoy = new Date().toDateString();
      l = l.filter(function(g) { return g && new Date(g.t).toDateString() === hoy; });
      l.unshift({ t: Date.now(), p: String(x.datos[2] || ''), f: x.fila, q: que });
      rbzEscribir_(RBZ_GUARD, l.slice(0, 8));
      rbzPintarUltimos_();
    }
    function rbzPintarUltimos_() {
      var cont = document.getElementById('rbz-ult'); if (!cont) return;
      var hoy = new Date().toDateString(), l = (rbzLeer_(RBZ_GUARD) || []).filter(function(g) { return g && new Date(g.t).toDateString() === hoy; });
      if (!l.length) { cont.innerHTML = ''; return; }
      cont.innerHTML = '<div class="rbz-ult"><div class="rbz-ut">🕘 Lo último que has guardado <span>hoy</span></div>' + l.slice(0, 5).map(function(g) {
        return '<div class="rbz-ul"><i>' + rbzHora_(g.t) + '</i><b>' + rbEsc(g.p || '(sin nº)') + '<small>fila ' + rbEsc(g.f) + ' · ' + rbEsc(g.q) + '</small></b><button data-p="' + rbEsc(g.p) + '" onclick="rbzVer_(this)">Ver</button></div>';
      }).join('') + '</div>';
    }
    function rbzVer_(btn) { var p = btn.getAttribute('data-p'); if (!p) return; document.getElementById('rb-texto').value = p; rbBuscar(); }
    function rbzPill_(i, n) {
      var cont = document.getElementById('rb-form-' + i); if (!cont) return;
      var barra = cont.querySelector('.rb-barra'), pill = cont.querySelector('.rbz-pend');
      if (n === 0) { if (pill) pill.remove(); return; }
      var txt = '<span class="rbz-pt"></span><div>' + n + ' cambio' + (n === 1 ? '' : 's') + ' sin guardar<small>Guardado en borrador por si cierras sin querer</small></div>';
      if (pill) pill.innerHTML = txt; else if (barra) barra.insertAdjacentHTML('beforebegin', '<div class="rbz-pend">' + txt + '</div>');
    }
    function rbLetra(c) { return String.fromCharCode(64 + c); }
    function rbInit() {
      var inp = document.getElementById('rb-texto');
      if (inp && inp.getAttribute('data-listo')) { try { rblCargar_(false); } catch (eRbl0) {} }
      if (!inp || inp.getAttribute('data-listo')) return;
      inp.setAttribute('data-listo', '1');
      inp.addEventListener('keydown', function(ev) { if (ev.key === 'Enter' || ev.keyCode === 13) { ev.preventDefault(); rbBuscar(); } });
      try {
        var cabRb = document.querySelector('#rb-box .rb-cab'), btnRb = document.querySelector('#rb-box .rb-btn');
        if (cabRb && !document.getElementById('rbz-aviso')) cabRb.insertAdjacentHTML('afterend', '<div id="rbz-aviso"></div>');
        if (btnRb && !document.getElementById('rbz-ult')) btnRb.insertAdjacentHTML('afterend', '<div id="rbz-ult"></div>');
        rbzPintarAviso_(); rbzPintarUltimos_();
        var ir = rbzLeer_('rbzIr');
        if (ir && ir.k && Date.now() - ir.t < 60000) {
          try { localStorage.removeItem('rbzIr'); } catch (eIr) {}
          var dIr = rbzBorradores_()[ir.k];
          if (dIr) { rbzAbrirTras = { k: ir.k, f: dIr.f }; inp.value = dIr.p || ''; setTimeout(rbBuscar, 50); }
        }
      } catch (eRbz) {}
      if (document.getElementById('rb-box').style.display !== 'none') { try { inp.focus(); } catch (e) {} }
      try { rblCargar_(false); } catch (eRbl) {}
      try { if (typeof accionActual !== 'undefined' && accionActual === 'retornos') mpRecoger_('retornos'); } catch (eMp) {}
    }
    function rbAbrir() {
      document.getElementById('rb-card').style.display = 'none';
      document.getElementById('rb-box').style.display = 'block';
      rbInit();
      try { document.getElementById('rb-texto').focus(); } catch (e) {}
    }
    function rbCerrar() {
      document.getElementById('rb-box').style.display = 'none';
      document.getElementById('rb-card').style.display = 'flex';
    }
    function rbModo(btn) {
      rbEstado.modo = btn.getAttribute('data-modo');
      [].forEach.call(document.querySelectorAll('.rb-modo'), function(b) { b.classList.toggle('on', b === btn); });
      if (document.getElementById('rb-texto').value.trim()) rbBuscar();
    }
    function rbBuscar() {
      var texto = document.getElementById('rb-texto').value.trim();
      var res = document.getElementById('rb-res');
      if (texto.length < 3) { res.innerHTML = '<div class="rb-consejo">Escribe al menos 3 letras o cifras.</div>'; return; }
      rbEstado.texto = texto; rbForms = {};
      var conHist = !!(document.getElementById('rb-hist') && document.getElementById('rb-hist').checked), loc = null, ficha = {};
      rbEstado.ficha = ficha;
      if (!conHist && RBL_ && rblDatos) { try { loc = RBL_.buscar(rblDatos, texto, rbEstado.modo); } catch (eL) { loc = null; } }
      if (loc && !loc.error) { loc.local = true; rbEstado.res = loc; rbEstado.filtro = 'todos'; rbPintar(); }
      else { loc = null; res.innerHTML = '<div class="rb-buscando"><svg viewBox="0 0 130 112"><use href="#bebe-rufo"/></svg><p>Buscando «' + rbEsc(texto) + '»…</p></div>'; }
      try { rblCargar_(false); } catch (eC) {}
      google.script.run.withSuccessHandler(function(r) {
        if (!document.getElementById('rb-texto') || document.getElementById('rb-texto').value.trim() !== texto || rbEstado.ficha !== ficha) return;
        if (r && r.error) { if (!loc) res.innerHTML = '<div class="rb-no">' + rbEsc(r.error) + '</div>'; else rblEstado_(''); return; }
        if (loc && rbEstado.res === loc) {
          if (rblFirma_(loc) === rblFirma_(r)) { rblCompletar_(loc, r); try { rbzTrasPintar_(); } catch (eRbz0) {} return; }
          rblViejo = true; try { rblCargar_(true); } catch (eC2) {}
          var hayForm = Object.keys(rbForms).some(function(i) { return rbForms[i]; });
          if (hayForm) { rblAvisoNuevo_(r); return; }
        }
        rbEstado.res = r; rbEstado.filtro = 'todos'; rbPintar();
        try { rbzTrasPintar_(); } catch (eRbz) {}
      }).withFailureHandler(function(err) {
        if (rbEstado.ficha !== ficha) return;
        if (loc && rbEstado.res === loc) { rblEstado_('<span class="rbl-comp">⚠️ No se ha podido comprobar en la hoja: ' + rbEsc(err && err.message ? err.message : err) + '</span>'); return; }
        res.innerHTML = '<div class="rb-no">⚠️ No se pudo buscar: ' + rbEsc(err && err.message ? err.message : err) + '</div>';
      }).backendBuscarRetornos(texto, rbEstado.modo, conHist);
    }
    function rblHueco_(r) { return '<div id="rbl-estado">' + (r && r.local ? '<span class="rbl-comp"><i></i>Al momento · comprobando en la hoja…</span>' : '') + '</div>'; }
    function rbPintar() {
      var r = rbEstado.res, res = document.getElementById('rb-res');
      if (!r || !r.resultados || r.resultados.length === 0) {
        res.innerHTML = rblHueco_(r) + '<div class="rb-no">❌ No hay ningún artículo ni cliente igual o parecido a «' + rbEsc(rbEstado.texto) + '» en Retornos.</div>' + (r && typeof crHistorialHtml === 'function' ? crHistorialHtml(r.historial) : '');
        return;
      }
      var lista = r.resultados;
      var h = rblHueco_(r) + '<div id="rb-cabres">' + rbCabHtml_(lista, r.total) + '</div>';
      res.innerHTML = rbPintarLista_(lista, h, r);
    }
    function rbCabHtml_(lista, total) {
      var proc = 0, exactas = 0;
      lista.forEach(function(x) { if (x.procesado) proc++; if (x.exacta) exactas++; });
      var pend = lista.length - proc;
      var h = '<div class="rb-resumen"><b>¡He encontrado ' + (total > lista.length ? total + ' (enseño las ' + lista.length + ' más parecidas)' : lista.length) + '!</b> ' +
        exactas + ' exacta' + (exactas === 1 ? '' : 's') + ' y ' + (lista.length - exactas) + ' parecida' + ((lista.length - exactas) === 1 ? '' : 's') + '</div>' +
        '<div class="rb-filtros">' +
          '<button class="rb-filtro' + (rbEstado.filtro === 'todos' ? ' on' : '') + '" onclick="rbFiltrar(this, &quot;todos&quot;)">Todos · ' + lista.length + '</button>' +
          '<button class="rb-filtro v' + (rbEstado.filtro === 'proc' ? ' on' : '') + '" onclick="rbFiltrar(this, &quot;proc&quot;)">✅ Procesados · ' + proc + '</button>' +
          '<button class="rb-filtro a' + (rbEstado.filtro === 'pend' ? ' on' : '') + '" onclick="rbFiltrar(this, &quot;pend&quot;)">⏳ Sin procesar · ' + pend + '</button>' +
        '</div>';
      return h;
    }
    function rbPintarLista_(lista, h, r) {
      var ocultos = 0;
      lista.forEach(function(x, i) {
        if ((rbEstado.filtro === 'proc' && !x.procesado) || (rbEstado.filtro === 'pend' && x.procesado)) { ocultos++; return; }
        h += '<div class="rb-item ' + (x.procesado ? 'proc' : 'pend') + '" id="rb-item-' + i + '">' + rbItemHtml(x, i) + '</div>';
      });
      if (ocultos > 0) h += '<div class="rb-consejo" style="text-align:center">Ocultos por el filtro: ' + ocultos + ' (pulsa «Todos» para verlos)</div>';
      if (typeof crHistorialHtml === 'function') h += crHistorialHtml(r.historial);
      return h;
    }
    function rbFiltrar(btn, f) { rbEstado.filtro = f; rbForms = {}; rbPintar(); }
    function rbItemHtml(x, i) {
      var cab = (rbEstado.res && rbEstado.res.cabeceras) || [];
      var d = x.datos || [];
      var fondo = x.procesado ? (x.color || '#b6d7a8') : '#ffffff', sinDato = x.parcial ? '…' : '—';
      var h = '<span class="rb-estado ' + (x.procesado ? 'proc' : 'pend') + '">' + (x.procesado ? '✅ Retorno procesado' : '⏳ Sin procesar') + '</span>' +
        (rbzBorradores_()[rbzClaveX_(x)] ? '<span class="rbz-badge">📝 Borrador</span>' : '') +
        '<div class="rb-top"><span>Retornos · fila ' + x.fila + '</span><span class="rb-badges">' +
          (x.exacta ? '<span class="rb-b-exacta">🎯 Exacta</span>' : '<span class="rb-b-similar">🟡 Parecida</span>') + '<span class="rb-b-pct">' + x.pct + '%</span></span></div>' +
        '<div class="rb-ref"><small>' + rbEsc(cab[2] || 'Nº pedido') + '</small>' + (rbEsc(d[2]) || '—') + (d[1] ? ' <span class="rb-chip">' + rbEsc(d[1]) + '</span>' : '') + '</div>' +
        '<div class="rb-match">' + x.tipo + ': <mark>' + rbEsc(x.valor) + '</mark></div><div class="rb-porque">' + rbEsc(x.porque) + '</div>' +
        '<div class="rb-vista" style="background:' + fondo + '"><div>' + (rbEsc(d[0]) || '—') + '</div><div>' + (rbEsc(d[1]) || '—') + '</div><div>' + (rbEsc(d[2]) || '—') + '</div></div>' +
        '<div class="rb-vista-cap">Así se ve la fila en la hoja' + (x.procesado ? ' (en verde: procesado)' : ' (sin verde: sin procesar)') + '</div>' +
        '<div class="rb-datos" id="rb-datos-' + i + '" style="display:none"><dl>';
      for (var c = 0; c < 14; c++) h += '<dt>' + rbEsc(cab[c] || rbLetra(c + 1)) + '</dt><dd' + (c + 1 === x.col ? ' class="resalte"' : '') + '>' + (rbEsc(d[c]) || sinDato) + '</dd>';
      h += '</dl></div><button class="rb-ver" onclick="rbVer(' + i + ')">▾ Ver todos los datos de la fila</button>' +
        '<div class="rb-acciones"><button class="rb-acc azul" title="Va a la fila y copia el nº de pedido" onclick="rbIr(' + i + ', this)">➡️ Ir a la fila</button>' +
        '<button class="rb-acc" id="rb-btn-rell-' + i + '" onclick="rbRellenar(' + i + ')">🛠️ Rellenar</button>' +
        '<button class="rb-acc" onclick="rbBuscarRef(' + i + ')">🔍 Buscar pedido</button>' +
        '<button class="rb-acc" data-cr="fila" data-i="' + i + '">🕵️ Cambios</button></div>' +
        '<div id="rb-camb-' + i + '"></div>' +
        '<div id="rb-form-' + i + '"></div>';
      return h;
    }
    function rbVer(i) {
      var el = document.getElementById('rb-datos-' + i);
      el.style.display = el.style.display === 'none' ? 'block' : 'none';
    }
    function rbX(i) { return rbEstado.res.resultados[i]; }
    function rbIr(i, btn) {
      var x = rbX(i), ped = String(x.datos[2] || '').trim();
      if (ped) copiarAlPortapapeles(ped, function(ok) { if (ok && btn) { btn.innerText = '📋 Ref. copiada ✔️'; btn.title = 'Copiado: ' + ped; } });
      google.script.run.activarCeldaEnHoja('Retornos', x.fila, 3);
    }
    function rbBuscarRef(i) {
      var ped = rbX(i).datos[2] || '';
      if (document.getElementById('b-ref')) {
        var nb = document.getElementById('b-nombre');
        if (nb && nb.checked) { nb.checked = false; cambiarModoNombre_(); }
        document.getElementById('b-ref').value = ped;
        try { window.scrollTo(0, 0); } catch (e) {}
        iniciarBusqueda();
      } else {
        try { localStorage.setItem('lastRef', ped); } catch (e) {}
        irSeccion_('buscar', '🔍 Buscar Referencia');
      }
    }
    function rbRellenar(i) {
      var cont = document.getElementById('rb-form-' + i), btn = document.getElementById('rb-btn-rell-' + i);
      if (rbForms[i] && rbForms[i].abierto) {
        var nCam = Object.keys(rbForms[i].cambios || {}).length;
        if (nCam > 0 && !cont.querySelector('.rbz-cerrar')) { rbzPreguntarCerrar_(i); return; }
        rbForms[i].abierto = false; cont.innerHTML = ''; btn.innerText = '🛠️ Rellenar'; btn.classList.remove('oscuro'); rbzPintarAviso_(); return;
      }
      btn.innerText = '✖️ Cerrar'; btn.classList.add('oscuro');
      cont.innerHTML = '<div class="rb-cargando"><span class="spinner"></span>Abriendo la fila…</div>';
      var x = rbX(i);
      google.script.run.withSuccessHandler(function(d) {
        if (d && d.error === 'fila_cambiada') { cont.innerHTML = rbAvisoFila(); return; }
        if (d && d.error) { cont.innerHTML = '<div class="rb-no">' + rbEsc(d.error) + '</div>'; return; }
        if (d.fila && d.fila !== x.fila) x.fila = d.fila;
        var agencia = (d.valores[1] || '').trim() || (d.agencias[0] || '');
        rbForms[i] = { abierto: true, d: d, cambios: {}, agencia: agencia, otra: false };
        var bz = rbzBorradores_()[rbzClaveX_(x)], recuperados = 0;
        if (bz && bz.c) Object.keys(bz.c).forEach(function(c) { if (String(bz.c[c]) !== String(d.valores[c - 1])) { rbForms[i].cambios[c] = bz.c[c]; recuperados++; } });
        rbPintarForm(i);
        if (recuperados) { var rr = document.querySelector('#rb-form-' + i + ' .rb-rell'); if (rr) rr.insertAdjacentHTML('afterbegin', '<div class="rbz-rec">📝 He recuperado lo que escribiste ' + rbzHace_(bz.t) + ' (' + recuperados + ' cambio' + (recuperados === 1 ? '' : 's') + '). <u>Aún no está en la hoja</u>: revísalo y pulsa 💾 Guardar.</div>'); var gb = document.querySelector('#rb-form-' + i + ' .rb-guardar'); if (gb) gb.classList.add('rbz-latir'); }
        rbzPintarAviso_();
      }).withFailureHandler(function(err) {
        cont.innerHTML = '<div class="rb-no">⚠️ No se pudo abrir: ' + rbEsc(err && err.message ? err.message : err) + '</div>';
      }).obtenerFilaRetornos(x.fila, x.datos[2] || '');
    }
    function rbzPreguntarCerrar_(i) {
      var f = rbForms[i], cont = document.getElementById('rb-form-' + i), n = Object.keys(f.cambios).length;
      cont.innerHTML = '<div class="rb-rell"><div class="rbz-cerrar"><div class="rbz-t">⚠️ Tienes ' + n + ' cambio' + (n === 1 ? '' : 's') + ' sin guardar en este retorno</div><ul>' +
        Object.keys(f.cambios).map(function(c) { return '<li><b>' + rbEsc(f.d.cabeceras[c - 1] || rbLetra(Number(c))) + ':</b> ' + (String(f.cambios[c]).trim() ? rbEsc(f.cambios[c]) : '<i>vacío</i>') + '</li>'; }).join('') +
        '</ul><button class="g rb-guardar" onclick="rbGuardar(' + i + ', false)">💾 Guardar ahora</button><button class="b" onclick="rbzDejar_(' + i + ')">📝 Dejarlo en borrador para luego</button><button class="x" onclick="rbzDescartarForm_(' + i + ')">Descartar los cambios</button></div></div>';
    }
    function rbzCerrarForm_(i) {
      var cont = document.getElementById('rb-form-' + i), btn = document.getElementById('rb-btn-rell-' + i);
      if (rbForms[i]) rbForms[i].abierto = false;
      if (cont) cont.innerHTML = ''; if (btn) { btn.innerText = '🛠️ Rellenar'; btn.classList.remove('oscuro'); }
    }
    function rbzDejar_(i) { rbzGuardarBorrador_(i); rbzCerrarForm_(i); var el = document.getElementById('rb-item-' + i); if (el && rbX(i)) el.innerHTML = rbItemHtml(rbX(i), i); rbzPintarAviso_(); }
    function rbzDescartarForm_(i) { var x = rbX(i); rbzCerrarForm_(i); rbzQuitarBorrador_(rbzClaveX_(x)); var el = document.getElementById('rb-item-' + i); if (el && x) el.innerHTML = rbItemHtml(x, i); }
    function rbAvisoFila() {
      return '<div class="rb-aviso-fila">⚠️ Ese pedido ya no está en su fila y no se ha podido localizar (¿lo han borrado?). No se ha escrito nada: vuelve a buscarlo.</div>';
    }
    function rbPintarForm(i) {
      var f = rbForms[i], d = f.d, cont = document.getElementById('rb-form-' + i);
      var h = '<div class="rb-rell">';
      if (!d.procesado) {
        h += '<div class="rb-proc"><button class="rb-proc-btn" onclick="rbGuardar(' + i + ', true)">✅ Marcar como procesado</button>' +
          '<div class="rb-proc-sub">Pone la fecha de hoy en ' + rbEsc(d.cabeceras[0] || 'A') + ' (si está vacía), la agencia en ' + rbEsc(d.cabeceras[1] || 'B') + ' y pinta toda la fila de verde.</div><div class="rb-ops">';
        d.agencias.forEach(function(a) { h += '<button class="rb-op' + (!f.otra && a === f.agencia ? ' on' : '') + '" onclick="rbAgencia(' + i + ', this)" data-v="' + rbEsc(a) + '">' + rbEsc(a) + '</button>'; });
        h += '<button class="rb-op' + (f.otra ? ' on' : '') + '" onclick="rbAgenciaOtra(' + i + ')">Otra…</button></div>' +
          (f.otra ? '<input class="rb-inp" id="rb-otra-' + i + '" placeholder="Escribe la agencia" value="' + rbEsc(f.agencia) + '" oninput="rbForms[' + i + '].agencia = this.value">' : '') + '</div>';
      } else {
        h += '<div class="rb-proc hecho"><b>✅ Esta fila ya está procesada (en verde).</b><button class="rb-quitar" onclick="rbGuardar(' + i + ', false, true)">Quitar el verde</button></div>';
      }
      var vacios = RB_CAMPOS.filter(function(c) { return String(d.valores[c - 1]).trim() === ''; }).length;
      h += '<div class="rb-rell-cab"><b>✏️ Datos del retorno</b><span' + (vacios ? '>' + vacios + ' de ' + RB_CAMPOS.length + ' vacíos (en amarillo)' : ' class="lleno">✔ Todo relleno') + '</span></div>' +
        '<div class="rb-rell-ayuda">Cambia lo que haga falta y pulsa <b>Guardar</b> (o Intro). Solo se escribe lo que cambies.</div>';
      RB_GRUPOS.forEach(function(g) {
        h += '<div class="rb-grupo"><div class="rb-grupo-cab">' + g[0] + '</div>';
        g[1].forEach(function(c) { h += rbCampoHtml(i, c); });
        h += '</div>';
      });
      h += '<div id="rb-resumen-' + i + '"></div>' +
        '<div class="rb-barra"><button class="rb-descartar" onclick="rbForms[' + i + '].cambios = {}; rbPintarForm(' + i + ')">Descartar</button>' +
        '<button class="rb-guardar" onclick="rbGuardar(' + i + ', false)">💾 Guardar</button></div></div>';
      cont.innerHTML = h;
      rbActualizarBarra(i);
    }
    function rbCampoHtml(i, c) {
      var f = rbForms[i], d = f.d, cambiado = f.cambios[c] !== undefined, actual = cambiado ? f.cambios[c] : d.valores[c - 1];
      var vacio = String(d.valores[c - 1]).trim() === '';
      var h = '<div class="rb-cf' + (vacio ? ' vacio' : '') + (cambiado ? ' cambiado' : '') + '" id="rb-cf-' + i + '-' + c + '"><label><span>' + rbEsc(d.cabeceras[c - 1] || rbLetra(c)) + '</span><i>' + (vacio ? 'vacío · ' : '') + 'col. ' + rbLetra(c) + '</i></label>';
      if (d.opciones[c]) {
        h += '<div class="rb-ops">';
        d.opciones[c].forEach(function(o) { h += '<button class="rb-op' + (o === actual ? ' on' + (cambiado ? ' cambiado' : '') : '') + '" data-v="' + rbEsc(o) + '" onclick="rbElegir(' + i + ',' + c + ', this)">' + rbEsc(o) + '</button>'; });
        h += '</div>';
      } else {
        var ph = c === 1 ? 'dd/mm/aaaa' : (c === 3 ? 'Nº de pedido' : (c === 8 ? 'Nombre del cliente' : 'Escribe lo que necesites…'));
        h += '<div class="rb-fila-inp"><input class="rb-inp' + (cambiado ? ' cambiado' : '') + '" id="rb-in-' + i + '-' + c + '" value="' + rbEsc(actual) + '" placeholder="' + ph + '" oninput="rbCambio(' + i + ',' + c + ', this)" onkeydown="if (event.key === &quot;Enter&quot;) { event.preventDefault(); rbGuardar(' + i + ', false); }">' +
          (c === 1 ? '<button class="rb-hoy" onclick="rbHoy(' + i + ')">Hoy</button>' : '') + '</div>';
      }
      return h + '</div>';
    }
    function rbCambio(i, c, el) {
      var f = rbForms[i];
      if (el.value === String(f.d.valores[c - 1])) delete f.cambios[c]; else f.cambios[c] = el.value;
      el.classList.toggle('cambiado', f.cambios[c] !== undefined);
      var cf = document.getElementById('rb-cf-' + i + '-' + c);
      if (cf) cf.classList.toggle('cambiado', f.cambios[c] !== undefined);
      rbActualizarBarra(i);
    }
    function rbActualizarBarra(i) {
      var f = rbForms[i], n = Object.keys(f.cambios).length, cont = document.getElementById('rb-form-' + i);
      try { rbzGuardarBorrador_(i); rbzPill_(i, n); } catch (eRbz) {}
      var btn = cont.querySelector('.rb-guardar');
      if (!btn) return;
      btn.disabled = n === 0;
      btn.innerText = '💾 ' + (n > 0 ? 'Guardar ' + n + ' cambio' + (n === 1 ? '' : 's') + ' (Intro)' : 'Guardar (no hay cambios)');
      var desc = cont.querySelector('.rb-descartar');
      if (desc) desc.style.display = n > 0 ? '' : 'none';
      var res = document.getElementById('rb-resumen-' + i);
      if (!res) return;
      res.innerHTML = n === 0 ? '' : '<div class="rb-resumen-cambios"><b>Vas a guardar:</b>' + Object.keys(f.cambios).map(function(k) {
        var c = Number(k), antes = String(f.d.valores[c - 1]), nuevo = String(f.cambios[k]);
        return '<div class="rb-rc"><span>' + rbEsc(f.d.cabeceras[c - 1] || rbLetra(c)) + ': ' + (antes.trim() ? '<s>' + rbEsc(antes) + '</s>' : '<i>vacío</i>') + ' → <b>' + (nuevo.trim() ? rbEsc(nuevo) : '<i>vacío</i>') + '</b></span>' +
          '<button class="rb-rc-x" title="Dejarlo como estaba" onclick="rbQuitarCambio(' + i + ',' + c + ')">✕</button></div>';
      }).join('') + '</div>';
    }
    function rbQuitarCambio(i, c) { delete rbForms[i].cambios[c]; rbPintarForm(i); }
    function rbElegir(i, c, el) {
      var f = rbForms[i], v = el.getAttribute('data-v');
      if (v === String(f.d.valores[c - 1]) || f.cambios[c] === v) delete f.cambios[c]; else f.cambios[c] = v;
      rbPintarForm(i);
    }
    function rbHoy(i) {
      var h = new Date(), t = ('0' + h.getDate()).slice(-2) + '/' + ('0' + (h.getMonth() + 1)).slice(-2) + '/' + h.getFullYear();
      var el = document.getElementById('rb-in-' + i + '-1');
      el.value = t; rbCambio(i, 1, el);
    }
    function rbAgencia(i, el) { rbForms[i].agencia = el.getAttribute('data-v'); rbForms[i].otra = false; rbPintarForm(i); }
    function rbAgenciaOtra(i) { rbForms[i].otra = true; rbForms[i].agencia = ''; rbPintarForm(i); try { document.getElementById('rb-otra-' + i).focus(); } catch (e) {} }
    function rbGuardar(i, procesar, quitarVerde) {
      var f = rbForms[i], x = rbX(i), cont = document.getElementById('rb-form-' + i);
      if (!f) return;
      if (!procesar && !quitarVerde && Object.keys(f.cambios).length === 0) return;
      if (procesar && !String(f.agencia || '').trim()) { alert('Elige o escribe la agencia.'); return; }
      [].forEach.call(cont.querySelectorAll('button'), function(b) { b.disabled = true; });
      var g = cont.querySelector(procesar ? '.rb-proc-btn' : (quitarVerde ? '.rb-quitar' : '.rb-guardar'));
      if (g) g.innerText = 'Guardando…';
      google.script.run.withSuccessHandler(function(r) {
        if (r && r.error === 'fila_cambiada') { cont.innerHTML = rbAvisoFila(); return; }
        if (r && r.error) { alert('No se pudo guardar: ' + r.error); rbPintarForm(i); return; }
        if (r.fila && r.fila !== x.fila) x.fila = r.fila;
        var kAntes = rbzClaveX_(x), queG = procesar ? '✅ procesado' + (f.agencia ? ' (' + f.agencia + ')' : '') : (quitarVerde ? 'quitado el verde' : Object.keys(f.cambios).map(function(c) { return f.d.cabeceras[c - 1] || rbLetra(Number(c)); }).join(', '));
        x.datos = r.valores; x.procesado = r.procesado; x.color = r.color; x.parcial = false; rblViejo = true;
        rbzQuitarBorrador_(kAntes); rbzQuitarBorrador_(rbzClaveX_(x));
        try { rbzAnotarGuardado_(x, queG); } catch (eRbz) {}
        var msg = (procesar ? '✅ Procesado: fila ' + x.fila + ' en verde' : (quitarVerde ? '↩️ Quitado el verde de la fila ' + x.fila : '✅ Guardado en la fila ' + x.fila)) + ' · ' + rbzHora_(Date.now());
        rbRefrescarItem(i, '<div class="rb-hecho"><span>' + msg + '</span><button class="rb-deshacer" onclick="rbDeshacer(' + i + ')">Deshacer</button></div>');
      }).withFailureHandler(function(err) {
        alert('No se pudo guardar: ' + (err && err.message ? err.message : err)); rbPintarForm(i);
      }).guardarFilaRetornos(x.fila, x.datos[2] || '', f.cambios, { procesar: !!procesar, quitarVerde: !!quitarVerde, agencia: f.agencia });
    }
    function rbRefrescarItem(i, extra) {
      var x = rbX(i), el = document.getElementById('rb-item-' + i);
      rbForms[i] = null;
      el.className = 'rb-item ' + (x.procesado ? 'proc' : 'pend');
      el.innerHTML = rbItemHtml(x, i);
      if (extra) document.getElementById('rb-form-' + i).innerHTML = extra;
    }
    function modoNombreActivo_() { var el = document.getElementById('b-nombre'); return !!(el && el.checked); }
    function cambiarModoNombre_() {
      var on = modoNombreActivo_(), inp = document.getElementById('b-ref'), et = document.getElementById('bn-toggle');
      if (et) et.classList.toggle('on', on);
      if (!inp) return;
      inp.placeholder = on ? 'Nombre del cliente (ej: maria garcia)' : 'Ej: 011324532';
      inp.value = on ? '' : (localStorage.getItem('lastRef') || '');
      document.getElementById('b-res').innerHTML = '<div class="empty-inline">' + (on ? '👤 Escribe el nombre (o parte) y pulsa Intro. Da igual tildes, mayúsculas y el orden.' : 'Tus resultados aparecerán aquí...') + '</div>';
      try { inp.focus(); inp.select(); } catch (e) {}
    }
    function iniciarBusquedaNombre_() {
      var texto = document.getElementById('b-ref').value.trim(), age = document.getElementById('b-age').value, cont = document.getElementById('b-res');
      if (texto.length < 3) { cont.innerHTML = '<div class="empty-inline">Escribe al menos 3 letras del nombre...</div>'; return; }
      cont.innerHTML = '<div class="loading-row"><span class="spinner"></span>Buscando el nombre «' + escAttr(texto) + '» en la columna I...</div>';
      if (busquedaEnCurso) return;
      busquedaEnCurso = true;
      google.script.run
        .withSuccessHandler(function(r) { terminarBusqueda_(texto, age, function() { pintarBusquedaNombre_(r, texto); }); })
        .withFailureHandler(function(err) { terminarBusqueda_(texto, age, function() { mostrarErrorBusqueda(err); }); })
        .backendBuscarNombre(texto, age);
    }
    function pintarBusquedaNombre_(r, texto) {
      var cont = document.getElementById('b-res'), t = escAttr(texto);
      if (!r || r.error) { cont.innerHTML = '<div style="color:var(--red-dark); text-align:center;">⚠️ ' + escAttr(r && r.error ? r.error : 'No se pudo buscar.') + '</div>'; return; }
      var lista = r.resultados || [];
      if (lista.length === 0) { cont.innerHTML = '<div style="color:var(--red-dark); text-align:center;">❌ Ningún nombre igual ni parecido a «' + t + '» en la columna I.</div>'; try { if (typeof mpAvisarSinResultados_ === 'function') mpAvisarSinResultados_(); } catch (eMp) {} return; }
      mostrarResultados(lista);
      var cab = r.exactas
        ? '👤 <b>' + r.total + ' fila' + (r.total === 1 ? '' : 's') + '</b> con «' + t + '» en el nombre' + (r.total > lista.length ? ' (enseño las ' + lista.length + ' más recientes)' : '')
        : '🟡 Ninguno igual. <b>' + (lista.length === 1 ? 'El más parecido' : 'Los ' + lista.length + ' más parecidos') + '</b> a «' + t + '»:';
      cont.insertAdjacentHTML('afterbegin', '<div class="bn-resumen">' + cab + '</div>');
    }
    function rbDeshacer(i) {
      var x = rbX(i), cont = document.getElementById('rb-form-' + i);
      cont.innerHTML = '<div class="rb-cargando"><span class="spinner"></span>Deshaciendo…</div>';
      google.script.run.withSuccessHandler(function(r) {
        if (r && r.error === 'fila_cambiada') { cont.innerHTML = rbAvisoFila(); return; }
        if (r && r.error) { cont.innerHTML = '<div class="rb-no">' + rbEsc(r.error) + '</div>'; return; }
        if (r.fila && r.fila !== x.fila) x.fila = r.fila;
        x.datos = r.valores; x.procesado = r.procesado; x.color = r.color; x.parcial = false; rblViejo = true;
        rbRefrescarItem(i, '<div class="rb-hecho"><span>↩️ Deshecho: la fila ' + x.fila + ' está como antes</span></div>');
      }).withFailureHandler(function(err) {
        cont.innerHTML = '<div class="rb-no">⚠️ No se pudo deshacer: ' + rbEsc(err && err.message ? err.message : err) + '</div>';
      }).deshacerFilaRetornos(x.fila);
    }
    function mpIr() {
      var cab = document.getElementById('mp-cabeza');
      if (!cab || cab.classList.contains('susto')) return;
      var destino = cab.getAttribute('data-destino') === 'retornos' ? 'retornos' : 'buscar';
      var campo = document.getElementById(destino === 'retornos' ? 'b-ref' : 'rb-texto');
      var texto = campo ? String(campo.value || '').trim().slice(0, 200) : '';
      mpQuitarAviso_();
      cab.classList.add('susto');
      try { if (texto) localStorage.setItem('mpIr', JSON.stringify({ d: destino, q: texto, t: Date.now() })); else localStorage.removeItem('mpIr'); } catch (e) {}
      setTimeout(function() { irSeccion_(destino, destino === 'retornos' ? '↩️ Retornos' : '🔍 Buscar Referencia'); setTimeout(function() { if (cab.isConnected) cab.classList.remove('susto'); }, 4000); }, 650);
    }
    function mpRecoger_(aqui) {
      var p = null;
      try { p = JSON.parse(localStorage.getItem('mpIr') || 'null'); if (p) localStorage.removeItem('mpIr'); } catch (e) { p = null; }
      if (!p || p.d !== aqui || !p.q || typeof p.t !== 'number') return false;
      var edad = Date.now() - p.t;
      if (!(edad >= -5000 && edad < 120000)) return false;
      var q = String(p.q);
      if (aqui === 'buscar') {
        var inp = document.getElementById('b-ref');
        if (!inp) return false;
        inp.value = q;
        if (q.length >= 3) iniciarBusqueda();
        return true;
      }
      var rb = document.getElementById('rb-texto');
      if (!rb) return false;
      rb.value = q;
      if (q.length < 3) return true;
      var bPed = document.querySelector('.rb-modo[data-modo="pedidos"]');
      if (/^[0-9]{6,}$/.test(q.replace(/[ .-]/g, '')) && bPed) rbModo(bPed); else rbBuscar();
      return true;
    }
    function mpAvisarSinResultados_() {
      var cab = document.getElementById('mp-cabeza'), z = cab ? cab.parentNode : null;
      if (!cab || !z || cab.getAttribute('data-destino') !== 'retornos' || cab.classList.contains('susto')) return;
      z.classList.add('aviso');
    }
    function mpQuitarAviso_() {
      var cab = document.getElementById('mp-cabeza');
      if (cab && cab.parentNode && cab.parentNode.classList) cab.parentNode.classList.remove('aviso');
    }
`.replace(/\n\s+/g, "\n");
}
function backendBuscarRetornos(_awr, _aws, _awt) {
    var _awu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_RETORNOS);
    if (!_awu)
        return { error: "No existe la pestaña Retornos." };
    var _awv = prepararBusquedaRetornos_(_awr);
    if (_awv.compacto.length < 3)
        return { error: "Escribe al menos 3 letras o cifras." };
    var _aww = [];
    if (_awt === true) {
        try {
            _aww = buscarEnHistorialRetornos_(_awv, _aws);
        }
        catch (_awx) {
            _aww = [];
        }
    }
    var _awy = _awu.getLastRow();
    if (_awy < 2)
        return { resultados: [], cabeceras: [], historial: _aww };
    var _awz = _aws !== "articulos" && _aws !== "pedidos", _axa = _aws !== "clientes" && _aws !== "pedidos";
    var _axb = _aws === "pedidos" || (_awz && _axa && _awv.parecePedido);
    var _axc = (_awz || _axa) ? _awu.getRange(2, COL_RET_CLIENTE_, _awy - 1, COL_RET_ARTICULO_ - COL_RET_CLIENTE_ + 1).getValues() : null;
    var _axd = _axb ? _awu.getRange(2, COLUMNA_REFERENCIAS, _awy - 1, 1).getValues() : null;
    var _axe = buscarEnDatosRetornos_(_awv, _aws, _awy - 1, function (_axf) { return _axc[_axf][0]; }, function (_axg) { return _axc[_axg][2]; }, function (_axh) { return _axd[_axh][0]; });
    var _axi = _axe.total, _axj = _axe.hallados;
    if (_axj.length === 0)
        return { resultados: [], total: 0, cabeceras: [], historial: _aww };
    var _axk = _awu.getRange(1, 1, 1, COL_RET_ULTIMA_).getDisplayValues()[0];
    var _axl = _axj.map(function (_axm) { return _axm.fila; });
    var _axn = leerFilasRetornos_(_awu, _axl);
    return {
        total: _axi,
        cabeceras: _axk,
        historial: _aww,
        resultados: _axj.map(function (_axo) {
            var _axp = _axn[_axo.fila];
            return {
                fila: _axo.fila, tipo: _axo.tipo, col: _axo.col, valor: String(_axo.valor),
                pct: _axo.r.pct, exacta: !!_axo.r.exacta, porque: porQueParecidoRetornos_(_axo.r, _axo.valor, _awr),
                datos: _axp.valores, procesado: esVerdeProcesado_(_axp.fondoA), color: _axp.fondoA
            };
        })
    };
}
function buscarEnDatosRetornos_(_axq, _axr, _axs, _axt, _axu, _axv) {
    var _axw = _axr !== "articulos" && _axr !== "pedidos", _axx = _axr !== "clientes" && _axr !== "pedidos";
    var _axy = _axr === "pedidos" || (_axw && _axx && _axq.parecePedido);
    var _axz = [], _aya = prepararFiltroRetornos_(_axq);
    for (var _ayb = 0; _ayb < _axs; _ayb++) {
        var _ayc = null;
        if (_axw) {
            var _ayd = _axt(_ayb);
            if (cotaParecidoRetornos_(_aya, _ayd) >= COTA_MINIMA_RETORNOS_) {
                var _aye = parecidoRetornos_(_axq, _ayd);
                if (_aye.pct >= UMBRAL_PARECIDO_RETORNOS_)
                    _ayc = { r: _aye, col: COL_RET_CLIENTE_, tipo: "Cliente", valor: _ayd };
            }
        }
        if (_axx) {
            var _ayf = _axu(_ayb);
            if (cotaParecidoRetornos_(_aya, _ayf) >= COTA_MINIMA_RETORNOS_) {
                var _ayg = parecidoRetornos_(_axq, _ayf);
                if (_ayg.pct >= UMBRAL_PARECIDO_RETORNOS_ && (!_ayc || _ayg.pct > _ayc.r.pct))
                    _ayc = { r: _ayg, col: COL_RET_ARTICULO_, tipo: "Artículo", valor: _ayf };
            }
        }
        if (_axy && (!_ayc || _ayc.r.pct < 100)) {
            var _ayh = _axv(_ayb), _ayi = pedidoCoincideRetornos_(_axq, _ayh);
            if (_ayi)
                _ayc = { r: _ayi, col: COLUMNA_REFERENCIAS, tipo: "Nº pedido", valor: textoCelda_(_ayh) };
        }
        if (_ayc) {
            _ayc.fila = _ayb + 2;
            _axz.push(_ayc);
        }
    }
    _axz.sort(function (_ayj, _ayk) { return (_ayk.r.pct - _ayj.r.pct) || (_ayk.fila - _ayj.fila); });
    return { total: _axz.length, hallados: _axz.slice(0, MAX_RESULTADOS_RETORNOS_) };
}
function buscarRetornosLocal_(_ayl, _aym, _ayn) {
    var _ayo = prepararBusquedaRetornos_(_aym);
    if (_ayo.compacto.length < 3)
        return { error: "Escribe al menos 3 letras o cifras." };
    if (!(_ayl.n > 0))
        return { resultados: [], cabeceras: [], historial: [] };
    var _ayp = buscarEnDatosRetornos_(_ayo, _ayn, _ayl.n, function (_ayq) { return _ayl.h[_ayq]; }, function (_ayr) { return _ayl.j[_ayr]; }, function (_ays) { return _ayl.p[_ays]; });
    if (_ayp.hallados.length === 0)
        return { resultados: [], total: 0, cabeceras: [], historial: [] };
    return {
        total: _ayp.total,
        cabeceras: _ayl.cab,
        historial: [],
        resultados: _ayp.hallados.map(function (_ayt) {
            var _ayu = _ayt.fila - 2, _ayv = [], _ayw = _ayl.df[_ayl.f.charCodeAt(_ayu) - 48];
            for (var _ayx = 0; _ayx < COL_RET_ULTIMA_; _ayx++)
                _ayv.push("");
            _ayv[0] = _ayl.da[_ayl.a.charCodeAt(_ayu) - 48];
            _ayv[1] = _ayl.db[_ayl.b.charCodeAt(_ayu) - 48];
            _ayv[COLUMNA_REFERENCIAS - 1] = _ayl.pv && _ayl.pv[_ayu] !== undefined ? _ayl.pv[_ayu] : textoCelda_(_ayl.p[_ayu]);
            _ayv[COL_RET_CLIENTE_ - 1] = _ayl.hv && _ayl.hv[_ayu] !== undefined ? _ayl.hv[_ayu] : textoCelda_(_ayl.h[_ayu]);
            _ayv[COL_RET_ARTICULO_ - 1] = _ayl.jv && _ayl.jv[_ayu] !== undefined ? _ayl.jv[_ayu] : textoCelda_(_ayl.j[_ayu]);
            return {
                fila: _ayt.fila, tipo: _ayt.tipo, col: _ayt.col, valor: String(_ayt.valor),
                pct: _ayt.r.pct, exacta: !!_ayt.r.exacta, porque: porQueParecidoRetornos_(_ayt.r, _ayt.valor, _aym),
                datos: _ayv, procesado: esVerdeProcesado_(_ayw), color: _ayw, parcial: true
            };
        })
    };
}
function obtenerDatosRetornosPanel() {
    if (!usuarioAutorizado_())
        return null;
    var _ayy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_RETORNOS);
    if (!_ayy)
        return null;
    var _ayz = Math.max(0, _ayy.getLastRow() - 1);
    var _aza = { v: VERSION_SISTEMA_, n: _ayz, cab: [], h: [], j: [], p: [], pv: {}, hv: {}, jv: {}, da: [], a: "", db: [], b: "", df: [], f: "" };
    if (!_ayz)
        return _aza;
    _aza.cab = _ayy.getRange(1, 1, 1, COL_RET_ULTIMA_).getDisplayValues()[0];
    var _azb = _ayy.getRange(2, COL_RET_CLIENTE_, _ayz, COL_RET_ARTICULO_ - COL_RET_CLIENTE_ + 1), _azc = _azb.getValues(), _azd = _azb.getDisplayValues();
    var _aze = _ayy.getRange(2, COLUMNA_REFERENCIAS, _ayz, 1).getValues(), _azf = _ayy.getRange(2, COLUMNA_REFERENCIAS, _ayz, 1).getDisplayValues();
    var _azg = _ayy.getRange(2, 1, _ayz, 2).getDisplayValues(), _azh = _ayy.getRange(2, 1, _ayz, 1).getBackgrounds();
    function _azi(_azj) { return (_azj === null || _azj === undefined || _azj instanceof Date) ? "" : _azj; }
    var _azk = { a: {}, b: {}, f: {} }, _azl = [], _azm = [], _azn = [];
    function _azo(_azp, _azq, _azr) {
        var _azs = _azk[_azp];
        if (!_azs.hasOwnProperty("k" + _azr)) {
            _azs["k" + _azr] = _azq.length;
            _azq.push(_azr);
        }
        var _azt = 48 + _azs["k" + _azr];
        if (_azt >= 0xD800)
            throw new Error("demasiados valores distintos");
        return String.fromCharCode(_azt);
    }
    for (var _azu = 0; _azu < _ayz; _azu++) {
        _aza.h.push(_azi(_azc[_azu][0]));
        _aza.j.push(_azi(_azc[_azu][2]));
        if (_azd[_azu][0] !== textoCelda_(_azc[_azu][0]))
            _aza.hv[_azu] = _azd[_azu][0];
        if (_azd[_azu][2] !== textoCelda_(_azc[_azu][2]))
            _aza.jv[_azu] = _azd[_azu][2];
        var _azv = textoCelda_(_aze[_azu][0]);
        _aza.p.push(_azv);
        if (_azf[_azu][0] !== _azv)
            _aza.pv[_azu] = _azf[_azu][0];
        _azl.push(_azo("a", _aza.da, _azg[_azu][0]));
        _azm.push(_azo("b", _aza.db, _azg[_azu][1]));
        _azn.push(_azo("f", _aza.df, _azh[_azu][0]));
    }
    _aza.a = _azl.join("");
    _aza.b = _azm.join("");
    _aza.f = _azn.join("");
    return _aza;
}
var codigoBusquedaRetornosHecho_ = null;
function codigoBusquedaRetornosCliente_() {
    if (codigoBusquedaRetornosHecho_ !== null)
        return codigoBusquedaRetornosHecho_;
    var _azw = "";
    try {
        var _azx = [normTextoRetornos_, compactoRetornos_, distanciaDentro_, levenshtein, similitudRetornos_, prepararBusquedaRetornos_,
            cuentaLetras_, comunesLetras_, prepararFiltroRetornos_, cotaParecidoRetornos_, parecidoRetornos_, porQueParecidoRetornos_,
            pedidoCoincideRetornos_, textoCelda_, esVerdeProcesado_, buscarEnDatosRetornos_, buscarRetornosLocal_];
        var _azy = { COL_RET_CLIENTE_: COL_RET_CLIENTE_, COL_RET_ARTICULO_: COL_RET_ARTICULO_, COL_RET_ULTIMA_: COL_RET_ULTIMA_, COLUMNA_REFERENCIAS: COLUMNA_REFERENCIAS,
            MAX_RESULTADOS_RETORNOS_: MAX_RESULTADOS_RETORNOS_, UMBRAL_PARECIDO_RETORNOS_: UMBRAL_PARECIDO_RETORNOS_, COTA_MINIMA_RETORNOS_: COTA_MINIMA_RETORNOS_ };
        var _azz = Object.keys(_azy).map(function (_baa) { return "var " + _baa + " = " + JSON.stringify(_azy[_baa]) + ";"; });
        _azx.forEach(function (_bab) {
            var _bac = String(_bab);
            if (!/^function [A-Za-z0-9_]+\s*\(/.test(_bac) || _bac.indexOf("[native code]") > -1)
                throw new Error("toString");
            _azz.push(_bac);
        });
        _azz.push("return { buscar: buscarRetornosLocal_ };");
        var _bad = _azz.join("\n");
        if (_bad.indexOf(String.fromCharCode(96)) > -1 || /<\/script/i.test(_bad))
            throw new Error("caracteres");
        var _bae = new Function(_bad)();
        var _baf = { n: 4, cab: ["A"], h: ["María García López", "Luis Pérez", 7000821, ""], j: ["ART-45210-B", "", "Crema 50ml", "maria garcia"], p: ["7000821406", "0123", "", "7000004"], pv: {},
            hv: {}, jv: {}, da: ["", "01/10/2026"], a: "0101", db: ["", "VELOX"], b: "1010", df: ["#ffffff", "#b6d7a8"], f: "0110" };
        [["maria garzia", "todo"], ["45210", "articulos"], ["7000821", "todo"], ["0123", "pedidos"], ["jose", "clientes"]].forEach(function (_bag) {
            if (JSON.stringify(_bae.buscar(_baf, _bag[0], _bag[1])) !== JSON.stringify(buscarRetornosLocal_(_baf, _bag[0], _bag[1])))
                throw new Error("distinto");
        });
        _azw = _bad;
    }
    catch (_bah) {
        _azw = "";
    }
    codigoBusquedaRetornosHecho_ = _azw;
    return _azw;
}
function pedidoCoincideRetornos_(_bai, _baj) {
    var _bak = compactoRetornos_(textoCelda_(_baj));
    return (_bak && _bai.compacto && _bak.indexOf(_bai.compacto) > -1) ? { pct: 100, exacta: true } : null;
}
function leerFilasRetornos_(_bal, _bam) {
    var _ban = {};
    var _bao = _bam.slice().sort(function (_bap, _baq) { return _bap - _baq; }), _bar = [];
    _bao.forEach(function (_bas) { var _bat = _bar[_bar.length - 1]; if (_bat && _bas - _bat.hasta <= 300)
        _bat.hasta = _bas;
    else
        _bar.push({ desde: _bas, hasta: _bas }); });
    _bar.forEach(function (_bau) {
        var _bav = _bau.hasta - _bau.desde + 1;
        var _baw = _bal.getRange(_bau.desde, 1, _bav, COL_RET_ULTIMA_).getDisplayValues(), _bax = _bal.getRange(_bau.desde, 1, _bav, 1).getBackgrounds();
        _bao.forEach(function (_bay) { if (_bay >= _bau.desde && _bay <= _bau.hasta)
            _ban[_bay] = { valores: _baw[_bay - _bau.desde], fondoA: _bax[_bay - _bau.desde][0] }; });
    });
    return _ban;
}
function leerFilasRetornosAntiguo_(_baz, _bba) {
    var _bbb = {};
    var _bbc = Math.min.apply(null, _bba), _bbd = Math.max.apply(null, _bba);
    if (_bba.length <= 6 || _bbd - _bbc + 1 > 1500) {
        _bba.forEach(function (_bbe) {
            var _bbf = _baz.getRange(_bbe, 1, 1, COL_RET_ULTIMA_);
            _bbb[_bbe] = { valores: _bbf.getDisplayValues()[0], fondoA: _bbf.getBackgrounds()[0][0] };
        });
    }
    else {
        var _bbg = _baz.getRange(_bbc, 1, _bbd - _bbc + 1, COL_RET_ULTIMA_);
        var _bbh = _bbg.getDisplayValues(), _bbi = _baz.getRange(_bbc, 1, _bbd - _bbc + 1, 1).getBackgrounds();
        _bba.forEach(function (_bbj) { _bbb[_bbj] = { valores: _bbh[_bbj - _bbc], fondoA: _bbi[_bbj - _bbc][0] }; });
    }
    return _bbb;
}
function agenciasRetornos_(_bbk) {
    var _bbl = null;
    try {
        _bbl = CacheService.getDocumentCache();
        var _bbm = _bbl && _bbl.get('RET_AGENCIAS');
        if (_bbm)
            return JSON.parse(_bbm);
    }
    catch (_bbn) { }
    var _bbo = _bbk.getLastRow() - 1, _bbp = [];
    if (_bbo > 0) {
        var _bbq = _bbk.getRange(2, COLUMNA_AGENCIAS, _bbo, 1).getValues();
        var _bbr = [];
        _bbq.forEach(function (_bbs) {
            var _bbt = String(_bbs[0] === null || _bbs[0] === undefined ? "" : _bbs[0]).trim().replace(/\s+/g, ' ');
            if (!_bbt)
                return;
            var _bbu = compactoRetornos_(_bbt);
            if (!_bbu)
                return;
            var _bbv = null;
            for (var _bbw = 0; _bbw < _bbr.length && !_bbv; _bbw++) {
                var _bbx = _bbr[_bbw].clave;
                if (_bbx === _bbu || (Math.abs(_bbx.length - _bbu.length) <= 2 && levenshtein(_bbx, _bbu) <= (Math.max(_bbx.length, _bbu.length) >= 10 ? 2 : 1)))
                    _bbv = _bbr[_bbw];
            }
            if (!_bbv) {
                _bbv = { clave: _bbu, formas: {}, total: 0 };
                _bbr.push(_bbv);
            }
            _bbv.formas[_bbt] = (_bbv.formas[_bbt] || 0) + 1;
            _bbv.total++;
        });
        _bbr.sort(function (_bby, _bbz) { return _bbz.total - _bby.total; });
        _bbp = _bbr.slice(0, 8).map(function (_bca) {
            var _bcb = null;
            Object.keys(_bca.formas).forEach(function (_bcc) { if (!_bcb || _bca.formas[_bcc] > _bca.formas[_bcb])
                _bcb = _bcc; });
            return _bcb;
        });
    }
    try {
        if (_bbl)
            _bbl.put('RET_AGENCIAS', JSON.stringify(_bbp), 600);
    }
    catch (_bcd) { }
    return _bbp;
}
function colorProcesadoRetornos_(_bce) {
    try {
        var _bcf = CacheService.getDocumentCache(), _bcg = _bcf.get('RET_VERDE');
        if (_bcg)
            return _bcg;
        var _bch = _bce.getLastRow(), _bci = Math.max(2, _bch - 400 + 1), _bcj = {}, _bck = null;
        if (_bch >= 2)
            _bce.getRange(_bci, 1, _bch - _bci + 1, 1).getBackgrounds().forEach(function (_bcl) {
                var _bcm = String(_bcl[0]).toLowerCase();
                if (esVerdeProcesado_(_bcm)) {
                    _bcj[_bcm] = (_bcj[_bcm] || 0) + 1;
                    if (!_bck || _bcj[_bcm] > _bcj[_bck])
                        _bck = _bcm;
                }
            });
        var _bcn = _bck || COLOR_PROCESADO_POR_DEFECTO_;
        _bcf.put('RET_VERDE', _bcn, 3600);
        return _bcn;
    }
    catch (_bco) {
        return COLOR_PROCESADO_POR_DEFECTO_;
    }
}
function opcionesDesplegablesFila_(_bcp, _bcq) {
    var _bcr = {};
    try {
        var _bcs = _bcp.getRange(_bcq, 1, 1, COL_RET_ULTIMA_).getDataValidations()[0];
        for (var _bct = 0; _bct < _bcs.length; _bct++) {
            var _bcu = _bcs[_bct];
            if (!_bcu)
                continue;
            var _bcv = _bcu.getCriteriaType(), _bcw = _bcu.getCriteriaValues();
            if (_bcv === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST)
                _bcr[_bct + 1] = (_bcw[0] || []).map(String).slice(0, 20);
            else if (_bcv === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE && _bcw[0]) {
                var _bcx = [];
                _bcw[0].getDisplayValues().forEach(function (_bcy) { _bcy.forEach(function (_bcz) { _bcz = String(_bcz).trim(); if (_bcz && _bcx.indexOf(_bcz) === -1)
                    _bcx.push(_bcz); }); });
                _bcr[_bct + 1] = _bcx.slice(0, 20);
            }
        }
    }
    catch (_bda) { }
    return _bcr;
}
var VENTANA_RELOCALIZAR_ = 400;
function relocalizarFilaPedido_(_bdb, _bdc, _bdd) {
    var _bde = normalizarRef(_bdd);
    if (!_bde)
        return 0;
    var _bdf = _bdb.getLastRow();
    if (_bdf < 2)
        return 0;
    var _bdg = Math.max(2, _bdc - VENTANA_RELOCALIZAR_), _bdh = Math.min(_bdf, _bdc + VENTANA_RELOCALIZAR_);
    if (_bdh < _bdg) {
        _bdg = Math.max(2, _bdf - VENTANA_RELOCALIZAR_);
        _bdh = _bdf;
    }
    var _bdi = _bdb.getRange(_bdg, COLUMNA_REFERENCIAS, _bdh - _bdg + 1, 1).getDisplayValues();
    var _bdj = 0, _bdk = Infinity, _bdl = false;
    for (var _bdm = 0; _bdm < _bdi.length; _bdm++) {
        if (normalizarRef(_bdi[_bdm][0]) !== _bde)
            continue;
        var _bdn = _bdg + _bdm, _bdo = Math.abs(_bdn - _bdc);
        if (_bdo < _bdk) {
            _bdj = _bdn;
            _bdk = _bdo;
            _bdl = false;
        }
        else if (_bdo === _bdk)
            _bdl = true;
    }
    return (_bdj && !_bdl) ? _bdj : 0;
}
function filaPedidoRetornos_(_bdp, _bdq, _bdr) {
    if (comprobarPedidoRetornos_(_bdp, _bdq, _bdr))
        return _bdq;
    return relocalizarFilaPedido_(_bdp, _bdq, _bdr);
}
function comprobarPedidoRetornos_(_bds, _bdt, _bdu) {
    if (!(_bdt >= 2) || _bdt > _bds.getLastRow())
        return false;
    var _bdv = _bds.getRange(_bdt, COLUMNA_REFERENCIAS).getDisplayValue();
    return normalizarRef(_bdv) === normalizarRef(_bdu);
}
function obtenerFilaRetornos(_bdw, _bdx) {
    var _bdy = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_RETORNOS);
    if (!_bdy)
        return { error: "No existe la pestaña Retornos." };
    _bdw = filaPedidoRetornos_(_bdy, _bdw, _bdx);
    if (!_bdw)
        return { error: "fila_cambiada" };
    var _bdz = _bdy.getRange(_bdw, 1, 1, COL_RET_ULTIMA_);
    return {
        fila: _bdw,
        cabeceras: _bdy.getRange(1, 1, 1, COL_RET_ULTIMA_).getDisplayValues()[0],
        valores: _bdz.getDisplayValues()[0],
        procesado: esVerdeProcesado_(_bdz.getBackgrounds()[0][0]),
        opciones: opcionesDesplegablesFila_(_bdy, _bdw),
        agencias: agenciasRetornos_(_bdy)
    };
}
function guardarFilaRetornos(_bea, _beb, _bec, _bed) {
    PLAZO_REGISTRO_ = new Date().getTime() + 20000;
    if (!usuarioAutorizado_())
        return { error: "No tienes permiso para usar este sistema." };
    _bed = _bed || {};
    _bec = _bec || {};
    var _bee = SpreadsheetApp.getActiveSpreadsheet(), _bef = _bee.getSheetByName(HOJA_RETORNOS);
    if (!_bef)
        return { error: "No existe la pestaña Retornos." };
    _bea = filaPedidoRetornos_(_bef, _bea, _beb);
    if (!_bea)
        return { error: "fila_cambiada" };
    var _beg = _bef.getRange(_bea, 1, 1, COL_RET_ULTIMA_);
    var _beh = _beg.getValues()[0], _bei = _beg.getDisplayValues()[0], _bej = _beg.getBackgrounds()[0];
    var _bek = {};
    Object.keys(_bec).forEach(function (_bel) {
        var _bem = Number(_bel);
        if (_bem >= 1 && _bem <= COL_RET_ULTIMA_ && String(_bec[_bel]) !== String(_bei[_bem - 1]))
            _bek[_bem] = String(_bec[_bel]);
    });
    if (_bed.procesar) {
        if (String(_bei[COLUMNA_RETORNOS_FECHA - 1]).trim() === "" && _bek[COLUMNA_RETORNOS_FECHA] === undefined)
            _bek[COLUMNA_RETORNOS_FECHA] = "__HOY__";
        var _ben = String(_bed.agencia || "").trim();
        if (_ben && _bek[COLUMNA_AGENCIAS] === undefined && String(_bei[COLUMNA_AGENCIAS - 1]).trim() !== _ben)
            _bek[COLUMNA_AGENCIAS] = _ben;
    }
    var _beo = Object.keys(_bek).map(Number);
    _beo.forEach(function (_bep) {
        var _beq = _bef.getRange(_bea, _bep);
        if (_bek[_bep] === "__HOY__") {
            var _ber = new Date();
            _ber.setHours(0, 0, 0, 0);
            _beq.setNumberFormat('dd/mm/yyyy');
            _beq.setValue(_ber);
        }
        else {
            if (_bep === COLUMNA_REFERENCIAS)
                _beq.setNumberFormat('@');
            _beq.setValue(_bek[_bep]);
        }
    });
    var _bes = false;
    if (_bed.procesar) {
        _beg.setBackground(colorProcesadoRetornos_(_bef));
        _bes = true;
    }
    else if (_bed.quitarVerde) {
        _beg.setBackground(null);
        _bes = true;
    }
    try {
        var _bet = {};
        _beo.forEach(function (_beu) { var _bev = _beh[_beu - 1]; _bet[_beu] = (_bev instanceof Date) ? { fecha: _bev.getTime() } : _bev; });
        CacheService.getUserCache().put('RET_DESHACER_' + _bea, JSON.stringify({ valores: _bet, fondos: _bes ? _bej : null, pedido: _bek[COLUMNA_REFERENCIAS] !== undefined ? _bek[COLUMNA_REFERENCIAS] : _beb }), 300);
    }
    catch (_bew) { }
    if (_beo.length > 0) {
        var _bex = Math.min.apply(null, _beo), _bey = Math.max.apply(null, _beo);
        try {
            registrarCambioCelda_(_bee, _bef, _bea, _bex, 1, _bey - _bex + 1);
        }
        catch (_bez) { }
        var _bfa = {};
        _bfa[_bea] = _beh;
        try {
            registrarCambiosRetornos_(_bef, _bea, _bex, 1, _bey - _bex + 1, { antesFilas: _bfa, tipo: "rellenar" });
        }
        catch (_bfb) { }
    }
    SpreadsheetApp.flush();
    var _bfc = _bef.getRange(_bea, 1).getBackground();
    return { ok: true, fila: _bea, escritas: _beo.length, valores: _beg.getDisplayValues()[0], procesado: esVerdeProcesado_(_bfc), color: _bfc };
}
function deshacerFilaRetornos(_bfd) {
    PLAZO_REGISTRO_ = new Date().getTime() + 20000;
    if (!usuarioAutorizado_())
        return { error: "No tienes permiso para usar este sistema." };
    var _bfe = SpreadsheetApp.getActiveSpreadsheet(), _bff = _bfe.getSheetByName(HOJA_RETORNOS);
    var _bfg = _bfd;
    var _bfh = CacheService.getUserCache(), _bfi = _bfh.get('RET_DESHACER_' + _bfg);
    if (!_bff || !_bfi)
        return { error: "Ya no se puede deshacer (han pasado más de 5 minutos)." };
    var _bfj = JSON.parse(_bfi);
    _bfd = filaPedidoRetornos_(_bff, _bfd, _bfj.pedido);
    if (!_bfd)
        return { error: "fila_cambiada" };
    var _bfk = Object.keys(_bfj.valores).map(Number);
    var _bfl = _bff.getRange(_bfd, 1, 1, COL_RET_ULTIMA_).getValues()[0];
    _bfk.forEach(function (_bfm) {
        var _bfn = _bfj.valores[_bfm];
        _bff.getRange(_bfd, _bfm).setValue(_bfn && typeof _bfn === 'object' && _bfn.fecha ? new Date(_bfn.fecha) : _bfn);
    });
    if (_bfj.fondos)
        _bff.getRange(_bfd, 1, 1, COL_RET_ULTIMA_).setBackgrounds([_bfj.fondos]);
    if (_bfk.length > 0) {
        var _bfo = Math.min.apply(null, _bfk), _bfp = Math.max.apply(null, _bfk);
        try {
            registrarCambioCelda_(_bfe, _bff, _bfd, _bfo, 1, _bfp - _bfo + 1);
        }
        catch (_bfq) { }
        var _bfr = {};
        _bfr[_bfd] = _bfl;
        try {
            registrarCambiosRetornos_(_bff, _bfd, _bfo, 1, _bfp - _bfo + 1, { antesFilas: _bfr, tipo: "deshacer" });
        }
        catch (_bfs) { }
    }
    _bfh.remove('RET_DESHACER_' + _bfg);
    SpreadsheetApp.flush();
    var _bft = _bff.getRange(_bfd, 1, 1, COL_RET_ULTIMA_), _bfu = _bff.getRange(_bfd, 1).getBackground();
    return { ok: true, fila: _bfd, valores: _bft.getDisplayValues()[0], procesado: esVerdeProcesado_(_bfu), color: _bfu };
}
var VENTANA_COPIA_RET_ = 2500;
var VENTANA_COPIA_MAX_ = 3500;
var MARGEN_COPIA_RET_ = 200;
var DIAS_CAMBIOS_RET_ = 30;
var MAX_CELDAS_EVENTO_RET_ = 200;
var COLS_COPIA_RET_ = 15;
var CABECERA_CAMBIOS_RET_ = ["Fecha", "Usuario", "Tipo", "Filas", "Pedido", "Cliente", "Columnas", "Cambios", "Total", "Id", "Restaurado"];
function trozosACadena_(_bfv) {
    var _bfw = "";
    for (var _bfx = 0; _bfx < _bfv.length; _bfx += 8192)
        _bfw += String.fromCharCode.apply(null, _bfv.slice(_bfx, _bfx + 8192));
    return _bfw;
}
function codificarTexto_(_bfy) {
    _bfy = (_bfy === undefined || _bfy === null) ? "" : String(_bfy);
    if (_bfy === "")
        return "";
    var _bfz = new Array(_bfy.length + 1);
    _bfz[0] = 0xE001;
    for (var _bga = 0; _bga < _bfy.length; _bga++) {
        var _bgb = _bfy.charCodeAt(_bga);
        _bfz[_bga + 1] = _bgb < 0x1800 ? 0xE100 + _bgb : _bgb;
    }
    return trozosACadena_(_bfz);
}
function decodificarTexto_(_bgc) {
    if (_bgc === undefined || _bgc === null)
        return "";
    if (typeof _bgc !== "string")
        return String(_bgc);
    if (_bgc.charCodeAt(0) !== 0xE001)
        return _bgc;
    var _bgd = new Array(_bgc.length - 1);
    for (var _bge = 1; _bge < _bgc.length; _bge++) {
        var _bgf = _bgc.charCodeAt(_bge);
        _bgd[_bge - 1] = (_bgf >= 0xE100 && _bgf < 0xF900) ? _bgf - 0xE100 : _bgf;
    }
    return trozosACadena_(_bgd);
}
function valorACodigo_(_bgg) {
    if (_bgg === "" || _bgg === null || _bgg === undefined)
        return "";
    if (_bgg instanceof Date)
        return "D" + _bgg.getTime();
    if (typeof _bgg === "number")
        return "N" + _bgg;
    if (typeof _bgg === "boolean")
        return "B" + (_bgg ? 1 : 0);
    return "S" + String(_bgg);
}
function codigoAValor_(_bgh) {
    if (!_bgh)
        return "";
    var _bgi = _bgh.charAt(0), _bgj = _bgh.substring(1);
    if (_bgi === "D")
        return new Date(Number(_bgj));
    if (_bgi === "N")
        return Number(_bgj);
    if (_bgi === "B")
        return _bgj === "1";
    return _bgj;
}
function codigoATexto_(_bgk) {
    if (!_bgk)
        return "";
    var _bgl = _bgk.charAt(0), _bgm = _bgk.substring(1);
    if (_bgl === "D") {
        var _bgn = new Date(Number(_bgm));
        return ('0' + _bgn.getDate()).slice(-2) + '/' + ('0' + (_bgn.getMonth() + 1)).slice(-2) + '/' + _bgn.getFullYear();
    }
    if (_bgl === "B")
        return _bgm === "1" ? "Sí" : "No";
    return _bgm;
}
function pedidoDeValor_(_bgo) { return normalizarRef(textoCelda_(_bgo)); }
function hojaCambiosRetornos_(_bgp) {
    var _bgq = _bgp.getSheetByName(HOJA_CAMBIOS_RET_);
    if (!_bgq) {
        _bgq = _bgp.insertSheet(HOJA_CAMBIOS_RET_);
        _bgq.getRange(1, 1, 1, CABECERA_CAMBIOS_RET_.length).setValues([CABECERA_CAMBIOS_RET_]);
        try {
            _bgq.hideSheet();
        }
        catch (_bgr) { }
    }
    return _bgq;
}
function registrarCambiosRetornos_(_bgs, _bgt, _bgu, _bgv, _bgw, _bgx) {
    _bgx = _bgx || {};
    if (_bgt < 2) {
        _bgv -= (2 - _bgt);
        _bgt = 2;
    }
    if (_bgv <= 0 || _bgu > COL_RET_ULTIMA_)
        return 0;
    var _bgy = Math.min(_bgu + _bgw - 1, COL_RET_ULTIMA_);
    var _bgz = SpreadsheetApp.getActiveSpreadsheet();
    var _bha = PropertiesService.getDocumentProperties(), _bhb = _bha.getProperties();
    var _bhc = Number(_bhb.RET_COPIA_INICIO || 0), _bhd = Number(_bhb.RET_COPIA_FILAS || 0), _bhe = _bhb.RET_COPIA_DESCUADRE === '1';
    var _bhf = (_bhc > 0 && _bhd > 0 && !_bhe) ? _bgz.getSheetByName(HOJA_COPIA_RET_) : null;
    var _bhg = _bgt + _bgv - 1, _bhh = _bhc + _bhd - 1;
    var _bhi = _bgs.getRange(_bgt, 1, _bgv, COL_RET_ULTIMA_).getValues();
    var _bhj = null, _bhk = Math.max(_bgt, _bhc), _bhl = Math.min(_bhg, _bhh);
    if (_bhf && _bhk <= _bhl) {
        var _bhm = _bhf.getRange(_bhk - _bhc + 2, 1, _bhl - _bhk + 1, COLS_COPIA_RET_).getValues();
        _bhj = {};
        var _bhn = true;
        for (var _bho = 0; _bho < _bhm.length && _bhn; _bho++) {
            var _bhp = _bhk + _bho, _bhq = _bhm[_bho].map(decodificarTexto_), _bhr = _bhi[_bhp - _bgt];
            if ((_bgu > COLUMNA_REFERENCIAS || _bgy < COLUMNA_REFERENCIAS) && _bhq[COL_RET_ULTIMA_] !== pedidoDeValor_(_bhr[COLUMNA_REFERENCIAS - 1]))
                _bhn = false;
            for (var _bhs = 1; _bhn && _bhs <= COL_RET_ULTIMA_; _bhs++) {
                if (_bhs >= _bgu && _bhs <= _bgy)
                    continue;
                if (_bhq[_bhs - 1] !== valorACodigo_(_bhr[_bhs - 1]))
                    _bhn = false;
            }
            _bhj[_bhp] = _bhq;
        }
        if (!_bhn) {
            _bhj = null;
            _bhf = null;
            try {
                _bha.setProperty('RET_COPIA_DESCUADRE', '1');
            }
            catch (_bht) { }
        }
    }
    var _bhu = [], _bhv = 0, _bhw = {};
    for (var _bhx = _bgt; _bhx <= _bhg; _bhx++) {
        var _bhy = _bhi[_bhx - _bgt], _bhz = null;
        if (_bgx.antesFilas && _bgx.antesFilas[_bhx])
            _bhz = _bgx.antesFilas[_bhx].map(valorACodigo_);
        else if (_bhj && _bhj[_bhx])
            _bhz = _bhj[_bhx];
        for (var _bia = _bgu; _bia <= _bgy; _bia++) {
            var _bib = valorACodigo_(_bhy[_bia - 1]), _bic;
            if (_bhz)
                _bic = _bhz[_bia - 1];
            else if (_bgv === 1 && _bgu === _bgy && _bgx.oldValue !== undefined && _bgx.oldValue !== null)
                _bic = String(_bgx.oldValue) === "" ? "" : "T" + _bgx.oldValue;
            else
                continue;
            if (!_bic || _bic === _bib)
                continue;
            if (_bic.charAt(0) === "T" && codigoATexto_(_bib) === _bic.substring(1))
                continue;
            _bhv++;
            _bhw[_bia] = true;
            if (_bhu.length < MAX_CELDAS_EVENTO_RET_) {
                var _bid = _bhz ? codigoATexto_(_bhz[COLUMNA_REFERENCIAS - 1]) : textoCelda_(_bhy[COLUMNA_REFERENCIAS - 1]);
                var _bie = _bhz ? codigoATexto_(_bhz[COL_RET_CLIENTE_ - 1]) : textoCelda_(_bhy[COL_RET_CLIENTE_ - 1]);
                var _bif = String(_bid || textoCelda_(_bhy[COLUMNA_REFERENCIAS - 1]) || "").trim();
                var _big = { f: _bhx, c: _bia, a: _bic, d: _bib, p: normalizarRef(_bif), n: _bie || textoCelda_(_bhy[COL_RET_CLIENTE_ - 1]) };
                if (_bif && _bif !== _big.p)
                    _big.pm = _bif;
                _bhu.push(_big);
            }
        }
    }
    if (_bhf) {
        if (_bhj && _bhk <= _bhl) {
            var _bih = [];
            for (var _bii = _bhk; _bii <= _bhl; _bii++)
                _bih.push(filaACopia_(_bhi[_bii - _bgt]));
            _bhf.getRange(_bhk - _bhc + 2, 1, _bih.length, COLS_COPIA_RET_).setValues(_bih);
        }
        if (_bhg > _bhh) {
            var _bij = _bhh + 1, _bik = [];
            if (_bgt > _bij) {
                if (_bgt - _bij > 500) {
                    try {
                        _bha.setProperty('RET_COPIA_DESCUADRE', '1');
                    }
                    catch (_bil) { }
                    _bij = -1;
                }
                else
                    _bgs.getRange(_bij, 1, _bgt - _bij, COL_RET_ULTIMA_).getValues().forEach(function (_bim) { _bik.push(filaACopia_(_bim)); });
            }
            if (_bij > 0) {
                for (var _bin = Math.max(_bgt, _bij); _bin <= _bhg; _bin++)
                    _bik.push(filaACopia_(_bhi[_bin - _bgt]));
                if (_bhf.getMaxRows() < _bhd + 1 + _bik.length)
                    _bhf.insertRowsAfter(_bhf.getMaxRows(), _bhd + 1 + _bik.length - _bhf.getMaxRows());
                _bhf.getRange(_bhd + 2, 1, _bik.length, COLS_COPIA_RET_).setValues(_bik);
                try {
                    _bha.setProperty('RET_COPIA_FILAS', String(_bhd + _bik.length));
                }
                catch (_bio) { }
            }
        }
    }
    if (_bhv === 0)
        return 0;
    var _bip = _bgx.usuario || "";
    if (!_bip) {
        try {
            _bip = Session.getActiveUser().getEmail() || "";
        }
        catch (_biq) { }
    }
    if (!_bip) {
        try {
            _bip = PropertiesService.getUserProperties().getProperty('devoluciones_email') || "";
        }
        catch (_bir) { }
    }
    var _bis = _bgx.tipo;
    if (!_bis) {
        var _bit = _bhu.every(function (_biu) { return _biu.d === ""; });
        _bis = (_bhv === 1 ? (_bit ? "borrado" : "cambio") : (_bit ? "borrado_varias" : "cambio_varias"));
    }
    var _biv = _bgt === _bhg ? String(_bgt) : (_bgt + "-" + _bhg);
    var _biw = Object.keys(_bhw).map(function (_bix) { return columnToLetter(Number(_bix)); }).join(",");
    var _biy = new Date();
    var _biz = _biy.getTime() + "-" + _bgt + "-" + _bgu + "-" + Math.floor(Math.random() * 1000);
    hojaCambiosRetornos_(_bgz).appendRow([_biy, codificarTexto_(_bip || "Usuario"), codificarTexto_(_bis), codificarTexto_(_biv), codificarTexto_(_bhu[0].pm || _bhu[0].p), codificarTexto_(_bhu[0].n),
        codificarTexto_(_biw), codificarTexto_(JSON.stringify(_bhu)), _bhv, codificarTexto_(_biz), ""]);
    return _bhv;
}
function filaACopia_(_bja) {
    var _bjb = [];
    for (var _bjc = 0; _bjc < COL_RET_ULTIMA_; _bjc++)
        _bjb.push(codificarTexto_(valorACodigo_(_bja[_bjc])));
    _bjb.push(codificarTexto_(pedidoDeValor_(_bja[COLUMNA_REFERENCIAS - 1])));
    return _bjb;
}
function reconstruirCopiaRetornos_(_bjd, _bje, _bjf) {
    try {
        _bjf.setProperty('RET_COPIA_DESCUADRE', '1');
    }
    catch (_bjg) { }
    var _bjh = _bjd.getSheetByName(HOJA_COPIA_RET_);
    if (!_bjh) {
        _bjh = _bjd.insertSheet(HOJA_COPIA_RET_);
        try {
            _bjh.hideSheet();
        }
        catch (_bji) { }
    }
    var _bjj = _bje.getLastRow();
    var _bjk = Math.max(2, _bjj - VENTANA_COPIA_RET_ + 1), _bjl = Math.max(0, _bjj - _bjk + 1);
    var _bjm = _bjl > 0 ? _bje.getRange(_bjk, 1, _bjl, COL_RET_ULTIMA_).getValues().map(filaACopia_) : [];
    if (_bjh.getMaxColumns() < COLS_COPIA_RET_)
        _bjh.insertColumnsAfter(_bjh.getMaxColumns(), COLS_COPIA_RET_ - _bjh.getMaxColumns());
    if (_bjh.getMaxRows() < _bjl + 1)
        _bjh.insertRowsAfter(_bjh.getMaxRows(), _bjl + 1 - _bjh.getMaxRows());
    _bjh.clearContents();
    var _bjn = [];
    for (var _bjo = 0; _bjo < COLS_COPIA_RET_; _bjo++)
        _bjn.push(_bjo === 0 ? "Copia interna de Retornos (no tocar)" : "");
    _bjh.getRange(1, 1, 1, COLS_COPIA_RET_).setValues([_bjn]);
    if (_bjl > 0)
        _bjh.getRange(2, 1, _bjl, COLS_COPIA_RET_).setValues(_bjm);
    _bjf.setProperties({ RET_COPIA_INICIO: String(_bjk), RET_COPIA_FILAS: String(_bjl), RET_COPIA_DESCUADRE: '0', RET_COPIA_T: String(new Date().getTime()) });
    return _bjl;
}
function sincronizarCopiaRetornos_(_bjp) {
    _bjp = _bjp || SpreadsheetApp.getActiveSpreadsheet();
    var _bjq = _bjp.getSheetByName(HOJA_RETORNOS);
    if (!_bjq)
        return { hecho: "sin_retornos" };
    var _bjr = PropertiesService.getDocumentProperties(), _bjs = _bjr.getProperties();
    var _bjt = Number(_bjs.RET_COPIA_INICIO || 0), _bju = Number(_bjs.RET_COPIA_FILAS || 0);
    var _bjv = _bjp.getSheetByName(HOJA_COPIA_RET_);
    var _bjw = _bjq.getLastRow();
    if (!_bjv || !_bjt) {
        reconstruirCopiaRetornos_(_bjp, _bjq, _bjr);
        return { hecho: "creada" };
    }
    var _bjx = Number(_bjs.RET_COPIA_T || 0), _bjy = new Date().getTime();
    var _bjz = Math.max(2, _bjt - MARGEN_COPIA_RET_);
    var _bka = _bjw >= _bjz ? _bjq.getRange(_bjz, COLUMNA_REFERENCIAS, _bjw - _bjz + 1, 1).getValues().map(function (_bkb) { return pedidoDeValor_(_bkb[0]); }) : [];
    var _bkc = _bju > 0 ? _bjv.getRange(2, COLS_COPIA_RET_, _bju, 1).getValues().map(function (_bkd) { return decodificarTexto_(_bkd[0]); }) : [];
    var _bke = _bjt - _bjz, _bkf = _bjs.RET_COPIA_DESCUADRE !== '1' && _bka.length >= _bke + _bju;
    for (var _bkg = 0; _bkf && _bkg < _bju; _bkg++)
        if (_bkc[_bkg] !== _bka[_bke + _bkg])
            _bkf = false;
    if (_bkf) {
        if (_bjw - _bjt + 1 > VENTANA_COPIA_MAX_) {
            reconstruirCopiaRetornos_(_bjp, _bjq, _bjr);
            return { hecho: "rehecha" };
        }
        if (_bjw > _bjt + _bju - 1) {
            var _bkh = _bjq.getRange(_bjt + _bju, 1, _bjw - _bjt - _bju + 1, COL_RET_ULTIMA_).getValues().map(filaACopia_);
            if (_bjv.getMaxRows() < _bju + 1 + _bkh.length)
                _bjv.insertRowsAfter(_bjv.getMaxRows(), _bju + 1 + _bkh.length - _bjv.getMaxRows());
            _bjv.getRange(_bju + 2, 1, _bkh.length, COLS_COPIA_RET_).setValues(_bkh);
            _bjr.setProperties({ RET_COPIA_FILAS: String(_bju + _bkh.length), RET_COPIA_T: String(_bjy) });
            return { hecho: "ampliada" };
        }
        _bjr.setProperty('RET_COPIA_T', String(_bjy));
        return { hecho: "igual" };
    }
    var _bki = {};
    _bka.forEach(function (_bkj) { if (_bkj)
        _bki[_bkj] = (_bki[_bkj] || 0) + 1; });
    var _bkk = _bjv.getRange(2, 1, _bju, COLS_COPIA_RET_).getValues();
    var _bkl = [];
    for (var _bkm = 0; _bkm < _bju; _bkm++) {
        var _bkn = _bkc[_bkm];
        if (!_bkn)
            continue;
        if (_bki[_bkn] > 0) {
            _bki[_bkn]--;
            continue;
        }
        _bkl.push({ fila: _bjt + _bkm, codigos: _bkk[_bkm].slice(0, COL_RET_ULTIMA_).map(decodificarTexto_) });
    }
    if (_bkl.length > 0) {
        var _bko = hojaCambiosRetornos_(_bjp), _bkp = [], _bkq = new Date();
        _bkl.slice(0, 60).forEach(function (_bkr, _bks) {
            var _bkt = [{ f: _bkr.fila, c: 0, fila: _bkr.codigos, p: normalizarRef(codigoATexto_(_bkr.codigos[COLUMNA_REFERENCIAS - 1])), n: codigoATexto_(_bkr.codigos[COL_RET_CLIENTE_ - 1]), desde: _bjx }];
            _bkp.push([_bkq, "", codificarTexto_("fila_eliminada"), codificarTexto_(String(_bkr.fila)), codificarTexto_(String(codigoATexto_(_bkr.codigos[COLUMNA_REFERENCIAS - 1]) || _bkt[0].p).trim()), codificarTexto_(_bkt[0].n), codificarTexto_("A:N"),
                codificarTexto_(JSON.stringify(_bkt)), 1, codificarTexto_(_bkq.getTime() + "-e" + _bks), ""]);
        });
        _bko.getRange(_bko.getLastRow() + 1, 1, _bkp.length, CABECERA_CAMBIOS_RET_.length).setValues(_bkp);
    }
    reconstruirCopiaRetornos_(_bjp, _bjq, _bjr);
    return { hecho: "rehecha", eliminadas: _bkl.length };
}
function purgarCambiosRetornos_(_bku) {
    var _bkv = _bku.getSheetByName(HOJA_CAMBIOS_RET_);
    if (!_bkv || _bkv.getLastRow() < 2)
        return 0;
    var _bkw = _bkv.getRange(2, 1, _bkv.getLastRow() - 1, 1).getValues(), _bkx = new Date().getTime() - DIAS_CAMBIOS_RET_ * 86400000, _bky = 0;
    while (_bky < _bkw.length && _bkw[_bky][0] instanceof Date && _bkw[_bky][0].getTime() < _bkx)
        _bky++;
    if (_bky === 0)
        return 0;
    if (_bkv.getMaxRows() <= _bkv.getLastRow())
        _bkv.insertRowsAfter(_bkv.getMaxRows(), 1);
    _bkv.deleteRows(2, _bky);
    return _bky;
}
function leerCambiosRetornos_(_bkz, _bla, _blb) {
    var _blc = _bkz.getSheetByName(HOJA_CAMBIOS_RET_);
    if (!_blc || _blc.getLastRow() < 2)
        return [];
    var _bld = new Date().getTime() - _bla * 86400000;
    var _ble = _blc.getLastRow() - 1, _blf = Math.max(0, _ble - 3000);
    var _blg = _blc.getRange(2 + _blf, 1, _ble - _blf, CABECERA_CAMBIOS_RET_.length).getValues();
    var _blh = [];
    for (var _bli = _blg.length - 1; _bli >= 0; _bli--) {
        var _blj = _blg[_bli][0];
        if (!(_blj instanceof Date) || _blj.getTime() < _bld)
            break;
        var _blk = { id: decodificarTexto_(_blg[_bli][9]), t: _blj.getTime(), usuario: decodificarTexto_(_blg[_bli][1]), tipo: decodificarTexto_(_blg[_bli][2]), filas: decodificarTexto_(_blg[_bli][3]),
            pedido: decodificarTexto_(_blg[_bli][4]), cliente: decodificarTexto_(_blg[_bli][5]), columnas: decodificarTexto_(_blg[_bli][6]), total: Number(_blg[_bli][8]) || 0,
            restaurado: decodificarTexto_(_blg[_bli][10]), cambios: [] };
        try {
            _blk.cambios = JSON.parse(decodificarTexto_(_blg[_bli][7]));
        }
        catch (_bll) {
            _blk.cambios = [];
        }
        if (_blb && !_blb(_blk))
            continue;
        _blh.push(_blk);
    }
    return _blh;
}
function prepararEventoCambios_(_blm, _bln) {
    var _blo = _blm.usuario ? _blm.usuario.split("@")[0] : "";
    var _blp = { id: _blm.id, t: _blm.t, usuario: _blo, tipo: _blm.tipo, filas: _blm.filas, pedido: _blm.pedido, cliente: _blm.cliente, total: _blm.total, restaurado: _blm.restaurado, cambios: [] };
    if (_blm.tipo === "fila_eliminada" && _blm.cambios[0]) {
        var _blq = _blm.cambios[0];
        _blp.desde = _blq.desde || 0;
        _blp.fila = (_blq.fila || []).map(function (_blr, _bls) { return { col: columnToLetter(_bls + 1), nombre: _bln[_bls] || columnToLetter(_bls + 1), valor: codigoATexto_(_blr) }; }).filter(function (_blt) { return _blt.valor !== ""; });
        return _blp;
    }
    _blm.cambios.slice(0, 40).forEach(function (_blu) {
        _blp.cambios.push({ f: _blu.f, col: columnToLetter(_blu.c), nombre: _bln[_blu.c - 1] || columnToLetter(_blu.c), antes: codigoATexto_(_blu.a), despues: codigoATexto_(_blu.d), p: _blu.p });
    });
    return _blp;
}
function anotarQuienEditaba_(_blv, _blw) {
    _blv.forEach(function (_blx) {
        if (_blx.tipo !== "fila_eliminada")
            return;
        var _bly = (_blx.desde || (_blx.t - 6 * 60000)) - 60000, _blz = _blx.t + 60000, _bma = {};
        _blw.forEach(function (_bmb) {
            if (_bmb.tipo === "fila_eliminada" || _bmb.tipo === "restaurado" || _bmb.t < _bly || _bmb.t > _blz)
                return;
            var _bmc = _bmb.usuario ? String(_bmb.usuario).split("@")[0] : "";
            if (_bmc && _bmc !== "Usuario")
                _bma[_bmc] = true;
        });
        _blx.editando = Object.keys(_bma).slice(0, 5);
    });
}
function pedidosComoEnLaHoja_(_bmd, _bme) {
    if (!_bmd)
        return;
    var _bmf = [];
    _bme.forEach(function (_bmg) { var _bmh = parseInt(_bmg.filas, 10); if (_bmg.tipo !== "fila_eliminada" && _bmh >= 2 && _bmg.pedido)
        _bmf.push(_bmh); });
    if (!_bmf.length)
        return;
    var _bmi = Math.min.apply(null, _bmf), _bmj = Math.min(Math.max.apply(null, _bmf), _bmd.getLastRow());
    if (_bmj < _bmi || _bmj - _bmi > 6000)
        return;
    var _bmk = _bmd.getRange(_bmi, COLUMNA_REFERENCIAS, _bmj - _bmi + 1, 1).getDisplayValues();
    _bme.forEach(function (_bml) {
        var _bmm = parseInt(_bml.filas, 10);
        if (_bml.tipo === "fila_eliminada" || !(_bmm >= _bmi && _bmm <= _bmj) || !_bml.pedido)
            return;
        var _bmn = String(_bmk[_bmm - _bmi][0] || "").trim();
        if (_bmn && normalizarRef(_bmn) === normalizarRef(_bml.pedido))
            _bml.pedido = _bmn;
    });
}
function obtenerCambiosRetornos() {
    if (!usuarioAutorizado_())
        return { error: "Herramienta de uso exclusivo para Devoluciones." };
    var _bmo = SpreadsheetApp.getActiveSpreadsheet(), _bmp = _bmo.getSheetByName(HOJA_RETORNOS);
    var _bmq = _bmp ? _bmp.getRange(1, 1, 1, COL_RET_ULTIMA_).getDisplayValues()[0] : [];
    var _bmr = leerCambiosRetornos_(_bmo, 7);
    var _bms = _bmr.slice(0, 300).map(function (_bmt) { return prepararEventoCambios_(_bmt, _bmq); });
    anotarQuienEditaba_(_bms, _bmr);
    try {
        pedidosComoEnLaHoja_(_bmp, _bms);
    }
    catch (_bmu) { }
    return { eventos: _bms };
}
function obtenerCambiosFilaRetornos(_bmv, _bmw) {
    if (!usuarioAutorizado_())
        return { error: "Herramienta de uso exclusivo para Devoluciones." };
    var _bmx = SpreadsheetApp.getActiveSpreadsheet(), _bmy = _bmx.getSheetByName(HOJA_RETORNOS);
    var _bmz = _bmy ? _bmy.getRange(1, 1, 1, COL_RET_ULTIMA_).getDisplayValues()[0] : [];
    var _bna = normalizarRef(_bmw), _bnb = Number(_bmv);
    var _bnc = leerCambiosRetornos_(_bmx, DIAS_CAMBIOS_RET_);
    var _bnd = _bnc.filter(function (_bne) {
        return _bne.cambios.some(function (_bnf) { return _bna ? _bnf.p === _bna : _bnf.f === _bnb; });
    }).slice(0, 30).map(function (_bng) { return prepararEventoCambios_(_bng, _bmz); });
    anotarQuienEditaba_(_bnd, _bnc);
    try {
        pedidosComoEnLaHoja_(_bmy, _bnd);
    }
    catch (_bnh) { }
    return { eventos: _bnd };
}
function restaurarCambioRetornos(_bni) {
    if (!usuarioAutorizado_())
        return { error: "Herramienta de uso exclusivo para Devoluciones." };
    var _bnj = SpreadsheetApp.getActiveSpreadsheet(), _bnk = _bnj.getSheetByName(HOJA_RETORNOS), _bnl = _bnj.getSheetByName(HOJA_CAMBIOS_RET_);
    if (!_bnk || !_bnl || _bnl.getLastRow() < 2)
        return { error: "No se encuentra ese cambio." };
    var _bnm = _bnl.getLastRow() - 1, _bnn = Math.max(0, _bnm - 3000);
    var _bno = _bnl.getRange(2 + _bnn, 10, _bnm - _bnn, 1).getValues(), _bnp = -1;
    for (var _bnq = _bno.length - 1; _bnq >= 0; _bnq--)
        if (decodificarTexto_(_bno[_bnq][0]) === String(_bni)) {
            _bnp = 2 + _bnn + _bnq;
            break;
        }
    if (_bnp < 0)
        return { error: "No se encuentra ese cambio (quizá tiene más de 30 días)." };
    var _bnr = _bnl.getRange(_bnp, 1, 1, CABECERA_CAMBIOS_RET_.length).getValues()[0];
    if (decodificarTexto_(_bnr[10]))
        return { error: "Ese cambio ya se restauró (" + decodificarTexto_(_bnr[10]) + ")." };
    var _bns = [];
    try {
        _bns = JSON.parse(decodificarTexto_(_bnr[7]));
    }
    catch (_bnt) { }
    var _bnu = "";
    try {
        _bnu = Session.getActiveUser().getEmail() || "";
    }
    catch (_bnv) { }
    var _bnw = new Date(), _bnx = (_bnu ? _bnu.split("@")[0] : "alguien") + " · " + ('0' + _bnw.getDate()).slice(-2) + "/" + ('0' + (_bnw.getMonth() + 1)).slice(-2) + " " + ('0' + _bnw.getHours()).slice(-2) + ":" + ('0' + _bnw.getMinutes()).slice(-2);
    if (decodificarTexto_(_bnr[2]) === "fila_eliminada" && _bns[0] && _bns[0].fila) {
        var _bny = _bns[0].fila.map(codigoAValor_);
        var _bnz = _bnk.getLastRow() + 1;
        if (_bnk.getMaxRows() < _bnz)
            _bnk.insertRowsAfter(_bnk.getMaxRows(), 1);
        _bnk.getRange(_bnz, COLUMNA_REFERENCIAS).setNumberFormat('@');
        _bnk.getRange(_bnz, 1, 1, COL_RET_ULTIMA_).setValues([_bny]);
        try {
            registrarCambiosRetornos_(_bnk, _bnz, 1, 1, COL_RET_ULTIMA_, { usuario: _bnu, tipo: "restaurado" });
        }
        catch (_boa) { }
        var _bob = _bnx + " (vuelta a añadir en la fila " + _bnz + ")";
        _bnl.getRange(_bnp, 11).setValue(codificarTexto_(_bob));
        return { ok: true, restauradas: 1, conflictos: [], filaNueva: _bnz, marca: _bob };
    }
    var _boc = {};
    _bns.forEach(function (_bod) { if (_bod.c >= 1)
        (_boc[_bod.f] = _boc[_bod.f] || []).push(_bod); });
    var _boe = 0, _bof = [], _bog = {}, _boh = [];
    var _boi = _bnk.getLastRow();
    Object.keys(_boc).map(Number).sort(function (_boj, _bok) { return _boj - _bok; }).forEach(function (_bol) {
        var _bom = localizarFilaCambio_(_bnk, _bol, _boc[_bol], _boi);
        if (_bom < 0) {
            _bof.push({ f: _bol, motivo: _bom === -2 ? "Hay varias filas con ese pedido y no se sabe cuál es." : "La fila ya no es la de ese pedido (se han movido o eliminado filas)." });
            return;
        }
        var _bon = _bnk.getRange(_bom, 1, 1, COL_RET_ULTIMA_).getValues()[0], _boo = _bon.slice(), _bop = false;
        _boc[_bol].forEach(function (_boq) {
            var _bor = valorACodigo_(_bon[_boq.c - 1]);
            var _bos = _bor === _boq.d || (_boq.d.charAt(0) === "T" && codigoATexto_(_bor) === _boq.d.substring(1));
            if (!_bos) {
                _bof.push({ f: _bom, col: columnToLetter(_boq.c), motivo: "Ya se ha vuelto a cambiar (ahora pone «" + codigoATexto_(_bor) + "»)." });
                return;
            }
            _boo[_boq.c - 1] = codigoAValor_(_boq.a);
            _bop = true;
            _boe++;
        });
        if (!_bop)
            return;
        _bog[_bom] = _bon;
        var _bot = _boc[_bol].map(function (_bou) { return _bou.c; }), _bov = Math.min.apply(null, _bot), _bow = Math.max.apply(null, _bot);
        _bnk.getRange(_bom, _bov, 1, _bow - _bov + 1).setValues([_boo.slice(_bov - 1, _bow)]);
        _boh.push({ f: _bom, cMin: _bov, cMax: _bow });
    });
    _boh.forEach(function (_box) {
        try {
            registrarCambioCelda_(_bnj, _bnk, _box.f, _box.cMin, 1, _box.cMax - _box.cMin + 1);
        }
        catch (_boy) { }
        var _boz = {};
        _boz[_box.f] = _bog[_box.f];
        try {
            registrarCambiosRetornos_(_bnk, _box.f, _box.cMin, 1, _box.cMax - _box.cMin + 1, { antesFilas: _boz, usuario: _bnu, tipo: "restaurado" });
        }
        catch (_bpa) { }
    });
    var _bpb = _bnx + (_bof.length ? " (en parte)" : "");
    if (_boe > 0)
        _bnl.getRange(_bnp, 11).setValue(codificarTexto_(_bpb));
    return { ok: _boe > 0, restauradas: _boe, conflictos: _bof, marca: _boe > 0 ? _bpb : "" };
}
function localizarFilaCambio_(_bpc, _bpd, _bpe, _bpf) {
    var _bpg = "", _bph = false;
    _bpe.forEach(function (_bpi) { if (_bpi.p)
        _bpg = _bpi.p; if (_bpi.c === COLUMNA_REFERENCIAS)
        _bph = true; });
    if (_bpd >= 2 && _bpd <= _bpf) {
        var _bpj = pedidoDeValor_(_bpc.getRange(_bpd, COLUMNA_REFERENCIAS).getValue());
        if (!_bpg || _bph || _bpj === _bpg)
            return _bpd;
    }
    if (!_bpg || _bph)
        return -1;
    var _bpk = Math.max(2, _bpd - 60), _bpl = Math.min(_bpf, _bpd + 60);
    if (_bpl < _bpk)
        return -1;
    var _bpm = _bpc.getRange(_bpk, 1, _bpl - _bpk + 1, COL_RET_ULTIMA_).getValues(), _bpn = [];
    for (var _bpo = 0; _bpo < _bpm.length; _bpo++) {
        if (pedidoDeValor_(_bpm[_bpo][COLUMNA_REFERENCIAS - 1]) !== _bpg)
            continue;
        var _bpp = _bpe.every(function (_bpq) { var _bpr = valorACodigo_(_bpm[_bpo][_bpq.c - 1]); return _bpr === _bpq.d || (_bpq.d.charAt(0) === "T" && codigoATexto_(_bpr) === _bpq.d.substring(1)); });
        if (_bpp)
            _bpn.push(_bpk + _bpo);
    }
    return _bpn.length === 1 ? _bpn[0] : (_bpn.length > 1 ? -2 : -1);
}
function buscarEnHistorialRetornos_(_bps, _bpt) {
    var _bpu = SpreadsheetApp.getActiveSpreadsheet();
    var _bpv = _bpt !== "articulos" && _bpt !== "pedidos", _bpw = _bpt !== "clientes" && _bpt !== "pedidos", _bpx = [];
    var _bpy = _bpt === "pedidos" || (_bpv && _bpw && _bps.parecePedido);
    leerCambiosRetornos_(_bpu, DIAS_CAMBIOS_RET_).forEach(function (_bpz) {
        if (_bpz.tipo === "restaurado")
            return;
        var _bqa = [];
        if (_bpz.tipo === "fila_eliminada" && _bpz.cambios[0] && _bpz.cambios[0].fila) {
            var _bqb = _bpz.cambios[0].fila;
            if (_bpv)
                _bqa.push({ col: COL_RET_CLIENTE_, valor: codigoATexto_(_bqb[COL_RET_CLIENTE_ - 1]), f: _bpz.cambios[0].f });
            if (_bpw)
                _bqa.push({ col: COL_RET_ARTICULO_, valor: codigoATexto_(_bqb[COL_RET_ARTICULO_ - 1]), f: _bpz.cambios[0].f });
            if (_bpy)
                _bqa.push({ col: COLUMNA_REFERENCIAS, valor: codigoATexto_(_bqb[COLUMNA_REFERENCIAS - 1]), f: _bpz.cambios[0].f });
        }
        else {
            _bpz.cambios.forEach(function (_bqc) {
                if ((_bqc.c === COL_RET_CLIENTE_ && _bpv) || (_bqc.c === COL_RET_ARTICULO_ && _bpw) || (_bqc.c === COLUMNA_REFERENCIAS && _bpy))
                    _bqa.push({ col: _bqc.c, valor: codigoATexto_(_bqc.a), ahora: codigoATexto_(_bqc.d), f: _bqc.f });
            });
        }
        var _bqd = null;
        _bqa.forEach(function (_bqe) {
            if (!_bqe.valor)
                return;
            var _bqf = _bqe.col === COLUMNA_REFERENCIAS ? (pedidoCoincideRetornos_(_bps, _bqe.valor) || { pct: 0 }) : parecidoRetornos_(_bps, _bqe.valor);
            if (_bqf.pct >= UMBRAL_PARECIDO_RETORNOS_ && (!_bqd || _bqf.pct > _bqd.pct))
                _bqd = { pct: _bqf.pct, exacta: !!_bqf.exacta, col: _bqe.col, valor: _bqe.valor, ahora: _bqe.ahora, f: _bqe.f };
        });
        if (_bqd)
            _bpx.push({ id: _bpz.id, t: _bpz.t, usuario: _bpz.usuario ? _bpz.usuario.split("@")[0] : "", tipo: _bpz.tipo, pedido: _bpz.pedido, cliente: _bpz.cliente,
                pct: _bqd.pct, exacta: _bqd.exacta, campo: _bqd.col === COL_RET_CLIENTE_ ? "Cliente" : (_bqd.col === COLUMNA_REFERENCIAS ? "Nº pedido" : "Artículo"), valor: _bqd.valor, ahora: _bqd.ahora || "", fila: _bqd.f, restaurado: _bpz.restaurado });
    });
    _bpx.sort(function (_bqg, _bqh) { return (_bqh.pct - _bqg.pct) || (_bqh.t - _bqg.t); });
    return _bpx.slice(0, 20);
}
function sincronizarRoturasHoy() {
    var _bqi = SpreadsheetApp.getActiveSpreadsheet(), _bqj = new Date();
    var _bqk = TODAS_HOJAS_HISTORIAL_;
    var _bql = getOrCreateSheet(_bqi, HOJA_ROTURAS);
    var _bqm = {};
    var _bqn = _bql.getLastRow();
    if (_bqn >= 2) {
        var _bqo = _bql.getRange(2, 1, _bqn - 1, 5).getValues();
        _bqo.forEach(function (_bqp) {
            if (esMismoDia(new Date(_bqp[0]), _bqj))
                _bqm[_bqp[2] + "_" + _bqp[4]] = true;
        });
    }
    var _bqq = "Usuario anónimo";
    try {
        _bqq = Session.getActiveUser().getEmail() || "Usuario anónimo";
    }
    catch (_bqr) { }
    var _bqs = [];
    _bqi.getSheets().forEach(function (_bqt) {
        var _bqu = _bqt.getName();
        if (_bqk.indexOf(_bqu) > -1)
            return;
        var _bqv = _bqt.getLastRow();
        if (_bqv < 2)
            return;
        var _bqw = _bqt.getRange(2, COLUMNA_AGENCIAS, _bqv - 1, COLUMNA_ROTURA - COLUMNA_AGENCIAS + 1).getDisplayValues();
        for (var _bqx = 0; _bqx < _bqw.length; _bqx++) {
            var _bqy = (_bqw[_bqx][0] || "").toString().trim();
            var _bqz = (_bqw[_bqx][_bqw[_bqx].length - 1] || "").toString().trim().toUpperCase();
            var _bra = _bqx + 2;
            if (_bqz === "SI" && !_bqm[_bqu + "_" + _bra]) {
                _bqs.push([_bqj, _bqq, _bqu, _bqy, _bra]);
            }
        }
    });
    if (_bqs.length > 0)
        _bql.getRange(_bql.getLastRow() + 1, 1, _bqs.length, 5).setValues(_bqs);
    SpreadsheetApp.getUi().alert("✅ Sincronizadas " + _bqs.length + " roturas de hoy que ya estaban marcadas.");
}
function activarDisparadorDiario() {
    if (!soloAdminOAvisar_())
        return;
    eliminarDisparadoresDe_('archivarNotasAntiguas');
    ScriptApp.newTrigger('archivarNotasAntiguas').timeBased().everyDays(1).atHour(1).create();
    try {
        archivarNotasAntiguas();
    }
    catch (_brb) { }
    eliminarDisparadoresDe_('archivarHistorialesAntiguos');
    ScriptApp.newTrigger('archivarHistorialesAntiguos').timeBased().everyDays(1).atHour(1).create();
    try {
        archivarHistorialesAntiguos();
    }
    catch (_brc) { }
    eliminarDisparadoresDe_('revisarRegistrosAutomatico');
    ScriptApp.newTrigger('revisarRegistrosAutomatico').timeBased().everyMinutes(5).create();
    try {
        revisarRegistrosAutomatico();
    }
    catch (_brd) { }
    try {
        ponerCabeceraReferencia_();
    }
    catch (_bre) { }
    try {
        aplicarProteccionesAviso_();
    }
    catch (_brf) { }
    eliminarDisparadoresDe_('revisarSaludDiaria');
    ScriptApp.newTrigger('revisarSaludDiaria').timeBased().everyDays(1).atHour(8).create();
    SpreadsheetApp.getUi().alert("✅ Auto-Archivado configurado a la 1:00 AM (notas e historiales de agencias).\n\n🔄 Revisión automática de lo contado activada: cada 5 minutos se comprueba que todas las agencias, OK, Reclamar y roturas estén registradas, y se añade lo que falte.\n\n🩺 Revisión de salud activada: cada mañana a las 8:00 se comprueba el estado del documento y, si algo no está bien, llega un correo a " + CORREO_AVISOS_SALUD_ + ".");
}
function activarAperturaAutomaticaRufo() {
    if (!soloAdminOAvisar_())
        return;
    eliminarDisparadoresDe_('disparadorAbrirRufo');
    ScriptApp.newTrigger('disparadorAbrirRufo')
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onOpen()
        .create();
    SpreadsheetApp.getUi().alert("✅ Listo. A partir de ahora, Rufo se abrirá solo cada vez que se abra el documento (no hace falta volver a pulsar esto).");
    try {
        abrirMotorLateral("rufo", "🐾 Rufo");
    }
    catch (_brg) { }
}
function disparadorAbrirRufo(_brh) {
    if (!usuarioAutorizado_())
        return;
    try {
        abrirMotorLateral("rufo", "🐾 Rufo");
    }
    catch (_bri) {
        try {
            SpreadsheetApp.getActiveSpreadsheet().toast("No se pudo abrir Rufo: " + _bri.message, "⚠️ Rufo", 8);
        }
        catch (_brj) { }
    }
}
function ponerCabeceraReferencia_() {
    var _brk = SpreadsheetApp.getActiveSpreadsheet();
    [[HOJA_AGENCIAS, 6], [HOJA_ROTURAS, 6], [HOJA_OK, 9], [HOJA_RECLAMAR, 9]].forEach(function (_brl) {
        var _brm = _brk.getSheetByName(_brl[0]);
        if (!_brm || _brm.getMaxColumns() < _brl[1])
            return;
        var _brn = _brm.getRange(1, _brl[1]);
        if (String(_brn.getValue()).trim() === "")
            _brn.setValue("Referencia").setFontWeight("bold");
    });
}
function eliminarDisparadoresDe_(_bro) {
    var _brp = ScriptApp.getProjectTriggers();
    for (var _brq = 0; _brq < _brp.length; _brq++) {
        if (_brp[_brq].getHandlerFunction() === _bro)
            ScriptApp.deleteTrigger(_brp[_brq]);
    }
}
function leerFilasEnRangoFechas_(_brr, _brs, _brt) {
    var _bru = _brr.getLastRow();
    if (_bru < 2)
        return [];
    var _brv = _brr.getLastColumn();
    var _brw = _brr.getRange(2, 1, _bru - 1, 1).getValues();
    var _brx = -1, _bry = -1;
    for (var _brz = 0; _brz < _brw.length; _brz++) {
        var _bsa = new Date(_brw[_brz][0]);
        if (_bsa >= _brs && _bsa <= _brt) {
            if (_brx === -1)
                _brx = _brz;
            _bry = _brz;
        }
    }
    if (_brx === -1)
        return [];
    return _brr.getRange(2 + _brx, 1, _bry - _brx + 1, _brv).getValues();
}
var AVISO_COMPANIAS_ACTIVO_ = false;
var AVISO_COMP_DIAS_ = 7, AVISO_COMP_MIN_ = 20, CLAVE_CACHE_AVISO_COMP_ = "AVISO_COMPANIAS_v2";
var COMPANIAS_AVISO_ = ["VELOX", "PAQNORTE", "CORREOMAX", "RUTASUR", "ATLAS", "PRONTO", "BOLIDO", "FARO"];
function calcularAvisoCompanias_() {
    var _bsb = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AGENCIAS), _bsc = [];
    var _bsd = new Date();
    _bsd.setHours(0, 0, 0, 0);
    var _bse = new Date(_bsd);
    _bse.setDate(_bsd.getDate() - DIAS_CADUCIDAD_HISTORIALES);
    var _bsf = new Date(_bsd);
    _bsf.setHours(23, 59, 59, 999);
    var _bsg = _bsb ? leerFilasEnRangoFechas_(_bsb, _bse, _bsf) : [], _bsh = {}, _bsi = {};
    COMPANIAS_AVISO_.forEach(function (_bsj) { _bsi[_bsj] = {}; });
    for (var _bsk = 0; _bsk < _bsg.length; _bsk++) {
        var _bsl = new Date(_bsg[_bsk][0]);
        if (isNaN(_bsl.getTime()))
            continue;
        var _bsm = (_bsg[_bsk][3] || "").toString().toUpperCase().trim(), _bsn = _bsg[_bsk][2];
        if (_bsm === "" || _bsn === HOJA_RETORNOS)
            continue;
        var _bso = claveFecha_(_bsl), _bsp = _bso + "_" + _bsn + "_" + _bsg[_bsk][4];
        if (_bsh[_bsp])
            continue;
        _bsh[_bsp] = true;
        COMPANIAS_AVISO_.forEach(function (_bsq) { if (_bsm.indexOf(_bsq) !== -1)
            _bsi[_bsq][_bso] = (_bsi[_bsq][_bso] || 0) + 1; });
    }
    COMPANIAS_AVISO_.forEach(function (_bsr) {
        var _bss = null, _bst = 0;
        Object.keys(_bsi[_bsr]).forEach(function (_bsu) { if (_bsi[_bsr][_bsu] >= AVISO_COMP_MIN_ && (!_bss || _bsu > _bss)) {
            _bss = _bsu;
            _bst = _bsi[_bsr][_bsu];
        } });
        var _bsv = null;
        if (_bss) {
            var _bsw = new Date(Number(_bss.substring(0, 4)), Number(_bss.substring(4, 6)) - 1, Number(_bss.substring(6, 8)));
            _bsv = Math.round((_bsd.getTime() - _bsw.getTime()) / 86400000);
        }
        if (_bss && _bsv <= AVISO_COMP_DIAS_)
            return;
        _bsc.push({ c: _bsr, dias: _bsv === null ? -1 : _bsv, ultima: _bss ? _bss.substring(6, 8) + "/" + _bss.substring(4, 6) : "", n: _bst });
    });
    _bsc.sort(function (_bsx, _bsy) { return (_bsy.dias < 0 ? 999 : _bsy.dias) - (_bsx.dias < 0 ? 999 : _bsx.dias); });
    return _bsc;
}
function avisoCompaniasPanel() {
    if (!AVISO_COMPANIAS_ACTIVO_)
        return [];
    if (!usuarioAutorizado_())
        return [];
    var _bsz = null;
    try {
        _bsz = CacheService.getDocumentCache();
        var _bta = _bsz.get(CLAVE_CACHE_AVISO_COMP_);
        if (_bta)
            return JSON.parse(_bta);
    }
    catch (_btb) { }
    var _btc = [];
    try {
        _btc = calcularAvisoCompanias_();
    }
    catch (_btd) {
        return [];
    }
    try {
        if (_bsz)
            _bsz.put(CLAVE_CACHE_AVISO_COMP_, JSON.stringify(_btc), 3600);
    }
    catch (_bte) { }
    return _btc;
}
function obtenerDatosCalendario(_btf) {
    var _btg = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AGENCIAS);
    var _bth = new Date(), _bti = _bth.getDay() || 7;
    var _btj = new Date(_bth);
    _btj.setDate(_bth.getDate() - _bti + 1 + (_btf * 7));
    _btj.setHours(0, 0, 0, 0);
    var _btk = ["DOM", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁB.", "DOM."];
    var _btl = ["VELOX", "PAQNORTE", "CORREOMAX", "RUTASUR", "ATLAS", "PRONTO", "BOLIDO", "FARO"], _btm = [];
    for (var _btn = 0; _btn < 7; _btn++) {
        var _bto = new Date(_btj);
        _bto.setDate(_btj.getDate() + _btn);
        var _btp = { nombre: _btk[_bto.getDay() === 0 ? 7 : _bto.getDay()] + " " + ('0' + _bto.getDate()).slice(-2), esHoy: esMismoDia(_bto, _bth), total: 0, agencias: {}, agenciasPendiente: {}, agenciasRotos: {}, agenciasRetornos: {} };
        _btl.forEach(function (_btq) { _btp.agencias[_btq] = 0; _btp.agenciasPendiente[_btq] = 0; _btp.agenciasRotos[_btq] = 0; _btp.agenciasRetornos[_btq] = 0; });
        _btm.push({ fObj: _bto, datos: _btp });
    }
    var _btr = _btm[0].fObj;
    var _bts = new Date(_btm[6].fObj);
    _bts.setHours(23, 59, 59, 999);
    var _btt = {};
    if (_btg) {
        var _btu = leerFilasEnRangoFechas_(_btg, _btr, _bts);
        for (var _btv = 0; _btv < _btu.length; _btv++) {
            for (var _btw = 0; _btw < 7; _btw++) {
                if (esMismoDia(new Date(_btu[_btv][0]), _btm[_btw].fObj)) {
                    _btt[_btw + "_" + _btu[_btv][2] + "_" + _btu[_btv][4]] = { idx: _btw, ag: (_btu[_btv][3] || "").toString().toUpperCase().trim(), origen: _btu[_btv][2], fila: _btu[_btv][4] };
                    break;
                }
            }
        }
        for (var _btx in _btt) {
            if (_btt[_btx].ag !== "" && _btt[_btx].origen !== HOJA_RETORNOS) {
                _btm[_btt[_btx].idx].datos.total++;
                _btl.forEach(function (_bty) { if (_btt[_btx].ag.indexOf(_bty) !== -1)
                    _btm[_btt[_btx].idx].datos.agencias[_bty]++; });
            }
        }
    }
    var _btz = {};
    for (var _bua in _btt) {
        var _bub = _btt[_bua];
        if (_bub.ag === "" || _bub.origen === HOJA_RETORNOS)
            continue;
        if (!_btz[_bub.origen])
            _btz[_bub.origen] = [];
        _btz[_bub.origen].push(_bub);
    }
    for (var _buc in _btz) {
        var _bud = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_buc);
        if (!_bud)
            continue;
        var _bue = _bud.getLastRow();
        if (_bue < 2)
            continue;
        var _buf = _bud.getRange(2, COLUMNA_NOTA_L, _bue - 1, 1).getValues();
        _btz[_buc].forEach(function (_bug) {
            var _buh = _bug.fila - 2;
            var _bui = (_buh >= 0 && _buh < _buf.length) ? (_buf[_buh][0] || "").toString().trim() : "";
            if (_bui === "") {
                _btl.forEach(function (_buj) { if (_bug.ag.indexOf(_buj) !== -1)
                    _btm[_bug.idx].datos.agenciasPendiente[_buj]++; });
            }
        });
    }
    var _buk = [];
    for (var _bul in _btt) {
        var _bum = _btt[_bul];
        if (_bum.ag === "" || _bum.origen !== HOJA_RETORNOS)
            continue;
        _buk.push(_bum);
    }
    if (_buk.length > 0) {
        var _bun = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_RETORNOS);
        if (_bun) {
            var _buo = _bun.getLastRow();
            if (_buo >= 2) {
                var _bup = _bun.getRange(2, COLUMNA_RETORNOS_FECHA, _buo - 1, 1).getValues();
                _buk.forEach(function (_buq) {
                    var _bur = _buq.fila - 2;
                    var _bus = (_bur >= 0 && _bur < _bup.length) ? (_bup[_bur][0] || "").toString().trim() : "";
                    if (_bus !== "") {
                        _btl.forEach(function (_but) { if (_buq.ag.indexOf(_but) !== -1)
                            _btm[_buq.idx].datos.agenciasRetornos[_but]++; });
                    }
                });
            }
        }
    }
    var _buu = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_ROTURAS);
    if (_buu && _buu.getLastRow() >= 2) {
        var _buv = leerFilasEnRangoFechas_(_buu, _btr, _bts);
        var _buw = {};
        for (var _bux = 0; _bux < _buv.length; _bux++) {
            var _buy = (_buv[_bux][3] || "").toString().toUpperCase().trim();
            if (_buy === "")
                continue;
            for (var _buz = 0; _buz < 7; _buz++) {
                if (esMismoDia(new Date(_buv[_bux][0]), _btm[_buz].fObj)) {
                    if (_buv[_bux][2] !== "" && Number(_buv[_bux][4]) >= 2) {
                        var _bva = _buz + "#" + _buv[_bux][2] + "#" + Number(_buv[_bux][4]);
                        if (_buw[_bva])
                            break;
                        _buw[_bva] = true;
                    }
                    _btl.forEach(function (_bvb) { if (_buy.indexOf(_bvb) !== -1)
                        _btm[_buz].datos.agenciasRotos[_bvb]++; });
                    break;
                }
            }
        }
    }
    var _bvc = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
    function _bvd(_bve) { return ('0' + _bve.getDate()).slice(-2) + '/' + ('0' + (_bve.getMonth() + 1)).slice(-2) + '/' + _bve.getFullYear(); }
    var _bvf = _bvd(_btm[0].fObj) + ' – ' + _bvd(_btm[6].fObj);
    var _bvg = '';
    try {
        _bvg = Session.getActiveUser().getEmail() || '';
    }
    catch (_bvh) {
        _bvg = '';
    }
    var _bvi = null;
    try {
        _bvi = estadoEnviosChecklist_(_btm.map(function (_bvj) { return _bvj.fObj; }), leerEnviosChecklist_());
    }
    catch (_bvk) {
        _bvi = null;
    }
    return { envios: _bvi, mes: _bvc[_btj.getMonth()] + " " + _btj.getFullYear(), dias: _btm.map(function (_bvl) { return _bvl.datos; }), compañias: _btl, rangoFechas: _bvf, generadoPor: _bvg, correoConfigurado: !!CORREO_JEFATURA, correoEsPrueba: CORREO_MODO_PRUEBA, correoDestino: CORREO_JEFATURA, correoCC: ccChecklist_() };
}
var CORREO_JEFATURA = "luis.prado@ejemplo.com";
var CORREO_CC_ENCARGADO = "jefatura@ejemplo.com";
var CORREO_CC_EXTRA_CHECKLIST_ = "ana.soler@ejemplo.com";
function ccChecklist_() { return [CORREO_CC_ENCARGADO, CORREO_CC_EXTRA_CHECKLIST_].filter(function (_bvm) { return String(_bvm || "").trim() !== ""; }).join(","); }
var CORREO_MODO_PRUEBA = false;
function construirTextoChecklist_(_bvn, _bvo) {
    var _bvp = [];
    var _bvq = _bvo === 'dia' ? (_bvn.diaElegidoTexto ? _bvn.diaElegidoTexto.toUpperCase() + (_bvn.reenvio ? ' (ACTUALIZADO EL ' + _bvn.reenvio.hoy + ')' : ' (ENVIADO CON RETRASO EL ' + _bvn.envioTardeTexto + ')') : 'HOY') : 'SEMANA (' + _bvn.rangoFechas + ')';
    _bvp.push('CHECKLIST DE AGENCIAS — ' + _bvq + (_bvn.reenvio && !(_bvo === 'dia' && _bvn.diaElegidoTexto) ? ' — ACTUALIZADO' : ''));
    if (_bvn.reenvio)
        _bvp.push('Versión actualizada: sustituye a la enviada el ' + _bvn.reenvio.texto + '.');
    var _bvr = new Date();
    var _bvs = ('0' + _bvr.getDate()).slice(-2) + '/' + ('0' + (_bvr.getMonth() + 1)).slice(-2) + '/' + _bvr.getFullYear() +
        ' ' + ('0' + _bvr.getHours()).slice(-2) + ':' + ('0' + _bvr.getMinutes()).slice(-2);
    _bvp.push('Generado: ' + _bvs + (_bvn.generadoPor ? ' · ' + _bvn.generadoPor : ''));
    _bvp.push('');
    var _bvt = diasDelChecklist_(_bvn, _bvo);
    var _bvu = 0, _bvv = 0, _bvw = 0, _bvx = 0;
    if (_bvo === 'dia') {
        _bvt.forEach(function (_bvy) {
            _bvp.push('— ' + _bvy.nombre + (_bvy.esHoy ? ' (HOY)' : '') + ' —');
            var _bvz = false;
            _bvn.compañias.forEach(function (_bwa) {
                var _bwb = _bvy.agencias[_bwa] || 0;
                var _bwc = (_bvy.agenciasRotos && _bvy.agenciasRotos[_bwa]) || 0;
                var _bwd = (_bvy.agenciasPendiente && _bvy.agenciasPendiente[_bwa]) || 0;
                var _bwe = (_bvy.agenciasRetornos && _bvy.agenciasRetornos[_bwa]) || 0;
                if (_bwb === 0 && _bwe === 0)
                    return;
                _bvz = true;
                _bvu += _bwb;
                if (_bwb > 0)
                    _bvv += _bwd;
                _bvx += _bwe;
                var _bwf = _bwb + ' traídas';
                if (_bwb > 0)
                    _bwf += ', ' + _bwd + ' pendiente(s) por abrir';
                if (_bwe > 0)
                    _bwf += ', ' + _bwe + ' retorno(s)';
                _bvp.push('☐ ' + _bwa + ' — ' + _bwf);
            });
            if (!_bvz)
                _bvp.push('  (sin movimiento)');
            _bvp.push('');
        });
    }
    else {
        var _bwg = false;
        _bvn.compañias.forEach(function (_bwh) {
            var _bwi = 0, _bwj = 0, _bwk = 0, _bwl = 0;
            _bvt.forEach(function (_bwm) {
                _bwi += _bwm.agencias[_bwh] || 0;
                _bwj += (_bwm.agenciasRotos && _bwm.agenciasRotos[_bwh]) || 0;
                _bwk += (_bwm.agenciasPendiente && _bwm.agenciasPendiente[_bwh]) || 0;
                _bwl += (_bwm.agenciasRetornos && _bwm.agenciasRetornos[_bwh]) || 0;
            });
            if (_bwi === 0 && _bwj === 0 && _bwl === 0)
                return;
            _bwg = true;
            _bvu += _bwi;
            _bvw += _bwj;
            _bvv += _bwk;
            _bvx += _bwl;
            var _bwn = _bwi + ' traídas';
            if (_bwj > 0)
                _bwn += ', ' + _bwj + ' rota(s)';
            if (_bwi > 0)
                _bwn += ', ' + _bwk + ' pendiente(s) por abrir';
            if (_bwl > 0)
                _bwn += ', ' + _bwl + ' retorno(s)';
            _bvp.push('☐ ' + _bwh + ' — ' + _bwn);
        });
        if (!_bwg)
            _bvp.push('(sin movimiento en toda la semana)');
        _bvp.push('');
    }
    if (_bvo === 'dia')
        _bvp.push('TOTAL: ' + _bvu + ' traídas · ' + _bvv + ' pendiente(s) por abrir · ' + _bvx + ' retorno(s)');
    else
        _bvp.push('TOTAL: ' + _bvu + ' traídas · ' + _bvw + ' rota(s) · ' + _bvv + ' pendiente(s) por abrir · ' + _bvx + ' retorno(s)');
    return _bvp.join('\n');
}
function construirHtmlChecklistCorreo_(_bwo, _bwp, _bwq, _bwr) {
    var _bws = "font-family:Arial, Helvetica, sans-serif;";
    function _bwt(_bwu) {
        if (_bwr) {
            var _bwv = { "VELOX": "#2F6FE4", "PAQNORTE": "#0F8A7E", "CORREOMAX": "#C99700", "RUTASUR": "#D85A30", "ATLAS": "#6D4BC9", "PRONTO": "#2E9E44", "BOLIDO": "#C62E3B", "FARO": "#26354A" }[_bwu] || "#64748b";
            return '<table cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;"><tr><td bgcolor="' + _bwv + '" style="background:' + _bwv + ';color:#ffffff;font-family:Arial, Helvetica, sans-serif;font-size:11px;font-weight:bold;padding:3px 8px;white-space:nowrap;">' + _bwu + '</td></tr></table>';
        }
        var _bww = _bws + 'display:inline-block;width:56px;box-sizing:border-box;text-align:center;vertical-align:middle;border-radius:6px;background:#ffffff;border:1px solid #e2e8f0;padding:3px 3px 2px;';
        var _bwx = 'width:100%;height:22px;line-height:22px;overflow:hidden;border-radius:4px;';
        var _bwy = 'font-size:7.5px;font-weight:700;color:#64748b;letter-spacing:.02em;margin-top:2px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;';
        var _bwz = '<div style="' + _bwy + '">' + _bwu + '</div>';
        if (AG_LOGOS_B64_[_bwu] && !_bwr) {
            var _bxa = _bwq ? 'cid:ag_logo_' + _bwu : 'data:image/png;base64,' + AG_LOGOS_B64_[_bwu];
            return '<div style="' + _bww + '"><div style="' + _bwx + '">' +
                '<img src="' + _bxa + '" alt="' + _bwu + '" style="max-width:48px;max-height:18px;width:auto;height:auto;vertical-align:middle;">' + '</div>' + _bwz + '</div>';
        }
        var _bxb = { "VELOX": { c: "#2F6FE4", t: "VELOX" }, "PAQNORTE": { c: "#0F8A7E", t: "PAQN" }, "CORREOMAX": { c: "#C99700", t: "CMAX" }, "RUTASUR": { c: "#D85A30", t: "RSUR" }, "ATLAS": { c: "#6D4BC9", t: "ATLAS" }, "PRONTO": { c: "#2E9E44", t: "PRON" }, "BOLIDO": { c: "#C62E3B", t: "BOLI" }, "FARO": { c: "#26354A", t: "FARO" } };
        var _bxc = _bxb[_bwu] || { c: "#64748b", t: _bwu.substring(0, 3) };
        return '<div style="' + _bww + '"><div style="' + _bwx + 'color:#fff;font-size:11px;font-weight:800;letter-spacing:.03em;background:' + _bxc.c + ';">' + _bxc.t + '</div>' + _bwz + '</div>';
    }
    var _bxd = diasDelChecklist_(_bwo, _bwp);
    var _bxe = _bwp === 'dia' ? 'Checklist Diario de Agencias' : 'Resumen Semanal de Agencias';
    var _bxf = _bwp === 'dia' ? (_bwo.diaElegidoTexto ? 'Traídas y pendientes del ' + _bwo.diaElegidoTexto.toLowerCase() : 'Traídas y pendientes de hoy') : ('Por agencia · Semana del ' + (_bwo.rangoFechas || '—'));
    var _bxg = new Date();
    var _bxh = _bxg.toLocaleDateString('es-ES') + ' ' + _bxg.toLocaleTimeString('es-ES', { hour: '2-digit', minute: '2-digit' });
    var _bxi = _bws + "padding:5px 8px;border-bottom:1px solid #eee;font-size:12.5px;vertical-align:middle;";
    var _bxj = _bws + "background:#f1f1f1;border-bottom:1px solid #ccc;padding:5px 8px;font-size:10px;text-transform:uppercase;color:#444;text-align:left;";
    var _bxk = _bws + "padding:6px 8px;font-weight:bold;border-top:1.5px solid #333;background:#f8f8f8;font-size:12.5px;";
    var _bxl = _bwp !== 'dia' ? '' : _bxd.map(function (_bxm) {
        var _bxn = '', _bxo = 0, _bxp = 0;
        _bwo.compañias.forEach(function (_bxq) {
            var _bxr = _bxm.agencias[_bxq] || 0;
            var _bxs = (_bxm.agenciasPendiente && _bxm.agenciasPendiente[_bxq]) || 0;
            var _bxt = (_bxm.agenciasRetornos && _bxm.agenciasRetornos[_bxq]) || 0;
            _bxo += _bxs;
            _bxp += _bxt;
            _bxn += '<tr>' +
                '<td style="' + _bxi + '">' + _bwt(_bxq) + '</td>' +
                '<td style="' + _bxi + '">' + _bxr + '</td>' +
                '<td style="' + _bxi + '">' + (_bxr > 0 ? _bxs : '—') + '</td>' +
                '<td style="' + _bxi + '">' + (_bxt > 0 ? _bxt : '—') + '</td>' +
                '</tr>';
        });
        _bxn += '<tr>' +
            '<td style="' + _bxk + '">TOTAL</td>' +
            '<td style="' + _bxk + '">' + _bxm.total + '</td>' +
            '<td style="' + _bxk + '">' + _bxo + '</td>' +
            '<td style="' + _bxk + '">' + _bxp + '</td>' +
            '</tr>';
        var _bxu = _bxm.esHoy ? '#E07A1F' : '#1a1a1a';
        var _bxv = _bxm.esHoy ? ' <span style="background:#fff;color:#E07A1F;font-size:9.5px;font-weight:bold;padding:1.5px 5px;border-radius:3px;margin-left:5px;">HOY</span>' : '';
        return '<table cellpadding="0" cellspacing="0" style="' + _bws + 'width:100%;max-width:560px;border-collapse:collapse;margin-bottom:14px;">' +
            '<tr><th colspan="4" style="' + _bws + 'background:' + _bxu + ';color:#fff;padding:7px 8px;font-size:13px;text-align:left;">' + _bxm.nombre + _bxv + '</th></tr>' +
            '<tr>' +
            '<th style="' + _bxj + '">Agencia</th>' +
            '<th style="' + _bxj + '">Traídas</th>' +
            '<th style="' + _bxj + '">Pendientes por abrir</th>' +
            '<th style="' + _bxj + '">Retornos</th>' +
            '</tr>' +
            _bxn +
            '</table>';
    }).join('');
    var _bxw = '';
    if (_bwp !== 'dia') {
        var _bxx = {};
        _bwo.compañias.forEach(function (_bxy) { _bxx[_bxy] = { v: 0, rot: 0, pend: 0, retornos: 0 }; });
        _bxd.forEach(function (_bxz) {
            _bwo.compañias.forEach(function (_bya) {
                _bxx[_bya].v += _bxz.agencias[_bya] || 0;
                _bxx[_bya].rot += (_bxz.agenciasRotos && _bxz.agenciasRotos[_bya]) || 0;
                _bxx[_bya].pend += (_bxz.agenciasPendiente && _bxz.agenciasPendiente[_bya]) || 0;
                _bxx[_bya].retornos += (_bxz.agenciasRetornos && _bxz.agenciasRetornos[_bya]) || 0;
            });
        });
        var _byb = '', _byc = 0, _byd = 0, _bye = 0, _byf = 0;
        _bwo.compañias.forEach(function (_byg) {
            var _byh = _bxx[_byg];
            _byc += _byh.v;
            _byd += _byh.rot;
            _bye += _byh.pend;
            _byf += _byh.retornos;
            var _byi = "—", _byj = "text-align:right;";
            if (_byh.rot > 0 && _byh.v > 0) {
                var _byk = Math.round((_byh.rot / _byh.v) * 1000) / 10;
                _byi = (_byk % 1 === 0 ? _byk.toFixed(0) : _byk.toFixed(1)) + "%";
                _byj = "text-align:right;color:#E07A1F;font-weight:bold;";
            }
            _byb += '<tr>' +
                '<td style="' + _bxi + '">' + _bwt(_byg) + '</td>' +
                '<td style="' + _bxi + '">' + _byh.v + '</td>' +
                '<td style="' + _bxi + '">' + (_byh.rot > 0 ? _byh.rot : '—') + '</td>' +
                '<td style="' + _bxi + _byj + '">' + _byi + '</td>' +
                '<td style="' + _bxi + '">' + (_byh.v > 0 ? _byh.pend : '—') + '</td>' +
                '<td style="' + _bxi + '">' + (_byh.retornos > 0 ? _byh.retornos : '—') + '</td>' +
                '</tr>';
        });
        var _byl = "—", _bym = "text-align:right;";
        if (_byd > 0 && _byc > 0) {
            var _byn = Math.round((_byd / _byc) * 1000) / 10;
            _byl = (_byn % 1 === 0 ? _byn.toFixed(0) : _byn.toFixed(1)) + "%";
            _bym = "text-align:right;color:#E07A1F;font-weight:bold;";
        }
        _byb += '<tr>' +
            '<td style="' + _bxk + '">TOTAL SEMANA</td>' +
            '<td style="' + _bxk + '">' + _byc + '</td>' +
            '<td style="' + _bxk + '">' + _byd + '</td>' +
            '<td style="' + _bxk + _bym + '">' + _byl + '</td>' +
            '<td style="' + _bxk + '">' + _bye + '</td>' +
            '<td style="' + _bxk + '">' + _byf + '</td>' +
            '</tr>';
        _bxw = '<table cellpadding="0" cellspacing="0" style="' + _bws + 'width:100%;max-width:560px;border-collapse:collapse;margin-bottom:14px;">' +
            '<tr><th colspan="6" style="' + _bws + 'background:#E07A1F;color:#fff;padding:7px 8px;font-size:13px;text-align:left;">📊 Resumen de la semana</th></tr>' +
            '<tr>' +
            '<th style="' + _bxj + '">Agencia</th>' +
            '<th style="' + _bxj + '">Traídas</th>' +
            '<th style="' + _bxj + '">Roto</th>' +
            '<th style="' + _bxj + 'text-align:right;">% Rotura</th>' +
            '<th style="' + _bxj + '">Pendientes por abrir</th>' +
            '<th style="' + _bxj + '">Retornos</th>' +
            '</tr>' +
            _byb +
            '</table>';
    }
    return '<div style="' + _bws + 'max-width:640px;margin:0 auto;padding:18px;background:#ffffff;color:#111;">' +
        '<table cellpadding="0" cellspacing="0" style="width:100%;border-collapse:collapse;margin-bottom:4px;">' +
        '<tr>' +
        '<td style="' + _bws + 'vertical-align:top;width:34%;">' +
        '<div style="font-family:Georgia, \'Times New Roman\', serif;font-size:26px;font-weight:bold;letter-spacing:1px;color:#111;">RAS<span style="color:#E07A1F;">TRO</span></div>' +
        '<div style="' + _bws + 'font-size:10.5px;text-transform:uppercase;letter-spacing:1.5px;color:#666;margin-top:2px;">Departamento de Devoluciones</div>' +
        '</td>' +
        '<td style="' + _bws + 'vertical-align:top;width:32%;text-align:center;">' +
        '<div style="' + _bws + 'font-size:18px;font-weight:bold;color:#111;margin-bottom:3px;">' + _bxe + '</div>' +
        '<div style="' + _bws + 'font-size:12px;color:#555;">' + _bxf + '</div>' +
        '</td>' +
        '<td style="' + _bws + 'vertical-align:top;width:34%;text-align:right;font-size:11px;color:#444;line-height:1.7;">' +
        '<div><strong>' + (_bwp === 'dia' && _bwo.fechaDiaTexto ? 'Día' : 'Periodo') + ':</strong> ' + (_bwp === 'dia' && _bwo.fechaDiaTexto ? _bwo.fechaDiaTexto : (_bwo.rangoFechas || '—')) + '</div>' +
        '<div><strong>Generado:</strong> ' + _bxh + '</div>' +
        '<div><strong>Por:</strong> ' + (_bwo.generadoPor || '—') + '</div>' +
        '</td>' +
        '</tr>' +
        '</table>' +
        '<div style="border-bottom:3px solid #E07A1F;margin-bottom:16px;"></div>' +
        (_bwo.reenvio ? '<div style="' + _bws + 'background:#eff6ff;border:1px solid #93c5fd;color:#1e40af;border-radius:6px;padding:8px 10px;font-size:12px;margin-bottom:14px;">🔄 <strong>Versión actualizada' + (_bwp === 'dia' ? ' del checklist ' + (_bwo.diaElegidoTexto ? 'del ' + _bwo.diaElegidoTexto.toLowerCase() : 'de hoy') : ' del resumen de la semana') + '.</strong> Sustituye al enviado el ' + _bwo.reenvio.texto + '.</div>' :
            (_bwp === 'dia' && _bwo.diaElegidoTexto ? '<div style="' + _bws + 'background:#fff7ed;border:1px solid #fdba74;color:#9a3412;border-radius:6px;padding:8px 10px;font-size:12px;margin-bottom:14px;">⏰ <strong>Checklist del ' + _bwo.diaElegidoTexto.toLowerCase() + ', enviado con retraso el ' + _bwo.envioTardeTexto + '.</strong></div>' : '')) +
        _bxl +
        _bxw +
        '<div style="margin-top:14px;padding-top:8px;border-top:1px solid #ccc;font-size:10px;color:#999;text-align:center;">Documento de uso interno · Sistema de Auditoría, Agencias y Tablón de Notas · Rastro</div>' +
        '</div>';
}
function datosParaChecklist(_byo, _byp, _byq, _byr) {
    if (_byr !== true) {
        try {
            revisarRegistros_(15000);
        }
        catch (_bys) { }
    }
    var _byt = obtenerDatosCalendario(_byo);
    var _byu = _byp === 'dia' ? elegirDiaChecklist_(_byt, _byo, _byq) : null;
    if (_byu)
        _byt.errorDia = _byu;
    if (!_byu)
        marcarReenvioChecklist_(_byt, _byp, _byo, _byq);
    if (!_byu)
        fechaDiaChecklist_(_byt, _byp, _byq);
    _byt.checklistAsunto = asuntoChecklist_(_byt, _byp);
    _byt.checklistHtml = construirHtmlChecklistCorreo_(_byt, _byp);
    _byt.checklistHtmlOutlook = '<table width="600" cellpadding="0" cellspacing="0" border="0" style="width:600px;border-collapse:collapse;"><tr><td style="padding:0;">' + construirHtmlChecklistCorreo_(_byt, _byp, false, true) + '</td></tr></table>';
    _byt.checklistTexto = construirTextoChecklist_(_byt, _byp);
    return _byt;
}
function revisarParaChecklist() {
    var _byv = null;
    try {
        _byv = revisarRegistros_(15000);
    }
    catch (_byw) {
        return { cambios: false, error: true };
    }
    return { cambios: !!(_byv && _byv.total > 0), ocupado: !!(_byv && _byv.ocupado) };
}
function construirInlineImagesLogos_() {
    var _byx = {};
    Object.keys(AG_LOGOS_B64_).forEach(function (_byy) {
        _byx['ag_logo_' + _byy] = Utilities.newBlob(Utilities.base64Decode(AG_LOGOS_B64_[_byy]), 'image/png', _byy + '.png');
    });
    return _byx;
}
var PROP_ENVIOS_CHECKLIST_ = "CHK_ENVIOS";
function leerEnviosChecklist_() {
    var _byz = PropertiesService.getDocumentProperties(), _bza = null;
    try {
        _bza = JSON.parse(_byz.getProperty(PROP_ENVIOS_CHECKLIST_) || "null");
    }
    catch (_bzb) {
        _bza = null;
    }
    if (!_bza || typeof _bza !== "object" || !_bza.desde) {
        _bza = { desde: claveFecha_(new Date()), e: {} };
        try {
            _byz.setProperty(PROP_ENVIOS_CHECKLIST_, JSON.stringify(_bza));
        }
        catch (_bzc) { }
    }
    if (!_bza.e || typeof _bza.e !== "object")
        _bza.e = {};
    return _bza;
}
function fechasSemanaChecklist_(_bzd) {
    var _bze = new Date(), _bzf = _bze.getDay() || 7, _bzg = new Date(_bze), _bzh = [];
    _bzg.setDate(_bze.getDate() - _bzf + 1 + ((Number(_bzd) || 0) * 7));
    _bzg.setHours(0, 0, 0, 0);
    for (var _bzi = 0; _bzi < 7; _bzi++) {
        var _bzj = new Date(_bzg);
        _bzj.setDate(_bzg.getDate() + _bzi);
        _bzh.push(_bzj);
    }
    return _bzh;
}
function anotarEnvioChecklist_(_bzk, _bzl, _bzm, _bzn) {
    var _bzo = LockService.getDocumentLock(), _bzp = false;
    try {
        _bzp = _bzo.tryLock(5000);
    }
    catch (_bzq) { }
    try {
        var _bzr = leerEnviosChecklist_(), _bzs = new Date(), _bzt = "";
        try {
            _bzt = String(Session.getActiveUser().getEmail() || "").split("@")[0];
        }
        catch (_bzu) { }
        var _bzv = _bzk === "dia" ? "d" + (_bzm || claveFecha_(_bzs)) : "s" + claveFecha_(_bzl[0]);
        var _bzw = _bzr.e[_bzv];
        if (_bzw && _bzw.t) {
            _bzw.a = _bzs.getTime();
            _bzw.au = _bzt;
            _bzw.r = (Number(_bzw.r) || 0) + 1;
        }
        else {
            _bzr.e[_bzv] = { t: _bzs.getTime(), u: _bzt };
            if (_bzk === "dia" ? (_bzm && _bzm !== claveFecha_(_bzs)) : claveFecha_(lunesSiguienteChecklist_(_bzl)) < claveFecha_(_bzs))
                _bzr.e[_bzv].tarde = 1;
            if (_bzn)
                _bzr.e[_bzv].o = 1;
        }
        var _bzx = claveFecha_(new Date(_bzs.getTime() - 70 * 86400000));
        Object.keys(_bzr.e).forEach(function (_bzy) { if (_bzy.substring(1) < _bzx)
            delete _bzr.e[_bzy]; });
        PropertiesService.getDocumentProperties().setProperty(PROP_ENVIOS_CHECKLIST_, JSON.stringify(_bzr));
        return estadoEnviosChecklist_(_bzl, _bzr);
    }
    finally {
        if (_bzp) {
            try {
                _bzo.releaseLock();
            }
            catch (_bzz) { }
        }
    }
}
var SABADO_CUENTA_DESDE_CHK_ = "20261003";
function lunesSiguienteChecklist_(_caa) {
    var _cab = new Date(_caa[6]);
    _cab.setDate(_cab.getDate() + 1);
    return _cab;
}
function estadoSemanalChecklist_(_cac, _cad, _cae) {
    var _caf = new Date(_cac);
    _caf.setDate(_caf.getDate() + 6);
    var _cag = new Date(_cac);
    _cag.setDate(_cag.getDate() + 7);
    var _cah = _cad.e["s" + claveFecha_(_cac)], _cai = claveFecha_(_caf), _caj = claveFecha_(_cag), _cak;
    if (_cah)
        _cak = "si";
    else if (_cai < _cad.desde)
        _cak = "nada";
    else if (_cae > _caj)
        _cak = "no";
    else if (_cae === _caj)
        _cak = "hoy";
    else
        _cak = "pronto";
    return { estado: _cak, t: _cah ? _cah.t : 0, u: _cah ? _cah.u : "", tarde: !!(_cah && _cah.tarde), o: !!(_cah && _cah.o), k: claveFecha_(_cac), a: _cah && _cah.a ? _cah.a : 0, au: _cah && _cah.au ? _cah.au : "" };
}
function estadoEnviosChecklist_(_cal, _cam) {
    var _can = claveFecha_(new Date()), _cao = [];
    _cal.forEach(function (_cap, _caq) {
        var _car = claveFecha_(_cap), _cas = _cam.e["d" + _car], _cat;
        if (_cas)
            _cat = "si";
        else if (_caq >= 6 || _car > _can || _car < _cam.desde || (_caq === 5 && _car < SABADO_CUENTA_DESDE_CHK_))
            _cat = "nada";
        else
            _cat = (_car === _can) ? "hoy" : "no";
        _cao.push({ estado: _cat, t: _cas ? _cas.t : 0, u: _cas ? _cas.u : "", tarde: !!(_cas && _cas.tarde), o: !!(_cas && _cas.o), k: _car, a: _cas && _cas.a ? _cas.a : 0, au: _cas && _cas.au ? _cas.au : "" });
    });
    var _cau = new Date(_cal[0]);
    _cau.setDate(_cau.getDate() - 7);
    return { dias: _cao, semana: estadoSemanalChecklist_(_cal[0], _cam, _can), anterior: estadoSemanalChecklist_(_cau, _cam, _can) };
}
var DIAS_LARGOS_CHK_ = ["Domingo", "Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"];
function elegirDiaChecklist_(_cav, _caw, _cax) {
    if (!_cax)
        return null;
    _cax = String(_cax);
    var _cay = fechasSemanaChecklist_(_caw), _caz = claveFecha_(new Date());
    if (_cax === _caz)
        return null;
    if (_cax > _caz)
        return "Ese día todavía no ha llegado.";
    for (var _cba = 0; _cba < _cay.length; _cba++) {
        if (claveFecha_(_cay[_cba]) !== _cax)
            continue;
        var _cbb = _cay[_cba], _cbc = new Date();
        _cav.dias[_cba].elegido = true;
        _cav.diaElegidoTexto = DIAS_LARGOS_CHK_[_cbb.getDay()] + " " + ('0' + _cbb.getDate()).slice(-2) + "/" + ('0' + (_cbb.getMonth() + 1)).slice(-2) + "/" + _cbb.getFullYear();
        _cav.envioTardeTexto = ('0' + _cbc.getDate()).slice(-2) + "/" + ('0' + (_cbc.getMonth() + 1)).slice(-2);
        return null;
    }
    return "Ese día no es de la semana que se está viendo.";
}
function marcarReenvioChecklist_(_cbd, _cbe, _cbf, _cbg) {
    try {
        var _cbh = leerEnviosChecklist_(), _cbi = new Date();
        var _cbj = _cbe === 'dia' ? "d" + (_cbd.diaElegidoTexto && _cbg ? String(_cbg) : claveFecha_(_cbi)) : "s" + claveFecha_(fechasSemanaChecklist_(_cbf)[0]);
        var _cbk = _cbh.e[_cbj];
        if (!_cbk || !_cbk.t)
            return;
        var _cbl = new Date(_cbk.t);
        _cbd.reenvio = { t: _cbk.t, u: _cbk.u || "", texto: ('0' + _cbl.getDate()).slice(-2) + "/" + ('0' + (_cbl.getMonth() + 1)).slice(-2) + " a las " + ('0' + _cbl.getHours()).slice(-2) + ":" + ('0' + _cbl.getMinutes()).slice(-2), hoy: ('0' + _cbi.getDate()).slice(-2) + "/" + ('0' + (_cbi.getMonth() + 1)).slice(-2) };
    }
    catch (_cbm) { }
}
function fechaDiaChecklist_(_cbn, _cbo, _cbp) {
    if (_cbo !== 'dia')
        return;
    var _cbq = new Date(), _cbr = String(_cbp || "");
    if (_cbn.diaElegidoTexto && /^\d{8}$/.test(_cbr))
        _cbq = new Date(Number(_cbr.substring(0, 4)), Number(_cbr.substring(4, 6)) - 1, Number(_cbr.substring(6, 8)));
    _cbn.fechaDiaTexto = ('0' + _cbq.getDate()).slice(-2) + '/' + ('0' + (_cbq.getMonth() + 1)).slice(-2) + '/' + _cbq.getFullYear();
}
function asuntoChecklist_(_cbs, _cbt) {
    var _cbu = _cbs.reenvio;
    var _cbv = _cbt !== 'dia' ? 'Semana' : (_cbs.diaElegidoTexto ? _cbs.diaElegidoTexto + (_cbu ? ' (actualizado el ' + _cbu.hoy + ')' : ' (enviado con retraso el ' + _cbs.envioTardeTexto + ')') : 'Hoy');
    return 'Checklist de Agencias — ' + _cbv + (_cbt === 'dia' && _cbs.diaElegidoTexto ? '' : ' (' + (_cbt === 'dia' && _cbs.fechaDiaTexto ? _cbs.fechaDiaTexto : _cbs.rangoFechas) + ')' + (_cbu ? ' — ACTUALIZADO' : ''));
}
function diasDelChecklist_(_cbw, _cbx) {
    if (_cbx !== 'dia')
        return _cbw.dias;
    return _cbw.dias.filter(function (_cby) { return _cbw.diaElegidoTexto ? _cby.elegido === true : _cby.esHoy; });
}
function marcarChecklistPreparado(_cbz, _cca, _ccb) {
    if (!usuarioAutorizado_())
        return { error: "No tienes permiso para usar este sistema." };
    _cbz = _cbz === "semana" ? "semana" : "dia";
    var _ccc = fechasSemanaChecklist_(_cca), _ccd = "";
    if (_cbz === "dia" && _ccb) {
        var _cce = { dias: _ccc.map(function () { return {}; }) };
        if (elegirDiaChecklist_(_cce, _cca, _ccb))
            return { error: "Ese día no es válido." };
        if (_cce.diaElegidoTexto)
            _ccd = String(_ccb);
    }
    return { ok: true, envios: anotarEnvioChecklist_(_cbz, _ccc, _ccd, true) };
}
function enviarChecklistPorCorreo(_ccf, _ccg, _cch) {
    try {
        revisarRegistros_(15000);
    }
    catch (_cci) { }
    var _ccj = obtenerDatosCalendario(_ccg);
    var _cck = _ccf === 'dia' ? elegirDiaChecklist_(_ccj, _ccg, _cch) : null;
    if (_cck)
        return { enviado: false, motivo: 'dia_no_valido', error: _cck };
    marcarReenvioChecklist_(_ccj, _ccf, _ccg, _cch);
    fechaDiaChecklist_(_ccj, _ccf, _cch);
    var _ccl = asuntoChecklist_(_ccj, _ccf);
    var _ccm = construirTextoChecklist_(_ccj, _ccf);
    if (!CORREO_JEFATURA) {
        return { enviado: false, motivo: 'sin_correo_configurado', asunto: _ccl };
    }
    var _ccn = construirHtmlChecklistCorreo_(_ccj, _ccf, true);
    var _cco = { htmlBody: _ccn, inlineImages: construirInlineImagesLogos_() };
    if (ccChecklist_())
        _cco.cc = ccChecklist_();
    var _ccp = "";
    try {
        _ccp = String(Session.getActiveUser().getEmail() || "").trim();
    }
    catch (_ccq) { }
    if (!_ccp) {
        try {
            _ccp = String(PropertiesService.getUserProperties().getProperty('devoluciones_email') || "").trim();
        }
        catch (_ccr) { }
    }
    if (_ccp && [CORREO_JEFATURA, CORREO_CC_ENCARGADO, CORREO_CC_EXTRA_CHECKLIST_].some(function (_ccs) { return String(_ccs || "").trim().toLowerCase() === _ccp.toLowerCase(); }))
        _ccp = "";
    if (_ccp)
        _cco.bcc = _ccp;
    MailApp.sendEmail(CORREO_JEFATURA, _ccl, _ccm, _cco);
    var _cct = null;
    try {
        _cct = anotarEnvioChecklist_(_ccf, fechasSemanaChecklist_(_ccg), _ccj.diaElegidoTexto ? String(_cch) : "");
    }
    catch (_ccu) { }
    return { enviado: true, asunto: _ccl, destinatario: CORREO_JEFATURA, cc: ccChecklist_(), copiaOculta: _ccp, esPrueba: CORREO_MODO_PRUEBA, envios: _cct };
}
function obtenerPendientesCalendario(_ccv) {
    var _ccw = ["VELOX", "PAQNORTE", "CORREOMAX", "RUTASUR", "ATLAS", "PRONTO", "BOLIDO", "FARO"];
    var _ccx = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_AGENCIAS);
    var _ccy = [];
    var _ccz = new Date(), _cda = _ccz.getDay() || 7;
    var _cdb = new Date(_ccz);
    _cdb.setDate(_ccz.getDate() - _cda + 1 + (_ccv * 7));
    _cdb.setHours(0, 0, 0, 0);
    var _cdc = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO", "JULIO", "AGOSTO", "SEPTIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"];
    var _cdd = new Date(_cdb);
    _cdd.setDate(_cdb.getDate() + 6);
    var _cde = ["ene.", "feb.", "mar.", "abr.", "may.", "jun.", "jul.", "ago.", "sept.", "oct.", "nov.", "dic."];
    var _cdf = _cdb.getMonth() === _cdd.getMonth();
    var _cdg = _cdf
        ? ('Semana del ' + _cdb.getDate() + ' al ' + _cdd.getDate() + ' · ' + _cdc[_cdb.getMonth()] + ' ' + _cdb.getFullYear())
        : ('Semana del ' + _cdb.getDate() + ' ' + _cde[_cdb.getMonth()] + ' al ' + _cdd.getDate() + ' ' + _cde[_cdd.getMonth()] + ' ' + _cdd.getFullYear());
    if (!_ccx)
        return { pendientes: _ccy, compañias: _ccw, total: 0, truncado: false, mes: _cdg };
    var _cdh = ["DOM", "LUNES", "MARTES", "MIÉRCOLES", "JUEVES", "VIERNES", "SÁB.", "DOM."];
    var _cdi = [];
    for (var _cdj = 0; _cdj < 7; _cdj++) {
        var _cdk = new Date(_cdb);
        _cdk.setDate(_cdb.getDate() + _cdj);
        _cdi.push({ fObj: _cdk, label: _cdh[_cdk.getDay() === 0 ? 7 : _cdk.getDay()] + " " + ('0' + _cdk.getDate()).slice(-2) });
    }
    var _cdl = _ccx.getDataRange().getValues(), _cdm = {};
    for (var _cdn = 1; _cdn < _cdl.length; _cdn++) {
        for (var _cdo = 0; _cdo < 7; _cdo++) {
            if (esMismoDia(new Date(_cdl[_cdn][0]), _cdi[_cdo].fObj)) {
                _cdm[_cdo + "_" + _cdl[_cdn][2] + "_" + _cdl[_cdn][4]] = { idx: _cdo, ag: (_cdl[_cdn][3] || "").toString().toUpperCase().trim(), origen: _cdl[_cdn][2], fila: _cdl[_cdn][4] };
                break;
            }
        }
    }
    var _cdp = {};
    for (var _cdq in _cdm) {
        var _cdr = _cdm[_cdq];
        if (_cdr.ag === "" || _cdr.origen === HOJA_RETORNOS)
            continue;
        if (!_cdp[_cdr.origen])
            _cdp[_cdr.origen] = [];
        _cdp[_cdr.origen].push(_cdr);
    }
    for (var _cds in _cdp) {
        var _cdt = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(_cds);
        if (!_cdt)
            continue;
        var _cdu = _cdt.getLastRow();
        if (_cdu < 2)
            continue;
        var _cdv = COLUMNA_NOTA_L - COLUMNA_REFERENCIAS + 1;
        var _cdw = _cdt.getRange(2, COLUMNA_REFERENCIAS, _cdu - 1, _cdv).getValues();
        var _cdx = 0;
        var _cdy = COLUMNA_NOTA_L - COLUMNA_REFERENCIAS;
        _cdp[_cds].forEach(function (_cdz) {
            var _cea = _cdz.fila - 2;
            var _ceb = (_cea >= 0 && _cea < _cdw.length) ? _cdw[_cea] : null;
            var _cec = _ceb ? (_ceb[_cdy] || "").toString().trim() : "";
            if (_cec === "") {
                var _ced = _ceb ? (_ceb[_cdx] || "").toString().trim() : "";
                _ccy.push({ hoja: _cds, fila: _cdz.fila, col: COLUMNA_NOTA_L, agencia: _cdz.ag, dia: _cdi[_cdz.idx].label, referencia: _ced });
            }
        });
    }
    _ccy.sort(function (_cee, _cef) { return _cee.agencia > _cef.agencia ? 1 : (_cee.agencia < _cef.agencia ? -1 : _cee.fila - _cef.fila); });
    var _ceg = _ccy.length;
    var _ceh = _ceg > 200;
    return { pendientes: _ccy.slice(0, 200), compañias: _ccw, total: _ceg, truncado: _ceh, mes: _cdg };
}
function esMismoDia(_cei, _cej) { return _cei.getDate() === _cej.getDate() && _cei.getMonth() === _cej.getMonth() && _cei.getFullYear() === _cej.getFullYear(); }
function columnToLetter(_cek) { var _cel, _cem = ''; while (_cek > 0) {
    _cel = (_cek - 1) % 26;
    _cem = String.fromCharCode(65 + _cel) + _cem;
    _cek = (_cek - _cel) / 26 | 0;
} return _cem; }
function getOrCreateSheet(_cen, _ceo) {
    var _cep = _cen.getSheetByName(_ceo);
    if (!_cep) {
        _cep = _cen.insertSheet(_ceo);
        if (_ceo === HOJA_AGENCIAS || _ceo === HOJA_ROTURAS || _ceo === HOJA_AGENCIAS_ARCHIVO || _ceo === HOJA_ROTURAS_ARCHIVO)
            _cep.appendRow(["Fecha", "Usuario", "Origen", "Agencia", "Fila"]);
        else if (_ceo === HOJA_NOTAS || _ceo === HOJA_NOTAS_ARCHIVO)
            _cep.appendRow(["Fecha", "Usuario", "Nota", "Estado"]);
        else if (_ceo === HOJA_AVISOS)
            _cep.appendRow(["Fecha", "Usuario", "Referencia", "Motivo", "Contador", "Última vez", "Último origen"]);
        else if (_ceo === HOJA_AVISOS_EVENTOS)
            _cep.appendRow(["Id", "Fecha", "ReferenciaNormalizada", "Referencia", "Hoja", "Fila", "Motivo", "Estado"]);
        else if (_ceo === HOJA_CONTROL_REVISION)
            _cep.appendRow(["Hoja", "Fila", "Referencia", "Agencia", "Estado (L)", "Rotura (E)"]);
        else
            _cep.appendRow(["Fecha", "Usuario", "Origen", "Estado", "Col", "Letra", "Fila", "Agencia"]);
        _cep.getRange("A1:H1").setFontWeight("bold");
        _cep.hideSheet();
    }
    return _cep;
}
function normalizarRef(_ceq) {
    return (_ceq || "").toString().toLowerCase().trim().replace(/^0+(?=\d)/, "");
}
function levenshtein(_cer, _ces) {
    var _cet = _cer.length, _ceu = _ces.length;
    if (_cet === 0)
        return _ceu;
    if (_ceu === 0)
        return _cet;
    var _cev = [];
    for (var _cew = 0; _cew <= _cet; _cew++)
        _cev[_cew] = _cew;
    for (var _cex = 1; _cex <= _ceu; _cex++) {
        var _cey = _cev[0];
        _cev[0] = _cex;
        for (var _cew = 1; _cew <= _cet; _cew++) {
            var _cez = _cev[_cew];
            if (_cer.charAt(_cew - 1) === _ces.charAt(_cex - 1)) {
                _cev[_cew] = _cey;
            }
            else {
                _cev[_cew] = Math.min(_cey + 1, _cev[_cew] + 1, _cev[_cew - 1] + 1);
            }
            _cey = _cez;
        }
    }
    return _cev[_cet];
}
function calcularPorcentajeSimilitud_(_cfa, _cfb, _cfc, _cfd) {
    if (_cfa)
        return 100;
    var _cfe = Math.max(_cfc, _cfd, 1);
    var _cff = Math.round((1 - (_cfb / _cfe)) * 100);
    if (_cff > 99)
        _cff = 99;
    if (_cff < 0)
        _cff = 0;
    return _cff;
}
function backendBuscarReferencia(_cfg, _cfh, _cfi, _cfj) {
    var _cfk = new Date().getTime(), _cfl = [], _cfm = _cfk;
    function _cfn(_cfo) { if (_cfl.length)
        _cfl[_cfl.length - 1] += ' ' + (_cfo - _cfm) + ' ms'; _cfm = _cfo; }
    try {
        return backendBuscarReferenciaPasos_(_cfg, _cfh, _cfi, _cfj, function (_cfp) { _cfn(new Date().getTime()); _cfl.push(_cfp); });
    }
    finally {
        var _cfq = new Date().getTime(), _cfr = _cfq - _cfk;
        _cfn(_cfq);
        try {
            console.log('Búsqueda «' + _cfg + '»: ' + _cfl.join(' → ') + ' · total ' + _cfr + ' ms');
        }
        catch (_cfs) { }
        if (_cfr > 10000)
            sumarContador_("BUSQLENTAS", 1);
    }
}
function backendBuscarReferenciaPasos_(_cft, _cfu, _cfv, _cfw, _cfx) {
    if (_cfv !== true) {
        var _cfy = null;
        _cfx("recientes");
        try {
            _cfy = buscarReferenciaReciente_(_cft, _cfu, _cfw);
        }
        catch (_cfz) {
            _cfy = null;
        }
        if (_cfy && _cfy.length > 0)
            return _cfy;
        var _cga = null;
        _cfx("columna C");
        try {
            _cga = buscarReferenciaExactaColumnas_(_cft, _cfu);
        }
        catch (_cgb) {
            _cga = undefined;
        }
        if (_cga === undefined) {
            _cfx("todo el documento");
            try {
                _cga = buscarReferenciaExactaRapida_(_cft, _cfu);
            }
            catch (_cgc) {
                _cga = null;
            }
        }
        if (_cga)
            return _cga;
    }
    _cfx(_cfv === true ? "parecidas" : "completa");
    return buscarReferenciaCompleta_(_cft, _cfu);
}
var COLUMNA_NOMBRE_HOJA1_ = 9;
var HOJA_NOMBRES_ = "Entradas";
var UMBRAL_PARECIDO_NOMBRE_ = 75;
function backendBuscarNombre(_cgd, _cge) {
    var _cgf = new Date().getTime(), _cgg = 0;
    try {
        var _cgh = buscarNombreHoja1_(_cgd, _cge);
        _cgg = _cgh && _cgh.resultados ? _cgh.resultados.length : 0;
        return _cgh;
    }
    finally {
        try {
            console.log('Búsqueda por nombre «' + _cgd + '»: ' + _cgg + ' resultados en ' + (new Date().getTime() - _cgf) + ' ms');
        }
        catch (_cgi) { }
    }
}
function parecidoNombre_(_cgj, _cgk, _cgl) {
    var _cgm = _cgk.split(' '), _cgn = _cgj.length, _cgo = 0, _cgp = UMBRAL_PARECIDO_NOMBRE_ / 100;
    for (var _cgq = 0; _cgq < _cgn; _cgq++) {
        var _cgr = _cgj[_cgq], _cgs = _cgl[_cgr] || (_cgl[_cgr] = Object.create(null)), _cgt = 0;
        for (var _cgu = 0; _cgu < _cgm.length; _cgu++) {
            var _cgv = _cgm[_cgu], _cgw = _cgs[_cgv];
            if (_cgw === undefined) {
                var _cgx = Math.max(_cgr.length, _cgv.length);
                _cgw = _cgs[_cgv] = (!_cgx || Math.abs(_cgr.length - _cgv.length) * 2 > _cgx) ? 0 : 1 - levenshtein(_cgr, _cgv) / _cgx;
            }
            if (_cgw > _cgt)
                _cgt = _cgw;
        }
        _cgo += _cgt;
        if ((_cgo + (_cgn - _cgq - 1)) / _cgn < _cgp)
            return 0;
    }
    return Math.min(99, Math.round(_cgo / _cgn * 100));
}
function buscarNombreHoja1_(_cgy, _cgz) {
    var _cha = 30;
    var _chb = prepararBusquedaRetornos_(_cgy);
    if (_chb.compacto.length < 3)
        return { error: "Escribe al menos 3 letras del nombre." };
    var _chc = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_NOMBRES_);
    if (!_chc)
        return { error: "No existe la pestaña " + HOJA_NOMBRES_ + "." };
    var _chd = _chc.getLastRow();
    if (_chd < 2 || _chc.getMaxColumns() < COLUMNA_NOMBRE_HOJA1_)
        return { resultados: [], total: 0, exactas: false };
    var _che = _chc.getRange(2, COLUMNA_NOMBRE_HOJA1_, _chd - 1, 1).getValues();
    var _chf = String(_cgz || "").toUpperCase().trim();
    var _chg = _chf ? _chc.getRange(2, COLUMNA_AGENCIAS, _chd - 1, 1).getValues() : null;
    var _chh = [];
    function _chi(_chj) {
        for (var _chk = _che.length - 1; _chk >= 0; _chk--) {
            var _chl = textoCelda_(_che[_chk][0]);
            if (!_chl)
                continue;
            var _chm = _chh[_chk];
            if (_chm === undefined)
                _chm = _chh[_chk] = normTextoRetornos_(_chl);
            if (!/[a-z]/.test(_chm))
                continue;
            if (_chg && String(_chg[_chk][0] === null || _chg[_chk][0] === undefined ? "" : _chg[_chk][0]).toUpperCase().trim().indexOf(_chf) === -1)
                continue;
            _chj(_chk + 2, _chl, _chm);
        }
    }
    var _chn = [];
    _chi(function (_cho, _chp, _chq) {
        var _chr = _chb.palabras.length > 0 && _chb.palabras.every(function (_chs) { return _chq.indexOf(_chs) > -1; });
        if (_chr || compactoRetornos_(_chp).indexOf(_chb.compacto) > -1)
            _chn.push({ fila: _cho, valor: _chp, pct: 100, exacta: true });
    });
    var _cht = _chn.length > 0;
    if (!_cht && _chb.palabras.length > 0 && _chb.compacto.length >= 4) {
        var _chu = Object.create(null);
        _chi(function (_chv, _chw, _chx) {
            var _chy = parecidoNombre_(_chb.palabras, _chx, _chu);
            if (_chy >= UMBRAL_PARECIDO_NOMBRE_)
                _chn.push({ fila: _chv, valor: _chw, pct: _chy, exacta: false });
        });
        _chn.sort(function (_chz, _cia) { return (_cia.pct - _chz.pct) || (_cia.fila - _chz.fila); });
    }
    var _cib = _chn.length;
    _chn = _chn.slice(0, _cha);
    if (_chn.length === 0)
        return { resultados: [], total: 0, exactas: false };
    var _cic = Math.min(COLUMNA_NOTA_L, _chc.getMaxColumns()) - COLUMNA_AGENCIAS + 1;
    var _cid = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _cie = COLUMNA_NOTA_L - COLUMNA_AGENCIAS;
    var _cif = leerFilasSueltas_(_chc, _chn.map(function (_cig) { return _cig.fila; }), _cic);
    return {
        total: _cib, exactas: _cht,
        resultados: _chn.map(function (_cih) {
            var _cii = _cif[_cih.fila] || [];
            return {
                hoja: HOJA_NOMBRES_, fila: _cih.fila, col: COLUMNA_REFERENCIAS,
                agencia: String(_cii[0] || "").toUpperCase().trim() || "SIN AGENCIA",
                matchStr: String(_cii[_cid] || ""), exacta: _cih.exacta, distancia: _cih.exacta ? 0 : 1,
                porcentaje: _cih.pct, notaL: _cic > _cie ? String(_cii[_cie] || "").trim() : "", nombre: _cih.valor
            };
        })
    };
}
function buscarReferenciaExactaColumnas_(_cij, _cik) {
    var _cil = 30, _cim = 60;
    var _cin = normalizarRef(_cij);
    if (_cin === "")
        return null;
    var _cio = (_cik || "").toString().toUpperCase().trim();
    var _cip = SpreadsheetApp.getActiveSpreadsheet(), _ciq = _cip.getSheets(), _cir = [];
    var _cis = columnToLetter(COLUMNA_REFERENCIAS);
    for (var _cit = 0; _cit < _ciq.length; _cit++) {
        var _ciu = _ciq[_cit], _civ = _ciu.getName();
        if (TODAS_HOJAS_HISTORIAL_.indexOf(_civ) > -1)
            continue;
        var _ciw = [];
        if (HOJAS_TRABAJO_PERMITIDAS_.indexOf(_civ) > -1) {
            var _cix = _ciu.getLastRow();
            if (_cix < 2)
                continue;
            var _ciy = _ciu.getRange(2, COLUMNA_REFERENCIAS, _cix - 1, 1).getValues();
            for (var _ciz = _ciy.length - 1; _ciz >= 0; _ciz--) {
                var _cja = normalizarRef(textoCelda_(_ciy[_ciz][0]));
                if (_cja !== "" && _cja.indexOf(_cin) !== -1)
                    _ciw.push(_ciz + 2);
            }
        }
        else {
            var _cjb;
            try {
                _cjb = _ciu.getRange(_cis + "2:" + _cis).createTextFinder(_cin).matchCase(false).findAll();
            }
            catch (_cjc) {
                if (_ciu.getMaxRows() < 2)
                    continue;
                throw _cjc;
            }
            _ciw = _cjb.map(function (_cjd) { return _cjd.getRow(); }).sort(function (_cje, _cjf) { return _cjf - _cje; });
        }
        if (_ciw.length === 0)
            continue;
        var _cjg = _ciu.getLastColumn() >= COLUMNA_NOTA_L;
        for (var _cjh = 0; _cjh < _ciw.length; _cjh++)
            _cir.push({ hoja: _ciu, fila: _ciw[_cjh], tieneColL: _cjg });
        if (_cir.length > _cim)
            return null;
    }
    var _cji = [];
    var _cjj = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _cjk = COLUMNA_NOTA_L - COLUMNA_AGENCIAS;
    for (var _cjl = 0; _cjl < _cir.length && _cji.length < _cil; _cjl++) {
        var _cjm = _cir[_cjl];
        var _cjn = (_cjm.tieneColL ? COLUMNA_NOTA_L : COLUMNA_REFERENCIAS) - COLUMNA_AGENCIAS + 1;
        var _cjo = _cjm.hoja.getRange(_cjm.fila, COLUMNA_AGENCIAS, 1, _cjn).getDisplayValues()[0];
        var _cjp = (_cjo[0] || "").toString().toUpperCase().trim();
        if (_cio !== "" && _cjp.indexOf(_cio) === -1)
            continue;
        var _cjq = _cjo[_cjj];
        var _cjr = normalizarRef(_cjq);
        if (_cjr === "" || _cjr.indexOf(_cin) === -1)
            continue;
        _cji.push({
            hoja: _cjm.hoja.getName(),
            fila: _cjm.fila,
            col: COLUMNA_REFERENCIAS,
            agencia: _cjp || "SIN AGENCIA",
            matchStr: _cjq,
            exacta: true,
            distancia: 0,
            porcentaje: calcularPorcentajeSimilitud_(true, 0, _cin.length, _cjr.length),
            notaL: _cjm.tieneColL ? (_cjo[_cjk] || "").toString().trim() : "",
            multi: _cjm.tieneColL ? (String(_cjo[COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS + 1] || "").trim().toUpperCase().replace("Í", "I") === "SI") : false
        });
    }
    return _cji;
}
function buscarReferenciaExactaRapida_(_cjs, _cjt) {
    var _cju = 30, _cjv = 60;
    var _cjw = normalizarRef(_cjs);
    if (_cjw === "")
        return null;
    var _cjx = (_cjt || "").toString().toUpperCase().trim();
    var _cjy = SpreadsheetApp.getActiveSpreadsheet();
    var _cjz = (_cjw.length >= 6) ? buscarCandidatasGlobal_(_cjy, _cjw, _cjv) : undefined;
    if (_cjz === null)
        return null;
    var _cka = _cjz ? [] : _cjy.getSheets();
    if (!_cjz)
        _cjz = [];
    for (var _ckb = 0; _ckb < _cka.length; _ckb++) {
        var _ckc = _cka[_ckb];
        if (TODAS_HOJAS_HISTORIAL_.indexOf(_ckc.getName()) > -1)
            continue;
        var _ckd = columnToLetter(COLUMNA_REFERENCIAS);
        var _cke;
        try {
            _cke = _ckc.getRange(_ckd + "2:" + _ckd).createTextFinder(_cjw).matchCase(false).findAll();
        }
        catch (_ckf) {
            if (_ckc.getMaxRows() < 2)
                continue;
            throw _ckf;
        }
        if (_cke.length === 0)
            continue;
        var _ckg = _ckc.getLastColumn() >= COLUMNA_NOTA_L;
        var _ckh = _cke.map(function (_cki) { return _cki.getRow(); }).sort(function (_ckj, _ckk) { return _ckk - _ckj; });
        for (var _ckl = 0; _ckl < _ckh.length; _ckl++)
            _cjz.push({ hoja: _ckc, fila: _ckh[_ckl], tieneColL: _ckg });
        if (_cjz.length > _cjv)
            return null;
    }
    var _ckm = [];
    var _ckn = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _cko = COLUMNA_NOTA_L - COLUMNA_AGENCIAS;
    for (var _ckp = 0; _ckp < _cjz.length && _ckm.length < _cju; _ckp++) {
        var _ckq = _cjz[_ckp];
        var _ckr = (_ckq.tieneColL ? COLUMNA_NOTA_L : COLUMNA_REFERENCIAS) - COLUMNA_AGENCIAS + 1;
        var _cks = _ckq.hoja.getRange(_ckq.fila, COLUMNA_AGENCIAS, 1, _ckr).getDisplayValues()[0];
        var _ckt = (_cks[0] || "").toString().toUpperCase().trim();
        if (_cjx !== "" && _ckt.indexOf(_cjx) === -1)
            continue;
        var _cku = _cks[_ckn];
        var _ckv = normalizarRef(_cku);
        if (_ckv === "" || _ckv.indexOf(_cjw) === -1)
            continue;
        _ckm.push({
            hoja: _ckq.hoja.getName(),
            fila: _ckq.fila,
            col: COLUMNA_REFERENCIAS,
            agencia: _ckt || "SIN AGENCIA",
            matchStr: _cku,
            exacta: true,
            distancia: 0,
            porcentaje: calcularPorcentajeSimilitud_(true, 0, _cjw.length, _ckv.length),
            notaL: _ckq.tieneColL ? (_cks[_cko] || "").toString().trim() : "",
            multi: _ckq.tieneColL ? (String(_cks[COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS + 1] || "").trim().toUpperCase().replace("Í", "I") === "SI") : false
        });
    }
    return _ckm;
}
var FILAS_RECIENTES_BUSQUEDA_ = 1200;
var FILAS_PARECIDAS_PANEL_ = 3000;
var CLAVE_CACHE_MARCAS_BUSCAR_ = "BUSCAR_MARCAS_v2";
function fechaCortaMarca_(_ckw) {
    if (_ckw instanceof Date && !isNaN(_ckw.getTime()))
        return ('0' + _ckw.getDate()).slice(-2) + '/' + ('0' + (_ckw.getMonth() + 1)).slice(-2);
    var _ckx = String(_ckw === null || _ckw === undefined ? "" : _ckw).trim(), _cky = /^(\d{1,2})[\/\-.](\d{1,2})/.exec(_ckx);
    return _cky ? ('0' + _cky[1]).slice(-2) + '/' + ('0' + _cky[2]).slice(-2) : _ckx;
}
var CLAVE_CALCULO_MARCAS_BUSCAR_ = "BUSCAR_MARCAS_CALCULANDO";
function marcasRecientesBuscar_(_ckz, _cla) {
    var _clb = null;
    try {
        _clb = CacheService.getDocumentCache();
        var _clc = _clb.get(CLAVE_CACHE_MARCAS_BUSCAR_);
        if (_clc)
            return JSON.parse(_clc);
    }
    catch (_cld) { }
    try {
        if (_clb) {
            if (_clb.get(CLAVE_CALCULO_MARCAS_BUSCAR_))
                return null;
            _clb.put(CLAVE_CALCULO_MARCAS_BUSCAR_, "1", 120);
        }
    }
    catch (_cle) { }
    try {
        return marcasRecientesBuscarCalculo_(_ckz, _cla, _clb);
    }
    finally {
        try {
            if (_clb)
                _clb.remove(CLAVE_CALCULO_MARCAS_BUSCAR_);
        }
        catch (_clf) { }
    }
}
function marcasRecientesBuscarCalculo_(_clg, _clh, _cli) {
    var _clj = { rep: {}, color: {} }, _clk = null;
    _clh.forEach(function (_cll) { if (_cll.hoja === "Entradas")
        _clk = _cll.filas; });
    var _clm = _clg.getSheetByName("Entradas");
    if (!_clm || !_clk || !_clk.length)
        return _clj;
    var _cln = _clm.getLastRow();
    if (_cln < 2)
        return _clj;
    var _clo = {}, _clp = {};
    _clk.forEach(function (_clq) { _clo[_clq[1]] = []; if (_clq[5] === "SI")
        _clp[_clq[1]] = true; });
    var _clr = _clm.getRange(2, 1, _cln - 1, COLUMNA_REFERENCIAS).getValues();
    for (var _cls = 0; _cls < _clr.length; _cls++) {
        var _clt = normalizarRef(textoCelda_(_clr[_cls][COLUMNA_REFERENCIAS - 1]));
        if (_clt && _clo[_clt])
            _clo[_clt].push(fechaCortaMarca_(_clr[_cls][0]));
    }
    Object.keys(_clo).forEach(function (_clu) {
        var _clv = _clo[_clu], _clw = {}, _clx = [];
        _clv.forEach(function (_cly) { if (!(_cly in _clw)) {
            _clw[_cly] = 0;
            _clx.push(_cly);
        } _clw[_cly]++; });
        var _clz = _clx.slice(0, 5).map(function (_cma) { return [_cma, _clw[_cma]]; });
        if (_clp[_clu])
            _clj.rep[_clu] = { t: 'M', n: _clv.length, d: _clz };
        else if (_clv.length > 1)
            _clj.rep[_clu] = { t: 'R', n: _clv.length, d: _clz };
    });
    var _cmb = _clk.map(function (_cmc) { return _cmc[0]; }), _cmd = Math.min.apply(null, _cmb), _cme = Math.max.apply(null, _cmb);
    var _cmf = Math.min(COLUMNA_NOTA_L, _clm.getMaxColumns()), _cmg = _clm.getRange(_cmd, 1, _cme - _cmd + 1, _cmf).getBackgrounds();
    _cmb.forEach(function (_cmh) {
        var _cmi = _cmg[_cmh - _cmd];
        if (!_cmi)
            return;
        var _cmj = String(_cmi[0] || "").toLowerCase();
        if (!_cmj || _cmj === "#ffffff" || _cmj === "white")
            return;
        for (var _cmk = 1; _cmk < _cmi.length; _cmk++)
            if (String(_cmi[_cmk] || "").toLowerCase() !== _cmj)
                return;
        _clj.color[_cmh] = _cmj;
    });
    try {
        var _cml = JSON.stringify(_clj);
        if (_cli && _cml.length < 95000)
            _cli.put(CLAVE_CACHE_MARCAS_BUSCAR_, _cml, 600);
    }
    catch (_cmm) { }
    return _clj;
}
function obtenerRecientesBuscar() {
    if (!usuarioAutorizado_())
        return null;
    var _cmn = new Date().getTime(), _cmo = SpreadsheetApp.getActiveSpreadsheet(), _cmp = { t: _cmn, hojas: [] };
    var _cmq = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _cmr = COLUMNA_NOTA_L - COLUMNA_AGENCIAS, _cms = COLUMNA_NOTA_L - COLUMNA_AGENCIAS + 1;
    for (var _cmt = 0; _cmt < HOJAS_TRABAJO_PERMITIDAS_.length; _cmt++) {
        var _cmu = _cmo.getSheetByName(HOJAS_TRABAJO_PERMITIDAS_[_cmt]);
        if (!_cmu)
            continue;
        var _cmv = _cmu.getLastRow(), _cmw = [], _cmx = [], _cmy = [];
        if (_cmv >= 2) {
            var _cmz = Math.max(2, _cmv - FILAS_PARECIDAS_PANEL_ - 1000 + 1);
            var _cna = _cmu.getRange(_cmz, COLUMNA_REFERENCIAS, _cmv - _cmz + 1, 1).getValues();
            var _cnb = _cna.length - 1;
            while (_cnb >= 0 && !/\d/.test(textoCelda_(_cna[_cnb][0])))
                _cnb--;
            var _cnc = Math.max(0, _cnb - FILAS_RECIENTES_BUSQUEDA_ + 1);
            var _cnd = Math.max(0, _cnb - FILAS_PARECIDAS_PANEL_ + 1);
            if (_cnb >= 0 && _cnd < _cnc) {
                try {
                    var _cne = _cmu.getRange(_cmz + _cnd, COLUMNA_AGENCIAS, _cnc - _cnd, 1).getValues(), _cnf = {}, _cng = [], _cnh = [];
                    for (var _cni = _cnc - 1; _cni >= _cnd; _cni--) {
                        var _cnj = normalizarRef(textoCelda_(_cna[_cni][0])).replace(/\|/g, "");
                        var _cnk = _cnj === "" ? "" : (_cne[_cni - _cnd][0] || "").toString().toUpperCase().trim();
                        if (_cnj !== "" && _cnf[_cnk] === undefined && _cmy.length < 36) {
                            _cnf[_cnk] = _cmy.length;
                            _cmy.push(_cnk);
                        }
                        _cng.push(_cnj);
                        _cnh.push(_cnj === "" || _cnf[_cnk] === undefined ? "-" : _cnf[_cnk].toString(36));
                    }
                    _cmx = { top: _cmz + _cnc - 1, n: _cng.join("|"), a: _cnh.join("") };
                }
                catch (_cnl) {
                    _cmx = [];
                    _cmy = [];
                }
            }
            if (_cnb >= 0) {
                var _cnm;
                try {
                    _cnm = _cmu.getRange(_cmz + _cnc, COLUMNA_AGENCIAS, _cnb - _cnc + 1, _cms).getDisplayValues();
                }
                catch (_cnn) {
                    _cms = Math.min(COLUMNA_NOTA_L, _cmu.getMaxColumns()) - COLUMNA_AGENCIAS + 1;
                    _cnm = _cmu.getRange(_cmz + _cnc, COLUMNA_AGENCIAS, _cnb - _cnc + 1, _cms).getDisplayValues();
                }
                for (var _cno = _cnb; _cno >= _cnc; _cno--) {
                    var _cnp = normalizarRef(textoCelda_(_cna[_cno][0]));
                    if (_cnp === "")
                        continue;
                    var _cnq = _cnm[_cno - _cnc];
                    _cmw.push([_cmz + _cno, _cnp, String(_cnq[_cmq] || ""), (_cnq[0] || "").toString().toUpperCase().trim(), _cms > _cmr ? (_cnq[_cmr] || "").toString().trim() : "", String(_cnq[_cmq + 1] || "").trim().toUpperCase().replace("Í", "I")]);
                }
            }
        }
        _cmp.hojas.push({ hoja: _cmu.getName(), filas: _cmw, extra: _cmx, ag: _cmy });
    }
    try {
        _cmp.notas = mapaNotasPedido_(leerNotas_());
    }
    catch (_cnr) {
        _cmp.notas = null;
    }
    try {
        var _cns = (new Date().getTime() - _cmn > 15000) ? null : marcasRecientesBuscar_(_cmo, _cmp.hojas);
        _cmp.rep = _cns ? _cns.rep : null;
        _cmp.color = _cns ? _cns.color : null;
    }
    catch (_cnt) {
        _cmp.rep = null;
        _cmp.color = null;
    }
    try {
        console.log('Filas recientes para Buscar: ' + _cmp.hojas.map(function (_cnu) { return _cnu.hoja + ' ' + _cnu.filas.length; }).join(', ') + ' · ' + (new Date().getTime() - _cmn) + ' ms');
    }
    catch (_cnv) { }
    return _cmp;
}
function obtenerLFilasBuscar(_cnw) {
    if (!usuarioAutorizado_())
        return null;
    var _cnx = SpreadsheetApp.getActiveSpreadsheet(), _cny = {}, _cnz = {};
    (_cnw || []).slice(0, 30).forEach(function (_coa) {
        if (!_coa || HOJAS_TRABAJO_PERMITIDAS_.indexOf(_coa.hoja) === -1 || !(Number(_coa.fila) >= 2))
            return;
        (_cnz[_coa.hoja] = _cnz[_coa.hoja] || []).push(Number(_coa.fila));
    });
    Object.keys(_cnz).forEach(function (_cob) {
        var _coc = _cnx.getSheetByName(_cob);
        if (!_coc || _coc.getLastColumn() < COLUMNA_NOTA_L)
            return;
        var _cod = _cnz[_cob].sort(function (_coe, _cof) { return _coe - _cof; }), _cog = _cod[0], _coh = _cod[_cod.length - 1];
        if (_coh - _cog < 400) {
            var _coi = _coc.getRange(_cog, COLUMNA_NOTA_L, _coh - _cog + 1, 1).getDisplayValues();
            _cod.forEach(function (_coj) { _cny[_cob + '#' + _coj] = String(_coi[_coj - _cog][0] || '').trim(); });
        }
        else
            _cod.forEach(function (_cok) { _cny[_cob + '#' + _cok] = String(_coc.getRange(_cok, COLUMNA_NOTA_L).getDisplayValue() || '').trim(); });
    });
    return _cny;
}
var FILAS_DELTA_BUSCAR_ = 150;
function obtenerFilasNuevasBuscar(_col) {
    if (!usuarioAutorizado_())
        return null;
    _col = _col || {};
    var _com = new Date().getTime(), _con = SpreadsheetApp.getActiveSpreadsheet(), _coo = { t: _com, hojas: [] };
    var _cop = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _coq = COLUMNA_NOTA_L - COLUMNA_AGENCIAS;
    for (var _cor = 0; _cor < HOJAS_TRABAJO_PERMITIDAS_.length; _cor++) {
        var _cos = _con.getSheetByName(HOJAS_TRABAJO_PERMITIDAS_[_cor]);
        if (!_cos)
            continue;
        var _cot = _cos.getLastRow(), _cou = _cos.getName(), _cov = Number(_col[_cou]) || 0;
        if (_cov && _cot - _cov > FILAS_DELTA_BUSCAR_ + 500) {
            _coo.recargar = true;
            return _coo;
        }
        var _cow = Math.max(2, _cot - FILAS_DELTA_BUSCAR_ + 1), _cox = [];
        if (_cot >= 2) {
            var _coy = _cot - _cow + 1;
            var _coz = _cos.getRange(_cow, COLUMNA_REFERENCIAS, _coy, 1).getValues();
            var _cpa = Math.min(COLUMNA_NOTA_L, _cos.getMaxColumns()) - COLUMNA_AGENCIAS + 1;
            var _cpb = _cos.getRange(_cow, COLUMNA_AGENCIAS, _coy, _cpa).getDisplayValues();
            for (var _cpc = _coy - 1; _cpc >= 0; _cpc--) {
                var _cpd = normalizarRef(textoCelda_(_coz[_cpc][0]));
                if (_cpd === "")
                    continue;
                var _cpe = _cpb[_cpc];
                _cox.push([_cow + _cpc, _cpd, String(_cpe[_cop] || ""), (_cpe[0] || "").toString().toUpperCase().trim(), _cpa > _coq ? (_cpe[_coq] || "").toString().trim() : "", String(_cpe[_cop + 1] || "").trim().toUpperCase().replace("Í", "I")]);
            }
        }
        _coo.hojas.push({ hoja: _cou, desde: _cow, filas: _cox });
    }
    return _coo;
}
function buscarReferenciaReciente_(_cpf, _cpg, _cph) {
    var _cpi = (_cph && typeof _cph === 'object') ? String(_cph.hoja || "") : "", _cpj = (_cph && typeof _cph === 'object') ? Math.floor(Number(_cph.fila)) : 0;
    var _cpk = 30;
    var _cpl = normalizarRef(_cpf);
    if (_cpl === "")
        return [];
    var _cpm = (_cpg || "").toString().toUpperCase().trim();
    var _cpn = SpreadsheetApp.getActiveSpreadsheet();
    var _cpo = COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS, _cpp = COLUMNA_NOTA_L - COLUMNA_AGENCIAS;
    var _cpq = [];
    for (var _cpr = 0; _cpr < HOJAS_TRABAJO_PERMITIDAS_.length && _cpq.length < _cpk; _cpr++) {
        var _cps = _cpn.getSheetByName(HOJAS_TRABAJO_PERMITIDAS_[_cpr]);
        if (!_cps)
            continue;
        var _cpt = _cps.getLastRow();
        if (_cpt < 2)
            continue;
        var _cpu = Math.max(2, _cpt - FILAS_RECIENTES_BUSQUEDA_ - 1000 + 1);
        var _cpv = _cps.getRange(_cpu, COLUMNA_REFERENCIAS, _cpt - _cpu + 1, 1).getValues();
        var _cpw = _cpv.length - 1;
        while (_cpw >= 0 && !/\d/.test(textoCelda_(_cpv[_cpw][0])))
            _cpw--;
        var _cpx = Math.max(0, _cpw - FILAS_RECIENTES_BUSQUEDA_ + 1);
        var _cpy = [];
        for (var _cpz = _cpw; _cpz >= _cpx; _cpz--) {
            var _cqa = normalizarRef(textoCelda_(_cpv[_cpz][0]));
            if (_cqa !== "" && _cqa.indexOf(_cpl) !== -1)
                _cpy.push(_cpu + _cpz);
        }
        var _cqb = _cpu + _cpx;
        if (_cpi && _cpj >= 2 && _cpj < _cqb && _cpi === _cps.getName()) {
            var _cqc = Math.max(2, _cpj - 400), _cqd = Math.min(_cqb - 1, _cpj + 400);
            if (_cqd >= _cqc) {
                var _cqe = _cps.getRange(_cqc, COLUMNA_REFERENCIAS, _cqd - _cqc + 1, 1).getValues();
                for (var _cqf = _cqe.length - 1; _cqf >= 0; _cqf--) {
                    var _cqg = normalizarRef(textoCelda_(_cqe[_cqf][0]));
                    if (_cqg !== "" && _cqg.indexOf(_cpl) !== -1)
                        _cpy.push(_cqc + _cqf);
                }
            }
        }
        if (_cpy.length === 0)
            continue;
        var _cqh = COLUMNA_NOTA_L - COLUMNA_AGENCIAS + 1, _cqi;
        try {
            _cqi = leerFilasSueltas_(_cps, _cpy, _cqh);
        }
        catch (_cqj) {
            _cqh = Math.min(COLUMNA_NOTA_L, _cps.getMaxColumns()) - COLUMNA_AGENCIAS + 1;
            _cqi = leerFilasSueltas_(_cps, _cpy, _cqh);
        }
        var _cqk = _cqh > _cpp;
        for (var _cql = 0; _cql < _cpy.length && _cpq.length < _cpk; _cql++) {
            var _cqm = _cpy[_cql], _cqn = _cqi[_cqm];
            var _cqo = (_cqn[0] || "").toString().toUpperCase().trim();
            if (_cpm !== "" && _cqo.indexOf(_cpm) === -1)
                continue;
            var _cqp = _cqn[_cpo];
            var _cqq = normalizarRef(_cqp);
            if (_cqq === "" || _cqq.indexOf(_cpl) === -1)
                continue;
            _cpq.push({
                hoja: _cps.getName(),
                fila: _cqm,
                col: COLUMNA_REFERENCIAS,
                agencia: _cqo || "SIN AGENCIA",
                matchStr: _cqp,
                exacta: true,
                distancia: 0,
                porcentaje: calcularPorcentajeSimilitud_(true, 0, _cpl.length, _cqq.length),
                notaL: _cqk ? (_cqn[_cpp] || "").toString().trim() : "",
                multi: _cqh > COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS + 1 ? (String(_cqn[COLUMNA_REFERENCIAS - COLUMNA_AGENCIAS + 1] || "").trim().toUpperCase().replace("Í", "I") === "SI") : false
            });
        }
    }
    return _cpq;
}
function textoCelda_(_cqr) {
    if (_cqr === null || _cqr === undefined || _cqr instanceof Date)
        return "";
    return String(_cqr);
}
function leerFilasSueltas_(_cqs, _cqt, _cqu) {
    var _cqv = {};
    if (_cqt.length === 0)
        return _cqv;
    var _cqw = Math.min.apply(null, _cqt), _cqx = Math.max.apply(null, _cqt);
    if (_cqt.length <= 6 || _cqx - _cqw + 1 > 1500) {
        for (var _cqy = 0; _cqy < _cqt.length; _cqy++)
            _cqv[_cqt[_cqy]] = _cqs.getRange(_cqt[_cqy], COLUMNA_AGENCIAS, 1, _cqu).getDisplayValues()[0];
    }
    else {
        var _cqz = _cqs.getRange(_cqw, COLUMNA_AGENCIAS, _cqx - _cqw + 1, _cqu).getDisplayValues();
        for (var _cra = 0; _cra < _cqt.length; _cra++)
            _cqv[_cqt[_cra]] = _cqz[_cqt[_cra] - _cqw];
    }
    return _cqv;
}
function buscarCandidatasGlobal_(_crb, _crc, _crd) {
    var _cre;
    try {
        _cre = _crb.createTextFinder(_crc).matchCase(false).findAll();
    }
    catch (_crf) {
        return undefined;
    }
    if (_cre.length > 400)
        return undefined;
    var _crg = {}, _crh = [];
    for (var _cri = 0; _cri < _cre.length; _cri++) {
        var _crj = _cre[_cri];
        if (_crj.getColumn() !== COLUMNA_REFERENCIAS || _crj.getRow() < 2)
            continue;
        var _crk = _crj.getSheet(), _crl = _crk.getName();
        if (TODAS_HOJAS_HISTORIAL_.indexOf(_crl) > -1)
            continue;
        if (!_crg[_crl])
            _crg[_crl] = { orden: _crk.getIndex(), tieneColL: _crk.getLastColumn() >= COLUMNA_NOTA_L };
        _crh.push({ hoja: _crk, fila: _crj.getRow(), tieneColL: _crg[_crl].tieneColL, orden: _crg[_crl].orden });
        if (_crh.length > _crd)
            return null;
    }
    _crh.sort(function (_crm, _crn) { return (_crm.orden - _crn.orden) || (_crn.fila - _crm.fila); });
    return _crh;
}
function buscarReferenciaCompleta_(_cro, _crp) {
    var _crq = SpreadsheetApp.getActiveSpreadsheet();
    var _crr = _crq.getSheets();
    var _crs = [];
    var _crt = normalizarRef(_cro);
    var _cru = _crp.toString().toUpperCase().trim();
    var _crv = TODAS_HOJAS_HISTORIAL_;
    var _crw = _crt.length <= 4 ? 1 : (_crt.length <= 8 ? 2 : 3);
    var _crx = _crt.length;
    var _cry = Math.min(COLUMNA_AGENCIAS, COLUMNA_REFERENCIAS);
    var _crz = Math.max(COLUMNA_AGENCIAS, COLUMNA_REFERENCIAS);
    var _csa = COLUMNA_AGENCIAS - _cry;
    var _csb = COLUMNA_REFERENCIAS - _cry;
    var _csc = 30;
    var _csd = 0;
    buscarCompleto: for (var _cse = 0; _cse < _crr.length; _cse++) {
        var _csf = _crr[_cse];
        if (_crv.indexOf(_csf.getName()) > -1)
            continue;
        var _csg = _csf.getLastRow();
        if (_csg < 2)
            continue;
        var _csh = _csf.getLastColumn() >= COLUMNA_NOTA_L;
        var _csi = (_cru === "");
        var _csj = _csi ? _csf.getRange(2, COLUMNA_REFERENCIAS, _csg - 1, 1).getValues() : _csf.getRange(2, _cry, _csg - 1, _crz - _cry + 1).getValues();
        var _csk = _csi ? 0 : _csb;
        for (var _csl = _csj.length - 1; _csl >= 0; _csl--) {
            var _csm = _csi ? null : (_csj[_csl][_csa] || "").toString().toUpperCase().trim();
            if (_cru !== "" && _csm.indexOf(_cru) === -1)
                continue;
            var _csn = textoCelda_(_csj[_csl][_csk]);
            var _cso = normalizarRef(_csn);
            if (_cso === "")
                continue;
            var _csp = _cso.indexOf(_crt) > -1;
            var _csq;
            if (_csp) {
                _csq = 0;
            }
            else if (Math.abs(_cso.length - _crx) > _crw) {
                _csq = _crw + 1;
            }
            else {
                _csq = levenshtein(_crt, _cso);
            }
            if (_csp || _csq <= _crw) {
                var _csr = null;
                _crs.push({
                    _hoja: _csf, _tieneL: _csh,
                    hoja: _csf.getName(),
                    fila: _csl + 2,
                    col: COLUMNA_REFERENCIAS,
                    agencia: _csm === null ? null : (_csm || "SIN AGENCIA"),
                    matchStr: _csn,
                    exacta: _csp,
                    distancia: _csq,
                    porcentaje: calcularPorcentajeSimilitud_(_csp, _csq, _crx, _cso.length),
                    notaL: _csr
                });
                if (_csp && ++_csd >= _csc)
                    break buscarCompleto;
            }
        }
    }
    _crs.sort(function (_css, _cst) { return _css.distancia - _cst.distancia; });
    var _csu = _crs.slice(0, _csc);
    completarAgenciaYL_(_csu);
    return _csu;
}
function completarAgenciaYL_(_csv) {
    var _csw = {};
    _csv.forEach(function (_csx) { var _csy = _csx.hoja; (_csw[_csy] = _csw[_csy] || { hoja: _csx._hoja, tieneL: _csx._tieneL, filas: [] }).filas.push(_csx.fila); });
    var _csz = {};
    Object.keys(_csw).forEach(function (_cta) {
        var _ctb = _csw[_cta], _ctc = _ctb.filas.slice().sort(function (_ctd, _cte) { return _ctd - _cte; }), _ctf = [], _ctg = _ctc[0], _cth = _ctc[0];
        for (var _cti = 1; _cti < _ctc.length; _cti++) {
            if (_ctc[_cti] - _cth <= 30)
                _cth = _ctc[_cti];
            else {
                _ctf.push([_ctg, _cth]);
                _ctg = _cth = _ctc[_cti];
            }
        }
        _ctf.push([_ctg, _cth]);
        _ctf.forEach(function (_ctj) {
            var _ctk = _ctj[1] - _ctj[0] + 1;
            var _ctl = _ctb.hoja.getRange(_ctj[0], COLUMNA_AGENCIAS, _ctk, 1).getValues();
            var _ctm = _ctb.tieneL ? _ctb.hoja.getRange(_ctj[0], COLUMNA_NOTA_L, _ctk, 1).getValues() : null;
            for (var _ctn = 0; _ctn < _ctk; _ctn++)
                _csz[_cta + '#' + (_ctj[0] + _ctn)] = { b: _ctl[_ctn][0], l: _ctm ? _ctm[_ctn][0] : null };
        });
    });
    _csv.forEach(function (_cto) {
        var _ctp = _csz[_cto.hoja + '#' + _cto.fila] || { b: "", l: null };
        if (_cto.agencia === null)
            _cto.agencia = (_ctp.b || "").toString().toUpperCase().trim() || "SIN AGENCIA";
        _cto.notaL = _cto._tieneL ? textoCelda_(_ctp.l).trim() : "";
        delete _cto._hoja;
        delete _cto._tieneL;
    });
}
function activarCeldaEnHoja(_ctq, _ctr, _cts) {
    var _ctt = SpreadsheetApp.getActiveSpreadsheet();
    var _ctu = _ctt.getSheetByName(_ctq);
    if (_ctu) {
        _ctu.activate();
        _ctu.getRange(_ctr, _cts).activate();
    }
}
var VALORES_ROTURA_PERMITIDOS = ["SI", "NO"];
var VALORES_VERIFICACION_PERMITIDOS = ["OK", "RECLAMAR", "RECLAMAR/PENDIENTE"];
function rellenarCampoRapido(_ctv, _ctw, _ctx, _cty, _ctz) {
    var _cua = new Date().getTime();
    PLAZO_REGISTRO_ = _cua + 20000;
    var _cub = 0;
    var _cuc = _ctx === "rotura" ? COLUMNA_ROTURA : (_ctx === "verificacion" ? COLUMNA_NOTA_L : null);
    var _cud = _ctx === "rotura" ? VALORES_ROTURA_PERMITIDOS : VALORES_VERIFICACION_PERMITIDOS;
    if (!_cuc || _cud.indexOf(_cty) === -1)
        return { error: true, mensaje: "Campo o valor no válido." };
    var _cue = SpreadsheetApp.getActiveSpreadsheet();
    var _cuf = _cue.getSheetByName(_ctv);
    if (!_cuf)
        return { error: true, mensaje: "No se encontró la hoja." };
    if (_ctz !== undefined && _ctz !== null && String(_ctz).trim() !== "") {
        var _cug = _cuf.getRange(_ctw, COLUMNA_REFERENCIAS), _cuh = normalizarRef(_ctz);
        if (normalizarRef(_cug.getDisplayValue()) !== _cuh && normalizarRef(textoCelda_(_cug.getValue())) !== _cuh) {
            var _cui = relocalizarFilaPedido_(_cuf, _ctw, _ctz);
            if (!_cui)
                return { error: true, cambiado: true, mensaje: "El pedido " + _ctz + " ya no está en su fila de " + _ctv + " y no se ha podido localizar (¿lo han borrado?). No se ha escrito nada: vuelve a buscarlo." };
            _cub = _ctw;
            _ctw = _cui;
        }
    }
    if (_ctx === "verificacion" && _cty !== "OK" && _cty !== "RECLAMAR")
        _cty = valorSegunDesplegable_(_cuf.getRange(_ctw, _cuc), _cty);
    var _cuj = (_ctv === HOJA_RETORNOS) ? _cuf.getRange(_ctw, 1, 1, COL_RET_ULTIMA_).getValues()[0] : null;
    _cuf.getRange(_ctw, _cuc).setValue(_cty);
    var _cuk = registrarCambioCelda_(_cue, _cuf, _ctw, _cuc, 1, 1);
    if (_cuj && new Date().getTime() - _cua > 12000) {
        try {
            PropertiesService.getDocumentProperties().setProperty('RET_COPIA_DESCUADRE', '1');
        }
        catch (_cul) { }
        sumarContador_("HISTRETSALTADO", 1);
        _cuj = null;
    }
    if (_cuj) {
        var _cum = {};
        _cum[_ctw] = _cuj;
        try {
            registrarCambiosRetornos_(_cuf, _ctw, _cuc, 1, 1, { antesFilas: _cum });
        }
        catch (_cun) { }
    }
    if (_cuk)
        _cue.toast("✅ Registrado.", "Actualizado", 3);
    return { error: false, hoja: _ctv, fila: _ctw, campo: _ctx, valor: _cty, movida: _cub ? true : false };
}
var HUCHO_VERSION_ = "1.1.0";
var HUCHO_CORREOS_ = ["admin@ejemplo.com"];
var HUCHO_CLAVE_ = "demo";
var HUCHO_HOJA_ = "Entradas";
var HUCHO_COL_AGENCIA_ = 2;
var HUCHO_COL_PEDIDO_ = 3;
var HUCHO_COL_FECHA_ = 7;
var HUCHO_COL_PAGO_ = 8;
var HUCHO_COL_NOMBRE_ = 9;
var HUCHO_COL_EMAIL_ = 10;
var HUCHO_UMBRAL_BLOQUEO_ = 3;
var HUCHO_UMBRAL_VIGILAR_ = 2;
var HUCHO_DESTACAR_ = {
    "TARJETA": 5, "PAYPAL": 5, "BIZUM": 4, "APLAZAME": 4, "TRANSFERENCIA": 4,
    "_OTROS": 4,
    "_TODOS": 6
};
var HUCHO_TODOS_ = "TODOS";
var HUCHO_MESES_ATRAS_ = 2;
var HUCHO_MAX_LISTA_ = 300;
var HUCHO_MAX_CORREO_ = 40;
var HUCHO_PARA_ = "luis.prado@ejemplo.com";
var HUCHO_CC_ = "jefatura@ejemplo.com";
var HUCHO_MESES_TXT_ = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"];
function huchoCorreoUsuario_() {
    var _cuo = "";
    try {
        _cuo = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    }
    catch (_cup) { }
    if (!_cuo) {
        try {
            _cuo = String(PropertiesService.getUserProperties().getProperty('devoluciones_email') || "").trim().toLowerCase();
        }
        catch (_cuq) { }
    }
    return _cuo;
}
function huchoPuedeVer_() {
    var _cur = huchoCorreoUsuario_();
    if (!_cur)
        return false;
    for (var _cus = 0; _cus < HUCHO_CORREOS_.length; _cus++) {
        if (String(HUCHO_CORREOS_[_cus]).trim().toLowerCase() === _cur)
            return true;
    }
    return false;
}
function huchoClaveOk_(_cut) {
    if (!huchoPuedeVer_())
        return false;
    return String(_cut || "") === HUCHO_CLAVE_;
}
function huchoExigir_(_cuu) {
    if (!huchoPuedeVer_())
        throw new Error("Hucho es solo para las personas autorizadas.");
    if (!huchoClaveOk_(_cuu))
        throw new Error("La contraseña no es correcta. Vuelve a abrir a Hucho.");
}
function huchoEntrar(_cuv) {
    if (!huchoPuedeVer_())
        return { ok: false, motivo: "Hucho es solo para las personas autorizadas." };
    var _cuw = null, _cux = "PCN_FALLOS";
    try {
        _cuw = CacheService.getUserCache();
    }
    catch (_cuy) { }
    var _cuz = 0;
    try {
        _cuz = Number(_cuw && _cuw.get(_cux)) || 0;
    }
    catch (_cva) { }
    if (_cuz >= 5)
        return { ok: false, motivo: "Demasiados intentos. Espera 5 minutos y vuelve a probar." };
    if (String(_cuv || "") !== HUCHO_CLAVE_) {
        try {
            if (_cuw)
                _cuw.put(_cux, String(_cuz + 1), 300);
        }
        catch (_cvb) { }
        return { ok: false, motivo: "Contraseña incorrecta" };
    }
    try {
        if (_cuw)
            _cuw.remove(_cux);
    }
    catch (_cvc) { }
    var _cvd = false;
    try {
        _cvd = (typeof usuarioAutorizado_ === 'function') && usuarioAutorizado_() === true;
    }
    catch (_cve) { }
    return { ok: true, equipo: _cvd };
}
function huchoAbrir(_cvf) {
    if (!huchoPuedeVer_()) {
        try {
            SpreadsheetApp.getUi().alert("🔒 No tienes permiso para usar esta herramienta.");
        }
        catch (_cvg) { }
        return false;
    }
    var _cvh = HtmlService.createTemplateFromFile('Contrarreembolso_panel');
    _cvh.despertar = _cvf === true;
    _cvh.version = HUCHO_VERSION_;
    var _cvi = _cvh.evaluate().setTitle("🐷 Hucho");
    SpreadsheetApp.getUi().showSidebar(_cvi);
    return true;
}
function mHucho() { huchoAbrir(false); }
function huchoVolverARufo() {
    if (typeof mRufoRapido === 'function')
        mRufoRapido();
}
function huchoTarjetaHtml() {
    if (!huchoPuedeVer_())
        return null;
    var _cvj = '<style>' +
        '.pcn-card{position:relative;margin:14px auto 0;max-width:260px;border-radius:18px;padding:10px 12px 8px;text-align:center;cursor:pointer;' +
        'background:radial-gradient(120% 90% at 50% 0%,#fff 0%,#fff1f4 55%,#ffe1e8 100%);border:1px solid #f9c3d0;' +
        'box-shadow:0 1px 2px rgba(15,23,42,.05),0 10px 24px -10px rgba(200,16,46,.35);transition:transform .15s}' +
        '.pcn-card:hover{transform:translateY(-2px)}' +
        '.pcn-card .pcn-pig{width:86px;height:94px;display:block;margin:0 auto}' +
        '.pcn-card b{display:block;font-size:14px;color:#1e293b}' +
        '.pcn-card small{display:block;font-size:11px;color:#64748b;margin-top:1px}' +
        '.pcn-card .pcn-cand{position:absolute;top:9px;right:10px;width:20px;height:24px;animation:pcnBamb 2.8s ease-in-out infinite}' +
        '.pcn-chip-pig{width:22px;height:24px;vertical-align:middle}' +
        '@keyframes pcnBamb{0%,100%{transform:rotate(-6deg)}50%{transform:rotate(6deg)}}' +
        huchoCssCerdito_() +
        '</style>';
    var _cvk = "var t=this;if(t.getAttribute('data-abriendo'))return;t.setAttribute('data-abriendo','1');t.style.opacity='.7';" +
        "google.script.run.withSuccessHandler(function(){}).withFailureHandler(function(){t.removeAttribute('data-abriendo');t.style.opacity='';}).huchoAbrir(true)";
    var _cvl = _cvj +
        '<div class="pcn-card" title="Hucho · contrarreembolso" onclick="' + _cvk + '">' +
        huchoSvgCandado_('pcn-cand') +
        huchoSvgCerdito_('') +
        '<b>Hucho</b><small>Contrarreembolso y otros pagos · pide contraseña</small>' +
        '</div>';
    var _cvm = '<button class="rufo-chip" onclick="' + _cvk + '"><span>🐷</span> Hucho · pagos devueltos 🔒</button>';
    return { tarjeta: _cvl, chip: _cvm };
}
function huchoFecha_(_cvn) {
    if (_cvn instanceof Date)
        return isNaN(_cvn.getTime()) ? null : _cvn;
    if (_cvn === null || _cvn === undefined || _cvn === "")
        return null;
    var _cvo = String(_cvn).trim(), _cvp = /^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/.exec(_cvo);
    if (_cvp) {
        var _cvq = Number(_cvp[3]);
        if (_cvq < 100)
            _cvq += 2000;
        var _cvr = new Date(_cvq, Number(_cvp[2]) - 1, Number(_cvp[1]));
        return (_cvr.getMonth() === Number(_cvp[2]) - 1) ? _cvr : null;
    }
    _cvp = /^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(_cvo);
    if (_cvp)
        return new Date(Number(_cvp[1]), Number(_cvp[2]) - 1, Number(_cvp[3]));
    return null;
}
function huchoIndiceMes_(_cvs) { return _cvs.getFullYear() * 12 + _cvs.getMonth(); }
function huchoNombreMes_(_cvt) { return HUCHO_MESES_TXT_[_cvt % 12] + " " + Math.floor(_cvt / 12); }
function huchoEsContra_(_cvu) {
    var _cvv = String(_cvu === null || _cvu === undefined ? "" : _cvu).trim().toUpperCase().replace(/\s+/g, " ");
    return _cvv === "CONTRA" || /^CONTRA[ -]?R?REEMBOLSO$/.test(_cvv) || /^CONTRA\b/.test(_cvv) && !/^CONTRAT/.test(_cvv);
}
function huchoNormPedido_(_cvw) {
    return String(_cvw === null || _cvw === undefined ? "" : _cvw).trim().toLowerCase().replace(/\s+/g, "").replace(/^0+(?=\d)/, "");
}
function huchoNormNombre_(_cvx) {
    return String(_cvx === null || _cvx === undefined ? "" : _cvx).normalize('NFD').replace(/[̀-ͯ]/g, '')
        .toLowerCase().replace(/[^a-z0-9 ]+/g, ' ').replace(/\s+/g, ' ').trim();
}
function huchoDosDigitos_(_cvy) { return ('0' + _cvy).slice(-2); }
function huchoHoy_() { var _cvz = new Date(); return huchoDosDigitos_(_cvz.getDate()) + '/' + huchoDosDigitos_(_cvz.getMonth() + 1) + '/' + _cvz.getFullYear(); }
function huchoMetodo_(_cwa) {
    if (huchoEsContra_(_cwa))
        return "CONTRA";
    return String(_cwa === null || _cwa === undefined ? "" : _cwa).trim().toUpperCase().replace(/\s+/g, " ").substring(0, 40);
}
function huchoUmbrales_(_cwb) {
    if (_cwb === "CONTRA")
        return { destacar: HUCHO_UMBRAL_BLOQUEO_, vigilar: HUCHO_UMBRAL_VIGILAR_, bloqueo: true };
    var _cwc = HUCHO_DESTACAR_._OTROS;
    if (_cwb === HUCHO_TODOS_)
        _cwc = HUCHO_DESTACAR_._TODOS;
    else
        for (var _cwd in HUCHO_DESTACAR_) {
            if (_cwd.charAt(0) !== "_" && _cwb.indexOf(_cwd) > -1) {
                _cwc = HUCHO_DESTACAR_[_cwd];
                break;
            }
        }
    return { destacar: _cwc, vigilar: Math.max(2, _cwc - 1), bloqueo: false };
}
function huchoCalcularMes_(_cwe) {
    var _cwf = new Date().getTime();
    var _cwg = new Date(), _cwh = huchoIndiceMes_(_cwg) + _cwe, _cwi = _cwh - HUCHO_MESES_ATRAS_;
    var _cwj = { mes: _cwh, mesTxt: huchoNombreMes_(_cwh), mesAntTxt: huchoNombreMes_(_cwh - 1), metodos: [], rankings: {}, filasLeidas: 0, celdasLeidas: 0, ms: 0 };
    var _cwk = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HUCHO_HOJA_);
    if (!_cwk) {
        _cwj.error = "No encuentro la pestaña «" + HUCHO_HOJA_ + "».";
        return _cwj;
    }
    var _cwl = _cwk.getLastRow();
    if (_cwl < 2)
        return _cwj;
    var _cwm = _cwk.getRange(2, HUCHO_COL_FECHA_, _cwl - 1, 1).getValues();
    var _cwn = -1, _cwo = -1;
    for (var _cwp = 0; _cwp < _cwm.length; _cwp++) {
        var _cwq = huchoFecha_(_cwm[_cwp][0]);
        if (!_cwq)
            continue;
        var _cwr = huchoIndiceMes_(_cwq);
        if (_cwr >= _cwi && _cwr <= _cwh) {
            if (_cwn < 0)
                _cwn = _cwp;
            _cwo = _cwp;
        }
    }
    _cwj.filasLeidas = _cwm.length;
    _cwj.celdasLeidas = _cwm.length;
    if (_cwn < 0) {
        _cwj.ms = new Date().getTime() - _cwf;
        return _cwj;
    }
    var _cws = _cwo - _cwn + 1;
    var _cwt = _cwk.getRange(2 + _cwn, HUCHO_COL_PAGO_, _cws, 3).getValues();
    var _cwu = _cwk.getRange(2 + _cwn, HUCHO_COL_AGENCIA_, _cws, 2).getDisplayValues();
    _cwj.filasLeidas += _cws;
    _cwj.celdasLeidas += _cws * 5;
    var _cwv = {};
    function _cww(_cwx) { return _cwv[_cwx] || (_cwv[_cwx] = { vistos: {}, clientes: {}, nombresMes: {}, pedidosMes: 0 }); }
    function _cwy(_cwz, _cxa, _cxb, _cxc, _cxd, _cxe, _cxf, _cxg, _cxh, _cxi) {
        if (_cwz.vistos[_cxa + "|" + _cxb])
            return;
        _cwz.vistos[_cxa + "|" + _cxb] = true;
        var _cxj = _cxf || (_cxh ? "nombre:" + _cxh : "");
        if (!_cxj)
            return;
        var _cxk = _cwz.clientes[_cxj];
        if (!_cxk)
            _cxk = _cwz.clientes[_cxj] = { k: _cxj, e: _cxf, n: _cxg, nNorm: _cxh, meses: {}, p: [] };
        if (!_cxk.n && _cxg) {
            _cxk.n = _cxg;
            _cxk.nNorm = _cxh;
        }
        _cxk.meses[_cxa] = (_cxk.meses[_cxa] || 0) + 1;
        if (_cxa === _cwh) {
            _cwz.pedidosMes++;
            _cxk.p.push([_cxc || "(sin nº)", huchoDosDigitos_(_cxd.getDate()) + "/" + huchoDosDigitos_(_cxd.getMonth() + 1), _cxi, _cxe]);
            if (_cxh && _cxf) {
                (_cwz.nombresMes[_cxh] = _cwz.nombresMes[_cxh] || {})[_cxf] = true;
            }
        }
    }
    for (var _cxl = 0; _cxl < _cws; _cxl++) {
        var _cxm = _cwt[_cxl], _cxn = huchoMetodo_(_cxm[0]);
        if (!_cxn)
            continue;
        var _cxo = huchoFecha_(_cwm[_cwn + _cxl][0]);
        if (!_cxo)
            continue;
        var _cxp = huchoIndiceMes_(_cxo);
        if (_cxp < _cwi || _cxp > _cwh)
            continue;
        var _cxq = 2 + _cwn + _cxl, _cxr = _cwu[_cxl];
        var _cxs = String(_cxr[1] || "").trim(), _cxt = huchoNormPedido_(_cxs) || ("fila" + _cxq);
        var _cxu = String(_cxm[2] === null || _cxm[2] === undefined ? "" : _cxm[2]).trim().toLowerCase();
        var _cxv = String(_cxm[1] === null || _cxm[1] === undefined ? "" : _cxm[1]).trim().replace(/\s+/g, " ");
        var _cxw = huchoNormNombre_(_cxv), _cxx = String(_cxr[0] || "").trim();
        _cwy(_cww(_cxn), _cxp, _cxt, _cxs, _cxo, _cxq, _cxu, _cxv, _cxw, _cxx);
        _cwy(_cww(HUCHO_TODOS_), _cxp, _cxt, _cxs, _cxo, _cxq, _cxu, _cxv, _cxw, _cxx);
    }
    for (var _cxy in _cwv) {
        var _cxz = _cwv[_cxy], _cya = huchoUmbrales_(_cxy), _cyb = [], _cyc = 0, _cyd = 0;
        for (var _cye in _cxz.clientes) {
            var _cyf = _cxz.clientes[_cye], _cyg = _cyf.meses[_cwh] || 0;
            if (!_cyg)
                continue;
            var _cyh = _cyf.meses[_cwh - 1] || 0, _cyi = _cyf.meses[_cwh - 2] || 0;
            var _cyj = 0;
            if (_cyg >= _cya.vigilar) {
                _cyj = 1;
                if (_cyh >= _cya.vigilar) {
                    _cyj = 2;
                    if (_cyi >= _cya.vigilar)
                        _cyj = 3;
                }
            }
            var _cyk = [];
            if (_cyf.nNorm && _cxz.nombresMes[_cyf.nNorm])
                for (var _cyl in _cxz.nombresMes[_cyf.nNorm])
                    if (_cyl !== _cyf.e)
                        _cyk.push(_cyl);
            if (_cyg >= _cya.destacar)
                _cyc++;
            else if (_cyg >= _cya.vigilar)
                _cyd++;
            _cyb.push({ k: _cyf.k, n: _cyf.n || "(sin nombre)", e: _cyf.e, c: _cyg, ant: _cyh, ant2: _cyi, rein: _cyj, otros: _cyk.slice(0, 2), p: _cyf.p });
        }
        if (!_cyb.length)
            continue;
        _cyb.sort(function (_cym, _cyn) { return _cyn.c - _cym.c || _cyn.ant - _cym.ant || (_cym.n < _cyn.n ? -1 : _cym.n > _cyn.n ? 1 : 0); });
        _cwj.rankings[_cxy] = {
            metodo: _cxy, pedidos: _cxz.pedidosMes, clientes: _cyb.length, conBloqueo: _cyc, conVigilar: _cyd,
            umbralBloqueo: _cya.destacar, umbralVigilar: _cya.vigilar, esContra: _cya.bloqueo,
            lista: _cyb.slice(0, HUCHO_MAX_LISTA_), recortado: _cyb.length > HUCHO_MAX_LISTA_
        };
        if (_cxy !== HUCHO_TODOS_)
            _cwj.metodos.push({ m: _cxy, pedidos: _cxz.pedidosMes });
    }
    _cwj.metodos.sort(function (_cyo, _cyp) { return (_cyp.m === "CONTRA") - (_cyo.m === "CONTRA") || _cyp.pedidos - _cyo.pedidos || (_cyo.m < _cyp.m ? -1 : 1); });
    if (_cwj.rankings[HUCHO_TODOS_])
        _cwj.metodos.push({ m: HUCHO_TODOS_, pedidos: _cwj.rankings[HUCHO_TODOS_].pedidos });
    _cwj.ms = new Date().getTime() - _cwf;
    return _cwj;
}
function huchoMes_(_cyq, _cyr) {
    var _cys = Math.max(-24, Math.min(0, Math.round(Number(_cyq) || 0)));
    var _cyt = String(_cyr || "CONTRA").trim().toUpperCase().substring(0, 40) || "CONTRA";
    var _cyu = huchoIndiceMes_(new Date()) + _cys, _cyv = "PCN_MES_v2_" + _cyu;
    var _cyw = null;
    try {
        _cyw = CacheService.getDocumentCache();
    }
    catch (_cyx) { }
    function _cyy(_cyz, _cza) {
        var _czb = _cza || { metodo: _cyt, pedidos: 0, clientes: 0, conBloqueo: 0, conVigilar: 0, lista: [], recortado: false };
        var _czc = huchoUmbrales_(_cyt);
        if (!_cza) {
            _czb.umbralBloqueo = _czc.destacar;
            _czb.umbralVigilar = _czc.vigilar;
            _czb.esContra = _czc.bloqueo;
        }
        _czb.mes = _cyz.mes;
        _czb.mesTxt = _cyz.mesTxt;
        _czb.mesAntTxt = _cyz.mesAntTxt;
        _czb.metodos = _cyz.metodos;
        if (_cyz.error)
            _czb.error = _cyz.error;
        return _czb;
    }
    if (_cyw) {
        try {
            var _czd = [_cyv + "_resumen", _cyv + "_m_" + _cyt], _cze = _cyw.getAll(_czd);
            if (_cze[_czd[0]]) {
                var _czf = JSON.parse(_cze[_czd[0]]), _czg = _cze[_czd[1]] ? JSON.parse(_cze[_czd[1]]) : null;
                var _czh = _czf.metodos.some(function (_czi) { return _czi.m === _cyt; });
                if (_czg || !_czh) {
                    var _czj = _cyy(_czf, _czg);
                    _czj.deCache = true;
                    return _czj;
                }
            }
        }
        catch (_czk) { }
    }
    var _czl = huchoCalcularMes_(_cys);
    var _czm = { mes: _czl.mes, mesTxt: _czl.mesTxt, mesAntTxt: _czl.mesAntTxt, metodos: _czl.metodos, error: _czl.error };
    if (_cyw && !_czl.error) {
        try {
            var _czn = {}, _czo = _cys === 0 ? 300 : 21600;
            _czn[_cyv + "_resumen"] = JSON.stringify(_czm);
            for (var _czp in _czl.rankings) {
                var _czq = JSON.stringify(_czl.rankings[_czp]);
                if (_czq.length < 95000)
                    _czn[_cyv + "_m_" + _czp] = _czq;
            }
            _cyw.putAll(_czn, _czo);
        }
        catch (_czr) { }
    }
    var _czs = _cyy(_czm, _czl.rankings[_cyt]);
    _czs.filasLeidas = _czl.filasLeidas;
    _czs.celdasLeidas = _czl.celdasLeidas;
    _czs.ms = _czl.ms;
    return _czs;
}
function huchoMarcas_() {
    var _czt = {};
    try {
        _czt = PropertiesService.getDocumentProperties().getProperties() || {};
    }
    catch (_czu) { }
    var _czv = {}, _czw = {};
    for (var _czx in _czt) {
        if (_czx.indexOf("PCN_BLOQ|") === 0) {
            try {
                _czv[_czx.substring(9)] = JSON.parse(_czt[_czx]);
            }
            catch (_czy) { }
        }
        else if (_czx.indexOf("PCN_ENV|") === 0) {
            try {
                _czw[_czx.substring(8)] = JSON.parse(_czt[_czx]);
            }
            catch (_czz) { }
        }
    }
    return { bloq: _czv, env: _czw };
}
function huchoClaveEnvio_(_daa, _dab) { return _dab === "CONTRA" ? String(_daa) : _daa + "|" + _dab; }
function huchoDatosMes(_dac, _dad, _dae) {
    huchoExigir_(_dac);
    var _daf = huchoMes_(_dad, _dae), _dag = huchoMarcas_();
    _daf.lista.forEach(function (_dah) { if (_dag.bloq[_dah.k])
        _dah.bloq = _dag.bloq[_dah.k].f || "sí"; });
    _daf.enviado = _dag.env[huchoClaveEnvio_(_daf.mes, _daf.metodo)] || null;
    _daf.offset = Math.max(-24, Math.min(0, Math.round(Number(_dad) || 0)));
    return _daf;
}
function huchoMarcarBloqueo(_dai, _daj, _dak) {
    huchoExigir_(_dai);
    var _dal = String(_daj || "").trim().toLowerCase();
    if (!_dal || _dal.length > 400)
        throw new Error("Cliente no válido.");
    var _dam = PropertiesService.getDocumentProperties();
    if (_dak) {
        var _dan = { f: huchoHoy_(), u: huchoCorreoUsuario_() };
        _dam.setProperty("PCN_BLOQ|" + _dal, JSON.stringify(_dan));
        return { ok: true, bloq: _dan.f };
    }
    _dam.deleteProperty("PCN_BLOQ|" + _dal);
    return { ok: true, bloq: null };
}
function huchoIrAFila(_dao, _dap) {
    huchoExigir_(_dao);
    var _daq = Math.round(Number(_dap));
    var _dar = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HUCHO_HOJA_);
    if (!_dar || !(_daq >= 2) || _daq > _dar.getMaxRows())
        return false;
    _dar.activate();
    _dar.getRange(_daq, HUCHO_COL_PEDIDO_).activate();
    return true;
}
function huchoEsc_(_das) {
    return String(_das === null || _das === undefined ? "" : _das).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
function huchoNombreMetodo_(_dat) { return _dat === HUCHO_TODOS_ ? "todos los métodos de pago" : (_dat === "CONTRA" ? "contrarreembolso" : _dat); }
function huchoSituacion_(_dau, _dav) {
    var _daw = [], _dax = "";
    if (_dau.bloq) {
        _daw.push("Ya bloqueado (" + _dau.bloq + ")");
        _dax = "ok";
    }
    else if (_dau.c >= _dav.umbralBloqueo) {
        var _day = _dav.esContra ? "propuesto bloquear" : "destacado";
        _daw.push(_dau.rein >= 2 ? "Reincidente (" + _dau.rein + " meses) · " + _day : (_dau.ant === 0 ? "Nuevo este mes · " + _day : _day.charAt(0).toUpperCase() + _day.slice(1)));
        _dax = "rojo";
    }
    else if (_dau.c >= _dav.umbralVigilar) {
        _daw.push(_dau.rein >= 2 ? "Reincidente (" + _dau.rein + " meses) · vigilar" : "Vigilar");
    }
    if (_dau.otros && _dau.otros.length) {
        _daw.push("mismo nombre con otro email · revisar");
        if (!_dax)
            _dax = "aviso";
    }
    return { texto: _daw.join(" · ") || "—", tipo: _dax };
}
function huchoCorreo(_daz, _dba, _dbb, _dbc) {
    huchoExigir_(_daz);
    var _dbd = huchoDatosMes(_daz, _dba, _dbc);
    var _dbe = Math.max(1, Math.min(9, Math.round(Number(_dbb) || _dbd.umbralVigilar)));
    var _dbf = _dbd.lista.filter(function (_dbg) { return _dbg.c >= _dbe; });
    var _dbh = _dbf.length > HUCHO_MAX_CORREO_;
    _dbf = _dbf.slice(0, HUCHO_MAX_CORREO_);
    var _dbi = _dbd.mesTxt.charAt(0).toUpperCase() + _dbd.mesTxt.slice(1);
    var _dbj = huchoNombreMetodo_(_dbd.metodo);
    var _dbk = _dbd.metodo === HUCHO_TODOS_ ? "todos los pagos" : _dbj;
    var _dbl = _dbd.esContra ? "propuesto bloquear" : "destacados";
    var _dbm = "Ranking " + _dbk + " — " + _dbi + " (" + _dbd.conBloqueo + " cliente" + (_dbd.conBloqueo === 1 ? "" : "s") + " con " + _dbd.umbralBloqueo + " o más)";
    var _dbn = "font-family:Segoe UI,Arial,sans-serif;";
    var _dbo = 'style="background:#1e293b;color:#ffffff;padding:7px 6px;font-weight:bold;' + _dbn + 'font-size:12.5px;"';
    var _dbp = function (_dbq) { return 'style="padding:7px 6px;border-bottom:1px solid #e2e8f0;' + _dbn + 'font-size:12.5px;color:#1e293b;' + (_dbq || '') + '"'; };
    var _dbr = { rojo: "background:#fee2e2;color:#991b1b;", ok: "background:#dcfce7;color:#166534;", aviso: "background:#fef3c7;color:#92400e;" };
    var _dbs = _dbf.map(function (_dbt, _dbu) {
        var _dbv = huchoSituacion_(_dbt, _dbd);
        var _dbw = _dbt.c >= _dbd.umbralBloqueo ? "background:#fee2e2;color:#dc2626;" : (_dbt.c >= _dbd.umbralVigilar ? "background:#fef3c7;color:#b45309;" : "");
        return '<tr>' +
            '<td ' + _dbp('font-weight:bold;') + '>' + (_dbu + 1) + '</td>' +
            '<td ' + _dbp() + '>' + huchoEsc_(_dbt.n) + '</td>' +
            '<td ' + _dbp() + '>' + huchoEsc_(_dbt.e || "—") + (_dbt.otros && _dbt.otros.length ? '<br><span style="color:#92400e;font-size:11px;">y ' + huchoEsc_(_dbt.otros.join(", ")) + '</span>' : '') + '</td>' +
            '<td ' + _dbp('text-align:center;font-weight:bold;' + _dbw) + '>' + _dbt.c + '</td>' +
            '<td ' + _dbp('text-align:center;') + '>' + _dbt.ant + '</td>' +
            '<td ' + _dbp(_dbr[_dbv.tipo] || '') + '>' + huchoEsc_(_dbv.texto) + '</td>' +
            '</tr>';
    }).join('');
    if (!_dbs)
        _dbs = '<tr><td colspan="6" ' + _dbp('text-align:center;color:#059669;') + '>Ningún cliente llega a ' + _dbe + ' pedido' + (_dbe === 1 ? '' : 's') + ' devuelto' + (_dbe === 1 ? '' : 's') + ' este mes.</td></tr>';
    var _dbx = '<table width="600" cellpadding="0" cellspacing="0" border="0" style="width:600px;border-collapse:collapse;border:1px solid #e2e8f0;' + _dbn + 'color:#1e293b;">' +
        '<tr><td style="background:#1B2636;padding:14px 16px;border-bottom:4px solid #E07A1F;">' +
        '<div style="color:#ffffff;font-size:18px;font-weight:bold;' + _dbn + '">RASTRO · Devoluciones</div>' +
        '<div style="color:#cbd5e1;font-size:12px;margin-top:2px;' + _dbn + '">Ranking mensual de pedidos devueltos · ' + huchoEsc_(_dbj) + '</div>' +
        '</td></tr>' +
        '<tr><td style="padding:12px 16px 4px;font-size:13px;' + _dbn + '">' +
        '<b>Periodo:</b> pedidos hechos en ' + huchoEsc_(_dbd.mesTxt) + ' (según la fecha de pedido).<br>' +
        '<b>Método de pago:</b> ' + huchoEsc_(_dbj) + '.<br>' +
        '<b>Resumen:</b> ' + _dbd.pedidos + ' pedidos devueltos · ' + _dbd.clientes + ' clientes · ' +
        '<span style="color:#dc2626;font-weight:bold;">' + _dbd.conBloqueo + ' con ' + _dbd.umbralBloqueo + ' o más (' + _dbl + ')</span>.' +
        '</td></tr>' +
        '<tr><td style="padding:10px 16px 6px;">' +
        '<table width="568" cellpadding="0" cellspacing="0" border="0" style="width:568px;border-collapse:collapse;">' +
        '<tr><td ' + _dbo + '>#</td><td ' + _dbo + '>Cliente</td><td ' + _dbo + '>Email</td><td ' + _dbo.replace('padding', 'text-align:center;padding') + '>Pedidos</td><td ' + _dbo.replace('padding', 'text-align:center;padding') + '>Mes ant.</td><td ' + _dbo + '>Situación</td></tr>' +
        _dbs +
        '</table>' +
        '</td></tr>' +
        '<tr><td style="padding:6px 16px 14px;font-size:11.5px;color:#64748b;' + _dbn + '">' +
        'Solo se incluyen clientes con ' + _dbe + ' o más pedidos devueltos en el mes' + (_dbh ? ' (los ' + HUCHO_MAX_CORREO_ + ' primeros)' : '') + '. ' +
        'Cada pedido cuenta una vez aunque llegue en varios bultos. «Reincidente» = también llegó a ' + _dbd.umbralVigilar + ' o más en los meses anteriores.' +
        '</td></tr>' +
        '</table>';
    var _dby = "Ranking " + _dbk + " — " + _dbi + "\n" + _dbd.pedidos + " pedidos · " + _dbd.clientes + " clientes · " + _dbd.conBloqueo + " con " + _dbd.umbralBloqueo + " o más\n\n" +
        _dbf.map(function (_dbz, _dca) { return (_dca + 1) + ". " + _dbz.n + " (" + (_dbz.e || "sin email") + "): " + _dbz.c + " pedidos (mes ant. " + _dbz.ant + ") - " + huchoSituacion_(_dbz, _dbd).texto; }).join("\n");
    var _dcb = huchoCorreoUsuario_();
    return { asunto: _dbm, para: HUCHO_PARA_, cc: HUCHO_CC_, yo: _dcb, html: _dbx, texto: _dby, filas: _dbf.length, mesTxt: _dbd.mesTxt, metodo: _dbd.metodo, enviado: _dbd.enviado, mes: _dbd.mes };
}
function huchoMarcarPreparado(_dcc, _dcd, _dce) {
    huchoExigir_(_dcc);
    var _dcf = Math.max(-24, Math.min(0, Math.round(Number(_dcd) || 0)));
    var _dcg = String(_dce || "CONTRA").trim().toUpperCase().substring(0, 40) || "CONTRA";
    var _dch = huchoIndiceMes_(new Date()) + _dcf, _dci = new Date();
    var _dcj = { f: huchoHoy_(), h: huchoDosDigitos_(_dci.getHours()) + ":" + huchoDosDigitos_(_dci.getMinutes()), u: huchoCorreoUsuario_() };
    PropertiesService.getDocumentProperties().setProperty("PCN_ENV|" + huchoClaveEnvio_(_dch, _dcg), JSON.stringify(_dcj));
    return _dcj;
}
function huchoSvgCandado_(_dck) {
    return '<svg class="' + _dck + '" viewBox="0 0 20 24" aria-hidden="true">' +
        '<path class="pcn-arco" d="M5 11 V7 a5 5 0 0 1 10 0 V11" fill="none" stroke="#1e293b" stroke-width="2.4" stroke-linecap="round"/>' +
        '<rect x="2" y="10" width="16" height="12" rx="3" fill="#E07A1F"/>' +
        '<circle cx="10" cy="15.5" r="1.8" fill="#fff"/><rect x="9.2" y="16" width="1.6" height="3" rx=".8" fill="#fff"/></svg>';
}
var HUCHO_SVG_N_ = 0;
function huchoSvgCerdito_(_dcl) {
    var _dcm = "pcn" + (HUCHO_SVG_N_++);
    var _dcn = function (_dco) { return 'url(#' + _dcm + _dco + ')'; };
    return '<svg class="pcn-pig ' + (_dcl || '') + '" viewBox="0 0 120 130" aria-hidden="true">' +
        '<defs>' +
        '<radialGradient id="' + _dcm + 'p" cx="38%" cy="28%" r="78%"><stop offset="0%" stop-color="#ffe0e7"/><stop offset="50%" stop-color="#fbb0c2"/><stop offset="100%" stop-color="#e2718e"/></radialGradient>' +
        '<radialGradient id="' + _dcm + 'h" cx="40%" cy="30%" r="80%"><stop offset="0%" stop-color="#ffc0cf"/><stop offset="100%" stop-color="#e06a88"/></radialGradient>' +
        '<radialGradient id="' + _dcm + 'o" cx="34%" cy="30%" r="75%"><stop offset="0%" stop-color="#3a3a3a"/><stop offset="60%" stop-color="#141414"/><stop offset="100%" stop-color="#000"/></radialGradient>' +
        '<radialGradient id="' + _dcm + 'm" cx="50%" cy="40%" r="60%"><stop offset="0%" stop-color="#e8354f" stop-opacity=".5"/><stop offset="100%" stop-color="#e8354f" stop-opacity="0"/></radialGradient>' +
        '<linearGradient id="' + _dcm + 'd" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#ffe89a"/><stop offset="55%" stop-color="#f5c542"/><stop offset="100%" stop-color="#cf961a"/></linearGradient>' +
        '<linearGradient id="' + _dcm + 'l" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#e13a52"/><stop offset="100%" stop-color="#8f1626"/></linearGradient>' +
        '<linearGradient id="' + _dcm + 'v" x1="0" y1="0" x2="0" y2="1"><stop offset="0%" stop-color="#34d399" stop-opacity=".85"/><stop offset="100%" stop-color="#047857" stop-opacity=".9"/></linearGradient>' +
        '<radialGradient id="' + _dcm + 'b" cx="35%" cy="30%" r="70%"><stop offset="0%" stop-color="#fff" stop-opacity=".95"/><stop offset="100%" stop-color="#bfdbfe" stop-opacity=".45"/></radialGradient>' +
        '</defs>' +
        '<ellipse cx="60" cy="126" rx="30" ry="4" fill="#000" opacity=".12"/>' +
        '<g class="pcn-todo"><g class="pcn-cuerpo">' +
        '<path class="pcn-rabito" d="M92 98 q9 -4 7 -11 q-2 -5 -6 -2 q-3 3 2 5" fill="none" stroke="#e2718e" stroke-width="3" stroke-linecap="round"/>' +
        '<ellipse cx="60" cy="104" rx="31" ry="19" fill="' + _dcn('p') + '"/>' +
        '<rect x="39" y="113" width="11" height="11" rx="4.5" fill="#e2718e"/><rect x="70" y="113" width="11" height="11" rx="4.5" fill="#e2718e"/>' +
        '<rect x="39" y="120" width="11" height="4" rx="2" fill="#a3485f"/><rect x="70" y="120" width="11" height="4" rx="2" fill="#a3485f"/>' +
        '<path d="M60 90 L48 83 L48 97 Z" fill="' + _dcn('l') + '"/><path d="M60 90 L72 83 L72 97 Z" fill="' + _dcn('l') + '"/><circle cx="60" cy="90" r="3.2" fill="#8f1626"/><circle cx="52" cy="87" r="1.6" fill="#fff" opacity=".45"/>' +
        '<g class="pcn-oreja-i"><path d="M30 42 L22 14 Q36 18 47 32 Z" fill="' + _dcn('p') + '"/><path d="M31 37 L27 21 Q36 24 41 32 Z" fill="#e06a88"/></g>' +
        '<g class="pcn-oreja-d"><path d="M90 42 L98 14 Q84 18 73 32 Z" fill="' + _dcn('p') + '"/><path d="M89 37 L93 21 Q84 24 79 32 Z" fill="#e06a88"/></g>' +
        '<circle cx="60" cy="62" r="34" fill="' + _dcn('p') + '"/>' +
        '<ellipse cx="48" cy="40" rx="11" ry="5" fill="#fff" opacity=".35" transform="rotate(-18 48 40)"/>' +
        '<rect x="51" y="29" width="18" height="4.2" rx="2.1" fill="#7a2c41"/>' +
        '<path d="M27 50 Q60 36 93 50 L95 45 Q60 28 25 45 Z" fill="' + _dcn('v') + '"/>' +
        '<path d="M28 49 Q60 36 92 49" fill="none" stroke="#fff" stroke-width="1" opacity=".5"/>' +
        '<circle cx="35" cy="73" r="8" fill="' + _dcn('m') + '"/><circle cx="85" cy="73" r="8" fill="' + _dcn('m') + '"/>' +
        '<g class="pcn-o-cerrados" fill="none" stroke="#141414" stroke-width="2.6" stroke-linecap="round"><path d="M40 60 Q46 65 52 60"/><path d="M68 60 Q74 65 80 60"/></g>' +
        '<g class="pcn-o-abiertos"><circle cx="46" cy="59" r="5.4" fill="' + _dcn('o') + '"/><circle cx="74" cy="59" r="5.4" fill="' + _dcn('o') + '"/><circle cx="47.9" cy="57" r="1.9" fill="#fff"/><circle cx="75.9" cy="57" r="1.9" fill="#fff"/><circle cx="44.7" cy="61" r=".8" fill="#fff" opacity=".8"/><circle cx="72.7" cy="61" r=".8" fill="#fff" opacity=".8"/></g>' +
        '<g class="pcn-o-felices" fill="none" stroke="#141414" stroke-width="2.6" stroke-linecap="round"><path d="M41 60 Q46 54 51 60"/><path d="M69 60 Q74 54 79 60"/></g>' +
        '<g class="pcn-o-susto"><circle cx="46" cy="58" r="8" fill="#fff" stroke="#141414" stroke-width="1.5"/><circle cx="74" cy="58" r="8" fill="#fff" stroke="#141414" stroke-width="1.5"/><circle cx="46" cy="58" r="2.6" fill="#141414"/><circle cx="74" cy="58" r="2.6" fill="#141414"/></g>' +
        '<g class="pcn-hocico"><ellipse cx="60" cy="75" rx="14.5" ry="10.5" fill="' + _dcn('h') + '"/><ellipse cx="55" cy="75" rx="2.7" ry="3.8" fill="#8a2a44"/><ellipse cx="65" cy="75" rx="2.7" ry="3.8" fill="#8a2a44"/><ellipse cx="56" cy="69.5" rx="5" ry="1.8" fill="#fff" opacity=".35"/></g>' +
        '<path class="pcn-boca-sueno" d="M56 89 Q60 91 64 89" fill="none" stroke="#8a2a44" stroke-width="1.8" stroke-linecap="round"/>' +
        '<path class="pcn-boca-feliz" d="M53 88 Q60 93 67 88" fill="none" stroke="#8a2a44" stroke-width="2" stroke-linecap="round"/>' +
        '<ellipse class="pcn-boca-o" cx="60" cy="90" rx="3.4" ry="4" fill="#8a2a44"/>' +
        '<circle class="pcn-burbuja" cx="71" cy="80" r="7" fill="' + _dcn('b') + '" stroke="#93c5fd" stroke-width=".8"/>' +
        '</g>' +
        '<g class="pcn-moneda"><circle cx="60" cy="22" r="7" fill="' + _dcn('d') + '" stroke="#b7811a" stroke-width="1"/><text x="60" y="25.5" font-size="9" font-weight="bold" fill="#8a5a00" text-anchor="middle" font-family="Arial">€</text></g>' +
        '<circle class="pcn-estallido" cx="60" cy="31" r="14" fill="none" stroke="#f5c542" stroke-width="3"/>' +
        '<g class="pcn-salto"><circle cx="60" cy="28" r="6" fill="' + _dcn('d') + '" stroke="#b7811a"/></g>' +
        '<g class="pcn-salto pcn-m2"><circle cx="60" cy="28" r="6" fill="' + _dcn('d') + '" stroke="#b7811a"/></g>' +
        '<g class="pcn-salto pcn-m3"><circle cx="60" cy="28" r="5" fill="' + _dcn('d') + '" stroke="#b7811a"/></g>' +
        '<g class="pcn-exclam"><path d="M100 4 L108 4 L106 24 L102 24 Z" fill="#E07A1F"/><circle cx="104" cy="30" r="3.2" fill="#E07A1F"/></g>' +
        '</g>' +
        '<g class="pcn-zz" font-family="Georgia,serif" font-weight="bold" fill="#94a3b8"><text x="92" y="30" font-size="12">z</text><text x="100" y="20" font-size="10">z</text><text x="106" y="11" font-size="8">z</text></g>' +
        '</svg>';
}
function huchoCssCerdito_() {
    return '' +
        '.pcn-pig{overflow:visible}' +
        '.pcn-pig .pcn-o-abiertos,.pcn-pig .pcn-o-felices,.pcn-pig .pcn-o-susto,.pcn-pig .pcn-exclam,.pcn-pig .pcn-boca-o,.pcn-pig .pcn-boca-feliz,.pcn-pig .pcn-moneda,.pcn-pig .pcn-salto,.pcn-pig .pcn-estallido{display:none}' +
        '.pcn-pig.pcn-despierto .pcn-o-cerrados,.pcn-pig.pcn-despierto .pcn-burbuja,.pcn-pig.pcn-despierto .pcn-zz,.pcn-pig.pcn-despierto .pcn-boca-sueno{display:none}' +
        '.pcn-pig.pcn-despierto .pcn-o-abiertos,.pcn-pig.pcn-despierto .pcn-boca-feliz,.pcn-pig.pcn-despierto .pcn-o-felices{display:inline}' +
        '.pcn-pig.pcn-susto .pcn-o-cerrados,.pcn-pig.pcn-susto .pcn-burbuja,.pcn-pig.pcn-susto .pcn-zz,.pcn-pig.pcn-susto .pcn-boca-sueno{display:none}' +
        '.pcn-pig.pcn-susto .pcn-o-susto,.pcn-pig.pcn-susto .pcn-exclam,.pcn-pig.pcn-susto .pcn-boca-o,.pcn-pig.pcn-susto .pcn-salto,.pcn-pig.pcn-susto .pcn-estallido{display:inline}' +
        '.pcn-pig .pcn-cuerpo{transform-box:view-box;transform-origin:60px 128px;animation:pcnRespira 3.2s ease-in-out infinite}' +
        '@keyframes pcnRespira{0%,100%{transform:scale(1,1)}50%{transform:scale(1.025,1.045)}}' +
        '.pcn-pig .pcn-burbuja{transform-box:fill-box;transform-origin:10% 20%;animation:pcnBurbuja 3.2s ease-in-out infinite}' +
        '@keyframes pcnBurbuja{0%,100%{transform:scale(.25)}50%{transform:scale(1)}}' +
        '.pcn-pig .pcn-zz text{transform-box:fill-box;opacity:0;animation:pcnZz 3.2s ease-in-out infinite}' +
        '.pcn-pig .pcn-zz text:nth-child(2){animation-delay:1s}.pcn-pig .pcn-zz text:nth-child(3){animation-delay:2s}' +
        '@keyframes pcnZz{0%{opacity:0;transform:translate(0,0)}25%{opacity:1}100%{opacity:0;transform:translate(10px,-18px)}}' +
        '.pcn-pig .pcn-rabito{transform-box:fill-box;transform-origin:0% 50%;animation:pcnRabito 2.4s ease-in-out infinite}' +
        '@keyframes pcnRabito{0%,100%{transform:rotate(-6deg)}50%{transform:rotate(8deg)}}' +
        '.pcn-pig.pcn-despierto .pcn-rabito{animation-duration:.9s}' +
        '.pcn-pig.pcn-despierto .pcn-cuerpo{animation:pcnBalanceo 3.6s ease-in-out infinite}' +
        '@keyframes pcnBalanceo{0%,100%{transform:rotate(-2deg)}50%{transform:rotate(2deg)}}' +
        '.pcn-pig.pcn-despierto .pcn-oreja-i{transform-box:fill-box;transform-origin:80% 90%;animation:pcnOrejaI 3.6s ease-in-out infinite}' +
        '.pcn-pig.pcn-despierto .pcn-oreja-d{transform-box:fill-box;transform-origin:20% 90%;animation:pcnOrejaD 3.6s ease-in-out infinite}' +
        '@keyframes pcnOrejaI{0%,70%,100%{transform:rotate(0)}76%{transform:rotate(-14deg)}82%{transform:rotate(4deg)}88%{transform:rotate(-8deg)}}' +
        '@keyframes pcnOrejaD{0%,70%,100%{transform:rotate(0)}76%{transform:rotate(14deg)}82%{transform:rotate(-4deg)}88%{transform:rotate(8deg)}}' +
        '.pcn-pig.pcn-despierto .pcn-hocico{transform-box:fill-box;transform-origin:center;animation:pcnHocico 2.2s ease-in-out infinite}' +
        '@keyframes pcnHocico{0%,60%,100%{transform:translateY(0) scale(1)}66%,78%{transform:translateY(-1.5px) scale(1.05,.95)}72%,84%{transform:translateY(0) scale(1)}}' +
        '.pcn-pig.pcn-despierto .pcn-moneda{display:inline;transform-box:fill-box;transform-origin:center;animation:pcnMoneda 3.6s cubic-bezier(.45,.05,.55,.95) infinite}' +
        '@keyframes pcnMoneda{0%,8%{transform:translate(-34px,40px) scale(0,1);opacity:0}12%{transform:translate(-34px,40px) scale(1,1);opacity:1}30%{transform:translate(-18px,-30px) scaleX(-1)}42%{transform:translate(-4px,-44px) scaleX(1)}54%{transform:translate(0,-24px) scaleX(-1)}62%{transform:translate(0,-4px) scaleX(.35);opacity:1}66%,100%{transform:translate(0,4px) scaleX(.35);opacity:0}}' +
        '.pcn-pig.pcn-despierto .pcn-o-abiertos{animation:pcnOjosA 3.6s linear infinite}' +
        '.pcn-pig.pcn-despierto .pcn-o-felices{opacity:0;animation:pcnOjosF 3.6s linear infinite}' +
        '@keyframes pcnOjosA{0%,63%{opacity:1}64%,80%{opacity:0}81%,100%{opacity:1}}' +
        '@keyframes pcnOjosF{0%,63%{opacity:0}64%,80%{opacity:1}81%,100%{opacity:0}}' +
        '.pcn-pig.pcn-susto .pcn-todo{transform-box:view-box;transform-origin:60px 128px;animation:pcnBote .75s cubic-bezier(.2,1.5,.4,1) both}' +
        '@keyframes pcnBote{0%{transform:translateY(0) scale(1.08,.9)}25%{transform:translateY(-22px) scale(.94,1.08) rotate(-4deg)}45%{transform:translateY(-22px) rotate(4deg)}60%{transform:translateY(-18px) rotate(-3deg)}100%{transform:translateY(0) scale(1)}}' +
        '.pcn-pig.pcn-susto .pcn-salto{transform-box:fill-box;transform-origin:center;animation:pcnVuela .8s ease-out both}' +
        '.pcn-pig.pcn-susto .pcn-m2{animation-name:pcnVuela2}.pcn-pig.pcn-susto .pcn-m3{animation-name:pcnVuela3}' +
        '@keyframes pcnVuela{0%{transform:translate(0,0) rotate(0);opacity:1}100%{transform:translate(-46px,-44px) rotate(-260deg);opacity:0}}' +
        '@keyframes pcnVuela2{0%{transform:translate(0,0) rotate(0);opacity:1}100%{transform:translate(44px,-50px) rotate(300deg);opacity:0}}' +
        '@keyframes pcnVuela3{0%{transform:translate(0,0);opacity:1}100%{transform:translate(4px,-66px) rotate(180deg);opacity:0}}' +
        '.pcn-pig.pcn-susto .pcn-exclam{transform-box:fill-box;transform-origin:50% 100%;animation:pcnPop .35s cubic-bezier(.3,1.8,.5,1) both}' +
        '@keyframes pcnPop{from{transform:scale(0)}to{transform:scale(1)}}' +
        '.pcn-pig.pcn-susto .pcn-estallido{transform-box:fill-box;transform-origin:center;animation:pcnEstalla .4s ease-out both}' +
        '@keyframes pcnEstalla{0%{transform:scale(.3);opacity:1}100%{transform:scale(1.6);opacity:0}}' +
        '@media (prefers-reduced-motion: reduce){.pcn-pig *{animation-duration:6s!important}}';
}
function huchoSvgVineta_() {
    return '<svg viewBox="0 0 270 150" aria-hidden="true">' +
        '<rect x="0" y="118" width="270" height="10" fill="#1e293b"/><g fill="#475569"><circle cx="15" cy="123" r="3"/><circle cx="55" cy="123" r="3"/><circle cx="95" cy="123" r="3"/><circle cx="135" cy="123" r="3"/><circle cx="175" cy="123" r="3"/><circle cx="215" cy="123" r="3"/><circle cx="255" cy="123" r="3"/></g>' +
        '<g class="pcn-v-caja"><g transform="translate(150 78)">' +
        '<rect x="0" y="0" width="56" height="40" rx="3" fill="#d6a15e" stroke="#111" stroke-width="2"/>' +
        '<rect x="22" y="0" width="12" height="40" fill="#E07A1F" opacity=".85"/>' +
        '<rect x="5" y="8" width="46" height="12" rx="2" fill="#fff" stroke="#111" stroke-width="1.2"/>' +
        '<text x="28" y="17.5" font-size="8.5" font-weight="900" text-anchor="middle" fill="#111" font-family="Arial">CONTRA</text>' +
        '<g class="pcn-v-marca"><rect x="3" y="23" width="50" height="13" rx="2" fill="#fff" fill-opacity=".92" stroke="#dc2626" stroke-width="2" transform="rotate(-8 28 29)"/>' +
        '<text x="28" y="32.5" font-size="7.6" font-weight="900" text-anchor="middle" fill="#dc2626" transform="rotate(-8 28 29)" font-family="Arial">BLOQUEADO</text></g>' +
        '</g></g>' +
        '<g transform="translate(18 38)">' +
        '<path d="M16 22 L10 2 Q22 6 30 16 Z" fill="#fbb0c2" stroke="#111" stroke-width="2"/>' +
        '<path d="M74 22 L80 2 Q68 6 60 16 Z" fill="#fbb0c2" stroke="#111" stroke-width="2"/>' +
        '<ellipse cx="45" cy="72" rx="28" ry="12" fill="#fbb0c2" stroke="#111" stroke-width="2"/>' +
        '<circle cx="45" cy="42" r="32" fill="#fbb0c2" stroke="#111" stroke-width="2.4"/>' +
        '<path d="M16 30 Q45 16 74 30 L75 25 Q45 10 15 25 Z" fill="#10b981" stroke="#111" stroke-width="1.6"/>' +
        '<g class="pcn-v-ceno"><path d="M26 34 L40 39" stroke="#111" stroke-width="3" stroke-linecap="round"/><path d="M64 34 L50 39" stroke="#111" stroke-width="3" stroke-linecap="round"/></g>' +
        '<circle cx="34" cy="44" r="4" fill="#111"/><circle cx="56" cy="44" r="4" fill="#111"/><circle cx="35.3" cy="42.6" r="1.3" fill="#fff"/><circle cx="57.3" cy="42.6" r="1.3" fill="#fff"/>' +
        '<ellipse cx="45" cy="56" rx="11" ry="8" fill="#f38aa4" stroke="#111" stroke-width="1.8"/><ellipse cx="41" cy="56" rx="2" ry="3" fill="#111"/><ellipse cx="49" cy="56" rx="2" ry="3" fill="#111"/>' +
        '<path d="M38 67 L52 67" stroke="#111" stroke-width="2.2" stroke-linecap="round"/>' +
        '<path d="M45 76 L37 71 L37 81 Z" fill="#E07A1F" stroke="#111" stroke-width="1.2"/><path d="M45 76 L53 71 L53 81 Z" fill="#E07A1F" stroke="#111" stroke-width="1.2"/>' +
        '</g>' +
        '<g class="pcn-v-brazo">' +
        '<path d="M96 96 Q140 58 170 52" fill="none" stroke="#111" stroke-width="10" stroke-linecap="round"/>' +
        '<path d="M96 96 Q140 58 170 52" fill="none" stroke="#fbb0c2" stroke-width="6.5" stroke-linecap="round"/>' +
        '<rect x="164" y="30" width="14" height="30" rx="5" fill="#7a2c41" stroke="#111" stroke-width="2"/>' +
        '<rect x="158" y="58" width="26" height="9" rx="2" fill="#555" stroke="#111" stroke-width="2"/>' +
        '<rect x="156" y="66" width="30" height="5" rx="1.5" fill="#dc2626" stroke="#111" stroke-width="1.5"/>' +
        '</g>' +
        '<g class="pcn-v-pum"><path d="M232 40 l7 -14 l3 13 l13 -7 l-6 13 l14 3 l-13 6 l8 12 l-14 -4 l-2 14 l-8 -11 l-9 10 l0 -14 l-14 1 l10 -10 l-10 -9 l14 -1 z" fill="#f5c542" stroke="#111" stroke-width="2"/>' +
        '<text x="236" y="54" font-size="11" font-weight="900" text-anchor="middle" fill="#E07A1F" font-family="Impact,Arial Black,Arial" transform="rotate(-10 236 50)">¡PUM!</text></g>' +
        '<g transform="translate(96 4)"><path d="M0 8 Q0 0 8 0 L92 0 Q100 0 100 8 L100 22 Q100 30 92 30 L22 30 L8 40 L12 30 L8 30 Q0 30 0 22 Z" fill="#fff" stroke="#111" stroke-width="2"/>' +
        '<text x="50" y="19.5" font-size="11" font-weight="800" text-anchor="middle" fill="#111" font-family="Comic Sans MS,Chalkboard SE,Segoe Print,sans-serif">¡Este no pasa!</text></g>' +
        '</svg>';
}
function revisarReglasFormato() { return lrfEjecutar_(false); }
function limpiarReglasFormato() { return lrfEjecutar_(true); }
function restaurarReglasFormato() { return lrfRestaurar_(); }
var LRF_HOJAS_TRABAJO_ = ['Entradas', 'Retornos'];
var LRF_COL_C_ = 3, LRF_COL_D_ = 4;
function lrfEjecutar_(_dcp) {
    var _dcq = SpreadsheetApp.getActiveSpreadsheet(), _dcr = [], _dcs = Date.now(), _dct = null;
    function _dcu(_dcv) { _dcr.push(_dcv); }
    _dcu('=== ' + (_dcp ? 'LIMPIEZA' : 'REVISIÓN (no cambia nada)') + ' de reglas de formato condicional ===');
    if (_dcp) {
        try {
            _dct = LockService.getDocumentLock();
            if (!_dct.tryLock(30000))
                _dct = null;
        }
        catch (_dcw) {
            _dct = null;
        }
        if (!_dct) {
            _dcu('El documento está ocupado. Prueba otra vez en un rato.');
            _dcr.forEach(function (_dcx) { console.log(_dcx); });
            return;
        }
    }
    var _dcy = 0, _dcz = 0, _dda = [];
    try {
        _dcq.getSheets().forEach(function (_ddb) {
            var _ddc = _ddb.getConditionalFormatRules();
            if (!_ddc.length)
                return;
            var _ddd = LRF_HOJAS_TRABAJO_.indexOf(_ddb.getName()) > -1;
            var _dde = lrfPlanHoja_(_ddb, _ddc, _ddd);
            _dcy += _ddc.length;
            _dcz += _dde.nuevas.length;
            if (!_dde.borrar.length && !_dde.cambios.length) {
                _dcu('\n«' + _ddb.getName() + '»: ' + _ddc.length + ' reglas, nada que limpiar.');
                return;
            }
            _dcu('\n«' + _ddb.getName() + '»: ' + _ddc.length + ' reglas → quedan ' + _dde.nuevas.length);
            _dde.borrar.forEach(function (_ddf) { _dcu('  ✂ quitar ' + _ddf.desc + '  [' + _ddf.motivo + ']'); });
            _dde.cambios.forEach(function (_ddg) { _dcu('  ✎ ' + _ddg); });
            _dde.avisos.forEach(function (_ddh) { _dcu('  ⓘ ' + _ddh); });
            if (_dcp)
                _dda.push({ sh: _ddb, reglas: _ddc, nuevas: _dde.nuevas });
        });
        if (_dcp && _dda.length) {
            var _ddi;
            try {
                _ddi = lrfGuardarCopia_(_dcq, _dda);
            }
            catch (_ddj) {
                _dcu('\n⛔ No se ha podido guardar la copia de seguridad (' + _ddj.message + '). No se ha cambiado nada.');
                _dcr.forEach(function (_ddk) { console.log(_ddk); });
                return _dcr.join('\n');
            }
            _dcu('\n💾 Copia guardada (' + _ddi + '). Para deshacer: «restaurarReglasFormato».');
            _dda.forEach(function (_ddl) { _ddl.sh.setConditionalFormatRules(_ddl.nuevas); _dcu('  ✔ «' + _ddl.sh.getName() + '» limpia'); });
        }
        if (_dcp && _dda.length && typeof actualizarReglaRepetidos_ === 'function') {
            try {
                actualizarReglaRepetidos_();
                _dcu('\nEntradas: ventana de la regla roja puesta al día (como cada noche).');
            }
            catch (_ddm) {
                _dcu('\nEntradas: no se pudo poner al día la ventana (' + _ddm.message + '); se hará esta noche.');
            }
        }
    }
    finally {
        if (_dct) {
            try {
                _dct.releaseLock();
            }
            catch (_ddn) { }
        }
    }
    _dcu('\nTotal: ' + _dcy + ' reglas → ' + _dcz + (_dcp ? '' : ' (si lo aplicas)') + ' · ' + Math.round((Date.now() - _dcs) / 1000) + ' s');
    if (!_dcp)
        _dcu('Si está bien, ejecuta «limpiarReglasFormato».');
    _dcr.forEach(function (_ddo) { console.log(_ddo); });
    return _dcr.join('\n');
}
function lrfPlanHoja_(_ddp, _ddq, _ddr) {
    var _dds = _ddq.map(function (_ddt, _ddu) { return lrfInfo_(_ddt, _ddu); });
    var _ddv = {}, _ddw = {}, _ddx = {}, _ddy = [], _ddz = [];
    function _dea(_deb, _dec) { _ddv[_deb.i] = true; _ddw[_deb.i] = _dec; }
    _dds.forEach(function (_ded) { if (_ded.rota)
        _dea(_ded, 'rota (#REF!), no pinta nada'); });
    if (_ddr) {
        var _dee = _dds.filter(function (_def) { return !_ddv[_def.i] && _def.esRepetidos; });
        var _deg = _dee.filter(function (_deh) { return _deh.celdasC > 0 && _deh.miraC; });
        if (_deg.length) {
            _deg.sort(function (_dei, _dej) {
                if (_ddp.getName() === 'Entradas' && _dei.todoC !== _dej.todoC)
                    return _dei.todoC ? -1 : 1;
                return _dej.celdasC - _dei.celdasC;
            });
            var _dek = _deg[0], _del = lrfUnir_(_dek.intC), _dem = [], _den = [];
            _dee.forEach(function (_deo) {
                if (_deo === _dek)
                    return;
                var _dep = lrfRestar_(_deo.intC, _del);
                if (!_dep.length) {
                    _dea(_deo, _deo.celdasC ? 'copia: la regla roja principal ya cubre esas celdas' : 'copia suelta fuera de la columna C (pinta celdas que no son pedidos)');
                    return;
                }
                if (_ddp.getName() === 'Entradas') {
                    _dea(_deo, 'copia vieja; deja de pintar ' + lrfTextoInt_(_dep) + ' (filas antiguas, fuera de la ventana de la regla principal)');
                }
                else {
                    _dea(_deo, 'copia; sus filas de la columna C (' + lrfTextoInt_(_dep) + ') pasan a la regla principal');
                    _dem = _dem.concat(_dep);
                    _del = lrfUnir_(_del.concat(_dep));
                    _den.push(_deo);
                }
            });
            var _deq = _dek.rangos.filter(function (_der) { return !lrfEnC_(_der); });
            if (_deq.length || _dem.length) {
                var _des = [];
                _dek.rangos.forEach(function (_det) {
                    if (lrfEnC_(_det))
                        _des.push(_det);
                    else if (_det.getColumn() <= LRF_COL_C_ && LRF_COL_C_ <= _det.getColumn() + _det.getNumColumns() - 1)
                        _des.push(_ddp.getRange(_det.getRow(), LRF_COL_C_, _det.getNumRows(), 1));
                });
                _des = _des.concat(_dem.map(function (_deu) { return _ddp.getRange(_deu[0], LRF_COL_C_, _deu[1] - _deu[0] + 1, 1); }));
                var _dev = _des.length && lrfEnC_(_dek.rangos[0]);
                var _dew = lrfEsquina_(_des) === lrfEsquina_(_dek.rangos);
                if (_dev && _dew) {
                    _ddx[_dek.i] = _dek.regla.copy().setRanges(_des).build();
                    if (_deq.length)
                        _ddy.push('regla roja principal: se le quitan ' + _deq.length + ' trozo(s) fuera de la columna C (' + lrfCorto_(_deq.map(function (_dex) { return _dex.getA1Notation(); }).join(', '), 80) + ')');
                    if (_dem.length)
                        _ddy.push('regla roja principal: se le añaden ' + lrfTextoInt_(_dem));
                }
                else {
                    _ddz.push('la regla roja principal no se puede recortar sin mover su referencia: se deja como está');
                    _den.forEach(function (_dey) { delete _ddv[_dey.i]; delete _ddw[_dey.i]; });
                }
            }
        }
        else if (_dee.length) {
            _dee.forEach(function (_dez) { if (!_dez.celdasC)
                _dea(_dez, 'copia suelta fuera de la columna C'); });
        }
        var _dfa = {};
        _dds.forEach(function (_dfb) { if (!_ddv[_dfb.i] && _dfb.clave)
            (_dfa[_dfb.clave] = _dfa[_dfb.clave] || []).push(_dfb); });
        Object.keys(_dfa).forEach(function (_dfc) {
            var _dfd = _dfa[_dfc];
            if (_dfd.length < 2)
                return;
            _dfd.sort(function (_dfe, _dff) { return _dff.celdas - _dfe.celdas; });
            var _dfg = lrfUnir_(_dfd[0].intD);
            _dfd.slice(1).forEach(function (_dfh) {
                if (_dfh.soloD && !lrfRestar_(_dfh.intD, _dfg).length)
                    _dea(_dfh, 'copia pequeña de «' + _dfh.textoCond + '»; la grande ya cubre esas celdas');
            });
        });
    }
    var _dfi = [];
    _dds.forEach(function (_dfj) { if (!_ddv[_dfj.i])
        _dfi.push(_ddx[_dfj.i] || _dfj.regla); });
    return {
        nuevas: _dfi, cambios: _ddy, avisos: _ddz,
        borrar: _dds.filter(function (_dfk) { return _ddv[_dfk.i]; }).map(function (_dfl) { return { desc: _dfl.desc, motivo: _ddw[_dfl.i] }; })
    };
}
function lrfInfo_(_dfm, _dfn) {
    var _dfo = _dfm.getRanges(), _dfp = _dfm.getBooleanCondition(), _dfq = _dfp ? String(_dfp.getCriteriaType()) : 'ESCALA';
    var _dfr = _dfp ? (_dfp.getCriteriaValues() || []).map(function (_dfs) { return (_dfs && _dfs.getA1Notation) ? _dfs.getA1Notation() : String(_dfs); }) : [];
    var _dft = _dfr.join(' | ');
    var _dfu = [], _dfv = [], _dfw = 0, _dfx = true, _dfy = true;
    _dfo.forEach(function (_dfz) {
        var _dga = _dfz.getColumn(), _dgb = _dga + _dfz.getNumColumns() - 1, _dgc = _dfz.getRow(), _dgd = _dgc + _dfz.getNumRows() - 1;
        _dfw += _dfz.getNumRows() * _dfz.getNumColumns();
        if (_dga <= LRF_COL_C_ && LRF_COL_C_ <= _dgb)
            _dfu.push([_dgc, _dgd]);
        if (_dga <= LRF_COL_D_ && LRF_COL_D_ <= _dgb)
            _dfv.push([_dgc, _dgd]);
        if (!(_dga === LRF_COL_D_ && _dgb === LRF_COL_D_))
            _dfx = false;
        if (!(_dga === LRF_COL_C_ && _dgb === LRF_COL_C_))
            _dfy = false;
    });
    var _dge = _dfu.reduce(function (_dgf, _dgg) { return _dgf + _dgg[1] - _dgg[0] + 1; }, 0);
    var _dgh = _dfq === 'CUSTOM_FORMULA';
    return {
        i: _dfn, regla: _dfm, rangos: _dfo, celdas: _dfw, celdasC: _dge, intC: _dfu, intD: _dfv, soloD: _dfx, todoC: _dfy,
        rota: _dft.indexOf('#REF!') > -1,
        esRepetidos: _dgh && /(CONTAR\.SI|COUNTIF)\s*\(/i.test(_dft) && />\s*1\s*\)?\s*$/.test(_dft),
        miraC: /(CONTAR\.SI|COUNTIF)\s*\(\s*C:C\s*[,;]/i.test(_dft),
        clave: (!_dgh && _dfp) ? _dfq + '|' + _dft : null,
        textoCond: _dft,
        desc: _dfo.length + ' rango(s) ' + lrfCorto_(_dfo.map(function (_dgi) { return _dgi.getA1Notation(); }).join(', '), 70) + ' → ' + (_dgh ? '' : _dfq + ' ') + lrfCorto_(_dft, 90)
    };
}
function lrfUnir_(_dgj) {
    var _dgk = _dgj.slice().sort(function (_dgl, _dgm) { return _dgl[0] - _dgm[0]; }), _dgn = [];
    _dgk.forEach(function (_dgo) {
        var _dgp = _dgn[_dgn.length - 1];
        if (_dgp && _dgo[0] <= _dgp[1] + 1)
            _dgp[1] = Math.max(_dgp[1], _dgo[1]);
        else
            _dgn.push([_dgo[0], _dgo[1]]);
    });
    return _dgn;
}
function lrfRestar_(_dgq, _dgr) {
    var _dgs = [];
    lrfUnir_(_dgq).forEach(function (_dgt) {
        var _dgu = _dgt[0];
        for (var _dgv = 0; _dgv < _dgr.length && _dgu <= _dgt[1]; _dgv++) {
            var _dgw = _dgr[_dgv];
            if (_dgw[1] < _dgu)
                continue;
            if (_dgw[0] > _dgt[1])
                break;
            if (_dgw[0] > _dgu)
                _dgs.push([_dgu, _dgw[0] - 1]);
            _dgu = Math.max(_dgu, _dgw[1] + 1);
        }
        if (_dgu <= _dgt[1])
            _dgs.push([_dgu, _dgt[1]]);
    });
    return _dgs;
}
function lrfTextoInt_(_dgx) {
    var _dgy = _dgx.reduce(function (_dgz, _dha) { return _dgz + _dha[1] - _dha[0] + 1; }, 0);
    return lrfCorto_(_dgx.map(function (_dhb) { return _dhb[0] === _dhb[1] ? 'C' + _dhb[0] : 'C' + _dhb[0] + ':C' + _dhb[1]; }).join(', '), 80) + ' (' + _dgy + ' celda' + (_dgy === 1 ? '' : 's') + ')';
}
function lrfEnC_(_dhc) { return _dhc.getColumn() === LRF_COL_C_ && _dhc.getNumColumns() === 1; }
function lrfEsquina_(_dhd) {
    var _dhe = Infinity, _dhf = Infinity;
    _dhd.forEach(function (_dhg) { _dhe = Math.min(_dhe, _dhg.getRow()); _dhf = Math.min(_dhf, _dhg.getColumn()); });
    return _dhe + ',' + _dhf;
}
function lrfCorto_(_dhh, _dhi) { _dhh = String(_dhh); return _dhh.length > _dhi ? _dhh.substring(0, _dhi - 1) + '…' : _dhh; }
var LRF_CLAVE_COPIA_ = 'LRF_COPIA_REGLAS';
function lrfGuardarCopia_(_dhj, _dhk) {
    var _dhl = { fecha: Utilities.formatDate(new Date(), _dhj.getSpreadsheetTimeZone(), 'dd/MM/yyyy HH:mm'), hojas: {} }, _dhm = 0;
    _dhk.forEach(function (_dhn) {
        var _dho = _dhn.reglas.map(lrfSerializar_);
        _dho.forEach(function (_dhp) { lrfConstruir_(_dhn.sh, _dhp); });
        _dhl.hojas[_dhn.sh.getName()] = _dho;
        _dhm += _dho.length;
    });
    var _dhq = PropertiesService.getDocumentProperties(), _dhr = JSON.stringify(_dhl), _dhs = 8000, _dht = Math.ceil(_dhr.length / _dhs);
    var _dhu = Number(_dhq.getProperty(LRF_CLAVE_COPIA_ + '_N') || 0);
    for (var _dhv = 0; _dhv < _dhu; _dhv++)
        _dhq.deleteProperty(LRF_CLAVE_COPIA_ + '_' + _dhv);
    var _dhw = {};
    for (var _dhx = 0; _dhx < _dht; _dhx++)
        _dhw[LRF_CLAVE_COPIA_ + '_' + _dhx] = _dhr.substring(_dhx * _dhs, (_dhx + 1) * _dhs);
    _dhw[LRF_CLAVE_COPIA_ + '_N'] = String(_dht);
    _dhq.setProperties(_dhw);
    if (JSON.stringify(lrfLeerCopia_()) !== _dhr)
        throw new Error('la copia no se ha guardado entera');
    return _dhm + ' reglas de ' + Object.keys(_dhl.hojas).length + ' pestaña(s), ' + _dhl.fecha;
}
function lrfLeerCopia_() {
    var _dhy = PropertiesService.getDocumentProperties(), _dhz = Number(_dhy.getProperty(LRF_CLAVE_COPIA_ + '_N') || 0), _dia = '';
    if (!_dhz)
        return null;
    for (var _dib = 0; _dib < _dhz; _dib++)
        _dia += _dhy.getProperty(LRF_CLAVE_COPIA_ + '_' + _dib) || '';
    return JSON.parse(_dia);
}
function lrfRestaurar_() {
    var _dic = SpreadsheetApp.getActiveSpreadsheet(), _did = [], _die = lrfLeerCopia_(), _dif = null;
    function _dig() { _did.forEach(function (_dih) { console.log(_dih); }); return _did.join('\n'); }
    if (!_die) {
        _did.push('No hay ninguna copia guardada (aún no se ha limpiado nada).');
        return _dig();
    }
    try {
        _dif = LockService.getDocumentLock();
        if (!_dif.tryLock(30000))
            _dif = null;
    }
    catch (_dii) {
        _dif = null;
    }
    if (!_dif) {
        _did.push('El documento está ocupado. Prueba otra vez en un rato.');
        return _dig();
    }
    try {
        _did.push('=== RESTAURAR reglas de formato condicional (copia del ' + _die.fecha + ') ===');
        Object.keys(_die.hojas).forEach(function (_dij) {
            var _dik = _dic.getSheetByName(_dij);
            if (!_dik) {
                _did.push('«' + _dij + '»: ya no existe, se salta');
                return;
            }
            var _dil = _die.hojas[_dij].map(function (_dim) { return lrfConstruir_(_dik, _dim); });
            _dik.setConditionalFormatRules(_dil);
            _did.push('«' + _dij + '»: ' + _dil.length + ' reglas restauradas');
        });
        _did.push('Hecho. La copia se conserva por si hiciera falta otra vez.');
    }
    finally {
        try {
            _dif.releaseLock();
        }
        catch (_din) { }
    }
    return _dig();
}
function lrfColor_(_dio, _dip) {
    try {
        if (_dio && _dio.asRgbColor)
            return _dio.asRgbColor().asHexString();
    }
    catch (_diq) { }
    try {
        return _dip ? _dip() : null;
    }
    catch (_dir) {
        return null;
    }
}
function lrfSerializar_(_dis) {
    var _dit = { g: _dis.getRanges().map(function (_diu) { return _diu.getA1Notation(); }) }, _div = _dis.getBooleanCondition();
    if (_div) {
        _dit.t = String(_div.getCriteriaType());
        _dit.v = (_div.getCriteriaValues() || []).map(function (_diw) {
            if (_diw === null || typeof _diw === 'string' || typeof _diw === 'number' || typeof _diw === 'boolean')
                return _diw;
            if (_diw instanceof Date)
                return { d: _diw.getTime() };
            return { e: String(_diw) };
        });
        _dit.bg = lrfColor_(_div.getBackgroundObject && _div.getBackgroundObject(), _div.getBackground && function () { return _div.getBackground(); });
        _dit.fc = lrfColor_(_div.getFontColorObject && _div.getFontColorObject(), _div.getFontColor && function () { return _div.getFontColor(); });
        _dit.b = _div.getBold();
        _dit.i = _div.getItalic();
        _dit.s = _div.getStrikethrough();
        _dit.u = _div.getUnderline();
        return _dit;
    }
    var _dix = _dis.getGradientCondition();
    if (!_dix)
        throw new Error('regla de un tipo desconocido en ' + _dit.g.join(','));
    _dit.t = 'GRADIENTE';
    _dit.min = [lrfColor_(_dix.getMinColorObject && _dix.getMinColorObject(), _dix.getMinColor && function () { return _dix.getMinColor(); }), String(_dix.getMinType()), _dix.getMinValue()];
    _dit.max = [lrfColor_(_dix.getMaxColorObject && _dix.getMaxColorObject(), _dix.getMaxColor && function () { return _dix.getMaxColor(); }), String(_dix.getMaxType()), _dix.getMaxValue()];
    if (_dix.getMidType && _dix.getMidType())
        _dit.mid = [lrfColor_(_dix.getMidColorObject && _dix.getMidColorObject(), _dix.getMidColor && function () { return _dix.getMidColor(); }), String(_dix.getMidType()), _dix.getMidValue()];
    return _dit;
}
function lrfConstruir_(_diy, _diz) {
    var _dja = SpreadsheetApp.newConditionalFormatRule().setRanges(_diz.g.map(function (_djb) { return _diy.getRange(_djb); }));
    if (_diz.t === 'GRADIENTE') {
        function _djc(_djd, _dje, _djf) {
            if (!_djd)
                return;
            var _djg = SpreadsheetApp.InterpolationType[_djd[1]];
            if (_djd[1] === 'MIN' || _djd[1] === 'MAX')
                _djf.call(_dja, _djd[0]);
            else
                _dje.call(_dja, _djd[0], _djg, String(_djd[2]));
        }
        _djc(_diz.min, _dja.setGradientMinpointWithValue, _dja.setGradientMinpoint);
        _djc(_diz.mid, _dja.setGradientMidpointWithValue, function () { });
        _djc(_diz.max, _dja.setGradientMaxpointWithValue, _dja.setGradientMaxpoint);
    }
    else {
        var _djh = _diz.v.map(function (_dji) {
            if (_dji && typeof _dji === 'object' && _dji.d !== undefined)
                return new Date(_dji.d);
            if (_dji && typeof _dji === 'object' && _dji.e !== undefined)
                return SpreadsheetApp.RelativeDate[_dji.e];
            return _dji;
        });
        _dja.withCriteria(SpreadsheetApp.BooleanCriteria[_diz.t], _djh);
        if (_diz.bg)
            _dja.setBackground(_diz.bg);
        if (_diz.fc)
            _dja.setFontColor(_diz.fc);
        if (_diz.b !== null && _diz.b !== undefined)
            _dja.setBold(_diz.b);
        if (_diz.i !== null && _diz.i !== undefined)
            _dja.setItalic(_diz.i);
        if (_diz.s !== null && _diz.s !== undefined)
            _dja.setStrikethrough(_diz.s);
        if (_diz.u !== null && _diz.u !== undefined)
            _dja.setUnderline(_diz.u);
    }
    return _dja.build();
}
