const HOJAS = {
  OPERADORES: 'CONFIG_OPERADORES',
  TELEFONOS: 'CONFIG_TELEFONOS',
  PCS: 'CONFIG_PC',
  FICHAJES: 'FICHAJES',
  RECEPCION: 'RECEPCION_TURNO',
  BANOS: 'BAÑO_PAUSAS',
  FUMAR: 'FUMAR_PAUSAS',
  LIMPIEZA: 'LIMPIEZA_PAUSAS',
  NOVEDADES: 'ERRORES_NOVEDADES',
  TAREAS: 'TAREAS_PC',
  COMENTARIOS: 'NOVEDADES_COMENTARIOS',
  NOTAS_TURNO: 'NOTAS_TURNO'
};

const TURNOS_CONFIG = {
  TM: { recepcion: '05:55', tardeDesde: '06:00', salida: '13:55', tolerancia: 5 },
  TT: { recepcion: '13:55', tardeDesde: '14:00', salida: '21:55', tolerancia: 5 },
  TN: { recepcion: '21:55', tardeDesde: '22:00', salida: '05:55', tolerancia: 5 }
};

const HORAS_MAX_FICHAJE = 9;

function doGet(e) {
  try {
    const datos = normalizarDatos_(e.parameter || {});
    return manejarAccion_(datos);
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    const datos = normalizarDatos_(body);
    return manejarAccion_(datos);
  } catch (err) {
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function manejarAccion_(datos) {
  if (datos.action === 'getconfig') return json_(getConfig_());
  if (datos.action === 'login') return json_(login_(datos));
  if (datos.action === 'ficharingreso') return json_(ficharIngreso_(datos));
  if (datos.action === 'ficharsalida') return json_(ficharSalida_(datos));
  if (datos.action === 'banoinicio') return json_(banoInicio_(datos));
  if (datos.action === 'banofin') return json_(banoFin_(datos));
  if (datos.action === 'fumarinicio') return json_(fumarInicio_(datos));
  if (datos.action === 'fumarfin') return json_(fumarFin_(datos));
  if (datos.action === 'recepcionturno') return json_(recepcionTurno_(datos));
  if (datos.action === 'novedad' || datos.action === 'reportarnovedad') return json_(reportarNovedad_(datos));
  if (datos.action === 'getnovedadespendientes') return json_(getNovedadesPendientes_(datos));
  if (datos.action === 'actualizarnovedad') return json_(actualizarNovedad_(datos));
  if (datos.action === 'cambiarpin') return json_(cambiarPin_(datos));
  if (datos.action === 'getpcs') return json_(getPCs_());
  if (datos.action === 'getlimpiezahoy') return json_(getLimpiezaHoy_());
  if (datos.action === 'gettelefonospc') return json_(getTelefonosPC_(datos));
  if (datos.action === 'getmispausas') return json_(getMisPausasAbiertas_(datos));
  if (datos.action === 'limpiezainicio') return json_(limpiezaInicio_(datos));
  if (datos.action === 'limpiezafin') return json_(limpiezaFin_(datos));
  if (datos.action === 'verificaractivo') return json_(verificarActivo_(datos));
  if (datos.action === 'gettareaspc') return json_(getTareasPC_(datos));
  if (datos.action === 'completartarea') return json_(completarTarea_(datos));
  if (datos.action === 'telefonocaido') return json_(telefonoCaido_(datos));
  if (datos.action === 'telefonolevantado') return json_(telefonoLevantado_(datos));
  if (datos.action === 'getlogtelefono') return json_(getLogTelefono_(datos));
  if (datos.action === 'notarnovedad') return json_(notarNovedad_(datos));
  if (datos.action === 'getpoll') return json_(getPoll_(datos));
  if (datos.action === 'getmisnovedades') return json_(getMisNovedades_(datos));
  if (datos.action === 'getcomentariosnovedad') return json_(getComentariosNovedad_(datos));
  if (datos.action === 'agregarcomentario') return json_(agregarComentario_(datos));
  if (datos.action === 'editarcomentario') return json_(editarComentario_(datos));
  if (datos.action === 'eliminarcomentario') return json_(eliminarComentario_(datos));
  if (datos.action === 'guardarnotaturno') return json_(guardarNotaTurno_(datos));
  if (datos.action === 'getnotaturnoanterior') return json_(getNotaTurnoAnterior_(datos));
  if (datos.action === 'getnotasturno') return json_(adminGuard_(datos, getNotasTurno_));

  // === ADMIN ===
  if (datos.action === 'loginadmin') return json_(loginAdmin_(datos));
  if (datos.action === 'getdashboard') return json_(adminGuard_(datos, getDashboard_));
  if (datos.action === 'getoperadoresadmin') return json_(adminGuard_(datos, getOperadoresAdmin_));
  if (datos.action === 'crearoperador') return json_(adminGuard_(datos, crearOperador_));
  if (datos.action === 'editaroperador') return json_(adminGuard_(datos, editarOperador_));
  if (datos.action === 'eliminaroperador') return json_(adminGuard_(datos, eliminarOperador_));
  if (datos.action === 'setactivoperador') return json_(adminGuard_(datos, setActivoOperador_));
  if (datos.action === 'setpinoperador') return json_(adminGuard_(datos, setPinOperador_));
  if (datos.action === 'gettelefonosadmin') return json_(adminGuard_(datos, getTelefonosAdmin_));
  if (datos.action === 'creartelefono') return json_(adminGuard_(datos, crearTelefono_));
  if (datos.action === 'editartelefono') return json_(adminGuard_(datos, editarTelefono_));
  if (datos.action === 'eliminartelefono') return json_(adminGuard_(datos, eliminarTelefono_));
  if (datos.action === 'creartarea') return json_(adminGuard_(datos, crearTarea_));
  if (datos.action === 'eliminartarea') return json_(adminGuard_(datos, eliminarTareaAdmin_));
  if (datos.action === 'gettareasadmin') return json_(adminGuard_(datos, getTareasAdmin_));
  if (datos.action === 'getnovedadesadmin') return json_(adminGuard_(datos, getNovedadesAdmin_));
  if (datos.action === 'eliminarnovedad') return json_(adminGuard_(datos, eliminarNovedad_));
  if (datos.action === 'getpcsadmin') return json_(adminGuard_(datos, getPCsAdmin_));
  if (datos.action === 'crearpc') return json_(adminGuard_(datos, crearPC_));
  if (datos.action === 'editarpc') return json_(adminGuard_(datos, editarPC_));
  if (datos.action === 'eliminarpc') return json_(adminGuard_(datos, eliminarPC_));
  if (datos.action === 'getfichajes') return json_(adminGuard_(datos, getFichajes_));
  if (datos.action === 'getpausasadmin') return json_(adminGuard_(datos, getPausasAdmin_));
  if (datos.action === 'getrecepciones') return json_(adminGuard_(datos, getRecepciones_));
  if (datos.action === 'getlogtelefonosadmin') return json_(adminGuard_(datos, getLogTelefonosAdmin_));
  if (datos.action === 'editarfichaje') return json_(adminGuard_(datos, editarFichaje_));
  if (datos.action === 'eliminarfichaje') return json_(adminGuard_(datos, eliminarFichaje_));
  if (datos.action === 'editarpausaadmin') return json_(adminGuard_(datos, editarPausaAdmin_));
  if (datos.action === 'eliminarpausa') return json_(adminGuard_(datos, eliminarPausa_));
  if (datos.action === 'editarrecepcion') return json_(adminGuard_(datos, editarRecepcion_));
  if (datos.action === 'eliminarrecepcion') return json_(adminGuard_(datos, eliminarRecepcion_));
  if (datos.action === 'editarnotaturno') return json_(adminGuard_(datos, editarNotaTurno_));
  if (datos.action === 'eliminarnotaturno') return json_(adminGuard_(datos, eliminarNotaTurno_));
  if (datos.action === 'editarlogtel') return json_(adminGuard_(datos, editarLogTel_));
  if (datos.action === 'eliminarlogtel') return json_(adminGuard_(datos, eliminarLogTel_));
  if (datos.action === 'editarnovedadcontenido') return json_(adminGuard_(datos, editarNovedadContenido_));
  if (datos.action === 'editartarea') return json_(adminGuard_(datos, editarTareaAdmin_));

  return json_({
    ok: true,
    mensaje: 'API activa',
    actionRecibida: datos.action || ''
  });
}

function normalizarDatos_(data) {
  return {
    action: String(data.action || '').toLowerCase().replace(/ñ/g, 'n'),

    nombre: data.nombre || '',
    pin: data.pin || '',
    pc: data.pc || '',

    recibe: data.recibe || data.nombre || '',
    entrega: data.entrega || '',
    puestoOrden: data.puestoOrden || '',
    whatsappOk: data.whatsappOk || '',
    cajaOk: data.cajaOk || '',
    diferenciaCaja: data.diferenciaCaja || data.detalleCaja || '',
    revinculacionOk: data.revinculacionOk || '',
    telefonosCaidos: data.telefonosCaidos || '',
    limpiezaCumplida: data.limpiezaCumplida || '',
    observaciones: data.observaciones || '',

    tipo: data.tipo || '',
    prioridad: data.prioridad || '',
    detalle: data.detalle || '',

    pinActual: data.pinActual || '',
    pinNuevo: data.pinNuevo || '',

    pcAsignada: data.pcAsignada || '',
    modoCovertura: data.modoCovertura === true || data.modoCovertura === 'true',

    linea: data.linea || '',
    motivo: data.motivo || '',
    nota: data.nota || '',
    tareaId: data.tareaId || data.fila || '',
    fila: data.fila || 0,

    // admin
    nombreOriginal: data.nombreOriginal || '',
    pinNuevo: data.pinNuevo || '',
    activo: typeof data.activo === 'boolean' ? data.activo : (data.activo === undefined ? null : String(data.activo)),
    turno: data.turno || '',
    rol: data.rol || '',
    tareaTexto: data.tareaTexto || '',
    lineaOriginal: data.lineaOriginal || '',
    pcOriginal: data.pcOriginal || '',
    estadoEsperado: data.estadoEsperado || '',
    activoEnSistema: data.activoEnSistema || '',
    numero: data.numero || '',
    pcNuevo: data.pcNuevo || '',
    orden: data.orden || 0,
    estadoPC: data.estadoPC || '',
    desde: data.desde || '',
    hasta: data.hasta || '',
    operador: data.operador || '',
    tipoPausa: data.tipoPausa || '',
    limit: parseInt(data.limit || 200),
    imagenBase64: data.imagenBase64 || '',
    accion: data.accion || '',
    hora: data.hora || '',
    minutosTarde: data.minutosTarde || '',
    horaSalida: data.horaSalida || '',
    horaRegreso: data.horaRegreso || '',
    minutos: data.minutos || '',
    excedio: data.excedio || '',
    horaCaida: data.horaCaida || '',
    horaLevantado: data.horaLevantado || '',
    duracion: data.duracion || '',
    operadorCaida: data.operadorCaida || '',
    operadorLevanta: data.operadorLevanta || '',
    fechaTexto: data.fechaTexto || data.fecha || '',

    id: data.id || '',
    estado: data.estado || '',
    revisadoPor: data.revisadoPor || data.revisadopor || data.nombre || '',
    resolucion: data.resolucion || data.detalle || ''
  };
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function hoja_(nombre) {
  const sh = SpreadsheetApp.getActive().getSheetByName(nombre);
  if (!sh) throw new Error('No existe la hoja: ' + nombre);
  return sh;
}

function filas_(nombreHoja) {
  const sh = hoja_(nombreHoja);
  const values = sh.getDataRange().getValues();

  if (values.length < 4) return [];

  const headers = values[2];
  const dataRows = values.slice(3);

  return dataRows
    .filter(r => r.join('') !== '')
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => {
        if (h) obj[String(h).trim()] = r[i];
      });
      return obj;
    });
}

function esSi_(valor) {
  const s = String(valor || '').trim().toLowerCase();
  return s === 'sí' || s === 'si';
}

function getConfig_() {
  const operadores = filas_(HOJAS.OPERADORES).filter(o => esSi_(o.Activo));
  const turnos = Object.keys(TURNOS_CONFIG).map(t => ({
    Turno: t,
    'Hora recepción': TURNOS_CONFIG[t].recepcion,
    'Tarde desde': TURNOS_CONFIG[t].tardeDesde,
    'Hora salida': TURNOS_CONFIG[t].salida,
    'Tolerancia min': TURNOS_CONFIG[t].tolerancia
  }));
  const telefonos = filas_(HOJAS.TELEFONOS).filter(t => esSi_(t['Activo en sistema']));

  return { ok: true, operadores, turnos, telefonos };
}

function getOperador_(nombre) {
  const operadores = filas_(HOJAS.OPERADORES);

  return operadores.find(o =>
    String(o.Nombre || '').trim().toLowerCase() === String(nombre || '').trim().toLowerCase()
    && esSi_(o.Activo)
  );
}

function getTurno_(turno) {
  const t = String(turno || '').trim().toUpperCase();
  const cfg = TURNOS_CONFIG[t];
  if (!cfg) return null;
  return {
    Turno: t,
    'Hora recepción': cfg.recepcion,
    'Tarde desde': cfg.tardeDesde,
    'Hora salida': cfg.salida,
    'Tolerancia min': cfg.tolerancia
  };
}

function esFranquero_(op) {
  return String(op && op.Tipo || '').trim().toLowerCase().indexOf('franq') !== -1;
}

function obtenerSesionActiva_(nombreOperador) {
  const sh = hoja_(HOJAS.FICHAJES);
  const values = sh.getDataRange().getValues();
  if (values.length < 4) return null;

  const nombreLower = String(nombreOperador || '').trim().toLowerCase();
  const ahora = new Date();
  const limiteMs = HORAS_MAX_FICHAJE * 3600 * 1000;

  // Find the most recent entry for this operator
  for (let i = values.length - 1; i >= 3; i--) {
    const ts = values[i][0];
    const nombre = String(values[i][2] || '').trim().toLowerCase();
    const accion = values[i][6];

    if (nombre !== nombreLower) continue;

    // Found the most recent entry for this operator
    if (accion === 'Ingreso' && ts instanceof Date) {
      // Check if it's within the 9-hour window
      if (ahora - ts <= limiteMs) {
        return {
          pc: values[i][4],
          hora: values[i][7],
          minutos: Math.max(0, Math.ceil((ahora - ts) / 60000))
        };
      }
    }

    break; // We found the most recent entry, stop looking
  }

  return null;
}

function login_(datos) {
  const op = getOperador_(datos.nombre);

  if (!op) return { ok: false, error: 'Operador no activo o no encontrado.' };
  if (String(op.PIN) !== String(datos.pin)) return { ok: false, error: 'PIN incorrecto.' };

  if (esFranquero_(op) && !datos.pcAsignada) {
    return { ok: true, requierePcSelector: true };
  }

  const cierres = cerrarSesionesExpiradas_();

  // Check dual login only for fixed operators (franqueros can have multiple sessions)
  if (!esFranquero_(op)) {
    const sesionActiva = obtenerSesionActiva_(op.Nombre);
    if (sesionActiva) {
      if (!datos.modoCovertura) {
        return {
          ok: true,
          requiereConfirmacionCobertura: true,
          sesionActiva: { pc: sesionActiva.pc, hora: sesionActiva.hora }
        };
      }
      // Coverage mode confirmed but PC not selected yet
      if (!datos.pcAsignada) {
        return { ok: true, requierePcSelector: true, modoCovertura: true };
      }
    }
  }

  const pcFinal = (esFranquero_(op) || datos.modoCovertura)
    ? String(datos.pcAsignada).trim().toUpperCase()
    : op.PC;

  return {
    ok: true,
    operador: {
      nombre: op.Nombre,
      turno: op.Turno,
      pc: pcFinal,
      tipo: op.Tipo,
      esFranquero: esFranquero_(op),
      modoCovertura: datos.modoCovertura === true
    },
    cierresForzosos: cierres
  };
}

function cerrarSesionesExpiradas_() {
  const sh = hoja_(HOJAS.FICHAJES);
  const values = sh.getDataRange().getValues();
  if (values.length < 4) return [];

  const ahora = new Date();
  const limiteMs = HORAS_MAX_FICHAJE * 3600 * 1000;

  const ultimosPorOperador = {};

  for (let i = 3; i < values.length; i++) {
    const ts = values[i][0];
    const nombre = values[i][2];
    const accion = values[i][6];

    if (!nombre) continue;

    const key = String(nombre).trim().toLowerCase();
    if (!ultimosPorOperador[key] || ts > ultimosPorOperador[key].timestamp) {
      ultimosPorOperador[key] = {
        rowIdx: i,
        accion: accion,
        timestamp: ts,
        nombre: nombre,
        turno: values[i][3],
        pc: values[i][4],
        tipo: values[i][5]
      };
    }
  }

  const cerrados = [];

  for (const key in ultimosPorOperador) {
    const ult = ultimosPorOperador[key];
    if (ult.accion === 'Ingreso' && ult.timestamp instanceof Date) {
      if (ahora - ult.timestamp > limiteMs) {
        const pausasCerradas = cerrarPausasAbiertasDe_(ult.nombre);

        sh.appendRow([
          new Date(),
          fecha_(ahora),
          ult.nombre,
          ult.turno || '',
          ult.pc || '',
          ult.tipo || '',
          'Salida',
          hora_(ahora),
          'Salida forzosa',
          '',
          'Turno expirado'
            + (pausasCerradas.length ? ' · cierre auto de pausas: ' + pausasCerradas.join(', ') : '')
        ]);
        cerrados.push(ult.nombre);
      }
    }
  }

  if (cerrados.length) SpreadsheetApp.flush();
  return cerrados;
}

function ficharIngreso_(datos) {
  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

  const turnoConfig = getTurno_(op.Turno);
  if (!turnoConfig) return { ok: false, error: 'No se encontró configuración del turno: ' + op.Turno };

  const ahora = new Date();
  const tardeDesde = horaDelDia_(ahora, turnoConfig['Tarde desde']);
  const estado = ahora >= tardeDesde ? 'Tarde' : 'A tiempo';
  const minutosTarde = estado === 'Tarde' ? Math.max(0, Math.ceil((ahora - tardeDesde) / 60000)) : 0;

  hoja_(HOJAS.FICHAJES).appendRow([
    new Date(),
    fecha_(ahora),
    op.Nombre,
    op.Turno,
    op.PC,
    op.Tipo,
    'Ingreso',
    hora_(ahora),
    estado,
    minutosTarde,
    datos.observaciones || ''
  ]);

  return { ok: true, accion: 'Ingreso', hora: hora_(ahora), estado, minutosTarde };
}

function ficharSalida_(datos) {
  // Buscar operador SIN filtrar por activo (para que pueda salir aún si fue desactivado mid-sesión)
  const todos = filas_(HOJAS.OPERADORES);
  const nombreLower = String(datos.nombre || '').trim().toLowerCase();
  const op = todos.find(o => String(o.Nombre || '').trim().toLowerCase() === nombreLower);

  if (!op) {
    // Aún si no se encuentra, escribir la salida con datos mínimos para no dejar la sesión abierta
    const ahora = new Date();
    hoja_(HOJAS.FICHAJES).appendRow([
      new Date(), fecha_(ahora), datos.nombre || '', '', '', '',
      'Salida', hora_(ahora), 'Salida OK (operador no resuelto)', '',
      datos.observaciones || ''
    ]);
    return { ok: true, accion: 'Salida', hora: hora_(ahora), pausasCerradas: [] };
  }

  const cerradas = cerrarPausasAbiertasDe_(op.Nombre);
  const ahora = new Date();

  hoja_(HOJAS.FICHAJES).appendRow([
    new Date(),
    fecha_(ahora),
    op.Nombre,
    op.Turno,
    op.PC,
    op.Tipo,
    'Salida',
    hora_(ahora),
    'Salida OK',
    '',
    cerradas.length
      ? 'Salida + cierre auto de pausas: ' + cerradas.join(', ')
      : (datos.observaciones || '')
  ]);

  return { ok: true, accion: 'Salida', hora: hora_(ahora), pausasCerradas: cerradas };
}

function cerrarPausasAbiertasDe_(nombreOperador) {
  const cerradas = [];
  const nombreLower = String(nombreOperador || '').trim().toLowerCase();
  if (!nombreLower) return cerradas;

  const hojas = [
    { nombre: HOJAS.BANOS,    label: 'baño' },
    { nombre: HOJAS.FUMAR,    label: 'fumar' },
    { nombre: HOJAS.LIMPIEZA, label: 'limpieza' }
  ];

  hojas.forEach(h => {
    let sh;
    try { sh = hoja_(h.nombre); } catch (e) { return; }

    const values = sh.getDataRange().getValues();
    const ahora = new Date();
    const horaRegreso = hora_(ahora);

    for (let i = values.length - 1; i >= 3; i--) {
      const op = String(values[i][2] || '').trim().toLowerCase();
      const estado = values[i][9];
      const tsInicio = values[i][0];
      const horaSalida = values[i][5];

      if (op !== nombreLower || estado !== 'Abierta') continue;

      let minutos = 0;
      if (tsInicio instanceof Date) {
        minutos = Math.max(0, Math.ceil((ahora - tsInicio) / 60000));
      } else {
        minutos = calcularMinutosEntreHoras_(String(horaSalida), horaRegreso);
      }
      const excedio = minutos > 10 ? 'Sí' : 'No';

      sh.getRange(i + 1, 7).setValue(horaRegreso);
      sh.getRange(i + 1, 8).setValue(minutos);
      sh.getRange(i + 1, 9).setValue(excedio);
      sh.getRange(i + 1, 10).setValue('Cerrada (auto)');

      cerradas.push(h.label);
    }
  });

  if (cerradas.length) SpreadsheetApp.flush();
  return cerradas;
}

function banoInicio_(datos) {
  return iniciarPausa_(datos, HOJAS.BANOS, 'baño', { bloqueoCruzado: true });
}

function banoFin_(datos) {
  return finalizarPausa_(datos, HOJAS.BANOS, 'baño');
}

function fumarInicio_(datos) {
  return iniciarPausa_(datos, HOJAS.FUMAR, 'fumar', { bloqueoCruzado: true });
}

function fumarFin_(datos) {
  return finalizarPausa_(datos, HOJAS.FUMAR, 'fumar');
}

function limpiezaInicio_(datos) {
  return iniciarPausa_(datos, HOJAS.LIMPIEZA, 'limpieza', { bloqueoCruzado: false });
}

function limpiezaFin_(datos) {
  return finalizarPausa_(datos, HOJAS.LIMPIEZA, 'limpieza');
}

// Revisa BAÑO y FUMAR (todas las hojas con bloqueo cruzado) buscando pausas
// abiertas de otros operadores. Auto-cierra pausas de operadores sin sesión activa.
// Devuelve el primer bloqueo real encontrado, o null si está libre.
function buscarPausaBloqueo_(nombreOperadorPropio) {
  const hojasBloqueo = [
    { nombre: HOJAS.BANOS, label: 'baño' },
    { nombre: HOJAS.FUMAR, label: 'fumar' }
  ];

  const propioLower = String(nombreOperadorPropio || '').trim().toLowerCase();

  for (var h = 0; h < hojasBloqueo.length; h++) {
    var entry = hojasBloqueo[h];
    var sh;
    try { sh = hoja_(entry.nombre); } catch (e) { continue; }

    var values = sh.getDataRange().getValues();
    var huboFlush = false;

    for (var i = values.length - 1; i >= 3; i--) {
      var operadorRow = values[i][2];
      var pcRow       = values[i][4];
      var tsInicio    = values[i][0];
      var estado      = values[i][9];

      if (estado !== 'Abierta') continue;
      if (String(operadorRow || '').trim().toLowerCase() === propioLower) continue;

      var mins = tsInicio instanceof Date
        ? Math.max(0, Math.floor((new Date() - tsInicio) / 60000))
        : 0;

      // Si el bloqueador ya no tiene sesión activa → cierre automático
      if (!obtenerSesionActiva_(String(operadorRow || ''))) {
        sh.getRange(i + 1, 7).setValue(hora_(new Date()));
        sh.getRange(i + 1, 8).setValue(mins);
        sh.getRange(i + 1, 9).setValue('Sí');
        sh.getRange(i + 1, 10).setValue('Cerrada (turno finalizado)');
        huboFlush = true;
        continue;
      }

      if (huboFlush) SpreadsheetApp.flush();
      return { operador: operadorRow, pc: pcRow, minutos: mins, tipo: entry.label };
    }

    if (huboFlush) SpreadsheetApp.flush();
  }

  return null;
}

function iniciarPausa_(datos, hojaNombre, tipoPausa, opciones) {
  opciones = opciones || {};
  const bloqueoCruzado = opciones.bloqueoCruzado === true;

  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

  const pcOperador = String(datos.pcAsignada || op.PC || '').trim();

  // Verificar si el mismo operador ya tiene esta pausa abierta
  const sh = hoja_(hojaNombre);
  const values = sh.getDataRange().getValues();
  for (var i = values.length - 1; i >= 3; i--) {
    if (values[i][9] !== 'Abierta') continue;
    if (String(values[i][2] || '').trim().toLowerCase() === String(op.Nombre || '').trim().toLowerCase()) {
      return { ok: false, error: 'Ya tenés una pausa abierta de ' + tipoPausa + '.' };
    }
  }

  // Bloqueo cruzado: revisar todas las hojas bloqueantes
  if (bloqueoCruzado) {
    const bloqueo = buscarPausaBloqueo_(op.Nombre);
    if (bloqueo) {
      return {
        ok: false,
        bloqueoCruzado: true,
        bloqueador: { operador: bloqueo.operador, pc: bloqueo.pc, minutos: bloqueo.minutos, tipo: bloqueo.tipo },
        error:
          'PC ' + bloqueo.pc + ' tiene una pausa de ' + bloqueo.tipo + ' abierta hace ' + bloqueo.minutos + ' min ('
          + bloqueo.operador + '). Esperá a que regrese y la cierre antes de salir. '
          + 'Si ya terminó, avisale a ' + bloqueo.operador + ' que cierre la pausa desde su PC.'
      };
    }
  }

  const ahora = new Date();

  sh.appendRow([
    new Date(),
    fecha_(ahora),
    op.Nombre,
    op.Turno,
    pcOperador,
    hora_(ahora),
    '',
    '',
    '',
    'Abierta',
    datos.observaciones || ''
  ]);

  return {
    ok: true,
    accion: tipoPausa + ' inicio',
    inicio: hora_(ahora),
    operador: op.Nombre
  };
}

function finalizarPausa_(datos, hojaNombre, tipoPausa) {
  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

  const sh = hoja_(hojaNombre);
  const values = sh.getDataRange().getValues();
  const ahora = new Date();
  const horaRegreso = hora_(ahora);

  for (let i = values.length - 1; i >= 3; i--) {
    const timestampInicio = values[i][0];
    const operador = values[i][2];
    const horaSalida = values[i][5];
    const estado = values[i][9];

    if (
      String(operador || '').trim().toLowerCase() === String(op.Nombre || '').trim().toLowerCase()
      && estado === 'Abierta'
    ) {
      let minutos = 0;

      if (timestampInicio instanceof Date) {
        minutos = Math.max(0, Math.ceil((ahora - timestampInicio) / 60000));
      } else {
        minutos = calcularMinutosEntreHoras_(String(horaSalida), horaRegreso);
      }

      const excedio = minutos > 10 ? 'Sí' : 'No';

      sh.getRange(i + 1, 7).setValue(horaRegreso);
      sh.getRange(i + 1, 8).setValue(minutos);
      sh.getRange(i + 1, 9).setValue(excedio);
      sh.getRange(i + 1, 10).setValue('Cerrada');

      return {
        ok: true,
        accion: tipoPausa + ' fin',
        salida: horaSalida,
        regreso: horaRegreso,
        minutos,
        excedio
      };
    }
  }

  return { ok: false, error: 'No hay pausa abierta para este operador.' };
}

function recepcionTurno_(datos) {
  const op = getOperador_(datos.recibe || datos.nombre);
  if (!op) return { ok: false, error: 'Operador recibe no encontrado.' };

  const ahora = new Date();

  hoja_(HOJAS.RECEPCION).appendRow([
    new Date(),
    fecha_(ahora),
    op.Turno,
    op.PC,
    datos.recibe || op.Nombre,
    datos.entrega || '',
    siNo_(datos.puestoOrden),
    siNo_(datos.whatsappOk),
    siNo_(datos.cajaOk),
    datos.diferenciaCaja || '',
    siNo_(datos.revinculacionOk),
    datos.telefonosCaidos || '',
    datos.limpiezaCumplida || 'No corresponde',
    datos.observaciones || ''
  ]);

  return { ok: true, accion: 'Recepción registrada' };
}

function reportarNovedad_(datos) {
  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

  const ahora = new Date();
  const id = generarIdNovedad_();

  hoja_(HOJAS.NOVEDADES).appendRow([
    new Date(),
    fecha_(ahora),
    op.Nombre,
    op.Turno,
    op.PC,
    datos.tipo || 'Otro',
    datos.prioridad || 'Media',
    datos.detalle || '',
    'No',
    datos.observaciones || '',
    id,
    'Pendiente',
    '',
    ''
  ]);

  return { ok: true, accion: 'Novedad registrada', id, estado: 'Pendiente' };
}

function getNovedadesPendientes_(datos) {
  const sh = hoja_(HOJAS.NOVEDADES);
  const values = sh.getDataRange().getValues();
  const pcFiltro = String(datos.pc || '').trim().toLowerCase();
  const novedades = [];

  for (let i = 3; i < values.length; i++) {
    const row = values[i];

    const id = row[10];
    const estado = row[11];
    const pc = row[4];

    if (!id) continue;

    const estadoNorm = String(estado || '').trim().toLowerCase();
    const pcNorm = String(pc || '').trim().toLowerCase();

    const estaPendiente = estadoNorm === 'pendiente' || estadoNorm === 'en seguimiento';
    const correspondePC = !pcFiltro || pcNorm === pcFiltro;

    if (estaPendiente && correspondePC) {
      novedades.push({
        fila: i + 1,
        timestamp: row[0],
        fecha: row[1],
        operador: row[2],
        turno: row[3],
        pc: row[4],
        tipo: row[5],
        prioridad: row[6],
        detalle: row[7],
        resuelto: row[8],
        observaciones: row[9],
        id: row[10],
        estado: row[11],
        revisadoPor: row[12],
        resolucion: row[13]
      });
    }
  }

  return { ok: true, cantidad: novedades.length, novedades };
}

function actualizarNovedad_(datos) {
  const idBuscado = String(datos.id || '').trim();
  if (!idBuscado) return { ok: false, error: 'Falta ID de novedad.' };

  const nuevoEstado = datos.estado || 'En seguimiento';
  const revisadoPor = datos.revisadoPor || datos.nombre || '';
  const resolucion = datos.resolucion || datos.detalle || '';

  const sh = hoja_(HOJAS.NOVEDADES);
  const values = sh.getDataRange().getValues();

  for (let i = 3; i < values.length; i++) {
    const id = String(values[i][10] || '').trim();

    if (id === idBuscado) {
      sh.getRange(i + 1, 12).setValue(nuevoEstado);
      sh.getRange(i + 1, 13).setValue(revisadoPor);
      sh.getRange(i + 1, 14).setValue(resolucion);

      if (String(nuevoEstado).trim().toLowerCase() === 'resuelta') {
        sh.getRange(i + 1, 9).setValue('Sí');
      }

      return { ok: true, accion: 'Novedad actualizada', id: idBuscado, estado: nuevoEstado };
    }
  }

  return { ok: false, error: 'No se encontró la novedad con ID: ' + idBuscado };
}

function getPCs_() {
  const filas = filas_(HOJAS.PCS);
  const pcs = filas
    .filter(p => String(p.Estado || '').trim().toLowerCase() === 'activa')
    .map(p => ({
      pc: String(p.PC || '').trim(),
      orden: Number(p['Orden rotación llmpieza'] || p['Orden rotación limpieza'] || p['Orden rotacion limpieza'] || 999)
    }))
    .filter(p => p.pc)
    .sort((a, b) => a.orden - b.orden);
  return { ok: true, pcs };
}

function getLimpiezaHoy_() {
  const r = getPCs_();
  if (!r.pcs.length) return { ok: true, pc: null };

  const ahora = new Date();
  const dia = Math.floor(
    (Date.UTC(ahora.getFullYear(), ahora.getMonth(), ahora.getDate())) / (24 * 3600 * 1000)
  );
  const n = r.pcs.length;
  const hoy    = r.pcs[((dia     % n) + n) % n];
  const manana = r.pcs[((dia + 1 % n) + n) % n];

  return {
    ok:      true,
    pc:      hoy.pc,
    fecha:   fecha_(ahora),
    manana:  manana.pc,
    rotacion: r.pcs.map(p => p.pc)
  };
}

function getTelefonosPC_(datos) {
  const pcFiltro = String(datos.pc || '').trim().toLowerCase();
  if (!pcFiltro) return { ok: true, telefonos: [] };

  const telefonos = filas_(HOJAS.TELEFONOS)
    .filter(t => esSi_(t['Activo en sistema']))
    .filter(t => String(t.PC || '').trim().toLowerCase() === pcFiltro)
    .map(t => ({
      linea: t.Linea,
      tipo: t.Tipo,
      numero: t.Notas
    }));

  return { ok: true, telefonos };
}

function getMisPausasAbiertas_(datos) {
  const nombre = String(datos.nombre || '').trim().toLowerCase();
  const result = { ok: true, bano: false, fumar: false, limpieza: false };

  if (!nombre) return result;

  result.bano     = tienePausaAbierta_(HOJAS.BANOS, nombre);
  result.fumar    = tienePausaAbierta_(HOJAS.FUMAR, nombre);
  result.limpieza = tienePausaAbierta_(HOJAS.LIMPIEZA, nombre);

  return result;
}

function tienePausaAbierta_(hojaNombre, nombreLower) {
  try {
    const sh = hoja_(hojaNombre);
    const values = sh.getDataRange().getValues();
    for (let i = values.length - 1; i >= 3; i--) {
      const op = String(values[i][2] || '').trim().toLowerCase();
      const estado = values[i][9];
      if (op === nombreLower && estado === 'Abierta') return true;
    }
  } catch (e) {}
  return false;
}

function cambiarPin_(datos) {
  const nombre = String(datos.nombre || '').trim();
  const pinActual = String(datos.pinActual || datos.pin || '').trim();
  const pinNuevo = String(datos.pinNuevo || '').trim();

  if (!pinNuevo) return { ok: false, error: 'El PIN nuevo no puede estar vacío.' };

  const sh = hoja_(HOJAS.OPERADORES);
  const values = sh.getDataRange().getValues();

  if (values.length < 4) return { ok: false, error: 'No hay operadores configurados.' };

  const headers = values[2];
  const nombreIdx = headers.findIndex(h => String(h).trim().toLowerCase() === 'nombre');
  const pinIdx = headers.findIndex(h => String(h).trim().toLowerCase() === 'pin');
  const activoIdx = headers.findIndex(h => String(h).trim().toLowerCase() === 'activo');

  if (nombreIdx === -1 || pinIdx === -1) return {
    ok: false,
    error: 'Estructura de hoja incorrecta.',
    debug: { headers: headers.map(h => String(h)) }
  };

  for (let i = 3; i < values.length; i++) {
    const row = values[i];
    const nombreFila = String(row[nombreIdx] || '').trim().toLowerCase();
    const activo = activoIdx !== -1 ? esSi_(row[activoIdx]) : true;

    if (nombreFila === nombre.toLowerCase() && activo) {
      const pinEnSheet = String(row[pinIdx]);
      if (pinEnSheet !== pinActual) {
        return { ok: false, error: 'PIN actual incorrecto.', debug: { pinEnSheet, pinActual } };
      }

      const celda = sh.getRange(i + 1, pinIdx + 1);
      celda.setValue(pinNuevo);
      SpreadsheetApp.flush();

      const verificar = String(celda.getValue());
      if (verificar !== pinNuevo) {
        return {
          ok: false,
          error: 'No se persistió el cambio en la hoja.',
          debug: { fila: i + 1, columna: pinIdx + 1, escrito: pinNuevo, leido: verificar }
        };
      }

      return { ok: true, debug: { fila: i + 1, columna: pinIdx + 1, nuevoValor: verificar } };
    }
  }

  return {
    ok: false,
    error: 'Operador no encontrado.',
    debug: { nombre, nombresBuscados: values.slice(3).map(r => String(r[nombreIdx] || '')) }
  };
}

function generarIdNovedad_() {
  return 'NV-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
}

function siNo_(v) {
  if (v === true) return 'Sí';
  if (v === false) return 'No';

  const s = String(v || '').trim().toLowerCase();

  if (s === 'si' || s === 'sí') return 'Sí';
  if (s === 'no') return 'No';

  return v || '';
}

function fecha_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function hora_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm:ss');
}

function horaDelDia_(base, hhmm) {
  const [h, m] = String(hhmm || '00:00').split(':').map(Number);
  return new Date(base.getFullYear(), base.getMonth(), base.getDate(), h || 0, m || 0, 0, 0);
}

function calcularMinutosEntreHoras_(horaSalida, horaRegreso) {
  const s1 = horaASegundos_(horaSalida);
  const s2 = horaASegundos_(horaRegreso);

  let diff = s2 - s1;

  if (diff < 0) diff += 24 * 3600;

  return Math.max(0, Math.ceil(diff / 60));
}

function horaASegundos_(hhmmss) {
  const partes = String(hhmmss || '00:00:00').split(':').map(Number);

  const h = partes[0] || 0;
  const m = partes[1] || 0;
  const s = partes[2] || 0;

  return h * 3600 + m * 60 + s;
}

// --- VERIFICAR OPERADOR ACTIVO ---

function verificarActivo_(datos) {
  const nombre = String(datos.nombre || '').trim().toLowerCase();
  if (!nombre) return { ok: true, activo: false };

  const todos = filas_(HOJAS.OPERADORES);
  const op = todos.find(o => String(o.Nombre || '').trim().toLowerCase() === nombre);
  if (!op) return { ok: true, activo: false };
  return { ok: true, activo: esSi_(op.Activo) };
}

// --- POLL CONSOLIDADO (1 fetch en lugar de 3) ---
function getPoll_(datos) {
  const tareasResp = getTareasPC_({ pc: datos.pc });
  const activoResp = verificarActivo_({ nombre: datos.nombre });
  const novResp    = getNovedadesPendientes_({ pc: datos.pc });
  const misResp    = getMisNovedades_({ nombre: datos.nombre });

  return {
    ok: true,
    activo: activoResp.activo,
    novedades: novResp.novedades || [],
    tareas: tareasResp.tareas || [],
    ultimaCompletada: tareasResp.ultimaCompletada || null,
    misNovedades: misResp.novedades || []
  };
}

// --- HISTORIAL DE NOVEDADES PROPIAS DEL OPERADOR ---
function getMisNovedades_(datos) {
  const sh = hoja_(HOJAS.NOVEDADES);
  const v = sh.getDataRange().getValues();
  const nombreF = String(datos.nombre || '').trim().toLowerCase();
  if (!nombreF) return { ok: true, novedades: [] };

  const out = [];
  for (let i = v.length - 1; i >= 3; i--) {
    if (!v[i][10]) continue;
    if (String(v[i][2] || '').trim().toLowerCase() !== nombreF) continue;
    out.push({
      timestamp: v[i][0], fecha: v[i][1], operador: v[i][2], turno: v[i][3], pc: v[i][4],
      tipo: v[i][5], prioridad: v[i][6], detalle: v[i][7], resuelto: v[i][8],
      observaciones: v[i][9], id: v[i][10], estado: v[i][11], revisadoPor: v[i][12], resolucion: v[i][13]
    });
    if (out.length >= 30) break;
  }
  return { ok: true, cantidad: out.length, novedades: out };
}

// --- TAREAS DE SUPERVISOR ---

function inicializarTareasPC_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJAS.TAREAS);
  if (!sh) sh = ss.insertSheet(HOJAS.TAREAS);

  const headerVal = String(sh.getRange(3, 1).getValue()).toLowerCase();
  if (headerVal === 'timestamp') return sh;

  if (sh.getLastRow() < 1) sh.getRange(1, 1).setValue('TAREAS_PC');
  sh.getRange(3, 1, 1, 7).setValues([[
    'Timestamp', 'Fecha', 'PC', 'Tarea', 'Estado', 'FechaCompletada', 'CompletoOperador'
  ]]);
  return sh;
}

function getTareasPC_(datos) {
  const pcFiltro = String(datos.pc || '').trim().toLowerCase();
  let sh;
  try { sh = inicializarTareasPC_(); } catch (e) { return { ok: true, tareas: [], ultimaCompletada: null }; }

  const values = sh.getDataRange().getValues();
  if (values.length < 4) return { ok: true, tareas: [], ultimaCompletada: null };

  const tareas = [];
  let ultimaCompletada = null;

  for (let i = 3; i < values.length; i++) {
    const row = values[i];
    const pc = String(row[2] || '').trim().toLowerCase();
    if (!row[3] || pc !== pcFiltro) continue;

    const estado = String(row[4] || '').trim().toLowerCase();
    const t = {
      fila: i + 1,
      fecha: row[1],
      pc: row[2],
      tarea: row[3],
      estado: row[4],
      fechaCompletada: row[5],
      completoPor: row[6],
      timestamp: row[0]
    };

    if (estado === 'completada') {
      if (!ultimaCompletada || t.timestamp > ultimaCompletada.timestamp) ultimaCompletada = t;
    } else {
      tareas.push(t);
    }
  }

  return { ok: true, tareas, ultimaCompletada };
}

function completarTarea_(datos) {
  const fila = parseInt(datos.tareaId || '0');
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };

  let sh;
  try { sh = inicializarTareasPC_(); } catch (e) { return { ok: false, error: e.message }; }

  const ahora = new Date();
  const tareaTexto = String(sh.getRange(fila, 4).getValue());
  sh.getRange(fila, 5).setValue('Completada');
  sh.getRange(fila, 6).setValue(fecha_(ahora) + ' ' + hora_(ahora));
  sh.getRange(fila, 7).setValue(datos.nombre || '');
  SpreadsheetApp.flush();

  return { ok: true, tarea: tareaTexto, fecha: fecha_(ahora) + ' ' + hora_(ahora) };
}

// --- LOG TELÉFONOS (en CONFIG_TELEFONOS desde col H) ---
// Cols H-Q (1-indexed 8-17): Timestamp | PC | Linea | HoraCaida | OperadorCaida | Motivo | HoraLevantado | DuracionMin | OperadorLevanta | Estado

function inicializarLogTelefonos_() {
  const sh = hoja_(HOJAS.TELEFONOS);
  if (sh.getRange(3, 8).getValue()) return;
  sh.getRange(3, 8, 1, 10).setValues([[
    'Timestamp_log', 'PC_log', 'Linea', 'HoraCaida',
    'OperadorCaida', 'Motivo', 'HoraLevantado',
    'DuracionMin', 'OperadorLevanta', 'Estado_log'
  ]]);
}

function telefonoCaido_(datos) {
  inicializarLogTelefonos_();
  const sh = hoja_(HOJAS.TELEFONOS);
  const ahora = new Date();
  const nextRow = sh.getLastRow() + 1;

  sh.getRange(nextRow, 8, 1, 10).setValues([[
    new Date(),
    datos.pc || '',
    datos.linea || '',
    hora_(ahora),
    datos.nombre || '',
    datos.motivo || '',
    '', '', '',
    'Abierta'
  ]]);
  SpreadsheetApp.flush();
  return { ok: true, hora: hora_(ahora) };
}

function telefonoLevantado_(datos) {
  const sh = hoja_(HOJAS.TELEFONOS);
  const lastRow = sh.getLastRow();
  if (lastRow < 4) return { ok: true };

  const logRange = sh.getRange(4, 8, lastRow - 3, 10).getValues();
  const lineaLower = String(datos.linea || '').trim().toLowerCase();
  const ahora = new Date();

  for (let i = logRange.length - 1; i >= 0; i--) {
    const ts     = logRange[i][0];
    const linea  = String(logRange[i][2] || '').trim().toLowerCase();
    const estado = logRange[i][9];

    if (linea !== lineaLower || estado !== 'Abierta') continue;

    const duracion = ts instanceof Date ? Math.max(0, Math.ceil((ahora - ts) / 60000)) : 0;
    const sheetRow = 4 + i;

    sh.getRange(sheetRow, 14).setValue(hora_(ahora));
    sh.getRange(sheetRow, 15).setValue(duracion);
    sh.getRange(sheetRow, 16).setValue(datos.nombre || '');
    sh.getRange(sheetRow, 17).setValue('Cerrada');
    SpreadsheetApp.flush();

    return { ok: true, duracion, hora: hora_(ahora) };
  }
  return { ok: true };
}

function getLogTelefono_(datos) {
  const sh = hoja_(HOJAS.TELEFONOS);
  const lastRow = sh.getLastRow();
  if (lastRow < 4) return { ok: true, log: [] };

  const logRange = sh.getRange(4, 8, lastRow - 3, 10).getValues();
  const lineaLower = String(datos.linea || '').trim().toLowerCase();
  const log = [];

  for (let i = logRange.length - 1; i >= 0; i--) {
    if (!logRange[i][0]) continue;
    const linea = String(logRange[i][2] || '').trim().toLowerCase();
    if (lineaLower && linea !== lineaLower) continue;

    log.push({
      pc: logRange[i][1],
      linea: logRange[i][2],
      horaCaida: logRange[i][3],
      operadorCaida: logRange[i][4],
      motivo: logRange[i][5],
      horaLevantado: logRange[i][6],
      duracion: logRange[i][7],
      operadorLevanta: logRange[i][8],
      estado: logRange[i][9]
    });
    if (log.length >= 20) break;
  }

  return { ok: true, cantidad: log.length, log };
}

// --- NOTA EN NOVEDAD ---

function notarNovedad_(datos) {
  const idBuscado = String(datos.id || '').trim();
  if (!idBuscado) return { ok: false, error: 'Falta ID.' };

  const sh = hoja_(HOJAS.NOVEDADES);
  const values = sh.getDataRange().getValues();

  for (let i = 3; i < values.length; i++) {
    if (String(values[i][10] || '').trim() !== idBuscado) continue;
    sh.getRange(i + 1, 14).setValue(datos.nota || '');
    SpreadsheetApp.flush();
    return { ok: true };
  }
  return { ok: false, error: 'Novedad no encontrada.' };
}

// ===========================
// ============ ADMIN ========
// ===========================

function esAdmin_(nombre) {
  if (!nombre) return false;
  const todos = filas_(HOJAS.OPERADORES);
  const op = todos.find(o => String(o.Nombre || '').trim().toLowerCase() === String(nombre).trim().toLowerCase());
  if (!op) return false;
  const tipo = String(op.Tipo || '').toLowerCase();
  const rol  = String(op.Rol  || '').toLowerCase();
  return tipo.indexOf('admin') !== -1 || rol.indexOf('admin') !== -1;
}

function adminGuard_(datos, fn) {
  if (!esAdmin_(datos.nombre)) return { ok: false, error: 'Acceso denegado: requiere rol admin.' };
  return fn(datos);
}

function loginAdmin_(datos) {
  const todos = filas_(HOJAS.OPERADORES);
  const nombre = String(datos.nombre || '').trim().toLowerCase();
  const op = todos.find(o => String(o.Nombre || '').trim().toLowerCase() === nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };
  if (String(op.PIN) !== String(datos.pin)) return { ok: false, error: 'PIN incorrecto.' };
  if (!esAdmin_(op.Nombre)) return { ok: false, error: 'No tenés permisos de administrador.' };
  return {
    ok: true,
    admin: { nombre: op.Nombre, tipo: op.Tipo, pc: op.PC, turno: op.Turno }
  };
}

// --- DASHBOARD ---

function getDashboard_() {
  const ahora = new Date();
  const pcsResp = getPCs_();
  const pcs = pcsResp.pcs || [];

  // Última acción por operador en FICHAJES
  const fichajes = hoja_(HOJAS.FICHAJES).getDataRange().getValues();
  const ultimosPorOp = {};
  for (let i = 3; i < fichajes.length; i++) {
    const ts = fichajes[i][0];
    const nombre = fichajes[i][2];
    if (!nombre || !ts) continue;
    const key = String(nombre).trim().toLowerCase();
    if (!ultimosPorOp[key] || ts > ultimosPorOp[key].ts) {
      ultimosPorOp[key] = {
        ts, nombre,
        turno: fichajes[i][3], pc: fichajes[i][4], tipo: fichajes[i][5],
        accion: fichajes[i][6], hora: fichajes[i][7]
      };
    }
  }

  // PCs activas: agrupar operadores con Ingreso vigente
  const pcsActivas = {};
  for (const k in ultimosPorOp) {
    const u = ultimosPorOp[k];
    if (u.accion === 'Ingreso') {
      const pcKey = String(u.pc || '').trim().toUpperCase();
      if (!pcsActivas[pcKey]) pcsActivas[pcKey] = [];
      pcsActivas[pcKey].push({
        nombre: u.nombre,
        turno: u.turno,
        desde: u.hora,
        minutosSesion: u.ts instanceof Date ? Math.max(0, Math.floor((ahora - u.ts) / 60000)) : 0
      });
    }
  }

  // Pausas abiertas por PC
  const pausasAbiertas = {};
  ['BANOS', 'FUMAR', 'LIMPIEZA'].forEach(k => {
    let sh; try { sh = hoja_(HOJAS[k]); } catch (e) { return; }
    const v = sh.getDataRange().getValues();
    for (let i = 3; i < v.length; i++) {
      if (v[i][9] !== 'Abierta') continue;
      const pcKey = String(v[i][4] || '').trim().toUpperCase();
      if (!pausasAbiertas[pcKey]) pausasAbiertas[pcKey] = [];
      pausasAbiertas[pcKey].push({
        operador: v[i][2],
        tipo: k.toLowerCase(),
        desde: v[i][5],
        minutos: v[i][0] instanceof Date ? Math.max(0, Math.floor((ahora - v[i][0]) / 60000)) : 0
      });
    }
  });

  // Conteos por PC
  const novedadesPorPC = {};
  const novV = hoja_(HOJAS.NOVEDADES).getDataRange().getValues();
  for (let i = 3; i < novV.length; i++) {
    const estado = String(novV[i][11] || '').toLowerCase();
    if (estado !== 'pendiente' && estado !== 'en seguimiento') continue;
    const pcKey = String(novV[i][4] || '').trim().toUpperCase();
    novedadesPorPC[pcKey] = (novedadesPorPC[pcKey] || 0) + 1;
  }

  const tareasPorPC = {};
  try {
    const tSh = inicializarTareasPC_();
    const tV = tSh.getDataRange().getValues();
    for (let i = 3; i < tV.length; i++) {
      const estado = String(tV[i][4] || '').toLowerCase();
      if (estado === 'completada') continue;
      if (!tV[i][3]) continue;
      const pcKey = String(tV[i][2] || '').trim().toUpperCase();
      tareasPorPC[pcKey] = (tareasPorPC[pcKey] || 0) + 1;
    }
  } catch (e) {}

  // Teléfonos caídos por PC desde el log
  const telCaidosPorPC = {};
  try {
    const shTel = hoja_(HOJAS.TELEFONOS);
    const lastRow = shTel.getLastRow();
    if (lastRow >= 4) {
      const logRange = shTel.getRange(4, 8, lastRow - 3, 10).getValues();
      for (let i = 0; i < logRange.length; i++) {
        if (logRange[i][9] !== 'Abierta') continue;
        const pcKey = String(logRange[i][1] || '').trim().toUpperCase();
        telCaidosPorPC[pcKey] = (telCaidosPorPC[pcKey] || 0) + 1;
      }
    }
  } catch (e) {}

  const result = pcs.map(p => {
    const pcKey = String(p.pc || '').trim().toUpperCase();
    return {
      pc: p.pc,
      orden: p.orden,
      operadores:        pcsActivas[pcKey] || [],
      pausasAbiertas:    pausasAbiertas[pcKey] || [],
      novedadesPendientes: novedadesPorPC[pcKey] || 0,
      tareasPendientes:    tareasPorPC[pcKey] || 0,
      telefonosCaidos:     telCaidosPorPC[pcKey] || 0
    };
  });

  return { ok: true, pcs: result, ahora: hora_(ahora), fecha: fecha_(ahora) };
}

// --- OPERADORES CRUD ---

function getOperadoresAdmin_(_datos) {
  const ops = filas_(HOJAS.OPERADORES);
  return { ok: true, operadores: ops };
}

function _filaOperador_(nombre) {
  const sh = hoja_(HOJAS.OPERADORES);
  const v = sh.getDataRange().getValues();
  const headers = v[2];
  const idxNombre = headers.findIndex(h => String(h).trim().toLowerCase() === 'nombre');
  if (idxNombre === -1) return null;
  for (let i = 3; i < v.length; i++) {
    if (String(v[i][idxNombre] || '').trim().toLowerCase() === String(nombre || '').trim().toLowerCase()) {
      return { sh, fila: i + 1, headers };
    }
  }
  return null;
}

function crearOperador_(datos) {
  if (!datos.nombre) return { ok: false, error: 'Falta nombre.' };
  if (_filaOperador_(datos.nombre)) return { ok: false, error: 'Ya existe un operador con ese nombre.' };

  const sh = hoja_(HOJAS.OPERADORES);
  const headers = sh.getDataRange().getValues()[2];
  const row = headers.map(h => {
    const k = String(h).trim().toLowerCase();
    if (k === 'nombre') return datos.nombre;
    if (k === 'pin') return datos.pin || '';
    if (k === 'turno') return datos.turno || '';
    if (k === 'pc') return datos.pc || '';
    if (k === 'tipo') return datos.tipo || 'Fijo';
    if (k === 'rol') return datos.rol || '';
    if (k === 'activo') return 'Sí';
    return '';
  });
  sh.appendRow(row);
  SpreadsheetApp.flush();
  return { ok: true };
}

function editarOperador_(datos) {
  const ref = _filaOperador_(datos.nombreOriginal || datos.nombre);
  if (!ref) return { ok: false, error: 'Operador no encontrado.' };

  const map = { nombre: datos.nombre, pin: datos.pin, turno: datos.turno, pc: datos.pc, tipo: datos.tipo, rol: datos.rol };
  for (const key in map) {
    if (map[key] === undefined || map[key] === '') continue;
    const idx = ref.headers.findIndex(h => String(h).trim().toLowerCase() === key);
    if (idx === -1) continue;
    ref.sh.getRange(ref.fila, idx + 1).setValue(map[key]);
  }
  SpreadsheetApp.flush();
  return { ok: true };
}

function eliminarOperador_(datos) {
  const ref = _filaOperador_(datos.nombre);
  if (!ref) return { ok: false, error: 'Operador no encontrado.' };
  ref.sh.deleteRow(ref.fila);
  return { ok: true };
}

function setActivoOperador_(datos) {
  const ref = _filaOperador_(datos.nombre);
  if (!ref) return { ok: false, error: 'Operador no encontrado.' };
  const idx = ref.headers.findIndex(h => String(h).trim().toLowerCase() === 'activo');
  if (idx === -1) return { ok: false, error: 'No existe columna Activo.' };
  const val = (datos.activo === true || String(datos.activo).toLowerCase() === 'true' || datos.activo === 'Sí') ? 'Sí' : 'No';
  ref.sh.getRange(ref.fila, idx + 1).setValue(val);
  SpreadsheetApp.flush();
  return { ok: true, activo: val };
}

function setPinOperador_(datos) {
  const ref = _filaOperador_(datos.nombre);
  if (!ref) return { ok: false, error: 'Operador no encontrado.' };
  const idx = ref.headers.findIndex(h => String(h).trim().toLowerCase() === 'pin');
  if (idx === -1) return { ok: false, error: 'No existe columna PIN.' };
  if (!datos.pinNuevo) return { ok: false, error: 'PIN nuevo vacío.' };
  ref.sh.getRange(ref.fila, idx + 1).setValue(datos.pinNuevo);
  SpreadsheetApp.flush();
  return { ok: true };
}

// --- TELÉFONOS CRUD ---

function getTelefonosAdmin_(_datos) {
  // Solo cols A-F (lista de teléfonos, no log)
  const sh = hoja_(HOJAS.TELEFONOS);
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 3; i < v.length; i++) {
    if (!v[i][0]) continue;
    out.push({
      fila: i + 1,
      linea: v[i][0],
      tipo: v[i][1],
      pc: v[i][2],
      estadoEsperado: v[i][3],
      activoEnSistema: v[i][4],
      numero: v[i][5]
    });
  }
  return { ok: true, telefonos: out };
}

function _filaTelefono_(linea, pc) {
  const sh = hoja_(HOJAS.TELEFONOS);
  const v = sh.getDataRange().getValues();
  for (let i = 3; i < v.length; i++) {
    if (String(v[i][0] || '').trim().toLowerCase() === String(linea || '').trim().toLowerCase()
        && String(v[i][2] || '').trim().toLowerCase() === String(pc || '').trim().toLowerCase()) {
      return { sh, fila: i + 1 };
    }
  }
  return null;
}

function crearTelefono_(datos) {
  if (!datos.linea || !datos.pc) return { ok: false, error: 'Falta línea o PC.' };
  if (_filaTelefono_(datos.linea, datos.pc)) return { ok: false, error: 'Ya existe ese teléfono en esa PC.' };
  const sh = hoja_(HOJAS.TELEFONOS);
  // Insertar nueva fila después de la última fila de teléfono (no tocar el log)
  // Estrategia: encontrar la última fila con dato en col A (entre 4 y getLastRow)
  const lastRow = sh.getLastRow();
  let ultimaFilaTel = 3;
  for (let i = 4; i <= lastRow; i++) {
    if (sh.getRange(i, 1).getValue()) ultimaFilaTel = i;
  }
  const insertRow = ultimaFilaTel + 1;
  sh.insertRowBefore(insertRow);
  sh.getRange(insertRow, 1, 1, 6).setValues([[
    datos.linea, datos.tipo || 'Principal', datos.pc,
    datos.estadoEsperado || 'Activo', datos.activoEnSistema || 'Sí', datos.numero || ''
  ]]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function editarTelefono_(datos) {
  const ref = _filaTelefono_(datos.lineaOriginal || datos.linea, datos.pcOriginal || datos.pc);
  if (!ref) return { ok: false, error: 'Teléfono no encontrado.' };
  const updates = [datos.linea, datos.tipo, datos.pc, datos.estadoEsperado, datos.activoEnSistema, datos.numero];
  updates.forEach((v, idx) => {
    if (v !== undefined && v !== '') ref.sh.getRange(ref.fila, idx + 1).setValue(v);
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

function eliminarTelefono_(datos) {
  const ref = _filaTelefono_(datos.linea, datos.pc);
  if (!ref) return { ok: false, error: 'Teléfono no encontrado.' };
  // No deleteRow porque puede haber log en cols H+. Limpiar A-F solamente.
  ref.sh.getRange(ref.fila, 1, 1, 6).clearContent();
  SpreadsheetApp.flush();
  return { ok: true };
}

// --- TAREAS ADMIN ---

function getTareasAdmin_(datos) {
  const sh = inicializarTareasPC_();
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 3; i < v.length; i++) {
    if (!v[i][3]) continue;
    out.push({
      fila: i + 1,
      timestamp: v[i][0], fecha: v[i][1], pc: v[i][2],
      tarea: v[i][3], estado: v[i][4],
      fechaCompletada: v[i][5], completoPor: v[i][6]
    });
  }
  return { ok: true, tareas: out };
}

function crearTarea_(datos) {
  if (!datos.pc || !datos.tareaTexto) return { ok: false, error: 'Falta PC o texto de tarea.' };
  const sh = inicializarTareasPC_();
  const ahora = new Date();
  sh.appendRow([
    new Date(), fecha_(ahora), datos.pc, datos.tareaTexto, 'Pendiente', '', ''
  ]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function eliminarTareaAdmin_(datos) {
  const fila = parseInt(datos.fila || datos.tareaId || '0');
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = inicializarTareasPC_();
  sh.deleteRow(fila);
  return { ok: true };
}

// --- NOVEDADES ADMIN ---

function getNovedadesAdmin_(datos) {
  const sh = hoja_(HOJAS.NOVEDADES);
  const v = sh.getDataRange().getValues();
  const pcF    = String(datos.pc || '').trim().toLowerCase();
  const estF   = String(datos.estado || '').trim().toLowerCase();
  const limit  = datos.limit || 200;
  const out = [];
  for (let i = v.length - 1; i >= 3; i--) {
    if (!v[i][10]) continue;
    const pcN = String(v[i][4] || '').trim().toLowerCase();
    const eN  = String(v[i][11] || '').trim().toLowerCase();
    if (pcF && pcN !== pcF) continue;
    if (estF && eN !== estF) continue;
    out.push({
      fila: i + 1,
      timestamp: v[i][0], fecha: v[i][1], operador: v[i][2], turno: v[i][3], pc: v[i][4],
      tipo: v[i][5], prioridad: v[i][6], detalle: v[i][7], resuelto: v[i][8],
      observaciones: v[i][9], id: v[i][10], estado: v[i][11], revisadoPor: v[i][12], resolucion: v[i][13]
    });
    if (out.length >= limit) break;
  }
  return { ok: true, cantidad: out.length, novedades: out };
}

function eliminarNovedad_(datos) {
  const idBuscado = String(datos.id || '').trim();
  if (!idBuscado) return { ok: false, error: 'Falta ID.' };
  const sh = hoja_(HOJAS.NOVEDADES);
  const v = sh.getDataRange().getValues();
  for (let i = 3; i < v.length; i++) {
    if (String(v[i][10] || '').trim() === idBuscado) {
      sh.deleteRow(i + 1);
      return { ok: true };
    }
  }
  return { ok: false, error: 'Novedad no encontrada.' };
}

// --- PCs CRUD ---

function getPCsAdmin_(_datos) {
  const filas = filas_(HOJAS.PCS);
  return { ok: true, pcs: filas };
}

function _filaPC_(pc) {
  const sh = hoja_(HOJAS.PCS);
  const v = sh.getDataRange().getValues();
  for (let i = 3; i < v.length; i++) {
    if (String(v[i][0] || '').trim().toLowerCase() === String(pc || '').trim().toLowerCase()) {
      return { sh, fila: i + 1, headers: v[2] };
    }
  }
  return null;
}

function crearPC_(datos) {
  if (!datos.pc) return { ok: false, error: 'Falta nombre de PC.' };
  if (_filaPC_(datos.pc)) return { ok: false, error: 'Ya existe esa PC.' };
  const sh = hoja_(HOJAS.PCS);
  const headers = sh.getDataRange().getValues()[2] || [];
  const row = headers.map(h => {
    const k = String(h).trim().toLowerCase();
    if (k === 'pc') return datos.pc;
    if (k.indexOf('orden') !== -1) return datos.orden || 999;
    if (k === 'estado') return datos.estadoPC || 'Activa';
    return '';
  });
  if (!row.length) sh.appendRow([datos.pc, datos.orden || 999, datos.estadoPC || 'Activa']);
  else sh.appendRow(row);
  SpreadsheetApp.flush();
  return { ok: true };
}

function editarPC_(datos) {
  const ref = _filaPC_(datos.pcOriginal || datos.pc);
  if (!ref) return { ok: false, error: 'PC no encontrada.' };
  const map = { pc: datos.pc, orden: datos.orden, estado: datos.estadoPC };
  ref.headers.forEach((h, idx) => {
    const k = String(h).trim().toLowerCase();
    if (k === 'pc' && map.pc) ref.sh.getRange(ref.fila, idx + 1).setValue(map.pc);
    else if (k.indexOf('orden') !== -1 && map.orden) ref.sh.getRange(ref.fila, idx + 1).setValue(map.orden);
    else if (k === 'estado' && map.estado) ref.sh.getRange(ref.fila, idx + 1).setValue(map.estado);
  });
  SpreadsheetApp.flush();
  return { ok: true };
}

function eliminarPC_(datos) {
  const ref = _filaPC_(datos.pc);
  if (!ref) return { ok: false, error: 'PC no encontrada.' };
  ref.sh.deleteRow(ref.fila);
  return { ok: true };
}

// --- HISTÓRICO ---

function _enRango_(ts, desdeStr, hastaStr) {
  if (!(ts instanceof Date)) return true;
  if (desdeStr) {
    const d = new Date(desdeStr); if (!isNaN(d) && ts < d) return false;
  }
  if (hastaStr) {
    const h = new Date(hastaStr); if (!isNaN(h)) { h.setHours(23,59,59,999); if (ts > h) return false; }
  }
  return true;
}

function getFichajes_(datos) {
  const sh = hoja_(HOJAS.FICHAJES);
  const v = sh.getDataRange().getValues();
  const opF = String(datos.operador || '').trim().toLowerCase();
  const limit = datos.limit || 200;
  const out = [];
  for (let i = v.length - 1; i >= 3; i--) {
    if (!v[i][2]) continue;
    if (!_enRango_(v[i][0], datos.desde, datos.hasta)) continue;
    if (opF && String(v[i][2]).trim().toLowerCase() !== opF) continue;
    out.push({
      fila: i + 1,
      timestamp: v[i][0], fecha: v[i][1], nombre: v[i][2], turno: v[i][3], pc: v[i][4],
      tipo: v[i][5], accion: v[i][6], hora: v[i][7], estado: v[i][8],
      minutosTarde: v[i][9], observaciones: v[i][10]
    });
    if (out.length >= limit) break;
  }
  return { ok: true, cantidad: out.length, fichajes: out };
}

function getPausasAdmin_(datos) {
  const tipos = { bano: HOJAS.BANOS, fumar: HOJAS.FUMAR, limpieza: HOJAS.LIMPIEZA };
  const tipoF = String(datos.tipoPausa || '').toLowerCase();
  const opF = String(datos.operador || '').trim().toLowerCase();
  const pcF = String(datos.pc || '').trim().toLowerCase();
  const limit = datos.limit || 200;
  const out = [];

  for (const k in tipos) {
    if (tipoF && k !== tipoF) continue;
    let sh; try { sh = hoja_(tipos[k]); } catch(e) { continue; }
    const v = sh.getDataRange().getValues();
    for (let i = v.length - 1; i >= 3; i--) {
      if (!v[i][2]) continue;
      if (!_enRango_(v[i][0], datos.desde, datos.hasta)) continue;
      if (opF && String(v[i][2]).trim().toLowerCase() !== opF) continue;
      if (pcF && String(v[i][4] || '').trim().toLowerCase() !== pcF) continue;
      out.push({
        fila: i + 1, tipoPausa: k,
        timestamp: v[i][0], fecha: v[i][1], operador: v[i][2], turno: v[i][3], pc: v[i][4],
        horaSalida: v[i][5], horaRegreso: v[i][6], minutos: v[i][7], excedio: v[i][8],
        estado: v[i][9], observaciones: v[i][10]
      });
      if (out.length >= limit) break;
    }
    if (out.length >= limit) break;
  }
  // Ordenar por timestamp desc
  out.sort((a, b) => (b.timestamp instanceof Date ? b.timestamp.getTime() : 0) - (a.timestamp instanceof Date ? a.timestamp.getTime() : 0));
  return { ok: true, cantidad: out.length, pausas: out };
}

function getRecepciones_(datos) {
  const sh = hoja_(HOJAS.RECEPCION);
  const v = sh.getDataRange().getValues();
  const pcF = String(datos.pc || '').trim().toLowerCase();
  const limit = datos.limit || 200;
  const out = [];
  for (let i = v.length - 1; i >= 3; i--) {
    if (!v[i][4]) continue;
    if (!_enRango_(v[i][0], datos.desde, datos.hasta)) continue;
    if (pcF && String(v[i][3] || '').trim().toLowerCase() !== pcF) continue;
    out.push({
      fila: i + 1,
      timestamp: v[i][0], fecha: v[i][1], turno: v[i][2], pc: v[i][3],
      recibe: v[i][4], entrega: v[i][5], puestoOrden: v[i][6], whatsappOk: v[i][7],
      cajaOk: v[i][8], diferenciaCaja: v[i][9], revinculacionOk: v[i][10],
      telefonosCaidos: v[i][11], limpieza: v[i][12], observaciones: v[i][13]
    });
    if (out.length >= limit) break;
  }
  return { ok: true, cantidad: out.length, recepciones: out };
}

function getLogTelefonosAdmin_(datos) {
  const sh = hoja_(HOJAS.TELEFONOS);
  const lastRow = sh.getLastRow();
  if (lastRow < 4) return { ok: true, log: [] };

  const logRange = sh.getRange(4, 8, lastRow - 3, 10).getValues();
  const pcF = String(datos.pc || '').trim().toLowerCase();
  const lineaF = String(datos.linea || '').trim().toLowerCase();
  const limit = datos.limit || 200;
  const log = [];

  for (let i = logRange.length - 1; i >= 0; i--) {
    if (!logRange[i][0]) continue;
    if (pcF && String(logRange[i][1] || '').trim().toLowerCase() !== pcF) continue;
    if (lineaF && String(logRange[i][2] || '').trim().toLowerCase() !== lineaF) continue;
    if (!_enRango_(logRange[i][0], datos.desde, datos.hasta)) continue;
    log.push({
      fila: 4 + i,
      timestamp: logRange[i][0], pc: logRange[i][1], linea: logRange[i][2],
      horaCaida: logRange[i][3], operadorCaida: logRange[i][4], motivo: logRange[i][5],
      horaLevantado: logRange[i][6], duracion: logRange[i][7],
      operadorLevanta: logRange[i][8], estado: logRange[i][9]
    });
    if (log.length >= limit) break;
  }
  return { ok: true, cantidad: log.length, log };
}

// ===========================
// === COMENTARIOS NOVEDAD ===
// ===========================

function inicializarComentarios_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJAS.COMENTARIOS);
  if (!sh) sh = ss.insertSheet(HOJAS.COMENTARIOS);
  if (String(sh.getRange(3, 1).getValue()).toLowerCase() === 'timestamp') {
    // Si ya existe pero falta col ImagenURL, la agregamos
    if (String(sh.getRange(3, 7).getValue()).toLowerCase() !== 'imagenurl') {
      sh.getRange(3, 7).setValue('ImagenURL');
    }
    return sh;
  }
  if (sh.getLastRow() < 1) sh.getRange(1, 1).setValue('NOVEDADES_COMENTARIOS');
  sh.getRange(3, 1, 1, 7).setValues([[
    'Timestamp', 'IdNovedad', 'Autor', 'Texto', 'Editado', 'Eliminado', 'ImagenURL'
  ]]);
  return sh;
}

// === DRIVE: carpeta auto-setup para imágenes de novedades ===
function _getDriveFolderId_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('NOVEDADES_FOLDER_ID');
  if (id) {
    try { DriveApp.getFolderById(id); return id; }
    catch (e) { /* folder borrado, recreamos */ }
  }
  const folder = DriveApp.createFolder('CTRL_NOVEDADES_IMAGENES');
  try { folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  id = folder.getId();
  props.setProperty('NOVEDADES_FOLDER_ID', id);
  return id;
}

function _subirImagenBase64_(base64DataUrl, nombreArchivo) {
  const match = String(base64DataUrl || '').match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Formato base64 inválido');
  const mime = match[1];
  const data = Utilities.base64Decode(match[2]);
  const blob = Utilities.newBlob(data, mime, nombreArchivo || 'imagen_' + Date.now());
  const folderId = _getDriveFolderId_();
  const file = DriveApp.getFolderById(folderId).createFile(blob);
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  return 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200';
}

function getComentariosNovedad_(datos) {
  const idF = String(datos.id || '').trim();
  if (!idF) return { ok: true, comentarios: [] };
  let sh; try { sh = inicializarComentarios_(); } catch (e) { return { ok: true, comentarios: [] }; }
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = 3; i < v.length; i++) {
    if (!v[i][0]) continue;
    if (String(v[i][1] || '').trim() !== idF) continue;
    if (v[i][5]) continue;
    out.push({
      fila: i + 1,
      timestamp: v[i][0],
      idNovedad: v[i][1],
      autor: v[i][2],
      texto: v[i][3],
      editado: v[i][4],
      imagenURL: v[i][6] || ''
    });
  }
  return { ok: true, comentarios: out };
}

function agregarComentario_(datos) {
  const idF = String(datos.id || '').trim();
  const autor = String(datos.nombre || '').trim();
  const texto = String(datos.nota || datos.detalle || '').trim();
  const imagenBase64 = String(datos.imagenBase64 || '').trim();
  if (!idF || !autor) return { ok: false, error: 'Falta novedad o autor.' };
  if (!texto && !imagenBase64) return { ok: false, error: 'Comentario vacío.' };

  let imagenURL = '';
  if (imagenBase64) {
    try {
      imagenURL = _subirImagenBase64_(imagenBase64, 'novedad_' + idF + '_' + Date.now());
    } catch (e) {
      return { ok: false, error: 'Error subiendo imagen: ' + e.message };
    }
  }

  const sh = inicializarComentarios_();
  sh.appendRow([new Date(), idF, autor, texto, '', '', imagenURL]);
  SpreadsheetApp.flush();
  return { ok: true, imagenURL };
}

function editarComentario_(datos) {
  const fila = parseInt(datos.fila || '0');
  const autor = String(datos.nombre || '').trim().toLowerCase();
  const texto = String(datos.nota || datos.detalle || '').trim();
  if (!fila || fila < 4 || !texto) return { ok: false, error: 'Datos inválidos.' };
  const sh = inicializarComentarios_();
  const filaAutor = String(sh.getRange(fila, 3).getValue() || '').trim().toLowerCase();
  if (filaAutor !== autor && !esAdmin_(datos.nombre)) return { ok: false, error: 'Solo el autor puede editar.' };
  sh.getRange(fila, 4).setValue(texto);
  sh.getRange(fila, 5).setValue('Sí ' + hora_(new Date()));
  SpreadsheetApp.flush();
  return { ok: true };
}

function eliminarComentario_(datos) {
  const fila = parseInt(datos.fila || '0');
  const autor = String(datos.nombre || '').trim().toLowerCase();
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = inicializarComentarios_();
  const filaAutor = String(sh.getRange(fila, 3).getValue() || '').trim().toLowerCase();
  if (filaAutor !== autor && !esAdmin_(datos.nombre)) return { ok: false, error: 'Solo el autor puede eliminar.' };
  sh.getRange(fila, 6).setValue('Sí');
  SpreadsheetApp.flush();
  return { ok: true };
}

// ===========================
// ===== NOTAS DE TURNO ======
// ===========================

function inicializarNotasTurno_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(HOJAS.NOTAS_TURNO);
  if (!sh) sh = ss.insertSheet(HOJAS.NOTAS_TURNO);
  if (String(sh.getRange(3, 1).getValue()).toLowerCase() === 'timestamp') return sh;
  if (sh.getLastRow() < 1) sh.getRange(1, 1).setValue('NOTAS_TURNO');
  sh.getRange(3, 1, 1, 6).setValues([[
    'Timestamp', 'Fecha', 'Turno', 'PC', 'Operador', 'Nota'
  ]]);
  return sh;
}

function guardarNotaTurno_(datos) {
  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };
  const texto = String(datos.nota || '').trim();
  if (!texto) return { ok: false, error: 'Falta texto de la nota.' };
  const sh = inicializarNotasTurno_();
  const ahora = new Date();
  sh.appendRow([new Date(), fecha_(ahora), op.Turno, datos.pc || op.PC, op.Nombre, texto]);
  SpreadsheetApp.flush();
  return { ok: true };
}

function getNotaTurnoAnterior_(datos) {
  const pcF = String(datos.pc || '').trim().toLowerCase();
  if (!pcF) return { ok: true, notas: [] };
  let sh; try { sh = inicializarNotasTurno_(); } catch (e) { return { ok: true, notas: [] }; }
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = v.length - 1; i >= 3 && out.length < 5; i--) {
    if (!v[i][3]) continue;
    if (String(v[i][3] || '').trim().toLowerCase() !== pcF) continue;
    out.push({
      timestamp: v[i][0], fecha: v[i][1], turno: v[i][2], pc: v[i][3], operador: v[i][4], nota: v[i][5]
    });
  }
  return { ok: true, notas: out };
}

// ===========================
// ====== EDICION ADMIN ======
// ===========================

// Helper: parsea "HH:MM" o "HH:MM:SS" a string normalizado HH:MM:SS
function _parseHora_(s) {
  if (!s) return '';
  const m = String(s).trim().match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return String(s);
  const h = ('0' + m[1]).slice(-2);
  const min = m[2];
  const sec = m[3] || '00';
  return h + ':' + min + ':' + sec;
}

// Aplica updates {col: valor} a una fila si valor !== undefined && valor !== null
function _setFila_(sh, fila, updates) {
  Object.keys(updates).forEach(col => {
    const v = updates[col];
    if (v === undefined || v === null) return;
    sh.getRange(fila, parseInt(col)).setValue(v);
  });
  SpreadsheetApp.flush();
}

// --- FICHAJES ---
function editarFichaje_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = hoja_(HOJAS.FICHAJES);
  const updates = {};
  if (datos.fechaTexto) updates[2] = datos.fechaTexto;
  if (datos.nombre) updates[3] = datos.nombre;
  if (datos.turno) updates[4] = datos.turno;
  if (datos.pc) updates[5] = datos.pc;
  if (datos.tipo) updates[6] = datos.tipo;
  if (datos.accion) updates[7] = datos.accion;
  if (datos.hora) updates[8] = _parseHora_(datos.hora);
  if (datos.estado) updates[9] = datos.estado;
  if (datos.minutosTarde !== '' && datos.minutosTarde !== undefined) updates[10] = parseInt(datos.minutosTarde) || 0;
  if (datos.observaciones !== undefined) updates[11] = datos.observaciones;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

function eliminarFichaje_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  hoja_(HOJAS.FICHAJES).deleteRow(fila);
  return { ok: true };
}

// --- PAUSAS ---
function _hojaPausa_(tipoPausa) {
  const t = String(tipoPausa || '').toLowerCase();
  if (t === 'bano' || t === 'baño' || t === 'baño') return HOJAS.BANOS;
  if (t === 'fumar') return HOJAS.FUMAR;
  if (t === 'limpieza') return HOJAS.LIMPIEZA;
  return null;
}

function editarPausaAdmin_(datos) {
  const fila = parseInt(datos.fila || 0);
  const hojaNombre = _hojaPausa_(datos.tipoPausa);
  if (!hojaNombre) return { ok: false, error: 'Tipo de pausa inválido.' };
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = hoja_(hojaNombre);
  const updates = {};
  if (datos.fechaTexto) updates[2] = datos.fechaTexto;
  if (datos.operador) updates[3] = datos.operador;
  if (datos.turno) updates[4] = datos.turno;
  if (datos.pc) updates[5] = datos.pc;
  if (datos.horaSalida) updates[6] = _parseHora_(datos.horaSalida);
  if (datos.horaRegreso !== undefined) updates[7] = datos.horaRegreso ? _parseHora_(datos.horaRegreso) : '';
  if (datos.minutos !== '' && datos.minutos !== undefined) updates[8] = parseInt(datos.minutos) || 0;
  if (datos.excedio) updates[9] = datos.excedio;
  if (datos.estado) updates[10] = datos.estado;
  if (datos.observaciones !== undefined) updates[11] = datos.observaciones;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

function eliminarPausa_(datos) {
  const fila = parseInt(datos.fila || 0);
  const hojaNombre = _hojaPausa_(datos.tipoPausa);
  if (!hojaNombre) return { ok: false, error: 'Tipo inválido.' };
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  hoja_(hojaNombre).deleteRow(fila);
  return { ok: true };
}

// --- RECEPCIONES ---
function editarRecepcion_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = hoja_(HOJAS.RECEPCION);
  const updates = {};
  if (datos.fechaTexto) updates[2] = datos.fechaTexto;
  if (datos.turno) updates[3] = datos.turno;
  if (datos.pc) updates[4] = datos.pc;
  if (datos.recibe) updates[5] = datos.recibe;
  if (datos.entrega !== undefined) updates[6] = datos.entrega;
  if (datos.puestoOrden) updates[7] = siNo_(datos.puestoOrden);
  if (datos.whatsappOk) updates[8] = siNo_(datos.whatsappOk);
  if (datos.cajaOk) updates[9] = siNo_(datos.cajaOk);
  if (datos.diferenciaCaja !== undefined) updates[10] = datos.diferenciaCaja;
  if (datos.revinculacionOk) updates[11] = siNo_(datos.revinculacionOk);
  if (datos.telefonosCaidos !== undefined) updates[12] = datos.telefonosCaidos;
  if (datos.limpiezaCumplida) updates[13] = datos.limpiezaCumplida;
  if (datos.observaciones !== undefined) updates[14] = datos.observaciones;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

function eliminarRecepcion_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  hoja_(HOJAS.RECEPCION).deleteRow(fila);
  return { ok: true };
}

// --- NOTAS DE TURNO ---
function editarNotaTurno_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = inicializarNotasTurno_();
  const updates = {};
  if (datos.fechaTexto) updates[2] = datos.fechaTexto;
  if (datos.turno) updates[3] = datos.turno;
  if (datos.pc) updates[4] = datos.pc;
  if (datos.operador) updates[5] = datos.operador;
  if (datos.nota !== undefined) updates[6] = datos.nota;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

function eliminarNotaTurno_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  inicializarNotasTurno_().deleteRow(fila);
  return { ok: true };
}

// --- LOG TELEFONOS (vive en CONFIG_TELEFONOS desde col H) ---
function editarLogTel_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = hoja_(HOJAS.TELEFONOS);
  // Cols H-Q = 8-17
  const updates = {};
  if (datos.pc) updates[9] = datos.pc;
  if (datos.linea) updates[10] = datos.linea;
  if (datos.horaCaida) updates[11] = _parseHora_(datos.horaCaida);
  if (datos.operadorCaida) updates[12] = datos.operadorCaida;
  if (datos.motivo !== undefined) updates[13] = datos.motivo;
  if (datos.horaLevantado !== undefined) updates[14] = datos.horaLevantado ? _parseHora_(datos.horaLevantado) : '';
  if (datos.duracion !== '' && datos.duracion !== undefined) updates[15] = parseInt(datos.duracion) || 0;
  if (datos.operadorLevanta !== undefined) updates[16] = datos.operadorLevanta;
  if (datos.estado) updates[17] = datos.estado;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

function eliminarLogTel_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  // No deleteRow porque cols A-F tienen la lista de teléfonos. Limpiar solo H-Q.
  const sh = hoja_(HOJAS.TELEFONOS);
  sh.getRange(fila, 8, 1, 10).clearContent();
  SpreadsheetApp.flush();
  return { ok: true };
}

// --- NOVEDAD CONTENIDO COMPLETO (no solo estado/resolución) ---
function editarNovedadContenido_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = hoja_(HOJAS.NOVEDADES);
  const updates = {};
  if (datos.fechaTexto) updates[2] = datos.fechaTexto;
  if (datos.operador) updates[3] = datos.operador;
  if (datos.turno) updates[4] = datos.turno;
  if (datos.pc) updates[5] = datos.pc;
  if (datos.tipo) updates[6] = datos.tipo;
  if (datos.prioridad) updates[7] = datos.prioridad;
  if (datos.detalle !== undefined) updates[8] = datos.detalle;
  if (datos.observaciones !== undefined) updates[10] = datos.observaciones;
  if (datos.estado) updates[12] = datos.estado;
  if (datos.revisadoPor !== undefined) updates[13] = datos.revisadoPor;
  if (datos.resolucion !== undefined) updates[14] = datos.resolucion;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

// --- TAREAS ---
function editarTareaAdmin_(datos) {
  const fila = parseInt(datos.fila || 0);
  if (!fila || fila < 4) return { ok: false, error: 'Fila inválida.' };
  const sh = inicializarTareasPC_();
  const updates = {};
  if (datos.pc) updates[3] = datos.pc;
  if (datos.tareaTexto) updates[4] = datos.tareaTexto;
  if (datos.estado) updates[5] = datos.estado;
  _setFila_(sh, fila, updates);
  return { ok: true };
}

function getNotasTurno_(datos) {
  const pcF = String(datos.pc || '').trim().toLowerCase();
  const opF = String(datos.operador || '').trim().toLowerCase();
  const limit = datos.limit || 200;
  let sh; try { sh = inicializarNotasTurno_(); } catch (e) { return { ok: true, notas: [] }; }
  const v = sh.getDataRange().getValues();
  const out = [];
  for (let i = v.length - 1; i >= 3 && out.length < limit; i--) {
    if (!v[i][0]) continue;
    if (pcF && String(v[i][3] || '').trim().toLowerCase() !== pcF) continue;
    if (opF && String(v[i][4] || '').trim().toLowerCase() !== opF) continue;
    if (!_enRango_(v[i][0], datos.desde, datos.hasta)) continue;
    out.push({
      fila: i + 1,
      timestamp: v[i][0], fecha: v[i][1], turno: v[i][2], pc: v[i][3], operador: v[i][4], nota: v[i][5]
    });
  }
  return { ok: true, cantidad: out.length, notas: out };
}