const HOJAS = {
  OPERADORES: 'CONFIG_OPERADORES',
  TELEFONOS: 'CONFIG_TELEFONOS',
  PCS: 'CONFIG_PC',
  FICHAJES: 'FICHAJES',
  RECEPCION: 'RECEPCION_TURNO',
  BANOS: 'BAÑO_PAUSAS',
  FUMAR: 'FUMAR_PAUSAS',
  LIMPIEZA: 'LIMPIEZA_PAUSAS',
  NOVEDADES: 'ERRORES_NOVEDADES'
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

function login_(datos) {
  const op = getOperador_(datos.nombre);

  if (!op) return { ok: false, error: 'Operador no activo o no encontrado.' };
  if (String(op.PIN) !== String(datos.pin)) return { ok: false, error: 'PIN incorrecto.' };

  if (esFranquero_(op) && !datos.pcAsignada) {
    return { ok: true, requierePcSelector: true };
  }

  const pcFinal = esFranquero_(op)
    ? String(datos.pcAsignada).trim().toUpperCase()
    : op.PC;

  const cierres = cerrarSesionesExpiradas_();

  return {
    ok: true,
    operador: {
      nombre: op.Nombre,
      turno: op.Turno,
      pc: pcFinal,
      tipo: op.Tipo,
      esFranquero: esFranquero_(op)
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
  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

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

function iniciarPausa_(datos, hojaNombre, tipoPausa, opciones) {
  opciones = opciones || {};
  const bloqueoCruzado = opciones.bloqueoCruzado === true;

  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

  const pcOperador = String(datos.pcAsignada || op.PC || '').trim();

  const sh = hoja_(hojaNombre);
  const values = sh.getDataRange().getValues();

  for (let i = values.length - 1; i >= 3; i--) {
    const operadorRow = values[i][2];
    const pcRow = values[i][4];
    const tsInicio = values[i][0];
    const estado = values[i][9];

    if (estado !== 'Abierta') continue;

    const mismoOp = String(operadorRow || '').trim().toLowerCase()
      === String(op.Nombre || '').trim().toLowerCase();

    if (mismoOp) {
      return { ok: false, error: 'Ya tenés una pausa abierta de ' + tipoPausa + '.' };
    }

    if (bloqueoCruzado) {
      let minutosAbierta = '?';
      if (tsInicio instanceof Date) {
        minutosAbierta = Math.max(0, Math.floor((new Date() - tsInicio) / 60000));
      }

      return {
        ok: false,
        bloqueoCruzado: true,
        bloqueador: { operador: operadorRow, pc: pcRow, minutos: minutosAbierta, tipo: tipoPausa },
        error:
          'PC ' + pcRow + ' tiene una pausa de ' + tipoPausa + ' abierta hace ' + minutosAbierta + ' min ('
          + operadorRow + '). Esperá a que regrese y la cierre antes de salir. '
          + 'Si ya terminó, avisale a ' + operadorRow + ' que cierre la pausa desde su PC.'
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
  const idx = ((dia % r.pcs.length) + r.pcs.length) % r.pcs.length;
  const elegida = r.pcs[idx];

  return {
    ok: true,
    pc: elegida.pc,
    orden: elegida.orden,
    fecha: fecha_(ahora),
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