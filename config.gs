const HOJAS = {
  OPERADORES: 'CONFIG_OPERADORES',
  TURNOS: 'CONFIG_TURNOS',
  TELEFONOS: 'CONFIG_TELEFONOS',
  FICHAJES: 'FICHAJES',
  RECEPCION: 'RECEPCION_TURNO',
  BANOS: 'BAÑO_PAUSAS',
  FUMAR: 'FUMAR_PAUSAS',
  NOVEDADES: 'ERRORES_NOVEDADES'
};

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
  const turnos = filas_(HOJAS.TURNOS);
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
  return filas_(HOJAS.TURNOS).find(t =>
    String(t.Turno || '').trim().toLowerCase() === String(turno || '').trim().toLowerCase()
  );
}

function login_(datos) {
  const op = getOperador_(datos.nombre);

  if (!op) return { ok: false, error: 'Operador no activo o no encontrado.' };
  if (String(op.PIN) !== String(datos.pin)) return { ok: false, error: 'PIN incorrecto.' };

  return {
    ok: true,
    operador: {
      nombre: op.Nombre,
      turno: op.Turno,
      pc: op.PC,
      tipo: op.Tipo
    }
  };
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
    datos.observaciones || ''
  ]);

  return { ok: true, accion: 'Salida', hora: hora_(ahora) };
}

function banoInicio_(datos) {
  return iniciarPausa_(datos, HOJAS.BANOS, 'baño');
}

function banoFin_(datos) {
  return finalizarPausa_(datos, HOJAS.BANOS, 'baño');
}

function fumarInicio_(datos) {
  return iniciarPausa_(datos, HOJAS.FUMAR, 'fumar');
}

function fumarFin_(datos) {
  return finalizarPausa_(datos, HOJAS.FUMAR, 'fumar');
}

function iniciarPausa_(datos, hojaNombre, tipoPausa) {
  const op = getOperador_(datos.nombre);
  if (!op) return { ok: false, error: 'Operador no encontrado.' };

  const sh = hoja_(hojaNombre);
  const values = sh.getDataRange().getValues();

  for (let i = values.length - 1; i >= 3; i--) {
    const operador = values[i][2];
    const estado = values[i][9];

    if (
      String(operador || '').trim().toLowerCase() === String(op.Nombre || '').trim().toLowerCase()
      && estado === 'Abierta'
    ) {
      return { ok: false, error: 'Ya hay una pausa abierta para este operador.' };
    }
  }

  const ahora = new Date();

  sh.appendRow([
    new Date(),
    fecha_(ahora),
    op.Nombre,
    op.Turno,
    op.PC,
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

  if (nombreIdx === -1 || pinIdx === -1) return { ok: false, error: 'Estructura de hoja incorrecta.' };

  for (let i = 3; i < values.length; i++) {
    const row = values[i];
    const nombreFila = String(row[nombreIdx] || '').trim().toLowerCase();
    const activo = activoIdx !== -1 ? esSi_(row[activoIdx]) : true;

    if (nombreFila === nombre.toLowerCase() && activo) {
      if (String(row[pinIdx]) !== pinActual) {
        return { ok: false, error: 'PIN actual incorrecto.' };
      }

      sh.getRange(i + 1, pinIdx + 1).setValue(pinNuevo);
      return { ok: true };
    }
  }

  return { ok: false, error: 'Operador no encontrado.' };
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