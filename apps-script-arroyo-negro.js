/**
 * ARROYO NEGRO — SISTEMA DE CONTROL GANADERO
 * Apps Script · Web App Deployment
 * 
 * INSTRUCCIONES DE DEPLOYMENT:
 * 1. Abre script.google.com → Nuevo proyecto → pega este código
 * 2. Cambia SHEET_ID por el ID del Sheets de Arroyo Negro
 * 3. Implementar → Nueva implementación → Aplicación web
 * 4. Ejecutar como: Yo / Quién accede: Cualquier persona → Implementar
 * 5. Copia el URL y reemplaza APPS_SCRIPT_URL_AQUI en reporte-diario.html
 */

// ── CONFIGURACIÓN ──────────────────────────────────────────────────────────
const SHEET_ID   = '15drNerKwx8xNvSUG3-P0k6wFiarEhgSYQh5aQ4M7F04';
const EMAIL_OSCAR = 'oscar@email.com'; // Reemplazar con tu email real

// Nombres exactos de las pestañas en el Sheets
const TABS = {
  REPORTES : '📋 Reportes Diarios',
  PESAJES  : '⚖️ Pesajes y GDP',
  POTREROS : '🌿 Potreros',
  ALERTAS  : 'Log Alertas',
};

// ── ENTRY POINT ────────────────────────────────────────────────────────────
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const tipo = data.tipo || 'reporte_diario';

    if (tipo === 'pesaje') {
      return registrarPesaje(data);
    } else {
      return registrarReporteDiario(data);
    }
  } catch(err) {
    logError('doPost', err.toString());
    return respuesta('error', err.toString());
  }
}

function doGet(e) {
  return ContentService.createTextOutput(
    JSON.stringify({ status: 'ok', sistema: 'Arroyo Negro · Sistema Ganadero', version: '1.0' })
  ).setMimeType(ContentService.MimeType.JSON);
}

// ── REPORTE DIARIO ─────────────────────────────────────────────────────────
function registrarReporteDiario(d) {
  const ss  = SpreadsheetApp.openById(SHEET_ID);
  const ws  = ss.getSheetByName(TABS.REPORTES);

  if (!ws) {
    inicializarTabReportes(ss);
    return registrarReporteDiario(d);
  }

  // Verificar si el encabezado existe
  if (ws.getLastRow() < 2) {
    crearEncabezadoReportes(ws);
  }

  const fila = [
    new Date(),                    // A: Timestamp
    d.fecha        || '',          // B: Fecha del reporte
    d.trabajador   || '',          // C: Trabajador
    d.horaLlegada  || '',          // D: Hora llegada
    d.horaSalida   || '',          // E: Hora salida
    d.productividad|| '',          // F: % Productividad
    Number(d.g1)   || 0,          // G: G1 bovinos
    Number(d.g2)   || 0,          // H: G2 bovinos
    Number(d.bufalo)|| 0,         // I: Búfalos
    d.sanidad      || '',          // J: Novedad sanitaria
    d.novSanidad   || '',          // K: Detalle sanidad
    d.pesaje       || '',          // L: ¿Pesaje realizado?
    d.pesoG1       || '',          // M: Peso G1 kg
    d.pesoG2       || '',          // N: Peso G2 kg
    d.potreroActivo|| '',          // O: Potrero activo
    d.estadoPasto  || '',          // P: Estado pasto
    d.cambioPotrero|| '',          // Q: Cambio potrero
    d.nuevoPotrero || '',          // R: Nuevo potrero
    d.salMineral   || '',          // S: Sal mineral
    d.agua         || '',          // T: Bebederos/agua
    d.actividades  || '',          // U: Actividades realizadas
    d.cercas       || '',          // V: Estado cercas
    d.cercaDetalle || '',          // W: Detalle cerca
    d.insumosAgotan|| '',          // X: Insumos agotándose
    d.observaciones|| '',          // Y: Observaciones
    d.decisionRequerida|| '',      // Z: ¿Decisión requerida?
    d.decisionDetalle  || '',      // AA: Detalle decisión
  ];

  ws.appendRow(fila);

  // Si hay pesaje, también registrar en pestaña de pesajes
  if (d.pesaje === 'SI' && (d.pesoG1 || d.pesoG2)) {
    registrarPesajeDesdeReporte(ss, d);
  }

  // Verificar alertas
  verificarAlertas(d);

  return respuesta('ok', 'Reporte registrado correctamente');
}

// ── PESAJE Y GDP ───────────────────────────────────────────────────────────
function registrarPesajeDesdeReporte(ss, d) {
  const ws = ss.getSheetByName(TABS.PESAJES);
  if (!ws) return;

  const hoy = new Date();

  // G1
  if (d.pesoG1) {
    const gdpG1 = calcularGDP('G1', Number(d.pesoG1), hoy, ss);
    ws.appendRow([
      hoy, d.fecha, 'G1', 30,
      Number(d.pesoG1), Number(d.pesoG1)*2.2046,
      gdpG1.dias, gdpG1.ganancia, gdpG1.gdp,
      semaforo(gdpG1.gdp),
      diasAlSacrificio(Number(d.pesoG1), gdpG1.gdp),
      d.trabajador
    ]);
  }

  // G2
  if (d.pesoG2) {
    const gdpG2 = calcularGDP('G2', Number(d.pesoG2), hoy, ss);
    ws.appendRow([
      hoy, d.fecha, 'G2', 35,
      Number(d.pesoG2), Number(d.pesoG2)*2.2046,
      gdpG2.dias, gdpG2.ganancia, gdpG2.gdp,
      semaforo(gdpG2.gdp),
      diasAlSacrificio(Number(d.pesoG2), gdpG2.gdp),
      d.trabajador
    ]);
  }
}

function calcularGDP(grupo, pesoActual, fechaActual, ss) {
  // Buscar el último pesaje de este grupo en el historial
  const ws = ss.getSheetByName(TABS.PESAJES);
  if (!ws || ws.getLastRow() < 3) {
    // Usar datos base del sistema
    const BASE = {
      'G1': { peso: 165, fecha: new Date('2026-05-12') },
      'G2': { peso: 360, fecha: new Date('2026-05-12') }
    };
    const base = BASE[grupo] || { peso: pesoActual, fecha: fechaActual };
    const dias = Math.max(1, Math.round((fechaActual - base.fecha)/(1000*60*60*24)));
    const ganancia = pesoActual - base.peso;
    const gdp = ganancia / dias;
    return { dias, ganancia: ganancia.toFixed(1), gdp: gdp.toFixed(2) };
  }

  // Buscar último registro del grupo
  const data = ws.getDataRange().getValues();
  let ultimoPeso = null, ultimaFecha = null;
  for (let i = data.length-1; i >= 2; i--) {
    if (data[i][2] === grupo && data[i][4]) {
      ultimoPeso  = Number(data[i][4]);
      ultimaFecha = new Date(data[i][0]);
      break;
    }
  }

  if (!ultimoPeso) {
    const BASE = {
      'G1': { peso: 165, fecha: new Date('2026-05-12') },
      'G2': { peso: 360, fecha: new Date('2026-05-12') }
    };
    const base = BASE[grupo] || { peso: pesoActual, fecha: fechaActual };
    ultimoPeso  = base.peso;
    ultimaFecha = base.fecha;
  }

  const dias     = Math.max(1, Math.round((fechaActual - ultimaFecha)/(1000*60*60*24)));
  const ganancia = pesoActual - ultimoPeso;
  const gdp      = ganancia / dias;
  return { dias, ganancia: ganancia.toFixed(1), gdp: gdp.toFixed(2) };
}

function semaforo(gdp) {
  const v = parseFloat(gdp);
  if (v >= 0.8) return '✅ EXCELENTE';
  if (v >= 0.5) return '⚠ MEJORAR';
  return '🚨 CRÍTICO';
}

function diasAlSacrificio(pesoActual, gdp) {
  const META_KG = 454; // 1.000 lb
  const v = parseFloat(gdp);
  if (v <= 0) return 'Sin dato';
  const pendiente = META_KG - pesoActual;
  if (pendiente <= 0) return '✅ LISTO PARA SACRIFICIO';
  return Math.ceil(pendiente / v) + ' días';
}

// ── ALERTAS AUTOMÁTICAS ────────────────────────────────────────────────────
function verificarAlertas(d) {
  const alertas = [];

  // Novedad sanitaria
  if (d.sanidad === 'SI') {
    alertas.push(`🚨 SANIDAD: ${d.novSanidad || 'Novedad reportada por ' + d.trabajador}`);
  }

  // Sal mineral agotándose
  if (d.salMineral === 'REPONER') {
    alertas.push('⚠ SAL MINERAL: Reponer urgente — ganado sin suplemento');
  }

  // Cerca con falla
  if (d.cercas === 'FALLA') {
    alertas.push(`⚠ CERCA: Falla reportada — ${d.cercaDetalle || 'ver detalle en reporte'}`);
  }

  // Decisión requerida
  if (d.decisionRequerida === 'SI') {
    alertas.push(`📋 DECISIÓN REQUERIDA: ${d.decisionDetalle}`);
  }

  // Pasto agotado
  if (d.estadoPasto === 'AGOTADO') {
    alertas.push(`⚠ PASTO AGOTADO en ${d.potreroActivo} — cambio de potrero urgente`);
  }

  // Agua con problema
  if (d.agua === 'PROBLEMA') {
    alertas.push('🚨 BEBEDEROS: Problema de agua — verificar hoy');
  }

  if (alertas.length > 0) {
    enviarAlertaEmail(alertas, d);
  }
}

function enviarAlertaEmail(alertas, d) {
  try {
    const asunto = `🚨 Arroyo Negro — ${alertas.length} alerta(s) · ${d.fecha}`;
    const cuerpo = `ALERTAS ARROYO NEGRO — ${d.fecha}
Reportado por: ${d.trabajador}

${alertas.join('\n\n')}

━━━━━━━━━━━━━━━━
Ver reporte completo en Google Sheets → Arroyo Negro Sistema Ganadero
`;
    GmailApp.sendEmail(EMAIL_OSCAR, asunto, cuerpo);
  } catch(e) {
    logError('enviarAlertaEmail', e.toString());
  }
}

// ── TRIGGER DIARIO — Sin reporte = alerta ─────────────────────────────────
function verificarReporteDiario() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const ws = ss.getSheetByName(TABS.REPORTES);
  if (!ws || ws.getLastRow() < 3) return;

  const data  = ws.getDataRange().getValues();
  const hoy   = new Date();
  const ayer  = new Date(hoy); ayer.setDate(hoy.getDate()-1);

  const hayReporteHoy = data.slice(2).some(row => {
    if (!row[0]) return false;
    const d = new Date(row[0]);
    return d.toDateString() === hoy.toDateString();
  });

  if (!hayReporteHoy && hoy.getHours() >= 8) {
    GmailApp.sendEmail(
      EMAIL_OSCAR,
      `⚠ Arroyo Negro — Sin reporte hoy ${hoy.toLocaleDateString('es-EC')}`,
      `No se ha recibido el reporte diario de Héctor.\n\nSon las ${hoy.getHours()}:${String(hoy.getMinutes()).padStart(2,'0')}.\n\nLlamar ahora.`
    );
  }
}

// ── INICIALIZACIÓN DE PESTAÑAS ─────────────────────────────────────────────
function crearEncabezadoReportes(ws) {
  ws.getRange(1,1,1,27).setValues([[
    'Timestamp','Fecha','Trabajador','Hora Llegada','Hora Salida','% Productividad',
    'G1 Bovinos','G2 Bovinos','Búfalos',
    'Novedad Sanitaria','Detalle Sanidad',
    'Pesaje Realizado','Peso G1 (kg)','Peso G2 (kg)',
    'Potrero Activo','Estado Pasto','Cambio Potrero','Nuevo Potrero',
    'Sal Mineral','Agua/Bebederos',
    'Actividades','Cercas','Detalle Cerca',
    'Insumos Agotando','Observaciones',
    'Decisión Requerida','Detalle Decisión'
  ]]);
  ws.getRange(1,1,1,27).setFontWeight('bold')
    .setBackground('#1A3A1A').setFontColor('#FFFFFF');
  ws.setFrozenRows(1);
}

function inicializarTabReportes(ss) {
  let ws = ss.getSheetByName(TABS.REPORTES);
  if (!ws) ws = ss.insertSheet(TABS.REPORTES);
  crearEncabezadoReportes(ws);
  return ws;
}

// ── SETUP TRIGGERS ─────────────────────────────────────────────────────────
function crearTriggers() {
  // Eliminar triggers existentes
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  // Verificar reporte a las 8 AM todos los días
  ScriptApp.newTrigger('verificarReporteDiario')
    .timeBased().everyDays(1).atHour(8).create();

  // Segunda verificación a las 5 PM
  ScriptApp.newTrigger('verificarReporteDiario')
    .timeBased().everyDays(1).atHour(17).create();

  Logger.log('Triggers creados: verificación a las 8 AM y 5 PM');
}

// ── UTILIDADES ─────────────────────────────────────────────────────────────
function respuesta(status, mensaje) {
  return ContentService
    .createTextOutput(JSON.stringify({ status, mensaje }))
    .setMimeType(ContentService.MimeType.JSON);
}

function logError(funcion, error) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let ws = ss.getSheetByName(TABS.ALERTAS);
    if (!ws) ws = ss.insertSheet(TABS.ALERTAS);
    ws.appendRow([new Date(), funcion, error]);
  } catch(e) { /* silencioso */ }
}

// ── TEST MANUAL ────────────────────────────────────────────────────────────
function testReporte() {
  const datos = {
    tipo: 'reporte_diario',
    fecha: 'miércoles 17 jun 2026',
    trabajador: 'Héctor',
    horaLlegada: '06:30',
    horaSalida: '15:30',
    productividad: '75%',
    g1: '30', g2: '35', bufalo: '10',
    sanidad: 'NO', pesaje: 'NO',
    potreroActivo: 'P-05',
    estadoPasto: 'BUENO',
    cambioPotrero: 'NO',
    salMineral: 'OK', agua: 'OK',
    actividades: '• Chapea / maleza: 3 hrs | • Cambio de potrero: 0.5 hrs',
    cercas: 'OK',
    insumosAgotan: '—',
    observaciones: 'Todo normal. Pasto en buen estado.',
    decisionRequerida: 'NO'
  };
  registrarReporteDiario(datos);
  Logger.log('Test completado — revisar pestaña Reportes Diarios');
}
