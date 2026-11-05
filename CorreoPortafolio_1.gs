/**
 * Correo "Novedades de mi portafolio" — María Alejandra Urdaneta
 * Crea BORRADORES personalizados en Gmail (no envía nada).
 *
 * Uso:
 *  1) probarConmigo()          -> crea 1 borrador dirigido a ti, para revisarlo.
 *  2) crearBorradoresPremium() -> crea 1 borrador por cada persona de la hoja "Premium".
 *     Marca en la columna "Estado correo" las filas ya procesadas, así no se duplican.
 */

const CONFIG = {
  SPREADSHEET_ID: '19UBpDQfsRxwoH54tbDhJynr-ozRyXiUgN00lJEKzIvM', // base_de_datos_verificada
  HOJA: 'Premium',
  COL_NOMBRE: 'Nombre',
  COL_CORREO: 'Correo principal',
  COL_ESTADO: 'Estado correo',          // se crea sola si no existe
  ASUNTO: '{{NOMBRE}}, te comparto las novedades de mi portafolio',
  CORREO_PRUEBA: 'maria.urdaneta@neurowellnesslearning.com',
  DESDE: 'info@marialeurdaneta.com',   // dirección "Enviar como" ya configurada en tu Gmail de NeuroWellness
  NOMBRE_REMITENTE: 'María Alejandra Urdaneta',
  RESPONDER_A: 'info@marialeurdaneta.com',
  ADJUNTAR_PDF: true,                   // false = solo enlace de descarga
  BASE: 'https://marialeurdaneta.com/correo/',
  URL_PDF: 'https://marialeurdaneta.com/portafolio.pdf',
  URL_NEWSLETTER: 'mailto:info@marialeurdaneta.com?subject=' +
      encodeURIComponent('Sí, quiero recibir Aprendizaje en acción') +
      '&body=' + encodeURIComponent('¡Hola, María Alejandra! Quiero recibir tu newsletter semanal Aprendizaje en acción.')
};

function probarConmigo() {
  comprobarRemitente_();
  const r = recursos_();
  crearBorrador_(CONFIG.CORREO_PRUEBA, 'Mariale', r);
  Logger.log('Borrador de prueba creado. Revísalo en Gmail > Borradores.');
}

function crearBorradoresPremium() {
  comprobarRemitente_();
  const hoja = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.HOJA);
  const datos = hoja.getDataRange().getValues();
  const cab = datos[0].map(String);
  const iNom = cab.indexOf(CONFIG.COL_NOMBRE);
  const iMail = cab.indexOf(CONFIG.COL_CORREO);
  let iEst = cab.indexOf(CONFIG.COL_ESTADO);
  if (iNom < 0 || iMail < 0) throw new Error('No encuentro las columnas "Nombre" o "Correo principal".');
  if (iEst < 0) { iEst = cab.length; hoja.getRange(1, iEst + 1).setValue(CONFIG.COL_ESTADO); }

  const r = recursos_();
  let creados = 0;
  for (let f = 1; f < datos.length; f++) {
    const correo = String(datos[f][iMail] || '').trim();
    const estado = String(datos[f][iEst] || '').trim();
    if (!correo || estado) continue;               // sin correo o ya procesado
    if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(correo)) {
      hoja.getRange(f + 1, iEst + 1).setValue('Correo no válido');
      continue;
    }
    const nombre = primerNombre_(datos[f][iNom]);
    crearBorrador_(correo, nombre, r);
    hoja.getRange(f + 1, iEst + 1).setValue('Borrador creado ' +
        Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy HH:mm'));
    creados++;
    Utilities.sleep(400);
  }
  Logger.log('Borradores creados: ' + creados);
}

/* ---------- internos ---------- */

function comprobarRemitente_() {
  const alias = GmailApp.getAliases();
  if (alias.indexOf(CONFIG.DESDE) < 0) {
    throw new Error('Esta cuenta no tiene configurado "Enviar como" ' + CONFIG.DESDE +
        '. Ejecuta el script desde maria.urdaneta@neurowellnesslearning.com. Alias encontrados: ' + alias.join(', '));
  }
}

function recursos_() {
  const get = (u) => UrlFetchApp.fetch(u).getBlob();
  const html = UrlFetchApp.fetch(CONFIG.BASE + 'plantilla.html').getContentText('UTF-8');
  return {
    html: html,
    hero: get(CONFIG.BASE + 'hero.jpg').setName('hero.jpg'),
    news: get(CONFIG.BASE + 'news.jpg').setName('news.jpg'),
    firma: get(CONFIG.BASE + 'firma.jpg').setName('firma.jpg'),
    pdf: CONFIG.ADJUNTAR_PDF
        ? get(CONFIG.URL_PDF).setName('Maria Alejandra Urdaneta - Portafolio.pdf').setContentType('application/pdf')
        : null
  };
}

function crearBorrador_(correo, nombre, r) {
  const html = r.html
      .split('{{NOMBRE}}').join(nombre)
      .split('{{URL_PDF}}').join(CONFIG.URL_PDF)
      .split('{{URL_NEWSLETTER}}').join(CONFIG.URL_NEWSLETTER);
  const asunto = CONFIG.ASUNTO.split('{{NOMBRE}}').join(nombre);
  const texto = 'Hola, ' + nombre + ', ¡qué gusto volver a saludarte!\n\n' +
      'Te escribo para compartirte las novedades de mi portafolio de formación y consultoría: ' +
      CONFIG.URL_PDF + '\n\nMe encantaría que conversáramos 20 minutos. ' +
      'Un abrazo,\nMaría Alejandra Urdaneta\nmarialeurdaneta.com · +34 603 777 334';
  const opciones = {
    htmlBody: html,
    name: CONFIG.NOMBRE_REMITENTE,
    from: CONFIG.DESDE,
    replyTo: CONFIG.RESPONDER_A,
    inlineImages: { hero: r.hero, news: r.news, firma: r.firma }
  };
  if (r.pdf) opciones.attachments = [r.pdf];
  GmailApp.createDraft(correo, asunto, texto, opciones);
}

function primerNombre_(v) {
  const n = String(v || '').trim().split(/\s+/)[0] || '';
  return n ? n.charAt(0).toUpperCase() + n.slice(1).toLowerCase() : '';
}
