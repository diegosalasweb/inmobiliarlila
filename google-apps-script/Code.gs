/**
 * Recibe las consultas del formulario de contacto del sitio, las agrega a la
 * hoja activa y envía confirmación al cliente + notificación interna.
 *
 * Deploy: Extensions > Apps Script > pegar este código > Deploy > New deployment
 * > Web app > Execute as: Me > Who has access: Anyone > Deploy.
 * Copiar la URL del deployment (termina en /exec) a src/components/ContactForm.astro.
 */
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const data = JSON.parse(e.postData.contents);

    sheet.appendRow([
      new Date(),
      data.name || '',
      data.email || '',
      data.phone || '',
      data.product || '',
      data.message || '',
      'Nuevo'
    ]);

    GmailApp.sendEmail(
      data.email,
      '✓ Hemos recibido tu consulta',
      'Gracias por contactarnos. Nos pondremos en contacto pronto.\n\nInmobiliaria Lila'
    );

    GmailApp.sendEmail(
      'web.inmobiliarlila@gmail.com',
      'Nuevo contacto: ' + data.name,
      'Nombre: ' + data.name + '\nEmail: ' + data.email + '\nTeléfono: ' + data.phone + '\nProducto: ' + data.product + '\nMensaje: ' + data.message
    );

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
