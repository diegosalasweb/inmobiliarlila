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

    const emailGeneral = 'web.inmobiliarialila@gmail.com';
    const product = (data.product || '').toLowerCase();

    let emailDestino = emailGeneral;
    if (product.includes('bodega')) {
      emailDestino = 'bodegasbuzeta@gmail.com';
    } else if (product.includes('estacionamiento')) {
      emailDestino = 'estacionamientoslila@gmail.com';
    }

    const asuntoInterno = 'Nuevo contacto: ' + data.name;
    const cuerpoInterno = 'Nombre: ' + data.name + '\nEmail: ' + data.email + '\nTeléfono: ' + data.phone + '\nProducto: ' + data.product + '\nMensaje: ' + data.message;

    GmailApp.sendEmail(emailDestino, asuntoInterno, cuerpoInterno);
    if (emailDestino !== emailGeneral) {
      GmailApp.sendEmail(emailGeneral, asuntoInterno, cuerpoInterno);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
