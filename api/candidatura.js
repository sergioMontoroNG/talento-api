/*
 * POST /api/candidatura
 *
 * Recibe el formulario de candidatura espontánea ("Trabaja con nosotros") del
 * tema de Shopify (sections/job-application-form.liquid) y replica el mismo
 * comportamiento que tenía el formulario de Contact Form 7 en WordPress
 * (ver WP/mail):
 *
 *   - Email interno a MAIL_TO (rrhh@newgarden.es) con el CV adjunto.
 *   - Email de confirmación al candidato.
 *
 * El CV se procesa en memoria y se adjunta directamente al correo: no se
 * guarda en ningún sitio ni se sube a almacenamiento de terceros.
 */

const Busboy = require('busboy');

/*
 * Microsoft Graph limita a 3MB el tamaño de un adjunto enviado como
 * contentBytes (base64) dentro de sendMail. El base64 infla el archivo
 * original ~1.33x, así que el límite real de archivo debe quedar más abajo
 * para no superar esos 3MB una vez codificado. Por eso el valor por defecto
 * baja de los 5MB provisionales iniciales a 2MB.
 */
const CV_MAX_SIZE_MB = Number(process.env.CV_MAX_SIZE_MB || 2);
const CV_MAX_SIZE_BYTES = CV_MAX_SIZE_MB * 1024 * 1024;
const ALLOWED_CV_EXTENSIONS = ['pdf', 'doc', 'docx'];
const ALLOWED_CV_MIME_TYPES = [
  'application/pdf',
  'application/msword',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
];

const DEFAULT_ALLOWED_ORIGINS = [
  'https://newgarden.es',
  'https://www.newgarden.es',
  'https://newgardenshop.es',
  'https://newgardenshop.pt',
  'https://newgardenshop.fr',
  'https://newgardenshop.it',
  'https://newgardenshopping.nl',
  'https://newgardenshopping.com',
  'https://newgardenstore.de',
  'https://newgardenshop.us',
  'https://newgardenstore.co.uk',
  'https://newgardenshop.mx',
  'https://proyecto-web-corporativa.myshopify.com'
];

function getAllowedOrigins() {
  if (!process.env.ALLOWED_ORIGINS) return DEFAULT_ALLOWED_ORIGINS;
  return process.env.ALLOWED_ORIGINS.split(',').map((origin) => origin.trim()).filter(Boolean);
}

function applyCors(req, res) {
  const origin = req.headers.origin;
  const allowedOrigins = getAllowedOrigins();

  if (origin && (allowedOrigins.includes(origin) || origin.endsWith('.myshopify.com'))) {
    res.setHeader('Access-Control-Allow-Origin', origin);
    res.setHeader('Vary', 'Origin');
  }

  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');
}

function getFileExtension(filename) {
  const parts = String(filename || '').split('.');
  return parts.length > 1 ? parts.pop().toLowerCase() : '';
}

function isVisiblyEmpty(value) {
  return String(value || '').trim() === '';
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(value || '').trim());
}

/*
 * Parsea el multipart/form-data del request con busboy.
 * Corta la subida del CV en cuanto supera CV_MAX_SIZE_BYTES (busboy limits)
 * en vez de dejar que el archivo entero llegue a memoria.
 */
function parseMultipartRequest(req) {
  return new Promise((resolve, reject) => {
    const busboy = Busboy({
      headers: req.headers,
      limits: { fileSize: CV_MAX_SIZE_BYTES, files: 1 },
    });

    const fields = {};
    let cvFile = null;
    let cvTruncated = false;

    busboy.on('field', (name, value) => {
      fields[name] = value;
    });

    busboy.on('file', (name, fileStream, info) => {
      if (name !== 'cv') {
        fileStream.resume();
        return;
      }

      const chunks = [];
      fileStream.on('data', (chunk) => chunks.push(chunk));
      fileStream.on('limit', () => {
        cvTruncated = true;
      });
      fileStream.on('end', () => {
        cvFile = {
          filename: info.filename,
          mimeType: info.mimeType,
          buffer: Buffer.concat(chunks),
        };
      });
    });

    busboy.on('error', (err) => reject(err));
    busboy.on('finish', () => resolve({ fields, cvFile, cvTruncated }));

    req.pipe(busboy);
  });
}

/*
 * Autenticación de aplicación (client credentials) contra la app ya
 * registrada en Microsoft Entra ID, para enviar por Microsoft Graph sin
 * depender de SMTP AUTH. La app necesita el permiso de aplicación
 * "Mail.Send" con consentimiento de administrador concedido, y MAIL_FROM
 * debe ser un buzón dentro del alcance permitido (todo el tenant, salvo que
 * haya una Application Access Policy que lo restrinja).
 */
async function getGraphAccessToken() {
  const tenantId = process.env.AZURE_TENANT_ID;
  const clientId = process.env.AZURE_CLIENT_ID;
  const clientSecret = process.env.AZURE_CLIENT_SECRET;

  const response = await fetch(`https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      client_id: clientId,
      client_secret: clientSecret,
      scope: 'https://graph.microsoft.com/.default',
      grant_type: 'client_credentials',
    }),
  });

  if (!response.ok) {
    throw new Error('No ha sido posible autenticar con Microsoft Graph.');
  }

  const data = await response.json();
  return data.access_token;
}

async function sendGraphMail(accessToken, { from, to, replyTo, subject, text, attachments }) {
  const message = {
    subject,
    body: { contentType: 'Text', content: text },
    toRecipients: [{ emailAddress: { address: to } }],
  };

  if (replyTo) {
    message.replyTo = [{ emailAddress: { address: replyTo } }];
  }

  if (attachments && attachments.length) {
    message.attachments = attachments.map((attachment) => ({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: attachment.filename,
      contentType: attachment.contentType || 'application/octet-stream',
      contentBytes: attachment.content.toString('base64'),
    }));
  }

  const response = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(from)}/sendMail`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${accessToken}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ message, saveToSentItems: false }),
  });

  if (!response.ok) {
    const errorBody = await response.text().catch(() => '');
    throw new Error(`Microsoft Graph sendMail ha devuelto ${response.status}: ${errorBody}`);
  }
}

function sendJson(res, statusCode, payload) {
  res.status(statusCode).json(payload);
}

module.exports = async function handler(req, res) {
  applyCors(req, res);

  if (req.method === 'OPTIONS') {
    res.status(204).end();
    return;
  }

  if (req.method !== 'POST') {
    sendJson(res, 405, { ok: false, message: 'Método no permitido.' });
    return;
  }

  if (!process.env.AZURE_TENANT_ID || !process.env.AZURE_CLIENT_ID || !process.env.AZURE_CLIENT_SECRET || !process.env.MAIL_FROM) {
    sendJson(res, 500, { ok: false, message: 'El envío de correo no está configurado.' });
    return;
  }

  let fields;
  let cvFile;
  let cvTruncated;

  try {
    ({ fields, cvFile, cvTruncated } = await parseMultipartRequest(req));
  } catch (err) {
    sendJson(res, 400, { ok: false, message: 'No ha sido posible procesar el formulario.' });
    return;
  }

  const name = String(fields.name || '').trim();
  const email = String(fields.email || '').trim();
  const message = String(fields.message || '').trim();
  const acceptPrivacy = String(fields.accept_privacy || '').toLowerCase();

  if (isVisiblyEmpty(name)) {
    sendJson(res, 422, { ok: false, message: 'El nombre es obligatorio.', field: 'name' });
    return;
  }

  if (name.length > 150) {
    sendJson(res, 422, { ok: false, message: 'El nombre no puede superar los 150 caracteres.', field: 'name' });
    return;
  }

  if (!isValidEmail(email)) {
    sendJson(res, 422, { ok: false, message: 'Introduce una dirección de email válida.', field: 'email' });
    return;
  }

  if (message.length > 2000) {
    sendJson(res, 422, { ok: false, message: 'El mensaje no puede superar los 2000 caracteres.', field: 'message' });
    return;
  }

  if (!['1', 'on', 'true', 'yes'].includes(acceptPrivacy)) {
    sendJson(res, 422, { ok: false, message: 'Debes aceptar la política de privacidad.', field: 'accept_privacy' });
    return;
  }

  if (!cvFile || !cvFile.buffer || cvFile.buffer.length === 0) {
    sendJson(res, 422, { ok: false, message: 'Adjunta tu CV en formato .pdf, .doc o .docx.', field: 'cv' });
    return;
  }

  if (cvTruncated) {
    sendJson(res, 422, {
      ok: false,
      message: 'El CV no puede superar los ' + CV_MAX_SIZE_MB + 'MB.',
      field: 'cv',
    });
    return;
  }

  const cvExtension = getFileExtension(cvFile.filename);
  if (
    !ALLOWED_CV_EXTENSIONS.includes(cvExtension) ||
    (cvFile.mimeType && !ALLOWED_CV_MIME_TYPES.includes(cvFile.mimeType))
  ) {
    sendJson(res, 422, {
      ok: false,
      message: 'El CV debe ser un archivo .pdf, .doc o .docx.',
      field: 'cv',
    });
    return;
  }

  const mailFrom = process.env.MAIL_FROM;
  const mailTo = process.env.MAIL_TO || 'rrhh@newgarden.es';

  try {
    const accessToken = await getGraphAccessToken();

    await sendGraphMail(accessToken, {
      from: mailFrom,
      to: mailTo,
      replyTo: email,
      subject: `Formulario de trabajo de ${name}, ${email}`,
      text: [
        `Nombre: ${name}`,
        `Email: ${email}`,
        '',
        'Mensaje:',
        message || '(sin mensaje)',
      ].join('\n'),
      attachments: [
        {
          filename: cvFile.filename,
          content: cvFile.buffer,
          contentType: cvFile.mimeType,
        },
      ],
    });

    await sendGraphMail(accessToken, {
      from: mailFrom,
      to: email,
      replyTo: mailTo,
      /*
       * Este es el mismo texto de "Asunto" que tenía configurado el mail de
       * confirmación en el Contact Form 7 original (WP/mail) — parece un
       * dato de remitente puesto por error en el campo Asunto, pero se
       * mantiene igual para replicar el comportamiento existente.
       * Override disponible vía MAIL_CANDIDATE_SUBJECT si se confirma que
       * fue un error y se quiere corregir.
       */
      subject: process.env.MAIL_CANDIDATE_SUBJECT || 'Recursos Humanos | NEWGARDEN <rrhh@newgarden.es>',
      text: [
        '¡Hola!',
        '',
        'Te confirmamos que hemos recibido tu currículum vitae.',
        '',
        'Los datos facilitados a través de esta solicitud de empleo serán tratados por NEWGARDEN SPAIN, S.L. Estos datos han sido facilitados por el interesado con la única finalidad de participar en los procesos de selección de personal realizados por NEWGARDEN SPAIN, S.L. y no serán cedidos a terceros salvo consentimiento del interesado.',
        '',
        'En el caso de producirse alguna modificación en sus datos, te rogamos que nos lo comuniques. Si no se produjera ninguna actualización, tus datos serán conservados durante un año, siendo posteriormente suprimidos.',
        '',
        'Puedes consultar información adicional y detallada, así como ejercitar tus derechos, solicitándolo al remitente de este correo electrónico.',
        'Agradecemos tu confianza depositada en Newgarden.',
        '',
        '¡Un saludo!',
      ].join('\n'),
    });
  } catch (err) {
    console.error('Error enviando el formulario de candidatura:', err);
    sendJson(res, 502, { ok: false, message: 'No ha sido posible enviar el formulario. Inténtalo más tarde.' });
    return;
  }

  sendJson(res, 200, { ok: true, message: 'Formulario enviado correctamente.' });
};

module.exports.config = {
  api: {
    bodyParser: false,
  },
};
