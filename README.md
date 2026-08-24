# talento-api

Backend del formulario de candidatura espontánea ("Trabaja con nosotros") de la
web corporativa de Shopify. Reemplaza al formulario de Contact Form 7 que
existía en WordPress: recibe el nombre, email, mensaje y CV del candidato, y
envía dos correos (interno a RRHH con el CV adjunto, y confirmación al
candidato), tal como hacía WordPress.

No guarda nada: el CV se procesa en memoria y se adjunta directamente al
correo saliente, sin pasar por ningún almacenamiento propio ni de terceros.

## Endpoint

`POST /api/candidatura` — `multipart/form-data` con los campos:

- `name` (texto, obligatorio)
- `email` (texto, obligatorio)
- `message` (texto, opcional)
- `accept_privacy` (`on`/`1`/`true`, obligatorio)
- `cv` (archivo `.pdf`/`.doc`/`.docx`, obligatorio)

Responde JSON `{ ok: true/false, message, field? }`.

Los emails se envían por Microsoft Graph (client credentials) usando la app
ya registrada en Microsoft Entra ID, en vez de SMTP. La app necesita el
permiso de aplicación **Mail.Send** con consentimiento de administrador
concedido, y `MAIL_FROM` debe ser un buzón dentro del alcance que tenga
permitido esa app.

## Desplegar en Vercel

1. `vercel` (o importar esta carpeta como proyecto desde el dashboard de Vercel).
2. Configurar las variables de entorno del proyecto en Vercel a partir de
   `.env.example` (tenant/client id/secret de la app de Entra ID).
3. Copiar la URL del deploy (`https://<proyecto>.vercel.app/api/candidatura`)
   y pegarla en el ajuste **"URL del backend (API)"** de la sección
   *Formulario de candidatura* en el editor de temas de Shopify.

## Desarrollo local

```
npm install
vercel dev
```
