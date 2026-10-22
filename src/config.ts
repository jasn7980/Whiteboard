// Configuración de la conexión con Microsoft (Outlook).
//
// El Client ID sale del registro de la app en el portal de Microsoft
// (un trámite único con la cuenta de Jaime). No es secreto: en apps de
// navegador el Client ID es público por diseño; la seguridad la da el
// inicio de sesión del propio usuario.
export const MSAL_CLIENT_ID = 'PENDIENTE-CLIENT-ID'

export const MSAL_AUTHORITY = 'https://login.microsoftonline.com/consumers'

export const SCOPES = {
  base: ['User.Read'],
  calendar: ['Calendars.ReadWrite'],
  tasks: ['Tasks.ReadWrite'],
  notes: ['Notes.Create'],
}

export const TIME_ZONE = 'America/Toronto'
