// Configuración de la conexión con Microsoft (Outlook).
//
// El Client ID sale del registro de la app en el portal de Microsoft
// (un trámite único con la cuenta de Jaime). No es secreto: en apps de
// navegador el Client ID es público por diseño; la seguridad la da el
// inicio de sesión del propio usuario.
export const MSAL_CLIENT_ID: string = '82cc697d-8b4b-4bff-a761-ca94bdc5f952'

export const MSAL_AUTHORITY = 'https://login.microsoftonline.com/consumers'

export const SCOPES = {
  base: ['User.Read'],
  calendar: ['Calendars.ReadWrite'],
  tasks: ['Tasks.ReadWrite'],
  notes: ['Notes.Create'],
  // Solo la carpeta propia de la app en OneDrive: el tablero no ve
  // ni toca ningún otro archivo del usuario.
  files: ['Files.ReadWrite.AppFolder'],
}

// Todos los permisos juntos al conectar: así Microsoft pide el
// consentimiento una sola vez y crear tareas/citas/ideas después
// funciona en un solo clic, sin redirecciones a mitad de la acción.
export const ALL_SCOPES = [
  ...SCOPES.base,
  ...SCOPES.calendar,
  ...SCOPES.tasks,
  ...SCOPES.notes,
  ...SCOPES.files,
]

export const TIME_ZONE = 'America/Toronto'

// Clave de licencia de tldraw (Trial/Hobby). Sin una clave válida, el SDK
// oculta el editor a los pocos segundos en producción. La clave viaja en el
// paquete público de la app (es un identificador del proyecto, no un secreto).
export const TLDRAW_LICENSE_KEY: string =
  'tldraw-2027-01-31/WyIxOTRmWFpsMiIsWyIqIl0sMTYsIjIwMjctMDEtMzEiXQ.raA3TBAqRB+bV5lGQWSgYOHsrENgMxR3PB4i2mKEdwO42VWdLwemZbfbA9E/qvQrcs0+QJk2azBwth0VYFn7Yg'
