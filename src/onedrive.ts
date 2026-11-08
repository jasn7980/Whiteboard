import type { Editor, TLStoreSnapshot } from 'tldraw'
import { MUSE_GIF_SHAPE_ID } from './museGif'

const GRAPH = 'https://graph.microsoft.com/v1.0'
// Carpeta propia de la app en OneDrive (se crea sola al primer uso):
// el tablero NO ve ni toca ningún otro archivo del usuario.
const BOARD_PATH = '/me/drive/special/approot:/tablero.json:/content'
const MARKER_KEY = 'tablero-nube-fecha'
const DIRTY_KEY = 'tablero-nube-sucia'

export interface CloudBoardFile {
  version: 1
  savedAt: string
  snapshot: TLStoreSnapshot
}

function getMarker(): string | null {
  try {
    return window.localStorage.getItem(MARKER_KEY)
  } catch {
    return null
  }
}

function setMarker(savedAt: string): void {
  try {
    window.localStorage.setItem(MARKER_KEY, savedAt)
  } catch {
    // sin almacenamiento disponible: se ignora
  }
}

function setDirty(): void {
  try {
    window.localStorage.setItem(DIRTY_KEY, '1')
  } catch {
    // sin almacenamiento disponible: se ignora
  }
}

function clearDirty(): void {
  try {
    window.localStorage.removeItem(DIRTY_KEY)
  } catch {
    // sin almacenamiento disponible: se ignora
  }
}

function isDirty(): boolean {
  try {
    return window.localStorage.getItem(DIRTY_KEY) === '1'
  } catch {
    return false
  }
}

async function fetchWithTimeout(
  url: string,
  init: RequestInit,
  ms = 10000,
): Promise<Response> {
  const controller = new AbortController()
  const timer = window.setTimeout(() => controller.abort(), ms)
  try {
    return await fetch(url, { ...init, signal: controller.signal })
  } finally {
    window.clearTimeout(timer)
  }
}

async function graphError(res: Response): Promise<string> {
  try {
    const body = (await res.json()) as {
      error?: { code?: string; message?: string }
    }
    const code = body.error?.code
    const msg = body.error?.message
    return `OneDrive ${res.status}${code ? ` ${code}` : ''}${msg ? `: ${msg}` : ''}`
  } catch {
    return `OneDrive ${res.status}`
  }
}

/** null si el archivo aún no existe en OneDrive. */
export async function downloadBoardFile(
  token: string,
): Promise<CloudBoardFile | null> {
  const res = await fetchWithTimeout(`${GRAPH}${BOARD_PATH}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (res.status === 404) return null
  if (!res.ok) throw new Error(await graphError(res))
  return (await res.json()) as CloudBoardFile
}

/** La subida real del tablero (un intento, con un reintento si OneDrive
 *  responde 409: ese error es el bloqueo por otra escritura en curso
 *  sobre el mismo archivo y suele bastar con esperar un momento). */
async function uploadBoardFileNow(
  token: string,
  editor: Editor,
): Promise<string> {
  const savedAt = new Date().toISOString()
  const file: CloudBoardFile = {
    version: 1,
    savedAt,
    // Solo el documento (tinta y tarjetas), sin la sesión/cámara.
    // El GIF de Muse viaja como cualquier figura: tiene ids fijos,
    // así que no se duplica entre dispositivos.
    snapshot: editor.getSnapshot().document,
  }
  for (let attempt = 0; attempt < 2; attempt++) {
    const res = await fetchWithTimeout(`${GRAPH}${BOARD_PATH}`, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify(file),
    })
    if (res.ok) {
      setMarker(savedAt)
      clearDirty()
      return savedAt
    }
    // 409 = otra escritura sobre tablero.json está en curso (por
    // ejemplo, el guardado automático que se dispara 2 s después de
    // un cambio y un toque manual a ⇧ Nube al mismo tiempo).
    if (res.status === 409 && attempt === 0) {
      await new Promise((r) => window.setTimeout(r, 1500))
      continue
    }
    throw new Error(await graphError(res))
  }
  throw new Error('OneDrive 409')
}

/**
 * Sube el tablero actual a OneDrive. Devuelve la fecha guardada.
 * Las subidas van en serie (una sola cadena para toda la app): el
 * guardado automático y el manual a ⇧ Nube ya no pueden chocar con
 * dos PUT al mismo archivo a la vez, que es lo que OneDrive
 * respondía con 409.
 */
let uploadChain: Promise<string> = Promise.resolve('')
export async function uploadBoardFile(
  token: string,
  editor: Editor,
): Promise<string> {
  const run = uploadChain.then(() => uploadBoardFileNow(token, editor))
  // Que la cadena no se rompa si esta subida falla: la siguiente
  // sigue intentándolo.
  uploadChain = run.catch(() => '')
  return run
}

/**
 * Sincronización inicial al abrir el tablero. Nunca redirige: solo usa
 * el token en silencio, así que sin sesión o sin permiso termina como
 * hoy (tablero local) sin molestar.
 *
 * Reglas (el último en guardar gana, pero el trabajo del dispositivo
 * que tienes en la mano nunca se borra en silencio al abrir):
 * - Hay cambios locales sin subir (se cerró antes de guardar) → se sube
 *   el tablero local, sin descargar nada.
 * - No hay nube → el tablero local con contenido real se sube como
 *   semilla (si está vacío, no se sube nada).
 * - Hay nube y este dispositivo ya sincronizó → si la nube es más
 *   nueva, se carga.
 * - Hay nube y este dispositivo nunca sincronizó → si el tablero local
 *   tiene contenido real, se sube el local (protege el tablero de la
 *   Chromebook frente a una semilla de otro dispositivo); si está
 *   vacío o solo tiene el GIF, se carga la nube.
 * Devuelve true si quedó conectado a la nube.
 */
export async function initCloudBoard(
  editor: Editor,
  getFilesTokenSilent: () => Promise<string | null>,
): Promise<boolean> {
  try {
    const token = await getFilesTokenSilent()
    if (!token) return false
    if (isDirty()) {
      await uploadBoardFile(token, editor)
      return true
    }
    const cloud = await downloadBoardFile(token)
    const hasRealContent = editor
      .getCurrentPageShapes()
      .some((s) => s.id !== MUSE_GIF_SHAPE_ID)
    if (!cloud) {
      if (hasRealContent) {
        await uploadBoardFile(token, editor)
        return true
      }
      return false
    }
    const marker = getMarker()
    if (marker) {
      if (cloud.savedAt > marker) {
        editor.loadSnapshot(cloud.snapshot)
        setMarker(cloud.savedAt)
      }
      return true
    }
    if (!hasRealContent) {
      editor.loadSnapshot(cloud.snapshot)
      setMarker(cloud.savedAt)
      return true
    }
    await uploadBoardFile(token, editor)
    return true
  } catch {
    return false
  }
}

/**
 * Guarda el tablero en OneDrive 2 s después del último cambio (y marca
 * los cambios como pendientes hasta que la subida tenga éxito, para
 * que un cierre rápido no pierda nada: al reabrir, lo pendiente se
 * sube antes de descargar). Sin token en silencio no hace nada: la
 * copia local sigue funcionando como siempre.
 */
export function attachCloudSave(
  editor: Editor,
  getFilesTokenSilent: () => Promise<string | null>,
  onSaved: () => void,
): () => void {
  let timer: number | undefined
  let saving = false
  let pending = false
  const save = async () => {
    if (saving) {
      pending = true
      return
    }
    saving = true
    try {
      const token = await getFilesTokenSilent()
      if (token) {
        await uploadBoardFile(token, editor)
        onSaved()
      }
    } catch {
      // Sin red o sin permiso: queda la copia local, se reintenta
      // en el próximo cambio.
    } finally {
      saving = false
      if (pending) {
        pending = false
        timer = window.setTimeout(save, 2000)
      }
    }
  }
  const dispose = editor.store.listen(
    () => {
      setDirty()
      window.clearTimeout(timer)
      timer = window.setTimeout(save, 2000)
    },
    { scope: 'document' },
  )
  return () => {
    window.clearTimeout(timer)
    dispose()
  }
}

/** ¿Hay cambios locales sin subir a OneDrive? */
export function isCloudDirty(): boolean {
  return isDirty()
}

/** Registra que este dispositivo ya tiene la versión de la nube indicada. */
export function setCloudMarker(savedAt: string): void {
  setMarker(savedAt)
}

/**
 * Descarga el tablero de OneDrive solo si es más nuevo que la última
 * versión que este dispositivo conoce. null si no hay nada nuevo
 * (o si este dispositivo nunca sincronizó: ese caso lo decide
 * initCloudBoard al abrir, no el sondeo).
 */
export async function fetchNewerCloudBoard(
  token: string,
): Promise<CloudBoardFile | null> {
  const marker = getMarker()
  if (!marker) return null
  const cloud = await downloadBoardFile(token)
  if (!cloud) return null
  return cloud.savedAt > marker ? cloud : null
}
