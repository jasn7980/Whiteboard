import type { Editor, TLStoreSnapshot } from 'tldraw'
import { MUSE_GIF_ASSET_ID, MUSE_GIF_SHAPE_ID } from './museGif'

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

/**
 * El GIF de Muse es por dispositivo (su baja queda en el tombstone
 * local) y se vuelve a sembrar al abrir: no viaja en la nube.
 */
function stripMuseGif(snapshot: TLStoreSnapshot): TLStoreSnapshot {
  const store = { ...(snapshot.store as Record<string, unknown>) }
  delete store[MUSE_GIF_SHAPE_ID as string]
  delete store[MUSE_GIF_ASSET_ID as string]
  return { ...snapshot, store: store as TLStoreSnapshot['store'] }
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

/** null si el archivo aún no existe en OneDrive. */
export async function downloadBoardFile(
  token: string,
): Promise<CloudBoardFile | null> {
  const res = await fetchWithTimeout(`${GRAPH}${BOARD_PATH}`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (res.status === 404) return null
  if (!res.ok) throw new Error(`OneDrive ${res.status}`)
  return (await res.json()) as CloudBoardFile
}

/** Sube el tablero actual a OneDrive. Devuelve la fecha guardada. */
export async function uploadBoardFile(
  token: string,
  editor: Editor,
): Promise<string> {
  const savedAt = new Date().toISOString()
  const file: CloudBoardFile = {
    version: 1,
    savedAt,
    // Solo el documento (tinta y tarjetas), sin la sesión/cámara.
    snapshot: stripMuseGif(editor.getSnapshot().document),
  }
  const res = await fetchWithTimeout(`${GRAPH}${BOARD_PATH}`, {
    method: 'PUT',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify(file),
  })
  if (!res.ok) throw new Error(`OneDrive ${res.status}`)
  setMarker(savedAt)
  clearDirty()
  return savedAt
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
