import { createShapeId, type Editor, type TLAssetId } from 'tldraw'

// Fotos del tablero: el archivo vive en OneDrive (carpeta propia de
// la app) y el tablero guarda solo la referencia. Así el tablero.json
// no se infla con las fotos adentro y se sincroniza igual que todo.
//
// El asset lleva la ruta en meta.onedrivePath y un src de blob válido
// solo en esta sesión; al abrir (o al cargar la nube) se vuelve a
// resolver desde OneDrive con el token del usuario.

const GRAPH = 'https://graph.microsoft.com/v1.0'
const PHOTOS_DIR = 'fotos'
const MAX_DIM = 1600
const JPEG_QUALITY = 0.85
const MAX_DISPLAY_W = 600

// Cache en memoria: ruta en OneDrive -> URL de blob válida en esta sesión.
const blobCache = new Map<string, string>()

function sanitizeFileName(name: string): string {
  const base = name.replace(/\.[^.]*$/, '').toLowerCase()
  const clean = base.replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '')
  return clean || 'foto'
}

export async function compressPhoto(file: File): Promise<{
  blob: Blob
  width: number
  height: number
}> {
  const url = URL.createObjectURL(file)
  try {
    const img = await new Promise<HTMLImageElement>((resolve, reject) => {
      const el = new Image()
      el.onload = () => resolve(el)
      el.onerror = () =>
        reject(new Error('No se pudo leer esa foto (formato no compatible)'))
      el.src = url
    })
    const scale = Math.min(1, MAX_DIM / Math.max(img.naturalWidth, img.naturalHeight))
    const width = Math.max(1, Math.round(img.naturalWidth * scale))
    const height = Math.max(1, Math.round(img.naturalHeight * scale))
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) throw new Error('No se pudo procesar la foto')
    ctx.fillStyle = '#ffffff'
    ctx.fillRect(0, 0, width, height)
    ctx.drawImage(img, 0, 0, width, height)
    const blob = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob(resolve, 'image/jpeg', JPEG_QUALITY),
    )
    if (!blob) throw new Error('No se pudo procesar la foto')
    return { blob, width, height }
  } finally {
    URL.revokeObjectURL(url)
  }
}

/** Lee el código de error de Graph (p. ej. nameAlreadyExists) para el mensaje. */
async function graphError(res: Response): Promise<string> {
  try {
    const body = await res.json()
    const code = (body as { error?: { code?: string } })?.error?.code
    const message = (body as { error?: { message?: string } })?.error?.message
    return code ? `${code}${message ? `: ${message}` : ''}` : `HTTP ${res.status}`
  } catch {
    return `HTTP ${res.status}`
  }
}

/** Crea la carpeta fotos/ en OneDrive si aún no existe (la subida no crea carpetas). */
export async function ensurePhotosFolder(token: string): Promise<void> {
  // Primero se resuelve la carpeta de la app a su id: POST /children
  // directo sobre special/approot devuelve 400 en OneDrive personal.
  const rootRes = await fetch(`${GRAPH}/me/drive/special/approot`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!rootRes.ok) {
    throw new Error(
      `OneDrive ${await graphError(rootRes)} al abrir la carpeta de la app`,
    )
  }
  const root = (await rootRes.json()) as { id: string }
  const res = await fetch(`${GRAPH}/me/drive/items/${root.id}/children`, {
    method: 'POST',
    headers: {
      Authorization: `Bearer ${token}`,
      'Content-Type': 'application/json',
    },
    body: JSON.stringify({ name: PHOTOS_DIR, folder: {} }),
  })
  // 409 = la carpeta ya existe.
  if (res.ok || res.status === 409) return
  throw new Error(
    `OneDrive ${await graphError(res)} al crear la carpeta fotos`,
  )
}

/** Sube la foto comprimida a la carpeta de la app en OneDrive. Devuelve su ruta. */
export async function uploadPhoto(
  token: string,
  blob: Blob,
  originalName: string,
): Promise<string> {
  const filename = `${Date.now()}-${sanitizeFileName(originalName)}.jpg`
  const path = `${PHOTOS_DIR}/${filename}`
  const res = await fetch(
    `${GRAPH}/me/drive/special/approot:/${PHOTOS_DIR}/${encodeURIComponent(filename)}:/content`,
    {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'image/jpeg',
      },
      body: blob,
    },
  )
  if (!res.ok) throw new Error(`OneDrive ${await graphError(res)} al subir la foto`)
  return path
}

/** Baja la foto de OneDrive y devuelve una URL de blob para mostrarla. */
export async function fetchPhotoBlobUrl(token: string, path: string): Promise<string> {
  const cached = blobCache.get(path)
  if (cached) return cached
  const res = await fetch(`${GRAPH}/me/drive/special/approot:/${path}:/content`, {
    headers: { Authorization: `Bearer ${token}` },
  })
  if (!res.ok) throw new Error(`OneDrive ${await graphError(res)} al bajar la foto`)
  const blob = await res.blob()
  const url = URL.createObjectURL(blob)
  blobCache.set(path, url)
  return url
}

/** Crea la figura de imagen en el tablero para una foto ya subida. */
export function addPhotoToBoard(
  editor: Editor,
  path: string,
  blobUrl: string,
  width: number,
  height: number,
  centerX: number,
  centerY: number,
  index = 0,
): void {
  const displayW = Math.min(width, MAX_DISPLAY_W)
  const displayH = Math.round((height * displayW) / width)
  const assetId = `asset:${crypto.randomUUID()}` as TLAssetId
  editor.createAssets([
    {
      id: assetId,
      typeName: 'asset',
      type: 'image',
      meta: { onedrivePath: path },
      props: {
        name: path.split('/').pop() ?? 'foto',
        src: blobUrl,
        w: width,
        h: height,
        mimeType: 'image/jpeg',
        isAnimated: false,
      },
    },
  ])
  editor.createShape({
    id: createShapeId(),
    type: 'image',
    x: centerX - displayW / 2 + index * 40,
    y: centerY - displayH / 2 + index * 40,
    props: {
      assetId,
      w: displayW,
      h: displayH,
      playing: true,
      url: '',
      crop: null,
      flipX: false,
      flipY: false,
    },
  })
}

/**
 * Resuelve las fotos del tablero: para cada asset con meta.onedrivePath,
 * baja la foto de OneDrive y pone su URL de blob como src. Se ejecuta
 * al abrir y después de cada carga de la nube (la nube guarda el src
 * de otro dispositivo, que en este navegador ya no vale).
 */
export async function resolvePhotoAssets(
  editor: Editor,
  getToken: () => Promise<string | null>,
): Promise<void> {
  const photos = editor
    .getAssets()
    .filter(
      (a) =>
        a.type === 'image' &&
        typeof (a.meta as Record<string, unknown>)?.onedrivePath === 'string',
    )
  if (photos.length === 0) return
  const token = await getToken()
  if (!token) return
  for (const asset of photos) {
    const path = (asset.meta as Record<string, unknown>).onedrivePath as string
    try {
      const url = await fetchPhotoBlobUrl(token, path)
      const current = editor.getAsset(asset.id)
      if (!current || current.props.src === url) continue
      editor.updateAssets([
        { ...current, props: { ...current.props, src: url } },
      ])
    } catch {
      // Sin red o sin permiso: la foto queda sin resolver hasta el próximo intento.
    }
  }
}
