import { createShapeId, type Editor, type TLAssetId } from 'tldraw'

// Fotos del tablero: el archivo original vive en OneDrive (carpeta
// propia de la app) y el tablero guarda la foto comprimida como data
// URL en el asset. tldraw solo acepta src http(s), data: o asset: —
// una URL blob: rompe la validación del asset y tumba el tablero.
// La ruta en OneDrive queda en meta.onedrivePath como referencia.

const GRAPH = 'https://graph.microsoft.com/v1.0'
const PHOTOS_DIR = 'fotos'
const MAX_DIM = 1600
const JPEG_QUALITY = 0.85
const MAX_DISPLAY_W = 600

/** SHA-256 (hex) de unos bytes: identifica el contenido de la foto. */
export async function hashBytes(bytes: Uint8Array): Promise<string> {
  const digest = await crypto.subtle.digest('SHA-256', bytes as BufferSource)
  return [...new Uint8Array(digest)]
    .map((b) => b.toString(16).padStart(2, '0'))
    .join('')
}

/** SHA-256 de un blob (la foto ya comprimida). */
export async function hashBlob(blob: Blob): Promise<string> {
  return hashBytes(new Uint8Array(await blob.arrayBuffer()))
}

/** SHA-256 del contenido de una data URL (las fotos ya guardadas en el tablero). */
export async function hashDataUrl(dataUrl: string): Promise<string | null> {
  const comma = dataUrl.indexOf(',')
  if (!dataUrl.startsWith('data:') || comma < 0) return null
  const binary = atob(dataUrl.slice(comma + 1))
  const bytes = new Uint8Array(binary.length)
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i)
  return hashBytes(bytes)
}

/** Convierte un blob a data URL (la forma en que tldraw guarda imágenes locales). */
export function blobToDataUrl(blob: Blob): Promise<string> {
  return new Promise((resolve, reject) => {
    const reader = new FileReader()
    reader.onload = () => resolve(reader.result as string)
    reader.onerror = () => reject(new Error('No se pudo leer esa foto'))
    reader.readAsDataURL(blob)
  })
}

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

/** Huellas SHA-256 de las fotos que ya están en el tablero. */
export async function boardPhotoHashes(editor: Editor): Promise<Set<string>> {
  const hashes = new Set<string>()
  for (const asset of editor.getAssets()) {
    if (asset.type !== 'image') continue
    const meta = asset.meta as Record<string, unknown> | undefined
    if (typeof meta?.onedrivePath !== 'string') continue
    try {
      const stored =
        typeof meta.photoHash === 'string' ? meta.photoHash : null
      const hash =
        stored ??
        (await hashDataUrl((asset.props as { src?: string }).src ?? ''))
      if (hash) hashes.add(hash)
    } catch {
      // Sin huella no se puede comparar: esa foto no bloquea nada.
    }
  }
  return hashes
}

/** Rutas en OneDrive de las fotos que ya están en el tablero. */
export function boardPhotoPaths(editor: Editor): Set<string> {
  const paths = new Set<string>()
  for (const asset of editor.getAssets()) {
    if (asset.type !== 'image') continue
    const meta = asset.meta as Record<string, unknown> | undefined
    if (typeof meta?.onedrivePath === 'string') paths.add(meta.onedrivePath)
  }
  return paths
}

function encodeDrivePath(path: string): string {
  return path
    .split('/')
    .map((seg) => encodeURIComponent(seg))
    .join('/')
}

export interface OneDrivePhoto {
  name: string
  path: string
  downloadUrl: string
}

const IMAGE_EXT = /\.(jpe?g|png|gif|webp|heic|heif|bmp|avif)$/i

/**
 * Lista las fotos de la carpeta fotos/ en OneDrive (la carpeta propia
 * de la app: con el permiso actual el tablero no ve más allá de ella).
 * downloadUrl es un enlace directo y temporal, útil como miniatura.
 */
export async function listAppPhotos(token: string): Promise<OneDrivePhoto[]> {
  const res = await fetch(
    `${GRAPH}/me/drive/special/approot:/${PHOTOS_DIR}:/children`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  if (!res.ok)
    throw new Error(
      `OneDrive ${await graphError(res)} al listar las fotos`,
    )
  const body = (await res.json()) as {
    value?: Array<{
      name: string
      file?: unknown
      '@microsoft.graph.downloadUrl'?: string
    }>
  }
  return (body.value ?? [])
    .filter((i) => i.file && IMAGE_EXT.test(i.name))
    .map((i) => ({
      name: i.name,
      path: `${PHOTOS_DIR}/${i.name}`,
      downloadUrl: i['@microsoft.graph.downloadUrl'] ?? '',
    }))
}

/** Descarga una foto de la carpeta fotos/ en OneDrive. */
export async function downloadPhoto(
  token: string,
  path: string,
): Promise<Blob> {
  const res = await fetch(
    `${GRAPH}/me/drive/special/approot:/${encodeDrivePath(path)}:/content`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  if (!res.ok)
    throw new Error(
      `OneDrive ${await graphError(res)} al descargar la foto`,
    )
  return res.blob()
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

/** Crea la figura de imagen en el tablero para una foto ya subida. */
export function addPhotoToBoard(
  editor: Editor,
  path: string,
  src: string,
  width: number,
  height: number,
  centerX: number,
  centerY: number,
  index = 0,
  hash = '',
): void {
  const displayW = Math.min(width, MAX_DISPLAY_W)
  const displayH = Math.round((height * displayW) / width)
  const assetId = `asset:${crypto.randomUUID()}` as TLAssetId
  editor.createAssets([
    {
      id: assetId,
      typeName: 'asset',
      type: 'image',
      // photoHash identifica el contenido: evita subir dos veces la
      // misma foto (el control de duplicados lo compara con el tablero).
      meta: { onedrivePath: path, ...(hash ? { photoHash: hash } : {}) },
      props: {
        name: path.split('/').pop() ?? 'foto',
        src,
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
