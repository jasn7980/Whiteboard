import { createShapeId, type Editor, type TLAssetId } from 'tldraw'

// Videos del tablero: el archivo original vive en OneDrive (carpeta
// propia de la app, videos/) y el tablero guarda solo un póster (una
// foto del primer segundo) como data URL en el asset, igual que las
// fotos. tldraw no reproduce video, así que el video completo solo se
// ve en el salvapantallas, que lo descarga de OneDrive y lo reproduce
// hasta el final. La ruta en OneDrive queda en meta.onedrivePath y
// meta.isVideo marca el asset como video.

const GRAPH = 'https://graph.microsoft.com/v1.0'
const VIDEOS_DIR = 'videos'
const MAX_DISPLAY_W = 600
const POSTER_MAX_DIM = 960
// La subida simple (PUT) solo acepta archivos de hasta 4 MB; más
// grandes van por sesión reanudable, en trozos de 5 MB.
const SIMPLE_PUT_MAX = 4 * 1024 * 1024
const CHUNK_SIZE = 5 * 1024 * 1024

function sanitizeVideoFileName(name: string): string {
  const match = name.match(/\.([a-z0-9]+)$/i)
  const ext = (match?.[1] ?? 'mp4').toLowerCase()
  const base = name
    .replace(/\.[^.]*$/, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-+|-+$/g, '')
  return `${base || 'video'}.${ext}`
}

/** Lee el código de error de Graph para el mensaje. */
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

/** Crea la carpeta videos/ en OneDrive si aún no existe. */
export async function ensureVideosFolder(token: string): Promise<void> {
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
    body: JSON.stringify({ name: VIDEOS_DIR, folder: {} }),
  })
  // 409 = la carpeta ya existe.
  if (res.ok || res.status === 409) return
  throw new Error(
    `OneDrive ${await graphError(res)} al crear la carpeta videos`,
  )
}

/**
 * Saca el póster del video (un fotograma cerca del inicio) como JPEG.
 * Si el navegador no puede decodificar el video, falla con mensaje claro.
 */
export async function extractPoster(file: File): Promise<{
  blob: Blob
  width: number
  height: number
}> {
  const url = URL.createObjectURL(file)
  try {
    const video = document.createElement('video')
    video.muted = true
    video.playsInline = true
    video.preload = 'auto'
    video.src = url
    await new Promise<void>((resolve, reject) => {
      video.onloadedmetadata = () => resolve()
      video.onerror = () =>
        reject(new Error('No se pudo leer ese video (formato no compatible)'))
    })
    if (!video.videoWidth || !video.videoHeight) {
      throw new Error('No se pudo leer ese video (formato no compatible)')
    }
    const duration = Number.isFinite(video.duration) ? video.duration : 0
    video.currentTime = duration > 1 ? Math.min(1, duration / 4) : 0
    // Algunos formatos nunca disparan 'seeked': no quedarse colgado.
    await Promise.race([
      new Promise<void>((resolve) => {
        video.onseeked = () => resolve()
      }),
      new Promise<void>((resolve) => window.setTimeout(resolve, 3000)),
    ])
    const scale = Math.min(
      1,
      POSTER_MAX_DIM / Math.max(video.videoWidth, video.videoHeight),
    )
    const width = Math.max(1, Math.round(video.videoWidth * scale))
    const height = Math.max(1, Math.round(video.videoHeight * scale))
    const canvas = document.createElement('canvas')
    canvas.width = width
    canvas.height = height
    const ctx = canvas.getContext('2d')
    if (!ctx) throw new Error('No se pudo procesar el video')
    ctx.drawImage(video, 0, 0, width, height)
    const blob = await new Promise<Blob | null>((resolve) =>
      canvas.toBlob(resolve, 'image/jpeg', 0.85),
    )
    if (!blob) throw new Error('No se pudo procesar el video')
    return { blob, width, height }
  } finally {
    URL.revokeObjectURL(url)
  }
}

function encodeDrivePath(path: string): string {
  return path
    .split('/')
    .map((seg) => encodeURIComponent(seg))
    .join('/')
}

/** Sube el video a la carpeta de la app en OneDrive. Devuelve su ruta. */
export async function uploadVideo(token: string, file: File): Promise<string> {
  const filename = `${Date.now()}-${sanitizeVideoFileName(file.name)}`
  const path = `${VIDEOS_DIR}/${filename}`
  const contentType = file.type || 'video/mp4'
  if (file.size <= SIMPLE_PUT_MAX) {
    const res = await fetch(
      `${GRAPH}/me/drive/special/approot:/${VIDEOS_DIR}/${encodeURIComponent(filename)}:/content`,
      {
        method: 'PUT',
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': contentType,
        },
        body: file,
      },
    )
    if (!res.ok)
      throw new Error(`OneDrive ${await graphError(res)} al subir el video`)
    return path
  }
  // Sesión reanudable: trozos de 5 MB (múltiplo de 320 KB, como pide Graph).
  const sessionRes = await fetch(
    `${GRAPH}/me/drive/special/approot:/${VIDEOS_DIR}/${encodeURIComponent(filename)}:/createUploadSession`,
    {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: JSON.stringify({
        item: {
          '@microsoft.graph.conflictBehavior': 'rename',
          name: filename,
        },
      }),
    },
  )
  if (!sessionRes.ok) {
    throw new Error(
      `OneDrive ${await graphError(sessionRes)} al preparar la subida del video`,
    )
  }
  const session = (await sessionRes.json()) as { uploadUrl?: string }
  if (!session.uploadUrl) {
    throw new Error('OneDrive no devolvió la dirección para subir el video')
  }
  let start = 0
  while (start < file.size) {
    const end = Math.min(start + CHUNK_SIZE, file.size)
    const chunk = file.slice(start, end)
    const res = await fetch(session.uploadUrl, {
      method: 'PUT',
      headers: {
        'Content-Length': String(chunk.size),
        'Content-Range': `bytes ${start}-${end - 1}/${file.size}`,
      },
      body: chunk,
    })
    if (!res.ok && res.status !== 202) {
      throw new Error(
        `OneDrive ${await graphError(res)} al subir el video (trozo ${start}-${end})`,
      )
    }
    if (res.status === 200 || res.status === 201) return path
    start = end
  }
  return path
}

/** Descarga un video de la carpeta videos/ en OneDrive. */
export async function downloadVideo(token: string, path: string): Promise<Blob> {
  const res = await fetch(
    `${GRAPH}/me/drive/special/approot:/${encodeDrivePath(path)}:/content`,
    { headers: { Authorization: `Bearer ${token}` } },
  )
  if (!res.ok) {
    throw new Error(`OneDrive ${await graphError(res)} al descargar el video`)
  }
  return res.blob()
}

/** Huellas SHA-256 de los videos que ya están en el tablero. */
export async function boardVideoHashes(editor: Editor): Promise<Set<string>> {
  const hashes = new Set<string>()
  for (const asset of editor.getAssets()) {
    if (asset.type !== 'image') continue
    const meta = asset.meta as Record<string, unknown> | undefined
    if (meta?.isVideo !== true) continue
    if (typeof meta.videoHash === 'string' && meta.videoHash) {
      hashes.add(meta.videoHash)
    }
  }
  return hashes
}

/** Rutas en OneDrive de los videos que ya están en el tablero. */
export function boardVideoPaths(editor: Editor): Set<string> {
  const paths = new Set<string>()
  for (const asset of editor.getAssets()) {
    if (asset.type !== 'image') continue
    const meta = asset.meta as Record<string, unknown> | undefined
    if (meta?.isVideo === true && typeof meta.onedrivePath === 'string') {
      paths.add(meta.onedrivePath)
    }
  }
  return paths
}

/** Crea la figura de imagen en el tablero para el póster de un video. */
export function addVideoToBoard(
  editor: Editor,
  path: string,
  posterSrc: string,
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
      // isVideo marca el asset: el salvapantallas lo reproduce entero
      // y la lista de fotos lo excluye (solo ve el póster).
      meta: {
        onedrivePath: path,
        isVideo: true,
        ...(hash ? { videoHash: hash } : {}),
      },
      props: {
        name: `▶ ${path.split('/').pop() ?? 'video'}`,
        src: posterSrc,
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
