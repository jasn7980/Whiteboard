import { createShapeId, type Editor, type TLAssetId, type TLShapeId } from 'tldraw'

const ASSET_ID = 'asset:muse-gif' as TLAssetId
const SHAPE_ID = createShapeId('muse-gif')
// Ids fijos del GIF, para que la sincronización en la nube pueda
// excluirlo: el GIF es por dispositivo y se vuelve a sembrar al abrir.
export const MUSE_GIF_ASSET_ID = ASSET_ID
export const MUSE_GIF_SHAPE_ID = SHAPE_ID
const OFF_KEY = 'tablero-muse-gif'
const DEFAULT_GIF = 'muse.gif'
const POLL_MS = 180_000

interface MuseLook {
  gif: string
  look: string
}

async function fetchCurrentLook(): Promise<MuseLook> {
  try {
    const res = await fetch(`${import.meta.env.BASE_URL}muse-current.json`, {
      cache: 'no-store',
    })
    if (!res.ok) throw new Error('sin manifiesto')
    const data = (await res.json()) as Partial<MuseLook>
    if (typeof data.gif === 'string' && data.gif.endsWith('.gif')) {
      return { gif: data.gif, look: data.look ?? '' }
    }
  } catch {
    // sin red o sin manifiesto: se usa el look por defecto
  }
  return { gif: DEFAULT_GIF, look: 'gato' }
}

/**
 * Muse vive en el tablero: un GIF animado con sus movimientos de
 * siempre y su disfraz del momento. El disfraz vigente lo dicta
 * muse-current.json (lo actualiza la rotación de Halloween); el
 * tablero lo revisa al abrir y cada 3 minutos, y cambia el GIF sin
 * mover la figura de donde Jaime la haya puesto. Se crea una sola
 * vez (ids fijos, sin duplicados); si Jaime lo borra, no vuelve a
 * aparecer (la baja queda registrada en este navegador).
 */
export function ensureMuseGif(editor: Editor): void {
  const wired = editor as Editor & { __museGifWired?: boolean }
  if (wired.__museGifWired) return
  wired.__museGifWired = true

  // Si Jaime lo borra, respetamos la baja en este navegador.
  editor.store.listen(({ changes }) => {
    if ((changes.removed as Record<string, unknown>)[SHAPE_ID]) {
      try {
        window.localStorage.setItem(OFF_KEY, 'off')
      } catch {
        // sin almacenamiento disponible: se ignora
      }
    }
  })

  let dismissed = false
  try {
    dismissed = window.localStorage.getItem(OFF_KEY) === 'off'
  } catch {
    dismissed = false
  }
  if (dismissed) return

  let currentGif = ''

  const applyLook = (gif: string) => {
    if (gif === currentGif && editor.getShape(SHAPE_ID as TLShapeId)) return
    currentGif = gif
    const src = `${import.meta.env.BASE_URL}${gif}`
    const props = {
      name: gif,
      src,
      w: 208,
      h: 208,
      mimeType: 'image/gif',
      isAnimated: true,
    }
    if (editor.getAsset(ASSET_ID)) {
      editor.updateAssets([{ id: ASSET_ID, type: 'image', props }])
    } else {
      editor.createAssets([
        { id: ASSET_ID, typeName: 'asset', type: 'image', meta: {}, props },
      ])
    }
    if (!editor.getShape(SHAPE_ID as TLShapeId)) {
      editor.createShape({
        id: SHAPE_ID,
        type: 'image',
        x: 36,
        y: 470,
        props: {
          assetId: ASSET_ID,
          w: 150,
          h: 150,
          playing: true,
          url: '',
          crop: null,
          flipX: false,
          flipY: false,
          altText: 'Muse, tu asistente',
        },
      })
    }
  }

  const refresh = () => {
    void fetchCurrentLook().then(({ gif }) => applyLook(gif))
  }
  refresh()
  window.setInterval(refresh, POLL_MS)
}
