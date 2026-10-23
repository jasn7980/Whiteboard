import { createShapeId, type Editor, type TLAssetId, type TLShapeId } from 'tldraw'

const ASSET_ID = 'asset:muse-gif' as TLAssetId
const SHAPE_ID = createShapeId('muse-gif')
const OFF_KEY = 'tablero-muse-gif'

/**
 * Muse vive en el tablero: un GIF animado con sus movimientos de
 * siempre. Se crea una sola vez (ids fijos, sin duplicados); Jaime lo
 * puede mover a donde quiera, y si lo borra, no vuelve a aparecer
 * (la baja queda registrada en este navegador).
 */
export function ensureMuseGif(editor: Editor): void {
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
  if (editor.getShape(SHAPE_ID as TLShapeId)) return

  if (!editor.getAsset(ASSET_ID)) {
    editor.createAssets([
      {
        id: ASSET_ID,
        typeName: 'asset',
        type: 'image',
        meta: {},
        props: {
          name: 'muse.gif',
          src: `${import.meta.env.BASE_URL}muse.gif`,
          w: 208,
          h: 208,
          mimeType: 'image/gif',
          isAnimated: true,
        },
      },
    ])
  }
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
