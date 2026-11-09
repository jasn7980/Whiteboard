import { useCallback, useEffect, useRef, useState } from 'react'

export interface ScreensaverPhoto {
  kind: 'photo'
  src: string
  name: string
}

export interface ScreensaverVideo {
  kind: 'video'
  /** Póster (la miniatura que se ve en el tablero). */
  src: string
  name: string
  /** URL del video (blob: descargado de OneDrive); '' si no se pudo descargar. */
  videoSrc: string
}

export type ScreensaverItem = ScreensaverPhoto | ScreensaverVideo

// Cada foto se muestra 8 segundos, con fundido de 0.6 s entre fotos.
// Un video se reproduce COMPLETO (hasta que termina) y luego sigue la
// rotación; si es el único ítem, se repite en bucle.
const PHOTO_MS = 8000
const FADE_MS = 600

/**
 * Salvapantallas del tablero: muestra las fotos y videos del tablero a
 * pantalla completa, uno a uno, sobre fondo negro. Cualquier toque, clic
 * o tecla lo cierra (el padre también escucha la actividad para salir).
 */
export function Screensaver({
  items,
  onExit,
}: {
  items: ScreensaverItem[]
  onExit: () => void
}) {
  const [index, setIndex] = useState(0)
  const [visible, setVisible] = useState(true)
  const [muted, setMuted] = useState(false)
  const videoRef = useRef<HTMLVideoElement | null>(null)

  const safeIndex = Math.min(index, Math.max(items.length - 1, 0))
  const item = items[safeIndex]

  const advance = useCallback(() => {
    setVisible(false)
    window.setTimeout(() => {
      setIndex((i) => (i + 1) % items.length)
      setVisible(true)
    }, FADE_MS)
  }, [items.length])

  // Las fotos avanzan solas a los 8 s. Un video con URL avanza solo
  // cuando termina (onEnded); sin URL se muestra el póster como foto.
  useEffect(() => {
    if (!item || items.length <= 1) return
    if (item.kind === 'video' && item.videoSrc) return
    const id = window.setTimeout(advance, PHOTO_MS)
    return () => window.clearTimeout(id)
  }, [item, items.length, advance])

  // El sonido: el navegador puede bloquear la reproducción con sonido
  // si no hubo un gesto reciente. Se intenta con sonido y, si lo
  // bloquea, se pasa a mudo para que la rotación no se detenga.
  useEffect(() => {
    const v = videoRef.current
    if (!v) return
    setMuted(false)
    v.muted = false
    const p = v.play()
    if (p) {
      p.catch(() => {
        v.muted = true
        setMuted(true)
        v.play().catch(() => {})
      })
    }
  }, [safeIndex])

  useEffect(() => {
    const onKey = () => onExit()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onExit])

  if (!item) return null
  const isVideo = item.kind === 'video' && !!item.videoSrc
  return (
    <div
      onClick={onExit}
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 1000,
        background: '#000',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        cursor: 'pointer',
      }}
    >
      {isVideo ? (
        <video
          key={(item as ScreensaverVideo).videoSrc}
          ref={videoRef}
          src={(item as ScreensaverVideo).videoSrc}
          autoPlay
          playsInline
          loop={items.length === 1}
          onEnded={() => {
            if (items.length > 1) advance()
          }}
          onError={() => {
            if (items.length > 1) advance()
          }}
          style={{
            maxWidth: '100%',
            maxHeight: '100%',
            objectFit: 'contain',
            opacity: visible ? 1 : 0,
            transition: `opacity ${FADE_MS}ms ease`,
          }}
        />
      ) : (
        <img
          src={item.src}
          alt={item.name}
          style={{
            maxWidth: '100%',
            maxHeight: '100%',
            objectFit: 'contain',
            opacity: visible ? 1 : 0,
            transition: `opacity ${FADE_MS}ms ease`,
          }}
        />
      )}
      <div
        style={{
          position: 'absolute',
          bottom: 24,
          left: 0,
          right: 0,
          textAlign: 'center',
          color: 'rgba(255,255,255,0.7)',
          fontSize: 13,
          fontFamily: 'sans-serif',
          pointerEvents: 'none',
        }}
      >
        {item.kind === 'video' ? '▶ ' : ''}
        {item.name} · {safeIndex + 1} / {items.length}
        {item.kind === 'video' && !item.videoSrc ? ' · video no disponible' : ''}
        {isVideo && muted ? ' · 🔇 sin sonido' : ''} · toca para salir
      </div>
    </div>
  )
}
