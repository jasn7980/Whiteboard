import { useEffect, useState } from 'react'

export interface ScreensaverPhoto {
  src: string
  name: string
}

// Cada foto se muestra 8 segundos, con fundido de 0.6 s entre fotos.
const PHOTO_MS = 8000
const FADE_MS = 600

/**
 * Salvapantallas del tablero: muestra las fotos del tablero a pantalla
 * completa, una a una, sobre fondo negro. Cualquier toque, clic o tecla
 * lo cierra (el padre también escucha la actividad para salir).
 */
export function Screensaver({
  photos,
  onExit,
}: {
  photos: ScreensaverPhoto[]
  onExit: () => void
}) {
  const [index, setIndex] = useState(0)
  const [visible, setVisible] = useState(true)

  useEffect(() => {
    if (photos.length <= 1) return
    const id = window.setInterval(() => {
      setVisible(false)
      window.setTimeout(() => {
        setIndex((i) => (i + 1) % photos.length)
        setVisible(true)
      }, FADE_MS)
    }, PHOTO_MS)
    return () => window.clearInterval(id)
  }, [photos.length])

  useEffect(() => {
    const onKey = () => onExit()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [onExit])

  const photo = photos[index]
  if (!photo) return null
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
      <img
        src={photo.src}
        alt={photo.name}
        style={{
          maxWidth: '100%',
          maxHeight: '100%',
          objectFit: 'contain',
          opacity: visible ? 1 : 0,
          transition: `opacity ${FADE_MS}ms ease`,
        }}
      />
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
        {photo.name} · {index + 1} / {photos.length} · toca para salir
      </div>
    </div>
  )
}
