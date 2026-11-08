import type { OneDrivePhoto } from './photos'

/**
 * Selector de fotos ya guardadas en OneDrive (carpeta fotos/ de la app).
 * Muestra miniaturas; las que ya están en el tablero llevan la marca
 * «✓ ya en el tablero». Elegir una la agrega al tablero sin volver a
 * subirla desde el dispositivo.
 */
export function PhotoPicker({
  loading,
  photos,
  onBoardPaths,
  busyPath,
  note,
  onPick,
  onClose,
}: {
  loading: boolean
  photos: OneDrivePhoto[]
  onBoardPaths: Set<string>
  busyPath: string | null
  note: string | null
  onPick: (photo: OneDrivePhoto) => void
  onClose: () => void
}) {
  return (
    <div
      onClick={onClose}
      style={{
        position: 'fixed',
        inset: 0,
        zIndex: 900,
        background: 'rgba(0,0,0,0.5)',
        display: 'flex',
        alignItems: 'center',
        justifyContent: 'center',
        fontFamily: 'sans-serif',
      }}
    >
      <div
        onClick={(e) => e.stopPropagation()}
        style={{
          background: '#fff',
          borderRadius: 16,
          padding: 16,
          maxWidth: 560,
          width: 'calc(100% - 32px)',
          maxHeight: '80vh',
          overflowY: 'auto',
          boxShadow: '0 8px 32px rgba(0,0,0,0.3)',
        }}
      >
        <div
          style={{
            display: 'flex',
            justifyContent: 'space-between',
            alignItems: 'center',
            marginBottom: 12,
          }}
        >
          <strong style={{ fontSize: 16 }}>Fotos en OneDrive</strong>
          <button
            onClick={onClose}
            style={{
              border: '1px solid #ddd',
              background: '#fff',
              borderRadius: 8,
              padding: '4px 10px',
              cursor: 'pointer',
              fontSize: 14,
            }}
          >
            ✕
          </button>
        </div>
        {note ? (
          <div
            style={{
              background: '#fff3cd',
              border: '1px solid #e6c200',
              borderRadius: 8,
              padding: '8px 10px',
              fontSize: 13,
              marginBottom: 12,
            }}
          >
            {note}
          </div>
        ) : null}
        {loading ? (
          <p style={{ color: '#666' }}>Buscando fotos en OneDrive…</p>
        ) : photos.length === 0 ? (
          <p style={{ color: '#666' }}>
            No hay fotos en la carpeta de OneDrive todavía. Sube una con 📷
            Foto y aparecerá aquí.
          </p>
        ) : (
          <div
            style={{
              display: 'grid',
              gridTemplateColumns: 'repeat(3, 1fr)',
              gap: 12,
            }}
          >
            {photos.map((p) => {
              const onBoard = onBoardPaths.has(p.path)
              return (
                <button
                  key={p.path}
                  onClick={() => onPick(p)}
                  disabled={busyPath === p.path}
                  style={{
                    border: '1px solid #ddd',
                    borderRadius: 12,
                    padding: 6,
                    background: '#fff',
                    cursor: 'pointer',
                    textAlign: 'left',
                  }}
                >
                  {p.downloadUrl ? (
                    <img
                      src={p.downloadUrl}
                      alt={p.name}
                      style={{
                        width: '100%',
                        height: 110,
                        objectFit: 'cover',
                        borderRadius: 8,
                      }}
                    />
                  ) : (
                    <div
                      style={{
                        height: 110,
                        display: 'flex',
                        alignItems: 'center',
                        justifyContent: 'center',
                        background: '#f0f0f0',
                        borderRadius: 8,
                        color: '#999',
                        fontSize: 12,
                      }}
                    >
                      sin vista previa
                    </div>
                  )}
                  <div
                    style={{
                      fontSize: 12,
                      marginTop: 6,
                      wordBreak: 'break-all',
                    }}
                  >
                    {p.name}
                  </div>
                  {onBoard ? (
                    <div style={{ fontSize: 11, color: '#2e7d32' }}>
                      ✓ ya en el tablero
                    </div>
                  ) : null}
                  {busyPath === p.path ? (
                    <div style={{ fontSize: 11 }}>Agregando…</div>
                  ) : null}
                </button>
              )
            })}
          </div>
        )}
      </div>
    </div>
  )
}
