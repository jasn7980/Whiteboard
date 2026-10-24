import { useCallback, useEffect, useRef, useState } from 'react'
import { Editor, Tldraw, createShapeId } from 'tldraw'
import type { AccountInfo } from '@azure/msal-browser'
import 'tldraw/tldraw.css'
import { UniversalCardUtil } from './UniversalCard'
import type { UniversalCardShapeProps } from './UniversalCard'
import { getAccount, getToken, getTokenSilent, isConfigured, signIn, signOut } from './auth'
import { fetchUpcomingEvents } from './graph'
import { fitCardSize } from './fit'
import { ensureMuseGif } from './museGif'
import { syncCalendarEvents } from './sync'
import { SCOPES, TLDRAW_LICENSE_KEY } from './config'

const shapeUtils = [UniversalCardUtil]

function createCard(editor: Editor, kind: string, title: string) {
  const bounds = editor.getViewportPageBounds()
  const jitter = Math.random() * 60 - 30
  const base: UniversalCardShapeProps = {
    w: 212,
    h: 120,
    kind,
    title,
    body: '',
    done: false,
  }
  // La tarjeta de progreso es un panel con gráficas: tamaño fijo grande,
  // sin el auto-ajuste de texto de las demás tarjetas.
  const { w, h } =
    kind === 'progreso' ? { w: 360, h: 560 } : fitCardSize(base)
  editor.createShape({
    id: createShapeId(),
    type: 'universal-card',
    x: bounds.center.x - w / 2 + jitter,
    y: bounds.center.y - h / 2 + jitter,
    props: { ...base, w, h },
  })
}

export default function App() {
  const editorRef = useRef<Editor | null>(null)
  const autoSyncedRef = useRef(false)
  const [account, setAccount] = useState<AccountInfo | null>(null)
  const [authBusy, setAuthBusy] = useState(false)
  const [syncBusy, setSyncBusy] = useState(false)
  const [syncMsg, setSyncMsg] = useState<string | null>(null)
  const [debugError, setDebugError] = useState<string | null>(null)

  useEffect(() => {
    const onError = (e: ErrorEvent) => setDebugError(e.message || 'Error desconocido')
    const onRejection = (e: PromiseRejectionEvent) =>
      setDebugError(e.reason instanceof Error ? e.reason.message : String(e.reason))
    window.addEventListener('error', onError)
    window.addEventListener('unhandledrejection', onRejection)
    return () => {
      window.removeEventListener('error', onError)
      window.removeEventListener('unhandledrejection', onRejection)
    }
  }, [])

  /**
   * Importa las citas de Outlook (2 semanas, más la semana pasada con palomita) como tarjetas enlazadas.
   * En modo automático solo usa el token en silencio: si no se puede,
   * no hace nada (nunca redirige al abrir el tablero).
   */
  const runSync = useCallback(async (manual: boolean) => {
    const editor = editorRef.current
    if (!editor) return
    setSyncBusy(true)
    try {
      const token = manual
        ? await getToken(SCOPES.calendar)
        : await getTokenSilent(SCOPES.calendar)
      if (!token) return
      const events = await fetchUpcomingEvents(token)
      const created = syncCalendarEvents(editor, events)
      if (manual || created > 0) {
        setSyncMsg(
          created > 0
            ? `✓ ${created} cita${created === 1 ? '' : 's'} nueva${created === 1 ? '' : 's'} de Outlook`
            : '✓ Calendario al día',
        )
        window.setTimeout(() => setSyncMsg(null), 6000)
      }
    } catch (e) {
      if (manual) {
        setSyncMsg(
          'No se pudieron traer las citas: ' +
            (e instanceof Error ? e.message : String(e)),
        )
      }
    } finally {
      setSyncBusy(false)
    }
  }, [])

  useEffect(() => {
    if (isConfigured()) {
      getAccount()
        .then((acc) => {
          setAccount(acc)
          if (acc && editorRef.current && !autoSyncedRef.current) {
            autoSyncedRef.current = true
            void runSync(false)
          }
        })
        .catch(() => setAccount(null))
    }
  }, [runSync])

  const handleSignIn = async () => {
    setAuthBusy(true)
    try {
      // Redirige a Microsoft; si regresamos aquí es que algo falló.
      await signIn()
    } catch (e) {
      // Mostrar el error en vez de fallar en silencio (diagnóstico).
      window.alert(
        'No se pudo conectar: ' + (e instanceof Error ? e.message : String(e)),
      )
      setAuthBusy(false)
    }
  }

  const handleMount = useCallback((editor: Editor) => {
    editorRef.current = editor
    // Muse presente en el tablero, con sus movimientos de siempre.
    ensureMuseGif(editor)
    // Si la cuenta ya se resolvió antes de montar el editor, sincroniza ahora.
    if (!autoSyncedRef.current) {
      getAccount()
        .then((acc) => {
          if (acc && !autoSyncedRef.current) {
            autoSyncedRef.current = true
            void runSync(false)
          }
        })
        .catch(() => undefined)
    }
    // El tablero siempre en modo claro, sin importar el tema del sistema
    editor.user.updateUserPreferences({ colorScheme: 'light' })
    // Si el tablero está vacío, dejamos dos tarjetas de ejemplo para
    // mostrar que los objetos son datos vivos, no dibujos.
    if (editor.getCurrentPageShapes().length === 0) {
      const demo1: UniversalCardShapeProps = {
        w: 212,
        h: 120,
        kind: 'todo',
        title: 'Probar el tablero en la Chromebook',
        body: 'Doble toque en la tarjeta para interactuar con ella.',
        done: false,
      }
      editor.createShape({
        id: createShapeId(),
        type: 'universal-card',
        x: 120,
        y: 120,
        props: { ...demo1, ...fitCardSize(demo1) },
      })
      const demo2: UniversalCardShapeProps = {
        w: 212,
        h: 120,
        kind: 'nota',
        title: 'Las tarjetas son objetos',
        body: 'Su comportamiento lo definen sus datos: cita, goal, idea, correo…',
        done: false,
      }
      editor.createShape({
        id: createShapeId(),
        type: 'universal-card',
        x: 380,
        y: 120,
        props: { ...demo2, ...fitCardSize(demo2) },
      })
    }
  }, [])

  const handleSignOut = async () => {
    await signOut()
    setAccount(null)
  }

  const barButton: React.CSSProperties = {
    fontSize: 13,
    padding: '6px 10px',
    borderRadius: 8,
    border: '1px solid #ccc',
    background: '#fff',
    cursor: 'pointer',
  }

  return (
    <div style={{ position: 'fixed', inset: 0 }}>
      <Tldraw
        persistenceKey="tablero-de-jaime-v2"
        shapeUtils={shapeUtils}
        onMount={handleMount}
        licenseKey={TLDRAW_LICENSE_KEY ? TLDRAW_LICENSE_KEY : undefined}
        options={{ createTextOnCanvasDoubleClick: false }}
      />
      <div
        style={{
          position: 'fixed',
          top: 12,
          left: 12,
          zIndex: 300,
          display: 'flex',
          gap: 8,
          alignItems: 'center',
          background: 'rgba(255,255,255,0.92)',
          border: '1px solid #ddd',
          borderRadius: 12,
          padding: '8px 10px',
          fontFamily: 'sans-serif',
          boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
        }}
      >
        <button style={barButton} onClick={() => editorRef.current && createCard(editorRef.current, 'todo', 'Nuevo to-do')}>
          + To-do
        </button>
        <button style={barButton} onClick={() => editorRef.current && createCard(editorRef.current, 'cita', 'Nueva cita')}>
          + Cita
        </button>
        <button style={barButton} onClick={() => editorRef.current && createCard(editorRef.current, 'nota', 'Nueva nota')}>
          + Nota
        </button>
        <button style={barButton} onClick={() => editorRef.current && createCard(editorRef.current, 'idea', 'Nueva idea')}>
          + Idea
        </button>
        <button
          style={barButton}
          title="Tu progreso de maratón 2026 hacia la meta 3:45, con gráficas e intervalos"
          onClick={() => editorRef.current && createCard(editorRef.current, 'progreso', 'Mi progreso a 3:45')}
        >
          🏃 Progreso
        </button>
        <button
          style={barButton}
          title="Encuadrar todo el contenido del tablero"
          onClick={() => {
            const ed = editorRef.current
            if (!ed) return
            if (ed.getCurrentPageShapes().length > 0) ed.zoomToFit()
            else ed.resetZoom()
          }}
        >
          🎯 Centrar
        </button>
        <div style={{ width: 1, alignSelf: 'stretch', background: '#ddd' }} />
        {isConfigured() ? (
          account ? (
            <>
              <span style={{ fontSize: 12, color: '#333' }}>☁ {account.username}</span>
              <button
                style={barButton}
                onClick={() => void runSync(true)}
                disabled={syncBusy}
                title="Traer tus citas de Outlook (2 semanas) como tarjetas"
              >
                {syncBusy ? 'Trayendo…' : '⇩ Citas'}
              </button>
              {syncMsg ? (
                <span style={{ fontSize: 12, color: '#2e7d32', maxWidth: 260 }}>{syncMsg}</span>
              ) : null}
              <button style={barButton} onClick={handleSignOut}>
                Salir
              </button>
            </>
          ) : (
            <button style={barButton} onClick={handleSignIn} disabled={authBusy}>
              {authBusy ? 'Conectando…' : 'Conectar con Outlook'}
            </button>
          )
        ) : (
          <span style={{ fontSize: 12, color: '#999' }}>Outlook: falta el Client ID</span>
        )}
      </div>
      {debugError ? (
        <div
          style={{
            position: 'fixed',
            bottom: 12,
            left: 12,
            right: 12,
            zIndex: 400,
            background: '#c62828',
            color: '#fff',
            borderRadius: 8,
            padding: '8px 12px',
            fontFamily: 'monospace',
            fontSize: 12,
          }}
        >
          Error: {debugError}
        </div>
      ) : null}
    </div>
  )
}
