import { useEffect, useRef, useState } from 'react'
import { Editor, Tldraw, createShapeId } from 'tldraw'
import type { AccountInfo } from '@azure/msal-browser'
import 'tldraw/tldraw.css'
import { UniversalCardUtil } from './UniversalCard'
import type { UniversalCardShapeProps } from './UniversalCard'
import { getAccount, isConfigured, signIn, signOut } from './auth'

const shapeUtils = [UniversalCardUtil]

function createCard(editor: Editor, kind: string, title: string) {
  const bounds = editor.getViewportPageBounds()
  const jitter = Math.random() * 60 - 30
  const base: UniversalCardShapeProps = {
    w: 260,
    h: 200,
    kind,
    title,
    body: '',
    done: false,
  }
  editor.createShape({
    id: createShapeId(),
    type: 'universal-card',
    x: bounds.center.x - 130 + jitter,
    y: bounds.center.y - 100 + jitter,
    props: base,
  })
}

export default function App() {
  const editorRef = useRef<Editor | null>(null)
  const [account, setAccount] = useState<AccountInfo | null>(null)
  const [authBusy, setAuthBusy] = useState(false)

  useEffect(() => {
    if (isConfigured()) {
      getAccount()
        .then(setAccount)
        .catch(() => setAccount(null))
    }
  }, [])

  const handleSignIn = async () => {
    setAuthBusy(true)
    try {
      setAccount(await signIn())
    } catch (e) {
      // Mostrar el error en vez de fallar en silencio (diagnóstico).
      window.alert(
        'No se pudo conectar: ' + (e instanceof Error ? e.message : String(e)),
      )
    } finally {
      setAuthBusy(false)
    }
  }

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
        persistenceKey="tablero-de-jaime"
        shapeUtils={shapeUtils}
        onMount={(editor) => {
          editorRef.current = editor
          // El tablero siempre en modo claro, sin importar el tema del sistema
          editor.user.updateUserPreferences({ colorScheme: 'light' })
          // Si el tablero está vacío, dejamos dos tarjetas de ejemplo para
          // mostrar que los objetos son datos vivos, no dibujos.
          if (editor.getCurrentPageShapes().length === 0) {
            editor.createShape({
              id: createShapeId(),
              type: 'universal-card',
              x: 120,
              y: 120,
              props: {
                w: 260,
                h: 200,
                kind: 'todo',
                title: 'Probar el tablero en la Chromebook',
                body: 'Doble toque en la tarjeta para interactuar con ella.',
                done: false,
              },
            })
            editor.createShape({
              id: createShapeId(),
              type: 'universal-card',
              x: 440,
              y: 120,
              props: {
                w: 260,
                h: 200,
                kind: 'nota',
                title: 'Las tarjetas son objetos',
                body: 'Su comportamiento lo definen sus datos: cita, goal, idea, correo…',
                done: false,
              },
            })
          }
        }}
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
    </div>
  )
}
