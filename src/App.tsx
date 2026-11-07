import { useCallback, useEffect, useRef, useState } from 'react'
import { Editor, Tldraw, createShapeId } from 'tldraw'
import type { AccountInfo } from '@azure/msal-browser'
import 'tldraw/tldraw.css'
import { UniversalCardUtil } from './UniversalCard'
import type { UniversalCardShapeProps } from './UniversalCard'
import { getAccount, getToken, getTokenSilent, isConfigured, signIn, signOut } from './auth'
import { fetchUpcomingEvents } from './graph'
import { fitCardSize } from './fit'
import { ensureMuseGif, restoreMuseGif } from './museGif'
import {
  addPhotoToBoard,
  compressPhoto,
  ensurePhotosFolder,
  fetchPhotoBlobUrl,
  resolvePhotoAssets,
  uploadPhoto,
} from './photos'
import { syncCalendarEvents } from './sync'
import { attachCloudSave, initCloudBoard, uploadBoardFile, downloadBoardFile, fetchNewerCloudBoard, isCloudDirty, setCloudMarker } from './onedrive'
import type { CloudBoardFile } from './onedrive'
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
  const cloudInitRef = useRef<Promise<boolean> | null>(null)
  const pollRef = useRef<number | null>(null)
  const photoInputRef = useRef<HTMLInputElement | null>(null)
  const [account, setAccount] = useState<AccountInfo | null>(null)
  const [authBusy, setAuthBusy] = useState(false)
  const [syncBusy, setSyncBusy] = useState(false)
  const [syncMsg, setSyncMsg] = useState<string | null>(null)
  const [cloudReady, setCloudReady] = useState(false)
  const [cloudInitDone, setCloudInitDone] = useState(false)
  const [cloudUpdate, setCloudUpdate] = useState<CloudBoardFile | null>(null)
  const [cloudMsg, setCloudMsg] = useState<string | null>(null)
  const [cloudBusy, setCloudBusy] = useState(false)
  const [photoBusy, setPhotoBusy] = useState(false)
  const [debugError, setDebugError] = useState<string | null>(null)

  useEffect(() => {
    return () => {
      if (pollRef.current !== null) window.clearInterval(pollRef.current)
    }
  }, [])

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
    // Las citas importadas no deben caer antes de que la nube haya
    // decidido qué tablero manda: si no, la carga la borraría.
    if (cloudInitRef.current) await cloudInitRef.current
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

  /**
   * Revisa cada 3 min si OneDrive tiene una versión más nueva del tablero.
   * Si este dispositivo no tiene cambios sin subir, la carga sola; si los
   * tiene, avisa y deja que el usuario decida (cargar la nube descartaría
   * sus cambios locales).
   */
  const checkCloudUpdates = useCallback(async () => {
    const editor = editorRef.current
    if (!editor) return
    try {
      const token = await getTokenSilent(SCOPES.files)
      if (!token) return
      const file = await fetchNewerCloudBoard(token)
      if (!file) {
        // Sin versión nueva: igual resolver las fotos pendientes.
        await resolvePhotoAssets(editor, () => Promise.resolve(token))
        return
      }
      if (isCloudDirty()) {
        setCloudUpdate(file)
        return
      }
      editor.loadSnapshot(file.snapshot)
      setCloudMarker(file.savedAt)
      setCloudUpdate(null)
      setCloudMsg('✓ Tablero actualizado desde OneDrive')
      window.setTimeout(() => setCloudMsg(null), 6000)
      // Las fotos viajan como referencias: resolverlas tras cada carga.
      await resolvePhotoAssets(editor, () => Promise.resolve(token))
    } catch {
      // Sin red o sin permiso: se reintenta en el próximo sondeo.
    }
  }, [])

  /** Carga la versión de la nube aunque haya cambios locales sin subir. */
  const loadCloudUpdate = useCallback(() => {
    const editor = editorRef.current
    if (!editor || !cloudUpdate) return
    if (
      !window.confirm(
        'Tienes cambios sin subir en este dispositivo. Si cargas la versión de OneDrive, esos cambios se pierden. ¿Cargar de todos modos?',
      )
    )
      return
    editor.loadSnapshot(cloudUpdate.snapshot)
    setCloudMarker(cloudUpdate.savedAt)
    setCloudUpdate(null)
    setCloudMsg('✓ Tablero actualizado desde OneDrive')
    window.setTimeout(() => setCloudMsg(null), 6000)
    void resolvePhotoAssets(editor, () => getTokenSilent(SCOPES.files))
  }, [cloudUpdate])

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

  const handleMount = useCallback(
    (editor: Editor) => {
      editorRef.current = editor
      // El tablero siempre en modo claro, sin importar el tema del sistema
      editor.user.updateUserPreferences({ colorScheme: 'light' })
      // Primero la nube decide qué tablero manda (o deja el local si no
      // hay sesión o permiso): solo después se siembra el GIF, las
      // tarjetas de ejemplo y la importación de citas, para que la
      // carga de la nube no las borre ni el GIF dispare su baja.
      cloudInitRef.current = initCloudBoard(editor, () =>
        getTokenSilent(SCOPES.files),
      )
      void cloudInitRef.current.then((ok) => {
        if (ok) setCloudReady(true)
        setCloudInitDone(true)
        attachCloudSave(editor, () => getTokenSilent(SCOPES.files), () =>
          setCloudReady(true),
        )
        // Sondeo de la nube cada 3 min: trae la versión de OneDrive si
        // otro dispositivo guardó después (sin recargar la página).
        if (pollRef.current === null) {
          pollRef.current = window.setInterval(
            () => void checkCloudUpdates(),
            180000,
          )
        }
        // Las fotos viajan como referencias: resolverlas al abrir.
        void resolvePhotoAssets(editor, () => getTokenSilent(SCOPES.files))
        // Muse presente en el tablero, con sus movimientos de siempre.
        ensureMuseGif(editor)
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
      })
    },
    [runSync, checkCloudUpdates],
  )

  /** Guarda el tablero en OneDrive ahora mismo (pide el permiso si falta). */
  const handleCloudSave = async () => {
    const editor = editorRef.current
    if (!editor) return
    setCloudBusy(true)
    try {
      const token = await getToken(SCOPES.files)
      await uploadBoardFile(token, editor)
      setCloudReady(true)
      setCloudMsg('✓ Guardado en OneDrive')
      window.setTimeout(() => setCloudMsg(null), 6000)
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      // Redirigir a Microsoft para autorizar no es un error: la página
      // regresa ya conectada y la nube se activa sola.
      if (!msg.includes('Redirigiendo')) {
        setCloudMsg('No se pudo guardar en OneDrive: ' + msg)
      }
    } finally {
      setCloudBusy(false)
    }
  }

  /**
   * Baja el tablero guardado en OneDrive y lo carga, forzado por el
   * usuario: si hay cambios sin subir en este dispositivo, se pierden.
   * Si no hay sesión o permiso, pide autorizarlo (igual que ⇧ Nube).
   */
  const handleCloudDownload = async () => {
    const editor = editorRef.current
    if (!editor) return
    setCloudBusy(true)
    try {
      const token = await getToken(SCOPES.files)
      const cloud = await downloadBoardFile(token)
      if (!cloud) {
        setCloudMsg('No hay tablero guardado en OneDrive')
        window.setTimeout(() => setCloudMsg(null), 6000)
        return
      }
      if (isCloudDirty()) {
        if (
          !window.confirm(
            'Tienes cambios sin subir en este dispositivo. Si descargas la versión de OneDrive, esos cambios se pierden. ¿Descargar de todos modos?',
          )
        )
          return
      }
      editor.loadSnapshot(cloud.snapshot)
      setCloudMarker(cloud.savedAt)
      setCloudUpdate(null)
      setCloudReady(true)
      setCloudMsg('✓ Tablero descargado desde OneDrive')
      window.setTimeout(() => setCloudMsg(null), 6000)
      await resolvePhotoAssets(editor, () => Promise.resolve(token))
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      // Redirigir a Microsoft para autorizar no es un error: la página
      // regresa ya conectada y la nube se activa sola.
      if (!msg.includes('Redirigiendo')) {
        setCloudMsg('No se pudo descargar de OneDrive: ' + msg)
      }
    } finally {
      setCloudBusy(false)
    }
  }

  /**
   * Sube las fotos elegidas a OneDrive (comprimidas) y las pone en el
   * tablero como imágenes, en el centro de la vista. En el teléfono el
   * selector ofrece tomar la foto o elegirla de la galería.
   */
  const handlePhotoFiles = async (files: FileList | null) => {
    const editor = editorRef.current
    if (!editor || !files || files.length === 0) return
    setPhotoBusy(true)
    try {
      const token = await getToken(SCOPES.files)
      await ensurePhotosFolder(token)
      const bounds = editor.getViewportPageBounds()
      let added = 0
      let firstError: string | null = null
      for (const file of Array.from(files)) {
        // Sin filtro previo por tipo o extensión: se intenta decodificar
        // cada archivo y solo el decodificador decide si es una foto.
        try {
          const { blob, width, height } = await compressPhoto(file)
          const path = await uploadPhoto(token, blob, file.name)
          const url = await fetchPhotoBlobUrl(token, path)
          addPhotoToBoard(
            editor,
            path,
            url,
            width,
            height,
            bounds.center.x,
            bounds.center.y,
            added,
          )
          added++
        } catch (e) {
          // Una foto mala no aborta el resto del lote.
          if (!firstError) firstError = e instanceof Error ? e.message : String(e)
        }
      }
      setCloudReady(true)
      setCloudMsg(
        added === 0
          ? firstError
            ? 'No se pudo subir la foto: ' + firstError
            : 'No se pudo agregar ninguna foto: no se reconoció el formato de imagen'
          : added === 1
            ? '✓ Foto guardada en OneDrive'
            : `✓ ${added} fotos guardadas en OneDrive`,
      )
      // Solo el éxito se borra solo: el error se queda hasta tocarlo.
      if (added > 0) window.setTimeout(() => setCloudMsg(null), 6000)
    } catch (e) {
      const msg = e instanceof Error ? e.message : String(e)
      if (!msg.includes('Redirigiendo')) {
        // El error se queda visible hasta tocarlo: el mensaje de 6 s
        // se perdía sin que se alcanzara a leer.
        setCloudMsg('No se pudo subir la foto: ' + msg)
      }
    } finally {
      setPhotoBusy(false)
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
    flexShrink: 0,
    whiteSpace: 'nowrap',
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
          right: 12,
          zIndex: 300,
          display: 'flex',
          gap: 8,
          alignItems: 'center',
          flexWrap: 'nowrap',
          overflowX: 'auto',
          WebkitOverflowScrolling: 'touch',
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
        <button
          style={barButton}
          title="Abrir el mapa de búsqueda de casas en otra pestaña"
          onClick={() => window.open('https://jasn7980.github.io/Mapa-casas/', '_blank', 'noopener')}
        >
          🗺 Mapa de casas
        </button>
        <button
          style={barButton}
          title="Volver a poner a Muse en el tablero"
          onClick={() => {
            const ed = editorRef.current
            if (ed) restoreMuseGif(ed)
          }}
        >
          🐈‍⬛ Muse
        </button>
        <button
          style={barButton}
          title="Agregar una foto al tablero (se guarda en OneDrive)"
          disabled={photoBusy}
          onClick={() => {
            // Renueva el permiso ANTES de elegir las fotos: si hace
            // falta ir a Microsoft, la redirección pasa ahora y no se
            // pierde la selección (al volver, toca 📷 Foto otra vez).
            void getToken(SCOPES.files).catch(() => {})
            photoInputRef.current?.click()
          }}
        >
          {photoBusy ? 'Subiendo…' : '📷 Foto'}
        </button>
        <input
          ref={photoInputRef}
          type="file"
          accept="image/*"
          multiple
          style={{ display: 'none' }}
          onChange={(e) => {
            void handlePhotoFiles(e.target.files)
            e.target.value = ''
          }}
        />
        <div style={{ width: 1, flexShrink: 0, alignSelf: 'stretch', background: '#ddd' }} />
        {isConfigured() ? (
          account ? (
            <>
              <span style={{ fontSize: 12, color: '#333', flexShrink: 0, whiteSpace: 'nowrap' }}>☁ {account.username}</span>
              <button
                style={barButton}
                onClick={() => void runSync(true)}
                disabled={syncBusy}
                title="Traer tus citas de Outlook (2 semanas) como tarjetas"
              >
                {syncBusy ? 'Trayendo…' : '⇩ Citas'}
              </button>
              {syncMsg ? (
                <span style={{ fontSize: 12, color: '#2e7d32', maxWidth: 260, flexShrink: 0 }}>{syncMsg}</span>
              ) : null}
              <button
                style={barButton}
                onClick={() => void handleCloudSave()}
                disabled={cloudBusy}
                title="Guardar este tablero en OneDrive ahora (si falta el permiso, te pedirá autorizarlo)"
              >
                {cloudBusy ? 'Guardando…' : '⇧ Nube'}
              </button>
              <button
                style={barButton}
                onClick={() => void handleCloudDownload()}
                disabled={cloudBusy}
                title="Bajar el tablero guardado en OneDrive (si tienes cambios sin subir, se pierden)"
              >
                {cloudBusy ? 'Bajando…' : '⇩ Nube'}
              </button>
              <span style={{ fontSize: 12, color: cloudReady ? '#2e7d32' : '#999', flexShrink: 0, whiteSpace: 'nowrap' }}>
                {cloudReady ? '☁ OneDrive' : '○ este dispositivo'}
              </span>
              {cloudMsg ? (
                <span
                  style={{ fontSize: 12, color: '#2e7d32', maxWidth: 260, flexShrink: 0, cursor: 'pointer' }}
                  title="Toca para ocultar este mensaje"
                  onClick={() => setCloudMsg(null)}
                >
                  {cloudMsg}
                </span>
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
        {import.meta.env.VITE_APP_VERSION ? (
          <span style={{ fontSize: 10, color: '#aaa', flexShrink: 0 }} title="Versión del tablero">
            v{import.meta.env.VITE_APP_VERSION}
          </span>
        ) : null}
      </div>
      {cloudInitDone && !cloudReady ? (
        <div
          style={{
            position: 'fixed',
            top: 68,
            left: 12,
            right: 12,
            zIndex: 300,
            display: 'flex',
            gap: 8,
            alignItems: 'center',
            background: '#fff3cd',
            border: '1px solid #e6c200',
            borderRadius: 12,
            padding: '8px 10px',
            fontFamily: 'sans-serif',
            fontSize: 13,
            boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
          }}
        >
          <span>⚠ Sin conexión a OneDrive: los cambios solo se guardan en este dispositivo.</span>
          {isConfigured() ? (
            <button style={barButton} onClick={handleSignIn} disabled={authBusy}>
              {authBusy ? 'Conectando…' : 'Conectar'}
            </button>
          ) : null}
        </div>
      ) : null}
      {cloudUpdate ? (
        <div
          style={{
            position: 'fixed',
            top: cloudInitDone && !cloudReady ? 124 : 68,
            left: 12,
            right: 12,
            zIndex: 300,
            display: 'flex',
            gap: 8,
            alignItems: 'center',
            background: '#e3f2fd',
            border: '1px solid #90caf9',
            borderRadius: 12,
            padding: '8px 10px',
            fontFamily: 'sans-serif',
            fontSize: 13,
            boxShadow: '0 2px 8px rgba(0,0,0,0.12)',
          }}
        >
          <span>Hay una versión más nueva en OneDrive y tienes cambios sin subir aquí.</span>
          <button style={barButton} onClick={loadCloudUpdate}>
            Cargar la de la nube
          </button>
          <button style={barButton} onClick={() => setCloudUpdate(null)}>
            Descartar
          </button>
        </div>
      ) : null}
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
