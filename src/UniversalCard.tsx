import { useState } from 'react'
import {
  HTMLContainer,
  Rectangle2d,
  ShapeUtil,
  T,
  useEditor,
  useIsEditing,
} from 'tldraw'
import type { RecordProps, TLBaseShape } from 'tldraw'
import { getToken, isConfigured } from './auth'
import { SCOPES } from './config'
import { createCalendarEvent, createOneNotePage, createTodoTask, setTodoTaskDone } from './graph'

// ---------------------------------------------------------------------------
// La "tarjeta universal": UN solo tipo de objeto cuyo comportamiento lo
// definen sus datos (kind + campos). Hoy sabe ser to-do, nota y cita, y los
// to-dos y citas se enlazan con objetos reales de Outlook (To Do / Calendar)
// guardando su identificador en linkId.
// ---------------------------------------------------------------------------

export type UniversalCardShapeProps = {
  w: number
  h: number
  kind: string // 'todo' | 'nota' | 'cita' | ... lo que definamos
  title: string
  body: string
  done: boolean
  linkId?: string // id del objeto real en Outlook (tarea o evento)
  linkListId?: string // solo to-do: id de la lista de To Do
  startAt?: string // solo cita: inicio local ISO (YYYY-MM-DDTHH:mm)
  endAt?: string // solo cita: fin local ISO
}

export type UniversalCardShape = TLBaseShape<'universal-card', UniversalCardShapeProps>

declare module 'tldraw' {
  interface TLGlobalShapePropsMap {
    'universal-card': UniversalCardShapeProps
  }
}

function toInputValue(d: Date): string {
  const pad = (n: number) => String(n).padStart(2, '0')
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}T${pad(d.getHours())}:${pad(d.getMinutes())}`
}

function defaultStart(): string {
  const d = new Date(Date.now() + 60 * 60 * 1000)
  d.setMinutes(d.getMinutes() < 30 ? 30 : 60, 0, 0)
  return toInputValue(d)
}

function defaultEnd(startIso: string): string {
  const d = new Date(startIso)
  d.setHours(d.getHours() + 1)
  return toInputValue(d)
}

function formatWhen(iso: string): string {
  const d = new Date(iso)
  return d.toLocaleString('es', {
    weekday: 'short',
    day: 'numeric',
    month: 'short',
    hour: 'numeric',
    minute: '2-digit',
  })
}

function UniversalCardComponent({ shape }: { shape: UniversalCardShape }) {
  const editor = useEditor()
  const isEditing = useIsEditing(shape.id)
  const [busy, setBusy] = useState(false)
  const [error, setError] = useState<string | null>(null)
  const { kind, title, body, done, linkId, linkListId } = shape.props
  const startAt = shape.props.startAt || defaultStart()
  const endAt = shape.props.endAt || defaultEnd(startAt)
  const linked = Boolean(linkId)

  const patch = (props: Partial<UniversalCardShapeProps>) => {
    editor.updateShape<UniversalCardShape>({
      id: shape.id,
      type: 'universal-card',
      props,
    })
  }

  const run = async (fn: () => Promise<void>) => {
    setBusy(true)
    setError(null)
    try {
      await fn()
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e))
    } finally {
      setBusy(false)
    }
  }

  const toggleDone = () => {
    const next = !done
    patch({ done: next })
    if (kind === 'todo' && linked && linkListId && isConfigured()) {
      void run(async () => {
        const token = await getToken(SCOPES.tasks)
        await setTodoTaskDone(token, linkListId, linkId!, next)
      })
    }
  }

  const createRemote = () => {
    void run(async () => {
      if (kind === 'todo') {
        const token = await getToken(SCOPES.tasks)
        const { listId, taskId } = await createTodoTask(token, title)
        patch({ linkId: taskId, linkListId: listId })
      } else if (kind === 'cita') {
        const token = await getToken(SCOPES.calendar)
        const eventId = await createCalendarEvent(token, title, startAt, endAt)
        patch({ linkId: eventId, startAt, endAt })
      } else if (kind === 'idea') {
        const token = await getToken(SCOPES.notes)
        const pageId = await createOneNotePage(token, title, body)
        patch({ linkId: pageId })
      }
    })
  }

  const linkable = kind === 'todo' || kind === 'cita' || kind === 'idea'
  const serviceName = kind === 'idea' ? 'OneNote' : 'Outlook'
  const createLabel =
    kind === 'todo'
      ? 'Crear tarea en Microsoft To Do'
      : kind === 'cita'
        ? 'Crear evento en Outlook'
        : 'Guardar idea en OneNote'
  const bg =
    kind === 'todo'
      ? '#fff8e1'
      : kind === 'cita'
        ? '#e8fde8'
        : kind === 'idea'
          ? '#f3e8fd'
          : '#e8f4fd'

  return (
    <HTMLContainer
      style={{
        width: '100%',
        height: '100%',
        pointerEvents: isEditing ? 'all' : 'none',
        backgroundColor: bg,
        border: '2px solid #333',
        borderRadius: 12,
        padding: 12,
        boxSizing: 'border-box',
        fontFamily: 'sans-serif',
        overflow: 'hidden',
        display: 'flex',
        flexDirection: 'column',
        gap: 6,
      }}
    >
      <div
        style={{
          display: 'flex',
          justifyContent: 'space-between',
          alignItems: 'center',
          fontSize: 10,
          fontWeight: 700,
          letterSpacing: 1,
          textTransform: 'uppercase',
          color: '#666',
        }}
      >
        <span>{kind}</span>
        {linked ? <span title="Guardado en el servicio enlazado">☁ {serviceName}</span> : null}
      </div>
      <div
        style={{
          fontSize: 16,
          fontWeight: 600,
          textDecoration: done ? 'line-through' : 'none',
          color: done ? '#888' : '#111',
        }}
      >
        {title}
      </div>
      {body ? <div style={{ fontSize: 13, color: '#444' }}>{body}</div> : null}

      {kind === 'cita' && !isEditing ? (
        <div style={{ fontSize: 13, color: '#333' }}>🗓 {formatWhen(startAt)}</div>
      ) : null}

      {isEditing && kind === 'cita' ? (
        <div style={{ display: 'flex', flexDirection: 'column', gap: 4, fontSize: 12 }}>
          <label>
            Inicio{' '}
            <input
              type="datetime-local"
              value={startAt}
              onChange={(e) => patch({ startAt: e.target.value })}
            />
          </label>
          <label>
            Fin{' '}
            <input
              type="datetime-local"
              value={endAt}
              onChange={(e) => patch({ endAt: e.target.value })}
            />
          </label>
        </div>
      ) : null}

      {kind === 'todo' ? (
        <label
          style={{
            marginTop: 'auto',
            fontSize: 13,
            display: 'flex',
            alignItems: 'center',
            gap: 6,
            cursor: 'pointer',
          }}
        >
          <input type="checkbox" checked={done} onChange={toggleDone} />
          {done ? 'Hecho' : 'Pendiente'}
        </label>
      ) : null}

      {isEditing && linkable ? (
        linked ? (
          <div style={{ fontSize: 12, color: '#2e7d32' }}>✓ Guardado en {serviceName}</div>
        ) : (
          <button
            onClick={createRemote}
            disabled={busy || !isConfigured()}
            style={{ fontSize: 13, padding: '6px 8px', cursor: 'pointer', marginTop: 4 }}
          >
            {busy ? 'Guardando…' : createLabel}
          </button>
        )
      ) : null}
      {isEditing && !isConfigured() && linkable ? (
        <div style={{ fontSize: 11, color: '#999' }}>
          (La conexión se activa al configurar el Client ID)
        </div>
      ) : null}
      {error ? (
        <div style={{ fontSize: 11, color: '#c62828' }}>Error: {error.slice(0, 140)}</div>
      ) : null}
    </HTMLContainer>
  )
}

export class UniversalCardUtil extends ShapeUtil<UniversalCardShape> {
  static override type = 'universal-card' as const

  static override props: RecordProps<UniversalCardShape> = {
    w: T.number,
    h: T.number,
    kind: T.string,
    title: T.string,
    body: T.string,
    done: T.boolean,
    linkId: T.optional(T.string),
    linkListId: T.optional(T.string),
    startAt: T.optional(T.string),
    endAt: T.optional(T.string),
  }

  getDefaultProps(): UniversalCardShape['props'] {
    return {
      w: 240,
      h: 190,
      kind: 'todo',
      title: 'Nueva tarjeta',
      body: '',
      done: false,
    }
  }

  getGeometry(shape: UniversalCardShape) {
    return new Rectangle2d({
      width: shape.props.w,
      height: shape.props.h,
      isFilled: true,
    })
  }

  component(shape: UniversalCardShape) {
    return <UniversalCardComponent shape={shape} />
  }

  getIndicatorPath(shape: UniversalCardShape) {
    const path = new Path2D()
    path.roundRect(0, 0, shape.props.w, shape.props.h, 12)
    return path
  }
}
