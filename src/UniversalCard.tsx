import {
  HTMLContainer,
  Rectangle2d,
  ShapeUtil,
  T,
  useEditor,
  useIsEditing,
} from 'tldraw'
import type { RecordProps, TLBaseShape } from 'tldraw'

// ---------------------------------------------------------------------------
// La "tarjeta universal": UN solo tipo de objeto cuyo comportamiento lo
// definen sus datos (kind + campos). Hoy sabe ser un to-do; mañana puede ser
// una cita, un goal, una idea o un correo sin tocar el lienzo.
// ---------------------------------------------------------------------------

export type UniversalCardShapeProps = {
  w: number
  h: number
  kind: string // 'todo' | 'nota' | 'cita' | ... lo que definamos
  title: string
  body: string
  done: boolean
}

export type UniversalCardShape = TLBaseShape<'universal-card', UniversalCardShapeProps>

declare module 'tldraw' {
  interface TLGlobalShapePropsMap {
    'universal-card': UniversalCardShapeProps
  }
}

function UniversalCardComponent({ shape }: { shape: UniversalCardShape }) {
  const editor = useEditor()
  const isEditing = useIsEditing(shape.id)
  const { kind, title, body, done } = shape.props

  const toggleDone = () => {
    editor.updateShape<UniversalCardShape>({
      id: shape.id,
      type: 'universal-card',
      props: { done: !done },
    })
  }

  return (
    <HTMLContainer
      style={{
        width: '100%',
        height: '100%',
        pointerEvents: isEditing ? 'all' : 'none',
        backgroundColor: kind === 'todo' ? '#fff8e1' : '#e8f4fd',
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
          fontSize: 10,
          fontWeight: 700,
          letterSpacing: 1,
          textTransform: 'uppercase',
          color: '#666',
        }}
      >
        {kind}
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
      {isEditing ? (
        <div style={{ fontSize: 11, color: '#888' }}>(modo edición: la tarjeta es interactiva)</div>
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
  }

  getDefaultProps(): UniversalCardShape['props'] {
    return {
      w: 240,
      h: 150,
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
