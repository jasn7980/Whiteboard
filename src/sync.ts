import { createShapeId, type Editor } from 'tldraw'
import type { UniversalCardShape, UniversalCardShapeProps } from './UniversalCard'
import { fitCardSize } from './fit'
import type { CalendarEventItem } from './graph'

/** "2026-10-25T10:30:00.0000000" → "2026-10-25T10:30" (formato datetime-local). */
function toLocalInput(dateTime: string): string {
  return dateTime.slice(0, 16)
}

/** Ahora mismo en hora de Toronto, en el mismo formato comparable. */
function torontoNowLocal(): string {
  const parts = new Intl.DateTimeFormat('en-CA', {
    timeZone: 'America/Toronto',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hour12: false,
  }).formatToParts(new Date())
  const get = (t: string) => parts.find((p) => p.type === t)?.value ?? ''
  const hour = get('hour') === '24' ? '00' : get('hour')
  return `${get('year')}-${get('month')}-${get('day')}T${hour}:${get('minute')}`
}

/**
 * Trae los eventos de Outlook al tablero como tarjetas de cita:
 * - Crea las que falten (enlazadas por el id del evento: sin duplicados).
 * - Actualiza título y horario de las ya enlazadas si el evento cambió.
 * Devuelve cuántas tarjetas NUEVAS se crearon.
 */
export function syncCalendarEvents(
  editor: Editor,
  events: CalendarEventItem[],
): number {
  const shapes = editor.getCurrentPageShapes()
  const citas = shapes.filter(
    (s): s is UniversalCardShape =>
      s.type === 'universal-card' && (s.props as UniversalCardShapeProps).kind === 'cita',
  )
  const byLinkId = new Map<string, UniversalCardShape>()
  for (const shape of citas) {
    const linkId = (shape.props as UniversalCardShapeProps).linkId
    if (linkId) byLinkId.set(linkId, shape)
  }

  // Columna de aterrizaje para las tarjetas nuevas: a la derecha de todo
  // el contenido actual, apiladas hacia abajo.
  let columnX = 120
  let startY = 120
  if (shapes.length > 0) {
    let maxX = 0
    let minY = Number.POSITIVE_INFINITY
    for (const shape of shapes) {
      const bounds = editor.getShapePageBounds(shape.id)
      if (!bounds) continue
      if (bounds.maxX > maxX) maxX = bounds.maxX
      if (bounds.y < minY) minY = bounds.y
    }
    if (maxX > 0) columnX = maxX + 60
    if (Number.isFinite(minY)) startY = minY
  }

  let created = 0
  let nextY = startY
  const nowLocal = torontoNowLocal()
  for (const event of events) {
    const title = event.subject?.trim() || 'Sin título'
    const startAt = toLocalInput(event.start.dateTime)
    const endAt = toLocalInput(event.end.dateTime)
    // Evento que ya terminó → la tarjeta queda marcada con palomita.
    const passed = endAt < nowLocal
    const existing = byLinkId.get(event.id)
    if (existing) {
      const props = existing.props as UniversalCardShapeProps
      if (
        props.title !== title ||
        props.startAt !== startAt ||
        props.endAt !== endAt ||
        props.done !== passed
      ) {
        editor.updateShape({
          id: existing.id,
          type: 'universal-card',
          props: { ...props, title, startAt, endAt, done: passed },
        })
      }
      continue
    }
    const baseProps: UniversalCardShapeProps = {
      w: 252,
      h: 160,
      kind: 'cita',
      title,
      body: event.location?.displayName ?? '',
      done: passed,
      linkId: event.id,
      startAt,
      endAt,
    }
    const { w, h } = fitCardSize(baseProps)
    editor.createShape({
      id: createShapeId(),
      type: 'universal-card',
      x: columnX,
      y: nextY,
      props: { ...baseProps, w, h },
    })
    nextY += h + 20
    created++
  }
  return created
}
