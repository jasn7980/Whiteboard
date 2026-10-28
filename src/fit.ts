import type { UniversalCardShapeProps } from './UniversalCard'

/**
 * Tamaño mínimo que necesita una tarjeta para mostrar su texto sin
 * recortes, pensado para pantallas chicas: ancho fijo y compacto por
 * tipo, y alto calculado según cuántas líneas ocupa el texto.
 * Es una estimación conservadora (mejor que sobre un poco de alto a
 * que se corte texto, porque en modo lectura el contenido no se
 * desplaza).
 */
export function fitCardSize(
  props: Pick<UniversalCardShapeProps, 'kind' | 'title' | 'body' | 'startAt' | 'endAt'>,
): { w: number; h: number } {
  const w = props.kind === 'cita' ? 252 : 212
  const contentW = w - 24 // padding horizontal de la tarjeta

  const lines = (text: string, pxPerChar: number) => {
    const perLine = Math.max(8, Math.floor(contentW / pxPerChar))
    return text
      .split('\n')
      .reduce((acc, paragraph) => acc + Math.max(1, Math.ceil(paragraph.length / perLine)), 0)
  }

  let h = 10 // padding superior
  h += 16 // fila de encabezado (tipo + estado)
  h += 6
  h += lines(props.title || '', 8.4) * 20 // título 15px
  if (props.body && props.body.trim().length > 0) {
    h += 4 + lines(props.body, 7.0) * 18 // cuerpo 13px
  }
  if (props.kind === 'cita') {
    h += 4 + 16 // fila de horario
  }
  h += 10 // padding inferior
  h += 6 // holgura de seguridad

  return { w, h: Math.min(430, Math.max(150, Math.round(h))) }
}
