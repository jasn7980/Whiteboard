import { Tldraw, createShapeId } from 'tldraw'
import 'tldraw/tldraw.css'
import { UniversalCardUtil } from './UniversalCard'

const shapeUtils = [UniversalCardUtil]

export default function App() {
  return (
    <div style={{ position: 'fixed', inset: 0 }}>
      <Tldraw
        persistenceKey="tablero-de-jaime"
        shapeUtils={shapeUtils}
        onMount={(editor) => {
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
                h: 170,
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
                h: 170,
                kind: 'nota',
                title: 'Las tarjetas son objetos',
                body: 'Su comportamiento lo definen sus datos: cita, goal, idea, correo…',
                done: false,
              },
            })
          }
        }}
      />
    </div>
  )
}
