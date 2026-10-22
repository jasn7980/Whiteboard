# Whiteboard — Tablero personal de Jaime

Un pizarrón infinito pensado para la Lenovo Yoga Chromebook (touchscreen + stylus):
dibujo libre con tinta sensible a la presión, más **tarjetas-objeto vivas** cuyo
comportamiento lo definen sus datos.

## La tarjeta universal

Un solo tipo de objeto (`universal-card`) con propiedades `kind`, `title`, `body`
y `done`. Hoy sabe ser:

- **todo** — con casilla que funciona de verdad (doble toque para interactuar)
- **nota** — tarjeta informativa

Mañana puede ser una cita, un goal, una idea, un recordatorio o un correo, sin
tocar el lienzo: basta agregarle el comportamiento a la tarjeta.

## Tecnología

- [tldraw](https://tldraw.dev) 5.5 (SDK de lienzo infinito) — licencia hobby para uso personal
- React + TypeScript + Vite
- Persistencia local del tablero en el navegador
- Publicado con GitHub Pages vía Actions

## Desarrollo

```bash
npm install
npm run dev     # servidor local
npm run build   # build de producción en dist/
```
