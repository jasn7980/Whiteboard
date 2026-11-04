"""Genera los íconos PNG del Whiteboard para instalarlo como app.

Fondo morado del logo tldraw (#863bff) con un rayo blanco al centro.
Sin transparencia: iOS pone su propio redondeo sobre el ícono.
El rayo queda dentro de la zona segura central, así el 512 sirve
también como maskable.
"""
from PIL import Image, ImageDraw

PURPLE = (134, 59, 255, 255)
WHITE = (255, 255, 255, 255)

# Rayo en coordenadas de un lienzo 512x512 (forma clásica en zigzag).
BOLT = [
    (318, 28), (178, 302), (258, 302), (198, 484),
    (362, 198), (276, 198),
]


def make_icon(size: int, path: str) -> None:
    img = Image.new("RGBA", (size, size), PURPLE)
    draw = ImageDraw.Draw(img)
    # Fondo con esquinas redondeadas (el sistema recorta de todos modos).
    r = int(size * 0.225)
    draw.rounded_rectangle([0, 0, size - 1, size - 1], radius=r, fill=PURPLE)
    scale = size / 512
    bolt = [(int(x * scale), int(y * scale)) for x, y in BOLT]
    draw.polygon(bolt, fill=WHITE)
    img.convert("RGB").save(path, "PNG")
    print(f"{path}: {size}x{size}")


if __name__ == "__main__":
    import os

    out = os.path.expanduser("~/workspace/whiteboard/public")
    make_icon(180, os.path.join(out, "icon-180.png"))
    make_icon(192, os.path.join(out, "icon-192.png"))
    make_icon(512, os.path.join(out, "icon-512.png"))
