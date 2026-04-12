import struct
from pathlib import Path

from PyQt5.QtCore import QByteArray, QBuffer, QIODevice, QLineF, QRectF, Qt
from PyQt5.QtGui import QColor, QIcon, QPainter, QPainterPath, QPen, QPixmap

try:
    from PyQt5.QtSvg import QSvgRenderer
except ImportError:
    QSvgRenderer = None

from .resources import app_data_dir, resource_path


ICON_SIZES = (16, 20, 24, 32, 40, 48, 64, 128, 256)
ICON_BACKGROUND = "#0F6CBD"
ICON_BORDER = "#084C84"
ICON_GLYPH = "#FFFFFF"
TRAY_ICON_FILE_NAME = "tray_keyboard_badge.ico"

_ICON_CACHE = None
_SVG_CACHE = None


def app_icon():
    global _ICON_CACHE

    if _ICON_CACHE is None:
        icon = QIcon()
        for size in ICON_SIZES:
            icon.addPixmap(render_keyboard_badge(size))
        _ICON_CACHE = icon

    return _ICON_CACHE


def ensure_tray_icon_file():
    target = app_data_dir().joinpath(TRAY_ICON_FILE_NAME)

    try:
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_bytes(build_ico_bytes())
        return str(target)
    except OSError:
        return resource_path("img", "ms_keyboard_tray.ico")


def build_ico_bytes():
    payloads = []
    for size in ICON_SIZES:
        payloads.append((size, pixmap_to_png_bytes(render_keyboard_badge(size))))

    directory = struct.pack("<HHH", 0, 1, len(payloads))
    offset = 6 + (16 * len(payloads))
    entries = []
    images = bytearray()

    for size, payload in payloads:
        entries.append(
            struct.pack(
                "<BBBBHHII",
                0 if size >= 256 else size,
                0 if size >= 256 else size,
                0,
                0,
                1,
                32,
                len(payload),
                offset,
            )
        )
        images.extend(payload)
        offset += len(payload)

    return directory + b"".join(entries) + bytes(images)


def pixmap_to_png_bytes(pixmap):
    array = QByteArray()
    buffer = QBuffer(array)
    buffer.open(QIODevice.WriteOnly)
    pixmap.save(buffer, "PNG")
    buffer.close()
    return bytes(array)


def render_keyboard_badge(size):
    pixmap = QPixmap(size, size)
    pixmap.fill(Qt.transparent)

    painter = QPainter(pixmap)
    painter.setRenderHint(QPainter.Antialiasing, True)
    painter.setRenderHint(QPainter.SmoothPixmapTransform, True)

    outer_margin = max(1.0, size * 0.08)
    badge_rect = QRectF(
        outer_margin,
        outer_margin,
        size - (outer_margin * 2),
        size - (outer_margin * 2),
    )
    radius = max(3.0, size * 0.22)

    badge = QPainterPath()
    badge.addRoundedRect(badge_rect, radius, radius)
    painter.fillPath(badge, QColor(ICON_BACKGROUND))

    border_pen = QPen(QColor(ICON_BORDER))
    border_pen.setWidthF(max(1.0, size * 0.05))
    painter.setPen(border_pen)
    painter.drawPath(badge)

    glyph_margin = max(2.0, size * 0.18)
    glyph_rect = QRectF(
        glyph_margin,
        glyph_margin,
        size - (glyph_margin * 2),
        size - (glyph_margin * 2),
    )
    render_keyboard_glyph(painter, glyph_rect)

    painter.end()
    return pixmap


def render_keyboard_glyph(painter, target_rect):
    svg_data = keyboard_svg_markup()
    if svg_data and QSvgRenderer is not None:
        renderer = QSvgRenderer(QByteArray(svg_data.encode("utf-8")))
        if renderer.isValid():
            renderer.render(painter, target_rect)
            return

    draw_fallback_keyboard_glyph(painter, target_rect)


def keyboard_svg_markup():
    global _SVG_CACHE

    if _SVG_CACHE is not None:
        return _SVG_CACHE

    svg_path = Path(resource_path("img", "ic_fluent_keyboard_24_regular.svg"))
    try:
        svg_markup = svg_path.read_text(encoding="utf-8")
    except OSError:
        _SVG_CACHE = ""
        return _SVG_CACHE

    _SVG_CACHE = svg_markup.replace('fill="#212121"', f'fill="{ICON_GLYPH}"')
    return _SVG_CACHE


def draw_fallback_keyboard_glyph(painter, target_rect):
    painter.save()

    keyboard_pen = QPen(QColor(ICON_GLYPH))
    keyboard_pen.setWidthF(max(1.2, target_rect.width() * 0.08))
    painter.setPen(keyboard_pen)
    painter.setBrush(Qt.NoBrush)

    shell = QPainterPath()
    shell.addRoundedRect(
        target_rect.adjusted(0.0, target_rect.height() * 0.04, 0.0, -target_rect.height() * 0.02),
        target_rect.width() * 0.08,
        target_rect.width() * 0.08,
    )
    painter.drawPath(shell)

    key_pen = QPen(QColor(ICON_GLYPH))
    key_pen.setWidthF(max(1.0, target_rect.width() * 0.09))
    painter.setPen(key_pen)

    left = target_rect.left() + (target_rect.width() * 0.18)
    right = target_rect.right() - (target_rect.width() * 0.18)
    top = target_rect.top() + (target_rect.height() * 0.26)
    mid = target_rect.top() + (target_rect.height() * 0.49)
    bottom = target_rect.bottom() - (target_rect.height() * 0.24)

    painter.drawLine(QLineF(left, top, right, top))
    painter.drawLine(QLineF(left, mid, right, mid))
    painter.drawLine(
        QLineF(
            left + (target_rect.width() * 0.12),
            bottom,
            right - (target_rect.width() * 0.12),
            bottom,
        )
    )

    painter.restore()
