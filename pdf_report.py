"""PDF-звіт для батьків: усі спостереження та фото дитини за місяць.

На Android PDF малюється вбудованим системним засобом (android.graphics.pdf.PdfDocument),
тому додаткові бібліотеки не потрібні, а українські літери відображаються системним шрифтом.
На комп'ютері (для перевірки) використовується reportlab, якщо він встановлений.
"""
import os
from datetime import datetime

from storage import (CRITERIA, IMAGE_EXT, VIDEO_EXT, format_day, format_day_short,
                     format_month, plural)

PAGE_W, PAGE_H = 595.0, 842.0  # A4 у пунктах
M = 46.0                       # поля
CONTENT_W = PAGE_W - 2 * M

GREEN = (46, 125, 96)
ORANGE = (230, 120, 60)
DARK = (33, 37, 41)
GREY = (110, 117, 125)
LIGHT = (236, 243, 239)
LINE = (210, 218, 214)
WHITE = (255, 255, 255)


# ----------------------------------------------------------------------
# Бекенди малювання. Координата y відраховується згори сторінки.
# ----------------------------------------------------------------------
class AndroidBackend:
    def __init__(self):
        from jnius import autoclass
        self.PdfDocument = autoclass("android.graphics.pdf.PdfDocument")
        self.PageInfoBuilder = autoclass("android.graphics.pdf.PdfDocument$PageInfo$Builder")
        self.Paint = autoclass("android.graphics.Paint")
        self.PaintStyle = autoclass("android.graphics.Paint$Style")
        self.Typeface = autoclass("android.graphics.Typeface")
        self.Color = autoclass("android.graphics.Color")
        self.BitmapFactory = autoclass("android.graphics.BitmapFactory")
        self.BitmapOptions = autoclass("android.graphics.BitmapFactory$Options")
        self.Bitmap = autoclass("android.graphics.Bitmap")
        self.Matrix = autoclass("android.graphics.Matrix")
        self.Rect = autoclass("android.graphics.Rect")
        self.RectF = autoclass("android.graphics.RectF")
        self.ExifInterface = autoclass("android.media.ExifInterface")
        self.FileOutputStream = autoclass("java.io.FileOutputStream")
        self.doc = self.PdfDocument()
        self.page = None
        self.canvas = None
        self.page_no = 0
        self.paint = self.Paint()
        self.paint.setAntiAlias(True)

    def new_page(self):
        if self.page is not None:
            self.doc.finishPage(self.page)
        self.page_no += 1
        info = self.PageInfoBuilder(int(PAGE_W), int(PAGE_H), self.page_no).create()
        self.page = self.doc.startPage(info)
        self.canvas = self.page.getCanvas()

    def _paint(self, size=11, bold=False, color=DARK, fill=True, stroke=1.0):
        p = self.paint
        p.setTextSize(float(size))
        p.setTypeface(self.Typeface.DEFAULT_BOLD if bold else self.Typeface.DEFAULT)
        p.setColor(self.Color.rgb(int(color[0]), int(color[1]), int(color[2])))
        p.setStyle(self.PaintStyle.FILL if fill else self.PaintStyle.STROKE)
        p.setStrokeWidth(float(stroke))
        return p

    def text_width(self, text, size, bold=False):
        return self._paint(size, bold).measureText(text)

    def text(self, x, y, text, size=11, bold=False, color=DARK):
        p = self._paint(size, bold, color)
        self.canvas.drawText(text, float(x), float(y + size * 0.82), p)

    def rect(self, x, y, w, h, color):
        self.canvas.drawRect(float(x), float(y), float(x + w), float(y + h), self._paint(color=color))

    def line(self, x1, y1, x2, y2, color=LINE, width=0.8):
        self.canvas.drawLine(float(x1), float(y1), float(x2), float(y2),
                             self._paint(color=color, fill=False, stroke=width))

    def _rotation(self, path):
        try:
            o = self.ExifInterface(path).getAttributeInt("Orientation", 1)
            return {6: 90, 3: 180, 8: 270}.get(o, 0)
        except Exception:
            return 0

    def image_size(self, path):
        opts = self.BitmapOptions()
        opts.inJustDecodeBounds = True
        self.BitmapFactory.decodeFile(path, opts)
        w, h = opts.outWidth, opts.outHeight
        if w <= 0 or h <= 0:
            return None
        return (h, w) if self._rotation(path) in (90, 270) else (w, h)

    def image(self, path, x, y, w, h):
        size = self.image_size(path)
        if not size:
            return
        # Зменшуємо фото під розмір у звіті (~180 dpi), щоб не забракло пам'яті
        target = max(w, h) * 2.5
        sample = 1
        while max(size) / (sample * 2) >= target:
            sample *= 2
        opts = self.BitmapOptions()
        opts.inSampleSize = sample
        bmp = self.BitmapFactory.decodeFile(path, opts)
        if bmp is None:
            return
        rot = self._rotation(path)
        if rot:
            m = self.Matrix()
            m.postRotate(float(rot))
            rotated = self.Bitmap.createBitmap(bmp, 0, 0, bmp.getWidth(), bmp.getHeight(), m, True)
            bmp.recycle()
            bmp = rotated
        src = self.Rect(0, 0, bmp.getWidth(), bmp.getHeight())
        dst = self.RectF(float(x), float(y), float(x + w), float(y + h))
        self.canvas.drawBitmap(bmp, src, dst, self._paint())
        bmp.recycle()

    def save(self, path):
        if self.page is not None:
            self.doc.finishPage(self.page)
            self.page = None
        fos = self.FileOutputStream(path)
        try:
            self.doc.writeTo(fos)
        finally:
            fos.close()
            self.doc.close()


class ReportLabBackend:
    """Для перевірки на комп'ютері."""
    FONT_DIRS = ["/usr/share/fonts/truetype/dejavu", "C:/Windows/Fonts", "/Library/Fonts"]

    def __init__(self, path):
        from reportlab.pdfgen import canvas
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont
        self.pdfmetrics = pdfmetrics
        for d in self.FONT_DIRS:
            reg, bold = os.path.join(d, "DejaVuSans.ttf"), os.path.join(d, "DejaVuSans-Bold.ttf")
            if os.path.exists(reg) and os.path.exists(bold):
                pdfmetrics.registerFont(TTFont("UA", reg))
                pdfmetrics.registerFont(TTFont("UA-Bold", bold))
                break
        else:
            raise RuntimeError("Не знайдено шрифт DejaVuSans для PDF")
        self.c = canvas.Canvas(path, pagesize=(PAGE_W, PAGE_H))
        self.started = False

    @staticmethod
    def _col(c):
        return c[0] / 255, c[1] / 255, c[2] / 255

    def new_page(self):
        if self.started:
            self.c.showPage()
        self.started = True

    def text_width(self, text, size, bold=False):
        return self.pdfmetrics.stringWidth(text, "UA-Bold" if bold else "UA", size)

    def text(self, x, y, text, size=11, bold=False, color=DARK):
        self.c.setFillColorRGB(*self._col(color))
        self.c.setFont("UA-Bold" if bold else "UA", size)
        self.c.drawString(x, PAGE_H - y - size * 0.82, text)

    def rect(self, x, y, w, h, color):
        self.c.setFillColorRGB(*self._col(color))
        self.c.rect(x, PAGE_H - y - h, w, h, stroke=0, fill=1)

    def line(self, x1, y1, x2, y2, color=LINE, width=0.8):
        self.c.setStrokeColorRGB(*self._col(color))
        self.c.setLineWidth(width)
        self.c.line(x1, PAGE_H - y1, x2, PAGE_H - y2)

    def _open(self, path):
        from PIL import Image, ImageOps
        return ImageOps.exif_transpose(Image.open(path))

    def image_size(self, path):
        try:
            return self._open(path).size
        except Exception:
            return None

    def image(self, path, x, y, w, h):
        from reportlab.lib.utils import ImageReader
        self.c.drawImage(ImageReader(self._open(path)), x, PAGE_H - y - h, w, h)

    def save(self, path):
        self.c.save()


# ----------------------------------------------------------------------
# Верстка
# ----------------------------------------------------------------------
class Layout:
    def __init__(self, backend, footer):
        self.be = backend
        self.footer = footer
        self.page = 0
        self.y = M
        self.new_page()

    def new_page(self):
        self.be.new_page()
        self.page += 1
        self.y = M
        self.be.line(M, PAGE_H - M + 4, PAGE_W - M, PAGE_H - M + 4, LINE, 0.6)
        self.be.text(M, PAGE_H - M + 10, f"{self.footer}  ·  сторінка {self.page}", 8.5, False, GREY)

    def ensure(self, h):
        if self.y + h > PAGE_H - M:
            self.new_page()

    def wrap(self, text, size, bold=False, width=CONTENT_W):
        lines = []
        for para in str(text).split("\n"):
            words = para.split()
            if not words:
                lines.append("")
                continue
            line = ""
            for word in words:
                test = f"{line} {word}".strip()
                if self.be.text_width(test, size, bold) <= width:
                    line = test
                    continue
                if line:
                    lines.append(line)
                # дуже довге слово ріжемо по літерах
                while self.be.text_width(word, size, bold) > width:
                    cut = len(word)
                    while cut > 1 and self.be.text_width(word[:cut], size, bold) > width:
                        cut -= 1
                    lines.append(word[:cut])
                    word = word[cut:]
                line = word
            lines.append(line)
        return lines

    def paragraph(self, text, size=11, bold=False, color=DARK, indent=0.0, after=4.0):
        lh = size * 1.38
        for line in self.wrap(text, size, bold, CONTENT_W - indent):
            self.ensure(lh)
            if line:
                self.be.text(M + indent, self.y, line, size, bold, color)
            self.y += lh
        self.y += after

    def heading(self, text, keep=80):
        self.ensure(42 + keep)  # щоб заголовок не залишився сам унизу сторінки
        self.y += 10
        self.be.text(M, self.y, text, 15, True, GREEN)
        self.y += 22
        self.be.line(M, self.y, M + CONTENT_W, self.y, GREEN, 1.2)
        self.y += 10


def build_report(backend, out_path, child, month, profile, records, media_dir):
    """records: список (day, record) за місяць для цієї дитини."""
    month_label = format_month(month)
    lay = Layout(backend, f"Щоденник спостережень · {child['name']} · {month_label}")
    be = backend

    # --- шапка ---
    be.rect(0, 0, PAGE_W, 118, GREEN)
    be.rect(0, 118, PAGE_W, 5, ORANGE)
    be.text(M, 34, "Звіт для батьків", 26, True, WHITE)
    be.text(M, 74, f"Досягнення та спостереження за {month_label}", 13, False, WHITE)
    lay.y = 145

    # --- картка дитини ---
    info = [("Дитина", child["name"])]
    if child.get("age"):
        info.append(("Вік", child["age"]))
    group = ", ".join(p for p in (profile.get("group_number", ""), profile.get("group_spec", "")) if p)
    if group:
        info.append(("Група", group))
    if profile.get("full_name"):
        role = f" ({profile['role'].lower()})" if profile.get("role") else ""
        info.append(("Педагог", profile["full_name"] + role))
    card_h = 18 + len(info) * 22
    be.rect(M, lay.y, CONTENT_W, card_h, LIGHT)
    y = lay.y + 12
    for label, value in info:
        be.text(M + 14, y, f"{label}:", 11.5, True, GREY)
        be.text(M + 100, y, value, 11.5, False, DARK)
        y += 22
    lay.y += card_h + 8

    # --- підсумок ---
    images, others = [], []
    for day, r in records:
        for name in r.get("media", []):
            path = os.path.join(media_dir, name)
            if not os.path.exists(path):
                continue
            if name.lower().endswith(IMAGE_EXT):
                images.append((day, path))
            else:
                others.append((day, name))

    lay.heading("Підсумок місяця")
    n = len(records)
    if n == 0:
        lay.paragraph("За цей місяць записів про дитину ще немає.", 11.5)
    else:
        lay.paragraph(f"За місяць педагог зробив {n} {plural(n, 'спостереження', 'спостереження', 'спостережень')}"
                      f", додано {len(images)} {plural(len(images), 'фото', 'фото', 'фото')}"
                      + (f" і {len(others)} {plural(len(others), 'інший медіафайл', 'інші медіафайли', 'інших медіафайлів')}."
                         if others else "."), 11.5)
        counts = {c: sum(1 for _, r in records if c in r.get("criteria", [])) for c in CRITERIA}
        top = max(counts.values()) or 1
        bar_x, bar_w = M + 150, CONTENT_W - 190
        for c in CRITERIA:
            lay.ensure(24)
            be.text(M, lay.y + 2, c, 11, False, DARK)
            be.rect(bar_x, lay.y + 3, bar_w, 11, LIGHT)
            if counts[c]:
                be.rect(bar_x, lay.y + 3, bar_w * counts[c] / top, 11, GREEN)
            be.text(bar_x + bar_w + 10, lay.y + 2, str(counts[c]), 11, True, DARK)
            lay.y += 22
        lay.y += 4

    # --- спостереження ---
    if records:
        lay.heading("Спостереження")
        for day, r in records:
            lay.ensure(56)
            be.text(M, lay.y, f"{format_day_short(day)}, {r['time']}", 11.5, True, DARK)
            lay.y += 18
            if r.get("criteria"):
                lay.paragraph(" · ".join(r["criteria"]), 10.5, True, GREEN, after=2)
            if r.get("note"):
                lay.paragraph(r["note"], 11, after=2)
            media_n = len(r.get("media", []))
            if media_n:
                lay.paragraph(f"Медіафайлів: {media_n}", 9.5, False, GREY, after=0)
            lay.y += 6
            lay.ensure(10)
            be.line(M, lay.y, M + CONTENT_W, lay.y, LINE, 0.6)
            lay.y += 10

    # --- фото ---
    if images:
        gap = 14
        cell_w = (CONTENT_W - gap) / 2
        max_h = 220
        lay.heading("Фотографії", keep=max_h + 30)
        for i in range(0, len(images), 2):
            row = []
            for day, path in images[i:i + 2]:
                try:
                    size = be.image_size(path)
                except Exception:
                    size = None
                if not size:
                    continue
                iw, ih = size
                scale = min(cell_w / iw, max_h / ih)
                row.append((day, path, iw * scale, ih * scale))
            if not row:
                continue
            row_h = max(h for _, _, _, h in row) + 22
            lay.ensure(row_h)
            for col, (day, path, w, h) in enumerate(row):
                x = M + col * (cell_w + gap) + (cell_w - w) / 2
                try:
                    be.image(path, x, lay.y, w, h)
                except Exception:
                    pass
                be.text(x, lay.y + h + 5, format_day_short(day), 9.5, False, GREY)
            lay.y += row_h + 8

    if others:
        lay.heading("Відео та аудіо")
        videos = sum(1 for _, name in others if name.lower().endswith(VIDEO_EXT))
        rest = len(others) - videos
        summary = f"Відеозаписів: {videos}" + (f", аудіо та інших файлів: {rest}" if rest else "")
        lay.paragraph(summary + ". Ці матеріали можна отримати у вихователя.", 11)
        for day, name in others:
            lay.paragraph(f"{format_day_short(day)} — {name}", 10, False, GREY, indent=10, after=0)

    lay.y += 16
    lay.paragraph(f"Звіт сформовано {format_day(datetime.now().strftime('%Y-%m-%d'))} "
                  "у застосунку «Щоденник спостережень».", 9.5, False, GREY)
    be.save(out_path)
    return out_path


def generate(out_path, child, month, profile, records, media_dir, android=False):
    backend = AndroidBackend() if android else ReportLabBackend(out_path)
    return build_report(backend, out_path, child, month, profile, records, media_dir)
