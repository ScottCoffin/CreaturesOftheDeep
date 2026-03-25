from pathlib import Path

try:
    from PIL import Image, ImageColor, ImageDraw, ImageFont, ImageOps
except ModuleNotFoundError as exc:
    raise SystemExit(
        "This script requires Pillow. Install it in the active interpreter with "
        "`python -m pip install pillow` or `%pip install pillow`."
    ) from exc


REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PDF = REPO_ROOT / "output" / "pdf" / "pelagic_acoustics_lab_poster.pdf"
OUTPUT_PNG = REPO_ROOT / "output" / "pdf" / "pelagic_acoustics_lab_poster.png"
LOGO_PATH = REPO_ROOT / "images" / "logo_small.png"

DPI = 300
PAGE_WIDTH = int(8.5 * DPI)
PAGE_HEIGHT = int(11 * DPI)

AGENCY_BLUE = ImageColor.getrgb("#005A8B")
DEEP_SEA = ImageColor.getrgb("#143B5D")
LAB_LINE = ImageColor.getrgb("#A9BBC8")
PANEL_BG = ImageColor.getrgb("#F6F8FA")
TITLE_BG = ImageColor.getrgb("#EAF2F7")
TEXT_GRAY = ImageColor.getrgb("#2F3B45")
GOLD = ImageColor.getrgb("#C9993A")
WHITE = (255, 255, 255)


def load_font(size: int, *, bold: bool = False) -> ImageFont.FreeTypeFont | ImageFont.ImageFont:
    windows_fonts = Path(r"C:\Windows\Fonts")
    candidates = [
        windows_fonts / ("arialbd.ttf" if bold else "arial.ttf"),
        windows_fonts / ("ARIALBD.TTF" if bold else "ARIAL.TTF"),
        windows_fonts / ("calibrib.ttf" if bold else "calibri.ttf"),
        windows_fonts / ("segoeuib.ttf" if bold else "segoeui.ttf"),
    ]
    for path in candidates:
        if path.exists():
            return ImageFont.truetype(str(path), size=size)
    return ImageFont.load_default()


def measure_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
    tracking: int = 0,
) -> tuple[int, int]:
    if not text:
        return 0, 0
    bbox = draw.textbbox((0, 0), text, font=font)
    width = (bbox[2] - bbox[0]) + max(0, len(text) - 1) * tracking
    height = bbox[3] - bbox[1]
    return width, height


def draw_tracked_text(
    draw: ImageDraw.ImageDraw,
    xy: tuple[float, float],
    text: str,
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
    fill: tuple[int, int, int],
    tracking: int = 0,
) -> None:
    x, y = xy
    cursor = x
    for char in text:
        draw.text((cursor, y), char, font=font, fill=fill)
        if char:
            bbox = draw.textbbox((0, 0), char, font=font)
            cursor += (bbox[2] - bbox[0]) + tracking


def draw_centered_text(
    draw: ImageDraw.ImageDraw,
    center_x: float,
    top_y: float,
    text: str,
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
    fill: tuple[int, int, int],
    tracking: int = 0,
) -> None:
    width, _ = measure_text(draw, text, font, tracking)
    draw_tracked_text(draw, (center_x - (width / 2.0), top_y), text, font, fill, tracking)


def wrap_text(
    draw: ImageDraw.ImageDraw,
    text: str,
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
    max_width: int,
) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current = ""
    for word in words:
        candidate = word if not current else f"{current} {word}"
        width, _ = measure_text(draw, candidate, font)
        if width <= max_width:
            current = candidate
            continue
        if current:
            lines.append(current)
        current = word
    if current:
        lines.append(current)
    return lines


def draw_centered_lines(
    draw: ImageDraw.ImageDraw,
    center_x: float,
    top_y: float,
    lines: list[str],
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont,
    fill: tuple[int, int, int],
    line_height: int,
) -> None:
    for index, line in enumerate(lines):
        draw_centered_text(draw, center_x, top_y + (index * line_height), line, font, fill)


def paste_image_contained(
    canvas: Image.Image,
    source_path: Path,
    box: tuple[int, int, int, int],
    *,
    alpha_scale: float = 1.0,
) -> None:
    left, top, right, bottom = box
    image = Image.open(source_path).convert("RGBA")
    fitted = ImageOps.contain(image, (right - left, bottom - top), Image.Resampling.LANCZOS)
    if alpha_scale < 1.0:
        alpha = fitted.getchannel("A").point(lambda value: int(value * alpha_scale))
        fitted.putalpha(alpha)
    paste_x = left + ((right - left - fitted.width) // 2)
    paste_y = top + ((bottom - top - fitted.height) // 2)
    canvas.paste(fitted, (paste_x, paste_y), fitted)


def build_poster() -> tuple[Path, Path]:
    OUTPUT_PDF.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_PNG.parent.mkdir(parents=True, exist_ok=True)

    image = Image.new("RGB", (PAGE_WIDTH, PAGE_HEIGHT), PANEL_BG)
    draw = ImageDraw.Draw(image)

    panel_margin = 105
    panel_left = panel_margin
    panel_top = panel_margin
    panel_right = PAGE_WIDTH - panel_margin
    panel_bottom = PAGE_HEIGHT - panel_margin
    center_x = PAGE_WIDTH / 2.0

    draw.rectangle((panel_left, panel_top, panel_right, panel_bottom), fill=WHITE, outline=DEEP_SEA, width=5)

    header_height = 280
    header_bottom = panel_top + header_height
    draw.rectangle((panel_left, panel_top, panel_right, header_bottom), fill=TITLE_BG)
    draw.line((panel_left, header_bottom, panel_right, header_bottom), fill=LAB_LINE, width=3)

    logo_box = (panel_left + 40, panel_top + 28, panel_left + 280, panel_top + 250)
    paste_image_contained(image, LOGO_PATH, logo_box)

    header_font = load_font(56, bold=True)
    subhead_font = load_font(32, bold=False)
    header_x = logo_box[2] + 28
    draw.text((header_x, panel_top + 46), "W.E.T. LAB", font=header_font, fill=AGENCY_BLUE)
    draw.text((header_x, panel_top + 118), "SONIC OPERATIONS DIVISION", font=header_font, fill=AGENCY_BLUE)
    draw.text((header_x, panel_top + 214), "Station 744 Signal Processing and Resonance Unit", font=subhead_font, fill=AGENCY_BLUE)

    watermark_box = (
        int(center_x - 300),
        header_bottom + 80,
        int(center_x + 300),
        header_bottom + 680,
    )
    paste_image_contained(image, LOGO_PATH, watermark_box, alpha_scale=0.08)

    title_font = load_font(112, bold=True)
    body_copy = (
        "Experimental grooves, percussion cascades, and unstable bass events "
        "must remain inside the room until properly field-tested."
    )
    body_font = load_font(56, bold=False)
    body_box_left = panel_left + 120
    body_box_right = panel_right - 120
    accent_width = 28
    content_left = body_box_left + accent_width + 44
    content_right = body_box_right - 44
    body_lines = wrap_text(draw, body_copy, body_font, content_right - content_left)
    body_line_height = 74
    body_text_height = len(body_lines) * body_line_height
    body_box_height = body_text_height + 120
    body_cluster_height = 268 + 82 + body_box_height
    title_top = int((PAGE_HEIGHT / 2.0) - (body_cluster_height / 2.0))

    draw_centered_text(draw, center_x, title_top, "PELAGIC ACOUSTICS", title_font, DEEP_SEA, tracking=4)
    draw_centered_text(draw, center_x, title_top + 126, "LAB", title_font, DEEP_SEA, tracking=6)

    divider_y = title_top + 268
    draw.line((panel_left + 210, divider_y, panel_right - 210, divider_y), fill=GOLD, width=4)

    body_box_top = divider_y + 82
    body_box_bottom = body_box_top + body_box_height
    draw.rectangle((body_box_left, body_box_top, body_box_right, body_box_bottom), fill=PANEL_BG, outline=LAB_LINE, width=3)
    draw.rectangle((body_box_left, body_box_top, body_box_left + accent_width, body_box_bottom), fill=TITLE_BG)

    body_text_top = body_box_top + ((body_box_height - body_text_height) // 2) - 10
    draw_centered_lines(
        draw,
        (content_left + content_right) / 2.0,
        body_text_top,
        body_lines,
        body_font,
        TEXT_GRAY,
        body_line_height,
    )

    footer_line_y = panel_bottom - 210
    draw.line((panel_left + 95, footer_line_y, panel_right - 95, footer_line_y), fill=LAB_LINE, width=3)
    footer_font = load_font(36, bold=True)
    draw_centered_text(
        draw,
        center_x,
        footer_line_y + 34,
        "Admin approval required for volume escalation.",
        footer_font,
        AGENCY_BLUE,
    )

    image.save(OUTPUT_PNG, format="PNG", optimize=True)
    image.save(OUTPUT_PDF, format="PDF", resolution=DPI)
    return OUTPUT_PDF, OUTPUT_PNG


if __name__ == "__main__":
    pdf_path, png_path = build_poster()
    print(pdf_path)
    print(png_path)
