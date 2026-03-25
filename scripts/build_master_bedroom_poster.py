from pathlib import Path

try:
    from PIL import Image, ImageColor, ImageDraw, ImageFont, ImageOps
except ModuleNotFoundError as exc:
    raise SystemExit(
        "This script requires Pillow. Install it in the active interpreter with "
        "`python -m pip install pillow` or `%pip install pillow`."
    ) from exc


REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PDF = REPO_ROOT / "output" / "pdf" / "restorative_conversation_suite_poster.pdf"
OUTPUT_PNG = REPO_ROOT / "output" / "pdf" / "restorative_conversation_suite_poster.png"
LOGO_PATH = REPO_ROOT / "images" / "logo_small.png"
HERO_IMAGE_PATH = REPO_ROOT / "images" / "vibrant-marine-life-stockcake.webp"

DPI = 300
PAGE_WIDTH = int(8.5 * DPI)
PAGE_HEIGHT = int(11 * DPI)

AGENCY_BLUE = ImageColor.getrgb("#005A8B")
DEEP_SEA = ImageColor.getrgb("#143B5D")
LAB_LINE = ImageColor.getrgb("#A9BBC8")
PANEL_BG = ImageColor.getrgb("#F6F8FA")
TITLE_BG = ImageColor.getrgb("#EAF2F7")
TEXT_GRAY = ImageColor.getrgb("#2F3B45")
CALM_TEAL = ImageColor.getrgb("#6BA8B0")
SEAFOAM = ImageColor.getrgb("#DCECEF")
REEF_GOLD = ImageColor.getrgb("#D0A24A")
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

    header_height = 210
    header_bottom = panel_top + header_height
    draw.rectangle((panel_left, panel_top, panel_right, header_bottom), fill=TITLE_BG)
    draw.line((panel_left, header_bottom, panel_right, header_bottom), fill=LAB_LINE, width=3)

    logo_box = (panel_left + 40, panel_top + 18, panel_left + 220, panel_top + 190)
    paste_image_contained(image, LOGO_PATH, logo_box)

    header_font = load_font(48, bold=True)
    subhead_font = load_font(28, bold=False)
    header_x = logo_box[2] + 28
    draw.text((header_x, panel_top + 34), "W.E.T. LAB", font=header_font, fill=AGENCY_BLUE)
    draw.text((header_x, panel_top + 92), "HOSPITALITY DIVISION", font=header_font, fill=AGENCY_BLUE)
    draw.text((header_x, panel_top + 152), "Station 744 Restorative Social Habitat", font=subhead_font, fill=AGENCY_BLUE)

    title_font = load_font(94, bold=True)
    subtitle_font = load_font(42, bold=False)
    body_font = load_font(46, bold=False)
    body_copy = (
        "This chamber is reserved for quiet conversation, nervous system "
        "recalibration, horizontal observation, and temporary withdrawal "
        "from high-intensity marine activity."
    )

    body_box_left = panel_left + 120
    body_box_right = panel_right - 120
    accent_width = 28
    content_left = body_box_left + accent_width + 44
    content_right = body_box_right - 44
    body_lines = wrap_text(draw, body_copy, body_font, content_right - content_left)
    body_line_height = 62
    body_text_height = len(body_lines) * body_line_height
    body_box_height = body_text_height + 120

    hero_height = 940
    hero_title_height = 330
    body_section_gap = 72
    footer_line_y = panel_bottom - 210
    body_region_top = header_bottom + 50
    body_region_bottom = footer_line_y - 72
    centered_stack_height = hero_height + body_section_gap + body_box_height
    centered_stack_top = int(body_region_top + ((body_region_bottom - body_region_top - centered_stack_height) / 2.0))

    hero_top = centered_stack_top
    hero_bottom = hero_top + hero_height
    hero_left = panel_left + 150
    hero_right = panel_right - 150
    hero_frame_pad = 14
    draw.rectangle(
        (hero_left - hero_frame_pad, hero_top - hero_frame_pad, hero_right + hero_frame_pad, hero_bottom + hero_frame_pad),
        fill=SEAFOAM,
        outline=LAB_LINE,
        width=3,
    )
    paste_image_contained(image, HERO_IMAGE_PATH, (hero_left, hero_top, hero_right, hero_bottom))
    hero_overlay = Image.new("RGBA", (hero_right - hero_left, hero_height), (*DEEP_SEA, 92))
    image.paste(hero_overlay, (hero_left, hero_top), hero_overlay)

    hero_title_top = int(hero_top + ((hero_height - hero_title_height) / 2.0))
    draw_centered_text(draw, center_x, hero_title_top, "RESTORATIVE", title_font, WHITE, tracking=5)
    draw_centered_text(draw, center_x, hero_title_top + 110, "CONVERSATION SUITE", title_font, WHITE, tracking=5)
    draw_centered_text(draw, center_x, hero_title_top + 238, "Quiet conversation, recalibration, and reef-adjacent repose", subtitle_font, SEAFOAM)

    divider_y = hero_bottom + 36
    draw.line((panel_left + 170, divider_y, panel_right - 170, divider_y), fill=REEF_GOLD, width=4)

    body_box_top = divider_y + body_section_gap
    body_box_bottom = body_box_top + body_box_height
    draw.rectangle((body_box_left, body_box_top, body_box_right, body_box_bottom), fill=PANEL_BG)
    draw.rectangle((body_box_left, body_box_top, body_box_left + accent_width, body_box_bottom), fill=TITLE_BG)
    draw.rectangle((body_box_left, body_box_top, body_box_right, body_box_bottom), outline=LAB_LINE, width=3)

    body_text_top = body_box_top + ((body_box_height - body_text_height) // 2) - 8
    draw_centered_lines(
        draw,
        (content_left + content_right) / 2.0,
        body_text_top,
        body_lines,
        body_font,
        TEXT_GRAY,
        body_line_height,
    )

    draw.line((panel_left + 95, footer_line_y, panel_right - 95, footer_line_y), fill=LAB_LINE, width=3)
    footer_font = load_font(36, bold=True)
    draw_centered_text(
        draw,
        center_x,
        footer_line_y + 34,
        "Restore baseline before returning to the field.",
        footer_font,
        CALM_TEAL,
    )

    image.save(OUTPUT_PNG, format="PNG", optimize=True)
    image.save(OUTPUT_PDF, format="PDF", resolution=DPI)
    return OUTPUT_PDF, OUTPUT_PNG


if __name__ == "__main__":
    pdf_path, png_path = build_poster()
    print(pdf_path)
    print(png_path)
