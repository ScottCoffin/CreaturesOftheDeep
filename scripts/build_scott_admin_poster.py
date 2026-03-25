from pathlib import Path

try:
    from PIL import Image, ImageColor, ImageDraw, ImageFont, ImageOps
except ModuleNotFoundError as exc:
    raise SystemExit(
        "This script requires Pillow. Install it in the active interpreter with "
        "`python -m pip install pillow` or `%pip install pillow`."
    ) from exc


REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PDF = REPO_ROOT / "output" / "pdf" / "scott_admin_office_poster.pdf"
OUTPUT_PNG = REPO_ROOT / "output" / "pdf" / "scott_admin_office_poster.png"
PORTRAIT_PATH = REPO_ROOT / "images" / "WETLAB_SCOTT.png"
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
            char_bbox = draw.textbbox((0, 0), char, font=font)
            cursor += (char_bbox[2] - char_bbox[0]) + tracking


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
    current_line = ""

    for word in words:
        candidate = word if not current_line else f"{current_line} {word}"
        candidate_width, _ = measure_text(draw, candidate, font)
        if candidate_width <= max_width:
            current_line = candidate
            continue

        if current_line:
            lines.append(current_line)
        current_line = word

    if current_line:
        lines.append(current_line)

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
) -> None:
    left, top, right, bottom = box
    box_width = right - left
    box_height = bottom - top

    image = Image.open(source_path).convert("RGBA")
    fitted = ImageOps.contain(image, (box_width, box_height), Image.Resampling.LANCZOS)
    paste_x = left + ((box_width - fitted.width) // 2)
    paste_y = top + ((box_height - fitted.height) // 2)
    canvas.paste(fitted, (paste_x, paste_y), fitted)


def paste_image_scaled(
    canvas: Image.Image,
    source_path: Path,
    xy: tuple[int, int],
    size: tuple[int, int],
    *,
    alpha_scale: float = 1.0,
) -> None:
    image = Image.open(source_path).convert("RGBA")
    resized = image.resize(size, Image.Resampling.LANCZOS)
    if alpha_scale < 1.0:
        alpha = resized.getchannel("A").point(lambda value: int(value * alpha_scale))
        resized.putalpha(alpha)
    canvas.paste(resized, xy, resized)


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
    panel_width = panel_right - panel_left
    panel_height = panel_bottom - panel_top
    center_x = PAGE_WIDTH / 2.0

    draw.rectangle((panel_left, panel_top, panel_right, panel_bottom), fill=WHITE, outline=DEEP_SEA, width=5)

    header_height = 440
    header_bottom = panel_top + header_height
    draw.rectangle((panel_left, panel_top, panel_right, header_bottom), fill=TITLE_BG)
    draw.line((panel_left, header_bottom, panel_right, header_bottom), fill=LAB_LINE, width=3)

    logo_box = (panel_left + 36, panel_top + 28, panel_left + 440, panel_top + 400)
    paste_image_contained(image, LOGO_PATH, logo_box)

    header_font = load_font(116, bold=True)
    subhead_font = load_font(68, bold=False)
    header_x = logo_box[2] + 34
    draw.text((header_x, panel_top + 54), "W.E.T. LAB", font=header_font, fill=AGENCY_BLUE)
    draw.text((header_x, panel_top + 196), "EXECUTIVE OFFICE", font=header_font, fill=AGENCY_BLUE)
    draw.text((header_x, panel_top + 340), "Station 744 Administrative Command", font=subhead_font, fill=AGENCY_BLUE)

    portrait_width = 1120
    portrait_height = 1680
    portrait_left = int(center_x - (portrait_width / 2.0))
    portrait_top = int((PAGE_HEIGHT - portrait_height) / 2.0)
    portrait_right = portrait_left + portrait_width
    portrait_bottom = portrait_top + portrait_height
    frame_pad = 44
    draw.rectangle(
        (portrait_left - frame_pad, portrait_top - frame_pad, portrait_right + frame_pad, portrait_bottom + frame_pad),
        fill=PANEL_BG,
        outline=LAB_LINE,
        width=3,
    )
    paste_image_contained(image, PORTRAIT_PATH, (portrait_left, portrait_top, portrait_right, portrait_bottom))

    name_font = load_font(112, bold=True)
    title_font = load_font(54, bold=True)
    name_top = portrait_bottom + 74
    draw_centered_text(draw, center_x, name_top, "DR. SCOTT COFFIN", name_font, DEEP_SEA, tracking=4)
    draw_centered_text(draw, center_x, name_top + 132, "W.E.T. LAB ADMIN", title_font, AGENCY_BLUE, tracking=6)

    divider_y = name_top + 250
    draw.line((panel_left + 275, divider_y, panel_right - 275, divider_y), fill=GOLD, width=4)

    image.save(OUTPUT_PNG, format="PNG", optimize=True)
    image.save(OUTPUT_PDF, format="PDF", resolution=DPI)
    return OUTPUT_PDF, OUTPUT_PNG


if __name__ == "__main__":
    pdf_path, png_path = build_poster()
    print(pdf_path)
    print(png_path)
