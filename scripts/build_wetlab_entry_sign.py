from pathlib import Path

from reportlab.lib.colors import HexColor, white
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfgen import canvas


REPO_ROOT = Path(__file__).resolve().parents[1]
OUTPUT_PATH = REPO_ROOT / "output" / "pdf" / "wetlab_entry_sign.pdf"
LOGO_PATH = REPO_ROOT / "images" / "logo_small.png"

AGENCY_BLUE = HexColor("#005A8B")
DEEP_SEA = HexColor("#143B5D")
LAB_LINE = HexColor("#A9BBC8")
PANEL_BG = HexColor("#F6F8FA")
TITLE_BG = HexColor("#EAF2F7")


def centered_text(
    pdf: canvas.Canvas,
    text: str,
    center_x: float,
    baseline_y: float,
    font_name: str,
    font_size: float,
    color,
    char_space: float = 0.0,
) -> None:
    text_width = pdfmetrics.stringWidth(text, font_name, font_size)
    if text:
        text_width += char_space * (len(text) - 1)

    text_object = pdf.beginText()
    text_object.setTextOrigin(center_x - (text_width / 2.0), baseline_y)
    text_object.setFont(font_name, font_size)
    text_object.setFillColor(color)
    text_object.setCharSpace(char_space)
    text_object.textLine(text)
    pdf.drawText(text_object)


def wrap_text_to_width(
    text: str,
    font_name: str,
    font_size: float,
    max_width: float,
) -> list[str]:
    words = text.split()
    lines: list[str] = []
    current_line = ""

    for word in words:
        candidate = word if not current_line else f"{current_line} {word}"
        if pdfmetrics.stringWidth(candidate, font_name, font_size) <= max_width:
            current_line = candidate
            continue

        if current_line:
            lines.append(current_line)
        current_line = word

    if current_line:
        lines.append(current_line)

    return lines


def build_sign() -> Path:
    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)

    page_width, page_height = letter
    pdf = canvas.Canvas(str(OUTPUT_PATH), pagesize=letter, pageCompression=1)
    pdf.setTitle("WETLAB Entry Sign")
    pdf.setAuthor("Codex")
    pdf.setSubject("Scientific laboratory style entry placard for W.E.T. LAB")

    panel_margin = 0.4 * inch
    panel_x = panel_margin
    panel_y = panel_margin
    panel_width = page_width - (2 * panel_margin)
    panel_height = page_height - (2 * panel_margin)
    panel_top = panel_y + panel_height

    pdf.setFillColor(white)
    pdf.rect(0, 0, page_width, page_height, fill=1, stroke=0)

    pdf.setFillColor(PANEL_BG)
    pdf.rect(panel_x, panel_y, panel_width, panel_height, fill=1, stroke=0)
    pdf.setStrokeColor(DEEP_SEA)
    pdf.setLineWidth(1.4)
    pdf.rect(panel_x, panel_y, panel_width, panel_height, fill=0, stroke=1)

    header_height = 0.75 * inch
    header_y = panel_top - header_height
    pdf.setFillColor(TITLE_BG)
    pdf.rect(panel_x, header_y, panel_width, header_height, fill=1, stroke=0)
    pdf.setStrokeColor(LAB_LINE)
    pdf.setLineWidth(0.8)
    pdf.line(panel_x, header_y, panel_x + panel_width, header_y)

    pdf.setFont("Helvetica-Bold", 13.2)
    pdf.setFillColor(AGENCY_BLUE)
    pdf.drawString(panel_x + 18, panel_top - 34, "W.E.T. LAB FIELD OPERATIONS UNIT")

    badge_width = 1.52 * inch
    badge_height = 0.34 * inch
    badge_x = panel_x + panel_width - badge_width - 18
    badge_y = panel_top - badge_height - 21
    pdf.setFillColor(white)
    pdf.setStrokeColor(LAB_LINE)
    pdf.setLineWidth(0.8)
    pdf.rect(badge_x, badge_y, badge_width, badge_height, fill=1, stroke=1)
    centered_text(
        pdf,
        "STATION 744",
        badge_x + (badge_width / 2.0),
        badge_y + 7,
        "Helvetica-Bold",
        10,
        DEEP_SEA,
        0.6,
    )

    logo_reader = ImageReader(str(LOGO_PATH))
    logo_width = 3.15 * inch * 1.5
    logo_height = logo_width * (481.0 / 520.0)
    logo_x = (page_width - logo_width) / 2.0
    logo_y = header_y - logo_height - 34
    pdf.drawImage(
        logo_reader,
        logo_x,
        logo_y,
        width=logo_width,
        height=logo_height,
        mask="auto",
        preserveAspectRatio=True,
    )

    centered_text(
        pdf,
        "WETLAB",
        page_width / 2.0,
        logo_y - 46,
        "Helvetica-Bold",
        42,
        DEEP_SEA,
        2.0,
    )
    centered_text(
        pdf,
        "WEST SACRAMENTO EXPERIMENTAL TROPHIC LABORATORY",
        page_width / 2.0,
        logo_y - 74,
        "Helvetica-Bold",
        11.5,
        AGENCY_BLUE,
        0.85,
    )

    divider_y = logo_y - 84
    pdf.setStrokeColor(LAB_LINE)
    pdf.setLineWidth(1.0)
    pdf.line(panel_x + 70, divider_y, panel_x + panel_width - 70, divider_y)

    centered_text(
        pdf,
        "MARINE RESEARCH LAB ENTRANCE",
        page_width / 2.0,
        divider_y - 32,
        "Helvetica-Bold",
        18,
        DEEP_SEA,
        1.2,
    )

    notice_width = panel_width - 72
    notice_height = 1.05 * inch
    notice_x = panel_x + 36
    notice_y = panel_y + 1.08 * inch
    notice_accent_width = 11
    pdf.setFillColor(white)
    pdf.setStrokeColor(LAB_LINE)
    pdf.setLineWidth(0.85)
    pdf.rect(notice_x, notice_y, notice_width, notice_height, fill=1, stroke=1)
    pdf.setFillColor(TITLE_BG)
    pdf.rect(notice_x, notice_y, notice_accent_width, notice_height, fill=1, stroke=0)

    notice_content_x = notice_x + notice_accent_width
    notice_content_width = notice_width - notice_accent_width
    notice_center_x = notice_content_x + (notice_content_width / 2.0)

    centered_text(
        pdf,
        "AUTHORIZED ENTRY",
        notice_center_x,
        notice_y + 45,
        "Helvetica-Bold",
        16,
        DEEP_SEA,
        1.0,
    )
    notice_font_name = "Helvetica"
    notice_font_size = 9.4
    notice_lines = [
        "Invited personnel, field scientists, and",
        "approved marine organisms may proceed.",
    ]
    line_height = 10.5
    total_text_height = len(notice_lines) * line_height
    first_line_y = notice_y + 15 + total_text_height - line_height
    pdf.setFillColor(AGENCY_BLUE)
    for index, line in enumerate(notice_lines):
        centered_text(
            pdf,
            line,
            notice_center_x,
            first_line_y - (index * line_height),
            notice_font_name,
            notice_font_size,
            AGENCY_BLUE,
        )

    footer_y = panel_y + 26
    pdf.setStrokeColor(LAB_LINE)
    pdf.setLineWidth(0.8)
    pdf.line(panel_x + 24, footer_y + 16, panel_x + panel_width - 24, footer_y + 16)
    centered_text(
        pdf,
        "INTERNAL USE ONLY",
        page_width / 2.0,
        footer_y,
        "Helvetica-Bold",
        11.4,
        AGENCY_BLUE,
        0.8,
    )

    pdf.showPage()
    pdf.save()
    return OUTPUT_PATH


if __name__ == "__main__":
    result = build_sign()
    print(result)
