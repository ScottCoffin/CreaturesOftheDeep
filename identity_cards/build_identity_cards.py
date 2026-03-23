from __future__ import annotations

import re
import subprocess
import textwrap
from dataclasses import dataclass
from pathlib import Path
from zipfile import ZipFile

from lxml import etree
from PIL import Image, ImageOps


ROOT = Path(__file__).resolve().parent
PPTX = ROOT / "identity_cards.pptx"
PDF = ROOT / "identity_cards.pdf"
RENDER_DIR = ROOT / "rendered_cards"
ASSET_DIR = ROOT / "card_assets"
TEX_OUT = ROOT / "identity_cards_sheet.tex"

NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}

CLASS_ABILITIES = {
    "Data Scientist": 'May declare "model overfit" before a mini-game; a successful result doubles that station payout.',
    "Geneticist": "May engineer one hybrid organism; the hybrid counts as both species during play.",
    "Oceanographer": "Receives three clues to hidden $100 discoveries and may search for them during the game phase.",
    "Principal Investigator": "Maintains the team's notebook and records wins, discoveries, and collaborations; valid entries require an opposing team signature.",
    "Dolphin": "Acts as the team's dolphin during Dolphin Trainer and may respond only to approved cue phrases.",
    "Whale": "Grants the team a $200 prestige advantage for each whale assigned to the roster.",
    "Squid": "At the end of the game, may siphon $200 from any rival team.",
}

SCIENTIFIC_NAMES = [
    "Aethalodelphis obliquidens",
    "Balaenoptera acutorostrata",
    "Balaenoptera musculus",
    "Delphinapterus leucas",
    "Grampus griseus",
    "Inia geoffrensis",
    "Joubiniteuthis portieri",
    "Lagenorhynchus obscurus",
    "Megaptera novaeangliae",
    "Mondon monoceros",
    "Orcaella brevirostris",
    "Physeter macrocephalus",
    "Sepiadarium kochi",
    "Stenella frontalis",
    "Vampyroteuthis infernalis",
    "Watasenia scintillans",
]

DESCRIPTION_OVERRIDES = {
    9: "Two supermodels with extraordinary cheekbones and limited analytical range who still approach every problem with complete conviction.",
    11: "A mathematician who realized Babbage's machine could do more than arithmetic and drafted one of the earliest computer algorithms.",
    12: "A nurse and statistician who transformed hospital sanitation and used data graphics to expose preventable military deaths.",
    14: "A pioneering mathematician whose geodesy work for the U.S. Navy helped establish the mathematical foundation for GPS.",
    33: "A warm-water Atlantic dolphin whose spots accumulate with age, shifting from smooth gray juveniles to heavily speckled adults.",
    45: "A small Indo-West Pacific bobtail squid that buries itself in sand by day, shifts from orange to pale white, and releases heavy defensive slime.",
}


@dataclass
class SlideRecord:
    index: int
    card_class: str
    name: str
    description: str
    pic_x: int
    pic_y: int
    pic_cx: int
    pic_cy: int


def normalize_text(value: str) -> str:
    value = value.replace("Ph.D.", "Ph<DOT>D<DOT>")
    replacements = {
        "\u2018": "'",
        "\u2019": "'",
        "\u201c": '"',
        "\u201d": '"',
        "\u2013": "-",
        "\u2014": "-",
        "\u2026": "...",
        "\u00a0": " ",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    value = re.sub(r"(?<=[a-z0-9])([.!?])([A-Z])", r"\1 \2", value)
    value = re.sub(r"\s+", " ", value)
    value = value.replace("Ph<DOT>D<DOT>", "Ph.D.")
    return value.strip()


def split_name_and_description(paragraphs: list[str]) -> tuple[str, str]:
    if not paragraphs:
        return "Unknown Subject", ""

    if len(paragraphs) >= 2:
        name = normalize_text(paragraphs[0])
        desc = normalize_text(" ".join(paragraphs[1:]))
        return name, desc

    raw = normalize_text(paragraphs[0])
    raw = re.sub(r"\)(?=[A-Z])", ") | ", raw)
    if " | " in raw:
        name, desc = raw.split(" | ", 1)
        return normalize_text(name), normalize_text(desc)

    article_match = re.search(r"(?:(?<=\s)|(?<=[a-z\)]))(A|An|The)\s+[A-Za-z]", raw)
    if article_match and article_match.start() >= 8:
        name = raw[: article_match.start()].strip()
        desc = raw[article_match.start() :].strip()
        return normalize_text(name), normalize_text(desc)

    return raw, ""


def condense_description(text: str) -> str:
    text = normalize_text(text)
    if not text:
        return ""

    protected = text.replace("Dr. ", "Dr<DOT> ")
    protected = protected.replace("Ph.D.", "Ph<DOT>D<DOT>")
    protected = re.sub(
        r"(?:[A-Z]\.){2,}[A-Z]?\.?",
        lambda match: match.group(0).replace(".", "<DOT>"),
        protected,
    )
    sentences = [s.strip().replace("<DOT>", ".") for s in re.split(r"(?<=[.!?])\s+", protected) if s.strip()]
    chosen: list[str] = []
    max_chars = 165

    for sentence in sentences:
        candidate = " ".join(chosen + [sentence]).strip()
        if len(candidate) <= max_chars:
            chosen.append(sentence)
        else:
            break

    if not chosen:
        return textwrap.shorten(text, width=max_chars, placeholder="...")

    result = " ".join(chosen)
    if len(result) > max_chars:
        result = textwrap.shorten(result, width=max_chars, placeholder="...")
    return result


def latex_escape(value: str) -> str:
    value = normalize_text(value)
    replacements = {
        "\\": r"\textbackslash{}",
        "&": r"\&",
        "%": r"\%",
        "$": r"\$",
        "#": r"\#",
        "_": r"\_",
        "{": r"\{",
        "}": r"\}",
        "~": r"\textasciitilde{}",
        "^": r"\textasciicircum{}",
    }
    for old, new in replacements.items():
        value = value.replace(old, new)
    return value


def class_ability_text(card_class: str) -> str:
    return CLASS_ABILITIES.get(card_class, "No special field effect assigned.")


def apply_scientific_italics(value: str) -> str:
    for name in sorted(SCIENTIFIC_NAMES, key=len, reverse=True):
        value = value.replace(name, rf"\textit{{{name}}}")
    return value


def latex_format(value: str, *, italicize_scientific: bool = False) -> str:
    formatted = latex_escape(value)
    if italicize_scientific:
        formatted = apply_scientific_italics(formatted)
    return formatted


def slide_sort_key(path: str) -> int:
    return int(Path(path).stem.replace("slide", ""))


def ensure_pdf() -> None:
    if PDF.exists() and PDF.stat().st_mtime >= PPTX.stat().st_mtime:
        return

    soffice = Path(r"C:\Program Files\LibreOffice\program\soffice.exe")
    if not soffice.exists():
        raise FileNotFoundError(f"LibreOffice not found at {soffice}")

    result = subprocess.run(
        [str(soffice), "--headless", "--convert-to", "pdf", "--outdir", str(ROOT), str(PPTX)],
        capture_output=True,
        text=True,
    )
    if result.returncode != 0:
        details = "\n".join(part for part in [result.stdout.strip(), result.stderr.strip()] if part)
        raise RuntimeError(
            "LibreOffice failed to convert identity_cards.pptx to PDF. "
            "Close PowerPoint, LibreOffice, and any open PDF preview on the identity cards files, "
            "then rerun the script."
            + (f"\n\nLibreOffice output:\n{details}" if details else "")
        )


def render_slide_pngs() -> None:
    RENDER_DIR.mkdir(parents=True, exist_ok=True)
    existing = sorted(RENDER_DIR.glob("identity_cards-*.png"))
    if len(existing) == 45 and all(p.stat().st_mtime >= PDF.stat().st_mtime for p in existing):
        return

    prefix = RENDER_DIR / "identity_cards"
    subprocess.run(
        ["pdftoppm", "-r", "300", "-png", str(PDF), str(prefix)],
        check=True,
    )


def extract_records() -> tuple[int, int, list[SlideRecord]]:
    records: list[SlideRecord] = []
    with ZipFile(PPTX) as archive:
        presentation = etree.fromstring(archive.read("ppt/presentation.xml"))
        slide_size = presentation.find(".//p:sldSz", namespaces=NS)
        slide_w = int(slide_size.get("cx"))
        slide_h = int(slide_size.get("cy"))

        slide_files = sorted(
            [name for name in archive.namelist() if name.startswith("ppt/slides/slide") and name.endswith(".xml")],
            key=slide_sort_key,
        )

        for slide_file in slide_files:
            slide_index = slide_sort_key(slide_file)
            root = etree.fromstring(archive.read(slide_file))

            text_shapes = []
            for shape in root.xpath(".//p:sp", namespaces=NS):
                xfrm = shape.find(".//p:spPr/a:xfrm", namespaces=NS)
                if xfrm is None:
                    continue
                off = xfrm.find("./a:off", namespaces=NS)
                ext = xfrm.find("./a:ext", namespaces=NS)
                if off is None or ext is None:
                    continue
                paragraphs = []
                for para in shape.xpath("./p:txBody/a:p", namespaces=NS):
                    runs = [t.text for t in para.xpath(".//a:t", namespaces=NS) if t.text]
                    if runs:
                        paragraphs.append("".join(runs))
                if paragraphs:
                    text_shapes.append(
                        {
                            "y": int(off.get("y")),
                            "paragraphs": paragraphs,
                        }
                    )

            text_shapes.sort(key=lambda item: item["y"])
            card_class = normalize_text(text_shapes[0]["paragraphs"][0]) if text_shapes else "Unknown Class"
            body_paragraphs = []
            for shape in text_shapes[1:]:
                body_paragraphs.extend(shape["paragraphs"])
            name, description = split_name_and_description(body_paragraphs)

            pic = root.xpath(".//p:pic", namespaces=NS)[0]
            xfrm = pic.find(".//p:spPr/a:xfrm", namespaces=NS)
            off = xfrm.find("./a:off", namespaces=NS)
            ext = xfrm.find("./a:ext", namespaces=NS)

            records.append(
                SlideRecord(
                    index=slide_index,
                    card_class=card_class,
                    name=name,
                    description=DESCRIPTION_OVERRIDES.get(slide_index, condense_description(description)),
                    pic_x=int(off.get("x")),
                    pic_y=int(off.get("y")),
                    pic_cx=int(ext.get("cx")),
                    pic_cy=int(ext.get("cy")),
                )
            )

    return slide_w, slide_h, records


def generate_crops(slide_w: int, slide_h: int, records: list[SlideRecord]) -> None:
    ASSET_DIR.mkdir(parents=True, exist_ok=True)

    for record in records:
        slide_png = RENDER_DIR / f"identity_cards-{record.index:02d}.png"
        image = Image.open(slide_png).convert("RGB")
        scale_x = image.width / slide_w
        scale_y = image.height / slide_h

        left = max(0, int(record.pic_x * scale_x))
        top = max(0, int(record.pic_y * scale_y))
        right = min(image.width, int((record.pic_x + record.pic_cx) * scale_x))
        bottom = min(image.height, int((record.pic_y + record.pic_cy) * scale_y))

        crop = image.crop((left, top, right, bottom))
        fitted = ImageOps.fit(crop, (900, 1050), method=Image.Resampling.LANCZOS, centering=(0.5, 0.45))
        final_path = ASSET_DIR / f"card_{record.index:02d}.png"
        temp_path = ASSET_DIR / f"card_{record.index:02d}.tmp.png"
        fitted.save(temp_path, format="PNG", optimize=True)
        temp_path.replace(final_path)


def generate_tex(records: list[SlideRecord]) -> None:
    header = r"""\documentclass[10pt]{article}

\usepackage[letterpaper,margin=0.34in]{geometry}
\usepackage[T1]{fontenc}
\usepackage[scaled=0.95]{helvet}
\renewcommand{\familydefault}{\sfdefault}
\usepackage{microtype}
\usepackage{xcolor}
\usepackage{graphicx}
\usepackage[most]{tcolorbox}
\pagestyle{empty}
\setlength{\parindent}{0pt}
\setlength{\parskip}{0pt}

\definecolor{agencyblue}{HTML}{005A8B}
\definecolor{deepsea}{HTML}{143B5D}
\definecolor{labline}{HTML}{A9BBC8}
\definecolor{panelbg}{HTML}{F6F8FA}
\definecolor{titlebg}{HTML}{EAF2F7}
\definecolor{textgray}{HTML}{2F3B45}

\color{textgray}
\tcbset{
  sharp corners,
  boxrule=0.65pt,
  colframe=deepsea,
  colback=white,
  left=8pt,
  right=8pt,
  top=5pt,
  bottom=15pt,
  enhanced
}

\newcommand{\IdentityCard}[6]{%
\begin{tcolorbox}[
  height=4.95in,
  valign=top,
  overlay={
    \fill[titlebg] ([yshift=-0.27in]frame.north west) rectangle (frame.north east);
    \draw[labline,line width=0.4pt] ([yshift=-0.27in]frame.north west) -- ([yshift=-0.27in]frame.north east);
    \fill[panelbg] (frame.south west) rectangle ([yshift=0.18in]frame.south east);
    \draw[labline,line width=0.4pt] ([yshift=0.18in]frame.south west) -- ([yshift=0.18in]frame.south east);
    \node[anchor=north west,font=\fontsize{6.5}{7.0}\selectfont\bfseries\color{agencyblue}] at ([xshift=6pt,yshift=-4pt]frame.north west) {W.E.T. LAB PERSONNEL DOSSIER};
    \node[anchor=center,font=\fontsize{5.8}{6.3}\selectfont\bfseries\color{deepsea}] at ([xshift=1.55in,yshift=0.09in]frame.south west) {STATION 744 | INTERNAL USE};
    \node[anchor=center,font=\fontsize{5.3}{5.8}\selectfont\color{agencyblue}] at ([xshift=-0.33in,yshift=0.09in]frame.south east) {ID-#1};
    \draw[labline,line width=0.35pt] ([xshift=-0.66in]frame.south east) -- ([xshift=-0.66in,yshift=0.18in]frame.south east);
  }
]
\vspace*{19pt}
{\centering\bfseries\fontsize{25.2}{26.8}\selectfont\color{deepsea} #2\par}
\vspace{8pt}
\begin{center}
\begin{tcolorbox}[colback=panelbg,colframe=labline,boxrule=0.4pt,width=0.82\linewidth,left=0pt,right=0pt,top=7pt,bottom=7pt]
\includegraphics[height=1.93in]{#4}
\end{tcolorbox}
\end{center}
\vspace{7pt}
\begin{center}
\begin{minipage}{0.82\linewidth}
\centering
{\bfseries\fontsize{11.2}{12.0}\selectfont\color{deepsea} #3\par}
\vspace{4pt}
{\fontsize{7.35}{8.0}\selectfont #5\par}
\vspace{5pt}
\color{labline}\rule{\linewidth}{0.35pt}\par
\vspace{3pt}
{\fontsize{6.15}{6.7}\selectfont\bfseries\color{agencyblue}CLASS ABILITY\par}
\vspace{2pt}
{\fontsize{6.7}{7.35}\selectfont #6\par}
\end{minipage}
\end{center}
\end{tcolorbox}%
}

\begin{document}
\small
"""

    page_open = r"""\begin{tcbraster}[raster columns=2,raster column skip=8pt,raster row skip=8pt]"""
    page_close = r"""\end{tcbraster}"""
    footer = r"""
\end{document}
"""

    chunks: list[str] = [header, page_open]
    for idx, record in enumerate(records, 1):
        card_number = f"{idx:02d}"
        image_path = f"card_assets/card_{record.index:02d}.png"
        chunks.append(
            "\\IdentityCard"
            f"{{{card_number}}}"
            f"{{{latex_format(record.card_class)}}}"
            f"{{{latex_format(record.name, italicize_scientific=True)}}}"
            f"{{{image_path}}}"
            f"{{{latex_format(record.description, italicize_scientific=True)}}}"
            f"{{{latex_format(class_ability_text(record.card_class))}}}"
        )
        if idx != len(records) and idx % 4 == 0:
            chunks.append(page_close)
            chunks.append("\\newpage")
            chunks.append(page_open)

    chunks.append(page_close)
    chunks.append(footer)
    temp_tex = TEX_OUT.with_suffix(".tmp.tex")
    temp_tex.write_text("\n".join(chunks), encoding="utf-8")
    temp_tex.replace(TEX_OUT)


def main() -> None:
    ensure_pdf()
    render_slide_pngs()
    slide_w, slide_h, records = extract_records()
    generate_crops(slide_w, slide_h, records)
    generate_tex(records)
    print(f"Generated {TEX_OUT}")


if __name__ == "__main__":
    main()
