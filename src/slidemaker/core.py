"""Core utilities for manipulating python-pptx presentations.

Provides low-level helper functions used by the template module
to find shapes, set text, add bullet lists, and clone slides.
"""

from __future__ import annotations

import copy
import re
from typing import Any, Optional

from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.xmlchemy import OxmlElement
from pptx.presentation import Presentation
from pptx.slide import Slide
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


# ── Style constants ─────────────────────────────────────────────
FONT_NAME = "Montserrat"
FONT_COLOR = RGBColor(0x0B, 0x1F, 0x33)
TITLE_FONT_SIZE = Pt(51)
SUBTITLE_FONT_SIZE = Pt(36)
BODY_FONT_SIZE = Pt(30)
SLIDE_W = Emu(18288000)  # 20 inches
SLIDE_H = Emu(10287000)  # 11.25 inches

# Content area defaults (below the title + decorative line)
CONTENT_LEFT = Inches(0.9)
CONTENT_TOP = Inches(2.6)
CONTENT_WIDTH = Inches(12.5)
CONTENT_HEIGHT = Inches(7.5)

# Brand palette — reusable colours for flow boxes etc.
BRAND = RGBColor(0x19, 0x39, 0x52)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
ACCENT = RGBColor(0x2E, 0x86, 0xAB)
ACCENT2 = RGBColor(0xE8, 0x6F, 0x51)
GREEN = RGBColor(0x48, 0xA9, 0x9A)
CODE_FONT = "Consolas"
CODE_FONT_SIZE = Pt(20)

_ALIGNMENTS = {
    "left": PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right": PP_ALIGN.RIGHT,
    "justify": PP_ALIGN.JUSTIFY,
}
_CODE_LINE_PREFIX_RE = re.compile(r"^\s{0,2}\d+\s{1,2}")
_INLINE_BOLD_RE = re.compile(r"\*\*(.+?)\*\*")


def _normalize_style(style: Optional[dict[str, Any]]) -> dict[str, Any]:
    """Normalize style keys to kebab-case lowercase."""
    if not style:
        return {}
    normalized: dict[str, Any] = {}
    for key, value in style.items():
        if not isinstance(key, str):
            continue
        normalized[key.strip().lower().replace("_", "-")] = value
    return normalized


def _merge_style(
    base: Optional[dict[str, Any]],
    override: Optional[dict[str, Any]],
) -> dict[str, Any]:
    """Return a shallow merge where ``override`` wins."""
    merged = dict(base or {})
    if override:
        merged.update(_normalize_style(override))
    return merged


def _as_rgb_color(
    value: Any,
    default: Optional[RGBColor] = None,
) -> Optional[RGBColor]:
    """Parse an RGBColor from common style representations."""
    if value is None:
        return default
    if isinstance(value, RGBColor):
        return value
    if isinstance(value, str):
        text = value.strip()
        if text.startswith("#"):
            text = text[1:]
        if len(text) == 6:
            try:
                return RGBColor(
                    int(text[0:2], 16),
                    int(text[2:4], 16),
                    int(text[4:6], 16),
                )
            except ValueError:
                return default
    if isinstance(value, (tuple, list)) and len(value) == 3:
        try:
            r, g, b = (int(v) for v in value)
        except (TypeError, ValueError):
            return default
        if all(0 <= v <= 255 for v in (r, g, b)):
            return RGBColor(r, g, b)
    return default


def _as_pt(value: Any, default: Optional[int] = None) -> Optional[int]:
    """Parse a point-sized value as EMU using ``Pt``."""
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return Pt(float(value))
    if isinstance(value, str):
        text = value.strip().lower()
        if text.endswith("pt"):
            text = text[:-2].strip()
        try:
            return Pt(float(text))
        except ValueError:
            return default
    return default


def _font_size_pt(font_size: Any) -> Optional[float]:
    """Extract font size in points from a ``Length`` value."""
    if font_size is None:
        return None
    pt = getattr(font_size, "pt", None)
    if pt is not None:
        try:
            return float(pt)
        except (TypeError, ValueError):
            return None
    return None


def _resolve_line_spacing(
    value: Any,
    font_size: Any,
    default: Any = None,
) -> Any:
    """Resolve line spacing for Canva-friendly output.

    Numeric values are treated as multipliers and converted
    to fixed point leading based on font size. Use ``\"pt\"``
    suffix for absolute point values.
    """
    if value is None:
        return default

    def as_points(multiplier: float) -> Any:
        font_pt = _font_size_pt(font_size)
        if font_pt is None:
            return multiplier
        return Pt(font_pt * multiplier)

    if isinstance(value, (int, float)):
        return as_points(float(value))

    if isinstance(value, str):
        text = value.strip().lower()
        if text.endswith("pt"):
            text = text[:-2].strip()
            try:
                return Pt(float(text))
            except ValueError:
                return default
        if text.endswith("x"):
            text = text[:-1].strip()
            try:
                return as_points(float(text))
            except ValueError:
                return default
        if text.endswith("%"):
            text = text[:-1].strip()
            try:
                return as_points(float(text) / 100.0)
            except ValueError:
                return default
        try:
            return as_points(float(text))
        except ValueError:
            return default

    return default


def _resolve_letter_spacing(
    value: Any,
    font_size: Any,
    default: Optional[int] = None,
) -> Optional[int]:
    """Resolve letter spacing for ``a:rPr@spc``.

    Numeric values are treated as tracking units relative to
    font size (Canva style). ``\"pt\"`` values are absolute.
    """
    if value is None:
        return default

    def tracking_to_spc(tracking: float) -> Optional[int]:
        font_pt = _font_size_pt(font_size)
        if font_pt is None:
            return int(round(tracking))
        # Canva-style tracking (per-thousand of em) to OOXML ST_TextPoint.
        return int(round(tracking * font_pt / 10.0))

    if isinstance(value, (int, float)):
        return tracking_to_spc(float(value))

    if isinstance(value, str):
        text = value.strip().lower()
        if text.endswith("pt"):
            text = text[:-2].strip()
            try:
                return int(round(float(text) * 100))
            except ValueError:
                return default
        try:
            return tracking_to_spc(float(text))
        except ValueError:
            return default

    return default


def _as_bool(value: Any, default: Optional[bool] = None) -> Optional[bool]:
    """Parse a boolean from common string/int representations."""
    if value is None:
        return default
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return bool(value)
    if isinstance(value, str):
        text = value.strip().lower()
        if text in {"true", "1", "yes", "on"}:
            return True
        if text in {"false", "0", "no", "off"}:
            return False
    return default


def _resolve_uppercase(
    normalized: dict[str, Any],
    default: bool = False,
) -> bool:
    """Resolve uppercase transform from style keys."""
    text_transform = normalized.get("text-transform")
    if isinstance(text_transform, str):
        mode = text_transform.strip().lower()
        if mode == "uppercase":
            return True
        if mode in {"none", "normal", "initial"}:
            return False
    return bool(_as_bool(normalized.get("uppercase"), default))


def _apply_uppercase(text: str, uppercase: bool) -> str:
    """Apply uppercase transform when enabled."""
    return text.upper() if uppercase else text


def _as_alignment(value: Any, default: Optional[int] = None) -> Optional[int]:
    """Parse paragraph alignment from text/int values."""
    if value is None:
        return default
    if isinstance(value, int):
        return value
    if isinstance(value, str):
        return _ALIGNMENTS.get(value.strip().lower(), default)
    return default


def _with_code_line_numbers(code_text: str) -> str:
    """Prefix code lines with right-aligned line numbers.

    Uses ``" X  "`` for one-digit lines and ``"XX  "`` for
    two-digit lines. If all non-empty lines already look
    numbered, the input is returned unchanged.
    """
    lines = code_text.splitlines()
    non_empty = [line for line in lines if line.strip()]
    if non_empty and all(_CODE_LINE_PREFIX_RE.match(line) for line in non_empty):
        return code_text
    return "\n".join(f"{idx:>2}  {line}" for idx, line in enumerate(lines, 1))


def _markdown_bold_segments(text: str) -> list[tuple[str, bool]]:
    """Split text into plain/bold segments based on ``**...**`` markup."""
    segments: list[tuple[str, bool]] = []
    last = 0
    for match in _INLINE_BOLD_RE.finditer(text):
        if match.start() > last:
            plain = text[last : match.start()].replace("**", "")
            if plain:
                segments.append((plain, False))
        bold_text = match.group(1)
        if bold_text:
            segments.append((bold_text, True))
        last = match.end()

    tail = text[last:].replace("**", "")
    if tail:
        segments.append((tail, False))

    if not segments:
        clean = text.replace("**", "")
        if clean:
            segments.append((clean, False))
    return segments


def _resolve_padding(
    normalized: dict[str, Any],
    default: Optional[int] = None,
) -> tuple[Optional[int], Optional[int], Optional[int], Optional[int]]:
    """Resolve padding values (left, top, right, bottom) in EMU."""
    pad_all = _as_pt(normalized.get("padding"), default)
    pad_x = _as_pt(normalized.get("padding-x"), pad_all)
    pad_y = _as_pt(normalized.get("padding-y"), pad_all)

    pad_left = _as_pt(normalized.get("padding-left"), pad_x)
    pad_top = _as_pt(normalized.get("padding-top"), pad_y)
    pad_right = _as_pt(normalized.get("padding-right"), pad_x)
    pad_bottom = _as_pt(normalized.get("padding-bottom"), pad_y)
    return pad_left, pad_top, pad_right, pad_bottom


def _apply_text_frame_padding(
    tf: Any,
    normalized: dict[str, Any],
    default: Optional[int] = None,
) -> None:
    """Apply resolved padding values to a text frame margins."""
    pad_left, pad_top, pad_right, pad_bottom = _resolve_padding(normalized, default)
    if pad_left is not None:
        tf.margin_left = pad_left
    if pad_top is not None:
        tf.margin_top = pad_top
    if pad_right is not None:
        tf.margin_right = pad_right
    if pad_bottom is not None:
        tf.margin_bottom = pad_bottom


def _apply_run_letter_spacing(run: Any, spacing: Optional[int]) -> None:
    """Apply letter spacing to a run via ``a:rPr@spc`` (centipoints)."""
    if spacing is None:
        return
    r_pr = run._r.get_or_add_rPr()
    r_pr.set("spc", str(spacing))


def find_group_textbox(slide: Slide, group_name: str) -> Any:
    """Find the first TextBox inside a named Group shape.

    Parameters
    ----------
    slide : Slide
        The slide to search.
    group_name : str
        The ``name`` attribute of the Group shape.

    Returns
    -------
    pptx.shapes.autoshape.Shape or None
        The TextBox shape if found, otherwise ``None``.
    """
    for shape in slide.shapes:
        if shape.name == group_name and shape.shape_type == 6:
            for child in shape.shapes:  # type: ignore[attr-defined]
                if (
                    child.has_text_frame and child.shape_type == 17  # TEXT_BOX
                ):
                    return child
    return None


def find_textbox_by_name(slide: Slide, name: str) -> Any:
    """Find a shape by its exact name on a slide.

    Parameters
    ----------
    slide : Slide
        The slide to search.
    name : str
        The ``name`` attribute of the shape.

    Returns
    -------
    pptx.shapes.autoshape.Shape or None
        The shape if found, otherwise ``None``.
    """
    for shape in slide.shapes:
        if shape.name == name:
            return shape
    return None


def set_textbox_text(
    shape: Any,
    text: str,
    font_size: Optional[int] = None,
    font_color: Optional[RGBColor] = None,
    font_name: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    alignment: Optional[int] = None,
    style: Optional[dict[str, Any]] = None,
) -> None:
    """Replace all text in a shape's text frame.

    Parameters
    ----------
    shape : Shape
        A shape with a ``text_frame`` attribute.
    text : str
        The replacement text.
    font_size : int, optional
        Font size in EMU (use ``Pt()``).
    font_color : RGBColor, optional
        Font colour.
    font_name : str, optional
        Font family name.
    bold : bool, optional
        Whether the text is bold.
    alignment : int, optional
        Paragraph alignment constant.
    """
    normalized = _normalize_style(style)
    resolved_uppercase = _resolve_uppercase(normalized, False)
    tf = shape.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    p.text = _apply_uppercase(text, resolved_uppercase)
    _apply_text_frame_padding(tf, normalized)

    resolved_alignment = alignment
    if resolved_alignment is None:
        resolved_alignment = _as_alignment(
            normalized.get("alignment", normalized.get("align")),
        )
    if resolved_alignment is not None:
        p.alignment = resolved_alignment

    # python-pptx may produce zero runs when text is empty.
    if not p.runs:
        return

    resolved_font_size = (
        font_size if font_size is not None else _as_pt(normalized.get("font-size"))
    )
    resolved_font_color = (
        font_color
        if font_color is not None
        else _as_rgb_color(normalized.get("font-color"))
    )
    resolved_font_name = (
        font_name if font_name is not None else normalized.get("font-name")
    )
    if resolved_font_name is not None:
        resolved_font_name = str(resolved_font_name)
    resolved_bold = bold if bold is not None else _as_bool(normalized.get("bold"))
    resolved_italic = (
        italic if italic is not None else _as_bool(normalized.get("italic"))
    )
    resolved_line_spacing = _resolve_line_spacing(
        normalized.get("line-spacing", normalized.get("line-height")),
        resolved_font_size,
    )
    if resolved_line_spacing is not None:
        p.line_spacing = resolved_line_spacing
    resolved_letter_spacing = _resolve_letter_spacing(
        normalized.get("letter-spacing"),
        resolved_font_size,
    )

    run = p.runs[0]
    if resolved_font_size is not None:
        run.font.size = resolved_font_size
    if resolved_font_color is not None:
        run.font.color.rgb = resolved_font_color
    if resolved_font_name is not None:
        run.font.name = resolved_font_name
    if resolved_bold is not None:
        run.font.bold = resolved_bold
    if resolved_italic is not None:
        run.font.italic = resolved_italic
    _apply_run_letter_spacing(run, resolved_letter_spacing)


def add_textbox(
    slide: Slide,
    left: int,
    top: int,
    width: int,
    height: int,
    text: str,
    font_size: Optional[int] = None,
    font_color: Optional[RGBColor] = None,
    font_name: Optional[str] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    alignment: Optional[int] = None,
    style: Optional[dict[str, Any]] = None,
) -> object:
    """Add a new text box to a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    left : int
        Horizontal position in EMU.
    top : int
        Vertical position in EMU.
    width : int
        Box width in EMU.
    height : int
        Box height in EMU.
    text : str
        The text content.
    font_size : int
        Font size in EMU (use ``Pt()``).
    font_color : RGBColor
        Font colour.
    font_name : str
        Font family name.
    bold : bool
        Whether the text is bold.
    alignment : int
        Paragraph alignment constant.

    Returns
    -------
    pptx.shapes.autoshape.Shape
        The newly created text box shape.
    """
    normalized = _normalize_style(style)
    resolved_font_size = (
        font_size
        if font_size is not None
        else _as_pt(normalized.get("font-size"), BODY_FONT_SIZE)
    )
    resolved_font_color = (
        font_color
        if font_color is not None
        else _as_rgb_color(normalized.get("font-color"), FONT_COLOR)
    )
    resolved_font_name_raw = (
        font_name if font_name is not None else normalized.get("font-name", FONT_NAME)
    )
    resolved_font_name = (
        FONT_NAME if resolved_font_name_raw is None else str(resolved_font_name_raw)
    )
    resolved_bold = (
        bold if bold is not None else _as_bool(normalized.get("bold"), False)
    )
    resolved_italic = (
        italic if italic is not None else _as_bool(normalized.get("italic"), False)
    )
    resolved_alignment = (
        alignment
        if alignment is not None
        else _as_alignment(
            normalized.get("alignment", normalized.get("align")),
            PP_ALIGN.LEFT,
        )
    )
    resolved_line_spacing = _resolve_line_spacing(
        normalized.get("line-spacing", normalized.get("line-height")),
        resolved_font_size,
    )
    resolved_letter_spacing = _resolve_letter_spacing(
        normalized.get("letter-spacing"),
        resolved_font_size,
    )
    resolved_uppercase = _resolve_uppercase(normalized, False)

    txbox = slide.shapes.add_textbox(left, top, width, height)  # type: ignore[arg-type]
    tf = txbox.text_frame
    tf.word_wrap = True
    _apply_text_frame_padding(tf, normalized)
    p = tf.paragraphs[0]
    p.text = _apply_uppercase(text, resolved_uppercase)
    p.font.size = resolved_font_size
    p.font.color.rgb = resolved_font_color
    p.font.name = resolved_font_name
    p.font.bold = resolved_bold
    p.font.italic = resolved_italic
    p.alignment = resolved_alignment
    if resolved_line_spacing is not None:
        p.line_spacing = resolved_line_spacing
    for run in p.runs:
        _apply_run_letter_spacing(run, resolved_letter_spacing)
    return txbox


def add_bullet_list(
    slide: Slide,
    left: int,
    top: int,
    width: int,
    height: int,
    items: list[str],
    font_size: Optional[int] = None,
    font_color: Optional[RGBColor] = None,
    font_name: Optional[str] = None,
    spacing: Optional[int] = None,
    bullet_char: Optional[str] = None,
    bold_prefixes: Optional[bool] = None,
    style: Optional[dict[str, Any]] = None,
) -> object:
    """Add a bulleted list as a text box on a slide.

    Supports basic inline markdown bold: ``**text**`` segments
    render in bold when ``bold_prefixes`` is ``True``.

    Parameters
    ----------
    slide : Slide
        Target slide.
    left : int
        Horizontal position in EMU.
    top : int
        Vertical position in EMU.
    width : int
        Box width in EMU.
    height : int
        Box height in EMU.
    items : list of str
        Bullet point strings.
    font_size : int
        Font size in EMU.
    font_color : RGBColor
        Font colour.
    font_name : str
        Font family name.
    spacing : int
        Space after each paragraph in EMU.
    bullet_char : str
        Character used for real paragraph bullets.
    bold_prefixes : bool
        If ``True``, parse and render ``**...**`` segments in
        bold within each item.

    Returns
    -------
    pptx.shapes.autoshape.Shape
        The newly created text box shape.
    """
    normalized = _normalize_style(style)
    resolved_font_size = (
        font_size
        if font_size is not None
        else _as_pt(normalized.get("font-size"), BODY_FONT_SIZE)
    )
    resolved_font_color = (
        font_color
        if font_color is not None
        else _as_rgb_color(normalized.get("font-color"), FONT_COLOR)
    )
    resolved_font_name_raw = (
        font_name if font_name is not None else normalized.get("font-name", FONT_NAME)
    )
    resolved_font_name = (
        FONT_NAME if resolved_font_name_raw is None else str(resolved_font_name_raw)
    )
    resolved_spacing = (
        spacing
        if spacing is not None
        else _as_pt(
            normalized.get("spacing", normalized.get("space-after")),
            Pt(10),
        )
    )
    resolved_bullet_char = (
        bullet_char
        if bullet_char is not None
        else str(normalized.get("bullet-char", "•"))
    )
    resolved_bold_prefixes = (
        bold_prefixes
        if bold_prefixes is not None
        else _as_bool(normalized.get("bold-prefixes"), True)
    )
    resolved_alignment = _as_alignment(
        normalized.get("alignment", normalized.get("align")),
        PP_ALIGN.LEFT,
    )
    resolved_space_before = _as_pt(normalized.get("space-before"), Pt(2))
    resolved_italic = _as_bool(normalized.get("italic"), False)
    resolved_line_spacing = _resolve_line_spacing(
        normalized.get("line-spacing", normalized.get("line-height")),
        resolved_font_size,
    )
    resolved_letter_spacing = _resolve_letter_spacing(
        normalized.get("letter-spacing"),
        resolved_font_size,
    )
    resolved_uppercase = _resolve_uppercase(normalized, False)

    txbox = slide.shapes.add_textbox(left, top, width, height)  # type: ignore[arg-type]
    tf = txbox.text_frame
    tf.word_wrap = True
    _apply_text_frame_padding(tf, normalized)

    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = resolved_spacing
        p.space_before = resolved_space_before
        p.alignment = resolved_alignment
        if resolved_line_spacing is not None:
            p.line_spacing = resolved_line_spacing

        # Configure a real paragraph bullet (not a text prefix).
        pPr = p._p.get_or_add_pPr()
        for child in list(pPr):
            if child.tag.endswith(("}buNone", "}buAutoNum", "}buChar", "}buBlip")):
                pPr.remove(child)
        bu_char = OxmlElement("a:buChar")
        bu_char.set("char", resolved_bullet_char)
        pPr.append(bu_char)

        item_text = _apply_uppercase(item.strip(), resolved_uppercase)
        if resolved_bold_prefixes:
            segments = _markdown_bold_segments(item_text)
        else:
            segments = [(item_text.replace("**", ""), False)]

        for segment_text, is_bold in segments:
            run = p.add_run()
            run.text = segment_text
            run.font.size = resolved_font_size
            run.font.color.rgb = resolved_font_color
            run.font.name = resolved_font_name
            run.font.bold = is_bold
            run.font.italic = resolved_italic
            _apply_run_letter_spacing(run, resolved_letter_spacing)

    return txbox


def add_shape_rect(
    slide: Slide,
    left: int,
    top: int,
    width: int,
    height: int,
    fill_color: Optional[RGBColor] = None,
    line_color: Optional[RGBColor] = None,
    line_width: Optional[int] = None,
    style: Optional[dict[str, Any]] = None,
) -> object:
    """Add a rectangle shape to a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    left : int
        Horizontal position in EMU.
    top : int
        Vertical position in EMU.
    width : int
        Shape width in EMU.
    height : int
        Shape height in EMU.
    fill_color : RGBColor, optional
        Fill colour. If ``None`` the shape has no fill.

    Returns
    -------
    pptx.shapes.autoshape.Shape
        The newly created rectangle shape.
    """
    normalized = _normalize_style(style)
    resolved_fill = (
        fill_color
        if fill_color is not None
        else _as_rgb_color(normalized.get("fill-color"))
    )
    resolved_line_color = (
        line_color
        if line_color is not None
        else _as_rgb_color(normalized.get("line-color"))
    )
    resolved_line_width = (
        line_width if line_width is not None else _as_pt(normalized.get("line-width"))
    )

    shp = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)  # type: ignore[arg-type]
    if resolved_line_color is not None:
        shp.line.color.rgb = resolved_line_color
        if resolved_line_width is not None:
            shp.line.width = resolved_line_width
    else:
        shp.line.fill.background()
    if resolved_fill:
        shp.fill.solid()
        shp.fill.fore_color.rgb = resolved_fill
    else:
        shp.fill.background()
    return shp


def add_code_block(
    slide: Slide,
    left: int,
    top: int,
    width: int,
    height: int,
    code_text: str,
    bg_color: Optional[RGBColor] = None,
    font_size: Optional[int] = None,
    font_color: Optional[RGBColor] = None,
    font_name: Optional[str] = None,
    style: Optional[dict[str, Any]] = None,
) -> None:
    """Add a code block with a dark background to a slide.

    Renders monospace text on a solid-colour rectangle.
    Style options include ``bg-color``, ``font-color``,
    ``font-name``, ``font-size``, ``line-numbers``,
    ``letter-spacing``, ``line-spacing`` (or ``line-height``),
    ``uppercase`` (or ``text-transform: uppercase``), and
    ``padding`` (plus side-specific variants).

    Parameters
    ----------
    slide : Slide
        Target slide.
    left : int
        Horizontal position in EMU.
    top : int
        Vertical position in EMU.
    width : int
        Block width in EMU.
    height : int
        Block height in EMU.
    code_text : str
        The source code to display.
    bg_color : RGBColor
        Background rectangle colour.
    font_size : int
        Font size for the code text.
    """
    normalized = _normalize_style(style)
    resolved_bg_color = (
        bg_color
        if bg_color is not None
        else _as_rgb_color(
            normalized.get("bg-color", normalized.get("fill-color")),
            BRAND,
        )
    )
    resolved_font_size = (
        font_size
        if font_size is not None
        else _as_pt(normalized.get("font-size"), CODE_FONT_SIZE)
    )
    resolved_font_color = (
        font_color
        if font_color is not None
        else _as_rgb_color(normalized.get("font-color"), WHITE)
    )
    resolved_font_name_raw = (
        font_name if font_name is not None else normalized.get("font-name", CODE_FONT)
    )
    resolved_font_name = (
        CODE_FONT if resolved_font_name_raw is None else str(resolved_font_name_raw)
    )
    resolved_bold = _as_bool(normalized.get("bold"), False)
    resolved_line_spacing = _resolve_line_spacing(
        normalized.get("line-spacing", normalized.get("line-height")),
        resolved_font_size,
    )
    resolved_letter_spacing = _resolve_letter_spacing(
        normalized.get("letter-spacing"),
        resolved_font_size,
    )
    resolved_uppercase = _resolve_uppercase(normalized, False)
    resolved_line_numbers = _as_bool(normalized.get("line-numbers"), False)
    if resolved_line_numbers:
        code_text = _with_code_line_numbers(code_text)

    shp = add_shape_rect(slide, left, top, width, height, fill_color=resolved_bg_color)
    shp.shadow.inherit = False  # type: ignore[attr-defined]

    pad_left, pad_top, pad_right, pad_bottom = _resolve_padding(
        normalized, Inches(0.25)
    )
    inner_width = max(1, width - (pad_left or 0) - (pad_right or 0))
    inner_height = max(1, height - (pad_top or 0) - (pad_bottom or 0))
    txbox = slide.shapes.add_textbox(
        left + (pad_left or 0),  # type: ignore[arg-type]
        top + (pad_top or 0),  # type: ignore[arg-type]
        inner_width,  # type: ignore[arg-type]
        inner_height,  # type: ignore[arg-type]
    )
    tf = txbox.text_frame
    tf.word_wrap = True

    for i, line in enumerate(code_text.strip("\n").split("\n")):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.space_after = Pt(2)
        p.space_before = Pt(0)
        if resolved_line_spacing is not None:
            p.line_spacing = resolved_line_spacing
        run = p.add_run()
        run.text = _apply_uppercase(line, resolved_uppercase)
        run.font.size = resolved_font_size
        run.font.color.rgb = resolved_font_color
        run.font.name = resolved_font_name
        run.font.bold = resolved_bold
        _apply_run_letter_spacing(run, resolved_letter_spacing)


def add_flow_boxes(
    slide: Slide,
    boxes: list[dict],
    left: int = CONTENT_LEFT,
    top: int = CONTENT_TOP,
    box_width: Optional[int] = None,
    box_height: int = Inches(2.4),
    gap: int = Inches(0.5),
    style: Optional[dict[str, Any]] = None,
) -> None:
    """Add a horizontal flow diagram with coloured boxes and arrows.

    Each box is a dictionary with keys:

    - ``label`` (str): bold heading inside the box.
    - ``desc`` (str): smaller description text (may contain
      newlines).
    - ``style`` (dict): style attributes for this box
      (for example ``fill-color``/``font-color``).
    - ``color`` (str or RGBColor): legacy fill colour key.

    Parameters
    ----------
    slide : Slide
        Target slide.
    boxes : list of dict
        Box definitions (see above).
    left : int
        Horizontal start position in EMU.
    top : int
        Vertical position in EMU.
    box_width : int, optional
        Width of each box.  If ``None`` it is computed to fill
        the available content width.
    box_height : int
        Height of each box.
    gap : int
        Space between boxes (includes arrow).
    style : dict, optional
        Default style for all boxes in the flow.
    """
    n = len(boxes)
    if n == 0:
        return

    if box_width is None:
        avail = CONTENT_WIDTH - gap * (n - 1)
        box_width = int(avail / n)

    base_style = _normalize_style(style)

    for i, box in enumerate(boxes):
        x = left + i * (box_width + gap)
        box_style = _merge_style(base_style, box.get("style"))
        color = _as_rgb_color(
            box_style.get("fill-color", box_style.get("color", box.get("color"))),
            ACCENT,
        )
        label_style = _merge_style(
            {
                "font-size": 28,
                "font-color": "#FFFFFF",
                "bold": True,
                "alignment": "center",
            },
            box_style,
        )
        desc_style = _merge_style(
            {
                "font-size": 22,
                "font-color": "#FFFFFF",
                "bold": False,
                "alignment": "center",
            },
            box_style,
        )
        arrow_style = _merge_style(
            {
                "font-size": 44,
                "font-color": "#193952",
                "bold": True,
                "alignment": "center",
            },
            box_style,
        )
        if "arrow-color" in box_style:
            arrow_style["font-color"] = box_style["arrow-color"]
        if "arrow-font-size" in box_style:
            arrow_style["font-size"] = box_style["arrow-font-size"]

        shp = add_shape_rect(
            slide,
            x,
            top,
            box_width,
            box_height,
            fill_color=color,
        )
        shp.shadow.inherit = False  # type: ignore[attr-defined]

        # Label
        add_textbox(
            slide,
            x,
            top + Inches(0.3),
            box_width,
            Inches(0.7),
            box["label"],
            style=label_style,
        )
        # Description
        if box.get("desc"):
            add_textbox(
                slide,
                x + Inches(0.2),
                top + Inches(1.1),
                box_width - Inches(0.4),
                box_height - Inches(1.3),
                box["desc"],
                style=desc_style,
            )

        # Arrow between boxes
        if i < n - 1:
            ax = x + box_width + Inches(0.02)
            add_textbox(
                slide,
                ax,
                top + Inches(0.6),
                gap - Inches(0.04),
                Inches(0.7),
                "→",
                style=arrow_style,
            )


def set_notes(slide: Slide, text: str) -> None:
    """Set the speaker notes for a slide.

    Parameters
    ----------
    slide : Slide
        Target slide.
    text : str
        The notes content.
    """
    notes_slide = slide.notes_slide
    tf = notes_slide.notes_text_frame
    if tf is not None:
        tf.text = text


def clone_slide(prs: Presentation, template_idx: int) -> Slide:
    """Clone a slide from the presentation by index.

    Creates a deep copy of the slide's XML and all
    relationships, appending the new slide at the end.

    Parameters
    ----------
    prs : Presentation
        The presentation object.
    template_idx : int
        Zero-based index of the slide to clone.

    Returns
    -------
    Slide
        The newly created slide.
    """
    template_slide = prs.slides[template_idx]
    slide_layout = template_slide.slide_layout

    new_slide = prs.slides.add_slide(slide_layout)

    # Remove default shapes from the new slide
    for shape in list(new_slide.shapes):
        sp_elem = shape._element
        sp_elem.getparent().remove(sp_elem)

    # Copy all elements from template slide
    for shape in template_slide.shapes:
        el = copy.deepcopy(shape._element)
        new_slide.shapes._spTree.append(el)

    # Copy slide background
    if template_slide.background._element is not None:
        # Use the csld element's bg
        csld = new_slide._element.find(
            "{http://schemas.openxmlformats.org/presentationml/2006/main}cSld"
        )
        if csld is not None:
            old_bg = csld.find(
                "{http://schemas.openxmlformats.org/presentationml/2006/main}bg"
            )
            if old_bg is not None:
                csld.remove(old_bg)

    # Copy relationships (images, etc.)
    for rel in template_slide.part.rels.values():
        if "image" in rel.reltype:
            new_slide.part.rels.get_or_add(rel.reltype, rel._target)

    return new_slide


def delete_slide(prs: Presentation, slide_idx: int) -> None:
    """Delete a slide from the presentation by index.

    Parameters
    ----------
    prs : Presentation
        The presentation object.
    slide_idx : int
        Zero-based index of the slide to delete.
    """
    rId = prs.slides._sldIdLst[slide_idx].get(
        "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"
    )
    prs.part.drop_rel(rId)
    sldId = prs.slides._sldIdLst[slide_idx]
    prs.slides._sldIdLst.remove(sldId)


def move_slide(prs: Presentation, old_idx: int, new_idx: int) -> None:
    """Move a slide from one position to another.

    Parameters
    ----------
    prs : Presentation
        The presentation object.
    old_idx : int
        Current zero-based index of the slide.
    new_idx : int
        Desired zero-based index for the slide.
    """
    sld_id_lst = prs.slides._sldIdLst
    el = sld_id_lst[old_idx]
    sld_id_lst.remove(el)
    if new_idx >= len(sld_id_lst):
        sld_id_lst.append(el)
    else:
        sld_id_lst.insert(new_idx, el)
