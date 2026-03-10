"""
title: Slide-specific template builders for presentations.
summary: |-
  Each function populates one slide type from the branded template.
  The template has 9 slides mapped as follows:

  Template   Generated
  --------   ---------
  1          1          Title page
  2          2          Learning Objectives
  3          3          Core Toolkit Recap
  4          4          What's New in This Lesson
  5          5..n-4     Default content slides (cloned)
  6          n-3        Validation Checkpoints
  7          n-2        Exercise Playbook
  8          n-1        Debugging Guide
  9          n          Recap and Next Steps
"""

from __future__ import annotations

from typing import Any, Optional

from pptx.slide import Slide
from pptx.util import Inches

from slidemaker.anchors import default_anchor_map
from slidemaker.core import (
    CONTENT_TOP,
    FONT_NAME,
    add_bullet_list,
    add_code_block,
    add_flow_boxes,
    add_textbox,
    find_group_textbox,
    find_textbox_by_name,
    set_notes,
    set_textbox_text,
)

# ── Title group names per slide index (0-based) ────────────────
# The title TextBox lives inside a Group shape whose name
# varies across slides.  These mappings were read from the
# template inspection.
_TITLE_GROUPS: dict[int, str] = {
    0: "Group 7",  # slide 1 — lesson number label
    1: "Group 3",  # slide 2
    2: "Group 4",  # slide 3
    3: "Group 4",  # slide 4
    4: "Group 4",  # slide 5 (default)
    5: "Group 4",  # slide 6 (checkpoints)
    6: "Group 4",  # slide 7 (exercise)
    7: "Group 4",  # slide 8 (debugging)
    8: "Group 4",  # slide 9 (recap)
}

StyleAttrs = dict[str, Any]
StyleMap = dict[str, StyleAttrs]
AnchorMap = dict[str, Any]
_FALLBACK_ANCHORS: AnchorMap = default_anchor_map()


def _pick(mapping: dict[Any, Any], key: Any, default: Any = None) -> Any:
    """
    title: Get value by key supporting both int and string page keys.
    parameters:
      mapping:
        type: dict[Any, Any]
      key:
        type: Any
      default:
        type: Any
    returns:
      type: Any
    """
    if key in mapping:
        return mapping[key]
    str_key = str(key)
    if str_key in mapping:
        return mapping[str_key]
    return default


def _title_group_for(anchors: Optional[AnchorMap], page_idx: int) -> str:
    """
    title: Resolve title group name for a 0-based template page index.
    parameters:
      anchors:
        type: Optional[AnchorMap]
      page_idx:
        type: int
    returns:
      type: str
    """
    page = page_idx + 1
    groups = anchors.get("title-groups", {}) if isinstance(anchors, dict) else {}
    if not isinstance(groups, dict):
        groups = {}
    fallback = _pick(_FALLBACK_ANCHORS["title-groups"], page, _TITLE_GROUPS[page_idx])
    value = _pick(groups, page, fallback)
    return str(value)


def _title_slide_group(
    anchors: Optional[AnchorMap],
    key: str,
    fallback: str,
) -> str:
    """
    title: Resolve title-slide group names (title/subtitle).
    parameters:
      anchors:
        type: Optional[AnchorMap]
      key:
        type: str
      fallback:
        type: str
    returns:
      type: str
    """
    section = anchors.get("title-slide", {}) if isinstance(anchors, dict) else {}
    if not isinstance(section, dict):
        section = {}
    fallback_section = _FALLBACK_ANCHORS.get("title-slide", {})
    if isinstance(fallback_section, dict):
        fallback = str(fallback_section.get(key, fallback))
    return str(section.get(key, fallback))


def _area(
    anchors: Optional[AnchorMap],
    key: str,
    default: dict[str, float],
) -> dict[str, int]:
    """
    title: Resolve an area in inches and convert to EMU.
    parameters:
      anchors:
        type: Optional[AnchorMap]
      key:
        type: str
      default:
        type: dict[str, float]
    returns:
      type: dict[str, int]
    """
    areas = anchors.get("areas", {}) if isinstance(anchors, dict) else {}
    if not isinstance(areas, dict):
        areas = {}
    area = areas.get(key, {})
    if not isinstance(area, dict):
        area = {}

    resolved = dict(default)
    for field in ("left", "top", "width", "height", "offset-top"):
        if field in area:
            try:
                resolved[field] = float(area[field])
            except (TypeError, ValueError):
                pass

    emu: dict[str, int] = {}
    for field, value in resolved.items():
        emu[field] = Inches(value)
    return emu


def _remove_shapes_for(
    anchors: Optional[AnchorMap],
    key: str,
    default: list[str],
) -> list[str]:
    """
    title: Resolve a list of shape names to remove for a section.
    parameters:
      anchors:
        type: Optional[AnchorMap]
      key:
        type: str
      default:
        type: list[str]
    returns:
      type: list[str]
    """
    section = anchors.get("remove-shapes", {}) if isinstance(anchors, dict) else {}
    if not isinstance(section, dict):
        section = {}
    value = section.get(key, default)
    if not isinstance(value, list):
        return list(default)
    names: list[str] = []
    for item in value:
        if isinstance(item, str):
            names.append(item)
    return names if names else list(default)


def _checkpoint_textboxes(anchors: Optional[AnchorMap]) -> list[str]:
    """
    title: Resolve checkpoint textbox names.
    parameters:
      anchors:
        type: Optional[AnchorMap]
    returns:
      type: list[str]
    """
    section = anchors.get("checkpoints", {}) if isinstance(anchors, dict) else {}
    if not isinstance(section, dict):
        section = {}
    value = section.get("textbox-names", [])
    if not isinstance(value, list):
        value = []
    names = [item for item in value if isinstance(item, str)]
    if names:
        return names
    fallback = _FALLBACK_ANCHORS.get("checkpoints", {})
    if isinstance(fallback, dict):
        fb_names = fallback.get("textbox-names", [])
        if isinstance(fb_names, list):
            return [item for item in fb_names if isinstance(item, str)]
    return []


def _merge_style(
    base: Optional[StyleAttrs],
    override: Optional[StyleAttrs],
) -> StyleAttrs:
    """
    title: Merge two style dictionaries where override wins.
    parameters:
      base:
        type: Optional[StyleAttrs]
      override:
        type: Optional[StyleAttrs]
    returns:
      type: StyleAttrs
    """
    merged = dict(base or {})
    if override:
        merged.update(override)
    return merged


def _style_for(styles: Optional[StyleMap], name: str) -> StyleAttrs:
    """
    title: Resolve a system style with ``.slide`` fallback.
    parameters:
      styles:
        type: Optional[StyleMap]
      name:
        type: str
    returns:
      type: StyleAttrs
    """
    slide_style = dict((styles or {}).get(".slide", {}))
    if name == ".slide":
        return slide_style
    return _merge_style(slide_style, (styles or {}).get(name))


def _style_with_fallback(
    styles: Optional[StyleMap],
    name: str,
    fallback: Optional[StyleAttrs] = None,
) -> StyleAttrs:
    """
    title: Resolve style and apply fallback defaults.
    parameters:
      styles:
        type: Optional[StyleMap]
      name:
        type: str
      fallback:
        type: Optional[StyleAttrs]
    returns:
      type: StyleAttrs
    """
    return _merge_style(fallback, _style_for(styles, name))


def _apply_group_title_style(
    slide: Slide,
    group_name: str,
    style: StyleAttrs,
) -> None:
    """
    title: Apply style to an existing group title textbox.
    parameters:
      slide:
        type: Slide
      group_name:
        type: str
      style:
        type: StyleAttrs
    """
    if not style:
        return
    title_tb = find_group_textbox(slide, group_name)
    if title_tb is None:
        return
    current = title_tb.text_frame.text
    set_textbox_text(title_tb, current, style=style)


def slide_title(
    slide: Slide,
    title: str,
    subtitle: str,
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the title page (slide 1).
    parameters:
      slide:
        type: Slide
        description: The first slide of the presentation.
      title:
        type: str
        description: Main title text in the top-left label area.
      subtitle:
        type: str
        description: Subtitle text in the main title area.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    title_style = _style_for(styles, ".title")
    subtitle_style = _style_with_fallback(styles, ".subtitle", title_style)

    title_group = _title_slide_group(anchors, "title-group", "Group 7")
    subtitle_group = _title_slide_group(anchors, "subtitle-group", "Group 4")

    # Top-left title label
    title_tb = find_group_textbox(slide, title_group)
    if title_tb is not None:
        set_textbox_text(title_tb, title, style=title_style)

    # Main subtitle line
    subtitle_tb = find_group_textbox(slide, subtitle_group)
    if subtitle_tb is not None:
        set_textbox_text(subtitle_tb, subtitle, style=subtitle_style)

    if notes:
        set_notes(slide, notes)


def slide_objectives(
    slide: Slide,
    items: list[str],
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the Learning Objectives slide (slide 2).
    parameters:
      slide:
        type: Slide
        description: The second slide of the presentation.
      items:
        type: list[str]
        description: Learning objective bullet points.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    for name in _remove_shapes_for(anchors, "objectives", ["TextBox 6"]):
        tb = find_textbox_by_name(slide, name)
        if tb is not None:
            # Remove placeholder(s) before adding custom content.
            sp = tb._element
            sp.getparent().remove(sp)

    _apply_group_title_style(
        slide, _title_group_for(anchors, 1), _style_for(styles, ".title")
    )
    box = _area(
        anchors,
        "objectives-bullets",
        {"left": 0.9, "top": 3.6, "width": 13.0, "height": 6.5},
    )

    add_bullet_list(
        slide,
        left=box["left"],
        top=box["top"],
        width=box["width"],
        height=box["height"],
        items=items,
        style=_style_with_fallback(
            styles,
            ".slide",
            {"font-size": 30, "spacing": 14},
        ),
    )

    if notes:
        set_notes(slide, notes)


def slide_toolkit(
    slide: Slide,
    items: list[str],
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the Core Toolkit Recap slide (slide 3).
    parameters:
      slide:
        type: Slide
        description: The third slide of the presentation.
      items:
        type: list[str]
        description: Toolkit recap bullet points.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    _apply_group_title_style(
        slide, _title_group_for(anchors, 2), _style_for(styles, ".title")
    )
    box = _area(
        anchors,
        "toolkit-bullets",
        {"left": 0.9, "top": 2.6, "width": 12.0, "height": 7.5},
    )

    add_bullet_list(
        slide,
        left=box["left"],
        top=box["top"],
        width=box["width"],
        height=box["height"],
        items=items,
        style=_style_with_fallback(
            styles,
            ".slide",
            {"font-size": 30, "spacing": 14},
        ),
    )

    if notes:
        set_notes(slide, notes)


def slide_whats_new(
    slide: Slide,
    items: list[str],
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the What's New in This Lesson slide (slide 4).
    parameters:
      slide:
        type: Slide
        description: The fourth slide of the presentation.
      items:
        type: list[str]
        description: New concepts bullet points.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    _apply_group_title_style(
        slide, _title_group_for(anchors, 3), _style_for(styles, ".title")
    )
    box = _area(
        anchors,
        "whats-new-bullets",
        {"left": 0.9, "top": 2.6, "width": 12.0, "height": 7.5},
    )

    add_bullet_list(
        slide,
        left=box["left"],
        top=box["top"],
        width=box["width"],
        height=box["height"],
        items=items,
        style=_style_with_fallback(
            styles,
            ".slide",
            {"font-size": 30, "spacing": 14},
        ),
    )

    if notes:
        set_notes(slide, notes)


def slide_default(
    slide: Slide,
    title: str,
    items: list[str] | None = None,
    code: str | None = None,
    flow_boxes: list[dict] | None = None,
    callout: str | None = None,
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate a generic content slide (cloned from template 5).
    summary: |-
      Supports several content combinations:

      - **Bullets only** — pass ``items``.
      - **Bullets + code block** — pass ``items`` and ``code``.
      Bullets are placed on the left half, code on the right.
      If no ``items``, the code block spans the full width.
      - **Flow diagram** — pass ``flow_boxes`` (list of dicts
      with ``label``, ``desc``, and optional ``style``
      (or legacy ``color``) keys.
      - **Callout** — pass ``callout`` for a highlighted text
      line below other content.
    parameters:
      slide:
        type: Slide
        description: A slide cloned from template slide 5.
      title:
        type: str
        description: Slide heading text.
      items:
        type: list[str] | None
        description: Bullet point strings.
      code:
        type: str | None
        description: Source code to display in a dark code block.
      flow_boxes:
        type: list[dict] | None
        description: >-
          Flow-diagram box definitions.  Each dict must have ``label`` (str)
          and optionally ``desc`` (str) and ``style`` (dict of box style
          attributes) or ``color`` (str hex or RGBColor).
      callout:
        type: str | None
        description: A bold callout line placed below main content.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
        description: Effective style map for this slide call.
      anchors:
        type: Optional[AnchorMap]
    """
    title_style = _style_for(styles, ".title")
    slide_style = _style_for(styles, ".slide")

    title_group = _title_group_for(anchors, 4)
    # Set title text in mapped title group.
    title_tb = find_group_textbox(slide, title_group)
    if title_tb is not None:
        set_textbox_text(title_tb, title, style=title_style)

    # Track the bottom of the last placed element
    bottom: int = CONTENT_TOP

    # ── Flow diagram ────────────────────────────────────────
    if flow_boxes:
        flow_anchor = _area(
            anchors,
            "generic-flow",
            {"left": 0.9, "top": 2.6},
        )
        add_flow_boxes(
            slide,
            boxes=flow_boxes,
            left=flow_anchor["left"],
            top=flow_anchor["top"],
            style=slide_style,
        )
        bottom = CONTENT_TOP + Inches(3.0)

    # ── Bullets + optional code block ───────────────────────
    elif items and code:
        # Side-by-side layout: bullets left, code right
        bullets_box = _area(
            anchors,
            "generic-bullets-code-bullets",
            {"left": 0.9, "top": 2.6, "width": 8.5, "height": 3.5},
        )
        code_box = _area(
            anchors,
            "generic-bullets-code-code",
            {"left": 0.9, "top": 6.4, "width": 13.0, "height": 4.5},
        )
        add_bullet_list(
            slide,
            left=bullets_box["left"],
            top=bullets_box["top"],
            width=bullets_box["width"],
            height=bullets_box["height"],
            items=items,
            style=_style_with_fallback(
                styles,
                ".slide",
                {"font-size": 26, "spacing": 10},
            ),
        )
        add_code_block(
            slide,
            left=code_box["left"],
            top=code_box["top"],
            width=code_box["width"],
            height=code_box["height"],
            code_text=code,
            style=_style_with_fallback(
                styles,
                ".code",
                {
                    "bg-color": "#193952",
                    "font-color": "#FFFFFF",
                    "font-size": 20,
                    "font-name": "Consolas",
                },
            ),
        )
        bottom = code_box["top"] + code_box["height"]

    elif code:
        # Code block only — full width
        code_box = _area(
            anchors,
            "generic-code-only",
            {"left": 0.9, "top": 2.6, "width": 13.0, "height": 6.0},
        )
        add_code_block(
            slide,
            left=code_box["left"],
            top=code_box["top"],
            width=code_box["width"],
            height=code_box["height"],
            code_text=code,
            style=_style_with_fallback(
                styles,
                ".code",
                {
                    "bg-color": "#193952",
                    "font-color": "#FFFFFF",
                    "font-size": 20,
                    "font-name": "Consolas",
                },
            ),
        )
        bottom = code_box["top"] + code_box["height"]

    elif items:
        # Bullets only — full width
        box = _area(
            anchors,
            "generic-bullets-only",
            {"left": 0.9, "top": 2.6, "width": 12.5, "height": 7.5},
        )
        add_bullet_list(
            slide,
            left=box["left"],
            top=box["top"],
            width=box["width"],
            height=box["height"],
            items=items,
            style=_style_with_fallback(
                styles,
                ".slide",
                {"font-size": 30, "spacing": 14},
            ),
        )
        bottom = box["top"] + box["height"]

    # ── Callout text ────────────────────────────────────────
    if callout:
        callout_box = _area(
            anchors,
            "generic-callout",
            {"left": 0.9, "width": 12.5, "height": 0.8, "offset-top": 0.3},
        )
        add_textbox(
            slide,
            left=callout_box["left"],
            top=bottom + callout_box["offset-top"],
            width=callout_box["width"],
            height=callout_box["height"],
            text=callout,
            style=_merge_style(
                {"font-size": 26, "bold": True},
                slide_style,
            ),
        )

    if notes:
        set_notes(slide, notes)


def slide_checkpoints(
    slide: Slide,
    items: list[str],
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the Validation Checkpoints slide (n-3).
    summary: |-
      The template slide 6 has 5 pre-built checklist rows with
      TextBoxes named ``TextBox 34`` through ``TextBox 38``.
      This function fills those text boxes with the provided items.
    parameters:
      slide:
        type: Slide
        description: The Validation Checkpoints slide.
      items:
        type: list[str]
        description: Up to 5 checkpoint descriptions.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    _apply_group_title_style(
        slide, _title_group_for(anchors, 5), _style_for(styles, ".title")
    )
    checkpoint_style = _style_with_fallback(
        styles,
        ".slide",
        {"font-size": 30, "font-color": "#0B1F33", "font-name": FONT_NAME},
    )

    textbox_names = _checkpoint_textboxes(anchors)
    for i, name in enumerate(textbox_names):
        tb = find_textbox_by_name(slide, name)
        if tb is not None and i < len(items):
            set_textbox_text(
                tb,
                items[i],
                style=checkpoint_style,
            )
        elif tb is not None:
            # Clear unused rows
            set_textbox_text(tb, "")

    if notes:
        set_notes(slide, notes)


def slide_exercise(
    slide: Slide,
    items: list[str],
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the Exercise Playbook slide (n-2).
    parameters:
      slide:
        type: Slide
        description: The Exercise Playbook slide.
      items:
        type: list[str]
        description: Exercise strategy bullet points.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    _apply_group_title_style(
        slide, _title_group_for(anchors, 6), _style_for(styles, ".title")
    )
    box = _area(
        anchors,
        "exercise-bullets",
        {"left": 0.9, "top": 2.6, "width": 12.0, "height": 7.5},
    )

    add_bullet_list(
        slide,
        left=box["left"],
        top=box["top"],
        width=box["width"],
        height=box["height"],
        items=items,
        style=_style_with_fallback(
            styles,
            ".slide",
            {"font-size": 30, "spacing": 14},
        ),
    )

    if notes:
        set_notes(slide, notes)


def slide_debugging(
    slide: Slide,
    items: list[str],
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the Debugging Guide slide (n-1).
    parameters:
      slide:
        type: Slide
        description: The Debugging Guide slide.
      items:
        type: list[str]
        description: Debugging tip bullet points.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    _apply_group_title_style(
        slide, _title_group_for(anchors, 7), _style_for(styles, ".title")
    )
    box = _area(
        anchors,
        "debugging-bullets",
        {"left": 0.9, "top": 2.6, "width": 12.0, "height": 7.5},
    )

    add_bullet_list(
        slide,
        left=box["left"],
        top=box["top"],
        width=box["width"],
        height=box["height"],
        items=items,
        style=_style_with_fallback(
            styles,
            ".slide",
            {"font-size": 30, "spacing": 14},
        ),
    )

    if notes:
        set_notes(slide, notes)


def slide_recap(
    slide: Slide,
    items: list[str],
    next_topic: str = "",
    notes: str = "",
    styles: Optional[StyleMap] = None,
    anchors: Optional[AnchorMap] = None,
) -> None:
    """
    title: Populate the Recap and Next Steps slide (n, final).
    parameters:
      slide:
        type: Slide
        description: The final slide.
      items:
        type: list[str]
        description: Recap bullet points.
      next_topic:
        type: str
        description: Short description of the next lesson topic.
      notes:
        type: str
        description: Speaker notes.
      styles:
        type: Optional[StyleMap]
      anchors:
        type: Optional[AnchorMap]
    """
    _apply_group_title_style(
        slide, _title_group_for(anchors, 8), _style_for(styles, ".title")
    )
    bullets_box = _area(
        anchors,
        "recap-bullets",
        {"left": 0.9, "top": 2.6, "width": 12.0, "height": 5.0},
    )

    add_bullet_list(
        slide,
        left=bullets_box["left"],
        top=bullets_box["top"],
        width=bullets_box["width"],
        height=bullets_box["height"],
        items=items,
        style=_style_with_fallback(
            styles,
            ".slide",
            {"font-size": 30, "spacing": 14},
        ),
    )

    if next_topic:
        next_box = _area(
            anchors,
            "recap-next-topic",
            {"left": 0.9, "top": 8.5, "width": 12.0, "height": 1.5},
        )
        add_textbox(
            slide,
            left=next_box["left"],
            top=next_box["top"],
            width=next_box["width"],
            height=next_box["height"],
            text=f"COMING NEXT: {next_topic}",
            style=_style_with_fallback(
                styles,
                ".title",
                {"font-size": 24, "bold": True},
            ),
        )

    if notes:
        set_notes(slide, notes)
