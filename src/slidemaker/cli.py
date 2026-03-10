"""
title: High-level API for building slide decks.
summary: |-
  ``SlideBuilder`` is the main entry point.  A user script creates
  an instance, calls ``add_slide`` to clone and populate slides, and
  finally calls :meth:`save` to write the ``.pptx`` file.

  Example
  -------
  ::

  from slidemaker.cli import SlideBuilder

  sb = SlideBuilder("template.pptx")
  sb.add_slide("The ETL Paradigm", items=["Extract", "Transform", "Load"])
  sb.add_slide("Code Example", code="print('hello')")
  sb.save("output.pptx")
"""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any, Optional

from pptx import Presentation

from slidemaker.anchors import (
    DEFAULT_TEMPLATE_PAGE,
    default_anchor_map,
    dump_anchor_map,
    generate_anchor_map,
    load_anchor_map,
)
from slidemaker.core import clone_slide, delete_slide
from slidemaker.template import slide_default

_SYSTEM_STYLES = {".slide", ".title", ".subtitle", ".code"}

StyleAttrs = dict[str, Any]
StyleMap = dict[str, StyleAttrs]


class SlideBuilder:
    """
    title: Builds a slide deck from a PowerPoint template.
    summary: |-
      Each ``add_slide`` call clones the template's generic slide and
      populates it with content.  Call :meth:`save` to assemble the
      final deck.
    attributes:
      _template_path:
        description: Path to the template file.
      _prs:
        description: The python-pptx Presentation being built.
      _default_template_page:
        description: 1-based index of the template slide to clone.
      _slide_count:
        description: Number of content slides added so far.
      _anchors:
        type: dict[str, Any]
        description: Anchor map for shape placement.
      _styles:
        type: StyleMap
        description: Registered named styles.
    """

    def __init__(
        self,
        template: str | Path,
        style: Optional[StyleMap] = None,
        default_template_page: int = DEFAULT_TEMPLATE_PAGE,
        anchor_map: Optional[dict[str, Any] | str | Path] = None,
    ) -> None:
        """
        title: Load template, anchor map, styles, and prepare for building.
        parameters:
          template:
            type: str | Path
          style:
            type: Optional[StyleMap]
          default_template_page:
            type: int
          anchor_map:
            type: Optional[dict[str, Any] | str | Path]
        """
        self._template_path = Path(template)
        self._prs = Presentation(str(self._template_path))
        self._default_template_page = default_template_page
        self._slide_count = 0
        loaded_anchor_map = load_anchor_map(anchor_map)
        self._anchors: dict[str, Any] = (
            loaded_anchor_map
            if loaded_anchor_map is not None
            else default_anchor_map(default_template_page)
        )
        self._styles: StyleMap = {
            ".slide": {},
            ".title": {},
            ".subtitle": {},
            ".code": {},
        }
        if style:
            self.add_style(style)

    @staticmethod
    def generate_anchor_map_file(
        out: str | Path,
        template: str | Path,
        default_template_page: int = DEFAULT_TEMPLATE_PAGE,
        include_shape_catalog: bool = True,
    ) -> Path:
        """
        title: Generate an editable anchor map file for a template.
        parameters:
          out:
            type: str | Path
          template:
            type: str | Path
          default_template_page:
            type: int
          include_shape_catalog:
            type: bool
        returns:
          type: Path
        """
        template_path = Path(template)
        anchor_map = generate_anchor_map(
            template=template_path,
            default_template_page=default_template_page,
            include_shape_catalog=include_shape_catalog,
        )
        return dump_anchor_map(anchor_map, out)

    def add_style(self, style: StyleMap) -> None:
        """
        title: Register or update named style definitions.
        summary: |-
          System style names:

          - ``.slide``: base style for all text.
          - ``.title``: title style (falls back to ``.slide``).
          - ``.subtitle``: subtitle style (falls back to ``.title``).
          - ``.code``: code style (falls back to ``.slide``).
        parameters:
          style:
            type: StyleMap
            description: >-
              Mapping of style names to style attribute dictionaries.
        """
        for name, attrs in style.items():
            if not isinstance(name, str):
                raise TypeError("style name keys must be strings")
            if not isinstance(attrs, dict):
                raise TypeError(
                    f"style '{name}' must map to a dictionary of attributes"
                )
            if name not in self._styles:
                self._styles[name] = {}
            self._styles[name].update(attrs)

    def _apply_named_style(
        self,
        styles: StyleMap,
        style_name: str,
    ) -> None:
        """
        title: Apply a named style as slide-level style overrides.
        parameters:
          styles:
            type: StyleMap
          style_name:
            type: str
        """
        if style_name not in self._styles:
            raise KeyError(f"unknown style name: {style_name}")
        styles[".slide"].update(self._styles[style_name])

    def _resolve_styles(
        self,
        style: Optional[str | dict[str, Any]],
    ) -> StyleMap:
        """
        title: Resolve effective styles for one slide call.
        parameters:
          style:
            type: Optional[str | dict[str, Any]]
        returns:
          type: StyleMap
        """
        styles: StyleMap = deepcopy(self._styles)

        if style is None:
            return styles

        if isinstance(style, str):
            self._apply_named_style(styles, style)
            return styles

        if not isinstance(style, dict):
            raise TypeError("style must be None, a style name, or a style dictionary")

        if any(isinstance(value, dict) for value in style.values()):
            use_names = style.get("use")
            if isinstance(use_names, str):
                self._apply_named_style(styles, use_names)
            elif isinstance(use_names, list):
                for name in use_names:
                    if isinstance(name, str):
                        self._apply_named_style(styles, name)

            for name, attrs in style.items():
                if name == "use":
                    continue
                if not isinstance(name, str):
                    continue
                if isinstance(attrs, dict):
                    if name in _SYSTEM_STYLES:
                        styles.setdefault(name, {}).update(attrs)
                    else:
                        styles[".slide"].update(attrs)
                else:
                    styles[".slide"][name] = attrs
            return styles

        styles[".slide"].update(style)
        return styles

    # ── Public API ────────────────────────────────────────────

    def add_slide(
        self,
        title: str,
        items: list[str] | None = None,
        code: str | None = None,
        flow_boxes: list[dict] | None = None,
        callout: str | None = None,
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """
        title: Add a content slide by cloning the template's generic page.
        parameters:
          title:
            type: str
            description: Slide heading.
          items:
            type: list[str] | None
            description: Bullet point strings.
          code:
            type: str | None
            description: Source code for a dark code block.
          flow_boxes:
            type: list[dict] | None
            description: Flow-diagram boxes (label, desc, optional style).
          callout:
            type: str | None
            description: Bold callout line below other content.
          notes:
            type: str
            description: Speaker notes.
          style:
            type: Optional[str | dict[str, Any]]
            description: Per-slide style override.
        """
        styles = self._resolve_styles(style)
        # Clone the generic template slide (0-based index)
        idx = self._default_template_page - 1
        new_slide = clone_slide(self._prs, idx)
        self._slide_count += 1
        slide_default(
            new_slide,
            title=title,
            items=items,
            code=code,
            flow_boxes=flow_boxes,
            callout=callout,
            notes=notes,
            styles=styles,
            anchors=self._anchors,
        )

    # ── Build and save ──────────────────────────────────────

    def save(self, path: str) -> None:
        """
        title: Assemble the final deck and write to disk.
        summary: Removes the original generic template slide and saves.
        parameters:
          path:
            type: str
            description: Output file path for the .pptx file.
        """
        # Remove the original template slide used for cloning
        idx = self._default_template_page - 1
        delete_slide(self._prs, idx)

        self._prs.save(path)
        n_slides = len(self._prs.slides)
        print(f"Saved → {path}  ({n_slides} slides)")
