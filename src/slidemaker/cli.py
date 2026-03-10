"""High-level API for building slide decks.

``SlideBuilder`` is the main entry point.  A user script creates
an instance, calls one method per slide in order, and finally
calls :meth:`save` to write the ``.pptx`` file.

Example
-------
::

    from slidemaker.cli import SlideBuilder

    sb = SlideBuilder("template.pptx")
    sb.add_title(title="LESSON 7.2", subtitle="ETL Pipelines")
    sb.add_objectives(["Explain ETL", "Query MongoDB"])
    sb.add_toolkit(["MongoClient", "pd.DataFrame"])
    sb.add_whats_new(["Date-range queries", "update_one"])
    sb.add_generic("The ETL Paradigm", ["Extract", "Transform"])
    sb.add_checkpoints(["After Extract", "After Load"])
    sb.add_exercise(["Work in order", "Test each step"])
    sb.add_debugging(["Empty results?", "KeyError?"])
    sb.add_recap(["ETL separates stages"], next_topic="Chi-sq")
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
from slidemaker.core import clone_slide, delete_slide, move_slide
from slidemaker.template import (
    slide_checkpoints,
    slide_debugging,
    slide_default,
    slide_exercise,
    slide_objectives,
    slide_recap,
    slide_title,
    slide_toolkit,
    slide_whats_new,
)

# Template slide indices (0-based)
_IDX_TITLE = 0
_IDX_OBJECTIVES = 1
_IDX_TOOLKIT = 2
_IDX_WHATS_NEW = 3
_IDX_DEFAULT = 4
_IDX_CHECKPOINTS = 5
_IDX_EXERCISE = 6
_IDX_DEBUGGING = 7
_IDX_RECAP = 8

_SYSTEM_STYLES = {".slide", ".title", ".subtitle", ".code"}

StyleAttrs = dict[str, Any]
StyleMap = dict[str, StyleAttrs]


class SlideBuilder:
    """Builds a slide deck from the branded template.

    Slides are added in order by calling the ``add_*`` methods.
    The four fixed-position ending slides (checkpoints, exercise,
    debugging, recap) are held until :meth:`save` is called, so
    they always appear at the end regardless of call order.

    Attributes
    ----------
    prs : Presentation
        The python-pptx ``Presentation`` being built.
    """

    def __init__(
        self,
        template: str | Path,
        style: Optional[StyleMap] = None,
        default_template_page: int = DEFAULT_TEMPLATE_PAGE,
        anchor_map: Optional[dict[str, Any] | str | Path] = None,
    ) -> None:
        """Load template, anchor map, styles, and prepare for building."""
        self._template_path = Path(template)
        self._prs = Presentation(str(self._template_path))
        self._default_count = 0
        # Deferred ending slides: stored as (func, kwargs)
        self._ending: dict[str, tuple] = {}
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
        """Generate an editable anchor map file for a template."""
        template_path = Path(template)
        anchor_map = generate_anchor_map(
            template=template_path,
            default_template_page=default_template_page,
            include_shape_catalog=include_shape_catalog,
        )
        return dump_anchor_map(anchor_map, out)

    def add_style(self, style: StyleMap) -> None:
        """Register or update named style definitions.

        System style names:

        - ``.slide``: base style for all text.
        - ``.title``: title style (falls back to ``.slide``).
        - ``.subtitle``: subtitle style (falls back to ``.title``).
        - ``.code``: code style (falls back to ``.slide``).

        Parameters
        ----------
        style : dict
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
        """Apply a named style as slide-level style overrides."""
        if style_name not in self._styles:
            raise KeyError(f"unknown style name: {style_name}")
        styles[".slide"].update(self._styles[style_name])

    def _resolve_styles(
        self,
        style: Optional[str | dict[str, Any]],
    ) -> StyleMap:
        """Resolve effective styles for one slide call."""
        styles: StyleMap = deepcopy(self._styles)

        if style is None:
            return styles

        if isinstance(style, str):
            self._apply_named_style(styles, style)
            return styles

        if not isinstance(style, dict):
            raise TypeError("style must be None, a style name, or a style dictionary")

        # Namespaced style map form:
        # {
        #   ".slide": {...},
        #   ".title": {...},
        #   ".subtitle": {...},
        #   ".code": {...},
        #   "use": "mycustomstyle"
        # }
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
                        # Inline custom style blocks are treated as
                        # slide-level overrides for this call.
                        styles[".slide"].update(attrs)
                else:
                    styles[".slide"][name] = attrs
            return styles

        # Flat attribute map applies to slide text.
        styles[".slide"].update(style)
        return styles

    # ── Fixed slides (1-4) ──────────────────────────────────

    def add_title(
        self,
        title: str,
        subtitle: str,
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for the title slide (slide 1).

        Parameters
        ----------
        title : str
            Main title text in the top-left label area.
        subtitle : str
            Subtitle text in the main title area.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        slide = self._prs.slides[_IDX_TITLE]
        slide_title(
            slide,
            title=title,
            subtitle=subtitle,
            notes=notes,
            styles=styles,
            anchors=self._anchors,
        )

    def add_objectives(
        self,
        items: list[str],
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for the Learning Objectives slide (slide 2).

        Parameters
        ----------
        items : list of str
            Learning objective bullet points.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        slide = self._prs.slides[_IDX_OBJECTIVES]
        slide_objectives(
            slide,
            items=items,
            notes=notes,
            styles=styles,
            anchors=self._anchors,
        )

    def add_toolkit(
        self,
        items: list[str],
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for the Core Toolkit Recap slide (slide 3).

        Parameters
        ----------
        items : list of str
            Toolkit recap bullet points.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        slide = self._prs.slides[_IDX_TOOLKIT]
        slide_toolkit(
            slide,
            items=items,
            notes=notes,
            styles=styles,
            anchors=self._anchors,
        )

    def add_whats_new(
        self,
        items: list[str],
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for the What's New slide (slide 4).

        Parameters
        ----------
        items : list of str
            New concepts bullet points.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        slide = self._prs.slides[_IDX_WHATS_NEW]
        slide_whats_new(
            slide,
            items=items,
            notes=notes,
            styles=styles,
            anchors=self._anchors,
        )

    # ── Generic content slides (5..n-4) ─────────────────────

    def add_generic(
        self,
        title: str,
        items: list[str] | None = None,
        code: str | None = None,
        flow_boxes: list[dict] | None = None,
        callout: str | None = None,
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Add a generic content slide.

        Clones template slide 5 and populates it with the
        given title and optional rich content: bullets, code
        blocks, flow diagrams, or callout text.

        Parameters
        ----------
        title : str
            Slide heading.
        items : list of str, optional
            Bullet point strings. Supports ``**bold**: rest``
            pattern for bold prefixes.
        code : str, optional
            Source code for a dark code block.
        flow_boxes : list of dict, optional
            Flow-diagram boxes (``label``, ``desc``, and
            optional ``style``/``color``).
        callout : str, optional
            Bold callout line below other content.
        notes : str
            Speaker notes.
        style : str or dict, optional
            Per-slide style override. Use a style name
            registered with :meth:`add_style`, a flat
            style dict for ``.slide``, or a namespaced
            dict (``.slide``, ``.title``, ``.subtitle``, ``.code``).
        """
        styles = self._resolve_styles(style)
        new_slide = clone_slide(self._prs, _IDX_DEFAULT)
        self._default_count += 1
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

    # ── Ending slides (n-3 .. n) ────────────────────────────

    def add_checkpoints(
        self,
        items: list[str],
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for Validation Checkpoints (n-3).

        Parameters
        ----------
        items : list of str
            Up to 5 checkpoint descriptions.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        self._ending["checkpoints"] = (
            slide_checkpoints,
            {
                "items": items,
                "notes": notes,
                "styles": styles,
                "anchors": self._anchors,
            },
        )

    def add_exercise(
        self,
        items: list[str],
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for Exercise Playbook (n-2).

        Parameters
        ----------
        items : list of str
            Exercise strategy bullet points.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        self._ending["exercise"] = (
            slide_exercise,
            {
                "items": items,
                "notes": notes,
                "styles": styles,
                "anchors": self._anchors,
            },
        )

    def add_debugging(
        self,
        items: list[str],
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for Debugging Guide (n-1).

        Parameters
        ----------
        items : list of str
            Debugging tip bullet points.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        self._ending["debugging"] = (
            slide_debugging,
            {
                "items": items,
                "notes": notes,
                "styles": styles,
                "anchors": self._anchors,
            },
        )

    def add_recap(
        self,
        items: list[str],
        next_topic: str = "",
        notes: str = "",
        style: Optional[str | dict[str, Any]] = None,
    ) -> None:
        """Set content for Recap and Next Steps (n).

        Parameters
        ----------
        items : list of str
            Recap bullet points.
        next_topic : str
            Short description of the next lesson topic.
        notes : str
            Speaker notes.
        """
        styles = self._resolve_styles(style)
        self._ending["recap"] = (
            slide_recap,
            {
                "items": items,
                "next_topic": next_topic,
                "notes": notes,
                "styles": styles,
                "anchors": self._anchors,
            },
        )

    # ── Build and save ──────────────────────────────────────

    def save(self, path: str) -> None:
        """Assemble the final deck and write to disk.

        This method:

        1. Applies deferred content to the ending slides
           (checkpoints, exercise, debugging, recap).
        2. Removes unused template slides (the original
           default slide 5 and any ending slides that were
           not populated).
        3. Saves the presentation to ``path``.

        Parameters
        ----------
        path : str
            Output file path for the ``.pptx`` file.
        """
        # Apply ending slide content to template slides
        ending_map = {
            "checkpoints": _IDX_CHECKPOINTS,
            "exercise": _IDX_EXERCISE,
            "debugging": _IDX_DEBUGGING,
            "recap": _IDX_RECAP,
        }
        for key, idx in ending_map.items():
            if key in self._ending:
                func, kwargs = self._ending[key]
                func(self._prs.slides[idx], **kwargs)

        # Delete unpopulated ending slides and original default
        # (in reverse order so indices stay valid)
        to_delete = [_IDX_DEFAULT]
        for key, idx in ending_map.items():
            if key not in self._ending:
                to_delete.append(idx)

        for idx in sorted(set(to_delete), reverse=True):
            delete_slide(self._prs, idx)

        # After deletion the layout is:
        #   [fixed 0-3] [endings] [cloned defaults]
        # We want: [fixed] [defaults] [endings]
        # Strategy: move each ending from position 4 to
        # the last position, repeating n_endings times.
        n_endings = len(self._ending)
        n_defaults = self._default_count
        total = len(self._prs.slides)
        if n_defaults > 0 and n_endings > 0:
            for _ in range(n_endings):
                move_slide(self._prs, 4, total - 1)

        self._prs.save(path)
        n_slides = len(self._prs.slides)
        print(f"Saved → {path}  ({n_slides} slides)")
