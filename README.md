# SlideMaker

`slidemaker` is a Python API for generating branded lesson decks from a fixed
PowerPoint template.

It is built on top of `python-pptx` and exposes one high-level class,
`SlideBuilder`, that handles:

- template loading
- slide population
- style resolution
- ending-slide placement
- final deck assembly and save

This package is opinionated: it targets a specific lesson format and its
predefined slide layouts.

## Table of Contents

- [Installation](#installation)
- [Quick Start](#quick-start)
- [How Deck Assembly Works](#how-deck-assembly-works)
- [SlideBuilder API](#slidebuilder-api)
- [Anchor Maps](#anchor-maps)
- [Style System](#style-system)
- [Supported Style Attributes](#supported-style-attributes)
- [Slide Content Options](#slide-content-options)
- [Flow Box Format](#flow-box-format)
- [Speaker Notes](#speaker-notes)
- [Troubleshooting and Gotchas](#troubleshooting-and-gotchas)
- [Advanced: Low-Level Core Utilities](#advanced-low-level-core-utilities)

## Installation

From the `slidemaker/` directory:

```bash
poetry install
```

Or with pip editable install:

```bash
pip install -e .
```

Dependencies:

- Python `>=3.10`
- `python-pptx >= 1`

## Quick Start

```python
from slidemaker import SlideBuilder

sb = SlideBuilder(
    template="my_template.pptx",
    anchor_map="template_anchor_map.yaml",       # optional
    default_template_page=5,                     # optional
    style={
        ".slide": {
            "font-name": "Montserrat",
            "font-size": 24,
            "font-color": "#193952",
            "line-spacing": 1.2,
        },
        ".title": {
            "font-size": 51,
            "bold": True,
            "uppercase": True,
        },
        ".subtitle": {
            "font-size": 41,
        },
        ".code": {
            "font-name": "Consolas",
            "font-size": 20,
            "bg-color": "#193952",
            "font-color": "#FFFFFF",
            "line-numbers": True,
        },
    }
)

sb.add_style({
    "dense": {"font-size": 26, "spacing": 10},
    "warm": {"font-color": "#E86F51"},
})

sb.add_title(
    title="LESSON 7.2",
    subtitle="ETL Pipelines for Analytics",
    notes="Intro narration...",
    style={
        ".title": {"letter-spacing": 90},
        ".subtitle": {"letter-spacing": 90},
    },
)

sb.add_objectives(
    items=[
        "Explain ETL",
        "Write date-range queries",
        "Update MongoDB documents",
    ],
    style="dense",
)

sb.add_generic(
    title="Extract Step",
    items=[
        "Use **pd.to_datetime** for start date",
        "Use **$gte/$lt** for half-open interval",
    ],
    code='''query = {
    "createdAt": {"$gte": start, "$lt": end}
}''',
)

sb.add_checkpoints(["After extract", "After transform", "After load"])
sb.add_recap(["ETL = Extract, Transform, Load"], next_topic="Hypothesis Testing")

sb.save("lesson_7_2.pptx")
```

## How Deck Assembly Works

The template has 9 slide prototypes. `SlideBuilder` maps them as:

- Template 1 -> Title (fixed)
- Template 2 -> Objectives (fixed)
- Template 3 -> Toolkit (fixed)
- Template 4 -> What's New (fixed)
- Template 5 -> Generic content slide (cloned for each `add_generic` call)
- Template 6 -> Checkpoints (ending)
- Template 7 -> Exercise (ending)
- Template 8 -> Debugging (ending)
- Template 9 -> Recap (ending)

Assembly behavior at `save()`:

1. Applies deferred ending slide content (`checkpoints`, `exercise`,
   `debugging`, `recap`) if provided.
2. Removes unused template slides:
   - always removes the original template generic slide
   - removes ending slides that were never populated
3. Reorders slides so final sequence is:
   - fixed slides 1-4
   - all cloned generic slides (in call order)
   - populated ending slides (in ending order)

Important:

- If you never call an ending method, that ending slide is deleted.
- If you never call `add_generic`, no generic slides are inserted.

## SlideBuilder API

`SlideBuilder` is the only public class exported by the package:

```python
from slidemaker import SlideBuilder
```

### Constructor

```python
SlideBuilder(
    template: str | Path,
    style: dict[str, dict] | None = None,
    default_template_page: int = 5,
    anchor_map: dict | str | Path | None = None,
)
```

- Loads the template presentation from the given path.
- Loads `anchor_map` if provided, otherwise built-in anchors.
- Initializes global styles (`.slide`, `.title`, `.subtitle`, `.code`).
- If `style` is passed, it is merged into those global styles.

### Generate Anchor Map File

```python
SlideBuilder.generate_anchor_map_file(
    out="template_anchor_map.yaml",
    template="my_template.pptx",
    default_template_page=5,
    include_shape_catalog=True,
)
```

Or from CLI:

```bash
python -m slidemaker.main generate-anchor-map \
  --template my_template.pptx \
  --out template_anchor_map.yaml
```

### Style registration

```python
add_style(style: dict[str, dict]) -> None
```

Registers named style classes.

Example:

```python
sb.add_style({
    "dense": {"font-size": 26, "spacing": 10},
    "highlight": {"font-color": "#193952"},
})
```

### Fixed slides

```python
add_title(title, subtitle, notes="", style=None)
add_objectives(items, notes="", style=None)
add_toolkit(items, notes="", style=None)
add_whats_new(items, notes="", style=None)
```

### Generic content slides

```python
add_generic(
    title,
    items=None,
    code=None,
    flow_boxes=None,
    callout=None,
    notes="",
    style=None,
)
```

### Ending slides (deferred until `save`)

```python
add_checkpoints(items, notes="", style=None)
add_exercise(items, notes="", style=None)
add_debugging(items, notes="", style=None)
add_recap(items, next_topic="", notes="", style=None)
```

### Save

```python
save(path: str) -> None
```

Writes the final `.pptx` and prints slide count.

## Anchor Maps

Anchor maps define **where** SlideBuilder writes content in a template. This
removes hardcoded shape dependencies from code and makes templates editable
through config.

Typical file: `template_anchor_map.yaml`:

```yaml
version: 1
default-template-page: 5
title-slide:
  title-group: Group 7
  subtitle-group: Group 4
title-groups:
  1: Group 7
  2: Group 3
  3: Group 4
checkpoints:
  textbox-names:
    - TextBox 34
    - TextBox 35
areas:
  objectives-bullets:
    left: 0.9
    top: 3.6
    width: 13.0
    height: 6.5
```

Key points:

- `title-slide` maps groups used by `add_title`.
- `title-groups` maps title group per template page.
- `areas` are inches-based boxes for content placement.
- `checkpoints.textbox-names` maps the prebuilt checklist rows.
- `remove-shapes.objectives` lists placeholder shapes to remove first.

## Style System

Style resolution happens per slide call, with this model:

- `.slide`: base/default text style
- `.title`: title style (falls back to `.slide`)
- `.subtitle`: subtitle style (falls back to `.title` for title slide)
- `.code`: code style (falls back to `.slide`)

### Where styles can be provided

1. Global: constructor `SlideBuilder(style=...)`
2. Named classes: `add_style({"dense": {...}})` and use by name in slide calls
3. Per-slide: `style=...` in each `add_*` method

### Per-slide `style` accepted forms

#### 1) Named style class

```python
style="dense"
```

Applies class attributes to `.slide`.

#### 2) Flat dict

```python
style={"font-size": 26, "font-color": "#193952"}
```

Applies directly to `.slide`.

#### 3) Namespaced dict

```python
style={
    "use": "dense",           # or ["dense", "warm"]
    ".slide": {"font-size": 24},
    ".title": {"font-size": 48},
    ".subtitle": {"uppercase": True},
    ".code": {"line-numbers": True},
}
```

Notes:

- In namespaced mode, **system keys must include the leading dot** (`.title`,
  not `title`).
- Custom keys in this form are treated as `.slide` overrides.

### Key normalization

Style keys are normalized to lowercase kebab-case:

- `font_size` -> `font-size`
- `text_transform` -> `text-transform`

## Supported Style Attributes

These keys are supported by current text/shape utilities.

### Common text style keys

Use these in `.slide`, `.title`, `.subtitle`, `.code`, or component-specific
styles.

- `font-name`: string
- `font-size`: number (points) or string like `"24pt"`
- `font-color`: `"#RRGGBB"`, `(r, g, b)`, or `RGBColor`
- `bold`: bool
- `italic`: bool
- `alignment` / `align`: `left|center|right|justify`
- `uppercase`: bool
- `text-transform`: `"uppercase"|"none"|"normal"|"initial"`
- `padding`: number (pt) or `"Npt"`
- `padding-x`, `padding-y`, `padding-left`, `padding-right`, `padding-top`,
  `padding-bottom`

### Spacing and tracking keys

- `line-spacing` / `line-height`:
  - numeric values are treated as multipliers and converted to point leading
    based on font size
  - accepts forms like `1.2`, `"1.2x"`, `"120%"`, or absolute `"36pt"`
- `letter-spacing`:
  - numeric values are treated as tracking units relative to font size
    (Canva-friendly behavior)
  - absolute values can be passed as points using `"0.9pt"`

### Bullet list-specific keys

- `spacing` / `space-after`
- `space-before`
- `bullet-char` (real paragraph bullet character)
- `bold-prefixes` (controls inline markdown bold parsing)

### Code block-specific keys

- `bg-color` (or `fill-color`)
- `line-numbers`: bool

When `line-numbers` is true, code lines are auto-prefixed unless already
numbered.

### Shape keys (`add_shape_rect` / flow boxes)

- `fill-color`
- `line-color`
- `line-width`

## Slide Content Options

### `add_generic` content branching

`add_generic` uses this priority:

1. `flow_boxes` present -> renders flow diagram branch
2. `items` + `code` -> split/stacked bullets + code layout
3. `code` only -> full-width code block
4. `items` only -> full-width bullets

`callout` can be added in any branch and is rendered below the main content.

### Inline markdown in bullets

Bullet items support inline bold with `**...**`, for example:

```python
items=[
    "Use **len(obs) // 2** for midpoint",
    "**First half** -> control",
]
```

These are rendered as real bullet paragraphs (not simulated with a text prefix).

## Flow Box Format

`flow_boxes` is a list of dictionaries:

```python
flow_boxes=[
    {
        "label": "EXTRACT",         # required
        "desc": "Get raw data",     # optional
        "style": {
            "fill-color": "#2E86AB",
            "font-color": "#FFFFFF",
            "arrow-color": "#193952",
            "arrow-font-size": 44,
        },
        # "color": "#2E86AB"        # legacy fill-color alias
    },
]
```

Defaults:

- auto box width from available content area
- horizontal arrow between boxes
- center-aligned label and description

## Speaker Notes

Every `add_*` method accepts `notes="..."`.

Notes are written into the slide speaker-notes pane via `python-pptx` notes
APIs.

## Troubleshooting and Gotchas

- Use `.title`, `.subtitle`, `.slide`, `.code` with leading dots in namespaced
  style dicts.
- Named styles from `add_style()` are applied to `.slide` when used via
  `style="name"`.
- `add_checkpoints` fills at most 5 pre-existing rows from the template.
- If text styling appears off in Canva, prefer explicit values for
  `line-spacing` and `letter-spacing` and test with a small sample deck.
- `uppercase` is style-based. Apply in a style class/system style, for example:

```python
style={
    ".title": {"uppercase": True}
}
```

## Advanced: Low-Level Core Utilities

`slidemaker.core` also provides lower-level helpers if you need direct control:

- text: `add_textbox`, `set_textbox_text`, `add_bullet_list`, `add_code_block`
- shapes/flow: `add_shape_rect`, `add_flow_boxes`
- notes: `set_notes`
- slide structure: `clone_slide`, `delete_slide`, `move_slide`
- shape lookup: `find_group_textbox`, `find_textbox_by_name`

These are stable enough for internal use, but `SlideBuilder` should be the
default entry point.

## Minimal Full Deck Skeleton

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx")

sb.add_title("LESSON 1.1", "Topic")
sb.add_objectives(["Obj 1", "Obj 2"])
sb.add_toolkit(["Tool A", "Tool B"])
sb.add_whats_new(["New 1", "New 2"])

sb.add_generic("Main Concept", items=["Point A", "Point B"])

sb.add_checkpoints(["Check 1", "Check 2"])
sb.add_exercise(["Do X", "Do Y"])
sb.add_debugging(["If A fails...", "If B fails..."])
sb.add_recap(["Summary A", "Summary B"], next_topic="Next lesson")

sb.save("lesson.pptx")
```
