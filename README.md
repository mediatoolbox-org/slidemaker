# SlideMaker

`slidemaker` is a Python library for generating slide decks from PowerPoint
templates. Built on `python-pptx`, it exposes one high-level class
(`SlideBuilder`) that handles template loading, slide cloning, content
population, style resolution, and final deck assembly.

## Installation

```bash
poetry install
```

Or with pip:

```bash
pip install -e .
```

Dependencies:

- Python `>=3.10`
- `python-pptx >= 1`

## Quick Start

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx")

sb.add_slide(
    title="Introduction",
    items=["Welcome to the course", "Today we cover ETL"],
)

sb.add_slide(
    title="Extract Step",
    items=["Use pd.to_datetime for start date", "Use $gte/$lt for interval"],
    code='''query = {
    "createdAt": {"$gte": start, "$lt": end}
}''',
)

sb.save("output.pptx")
```

## SlideBuilder API

```python
from slidemaker import SlideBuilder
```

### Constructor

```python
SlideBuilder(
    template: str | Path,
    style: dict[str, dict] | None = None,
    anchor_map: dict | str | Path | None = None,
)
```

- `template`: path to the `.pptx` template (required).
- `style`: global style definitions.
- `anchor_map`: optional anchor map for content placement.

### Adding slides

```python
sb.add_slide(
    title="Slide Title",
    items=None,
    code=None,
    flow_boxes=None,
    callout=None,
    notes="",
    style=None,
)
```

Each `add_slide` call clones a template slide and populates it with the given
content. Content branching:

1. `flow_boxes` present -> renders flow diagram
2. `items` + `code` -> bullets + code block
3. `code` only -> full-width code block
4. `items` only -> full-width bullets

`callout` can be added in any branch and is rendered below the main content.

### Style registration

```python
sb.add_style({
    "dense": {"font-size": 26, "spacing": 10},
    "highlight": {"font-color": "#193952"},
})
```

### Save

```python
sb.save("output.pptx")
```

## Style System

Style resolution happens per slide call:

- `.slide`: base/default text style
- `.title`: title style (falls back to `.slide`)
- `.subtitle`: subtitle style (falls back to `.title`)
- `.code`: code style (falls back to `.slide`)

### Where styles can be provided

1. Global: `SlideBuilder(style=...)`
2. Named classes: `add_style({"dense": {...}})` and use by name
3. Per-slide: `style=...` in `add_slide`

### Per-slide `style` forms

```python
# Named style class
style="dense"

# Flat dict (applies to .slide)
style={"font-size": 26, "font-color": "#193952"}

# Namespaced dict
style={
    "use": "dense",
    ".slide": {"font-size": 24},
    ".title": {"font-size": 48},
    ".code": {"line-numbers": True},
}
```

## Supported Style Attributes

### Text style keys

- `font-name`: string
- `font-size`: number (points) or string like `"24pt"`
- `font-color`: `"#RRGGBB"`, `(r, g, b)`, or `RGBColor`
- `bold`: bool
- `italic`: bool
- `alignment` / `align`: `left|center|right|justify`
- `uppercase`: bool
- `line-spacing` / `line-height`: multiplier, `"1.2x"`, `"120%"`, or `"36pt"`
- `letter-spacing`: tracking units or `"0.9pt"`
- `padding`, `padding-x`, `padding-y`, `padding-left`, `padding-right`,
  `padding-top`, `padding-bottom`

### Bullet list keys

- `spacing` / `space-after`
- `space-before`
- `bullet-char`
- `bold-prefixes`

### Code block keys

- `bg-color` (or `fill-color`)
- `line-numbers`: bool

### Shape keys (flow boxes)

- `fill-color`
- `line-color`
- `line-width`

## Flow Box Format

```python
flow_boxes=[
    {
        "label": "EXTRACT",
        "desc": "Get raw data",
        "style": {
            "fill-color": "#2E86AB",
            "font-color": "#FFFFFF",
        },
    },
]
```

## Inline Markdown in Bullets

Bullet items support inline bold with `**...**`:

```python
items=[
    "Use **len(obs) // 2** for midpoint",
    "**First half** -> control",
]
```

## Speaker Notes

Every `add_slide` call accepts `notes="..."` for speaker notes.

## Low-Level Core Utilities

`slidemaker.core` provides lower-level helpers:

- text: `add_textbox`, `set_textbox_text`, `add_bullet_list`, `add_code_block`
- shapes/flow: `add_shape_rect`, `add_flow_boxes`
- notes: `set_notes`
- slide structure: `clone_slide`, `delete_slide`, `move_slide`
- shape lookup: `find_group_textbox`, `find_textbox_by_name`
