# SlideMaker

`slidemaker` is a Python library for generating slide decks from PowerPoint
templates. Built on `python-pptx`, it exposes one high-level class
(`SlideBuilder`) that handles template loading, slide cloning, placeholder
replacement, content layout, style resolution, and final deck assembly.

## Installation

```bash
pip install slidemaker
```

Or for development:

```bash
pip install -e .
```

Dependencies:

- Python `>=3.10`
- `python-pptx >= 1`

## Quick Start

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx", template_default_page=5)

# Replace placeholders in a template slide
sb.add_slide(
    template_page=1,
    content={
        "title": "LESSON 7.2",
        "subtitle": "ETL Pipelines for Analytics",
    },
)

# Create new shapes on a cloned slide
sb.add_slide(
    content={"title": "Extract Step"},
    items=["Use pd.to_datetime for start date", "Use $gte/$lt for interval"],
    code='query = {\n    "createdAt": {"$gte": start, "$lt": end}\n}',
)

# Bullets-only slide using default template page
sb.add_slide(
    content={"title": "Key Takeaways"},
    items=["ETL separates concerns", "Classes bundle logic"],
)

sb.save("output.pptx")
```

## How It Works

### Template placeholders

Place `{{placeholder}}` text in your PowerPoint template slides. When
`add_slide` clones a template page, shapes containing `{{key}}` are found and
replaced with the matching value from the `content` dict.

Placeholder matching is **case-insensitive**: `{{TITLE}}`, `{{Title}}`, and
`{{title}}` all match `content={"title": "..."}`.

### Content values

- `str` — replaces the shape text.
- `list[str]` — replaces the shape with a bullet list at the same position and
  size.
- `None` — clears the shape text.

### New shapes

Pass `items`, `markdown`, `code`, `table`, `image`, `flow_boxes`, or `callout`
to create new shapes on top of the cloned slide. These are laid out
automatically:

| Combination          | Layout                            |
| -------------------- | --------------------------------- |
| `flow_boxes`         | Flow diagram at content top       |
| `items` + `code`     | Bullets on top, code block below  |
| `markdown` + `code`  | Markdown on top, code block below |
| `items` + `table`    | Bullets on top, table below       |
| `markdown` + `table` | Markdown on top, table below      |
| `code` + `table`     | Code block on top, table below    |
| `items` + `image`    | Bullets on top, image below       |
| `markdown` + `image` | Markdown on top, image below      |
| `code` + `image`     | Code block on top, image below    |
| `items` only         | Full content area                 |
| `markdown` only      | Full content area                 |
| `code` only          | Full content area                 |
| `table` only         | Full content area                 |
| `image` only         | Full content area                 |
| `callout`            | Placed below other content        |

Both modes (placeholder replacement and new shapes) can be used together on the
same slide.

`table` can be combined with either `items` or `code`, but not both, and it
cannot be combined with `flow_boxes`. `image` can be combined with either
`items` or `code`, but not both, and it cannot be combined with `table` or
`flow_boxes`. `markdown` can be combined with `code`, `table`, or `image`, but
not with `items` or `flow_boxes`.

## SlideBuilder API

```python
from slidemaker import SlideBuilder
```

### Constructor

```python
SlideBuilder(
    template: str | Path,
    style: dict[str, dict] | None = None,
    template_default_page: int = 5,
)
```

- `template`: path to the `.pptx` template (required).
- `style`: global style definitions.
- `template_default_page`: 1-based index of the default template slide to clone
  when `template_page` is not specified.

### Adding slides

```python
sb.add_slide(
    content={"title": "Slide Title", "body": ["Point A", "Point B"]},
    items=None,
    markdown=None,
    code=None,
    table=None,
    image=None,
    flow_boxes=None,
    callout=None,
    notes="",
    style=None,
    template_page=None,
)
```

- `content`: dict of placeholder replacements (`{{key}}` in template).
- `items`: bullet points (creates new shapes).
- `markdown`: free-form markdown block (creates new shape).
- `code`: source code block (creates new shape).
- `table`: generated table definition (creates new shape).
- `image`: image path or image spec (creates new shape).
- `flow_boxes`: flow diagram boxes (creates new shapes).
- `callout`: bold callout text below other content.
- `notes`: speaker notes.
- `style`: per-slide style override (string name, flat dict, or structured
  dict).
- `template_page`: 1-based template page to clone (defaults to
  `template_default_page`).

### Style registration

```python
sb.add_style({
    "dense": {"font-size": 26, "spacing": 10},
    "highlight": {"font-color": "#193952"},
    "#title": {"font-size": 51, "bold": True},
})
```

### Save

```python
sb.save("output.pptx")
```

Removes all original template slides and writes the final deck.

## Style System

### Style keys

- `.slide` — base style for all new text shapes and placeholder fallback.
- `.code` — code block style.
- `.table` — base style for generated tables.
- `.table-header` — header-row overrides for generated tables.
- `.table-cell` — body-cell overrides for generated tables.
- `#placeholder` — style for a specific `{{placeholder}}`. Falls back to
  `.slide`.
- Any other name — a named preset (e.g. `"dense"`) that can be applied via
  `style="dense"` or `style={"use": "dense"}`.

### Where styles can be provided

1. **Global**: `SlideBuilder(style={...})`
2. **Named presets**: `add_style({"dense": {...}})`, then use by name.
3. **Per-slide**: `style=...` in `add_slide`.

### Per-slide `style` forms

```python
# Named preset
style="dense"

# Flat dict (applies to .slide)
style={"font-size": 26, "font-color": "#193952"}

# Structured dict with placeholder overrides
style={
    "use": "dense",
    ".slide": {"font-size": 24},
    ".code": {"line-numbers": True},
    ".table": {"font-size": 20, "padding": "6pt"},
    ".table-header": {"fill-color": "#193952", "font-color": "#FFFFFF"},
    ".table-cell": {"font-color": "#0B1F33"},
    "#title": {"font-size": 48, "letter-spacing": 90},
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

These keys also apply to bullet paragraphs inside `markdown` blocks.

### Code block keys

- `bg-color` (or `fill-color`)
- `line-numbers`: bool

### Table keys

- `columns`: optional header-row cell values
- `rows`: body rows as `list[list[str | None]]`
- `column_widths`: optional column widths; numeric values are inches
- `row_heights`: optional row heights; numeric values are inches
- `banded_rows`: bool
- `style`: per-table override merged into `.table`
- `header_style`: per-table override merged into `.table-header`
- `cell_style`: per-table override merged into `.table-cell`

### Image keys

- `path` / `src`: image file path
- `caption`: optional caption below the image
- `fit`: `"contain"` (default) or `"stretch"`
- `caption_style`: optional caption text override merged into `.slide`

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

## Table Format

```python
table={
    "columns": ["Field", "Type", "Notes"],
    "rows": [
        ["_id", "ObjectId", "Primary key"],
        ["createdAt", "datetime", "UTC timestamp"],
        ["tags", "list[str]", "Optional labels"],
    ],
    "column_widths": [2.2, 2.0, 5.8],
    "banded_rows": True,
}
```

## Image Format

```python
image={
    "path": "artifacts/confusion_matrix.png",
    "caption": "Validation confusion matrix",
    "fit": "contain",
}
```

## Markdown Format

```python
markdown = """# Why This Matters

This pipeline keeps **data movement** and *transformation* clear.
- Reproducible steps
- Easier debugging
"""
```

Supported markdown subset:

- paragraphs separated by blank lines
- `#`, `##`, `###` headings
- unordered bullets with `-` or `*` rendered as real PowerPoint bullets
- nested bullets using two leading spaces per level
- inline `**bold**`, `*italic*`, and `` `code` ``

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

- text: `add_textbox`, `set_textbox_text`, `add_bullet_list`,
  `add_markdown_textbox`, `add_code_block`
- media: `add_image`
- tables: `add_table`
- layout: `layout_content_shapes`
- placeholders: `replace_placeholders`
- shapes/flow: `add_shape_rect`, `add_flow_boxes`
- notes: `set_notes`
- slide structure: `clone_slide`, `delete_slide`, `move_slide`
- shape lookup: `find_group_textbox`, `find_textbox_by_name`
