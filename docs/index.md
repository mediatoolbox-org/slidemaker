# SlideMaker

`slidemaker` is a Python library for generating slide decks from PowerPoint
templates. Built on `python-pptx`, it exposes one high-level class
(`SlideBuilder`) that handles template loading, slide cloning, placeholder
replacement, content layout, style resolution, and final deck assembly.

## Architecture

- `src/slidemaker/core.py`: low-level text, shape, layout, and slide utilities.
- `src/slidemaker/cli.py`: `SlideBuilder` high-level API.
- `src/slidemaker/media.py`: image URL resolution, caching, and prompt-based
  generation helpers.

## Quick Start

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx")

# Replace {{title}} placeholder in template slide 1
sb.add_slide(
    template_page=1,
    content={"title": "Introduction", "subtitle": "Getting Started"},
)

# Create new shapes on default template page
sb.add_slide(
    content={"title": "Key Points"},
    items=["Point A", "Point B"],
)

sb.save("output.pptx")
```

See [Getting Started](getting-started.md) for setup and examples.
