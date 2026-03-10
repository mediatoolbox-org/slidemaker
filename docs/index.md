# SlideMaker

`slidemaker` is a Python library for generating slide decks from PowerPoint
templates. Built on `python-pptx`, it exposes one high-level class
(`SlideBuilder`) that handles template loading, slide cloning, content
population, style resolution, and final deck assembly.

## Architecture

- `src/slidemaker/core.py`: low-level text, shape, and slide utilities.
- `src/slidemaker/anchors.py`: anchor map loading and validation.
- `src/slidemaker/cli.py`: `SlideBuilder` high-level API.
- `src/slidemaker/main.py`: CLI entry point.

## Quick Start

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx")
sb.add_slide("Introduction", items=["Point A", "Point B"])
sb.add_slide("Code Example", code="print('hello')")
sb.save("output.pptx")
```

See [Getting Started](getting-started.md) for setup and [CLI](commands.md) for
command-line usage.
