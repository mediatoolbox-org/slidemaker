# SlideMaker

`slidemaker` is a Python API for generating branded lesson decks from a fixed
PowerPoint template. Built on `python-pptx`, it exposes one high-level class
(`SlideBuilder`) that handles template loading, slide population, style
resolution, and final deck assembly.

## Architecture

- `src/slidemaker/core.py`: low-level text, shape, and slide utilities.
- `src/slidemaker/anchors.py`: anchor map loading and validation.
- `src/slidemaker/template.py`: `SlideBuilder` high-level API.
- `src/slidemaker/cli.py`: Typer CLI layer.
- `src/slidemaker/main.py`: CLI entry point.

## Quick Start

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx")
sb.add_title("LESSON 1.1", "Topic")
sb.add_objectives(["Obj 1", "Obj 2"])
sb.add_generic("Main Concept", items=["Point A", "Point B"])
sb.add_recap(["Summary"], next_topic="Next lesson")
sb.save("lesson.pptx")
```

See [Getting Started](getting-started.md) for setup and [CLI](commands.md) for
command-line usage.
