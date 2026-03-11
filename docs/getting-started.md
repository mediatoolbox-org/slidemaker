# Getting Started

This guide sets up `slidemaker` for local development.

## Prerequisites

- Conda or Mamba
- Poetry

## Setup

```bash
git clone https://github.com/mediatoolbox-org/slidemaker.git
cd slidemaker
mamba env create --file conda/dev.yaml
conda activate slidemaker
poetry config virtualenvs.create false
poetry install --with dev
```

## Verify Installation

```bash
python -c "from slidemaker import SlideBuilder; print('OK')"
```

## Prepare Your Template

Add `{{placeholder}}` text to shapes in your PowerPoint template. For example,
put `{{title}}` in the title textbox and `{{body}}` in the content area.

## Run Example

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx", template_default_page=5)

# Replace placeholders in template slide 1
sb.add_slide(
    template_page=1,
    content={"title": "Introduction", "subtitle": "Course Overview"},
    notes="Welcome to the course.",
)

# Create new shapes on default template page
sb.add_slide(
    content={"title": "Key Concepts"},
    items=["Concept A", "Concept B"],
)

sb.save("test_deck.pptx")
```

## Local Quality Gates

```bash
makim tests.unit
makim tests.linter
makim docs.build
```
