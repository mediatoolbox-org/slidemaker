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

## Run Example

```python
from slidemaker import SlideBuilder

sb = SlideBuilder("my_template.pptx")
sb.add_title("LESSON 1.1", "Topic")
sb.add_objectives(["Obj 1", "Obj 2"])
sb.add_generic("Main Concept", items=["Point A", "Point B"])
sb.save("test_deck.pptx")
```

## Local Quality Gates

```bash
makim tests.unit
makim tests.linter
makim docs.build
```
