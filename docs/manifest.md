# Anchor Maps

Anchor maps define **where** SlideBuilder writes content in a PowerPoint
template. They remove hardcoded shape dependencies from code and make templates
editable through config.

## File Format

Typical file: `template_anchor_map.yaml`

```yaml
version: 1
template-default-page: 5
```

## Generating an Anchor Map

```bash
python -m slidemaker.main generate-anchor-map \
  --template my_template.pptx \
  --out template_anchor_map.yaml
```

This inspects every slide in the template and writes a shape catalog with names,
types, and text previews. Use this to identify shapes for your placeholders.

## Template Placeholders

The primary way to bind content to template shapes is through `{{placeholder}}`
text. Place `{{title}}`, `{{body}}`, `{{subtitle}}`, etc. directly in the
template's shape text. SlideBuilder scans all shapes (including inside Group
shapes) and replaces matches from the `content` dict.

Matching is case-insensitive: `{{TITLE}}` matches `content={"title": "..."}`.

## Using with SlideBuilder

```python
sb = SlideBuilder(
    "my_template.pptx",
    template_default_page=5,
    anchor_map="template_anchor_map.yaml",  # optional
)
```
