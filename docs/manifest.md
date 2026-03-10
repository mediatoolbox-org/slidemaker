# Anchor Maps

Anchor maps define **where** SlideBuilder writes content in a PowerPoint
template. They remove hardcoded shape dependencies from code and make templates
editable through config.

## File Format

Typical file: `template_anchor_map.yaml`

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

## Key Sections

- `title-slide`: maps groups used by `add_title`.
- `title-groups`: maps title group per template page.
- `areas`: inches-based boxes for content placement.
- `checkpoints.textbox-names`: maps the prebuilt checklist rows.
- `remove-shapes`: lists placeholder shapes to remove before populating.

## Generating an Anchor Map

```bash
python -m slidemaker.main generate-anchor-map \
  --template my_template.pptx \
  --out template_anchor_map.yaml \
  --include-shape-catalog
```

This inspects every slide layout in the template and writes shape names, groups,
and suggested area coordinates.
