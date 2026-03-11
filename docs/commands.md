# CLI

## `slidemaker generate-anchor-map`

Generate an anchor map YAML file from a PowerPoint template. The anchor map
defines where SlideBuilder writes content in each slide layout.

```bash
python -m slidemaker.main generate-anchor-map \
  --template my_template.pptx \
  --out template_anchor_map.yaml
```

Options:

- `--template`: path to the `.pptx` template (required)
- `--out`: output YAML path (default `template_anchor_map.yaml`)
- `--template-default-page`: default template page index (default `5`)
- `--no-shape-catalog`: omit the shape catalog from output
