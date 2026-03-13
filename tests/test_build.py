"""
title: Quick smoke test for SlideBuilder.
"""

import sys

sys.path.insert(0, "src")

from slidemaker import SlideBuilder  # noqa: E402

TEMPLATE = "tests/data/template.pptx"

sb = SlideBuilder(TEMPLATE, template_default_page=4)

sb.add_slide(
    content={"title": "The ETL Paradigm"},
    items=[
        "Extract: Retrieve raw data from a source",
        "Transform: Clean, filter, enrich or reshape",
        "Load: Write results to a destination",
    ],
    notes="ETL stands for Extract Transform Load.",
)

sb.add_slide(
    content={"title": "Extract: Date-Range Queries"},
    items=[
        "Convert date string to Timestamp",
        "Compute end of day with DateOffset",
        "Use $gte and $lt for half-open interval",
    ],
    notes="The extract step.",
)

sb.save("/tmp/test_slidemaker_output.pptx")
