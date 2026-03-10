"""Quick smoke test for SlideBuilder."""

import sys

sys.path.insert(0, "src")

from slidemaker import SlideBuilder  # noqa: E402

TEMPLATE = "tests/data/template.pptx"

sb = SlideBuilder(TEMPLATE)

sb.add_title(
    title="LESSON 7.2",
    subtitle="ETL Pipelines for Analytics",
    notes="Welcome to lesson 7.2.",
)

sb.add_objectives(
    items=[
        "Explain the ETL paradigm and why it matters",
        "Query MongoDB with date-based filters",
        "Transform documents by assigning groups",
        "Update documents back into MongoDB",
        "Define a Python class with __init__",
        "Build a reusable MongoRepository class",
    ],
    notes="Here is what you will be able to do.",
)

sb.add_toolkit(
    items=[
        "MongoClient connects to MongoDB",
        "collection.find(query) returns a cursor",
        "collection.aggregate([...]) runs pipelines",
        "pd.DataFrame(list_of_dicts) converts documents",
    ],
    notes="Recap of tools you already know.",
)

sb.add_whats_new(
    items=[
        "ETL pattern: Extract, Transform, Load",
        "Date-range queries using $gte and $lt",
        "update_one: write modified fields back",
        "Python classes for clean interfaces",
    ],
    notes="Five new things in this lesson.",
)

sb.add_generic(
    title="The ETL Paradigm",
    items=[
        "Extract: Retrieve raw data from a source",
        "Transform: Clean, filter, enrich or reshape",
        "Load: Write results to a destination",
    ],
    notes="ETL stands for Extract Transform Load.",
)

sb.add_generic(
    title="Extract: Date-Range Queries",
    items=[
        "Convert date string to Timestamp",
        "Compute end of day with DateOffset",
        "Use $gte and $lt for half-open interval",
    ],
    notes="The extract step.",
)

sb.add_checkpoints(
    items=[
        "After Extract: type is list, len > 0",
        "After Transform: inExperiment in doc",
        "After Load: result['n'] matches len(obs)",
        "After CSV: file exists, columns correct",
        "Re-run load: nModified == 0 is expected",
    ],
    notes="How to check your work at each stage.",
)

sb.add_exercise(
    items=[
        "Work through stages in order",
        "Write and test each function first",
        "Run checkpoint assertions after each step",
        "Print intermediate outputs for debugging",
    ],
    notes="Strategy for the exercises.",
)

sb.add_debugging(
    items=[
        "Empty results? Check date format",
        "KeyError? Print doc.keys()",
        "nModified is 0? Documents already updated",
        "Cursor exhausted? Re-run find()",
    ],
    notes="Common problems and fixes.",
)

sb.add_recap(
    items=[
        "ETL separates extraction, transform, load",
        "Date-range queries use $gte / $lt",
        "update_one with $set persists new fields",
        "Python classes bundle related logic",
    ],
    next_topic="Hypothesis testing with chi-square",
    notes="Wrap up.",
)

sb.save("/tmp/test_slidemaker_output.pptx")
