"""Test rich content in SlideBuilder (code, flow, callout)."""

from __future__ import annotations

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
)

sb.add_toolkit(
    items=[
        "MongoClient connects to MongoDB",
        "collection.find(query) returns a cursor",
        "collection.aggregate([...]) runs pipelines",
        "pd.DataFrame(list_of_dicts) converts documents",
    ],
)

sb.add_whats_new(
    items=[
        "**ETL pattern**: Extract → Transform → Load as separate, testable stages",
        "**Date-range queries**: using $gte and $lt in MongoDB",
        "**update_one**: write modified fields back into the database",
        "**Python classes**: bundle the full pipeline behind a clean interface",
        "**Randomised group assignment**: with random.shuffle and random.seed",
    ],
)

# ── Flow diagram slide ──────────────────────────────────────
sb.add_generic(
    title="The ETL Paradigm",
    flow_boxes=[
        {
            "label": "EXTRACT",
            "desc": "Retrieve raw data\nfrom a source",
            "style": {"fill-color": "#2E86AB"},
        },
        {
            "label": "TRANSFORM",
            "desc": "Clean, filter, enrich\nor reshape",
            "style": {"fill-color": "#48A99A"},
        },
        {
            "label": "LOAD",
            "desc": "Write results to\na destination",
            "style": {"fill-color": "#E86F51"},
        },
    ],
    callout=("Separating stages makes each one independently testable and swappable"),
    notes="ETL stands for Extract Transform Load.",
)

# ── Bullets + code block slide ──────────────────────────────
sb.add_generic(
    title="Extract: Date-Range Queries in MongoDB",
    items=[
        "Convert date string to Timestamp with pd.to_datetime",
        "Compute end of day with pd.DateOffset(days=1)",
        "Use $gte and $lt for a half-open interval",
        "Combine date filter with field filter",
    ],
    code="""start = pd.to_datetime("2022-05-02",
                       format="%Y-%m-%d")
end = start + pd.DateOffset(days=1)
query = {
    "createdAt": {"$gte": start, "$lt": end},
    "admissionsQuiz": "incomplete",
}
results = list(collection.find(query))""",
    notes="The extract step.",
)

# ── Bullets + code block slide ──────────────────────────────
sb.add_generic(
    title="Transform: Randomised Group Assignment",
    items=[
        "random.seed(42) makes the split reproducible",
        "random.shuffle(observations) reorders in place",
        "Integer division len(obs) // 2 finds the midpoint",
        "First half → control, second half → treatment",
    ],
    code='''random.seed(42)
random.shuffle(observations)
mid = len(observations) // 2
for doc in observations[:mid]:
    doc["inExperiment"] = True
    doc["group"] = "no email (control)"
for doc in observations[mid:]:
    doc["inExperiment"] = True
    doc["group"] = "email (treatment)"''',
    notes="Transform step.",
)

# ── Bullets only slide ──────────────────────────────────────
sb.add_generic(
    title="Transform: Exporting Treatment Emails to CSV",
    items=[
        "Convert assigned documents to DataFrame",
        'Add a tracking column "tag" with a fixed value',
        "Filter with boolean mask for treatment group",
        "Build dated filename with strftime",
        "Save selected columns with to_csv(index=False)",
    ],
    notes="Secondary transform step.",
)

# ── Bullets + code block slide ──────────────────────────────
sb.add_generic(
    title="Load: Updating Documents with update_one",
    items=[
        "update_one(filter, update) modifies one document",
        'Filter identifies the target: {"_id": doc["_id"]}',
        '{"$set": doc} adds or overwrites fields',
        "matched_count = 1 if filter found a document",
        "modified_count = 0 on a second run",
    ],
    code="""for doc in observations:
    result = collection.update_one(
        {"_id": doc["_id"]},
        {"$set": doc},
    )""",
    notes="Load step.",
)

# ── Bullets + code block slide ──────────────────────────────
sb.add_generic(
    title="Python Classes: Bundling Data and Behaviour",
    items=[
        "A class groups attributes and methods into one object",
        "__init__ runs at creation time and sets instance attributes via self",
        "Methods receive self as their first argument",
        "You already use classes: DataFrame has .shape, .head(), .describe()",
    ],
    code="""class Greeter:
    def __init__(self, name="World"):
        self.name = name

    def greet(self):
        return f"Hello, {self.name}!"

g = Greeter("Data Science")
print(g.greet())""",
    notes="Python classes intro.",
)

# ── Flow diagram slide ──────────────────────────────────────
sb.add_generic(
    title="Building the MongoRepository Class",
    flow_boxes=[
        {
            "label": "__init__",
            "desc": "Stores\nself.collection",
            "style": {"fill-color": "#193952"},
        },
        {
            "label": "find_by_date",
            "desc": "Extract step\nusing self.collection",
            "style": {"fill-color": "#2E86AB"},
        },
        {
            "label": "update_applicants",
            "desc": "Load step\nreturns summary dict",
            "style": {"fill-color": "#48A99A"},
        },
        {
            "label": "assign_to_groups",
            "desc": "Full ETL\nin one call",
            "style": {"fill-color": "#E86F51"},
        },
    ],
    callout="One method call does the entire ETL pipeline",
    notes="MongoRepository class overview.",
)

# ── Bullets only slide ──────────────────────────────────────
sb.add_generic(
    title="Inspecting Unfamiliar Objects with dir",
    items=[
        "dir(obj) lists all attributes and methods",
        'Filter internals: [a for a in dir(obj) if not a.startswith("_")]',
        "Useful for PyMongo return types like UpdateResult",
        "Check raw_result for the full MongoDB response",
    ],
    notes="Object inspection.",
)

# ── Ending slides ───────────────────────────────────────────
sb.add_checkpoints(
    items=[
        "After Extract: type is list, len > 0, _id in results[0]",
        "After Transform: inExperiment and group in doc, two group labels exist",
        'After Load: result["n"] matches len(obs), nModified is sensible',
        "After CSV: file exists, columns [email, tag], row count matches",
        "Re-running load: nModified == 0 is expected, not an error",
    ],
    notes="Validation.",
)

sb.add_exercise(
    items=[
        "Work through ETL stages in order",
        "Write and test each standalone function first",
        "Run checkpoint assertions after each step",
        "If results look wrong, print intermediate outputs",
        "Class methods reuse standalone function logic",
    ],
    notes="Strategy.",
)

sb.add_debugging(
    items=[
        '**Empty results from find?**: Check date format matches "%Y-%m-%d"',
        "**KeyError on document?**: Print doc.keys(); field names are case-sensitive",
        "**nModified is 0?**: Documents already updated; reset database and rerun",
        "**Cursor exhausted?**: Re-run find() or aggregate(); cursors are single-use",
        "**Class attribute missing?**: Ensure self.attr is set in __init__",
    ],
    notes="Debugging tips.",
)

sb.add_recap(
    items=[
        "ETL separates extraction, transformation, and loading into testable stages",
        "Date-range queries use $gte / $lt with pandas Timestamps",
        "update_one with $set persists new fields back into MongoDB",
        "Python classes bundle related logic behind a clean interface",
    ],
    next_topic=(
        "Use experimental data to perform hypothesis testing with a chi-square test"
    ),
    notes="Wrap up.",
)

sb.save("/tmp/test_rich_output.pptx")
