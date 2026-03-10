"""Test rich content in SlideBuilder (code, flow, callout)."""

from __future__ import annotations

import sys

sys.path.insert(0, "src")

from slidemaker import SlideBuilder  # noqa: E402

TEMPLATE = "tests/data/template.pptx"

sb = SlideBuilder(TEMPLATE, default_template_page=4)

# ── Flow diagram slide ──────────────────────────────────────
sb.add_slide(
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
    callout="Separating stages makes each one independently testable and swappable",
    notes="ETL stands for Extract Transform Load.",
)

# ── Bullets + code block slide ──────────────────────────────
sb.add_slide(
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
sb.add_slide(
    title="Transform: Randomised Group Assignment",
    items=[
        "random.seed(42) makes the split reproducible",
        "random.shuffle(observations) reorders in place",
        "Integer division len(obs) // 2 finds the midpoint",
        "First half -> control, second half -> treatment",
    ],
    code="""random.seed(42)
random.shuffle(observations)
mid = len(observations) // 2
for doc in observations[:mid]:
    doc["inExperiment"] = True
    doc["group"] = "no email (control)"
for doc in observations[mid:]:
    doc["inExperiment"] = True
    doc["group"] = "email (treatment)" """,
    notes="Transform step.",
)

# ── Bullets only slide ──────────────────────────────────────
sb.add_slide(
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
sb.add_slide(
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
sb.add_slide(
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
sb.add_slide(
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
sb.add_slide(
    title="Inspecting Unfamiliar Objects with dir",
    items=[
        "dir(obj) lists all attributes and methods",
        'Filter internals: [a for a in dir(obj) if not a.startswith("_")]',
        "Useful for PyMongo return types like UpdateResult",
        "Check raw_result for the full MongoDB response",
    ],
    notes="Object inspection.",
)

sb.save("/tmp/test_rich_output.pptx")
