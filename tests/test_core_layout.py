from __future__ import annotations

import unittest
from unittest.mock import patch

from tests._util import new_slide

from slidemaker import core


class LayoutContentShapeTests(unittest.TestCase):
    def test_layout_rejects_invalid_table_combinations(self) -> None:
        _, slide = new_slide()

        with self.assertRaisesRegex(TypeError, "table must be a dictionary"):
            core.layout_content_shapes(slide, table=[])  # type: ignore[arg-type]

        with self.assertRaisesRegex(
            ValueError, "table cannot be combined with flow_boxes"
        ):
            core.layout_content_shapes(
                slide,
                table={"columns": ["A"], "rows": [["1"]]},
                flow_boxes=[{"label": "A"}],
            )

        with self.assertRaisesRegex(
            ValueError, "table can be combined with items or code, not both"
        ):
            core.layout_content_shapes(
                slide,
                items=["A"],
                code="print('x')",
                table={"columns": ["A"], "rows": [["1"]]},
            )

    def test_layout_places_flow_boxes_and_callout(self) -> None:
        _, slide = new_slide()

        with (
            patch("slidemaker.core.add_flow_boxes") as add_flow_boxes,
            patch("slidemaker.core.add_textbox") as add_textbox,
        ):
            core.layout_content_shapes(
                slide,
                flow_boxes=[{"label": "Extract"}],
                callout="Summary",
                slide_style={"font-size": 28},
            )

        add_flow_boxes.assert_called_once_with(
            slide,
            boxes=[{"label": "Extract"}],
            left=core.CONTENT_LEFT,
            top=core.CONTENT_TOP,
            style={"font-size": 28},
        )
        _, kwargs = add_textbox.call_args
        self.assertEqual(kwargs["text"], "Summary")
        self.assertGreater(kwargs["top"], core.CONTENT_TOP + core.Inches(3.0))

    def test_layout_places_items_and_code_without_overlap(self) -> None:
        _, slide = new_slide()

        with (
            patch("slidemaker.core.add_bullet_list") as add_bullet_list,
            patch("slidemaker.core.add_code_block") as add_code_block,
        ):
            core.layout_content_shapes(
                slide,
                items=["A", "B"],
                code="print('x')",
                slide_style={"italic": True},
                code_style={"line-numbers": True},
            )

        bullet_kwargs = add_bullet_list.call_args.kwargs
        code_kwargs = add_code_block.call_args.kwargs
        self.assertGreater(
            code_kwargs["top"], bullet_kwargs["top"] + bullet_kwargs["height"]
        )
        self.assertTrue(bullet_kwargs["style"]["italic"])
        self.assertTrue(code_kwargs["style"]["line-numbers"])

    def test_layout_places_items_and_table_with_merged_styles(self) -> None:
        _, slide = new_slide()

        with (
            patch("slidemaker.core.add_bullet_list") as add_bullet_list,
            patch("slidemaker.core.add_table") as add_table,
            patch("slidemaker.core.add_textbox") as add_textbox,
        ):
            core.layout_content_shapes(
                slide,
                items=["A"],
                table={
                    "columns": ["Field"],
                    "rows": [["Value"]],
                    "style": {"font-size": 18, "fill-color": "#EEEEEE"},
                    "header_style": {"font-color": "#111111"},
                    "cell_style": {"italic": True},
                    "banded_rows": True,
                },
                callout="Table summary",
                slide_style={"font-size": 24},
                table_style={"padding": "6pt"},
                table_header_style={"bold": True},
                table_cell_style={"font-color": "#333333"},
            )

        bullet_kwargs = add_bullet_list.call_args.kwargs
        table_kwargs = add_table.call_args.kwargs
        callout_kwargs = add_textbox.call_args.kwargs
        self.assertGreater(
            table_kwargs["top"], bullet_kwargs["top"] + bullet_kwargs["height"]
        )
        self.assertEqual(table_kwargs["style"]["font-size"], 18)
        self.assertEqual(table_kwargs["style"]["padding"], "6pt")
        self.assertTrue(table_kwargs["header_style"]["bold"])
        self.assertEqual(table_kwargs["header_style"]["font-color"], "#111111")
        self.assertTrue(table_kwargs["cell_style"]["italic"])
        self.assertEqual(table_kwargs["cell_style"]["font-color"], "#333333")
        self.assertTrue(table_kwargs["banded_rows"])
        self.assertGreater(
            callout_kwargs["top"], table_kwargs["top"] + table_kwargs["height"]
        )

    def test_layout_places_code_and_table_without_overlap(self) -> None:
        _, slide = new_slide()

        with (
            patch("slidemaker.core.add_code_block") as add_code_block,
            patch("slidemaker.core.add_table") as add_table,
        ):
            core.layout_content_shapes(
                slide,
                code="print('x')",
                table={"columns": ["Field"], "rows": [["Value"]]},
            )

        code_kwargs = add_code_block.call_args.kwargs
        table_kwargs = add_table.call_args.kwargs
        self.assertGreater(
            table_kwargs["top"], code_kwargs["top"] + code_kwargs["height"]
        )

    def test_layout_places_single_content_shapes(self) -> None:
        _, slide = new_slide()

        with (
            patch("slidemaker.core.add_bullet_list") as add_bullet_list,
            patch("slidemaker.core.add_code_block") as add_code_block,
            patch("slidemaker.core.add_table") as add_table,
        ):
            core.layout_content_shapes(slide, items=["A"])
            core.layout_content_shapes(slide, code="print('x')")
            core.layout_content_shapes(
                slide, table={"columns": ["Field"], "rows": [["Value"]]}
            )

        self.assertEqual(add_bullet_list.call_count, 1)
        self.assertEqual(add_code_block.call_count, 1)
        self.assertEqual(add_table.call_count, 1)
        self.assertEqual(
            add_bullet_list.call_args.kwargs["height"], core.CONTENT_HEIGHT
        )
        self.assertEqual(add_code_block.call_args.kwargs["height"], core.CONTENT_HEIGHT)
        self.assertEqual(add_table.call_args.kwargs["height"], core.CONTENT_HEIGHT)

    def test_layout_validates_table_specs(self) -> None:
        _, slide = new_slide()

        bad_specs = [
            (
                {"rows": "bad", "columns": ["A"]},
                "table rows must be a list of row lists",
            ),
            ({"rows": [["1"]], "columns": "bad"}, "table columns must be a list"),
            (
                {"rows": [["1"]], "columns": ["A"], "column_widths": "bad"},
                "table column_widths must be a list",
            ),
            (
                {"rows": [["1"]], "columns": ["A"], "row_heights": "bad"},
                "table row_heights must be a list",
            ),
            (
                {"rows": [["1"]], "columns": ["A"], "style": "bad"},
                "table style must be a dictionary",
            ),
            (
                {"rows": [["1"]], "columns": ["A"], "header_style": "bad"},
                "table header_style must be a dictionary",
            ),
            (
                {"rows": [["1"]], "columns": ["A"], "cell_style": "bad"},
                "table cell_style must be a dictionary",
            ),
        ]

        for spec, message in bad_specs:
            with self.subTest(spec=spec):
                with self.assertRaisesRegex(TypeError, message):
                    core.layout_content_shapes(slide, table=spec)


if __name__ == "__main__":
    unittest.main()
