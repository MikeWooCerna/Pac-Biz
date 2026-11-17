"""Regression coverage for persistent Coaching-report exclusions."""
import ast
from pathlib import Path
import unittest

import pandas as pd


class CoachingExclusionTests(unittest.TestCase):
    EXCLUDED_ID = "1219721323810776"

    def setUp(self):
        source = Path(__file__).resolve().parents[1] / "dashboard.py"
        tree = ast.parse(source.read_text(encoding="utf-8"))
        function = next(
            node
            for node in tree.body
            if isinstance(node, ast.FunctionDef) and node.name == "transform_coaching_logs"
        )
        self.output_columns = [
            "Coaching ID", "Emp Name", "Coached by", "Title/Concern",
            "Incident Overview", "Employee Explanation", "Agreed Action Steps",
            "Remarks/Comment", "Coaching Category", "Category Status",
            "Confidence Level", "Reason", "Coaching Date", "Coaching Status",
            "Created Date", "Created Time",
        ]
        namespace = {
            "pd": pd,
            "COACHING_OUTPUT_COLUMNS": self.output_columns,
            "EXCLUDED_COACHING_GIDS": {self.EXCLUDED_ID},
            "assign_category_status": lambda frame: frame,
        }
        exec(compile(ast.Module(body=[function], type_ignores=[]), str(source), "exec"), namespace)
        self.transform = namespace["transform_coaching_logs"]

    def test_excluded_id_is_removed_from_normalized_coaching_source(self):
        source = pd.DataFrame([
            {"Task GID": self.EXCLUDED_ID, **{column: "excluded" for column in self.output_columns}},
            {"Task GID": "kept-id", **{column: "kept" for column in self.output_columns}},
        ])

        result = self.transform(source, pd.DataFrame())

        self.assertEqual(result["Coaching ID"].tolist(), ["kept"])


if __name__ == "__main__":
    unittest.main()
