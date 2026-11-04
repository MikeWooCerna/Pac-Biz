"""Exercise Skyline's request handling without loading local credentials."""
import ast
from datetime import date
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock

import requests


class SkylineTransportTests(unittest.TestCase):
    script = "Skyline/Skyline_pull.py"

    def setUp(self):
        source = Path(__file__).resolve().parents[2] / "Quality" / self.script
        tree = ast.parse(source.read_text(encoding="utf-8-sig"))
        function = next(n for n in tree.body if isinstance(n, ast.FunctionDef) and n.name == "fetch_week")
        namespace = dict(requests=requests, time=SimpleNamespace(sleep=lambda _: None),
                         QA_BASE_URL="https://example.invalid", QA_ACCOUNT_ID="skyline",
                         REQUEST_TIMEOUT_SECONDS=90, REQUEST_MAX_ATTEMPTS=4)
        exec(compile(ast.Module(body=[function], type_ignores=[]), str(source), "exec"), namespace)
        self.fetch = namespace["fetch_week"]
        self.start, self.end = date(2026, 6, 8), date(2026, 6, 14)

    def response(self, status=200, body=b'{"data":[{"evaluation_id":"one"}]}'):
        response = requests.Response()
        response.status_code = status
        response._content = body
        return response

    def test_interrupted_week_recovers_without_losing_records(self):
        for error in [requests.ConnectionError("connection reset"),
                      requests.exceptions.ChunkedEncodingError("interrupted body"),
                      requests.Timeout("read timeout")]:
            with self.subTest(error=type(error).__name__):
                session = Mock()
                session.get.side_effect = [error, self.response()]
                self.assertEqual(self.fetch(session, "test-token", self.start, self.end),
                                 [{"evaluation_id": "one"}])

    def test_truncated_json_recovers(self):
        session = Mock()
        session.get.side_effect = [self.response(body=b'{"data":['), self.response()]
        self.assertEqual(self.fetch(session, "test-token", self.start, self.end),
                         [{"evaluation_id": "one"}])

    def test_server_error_recovers(self):
        session = Mock()
        session.get.side_effect = [self.response(503), self.response()]
        self.assertEqual(self.fetch(session, "test-token", self.start, self.end),
                         [{"evaluation_id": "one"}])

    def test_authentication_failure_does_not_retry(self):
        session = Mock()
        session.get.side_effect = [self.response(401), self.response()]
        with self.assertRaises(requests.HTTPError):
            self.fetch(session, "test-token", self.start, self.end)
        self.assertEqual(session.get.call_count, 1)

    def test_exhaustion_raises_instead_of_returning_partial_data(self):
        session = Mock()
        session.get.side_effect = requests.ConnectionError("connection reset")
        with self.assertRaisesRegex(RuntimeError, "2026-06-08.*2026-06-14"):
            self.fetch(session, "test-token", self.start, self.end)
        self.assertEqual(session.get.call_count, 4)


class DataCarzTransportTests(SkylineTransportTests):
    script = "Data Carz/dc_pull.py"


class KelownaTransportTests(SkylineTransportTests):
    script = "Kelowna/kel_pull.py"


class OlliesTransportTests(SkylineTransportTests):
    script = "Ollies/ol_pull.py"


if __name__ == "__main__":
    unittest.main()
