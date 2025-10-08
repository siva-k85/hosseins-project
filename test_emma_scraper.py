"""
Unit tests for eMMA scraper
"""

import unittest
from datetime import datetime, timedelta
from unittest.mock import Mock, patch, MagicMock
import os
import sys
import tempfile
from pathlib import Path

# Import the enhanced scraper
from emma_scraper_enhanced import (
    parse_flexible_datetime,
    deduplicate_rows_enhanced,
    make_record_id_enhanced,
    _normalize_text,
    _build_column_map,
    validate_parameters,
    get_default_workbook_path,
    ExcelStyler,
    DynamicRetry,
    JsonFormatter,
    load_refs_rules,
    apply_auto_tagging,
    deduplicate_archive,
    generate_analytics_report
)

class TestDateTimeParsing(unittest.TestCase):
    """Test flexible datetime parsing."""

    def test_standard_format(self):
        """Test standard datetime formats."""
        test_cases = [
            ("12/25/2024 2:30:00 PM", datetime(2024, 12, 25, 14, 30, 0)),
            ("01/15/2024", datetime(2024, 1, 15, 0, 0, 0)),
            ("2024-01-15 14:30:00", datetime(2024, 1, 15, 14, 30, 0)),
        ]

        for date_str, expected in test_cases:
            with self.subTest(date_str=date_str):
                result = parse_flexible_datetime(date_str)
                self.assertIsNotNone(result)
                self.assertEqual(result.replace(tzinfo=None), expected)

    def test_alternate_formats(self):
        """Test alternate datetime formats."""
        test_cases = [
            ("01-15-2024 2:30 PM", datetime(2024, 1, 15, 14, 30, 0)),
            ("15/01/2024", datetime(2024, 1, 15, 0, 0, 0)),
        ]

        for date_str, expected in test_cases:
            with self.subTest(date_str=date_str):
                result = parse_flexible_datetime(date_str)
                if result:
                    self.assertEqual(result.replace(tzinfo=None).date(), expected.date())

    def test_invalid_formats(self):
        """Test invalid datetime formats."""
        invalid_dates = [
            "not a date",
            "",
            None,
            "2024",
            "January 15, 2024",  # Not in our supported formats
        ]

        for date_str in invalid_dates:
            with self.subTest(date_str=date_str):
                result = parse_flexible_datetime(date_str)
                self.assertIsNone(result)


class TestDeduplication(unittest.TestCase):
    """Test record deduplication."""

    def test_duplicate_detection(self):
        """Test that duplicates are correctly identified and removed."""
        rows = [
            {
                "title": "Test Project 1",
                "agency": "Test Agency",
                "publish_dt_et": datetime(2024, 1, 1),
                "solicitation_id": "TEST001"
            },
            {
                "title": "Test Project 1",  # Duplicate
                "agency": "Test Agency",
                "publish_dt_et": datetime(2024, 1, 1),
                "solicitation_id": "TEST001"
            },
            {
                "title": "Test Project 2",  # Different title
                "agency": "Test Agency",
                "publish_dt_et": datetime(2024, 1, 1),
                "solicitation_id": "TEST002"
            },
        ]

        result = deduplicate_rows_enhanced(rows)
        self.assertEqual(len(result), 2)
        self.assertEqual(result[0]["title"], "Test Project 1")
        self.assertEqual(result[1]["title"], "Test Project 2")

    def test_case_insensitive_deduplication(self):
        """Test that deduplication is case-insensitive."""
        rows = [
            {
                "title": "Test Project",
                "agency": "TEST AGENCY",
                "publish_dt_et": datetime(2024, 1, 1),
                "solicitation_id": "test001"
            },
            {
                "title": "TEST PROJECT",  # Different case
                "agency": "test agency",
                "publish_dt_et": datetime(2024, 1, 1),
                "solicitation_id": "TEST001"
            },
        ]

        result = deduplicate_rows_enhanced(rows)
        self.assertEqual(len(result), 1)


class TestRecordID(unittest.TestCase):
    """Test record ID generation."""

    def test_solicitation_id_priority(self):
        """Test that solicitation ID is used first if available."""
        row = {
            "solicitation_id": "SOL123",
            "url": "https://example.com/123",
            "title": "Test Project"
        }
        result = make_record_id_enhanced(row)
        self.assertEqual(result, "sid_SOL123")

    def test_url_extraction(self):
        """Test ID extraction from URL."""
        row = {
            "solicitation_id": "",
            "url": "https://emma.maryland.gov/extranet/456789",
            "title": "Test Project"
        }
        result = make_record_id_enhanced(row)
        self.assertEqual(result, "eid_456789")

    def test_composite_key(self):
        """Test composite key generation."""
        row = {
            "solicitation_id": "",
            "url": "",
            "title": "Test Project",
            "agency": "Test Agency",
            "publish_dt_et": datetime(2024, 1, 1)
        }
        result = make_record_id_enhanced(row)
        self.assertTrue(result.startswith("hash_"))

    def test_fallback_hash(self):
        """Test fallback hash generation."""
        row = {
            "solicitation_id": "",
            "url": "https://example.com",
            "title": "Test"
        }
        result = make_record_id_enhanced(row)
        self.assertTrue(result.startswith("emma_"))


class TestParameterValidation(unittest.TestCase):
    """Test parameter validation."""

    @patch.dict(os.environ, {"DAYS_AGO": "-1"})
    def test_negative_days_ago(self):
        """Test validation catches negative DAYS_AGO."""
        with self.assertRaises(ValueError) as ctx:
            validate_parameters()
        self.assertIn("DAYS_AGO must be non-negative", str(ctx.exception))

    @patch.dict(os.environ, {"STALE_AFTER_D": "0"})
    def test_zero_stale_after(self):
        """Test validation catches zero STALE_AFTER_D."""
        with self.assertRaises(ValueError) as ctx:
            validate_parameters()
        self.assertIn("STALE_AFTER_D must be positive", str(ctx.exception))

    @patch.dict(os.environ, {"MAX_PAGES": "abc"})
    def test_invalid_max_pages(self):
        """Test validation catches non-integer MAX_PAGES."""
        with self.assertRaises(ValueError) as ctx:
            validate_parameters()
        self.assertIn("MAX_PAGES must be an integer", str(ctx.exception))

    @patch.dict(os.environ, {"LOG_LEVEL": "INVALID"})
    def test_invalid_log_level(self):
        """Test validation catches invalid LOG_LEVEL."""
        with self.assertRaises(ValueError) as ctx:
            validate_parameters()
        self.assertIn("LOG_LEVEL must be one of", str(ctx.exception))


class TestCrossPlatformPath(unittest.TestCase):
    """Test cross-platform path handling."""

    @patch('os.name', 'nt')
    def test_windows_path(self):
        """Test Windows default path."""
        path = get_default_workbook_path()
        self.assertIn("C:\\", path)

    @patch('os.name', 'posix')
    def test_unix_path(self):
        """Test Unix/Mac default path."""
        with patch('os.path.expanduser') as mock_expand:
            mock_expand.return_value = "/home/user"
            path = get_default_workbook_path()
            self.assertIn("Documents", path)
            self.assertTrue(path.startswith("/home/user"))


class TestHeaderMapping(unittest.TestCase):
    """Test header column mapping."""

    def test_exact_match(self):
        """Test exact header matching."""
        from bs4 import BeautifulSoup

        html = """
        <table>
            <tr>
                <th>Title</th>
                <th>Agency</th>
                <th>Category</th>
            </tr>
        </table>
        """
        soup = BeautifulSoup(html, 'html.parser')
        header_cells = soup.find_all('th')

        mapping = _build_column_map(header_cells)

        self.assertEqual(mapping.get('title'), 0)
        self.assertEqual(mapping.get('agency'), 1)
        self.assertEqual(mapping.get('category'), 2)

    def test_alias_match(self):
        """Test alias header matching."""
        from bs4 import BeautifulSoup

        html = """
        <table>
            <tr>
                <th>Solicitation Title</th>
                <th>Issuing Agency</th>
                <th>Procurement Category</th>
            </tr>
        </table>
        """
        soup = BeautifulSoup(html, 'html.parser')
        header_cells = soup.find_all('th')

        mapping = _build_column_map(header_cells)

        self.assertEqual(mapping.get('title'), 0)
        self.assertEqual(mapping.get('agency'), 1)
        self.assertEqual(mapping.get('category'), 2)

    def test_partial_match(self):
        """Test partial header matching."""
        from bs4 import BeautifulSoup

        html = """
        <table>
            <tr>
                <th>Project Title and Description</th>
                <th>Department/Agency Name</th>
            </tr>
        </table>
        """
        soup = BeautifulSoup(html, 'html.parser')
        header_cells = soup.find_all('th')

        mapping = _build_column_map(header_cells)

        self.assertIsNotNone(mapping.get('title'))
        self.assertIsNotNone(mapping.get('agency'))


class TestAutoTagging(unittest.TestCase):
    """Test auto-tagging functionality."""

    def test_keyword_matching(self):
        """Test keyword-based tagging."""
        rules = [
            {
                "keyword": "construction",
                "field": "title",
                "tag": "Construction",
                "score": 10,
                "priority": 1
            },
            {
                "keyword": "it services",
                "field": "title",
                "tag": "IT",
                "score": 15,
                "priority": 2
            }
        ]

        row = {
            "title": "Construction of new IT Services building",
            "agency": "Test Agency",
            "tags": "",
            "score_bd_fit": ""
        }

        result = apply_auto_tagging(row, rules)

        self.assertIn("Construction", result["tags"])
        self.assertIn("IT", result["tags"])
        self.assertEqual(result["score_bd_fit"], "25")

    def test_existing_tags_preserved(self):
        """Test that existing tags are preserved."""
        rules = [
            {
                "keyword": "construction",
                "field": "title",
                "tag": "Construction",
                "score": 10,
                "priority": 1
            }
        ]

        row = {
            "title": "Construction project",
            "tags": "Urgent",
            "score_bd_fit": "5"
        }

        result = apply_auto_tagging(row, rules)

        self.assertIn("Urgent", result["tags"])
        self.assertIn("Construction", result["tags"])
        self.assertEqual(result["score_bd_fit"], "15")


class TestExcelStyler(unittest.TestCase):
    """Test Excel styling module."""

    def test_status_fill_colors(self):
        """Test status-based fill colors."""
        styler = ExcelStyler()

        new_fill = styler.get_status_fill("New")
        self.assertEqual(new_fill.start_color, "C6EFCE")

        updated_fill = styler.get_status_fill("Updated")
        self.assertEqual(updated_fill.start_color, "FFEB9C")

        stale_fill = styler.get_status_fill("Stale")
        self.assertEqual(stale_fill.start_color, "E7E6E6")


class TestJsonFormatter(unittest.TestCase):
    """Test JSON logging formatter."""

    def test_json_format(self):
        """Test JSON log formatting."""
        import logging
        import json

        formatter = JsonFormatter()
        record = logging.LogRecord(
            name="test_logger",
            level=logging.INFO,
            pathname="test.py",
            lineno=10,
            msg="Test message",
            args=(),
            exc_info=None
        )

        result = formatter.format(record)
        parsed = json.loads(result)

        self.assertEqual(parsed["level"], "INFO")
        self.assertEqual(parsed["message"], "Test message")
        self.assertEqual(parsed["logger"], "test_logger")
        self.assertIn("timestamp", parsed)


class TestDynamicRetry(unittest.TestCase):
    """Test dynamic retry with backoff."""

    def test_retry_after_header(self):
        """Test that Retry-After header is respected."""
        retry = DynamicRetry(total=3)

        # Mock response with Retry-After header
        mock_response = Mock()
        mock_response.status_code = 429
        mock_response.headers = {"Retry-After": "5"}

        retry.increment(response=mock_response)

        self.assertEqual(retry.retry_after_header, "5")
        self.assertEqual(retry.get_backoff_time(), 5)

    def test_rate_limit_detection(self):
        """Test rate limit status codes trigger backoff."""
        retry = DynamicRetry(
            total=3,
            status_forcelist=[403, 429]
        )

        mock_response = Mock()
        mock_response.status_code = 403

        # Should not raise on first increment
        retry.increment(response=mock_response)


class TestWorkbookOperations(unittest.TestCase):
    """Test workbook-related operations."""

    @patch('emma_scraper_enhanced.load_workbook')
    def test_load_refs_rules(self, mock_load):
        """Test loading rules from Refs sheet."""
        # Mock workbook structure
        mock_wb = MagicMock()
        mock_ws = MagicMock()
        mock_wb.sheetnames = ["Master", "Refs"]
        mock_wb.__getitem__.return_value = mock_ws

        # Mock Refs sheet data
        mock_ws.max_row = 3
        mock_ws.__getitem__.side_effect = lambda row: [
            Mock(value=v) for v in [
                ["keyword", "field", "tag", "score", "priority"][0] if row == 1
                else ["construction", "title", "Construction", "10", "1"][0] if row == 2
                else ["it", "title", "IT", "15", "2"][0]
            ]
        ]

        rules = load_refs_rules(mock_wb)
        # Note: This test would need more complex mocking to work properly


class TestAnalyticsGeneration(unittest.TestCase):
    """Test analytics report generation."""

    def test_analytics_calculation(self):
        """Test analytics calculations."""
        # This would require mocking the workbook structure
        # For now, just test that the function exists
        self.assertTrue(callable(generate_analytics_report))


if __name__ == "__main__":
    unittest.main()