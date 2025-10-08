from datetime import datetime

import pytest


def test_parse_publish_dt_formats(main_code):
    samples = [
        ("12/14/2024 09:15:30 AM", datetime(2024, 12, 14, 9, 15, 30)),
        ("12/14/2024 09:15 AM", datetime(2024, 12, 14, 9, 15)),
        ("2024-12-14 09:15:30", datetime(2024, 12, 14, 9, 15, 30)),
        ("12/14/24 09:15 AM", datetime(2024, 12, 14, 9, 15)),
    ]

    for raw, expected in samples:
        parsed = main_code.parse_publish_dt(raw)
        assert parsed is not None, f"Expected parse for {raw}"
        expected_local = main_code.localize_et(expected)
        assert parsed == expected_local


def test_parse_publish_dt_invalid(main_code):
    assert main_code.parse_publish_dt("not a date") is None
    assert main_code.parse_publish_dt("") is None
