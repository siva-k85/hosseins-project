from pathlib import Path

from bs4 import BeautifulSoup


def test_extract_rows_reordered_columns(main_code):
    fixture_path = Path(__file__).resolve().parent / "fixtures" / "emma_reordered_columns.html"
    html = fixture_path.read_text()
    soup = BeautifulSoup(html, "html.parser")

    rows = main_code.extract_rows(soup)
    assert len(rows) == 2
    first = rows[0]
    assert first["solicitation_id"] == "IFB-2024-020"
    assert first["category"] == "Construction"
    assert first["procurement_method"] == "Invitation for Bid"
    assert first["due_dt_raw"].startswith("01/20/2025")


def test_deduplicate_rows(main_code):
    rows = [
        {"title": "Project A", "agency": "Agency 1", "_publish_dt_key": "2024-01-01", "solicitation_id": "A1"},
        {"title": "Project A", "agency": "Agency 1", "_publish_dt_key": "2024-01-01", "solicitation_id": "A1"},
        {"title": "Project B", "agency": "Agency 2", "_publish_dt_key": "2024-01-02", "solicitation_id": ""},
    ]
    deduped, duplicates = main_code._deduplicate_rows(rows)
    assert len(deduped) == 2
    assert duplicates == 1


def test_make_record_id(main_code):
    row_with_id = {"solicitation_id": "ABC123", "url": "/page?id=1", "title": "Example"}
    assert main_code._make_record_id(row_with_id) == "ABC123"

    row_with_url = {"solicitation_id": "", "url": "/extranet/456", "title": "Example"}
    assert main_code._make_record_id(row_with_url) == "456"

    row_hash = {"solicitation_id": "", "url": "", "title": "Example"}
    assert main_code._make_record_id(row_hash).startswith("emma_")
