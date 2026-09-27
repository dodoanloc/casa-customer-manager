from pathlib import Path
from server import report_casa_decrease

ROOT = Path(__file__).resolve().parents[1]

def test_casa_decrease_ui_contract():
    html = (ROOT / "public/index.html").read_text()
    js = (ROOT / "public/app.js").read_text()
    assert "casaDecreaseReportTab" in html
    assert "Báo cáo KH giảm CASA" in html
    assert "setReportMode('casa_decrease')" in js
    assert "casa-decrease" in js

def test_report_casa_decrease_function():
    data = report_casa_decrease(report_date="2026-09-25")
    assert data["report_date"] == "2026-09-25"
    assert data["prev_date"] == "2026-09-24"
    assert data["count"] > 0
    assert data["total_decrease"] > 0
    assert len(data["rows"]) == data["count"]
    # Check decrease_amount order (descending within each currency)
    by_currency = {}
    for r in data["rows"]:
        by_currency.setdefault(r["currency"], []).append(r["decrease_amount"])
    for curr, decreases in by_currency.items():
        assert decreases == sorted(decreases, reverse=True)
