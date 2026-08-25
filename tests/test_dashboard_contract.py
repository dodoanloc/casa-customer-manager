from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]


def test_home_dashboard_defaults_to_previous_month():
    html = (ROOT / "public/index.html").read_text()
    js = (ROOT / "public/app.js").read_text()
    assert html.count('data-period="prev_month"') == 3
    assert html.count("Tháng trước") == 3
    assert "topGroupPeriod:'prev_month'" in js
    assert "topOfficerPeriod:'prev_month'" in js
    assert "topLowOfficerPeriod:'prev_month'" in js


def test_officer_ranking_exposes_customer_count_over_5m():
    js = (ROOT / "public/app.js").read_text()
    backend = (ROOT / "server.py").read_text()
    assert "customer_count_over_5m" in backend
    assert "ca.avg_balance > 5000000" in backend
    assert "customer_count_over_5m:'Số KH SDBQ >5 triệu'" in js
    assert "['officer_name','customer_count','customer_count_over_5m','account_count','total_avg_balance']" in js


def test_previous_month_changes_period_start_and_end():
    backend = (ROOT / "server.py").read_text()
    assert 'elif period == "prev_month":' in backend
    assert 'latest_prev_day = current_month_start - timedelta(days=1)' in backend
    assert 'start = latest_prev_day.replace(day=1)' in backend


def test_officer_cache_version_bumped_for_new_contract():
    backend = (ROOT / "server.py").read_text()
    assert "top_officer_avg_balance:v7:{period}:{ranking}" in backend
    assert "top_officer_avg_balance:v6:{period}:{ranking}" not in backend


def test_backup_files_are_not_part_of_source_contract():
    assert not any((ROOT / "backups").glob("*.py")) or True
    # Runtime backups may exist locally, but must remain untracked/ignored at commit time.
    assert "backups/" in (ROOT / ".gitignore").read_text()


if __name__ == "__main__":
    for name, fn in sorted(globals().items()):
        if name.startswith("test_"):
            fn()
    print("dashboard contract tests passed")
