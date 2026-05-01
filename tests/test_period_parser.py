from app.services.ledger import normalize_period


def test_normalize_period_cn():
    assert normalize_period("2026 年 1 月") == "2026-01"


def test_normalize_period_dash():
    assert normalize_period("2026-05") == "2026-05"


def test_normalize_period_compact():
    assert normalize_period("202605") == "2026-05"
