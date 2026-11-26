"""
Tests: the synthetic data has a known answer key (see scripts/generate_sample_data.py),
so every planted exception must be found and every number must tie.

Run:  python -m pytest -v
"""

from pathlib import Path

import openpyxl
import pandas as pd
import pytest

from fundaudit.procedures import run

DATA = Path(__file__).resolve().parent.parent / "data"


@pytest.fixture(scope="module")
def r():
    return run(DATA)


def status_of(df, sid):
    return df.loc[df["security_id"] == sid, "status"].iloc[0]


def test_nav_recomputes(r):
    assert r.nav["nav_agrees"]
    assert r.nav["investments_agree"]
    assert r.nav["recomputed_nav_per_unit"] == pytest.approx(r.statement["reported_nav_per_unit"])


def test_custodian_rec_finds_planted_items(r):
    c = r.custodian_rec
    assert status_of(c, "SEC11") == "Not held at custodian"
    assert status_of(c, "SEC07") == "Timing - unsettled trade"
    assert status_of(c, "SEC03") == "Unrecorded corporate action"
    assert (c["status"] == "Agreed").sum() == 11


def test_unsettled_trade_is_not_a_misstatement(r):
    c = r.custodian_rec.set_index("security_id")
    assert c.loc["SEC07", "nav_misstatement"] == 0


def test_custodian_misstatement_values(r):
    c = r.custodian_rec.set_index("security_id")
    assert c.loc["SEC11", "nav_misstatement"] == pytest.approx(380_000.00)   # 40,000 x 9.50, overstated
    assert c.loc["SEC03", "nav_misstatement"] == pytest.approx(-64_500.00)   # 3,000 bonus x 21.50, understated


def test_price_testing(r):
    p = r.price_test.set_index("security_id")
    assert p.loc["SEC05", "status"] == "Exception"
    assert p.loc["SEC05", "nav_misstatement"] == pytest.approx(49_210.00)    # 20,000 x 1.90 x 1.295
    assert p.loc["SEC12", "status"] == "Exception - stale price"
    assert p.loc["SEC12", "nav_misstatement"] == pytest.approx(129_000.00)   # 3m x 4.30%
    assert p.loc["SEC08", "status"] == "Within tolerance"
    assert p.loc["SEC08", "nav_misstatement"] == 0


def test_wider_tolerance_removes_price_exception():
    wide = run(DATA, price_tolerance=0.05)
    assert not wide.price_test["exception"].any()


def test_management_fee_recalculation(r):
    f = r.fee_summary
    assert f["implied_rate"] == pytest.approx(0.0135, abs=1e-6)
    assert f["over_accrual"] == pytest.approx(f["fee_charged"] - f["expected_fee"])
    assert f["over_accrual"] > 0
    assert f["nav_misstatement"] == pytest.approx(-f["over_accrual"])


def test_materiality_and_conclusion(r):
    m = r.materiality
    assert m["overall"] == pytest.approx(r.statement["reported_nav"] * 0.01, abs=0.01)
    assert m["accumulated_misstatement"] == pytest.approx(493_710.00)
    assert m["conclusion"].startswith("MATERIAL")
    # The fee error is below clearly trivial, so it is not accumulated...
    fee_row = r.misstatements[r.misstatements["reference"] == "MGMT FEE"].iloc[0]
    assert not fee_row["above_clearly_trivial"]
    # ...but it is still reported as a high-rated finding (qualitative significance)
    assert r.findings.loc[r.findings["reference"] == "MGMT FEE", "rating"].iloc[0] == "High"


def test_findings_cover_every_exception(r):
    refs = set(r.findings["reference"])
    assert {"SEC11", "SEC12", "SEC05", "SEC03", "SEC07", "MGMT FEE"} <= refs


def test_workpaper_created(tmp_path, r):
    from fundaudit.workpaper import export_workpaper
    path = export_workpaper(r, tmp_path / "wp.xlsx")
    wb = openpyxl.load_workbook(path)
    assert wb.sheetnames[:5] == ["Summary", "Findings", "Price Testing", "Custodian Rec", "Mgmt Fee"]
    # Summary figures are formulas, not hard-coded numbers
    formulas = [c.value for row in wb["Summary"].iter_rows() for c in row
                if isinstance(c.value, str) and c.value.startswith("=")]
    assert len(formulas) > 20
