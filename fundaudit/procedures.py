"""
Fund audit procedures for an investment fund at its valuation date.

Procedures (each maps to an audit assertion)
--------------------------------------------
1. Custodian reconciliation   - existence & completeness of investments
2. Price testing              - valuation of investments
3. Management fee recalculation - accuracy of expenses / accrued liabilities
4. NAV recomputation          - arithmetic accuracy of the reported NAV
5. Misstatement evaluation    - aggregate misstatements vs materiality

Sign convention for misstatements: a POSITIVE amount means the reported NAV
is OVERSTATED; a negative amount means it is understated.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

DEFAULTS = {
    "price_tolerance": 0.01,       # price differences above 1% are exceptions
    "stale_days": 5,               # fund price older than 5 days before valuation date = stale
    "materiality_pct": 0.01,       # overall materiality = 1% of NAV (illustrative)
    "performance_pct": 0.75,       # performance materiality = 75% of overall
    "trivial_pct": 0.05,           # clearly trivial = 5% of overall
}


@dataclass
class AuditResult:
    terms: dict
    statement: dict
    settings: dict
    holdings: pd.DataFrame
    custodian_rec: pd.DataFrame
    price_test: pd.DataFrame
    fee_daily: pd.DataFrame
    fee_summary: dict
    nav: dict
    misstatements: pd.DataFrame
    materiality: dict
    findings: pd.DataFrame
    raw: dict = field(default_factory=dict)


# ---------------------------------------------------------------------------
# Loading
# ---------------------------------------------------------------------------
def load(data_dir: Path) -> dict:
    d = Path(data_dir)
    with open(d / "fund_terms.json") as f:
        terms = json.load(f)
    with open(d / "nav_statement.json") as f:
        statement = json.load(f)
    return {
        "terms": terms,
        "statement": statement,
        "holdings": pd.read_csv(d / "fund_holdings.csv", parse_dates=["fund_price_date"]),
        "custodian": pd.read_csv(d / "custodian_statement.csv"),
        "prices": pd.read_csv(d / "independent_prices.csv", parse_dates=["price_date"]),
        "pending": pd.read_csv(d / "pending_trades.csv", parse_dates=["trade_date", "settle_date"]),
        "corp_actions": pd.read_csv(d / "corporate_actions.csv", parse_dates=["ex_date", "pay_date"]),
        "daily_nav": pd.read_csv(d / "daily_nav_oct2026.csv", parse_dates=["date"]),
    }


def price_factor(unit: str) -> float:
    """Bonds are quoted per 100 nominal, equities per share."""
    return 0.01 if unit == "nominal" else 1.0


def add_values(holdings: pd.DataFrame, fx: dict) -> pd.DataFrame:
    h = holdings.copy()
    h["factor"] = h["unit"].map(price_factor)
    h["fx_rate"] = h["currency"].map(fx)
    h["mv_fund_sgd"] = (h["quantity"] * h["fund_price"] * h["factor"] * h["fx_rate"]).round(2)
    return h


# ---------------------------------------------------------------------------
# 1. Custodian reconciliation
# ---------------------------------------------------------------------------
def custodian_reconciliation(holdings, custodian, pending, corp_actions, prices, valuation_date) -> pd.DataFrame:
    rec = holdings[["security_id", "security_name", "currency", "unit", "quantity", "factor", "fx_rate"]].merge(
        custodian, on="security_id", how="outer")
    rec["quantity"] = rec["quantity"].fillna(0)
    rec["quantity_held"] = rec["quantity_held"].fillna(0)
    rec["difference"] = rec["quantity_held"] - rec["quantity"]  # custodian minus books
    rec = rec.merge(prices[["security_id", "independent_price"]], on="security_id", how="left")

    # Expected differences
    unsettled = pending[(pending["trade_date"] <= valuation_date) & (pending["settle_date"] > valuation_date)]
    buys = unsettled.assign(signed=lambda t: t["quantity"].where(t["side"] == "BUY", -t["quantity"]))
    rec["unsettled_trades"] = rec["security_id"].map(buys.groupby("security_id")["signed"].sum()).fillna(0)

    paid = corp_actions[corp_actions["pay_date"] <= valuation_date]
    ca_shares = {}
    for _, c in paid.iterrows():
        base_qty = rec.loc[rec["security_id"] == c["security_id"], "quantity"].sum()
        ca_shares[c["security_id"]] = base_qty * c["new_shares"] / c["per_old_shares"]
    rec["corporate_action_shares"] = rec["security_id"].map(ca_shares).fillna(0)

    statuses, explanations, impacts, assertions = [], [], [], []
    for _, r in rec.iterrows():
        diff = r["difference"]
        value_per_unit = r["independent_price"] * r["factor"] * r["fx_rate"]
        if diff == 0:
            s, e, impact, a = "Agreed", "Quantity agrees to custodian.", 0.0, ""
        elif r["quantity_held"] == 0 and r["quantity"] > 0 and r["unsettled_trades"] == 0:
            s = "Not held at custodian"
            e = "Recorded in the books but not on the custodian statement, and no pending trade explains it."
            impact, a = round(r["quantity"] * value_per_unit, 2), "Existence"
        elif diff == -r["unsettled_trades"] and r["unsettled_trades"] != 0:
            s = "Timing - unsettled trade"
            e = (f"Purchase of {abs(r['unsettled_trades']):,.0f} traded before year end, settles after; "
                 "recorded on trade date (correct).")
            impact, a = 0.0, "Cut-off"
        elif diff == r["corporate_action_shares"] and diff != 0:
            s = "Unrecorded corporate action"
            e = f"Bonus shares of {diff:,.0f} received at custodian but not recorded in the books."
            impact, a = round(-diff * value_per_unit, 2), "Completeness"
        else:
            s, e = "Unexplained difference", "Difference not explained by pending trades or corporate actions."
            impact, a = round(-diff * value_per_unit, 2), "Existence / Completeness"
        statuses.append(s); explanations.append(e); impacts.append(impact); assertions.append(a)

    rec["status"], rec["explanation"] = statuses, explanations
    rec["nav_misstatement"], rec["assertion"] = impacts, assertions
    return rec.sort_values("security_id").reset_index(drop=True)


# ---------------------------------------------------------------------------
# 2. Price testing
# ---------------------------------------------------------------------------
def price_testing(holdings, prices, valuation_date, tolerance, stale_days) -> pd.DataFrame:
    p = holdings.merge(prices[["security_id", "independent_price", "source"]], on="security_id", how="left")
    p["mv_independent_sgd"] = (p["quantity"] * p["independent_price"] * p["factor"] * p["fx_rate"]).round(2)
    p["value_difference"] = (p["mv_fund_sgd"] - p["mv_independent_sgd"]).round(2)  # + = overvalued
    p["price_diff_pct"] = (p["fund_price"] - p["independent_price"]) / p["independent_price"]
    p["price_age_days"] = (valuation_date - p["fund_price_date"]).dt.days
    p["stale"] = p["price_age_days"] > stale_days
    p["exception"] = p["price_diff_pct"].abs() > tolerance

    def status(r):
        if r["exception"] and r["stale"]:
            return "Exception - stale price"
        if r["exception"]:
            return "Exception"
        if r["stale"]:
            return "Stale price"
        return "Within tolerance" if r["price_diff_pct"] != 0 else "Agreed"
    p["status"] = p.apply(status, axis=1)
    p["nav_misstatement"] = p["value_difference"].where(p["exception"], 0.0)
    return p


# ---------------------------------------------------------------------------
# 3. Management fee recalculation
# ---------------------------------------------------------------------------
def fee_recalculation(daily_nav, rate, day_count, fee_charged):
    d = daily_nav.copy()
    d["daily_fee"] = d["net_assets_sgd"] * rate / day_count
    expected = round(float(d["daily_fee"].sum()), 2)
    over = round(fee_charged - expected, 2)
    implied = fee_charged * day_count / float(d["net_assets_sgd"].sum())
    summary = {
        "contract_rate": rate, "expected_fee": expected, "fee_charged": fee_charged,
        "over_accrual": over, "implied_rate": round(implied, 6),
        # Over-accrued liability means NAV is UNDERSTATED -> negative overstatement
        "nav_misstatement": -over,
    }
    return d, summary


# ---------------------------------------------------------------------------
# 4. NAV recomputation
# ---------------------------------------------------------------------------
def recompute_nav(statement, holdings):
    inv_books = round(float(holdings["mv_fund_sgd"].sum()), 2)
    nav = (statement["investments_at_market_value"] + statement["cash_at_bank"]
           + statement["dividends_receivable"] - statement["payable_for_investments_purchased"]
           - statement["redemptions_payable"] - statement["other_payables"]
           - statement["accrued_management_fee"])
    nav = round(nav, 2)
    return {
        "investments_per_holdings": inv_books,
        "investments_agree": bool(abs(inv_books - statement["investments_at_market_value"]) < 0.01),
        "recomputed_nav": nav,
        "nav_agrees": abs(nav - statement["reported_nav"]) < 0.01,
        "recomputed_nav_per_unit": round(nav / statement["units_in_issue"], 4),
    }


# ---------------------------------------------------------------------------
# 5. Misstatements, materiality, findings
# ---------------------------------------------------------------------------
def evaluate(statement, cust, price, fee, settings):
    nav = statement["reported_nav"]
    overall = round(nav * settings["materiality_pct"], 2)
    pm = round(overall * settings["performance_pct"], 2)
    ct = round(overall * settings["trivial_pct"], 2)

    rows = []
    for _, r in price[price["exception"]].iterrows():
        rows.append(["Valuation", r["security_id"], f"{r['security_name']}: {r['status'].lower()}",
                     r["nav_misstatement"]])
    for _, r in cust[cust["nav_misstatement"] != 0].iterrows():
        rows.append([r["assertion"], r["security_id"], f"{r['security_name']}: {r['status'].lower()}",
                     r["nav_misstatement"]])
    rows.append(["Accuracy", "MGMT FEE", "Management fee over-accrued vs prospectus rate",
                 fee["nav_misstatement"]])

    m = pd.DataFrame(rows, columns=["area", "reference", "description", "nav_overstatement"])
    m["above_clearly_trivial"] = m["nav_overstatement"].abs() >= ct
    m.insert(0, "ref", [f"M{i}" for i in range(1, len(m) + 1)])

    accumulated = round(float(m.loc[m["above_clearly_trivial"], "nav_overstatement"].sum()), 2)
    adjusted_nav = round(nav - accumulated, 2)
    if abs(accumulated) >= overall:
        conclusion = "MATERIAL - correction required before the NAV can be relied on"
    elif abs(accumulated) >= pm:
        conclusion = "Above performance materiality - request correction"
    else:
        conclusion = "Below performance materiality"
    mat = {
        "basis": "Reported NAV", "overall": overall, "performance": pm, "clearly_trivial": ct,
        "accumulated_misstatement": accumulated, "adjusted_nav": adjusted_nav,
        "adjusted_nav_per_unit": round(adjusted_nav / statement["units_in_issue"], 4),
        "nav_impact_pct": round(accumulated / nav, 6), "conclusion": conclusion,
    }
    return m, mat


def build_findings(cust, price, fee, mat) -> pd.DataFrame:
    f = []
    for _, r in cust[cust["status"] == "Not held at custodian"].iterrows():
        f.append(["Custodian reconciliation", "Existence", r["security_id"],
                  f"{r['security_name']} ({r['quantity']:,.0f} shares) is in the fund's records but not on the "
                  f"custodian statement, and no unsettled trade explains it.",
                  r["nav_misstatement"], "High",
                  "Obtain direct confirmation from the custodian and the purchase trade ticket. "
                  "If the holding cannot be evidenced, remove it from the NAV and investigate how it was recorded."])
    for _, r in price[price["exception"]].iterrows():
        stale = r["stale"]
        f.append(["Price testing", "Valuation", r["security_id"],
                  f"{r['security_name']} priced at {r['fund_price']:,.2f} vs independent {r['independent_price']:,.2f} "
                  f"({r['price_diff_pct']:+.1%})" + (f"; fund price is {r['price_age_days']} days old (stale)." if stale else "."),
                  r["nav_misstatement"], "High" if stale else "Medium",
                  "Revalue using the independent price. " + (
                      "Review the pricing policy for illiquid bonds and require prices at the valuation date."
                      if stale else "Ask the manager to explain the price source and any override.")])
    for _, r in cust[cust["status"] == "Unrecorded corporate action"].iterrows():
        f.append(["Custodian reconciliation", "Completeness", r["security_id"],
                  f"{r['security_name']}: {r['difference']:,.0f} bonus shares paid on the custodian account "
                  "were not recorded in the fund's books.",
                  r["nav_misstatement"], "Medium",
                  "Record the bonus shares. Review the corporate action process so events are booked by pay date."])
    f.append(["Management fee recalculation", "Accuracy", "MGMT FEE",
              f"Fee accrued at an implied {fee['implied_rate']:.2%} p.a. vs {fee['contract_rate']:.2%} in the "
              f"prospectus: over-accrued by SGD {fee['over_accrual']:,.2f} for October.",
              fee["nav_misstatement"], "High",
              "Below clearly trivial in amount but qualitatively significant: investors are overcharged every month. "
              "Correct the rate in the fee system, refund prior periods and confirm the rate with the trustee."])
    for _, r in cust[cust["status"] == "Timing - unsettled trade"].iterrows():
        f.append(["Custodian reconciliation", "Cut-off", r["security_id"],
                  f"{r['security_name']}: purchase of {abs(r['unsettled_trades']):,.0f} shares traded 30 Oct, settles 3 Nov.",
                  0.0, "Low", "No adjustment. Trace settlement to the November custodian statement."])

    df = pd.DataFrame(f, columns=["procedure", "assertion", "reference", "observation", "nav_overstatement",
                                  "rating", "recommendation"])
    order = {"High": 0, "Medium": 1, "Low": 2}
    df = df.sort_values("rating", key=lambda s: s.map(order), kind="stable").reset_index(drop=True)
    df.insert(0, "finding", [f"F{i}" for i in range(1, len(df) + 1)])
    df["management_response"] = "[To be obtained]"
    return df


def run(data_dir: Path, **overrides) -> AuditResult:
    raw = load(data_dir)
    settings = {**DEFAULTS, **overrides}
    terms, statement = raw["terms"], raw["statement"]
    vdate = pd.Timestamp(terms["valuation_date"])

    holdings = add_values(raw["holdings"], terms["fx_sgd_per_unit"])
    cust = custodian_reconciliation(holdings, raw["custodian"], raw["pending"], raw["corp_actions"],
                                    raw["prices"], vdate)
    price = price_testing(holdings, raw["prices"], vdate, settings["price_tolerance"], settings["stale_days"])
    fee_daily, fee = fee_recalculation(raw["daily_nav"], terms["management_fee_rate"], terms["day_count"],
                                       statement["accrued_management_fee"])
    nav = recompute_nav(statement, holdings)
    mis, mat = evaluate(statement, cust, price, fee, settings)
    findings = build_findings(cust, price, fee, mat)
    return AuditResult(terms, statement, settings, holdings, cust, price, fee_daily, fee, nav, mis, mat,
                       findings, raw)
