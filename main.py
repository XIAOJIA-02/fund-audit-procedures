"""
Run the fund audit procedures end to end.

    python main.py                          # uses ./data, writes ./output
    python main.py --preparer "Your Name"
    python main.py --tolerance 0.02         # change the price tolerance (2%)
    python main.py --materiality 0.005      # overall materiality = 0.5% of NAV
"""

import argparse
from pathlib import Path

from fundaudit.procedures import DEFAULTS, run
from fundaudit.workpaper import export_workpaper


def main():
    ap = argparse.ArgumentParser(description="Fund audit: custodian rec, price testing, fee recalc, NAV")
    ap.add_argument("--data-dir", default="data")
    ap.add_argument("--output-dir", default="output")
    ap.add_argument("--preparer", default="[Your name]")
    ap.add_argument("--tolerance", type=float, default=DEFAULTS["price_tolerance"])
    ap.add_argument("--materiality", type=float, default=DEFAULTS["materiality_pct"])
    args = ap.parse_args()

    r = run(Path(args.data_dir), price_tolerance=args.tolerance, materiality_pct=args.materiality)
    out = Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)
    wp = export_workpaper(r, out / "fund_audit_workpaper_oct2026.xlsx", preparer=args.preparer)
    r.findings.to_csv(out / "audit_findings.csv", index=False)
    r.misstatements.to_csv(out / "misstatement_schedule.csv", index=False)

    m, f, n = r.materiality, r.fee_summary, r.nav
    print("=" * 66)
    print(f"Fund audit - {r.terms['fund']}")
    print(f"Valuation date {r.terms['valuation_date']}  |  base currency {r.terms['base_currency']}")
    print("=" * 66)
    print(f"Reported NAV                       {r.statement['reported_nav']:>16,.2f}")
    print(f"NAV recomputes?                    {'Yes' if n['nav_agrees'] else 'NO':>16}")
    print(f"Overall materiality ({r.settings['materiality_pct']:.1%})        {m['overall']:>16,.2f}")
    print(f"Performance materiality            {m['performance']:>16,.2f}")
    print(f"Clearly trivial                    {m['clearly_trivial']:>16,.2f}")
    print("-" * 66)
    print("Custodian reconciliation:")
    for s_, cnt in r.custodian_rec["status"].value_counts().items():
        print(f"  {s_:<34}{cnt:>3}")
    print("Price testing:")
    for s_, cnt in r.price_test["status"].value_counts().items():
        print(f"  {s_:<34}{cnt:>3}")
    print(f"Management fee: expected {f['expected_fee']:,.2f}, charged {f['fee_charged']:,.2f}, "
          f"over-accrued {f['over_accrual']:,.2f} (implied rate {f['implied_rate']:.2%})")
    print("-" * 66)
    for row in r.misstatements.itertuples(index=False):
        flag = "" if row.above_clearly_trivial else "  (below trivial)"
        print(f"  {row.ref} {row.reference:<9}{row.nav_overstatement:>14,.2f}{flag}")
    print(f"Accumulated misstatement           {m['accumulated_misstatement']:>16,.2f}  ({m['nav_impact_pct']:.2%} of NAV)")
    print(f"Adjusted NAV per unit              {m['adjusted_nav_per_unit']:>16.4f}  "
          f"(reported {r.statement['reported_nav_per_unit']:.4f})")
    print(f"Evaluation: {m['conclusion']}")
    print(f"Findings: {len(r.findings)}  ({(r.findings['rating'] == 'High').sum()} high)")
    print(f"Workpaper: {wp}")


if __name__ == "__main__":
    main()
