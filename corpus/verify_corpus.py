#!/usr/bin/env python3
"""
Verification of the SBSC-1 corpus.

Recomputes, from the generated artefacts alone, the figures a paper using this
corpus would report, and checks file integrity against the manifest:

  1. Corpus composition          -> the composition table
  2. Per-field coverage c_j      -> the field-mapping and coverage table
  3. Label heterogeneity         -> distinct source column labels per field
  4. Reference reconciliation    -> how many rows and accounts a correct
                                    arithmetic check would flag, given the
                                    injected balance and direction defects
  5. Manifest integrity          -> SHA-256 of every file

Item 4 is a reference, not a requirement. It states what an ideal detector
would find on this corpus, so that a pipeline's own reconciliation numbers can
be compared against a known upper bound.

Usage:
    python verify_corpus.py [--root DIR] [--tolerance 0.01]
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
from collections import Counter, defaultdict

FIELDS = [
    "client_code", "client_identifier", "client_account", "statement_period",
    "operation_date", "direction", "amount", "currency", "balance_after",
    "category", "subcategory", "counterparty_name", "counterparty_identifier",
    "counterparty_account_or_channel", "counterparty_type",
]
MANDATORY = FIELDS[:8]


def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def rule(title):
    print(f"\n{title}\n{'-' * len(title)}")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--root", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--tolerance", type=float, default=0.01)
    args = ap.parse_args()

    gt = os.path.join(args.root, "ground_truth")
    journal = read_csv(os.path.join(gt, "normalized_journal.csv"))
    injection = read_csv(os.path.join(gt, "injection_log.csv"))
    layouts = read_csv(os.path.join(gt, "layout_registry.csv"))
    labels = read_csv(os.path.join(gt, "field_label_index.csv"))

    n = len(journal)
    failures = []

    # ---------------------------------------------------------------- 1
    rule("1. Corpus composition")
    files = {r["source_file_id"] for r in journal}
    accounts = {r["client_account"] for r in journal}
    cps = {r["counterparty_ref"] for r in journal}
    ccy = sorted({r["currency"] for r in journal})
    comp = [
        ("Statement files", len(files)),
        ("Distinct source layouts", len(layouts)),
        ("Normalized transactions", n),
        ("Account holders", len(accounts)),
        ("Counterparties", len(cps)),
        ("Currencies represented", len(ccy)),
    ]
    for k, v in comp:
        print(f"  {k:<32} {v}")
    by_class = Counter(r["defect_class"] for r in injection)
    print(f"  {'Injected defects, total':<32} {len(injection)}")
    for k, v in sorted(by_class.items()):
        print(f"    {k:<30} {v}")
    lvl = Counter(r["level"] for r in injection)
    print(f"  {'Row-level / file-level':<32} {lvl['row']} / {lvl['file']}")

    # ---------------------------------------------------------------- 2
    rule("2. Per-field coverage c_j")
    print(f"  {'field':<34}{'mand.':<7}{'labels':<8}{'present':<9}{'c_j'}")
    label_counts = Counter(r["normalized_field"] for r in labels)
    for f in FIELDS:
        present = sum(1 for r in journal if r[f] != "")
        cj = present / n
        mand = "yes" if f in MANDATORY else "no"
        print(f"  {f:<34}{mand:<7}{label_counts[f]:<8}{present:<9}{cj:.3f}")
        if f in MANDATORY and present != n:
            failures.append(f"mandatory field {f} has coverage {cj:.3f} < 1.000")

    # ---------------------------------------------------------------- 3
    rule("3. Label heterogeneity")
    total = sum(label_counts[f] for f in FIELDS)
    print(f"  distinct source column labels in the corpus: {total}")
    worst = sorted(FIELDS, key=lambda f: -label_counts[f])[:3]
    for f in worst:
        print(f"    {f:<34}{label_counts[f]} labels")

    # ---------------------------------------------------------------- 4
    rule("4. Reference arithmetic reconciliation")
    flipped = {int(r["journal_index"]) for r in injection
               if r["defect_class"] == "DIRECTION_INCONSISTENCY"}
    offset = {int(r["journal_index"]): r["detail"] for r in injection
              if r["defect_class"] == "BALANCE_INCONSISTENCY"}

    per_account = defaultdict(list)
    for r in journal:
        per_account[r["client_account"]].append(r)

    exceed_rows = 0
    exceed_accounts = set()
    scored_rows = 0
    residuals = []
    for acct, rows in per_account.items():
        rows.sort(key=lambda r: (r["source_file_id"], int(r["source_row"])))
        prev = None
        for r in rows:
            if r["balance_after"] == "":
                prev = None
                continue
            bal = float(r["balance_after"])
            idx = int(r["journal_index"])
            printed = round(bal + float(r["printed_balance_delta"]), 4)
            amount = float(r["amount"])
            signed = amount if r["direction"] == "credit" else -amount
            if idx in flipped:
                signed = -signed
            if prev is not None:
                eps = abs(round(printed - (prev + signed), 4))
                residuals.append(eps)
                scored_rows += 1
                if eps > args.tolerance:
                    exceed_rows += 1
                    exceed_accounts.add(acct)
            prev = printed

    residuals.sort()

    def q(p):
        if not residuals:
            return 0.0
        return residuals[min(len(residuals) - 1, int(p * len(residuals)))]

    with_bal = sum(1 for r in journal if r["balance_after"] != "")
    print(f"  rows carrying balance_after            {with_bal} of {n}")
    print(f"  rows scorable (balance on both legs)   {scored_rows}")
    print(f"  tolerance tau_b                        {args.tolerance}")
    print(f"  rows exceeding tolerance               {exceed_rows}")
    print(f"  accounts raising a warning             "
          f"{len(exceed_accounts)} of {len(per_account)}")
    print(f"  residual quantiles  q50 / q90 / q99    "
          f"{q(0.50):.3f} / {q(0.90):.3f} / {q(0.99):.2f}")
    print("  note: direction defects placed in files without a balance column "
          "are\n        invisible to this check by construction, which is "
          "intended.")

    undetectable = [r for r in injection
                    if r["defect_class"] == "DIRECTION_INCONSISTENCY"
                    and "undetectable" in r["detail"]]
    print(f"  direction defects undetectable by construction: "
          f"{len(undetectable)} of {by_class['DIRECTION_INCONSISTENCY']}")

    # ---------------------------------------------------------------- 5
    rule("5. Manifest integrity")
    man = os.path.join(gt, "manifest_sha256.txt")
    checked = bad = missing = 0
    with open(man, encoding="utf-8") as fh:
        for line in fh:
            h, rel = line.strip().split("  ", 1)
            p = os.path.join(args.root, rel)
            if not os.path.exists(p):
                missing += 1
                continue
            checked += 1
            if sha256(p) != h:
                bad += 1
                failures.append(f"digest mismatch: {rel}")
    print(f"  files checked {checked}, mismatched {bad}, missing {missing}")

    rule("Result")
    if failures:
        for f in failures:
            print(f"  FAIL  {f}")
        raise SystemExit(1)
    print("  all checks passed")


if __name__ == "__main__":
    main()
