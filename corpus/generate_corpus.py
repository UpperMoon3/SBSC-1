#!/usr/bin/env python3
"""
SBSC-1: Synthetic Bank Statement Corpus for normalization research.

Companion artefact to:
    "A Reproducible Normalization Model for Forensic Analysis of
     Heterogeneous Bank Statements"

The corpus is generated, not collected. It contains no real persons,
organizations, account numbers, bank templates or case material. Every value
is drawn from a fixed seed, so the corpus is byte-reproducible from this file
alone.

What it produces
----------------
statements/          30 statement files across 15 layouts and 3 synthetic issuers
ground_truth/
    normalized_journal.csv   the 1200 records a correct pipeline must produce
    injection_log.csv        the 57 injected defects, with type and position
    layout_registry.csv      the documented field mapping M_b for each layout
    field_label_index.csv    every source column label and the field it maps to
    manifest_sha256.txt      SHA-256 of every generated file
reports/
    corpus_summary.json      the corpus composition table
    coverage.csv             per-field coverage c_j

Usage
-----
    python generate_corpus.py [--out DIR] [--seed N]

Dependencies: numpy, openpyxl (standard library otherwise).
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import datetime
import zipfile
import io
import json
import os
import random
import re
from collections import Counter, defaultdict
from datetime import date, timedelta

import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font

# ----------------------------------------------------------------------
# Corpus specification. These constants define the published composition.
# ----------------------------------------------------------------------

SPEC = {
    "files": 30,
    "layouts": 15,
    "issuers": 3,
    "rows": 1200,            # normalized transactions, duplicates included
    "subjects": 12,          # account holders, one account each
    "counterparties": 187,
    "currencies": ["KZT", "USD", "EUR"],
    "period_months": 6,
    "period_start": date(2025, 1, 1),
}

# Injected defects. Sum = 57; 53 row-level + 4 file-level.
DEFECTS = {
    "DUPLICATE_ROW": 14,
    "DIRECTION_INCONSISTENCY": 9,     # 2 of these placed in files with no balance column
    "BALANCE_INCONSISTENCY": 7,
    "INCOMPLETE_COUNTERPARTY": 11,    # 2 of these retain a channel token
    "CURRENCY_MIXING": 6,
    "UNMAPPED_FIELD": 6,
    "UNPARSED_LAYOUT": 4,             # file-level
}

# Number of DISTINCT source column labels each normalized field appears under
# across the 15 layouts. This is the label-heterogeneity column of the paper.
LABEL_COUNTS = {
    "client_code": 4,
    "client_identifier": 5,
    "client_account": 3,
    "statement_period": 3,
    "operation_date": 11,
    "direction": 7,
    "amount": 9,
    "currency": 5,
    "balance_after": 6,
    "category": 8,
    "subcategory": 14,
    "counterparty_name": 10,
    "counterparty_identifier": 6,
    "counterparty_account_or_channel": 9,
    "counterparty_type": 5,
}

MANDATORY = [
    "client_code", "client_identifier", "client_account", "statement_period",
    "operation_date", "direction", "amount", "currency",
]

OPTIONAL_COVERAGE = {          # target c_j over the 1200 journal rows
    "balance_after": 0.920,
    "category": 0.823,
    "subcategory": 0.702,
    "counterparty_name": 0.892,
    "counterparty_identifier": 0.578,
    "counterparty_account_or_channel": 0.765,
    "counterparty_type": 0.892,
}

CONTROL_FIELDS = [
    "source_file_id", "source_sheet", "source_row",
    "processing_version", "defect_flag",
]

PROCESSING_VERSION = "sbsc-1.1.0"

# ----------------------------------------------------------------------
# Label pools. Deliberately varied in language, casing and wording, so that
# the mapping M_b is a genuine mapping rather than an identity relabelling.
# ----------------------------------------------------------------------

LABEL_POOL = {
    "client_code": [
        "Client code", "Customer ID", "Klient kodu", "Holder ref",
    ],
    "client_identifier": [
        "Client identifier", "Holder tax ref", "Customer no.", "IIN/BIN",
        "Party ref",
    ],
    "client_account": [
        "Account", "Account number", "Acct.",
    ],
    "statement_period": [
        "Statement period", "Reporting period", "Period",
    ],
    "operation_date": [
        "Date", "Operation date", "Transaction date", "Value date",
        "Posting date", "Date of operation", "Dt", "Oper. date",
        "Trans. date", "Date/time", "Booking date",
    ],
    "direction": [
        "Direction", "Type", "Dr/Cr", "Operation type", "Flow",
        "Debit/Credit", "Sign",
    ],
    "amount": [
        "Amount", "Sum", "Transaction amount", "Operation amount", "Value",
        "Amt", "Turnover", "Amount, total", "Sum of operation",
    ],
    "currency": [
        "Currency", "Curr.", "CCY", "Currency code", "Val.",
    ],
    "balance_after": [
        "Balance", "Balance after", "Running balance", "Closing balance",
        "Balance after operation", "Bal.",
    ],
    "category": [
        "Category", "Operation category", "Class", "Group", "Type of expense",
        "Payment category", "Cat.", "Classification",
    ],
    "subcategory": [
        "Subcategory", "Sub-category", "Detail", "Subgroup", "Sub class",
        "Operation detail", "Subcat.", "Detailed category", "Sub type",
        "Purpose group", "Detail level 2", "Sub-class", "Item", "Sub-item",
    ],
    "counterparty_name": [
        "Counterparty", "Counterparty name", "Payer/Payee", "Beneficiary",
        "Correspondent", "Name of counterparty", "Partner",
        "Sender/Recipient", "Contragent", "Payee name",
    ],
    "counterparty_identifier": [
        "Counterparty ID", "Counterparty tax ref", "BIN/IIN of counterparty",
        "Partner ID", "Correspondent ref", "CP identifier",
    ],
    "counterparty_account_or_channel": [
        "Counterparty account", "Beneficiary account", "Channel",
        "Correspondent account", "CP account", "Account of counterparty",
        "Card/Channel", "Payment channel", "Destination",
    ],
    "counterparty_type": [
        "Counterparty type", "Partner type", "CP type", "Entity type", "Kind",
    ],
}

ISSUERS = ["Bank A", "Bank B", "Bank C"]

# Direction encodings used across layouts.
#   split  : separate Debit and Credit amount columns
#   signed : one signed amount column, no direction column
#   flag   : unsigned amount column plus a direction flag column
DIRECTION_ENCODINGS = ["split", "signed", "flag"]

CATEGORIES = {
    "Transfer": ["Internal transfer", "Outgoing transfer", "Incoming transfer"],
    "Goods and services": ["Retail", "Utilities", "Telecom", "Fuel"],
    "Payroll": ["Salary", "Bonus", "Advance"],
    "Cash": ["ATM withdrawal", "Cash deposit"],
    "Asset": ["Real estate", "Vehicle", "Deposit placement", "Securities"],
    "Fees": ["Bank fee", "Service charge"],
    "Loan": ["Loan repayment", "Loan disbursement"],
    "Tax": ["Tax payment", "Social contribution"],
}

CP_TYPES = ["company", "individual", "bank", "state", "marketplace",
            "utility", "payroll_aggregator"]

# Counterparty types that structurally lack a registered identifier.
NO_IDENTIFIER_TYPES = {"individual"}


# ----------------------------------------------------------------------
# Helpers
# ----------------------------------------------------------------------

def sha256(path: str) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def exact_mask(n: int, present: int, rng: random.Random) -> list[bool]:
    """Boolean mask of length n with exactly `present` True values."""
    idx = list(range(n))
    rng.shuffle(idx)
    keep = set(idx[:present])
    return [i in keep for i in range(n)]


def assign_labels(field: str, rng: random.Random) -> list[str]:
    """One label per layout, using exactly LABEL_COUNTS[field] distinct labels."""
    k = LABEL_COUNTS[field]
    pool = LABEL_POOL[field]
    assert len(pool) >= k, f"label pool for {field} too small"
    chosen = pool[:k]
    out = [chosen[i % k] for i in range(SPEC["layouts"])]
    rng.shuffle(out)
    # shuffling cannot lose a label because every label appears at least once
    # in the pre-shuffle list (15 >= k for every field)
    assert len(set(out)) == k, f"{field}: {len(set(out))} != {k}"
    return out


# ----------------------------------------------------------------------
# Entity construction
# ----------------------------------------------------------------------

def build_counterparties(rng: random.Random) -> list[dict]:
    """187 synthetic counterparties, including shared settlement hubs."""
    cps = []
    # 7 mass-payment hubs shared across subjects
    hubs = [
        ("Bank A", "bank"), ("Bank B", "bank"), ("Bank C", "bank"),
        ("State Revenue Office", "state"), ("City Utilities", "utility"),
        ("Marketplace Platform", "marketplace"),
        ("Payroll Aggregator", "payroll_aggregator"),
    ]
    for i, (name, t) in enumerate(hubs):
        cps.append({
            "cp_id": f"CP-{i+1:03d}",
            "name": f"SYNTH {name}",
            "type": t,
            "identifier": f"BIN{700000000 + i:09d}",
            "is_hub": True,
        })
    n_rest = SPEC["counterparties"] - len(hubs)
    for i in range(n_rest):
        j = len(hubs) + i + 1
        t = rng.choices(
            ["company", "individual", "marketplace", "utility"],
            weights=[0.55, 0.31, 0.08, 0.06],
        )[0]
        cps.append({
            "cp_id": f"CP-{j:03d}",
            "name": (f"SYNTH Individual {j:03d}" if t == "individual"
                     else f"SYNTH Entity {j:03d} LLP"),
            "type": t,
            "identifier": (None if t in NO_IDENTIFIER_TYPES
                           else f"BIN{800000000 + j:09d}"),
            "is_hub": False,
        })
    assert len(cps) == SPEC["counterparties"]
    return cps


def build_accounts(rng: random.Random) -> list[dict]:
    """12 subjects, one account each, distributed over 3 issuers."""
    ccy_plan = ["KZT"] * 8 + ["USD"] * 2 + ["EUR"] * 2
    accounts = []
    for i in range(SPEC["subjects"]):
        issuer = i % SPEC["issuers"]
        accounts.append({
            "client_code": f"SUBJ-{i+1:03d}",
            "client_identifier": f"IIN{900000000 + i:09d}",
            "client_account": f"ACC-{i+1:03d}",
            "issuer": issuer,
            "currency": ccy_plan[i],
            "opening_balance": round(rng.uniform(50_000, 900_000), 2),
        })
    return accounts


def build_layouts(rng: random.Random) -> list[dict]:
    """15 layouts, 5 per issuer, each with its own labels and conventions."""
    labels = {f: assign_labels(f, rng) for f in LABEL_COUNTS}
    layouts = []
    for b in range(SPEC["layouts"]):
        layouts.append({
            "layout_id": f"L{b+1:02d}",
            "issuer": b // 5,
            "labels": {f: labels[f][b] for f in LABEL_COUNTS},
            "direction_encoding": DIRECTION_ENCODINGS[b % 3],
            "has_preamble": b % 4 == 0,
            "has_summary_rows": b % 5 == 0,
            "two_row_header": b % 7 == 3,
            "date_format": ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y",
                            "%m/%d/%Y"][b % 4],
            "decimal_comma": b % 3 == 1,
            "has_currency_column": b % 6 != 2,
            "has_balance_column": True,       # overridden below
        })
    # three layouts publish no running balance at all
    for b in (4, 9, 13):
        layouts[b]["has_balance_column"] = False
    # v1.0 added a sub-unit "display rounding" offset to the printed balance
    # of four layouts; balances are written with two decimals, so the offset
    # never reached the files and existed only in the ground truth. It is
    # removed in v1.1 so that the ground truth matches what the files print.
    for b in range(SPEC["layouts"]):
        layouts[b]["balance_jitter"] = 0.0
    return layouts


def build_files(accounts, layouts, rng: random.Random) -> list[dict]:
    """30 files: each layout used exactly twice; files tied to their issuer."""
    by_issuer = defaultdict(list)
    for b, lay in enumerate(layouts):
        by_issuer[lay["issuer"]].extend([b, b])   # two files per layout

    files = []
    fid = 0
    for issuer in range(SPEC["issuers"]):
        accs = [a for a in accounts if a["issuer"] == issuer]
        slots = by_issuer[issuer]
        rng.shuffle(slots)
        # 4 accounts per issuer produce 10 files: two accounts x 3, two x 2
        counts = [3, 3, 2, 2]
        k = 0
        for acc, cnt in zip(accs, counts):
            span = SPEC["period_months"] // cnt
            for part in range(cnt):
                start = SPEC["period_start"] + timedelta(days=30 * span * part)
                end = start + timedelta(days=30 * span - 1)
                fid += 1
                files.append({
                    "source_file_id": f"F{fid:03d}",
                    "layout": slots[k],
                    "account": acc,
                    "period_start": start,
                    "period_end": end,
                })
                k += 1
    assert len(files) == SPEC["files"], len(files)
    assert Counter(f["layout"] for f in files) == \
        Counter({b: 2 for b in range(SPEC["layouts"])})
    return files


# ----------------------------------------------------------------------
# Transaction construction
# ----------------------------------------------------------------------

def build_transactions(files, counterparties, rng: random.Random) -> list[dict]:
    """1186 unique transactions; 14 duplicates are added afterwards."""
    n_unique = SPEC["rows"] - DEFECTS["DUPLICATE_ROW"]
    per_file = [n_unique // len(files)] * len(files)
    for i in range(n_unique - sum(per_file)):
        per_file[i] += 1

    hubs = [c for c in counterparties if c["is_hub"]]
    normal = [c for c in counterparties if not c["is_hub"]]

    rows = []
    for f, n in zip(files, per_file):
        acc = f["account"]
        span = (f["period_end"] - f["period_start"]).days
        days = sorted(rng.randint(0, span) for _ in range(n))
        for d in days:
            cp = rng.choice(hubs) if rng.random() < 0.22 else rng.choice(normal)
            cat = rng.choice(list(CATEGORIES))
            sub = rng.choice(CATEGORIES[cat])
            direction = "credit" if rng.random() < 0.42 else "debit"
            if cat == "Payroll":
                direction = "credit"
            if cat in ("Asset", "Tax", "Fees"):
                direction = "debit"
            amount = round(abs(rng.lognormvariate(10.2, 1.05)), 2)
            if cat == "Asset":
                amount = round(amount * 4, 2)
            rows.append({
                "client_code": acc["client_code"],
                "client_identifier": acc["client_identifier"],
                "client_account": acc["client_account"],
                "statement_period": (f"{f['period_start'].isoformat()}/"
                                     f"{f['period_end'].isoformat()}"),
                "operation_date": f["period_start"] + timedelta(days=d),
                "direction": direction,
                "amount": amount,
                "currency": acc["currency"],
                "balance_after": None,           # filled after ordering
                "category": cat,
                "subcategory": sub,
                "counterparty_name": cp["name"],
                "_cp_name_full": cp["name"],
                "counterparty_identifier": cp["identifier"],
                "counterparty_account_or_channel":
                    f"CHN-{cp['cp_id'][3:]}" if cp["type"] == "individual"
                    else f"ACCT-{cp['cp_id'][3:]}",
                "counterparty_type": cp["type"],
                "source_file_id": f["source_file_id"],
                "source_sheet": "Sheet1",
                "source_row": None,              # filled at render time
                "processing_version": PROCESSING_VERSION,
                "defect_flag": "",
                "_cp_id": cp["cp_id"],
                "_is_duplicate_of": "",
            })
    return rows


def ensure_all_counterparties(rows, counterparties, rng: random.Random):
    """Guarantee that every generated counterparty appears at least once, so
    the corpus really contains the number of entities it declares."""
    used = {r["_cp_id"] for r in rows}
    unused = [c for c in counterparties if c["cp_id"] not in used]
    if not unused:
        return
    # reassign rows whose counterparty is already well represented
    counts = Counter(r["_cp_id"] for r in rows)
    spare = [r for r in rows
             if counts[r["_cp_id"]] > 1 and not r["_cp_id"].startswith("CP-00")]
    for cp, r in zip(unused, rng.sample(spare, len(unused))):
        r["_cp_id"] = cp["cp_id"]
        r["counterparty_name"] = cp["name"]
        r["counterparty_identifier"] = cp["identifier"]
        r["counterparty_type"] = cp["type"]
        r["counterparty_account_or_channel"] = (
            f"CHN-{cp['cp_id'][3:]}" if cp["type"] == "individual"
            else f"ACCT-{cp['cp_id'][3:]}")


def insert_duplicates(rows, rng: random.Random, injection: list) -> list[dict]:
    """Duplicate 14 rows in place, so the journal holds 1200 records."""
    out = list(rows)
    picks = rng.sample(range(len(rows)), DEFECTS["DUPLICATE_ROW"])
    for p in sorted(picks, reverse=True):
        clone = dict(out[p])
        clone["_is_duplicate_of"] = "pending"
        out.insert(p + 1, clone)
    for i, r in enumerate(out):
        r["_journal_index"] = i
    for i, r in enumerate(out):
        if r["_is_duplicate_of"] == "pending":
            r["_is_duplicate_of"] = str(out[i - 1]["_journal_index"])
            injection.append({
                "defect_id": "",
                "defect_class": "DUPLICATE_ROW",
                "level": "row",
                "source_file_id": r["source_file_id"],
                "journal_index": i,
                "detail": f"copy of journal index {r['_is_duplicate_of']}",
            })
    assert len(out) == SPEC["rows"], len(out)
    return out


def order_and_balance(rows, files, layouts):
    """Sort per file by date, assign source_row and a consistent running balance."""
    by_file = defaultdict(list)
    for r in rows:
        by_file[r["source_file_id"]].append(r)

    balances = {}
    for f in files:
        acc = f["account"]
        balances.setdefault(acc["client_account"], acc["opening_balance"])

    for f in sorted(files, key=lambda x: (x["account"]["client_account"],
                                          x["period_start"])):
        rs = sorted(by_file[f["source_file_id"]],
                    key=lambda r: (r["operation_date"], r["_journal_index"]))
        acc_key = f["account"]["client_account"]
        bal = balances[acc_key]
        for i, r in enumerate(rs, start=1):
            r["source_row"] = i
            signed = r["amount"] if r["direction"] == "credit" else -r["amount"]
            bal = round(bal + signed, 2)
            r["balance_after"] = bal
        balances[acc_key] = bal
        by_file[f["source_file_id"]] = rs
    return by_file


# ----------------------------------------------------------------------
# Coverage masks and defect injection
# ----------------------------------------------------------------------

def apply_coverage(rows, by_file, files, layouts, rng: random.Random):
    """
    Set the seven optional fields to None so that each reaches its target
    coverage exactly. Missingness is structured, not uniform noise.
    """
    n = len(rows)
    idx = {r["_journal_index"]: r for r in rows}

    # balance_after: absent for every row of a file whose layout publishes no
    # running balance, so that the ground truth never holds a balance the
    # statement file does not print (v1.0 topped this up to a fixed target
    # and left balances in the journal for four files that print none).
    for f in files:
        if not layouts[f["layout"]]["has_balance_column"]:
            for r in by_file[f["source_file_id"]]:
                r["balance_after"] = None

    # counterparty_identifier: individuals structurally lack one, then trim
    # or extend to hit the target exactly.
    tgt_present = round(OPTIONAL_COVERAGE["counterparty_identifier"] * n)
    have = [r for r in rows if r["counterparty_identifier"] is not None]
    if len(have) > tgt_present:
        for r in rng.sample(have, len(have) - tgt_present):
            r["counterparty_identifier"] = None
    else:
        lack = [r for r in rows if r["counterparty_identifier"] is None]
        for r in rng.sample(lack, tgt_present - len(have)):
            r["counterparty_identifier"] = f"BIN{810000000 + r['_journal_index']:09d}"

    # remaining optional fields: exact random masks
    for field in ("category", "subcategory", "counterparty_name",
                  "counterparty_account_or_channel", "counterparty_type"):
        tgt = round(OPTIONAL_COVERAGE[field] * n)
        mask = exact_mask(n, tgt, rng)
        for i, keep in enumerate(mask):
            if not keep:
                idx[i][field] = None

    # an issuer prints at least a name or an identifier for every
    # counterparty; a record lacking both is a defect and is only ever
    # injected (INCOMPLETE_COUNTERPARTY), never produced by a coverage mask
    for r in rows:
        if r["counterparty_name"] is None and r["counterparty_identifier"] is None:
            r["counterparty_name"] = r["_cp_name_full"]

    # subcategory cannot survive a missing category
    for r in rows:
        if r["category"] is None and r["subcategory"] is not None:
            r["subcategory"] = None
    # restore the exact subcategory count after that dependency pass
    tgt = round(OPTIONAL_COVERAGE["subcategory"] * n)
    have = [r for r in rows if r["subcategory"] is not None]
    if len(have) < tgt:
        cands = [r for r in rows
                 if r["subcategory"] is None and r["category"] is not None]
        for r in rng.sample(cands, min(tgt - len(have), len(cands))):
            r["subcategory"] = rng.choice(CATEGORIES[r["category"]])
    elif len(have) > tgt:
        for r in rng.sample(have, len(have) - tgt):
            r["subcategory"] = None


def inject_row_defects(rows, by_file, files, layouts, rng: random.Random,
                       injection: list):
    """
    Inject the remaining row-level defect classes. The ground-truth journal
    keeps the correct value; the rendered statement file carries the defect.
    """
    idx_by_file = {f["source_file_id"]: f for f in files}

    def has_balance(r):
        return r["balance_after"] is not None

    # ------------------------------------------------------------------
    # Balance-bearing defects are concentrated in five accounts. Clustering
    # is realistic (an export fault affects the account it was raised on)
    # and it is what the published corpus exhibits.
    #
    # Two sub-modes of BALANCE_INCONSISTENCY are used:
    #   isolated        a single mistyped balance cell -> two residuals
    #   carried_forward the balance drifts from that row to the end of the
    #                   account -> one residual
    # Magnitudes are fixed constants rather than draws, so that the residual
    # distribution of the corpus is reproducible exactly.
    # ------------------------------------------------------------------
    CARRIED_DELTAS = [1.45, 1.20, 0.95, 0.87, 0.66]
    ISOLATED_DELTAS = [0.50, 0.40]

    per_account = defaultdict(list)
    for r in rows:
        per_account[r["client_account"]].append(r)
    for acct, rs in per_account.items():
        rs.sort(key=lambda r: (r["source_file_id"], r["source_row"]))

    def scorable(rs):
        """Rows whose residual is computable: this row and the previous one
        both carry a balance."""
        out = []
        for i in range(1, len(rs)):
            if (rs[i]["balance_after"] is not None
                    and rs[i - 1]["balance_after"] is not None
                    and not rs[i]["defect_flag"]):
                out.append(i)
        return out

    ranked = sorted(per_account.items(),
                    key=lambda kv: -len(scorable(kv[1])))
    defective = [a for a, _ in ranked[:5]]

    # --- BALANCE_INCONSISTENCY: 7, one carried_forward per defective
    #     account plus two isolated ---
    for acct, delta in zip(defective, CARRIED_DELTAS):
        rs = per_account[acct]
        i = rng.choice(scorable(rs)[2:-2])
        r = rs[i]
        for later in rs[i:]:
            if later["balance_after"] is not None:
                later["_printed_delta"] += delta
        r["defect_flag"] = "BALANCE_INCONSISTENCY"
        injection.append({
            "defect_id": "", "defect_class": "BALANCE_INCONSISTENCY",
            "level": "row", "source_file_id": r["source_file_id"],
            "journal_index": r["_journal_index"],
            "detail": f"carried_forward; printed balance drifts by {delta} "
                      f"from this row to the end of the account",
        })
    for acct, delta in zip(defective[:2], ISOLATED_DELTAS):
        rs = per_account[acct]
        cand = [i for i in scorable(rs)[2:-2] if i + 1 in scorable(rs)]
        i = rng.choice(cand)
        r = rs[i]
        r["_printed_delta"] += delta
        r["defect_flag"] = "BALANCE_INCONSISTENCY"
        injection.append({
            "defect_id": "", "defect_class": "BALANCE_INCONSISTENCY",
            "level": "row", "source_file_id": r["source_file_id"],
            "journal_index": r["_journal_index"],
            "detail": f"isolated; single printed balance cell offset by {delta}",
        })

    # --- DIRECTION_INCONSISTENCY: 9, of which 7 in the defective accounts
    #     on reconcilable rows and 2 in rows with no balance column ---
    picks = []
    per_acct_share = [2, 2, 1, 1, 1]
    for acct, share in zip(defective, per_acct_share):
        rs = per_account[acct]
        cand = [i for i in scorable(rs)[2:-2] if not rs[i]["defect_flag"]]
        for i in rng.sample(cand, share):
            picks.append(rs[i])
    no_bal = [r for r in rows
              if r["balance_after"] is None and not r["defect_flag"]]
    picks += rng.sample(no_bal, 2)
    for r in picks:
        r["_render_direction_flipped"] = True
        r["defect_flag"] = "DIRECTION_INCONSISTENCY"
        injection.append({
            "defect_id": "", "defect_class": "DIRECTION_INCONSISTENCY",
            "level": "row", "source_file_id": r["source_file_id"],
            "journal_index": r["_journal_index"],
            "detail": ("undetectable by reconciliation: no balance column"
                       if r["balance_after"] is None
                       else "sign inverted in the file"),
        })

    # --- INCOMPLETE_COUNTERPARTY: 11, of which 2 keep a channel token ---
    pool = [r for r in rows if r["counterparty_name"] is not None
            and not r.get("_render_direction_flipped")
            and not r.get("_render_balance_delta")]
    picks = rng.sample(pool, DEFECTS["INCOMPLETE_COUNTERPARTY"])
    for i, r in enumerate(picks):
        keep_channel = i < 2
        r["_render_blank_cp"] = "channel_only" if keep_channel else "full"
        r["defect_flag"] = "INCOMPLETE_COUNTERPARTY"
        injection.append({
            "defect_id": "", "defect_class": "INCOMPLETE_COUNTERPARTY",
            "level": "row", "source_file_id": r["source_file_id"],
            "journal_index": r["_journal_index"],
            "detail": ("name and identifier both blank; channel token retained"
                       if keep_channel else "name and identifier both blank"),
        })

    # --- CURRENCY_MIXING: 6 rows of a foreign currency in a sheet that
    #     publishes no currency column ---
    cand_files = [f for f in files
                  if not layouts[f["layout"]]["has_currency_column"]]
    picks = []
    for f in cand_files:
        picks.extend(by_file[f["source_file_id"]])
    picks = [r for r in picks if not r["defect_flag"]]
    for r in rng.sample(picks, DEFECTS["CURRENCY_MIXING"]):
        foreign = rng.choice([c for c in SPEC["currencies"]
                              if c != r["currency"]])
        home = r["currency"]
        r["_render_currency"] = foreign
        # the operation is in the foreign currency; the sheet states it only
        # as a code printed next to the amount, which a reader that assigns
        # the account currency to every row of the sheet will miss
        r["currency"] = foreign
        r["defect_flag"] = "CURRENCY_MIXING"
        injection.append({
            "defect_id": "", "defect_class": "CURRENCY_MIXING",
            "level": "row", "source_file_id": r["source_file_id"],
            "journal_index": r["_journal_index"],
            "detail": (f"row is in {foreign}; sheet has no currency column "
                       f"and is otherwise {home}"),
        })

    # --- UNMAPPED_FIELD: 6 rows carrying a value in a column whose header
    #     matches no known label ---
    pool = [r for r in rows if not r["defect_flag"]]
    for r in rng.sample(pool, DEFECTS["UNMAPPED_FIELD"]):
        r["_render_unmapped"] = f"REF{rng.randint(10**6, 10**7 - 1)}"
        r["defect_flag"] = "UNMAPPED_FIELD"
        injection.append({
            "defect_id": "", "defect_class": "UNMAPPED_FIELD",
            "level": "row", "source_file_id": r["source_file_id"],
            "journal_index": r["_journal_index"],
            "detail": "value present in an unlabelled trailing column",
        })


def inject_file_defects(files, layouts, rng: random.Random, injection: list):
    """UNPARSED_LAYOUT: 4 files whose transaction table resists naive parsing."""
    modes = ["metadata_preamble", "summary_rows_before_table",
             "two_row_header", "trailing_totals_block"]
    # an obstacle is a departure from the file's documented layout, so it is
    # only placed in a file whose layout does not already have that feature
    native = {"metadata_preamble": "has_preamble",
              "summary_rows_before_table": "has_summary_rows",
              "two_row_header": "two_row_header"}
    picks, used = [], set()
    for mode in modes:
        feat = native.get(mode)
        elig = [f for f in files if f["source_file_id"] not in used
                and not (feat and layouts[f["layout"]][feat])]
        f = rng.choice(elig)
        used.add(f["source_file_id"])
        picks.append(f)
    for f, mode in zip(picks, modes):
        f["_unparsed_mode"] = mode
        injection.append({
            "defect_id": "", "defect_class": "UNPARSED_LAYOUT",
            "level": "file", "source_file_id": f["source_file_id"],
            "journal_index": "", "detail": mode,
        })
    return picks


# ----------------------------------------------------------------------
# Rendering
# ----------------------------------------------------------------------

def fmt_amount(value, decimal_comma: bool) -> str:
    s = f"{value:,.2f}"
    if decimal_comma:
        s = s.replace(",", "\u0001").replace(".", ",").replace("\u0001", " ")
    return s


def render_file(f, layout, rows, out_dir) -> str:
    """Write one statement file in its layout's own dialect."""
    lab = layout["labels"]
    enc = layout["direction_encoding"]
    mode = f.get("_unparsed_mode")

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    r = 1

    if layout["has_preamble"] or mode == "metadata_preamble":
        acc = f["account"]
        for k, v in [
            ("Issuer", f"SYNTH {ISSUERS[layout['issuer']]}"),
            (lab["client_code"], acc["client_code"]),
            (lab["client_identifier"], acc["client_identifier"]),
            (lab["client_account"], acc["client_account"]),
            (lab["statement_period"],
             f"{f['period_start'].isoformat()} - {f['period_end'].isoformat()}"),
            ("Generated", "synthetic corpus SBSC-1, no real data"),
        ]:
            ws.cell(row=r, column=1, value=k).font = Font(bold=True)
            ws.cell(row=r, column=2, value=v)
            r += 1
        r += 1

    if mode == "summary_rows_before_table" or layout["has_summary_rows"]:
        ws.cell(row=r, column=1, value="Summary").font = Font(bold=True)
        r += 1
        ws.cell(row=r, column=1, value="Total credit")
        ws.cell(row=r, column=2,
                value=sum(x["amount"] for x in rows if x["direction"] == "credit"))
        r += 1
        ws.cell(row=r, column=1, value="Total debit")
        ws.cell(row=r, column=2,
                value=sum(x["amount"] for x in rows if x["direction"] == "debit"))
        r += 2

    # ---- header ----
    cols = [lab["operation_date"]]
    if enc == "split":
        cols += [f"{lab['amount']} (Dr)", f"{lab['amount']} (Cr)"]
    elif enc == "signed":
        cols += [lab["amount"]]
    else:
        cols += [lab["direction"], lab["amount"]]
    if layout["has_currency_column"]:
        cols += [lab["currency"]]
    if layout["has_balance_column"]:
        cols += [lab["balance_after"]]
    cols += [lab["counterparty_name"], lab["counterparty_identifier"],
             lab["counterparty_account_or_channel"], lab["counterparty_type"],
             lab["category"], lab["subcategory"]]
    if any(x.get("_render_unmapped") for x in rows):
        cols += [""]                       # deliberately unlabelled column

    header_row = r
    if layout["two_row_header"] or mode == "two_row_header":
        ws.cell(row=r, column=1,
                value=f"Statement of {f['account']['client_account']}")
        r += 1
        header_row = r
    for c, name in enumerate(cols, start=1):
        cell = ws.cell(row=header_row, column=c, value=name)
        cell.font = Font(bold=True)
        cell.alignment = Alignment(wrap_text=True)
    r = header_row + 1

    # ---- data ----
    for row in rows:
        direction = row["direction"]
        if row.get("_render_direction_flipped"):
            direction = "debit" if direction == "credit" else "credit"
        bal = row["balance_after"]
        if bal is not None:
            bal = round(bal + row["_printed_delta"], 4)
        ccy = row.get("_render_currency") or row["currency"]

        vals = [row["operation_date"].strftime(layout["date_format"])]
        if enc == "split":
            vals += [fmt_amount(row["amount"], layout["decimal_comma"])
                     if direction == "debit" else None,
                     fmt_amount(row["amount"], layout["decimal_comma"])
                     if direction == "credit" else None]
        elif enc == "signed":
            signed = row["amount"] if direction == "credit" else -row["amount"]
            vals += [fmt_amount(signed, layout["decimal_comma"])]
        else:
            vals += [("CR" if direction == "credit" else "DR"),
                     fmt_amount(row["amount"], layout["decimal_comma"])]
        if not layout["has_currency_column"] and row.get("_render_currency"):
            vals[-1] = f"{vals[-1]} {ccy}"
        if layout["has_currency_column"]:
            vals += [ccy]
        if layout["has_balance_column"]:
            vals += [None if bal is None
                     else fmt_amount(bal, layout["decimal_comma"])]

        blank = row.get("_render_blank_cp")
        vals += [
            None if blank else row["counterparty_name"],
            None if blank else row["counterparty_identifier"],
            (row["counterparty_account_or_channel"]
             if blank != "full" else None),
            None if blank else row["counterparty_type"],
            row["category"],
            row["subcategory"],
        ]
        if any(x.get("_render_unmapped") for x in rows):
            vals += [row.get("_render_unmapped")]

        for c, v in enumerate(vals, start=1):
            ws.cell(row=r, column=c, value=v)
        r += 1

    if mode == "trailing_totals_block":
        r += 1
        ws.cell(row=r, column=1, value="TOTALS").font = Font(bold=True)
        ws.cell(row=r, column=2, value=sum(x["amount"] for x in rows))

    for c in range(1, len(cols) + 1):
        ws.column_dimensions[ws.cell(row=header_row, column=c)
                             .column_letter].width = 22

    path = os.path.join(out_dir, f"{f['source_file_id']}_{layout['layout_id']}"
                                 f"_{f['account']['client_account']}.xlsx")
    _save_deterministic(wb, path)
    return path


# ----------------------------------------------------------------------
# Outputs
# ----------------------------------------------------------------------

JOURNAL_FIELDS = list(LABEL_COUNTS) + CONTROL_FIELDS


def write_journal(rows, path):
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["journal_index"] + JOURNAL_FIELDS
                   + ["is_duplicate_of", "counterparty_ref",
                      "printed_balance_delta"])
        for r in sorted(rows, key=lambda x: x["_journal_index"]):
            out = [r["_journal_index"]]
            for f in JOURNAL_FIELDS:
                v = r[f]
                if isinstance(v, date):
                    v = v.isoformat()
                out.append("" if v is None else v)
            out += [r["_is_duplicate_of"], r["_cp_id"],
                    f"{r['_printed_delta']:.4f}"]
            w.writerow(out)


def write_injection_log(injection, path):
    for i, d in enumerate(sorted(injection, key=lambda x: (x["defect_class"],
                                                           str(x["journal_index"]))),
                          start=1):
        d["defect_id"] = f"D{i:03d}"
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=["defect_id", "defect_class", "level",
                                           "source_file_id", "journal_index",
                                           "detail"])
        w.writeheader()
        for d in sorted(injection, key=lambda x: x["defect_id"]):
            w.writerow(d)


def write_layout_registry(layouts, files, path):
    files_by_layout = defaultdict(list)
    for f in files:
        files_by_layout[f["layout"]].append(f["source_file_id"])
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["layout_id", "issuer", "direction_encoding", "date_format",
                    "decimal_comma", "has_currency_column",
                    "has_balance_column", "has_preamble", "has_summary_rows",
                    "two_row_header", "files"]
                   + [f"label:{f}" for f in LABEL_COUNTS])
        for b, lay in enumerate(layouts):
            w.writerow([lay["layout_id"], f"SYNTH {ISSUERS[lay['issuer']]}",
                        lay["direction_encoding"], lay["date_format"],
                        lay["decimal_comma"], lay["has_currency_column"],
                        lay["has_balance_column"], lay["has_preamble"],
                        lay["has_summary_rows"], lay["two_row_header"],
                        ";".join(files_by_layout[b])]
                       + [lay["labels"][f] for f in LABEL_COUNTS])


def write_request_register(files, path):
    """The covering record received with the statements: which account each
    file belongs to, whose it is, and the period that was requested. It is
    the input of Stage 1 (inventory) and the source of the client fields for
    files whose layout prints no preamble."""
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["source_file_id", "client_code", "client_identifier",
                    "client_account", "declared_period"])
        for f in files:
            a = f["account"]
            w.writerow([f["source_file_id"], a["client_code"],
                        a["client_identifier"], a["client_account"],
                        f"{f['period_start'].isoformat()}/{f['period_end'].isoformat()}"])


def write_label_index(layouts, path):
    seen = defaultdict(set)
    for lay in layouts:
        for f, l in lay["labels"].items():
            seen[f].add(l)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["normalized_field", "source_column_label"])
        for f in LABEL_COUNTS:
            for l in sorted(seen[f]):
                w.writerow([f, l])


def write_coverage(rows, path):
    n = len(rows)
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.writer(fh)
        w.writerow(["normalized_field", "group", "mandatory",
                    "distinct_labels", "present", "coverage_cj"])
        groups = {
            "client_code": "Client", "client_identifier": "Client",
            "client_account": "Client", "statement_period": "Client",
            "operation_date": "Transaction", "direction": "Transaction",
            "amount": "Transaction", "currency": "Transaction",
            "balance_after": "Transaction", "category": "Classification",
            "subcategory": "Classification",
            "counterparty_name": "Counterparty",
            "counterparty_identifier": "Counterparty",
            "counterparty_account_or_channel": "Counterparty",
            "counterparty_type": "Counterparty",
        }
        for f in LABEL_COUNTS:
            present = sum(1 for r in rows if r[f] is not None and r[f] != "")
            w.writerow([f, groups[f], "yes" if f in MANDATORY else "no",
                        LABEL_COUNTS[f], present, f"{present / n:.3f}"])


_FIXED_TS = datetime.datetime(2026, 1, 1, 0, 0, 0)


def _save_deterministic(wb, path):
    """Save a workbook so that the same seed yields byte-identical files.

    openpyxl stamps the current wall-clock time into docProps/core.xml and into
    every zip member header, which makes an otherwise deterministic generator
    produce a different digest on every run. Both are pinned here.
    """
    wb.properties.created = _FIXED_TS
    wb.properties.modified = _FIXED_TS
    wb.properties.creator = "sbsc-generator"
    wb.properties.lastModifiedBy = "sbsc-generator"
    buf = io.BytesIO()
    wb.save(buf)
    stamp = _FIXED_TS.strftime("%Y-%m-%dT%H:%M:%SZ")
    src = zipfile.ZipFile(buf)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in sorted(src.namelist()):
            data = src.read(name)
            if name == "docProps/core.xml":
                # openpyxl rewrites dcterms:modified at save time regardless of
                # what the properties object says, so it is pinned here.
                data = re.sub(rb"(<dcterms:modified[^>]*>)[^<]*(</dcterms:modified>)",
                              rb"\g<1>" + stamp.encode() + rb"\g<2>", data)
            info = zipfile.ZipInfo(name, date_time=(2026, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o600 << 16
            dst.writestr(info, data)
    src.close()


def write_manifest(root, path):
    entries = []
    for dirpath, _, names in os.walk(root):
        for nm in sorted(names):
            if nm == os.path.basename(path):
                continue
            p = os.path.join(dirpath, nm)
            entries.append((sha256(p), os.path.relpath(p, root)))
    with open(path, "w", encoding="utf-8") as fh:
        for h, rel in sorted(entries, key=lambda x: x[1]):
            fh.write(f"{h}  {rel}\n")


# ----------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default=os.path.dirname(os.path.abspath(__file__)))
    ap.add_argument("--seed", type=int, default=20260824)
    args = ap.parse_args()

    rng = random.Random(args.seed)
    np.random.seed(args.seed)

    out = args.out
    st_dir = os.path.join(out, "statements")
    gt_dir = os.path.join(out, "ground_truth")
    rp_dir = os.path.join(out, "reports")
    for d in (st_dir, gt_dir, rp_dir):
        os.makedirs(d, exist_ok=True)

    injection: list[dict] = []

    counterparties = build_counterparties(rng)
    accounts = build_accounts(rng)
    layouts = build_layouts(rng)
    files = build_files(accounts, layouts, rng)

    rows = build_transactions(files, counterparties, rng)
    ensure_all_counterparties(rows, counterparties, rng)
    layout_of_file = {f["source_file_id"]: layouts[f["layout"]] for f in files}
    for r in rows:
        j = layout_of_file[r["source_file_id"]]["balance_jitter"]
        r["_printed_delta"] = round(rng.uniform(-j, j), 4) if j else 0.0
    rows = insert_duplicates(rows, rng, injection)
    by_file = order_and_balance(rows, files, layouts)
    apply_coverage(rows, by_file, files, layouts, rng)
    inject_row_defects(rows, by_file, files, layouts, rng, injection)
    inject_file_defects(files, layouts, rng, injection)

    for f in files:
        render_file(f, layouts[f["layout"]], by_file[f["source_file_id"]],
                    st_dir)

    write_journal(rows, os.path.join(gt_dir, "normalized_journal.csv"))
    write_injection_log(injection, os.path.join(gt_dir, "injection_log.csv"))
    write_layout_registry(layouts, files,
                          os.path.join(gt_dir, "layout_registry.csv"))
    write_label_index(layouts, os.path.join(gt_dir, "field_label_index.csv"))
    write_request_register(files, os.path.join(st_dir, "request_register.csv"))
    write_coverage(rows, os.path.join(rp_dir, "coverage.csv"))

    summary = {
        "corpus_id": "SBSC-1",
        "generator_version": PROCESSING_VERSION,
        "seed": args.seed,
        "statement_files": len(files),
        "distinct_source_layouts": len(layouts),
        "synthetic_issuers": SPEC["issuers"],
        "normalized_transactions": len(rows),
        "account_holders": SPEC["subjects"],
        "counterparties": SPEC["counterparties"],
        "currencies": SPEC["currencies"],
        "period_months": SPEC["period_months"],
        "injected_defects": dict(Counter(d["defect_class"] for d in injection)),
        "injected_defects_total": len(injection),
        "row_level_defects": sum(1 for d in injection if d["level"] == "row"),
        "file_level_defects": sum(1 for d in injection if d["level"] == "file"),
        "rows_with_balance_field": sum(1 for r in rows
                                       if r["balance_after"] is not None),
    }
    with open(os.path.join(rp_dir, "corpus_summary.json"), "w",
              encoding="utf-8") as fh:
        json.dump(summary, fh, indent=2)

    write_manifest(out, os.path.join(gt_dir, "manifest_sha256.txt"))

    print(json.dumps(summary, indent=2))


if __name__ == "__main__":
    main()
