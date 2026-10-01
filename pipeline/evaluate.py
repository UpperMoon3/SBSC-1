#!/usr/bin/env python3
"""
Score the reference pipeline against the SBSC-1 ground truth.

    python evaluate.py --corpus ../corpus --out results/

Rows are aligned to the ground-truth journal by file and by order within the
file (the ground truth numbers rows within the transaction table, the pipeline
numbers spreadsheet rows). Detection is scored per defect class against
injection_log.csv:

  * a row-level defect is detected when the pipeline logs a candidate of the
    same class on that row (for BALANCE_INCONSISTENCY also on the next scorable
    row, since an isolated offset shows up in the successor's residual);
  * a file-level defect is detected when the pipeline logs a layout obstacle
    for that file;
  * every other logged candidate of a class is a false positive of that class.
"""
from __future__ import annotations

import argparse
import csv
import json
import math
import os
from collections import Counter, defaultdict

from sbsc_pipeline import FIELDS, run

MANDATORY = FIELDS[:8]
CLASSES = ["DUPLICATE_ROW", "DIRECTION_INCONSISTENCY", "BALANCE_INCONSISTENCY",
           "INCOMPLETE_COUNTERPARTY", "CURRENCY_MIXING", "UNPARSED_LAYOUT",
           "UNMAPPED_FIELD"]


def wilson(k, n, z=1.959963985):
    if n == 0:
        return (float("nan"), float("nan"))
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    m = z / d * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return (max(0.0, c - m), min(1.0, c + m))


def read_csv(p):
    with open(p, newline="", encoding="utf-8") as fh:
        return list(csv.DictReader(fh))


def align(journal, gt):
    by_file_p = defaultdict(list)
    for i, r in enumerate(journal):
        by_file_p[r["_eval_file"]].append(i)
    by_file_g = defaultdict(list)
    for r in gt:
        by_file_g[r["source_file_id"]].append(int(r["journal_index"]))
    p2g, g2p = {}, {}
    problems = []
    for f in sorted(set(by_file_p) | set(by_file_g)):
        ps = sorted(by_file_p[f], key=lambda i: journal[i]["_eval_row"])
        gs = sorted(by_file_g[f], key=lambda j: int(gt[j]["source_row"]))
        if len(ps) != len(gs):
            problems.append(f"{f}: pipeline {len(ps)} rows, ground truth {len(gs)}")
        for i, j in zip(ps, gs):
            p2g[i] = j
            g2p[j] = i
    return p2g, g2p, problems


def field_equal(f, pv, gv):
    if gv in ("", None):
        return pv in ("", None)
    if pv in ("", None):
        return False
    if f in ("amount", "balance_after"):
        return abs(float(pv) - float(gv)) < 0.006
    return str(pv) == str(gv)


def score(res, corpus):
    gt = read_csv(os.path.join(corpus, "ground_truth", "normalized_journal.csv"))
    inj = read_csv(os.path.join(corpus, "ground_truth", "injection_log.csv"))
    journal, defects = res["journal"], res["defects"]
    p2g, g2p, problems = align(journal, gt)
    n = len(journal)

    # traceability and coverage
    T = sum(1 for r in journal if all(r.get(k) not in (None, "") for k in
                                      ("source_file_id", "source_sheet",
                                       "source_row", "processing_version"))) / n
    coverage = {f: sum(1 for r in journal if r[f] not in (None, "")) / n
                for f in FIELDS}
    # agreement of the mapped values with the ground truth, on rows the
    # corpus did not deliberately corrupt in that field
    agree = {}
    for f in FIELDS:
        ok = tot = 0
        for i, j in p2g.items():
            if gt[j]["defect_flag"] and f in ("direction", "currency",
                                              "counterparty_name",
                                              "counterparty_identifier",
                                              "counterparty_account_or_channel",
                                              "counterparty_type", "balance_after"):
                continue
            tot += 1
            gv = gt[j][f]
            if f == "balance_after" and gv not in ("", None):
                # the file prints the balance with its display rounding
                gv = float(gv) + float(gt[j]["printed_balance_delta"])
            ok += field_equal(f, journal[i][f], gv)
        agree[f] = ok / tot if tot else float("nan")

    # detection
    inj_rows = defaultdict(set)     # class -> set of gt indices
    inj_files = set()
    for d in inj:
        if d["level"] == "file":
            inj_files.add(d["source_file_id"])
        else:
            inj_rows[d["defect_class"]].add(int(d["journal_index"]))

    # successor of an isolated balance offset, in the pipeline's account order
    per_acct = defaultdict(list)
    for i, r in enumerate(journal):
        per_acct[r["client_account"]].append(i)
    succ = {}
    for idx in per_acct.values():
        idx.sort(key=lambda i: (journal[i]["_eval_file"], journal[i]["_eval_row"]))
        for a, b in zip(idx, idx[1:]):
            succ[a] = b

    flagged = defaultdict(set)      # class -> pipeline row indices
    file_flags = set()
    for d in defects:
        if d["class"] == "UNPARSED_LAYOUT" and d["code"] == "LAYOUT_OBSTACLE":
            file_flags.add(d["source_file_id"])
        elif d["class"] and "_i" in d:
            flagged[d["class"]].add(d["_i"])

    table = {}
    for c in CLASSES:
        if c == "UNPARSED_LAYOUT":
            tp = len(inj_files & file_flags)
            fp = len(file_flags - inj_files)
            n_inj = len(inj_files)
        else:
            targets = {g2p[j] for j in inj_rows[c] if j in g2p}
            hit = set()
            used = set()
            for t in targets:
                cand = {t}
                if c == "BALANCE_INCONSISTENCY" and t in succ:
                    cand.add(succ[t])
                m = cand & flagged[c]
                if m:
                    hit.add(t)
                    used |= m
            tp = len(hit)
            # a residual caused by a flipped sign that is labelled as a balance
            # defect (or vice versa) is counted as a false positive of the
            # label it was given, which is what a reviewer would see
            fp = len(flagged[c] - used)
            n_inj = len(inj_rows[c])
        fn = n_inj - tp
        P = tp / (tp + fp) if tp + fp else float("nan")
        R = tp / n_inj if n_inj else float("nan")
        F = 2 * P * R / (P + R) if tp else 0.0
        lo, hi = wilson(tp, n_inj)
        table[c] = {"injected": n_inj, "TP": tp, "FP": fp, "FN": fn,
                    "P": P, "R": R, "F1": F, "R_CI": [lo, hi]}
    TP = sum(v["TP"] for v in table.values())
    FP = sum(v["FP"] for v in table.values())
    NI = sum(v["injected"] for v in table.values())
    P = TP / (TP + FP)
    R = TP / NI
    micro = {"injected": NI, "TP": TP, "FP": FP, "FN": NI - TP, "P": P, "R": R,
             "F1": 2 * P * R / (P + R), "R_CI": list(wilson(TP, NI))}

    # direction defects that reconciliation cannot reach, and whether the run
    # reported them as unreachable (NO_BALANCE_FIELD on that row)
    nobal = {d["_i"] for d in defects if d["code"] == "NO_BALANCE_FIELD"}
    unreach = [int(d["journal_index"]) for d in inj
               if d["defect_class"] == "DIRECTION_INCONSISTENCY"
               and "undetectable" in d["detail"]]
    reported_unreachable = sum(1 for j in unreach if g2p.get(j) in nobal)

    # visibility: a defect is visible when the artefact flags its row or
    # records it (file-level obstacle, unreachable check) in the defect log;
    # without a defect log only in-row flags remain
    detected_rows = set()
    for c in CLASSES:
        if c == "UNPARSED_LAYOUT":
            continue
        for j in inj_rows[c]:
            i = g2p.get(j)
            if i is not None and journal[i]["defect_flag"]:
                detected_rows.add((c, j))
    visible = len(detected_rows)
    if res["log_visible"]:
        visible += len(inj_files & file_flags)
        visible += sum(1 for j in unreach if g2p.get(j) in nobal
                       and ("DIRECTION_INCONSISTENCY", j) not in detected_rows)
    traceable = T > 0
    res_abs = sorted(abs(x["eps"]) for x in res["residuals"]) or [0.0]
    q = lambda p: res_abs[min(len(res_abs) - 1, int(p * len(res_abs)))]
    exceed = [x for x in res["residuals"] if abs(x["eps"]) > 0.01]
    return {
        "alignment_problems": problems,
        "rows": n, "files": len(res["inventory"]),
        "traceability_T": T, "hash_coverage": len(res["inventory"]),
        "coverage": coverage, "value_agreement": agree,
        "reconciliation": {"scorable": len(res["residuals"]),
                           "unreconcilable": res["no_balance_rows"],
                           "exceeding": len(exceed),
                           "accounts_exceeding": len({x["acct"] for x in exceed}),
                           "q50": q(.5), "q90": q(.9), "q99": q(.99)},
        "chronology_violations": res["chronology_violations"],
        "detection": table, "micro": micro,
        "unreachable_direction": {"injected": len(unreach),
                                  "reported_as_unreachable": reported_unreachable},
        "visible": visible, "visible_and_traceable": visible if traceable else 0,
        "integrity_failures": res["integrity_failures"],
        "defect_log_size": len(defects) if res["log_visible"] else 0,
        "defect_codes": Counter(d["code"] for d in defects),
        "elapsed_s": res["elapsed_s"],
    }


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", default="results")
    ap.add_argument("--no-fallback", action="store_true")
    a = ap.parse_args()
    res = run(a.corpus, os.path.join(a.out, "run"), fallback=not a.no_fallback)
    s = score(res, a.corpus)
    os.makedirs(a.out, exist_ok=True)
    with open(os.path.join(a.out, "evaluation.json"), "w") as fh:
        json.dump(s, fh, indent=2, default=str)
    print(json.dumps(s, indent=2, default=str))
