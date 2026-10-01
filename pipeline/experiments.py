#!/usr/bin/env python3
"""
Revision experiments on SBSC-1 v1.1 with the reference pipeline.

    python experiments.py --corpus ../corpus --out results/

  ablation      one stage and pairs of stages removed (R1.d5)
  tolerance     reconciliation tolerance sweep (R1.4)
  duplicates    exact key, fallback key, near-duplicates, overlapping
                deliveries and legitimate same-day repeats (R1.6, R2.7)
  unseen        defect types the generator does not produce (R1.3, R2.3)
  scaling       runtime from 300 to 9600 rows, measured on this machine (R1.10)
  indicators    null flag rates across theta, l and Delta (R2.5, R1.12)
  routes        candidate links, admissibility and sensitivity of the ranking
                to omega and beta (R1.7)
  naive         naive concatenation baseline, computed (R2.6)
"""
from __future__ import annotations

import argparse
import copy
import csv
import itertools
import json
import math
import os
import platform
import random
import re
import shutil
import statistics
import subprocess
import sys
import tempfile
import time
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

import openpyxl

from sbsc_pipeline import (FIELDS, STAGES, dup_key, nu, qc_duplicates,
                           read_registry, run)
from evaluate import CLASSES, read_csv, score


def dump(out, name, obj):
    with open(os.path.join(out, name), "w") as fh:
        json.dump(obj, fh, indent=2, default=str)


# ------------------------------------------------------------------ ablation
def ablation(corpus):
    configs = [("full", ())]
    configs += [(f"without {s}", (s,)) for s in STAGES]
    configs += [("without lineage and defect_log", ("lineage", "defect_log")),
                ("without reconciliation and duplicates", ("reconciliation", "duplicates")),
                ("without hash and lineage", ("hash", "lineage"))]
    rows = []
    for name, dis in configs:
        res = run(corpus, None, disable=dis)
        s = score(res, corpus)
        tam = run(corpus, None, disable=dis, tamper="F001_L02_ACC-001.xlsx")
        rows.append({"config": name, "detected": s["micro"]["TP"],
                     "injected": s["micro"]["injected"], "FP": s["micro"]["FP"],
                     "rate": s["micro"]["R"], "T": s["traceability_T"],
                     "visible": s["visible"],
                     "visible_and_traceable": s["visible_and_traceable"],
                     "tamper_detected": bool(tam["integrity_failures"])})
    return rows


# ----------------------------------------------------------------- tolerance
def tolerance(corpus):
    out = []
    for tau in [0.0001, 0.001, 0.005, 0.01, 0.02, 0.05, 0.1, 0.2, 0.3, 0.5,
                1.0, 2.0]:
        res = run(corpus, None, tau=tau)
        s = score(res, corpus)
        ex = [x for x in res["residuals"] if abs(x["eps"]) > tau]
        out.append({"tau": tau, "rows_exceeding": len(ex),
                    "accounts": len({x["acct"] for x in ex}),
                    "balance_TP": s["detection"]["BALANCE_INCONSISTENCY"]["TP"],
                    "balance_FP": s["detection"]["BALANCE_INCONSISTENCY"]["FP"],
                    "direction_TP": s["detection"]["DIRECTION_INCONSISTENCY"]["TP"],
                    "direction_FP": s["detection"]["DIRECTION_INCONSISTENCY"]["FP"]})
    res = run(corpus, None)
    per = defaultdict(list)
    for x in res["residuals"]:
        per[x["ccy"]].append(abs(x["eps"]))
    by_ccy = {}
    for c, v in per.items():
        small = [e for e in v if e <= 0.01]
        large = [e for e in v if e > 0.01]
        by_ccy[c] = {"scorable": len(v), "max_conforming": max(small) if small else None,
                     "min_exceeding": min(large) if large else None,
                     "exceeding": len(large)}
    return {"sweep": out, "per_currency": by_ccy}


# ---------------------------------------------------------------- duplicates
def duplicates(corpus, seed=7):
    """Journal-level stress test of duplicate screening. The pipeline journal is
    taken as produced, the 14 injected duplicates are located, and variants are
    built from it; screening is re-run on each variant."""
    rng = random.Random(seed)
    base = run(corpus, None)
    gt = read_csv(os.path.join(corpus, "ground_truth", "normalized_journal.csv"))
    from evaluate import align
    p2g, g2p, _ = align(base["journal"], gt)
    pairs = [(g2p[int(r["is_duplicate_of"])], g2p[int(r["journal_index"])])
             for r in gt if r["is_duplicate_of"]]

    def clean(j):
        j = copy.deepcopy(j)
        for r in j:
            r["defect_flag"] = ""
        return j

    def screen(journal, fallback=True, fuzzy=False):
        d = []
        qc_duplicates(journal, d, fallback)
        found = {x["_i"]: x["code"] for x in d}
        if fuzzy:
            for i in fuzzy_candidates(journal):
                found.setdefault(i, "DUP_CANDIDATE_FUZZY")
        return found

    out = {}
    # 1. exact key vs fallback on the corpus as delivered
    for fb in (False, True):
        f = screen(clean(base["journal"]), fallback=fb)
        tp = sum(1 for a, b in pairs if b in f)
        out[f"corpus_fallback_{fb}"] = {"TP": tp, "FP": len(f) - tp}

    # 2. near-duplicates: each injected copy perturbed one way
    def typo(s):
        if not s or len(s) < 4:
            return s
        i = rng.randrange(1, len(s) - 2)
        return s[:i] + s[i + 1] + s[i] + s[i + 2:]
    modes = {
        "name_misspelled": lambda r: r.update(counterparty_name=typo(r["counterparty_name"])),
        "name_legal_form_dropped": lambda r: r.update(
            counterparty_name=(r["counterparty_name"] or "").replace(" LLP", "")),
        "amount_rounded_to_unit": lambda r: r.update(amount=float(round(r["amount"]) or 1)),
        "date_shifted_one_day": lambda r: r.update(operation_date=(
            date.fromisoformat(r["operation_date"]) + timedelta(days=1)).isoformat()),
    }
    near = {}
    for m, fn in modes.items():
        res = {}
        for variant in ("exact+fallback", "exact+fallback+fuzzy"):
            j = clean(base["journal"])
            applicable = 0
            for a, b in pairs:
                before = dict(j[b])
                fn(j[b])
                applicable += j[b] != before
            f = screen(j, fuzzy=variant.endswith("fuzzy"))
            tp = sum(1 for a, b in pairs if b in f)
            res[variant] = {"TP": tp, "FP": len(f) - tp, "n": len(pairs),
                            "perturbed": applicable}
        near[m] = res
    out["near_duplicates"] = near

    # 3. overlapping delivery: the last 10 days of each account's first file are
    # delivered again as part of a second export, and 4. legitimate same-day
    # repeats: 10 genuine second payments identical to an existing one
    j = clean(base["journal"])
    by_acct = defaultdict(list)
    for i, r in enumerate(j):
        by_acct[r["client_account"]].append(i)
    overlap_rows = []
    for acct, idx in by_acct.items():
        first_file = min(j[i]["source_file_id"] for i in idx)
        rows = [i for i in idx if j[i]["source_file_id"] == first_file]
        last = max(j[i]["operation_date"] for i in rows)
        cutoff = (date.fromisoformat(last) - timedelta(days=10)).isoformat()
        for i in rows:
            if j[i]["operation_date"] > cutoff:
                r = dict(j[i])
                r["source_file_id"] = f"X{first_file}"
                r["source_row"] = 1000 + i
                overlap_rows.append(len(j))
                j.append(r)
    repeat_rows = []
    for i in rng.sample(range(len(base["journal"])), 10):
        r = dict(j[i])
        r["source_row"] = r["source_row"] + 0.5        # printed on the next line
        repeat_rows.append(len(j))
        j.append(r)
    f = screen(j)
    out["overlap_and_repeat"] = {
        "overlap_injected": len(overlap_rows),
        "overlap_found": sum(1 for i in overlap_rows if i in f),
        "overlap_labelled_overlap": sum(1 for i in overlap_rows
                                        if f.get(i) == "DUP_CANDIDATE_OVERLAP"),
        "repeat_injected": len(repeat_rows),
        "repeat_found": sum(1 for i in repeat_rows if i in f),
        "repeat_labelled_repeat": sum(1 for i in repeat_rows
                                      if f.get(i) == "DUP_CANDIDATE_REPEAT"),
        "legitimate_repeats_removed": 0,
    }
    return out


def fuzzy_candidates(journal, name_sim=0.85, amount_tol=1.0, days=1):
    """Third pass for near-duplicates: same account and direction, dates within
    one day, amounts within one major unit, and normalized names similar (or one
    missing). Candidates only; nothing is removed."""
    import difflib
    per = defaultdict(list)
    for i, r in enumerate(journal):
        if r["amount"] is not None:
            per[(r["client_account"], r["direction"])].append(i)
    found = set()
    for idx in per.values():
        idx.sort(key=lambda i: journal[i]["operation_date"])
        for x in range(len(idx)):
            a = journal[idx[x]]
            da = date.fromisoformat(a["operation_date"])
            for y in range(x + 1, len(idx)):
                b = journal[idx[y]]
                if (date.fromisoformat(b["operation_date"]) - da).days > days:
                    break
                if abs(a["amount"] - b["amount"]) > amount_tol:
                    continue
                na, nb = nu(a["counterparty_name"]), nu(b["counterparty_name"])
                if na and nb and difflib.SequenceMatcher(None, na, nb).ratio() < name_sim:
                    continue
                found.add(max(idx[x], idx[y]))
    return found


# ------------------------------------------------------------------- unseen
def unseen(corpus, out_dir):
    """Defect types the generator never produces, applied to copies of the
    statement files. For each, the question is whether the run surfaces the
    problem in the defect log or absorbs it silently into the journal."""
    reg = read_registry(corpus)
    base = run(corpus, None)
    base_codes = Counter(d["code"] for d in base["defects"])

    def mutate(name, fn):
        tmp = tempfile.mkdtemp(prefix="sbsc_mut_")
        shutil.copytree(corpus, os.path.join(tmp, "c"))
        c = os.path.join(tmp, "c")
        info = fn(c)
        res = run(c, None)
        codes = Counter(d["code"] for d in res["defects"])
        new = {k: v - base_codes.get(k, 0) for k, v in codes.items()
               if v != base_codes.get(k, 0)}
        rows_delta = len(res["journal"]) - len(base["journal"])
        shutil.rmtree(tmp)
        return {"mutation": name, "detail": info, "new_log_entries": new,
                "journal_rows_delta": rows_delta,
                "surfaced": any(v > 0 for v in new.values())}

    def edit(c, fname, fn):
        p = os.path.join(c, "statements", fname)
        wb = openpyxl.load_workbook(p)
        ws = wb.worksheets[0]
        info = fn(ws)
        wb.save(p)
        return info

    def header_row(ws):
        for r in range(1, 15):
            vals = [ws.cell(row=r, column=k).value for k in range(1, ws.max_column + 1)]
            if sum(v is not None for v in vals) >= 4 and r > 0 and \
               ws.cell(row=r + 1, column=1).value is not None:
                if any(isinstance(v, str) and re.match(r"\d", v or "") for v in
                       [ws.cell(row=r + 1, column=1).value]):
                    return r
        return 1

    results = []
    # 1. a column renamed to a label absent from the registry
    def m1(c):
        return edit(c, "F012_L09_ACC-002.xlsx", lambda ws: (
            ws.cell(row=header_row(ws), column=3, value="Amount (net)"), "amount column renamed")[1])
    results.append(mutate("renamed_amount_column", m1))

    # 2. a file exported with day and month swapped for days <= 12
    def m2(c):
        def f(ws):
            h = header_row(ws)
            n = 0
            for r in range(h + 1, ws.max_row + 1):
                v = ws.cell(row=r, column=1).value
                if isinstance(v, str) and re.match(r"\d{2}\.\d{2}\.\d{4}$", v):
                    d, m, y = v.split(".")
                    if int(d) <= 12:
                        ws.cell(row=r, column=1, value=f"{m}.{d}.{y}")
                        n += 1
            return f"{n} dates with day and month swapped"
        return edit(c, "F001_L02_ACC-001.xlsx", f)
    results.append(mutate("day_month_swapped", m2))

    # 3. the last 8 transaction rows of a file cut off (truncated export)
    def m3(c):
        def f(ws):
            ws.delete_rows(ws.max_row - 7, 8)
            return "last 8 rows removed"
        return edit(c, "F002_L01_ACC-001.xlsx", f)
    results.append(mutate("truncated_file", m3))

    # 4. one amount cell with thousands separator and decimal mark swapped
    def m4(c):
        def f(ws):
            h = header_row(ws)
            ws.cell(row=h + 5, column=2, value="12.345,67")
            return "one amount written in the other decimal convention"
        return edit(c, "F006_L04_ACC-004.xlsx", f)
    results.append(mutate("decimal_convention_flip", m4))

    # 5. a transfer between two of the subject's own accounts appearing in both
    #    statements is a legitimate pair, not a duplicate; here instead the same
    #    file is delivered twice under two names
    def m5(c):
        src = os.path.join(c, "statements", "F013_L09_ACC-002.xlsx")
        shutil.copy(src, os.path.join(c, "statements", "F031_L09_ACC-002.xlsx"))
        with open(os.path.join(c, "statements", "request_register.csv"), "a",
                  newline="") as fh:
            rr = read_csv(os.path.join(c, "statements", "request_register.csv"))
            r13 = [r for r in rr if r["source_file_id"] == "F013"][0]
            csv.writer(fh).writerow(["F031", r13["client_code"], r13["client_identifier"],
                                     r13["client_account"], r13["declared_period"]])
        return "F013 delivered a second time as F031"
    results.append(mutate("file_delivered_twice", m5))

    # 6. a direction token outside the documented vocabulary
    def m6(c):
        def f(ws):
            h = header_row(ws)
            ws.cell(row=h + 3, column=2, value="REV")
            return "direction token REV (reversal)"
        return edit(c, "F012_L09_ACC-002.xlsx", f)
    results.append(mutate("unknown_direction_token", m6))

    # 7. a value shifted one column to the right on one row
    def m7(c):
        def f(ws):
            h = header_row(ws)
            r = h + 4
            vals = [ws.cell(row=r, column=k).value for k in range(1, ws.max_column + 1)]
            for k, v in enumerate([vals[0], None] + vals[1:-1], start=1):
                ws.cell(row=r, column=k, value=v)
            return "one row shifted right from the amount column"
        return edit(c, "F021_L12_ACC-003.xlsx", f)
    results.append(mutate("row_shifted_one_column", m7))

    # 8. a currency code in lower case in a currency column
    def m8(c):
        def f(ws):
            h = header_row(ws)
            hdr = [ws.cell(row=h, column=k).value for k in range(1, ws.max_column + 1)]
            k = next(i for i, v in enumerate(hdr, 1) if v in ("CCY", "Currency", "Curr.",
                                                              "Currency code", "Val."))
            ws.cell(row=h + 2, column=k, value="kzt")
            return "currency written as 'kzt'"
        return edit(c, "F016_L08_ACC-005.xlsx", f)
    results.append(mutate("currency_code_lowercase", m8))
    return results


# ------------------------------------------------------------------ scaling
def build_scaled(corpus, k, tmp):
    """Merge k corpora generated with different seeds into one corpus with
    prefixed file, layout and account identifiers."""
    gen = os.path.join(corpus, "generate_corpus.py")
    dst = os.path.join(tmp, f"scaled_{k}")
    if os.path.exists(dst):
        return dst
    os.makedirs(os.path.join(dst, "statements"))
    os.makedirs(os.path.join(dst, "ground_truth"))
    reg_rows, rr_rows, header, rr_header = [], [], None, None
    for s in range(k):
        sd = os.path.join(tmp, f"seed_{s}")
        if not os.path.exists(sd):
            subprocess.run([sys.executable, gen, "--out", sd, "--seed", str(20260824 + s)],
                           check=True, capture_output=True)
        p = f"S{s}"
        with open(os.path.join(sd, "ground_truth", "layout_registry.csv"), newline="") as fh:
            rd = csv.reader(fh)
            header = next(rd)
            for row in rd:
                row[0] = p + row[0]
                reg_rows.append(row)
        with open(os.path.join(sd, "statements", "request_register.csv"), newline="") as fh:
            rd = csv.reader(fh)
            rr_header = next(rd)
            for row in rd:
                row[0] = p + row[0]
                row[3] = p + row[3]
                rr_rows.append(row)
        for name in os.listdir(os.path.join(sd, "statements")):
            if name.endswith(".xlsx"):
                fid, lay, acct = name[:-5].split("_", 2)
                shutil.copy(os.path.join(sd, "statements", name),
                            os.path.join(dst, "statements", f"{p}{fid}_{p}{lay}_{p}{acct}.xlsx"))
    with open(os.path.join(dst, "ground_truth", "layout_registry.csv"), "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(header)
        w.writerows(reg_rows)
    with open(os.path.join(dst, "statements", "request_register.csv"), "w", newline="") as fh:
        w = csv.writer(fh)
        w.writerow(rr_header)
        w.writerows(rr_rows)
    return dst


def subset(corpus_dir, n_files, tmp):
    dst = os.path.join(tmp, f"subset_{n_files}")
    shutil.copytree(corpus_dir, dst)
    files = sorted(f for f in os.listdir(os.path.join(dst, "statements")) if f.endswith(".xlsx"))
    for f in files[n_files:]:
        os.remove(os.path.join(dst, "statements", f))
    return dst


def scaling(corpus, reps=5):
    import resource
    tmp = tempfile.mkdtemp(prefix="sbsc_scale_")
    one = build_scaled(corpus, 1, tmp)
    sizes = [("subset", 8), ("subset", 15), ("k", 1), ("k", 2), ("k", 4), ("k", 8)]
    out = []
    for kind, v in sizes:
        d = subset(one, v, tmp) if kind == "subset" else build_scaled(corpus, v, tmp)
        times, rows, files = [], None, None
        for _ in range(reps):
            res = run(d, None)
            times.append(res["elapsed_s"])
            rows, files = len(res["journal"]), len(res["inventory"])
        times.sort()
        q1, q3 = times[len(times) // 4], times[(3 * len(times)) // 4]
        med = statistics.median(times)
        out.append({"rows": rows, "files": files, "runs": reps, "median_s": med,
                    "iqr_s": q3 - q1, "rows_per_s": rows / med})
    peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
    shutil.rmtree(tmp)
    # power-law fit on log-log scale
    xs = [math.log(o["rows"]) for o in out]
    ys = [math.log(o["median_s"]) for o in out]
    mx, my = statistics.mean(xs), statistics.mean(ys)
    b = sum((x - mx) * (y - my) for x, y in zip(xs, ys)) / sum((x - mx) ** 2 for x in xs)
    a = math.exp(my - b * mx)
    ss_res = sum((y - (math.log(a) + b * x)) ** 2 for x, y in zip(xs, ys))
    ss_tot = sum((y - my) ** 2 for y in ys)
    try:
        cpu = subprocess.run(["sysctl", "-n", "machdep.cpu.brand_string"],
                             capture_output=True, text=True).stdout.strip()
        mem = int(subprocess.run(["sysctl", "-n", "hw.memsize"], capture_output=True,
                                 text=True).stdout.strip()) / 2 ** 30
    except Exception:
        cpu, mem = platform.processor(), None
    return {"sizes": out, "fit": {"a": a, "b": b, "R2_log": 1 - ss_res / ss_tot},
            "peak_rss_mb_process": peak / 2 ** 20 if sys.platform == "darwin" else peak / 1024,
            "machine": {"cpu": cpu, "memory_gb": mem, "os": platform.platform(),
                        "python": platform.python_version(),
                        "openpyxl": openpyxl.__version__}}


# --------------------------------------------------------------- indicators
EXCLUDED_TYPES = {"bank", "state", "utility", "marketplace", "payroll_aggregator"}
ASSET = {"Asset"}


def indicator_values(journal, ell=3, delta_h=72):
    per = defaultdict(list)
    for r in journal:
        if r["amount"] is None or r["direction"] is None:
            continue
        per[(r["client_code"], r["currency"])].append(r)
    # each subject is evaluated in the currency partition carrying most rows
    subj = {}
    for (x, c), rows in per.items():
        if x not in subj or len(rows) > len(subj[x]):
            subj[x] = rows
    vals = {}
    cps = {}
    for x, rows in subj.items():
        signed = [(date.fromisoformat(r["operation_date"]),
                   r["amount"] if r["direction"] == "credit" else -r["amount"], r)
                  for r in rows]
        turn = Counter()
        for _, a, r in signed:
            key = r["counterparty_identifier"] or nu(r["counterparty_name"])
            if key:
                turn[key] += abs(a)
        tot = sum(turn.values())
        hhi = sum((v / tot) ** 2 for v in turn.values()) if tot else 0
        rho = sum(1 for _, a, _ in signed if abs(a) / 10 ** ell == int(abs(a) / 10 ** ell)
                  and abs(a) >= 10 ** ell) / len(signed)
        ins = [(d, a) for d, a, _ in signed if a > 0]
        outs = [(d, -a) for d, a, _ in signed if a < 0]
        num = den = 0.0
        days = delta_h / 24
        for d, a in ins:
            o = sum(v for dd, v in outs if 0 < (dd - d).days <= days)
            num += min(o, a)
            den += a
        pi = num / den if den else 0
        out_tot = sum(-a for _, a, _ in signed if a < 0)
        asset = sum(-a for _, a, r in signed if a < 0 and r["category"] in ASSET)
        alpha = asset / out_tot if out_tot else 0
        cps[x] = {(r["counterparty_identifier"] or nu(r["counterparty_name"]))
                  for _, _, r in signed
                  if (r["counterparty_type"] not in EXCLUDED_TYPES)
                  and (r["counterparty_identifier"] or r["counterparty_name"])}
        vals[x] = {"HHI": hhi, "rounded": rho, "transit": pi, "asset": alpha}
    for x in vals:
        best = 0
        for y in vals:
            if y == x:
                continue
            u = cps[x] | cps[y]
            best = max(best, len(cps[x] & cps[y]) / len(u) if u else 0)
        vals[x]["shared"] = best
    return vals


def indicators(corpus):
    res = run(corpus, None)
    j = res["journal"]
    grid = [round(0.01 * i, 2) for i in range(0, 101)]
    base = indicator_values(j)
    sweep = {k: [{"theta": t, "flagged": sum(1 for v in base.values() if v[k] >= t)}
                 for t in grid] for k in ("HHI", "rounded", "transit", "shared", "asset")}
    op = {"HHI": 0.50, "rounded": 0.30, "transit": 0.60, "shared": 0.20, "asset": 0.25}
    at_op = {k: sum(1 for v in base.values() if v[k] >= t) for k, t in op.items()}
    ell_sens = {ell: sorted(round(v["rounded"], 3) for v in indicator_values(j, ell=ell).values())
                for ell in (2, 3, 4)}
    delta_sens = {d: sorted(round(v["transit"], 3) for v in indicator_values(j, delta_h=d).values())
                  for d in (24, 72, 168)}
    return {"values": base, "sweep": sweep, "operating_points": op,
            "flagged_at_operating_points": at_op,
            "rounded_by_ell": ell_sens, "transit_by_delta": delta_sens}


# ------------------------------------------------------------------- routes
def routes(corpus, window_days=14):
    """Candidate links u -> v within an account: incoming u, later outgoing v
    within the window, admissible only if the balance before v could carry it
    (where a balance is printed). Scores follow Eq. (score)."""
    res = run(corpus, None)
    j = [r for r in res["journal"] if r["amount"] is not None and r["direction"]]
    stop = {"transfer", "payment", "incoming", "outgoing", "internal", "for"}

    def toks(r):
        t = set(nu(" ".join(str(r[k] or "") for k in ("category", "subcategory"))).split())
        return t - stop

    per = defaultdict(list)
    for r in j:
        per[r["client_account"]].append(r)
    links = []
    inadmissible = 0
    for acct, rows in per.items():
        rows.sort(key=lambda r: (r["operation_date"], r["source_file_id"], r["source_row"]))
        for x, u in enumerate(rows):
            if u["direction"] != "credit":
                continue
            du = date.fromisoformat(u["operation_date"])
            for v in rows[x + 1:]:
                dv = date.fromisoformat(v["operation_date"])
                if (dv - du).days > window_days:
                    break
                if v["direction"] != "debit":
                    continue
                gA = max(0.0, 1 - abs(v["amount"] - u["amount"]) / u["amount"])
                dt = max((dv - du).days, 0)
                W1, W2 = toks(u), toks(v)
                gP = len(W1 & W2) / len(W1 | W2) if (W1 | W2) else 0.0
                links.append({"acct": acct, "u": (u["source_file_id"], u["source_row"]),
                              "v": (v["source_file_id"], v["source_row"]),
                              "gA": gA, "dt": dt, "gP": gP,
                              "next_statement": any(
                                  r["client_account"] == v.get("counterparty_account_or_channel")
                                  and r["operation_date"] >= v["operation_date"] for r in j)})

    def scores(w1, w2, w3, beta):
        return [w1 * l["gA"] + w2 * math.exp(-l["dt"] / beta) + w3 * l["gP"] for l in links]

    base = scores(0.4, 0.4, 0.2, 7)

    def kendall(a, b, sample=4000, seed=1):
        rng = random.Random(seed)
        n = len(a)
        c = d = 0
        for _ in range(sample):
            i, k = rng.randrange(n), rng.randrange(n)
            if i == k:
                continue
            s = (a[i] - a[k]) * (b[i] - b[k])
            c += s > 0
            d += s < 0
        return (c - d) / (c + d) if c + d else float("nan")

    def top_overlap(a, b, q=0.1):
        n = max(1, int(len(a) * q))
        ta = set(sorted(range(len(a)), key=lambda i: -a[i])[:n])
        tb = set(sorted(range(len(b)), key=lambda i: -b[i])[:n])
        return len(ta & tb) / n

    sens = []
    for (w1, w2, w3) in [(0.4, 0.4, 0.2), (1 / 3, 1 / 3, 1 / 3), (0.5, 0.3, 0.2),
                         (0.3, 0.5, 0.2), (0.5, 0.5, 0.0), (0.45, 0.45, 0.1)]:
        for beta in (3, 7, 14):
            s = scores(w1, w2, w3, beta)
            sens.append({"w": [round(w1, 3), round(w2, 3), round(w3, 3)], "beta": beta,
                         "kendall_tau": kendall(base, s), "top10_overlap": top_overlap(base, s)})
    ws = sorted(base)
    q = lambda p: ws[min(len(ws) - 1, int(p * len(ws)))]
    return {"candidate_links": len(links),
            "without_next_statement": sum(1 for l in links if not l["next_statement"]),
            "w_quantiles": {"q10": q(.1), "q50": q(.5), "q90": q(.9)},
            "share_w_ge_075": sum(1 for x in base if x >= 0.75) / len(base),
            "sensitivity": sens}


# -------------------------------------------------------------------- naive
def naive(corpus):
    """Stack the source tables on the first row of each sheet, positionally,
    with no field mapping, no sign canonicalization and no currency
    partitioning, and count what survives into totals."""
    gt = read_csv(os.path.join(corpus, "ground_truth", "normalized_journal.csv"))
    inj = read_csv(os.path.join(corpus, "ground_truth", "injection_log.csv"))
    reg = read_registry(corpus)
    files = sorted(f for f in os.listdir(os.path.join(corpus, "statements")) if f.endswith(".xlsx"))
    stacked_rows = 0
    header_wrong = 0
    unsigned_debits = 0
    no_ccy_rows = 0
    for f in files:
        fid, lay, _ = f[:-5].split("_", 2)
        L = reg[lay]
        ws = openpyxl.load_workbook(os.path.join(corpus, "statements", f)).worksheets[0]
        rows = list(ws.iter_rows(values_only=True))
        first = rows[0]
        labs = set(L["labels"].values())
        if sum(1 for c in first if c is not None and str(c).replace(" (Dr)", "").replace(" (Cr)", "") in labs) < 4:
            header_wrong += 1
        stacked_rows += len(rows) - 1
        gtf = [r for r in gt if r["source_file_id"] == fid]
        if L["direction_encoding"] in ("flag", "split"):
            unsigned_debits += sum(1 for r in gtf if r["direction"] == "debit")
        if not L["has_currency_column"]:
            no_ccy_rows += len(gtf)
    cls = Counter(d["defect_class"] for d in inj)
    return {"files_misread_on_first_row_header": header_wrong,
            "rows_stacked_including_non_transaction_rows": stacked_rows,
            "rows_retaining_lineage": 0,
            "duplicates_surviving": cls["DUPLICATE_ROW"],
            "sign_errors_surviving_injected": cls["DIRECTION_INCONSISTENCY"],
            "debits_entering_totals_as_positive": unsigned_debits,
            "rows_without_stated_currency": no_ccy_rows,
            "currency_mixing_rows_aggregated": cls["CURRENCY_MIXING"],
            "defects_visible": 0}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", default="results")
    ap.add_argument("--only", nargs="*")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    todo = {"ablation": lambda: ablation(a.corpus),
            "tolerance": lambda: tolerance(a.corpus),
            "duplicates": lambda: duplicates(a.corpus),
            "unseen": lambda: unseen(a.corpus, a.out),
            "indicators": lambda: indicators(a.corpus),
            "routes": lambda: routes(a.corpus),
            "naive": lambda: naive(a.corpus),
            "scaling": lambda: scaling(a.corpus)}
    for k, fn in todo.items():
        if a.only and k not in a.only:
            continue
        t = time.time()
        r = fn()
        dump(a.out, f"{k}.json", r)
        print(f"== {k} ({time.time() - t:.1f}s)")
        print(json.dumps(r, indent=1, default=str)[:3000])
