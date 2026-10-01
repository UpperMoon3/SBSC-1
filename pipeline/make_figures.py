#!/usr/bin/env python3
"""Figures 2-7 of the revised manuscript, drawn from the result files written
by evaluate.py and experiments.py. No value is typed by hand.

    python make_figures.py --results results --corpus ../corpus --out ../figures
"""
import argparse
import csv
import json
import math
import os
from collections import Counter, defaultdict

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np

C1, C2, C3, C4 = "#5B7B9A", "#A8BBCB", "#7C9A7E", "#B5725B"
plt.rcParams.update({"font.family": "serif", "font.size": 8.4,
                     "axes.spines.top": False, "axes.spines.right": False,
                     "axes.titlesize": 8.8, "axes.labelsize": 8.4,
                     "legend.frameon": False, "savefig.bbox": "tight"})

FIELDS = ["client_code", "client_identifier", "client_account", "statement_period",
          "operation_date", "direction", "amount", "currency", "balance_after",
          "category", "subcategory", "counterparty_name", "counterparty_identifier",
          "counterparty_account_or_channel", "counterparty_type"]
MAND = set(FIELDS[:8])


def load(results, name):
    with open(os.path.join(results, name)) as fh:
        return json.load(fh)


def wilson(k, n, z=1.959963985):
    p = k / n
    d = 1 + z * z / n
    c = (p + z * z / (2 * n)) / d
    m = z / d * math.sqrt(p * (1 - p) / n + z * z / (4 * n * n))
    return c - m, c + m


def fig02(results, corpus, out):
    ev = load(results, "evaluation.json")
    with open(os.path.join(corpus, "ground_truth", "field_label_index.csv")) as fh:
        labels = Counter(r["normalized_field"] for r in csv.DictReader(fh))
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 3.6), sharey=True)
    y = np.arange(len(FIELDS))[::-1]
    names = [f.replace("_", r"\_") for f in FIELDS]
    ax[0].barh(y, [labels[f] for f in FIELDS],
               color=[C1 if f in MAND else C2 for f in FIELDS])
    ax[0].set_yticks(y)
    ax[0].set_yticklabels(FIELDS, fontsize=7.4)
    for t, f in zip(ax[0].get_yticklabels(), FIELDS):
        t.set_fontweight("bold" if f in MAND else "normal")
    ax[0].set_xlabel("distinct source column labels")
    ax[0].set_title("(a) label heterogeneity", loc="left")
    cov = [ev["coverage"][f] for f in FIELDS]
    ax[1].barh(y, cov, color=[C1 if f in MAND else C2 for f in FIELDS])
    for yy, c in zip(y, cov):
        ax[1].text(min(c, 1.0) + 0.01, yy, f"{c:.3f}", va="center", fontsize=6.8)
    ax[1].set_xlim(0, 1.12)
    ax[1].set_xlabel(r"coverage $c_j$")
    ax[1].set_title("(b) coverage in the produced journal", loc="left")
    fig.tight_layout()
    fig.savefig(os.path.join(out, "fig02_field_heterogeneity.pdf"))


def fig03(results, out, run_dir):
    # residuals are recomputed from the produced journal
    rows = []
    with open(os.path.join(run_dir, "journal.csv")) as fh:
        rows = list(csv.DictReader(fh))
    per = defaultdict(list)
    for r in rows:
        per[r["client_account"]].append(r)
    eps, worst = [], {}
    for a, rs in per.items():
        rs.sort(key=lambda r: (r["source_file_id"], int(r["source_row"])))
        prev = None
        worst[a] = 0.0
        for r in rs:
            if not r["balance_after"] or not r["amount"]:
                prev = None
                continue
            b = float(r["balance_after"])
            if prev is not None:
                s = float(r["amount"]) * (1 if r["direction"] == "credit" else -1)
                e = abs(round(b - (prev + s), 4))
                eps.append(e)
                worst[a] = max(worst[a], e)
            prev = b
    zero = sum(1 for e in eps if e == 0)
    nz = [e for e in eps if e > 0]
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.7))
    bins = np.logspace(-2.5, 6.5, 37)
    ax[0].hist(nz, bins=bins, color=C4)
    ax[0].set_xscale("log")
    ax[0].axvline(0.01, color="k", lw=0.8, ls="--")
    ax[0].text(0.011, ax[0].get_ylim()[1] * 0.9, r"$\tau_\gamma=0.01$", fontsize=7)
    ax[0].set_xlabel(r"$|\varepsilon_i|$ (non-zero residuals only)")
    ax[0].set_ylabel("rows")
    ax[0].set_title(f"(a) {len(eps)} scorable rows: {zero} with "
                    r"$\varepsilon_i=0$" f", {len(nz)} non-zero", loc="left")
    accs = sorted(worst)
    vals = [worst[a] for a in accs]
    ax[1].bar(range(len(accs)), [max(v, 1e-3) for v in vals],
              color=[C4 if v > 0.01 else C2 for v in vals])
    ax[1].set_yscale("log")
    ax[1].axhline(0.01, color="k", lw=0.8, ls="--")
    ax[1].set_xticks(range(len(accs)))
    ax[1].set_xticklabels([a.replace("ACC-", "") for a in accs], fontsize=7)
    ax[1].set_xlabel("account")
    ax[1].set_ylabel(r"max $|\varepsilon_i|$ (zero drawn at $10^{-3}$)")
    ax[1].set_title("(b) worst residual per account", loc="left")
    fig.tight_layout()
    fig.savefig(os.path.join(out, "fig03_reconciliation_residuals.pdf"))


def fig04(results, out):
    ev = load(results, "evaluation.json")
    order = ["DUPLICATE_ROW", "DIRECTION_INCONSISTENCY", "BALANCE_INCONSISTENCY",
             "INCOMPLETE_COUNTERPARTY", "CURRENCY_MIXING", "UNPARSED_LAYOUT",
             "UNMAPPED_FIELD"]
    names = ["Duplicate\nrows", "Direction\ninconsistency", "Balance\ninconsistency",
             "Incomplete\ncounterparty", "Currency\nmixing", "Layout\nobstacle",
             "Unmapped\nfield"]
    d = ev["detection"]
    P = [d[c]["P"] for c in order]
    R = [d[c]["R"] for c in order]
    F = [d[c]["F1"] for c in order]
    lo = [R[i] - wilson(d[c]["TP"], d[c]["injected"])[0] for i, c in enumerate(order)]
    hi = [wilson(d[c]["TP"], d[c]["injected"])[1] - R[i] for i, c in enumerate(order)]
    x = np.arange(len(order))
    w = 0.26
    fig, ax = plt.subplots(figsize=(7.2, 2.8))
    ax.bar(x - w, P, w, color=C1, label="precision")
    ax.bar(x, R, w, color=C3, label="recall", yerr=[lo, hi], capsize=2,
           error_kw={"lw": 0.7})
    ax.bar(x + w, F, w, color=C2, label=r"$F_1$")
    ax.set_xticks(x)
    ax.set_xticklabels(names, fontsize=7.2)
    ax.set_ylim(0, 1.08)
    ax.set_ylabel("score")
    ax.legend(ncol=3, loc="lower center", bbox_to_anchor=(0.5, 1.0), fontsize=7.4)
    fig.tight_layout()
    fig.savefig(os.path.join(out, "fig04_defect_detection.pdf"))


def fig05(results, out):
    ind = load(results, "indicators.json")
    lab = {"HHI": "concentration (HHI)", "rounded": r"rounded amounts $\rho_3$",
           "transit": r"onward movement $\pi_{72\,h}$", "shared": r"shared counterparty $S$",
           "asset": r"asset share $\alpha$"}
    col = {"HHI": C1, "rounded": C2, "transit": C3, "shared": C4, "asset": "#8C6D9E"}
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.8))
    for k in lab:
        th = [p["theta"] for p in ind["sweep"][k]]
        fl = [p["flagged"] for p in ind["sweep"][k]]
        ax[0].step(th, fl, where="post", color=col[k], label=lab[k], lw=1.1)
        t0 = ind["operating_points"][k]
        ax[0].plot([t0], [ind["flagged_at_operating_points"][k]], "o", color=col[k], ms=3.5)
    ax[0].set_xlabel(r"threshold $\theta$")
    ax[0].set_ylabel("subjects flagged (of 12)")
    ax[0].set_yticks(range(0, 13, 2))
    ax[0].set_title("(a) flag count across $\\theta$", loc="left")
    ax[0].legend(fontsize=6.4, loc="upper right")
    for k in lab:
        v = sorted(x[k] for x in ind["values"].values())
        ax[1].plot(v, range(1, len(v) + 1), drawstyle="steps-post", color=col[k], lw=1.1)
    ax[1].set_xlabel("indicator value")
    ax[1].set_ylabel("subjects with value $\\leq$ x")
    ax[1].set_title("(b) distribution over the 12 subjects", loc="left")
    fig.tight_layout()
    fig.savefig(os.path.join(out, "fig05_indicator_sensitivity.pdf"))


def fig06(results, out):
    r = load(results, "routes.json")
    sens = r["sensitivity"]
    ws = []
    for s in sens:
        k = tuple(s["w"])
        if k not in ws:
            ws.append(k)
    betas = sorted({s["beta"] for s in sens})
    M = np.zeros((len(ws), len(betas)))
    O = np.zeros((len(ws), len(betas)))
    for s in sens:
        M[ws.index(tuple(s["w"])), betas.index(s["beta"])] = s["kendall_tau"]
        O[ws.index(tuple(s["w"])), betas.index(s["beta"])] = s["top10_overlap"]
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.9))
    for a, Z, t in ((ax[0], M, r"(a) Kendall $\tau$ with the reference ranking"),
                    (ax[1], O, "(b) overlap of the top 10% of links")):
        im = a.imshow(Z, cmap="Greys", vmin=0.5, vmax=1.0, aspect="auto")
        for i in range(Z.shape[0]):
            for j in range(Z.shape[1]):
                a.text(j, i, f"{Z[i, j]:.2f}", ha="center", va="center", fontsize=6.8,
                       color="white" if Z[i, j] > 0.8 else "black")
        a.set_xticks(range(len(betas)))
        a.set_xticklabels([f"{b} d" for b in betas])
        a.set_xlabel(r"decay constant $\beta$")
        a.set_yticks(range(len(ws)))
        a.set_yticklabels([f"({w[0]:.2f}, {w[1]:.2f}, {w[2]:.2f})" for w in ws], fontsize=6.8)
        a.set_title(t, loc="left")
    ax[0].set_ylabel(r"weights $(\omega_1, \omega_2, \omega_3)$")
    fig.tight_layout()
    fig.savefig(os.path.join(out, "fig06_route_status.pdf"))


def fig07(results, out):
    s = load(results, "scaling.json")
    n = [x["rows"] for x in s["sizes"]]
    t = [x["median_s"] for x in s["sizes"]]
    th = [x["rows_per_s"] for x in s["sizes"]]
    a, b = s["fit"]["a"], s["fit"]["b"]
    fig, ax = plt.subplots(1, 2, figsize=(7.2, 2.6))
    ax[0].loglog(n, t, "o", color=C1, ms=4)
    xx = np.logspace(np.log10(min(n)), np.log10(max(n)), 50)
    ax[0].loglog(xx, a * xx ** b, color=C4, lw=0.9,
                 label=f"$t = {a:.2e}\\,n^{{{b:.3f}}}$")
    ax[0].set_xlabel("rows $n$")
    ax[0].set_ylabel("median time (s)")
    ax[0].legend(fontsize=7)
    ax[0].set_title("(a) end-to-end time", loc="left")
    ax[1].plot(n, th, "o-", color=C3, ms=4, lw=0.9)
    ax[1].set_xscale("log")
    ax[1].set_ylim(0, max(th) * 1.15)
    ax[1].set_xlabel("rows $n$")
    ax[1].set_ylabel("rows per second")
    ax[1].set_title("(b) throughput", loc="left")
    fig.tight_layout()
    fig.savefig(os.path.join(out, "fig07_throughput.pdf"))


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--results", default="results")
    ap.add_argument("--corpus", default="../corpus")
    ap.add_argument("--out", default="../figures")
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    fig02(a.results, a.corpus, a.out)
    fig03(a.results, a.out, os.path.join(a.results, "run"))
    fig04(a.results, a.out)
    fig05(a.results, a.out)
    fig06(a.results, a.out)
    fig07(a.results, a.out)
    print("figures written to", a.out)
