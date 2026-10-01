#!/usr/bin/env python3
"""
Reference implementation of the normalization model (Stages 1-7) for the
SBSC-1 corpus.

This is evaluation code written against synthetic data. It implements the
equations of the article as stated and contains no bank-specific parsing logic
beyond the documented layout registry M_b that ships with the corpus.

    python sbsc_pipeline.py --corpus ../corpus --out run/

Outputs (in --out):
    inventory.csv        Stage 1-3: file inventory with SHA-256
    journal.csv          Stage 6: normalized journal with lineage tuple
    defects.csv          Stage 7: defect log (candidates, never removals)
    summary.json         control-layer figures
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import shutil
import stat
import tempfile
import time
from collections import Counter, defaultdict
from datetime import datetime

import openpyxl

PROCESSING_VERSION = "sbsc-ref-1.0.1"
FIELDS = [
    "client_code", "client_identifier", "client_account", "statement_period",
    "operation_date", "direction", "amount", "currency", "balance_after",
    "category", "subcategory", "counterparty_name", "counterparty_identifier",
    "counterparty_account_or_channel", "counterparty_type",
]
LINEAGE = ["source_file_id", "source_sheet", "source_row",
           "processing_version", "defect_flag"]
LEGAL_FORMS = r"\b(llp|llc|jsc|ltd|too|ao|ip|inc|plc|gmbh)\b"


# ------------------------------------------------------------------ helpers
def sha256(path):
    h = hashlib.sha256()
    with open(path, "rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 16), b""):
            h.update(chunk)
    return h.hexdigest()


def nu(s):
    """String normalization nu(.) of Eq. (dupkey)."""
    if not s:
        return ""
    s = s.lower()
    s = re.sub(r"[^\w\s]", " ", s)
    s = re.sub(LEGAL_FORMS, " ", s)
    return " ".join(s.split())


CCY_CODE = re.compile(r"\s*([A-Z]{3})\s*$")


def split_code(v):
    """An amount cell may carry an ISO currency code after the number."""
    if isinstance(v, str):
        m = CCY_CODE.search(v)
        if m:
            return v[:m.start()], m.group(1)
    return v, None


def parse_amount(v, decimal_comma):
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace(" ", " ").strip()
    if decimal_comma:
        s = s.replace(" ", "").replace(".", "").replace(",", ".")
    else:
        s = s.replace(" ", "").replace(",", "")
    return float(s)


def read_registry(corpus):
    reg = {}
    with open(os.path.join(corpus, "ground_truth", "layout_registry.csv"),
              newline="", encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            r["labels"] = {f: r[f"label:{f}"] for f in FIELDS}
            r["decimal_comma"] = r["decimal_comma"] == "True"
            r["has_currency_column"] = r["has_currency_column"] == "True"
            r["has_balance_column"] = r["has_balance_column"] == "True"
            for k in ("has_preamble", "has_summary_rows", "two_row_header"):
                r[k] = r[k] == "True"
            reg[r["layout_id"]] = r
    return reg


# ------------------------------------------------------------ Stages 1 - 4
def stages_1_to_4(corpus, work):
    src = os.path.join(corpus, "statements")
    orig = os.path.join(work, "ORIGINAL")
    wc = os.path.join(work, "WORKING_COPY")
    os.makedirs(orig)
    os.makedirs(wc)
    register = {}
    with open(os.path.join(src, "request_register.csv"), newline="",
              encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            register[r["source_file_id"]] = r
    inventory = []
    for name in sorted(os.listdir(src)):
        if not name.endswith(".xlsx"):
            continue
        p = os.path.join(orig, name)
        shutil.copy2(os.path.join(src, name), p)
        os.chmod(p, stat.S_IRUSR | stat.S_IRGRP)          # read-only original
        shutil.copy2(p, os.path.join(wc, name))
        os.chmod(os.path.join(wc, name), stat.S_IRUSR | stat.S_IWUSR)
        fid, lay, acct = name[:-5].split("_", 2)
        reg = register[fid]
        inventory.append({"source_file_id": fid, "layout_id": lay,
                          "declared_account": reg["client_account"],
                          "client_code": reg["client_code"],
                          "client_identifier": reg["client_identifier"],
                          "declared_period": reg["declared_period"],
                          "file_name": name,
                          "size": os.path.getsize(p), "sha256": sha256(p)})
    return inventory, wc


# ------------------------------------------------------------ Stages 5 - 6
def locate_header(rows, labels):
    """Return index of the header row: the first row in which at least four
    cells are documented labels of this layout (Dr/Cr suffixes allowed)."""
    known = set(labels.values())
    for i, r in enumerate(rows):
        cells = [str(c).replace(" (Dr)", "").replace(" (Cr)", "")
                 for c in r if c is not None]
        if sum(c in known for c in cells) >= 4:
            return i
    return None


def normalize_file(path, inv, layout, defects):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    rows = list(ws.iter_rows(values_only=True))
    lab = layout["labels"]
    fid = inv["source_file_id"]
    h = locate_header(rows, lab)
    if h is None:
        defects.append({"source_file_id": fid, "source_row": "",
                        "code": "HEADER_NOT_FOUND", "class": "UNPARSED_LAYOUT",
                        "detail": "no row matches the documented labels"})
        return [], {}

    # file-level structure: anything that is not the header followed directly
    # by transaction rows is logged as a layout obstacle
    meta = {}
    obstacles = []
    for r in rows[:h]:
        if r[0] is not None and str(r[0]) == "Summary":
            break                       # summary block, not holder metadata
        if r[0] is not None and r[1] is not None:
            meta[str(r[0])] = r[1]
    if h > 0:
        pre = [r for r in rows[:h] if any(c is not None for c in r)]
        # structure is compared with the documented layout; only a departure
        # from it is a layout obstacle, a documented feature is not
        if any(str(r[0]) == "Summary" for r in pre) and not layout["has_summary_rows"]:
            obstacles.append("summary_rows_before_table")
        if any(str(r[0]).startswith("Statement of") for r in pre) and not layout["two_row_header"]:
            obstacles.append("two_row_header")
        if meta and not layout["has_preamble"]:
            obstacles.append("metadata_preamble")

    header = rows[h]
    col = {}
    unlabelled = []
    for j, c in enumerate(header):
        if c is None:
            unlabelled.append(j)
            continue
        c = str(c)
        for f, l in lab.items():
            if c == l:
                col[f] = j
            elif c == f"{l} (Dr)":
                col["amount_dr"] = j
            elif c == f"{l} (Cr)":
                col["amount_cr"] = j
    enc = layout["direction_encoding"]
    need = {"split": ["operation_date", "amount_dr", "amount_cr"],
            "signed": ["operation_date", "amount"],
            "flag": ["operation_date", "direction", "amount"]}[enc]
    if layout["has_currency_column"]:
        need.append("currency")
    if layout["has_balance_column"]:
        need.append("balance_after")
    missing = [f for f in need if f not in col]
    for f in missing:
        defects.append({"source_file_id": fid, "source_row": "",
                        "code": "UNMAPPED_MANDATORY", "class": "",
                        "detail": f"documented column for {f} not found"})

    def cell(r, f):
        j = col.get(f)
        return r[j] if j is not None and j < len(r) else None

    out = []
    body = rows[h + 1:]
    date_fmt = layout["date_format"]
    dc = layout["decimal_comma"]
    for k, r in enumerate(body):
        excel_row = h + 2 + k
        if all(c is None for c in r):
            continue
        first = cell(r, "operation_date")
        try:
            d = datetime.strptime(str(first), date_fmt).date()
        except ValueError:
            obstacles.append("trailing_non_transaction_rows")
            defects.append({"source_file_id": fid, "source_row": excel_row,
                            "code": "NON_TRANSACTION_ROW", "class": "UNPARSED_LAYOUT",
                            "detail": f"row starting {first!r} is not a transaction"})
            continue
        cell_ccy = None
        direction = amount = None
        problem = None
        try:
            if enc == "split":
                vdr, c1 = split_code(cell(r, "amount_dr"))
                vcr, c2 = split_code(cell(r, "amount_cr"))
                cell_ccy = c1 or c2
                dr = parse_amount(vdr, dc)
                cr = parse_amount(vcr, dc)
                if dr is not None and cr is not None:
                    problem = "both debit and credit populated"
                direction = "debit" if dr is not None else ("credit" if cr is not None else None)
                amount = dr if dr is not None else cr
            elif enc == "signed":
                v, cell_ccy = split_code(cell(r, "amount"))
                v = parse_amount(v, dc)
                if v is not None:
                    direction = "credit" if v >= 0 else "debit"
                    amount = abs(v)
            else:
                flag = str(cell(r, "direction") or "").upper()
                direction = {"CR": "credit", "DR": "debit"}.get(flag)
                if direction is None:
                    problem = f"direction token {flag!r} not recognized"
                v, cell_ccy = split_code(cell(r, "amount"))
                amount = parse_amount(v, dc)
            bal = parse_amount(cell(r, "balance_after"), dc) if "balance_after" in col else None
        except (ValueError, TypeError) as e:
            problem = f"value not parseable: {e}"
            bal = None
        if amount is None and problem is None:
            problem = "no amount"
        rec = {
            "client_code": inv["client_code"],
            "client_identifier": inv["client_identifier"],
            "client_account": inv["declared_account"],
            "statement_period": inv["declared_period"],
            "operation_date": d.isoformat(),
            "direction": direction,
            "amount": round(amount, 2) if amount is not None else None,
            "currency": (cell(r, "currency") if layout["has_currency_column"]
                         else cell_ccy),
            "balance_after": bal,
        }
        for f in ("category", "subcategory", "counterparty_name",
                  "counterparty_identifier", "counterparty_account_or_channel",
                  "counterparty_type"):
            rec[f] = cell(r, f)
        rec.update({"source_file_id": fid, "source_sheet": ws.title,
                    "source_row": excel_row,
                    "processing_version": PROCESSING_VERSION,
                    "defect_flag": ""})
        rec["_unlabelled_values"] = [r[j] for j in unlabelled
                                     if j < len(r) and r[j] not in (None, "")]
        rec["_cell_ccy"] = cell_ccy
        rec["_eval_file"], rec["_eval_row"] = fid, excel_row   # scoring only
        if problem:
            # the row is kept with what could be read; the problem is logged
            defects.append({"source_file_id": fid, "source_row": excel_row,
                            "code": "ROW_NOT_PARSED", "class": "",
                            "detail": problem, "_i_local": len(out)})
        out.append(rec)

    for kind in sorted(set(obstacles)):
        defects.append({"source_file_id": fid, "source_row": "",
                        "code": "LAYOUT_OBSTACLE", "class": "UNPARSED_LAYOUT",
                        "detail": kind})
    return out, meta


def fill_account_attributes(journal):
    """Client attributes printed in one file's preamble are carried to the other
    files of the same account; the period falls back to the observed date span
    of the file. Both are recorded, never guessed from outside the corpus."""
    by_acct = defaultdict(dict)
    for r in journal:
        for f in ("client_code", "client_identifier"):
            if r[f]:
                by_acct[r["client_account"]][f] = r[f]
    span = defaultdict(list)
    for r in journal:
        span[r["source_file_id"]].append(r["operation_date"])
    for r in journal:
        for f in ("client_code", "client_identifier"):
            if not r[f]:
                r[f] = by_acct[r["client_account"]].get(f)
        if not r["statement_period"]:
            ds = span[r["source_file_id"]]
            r["statement_period"] = f"{min(ds)}/{max(ds)}"


def fill_currency(journal):
    """A sheet without a currency column inherits the currency of the account,
    taken from the account's other files. Rows so filled are marked, because
    the sheet itself does not state the currency (Section 3.4.4)."""
    acct_ccy = defaultdict(Counter)
    for r in journal:
        if r["currency"] and not r["_cell_ccy"]:
            acct_ccy[r["client_account"]][r["currency"]] += 1
    for r in journal:
        r["_currency_inferred"] = False
        r["_account_ccy"] = (acct_ccy[r["client_account"]].most_common(1)[0][0]
                             if acct_ccy[r["client_account"]] else None)
        if not r["currency"]:
            r["currency"] = r["_account_ccy"]
            r["_currency_inferred"] = True


# ------------------------------------------------------------------ Stage 7
def qc_reconciliation(journal, defects, tau):
    per = defaultdict(list)
    for i, r in enumerate(journal):
        per[r["client_account"]].append(i)
    residuals = []
    no_bal = 0
    for acct, idx in per.items():
        idx.sort(key=lambda i: (journal[i]["source_file_id"], journal[i]["source_row"]))
        prev = None
        for i in idx:
            r = journal[i]
            if r["balance_after"] is None:
                prev = None
                no_bal += 1
                defects.append({"source_file_id": r["source_file_id"],
                                "source_row": r["source_row"], "code": "NO_BALANCE_FIELD",
                                "class": "", "detail": "reconciliation unavailable",
                                "_i": i})
                continue
            if prev is None:
                no_bal += 1
                defects.append({"source_file_id": r["source_file_id"],
                                "source_row": r["source_row"], "code": "NO_BALANCE_FIELD",
                                "class": "", "detail": "no predecessor balance",
                                "_i": i})
                prev = r["balance_after"]
                continue
            if r["amount"] is None or r["direction"] is None:
                prev = None
                continue
            s = r["amount"] if r["direction"] == "credit" else -r["amount"]
            eps = round(r["balance_after"] - (prev + s), 4)
            residuals.append({"i": i, "acct": acct, "ccy": r["currency"], "eps": eps})
            if abs(eps) > tau:
                # a flipped sign produces a residual of exactly 2|a|
                if abs(abs(eps) - 2 * r["amount"]) <= tau:
                    cls, code = "DIRECTION_INCONSISTENCY", "SIGN_RESIDUAL"
                else:
                    cls, code = "BALANCE_INCONSISTENCY", "BALANCE_RESIDUAL"
                defects.append({"source_file_id": r["source_file_id"],
                                "source_row": r["source_row"], "code": code,
                                "class": cls, "detail": f"eps={eps}", "_i": i})
            prev = r["balance_after"]
    return residuals, no_bal


def qc_chronology(journal, defects):
    """Source order within an account must be non-decreasing in date; a
    violation is logged on the row where the order breaks."""
    per = defaultdict(list)
    for i, r in enumerate(journal):
        per[r["client_account"]].append(i)
    bad = 0
    for idx in per.values():
        idx.sort(key=lambda i: (journal[i]["source_file_id"], journal[i]["source_row"]))
        for a, b in zip(idx, idx[1:]):
            if journal[b]["operation_date"] < journal[a]["operation_date"]:
                bad += 1
                defects.append({"source_file_id": journal[b]["source_file_id"],
                                "source_row": journal[b]["source_row"],
                                "code": "ORDER_VIOLATION", "class": "",
                                "detail": f"date {journal[b]['operation_date']} precedes "
                                          f"{journal[a]['operation_date']}", "_i": b})
    return bad


def dup_key(r, with_name=True):
    amt = "" if r["amount"] is None else f"{r['amount']:.2f}"
    parts = [r["client_account"], r["operation_date"], amt, str(r["direction"])]
    if with_name:
        parts.append(nu(r["counterparty_name"]))
    return hashlib.sha256("|".join(parts).encode()).hexdigest()


def qc_duplicates(journal, defects, fallback):
    groups = defaultdict(list)
    for i, r in enumerate(journal):
        groups[dup_key(r)].append(i)
    if fallback:
        # second pass: rows whose name is missing are compared without it
        g2 = defaultdict(list)
        for i, r in enumerate(journal):
            g2[dup_key(r, with_name=False)].append(i)
        for k, idx in g2.items():
            if len(idx) > 1 and any(not journal[i]["counterparty_name"] for i in idx):
                groups["fb:" + k] = idx
    seen = set()
    for idx in groups.values():
        if len(idx) < 2:
            continue
        idx = sorted(idx)
        for later in idx[1:]:
            if later in seen:
                continue
            seen.add(later)
            a, b = journal[idx[0]], journal[later]
            if a["source_file_id"] == b["source_file_id"]:
                label = "repeat"
            else:
                try:
                    a0, a1 = a["statement_period"].split("/")
                    b0, b1 = b["statement_period"].split("/")
                    overlap = max(a0, b0) <= a["operation_date"] <= min(a1, b1)
                except (AttributeError, ValueError, TypeError):
                    overlap = False
                label = "overlap" if overlap else "cross_file_review"
            defects.append({"source_file_id": b["source_file_id"],
                            "source_row": b["source_row"],
                            "code": f"DUP_CANDIDATE_{label.upper()}",
                            "class": "DUPLICATE_ROW",
                            "detail": f"shares key with {a['source_file_id']}:{a['source_row']}",
                            "_i": later})


def qc_counterparty(journal, defects):
    """A record is incomplete when neither a name nor an identifier is present."""
    for i, r in enumerate(journal):
        if not r["counterparty_name"] and not r["counterparty_identifier"]:
            defects.append({"source_file_id": r["source_file_id"],
                            "source_row": r["source_row"], "code": "CP_INCOMPLETE",
                            "class": "INCOMPLETE_COUNTERPARTY",
                            "detail": ("channel only" if r["counterparty_account_or_channel"]
                                       else "no counterparty attribute"), "_i": i})


def qc_unmapped(journal, defects):
    for i, r in enumerate(journal):
        if r["_unlabelled_values"]:
            defects.append({"source_file_id": r["source_file_id"],
                            "source_row": r["source_row"], "code": "UNMAPPED_VALUE",
                            "class": "UNMAPPED_FIELD",
                            "detail": f"value {r['_unlabelled_values'][0]!r} in unlabelled column",
                            "_i": i})


def qc_currency(journal, defects):
    """Currency partitioning (Eq. currency). A row whose currency is not stated
    by its own sheet cannot be placed in a partition with certainty; it is
    logged so that no total silently includes it."""
    for i, r in enumerate(journal):
        if r["_currency_inferred"]:
            defects.append({"source_file_id": r["source_file_id"],
                            "source_row": r["source_row"], "code": "CURRENCY_NOT_STATED",
                            "class": "", "detail": "currency inherited from account",
                            "_i": i})
        if r["_account_ccy"] and r["currency"] != r["_account_ccy"]:
            defects.append({"source_file_id": r["source_file_id"],
                            "source_row": r["source_row"], "code": "FOREIGN_CURRENCY_ROW",
                            "class": "CURRENCY_MIXING",
                            "detail": f"{r['currency']} row in {r['_account_ccy']} account",
                            "_i": i})


STAGES = ("hash", "reconciliation", "duplicates", "currency", "lineage",
          "defect_log")


def run(corpus, out, tau=0.01, fallback=True, keep_work=False, disable=(),
        tamper=None):
    """disable: stages to remove for the ablation (see STAGES).
    tamper: file name whose working copy is altered after receipt, to test
    whether the run notices (integrity test for hash control)."""
    disable = set(disable)
    t0 = time.perf_counter()
    reg = read_registry(corpus)
    work = tempfile.mkdtemp(prefix="sbsc_")
    inventory, wc = stages_1_to_4(corpus, work)
    journal, defects = [], []
    if tamper:
        wb = openpyxl.load_workbook(os.path.join(wc, tamper))
        ws = wb.worksheets[0]
        ws.cell(row=ws.max_row, column=2, value="1")      # one altered cell
        wb.save(os.path.join(wc, tamper))
    integrity_failures = []
    for inv in inventory:
        path = os.path.join(wc, inv["file_name"])
        if "hash" not in disable and sha256(path) != inv["sha256"]:
            integrity_failures.append(inv["source_file_id"])
            defects.append({"source_file_id": inv["source_file_id"],
                            "source_row": "", "code": "HASH_MISMATCH",
                            "class": "", "detail": "working copy differs from original"})
        rows, _ = normalize_file(path, inv, reg[inv["layout_id"]], defects)
        journal.extend(rows)
    fill_currency(journal)
    if "reconciliation" in disable:
        residuals, no_bal = [], 0
    else:
        residuals, no_bal = qc_reconciliation(journal, defects, tau)
    chrono_bad = qc_chronology(journal, defects)
    if "duplicates" not in disable:
        qc_duplicates(journal, defects, fallback)
    qc_counterparty(journal, defects)
    qc_unmapped(journal, defects)
    if "currency" not in disable:
        qc_currency(journal, defects)
    for d in defects:
        if "_i" in d and d["class"]:
            r = journal[d["_i"]]
            r["defect_flag"] = ";".join(filter(None, [r["defect_flag"], d["code"]]))
    if "lineage" in disable:
        for r in journal:
            for k in ("source_file_id", "source_sheet", "source_row",
                      "processing_version"):
                r[k] = None
    log_visible = "defect_log" not in disable
    elapsed = time.perf_counter() - t0
    if not keep_work:
        for root, dirs, files in os.walk(work):
            for f in files:
                os.chmod(os.path.join(root, f), stat.S_IRUSR | stat.S_IWUSR)
        shutil.rmtree(work)

    if out:
        os.makedirs(out, exist_ok=True)
        with open(os.path.join(out, "inventory.csv"), "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=list(inventory[0]))
            w.writeheader()
            w.writerows(inventory)
        with open(os.path.join(out, "journal.csv"), "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=FIELDS + LINEAGE, extrasaction="ignore")
            w.writeheader()
            w.writerows(journal)
        with open(os.path.join(out, "defects.csv"), "w", newline="") as fh:
            w = csv.DictWriter(fh, fieldnames=["source_file_id", "source_row",
                                               "code", "class", "detail"],
                               extrasaction="ignore")
            w.writeheader()
            w.writerows(defects)
    return {"journal": journal, "defects": defects, "inventory": inventory,
            "residuals": residuals, "no_balance_rows": no_bal,
            "chronology_violations": chrono_bad, "elapsed_s": elapsed,
            "integrity_failures": integrity_failures, "log_visible": log_visible,
            "disabled": sorted(disable)}


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--corpus", required=True)
    ap.add_argument("--out", default="run")
    ap.add_argument("--tau", type=float, default=0.01)
    ap.add_argument("--no-fallback", action="store_true")
    a = ap.parse_args()
    res = run(a.corpus, a.out, a.tau, not a.no_fallback)
    print(json.dumps({"rows": len(res["journal"]), "defect_log": len(res["defects"]),
                      "elapsed_s": round(res["elapsed_s"], 3)}, indent=2))
