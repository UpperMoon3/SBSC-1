# SBSC-1 — Synthetic Bank Statement Corpus for normalization research

A generated corpus of heterogeneous bank statements with defects of known type
and position, released as the companion artefact to:

> *A Reproducible Normalization Model for Forensic Analysis of Heterogeneous
> Bank Statements.*

## What this is, and what it is not

Public benchmarks in anti-money-laundering research are transaction streams:
clean, uniform tables intended for training and evaluating detection models.
This corpus is a different object. It is a set of **statement files** as an
analyst receives them — differing column labels, three debit/credit encodings,
mixed date and decimal formats, partial counterparty attributes, and defects
injected on purpose. The target task is not detection. It is the step that
precedes detection: turning that stack of files into one comparable journal
without losing the ability to check any result against its source line.

The corpus contains **no real persons, organizations, account numbers, bank
templates, transactions or case material**. Every value is generated. Direct
identifiers are synthetic tokens rather than removed attributes, because
attribute removal alone is not a sufficient protection against
re-identification.

## Contents

```
generate_corpus.py           the generator; regenerates everything below
verify_corpus.py             recomputes the published figures and checks digests
statements/                  30 statement files, 15 layouts, 3 synthetic issuers,
                             and request_register.csv (account, holder, period per file)
ground_truth/
  normalized_journal.csv     the 1200 records a correct pipeline must produce
  injection_log.csv          the 57 injected defects, typed and positioned
  layout_registry.csv        the documented field mapping M_b per layout
  field_label_index.csv      every source column label and the field it maps to
  manifest_sha256.txt        SHA-256 of every generated file
reports/
  corpus_summary.json        composition
  coverage.csv               per-field coverage c_j
```

## Composition

| Parameter | Value |
|---|---|
| Statement files | 30 |
| Distinct source layouts | 15 |
| Synthetic issuers | 3 |
| Normalized transactions | 1200 |
| Account holders | 12 |
| Counterparties | 187 |
| Currencies | KZT, USD, EUR |
| Period covered | 6 months |
| Injected defects | 57 (53 row-level, 4 file-level) |

Defect classes: duplicate rows (14), direction inconsistency (9), balance
inconsistency (7), incomplete counterparty record (11), currency mixing (6),
unmapped field (6), unparsed layout (4).

## Deliberately undetectable defects

Two of the nine direction inconsistencies are placed in files that publish no
running balance, so no arithmetic check can reach them. This is intentional. A
validation set in which every injected defect is detectable measures the
generator rather than the tool under test. These cases are marked in
`injection_log.csv` so that a pipeline reporting them as undetectable is scored
correctly, and a pipeline reporting them as absent is not.

## Heterogeneity

Fields do not appear under a single label. Across the 15 layouts there are 105
distinct source column labels for 15 normalized fields — `subcategory` appears
under 14, `operation_date` under 11, `counterparty_name` under 10. Each layout
is individually readable; the mapping between them is the work.

Coverage is uneven by design and follows what the layouts print. The eight
mandatory fields reach 1.000; the four client fields come from
`statements/request_register.csv`, the record received with the files, since
most layouts print no holder details. `balance_after` reaches 0.799 because
three layouts print no running balance; `counterparty_name` reaches 0.939 and
`counterparty_identifier` 0.578. Every counterparty record carries a name or
an identifier unless an INCOMPLETE_COUNTERPARTY defect was injected.

## Reproducing and verifying

```bash
python generate_corpus.py          # regenerates the corpus from the seed
python verify_corpus.py            # recomputes composition, coverage, digests
```

`verify_corpus.py` also reports a **reference reconciliation**: what an ideal
arithmetic check would find on this corpus, given the injected defects and a
tolerance of one minor currency unit.

| Quantity | Value |
|---|---|
| Rows carrying `balance_after` | 959 of 1200 |
| Rows scorable (balance on both legs) | 947 |
| Rows exceeding the tolerance | 16 |
| Accounts raising a warning | 5 of 12 |
| Residual quantiles q50 / q90 / q99 | 0.000 / 0.000 / 0.95 |

Conforming rows reconcile exactly; real exports would not. Balance defects
come in two sub-modes, both fixed constants: an *isolated* mistyped balance
cell, which produces two residuals, and a *carried-forward* drift running to
the end of the account, which produces one. Balance-bearing defects are
concentrated in five of the twelve accounts, all in KZT.

## Reference pipeline

`../pipeline/` holds a reference implementation of the normalization model
(`sbsc_pipeline.py`), the scorer (`evaluate.py`), the experiments of the
revised article (`experiments.py`) and the figure script (`make_figures.py`):

```bash
cd ../pipeline
python evaluate.py --corpus ../corpus --out results
python experiments.py --corpus ../corpus --out results
python make_figures.py --results results --corpus ../corpus --out ../figures
```

## Changes in version 1.1

Version 1.0 had a ground truth that did not match its own files in three
respects, found by running the reference pipeline against it:

* it held running balances for four files whose layout prints none, so its
  `balance_after` coverage (0.920) and reference reconciliation counted
  values no reader of the files could see;
* it added a sub-unit display-rounding offset to four layouts that the
  two-decimal files could not show;
* its currency-mixing rows left no trace in the file.

Version 1.1 removes all three, prints a currency code after the amount of a
foreign-currency row, records the structural features of each layout in the
registry so that a layout obstacle is a departure from the documented layout,
adds the request register, never lets a coverage mask blank both the name and
the identifier of a counterparty, and fixes a bug that wrote an invalid
`docProps/core.xml` into every workbook. The corpus is regenerated from the
same seed.

## Suggested use

Report, at minimum: traceability of the produced journal to source lines,
per-field coverage, and per-class detection against `injection_log.csv`.
Report undetected defects separately from defects reported as absent — the
distinction is the point of the corpus.

Scoring against a ground truth authored by the same group that authored the
tool is an internal consistency check, not an independent evaluation. Where a
corpus is used to evaluate a tool built by its own authors, that should be
stated.

## Requirements

Python 3.10+, `numpy`, `openpyxl`.

## License

CC BY 4.0. See `LICENSE`.

## Distinguishing latent and observed field coverage

The coverage values 0.939 (name) and 0.578 (identifier) above describe
the latent ground-truth journal, including values deliberately removed from
statement files. The reference pipeline reads those files and reports 0.930
and 0.573, respectively. Its channel/account and type coverages are 0.759
and 0.882. These are different denominators of knowledge, not different
versions of the observed data. The article reports observed coverage.
