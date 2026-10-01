# SBSC-1: Synthetic Bank Statement Corpus (v1.1) and reference pipeline

Companion artefact to

> Kaliyev, A.; Yessengabulov, Z.; Shayakhmetov, M.; Rzayeva, L.
> *A Reproducible Normalization Model for Forensic Analysis of Heterogeneous
> Bank Statements: Design and Internal-Consistency Evaluation on a Synthetic
> Corpus.* Journal of Cybersecurity and Privacy (under review).

All data in this repository are generated. It contains no real persons,
organizations, account numbers, bank templates, transactions or case material.

## Layout

| Path | Content |
|---|---|
| `corpus/` | 30 statement files, ground truth, injection log, layout registry, request register, generator and verification script. See `corpus/README.md`. |
| `pipeline/` | Reference implementation of the normalization model (`sbsc_pipeline.py`), scorer (`evaluate.py`), experiments of the article (`experiments.py`, `extended_validation.py`), revision checks (`check_revision.py`) and figure script (`make_figures.py`). |
| `verified/` | Result files of the runs reported in the article. |

## Reproducing

Requires Python 3.10+, `numpy`, `openpyxl`, `matplotlib`.

```bash
cd corpus
python generate_corpus.py        # regenerates the corpus from the seed (byte-identical)
python verify_corpus.py          # composition, coverage, reference reconciliation, digests

cd ../pipeline
python evaluate.py --corpus ../corpus --out results
python experiments.py --corpus ../corpus --out results
python make_figures.py --results results --corpus ../corpus --out ../figures
```

## Scope of the evaluation

The corpus and the control layer were written by the same authors, so the
detection scores are an internal-consistency check, not evidence of
performance on real bank exports. See the article, Sections 3.10 and 6.

## License

CC BY 4.0. See `LICENSE`.
