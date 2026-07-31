# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.3.1] - 2026-08-01

### Fixed
- Reject Excel scientific-notation Belegfeld1 (`1,22001E+11`) so different invoices are not falsely matched
- Do not use empty Belegfeld in key1 matching
- Skip match candidates with conflicting invoice numbers in text (`ref_ids`), including clone-pass pairs
- Match konto-recoded rows that share a long digit id in Buchungstext (Vodafone payment case)
- Extract ref ids without gluing space-separated numbers; still match spaced vs compact beleg (`2026 130000011` ↔ `2026130000011`)

## [0.3.0] - 2026-07-31

### Added
- HTML single-file interactive report (`--html`, GUI checkbox) with status filters
- PDF summary report (`--pdf`, requires reportlab)
- Folder batch compare (`--batch`, `--batch-mode stem|zip`)
- SQLite run archive (`--archive`, GUI auto-save to `data/`)
- Fast mode (`--fast`) skips per-file detail sheets for large runs
- Benchmark script `benchmarks/run_benchmark.py`
- Tests in `tests/test_faz23.py`

## [0.2.0] - 2026-07-31

### Added
- CLI mode: `python karsilastir.py file1.csv file2.xlsx` (also `python cli.py`)
- i18n: Turkish, Russian, German, English (`locales/`)
- Progress callback in the comparison engine and GUI progress bar
- Drag-and-drop file selection (optional `windnd` on Windows)
- Per-run logging under `logs/`
- Excel `Durum` column + AutoFilter; `FILTRE-KIRMIZI` / `FILTRE-TURUNCU` sheets
- Dual-license docs: MIT Community (`LICENSE`) + commercial draft (`LICENSE.COMMERCIAL`)
- `CONTRIBUTING.md`, `docs/FAQ.md`
- Deep regression tests for filters, progress stages, empty sheets, CLI, i18n, GUI lang (`tests/test_phase1_deep.py`)

### Fixed
- Excel loaders skip empty header sheets and continue scanning later sheets

## [0.1.0] - 2026-07-31

### Added
- DATEV / Journal format detection and multi-tier matching
- Colored Excel report and Tkinter GUI (TR/RU)
- Unittest suite and GitHub Actions CI
