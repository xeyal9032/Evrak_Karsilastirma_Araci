# Frequently Asked Questions

## Which file formats are supported?

DATEV Buchungsstapel (CSV or Excel) and Journal Excel exports with columns such
as `Belegdat.`, `Sollkto`, `Habenkto`, and `Betrag`. PDF scans and OCR are not
supported.

## Can I run this without the GUI?

Yes:

```bash
python karsilastir.py file1.csv file2.xlsx -o report.xlsx --html --pdf
python cli.py file1.csv file2.xlsx --lang en
```

Exit codes: `0` success (Excel written; HTML/PDF failures are warnings on stderr),
`1` format/IO error, `2` unexpected error before/during Excel write.

## HTML / PDF reports?

- `--html` or GUI **HTML** checkbox → single-file interactive report (filter by status)
- `--pdf` or GUI **PDF** checkbox → summary PDF (`pip install reportlab`)

## Batch / folder compare?

CLI:

```bash
python karsilastir.py folder_a folder_b --batch -o batch_out --html --fast
```

GUI: **Batch folders…** (same pairing; respects Fast / HTML / PDF / Archive checkboxes).

Pairs files by matching basename (`--batch-mode stem`, default) or sorted zip order.

## SQLite archive?

```bash
python karsilastir.py a.csv b.xlsx --archive data/compare_archive.db
```

GUI: **Save to archive** checkbox (default on) and **Archive…** to browse recent runs.
Archive DB defaults to `data/compare_archive.db` next to the app (source tree or exe
folder — same path for single compare, batch, CLI `--archive`, and Archive browser).
Archive failures show a warning (they no longer fail silently).

## Where is the report saved?

GUI mode writes to the Desktop by default. CLI mode writes next to the current
working directory unless you pass `-o` / `--output`.

## How do I filter only red or orange rows?

Excel: AutoFilter on `Durum`, or sheets `FILTRE-KIRMIZI` / `FILTRE-TURUNCU`.  
HTML: use the filter buttons at the top of the report.

## Community vs Enterprise?

| Feature | Community (MIT) | Enterprise (planned, paid) |
|---------|-----------------|----------------------------|
| Single-pair GUI / CLI | Yes | Yes |
| Excel / HTML / PDF, 4 languages, logs | Yes | Yes |
| Folder / batch compare | Yes | Yes + SLA |
| SQL result archive | Yes | Yes + SLA |
| Signed commercial builds / SLA | — | Planned |

See `LICENSE` and `LICENSE.COMMERCIAL`.

## Logs?

Each comparison appends to `logs/compare_YYYYMMDD.log` next to the app
(source tree or the folder containing the `.exe` — never under PyInstaller’s
temporary `_MEIPASS`). Archive DB defaults to `data/compare_archive.db` there.

## Benchmarks?

```bash
python benchmarks/run_benchmark.py --rows 5000 --fast
python benchmarks/run_benchmark.py --rows 100000 --fast
```

Example (this machine): ~5 000 rows, `--fast` ≈ 2.3 s compare+Excel write.
