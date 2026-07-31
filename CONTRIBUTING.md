# Contributing

Thanks for improving Evrak Karsilastirma Araci.

## Setup

```bash
git clone https://github.com/xeyal9032/Evrak_Karsilastirma_Araci.git
cd Evrak_Karsilastirma_Araci
python -m venv .venv
# Windows:
.\.venv\Scripts\activate
pip install -r requirements.txt
```

Optional drag-and-drop on Windows:

```bash
pip install windnd
```

## Tests

```bash
python -m unittest discover -s tests -v
```

CI runs the same command on every push / pull request. Tests that need local
desktop sample folders are skipped automatically when those paths are missing.

## Adding a language

1. Copy `locales/en.json` to `locales/<code>.json`
2. Translate all string values (keep keys identical)
3. Register the code in `i18n.SUPPORTED` if needed
4. Run tests; `tests/test_i18n.py` fails if keys drift between locales

## Code style

- Prefer small, focused changes
- Keep `karsilastir_motor.py` free of Tkinter / CLI UI details
- Pass optional `progress_cb` and `lang` into the motor instead of globals
- Do not commit `logs/`, `dist/`, `.venv*`, or generated Excel reports

## Pull requests

- Describe the problem and how you verified the fix
- Update `CHANGELOG.md` under `[Unreleased]` or the next version section
- Keep Community features under MIT; do not introduce paid license locks in PRs
  unless maintainers request Enterprise work explicitly
