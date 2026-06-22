# Contributing

Thanks for your interest in contributing!

## Development setup

```bash
git clone https://github.com/N1k4G/ha-google-outlook-contacts-sync
cd ha-google-outlook-contacts-sync
python -m venv .venv && source .venv/bin/activate  # or .venv\Scripts\activate on Windows
pip install -r requirements_test.txt
pre-commit install
```

Generate a secrets baseline on first run:

```bash
detect-secrets scan > .secrets.baseline
```

## Running tests

```bash
pytest --cov=custom_components/google_outlook_contacts_sync
```

## Code style

- `ruff` for linting and formatting, `mypy --strict` for types.
- No secrets in code, logs, or test fixtures — anonymize all contact data.
- Conventional Commits (`feat:`, `fix:`, `chore:`, etc.).

## Pull requests

- Open an issue first for significant changes.
- All CI checks must pass before merge.
- Add or update tests for any changed behaviour.
