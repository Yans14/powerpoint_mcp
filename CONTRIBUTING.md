# Contributing

## Development setup

1. Install dependencies:
   - `npm install`
   - `python3 -m pip install -r requirements.txt -r requirements-dev.txt`
2. Install pre-commit hooks:
   - `pre-commit install --install-hooks`

## Local validation before push

Run all checks:

```bash
npm run build
npm test
ruff check python
black --check python
npm run test:py
pre-commit run --all-files
```

Windows COM parity smoke run:

```powershell
.\scripts\run_windows_com_smoke.ps1
```

## Commit hygiene

- Keep commits focused and small.
- Use meaningful commit messages.
- Do not commit secrets, credentials, or private customer data.
