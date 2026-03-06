# PowerPoint MCP Server (v1)

PowerPoint MCP server with dual engine support:
- `COM` mode on Windows with Microsoft PowerPoint installed (highest fidelity)
- `OOXML` mode everywhere else via `python-pptx` and safe OOXML helpers

The TypeScript MCP server communicates with a local Python bridge over stdio JSON-RPC.

## Status

Implemented v1 Phase 1-2 surface:
- Session tools
- Discovery tools
- Slide tools
- Placeholder tools
- Slide snapshot tool
- MCP resources for presentations/slides/layouts/masters/theme/snapshots

Only implemented tools are registered.

## Runtime Targets

- Node.js `>=22`
- Python `>=3.12`

Note: local development can run with compatible lower patch/minor versions, but Node 22/Python 3.12 are the declared baseline.

## Install

```bash
npm install
python3 -m pip install -r requirements.txt -r requirements-dev.txt
pre-commit install --install-hooks
```

## Run

```bash
npm run dev
# or
npm run build && npm start
```

## Architecture

- `src/index.ts`: MCP stdio entrypoint
- `src/server.ts`: MCP tool/resource registration and validation
- `src/bridge/client.ts`: persistent Python subprocess client
- `src/tools/catalog.ts`: flat zod schemas for all v1 tools
- `python/bridge.py`: line-delimited JSON-RPC dispatcher
- `python/service.py`: engine selection + method routing
- `python/engines/ooxml_engine.py`: cross-platform implementation
- `python/engines/com_engine.py`: Windows COM implementation
- `python/com_worker.py`: STA-thread COM call isolation

## Implemented Tools

- `pptx_get_engine_info`
- `pptx_create_presentation`
- `pptx_open_presentation`
- `pptx_save_presentation`
- `pptx_close_presentation`
- `pptx_list_open_presentations`
- `pptx_get_presentation_state`
- `pptx_get_layouts`
- `pptx_get_layout_detail`
- `pptx_get_masters`
- `pptx_get_theme`
- `pptx_get_slide`
- `pptx_add_slide`
- `pptx_duplicate_slide`
- `pptx_delete_slide`
- `pptx_reorder_slides`
- `pptx_move_slide`
- `pptx_set_slide_background`
- `pptx_get_slide_snapshot`
- `pptx_get_placeholders`
- `pptx_set_placeholder_text`
- `pptx_set_placeholder_image`
- `pptx_clear_placeholder`
- `pptx_get_placeholder_text`

All mutating operations return `success` plus updated `presentation_state` (except close, which returns close confirmation).

## File Safety Model

- Opened/created presentations are edited in a temporary working copy.
- Original files are not modified until explicit `pptx_save_presentation`.
- Sessions are in-memory and do not survive process restart.

## Snapshots

- COM mode: native `Slide.Export`
- OOXML mode: `soffice --headless` + `pdftoppm`
- If dependencies are missing in OOXML mode, snapshot returns `dependency_missing` with install hint.

## Test

```bash
npm test
npm run test:py
```

## Quality and formatting

```bash
# Python lint and formatting
npm run lint:py
npm run format:check:py

# Full validation suite
npm run check

# Run all pre-commit hooks manually
npm run precommit:run
```

Python tool configuration is centralized in `pyproject.toml`:
- `black` for formatting
- `ruff` for linting/import sorting rules
- `pytest` defaults for test execution

## Security and secret protection

- Pre-commit secret scanning:
  - `detect-secrets` with `.secrets.baseline`
  - `detect-private-key` hook
- CI secret scanning with `gitleaks` (`.github/workflows/security.yml`)
- Recommended GitHub settings:
  - Enable branch protection on `main`
  - Require `ci` + `security` checks before merge
  - Enable secret scanning and push protection

Never commit real credentials. Use `.env.example` as the only tracked env template.

## Dependency automation

- `Dependabot` is configured in `.github/dependabot.yml` for npm, pip, and GitHub Actions updates.

## Windows COM Smoke Runner

Run this on a licensed Windows machine with PowerPoint installed to validate COM parity quickly.

PowerShell wrapper:

```powershell
.\scripts\run_windows_com_smoke.ps1
```

Python direct invocation:

```bash
PYTHONPATH=python python scripts/windows_com_smoke.py --output-dir /abs/path/to/artifacts/com-smoke
```

Useful options:
- `--input-pptx <absolute_path>`: run smoke checks against an existing deck
- `--layout-name \"Title Slide\"`: force a specific layout for add-slide
- `--skip-snapshot`: skip image snapshot validation
- `--allow-ooxml`: debug script flow on non-Windows hosts (not a COM parity check)

Outputs are written to `artifacts/com-smoke/` by default:
- `com_smoke_output_<timestamp>.pptx`
- `com_smoke_snapshot_<timestamp>.jpg` (unless skipped)
- `com_smoke_report_<timestamp>.json`
- `com_smoke_report_<timestamp>.md` (automated status + manual checklist)
