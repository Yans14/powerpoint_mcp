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

You can do the same with:

```bash
make install
make precommit-install
```

## Run

```bash
npm run dev
# or
npm run build && npm start
```

Make shortcuts:

```bash
make mcp
make mcp-start
```

## Web UI (Local Test Console)

The repo includes a lightweight FastAPI + static HTML UI for interactive testing of tools, sessions, and logs.

Run it from the project root:

```bash
python3 -m uvicorn web_ui:app --app-dir python --host 127.0.0.1 --port 8000 --reload
```

Then open:

- `http://127.0.0.1:8000`

What it exposes:

- Tool catalog browser grouped by category (`GET /api/tools`)
- Engine info (`GET /api/engine`)
- Open sessions list (`GET /api/sessions`)
- Generic method dispatcher (`POST /api/dispatch`)
- Live backend logs over WebSocket (`/ws/logs`)

Relevant files:

- `python/web_ui.py` (FastAPI app + API routes + WebSocket log broadcast)
- `python/tool_catalog.py` (UI-facing tool metadata/catalog shaping)
- `web/index.html` (single-page test UI)

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
- In OOXML mode, edits stay in memory and the working copy is written only when something needs the file
  (`pptx_save_presentation`, `pptx_get_theme`, snapshots, agent rollback). Unsaved edits are lost if the bridge dies.
- Working files live in `<tmp>/pptx-mcp/bridge-<pid>-*/` and are removed on shutdown; leftovers from a
  killed process are removed the next time the server starts.

## Failure handling

- A bridge call that exceeds its time budget returns `timeout`; the Python bridge is stopped and restarted on the next call.
- After a bridge restart or crash, calls using a presentation id from the previous process return `session_lost`:
  reopen the file and use the new id.
- A missing Python interpreter is reported as `dependency_missing` at startup.

## Environment variables

| Variable | Default | Purpose |
| --- | --- | --- |
| `PPTX_PYTHON_BIN` | `python3` | Python interpreter used to run the bridge |
| `PPTX_PY_BRIDGE` | bundled `python/bridge.py` | Alternative bridge script |
| `PPTX_BRIDGE_TIMEOUT_MS` | `120000` | Time budget for one bridge call (`pptx_agent_*` tools get at least 600000) |
| `PPTX_RENDER_TIMEOUT_SEC` | `90` | Time budget for rendering a snapshot request (LibreOffice + pdftoppm) |
| `PPTX_LOG_LEVEL` | `INFO` | Bridge log level on stderr (`DEBUG`, `INFO`, `WARNING`, ...) |
| `PPTX_WORKSPACE_DIR` | `<tmp>/pptx-mcp` | Root directory for session working files |
| `PPTX_TOOL_PROFILE` | `full` | `core` exposes 31 essential tools instead of all 69 (about half the tool-list size) |

## MCP surface

- **Tools** carry explicit annotations (`readOnlyHint`, `destructiveHint`, `idempotentHint`, `openWorldHint`).
  `pptx_save_presentation` is marked destructive: it overwrites `output_path`.
- **Responses** are compact JSON. Mutating tools return `slide_count`; pass `include_state: true` to also get
  the per-slide `presentation_state`.
- **`pptx_batch`** runs up to 50 operations on one presentation in a single call. All operations are validated
  first; with `atomic` (default) a failure restores the presentation (OOXML engine).
  Saving, closing and agent tools are not allowed inside a batch.
- **Errors** are `{code, message, details?, hint?}`. `hint` names the tool to call next
  (for example `pptx_get_layouts` after an unknown layout).
- **Structured output**: `pptx_list_open_presentations`, `pptx_get_presentation_state`, `pptx_get_layouts`,
  `pptx_get_layout_detail` and `pptx_get_placeholders` declare an `outputSchema` and return `structuredContent`.
- **Resources**: `pptx://presentations` lists open presentations; slides, layouts, masters, theme and
  snapshots (`image/jpeg` blob) are resource templates under `pptx://presentations/{presentation_id}`.
- **Prompts**: `create_deck_from_outline`, `apply_template_to_deck`.

## Snapshots

- COM mode: native `Slide.Export`
- OOXML mode: `soffice --headless` + `pdftoppm`. The PDF conversion is cached until the presentation changes.
- Pass `slide_indices` instead of `slide_index` to get up to 20 slides in one call.
- If dependencies are missing in OOXML mode, snapshot returns `dependency_missing` with install hint.

## Test

```bash
npm test
npm run test:py
```

## Notebook-based Validation (Windows)

Windows-focused Jupyter notebooks are available in `notebooks/` for Phase 1/2 validation:

1. `notebooks/01_engine_session_smoke.ipynb`
2. `notebooks/02_discovery_and_slide_management.ipynb`
3. `notebooks/03_placeholders_background.ipynb`
4. `notebooks/04_snapshot_and_reports.ipynb`

See `notebooks/README.md` for run guidance and expected outputs.

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

Equivalent make targets:

```bash
make check
make lint-py
make format-check-py
make precommit-run
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

Make shortcuts:

```bash
make com-smoke
make com-smoke-ooxml
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
