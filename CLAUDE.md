# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project overview

Desktop invoicing application (Windows only) built with Python + CustomTkinter. It generates Factur-X PDF invoices from Excel templates, embedding EN16931-compliant UBL XML.

The app is **open-core**: the public repository runs fully **local** (invoice numbers stored on disk, no auth). API-backed numbering + OAuth2 authentication live in a **private git submodule** at `controller/backends/api/` (see "Numbering backends" below). When that submodule is absent (public clone), the app automatically falls back to local mode.

**Companion API project** (needed only for the API backend): `\\wsl.localhost\Ubuntu\home\olivier\ez_facture_api_linux` — runs under Docker in WSL/Ubuntu.

## Commands

### Run the application
```
python main.py
```
Requires Windows with Excel installed (uses xlwings COM automation). Must be run from the project root.

### Run tests
```
# All tests (en16931 module only)
pytest en16931/tests/

# Single test file
pytest en16931/tests/test_invoice.py

# Single test
pytest en16931/tests/test_invoice.py::test_name
```

Tests are scoped to the `en16931/` module. There are no tests for the controller/view layer (requires Excel COM).

### Compile to executable
```
pyinstaller main.spec
```
After build, copy `templates/` and `images/` folders next to the generated executable.

### Validate an XML invoice against EN16931 XSLT rules
```
python validate/validate.py
```
Edit `xml_path` and `xslt_path` variables in that file before running.

### Launch the local API (in WSL)
```
docker-compose -f docker-compose.dev.yml up -d --build
```
Run from `\\wsl.localhost\Ubuntu\home\olivier\ez_facture_api_linux`.

## Configuration

Public switches are in [config.py](config.py):

- `MODE_API` — `None` (default): auto-select the backend by the presence of the private API submodule; `True`: force API; `False`: force local. Can be overridden by the `EZFACTURE_MODE=local|api` environment variable.

API-only config (URLs, OAuth `CLIENT_ID`, `PROD` flag) lives in the private submodule at `controller/backends/api/config.py`, not in the public `config.py`.

Data files loaded at startup (must be present in the project root):
- `clients.xlsx` — customer list
- `produits.xlsx` — product catalog
- `config.xlsx` — seller configuration (SIRET, TVA, address, etc.)

## Architecture

### MVC structure

```
main.py          → Entry point: creates AppView and Controller
view/view.py     → AppView (CustomTkinter UI)
controller/      → Business logic
models/          → Domain models (Facture, Devis extend en16931.Invoice)
en16931/         → EN16931/UBL XML library (self-contained, testable)
tools/utils.py   → Utilities (resource path, local invoice file integrity)
```

### controller/ internals

| File | Role |
|---|---|
| `controller.py` | `Controller` class — wires view actions to business logic, reads Excel cells, orchestrates validation flow. Talks to numbering/auth only through `self.backend`. |
| `ezfacture.py` | `Ezfacture` singleton — wraps a single Excel workbook (xlwings); enforces one-document-at-a-time via `__new__` |
| `eztransaction.py` | `Eztransaction` context manager — simple two-phase commit: `step_fn` runs immediately, `rollback_fn` runs on exception, `commit_fn` runs on success |
| `constantes.py` | Named cell lists for template validation, draft path, print area range |
| `backends/` | Numbering/auth backends (see below) |

### Numbering backends (`controller/backends/`)

The `Controller` never calls the API or the local file directly — it uses a `NumberingBackend` chosen at startup by `get_numbering_backend(controller)`:

| File | Visibility | Role |
|---|---|---|
| `backends/base.py` | public | `NumberingBackend` ABC: `login`, `integrity_ok`, `get_number`, `reserve`, `commit`, `cancel` |
| `backends/local.py` | public | `LocalBackend` — numbers stored in `invoices.jsonl`; `login()`/`cancel()` are no-ops |
| `backends/__init__.py` | public | `get_numbering_backend()` factory + automatic fallback to local |
| `backends/api/` | **private submodule** | `ApiBackend` (`api_backend.py`), OAuth2 PKCE flow (`ezauth.py`, local Werkzeug callback on port 8080, token in `.token_cache`), and API config (`config.py`) |

Backend selection order: `EZFACTURE_MODE` env var → `MODE_API` in config.py → auto (API submodule present ⇒ `ApiBackend`, else `LocalBackend`). The factory catches `ImportError` on `controller.backends.api`, so a public clone without the submodule runs local with zero config.

The backend holds a reference to the `Controller` (for `doc.onglet_config`, `view`, `model`).

### Invoice validation flow (Controller.validate)

The critical path when a user clicks "Valider" uses `Eztransaction` to ensure atomicity (all numbering steps go through `self.backend`):
1. Get invoice number — `backend.get_number()` (API GET or local read)
2. Generate XML (`model.save()`) → rollback: `model.unsave()`
3. Validate XML against XSD (`facturx.xml_check_xsd`)
4. Generate PDF from Excel sheet → rollback: delete PDF file
5. Embed XML into PDF via `factur-x` library
6. Reserve invoice number — `backend.reserve()` (API POST / local PREPARE) → rollback: `backend.cancel()` (API DELETE / local no-op)
7. Commit — `backend.commit()` (local: write to `invoices.jsonl`; API: no-op)

### en16931/ module

An embedded Python library implementing the EN16931 European invoice standard (UBL 2.1 / PEPPOL BIS 3). Key classes:
- `Invoice` — main object; serializes to XML via Jinja2 template (`en16931/templates/invoice.xml`)
- `Entity` — seller/buyer party
- `InvoiceLine` — individual line items with tax
- `Tax` — tax category/percent

`models/facture.py` subclasses `Invoice` to add `invoice_type_code` (380=facture, 381=avoir), `invoice_reference` (for avoirs), and an in-memory `xml` property instead of writing to disk.

### Invoice numbering — local mode (`LocalBackend`)

In local mode, invoice numbers are persisted in `invoices.jsonl` (append-only JSONL). Each entry is protected by two hashes:
- `file_hash_before`: SHA256 of the file content before this entry (chaining)
- `self_hash`: SHA256 of `number + type + timestamp + file_hash_before`

`tools/utils.verify_local_file()` re-validates the full chain, invoked via `LocalBackend.integrity_ok()` after login. If tampered, the UI shows an error and blocks document creation.

### Excel template conventions

Named cells (defined in `controller/constantes.py`) drive data extraction. All reads go through `Controller.get_value()` which checks cell existence, emptiness, and optionally format. Cell names are scoped: `facture!date_facture`, `devis!dev_num_devis`, etc.

Draft files are prefixed: `DRAFT_FAC_`, `DRAFT_AV_`, `DEV_` + Unix timestamp. Stored in `brouillons/`. Validated documents become PDFs in `pdf/`.

### `Ezfacture` singleton constraint

Only one Excel document can be open at a time. `Ezfacture._instance` enforces this — always clean up with `type(self.doc)._instance = None` and `gc.collect()` before creating a new instance.

### `check_ui` / `check_opened_doc` decorators

- `@check_ui` — guards any operation on an open document; if Excel was closed externally, blocks the UI and shows a restart prompt
- `@check_opened_doc` — used on `create_doc` / `ouvrir`; closes any existing document before opening a new one
