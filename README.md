
# Uniware Excel Verifier (MVP)

A local-first web application for pre-upload spreadsheet quality checks. It supports Excel/CSV upload, worksheet selection, preview, configurable mandatory columns, configurable unique columns, row-level issues, and Excel/CSV verification reports.

## Important scope statement

This MVP performs **generic data-quality validation plus only the rules you configure**. It does not include an official Uniware API integration, does not know your account-specific import templates, and cannot guarantee that Uniware will accept a file. Direct upload and AI analysis are not implemented in this MVP; the core validator works without AI credentials.

The application is designed as a single-user local prototype. It is **not safe to expose publicly as-is**: authentication, multi-user authorization, durable session storage, CSRF/abuse controls, and a production retention worker must be added before deployment for multiple users.

## Requirements

- Python 3.10 or newer
- Internet access for installing packages (not required at runtime after installation)
- For `.xls` files, `xlrd` is installed through `requirements.txt`

## Run locally

### Windows PowerShell

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
uvicorn app.main:app --reload
```

### macOS / Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env
uvicorn app.main:app --reload
```

Open http://127.0.0.1:8000.

## Use

1. Upload `.xlsx`, `.xls`, or `.csv` (default limit: 20 MB).
2. Select the worksheet if needed.
3. Optionally enter comma-separated required columns and columns that must be unique. These must reflect the actual operation/template; do not guess.
4. Run verification.
5. Filter issues and download an Excel or CSV report.
6. Download a data copy. For Excel, the MVP exports the selected sheet to a new workbook; it does not preserve original formatting or other sheets.
7. Delete the session when finished.

## Validation behavior

- Required column headers and required cell values: errors, but only when configured.
- Duplicate headers and blank headers: errors.
- Exact duplicate data rows: warnings.
- Repeated values in configured unique columns: errors.
- Numeric-looking headers (such as quantity, qty, price, amount, weight, stock): suspicious nonnumeric values are warnings, not hard failures.
- Missing configured unique columns: warning that the check was skipped.

A `PASS` means only that no configured checks reported an issue. It is not a guarantee of Uniware acceptance. If you have not supplied the official template/rules, treat the result as generic preflight checking, not Uniware-specific verification.

## Tests

```bash
pytest -q
```

The test suite covers basic validator cases. Add integration tests for your real templates and rules before relying on the application operationally.

## Configuration

Copy `.env.example` to `.env`. `MAX_UPLOAD_MB` controls the upload size limit. `DATA_DIR` can optionally point to a dedicated temporary-data directory. Restart the server after changing environment variables.

## Data handling and security notes

- Uploaded bytes are stored under `data/` until the user deletes the session or the application data is removed.
- The current session registry is in memory. Restarting the server clears sessions, but files in `data/` may remain; manually remove them if needed.
- Do not upload production/customer data into an exposed instance.
- Do not put credentials in frontend code or commit `.env`.
- This MVP does not call an AI provider or Uniware.
- Before public or multi-user deployment, implement authentication, authorization, durable isolated storage, retention cleanup, security review, HTTPS, request throttling, and audit controls.

## Planned integration path

To implement real Uniware upload, provide the official documentation, API endpoint specifications, supported payload/file format, authentication method, and sandbox credentials. To implement Uniware-specific validation, provide the official import template and business rules. Never infer endpoints or treat local validation as an API acceptance response.
