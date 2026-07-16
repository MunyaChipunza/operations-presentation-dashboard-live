# Operations Presentation Dashboard

Live dashboard bundle for `PPT presentation source data.xlsx`.

## What is here

- `index.html` renders the live dashboard with a table/chart switch for each workbook dataset.
- `scripts/refresh_dashboard_data.py` converts the workbook into `dashboard_data.json`.
- `scripts/publish_dashboard_data.py` refreshes the JSON and pushes it if the folder is inside a Git repo with an `origin`.
- `scripts/register_local_autopublish.ps1` creates a Windows scheduled task that refreshes every minute on this PC.

## Typical local flow

1. Run `python scripts/refresh_dashboard_data.py` to generate `dashboard_data.json`.
2. Preview the folder with a static server or publish it through GitHub Pages.
3. When the folder is in a GitHub-backed repo, run `powershell -ExecutionPolicy Bypass -File scripts/register_local_autopublish.ps1` to keep the JSON updated automatically.

## Local source config

- `scripts/workbook_source.local.json` is a local-only file used by the scheduled task.
- `workbookPath` points at the fallback local workbook on this PC.
- `workbookUrl` can hold a OneDrive or SharePoint share link so the refresh job reads the online workbook directly instead of relying on desktop OneDrive sync.
- Start from `scripts/workbook_source.local.example.json` if you need to recreate the file.

## Optional cloud refresh

The workflow in `.github/workflows/update-dashboard-data.yml` can run manually if you set a `WORKBOOK_URL` repository secret that points to a downloadable workbook.
