# Statement Generator Desktop

Python Windows desktop app for generating Nepali bank statements and exporting:

- Statement to Excel using your real template files
- Balance certificate to Word using your real template files
- Issue date as the next valid business day
- USD conversion from Nepal Rastra Bank rate or manual override

## Run

From this workspace root:

```powershell
python -m statement_generator.app
```

Or double-click:

- `statement_generator_app.pyw`
- `run_statement_generator.bat`

## Notes

- The app scans the template folder you choose, so you can add more formats later and click `Refresh`.
- Excel/Word export uses Microsoft Office on Windows to preserve layout as closely as possible.
- If Office is not installed or blocked, preview and generation still work, but export will fail.
