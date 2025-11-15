# Helper scripts

This folder contains small helper scripts used by the project. Each script is intended
to be simple and environment-safe; they either create a project-local `.venv` or
perform common developer tasks.

## setup-venv.ps1

- Location: `scripts/setup-venv.ps1`
- Purpose: Create a Windows PowerShell-friendly virtual environment at the project root
  (path `.venv`), upgrade pip, and install `requirements.txt`.
- Usage:

```powershell
# create the venv if it does not exist
.\scripts\setup-venv.ps1

# force recreate (backup existing .venv)
.\scripts\setup-venv.ps1 -Force
```

- Notes:
  - The script attempts to locate the newest Python available (looks in common
    Windows locations and PATH). If you need a specific interpreter, create the
    venv manually with `C:\path\to\python.exe -m venv .venv`.
  - Activation:

```powershell
.\.venv\Scripts\Activate.ps1
# then run: python cal2planner.py
```

## CI

- A minimal GitHub Actions workflow is included at `.github/workflows/setup-venv.yml`.
  It creates a venv on `windows-latest` and `ubuntu-latest`, restores a cached pip
  download cache, and installs requirements. It's a convenience to verify that
  dependencies install cleanly on supported platforms.
