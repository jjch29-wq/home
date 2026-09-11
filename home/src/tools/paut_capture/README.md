# PAUT Capture

Standalone Windows GUI utility that opens OmniPC inspection files and captures
the configured screen area automatically.

## Layout

- `src/paut_capture.py`: application implementation.
- `src/auto_capture_config.json`: current capture settings.

The legacy `home/src/paut capture.py` entry point remains as a compatibility
launcher for existing shortcuts and IDE run targets.

## Run

From the repository root:

```powershell
.venv\Scripts\python.exe "home\src\paut capture.py"
```
