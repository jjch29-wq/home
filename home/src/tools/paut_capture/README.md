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
.venv\Scripts\python.exe -m pip install -r "home\src\tools\paut_capture\requirements.txt"
.venv\Scripts\python.exe "home\src\paut capture.py"
```

## Capture profiles

The app keeps separate click coordinates and capture regions for three layouts:

- `OPD Single`: an `.opd` filename whose stem ends in the standalone number
  `90` or `270` (for example, `WELD_90.opd` or `WELD-270.opd`). Click step 1
  is skipped because the file already uses the Single layout.
- `OPD Dual`: every other `.opd` file.
- `NDE`: every `.nde` file.

Choose each layout from **화면 설정 유형**, calibrate its click positions and
capture region, and close the app normally to save it. During a batch run the
matching profile is selected automatically for every file.
