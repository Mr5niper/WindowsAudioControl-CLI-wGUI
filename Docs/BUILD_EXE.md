# Build (Windows, PowerShell): Create venv, install requirements, build EXE

1) Open PowerShell

2) Navigate to the project folder (example: Audio_Control)
```powershell
cd "C:\path\to\Audio_Control"
```

3) Create a virtual environment
```powershell
py -3 -m venv .\venv
```

4) Activate the virtual environment
```powershell
.\venv\Scripts\Activate.ps1
```

5) Upgrade pip and install requirements
```powershell
python -m pip install --upgrade pip
pip install -r requirements.txt
```

6) Build the executable with PyInstaller
```powershell
pyinstaller --clean --onefile --console --name audioctl --collect-all pycaw --collect-submodules comtypes --hidden-import comtypes --hidden-import comtypes.gen --hidden-import comtypes.automation --icon audio.ico --add-data "audio.ico;." --version-file version.txt audioctl_v1.2.1.py
```

Notes:
- Ensure `requirements.txt`, `audio.ico`, and `version.txt` are present in the project folder.
- After activation, your prompt will show `(venv)`. Run all build commands while it’s active.

