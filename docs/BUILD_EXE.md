# **Auto Build Process (Recommended)**
This method uses the included batch script to automate environment creation, dependency installation, and compilation.

**Prerequisite:** Install Python 3.14.3  
Download: https://www.python.org/downloads/release/python-3143/  
*(Ensure you check "Add Python to environment variables" during installation or manually add them before build)*

1.  Navigate to the project folder.
2.  Double-click **`BUILD_EXE.bat`**.
    *   The script will verify your Python version matches `3.14.3`.
    *   If the version is incorrect, it will pause and provide the download link.
3.  Wait for the process to complete.
    *   It will automatically create a temporary virtual environment (`venv`), install `requirements.txt`, and run PyInstaller.
4.  Locate your new executable in the **`dist`** folder.

# Manual Build Process (Windows, PowerShell): Create venv, install requirements, build EXE
  Prerequisite:
  Install Python 3.14.3 from:
  <BR>
  https://www.python.org/downloads/release/python-3143/
  <BR>
  *(Ensure you check "Add Python to environment variables" during installation or manually add them before build)*
  <BR>
1) Open PowerShell
2) Navigate to the project folder,
   <BR>*example:* `cd "C:\path\to\Audio_Control"`

4) Create a virtual environment
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
```
```powershell
pip install -r requirements.txt
```
6) Build the executable with PyInstaller
```powershell
pyinstaller -F --noupx --clean --console --name audioctl --collect-all pycaw --collect-all comtypes --hidden-import comtypes.automation --hidden-import comtypes._post_coinit --hidden-import comtypes._post_coinit.unknwn --hidden-import comtypes._post_coinit.misc --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl.py
```
Notes:
- Ensure `requirements.txt`, `audio.ico`, and `version.txt` are present in the project folder.
- After activation, your prompt will show `(venv)`. Run all build commands while it's active.
- The `--collect-all comtypes` ensures all comtypes modules are included (fixes COM cleanup crashes).
- The additional hidden imports ensure critical comtypes submodules are bundled.
