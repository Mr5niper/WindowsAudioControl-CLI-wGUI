@echo off
setlocal

:: ==========================================================================
:: Build Script for audioctl
:: ==========================================================================
:: This script creates a virtual environment, installs dependencies,
:: and builds a single-file executable using PyInstaller.
:: ==========================================================================

echo [INFO] Starting build process...

:: 1. Create Virtual Environment
echo [STEP 1/4] Creating virtual environment in '.\venv'...
if not exist .\venv (
    py -3 -m venv .\venv
    if errorlevel 1 (
        echo [ERROR] Failed to create virtual environment.
        goto :error
    )
) else (
    echo [INFO] Virtual environment '.\venv' already exists. Skipping creation.
)

:: 2. Activate Virtual Environment
echo [STEP 2/4] Activating virtual environment...
call .\venv\Scripts\activate.bat
if not defined VIRTUAL_ENV (
    echo [ERROR] Failed to activate the virtual environment. Make sure '.\venv\Scripts\activate.bat' exists.
    goto :error
)

:: 3. Install Dependencies
echo [STEP 3/4] Upgrading pip and installing dependencies from requirements.txt...
python -m pip install --upgrade pip > nul
if errorlevel 1 (
    echo [ERROR] Failed to upgrade pip.
    goto :error
)
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies from requirements.txt.
    goto :error
)

:: 4. Build with PyInstaller
echo [STEP 4/4] Building executable with PyInstaller...
pyinstaller -F --noupx --clean --console --name audioctl --collect-all pycaw --collect-all comtypes --hidden-import comtypes.automation --hidden-import comtypes._post_coinit --hidden-import comtypes._post_coinit.unknwn --hidden-import comtypes._post_coinit.misc --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl.py
if errorlevel 1 (
    echo [ERROR] PyInstaller build failed.
    goto :error
)

echo.
echo [SUCCESS] Build completed successfully.
echo The executable can be found in the '.\dist' directory.
goto :end

:error
echo.
echo [FAILURE] The build process failed. Please check the errors above.

:end
echo.
pause
endlocal
