# Build (Windows, PowerShell): Create venv, install requirements, build EXE

1)  **Open PowerShell**

2)  **Navigate to the project folder** (e.g., `D:\Audio_Control\`)
    ```powershell
    cd "D:\Audio_Control\WindowsAudioControl-CLI-wGUI-1.4.3.3"
    ```

3)  **Create a virtual environment**
    ```powershell
    py -3 -m venv .\venv
    ```

4)  **Activate the virtual environment**
    ```powershell
    .\venv\Scripts\Activate.ps1
    ```

5)  **Upgrade pip and install requirements**
    ```powershell
    python -m pip install --upgrade pip
    pip install -r requirements.txt
    ```

6)  **Build the Executable with PyInstaller**

    Choose one of the following two methods.

    ---
    ### Method A: Using Command-Line Arguments (Full Command)
    This method uses a single, long command with all options specified. It's useful for one-off builds or scripting without relying on a `.spec` file.

    ```powershell
    pyinstaller -F --noupx --clean --console --name audioctl --bootloader-ignore-signals --collect-all pycaw --collect-all comtypes --hidden-import comtypes.automation --hidden-import comtypes._post_coinit --hidden-import comtypes._post_coinit.unknwn --hidden-import comtypes._post_coinit.misc --icon audio.ico --add-data "audio.ico;." --version-file version.txt .\audioctl.py
    ```
    ---
    ### Method B: Using a `.spec` File (Recommended)
    This is the recommended method. It is simpler to run and less prone to typing errors once the `audioctl.spec` file is set up correctly.

    **1. First, ensure your `audioctl.spec` file contains the following:**
    ```python
    # -*- mode: python ; coding: utf-8 -*-
    from PyInstaller.utils.hooks import collect_all

    # This spec file is configured for a --onefile build.
    
    # Corresponds to: --add-data "audio.ico;."
    datas = [('audio.ico', '.')]
    binaries = []

    # Corresponds to all the --hidden-import flags
    hiddenimports = [
        'comtypes.automation',
        'comtypes._post_coinit',
        'comtypes._post_coinit.unknwn',
        'comtypes._post_coinit.misc'
    ]

    # Corresponds to: --collect-all pycaw
    tmp_ret = collect_all('pycaw')
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]

    # Corresponds to: --collect-all comtypes
    tmp_ret = collect_all('comtypes')
    datas += tmp_ret[0]
    binaries += tmp_ret[1]
    hiddenimports += tmp_ret[2]

    a = Analysis(
        ['audioctl.py'], # The main script
        pathex=[],
        binaries=binaries,
        datas=datas,
        hiddenimports=hiddenimports,
        hookspath=[],
        hooksconfig={},
        runtime_hooks=[],
        excludes=[],
        noarchive=False,
        optimize=0,
    )
    pyz = PYZ(a.pure)

    exe = EXE(
        pyz,
        a.scripts,
        a.binaries,
        a.datas,
        [],
        # Corresponds to: --name audioctl
        name='audioctl',
        debug=False,
        # Corresponds to: --bootloader-ignore-signals
        bootloader_ignore_signals=True,
        strip=False,
        # Corresponds to: --noupx
        upx=False,
        upx_exclude=[],
        runtime_tmpdir=None,
        # Corresponds to: --console
        console=True,
        disable_windowed_traceback=False,
        argv_emulation=False,
        target_arch=None,
        codesign_identity=None,
        entitlements_file=None,
        # Corresponds to: --version-file version.txt
        version='version.txt',
        # Corresponds to: --icon audio.ico
        icon=['audio.ico'],
    )
    ```

    **2. Then, run this simple command:**
    ```powershell
    pyinstaller audioctl.spec --clean
    ```
---

### Notes:
-   Ensure `requirements.txt`, `audio.ico`, and `version.txt` are present in the project folder.
-   After activation, your prompt will show `(venv)`. Run all build commands while it's active.
-   The `--collect-all comtypes` and additional hidden imports are critical to ensure all `comtypes` modules are included, which fixes COM cleanup crashes.
-   The `--bootloader-ignore-signals` flag is added to slightly alter the executable's wrapper, which can help avoid false positive detections by antivirus software.
