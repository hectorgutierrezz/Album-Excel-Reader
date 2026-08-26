This is an app to check and filter an album spreadsheet. It has two front-end options:

- Terminal-based
- GUI-based

## Windows

Build the portable Windows executable on a Windows computer. In PowerShell, from the project folder, run:

```powershell
py -m venv .venv
.venv\Scripts\activate
py -m pip install pandas openpyxl customtkinter pyinstaller
py -m PyInstaller --noconsole --onefile --clean --collect-all customtkinter readerUI.py
```

The executable will be created at `dist\readerUI.exe`. To share it, place these files in the same folder:

```text
readerUI.exe
Albums List.xlsx
```

The user can then double-click `readerUI.exe`. Python and the dependencies are not required on the target computer.

## Linux

Build the portable Linux executable on Linux. From the project folder, run:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install pandas openpyxl customtkinter pyinstaller
python -m PyInstaller --noconsole --onefile --clean --collect-all customtkinter readerUI.py
```

The executable will be created at `dist/readerUI`. To share it, place these files in the same folder:

```text
readerUI
Albums List.xlsx
```

Run it with:

```bash
chmod +x readerUI
./readerUI
```

Python and the dependencies are not required on the target computer. The Linux executable is built for 64-bit x86 Linux systems, and must be built separately from the Windows executable.

