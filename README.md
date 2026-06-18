This is an app to check and filter out an album excel that Ive been doing for a while. Its got 2 different front-end options:
    -Terminal based
    -GUI based
    
To share or run:
    - Copy the readerUI binary from the dist/ folder.
    - Put it in the same directory as your Albums List.xlsx spreadsheet.
    - You can now run the executable by double-clicking it or calling ./dist/readerUI (no python/dependencies needed). You can send this binary to any other Linux user along with the spreadsheet.
(Note: If you want to make a Windows version, copy your folder onto a Windows machine, install pyinstaller, and run the same command: python -m PyInstaller --noconsole --onefile --collect-all customtkinter readerUI.py)

