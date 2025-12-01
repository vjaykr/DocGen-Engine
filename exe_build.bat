@echo off
echo Building executable without Qt conflicts...
python -m PyInstaller --onefile --noconsole --exclude-module PyQt5 --exclude-module PyQt6 --exclude-module torch --exclude-module tensorflow --exclude-module matplotlib --name "DocGen Engine" --hidden-import win32com.client --add-data "app.py;." launcher.py

echo Listing build results:
if exist "dist" (
    echo Contents of dist folder:
    dir dist
    if exist "dist\DocGen Engine.exe" (
        echo SUCCESS: DocGen Engine.exe created successfully!
    ) else (
        echo ERROR: DocGen Engine.exe not found in dist folder
    )
) else (
    echo ERROR: dist folder not created!
)

pause