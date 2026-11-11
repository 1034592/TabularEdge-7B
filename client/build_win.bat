@echo off
rem Build TabularEdge.exe (PyInstaller onedir, windowed). Run on the Windows host.
set SRC=%~dp0..
set B=%USERPROFILE%\te-build
cd /d %SRC%
"%B%\venv\Scripts\pyinstaller.exe" --noconfirm --clean --windowed --name TabularEdge --icon "%SRC%\client\tabularedge.ico" ^
  --add-data "%SRC%\schema;schema" --add-data "%SRC%\client\tabularedge.ico;." --paths "%SRC%" ^
  --exclude-module numpy --exclude-module PIL --exclude-module matplotlib --exclude-module torch ^
  --distpath "%B%\dist" --workpath "%B%\work" --specpath "%B%" "%SRC%\client\app.py"
echo BUILD_EXIT %ERRORLEVEL%
