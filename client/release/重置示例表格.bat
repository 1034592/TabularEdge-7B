@echo off
rem Restore the sample workbooks and remove generated results / backups / split folders.
cd /d "%~dp0test_cases"
del /q *.xlsx 2>nul
for /d %%d in (*) do if /i not "%%d"=="_original" rmdir /s /q "%%d"
copy /Y "_original\*.xlsx" . >nul
echo Sample workbooks restored.
pause
