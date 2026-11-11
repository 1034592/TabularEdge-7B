@echo off
rem Assemble TabularEdge-Release on the Desktop. Weights are NOT bundled; config.json points to models\ (see README).
set SRC=%~dp0..
set B=%USERPROFILE%\te-build
set Z=%USERPROFILE%\tools\llama-b12119
set R=%USERPROFILE%\Desktop\TabularEdge-Release
if not exist "%R%" mkdir "%R%"
robocopy "%B%\dist\TabularEdge" "%R%" /E /NFL /NDL /NJH /NJS /NP >nul
for %%E in (cpu vulkan cuda) do if not exist "%R%\engine\%%E" mkdir "%R%\engine\%%E"
if not exist "%R%\engine\cpu\llama-server.exe" tar -xf "%Z%\llama-b12119-bin-win-cpu-x64.zip" -C "%R%\engine\cpu"
if not exist "%R%\engine\vulkan\llama-server.exe" tar -xf "%Z%\llama-b12119-bin-win-vulkan-x64.zip" -C "%R%\engine\vulkan"
if not exist "%R%\engine\cuda\llama-server.exe" (
  tar -xf "%Z%\llama-b12119-bin-win-cuda-12.4-x64.zip" -C "%R%\engine\cuda"
  tar -xf "%Z%\cudart-llama-bin-win-cuda-12.4-x64.zip" -C "%R%\engine\cuda"
)
if not exist "%R%\test_cases" mkdir "%R%\test_cases"
copy /Y "%SRC%\test_cases\*.xlsx" "%R%\test_cases\" >nul
copy /Y "%SRC%\test_cases\scenarios.json" "%R%\test_cases\" >nul
if not exist "%R%\test_cases\_original" mkdir "%R%\test_cases\_original"
copy /Y "%SRC%\test_cases\*.xlsx" "%R%\test_cases\_original\" >nul
copy /Y "%SRC%\client\release\*.*" "%R%\" >nul
echo ASSEMBLE_EXIT %ERRORLEVEL%
