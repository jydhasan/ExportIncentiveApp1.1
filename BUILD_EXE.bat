@echo off
echo ================================================
echo   Export Incentive App - Build EXE
echo ================================================
echo.

cd /d "%~dp0"

echo [1/4] Cleaning old build...
if exist ".\publish" rmdir /s /q ".\publish"

echo.
echo [2/4] Building EXE (folder mode)...
dotnet publish -c Release -r win-x64 --self-contained true -o ".\publish"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo ERROR: Build failed! See errors above.
    pause
    exit /b 1
)

echo.
echo [3/4] Copying libwkhtmltox.dll to publish folder...
if exist "libwkhtmltox.dll" (
    copy /Y "libwkhtmltox.dll" ".\publish\libwkhtmltox.dll" >nul
    echo     OK: libwkhtmltox.dll copied.
) else (
    echo.
    echo     WARNING: libwkhtmltox.dll NOT FOUND in project folder!
    echo     Please download it from:
    echo     https://github.com/rdvojmoc/DinkToPdf/raw/master/v0.12.4/64%%20bit/libwkhtmltox.dll
    echo     and place it in: %~dp0
    echo.
)

echo.
echo [4/4] Creating Desktop Shortcut...
set EXE_PATH=%~dp0publish\ExportIncentiveApp.exe
set ICON_PATH=%~dp0app.ico
set SHORTCUT=%USERPROFILE%\Desktop\Export Incentive App.lnk

powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "$s = (New-Object -COM WScript.Shell).CreateShortcut('%SHORTCUT%');" ^
  "$s.TargetPath = '%EXE_PATH%';" ^
  "$s.IconLocation = '%ICON_PATH%';" ^
  "$s.WorkingDirectory = '%~dp0publish';" ^
  "$s.Description = 'Export Incentive PDF Generator - Mahamud Sabuj & Co.';" ^
  "$s.Save()"

echo.
echo ================================================
echo   BUILD SUCCESSFUL!
echo ================================================
echo.
echo   EXE:     %~dp0publish\ExportIncentiveApp.exe
echo   Desktop: "Export Incentive App" shortcut created
echo.
echo   IMPORTANT: Do NOT move the 'publish' folder!
echo   The desktop shortcut points to this location.
echo ================================================
echo.

set /p OPEN="Open publish folder now? (Y/N): "
if /i "%OPEN%"=="Y" explorer "%~dp0publish"

pause
