@echo off
echo ================================================
echo   Running app and capturing errors...
echo ================================================
echo.

cd /d "%~dp0publish"

echo Running: ExportIncentiveApp.exe
ExportIncentiveApp.exe
set EXIT_CODE=%ERRORLEVEL%

echo.
echo Exit code: %EXIT_CODE%

if %EXIT_CODE% NEQ 0 (
    echo.
    echo --- Checking for missing files ---
    if not exist "libwkhtmltox.dll" (
        echo MISSING: libwkhtmltox.dll
    ) else (
        echo OK: libwkhtmltox.dll found
    )
    
    echo.
    echo --- Checking .NET runtime ---
    dotnet --version
    dotnet --list-runtimes
)

echo.
pause
