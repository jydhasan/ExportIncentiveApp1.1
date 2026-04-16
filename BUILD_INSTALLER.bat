@echo off
chcp 65001 >nul
title Export Incentive App - Build Installer
color 0A

echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║     Export Incentive PDF Generator - Installer       ║
echo  ║              Build Script v1.0                       ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

cd /d "%~dp0"

:: ── Step 1: Check dotnet ─────────────────────────────────────
echo  [1/4] Checking .NET SDK...
dotnet --version >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo  ERROR: .NET SDK not found!
    echo  Please install from: https://dotnet.microsoft.com/download
    pause & exit /b 1
)
echo        OK - .NET found.
echo.

:: ── Step 2: Check libwkhtmltox.dll ───────────────────────────
echo  [2/4] Checking libwkhtmltox.dll...
if not exist "libwkhtmltox.dll" (
    echo.
    echo  WARNING: libwkhtmltox.dll not found!
    echo  Download it from:
    echo  https://github.com/rdvojmoc/DinkToPdf/raw/master/v0.12.4/64%%20bit/libwkhtmltox.dll
    echo  Place it in this folder, then run this script again.
    echo.
    pause & exit /b 1
)
echo        OK - libwkhtmltox.dll found.
echo.

:: ── Step 3: Publish single EXE ───────────────────────────────
echo  [3/4] Building self-contained EXE (this takes ~30 seconds)...
if exist "publish" rmdir /s /q "publish"

dotnet publish -c Release -r win-x64 --self-contained true ^
  -p:PublishSingleFile=true ^
  -p:IncludeNativeLibrariesForSelfExtract=true ^
  -p:EnableCompressionInSingleFile=true ^
  -p:DebugType=none ^
  -p:DebugSymbols=false ^
  -o "publish" >build_log.txt 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo  ERROR: Build failed! See build_log.txt for details.
    type build_log.txt
    pause & exit /b 1
)
echo        OK - EXE created at publish\ExportIncentiveApp.exe
echo.

:: ── Step 4: Create Installer with Inno Setup ─────────────────
echo  [4/4] Creating Installer...

:: Check common Inno Setup install locations
set ISCC=""
if exist "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files\Inno Setup 6\ISCC.exe"       set ISCC="C:\Program Files\Inno Setup 6\ISCC.exe"
if exist "C:\Program Files (x86)\Inno Setup 5\ISCC.exe" set ISCC="C:\Program Files (x86)\Inno Setup 5\ISCC.exe"

if %ISCC%=="" (
    echo.
    echo  Inno Setup not found!
    echo.
    echo  Please install Inno Setup from:
    echo  https://jrsoftware.org/isdl.php
    echo  (Free download, takes 2 minutes)
    echo.
    echo  After installing, run this script again.
    echo.
    echo  NOTE: Your EXE is already built at:
    echo        %~dp0publish\ExportIncentiveApp.exe
    pause & exit /b 1
)

if not exist "installer_output" mkdir "installer_output"
%ISCC% "installer.iss" >installer_log.txt 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo  ERROR: Installer creation failed! See installer_log.txt
    type installer_log.txt
    pause & exit /b 1
)

echo        OK - Installer created!
echo.
echo.
echo  ╔══════════════════════════════════════════════════════╗
echo  ║                 BUILD COMPLETE!                      ║
echo  ╠══════════════════════════════════════════════════════╣
echo  ║                                                      ║
echo  ║  Installer:                                          ║
echo  ║  installer_output\ExportIncentiveApp_Setup_v1.0.exe  ║
echo  ║                                                      ║
echo  ║  Share this single file with anyone.                 ║
echo  ║  They just double-click to install!                  ║
echo  ║                                                      ║
echo  ╚══════════════════════════════════════════════════════╝
echo.

:: Open the output folder
explorer "installer_output"
pause
