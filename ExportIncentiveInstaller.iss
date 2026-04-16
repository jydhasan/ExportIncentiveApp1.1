; ===============================
; Export Incentive Installer
; ===============================

#define MyAppName "Export Incentive PDF Generator"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Mahamud Sabuj & Co. Chartered Accountants"
#define MyAppExeName "ExportIncentiveApp.exe"

[Setup]
AppId={{A1B2C3D4-E5F6-47A8-9B0C-123456789ABC}}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\ExportIncentivePDFGenerator
DefaultGroupName={#MyAppName}
OutputDir=installer_output
OutputBaseFilename=ExportIncentiveApp_Setup_v1.0
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

; ===============================
; Files Section
; ===============================

[Files]
; Copy ALL published files
Source: "publish\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; ===============================
; Shortcuts
; ===============================

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"

; ===============================
; Run After Install
; ===============================

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent