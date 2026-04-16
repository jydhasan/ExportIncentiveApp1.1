#define MyAppName "Export Incentive PDF Generator"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "Mahamud Sabuj & Co. Chartered Accountants"
#define MyAppExeName "ExportIncentiveApp.exe"
#define MyAppIcon "app.ico"

[Setup]
AppId={{8F3A2B1C-4D5E-6F7A-8B9C-0D1E2F3A4B5C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL=https://mahamudsabuj.com
DefaultDirName={autopf}\ExportIncentivePDFGenerator
DefaultGroupName={#MyAppName}
AllowNoIcons=no
OutputDir=installer_output
OutputBaseFilename=ExportIncentiveApp_Setup_v1.0
SetupIconFile=app.ico
UninstallDisplayIcon={app}\{#MyAppExeName}
UninstallDisplayName={#MyAppName}
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
WizardImageFile=compiler:WizModernImage.bmp
WizardSmallImageFile=compiler:WizModernSmallImage.bmp
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
MinVersion=10.0.17763
DisableProgramGroupPage=yes
CreateDesktopIcon=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"; Flags: checked

[Files]
; Main application EXE
Source: "publish\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion

; Required PDF native library
Source: "libwkhtmltox.dll"; DestDir: "{app}"; Flags: ignoreversion

; App icon (for uninstaller display)
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; Start Menu shortcut
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"; Comment: "Generate Export Incentive PDF Documents"

; Desktop shortcut
Name: "{userdesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; IconFilename: "{app}\app.ico"; Comment: "Generate Export Incentive PDF Documents"; Tasks: desktopicon

; Uninstall shortcut in Start Menu
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"

[Run]
; Ask user if they want to launch the app after installation
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;
