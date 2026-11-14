; Book Work Clock per-user installer built from a PyInstaller executable.
#define MyAppName "Book Work Clock"
#define MyAppVersion "0.3.4"
#define MyAppExeName "BookWorkClock.exe"

[Setup]
AppId={{BE2C890A-4E75-4EDB-911B-881DE34D05F4}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
DefaultDirName={localappdata}\Programs\BookWorkClock
DefaultGroupName={#MyAppName}
OutputDir=..\dist\installer
OutputBaseFilename=BookWorkClock-Setup
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
CloseApplications=yes

[Files]
Source: "..\dist\BookWorkClock.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\Book Work Clock"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\Book Work Clock"; Filename: "{app}\{#MyAppExeName}"

[Tasks]
Name: "autostart"; Description: "Start Book Work Clock in the background when Windows starts"; GroupDescription: "Startup:"; Flags: checkedonce

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "BookWorkClock"; ValueData: """{app}\{#MyAppExeName}"" --background"; Tasks: autostart; Flags: uninsdeletevalue

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch Book Work Clock"; Flags: nowait postinstall skipifsilent
