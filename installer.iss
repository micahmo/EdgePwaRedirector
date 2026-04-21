#define MyAppName "Teams Link Redirector"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "micahmo"
#define MyAppExeName "TeamsLinkRedirector.exe"

[Setup]
AppId={{A3F2B1C4-8D6E-4F3A-9B2C-1D5E7F8A0B3C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\TeamsLinkRedirector
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=TeamsLinkRedirector-Setup
SetupIconFile=app.ico
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "startup"; Description: "Start automatically when Windows starts"; GroupDescription: "Additional options:"

[Files]
Source: "publish\{#MyAppExeName}"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\app.ico"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\System.Management.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\TeamsLinkRedirector.deps.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\TeamsLinkRedirector.runtimeconfig.json"; DestDir: "{app}"; Flags: ignoreversion
Source: "publish\runtimes\*"; DestDir: "{app}\runtimes"; Flags: ignoreversion recursesubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: ""

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue; Tasks: startup

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent
