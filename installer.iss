#define MyAppName "Edge PWA Redirector"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "micahmo"
#define MyAppExeName "EdgePwaRedirector.exe"

[Setup]
AppId={{A3F2B1C4-8D6E-4F3A-9B2C-1D5E7F8A0B3C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\EdgePwaRedirector
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=EdgePwaRedirector-Setup
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
Source: "app.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: ""

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue; Tasks: startup

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent runasoriginaluser

[Code]
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ResultCode, I, F: Integer;
  ExePath: String;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM {#MyAppExeName}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  // Single-file .NET apps memory-map the exe; handles release asynchronously after
  // TerminateProcess. Poll until we can open the file for writing (up to 10s).
  ExePath := ExpandConstant('{app}\{#MyAppExeName}');
  for I := 1 to 50 do
  begin
    Sleep(200);
    F := FileOpen(ExePath, fmOpenReadWrite);
    if F >= 0 then
    begin
      FileClose(F);
      break;
    end;
  end;

  Result := '';
end;
