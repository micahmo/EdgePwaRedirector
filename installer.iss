#define MyAppName "Edge PWA Redirector"
#define _ExeVer GetVersionNumbersString("publish\EdgePwaRedirector.exe")
#define MyAppVersion Copy(_ExeVer, 1, Len(_ExeVer) - 2)
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

[InstallDelete]
Type: files; Name: "{app}\{#MyAppExeName}.bak"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: ""

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Windows\CurrentVersion\Run"; ValueType: string; ValueName: "{#MyAppName}"; ValueData: """{app}\{#MyAppExeName}"""; Flags: uninsdeletevalue; Tasks: startup

[Code]
function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  ExePath: String;
begin
  // Rename the old exe out of the way so Inno creates a fresh file rather than
  // overwriting. RenameFile works even on running exes (FILE_SHARE_DELETE), so
  // no process kill is needed here. The new instance kills the old one on startup.
  ExePath := ExpandConstant('{app}\{#MyAppExeName}');
  DeleteFile(ExePath + '.bak');
  RenameFile(ExePath, ExePath + '.bak');
  Result := '';
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ShellApp: Variant;
  ExePath: String;
begin
  if CurStep <> ssDone then Exit;
  // Launch the app after install. On domain machines where the installer runs
  // as a different elevated account, the app detects it wasn't launched from
  // Explorer and automatically relaunches itself via Task Scheduler to get a
  // proper interactive session context for the notification area.
  ExePath := ExpandConstant('{app}\{#MyAppExeName}');
  try
    ShellApp := CreateOleObject('Shell.Application');
    ShellApp.ShellExecute(ExePath, '', ExpandConstant('{app}'), 'open', 1);
  except
  end;
end;
