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
  ResultCode, I: Integer;
  ExePath, Log: String;
begin
  Exec(ExpandConstant('{sys}\taskkill.exe'), '/F /IM {#MyAppExeName}', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
  Log := 'taskkill exit code: ' + IntToStr(ResultCode) + #13#10;

  ExePath := ExpandConstant('{app}\{#MyAppExeName}');
  Log := Log + 'watching: ' + ExePath + #13#10;

  for I := 1 to 50 do
  begin
    Sleep(200);
    if RenameFile(ExePath, ExePath + '.old') then
    begin
      RenameFile(ExePath + '.old', ExePath);
      Log := Log + 'file free after ' + IntToStr(I) + ' attempt(s)' + #13#10;
      break;
    end;
    if I = 50 then
      Log := Log + 'file still locked after 50 attempts' + #13#10;
  end;

  SaveStringToFile(ExpandConstant('{app}\install_diag.txt'), Log, False);
  Result := '';
end;
