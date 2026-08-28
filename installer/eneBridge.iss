; eneBridge 2.0 installer.
;
; Build steps (see CLAUDE.md's Commands section):
;   dotnet publish src\eneBridge.Wpf\eneBridge.Wpf.csproj -c Release -o installer\publish
;   iscc installer\eneBridge.iss
; Produces installer\Output\eneBridge-Setup.exe
;
; OutputBaseFilename is deliberately NOT "setup" -- Inno Setup itself warns that "setup.exe" is
; subject to Windows app-compatibility DLL-hijacking shims (they unsafely load extra DLLs like
; version.dll next to any exe with that exact name).
;
; AppId is fixed permanently -- do not change it. It's what lets a newer installer build be
; recognized as an in-place upgrade of the same install rather than a separate parallel one.

#define MyAppName "eneBridge"
#define MyAppVersion "1.0.0"
#define MyAppPublisher "EnE Computer"
#define MyAppExeName "eneBridge.Wpf.exe"

[Setup]
AppId={{0A2AD934-EBEE-4779-9543-BC18E4E9AEDC}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=Output
OutputBaseFilename=eneBridge-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesInstallIn64BitMode=x64compatible
PrivilegesRequired=admin
UninstallDisplayIcon={app}\{#MyAppExeName}

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional shortcuts:"

[Files]
Source: "publish\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
Source: "..\prerequisites\accessdatabaseengine_X64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall
Source: "..\prerequisites\dotnet-runtime-8.0.23-win-x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{tmp}\accessdatabaseengine_X64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing Access Database Engine..."; Check: not IsAccessDatabaseEngineInstalled
Filename: "{tmp}\dotnet-runtime-8.0.23-win-x64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing .NET 8 Desktop Runtime..."; Check: not IsDotNet8DesktopRuntimeInstalled
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function IsAccessDatabaseEngineInstalled(): Boolean;
begin
  { Checks the actual COM registration for the ACE OleDb provider that DbfExportService depends
    on (Provider=Microsoft.ACE.OLEDB.12.0) -- the most direct signal that it's already usable. }
  Result := RegKeyExists(HKLM, 'SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0\CLSID') or
            RegKeyExists(HKLM, 'SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.12.0\CLSID');
end;

function IsDotNet8DesktopRuntimeInstalled(): Boolean;
var
  FindRec: TFindRec;
  BasePath: String;
begin
  Result := False;
  BasePath := ExpandConstant('{commonpf64}\dotnet\shared\Microsoft.WindowsDesktop.App');
  if DirExists(BasePath) then
  begin
    if FindFirst(BasePath + '\8.*', FindRec) then
    begin
      try
        Result := True;
      finally
        FindClose(FindRec);
      end;
    end;
  end;
end;
