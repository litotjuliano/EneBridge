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
#define MyAppVersion "1.0.1"
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
Source: "..\prerequisites\accessdatabaseengine_X64.exe"; DestDir: "{app}\prerequisites"; Flags: ignoreversion
Source: "..\prerequisites\dotnet-runtime-8.0.23-win-x64.exe"; DestDir: "{tmp}"; Flags: deleteafterinstall

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{group}\Uninstall {#MyAppName}"; Filename: "{uninstallexe}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\prerequisites\accessdatabaseengine_X64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing Access Database Engine..."; Check: not IsAccessDatabaseEngineInstalled
Filename: "{tmp}\dotnet-runtime-8.0.23-win-x64.exe"; Parameters: "/quiet /norestart"; StatusMsg: "Installing .NET 8 Desktop Runtime..."; Check: not IsDotNet8DesktopRuntimeInstalled
Filename: "{app}\{#MyAppExeName}"; Description: "Launch {#MyAppName}"; Flags: nowait postinstall skipifsilent

[Code]
function GetAceInprocServerPath(): String;
var
  ClsidStr: String;
  DllPath: String;
begin
  { Resolves what DLL Microsoft.ACE.OLEDB.12.0 actually points at: read the ProgID's CLSID, then
    that CLSID's InprocServer32 default value, checking both the native and Wow6432Node registry
    views since the provider can be registered under either depending on what else is installed. }
  Result := '';

  if not (RegQueryStringValue(HKLM, 'SOFTWARE\Classes\Microsoft.ACE.OLEDB.12.0\CLSID', '', ClsidStr) or
          RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Classes\Microsoft.ACE.OLEDB.12.0\CLSID', '', ClsidStr)) then
    exit;

  if RegQueryStringValue(HKLM, 'SOFTWARE\Classes\CLSID\' + ClsidStr + '\InprocServer32', '', DllPath) then
  begin
    Result := DllPath;
    exit;
  end;

  if RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Classes\CLSID\' + ClsidStr + '\InprocServer32', '', DllPath) then
    Result := DllPath;
end;

function IsAccessDatabaseEngineInstalled(): Boolean;
var
  DllPath: String;
begin
  { A key-existence check alone isn't enough: Office's Click-to-Run install also registers
    Microsoft.ACE.OLEDB.12.0, pointing at its own sandboxed copy of ACEOLEDB.DLL (path contains
    '\root\VFS\'), which crashes natively (AccessViolationException) on CREATE TABLE/write --
    confirmed live on a real machine, see
    docs/superpowers/specs/2026-09-07-dbf-export-crash-prevention-design.md. Only the standalone
    redistributable's copy (this installer's own accessdatabaseengine_X64.exe) counts as properly
    installed here. }
  DllPath := GetAceInprocServerPath();
  { Lowercase both sides before comparing: Pos is case-sensitive, and the VFS path's casing isn't
    a documented Microsoft contract -- failing to match here would fail OPEN (report "installed"
    for a still-broken Click-to-Run binding), the same silent-skip failure mode this fix exists to
    close. }
  Result := (DllPath <> '') and (Pos('\root\vfs\', Lowercase(DllPath)) = 0);
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
