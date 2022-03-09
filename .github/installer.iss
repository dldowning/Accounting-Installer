#define MyAppName "name_placeholder"
#define MyAppVersion "dev"
#define MyAppPublisher "publisher_placeholder"
#define MyAppURL "https://www.xlwings.org"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
; SignTool=signtool
AppId={{appid_placeholder}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={userpf}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename={#MyAppName}-{#MyAppVersion}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=none
UninstallDisplayName={#MyAppName}

[CustomMessages]
InstallingLabel=

[InstallDelete]
Type: filesandordirs; Name: "{app}"

[Files]
Source: "{#GetEnv('pythonLocation')}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Code]
procedure InitializeWizard;
begin
  with TNewStaticText.Create(WizardForm) do
  begin
    Parent := WizardForm.FilenameLabel.Parent;
    Left := WizardForm.FilenameLabel.Left;
    Top := WizardForm.FilenameLabel.Top;
    Width := WizardForm.FilenameLabel.Width;
    Height := WizardForm.FilenameLabel.Height;
    Caption := ExpandConstant('{cm:InstallingLabel}');
  end;
  WizardForm.FilenameLabel.Visible := False;
end;

[Run]
Filename: "cmd.exe"; Parameters: "/c ""{app}\python.exe"" Lib\site-packages\xlwings\cli.py addin install --dir addins"; WorkingDir: "{app}"; Flags: runhidden
