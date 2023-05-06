#define MyAppName "XLKitLearn"
#define MyAppVersion "version_placeholder"
#define MyAppPublisher "Dynamic Analytics LLC"
#define MyAppURL "https://www.xlkitlearn.com"

[Setup]
; SignTool=signtool
AppId={{YaocNmA99ZWqRnwKgQKZKayRJojMjq}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
DefaultDirName={localappdata}\{#MyAppName}
DisableDirPage=yes
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename={#MyAppName}
Compression=lzma
SolidCompression=yes
PrivilegesRequired=none
UninstallDisplayName={#MyAppName}
SetupIconFile="{#GetEnv('GITHUB_WORKSPACE')}\logos\logo.ico"
UninstallDisplayIcon={app}\logo.ico

[CustomMessages]
InstallingLabel=

[InstallDelete]
Type: filesandordirs; Name: "{app}"

[Files]
Source: "{#GetEnv('pythonLocation')}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "{#GetEnv('GITHUB_WORKSPACE')}\XLKitLearn.xltm"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#GetEnv('GITHUB_WORKSPACE')}\.github\package_compile.py"; DestDir: "{app}"; Flags: ignoreversion
Source: "{#GetEnv('GITHUB_WORKSPACE')}\logos\logo.ico"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{group}\XLKitLearn"; Filename: "{app}\XLKitLearn.xltm"

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
Filename: "cmd.exe"; Parameters: "/c ""{app}\python.exe"" package_compile.py"; WorkingDir: "{app}"