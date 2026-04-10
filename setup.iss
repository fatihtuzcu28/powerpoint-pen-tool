; PPTKalem Inno Setup Script
; Inno Setup 6.x ile derleyin: https://jrsoftware.org/isinfo.php

#define MyAppName "PPTKalem"
#define MyAppVersion "1.0"
#define MyAppPublisher "PPTKalem"
#define MyAppURL ""

[Setup]
AppId={{A3F2B8C1-4D5E-6F7A-8B9C-0D1E2F3A4B5C}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=installer_output
OutputBaseFilename=PPTKalem_Setup
Compression=lzma2
SolidCompression=yes
PrivilegesRequired=admin
WizardStyle=modern
SetupIconFile=compiler:SetupClassicIcon.ico
UninstallDisplayName=PPTKalem - PowerPoint Kalem Add-in
CloseApplications=force
CloseApplicationsFilter=POWERPNT.EXE

[Languages]
Name: "turkish"; MessagesFile: "compiler:Languages\Turkish.isl"

[Files]
; Ana DLL ve bağımlılıklar
Source: "PPTKalem\bin\Release\net481\PPTKalem.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "PPTKalem\bin\Release\net481\Extensibility.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "PPTKalem\bin\Release\net481\Microsoft.Office.Interop.PowerPoint.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "PPTKalem\bin\Release\net481\stdole.dll"; DestDir: "{app}"; Flags: ignoreversion

; AddIns klasörüne de kopyala (PowerPoint buradan yükler)
Source: "PPTKalem\bin\Release\net481\PPTKalem.dll"; DestDir: "{userappdata}\Microsoft\AddIns"; Flags: ignoreversion

[Code]
var
  RegAsmPath: String;

function GetPowerPointPath(): String;
var
  PPTPath: String;
begin
  Result := '';
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE', '', PPTPath) then
    Result := PPTPath
  else if RegQueryStringValue(HKCU, 'SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\POWERPNT.EXE', '', PPTPath) then
    Result := PPTPath;
end;

function IsOffice32Bit(): Boolean;
var
  PPTPath: String;
begin
  PPTPath := GetPowerPointPath();
  if PPTPath <> '' then
    Result := Pos('Program Files (x86)', PPTPath) > 0
  else
    Result := True; // varsayılan 32-bit
end;

function GetRegAsmPath(): String;
begin
  if IsOffice32Bit() then
    Result := 'C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe'
  else
    Result := 'C:\Windows\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe';
end;

procedure RegisterDLL();
var
  ResultCode: Integer;
  DllPath: String;
begin
  RegAsmPath := GetRegAsmPath();
  DllPath := ExpandConstant('{userappdata}\Microsoft\AddIns\PPTKalem.dll');

  Log('RegAsm path: ' + RegAsmPath);
  Log('DLL path: ' + DllPath);
  if IsOffice32Bit() then
    Log('Office 32-bit: True')
  else
    Log('Office 32-bit: False');

  if FileExists(RegAsmPath) and FileExists(DllPath) then
  begin
    Exec(RegAsmPath, '"' + DllPath + '" /codebase', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);
    if ResultCode <> 0 then
      MsgBox('COM kaydı sırasında hata oluştu (kod: ' + IntToStr(ResultCode) + '). ' +
             'PowerPoint''i yeniden başlatmayı deneyin.', mbError, MB_OK)
    else
      Log('COM registration successful');
  end
  else
  begin
    if not FileExists(RegAsmPath) then
      MsgBox('RegAsm bulunamadı: ' + RegAsmPath + #13#10 +
             '.NET Framework 4.x yüklü olduğundan emin olun.', mbError, MB_OK);
  end;
end;

procedure UnregisterDLL();
var
  ResultCode: Integer;
  DllPath: String;
begin
  RegAsmPath := GetRegAsmPath();
  DllPath := ExpandConstant('{userappdata}\Microsoft\AddIns\PPTKalem.dll');

  if FileExists(RegAsmPath) and FileExists(DllPath) then
    Exec(RegAsmPath, '"' + DllPath + '" /unregister', '', SW_HIDE, ewWaitUntilTerminated, ResultCode);

  // Registry'den de temizle
  RegDeleteKeyIncludingSubkeys(HKCU, 'Software\Microsoft\Office\PowerPoint\Addins\PPTKalem.Connect');
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    RegisterDLL();
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
begin
  if CurUninstallStep = usUninstall then
    UnregisterDLL();
end;

function InitializeSetup(): Boolean;
var
  PPTPath: String;
begin
  Result := True;
  PPTPath := GetPowerPointPath();
  if PPTPath = '' then
  begin
    if MsgBox('PowerPoint bulunamadı. Kuruluma devam etmek istiyor musunuz?',
              mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

[Icons]
Name: "{group}\{#MyAppName} Kaldır"; Filename: "{uninstallexe}"

[UninstallDelete]
Type: files; Name: "{userappdata}\Microsoft\AddIns\PPTKalem.dll"
