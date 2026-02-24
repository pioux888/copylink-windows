#define MyAppName "CopyLink for Windows"
#define MyAppVersion "1.0"
#define MyAppPublisher "Pioux"
#define MyAppURL "https://copylinkexcel.com"
#define MyInstallerGUID "{{9F2C4D6E-1A73-4B88-AF6C-2D9E5B7413C1}"
#define MyAppGUID "{{B2C3D4E5-F6A7-8901-BCDE-F12345678902}"

[Setup]
AppId={#MyInstallerGUID}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL}
AppSupportURL={#MyAppURL}
AppUpdatesURL={#MyAppURL}
SetupIconFile=..\icon.ico
DefaultDirName={autopf}\CopyLink
DisableProgramGroupPage=yes
LicenseFile=..\LICENSE
OutputDir=Output
OutputBaseFilename=CopyLink-Windows-Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
WizardImageFile=WizardImage.bmp
WizardSmallImageFile=WizardSmallImage.bmp
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\CopyLinkShellExtension\bin\x64\Release\CopyLinkShellExtension.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\CopyLinkShellExtension\bin\x64\Release\SharpShell.dll"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\LICENSE"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\README.md"; DestDir: "{app}"; Flags: ignoreversion isreadme

[Registry]
; Register context menu handler for files
Root: HKCR; Subkey: "*\shellex\ContextMenuHandlers\CopyLinkShellExtension"; ValueType: string; ValueName: ""; ValueData: "{#MyAppGUID}"; Flags: uninsdeletekey

; Register context menu handler for folders
Root: HKCR; Subkey: "Directory\shellex\ContextMenuHandlers\CopyLinkShellExtension"; ValueType: string; ValueName: ""; ValueData: "{#MyAppGUID}"; Flags: uninsdeletekey

; COM class cleanup on uninstall
Root: HKCR; Subkey: "CLSID\{#MyAppGUID}"; Flags: uninsdeletekey

[Code]
function IsWindows11(): Boolean;
var
  Version: TWindowsVersion;
begin
  GetWindowsVersionEx(Version);
  // Windows 11 is build 22000 or higher
  Result := Version.Build >= 22000;
end;

function InitializeSetup(): Boolean;
var
  Release: Cardinal;
begin
  // Check for .NET Framework 4.7.2 or higher (Release DWORD >= 461808)
  if RegQueryDWordValue(HKLM, 'SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full', 'Release', Release) then
  begin
    if Release >= 461808 then
    begin
      Result := True;
    end
    else
    begin
      MsgBox('.NET Framework 4.7.2 or higher is required.' + #13#10 +
             'You have an older version installed.' + #13#10 +
             'Please download .NET Framework 4.7.2 from Microsoft.', mbError, MB_OK);
      Result := False;
    end;
  end
  else
  begin
    MsgBox('.NET Framework 4.7.2 or higher is required.' + #13#10 +
           'Please install it from Microsoft before installing CopyLink.', mbError, MB_OK);
    Result := False;
  end;
end;

procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  RegAsmPath: String;
begin
  if CurStep = ssPostInstall then
  begin
    RegAsmPath := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\regasm.exe');

    // Register the shell extension
    if Exec(RegAsmPath,
            '/codebase "' + ExpandConstant('{app}\CopyLinkShellExtension.dll') + '"',
            '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      if ResultCode = 0 then
      begin
        // Registration successful
        MsgBox('Installation complete! Please log off Windows and log back on for the CopyLink context menu to appear.', mbInformation, MB_OK);

        // Check if Windows 11 and offer context menu fix
        if IsWindows11() then
        begin
          if MsgBox('Windows 11 Context Menu Options' + #13#10 + #13#10 +
                    'Windows 11 hides shell extensions like CopyLink under "Show more options" (extra click required).' + #13#10 + #13#10 +
                    'Would you like to restore the classic Windows 10-style context menu?' + #13#10 + #13#10 +
                    'YES = CopyLink appears immediately (affects all right-click menus)' + #13#10 +
                    'NO = Keep new Windows 11 menu (CopyLink under "Show more options")',
                    mbConfirmation, MB_YESNO) = IDYES then
          begin
            // Apply registry fix to restore classic context menu
            RegWriteStringValue(HKCU, 'Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32', '', '');
            MsgBox('Classic context menu enabled! Please log off and log back on for changes to take effect.', mbInformation, MB_OK);
          end;
        end;
      end
      else
      begin
        MsgBox('Warning: Shell extension registration completed with errors (code ' + IntToStr(ResultCode) + ').' + #13#10 +
               'The context menu may not appear. Please restart your computer and try again.', mbError, MB_OK);
      end;
    end
    else
    begin
      MsgBox('Error: Failed to register shell extension.' + #13#10 +
             'The CopyLink context menu will not appear.' + #13#10 +
             'Please contact support or try reinstalling.', mbError, MB_OK);
    end;
  end;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  ResultCode: Integer;
  RegAsmPath: String;
begin
  if CurUninstallStep = usUninstall then
  begin
    RegAsmPath := ExpandConstant('{win}\Microsoft.NET\Framework64\v4.0.30319\regasm.exe');

    // Unregister the shell extension
    if Exec(RegAsmPath,
            '/u "' + ExpandConstant('{app}\CopyLinkShellExtension.dll') + '"',
            '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    begin
      if ResultCode <> 0 then
      begin
        // Log warning but don't block uninstall
        Log('Warning: Shell extension unregistration completed with errors (code ' + IntToStr(ResultCode) + ')');
      end;
    end;
  end
  else if CurUninstallStep = usPostUninstall then
  begin
    // Inform user that restart is required
    if MsgBox('Uninstallation complete.' + #13#10 + #13#10 +
              'IMPORTANT: You must restart your computer before reinstalling CopyLink.' + #13#10 + #13#10 +
              'The shell extension DLL remains locked by Windows Explorer until restart.' + #13#10 +
              'Reinstalling without restarting may cause installation failures.' + #13#10 + #13#10 +
              'Would you like to restart now?',
              mbConfirmation, MB_YESNO) = IDYES then
    begin
      // Restart computer
      Exec('shutdown', '/r /t 5 /c "Restarting to complete CopyLink uninstallation..."', '', SW_HIDE, ewNoWait, ResultCode);
    end;
  end;
end;
