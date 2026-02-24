# Building CopyLink for Windows

## Prerequisites

- **Visual Studio 2022** (Community, Professional, or Enterprise) with:
  - .NET desktop development workload
  - .NET Framework 4.7.2 targeting pack
- Or **.NET Framework 4.7.2 SDK** and **MSBuild** (from Visual Studio Build Tools)

## Build Steps

1. Open `CopyLinkWindows.sln` in Visual Studio 2022
2. Set configuration to **Release** and platform to **x64**
3. Build → Build Solution (or press Ctrl+Shift+B)
4. Output: `CopyLinkShellExtension\bin\x64\Release\CopyLinkShellExtension.dll`

## Manual Registration (Testing)

Run Command Prompt **as Administrator**:

```batch
cd CopyLinkShellExtension\bin\x64\Release
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /codebase CopyLinkShellExtension.dll
taskkill /f /im explorer.exe
start explorer.exe
```

## Unregister (for rebuild)

```batch
%windir%\Microsoft.NET\Framework64\v4.0.30319\regasm.exe /u CopyLinkShellExtension.dll
taskkill /f /im explorer.exe
start explorer.exe
```

## Building the Installer

1. Install [Inno Setup](https://jrsoftware.org/isdl.php)
2. Open `Installer\CopyLinkWindows-Setup.iss` in Inno Setup Compiler
3. Build → Compile
4. Output: `Installer\Output\CopyLink-Windows-Setup.exe`

**Note:** Build the C# project first so the DLL exists in `bin\x64\Release\`.

## Testing

1. Install the generated `CopyLink-Windows-Setup.exe` on a test machine
2. Right-click files and folders to confirm CopyLink appears
3. Verify:
   - Copy File Path → UNC path output
   - Copy Folder Path → parent folder path output
   - Copy as File URL → `file://` URL output
4. Log off and log back on if context menu updates are not immediately visible

## Windows 11 Context Menu Behavior

The installer includes automatic Windows 11 detection and offers users a choice:

**How it works:**
1. Installer detects Windows 11 (build 22000+)
2. After successful installation, shows dialog explaining context menu options
3. User chooses whether to restore classic menu or keep new Windows 11 menu
4. Registry change applied if user selects "YES"

**Technical details:**
- Registry key: `HKCU\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32`
- Setting value to empty string restores classic menu
- Deleting the key restores Windows 11 menu
- Change requires logoff/logon to take effect

**Testing on Windows 11:**
- Test both YES and NO options
- Verify classic menu shows CopyLink immediately (YES option)
- Verify new menu requires "Show more options" (NO option)
- Confirm behavior persists after logoff/logon

## Uninstall Behavior and Restart Requirement

**The Problem:**

Shell extension DLLs are loaded into Explorer.exe memory and remain locked even after uninstallation. This prevents:
- Clean file deletion during uninstall
- Successful reinstallation without reboot
- Proper file replacement during upgrades

**The Solution:**

The installer implements a mandatory restart notification after uninstall:

1. **During uninstall (usUninstall step):**
   - RegAsm unregisters the COM class
   - Logs any errors but doesn't block uninstall

2. **After uninstall (usPostUninstall step):**
   - Shows clear dialog explaining restart requirement
   - Offers immediate restart via `shutdown /r /t 5`
   - User can choose to restart later

**Testing uninstall flow:**

1. Install CopyLink and log off/on (to load shell extension)
2. Uninstall via Settings → Apps
3. Verify restart prompt appears
4. Test "YES" option (computer restarts in 5 seconds)
5. Reinstall and test "NO" option (manual restart required)
6. Attempt reinstall without restart (should warn about locked files if using `restartreplace` flag)

**Technical details:**

- Restart command: `shutdown /r /t 5 /c "Restarting to complete CopyLink uninstallation..."`
- 5-second delay allows user to save work
- Restart is optional but strongly recommended
- Without restart, DLL remains at `C:\Program Files\CopyLink\CopyLinkShellExtension.dll`
