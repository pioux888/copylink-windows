# CopyLink for Windows

**Free Windows shell extension for copying network file paths as clickable links.**

Perfect for teams sharing files on network drives - no more broken "Z:\" paths!

---

## What It Does

Adds a **"Copy as File URL"** option to your Windows right-click menu that:

- ✅ Converts mapped drives (Z:\, P:\) to UNC paths (`\\server\share\...`)
- ✅ Creates clickable `file://` links for Outlook, Word, Excel
- ✅ Works on files AND folders
- ✅ Copies to clipboard - just paste anywhere

---

## FAQ

### Why use CopyLink when Windows already has "Copy as path"?

Great question! Windows' built-in "Copy as path" has three major limitations that CopyLink solves:

#### Problem 1: Mapped Drive Letters Don't Work for Recipients

**Windows "Copy as path" gives you:**
"Z:\Marketing\Q4 Campaign\Budget.xlsx"

**CopyLink gives you:**
\\fileserver\marketing\Q4 Campaign\Budget.xlsx

**Why this matters:**
- Your `Z:` drive might map to `\\fileserver\marketing`
- But your colleague's `Z:` might map to `\\fileserver\hr` (or not exist at all)
- When you email them `Z:\Marketing\...`, they can't open it
- CopyLink converts to the universal UNC path that works for everyone on the network

#### Problem 2: Annoying Quotation Marks

**Windows adds quotes:**
"Z:\Marketing\Q4 Campaign\Budget.xlsx"

**CopyLink doesn't:**
\\fileserver\marketing\Q4 Campaign\Budget.xlsx

Those quotes are annoying when pasting into emails, Slack, or Teams. You have to manually delete them every time.

#### Problem 3: Not Clickable in Outlook/Word/Excel

**Windows "Copy as path" in an email:**
"Z:\Marketing\Budget.xlsx"
→ Just plain text, not clickable

**CopyLink's "Copy as File URL":**
file://fileserver/marketing/Budget.xlsx
→ **Automatically becomes a blue clickable hyperlink** in Outlook, Word, and Excel

#### Real-World Example

**Using Windows "Copy as path":**
Hi Sarah,
Please review "Z:\Marketing\Q4 Campaign\Budget.xlsx"
- Sarah clicks → Nothing happens (it's just text)
- Her Z: drive maps to a different folder
- She emails back: "I can't find this file, where is it?"

**Using CopyLink "Copy as File URL":**
Hi Sarah,
Please review: file://fileserver/marketing/Q4 Campaign/Budget.xlsx
- Outlook makes it a clickable link automatically
- Sarah clicks → File opens immediately
- Works for everyone, regardless of drive mappings

#### Summary

| Feature | Windows "Copy as path" | CopyLink |
|---------|------------------------|----------|
| **Path type** | Mapped drive (`Z:\`) | Universal UNC (`\\server\share\`) |
| **Works for recipients?** | ❌ Only if same drive mapping | ✅ Works for anyone on network |
| **Adds quotes?** | ✅ Yes (annoying) | ❌ No |
| **Clickable in Outlook?** | ❌ No | ✅ Yes (file:// format) |
| **Folder path option?** | ❌ No | ✅ Yes |

**Bottom line:** Windows "Copy as path" = *"Copy this path that only works on MY computer"*  
**CopyLink** = *"Copy this path that works on EVERYONE's computer"*

### Why do I need to click "Show more options" in Windows 11?

Windows 11 redesigned the right-click context menu and hides classic shell extensions under "Show more options" by default.

**During installation, CopyLink will ask if you want to restore the classic menu.** This is optional and affects all right-click menus in Windows (not just CopyLink).

**Your choices:**
- **Restore classic menu** = CopyLink appears immediately when you right-click (like Windows 10)
- **Keep Windows 11 menu** = CopyLink appears under "Show more options" (extra click required)

You can change this setting anytime using the registry files provided in the "Windows 11 Users" section below.

**Alternative:** Press **Shift + F10** instead of right-clicking to open the classic menu directly.

### Do I need to restart after uninstalling?

**Yes, restarting is strongly recommended** before reinstalling CopyLink or installing a newer version.

**Why?**

Shell extensions are loaded into Windows Explorer's memory. Even after uninstalling, the DLL file remains locked by Explorer until you restart. If you try to reinstall without restarting:
- Installation may fail with "file in use" errors
- Installer may crash trying to overwrite locked files
- Old version may remain partially installed

**What happens during uninstall:**

After uninstallation completes, you'll see a dialog asking:
Would you like to restart now?

- **YES** = Computer restarts in 5 seconds (recommended)
- **NO** = You can restart manually later

**If you chose NO:** Remember to restart before reinstalling CopyLink.

---

## Why You Need This

**The Problem:**

When you right-click a file on a mapped drive (Z:\) and use Windows' built-in "Copy as path", you get:
```
Z:\Projects\Report.xlsx
```

This path **only works on your computer** because Z:\ is your personal drive mapping.

**The Solution:**

CopyLink gives you the **real network path** that works for everyone:
```
\\CompanyServer\Shared\Projects\Report.xlsx
```

Or as a clickable link in Outlook:
```
file://CompanyServer/Shared/Projects/Report.xlsx
```

**Your colleagues can click and open the file directly!**

---

## Installation

### Quick Install (30 seconds):

1. **Download** `CopyLink-Windows-Setup.exe` from the [Releases page](https://github.com/pioux888/copylink-windows/releases)
2. **Run the installer** (requires administrator rights)
3. Click "Yes" when Windows asks for permission
4. Follow the installation wizard
5. **Right-click any file** in Windows Explorer
6. **Select "Copy as File URL"** from the menu
7. **Paste** into Outlook, Teams, or any app!

That's it! ✅

### Windows 11 Users

Windows 11's redesigned context menu hides classic shell extensions under "Show more options," requiring an extra click to access CopyLink.

**During installation, you'll be asked:**
- **YES** = Restore classic Windows 10-style menu (CopyLink appears immediately)
- **NO** = Keep Windows 11 menu (CopyLink under "Show more options")

**To change this later:**

Enable classic menu (CopyLink visible immediately):
```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}\InprocServer32]
@=""
```

Restore Windows 11 menu (hide CopyLink under "Show more options"):
```reg
Windows Registry Editor Version 5.00

[-HKEY_CURRENT_USER\Software\Classes\CLSID\{86ca1aa0-34aa-4e8b-a509-50c905bae2a2}]
```

After applying either registry change, log off and log back on.

### Uninstalling

To uninstall CopyLink:

1. Open **Settings → Apps → Installed apps**
2. Find **CopyLink for Windows**
3. Click **Uninstall**
4. **Important:** You will be prompted to restart your computer

**Why restart is required:**

Shell extensions remain locked in Windows Explorer's memory after uninstall. Restarting ensures:
- Complete removal of all CopyLink components
- Clean reinstallation if needed later
- No installation failures or crashes

You can choose to restart immediately or restart manually later. **If you plan to reinstall CopyLink, you must restart first.**

### Security Note

This software is **not digitally signed** (code signing certificates cost $400+/year for free open-source software).

Windows will show: *"Windows protected your PC"*

**This is normal. To install:**
1. Click **"More info"**
2. Click **"Run anyway"**

💡 The code is **100% open source** and reviewable on [GitHub](https://github.com/pioux888/copylink-windows).

---

## System Requirements

- ✅ Windows 10 or Windows 11 (64-bit)
- ✅ .NET Framework 4.7.2 or higher (pre-installed on most systems)
- ✅ **Primary use:** Network files on mapped drives (Z:, P:, etc.) - automatically converts to UNC paths (`\\server\share\...`)
- ✅ **Also works with:** Local files on C: drive (copied as-is, with `file:///` URL support)

---

## Usage Examples

### Example 1: Share a File in Outlook

1. Right-click `Report.xlsx` on your Z: drive
2. Select **"Copy as File URL"**
3. Paste into Outlook email
4. Recipients get a clickable link: `file://server/share/Report.xlsx`

### Example 2: Link to a Folder in Teams

1. Right-click a folder in Windows Explorer
2. Select **"Copy as File URL"**
3. Paste in Teams chat
4. Team members click to open the folder directly

---

## Troubleshooting

### ❓ CopyLink doesn't appear in the menu

**Windows 11 24H2/25H2 users:**
- Try **Shift + Right-click** → "Show more options"
- Restart Windows Explorer:
```cmd
taskkill /f /im explorer.exe
start explorer.exe
```

### ❓ Windows Explorer crashes after install

- Uninstall from **Settings → Apps → CopyLink**
- [Report the issue on GitHub](https://github.com/pioux888/copylink-windows/issues)
- We track compatibility with latest Windows builds

### ❓ Link doesn't work when pasted

- Check that the file is on a **network drive** (not C:\ local drive)
- Verify your colleagues have **network access** to the share
- UNC paths require network connectivity

### ❓ Does it work with local C: drive files?

Yes. While CopyLink's main purpose is converting mapped network drives to UNC paths, it also works with local files:
- **Copy File Path:** Copies `C:\path\to\file.xlsx` as-is
- **Copy Folder Path:** Copies the containing folder path
- **Copy as File URL:** Creates `file:///C:/path/to/file.xlsx` (clickable in Outlook/Word)

---

## Uninstallation

**Method 1:** Settings → Apps → CopyLink for Windows → Uninstall

**Method 2:** Run `unins000.exe` from installation directory

Windows Explorer will restart automatically.

---

## Privacy & Security

- ✅ **No telemetry** - doesn't phone home or track usage
- ✅ **No internet connection** - works 100% offline
- ✅ **Open source** - all code visible on GitHub
- ✅ **No ads** - completely free forever

---

## Support the Project

CopyLink is **100% free** and always will be. If it saves you time, consider:

☕ **[Buy me a coffee on Ko-fi](https://ko-fi.com/pioux888)**

Your support helps maintain and improve CopyLink!

---

## Related Tools

**CopyLink Family:**

- 🟢 **[CopyLink for Excel](https://copylinkexcel.com)** - Excel add-in (Windows & Mac)
- 🔵 **CopyLink for Windows** - Shell extension (you are here!)

All tools are free and open source.

---

## License

**MIT License** - Free for personal and commercial use

---

## Support

- 💬 **Issues:** [GitHub Issues](https://github.com/pioux888/copylink-windows/issues)
- 📧 **Email:** admin@copylinkexcel.com
- 🌐 **Website:** [CopyLinkExcel.com](https://copylinkexcel.com)
- ☕ **Donate:** [Ko-fi](https://ko-fi.com/pioux888)

---

## Credits

Built by **Pioux** | Not affiliated with Microsoft Corporation

© 2026 | Version 1.0.0

---

**⭐ If CopyLink saves you time, please star the repo on GitHub or [buy me a coffee](https://ko-fi.com/pioux888)!**
