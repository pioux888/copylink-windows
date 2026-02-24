using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace CopyLinkShellExtension
{
    /// <summary>
    /// CopyLink context menu extension for Windows Explorer.
    /// Adds menu items to copy UNC paths to clipboard.
    /// </summary>
    [ComVisible(true)]
    [Guid("B2C3D4E5-F6A7-8901-BCDE-F12345678902")]
    [COMServerAssociation(AssociationType.AllFilesAndFolders)]
    public class CopyLinkContextMenu : SharpContextMenu
    {
        protected override bool CanShowMenu()
        {
            try
            {
                return SelectedItemPaths != null && SelectedItemPaths.Any();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CanShowMenu failed: {ex.Message}");
                return false;
            }
        }

        protected override ContextMenuStrip CreateMenu()
        {
            var menu = new ContextMenuStrip();

            try
            {
                System.Drawing.Image icon = null;
                try
                {
                    icon = Properties.Resources.CopyLinkIcon;
                }
                catch
                {
                    // Icon load failure - menu works without icon
                }

                var mainMenuItem = new ToolStripMenuItem
                {
                    Text = "CopyLink",
                    Image = icon
                };

                var copyFilePath = new ToolStripMenuItem
                {
                    Text = "Copy File Path",
                    ToolTipText = "Copy full UNC path to file"
                };
                copyFilePath.Click += (sender, args) => CopyPath(PathType.FilePath);
                mainMenuItem.DropDownItems.Add(copyFilePath);

                var copyFolderPath = new ToolStripMenuItem
                {
                    Text = "Copy Folder Path",
                    ToolTipText = "Copy UNC path to parent folder"
                };
                copyFolderPath.Click += (sender, args) => CopyPath(PathType.FolderPath);
                mainMenuItem.DropDownItems.Add(copyFolderPath);

                mainMenuItem.DropDownItems.Add(new ToolStripSeparator());

                var copyFileURL = new ToolStripMenuItem
                {
                    Text = "Copy as File URL",
                    ToolTipText = "Copy as file:// URL (clickable in Outlook, Word, Excel)"
                };
                copyFileURL.Click += (sender, args) => CopyPath(PathType.FileURL);
                mainMenuItem.DropDownItems.Add(copyFileURL);

                menu.Items.Add(mainMenuItem);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CreateMenu failed: {ex.Message}");
            }

            return menu;
        }

        private enum PathType
        {
            FilePath,
            FolderPath,
            FileURL
        }

        private void CopyPath(PathType pathType)
        {
            try
            {
                string selectedPath = SelectedItemPaths?.FirstOrDefault();

                if (string.IsNullOrEmpty(selectedPath))
                {
                    ShowError("No file or folder selected.");
                    return;
                }

                string uncPath = UNCPathConverter.ConvertToUNC(selectedPath);
                string textToCopy;

                switch (pathType)
                {
                    case PathType.FilePath:
                        textToCopy = uncPath;
                        break;
                    case PathType.FolderPath:
                        textToCopy = UNCPathConverter.GetDirectoryPath(uncPath);
                        break;
                    case PathType.FileURL:
                        textToCopy = ClipboardHelper.CreateFileURL(uncPath);
                        break;
                    default:
                        textToCopy = uncPath;
                        break;
                }

                if (string.IsNullOrEmpty(textToCopy))
                {
                    ShowError("Could not determine path to copy.");
                    return;
                }

                bool success = ClipboardHelper.CopyToClipboard(textToCopy);

                if (!success)
                {
                    ShowError("Failed to copy to clipboard. Please try again.");
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"CopyPath failed: {ex.Message}");
                ShowError($"Error: {ex.Message}");
            }
        }

        private void ShowError(string message)
        {
            try
            {
                MessageBox.Show(
                    message,
                    "CopyLink Error",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ShowError failed: {ex.Message}");
            }
        }
    }
}
