using System;
using System.Windows.Forms;

namespace CopyLinkShellExtension
{
    /// <summary>
    /// Clipboard operations with retry logic.
    /// </summary>
    public static class ClipboardHelper
    {
        /// <summary>
        /// Copies text to clipboard with retry logic.
        /// </summary>
        public static bool CopyToClipboard(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return false;
            }

            try
            {
                for (int i = 0; i < 3; i++)
                {
                    try
                    {
                        Clipboard.SetText(text, TextDataFormat.UnicodeText);
                        return true;
                    }
                    catch (System.Runtime.InteropServices.ExternalException)
                    {
                        System.Threading.Thread.Sleep(100);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Clipboard operation failed: {ex.Message}");
                return false;
            }

            return false;
        }

        /// <summary>
        /// Converts a UNC path to a file:// URL format.
        /// Outlook, Word, and Excel automatically convert file:// URLs into clickable links.
        /// </summary>
        /// <param name="uncPath">UNC path (e.g., \\server\share\folder\file.txt)</param>
        /// <returns>file:// URL (e.g., file://server/share/folder/file.txt)</returns>
        public static string CreateFileURL(string uncPath)
        {
            if (string.IsNullOrEmpty(uncPath))
            {
                return string.Empty;
            }

            try
            {
                string fileUrl = uncPath.Replace('\\', '/');

                // UNC path: \\server\share\folder\file.txt → file://server/share/folder/file.txt
                if (uncPath.StartsWith(@"\\") || uncPath.StartsWith("//"))
                {
                    fileUrl = fileUrl.TrimStart('/');
                    return "file://" + fileUrl;
                }

                // Local path: C:\folder\file.txt → file:///C:/folder/file.txt
                return "file:///" + fileUrl;
            }
            catch
            {
                return uncPath; // Fallback to original path on error
            }
        }
    }
}
