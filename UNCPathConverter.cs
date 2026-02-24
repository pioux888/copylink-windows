using System;
using System.IO;
using System.Runtime.InteropServices;

namespace CopyLinkShellExtension
{
    /// <summary>
    /// Converts local or mapped drive paths to UNC paths.
    /// </summary>
    public static class UNCPathConverter
    {
        private const int UNIVERSAL_NAME_INFO_LEVEL = 1;
        private const int ERROR_MORE_DATA = 234;
        private const int NOERROR = 0;

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct UNIVERSAL_NAME_INFO
        {
            [MarshalAs(UnmanagedType.LPWStr)]
            public string lpUniversalName;
        }

        [DllImport("mpr.dll", CharSet = CharSet.Unicode)]
        private static extern int WNetGetUniversalName(
            string lpLocalPath,
            int dwInfoLevel,
            IntPtr lpBuffer,
            ref int lpBufferSize);

        /// <summary>
        /// Converts a local or mapped drive path to UNC path.
        /// </summary>
        /// <param name="localPath">Path to convert (e.g., "Z:\folder\file.txt")</param>
        /// <returns>UNC path (e.g., "\\server\share\folder\file.txt") or original path if already UNC or conversion fails</returns>
        public static string ConvertToUNC(string localPath)
        {
            if (string.IsNullOrEmpty(localPath))
            {
                return localPath ?? string.Empty;
            }

            localPath = localPath.Trim('"');

            if (localPath.StartsWith(@"\\"))
            {
                return localPath;
            }

            if (localPath.StartsWith(@"C:\", StringComparison.OrdinalIgnoreCase))
            {
                return localPath;
            }

            try
            {
                int bufferSize = 0;
                int result = WNetGetUniversalName(localPath, UNIVERSAL_NAME_INFO_LEVEL, IntPtr.Zero, ref bufferSize);

                if (result == ERROR_MORE_DATA)
                {
                    IntPtr buffer = Marshal.AllocHGlobal(bufferSize);

                    try
                    {
                        result = WNetGetUniversalName(localPath, UNIVERSAL_NAME_INFO_LEVEL, buffer, ref bufferSize);

                        if (result == NOERROR)
                        {
                            var info = Marshal.PtrToStructure<UNIVERSAL_NAME_INFO>(buffer);
                            return info.lpUniversalName ?? localPath;
                        }
                    }
                    finally
                    {
                        Marshal.FreeHGlobal(buffer);
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"UNC conversion failed: {ex.Message}");
            }

            return localPath;
        }

        /// <summary>
        /// Gets the directory path from a full file path.
        /// </summary>
        public static string GetDirectoryPath(string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
            {
                return filePath ?? string.Empty;
            }

            try
            {
                if (File.Exists(filePath))
                {
                    return Path.GetDirectoryName(filePath) ?? filePath;
                }

                if (Directory.Exists(filePath))
                {
                    return filePath;
                }

                return Path.GetDirectoryName(filePath) ?? filePath;
            }
            catch
            {
                return filePath;
            }
        }
    }
}
