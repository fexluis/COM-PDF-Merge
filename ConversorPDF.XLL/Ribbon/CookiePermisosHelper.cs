using System;
using System.IO;
using System.Text;
using System.Security.Cryptography;

namespace ConversorPDF.XLL.Ribbon
{
    /// <summary>
    /// Helper to store user profile permissions locally.
    /// Uses %LOCALAPPDATA%\ConversorPDF\Data\userserial.txt and stores data encrypted via DPAPI
    /// to protect the information.
    /// </summary>
    internal static class CookiePermisosHelper
    {
        private static string GetCookieFilePath()
        {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrEmpty(appData))
                appData = Environment.GetEnvironmentVariable("LOCALAPPDATA") ?? @"C:\Temp";
            
            string dir = Path.Combine(appData, "ConversorPDF", "Data");
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            
            return Path.Combine(dir, "userserial.txt");
        }

        public static string LoadCookie(string username)
        {
            try
            {
                string path = GetCookieFilePath();
                if (File.Exists(path))
                {
                    byte[] encryptedBytes = File.ReadAllBytes(path);
                    if (encryptedBytes.Length > 0)
                    {
                        byte[] bytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
                        string content = Encoding.UTF8.GetString(bytes);

                        var parts = content.Split('|');
                        if (parts.Length == 2 && string.Equals(parts[0], username, StringComparison.OrdinalIgnoreCase))
                        {
                            return parts[1];
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Core.Helpers.LoggerHelper.Warning($"Error leyendo cookie de permisos: {ex.Message}");
            }
            return null;
        }

        public static void SaveCookie(string username, string permisos)
        {
            try
            {
                string path = GetCookieFilePath();
                string content = $"{username}|{permisos}";
                byte[] bytes = Encoding.UTF8.GetBytes(content);
                byte[] encryptedBytes = ProtectedData.Protect(bytes, null, DataProtectionScope.CurrentUser);
                
                File.WriteAllBytes(path, encryptedBytes);
            }
            catch (Exception ex)
            {
                Core.Helpers.LoggerHelper.Warning($"Error guardando cookie de permisos: {ex.Message}");
            }
        }
    }
}
