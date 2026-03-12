using System;
using System.IO;
using Microsoft.Win32;

namespace ConversorPDF
{
    public static class OneDriveHelper
    {
        /// <summary>
        /// Cross-platform equivalent to get the local path of OneDrive/SharePoint
        /// synchronized Microsoft Office files.
        /// </summary>
        public static string GetOneDriveLocalPath(string path)
        {
            if (string.IsNullOrEmpty(path))
                return path;

            if (path.Length >= 2 && path[1] == ':')
            {
                return path;
            }

            if (path.StartsWith(@"\\", StringComparison.Ordinal))
            {
                return path;
            }

            if (Uri.TryCreate(path, UriKind.Absolute, out Uri uri))
            {
                if (uri.IsFile)
                    return uri.LocalPath;
            }

            string urlHost = uri?.Host ?? string.Empty;
            if (urlHost.Length == 0 &&
                path.IndexOf("sharepoint", StringComparison.OrdinalIgnoreCase) == -1 &&
                path.IndexOf("onedrive", StringComparison.OrdinalIgnoreCase) == -1)
            {
                return path;
            }

            string oneDriveFullPath = GetOneDrivePathFromRegistryFast();

            if (string.IsNullOrEmpty(oneDriveFullPath))
            {
                oneDriveFullPath = GetOneDrivePathFromFolderFast(Environment.UserName);
            }

            if (string.IsNullOrEmpty(oneDriveFullPath))
            {
                return path;
            }

            string urlPath = path;
            if (uri != null && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
            {
                urlPath = Uri.UnescapeDataString(uri.AbsolutePath);
            }

            string relative = ExtractRelativePathFromUrlPath(urlPath);
            if (string.IsNullOrEmpty(relative))
            {
                return path;
            }

            relative = relative.Replace("/", "\\").TrimStart('\\');
            string resultado = Path.Combine(oneDriveFullPath, relative);
            resultado = resultado.Replace("\\\\", "\\");

            try
            {
                if (File.Exists(resultado))
                    return resultado;

                string parent = Path.GetDirectoryName(resultado);
                if (!string.IsNullOrEmpty(parent) && Directory.Exists(parent))
                    return resultado;
            }
            catch
            {
            }

            return resultado;
        }

        private static string GetOneDrivePathFromRegistryFast()
        {
            try
            {
                using (RegistryKey accountsKey = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\OneDrive\Accounts"))
                {
                    if (accountsKey != null)
                    {
                        string[] subKeys = accountsKey.GetSubKeyNames();
                        Array.Sort(subKeys, (a, b) =>
                        {
                            bool aBiz = a.StartsWith("Business", StringComparison.OrdinalIgnoreCase);
                            bool bBiz = b.StartsWith("Business", StringComparison.OrdinalIgnoreCase);
                            if (aBiz != bBiz) return aBiz ? -1 : 1;
                            if (string.Equals(a, "Personal", StringComparison.OrdinalIgnoreCase)) return 1;
                            if (string.Equals(b, "Personal", StringComparison.OrdinalIgnoreCase)) return -1;
                            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
                        });

                        foreach (string name in subKeys)
                        {
                            using (RegistryKey acct = accountsKey.OpenSubKey(name))
                            {
                                if (acct == null)
                                    continue;

                                object raw = acct.GetValue("UserFolder", null);
                                if (raw is string s && !string.IsNullOrWhiteSpace(s) && Directory.Exists(s))
                                    return s;
                            }
                        }
                    }
                }
            }
            catch
            {
            }

            try
            {
                string oneDrivePath = Registry.GetValue(@"HKEY_CURRENT_USER\Software\Microsoft\OneDrive\Accounts\Personal", "UserFolder", null) as string;
                if (!string.IsNullOrWhiteSpace(oneDrivePath) && Directory.Exists(oneDrivePath))
                    return oneDrivePath;
            }
            catch
            {
            }

            return string.Empty;
        }

        private static string GetOneDrivePathFromFolderFast(string usuario)
        {
            string userPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            if (string.IsNullOrWhiteSpace(userPath))
                userPath = $@"C:\Users\{usuario}";

            if (!Directory.Exists(userPath))
                return string.Empty;

            try
            {
                // Buscar solo la primera carpeta OneDrive comercial (con " - ")
                string[] directories = Directory.GetDirectories(userPath, "OneDrive*");
                
                foreach (string dir in directories)
                {
                    string folderName = Path.GetFileName(dir);
                    if (folderName != null && folderName.Contains(" - "))
                    {
                        return dir;
                    }
                }

                // Si no hay comercial, buscar personal
                foreach (string dir in directories)
                {
                    return dir;
                }
            }
            catch
            {
                // Ignorar errores (equivalente a On Error Resume Next de VBA)
            }

            return string.Empty;
        }

        private static string ExtractRelativePathFromUrlPath(string urlPath)
        {
            if (string.IsNullOrEmpty(urlPath))
                return string.Empty;

            string[] markers =
            {
                "/Documents/",
                "/Documentos/",
                "/Shared Documents/",
                "/Documentos compartidos/"
            };

            foreach (string marker in markers)
            {
                int idx = urlPath.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
                if (idx >= 0)
                    return urlPath.Substring(idx + marker.Length);
            }

            int lastSlash = urlPath.LastIndexOf("/", StringComparison.Ordinal);
            if (lastSlash >= 0 && lastSlash + 1 < urlPath.Length)
                return urlPath.Substring(lastSlash + 1);

            return string.Empty;
        }

        public static string GetCurrentOfficeUser()
        {
            string[] officeVersions = { "16.0", "15.0", "14.0" };

            foreach (string ver in officeVersions)
            {
                string email = TryGetOfficeEmailFromIdentity(ver);
                if (!string.IsNullOrWhiteSpace(email))
                    return email;
            }

            foreach (string ver in officeVersions)
            {
                string author = TryGetOfficeAuthorName(ver);
                if (!string.IsNullOrWhiteSpace(author))
                    return author;
            }

            return Environment.UserName;
        }

        private static string TryGetOfficeEmailFromIdentity(string officeVersion)
        {
            try
            {
                using (RegistryKey identitiesKey = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{officeVersion}\Common\Identity\Identities"))
                {
                    if (identitiesKey != null)
                    {
                        foreach (string subKeyName in identitiesKey.GetSubKeyNames())
                        {
                            using (RegistryKey identity = identitiesKey.OpenSubKey(subKeyName))
                            {
                                if (identity == null)
                                    continue;

                                string email = ReadFirstNonEmptyValue(identity, "EmailAddress", "SignInName", "UserEmail", "PrimaryEmailAddress", "SMTPAddress");
                                if (!string.IsNullOrWhiteSpace(email))
                                    return email;
                            }
                        }
                    }
                }

                using (RegistryKey signInKey = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{officeVersion}\Common\Identity\Signin"))
                {
                    if (signInKey != null)
                    {
                        string email = ReadFirstNonEmptyValue(signInKey, "EmailAddress", "SignInName", "UserEmail");
                        if (!string.IsNullOrWhiteSpace(email))
                            return email;
                    }
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        private static string TryGetOfficeAuthorName(string officeVersion)
        {
            try
            {
                using (RegistryKey userInfoKey = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{officeVersion}\Common\UserInfo"))
                {
                    if (userInfoKey == null)
                        return string.Empty;

                    string name = ReadFirstNonEmptyValue(userInfoKey, "UserName", "Name", "DisplayName");
                    if (!string.IsNullOrWhiteSpace(name))
                        return name;

                    string initials = ReadFirstNonEmptyValue(userInfoKey, "UserInitials", "Initials");
                    if (!string.IsNullOrWhiteSpace(initials))
                        return initials;
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }

        private static string ReadFirstNonEmptyValue(RegistryKey key, params string[] valueNames)
        {
            foreach (string valueName in valueNames)
            {
                object raw = key.GetValue(valueName, null);
                if (raw is string s && !string.IsNullOrWhiteSpace(s))
                    return s.Trim();
            }

            return string.Empty;
        }
    }
}
