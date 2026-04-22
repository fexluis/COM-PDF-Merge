using System;
using System.IO;
using System.Security;
using System.Security.Cryptography;
using System.Runtime.InteropServices;

namespace ConversorPDF.XLL.Core.Helpers
{
    /// <summary>
    /// Motor de encriptación a nivel de archivo para SQLite (PBKDF2 + AES-256-CBC + HMAC-SHA256).
    /// </summary>
    public static class DatabaseEncryptionEngine
    {
        private const int KeySize = 256;
        private const int Multiplier = 8;
        private const int KeyByteSize = KeySize / Multiplier;
        private const int BlockSize = 128;
        private const int BlockByteSize = BlockSize / Multiplier;
        private const int MacByteSize = 32; // HMAC-SHA256
        private const int SaltSize = 16;
        private const int Pbkdf2Iterations = 100000;

        public static byte[] Encrypt(byte[] plaintext, SecureString password)
        {
            if (plaintext == null || plaintext.Length == 0) return plaintext;
            if (password == null || password.Length == 0) throw new ArgumentNullException(nameof(password));

            byte[] salt = new byte[SaltSize];
            using (var rng = new RNGCryptoServiceProvider())
            {
                rng.GetBytes(salt);
            }

            byte[] pwdBytes = SecureStringToBytes(password);
            try
            {
                using (var deriveBytes = new Rfc2898DeriveBytes(pwdBytes, salt, Pbkdf2Iterations))
                {
                    byte[] keyInfo = deriveBytes.GetBytes(KeyByteSize + MacByteSize);
                    byte[] aesKey = new byte[KeyByteSize];
                    byte[] hmacKey = new byte[MacByteSize];
                    
                    Buffer.BlockCopy(keyInfo, 0, aesKey, 0, KeyByteSize);
                    Buffer.BlockCopy(keyInfo, KeyByteSize, hmacKey, 0, MacByteSize);

                    byte[] ciphertext;
                    byte[] iv;

                    using (var aes = new AesCryptoServiceProvider())
                    {
                        aes.KeySize = KeySize;
                        aes.BlockSize = BlockSize;
                        aes.Mode = CipherMode.CBC;
                        aes.Padding = PaddingMode.PKCS7;
                        aes.Key = aesKey;
                        aes.GenerateIV();
                        iv = aes.IV;

                        using (var encryptor = aes.CreateEncryptor(aes.Key, aes.IV))
                        using (var msTest = new MemoryStream())
                        {
                            using (var cs = new CryptoStream(msTest, encryptor, CryptoStreamMode.Write))
                            {
                                cs.Write(plaintext, 0, plaintext.Length);
                            }
                            ciphertext = msTest.ToArray();
                        }
                    }

                    // Format: Salt (16) + IV (16) + Ciphertext + MAC (32)
                    byte[] result = new byte[SaltSize + BlockByteSize + ciphertext.Length + MacByteSize];
                    Buffer.BlockCopy(salt, 0, result, 0, SaltSize);
                    Buffer.BlockCopy(iv, 0, result, SaltSize, BlockByteSize);
                    Buffer.BlockCopy(ciphertext, 0, result, SaltSize + BlockByteSize, ciphertext.Length);

                    // Compute MAC
                    using (var hmac = new HMACSHA256(hmacKey))
                    {
                        byte[] payloadToMac = new byte[SaltSize + BlockByteSize + ciphertext.Length];
                        Buffer.BlockCopy(result, 0, payloadToMac, 0, payloadToMac.Length);
                        byte[] mac = hmac.ComputeHash(payloadToMac);
                        Buffer.BlockCopy(mac, 0, result, payloadToMac.Length, MacByteSize);
                    }

                    return result;
                }
            }
            finally
            {
                Array.Clear(pwdBytes, 0, pwdBytes.Length);
            }
        }

        public static byte[] Decrypt(byte[] encryptedBlob, SecureString password)
        {
            if (encryptedBlob == null || encryptedBlob.Length == 0) return encryptedBlob;
            if (password == null || password.Length == 0) throw new ArgumentNullException(nameof(password));

            int minLength = SaltSize + BlockByteSize + MacByteSize;
            if (encryptedBlob.Length < minLength)
                throw new CryptographicException("Mensaje encriptado inválido (muy corto).");

            byte[] pwdBytes = SecureStringToBytes(password);
            try
            {
                byte[] salt = new byte[SaltSize];
                Buffer.BlockCopy(encryptedBlob, 0, salt, 0, SaltSize);

                using (var deriveBytes = new Rfc2898DeriveBytes(pwdBytes, salt, Pbkdf2Iterations))
                {
                    byte[] keyInfo = deriveBytes.GetBytes(KeyByteSize + MacByteSize);
                    byte[] aesKey = new byte[KeyByteSize];
                    byte[] hmacKey = new byte[MacByteSize];

                    Buffer.BlockCopy(keyInfo, 0, aesKey, 0, KeyByteSize);
                    Buffer.BlockCopy(keyInfo, KeyByteSize, hmacKey, 0, MacByteSize);

                    // Verify MAC
                    int payloadLen = encryptedBlob.Length - MacByteSize;
                    byte[] payloadToMac = new byte[payloadLen];
                    Buffer.BlockCopy(encryptedBlob, 0, payloadToMac, 0, payloadLen);

                    byte[] expectedMac = new byte[MacByteSize];
                    Buffer.BlockCopy(encryptedBlob, payloadLen, expectedMac, 0, MacByteSize);

                    using (var hmac = new HMACSHA256(hmacKey))
                    {
                        byte[] calculatedMac = hmac.ComputeHash(payloadToMac);
                        if (!CryptographicEquals(expectedMac, calculatedMac))
                        {
                            throw new CryptographicException("La autenticación del archivo falló. Contraseña incorrecta o archivo corrupto.");
                        }
                    }

                    byte[] iv = new byte[BlockByteSize];
                    Buffer.BlockCopy(encryptedBlob, SaltSize, iv, 0, BlockByteSize);

                    int cipherLen = payloadLen - SaltSize - BlockByteSize;
                    byte[] ciphertext = new byte[cipherLen];
                    Buffer.BlockCopy(encryptedBlob, SaltSize + BlockByteSize, ciphertext, 0, cipherLen);

                    using (var aes = new AesCryptoServiceProvider())
                    {
                        aes.KeySize = KeySize;
                        aes.BlockSize = BlockSize;
                        aes.Mode = CipherMode.CBC;
                        aes.Padding = PaddingMode.PKCS7;
                        aes.Key = aesKey;
                        aes.IV = iv;

                        using (var decryptor = aes.CreateDecryptor(aes.Key, aes.IV))
                        using (var ms = new MemoryStream(ciphertext))
                        using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                        using (var msDecrypted = new MemoryStream())
                        {
                            cs.CopyTo(msDecrypted);
                            return msDecrypted.ToArray();
                        }
                    }
                }
            }
            finally
            {
                Array.Clear(pwdBytes, 0, pwdBytes.Length);
            }
        }

        private static bool CryptographicEquals(byte[] a, byte[] b)
        {
            if (a.Length != b.Length) return false;
            int result = 0;
            for (int i = 0; i < a.Length; i++)
            {
                result |= a[i] ^ b[i];
            }
            return result == 0;
        }

        private static byte[] SecureStringToBytes(SecureString secureString)
        {
            IntPtr unmanagedBytes = IntPtr.Zero;
            try
            {
                unmanagedBytes = Marshal.SecureStringToGlobalAllocAnsi(secureString);
                byte[] b = new byte[secureString.Length];
                Marshal.Copy(unmanagedBytes, b, 0, b.Length);
                return b;
            }
            finally
            {
                if (unmanagedBytes != IntPtr.Zero)
                    Marshal.ZeroFreeGlobalAllocAnsi(unmanagedBytes);
            }
        }

        public static void SecureWipe(string filePath, int passes = 3)
        {
            if (!File.Exists(filePath)) return;

            try
            {
                FileInfo fileInfo = new FileInfo(filePath);
                long length = fileInfo.Length;
                
                // Ensure not read-only
                fileInfo.Attributes = FileAttributes.Normal;

                byte[] dummyBuffer = new byte[4096];
                using (var rng = new RNGCryptoServiceProvider())
                {
                    using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Write, FileShare.None))
                    {
                        for (int pass = 0; pass < passes; pass++)
                        {
                            fs.Position = 0;
                            // Pass 1: Zeros, Pass 2: Ones, Pass 3: Random
                            byte fillByte = (byte)(pass == 0 ? 0x00 : (pass == 1 ? 0xFF : 0));
                            
                            for (int i = 0; i < dummyBuffer.Length; i++) dummyBuffer[i] = fillByte;

                            long sectors = length / dummyBuffer.Length;
                            for (int i = 0; i < sectors; i++)
                            {
                                if (pass == 2) rng.GetBytes(dummyBuffer);
                                fs.Write(dummyBuffer, 0, dummyBuffer.Length);
                            }

                            int remainder = (int)(length % dummyBuffer.Length);
                            if (remainder > 0)
                            {
                                if (pass == 2) rng.GetBytes(dummyBuffer);
                                fs.Write(dummyBuffer, 0, remainder);
                            }
                            fs.Flush(true);
                        }
                        
                        // Resizing
                        fs.SetLength(0);
                    }
                }
                
                // Scramble date times
                File.SetCreationTime(filePath, new DateTime(1980, 1, 1));
                File.SetLastAccessTime(filePath, new DateTime(1980, 1, 1));
                File.SetLastWriteTime(filePath, new DateTime(1980, 1, 1));

                // Scramble name and delete
                string directory = Path.GetDirectoryName(filePath);
                string newName = Path.Combine(directory, Guid.NewGuid().ToString() + ".tmp");
                File.Move(filePath, newName);
                File.Delete(newName);
            }
            catch (Exception ex)
            {
                // Fallback delete
                if (File.Exists(filePath))
                {
                    try { File.Delete(filePath); } catch { /* ignore fallback error */ }
                }
            }
        }
    }
}
