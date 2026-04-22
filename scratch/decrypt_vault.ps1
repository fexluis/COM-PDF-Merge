$vaultPath = "$env:LOCALAPPDATA\ConversorPDF\Data\AcademicData.db.enc"
$password = "Sqlite@ESAP.2026"
# Usar ruta relativa para evitar problemas de codificación en el path absoluto
$outputPath = Join-Path $PSScriptRoot "decrypted.db"

$source = @"
using System;
using System.IO;
using System.Security;
using System.Security.Cryptography;

public class Decryptor {
    public static void DecryptFile(string inputPath, string outputPath, string password) {
        byte[] bytes = File.ReadAllBytes(inputPath);
        
        byte[] salt = new byte[16];
        byte[] iv = new byte[16];
        Buffer.BlockCopy(bytes, 0, salt, 0, 16);
        Buffer.BlockCopy(bytes, 16, iv, 0, 16);
        
        int cipherLen = bytes.Length - 16 - 16 - 32;
        byte[] ciphertext = new byte[cipherLen];
        Buffer.BlockCopy(bytes, 32, ciphertext, 0, cipherLen);
        
        using (var deriveBytes = new Rfc2898DeriveBytes(password, salt, 100000)) {
            byte[] keyInfo = deriveBytes.GetBytes(64);
            byte[] aesKey = new byte[32];
            Buffer.BlockCopy(keyInfo, 0, aesKey, 0, 32);
            
            using (var aes = Aes.Create()) {
                aes.Key = aesKey;
                aes.IV = iv;
                aes.Mode = CipherMode.CBC;
                aes.Padding = PaddingMode.PKCS7;
                
                using (var decryptor = aes.CreateDecryptor())
                using (var ms = new MemoryStream(ciphertext))
                using (var cs = new CryptoStream(ms, decryptor, CryptoStreamMode.Read))
                using (var outFile = new FileStream(outputPath, FileMode.Create, FileAccess.Write)) {
                    cs.CopyTo(outFile);
                }
            }
        }
    }
}
"@

# Remover tipo si ya existe (PowerShell session persistentce)
try {
    Add-Type -TypeDefinition $source -ErrorAction SilentlyContinue
} catch {}

[Decryptor]::DecryptFile($vaultPath, $outputPath, $password)
Write-Host "Decrypted to $outputPath"
