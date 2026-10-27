using System;
using System.IO;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Web.Script.Serialization;
using ConversorPDF.XLL.Core.Helpers;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Serialización y deserialización del formato binario .crsync.
    ///
    /// FORMATO EN DISCO:
    ///   [0..N-1]  → JSON comprimido con GZip (payload principal)
    ///   Los primeros 64 bytes del JSON (sin comprimir) contienen el checksum SHA256
    ///   como campo "Checksum" dentro del propio CrSyncEnvelope.
    ///
    /// POR QUÉ JSON + GZIP:
    ///   - JSON: legible por humanos para debugging, versionable, ya soportado por
    ///     JavaScriptSerializer (System.Web.Extensions, sin NuGet adicional).
    ///   - GZip: reduce el tamaño 5-10x para datos textuales (columnas de texto).
    ///   - SHA256: detecta corrupción durante el sync parcial de OneDrive.
    ///
    /// SERIALIZER ELEGIDO:
    ///   JavaScriptSerializer (System.Web.Extensions), ya referenciado en el .csproj.
    ///   Limitaciones conocidas: no soporta propiedades privadas, necesita constructor
    ///   sin parámetros en los objetos serializados. Los POCOs en ChunkedChangeSet.cs
    ///   cumplen estos requisitos.
    ///
    /// INTEGRIDAD:
    ///   El checksum SHA256 se calcula sobre el JSON sin comprimir y se incrusta
    ///   en el campo Checksum del envelope ANTES de comprimir. Al deserializar,
    ///   se extrae el checksum, se limpia el campo, y se recalcula para comparar.
    /// </summary>
    internal sealed class CrSyncFileWriter
    {
        // ====================================================================
        // CONSTANTES
        // ====================================================================

        /// <summary>
        /// Encoding usado para JSON. UTF-8 sin BOM para compatibilidad máxima.
        /// </summary>
        private static readonly Encoding JsonEncoding = new UTF8Encoding(
            encoderShouldEmitUTF8Identifier: false);  // sin BOM

        /// <summary>
        /// Instancia compartida del serializer (thread-safe para operaciones Read-only).
        /// MaxJsonLength aumentado para manejar changesets grandes.
        /// </summary>
        private static readonly JavaScriptSerializer Serializer = new JavaScriptSerializer
        {
            MaxJsonLength = 50 * 1024 * 1024  // 50 MB máximo (para batches grandes)
        };

        // ====================================================================
        // SERIALIZACIÓN (Envelope → bytes en disco)
        // ====================================================================

        /// <summary>
        /// Serializa un CrSyncEnvelope a bytes listos para escribir al disco.
        ///
        /// PROCESO:
        ///   1. Serializar a JSON (sin checksum aún).
        ///   2. Calcular SHA256 del JSON.
        ///   3. Insertar checksum en el envelope y re-serializar.
        ///   4. Comprimir con GZip.
        ///
        /// El checksum se calcula sobre el JSON con Checksum = null para que sea
        /// determinístico (no cambia aunque se llame dos veces con el mismo contenido).
        /// </summary>
        /// <param name="envelope">Envelope con los cambios a serializar.</param>
        /// <returns>Bytes comprimidos listos para escribir al archivo .crsync.</returns>
        internal byte[] Serialize(CrSyncEnvelope envelope)
        {
            if (envelope == null)
            {
                throw new ArgumentNullException(nameof(envelope));
            }

            // Paso 1: Calcular checksum sobre el JSON sin campo Checksum
            // (para que el checksum sea determinístico independientemente del orden de campos)
            string previousChecksum = envelope.Checksum;
            envelope.Checksum = null;

            string jsonForChecksum = Serializer.Serialize(envelope);
            string checksum = ComputeChecksum(jsonForChecksum);

            // Paso 2: Insertar checksum y serializar el JSON final
            envelope.Checksum = checksum;
            string finalJson = Serializer.Serialize(envelope);

            // Restaurar estado del envelope (por si el llamador lo reutiliza)
            envelope.Checksum = previousChecksum ?? checksum;

            // Paso 3: Comprimir con GZip
            byte[] jsonBytes = JsonEncoding.GetBytes(finalJson);
            return CompressGzip(jsonBytes);
        }

        /// <summary>
        /// Deserializa bytes de un archivo .crsync a un CrSyncEnvelope.
        ///
        /// PROCESO:
        ///   1. Descomprimir GZip → JSON bytes.
        ///   2. Deserializar JSON → CrSyncEnvelope.
        ///   3. Extraer y verificar checksum SHA256.
        ///
        /// LANZA:
        ///   - InvalidDataException: si el checksum no coincide (datos corruptos).
        ///   - FormatException: si el JSON es inválido.
        /// </summary>
        /// <param name="compressedBytes">Bytes leídos del archivo .crsync.</param>
        /// <returns>Envelope deserializado y verificado.</returns>
        internal CrSyncEnvelope Deserialize(byte[] compressedBytes)
        {
            if (compressedBytes == null || compressedBytes.Length == 0)
            {
                throw new ArgumentException("Los bytes del archivo .crsync no pueden ser nulos o vacíos.", nameof(compressedBytes));
            }

            // Paso 1: Descomprimir
            byte[] jsonBytes = DecompressGzip(compressedBytes);
            string json = JsonEncoding.GetString(jsonBytes);

            // Paso 2: Deserializar
            CrSyncEnvelope envelope;
            try
            {
                envelope = Serializer.Deserialize<CrSyncEnvelope>(json);
            }
            catch (Exception ex)
            {
                throw new FormatException($"[CrSyncFileWriter] El archivo .crsync contiene JSON inválido: {ex.Message}", ex);
            }

            if (envelope == null)
            {
                throw new FormatException("[CrSyncFileWriter] El archivo .crsync está vacío o es inválido.");
            }

            // Paso 3: Verificar checksum
            string storedChecksum = envelope.Checksum;
            if (!string.IsNullOrEmpty(storedChecksum))
            {
                // Recalcular sobre el JSON sin el campo Checksum (mismo método que Serialize)
                envelope.Checksum = null;
                string jsonForVerification = Serializer.Serialize(envelope);
                string computedChecksum = ComputeChecksum(jsonForVerification);
                envelope.Checksum = storedChecksum;

                if (!string.Equals(storedChecksum, computedChecksum, StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        $"[CrSyncFileWriter] Checksum inválido. " +
                        $"Esperado: {storedChecksum}, " +
                        $"Calculado: {computedChecksum}. " +
                        $"El archivo puede estar corrupto o fue modificado en tránsito.");
                }
            }
            else
            {
                LoggerHelper.Warning("[CrSyncFileWriter] Archivo .crsync sin checksum. Se omite verificación de integridad.");
            }

            return envelope;
        }

        // ====================================================================
        // OPERACIONES DE DISCO
        // ====================================================================

        /// <summary>
        /// Escribe el envelope serializado a disco como archivo .crsync.
        ///
        /// USA escritura atómica (archivo temporal + rename) para prevenir que
        /// el FileSystemWatcher de otra sede lea un archivo parcialmente escrito.
        /// </summary>
        /// <param name="envelope">Envelope a escribir.</param>
        /// <param name="filePath">Ruta completa del archivo de destino.</param>
        internal void WriteToFile(CrSyncEnvelope envelope, string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentNullException(nameof(filePath));
            }

            byte[] data = Serialize(envelope);

            // Escritura atómica: primero escribir a .tmp, luego renombrar
            // Esto evita que el FileSystemWatcher lea un archivo incompleto
            string tempPath = filePath + ".tmp";
            string directory = Path.GetDirectoryName(filePath);

            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
                LoggerHelper.Debug($"[CrSyncFileWriter] Carpeta creada: {directory}");
            }

            File.WriteAllBytes(tempPath, data);
            File.Move(tempPath, filePath);

            LoggerHelper.Info($"[CrSyncFileWriter] Archivo escrito: {Path.GetFileName(filePath)} ({data.Length:N0} bytes comprimidos)");
        }

        /// <summary>
        /// Lee y deserializa un archivo .crsync desde disco.
        /// Incluye verificación de integridad mediante checksum SHA256.
        /// </summary>
        /// <param name="filePath">Ruta completa del archivo .crsync a leer.</param>
        internal CrSyncEnvelope ReadFromFile(string filePath)
        {
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException($"Archivo .crsync no encontrado: {filePath}");
            }

            byte[] data = File.ReadAllBytes(filePath);
            return Deserialize(data);
        }

        // ====================================================================
        // UTILIDADES (internas al módulo)
        // ====================================================================

        /// <summary>
        /// Calcula el SHA256 de una cadena JSON y retorna el hash en formato hexadecimal lowercase.
        ///
        /// EJEMPLO: "d2a8..." (64 caracteres hex)
        ///
        /// THREAD SAFETY: SHA256.Create() crea una nueva instancia cada vez.
        /// No compartir instancias de SHA256 entre hilos.
        /// </summary>
        internal static string ComputeChecksum(string json)
        {
            byte[] inputBytes = JsonEncoding.GetBytes(json);
            using (SHA256 sha256 = SHA256.Create())
            {
                byte[] hashBytes = sha256.ComputeHash(inputBytes);
                return BitConverter.ToString(hashBytes)
                    .Replace("-", string.Empty)
                    .ToLowerInvariant();
            }
        }

        /// <summary>
        /// Comprime un array de bytes usando GZip (nivel óptimo).
        /// Para datos JSON típicos, reduce el tamaño 5-10x.
        /// </summary>
        private static byte[] CompressGzip(byte[] data)
        {
            using (MemoryStream output = new MemoryStream())
            {
                // CompressionLevel.Optimal: máxima compresión, más lento.
                // Para archivos .crsync (operación background), es aceptable.
                using (GZipStream gzip = new GZipStream(output, CompressionLevel.Optimal))
                {
                    gzip.Write(data, 0, data.Length);
                }
                return output.ToArray();
            }
        }

        /// <summary>
        /// Descomprime un array de bytes comprimidos con GZip.
        /// LANZA: InvalidDataException si los bytes no son GZip válido.
        /// </summary>
        private static byte[] DecompressGzip(byte[] compressedData)
        {
            using (MemoryStream input = new MemoryStream(compressedData))
            using (GZipStream gzip = new GZipStream(input, CompressionMode.Decompress))
            using (MemoryStream output = new MemoryStream())
            {
                gzip.CopyTo(output);
                return output.ToArray();
            }
        }
    }
}
