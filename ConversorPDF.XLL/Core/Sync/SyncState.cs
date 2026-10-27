using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// Modelo de datos para el estado de sincronización de un dispositivo (Phase 9.5).
    /// Se escribe en un archivo state_{DeviceId}.json en la carpeta csync/{subunidad}.
    /// </summary>
    public class SyncState
    {
        public string DeviceId { get; set; }
        
        /// <summary>
        /// El nombre del último archivo .crsync generado por este dispositivo tras un PUSH exitoso.
        /// </summary>
        public string LastGeneratedFile { get; set; }
        
        /// <summary>
        /// Diccionario que mapea SenderDeviceId -> Último nombre de archivo .crsync procesado exitosamente por este dispositivo.
        /// </summary>
        public Dictionary<string, string> LastProcessedFiles { get; set; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        
        public DateTime LastUpdatedAt { get; set; }
        public string AppVersion { get; set; }
    }
}
