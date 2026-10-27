using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Core.Sync
{
    /// <summary>
    /// PROPÓSITO:
    ///   Modelo de datos que representa el contenido de un archivo .crsync.
    ///   Un archivo .crsync es la unidad de transferencia entre una subunidad (sede)
    ///   y el directorio compartido de OneDrive.
    ///
    /// FORMATO EN DISCO:
    ///   GZip( JSON( CrSyncEnvelope ) ) + SHA256(32 bytes hex) al final del archivo.
    ///   Esto permite verificar integridad sin descomprimir todo el payload.
    ///
    /// COMPATIBILIDAD:
    ///   El campo Version permite versionar el formato. Si en el futuro se cambia
    ///   la estructura, se puede leer ambos formatos en paralelo durante la migración.
    /// </summary>
    internal sealed class CrSyncEnvelope
    {
        /// <summary>
        /// Versión del formato del archivo .crsync.
        /// Valor actual: "1.0". Nunca omitir — es la clave para compatibilidad futura.
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// Metadatos del origen del archivo: quién lo generó, cuándo y desde qué subunidad.
        /// Permite filtrar archivos propios (no procesar lo que yo mismo generé).
        /// </summary>
        public CrSyncSource Source { get; set; }

        /// <summary>
        /// Lista de cambios CRDT extraídos de la tabla crsql_changes.
        /// Cada elemento representa un cambio a nivel de celda (tabla + fila + columna).
        /// </summary>
        public List<CrSyncChange> Changes { get; set; }

        /// <summary>
        /// SHA256 hex del JSON sin comprimir.
        /// Se calcula ANTES de la compresión y se verifica DESPUÉS de la descompresión.
        /// Permite detectar corrupción de datos en tránsito (OneDrive sync parcial).
        /// </summary>
        public string Checksum { get; set; }

        /// <summary>
        /// Constructor sin parámetros requerido por JavaScriptSerializer.
        /// </summary>
        public CrSyncEnvelope()
        {
            Changes = new List<CrSyncChange>();
        }
    }

    /// <summary>
    /// Metadatos del origen de un archivo .crsync.
    /// Identifica inequívocamente el PC y subunidad que generó los cambios.
    /// </summary>
    internal sealed class CrSyncSource
    {
        /// <summary>
        /// Nombre del PC que generó los cambios (Environment.MachineName).
        /// Ej: "PC-DIRECCION-VALLE", "LAPTOP-JURIDICA-01"
        /// Se usa para evitar que una subunidad aplique sus propios cambios (loop).
        /// </summary>
        public string DeviceId { get; set; }

        /// <summary>
        /// Código de la subunidad que generó los cambios.
        /// Ej: "VALLE", "ANTIOQUIA", "CENTRAL"
        /// Normalizado a mayúsculas sin tildes.
        /// </summary>
        public string Subunidad { get; set; }

        /// <summary>
        /// Timestamp UTC del momento en que se generó el archivo .crsync.
        /// Formato ISO 8601: "2026-06-13T22:00:00Z"
        /// NO es el timestamp de los cambios individuales (eso está en col_version).
        /// </summary>
        public string TimestampUtc { get; set; }
    }

    /// <summary>
    /// Representa un cambio individual en la tabla virtual crsql_changes de CR-SQLite.
    ///
    /// CONTEXTO CR-SQLite:
    ///   CR-SQLite trackea cambios a nivel de celda (tabla + fila + columna).
    ///   Cada fila de crsql_changes corresponde a una modificación de UNA columna
    ///   en UNA fila de UNA tabla. Esto permite resolución de conflictos LWW (Last-Writer-Wins)
    ///   a nivel de columna, no a nivel de fila completa.
    ///
    /// CAMPOS CRDT:
    ///   - col_version: Versión lógica de esta columna (incrementa con cada UPDATE).
    ///   - db_version: Versión global de la DB cuando ocurrió el cambio.
    ///   - site_id: ID único del dispositivo que generó el cambio (asignado por CR-SQLite).
    ///   - cl: Causal Length (longitud de la cadena causal, para CRDTs avanzados).
    /// </summary>
    internal sealed class CrSyncChange
    {
        /// <summary>
        /// Nombre de la tabla donde ocurrió el cambio.
        /// Ej: "terceros", "vinculaciones", "sesiones"
        /// </summary>
        public string Table { get; set; }

        /// <summary>
        /// Clave primaria de la fila modificada, serializada como string.
        /// Para PKs compuestas, CR-SQLite las codifica internamente.
        /// </summary>
        public string Pk { get; set; }

        /// <summary>
        /// Nombre de la columna que cambió (cid = Column ID).
        /// Ej: "nombre", "subunidad", "activo"
        /// </summary>
        public string Cid { get; set; }

        /// <summary>
        /// Nuevo valor de la columna, serializado como string.
        /// null indica que la fila fue eliminada (DELETE).
        /// Para booleanos/enteros, el string representa el valor numérico ("0", "1").
        /// </summary>
        public string Val { get; set; }

        /// <summary>
        /// Versión lógica de esta columna en este dispositivo.
        /// Mayor = más reciente. En LWW, gana el mayor col_version.
        /// </summary>
        public long ColVersion { get; set; }

        /// <summary>
        /// Versión global de la base de datos al momento del cambio.
        /// Secuencial y monotónicamente creciente en cada dispositivo.
        /// </summary>
        public long DbVersion { get; set; }

        /// <summary>
        /// Identificador único del sitio (dispositivo) que generó el cambio.
        /// Asignado automáticamente por CR-SQLite. Sirve para desempatar conflictos
        /// cuando col_version es igual en dos dispositivos (caso raro pero posible).
        /// </summary>
        public long SiteId { get; set; }

        /// <summary>
        /// Causal Length: longitud de la cadena causal del CRDT.
        /// 0 = INSERT, 1 = UPDATE, -1 = DELETE (en la semántica de CR-SQLite).
        /// </summary>
        public long Cl { get; set; }
    }
}
