using System;

namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Resultado del proceso de verificación de presencia física de solicitudes
    /// sobre la hoja CONSULTAS (columna INFORME).
    /// </summary>
    public class ResultadoVerificacionSolicitudes
    {
        /// <summary> Mensaje de error de infraestructura; vacío si la ejecución fue normal. </summary>
        public string Error { get; set; }

        /// <summary> True cuando la ejecución terminó por error de infraestructura. </summary>
        public bool HayError { get { return !string.IsNullOrEmpty(Error); } }

        /// <summary> Filas de CONSULTAS evaluadas dentro del rango. </summary>
        public int FilasRevisadas { get; set; }

        /// <summary> Filas protegidas por INFORME = VERIFICADO o REINTEGRO (sin acceso a disco). </summary>
        public int OmitidasProtegidas { get; set; }

        /// <summary> Filas con INFORME no vacío distinto de las protegidas (se preservan). </summary>
        public int OmitidasOtrasMarcas { get; set; }

        /// <summary> Filas cuyo estado terminal se copió a INFORME. </summary>
        public int CopiadasEstado { get; set; }

        /// <summary> Filas que ya tenían el PDF correspondiente y marca válida. </summary>
        public int ConArchivo { get; set; }

        /// <summary> Filas sin PDF correspondiente. </summary>
        public int SinArchivo { get; set; }

        /// <summary> Filas que ya contenían el valor destino (idempotencia). </summary>
        public int YaMarcadas { get; set; }

        /// <summary> Celdas de INFORME escritas. </summary>
        public int CeldasEscritas { get; set; }

        /// <summary> Celdas de INFORME limpiadas por contener una marca errónea. </summary>
        public int CeldasLimpiadas { get; set; }

        /// <summary> True si el libro se guardó al finalizar. </summary>
        public bool LibroGuardado { get; set; }

        /// <summary> Mensaje de error si el guardado falló. </summary>
        public string ErrorGuardado { get; set; }
    }
}
