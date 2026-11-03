using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Resultado detallado del renombrado de solicitudes de viáticos.
    /// Permite distinguir "sin comisiones en el rango" de "comisiones sin PDF coincidentes".
    /// </summary>
    public class ResultadoRenombrado
    {
        /// <summary> Carpeta raíz recibida (carpeta del libro). </summary>
        public string CarpetaRaiz { get; set; }

        /// <summary> Carpetas realmente escaneadas (raíz y, si existe, Solicitudes). </summary>
        public List<string> CarpetasProcesadas { get; set; }

        /// <summary> Mensaje de error de infraestructura; vacío si la ejecución fue normal. </summary>
        public string Error { get; set; }

        /// <summary> True cuando la ejecución terminó por error de infraestructura. </summary>
        public bool HayError { get { return !string.IsNullOrEmpty(Error); } }

        /// <summary> Comisiones de CONSULTAS dentro del rango solicitado. </summary>
        public int ComisionesEnRango { get; set; }

        /// <summary> PDF examinados en todas las carpetas procesadas. </summary>
        public int PdfsExaminados { get; set; }

        /// <summary> Cantidad de PDFs renombrados exitosamente. </summary>
        public int Renombrados { get; set; }

        /// <summary> Cantidad de PDFs que ya poseían el nombre estandarizado correcto. </summary>
        public int YaCorrectos { get; set; }

        /// <summary> Cantidad de colisiones (el archivo destino ya existía con contenido distinto). </summary>
        public int Colisiones { get; set; }

        /// <summary> Cantidad de duplicados exactos movidos a la subcarpeta _Duplicados. </summary>
        public int Duplicados { get; set; }

        /// <summary> Cantidad de archivos ambiguos sin resolución única. </summary>
        public int Ambiguos { get; set; }

        /// <summary> Cantidad de archivos cuyo número no figura en CONSULTAS. </summary>
        public int SinComision { get; set; }

        /// <summary> Cantidad de archivos cuya comisión existe en CONSULTAS pero carece de datos completos. </summary>
        public int SinDatos { get; set; }

        /// <summary> Cantidad de archivos con nombre que no inicia con prefijo numérico. </summary>
        public int PrefijoInvalido { get; set; }

        /// <summary> Cantidad de errores de I/O al intentar renombrar. </summary>
        public int ErroresIo { get; set; }

        /// <summary> Nombre original del archivo → motivo (ver <see cref="MotivoRenombrado"/>). </summary>
        public Dictionary<string, string> Detalle { get; set; }

        /// <summary>
        /// Inicializa una nueva instancia de <see cref="ResultadoRenombrado"/>.
        /// </summary>
        public ResultadoRenombrado()
        {
            CarpetasProcesadas = new List<string>();
            Detalle = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }
}

