using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Resultado detallado del proceso "Buscar Informes": renombrado, unificación de
    /// parejas INF + CUM y auditoría de la columna INFORME de la hoja CONSULTAS.
    /// </summary>
    public class ResultadoBusquedaInformes
    {
        /// <summary> Inicializa contadores, carpetas y detalle. </summary>
        public ResultadoBusquedaInformes()
        {
            CarpetasProcesadas = new List<string>();
            Detalle = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }

        /// <summary> Carpeta raíz procesada (&lt;libro&gt;\Informes). </summary>
        public string CarpetaRaiz { get; set; }

        /// <summary> Carpetas realmente escaneadas (excluye VERIFICADOS y _Duplicados). </summary>
        public List<string> CarpetasProcesadas { get; set; }

        /// <summary> Mensaje de error de infraestructura; vacío si la ejecución fue normal. </summary>
        public string Error { get; set; }

        /// <summary> True cuando la ejecución terminó por error de infraestructura. </summary>
        public bool HayError
        {
            get { return !string.IsNullOrEmpty(Error); }
        }

        /// <summary> True si la ejecución se hizo en modo simulación (sin tocar disco ni Excel). </summary>
        public bool Simulado { get; set; }

        /// <summary> PDF examinados. </summary>
        public int PdfsExaminados { get; set; }

        /// <summary> PDF renombrados con éxito. </summary>
        public int Renombrados { get; set; }

        /// <summary> PDF que ya tenían el nombre estandarizado correcto. </summary>
        public int YaCorrectos { get; set; }

        /// <summary> Parejas INF + CUM unificadas en un archivo UNI. </summary>
        public int Unificados { get; set; }

        /// <summary> Unificaciones que fallaron (se conservaron los originales). </summary>
        public int MergeFallidos { get; set; }

        /// <summary> Duplicados exactos archivados en _Duplicados. </summary>
        public int Duplicados { get; set; }

        /// <summary> Colisiones resueltas con sufijo _1, _2… </summary>
        public int Colisiones { get; set; }

        /// <summary> Archivos INF o CUM que quedaron sin pareja. </summary>
        public int RechazadosSinPareja { get; set; }

        /// <summary> Errores de entrada/salida o de validación de archivos. </summary>
        public int Errores { get; set; }

        /// <summary> PDF sin capa de texto (escaneados) o ilegibles. </summary>
        public int SinTexto { get; set; }

        /// <summary> Archivos cuya solicitud no figura en CONSULTAS. </summary>
        public int SinComision { get; set; }

        /// <summary> Archivos cuya solicitud figura en CONSULTAS pero sin cédula o nombre. </summary>
        public int SinDatos { get; set; }

        /// <summary> Archivos cuyo nombre no inicia con un número de solicitud válido. </summary>
        public int PrefijoInvalido { get; set; }

        /// <summary> Filas de CONSULTAS evaluadas en la auditoría. </summary>
        public int FilasRevisadas { get; set; }

        /// <summary> Celdas de la columna INFORME escritas. </summary>
        public int CeldasEscritas { get; set; }

        /// <summary> Celdas de INFORME limpiadas por contener una marca inválida. </summary>
        public int CeldasLimpiadas { get; set; }

        /// <summary> Filas protegidas (VERIFICADO, REINTEGRO, REVOCADA, RECHAZADA, ANULADA). </summary>
        public int OmitidasProtegidas { get; set; }

        /// <summary> Filas marcadas como VERIFICADO. </summary>
        public int Verificados { get; set; }

        /// <summary> Filas marcadas como FALTA CUM. </summary>
        public int FaltaCum { get; set; }

        /// <summary> Filas marcadas como FALTA INF. </summary>
        public int FaltaInf { get; set; }

        /// <summary> Filas marcadas como ERROR. </summary>
        public int ConError { get; set; }

        /// <summary> Filas marcadas como SIN INFORME. </summary>
        public int SinInforme { get; set; }

        /// <summary> Parejas completas detectadas en la auditoría (no se reescriben). </summary>
        public int ParOk { get; set; }

        /// <summary> True si el libro se guardó al finalizar. </summary>
        public bool LibroGuardado { get; set; }

        /// <summary> Mensaje de error si el guardado falló. </summary>
        public string ErrorGuardado { get; set; }

        /// <summary> Duración total en milisegundos. </summary>
        public long TiempoMs { get; set; }

        /// <summary> Nombre del archivo (o fila) → motivo / estado. </summary>
        public Dictionary<string, string> Detalle { get; set; }

        /// <summary>
        /// Suma a esta instancia los contadores y el detalle de otro resultado.
        /// Se usa para combinar la fase de renombrado con la de auditoría.
        /// </summary>
        /// <param name="otro">Resultado a absorber; se ignora si es null.</param>
        public void Absorber(ResultadoBusquedaInformes otro)
        {
            if (otro == null) return;

            foreach (var carpeta in otro.CarpetasProcesadas)
            {
                if (!CarpetasProcesadas.Contains(carpeta)) CarpetasProcesadas.Add(carpeta);
            }

            foreach (var par in otro.Detalle)
            {
                Detalle[par.Key] = par.Value;
            }

            PdfsExaminados += otro.PdfsExaminados;
            Renombrados += otro.Renombrados;
            YaCorrectos += otro.YaCorrectos;
            Unificados += otro.Unificados;
            MergeFallidos += otro.MergeFallidos;
            Duplicados += otro.Duplicados;
            Colisiones += otro.Colisiones;
            RechazadosSinPareja += otro.RechazadosSinPareja;
            Errores += otro.Errores;
            SinTexto += otro.SinTexto;
            SinComision += otro.SinComision;
            SinDatos += otro.SinDatos;
            PrefijoInvalido += otro.PrefijoInvalido;
            FilasRevisadas += otro.FilasRevisadas;
            CeldasEscritas += otro.CeldasEscritas;
            CeldasLimpiadas += otro.CeldasLimpiadas;
            OmitidasProtegidas += otro.OmitidasProtegidas;
            Verificados += otro.Verificados;
            FaltaCum += otro.FaltaCum;
            FaltaInf += otro.FaltaInf;
            ConError += otro.ConError;
            SinInforme += otro.SinInforme;
            ParOk += otro.ParOk;
            TiempoMs += otro.TiempoMs;

            if (!string.IsNullOrEmpty(otro.Error)) Error = otro.Error;
            LibroGuardado = LibroGuardado || otro.LibroGuardado;
            if (string.IsNullOrEmpty(ErrorGuardado)) ErrorGuardado = otro.ErrorGuardado;
            Simulado = Simulado || otro.Simulado;
        }
    }
}
