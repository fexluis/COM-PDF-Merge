using System;
using ConversorPDF.XLL.Core.Enums;

namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Estructura de datos para los metadatos de un documento PDF.
    /// </summary>
    public class Metadata
    {
        /// <summary> Título principal del documento. </summary>
        public string Titulo { get; set; }
        
        /// <summary> Autor o usuario que generó el documento. </summary>
        public string Autor { get; set; }
        
        /// <summary> Asunto o descripción breve del contenido (ej. Objeto de comisión). </summary>
        public string Asunto { get; set; }
        
        /// <summary> Palabras clave para búsqueda, separadas por comas. </summary>
        public string PalabrasClave { get; set; }
        
        /// <summary> Fecha de relevancia para el documento (ej. Fecha de inicio de viaje). </summary>
        public DateTime? Fecha { get; set; }
        
        /// <summary> Tipo de documento según la enumeración de negocio (VIA-SOLICITUD, etc.). </summary>
        public TipoMetadata? MetaDocumento { get; set; }
        
        /// <summary> Número de Certificado de Disponibilidad Presupuestal (CDP). </summary>
        public string CDP { get; set; }
        
        /// <summary> Número de Registro Presupuestal (RP). </summary>
        public string RP { get; set; }
        
        /// <summary> Número de Obligación (OB). </summary>
        public string OB { get; set; }
        
        /// <summary> Número de Orden de Pago (OP). </summary>
        public string OP { get; set; }
        
        /// <summary> Tipo de gestión o trámite (ej. Solicitud, Cumplido). </summary>
        public string Tipo { get; set; }
        
        /// <summary> Número de comisión asociado (si se conoce). </summary>
        public string Comision { get; set; }
    }
}
