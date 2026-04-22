namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Modelo que encapsula los datos extraídos de documentos.
    /// </summary>
    public class ExtraerDatos
    {
        /// <summary> Número identificador de la obligación presupuestal (OB). </summary>
        public string xNumero { get; set; }
        
        /// <summary> Fecha de la obligación leída del documento. </summary>
        public string xFechaStr { get; set; }
        
        /// <summary> Nombre del beneficiario o tercero de la obligación. </summary>
        public string xTercero { get; set; }
        
        /// <summary> Documento de identidad (NIT/CC) del tercero. </summary>
        public string xIdentificacion { get; set; }
        
        /// <summary> Registro Presupuestal (RP) asociado a la obligación. </summary>
        public string xRP { get; set; }

        /// <summary> Certificado de Disponibilidad Presupuestal (CDP). </summary>
        public string xCDP { get; set; }

        /// <summary> Orden de Pago (OP). </summary>
        public string xOP { get; set; }

        /// <summary> Obligación presupuestal (OB). </summary>
        public string xOB { get; set; }

        /// <summary> Estado del documento (ej: Anulada). </summary>
        public string xEstado { get; set; }
        
        /// <summary> Dependencia o centro de costo responsable. </summary>
        public string xDep { get; set; }
        
        /// <summary> Descripción detallada del objeto de la obligación. </summary>
        public string xObjeto { get; set; }
        
        /// <summary> Categoría corta calculada (VIA, HON, SSO, etc.). </summary>
        public string xCategoria { get; set; }
        
        /// <summary> Nombre unificado sugerido para el archivo (sin extensión). </summary>
        public string xUnifiedFileName { get; set; }
    }
}
