namespace ConversorPDF.XLL.Core.Enums
{
    /// <summary>
    /// Clasificación de documentos según los metadatos inyectados en el PDF.
    /// Permite a otros sistemas identificar el tipo de trámite sin procesar el contenido.
    /// </summary>
    public enum TipoMetadata
    {
        /// <summary> Certificado de Disponibilidad Presupuestal. </summary>
        CDP,
        /// <summary> Registro Presupuestal. </summary>
        RP,
        /// <summary> Obligación presupuestal del contador. </summary>
        OB,
        /// <summary> Orden de Pago generada. </summary>
        OP,
        /// <summary> Resolución o formato inicial de solicitud de viáticos. </summary>
        VIA_SOLICITUD,
        /// <summary> Planilla de viáticos verificada por el área administrativa. </summary>
        VIA_VERIFICADO,
        /// <summary> Documento de viáticos con soportes unificados. </summary>
        VIA_UNIFICADO,
        /// <summary> Informe de actividades presentado por el comisionado. </summary>
        VIA_INFORME,
        /// <summary> Certificado de cumplimiento de la comisión. </summary>
        VIA_CUMPLIDO,
        /// <summary> Registro de obligación de viáticos pendiente de legalizar. </summary>
        OB_VIA_SINLEGALIZAR,
        /// <summary> Registro de obligación de viáticos ya legalizado. </summary>
        OB_VIA_LEGALIZADO
    }
}
