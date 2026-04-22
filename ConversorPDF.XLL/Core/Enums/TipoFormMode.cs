namespace ConversorPDF.XLL.Core.Enums
{
    /// <summary>
    /// Representa los estados de visualización y funcionalidad del formulario de programación de viáticos.
    /// Define cómo se comporta la UI en función del tipo de trámite.
    /// </summary>
    public enum TipoFormMode
    {
        /// <summary> Carga la interfaz para el procesamiento de una comisión de viáticos individual. </summary>
        verViaIndividual,
        
        /// <summary> Carga la interfaz para la generación masiva de múltiples PDF de viáticos. </summary>
        verViaMasivo,
        
        /// <summary> Carga la interfaz para la validación de cumplidos y generación de informes de actividades. </summary>
        verViaCumplidos
    }
}
