namespace ConversorPDF.XLL.Core.Enums
{
    /// <summary>
    /// Clasifica los tipos de vinculación laboral de los comisionados.
    /// Este valor determina qué hoja de Excel se utiliza como plantilla para la generación de PDFs.
    /// </summary>
    public enum TipoFormato
    {
        /// <summary> Personal contratado por prestación de servicios (Formato Honorarios). </summary>
        Contratista = 1,
        
        /// <summary> Personal vinculado a la planta global de la entidad. </summary>
        Servidor = 2,
        
        /// <summary> Docentes con vinculación especial por horas de cátedra. </summary>
        DocenteCatedra = 3
    }
}
