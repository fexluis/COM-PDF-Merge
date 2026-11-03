namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Códigos canónicos del segmento TIPO del nombre de archivo de informes:
    /// <c>{solicitud}-{cedula}-{tipo}-{nombre}</c>.
    /// </summary>
    public static class InformeTipo
    {
        /// <summary> Informe de actividades / informe de cumplimiento. </summary>
        public const string Inf = "INF";

        /// <summary> Certificado de cumplimiento. </summary>
        public const string Cum = "CUM";

        /// <summary> Archivo unificado (CUM + INF) generado por el proceso. </summary>
        public const string Uni = "UNI";

        /// <summary> Archivo que no pudo validarse (criterio fail-closed). </summary>
        public const string Err = "ERR";

        /// <summary>
        /// True si el tipo participa en la formación de parejas INF + CUM.
        /// </summary>
        public static bool EsParValido(string tipo)
        {
            return tipo == Inf || tipo == Cum;
        }

        /// <summary>
        /// True si el tipo es uno de los admitidos en el nombre estandarizado.
        /// </summary>
        public static bool EsTipoValido(string tipo)
        {
            return tipo == Inf || tipo == Cum || tipo == Uni || tipo == Err;
        }
    }
}
