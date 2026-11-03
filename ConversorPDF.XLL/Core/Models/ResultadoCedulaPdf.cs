namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Resultado de la verificación de la cédula dentro del contenido de un PDF.
    /// </summary>
    public sealed class ResultadoCedulaPdf
    {
        private ResultadoCedulaPdf()
        {
            Cedula = string.Empty;
            Evidencia = string.Empty;
            Confianza = ConfianzaCedula.Ninguna;
        }

        /// <summary> True si se halló alguna cédula (aunque sea de confianza baja). </summary>
        public bool Hallada { get; private set; }

        /// <summary> Cédula hallada en solo dígitos; vacía si no se halló. </summary>
        public string Cedula { get; private set; }

        /// <summary> Nivel de confianza del hallazgo. </summary>
        public ConfianzaCedula Confianza { get; private set; }

        /// <summary> Motivo o patrón que sustentó el hallazgo (para trazabilidad en el log). </summary>
        public string Evidencia { get; private set; }

        /// <summary> True cuando el hallazgo es lo bastante fuerte para validar el archivo. </summary>
        public bool EsVerificacionFiable
        {
            get { return Hallada && Confianza >= ConfianzaCedula.Media; }
        }

        /// <summary>
        /// Crea un resultado sin hallazgo.
        /// </summary>
        /// <param name="motivo">Motivo del fallo (queda en <see cref="Evidencia"/>).</param>
        public static ResultadoCedulaPdf SinHallazgo(string motivo)
        {
            return new ResultadoCedulaPdf { Hallada = false, Evidencia = motivo ?? string.Empty };
        }

        /// <summary>
        /// Crea un resultado con hallazgo.
        /// </summary>
        /// <param name="cedula">Cédula en solo dígitos.</param>
        /// <param name="confianza">Nivel de confianza.</param>
        /// <param name="evidencia">Patrón o etiqueta que lo sustentó.</param>
        public static ResultadoCedulaPdf Hallazgo(string cedula, ConfianzaCedula confianza, string evidencia)
        {
            return new ResultadoCedulaPdf
            {
                Hallada = true,
                Cedula = cedula ?? string.Empty,
                Confianza = confianza,
                Evidencia = evidencia ?? string.Empty
            };
        }
    }
}
