namespace ConversorPDF.XLL.Academic.Domain.Enums
{
    /// <summary>
    /// Origen o estrategia utilizada para resolver la ventana temporal del semestre.
    /// </summary>
    public enum AgOrigenVentana
    {
        /// <summary>Derivada de los rangos de tipo 'PRESENCIAL' configurados en la tabla festivos.</summary>
        Presencial = 1,

        /// <summary>Derivada de los campos de fecha configurados en la entidad periodo o config.</summary>
        ConfigFechas = 2,

        /// <summary>Especificada o ajustada manualmente por el usuario en los parámetros de la corrida.</summary>
        Manual = 3,

        /// <summary>Valor por defecto o fallback derivado de las fechas mínimas/máximas de las sesiones.</summary>
        Default = 4,

        /// <summary>Derivada directamente del rango configurado en la tabla periodos (periodo_inicio / periodo_fin).</summary>
        Periodos = 5
    }
}

