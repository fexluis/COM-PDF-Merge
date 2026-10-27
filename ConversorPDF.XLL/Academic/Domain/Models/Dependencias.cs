using System.ComponentModel;

namespace ConversorPDF.XLL.Academic.Domain.Models
{
    public enum Dependencias
    {
        [Description("PREGRADO")]
        PREGRADO,
        [Description("POSGRADO")]
        POSGRADO,
        [Description("APOYO ACADEMICO")]
        APOYO_ACADEMICO,
        [Description("FORTALECIMIENTO")]
        FORTALECIMIENTO,
        [Description("CAPACITACION")]
        CAPACITACION,
        [Description("DIRECCION")]
        DIRECCION,
        [Description("OTIC")]
        OTIC
    }
}
