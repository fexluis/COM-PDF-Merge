namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// DTO para la visualización de asignaturas con vacantes pendientes en <see cref="FormAutorizacionVacantes"/> (Fase 2037).
    /// </summary>
    internal class ItemVacanteDto
    {
        public int IdAsignatura { get; set; }
        public int IdSede { get; set; }
        public string Texto { get; set; }

        public string Clave => IdSede > 0 ? $"{IdAsignatura}:{IdSede}" : IdAsignatura.ToString();

        public override string ToString()
        {
            return Texto;
        }
    }
}
