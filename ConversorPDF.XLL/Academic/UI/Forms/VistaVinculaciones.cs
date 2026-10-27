namespace ConversorPDF.XLL.Academic.UI.Forms
{
    /// <summary>
    /// Define el modo de visualización para el árbol de vinculaciones.
    /// </summary>
    internal enum VistaVinculaciones
    {
        /// <summary>Docente → Sede|Semestre|Asignatura → Sesión.</summary>
        PorAsignaturas,

        /// <summary>Docente → Sesión (orden cronológico) con detección de conflictos.</summary>
        PorFechas,

        /// <summary>Programa → Semestre → Asignatura → Sede/Docente → Sesión.</summary>
        PorSemestres,

        /// <summary>Docente → Bloques Viáticos → Sesión.</summary>
        Viaticos
    }
}
