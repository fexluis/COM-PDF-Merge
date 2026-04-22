using System.Collections.Generic;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Core.Models.Academic;

namespace ConversorPDF.XLL.Core.Interfaces
{
    /// <summary>
    /// Interfaz para el repositorio de datos académicos.
    /// </summary>
    public interface IAcademicRepository
    {
        // Programas
        SqliteResult<IEnumerable<Programa>> GetProgramas();
        SqliteResult<Programa> GetProgramaById(int id);
        SqliteResult<int> InsertPrograma(Programa programa);

        // Asignaturas
        SqliteResult<IEnumerable<Asignatura>> GetAsignaturasByPrograma(int idPrograma, string semestre = null);
        SqliteResult<Asignatura> GetAsignaturaByAbreviatura(string abreviatura);
        SqliteResult<int> InsertAsignatura(Asignatura asignatura);

        // Microcurrículos
        SqliteResult<Microcurriculo> GetMicrocurriculoByAsignatura(int idAsignatura);
        SqliteResult<int> InsertMicrocurriculo(Microcurriculo microcurriculo);

        // Contenidos
        SqliteResult<IEnumerable<Contenido>> GetContenidosByMicrocurriculo(int idMicrocurriculo);
        SqliteResult<int> InsertContenido(Contenido contenido);

        // Terceros
        SqliteResult<IEnumerable<Tercero>> GetTerceros();
        SqliteResult<Tercero> GetTerceroByCedula(string cedula);
        SqliteResult<int> InsertTercero(Tercero tercero);

        // Sedes
        SqliteResult<IEnumerable<Sede>> GetSedes();
        SqliteResult<Sede> GetSedeById(int id);
        SqliteResult<int> InsertSede(Sede sede);

        // Vinculaciones
        SqliteResult<IEnumerable<Vinculacion>> GetVinculaciones();
        SqliteResult<Vinculacion> GetVinculacionById(int id);
        SqliteResult<int> InsertVinculacion(Vinculacion vinculacion);
        SqliteResult<int> UpdateVinculacion(Vinculacion vinculacion);

        // Sesiones
        SqliteResult<IEnumerable<Sesion>> GetSesionesByVinculacion(int idVinculacion);
        SqliteResult<int> InsertSesion(Sesion sesion);
        SqliteResult<int> InsertSesiones(IEnumerable<Sesion> sesiones);
        SqliteResult<int> DeleteSesionesByVinculacion(int idVinculacion);
        SqliteResult<int> ExecuteRawSql(string sql);
        SqliteResult<int> SeedCatalogosTransaccional();
    }
}
