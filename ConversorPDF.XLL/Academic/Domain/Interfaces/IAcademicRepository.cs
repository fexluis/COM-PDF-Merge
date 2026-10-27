using System.Collections.Generic;
using ConversorPDF.XLL.Core.Models;
using ConversorPDF.XLL.Academic.Domain.Models;

namespace ConversorPDF.XLL.Academic.Domain.Interfaces
{
    /// <summary>
    /// Interfaz para el repositorio de datos académicos.
    /// </summary>
    public interface IAcademicRepository
    {
        /// <summary> Obtiene todos los programas académicos. </summary>
        SqliteResult<IEnumerable<Programa>> GetProgramas();
        /// <summary> Obtiene un programa por su identificador. </summary>
        SqliteResult<Programa> GetProgramaById(int id);
        /// <summary> Inserta un nuevo programa. </summary>
        SqliteResult<int> InsertPrograma(Programa programa);

        // Asignaturas
        /// <summary> Obtiene las asignaturas de un programa, opcionalmente filtradas por semestre. </summary>
        SqliteResult<IEnumerable<Asignatura>> GetAsignaturasByPrograma(int idPrograma, string semestre = null);
        /// <summary> Obtiene una asignatura por su abreviatura. </summary>
        SqliteResult<Asignatura> GetAsignaturaByAbreviatura(string abreviatura);
        /// <summary> Inserta una nueva asignatura. </summary>
        SqliteResult<int> InsertAsignatura(Asignatura asignatura);

        // Microcurrículos
        /// <summary> Obtiene el microcurriculo asociado a una asignatura. </summary>
        SqliteResult<Microcurriculo> GetMicrocurriculoByAsignatura(int idAsignatura);
        /// <summary> Inserta un nuevo microcurriculo. </summary>
        SqliteResult<int> InsertMicrocurriculo(Microcurriculo microcurriculo);

        // Contenidos
        /// <summary> Obtiene los contenidos asociados a un microcurriculo. </summary>
        SqliteResult<IEnumerable<Contenido>> GetContenidosByMicrocurriculo(int idMicrocurriculo);
        /// <summary> Inserta un nuevo contenido programático. </summary>
        SqliteResult<int> InsertContenido(Contenido contenido);

        // Terceros
        /// <summary> Obtiene todos los terceros (docentes/contratistas). </summary>
        SqliteResult<IEnumerable<Tercero>> GetTerceros(bool soloActivos = true);
        /// <summary> Obtiene un tercero por su cédula. </summary>
        SqliteResult<Tercero> GetTerceroByCedula(string cedula);
        /// <summary> Inserta un nuevo tercero. </summary>
        SqliteResult<int> InsertTercero(Tercero tercero);
        /// <summary> Actualiza la información de un tercero existente. </summary>
        SqliteResult<int> UpdateTercero(Tercero tercero);
        /// <summary> Elimina un tercero por su identificador. </summary>
        SqliteResult<int> DeleteTercero(int id);

        // Sedes
        /// <summary> Obtiene todas las sedes registradas. </summary>
        SqliteResult<IEnumerable<Sede>> GetSedes();
        /// <summary> Obtiene una sede por su identificador. </summary>
        SqliteResult<Sede> GetSedeById(int id);
        /// <summary> Inserta una nueva sede. </summary>
        SqliteResult<int> InsertSede(Sede sede);
        /// <summary> Actualiza la información de una sede. </summary>
        SqliteResult<int> UpdateSede(Sede sede);
        /// <summary> Elimina una sede por su identificador. </summary>
        SqliteResult<int> DeleteSede(int id);

        // Terminales
        /// <summary> Obtiene la lista de terminales. </summary>
        SqliteResult<IEnumerable<Terminal>> GetTerminales();

        // Divipola
        /// <summary> Obtiene todos los registros de Divipola. </summary>
        SqliteResult<IEnumerable<Divipola>> GetDivipolas();
        /// <summary> Obtiene un registro de Divipola por su identificador. </summary>
        SqliteResult<Divipola> GetDivipolaById(int id);
        /// <summary> Busca en Divipola por nombre de municipio o departamento. </summary>
        SqliteResult<IEnumerable<Divipola>> SearchDivipola(string searchTerm);
        /// <summary> Inserta un nuevo registro de Divipola. </summary>
        SqliteResult<int> InsertDivipola(Divipola divipola);
        /// <summary> Actualiza la información de un registro de Divipola. </summary>
        SqliteResult<int> UpdateDivipola(Divipola divipola);
        /// <summary> Elimina un registro de Divipola por su identificador. </summary>
        SqliteResult<int> DeleteDivipola(int id);

        // Periodos
        /// <summary> Obtiene los períodos académicos. </summary>
        SqliteResult<IEnumerable<Periodo>> GetPeriodos();
        /// <summary> Obtiene un período por su identificador. </summary>
        SqliteResult<Periodo> GetPeriodoById(int id);
        /// <summary> Inserta un nuevo período académico. </summary>
        SqliteResult<int> InsertPeriodo(Periodo periodo);
        /// <summary> Actualiza un período académico existente. </summary>
        SqliteResult<int> UpdatePeriodo(Periodo periodo);
        /// <summary> Elimina un período académico verificando dependencias en vinculaciones, festivos y configuración. </summary>
        SqliteResult<int> DeletePeriodo(int id);
        /// <summary> Obtiene los períodos académicos proyectando la ventana temporal canónica. </summary>
        SqliteResult<IEnumerable<Periodo>> GetPeriodosConVentana();
        
        // Subunidades
        /// <summary> Obtiene la lista centralizada de subunidades desde territoriales. </summary>
        SqliteResult<IEnumerable<string>> GetSubunidades();

        // Vinculaciones
        /// <summary> Obtiene todas las vinculaciones académicas. </summary>
        SqliteResult<IEnumerable<Vinculacion>> GetVinculaciones();
        /// <summary> Obtiene una vinculación por su identificador. </summary>
        SqliteResult<Vinculacion> GetVinculacionById(int id);
        /// <summary> Inserta una nueva vinculación. </summary>
        SqliteResult<int> InsertVinculacion(Vinculacion vinculacion);
        /// <summary> Actualiza una vinculación existente. </summary>
        SqliteResult<int> UpdateVinculacion(Vinculacion vinculacion);
        /// <summary> Elimina una vinculación por su identificador. </summary>
        SqliteResult<int> DeleteVinculacion(int idVinculacion);
        /// <summary> Actualiza el flag de pernocta en último día de una vinculación. </summary>
        SqliteResult<int> SetPernoctaUltimoDia(int idVinculacion, int valor);

        // Sesiones
        /// <summary> Obtiene las sesiones asociadas a una vinculación. </summary>
        SqliteResult<IEnumerable<Sesion>> GetSesionesByVinculacion(int idVinculacion);
        /// <summary> Inserta una sesión individual. </summary>
        SqliteResult<int> InsertSesion(Sesion sesion);
        /// <summary> Inserta un lote de sesiones. </summary>
        SqliteResult<int> InsertSesiones(IEnumerable<Sesion> sesiones);
        /// <summary> Elimina todas las sesiones de una vinculación. </summary>
        SqliteResult<int> DeleteSesionesByVinculacion(int idVinculacion);
        // Festivos
        /// <summary> Obtiene todos los festivos. </summary>
        SqliteResult<IEnumerable<Festivo>> GetFestivos();
        /// <summary> Obtiene un festivo por su identificador. </summary>
        SqliteResult<Festivo> GetFestivoById(int id);
        /// <summary> Inserta un nuevo festivo. </summary>
        SqliteResult<int> InsertFestivo(Festivo festivo);
        /// <summary> Actualiza un festivo existente. </summary>
        SqliteResult<int> UpdateFestivo(Festivo festivo);
        /// <summary> Elimina un festivo por su identificador. </summary>
        SqliteResult<int> DeleteFestivo(int id);

        /// <summary> Ejecuta SQL crudo (uso interno). </summary>
        SqliteResult<int> ExecuteRawSql(string sql);
        /// <summary> Inicializa los catálogos con datos semilla de forma transaccional. </summary>
        SqliteResult<int> SeedCatalogosTransaccional();
        
        // Usuarios
        /// <summary> Obtiene todos los usuarios. </summary>
        SqliteResult<IEnumerable<Usuario>> GetUsuarios();
        /// <summary> Obtiene un usuario por nombre de usuario. </summary>
        SqliteResult<Usuario> GetUsuarioByNombreUsuario(string nombreUsuario);
        /// <summary> Inserta un nuevo usuario. </summary>
        SqliteResult<string> InsertUsuario(Usuario usuario);
        /// <summary> Actualiza un usuario existente. </summary>
        SqliteResult<int> UpdateUsuario(Usuario usuario);
        /// <summary> Elimina un usuario por su identificador. </summary>
        SqliteResult<int> DeleteUsuario(string id);
        
        /// <summary> Obtiene los permisos (identificadores separados por comas) asociados al perfil del usuario. </summary>
        SqliteResult<string> GetPermisosPorUsuario(string nombreUsuario);

        // CRUD Permisos
        /// <summary> Obtiene la lista de perfiles y sus permisos. </summary>
        SqliteResult<IEnumerable<Permiso>> GetTodosLosPermisos();
        /// <summary> Guarda (inserta o actualiza) los permisos de un perfil. </summary>
        SqliteResult<int> SavePermiso(Permiso permiso);
        /// <summary> Elimina un perfil y sus permisos. </summary>
        SqliteResult<int> DeletePermiso(string perfil);
    }
}
