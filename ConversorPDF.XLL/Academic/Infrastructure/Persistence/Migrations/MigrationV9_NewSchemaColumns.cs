using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    public class MigrationV9_NewSchemaColumns : IMigration
    {
        public int Version => 9;
        public string Name => "Alineación de Modelos con Schema Académico V9";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            // 1. ELIMINAR VISTAS DEPENDIENTES PRIMERO (CRÍTICO PARA SQLITE)
            connection.Execute("DROP VIEW IF EXISTS vw_VinculacionesDetalle;", transaction: transaction);
            connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction);

            // 2. RENOMBRAR Y AÑADIR COLUMNAS
            
            // Programas
            var progCols = connection.Query("PRAGMA table_info(programas)", transaction: transaction)
                                     .Select(c => (string)c.name).ToList();
            if (progCols.Contains("tipo", StringComparer.OrdinalIgnoreCase) && !progCols.Contains("modalidad", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE programas RENAME COLUMN tipo TO modalidad;", transaction: transaction);
            if (!progCols.Contains("pensum", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE programas ADD COLUMN pensum INTEGER;", transaction: transaction);
                connection.Execute("ALTER TABLE programas ADD COLUMN actual INTEGER DEFAULT 0;", transaction: transaction);
                connection.Execute("ALTER TABLE programas ADD COLUMN activo INTEGER DEFAULT 0;", transaction: transaction);
            }

            // Asignaturas
            var asigCols = connection.Query("PRAGMA table_info(asignaturas)", transaction: transaction)
                                     .Select(c => (string)c.name).ToList();
            if (asigCols.Contains("codigo", StringComparer.OrdinalIgnoreCase) && !asigCols.Contains("cod_materia", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE asignaturas RENAME COLUMN codigo TO cod_materia;", transaction: transaction);
            if (asigCols.Contains("nombreasignatura", StringComparer.OrdinalIgnoreCase) && !asigCols.Contains("nombre", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE asignaturas RENAME COLUMN nombreasignatura TO nombre;", transaction: transaction);
            if (asigCols.Contains("total_horas", StringComparer.OrdinalIgnoreCase) && !asigCols.Contains("int_horaria", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE asignaturas RENAME COLUMN total_horas TO int_horaria;", transaction: transaction);
            if (!asigCols.Contains("num_nivel", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE asignaturas ADD COLUMN num_nivel INTEGER;", transaction: transaction);
                connection.Execute("ALTER TABLE asignaturas ADD COLUMN cod_unidad TEXT;", transaction: transaction);
                connection.Execute("ALTER TABLE asignaturas ADD COLUMN uni_teorica INTEGER;", transaction: transaction);
                connection.Execute("ALTER TABLE asignaturas ADD COLUMN metodologia TEXT;", transaction: transaction);
            }

            // Terceros
            var terCols = connection.Query("PRAGMA table_info(terceros)", transaction: transaction)
                                    .Select(c => (string)c.name).ToList();
            if (terCols.Contains("rol", StringComparer.OrdinalIgnoreCase) && !terCols.Contains("categoria", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE terceros RENAME COLUMN rol TO categoria;", transaction: transaction);
            if (!terCols.Contains("pensionado", StringComparer.OrdinalIgnoreCase))
                connection.Execute("ALTER TABLE terceros ADD COLUMN pensionado INTEGER DEFAULT 0;", transaction: transaction);

            // 3. RECREAR VISTAS CON NUEVOS NOMBRES
            connection.Execute(@"
                CREATE VIEW vw_VinculacionesDetalle AS
                SELECT 
                    v.id AS IdVinculacion,
                    v.idsede AS IdSede,
                    v.idtercero AS IdTercero,
                    pe.idperiodo AS IdPeriodo,
                    pe.periodo AS PeriodoNombre,
                    p.modalidad AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    t.nombre AS TerceroNombre,
                    CASE 
                        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                        ELSE v.semestre || '-' || v.grupo 
                    END AS Semestre,
                    a.nombre AS AsignaturaNombre,
                    s.ciudad AS SedeCiudad,
                    v.fecha_inicio AS SesionFechaInicial,
                    v.fecha_fin AS SesionFechaFinal,
                    v.total_horas AS SesionHoras,
                    (SELECT CASE WHEN EXISTS (
                        SELECT 1 FROM sesiones ses WHERE ses.idvinculacion = v.id AND ses.viatica = 1
                    ) THEN 1 ELSE 0 END) AS SesionViatica
                FROM vinculaciones v
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s ON v.idsede = s.id
                LEFT JOIN terceros t ON v.idtercero = t.id
                LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;
            ", transaction: transaction);

            connection.Execute(@"
                CREATE VIEW vw_SesionesDetalle AS
                SELECT 
                    s.id AS IdSesion,
                    s.idvinculacion AS IdVinculacion,
                    v.idperiodo AS IdPeriodo,
                    v.idsede AS IdSede,
                    s.idtercero AS IdTercero,
                    p.modalidad AS ProgramaTipo,
                    p.abreviatura AS ProgramaAbreviatura,
                    CASE 
                        WHEN v.grupo IS NULL OR v.grupo = '' OR v.grupo = '--' THEN v.semestre
                        ELSE v.semestre || '-' || v.grupo 
                    END AS Semestre,
                    t.nombre AS DocenteNombre,
                    t.cedula AS CedulaTercero,
                    a.nombre AS AsignaturaNombre,
                    a.abreviatura AS AsignaturaAbreviatura,
                    s2.ciudad AS SedeCiudad,
                    CASE strftime('%w', s.fechasesionInicial)
                        WHEN '0' THEN 'DOMINGO'
                        WHEN '1' THEN 'LUNES'
                        WHEN '2' THEN 'MARTES'
                        WHEN '3' THEN 'MIÉRCOLES'
                        WHEN '4' THEN 'JUEVES'
                        WHEN '5' THEN 'VIERNES'
                        WHEN '6' THEN 'SÁBADO'
                    END AS DiaSemana,
                    s.horas AS Horas,
                    s.duracion AS Duracion,
                    s.fechasesionInicial AS FechaInicial,
                    s.fechasesionFinal AS FechaFinal,
                    s.viatica AS Viatica
                FROM sesiones s
                JOIN vinculaciones v ON s.idvinculacion = v.id
                JOIN programas p ON v.idprograma = p.id
                JOIN asignaturas a ON v.idasignatura = a.id
                JOIN sedes s2 ON v.idsede = s2.id
                LEFT JOIN terceros t ON s.idtercero = t.id;
            ", transaction: transaction);
        }
    }
}
