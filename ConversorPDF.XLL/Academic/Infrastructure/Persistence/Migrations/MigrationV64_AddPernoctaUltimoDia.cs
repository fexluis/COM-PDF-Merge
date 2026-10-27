using System;
using System.Data;
using System.Linq;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V64: Agrega la columna isPernoctaUltimoDia a vinculaciones si no existe,
    /// y actualiza la vista vw_SesionesDetalle para exponer IsPernoctaUltimoDia.
    /// </summary>
    public class MigrationV64_AddPernoctaUltimoDia : IMigration
    {
        public int Version => 64;
        public string Name => "AddPernoctaUltimoDia";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            if (connection.QueryFirstOrDefault<int>("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name='vinculaciones';", transaction: transaction) == 0)
                return;

            // 1. Agregar columna isPernoctaUltimoDia a vinculaciones de forma tolerante/idempotente
            var columns = connection.Query("PRAGMA table_info(vinculaciones)", transaction: transaction)
                .Select(r => (string)((dynamic)r).name)
                .ToList();

            if (!columns.Contains("isPernoctaUltimoDia", StringComparer.OrdinalIgnoreCase))
            {
                connection.Execute("ALTER TABLE vinculaciones ADD COLUMN isPernoctaUltimoDia INTEGER DEFAULT 0;", transaction: transaction);
            }

            // 2. Recrear la vista vw_SesionesDetalle incluyendo IsPernoctaUltimoDia
            connection.Execute("DROP VIEW IF EXISTS vw_SesionesDetalle;", transaction: transaction);

            connection.Execute(@"
CREATE VIEW vw_SesionesDetalle AS
SELECT 
    s.id AS IdSesion,
    s.idvinculacion AS IdVinculacion,
    v.idperiodo AS IdPeriodo,
    v.idsede AS IdSede,
    v.idprograma AS IdPrograma,
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
    s2.subsede AS Subsede,
    CASE strftime('%w', NULLIF(s.fechasesion, ''))
        WHEN '0' THEN 'DOMINGO'
        WHEN '1' THEN 'LUNES'
        WHEN '2' THEN 'MARTES'
        WHEN '3' THEN 'MIÉRCOLES'
        WHEN '4' THEN 'JUEVES'
        WHEN '5' THEN 'VIERNES'
        WHEN '6' THEN 'SÁBADO'
    END AS DiaSemana,
    s.horas AS Horas,
    s.horariosesion AS HorarioSesion,
    NULLIF(s.fechasesion, '') AS FechaSesion,
    s.viatica AS Viatica,
    s2.departamento AS SedeDepartamento,
    s2.transporte AS TransporteSede,
    p.nombreprograma AS ProgramaNombre,
    a.creditos AS AsignaturaCreditos,
    a.metodologia AS AsignaturaMetodologia,
    a.cod_materia AS AsignaturaCodMateria,
    a.num_nivel AS Nivel,
    v.grupo AS Grupo,
    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS VinculacionFechaInicio,
    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS VinculacionFechaFin,
    v.modalidadVinculacion AS Modalidad,
    v.isPernoctaUltimoDia AS IsPernoctaUltimoDia,
    v.isTerminalAereo AS IsTerminalAereo,
    v.idTerminales AS IdTerminales,
    v.gastoTerminalAerea AS GastoTerminalAerea,
    term.terminal AS TerminalAereaNombre,
    v.subunidad AS Subunidad
FROM sesiones s
JOIN vinculaciones v ON s.idvinculacion = v.id
JOIN programas p ON v.idprograma = p.id
JOIN asignaturas a ON v.idasignatura = a.id
JOIN sedes s2 ON v.idsede = s2.id
LEFT JOIN terceros t ON s.idtercero = t.id
LEFT JOIN terminales term ON (CASE WHEN v.idTerminales IS NULL OR v.idTerminales <= 0 THEN 1 ELSE v.idTerminales END) = term.id;
            ", transaction: transaction);
        }
    }
}
