using System;
using System.Data;
using Dapper;

namespace ConversorPDF.XLL.Academic.Infrastructure.Persistence.Migrations
{
    /// <summary>
    /// Migración V49: Crea las 13 tablas del motor del Algoritmo Genético, 2 vistas auxiliares de staging y siembra inicial de bloques horarios y aulas.
    /// </summary>
    public class MigrationV49_AddAgHorariosTables : IMigration
    {
        public int Version => 49;
        public string Name => "AddAgHorariosTables";

        public void Up(IDbConnection connection, IDbTransaction transaction)
        {
            // 1. Tablas Maestras del AG
            string ddlSql = @"
CREATE TABLE IF NOT EXISTS ag_aula (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    idsede                  INTEGER NOT NULL DEFAULT 0,
    idperiodo               INTEGER NOT NULL DEFAULT 0,
    idprograma              INTEGER NOT NULL DEFAULT 0,
    codigo                  TEXT    NOT NULL DEFAULT '',
    nombre                  TEXT    NOT NULL DEFAULT '',
    capacidad               INTEGER NOT NULL DEFAULT 30,
    tipo_aula               TEXT    NOT NULL DEFAULT 'ESTANDAR',
    es_virtual              INTEGER NOT NULL DEFAULT 0,
    estado                  INTEGER NOT NULL DEFAULT 1,
    caracteristicas_json    TEXT    NOT NULL DEFAULT '{}',
    semestre                TEXT    NOT NULL DEFAULT '',
    grupo                   TEXT    NOT NULL DEFAULT '',
    modalidadVinculacion    TEXT    NOT NULL DEFAULT 'Presencial',
    fecha_inicio            TEXT    NOT NULL DEFAULT '',
    fecha_fin               TEXT    NOT NULL DEFAULT '',
    semestre_inicio         TEXT    NOT NULL DEFAULT '',
    semestre_fin            TEXT    NOT NULL DEFAULT '',
    ventana_origen          TEXT    NOT NULL DEFAULT '',
    aula_ref                TEXT    NOT NULL DEFAULT '',
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_aula_sede ON ag_aula(idsede);

CREATE TABLE IF NOT EXISTS ag_aula_asignatura (
    id_aula                 INTEGER NOT NULL,
    id_asignatura           INTEGER NOT NULL,
    exclusivo               INTEGER NOT NULL DEFAULT 0,
    prioridad               INTEGER NOT NULL DEFAULT 0,
    horas_objetivo          REAL    NOT NULL DEFAULT 0.0,
    orden_curricular        INTEGER NOT NULL DEFAULT 0,
    requerida               INTEGER NOT NULL DEFAULT 1,
    subunidad               TEXT    NOT NULL DEFAULT '',
    PRIMARY KEY (id_aula, id_asignatura)
);

CREATE TABLE IF NOT EXISTS ag_bloque_horario (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    dia_semana              INTEGER NOT NULL DEFAULT 1,
    hora_inicio             TEXT    NOT NULL DEFAULT '08:00',
    hora_fin                TEXT    NOT NULL DEFAULT '12:00',
    duracion_horas          REAL    NOT NULL DEFAULT 4.0,
    franja                  TEXT    NOT NULL DEFAULT 'MANANA',
    activo                  INTEGER NOT NULL DEFAULT 1,
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_bloque_dia ON ag_bloque_horario(dia_semana, activo);

CREATE TABLE IF NOT EXISTS ag_corrida (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    idperiodo               INTEGER NOT NULL DEFAULT 0,
    idprograma              INTEGER NOT NULL DEFAULT 0,
    fecha_inicio_corrida    TEXT    NOT NULL DEFAULT '',
    fecha_fin_corrida       TEXT    NOT NULL DEFAULT '',
    estado                  INTEGER NOT NULL DEFAULT 0,
    poblacion_tamano        INTEGER NOT NULL DEFAULT 50,
    generaciones_max        INTEGER NOT NULL DEFAULT 200,
    generaciones_ejecutadas INTEGER NOT NULL DEFAULT 0,
    tasa_cruce              REAL    NOT NULL DEFAULT 0.85,
    tasa_mutacion           REAL    NOT NULL DEFAULT 0.10,
    elitismo                INTEGER NOT NULL DEFAULT 2,
    mejor_fitness           REAL    NOT NULL DEFAULT 0.0,
    hard_violations         INTEGER NOT NULL DEFAULT 0,
    soft_violations         REAL    NOT NULL DEFAULT 0.0,
    id_individuo_optimo     INTEGER NULL,
    parametros_json         TEXT    NOT NULL DEFAULT '{}',
    subunidad               TEXT    NOT NULL DEFAULT ''
);

CREATE TABLE IF NOT EXISTS ag_individuo (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_corrida              INTEGER NOT NULL,
    generacion              INTEGER NOT NULL DEFAULT 0,
    fitness                 REAL    NOT NULL DEFAULT 0.0,
    hard_violations         INTEGER NOT NULL DEFAULT 0,
    soft_violations         REAL    NOT NULL DEFAULT 0.0,
    estado                  INTEGER NOT NULL DEFAULT 0,
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_ind_corrida ON ag_individuo(id_corrida);

CREATE TABLE IF NOT EXISTS ag_gen (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_individuo            INTEGER NOT NULL,
    idperiodo               INTEGER NOT NULL DEFAULT 0,
    idasignatura            INTEGER NOT NULL,
    idprograma              INTEGER NOT NULL,
    idsede                  INTEGER NOT NULL,
    semestre                TEXT    NOT NULL DEFAULT '',
    grupo                   TEXT    NOT NULL DEFAULT 'A',
    idtercero               INTEGER NOT NULL DEFAULT 0,
    id_aula                 INTEGER NOT NULL DEFAULT 0,
    id_bloque_horario       INTEGER NOT NULL DEFAULT 0,
    dia_semana              INTEGER NOT NULL DEFAULT 1,
    horas_semanales         REAL    NOT NULL DEFAULT 4.0,
    total_horas_asignatura  REAL    NOT NULL DEFAULT 32.0,
    fecha_inicio_estimada   TEXT    NULL,
    fecha_fin_estimada      TEXT    NULL,
    es_fijo                 INTEGER NOT NULL DEFAULT 0,
    es_vacante              INTEGER NOT NULL DEFAULT 0,
    motivo_vacante          TEXT    NOT NULL DEFAULT '',
    dentro_de_ventana       INTEGER NOT NULL DEFAULT 1,
    estado                  INTEGER NOT NULL DEFAULT 0,
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_gen_ind ON ag_gen(id_individuo);

CREATE TABLE IF NOT EXISTS ag_violacion (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_individuo            INTEGER NOT NULL,
    id_gen                  INTEGER NULL,
    regla                   INTEGER NOT NULL,
    es_dura                 INTEGER NOT NULL DEFAULT 1,
    mensaje                 TEXT    NOT NULL DEFAULT '',
    detalle_json            TEXT    NOT NULL DEFAULT '{}',
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_viol_ind ON ag_violacion(id_individuo);

CREATE TABLE IF NOT EXISTS ag_bitacora (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_corrida              INTEGER NOT NULL,
    generacion              INTEGER NOT NULL,
    mejor_fitness           REAL    NOT NULL,
    promedio_fitness        REAL    NOT NULL,
    peor_fitness            REAL    NOT NULL DEFAULT 0.0,
    hard_violations         INTEGER NOT NULL,
    evaluados               INTEGER NOT NULL DEFAULT 0,
    descartados             INTEGER NOT NULL DEFAULT 0,
    vacantes                INTEGER NOT NULL DEFAULT 0,
    fuera_de_ventana        INTEGER NOT NULL DEFAULT 0,
    duracion_ms             INTEGER NOT NULL DEFAULT 0,
    fecha_registro          TEXT    NOT NULL DEFAULT '',
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_bitacora_corrida ON ag_bitacora(id_corrida);

CREATE TABLE IF NOT EXISTS ag_preferencia_docente (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    idtercero               INTEGER NOT NULL,
    dia_semana              INTEGER NOT NULL,
    franja                  TEXT    NOT NULL DEFAULT 'MANANA',
    preferencia             INTEGER NOT NULL DEFAULT 1,
    subunidad               TEXT    NOT NULL DEFAULT ''
);

-- 2. Tablas de Staging del AG
CREATE TABLE IF NOT EXISTS ag_vinculaciones (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_corrida              INTEGER NOT NULL DEFAULT 0,
    id_individuo            INTEGER NOT NULL DEFAULT 0,
    idtercero               INTEGER NOT NULL DEFAULT 0,
    idprograma              INTEGER NOT NULL DEFAULT 0,
    idsede                  INTEGER NOT NULL DEFAULT 0,
    idasignatura            INTEGER NOT NULL DEFAULT 0,
    idperiodo               INTEGER NOT NULL DEFAULT 0,
    semestre                TEXT    NOT NULL DEFAULT '',
    grupo                   TEXT    NOT NULL DEFAULT '',
    dia_horas_semanales     TEXT    NOT NULL DEFAULT '',
    modalidadVinculacion    TEXT    NOT NULL DEFAULT 'Presencial',
    total_horas             INTEGER NOT NULL DEFAULT 0,
    fecha_inicio            TEXT    NOT NULL DEFAULT '',
    fecha_fin               TEXT    NOT NULL DEFAULT '',
    estado                  INTEGER NOT NULL DEFAULT 0,
    isTerminalAereo         INTEGER NOT NULL DEFAULT 0,
    idTerminales            INTEGER NOT NULL DEFAULT 0,
    gastoTerminalAerea      REAL    NOT NULL DEFAULT 0.0,
    es_vacante              INTEGER NOT NULL DEFAULT 0,
    motivo_vacante          TEXT    NOT NULL DEFAULT '',
    candidatos_pool         INTEGER NOT NULL DEFAULT 0,
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_vinc_corrida ON ag_vinculaciones(id_corrida);

CREATE TABLE IF NOT EXISTS ag_vinculacion_gen (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    idagvinculacion         INTEGER NOT NULL,
    idgen                   INTEGER NOT NULL,
    subunidad               TEXT    DEFAULT ''
);
CREATE UNIQUE INDEX IF NOT EXISTS uq_ag_vinc_gen ON ag_vinculacion_gen(idagvinculacion, idgen);

CREATE TABLE IF NOT EXISTS ag_sesiones (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    idvinculacion           INTEGER NOT NULL,
    idtercero               INTEGER NOT NULL,
    id_aula                 INTEGER NOT NULL DEFAULT 0,
    fechasesion             TEXT    NOT NULL,
    horariosesion           TEXT    NOT NULL DEFAULT '',
    horas                   REAL    NOT NULL DEFAULT 0.0,
    viatica                 INTEGER NOT NULL DEFAULT 0,
    subunidad               TEXT    NOT NULL DEFAULT '',
    idindividuo             INTEGER NULL,
    idgen                   INTEGER NULL,
    origen_ag               INTEGER NOT NULL DEFAULT 1,
    estado_ag               TEXT    NOT NULL DEFAULT 'BORRADOR',
    es_vacante              INTEGER NOT NULL DEFAULT 0,
    motivo_vacante          TEXT    NOT NULL DEFAULT '',
    candidatos_pool         INTEGER NOT NULL DEFAULT 0,
    fecha_creacion          TEXT    NOT NULL DEFAULT (datetime('now'))
);
CREATE INDEX IF NOT EXISTS idx_ag_ses_vinc ON ag_sesiones(idvinculacion);

CREATE TABLE IF NOT EXISTS ag_salvedades_conflictos (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_corrida              INTEGER NOT NULL,
    idsesiona               INTEGER NOT NULL,
    idsesionb               INTEGER NOT NULL,
    idtercero               INTEGER NOT NULL,
    idperiodo               INTEGER NOT NULL,
    justificacion           TEXT    NOT NULL DEFAULT '',
    subunidad               TEXT    NOT NULL DEFAULT ''
);
CREATE INDEX IF NOT EXISTS idx_ag_salv_corrida ON ag_salvedades_conflictos(id_corrida);

CREATE TABLE IF NOT EXISTS ag_promocion (
    id                      INTEGER PRIMARY KEY AUTOINCREMENT,
    id_corrida              INTEGER NOT NULL,
    id_individuo            INTEGER NOT NULL,
    idaula                  INTEGER NOT NULL DEFAULT 0,
    idagvinculacion         INTEGER NOT NULL DEFAULT 0,
    idvinculacion           INTEGER NOT NULL DEFAULT 0,
    idtercero               INTEGER NOT NULL DEFAULT 0,
    idasignatura            INTEGER NOT NULL DEFAULT 0,
    fecha_inicio            TEXT    NOT NULL DEFAULT '',
    fecha_fin               TEXT    NOT NULL DEFAULT '',
    modo                    TEXT    NOT NULL DEFAULT 'COPIA_REMAPEO',
    justificacion_vacante   TEXT    NOT NULL DEFAULT '',
    fecha_promocion         TEXT    NOT NULL DEFAULT '',
    usuario                 TEXT    NOT NULL DEFAULT '',
    total_vinculaciones     INTEGER NOT NULL DEFAULT 0,
    total_sesiones          INTEGER NOT NULL DEFAULT 0,
    subunidad               TEXT    NOT NULL DEFAULT ''
);";

            connection.Execute(ddlSql, transaction: transaction);

            // 3. Vistas Auxiliares de Staging (Idempotentes con DROP VIEW IF EXISTS)
            string viewsSql = @"
DROP VIEW IF EXISTS vw_AgSesionesDetalle;
CREATE VIEW vw_AgSesionesDetalle AS
SELECT 
    s.id AS IdSesion,
    s.idvinculacion AS IdVinculacion,
    v.id_corrida AS IdCorrida,
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
    v.isTerminalAereo AS IsTerminalAereo,
    v.idTerminales AS IdTerminales,
    v.gastoTerminalAerea AS GastoTerminalAerea,
    term.terminal AS TerminalAereaNombre,
    v.subunidad AS Subunidad
FROM ag_sesiones s
JOIN ag_vinculaciones v ON s.idvinculacion = v.id AND v.estado = 0
LEFT JOIN programas p ON v.idprograma = p.id
LEFT JOIN asignaturas a ON v.idasignatura = a.id
LEFT JOIN sedes s2 ON v.idsede = s2.id
LEFT JOIN terceros t ON s.idtercero = t.id
LEFT JOIN terminales term ON (CASE WHEN v.idTerminales IS NULL OR v.idTerminales <= 0 THEN 1 ELSE v.idTerminales END) = term.id;

DROP VIEW IF EXISTS vw_AgVinculacionesDetalle;
CREATE VIEW vw_AgVinculacionesDetalle AS
SELECT v.id AS IdVinculacion, v.id_corrida AS IdCorrida, v.idsede AS IdSede, v.idtercero AS IdTercero,
       pe.idperiodo AS IdPeriodo, pe.periodo AS PeriodoNombre, p.modalidad AS ProgramaTipo,
       p.abreviatura AS ProgramaAbreviatura, t.nombre AS TerceroNombre,
       CASE WHEN v.grupo IS NULL OR v.grupo='' OR v.grupo='--' THEN v.semestre
            ELSE v.semestre || '-' || v.grupo END AS Semestre,
       a.nombre AS AsignaturaNombre, s.ciudad AS SedeCiudad,
       NULLIF(NULLIF(v.fecha_inicio,''),'--') AS SesionFechaInicial,
       NULLIF(NULLIF(v.fecha_fin,''),'--')    AS SesionFechaFinal,
       v.total_horas AS SesionHoras,
       (SELECT CASE WHEN EXISTS (SELECT 1 FROM ag_sesiones ses
                                 WHERE ses.idvinculacion=v.id AND ses.viatica=1)
                    THEN 1 ELSE 0 END) AS SesionViatica,
       v.modalidadVinculacion AS Modalidad, v.subunidad AS Subunidad
FROM ag_vinculaciones v
LEFT JOIN programas  p  ON v.idprograma  = p.id
LEFT JOIN asignaturas a ON v.idasignatura = a.id
LEFT JOIN sedes      s  ON v.idsede      = s.id
LEFT JOIN terceros   t  ON v.idtercero   = t.id
LEFT JOIN periodos   pe ON v.idperiodo   = pe.idperiodo;";

            connection.Execute(viewsSql, transaction: transaction);

            // Permisos para el botón del Ribbon
            string permSql = @"
UPDATE permisos
SET opcionespermitidas = TRIM(opcionespermitidas) || ',AgHorariosAcademicos'
WHERE perfil = 'Admin'
  AND (',' || REPLACE(opcionespermitidas, ' ', '') || ',') NOT LIKE '%,AgHorariosAcademicos,%';";
            connection.Execute(permSql, transaction: transaction);

            // 4. Siembra Canónica de Bloques Horarios (γ) - Plan v7 (Martes a Domingo)
            string seedBloquesSql = @"
DELETE FROM ag_bloque_horario WHERE subunidad = '';
INSERT INTO ag_bloque_horario (dia_semana, hora_inicio, hora_fin, duracion_horas, franja, activo, subunidad) VALUES
(2, '18:00', '22:00', 4.0, 'NOCHE', 1, ''),
(3, '18:00', '22:00', 4.0, 'NOCHE', 1, ''),
(4, '18:00', '22:00', 4.0, 'NOCHE', 1, ''),
(5, '18:00', '22:00', 4.0, 'NOCHE', 1, ''),
(6, '08:00', '12:00', 4.0, 'MANANA', 1, ''),
(6, '13:00', '17:00', 4.0, 'TARDE', 1, ''),
(7, '08:00', '12:00', 4.0, 'MANANA', 1, '');";

            connection.Execute(seedBloquesSql, transaction: transaction);

            // 5. Siembra Canónica de Aulas si la tabla está vacía (Fase 221: aisladas con subunidad 'DEMO')
            int countAulas = connection.ExecuteScalar<int>("SELECT COUNT(*) FROM ag_aula;", transaction: transaction);
            if (countAulas == 0)
            {
                string seedAulasSql = @"
INSERT INTO ag_aula (idsede, codigo, nombre, capacidad, tipo_aula, es_virtual, estado, subunidad) VALUES
(1, 'A-101', 'Aula Estándar 101', 40, 'ESTANDAR', 0, 1, 'DEMO'),
(1, 'A-102', 'Aula Estándar 102', 35, 'ESTANDAR', 0, 1, 'DEMO'),
(1, 'A-201', 'Aula Magistral 201', 50, 'MAGISTRAL', 0, 1, 'DEMO'),
(1, 'LAB-1', 'Laboratorio de Cómputo 1', 30, 'SISTEMAS', 0, 1, 'DEMO'),
(1, 'VIR-01', 'Sala Virtual Institucional', 100, 'VIRTUAL', 1, 1, 'DEMO');";

                connection.Execute(seedAulasSql, transaction: transaction);
            }

            // 6. Alineación de Esquema Defensiva e Idempotente (Fase 81)
            string[] alterStatements = new[]
            {
                // ag_aula
                "ALTER TABLE ag_aula ADD COLUMN idperiodo INTEGER DEFAULT 0;",
                "ALTER TABLE ag_aula ADD COLUMN semestre TEXT DEFAULT '';",
                "ALTER TABLE ag_aula ADD COLUMN modalidadVinculacion TEXT DEFAULT 'Presencial';",
                "ALTER TABLE ag_aula ADD COLUMN fecha_inicio TEXT DEFAULT '';",
                "ALTER TABLE ag_aula ADD COLUMN fecha_fin TEXT DEFAULT '';",
                "ALTER TABLE ag_aula ADD COLUMN semestre_inicio TEXT DEFAULT '';",
                "ALTER TABLE ag_aula ADD COLUMN semestre_fin TEXT DEFAULT '';",
                "ALTER TABLE ag_aula ADD COLUMN ventana_origen TEXT DEFAULT '';",
                "ALTER TABLE ag_aula ADD COLUMN aula_ref TEXT DEFAULT '';",
                // ag_vinculaciones
                "ALTER TABLE ag_vinculaciones ADD COLUMN total_horas INTEGER DEFAULT 0;",
                "ALTER TABLE ag_vinculaciones ADD COLUMN isTerminalAereo INTEGER DEFAULT 0;",
                "ALTER TABLE ag_vinculaciones ADD COLUMN idTerminales INTEGER DEFAULT 0;",
                "ALTER TABLE ag_vinculaciones ADD COLUMN gastoTerminalAerea REAL DEFAULT 0.0;",
                "ALTER TABLE ag_vinculaciones ADD COLUMN id_corrida INTEGER DEFAULT 0;",
                "ALTER TABLE ag_vinculaciones ADD COLUMN id_individuo INTEGER DEFAULT 0;",
                "ALTER TABLE ag_vinculaciones ADD COLUMN grupo TEXT DEFAULT '';",
                "ALTER TABLE ag_vinculaciones ADD COLUMN dia_horas_semanales TEXT DEFAULT '';",
                "ALTER TABLE ag_vinculaciones ADD COLUMN modalidadVinculacion TEXT DEFAULT '';",
                // ag_gen
                "ALTER TABLE ag_gen ADD COLUMN dia_semana INTEGER DEFAULT 1;",
                "ALTER TABLE ag_gen ADD COLUMN dentro_de_ventana INTEGER DEFAULT 1;",
                // ag_bitacora
                "ALTER TABLE ag_bitacora ADD COLUMN peor_fitness REAL DEFAULT 0;",
                "ALTER TABLE ag_bitacora ADD COLUMN duracion_ms INTEGER DEFAULT 0;",
                "ALTER TABLE ag_bitacora ADD COLUMN tiempo_ms INTEGER DEFAULT 0;",
                // ag_corrida
                "ALTER TABLE ag_corrida ADD COLUMN idaula INTEGER DEFAULT 0;",
                // ag_promocion
                "ALTER TABLE ag_promocion ADD COLUMN idaula INTEGER DEFAULT 0;",
                "ALTER TABLE ag_promocion ADD COLUMN idagvinculacion INTEGER DEFAULT 0;",
                "ALTER TABLE ag_promocion ADD COLUMN idvinculacion INTEGER DEFAULT 0;",
                "ALTER TABLE ag_promocion ADD COLUMN idtercero INTEGER DEFAULT 0;",
                "ALTER TABLE ag_promocion ADD COLUMN idasignatura INTEGER DEFAULT 0;",
                "ALTER TABLE ag_promocion ADD COLUMN fecha_inicio TEXT DEFAULT '';",
                "ALTER TABLE ag_promocion ADD COLUMN fecha_fin TEXT DEFAULT '';",
                "ALTER TABLE ag_promocion ADD COLUMN modo TEXT DEFAULT 'COPIA_REMAPEO';",
                "ALTER TABLE ag_promocion ADD COLUMN justificacion_vacante TEXT DEFAULT '';",
                // ag_aula_asignatura (Fase 122)
                "ALTER TABLE ag_aula_asignatura ADD COLUMN horas_objetivo REAL DEFAULT 0.0;",
                "ALTER TABLE ag_aula_asignatura ADD COLUMN orden_curricular INTEGER DEFAULT 0;",
                "ALTER TABLE ag_aula_asignatura ADD COLUMN requerida INTEGER DEFAULT 1;"
            };

            foreach (var alterSql in alterStatements)
            {
                try
                {
                    connection.Execute(alterSql, transaction: transaction);
                }
                catch
                {
                    // Ignora error si la columna ya existe en la base de datos (idempotente)
                }
            }
        }
    }
}
