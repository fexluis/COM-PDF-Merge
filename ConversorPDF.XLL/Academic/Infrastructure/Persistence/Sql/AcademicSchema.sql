CREATE TABLE IF NOT EXISTS __SchemaVersion (Version INTEGER PRIMARY KEY);
-- Schema Base v28: Contiene el esquema completo incluyendo tablas de sync V24, SIIF V26/V27 y CDP V28. Salta migraciones V2-V28.
INSERT OR IGNORE INTO __SchemaVersion (Version) VALUES (28);

CREATE TABLE IF NOT EXISTS festivos (
     idfestivo INTEGER PRIMARY KEY AUTOINCREMENT,
     idperiodo INTEGER,
     fechafestivoinicio TEXT,
     fechafestivofinal TEXT,
     detallefestivo TEXT,
     tipofestivo TEXT DEFAULT 'FESTIVO',  -- adicionado, rango períodos 'FESTIVO' / 'PRESENCIAL'
     subunidad TEXT DEFAULT '',
     FOREIGN KEY(idperiodo) REFERENCES periodos(idperiodo)
);

CREATE TABLE IF NOT EXISTS programas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nombreprograma TEXT,
    abreviatura TEXT,
    pensum INTEGER, -- adicionado
    actual INTEGER DEFAULT 0, -- adicionado, 1 si es el actual, 0 si no lo es
    activo INTEGER DEFAULT 0, -- adicionado, 1 si esta activo, 0 si no lo es
    modalidad TEXT, -- reemplaza a tipo ("PREGRADO" y "POSGRADO")
    subunidad TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS asignaturas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idprograma INTEGER, 
    semestre TEXT,
    num_nivel INTEGER, -- adicionado, numero de semestre
    cod_unidad TEXT,
    cod_materia TEXT, -- reemplaza a codigo
    uni_teorica INTEGER, -- adicionado, numero de unidad
    metodologia TEXT,
    abreviatura TEXT,
    nucleo TEXT,
    nombre TEXT,
    creditos INTEGER,
    int_horaria INTEGER, -- reemplaza a total_horas
    tipoasignatura TEXT,
    tipocredito TEXT,
    prerequisito TEXT,
    subunidad TEXT DEFAULT '',
    FOREIGN KEY(idprograma) REFERENCES programas(id)
);

CREATE TABLE IF NOT EXISTS microcurriculo (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idasignatura INTEGER,
    nucleo TEXT,
    objetivogeneral TEXT,
    competencias TEXT,
    subunidad TEXT DEFAULT '',
    FOREIGN KEY(idasignatura) REFERENCES asignaturas(id)
);

CREATE TABLE IF NOT EXISTS contenidos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idmicrocurriculo INTEGER,
    unidad TEXT,
    contenido TEXT,
    subunidad TEXT DEFAULT '',
    FOREIGN KEY(idmicrocurriculo) REFERENCES microcurriculo(id)
);

CREATE TABLE IF NOT EXISTS terceros (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    departamento TEXT,
    nombre TEXT,
    cedula TEXT,
    tipo TEXT,
    salario REAL,
    categoria TEXT, -- reemplaza a rol
    entidad_bancaria TEXT,
    cuenta_bancaria TEXT,
    tipo_cuenta_bancaria TEXT,
    telefono TEXT,
    correo TEXT,
    correo_personal TEXT,
    fecha_nacimiento TEXT,
    contrato TEXT,
    residencia TEXT,
    id_residencia_divipola INTEGER NULL,
    pensionado INTEGER DEFAULT 0, -- adicionado, 1 si es pensionado, 0 si no lo es
    activo INTEGER DEFAULT 1, -- 1 si esta activo, 0 si no lo es
    subunidad TEXT DEFAULT '',
    puntaje_ocasionales REAL DEFAULT 0
);

CREATE TABLE IF NOT EXISTS territoriales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    territorial TEXT, 
    subunidad TEXT
);

INSERT OR IGNORE INTO territoriales (id, territorial, subunidad) VALUES 
(1, 'ANTIOQUIA', 'ANTIOQUIA'),
(2, 'ATLÁNTICO - CESAR - MAGDALENA - LA GUAJIRA', 'ATLANTICO'),
(3, 'BOLIVAR - CÓRDOBA - SUCRE - SAN ANDRÉS', 'BOLIVAR'),
(4, 'BOYACÁ - CASANARE', 'BOYACA'),
(5, 'CALDAS', 'CALDAS'),
(6, 'CAUCA', 'CAUCA'),
(7, 'CHOCO', 'CHOCO'),
(8, 'CUNDINAMARCA', 'CUNDINAMARCA'),
(9, 'HUILA - CAQUETÁ - PUTUMAYO', 'HUILA'),
(10, 'META - GUAVIARE - GUAINÍA - VAUPÉS - VICHADA - AMAZONAS', 'META'),
(11, 'NARIÑO - ALTO PUTUMAYO', 'NARINO'),
(12, 'NORTE DE SANTANDER - ARAUCA', 'NORTE'),
(13, 'QUINDÍO - RISARALDA', 'RISARALDA'),
(14, 'SANTANDER', 'SANTANDER'),
(15, 'TOLIMA', 'TOLIMA'),
(16, 'VALLE', 'VALLE'),
(17, 'ESAP CENTRAL', 'ESAPCENTRAL');

CREATE TABLE IF NOT EXISTS sedes (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    iddivipola INTEGER, -- adicionado
    departamento TEXT,
    ciudad TEXT,
    departamentoSiif TEXT, -- adicionado Fase 26
    ciudadSiif TEXT, -- adicionado Fase 26
    subsede TEXT, -- adicionado
    categoria INTEGER DEFAULT 2, -- adicionado, 1 o 2 
    transporte REAL,
    activo INTEGER DEFAULT 1, -- adicionado, 1 si esta activa, 0 si no lo es
    subunidad TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS BaseGastosTerrestre (
    Id INTEGER PRIMARY KEY AUTOINCREMENT,
    Territorial TEXT NOT NULL,
    CodigoTrayecto TEXT NOT NULL UNIQUE,
    Trayecto TEXT NOT NULL,
    Valor REAL
);

CREATE TABLE IF NOT EXISTS periodos (
    idperiodo INTEGER PRIMARY KEY,
    periodo TEXT NOT NULL,
    periodo_inicio TEXT,
    periodo_fin TEXT,
    activo INTEGER DEFAULT 0,
    subunidad TEXT DEFAULT ''
);

CREATE TABLE IF NOT EXISTS vinculaciones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idsede INTEGER,
    idtercero INTEGER,
    idprograma INTEGER,
    idperiodo INTEGER,
    grupo TEXT,
    semestre TEXT,
    idasignatura INTEGER,
    total_horas INTEGER,
    dia_horas_semanales TEXT,
    fecha_inicio TEXT,
    fecha_fin TEXT,
    modalidadVinculacion TEXT,
    subunidad TEXT DEFAULT '',
    isPernoctaUltimoDia INTEGER DEFAULT 0, -- 0 si No Pernocta ultimo día, 1 si Pernocta ultimo día y suma 1 día a la comisión de viáticos 
    isTerminalAereo INTEGER DEFAULT 0, -- 0 si es 'Terrestre', 1 si es 'Aéreo' 
    idTerminales INTEGER DEFAULT 0, -- id = 0 si isTerminalAereo = 0 ('NO APLICA')
    gastoTerminalAerea REAL DEFAULT 0, -- 'maxvalor' de tabla 'terminales' si isTerminalAereo = 1, 0 si isTerminalAereo = 0
    FOREIGN KEY(idsede) REFERENCES sedes(id),
    FOREIGN KEY(idtercero) REFERENCES terceros(id),
    FOREIGN KEY(idprograma) REFERENCES programas(id),
    FOREIGN KEY(idperiodo) REFERENCES periodos(idperiodo),
    FOREIGN KEY(idasignatura) REFERENCES asignaturas(id)
);

-- Migracion segura: si la columna ya existe SQLite retorna un error que
-- SQLiteEngineBootstrap ignora; si no existe, la crea correctamente.
-- idtercero e idperiodo ya están en el CREATE TABLE.

CREATE TABLE IF NOT EXISTS sesiones (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idvinculacion INTEGER,
    idtercero INTEGER,
    fechasesion TEXT,
    horariosesion TEXT,
    horas REAL,
    viatica INTEGER,
    subunidad TEXT DEFAULT '',
    FOREIGN KEY(idvinculacion) REFERENCES vinculaciones(id),
    FOREIGN KEY(idtercero) REFERENCES terceros(id)
);

CREATE TABLE IF NOT EXISTS salvedades_conflictos (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idsesion_a INTEGER NOT NULL,
    idsesion_b INTEGER NOT NULL,
    idtercero INTEGER NOT NULL,
    idperiodo INTEGER,
    justificacion TEXT,
    fecha_creacion TEXT DEFAULT CURRENT_TIMESTAMP,
    subunidad TEXT DEFAULT '',
    UNIQUE(idsesion_a, idsesion_b)
);

CREATE INDEX IF NOT EXISTS idx_salvedades_periodo_tercero 
ON salvedades_conflictos(idperiodo, idtercero);

-- 2A. INSERCION DE DATOS: PERIODOS
INSERT OR IGNORE INTO periodos (idperiodo, periodo, periodo_inicio, periodo_fin, activo) VALUES 
(1, '2026-01', '2026-01-12', '2026-06-19', 0),
(2, '2026-02', '2026-08-15', '2026-12-12', 1),
(3, '2027-01', '2027-01-01', '2027-06-30', 0);

-- 3. INSERCION DE DATOS: SEDES (ESAP Context)
INSERT OR IGNORE INTO sedes (iddivipola, departamento, ciudad, departamentoSiif, ciudadSiif, subsede, categoria, transporte, activo, subunidad) VALUES
(1006,'VALLE','CALI','VALLE','CALI', NULL, 1, 90000.0, 1, 'VALLE'),
(1018, 'VALLE','CARTAGO','VALLE','CARTAGO', NULL, 2, 92000.0, 1, 'VALLE'),
(1012, 'VALLE','BUENAVENTURA','VALLE','BUENAVENTURA', NULL, 1, 122000.0, 1, 'VALLE'),
(1013, 'VALLE','GUADALAJARA DE BUGA','VALLE','BUGA', NULL, 1, 70000.0, 1, 'VALLE'),
(1013, 'VALLE','GUADALAJARA DE BUGA','VALLE','BUGA','SEVILLA', 1, 70000.0, 1, 'VALLE'),
(1013, 'VALLE','GUADALAJARA DE BUGA','VALLE','BUGA','FLORIDA-PRADERA', 1, 70000.0, 1, 'VALLE'),
(1015, 'VALLE','CAICEDONIA','VALLE','CAICEDONIA', NULL, 2, 120000.0, 1, 'VALLE'),
(1024, 'VALLE','FLORIDA','VALLE','FLORIDA', NULL, 2, 40000.0, 1, 'VALLE'),
(1024, 'VALLE','FLORIDA','VALLE','FLORIDA','PLATAFORMA', 2, 40000.0, 1, 'VALLE'),
(1029, 'VALLE','LA UNIÓN','VALLE','LA UNION', NULL, 2, 75000.0, 1, 'VALLE'),
(1033, 'VALLE','PRADERA','VALLE','PRADERA', NULL, 2, 40000.0, 1, 'VALLE'),
(1038, 'VALLE','SEVILLA','VALLE','SEVILLA', NULL, 2, 70000.0, 1, 'VALLE'),
(1041, 'VALLE','TULUÁ','VALLE','TULUA', NULL, 1, 50000.0, 1, 'VALLE'),
(1020, 'VALLE','EL ÁGUILA','VALLE','EL AGUILA', NULL, 2, 120000.0, 1, 'VALLE'),
(1006, 'VALLE','CALI','VALLE','CALI','PLATAFORMA', 1, 90000.0, 1, 'VALLE');

-- 8. FESTIVOS

INSERT OR IGNORE INTO festivos (idfestivo, idperiodo, fechafestivoinicio, fechafestivofinal, detallefestivo, tipofestivo) VALUES
(1, 2, '2026-07-20', '2026-07-20', 'Día de la Independencia de Colombia', 'FESTIVO'),
(2, 2, '2026-08-07', '2026-08-07', 'Batalla de Boyacá', 'FESTIVO'),
(3, 2, '2026-08-17', '2026-08-17', 'Asuncion de la Virgen', 'FESTIVO'),
(4, 2, '2026-10-05', '2026-10-11', 'Semana de Receso', 'FESTIVO'),
(5, 2, '2026-10-12', '2026-10-12', 'Día de la Raza', 'FESTIVO'),
(6, 2, '2026-11-02', '2026-11-02', 'Todos los Santos', 'FESTIVO'),
(7, 2, '2026-11-16', '2026-11-16', 'Independencia de CARTAGENA', 'FESTIVO'),
(8, 2, '2026-12-08', '2026-12-08', 'Inmaculada Concepcion', 'FESTIVO'),
(9, 2, '2026-12-25', '2026-12-25', 'Navidad', 'FESTIVO'),
(10, 2, '2026-08-18', '2026-08-23', 'Viaticos', 'PRESENCIAL'),
(11, 2, '2026-08-31', '2026-09-06', 'Viaticos', 'PRESENCIAL'),
(12, 2, '2026-09-14', '2026-09-20', 'Viaticos', 'PRESENCIAL'),
(13, 2, '2026-09-28', '2026-10-04', 'Viaticos', 'PRESENCIAL'),
(14, 2, '2026-10-19', '2026-10-25', 'Viaticos', 'PRESENCIAL'),
(15, 2, '2026-11-03', '2026-11-08', 'Viaticos', 'PRESENCIAL'),
(16, 2, '2026-11-17', '2026-11-22', 'Viaticos', 'PRESENCIAL'),
(17, 2, '2026-11-30', '2026-12-06', 'Viaticos', 'PRESENCIAL');

-- -- TABLAS DE SYNC (V24: config, usuarios, sync_control, sync_archive) ----

CREATE TABLE IF NOT EXISTS usuarios (
    id              TEXT PRIMARY KEY DEFAULT (lower(hex(randomblob(16)))),
    nombre          TEXT NOT NULL,
    nombreusuario   TEXT NOT NULL, -- Nombre de usuario OneDrive
    subunidad       TEXT NOT NULL,
    perfil          TEXT DEFAULT 'Admin', -- Rol del usuario: Admin, Academica, Contador, Pagador, Auditor
    activo          INTEGER DEFAULT 1,
    firma           TEXT DEFAULT NULL, -- Cadena Base64 completa de la imagen de la firma del usuario
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);

CREATE INDEX IF NOT EXISTS idx_usuarios_nombreusuario ON usuarios(nombreusuario);

CREATE TABLE IF NOT EXISTS permisos (
    perfil TEXT PRIMARY KEY NOT NULL,
    opcionespermitidas TEXT NOT NULL
);

INSERT OR IGNORE INTO permisos (perfil, opcionespermitidas) VALUES 
('Admin', 'ConvertirPDF,ConvertirObligacion,UnificarOB,UnificarOP,UnificarViaticos,MasivoViaticos,GenerarCumplidos,GenerarFilaSeleccionada,GestionDocumental,BuscarInformes,ProcesarFirma,SesionesAcademicas,Metadatos,Vinculaciones,Backups,Usuario,GestionUsuarios,Permisos'),
('Academica', 'ConvertirPDF,Vinculaciones,Backups,Usuario'),
('Contador', 'ConvertirPDF,ConvertirObligacion,Usuario'),
('Pagador', 'ConvertirPDF,UnificarOP,Usuario'),
('Auditor', 'ConvertirPDF,BuscarInformes,Usuario');

CREATE TABLE IF NOT EXISTS sync_control (
    subunidad       TEXT NOT NULL,
    device_id       TEXT NOT NULL,
    last_pushed_seq INTEGER DEFAULT 0,
    last_pulled_at  TEXT,
    PRIMARY KEY (subunidad, device_id)
);

CREATE TABLE IF NOT EXISTS sync_archive (
    filename        TEXT PRIMARY KEY,
    subunidad       TEXT NOT NULL,
    checksum        TEXT,
    processed_at    TEXT DEFAULT (datetime('now'))
);

-- 9. VISTAS DE CONSULTA CORREGIDAS

DROP VIEW IF EXISTS vw_VinculacionesDetalle;
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
    NULLIF(NULLIF(v.fecha_inicio, ''), '--') AS SesionFechaInicial,
    NULLIF(NULLIF(v.fecha_fin, ''), '--') AS SesionFechaFinal,
    v.total_horas AS SesionHoras,
    (SELECT CASE WHEN EXISTS (
        SELECT 1 FROM sesiones ses WHERE ses.idvinculacion = v.id AND ses.viatica = 1
    ) THEN 1 ELSE 0 END) AS SesionViatica,
    v.modalidadVinculacion AS Modalidad,
    v.subunidad AS Subunidad
FROM vinculaciones v
JOIN programas p ON v.idprograma = p.id
JOIN asignaturas a ON v.idasignatura = a.id
JOIN sedes s ON v.idsede = s.id
LEFT JOIN terceros t ON v.idtercero = t.id
LEFT JOIN periodos pe ON v.idperiodo = pe.idperiodo;

CREATE TABLE IF NOT EXISTS terminales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    terminal TEXT, 
    maxvalor REAL
);

INSERT OR IGNORE INTO terminales (id,terminal, maxvalor) VALUES 
(1,'NO APLICA', 0),
(2,'ANTIOQUIA (Rionegro)', 162634),
(3,'ATLANTICO (Soledad)', 130704),
(4,'CORDOBA (Los Garzones)', 118731),
(5,'MAGDALENA (Santa Marta)', 129708),
(6,'NORTE DE SANTANDER', 50689),
(7,'QUINDIO (La Tebaida)', 186581),
(8,'SANTANDER (Lebrija)', 186581),
(9,'SUCRE (Corozal)', 162634),
(10,'VALLE DEL CAUCA (Palmira)', 186581),
(11,'NARIÑO (Chachagüí)', 186581),
(12,'PUTUMAYO (Pto.Asís)', 93788),
(13,'OTRO', 50689);

DROP VIEW IF EXISTS vw_SesionesDetalle;
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

CREATE TABLE IF NOT EXISTS config (
    id                          INTEGER PRIMARY KEY,
    subunidad                   TEXT NOT NULL,
    id_territorial              INTEGER NOT NULL,
    id_periodo                  INTEGER NOT NULL,
    fecha_inicio                TEXT,
    fecha_fin                   TEXT,
    cdp                         TEXT,
    salario_decreto_613_2025    REAL,
    viatico_diario_tarifa_a     REAL,
    viatico_diario_tarifa_b     REAL,
    porcentaje_decreto_a        REAL,
    porcentaje_decreto_b        REAL,
    dt_nombre                   TEXT,
    dt_cargo                    TEXT,
    corporativo_cedula          TEXT,
    corporativo_nombre          TEXT,
    corporativo_cargo           TEXT,
    departamento_base           TEXT,
    ciudad_base                 TEXT,
    aeropuerto_base             TEXT,
    created_at TEXT DEFAULT (datetime('now')),
    updated_at TEXT DEFAULT (datetime('now'))
);

INSERT OR IGNORE INTO config (id, subunidad, id_territorial, id_periodo, fecha_inicio, fecha_fin, cdp, salario_decreto_613_2025, viatico_diario_tarifa_a, viatico_diario_tarifa_b, porcentaje_decreto_a, porcentaje_decreto_b, dt_nombre, dt_cargo, corporativo_cedula, corporativo_nombre, corporativo_cargo, departamento_base, ciudad_base, aeropuerto_base)
SELECT 1, 'VALLE', 16, 2, '14/05/2026', '21/05/2026', '10126', 5102609, 336470, 282580, 879578, 732547, 'Jacinto Pineda Jimenez', 'Director Territorial', '72170937', 'Jairo Alonso', 'Subdirector Corporativo', 'VALLE', 'CALI', 'VALLE DEL CAUCA (Palmira)'
WHERE NOT EXISTS (SELECT 1 FROM config);

CREATE TABLE IF NOT EXISTS cdpDependencias (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    idperiodo INTEGER,
    cdpnum TEXT, 
    fechacdp TEXT,
    rubro TEXT,
    responsable TEXT,
    cargo TEXT,
    elaboro TEXT,
    reviso TEXT,
    dependencia TEXT DEFAULT NULL,
    dependencia_full TEXT DEFAULT NULL,
    activo INTEGER DEFAULT 0, -- adicionado, 1 si esta activo, 0 si no lo es
    subunidad TEXT -- FASE 21: Agregado para soporte de sincronización multi-sede
);
