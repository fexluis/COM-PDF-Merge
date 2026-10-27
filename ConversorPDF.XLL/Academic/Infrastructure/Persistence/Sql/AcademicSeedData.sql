-- Datos Semilla Corregidos para el Entorno de Pruebas Academicas
-- Cuidado con los nombres de columnas para que coincidan con AcademicSchema.sql

-- 1. SEDES (ciudad es la columna de visualización)
INSERT OR IGNORE INTO sedes (id, iddivipola, departamento, ciudad, subsede, categoria, transporte) 
VALUES (1, 1006, 'VALLE', 'CALI', NULL, 1, 0.0);

-- 2. PROGRAMAS (nombreprograma)
INSERT OR IGNORE INTO programas (id, nombreprograma, abreviatura, modalidad) 
VALUES (1, 'ADMINISTRACION PUBLICA TERRITORIAL', 'APT', 'PREGRADO');

-- 3. ASIGNATURAS (nombreasignatura)
INSERT OR IGNORE INTO asignaturas (id, idprograma, semestre, cod_materia, abreviatura, nucleo, nombre, creditos, int_horaria, tipoasignatura, tipocredito, prerequisito) 
VALUES (1, 1, 'I', 'APT01-CRD', 'FDEPU', 'E P', 'Fundamentos de economía pública', 3, 48, 'FUNDAMENTACIÓN', 'Teórico', 'N. A.');

-- 4. TERCEROS (cedula)
INSERT OR IGNORE INTO terceros (id, departamento, nombre, cedula, tipo, categoria) 
VALUES (1, 'VALLE', 'LUIS BECHARA', '1098765432', 'PLANTA', 'PROFESOR');

-- 5. PERIODOS
INSERT OR IGNORE INTO periodos (idperiodo, periodo, activo) VALUES 
(1, '2026-01', 1),
(2, '2026-02', 0);

-- 6. VINCULACIONES (Eliminado en v1.2.51 para evitar conflictos de persistencia con datos de usuario)
-- 7. SESIONES (Eliminado en v1.2.51 para evitar "revival" de sesiones borradas)
