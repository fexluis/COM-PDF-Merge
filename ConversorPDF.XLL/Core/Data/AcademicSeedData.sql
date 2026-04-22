-- Datos Semilla Corregidos para el Entorno de Pruebas Academicas
-- Cuidado con los nombres de columnas para que coincidan con AcademicSchema.sql

-- 1. SEDES (ciudad es la columna de visualización)
INSERT OR IGNORE INTO sedes (id, departamento, ciudad, categoria, transporte) 
VALUES (1, 'VALLE', 'CALI', 'SEDE CENTRAL', 0.0);

-- 2. PROGRAMAS (nombreprograma)
INSERT OR IGNORE INTO programas (id, nombreprograma, abreviatura, tipo) 
VALUES (1, 'ADMINISTRACION PUBLICA TERRITORIAL', 'APT', 'PREGRADO');

-- 3. ASIGNATURAS (nombreasignatura)
INSERT OR IGNORE INTO asignaturas (id, idprograma, semestre, código, abreviatura, área, nombreasignatura, creditos, tipoasignatura, tipocredito, prerequisito) 
VALUES (1, 1, 'I', 'APT01-CRD', 'FDEPU', 'E P', 'Fundamentos de economía pública', 3, 'FUNDAMENTACIÓN', 'Teórico', 'N. A.');

-- 4. TERCEROS (cedula)
INSERT OR IGNORE INTO terceros (id, departamento, nombre, cedula, tipo, rol) 
VALUES (1, 'VALLE', 'LUIS BECHARA', '1098765432', 'PLANTA', 'PROFESOR');

-- 5. VINCULACIONES (Eliminado en v1.2.51 para evitar conflictos de persistencia con datos de usuario)
-- 6. SESIONES (Eliminado en v1.2.51 para evitar "revival" de sesiones borradas)
