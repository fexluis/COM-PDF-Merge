# FormTerceroAsignaturas: catálogo completo de programas — Fase 1470

## Purpose
El combo `cmbPrograma` de la ventana modal `FormTerceroAsignaturas` cargaba los programas con `WHERE activo = 1`. En la base de datos institucional únicamente 2 programas figuran con `activo = 1` (ambos de nivel pregrado), lo que imposibilitaba registrar o editar asignaturas habilitadas pertenecientes a programas de posgrado (19 programas). Con este ajuste, el combo pasa de 2 a 21 programas académicos utilizando la consulta canónica compartida con el AG (`AgProgramaCatalogoSql.SeleccionProgramas`).

## Changes
1. **`Academic/UI/Forms/FormTerceroAsignaturas.cs`**:
   - `CargarProgramas()`: sustitución del SQL propio con filtro restrictivo por `AgProgramaCatalogoSql.SeleccionProgramas`. Se agrega desuscripción/suscripción al evento `Format` (`CmbPrograma_Format`) y registro en log de conteo por nivel académico (`pregrado`, `posgrado`).
   - `CmbPrograma_Format`: formatea cada elemento del combo como `"ABREV [NIVEL]"` para distinguir claramente los programas de pregrado y posgrado.
   - `CargarDatos()`: se implementa blindaje y fallback para programas históricos que no figuren en el catálogo actual, insertando un elemento sintético para preservar la edición del registro sin excepciones.
2. **`ConversorPDF.Tests/Academic/Application/Services/AgProgramaCatalogoSqlTests.cs`**:
   - Normalización de cadena de conexión a SQLite en memoria (`New=True`).

## How to use
1. En el módulo Académico (o desde el Generador de Horarios / Vinculaciones), acceder a la edición de terceros y abrir «Habilitar Nueva Asignatura» o editar una existente.
2. Desplegar el combo `cmbPrograma`: se mostrará el catálogo completo de 21 programas formateados con su abreviatura y nivel académico entre corchetes (ej. `APT-53 [PREGRADO]`, `MDH-409 [POSGRADO]`).

## Test coverage
- `ConversorPDF.Tests.Academic.Application.Services.AgProgramaCatalogoSqlTests` (4 casos de prueba verificando catálogo completo, ordenación, mapeo seguro ante nulos y soporte de programas inactivos).
- `ConversorPDF.Tests.Academic.Application.Services.AgTipoProgramaCatalogTests` (8 casos de prueba sobre tipificación y niveles académicos).
