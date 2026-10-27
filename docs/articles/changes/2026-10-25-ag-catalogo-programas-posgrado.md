# Catálogo de programas AG alineado con Vinculaciones — Fase 1466

## Purpose
El combo `cmbProgramas` del generador AG se poblaba con `WHERE activo = 1`. En la base
institucional solo 2 programas están activos y ambos son PREGRADO, por lo que al seleccionar
`POSGRADO` el combo quedaba vacío (solo el pseudo-ítem) y el árbol no mostraba nodos de programa.
Vinculaciones ya lista todos los programas; AG queda alineado.

## Changes
1. **`Academic/Application/Services/AgProgramaCatalogoSql.cs`** (nuevo): SQL canónico
   `SeleccionProgramas`, sin filtro `activo`, con activos primero y `COALESCE(activo,0)`.
2. **`Academic/UI/Forms/AgFormGeneradorHorarios.cs`**: `CargarCatalogosAsync` usa el SQL canónico;
   se agrega traza `ProgramasPorModalidad` y advertencia si el catálogo llega vacío.
3. **Versión**: `1.6.3.76` en `ConversorPDF.XLL.csproj` y `Core/Helpers/App.cs`.

## How to use
Ribbon → AG → Generador de horarios. En «Tipo de programa» seleccionar `POSGRADO`:
el combo de programas lista los posgrados (p. ej. `ESP-ADE-614`, `MAP-391`, `MDH-409`).
`TODOS` lista el catálogo completo con los activos primero.

## Test coverage
- `AgTipoProgramaCatalogTests.Filtrar_WhenProgramasInactivos_ShouldIncludeThemRegardlessOfActivoFlag`
- `AgProgramaCatalogoSqlTests` (3 casos): catálogo completo, filtro POSGRADO con inactivos y orden.
