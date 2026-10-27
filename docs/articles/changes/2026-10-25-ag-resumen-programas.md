# Resumen de programas en filtros del Generador AG — Fase 1472

## Purpose
En `AgFormGeneradorHorarios`, el resumen de programas en el encabezado del grupo de selección evaluaba `progsSel.Count == 0 || progVisibles == 0 ? "Todos" : ...`. Si el usuario seleccionaba programas pero el filtro de modalidad (p. ej. `POSGRADO`) ocultaba todos los programas seleccionados (es decir, `progsSel.Count > 0` y `progVisibles == 0`), el texto mostraba engañosamente `Programas: Todos`, en abierta contradicción con el árbol vacío.

Esta fase corrige dicho comportamiento mediante el servicio puro `AgResumenFiltros` y cierra formalmente la Fase 1469.

## Changes
1. **`Academic/Application/Services/AgResumenFiltros.cs`** (nuevo):
   - Servicio puro `internal static` que encapsula la lógica de formateo del resumen de programas:
     - `0 seleccionados` → `"Todos"`
     - `visibles == seleccionados` → `"{seleccionados}"`
     - `visibles < seleccionados` → `"{visibles} (+{seleccionados - visibles} otra modalidad)"` (incluyendo `0 (+N otra modalidad)`).
2. **`Academic/UI/Forms/AgFormGeneradorHorarios.cs`**:
   - `ActualizarResumenFiltros()`: delega el cálculo de `progTxt` a `AgResumenFiltros.Programas(progsSel.Count, progVisibles)`.
3. **`ConversorPDF.XLL.csproj`**:
   - Registro de `AgResumenFiltros.cs`.

## How to use
1. En el Generador de Horarios AG, seleccionar programas de pregrado en `cmbProgramas`.
2. Cambiar «Tipo de programa» a `POSGRADO`.
3. El resumen de filtros en `grpSeleccion` mostrará `Programas: 0 (+N otra modalidad)` en lugar de `Programas: Todos`.

## Test coverage
- `ConversorPDF.Tests.Academic.Application.Services.AgResumenFiltrosTests`: prueba con 5 escenarios paramétricos (`(0,0)`, `(5,5)`, `(5,3)`, `(5,0)`, `(2,7)`).
