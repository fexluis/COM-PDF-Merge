# AgFormTerceros multi-territorial y catálogo canónico de territoriales — Fase 1471

## Purpose
En el generador de horarios por Algoritmo Genético, `AgFormTerceros` aplicaba el filtro de subunidad activa directamente en la consulta de carga a la tabla `terceros` (`subunidad = @Sub OR subunidad = '' OR subunidad IS NULL`), restringiendo artificialmente la lista visible (por ejemplo, en la territorial VALLE solo se mostraban 93 de los 199 terceros activos). Además, el combo de territoriales se alimentaba de valores crudos de `terceros.departamento` sin contemplar asignaciones compuestas (`BOLIVAR,VALLE`), sin consultar la tabla institucional `territoriales` y sin permitir filtrar por `subunidad`.

Esta fase amplía `AgFormTerceros` para cargar el padrón completo de terceros (multi-territorial), introduce el servicio `AgTerritorialCatalog`, añade la columna «Territorial» a la grilla y discrimina las habilitaciones locales de las habilitaciones en otras subunidades, manteniendo intacto el pool de generación del AG.

## Changes
1. **`Academic/Application/Services/AgTerritorialCatalog.cs`** (nuevo):
   - Separa cadenas compuestas delimitadas por comas.
   - Construye la lista canónica y unificada para el combo territorial combinando el catálogo oficial (`territoriales.subunidad`) y los valores presentes en terceros (`departamento` y `subunidad`).
2. **`Academic/Application/DTOs/AgTerceroSeleccionDto.cs`**:
   - Incorpora `HabilitacionesLocales` y `HabilitacionesOtras`.
   - Adapta `TotalHabilitaciones` y `TieneAsignaturasHabilitadas` manteniendo compatibilidad binaria y semántica (únicamente se consideran habilitadas aquellas asignaturas de la subunidad activa).
3. **`Academic/Application/Services/AgTerceroFiltro.cs`**:
   - Filtrado territorial contra `Departamento` y `Subunidad` reconociendo valores compuestos.
   - Soporte para la constante `AgTerritorialCatalog.Todas`.
   - Búsqueda textual incluyendo `Subunidad` y `Departamento`.
   - `soloHabilitadas` evalúa `TieneAsignaturasHabilitadas` (habilitación local).
4. **`Academic/UI/Forms/AgFormTerceros.cs`**:
   - `OnLoad`: adapta la etiqueta del checkbox para indicar explícitamente la subunidad activa (`Solo asignaturas habilitadas en {subActiva}`).
   - `CargarDatosAsync`: carga todos los terceros activos (excluyendo VACANTE y cédula 1111111111); calcula habilitaciones locales y de otras territoriales mediante dos consultas agrupadas indexadas; puebla el combo mediante `AgTerritorialCatalog.Construir`.
   - `ActualizarConteo`: muestra el desglose de `habilitados locales` y `solo en otras`.
   - Traza en log (`LoggerHelper.Warning`) al presionar Aceptar si existen docentes seleccionados sin habilitación en la territorial activa.
5. **`Academic/UI/Forms/AgFormTerceros.Designer.cs`**:
   - Agrega columna `colTerritorial` enlazada a `Subunidad` (120 px, solo lectura).
6. **`ConversorPDF.XLL.csproj`**:
   - Registro de `AgTerritorialCatalog.cs`.

## How to use
1. En el Generador de Horarios AG, pulsar el botón «Terceros».
2. La grilla mostrará los 199 terceros activos.
3. El combo «Territorial» contiene las 17 subunidades oficiales más la opción `TODAS LAS TERRITORIALES`.
4. Al seleccionar una territorial (p. ej. `VALLE` o `BOLIVAR`), docentes con asignación compuesta como `BOLIVAR,VALLE` aparecerán en ambas opciones.
5. Activar «Solo asignaturas habilitadas en VALLE» para acotar a los docentes con habilitaciones activas en la subunidad local.
6. La columna «Territorial» muestra la subunidad correspondiente de cada tercero.

> [!NOTE]
> El pool de generación del algoritmo genético continúa estrictamente acotado a la subunidad activa en `AgFormGeneradorHorarios` (L912-936). Seleccionar terceros de otra territorial que no tengan habilitación local no los incorporará al pool de cruce del AG y se reportará en el log como advertencia informativa.

## Test coverage
- `ConversorPDF.Tests.Academic.Application.Services.AgTerritorialCatalogTests` (5 casos: separación, vacíos, unión sin duplicados, normalización).
- `ConversorPDF.Tests.Academic.Application.Services.AgTerceroFiltroTests` (4 nuevos casos: subunidades compuestas, Todas, coincidencia en subunidad, exclusión de solo otras).
- `ConversorPDF.Tests.Academic.Application.DTOs.AgTerceroSeleccionDtoTests` (3 casos: propiedades locales, otras y total).
