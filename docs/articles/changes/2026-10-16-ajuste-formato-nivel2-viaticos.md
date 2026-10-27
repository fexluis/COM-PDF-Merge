# Ajuste de Formato y Resaltado del Nivel 2 en Vista Viáticos

**Fecha:** 2026-10-16  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Ronda:** 132 (Fases 714 a 720)

## Purpose
Optimizar la legibilidad y presentación visual del nodo de Nivel 2 (Bloque de viático) en el `TreeView` de viáticos (`FormVerVinculaciones`), eliminando información redundante (`Cierre: `, código de ruta en la cadena de vista), compactando la unidad de duración (`días` → `d`), y destacando visualmente en color naranja oscuro (`Color.DarkOrange`) aquellos viajes que incorporan la regla de pernocta en el último día (`isPernoctaUltimoDia = 1`).

## Changes
1. **`ViaticoBloqueFormatter.cs`:**
   - Modificada la firma de `FormatearNodoNivel2`: se retira el parámetro `codRuta`.
   - Modificado el template de cadena:
     `📅 [FechaInicio] - [FechaFin] | [Duración] d | [Aéreo/Terrestre] | [Ruta] | $[ValorTransporte] | [DIA HORA] | [Pernocta Ultimo Dia]`
   - Se elimina el prefijo literal `Cierre: `.
   - Se mantiene intacto `ObtenerTextoCierre` y `ObtenerRangoComision`.
2. **`ViaticoBloqueFormatterTests.cs`:**
   - Actualizadas pruebas existentes de `FormatearNodoNivel2` con y sin pernocta al nuevo formato exacto.
   - Añadidas pruebas unitarias `FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludeCierrePrefixNorCodigoRuta` y `FormatearNodoNivel2_WhenBloqueNull_ShouldReturnEmpty`.
3. **`FormVerVinculaciones.cs`:**
   - Se mantiene la llamada a `_trayectoEncoderService.GetOrCreateCodigoTrayecto(territorial, verificacion, transporteBase)` para asegurar la pre-creación en BD de rutas no catalogadas.
   - Se elimina la variable local `codRutaDisplay` y se invoca `FormatearNodoNivel2` con 6 argumentos.
   - Se condiciona el color de la fila: `ForeColor = bloque.TienePernoctaUltimoDia ? Color.DarkOrange : Color.DarkBlue`.
4. **`docs/articles/academic-viaticos.md`:**
   - Documentada la subsección descriptiva del Nivel 2 y el resaltado visual.

## How to use
1. En Excel, abrir la pestaña **Conversor Académico** → **Ver Vinculaciones**.
2. Cambiar a la vista **Viáticos**.
3. Observar los nodos de Nivel 2:
   - Bloques estándar se muestran en azul oscuro (`DarkBlue`) con duración abreviada en `d` y horario sin prefijo `Cierre:`.
   - Bloques con pernocta en último día se resaltan en naranja oscuro (`DarkOrange`) con el sufijo `| Pernocta Ultimo Dia`.
4. Al exportar a `FV-FO-033`, la columna BL (`COD_RUTA`) sigue completándose adecuadamente a partir de los datos en base de datos.

## Test Coverage
- `ViaticoBloqueFormatterTests.cs`:
  - `FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludePernoctaUltimoDiaSuffix`: Valida formato exacto sin pernocta ni código.
  - `FormatearNodoNivel2_WhenConPernocta_ShouldIncludePernoctaUltimoDiaSuffix`: Valida sufijo de pernocta y fecha extendida.
  - `FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludeCierrePrefixNorCodigoRuta`: Valida ausencia de `Cierre:` y `SIN-COD`.
  - `FormatearNodoNivel2_WhenBloqueNull_ShouldReturnEmpty`: Valida manejo robusto de nulo.
