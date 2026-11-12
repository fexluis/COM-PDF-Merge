# Handoff: Sincronización del Selector de Corridas del Módulo R (R-041)

**Fecha**: 2026-11-12  
**Plan asociado**: `.agents/plans/2026-11-12-sincronizacion-corridas-modulo-r-plan.md`  
**Registro**: Fila `R-041` en `.agents/plans/REGISTRY.md` (Ronda 319–325, Fases 1606–1612)  
**Estado inicial**: `executing`

---

## 1. Contexto y Problema

En Módulo R (`AgFormGeneradorHorarios`), tras ejecutar o cargar una corrida genética, se evidenciaba una desincronización entre:
- El selector `cmbCargarCorridas` (mostraba p. ej. `#60`)
- La etiqueta de estado `lblAgModo` (mostraba texto genérico o `#60`)
- El grid de resultados y staging (mostraba `#62`)

### Causas Raíz Identificadas
1. **DropDown bloqueado por caché**: El evento `DropDown` abortaba si la clave de filtros no cambiaba (`ClaveFiltroCorridas() == _claveFiltroCorridas`), impidiendo detectar nuevas corridas creadas por el propio AG o por otra ventana.
2. **Ausencia de refresco tras ejecución**: Al finalizar la ejecución del AG, no se invocaba el refresco del combo ni se actualizaba `lblAgModo` con `#N`.
3. **Fallback destructor a índice 0**: `AgCorridaSelector.ResolverSeleccionInicial` retornaba `CrearTodasLasCorridas()` si la corrida activa no aparecía en el catálogo filtrado. Al no pertenecer este ítem a `Items`, `items.IndexOf()` devolvía `-1`, degradando forzosamente a índice `0` (`SelectedIndex = 0`), lo cual disparaba `CargarCorridaDesdeComboAsync()` y borraba los resultados activos para cargar silenciosamente otra corrida ajena.

---

## 2. Decisiones de Diseño

1. **Política Pura `ResolverSeleccionModuloR`**: Separada de `ResolverSeleccionInicial` (que sigue sirviendo a `FormVinculaciones` con su comodín). Retorna `AgSeleccionModuloR` con la corrida seleccionada, índice, bandera de autorización de carga inicial y detección de activa ausente.
2. **Preservación Estricta de la Corrida Activa**: Si la corrida activa no está en la lista tras un refresco, `SelectedIndex = -1`, se conserva el resultado en pantalla y se notifica en `lblAgModo`. NUNCA se degrada a 0 ni se cargan resultados ajenos.
3. **Refresco Fresco y Concurrente**: `DropDown` siempre solicita datos frescos. Se controlan versiones y solicitudes pendientes para evitar que respuestas lentas u obsoletas pisen la selección actual.
4. **Ciclo de Vida Limpio**: En `IniciarOptimizacionAsync`, el refresco se ejecuta en `finally` si se creó/modificó cabecera, con token independiente. Si `CargarModuloResultadosAsync` falla, se limpia `_idCorridaCargada` y las vistas.

---

## 3. Olas de Ejecución

- **Wave 1 (Ronda 319 / Fase 1606)**: T1 (`integrator`, L0) - Esqueleto `AgSeleccionModuloR.cs`, registro en `ConversorPDF.XLL.csproj` y stub `ResolverSeleccionModuloR`.
- **Wave 2 (Ronda 320 / Fase 1607)**: T2 (`tdd-red-writer`) - Caracterización y pruebas unitarias RED en `AgCorridaSelectorTests.cs`.
- **Wave 3 (Ronda 321 / Fase 1608)**: T3 (`implementer`) - Implementación `ResolverSeleccionModuloR` y enlace en `AgFormGeneradorHorarios.cs` (GREEN).
- **Wave 4 (Ronda 322 / Fase 1609)**: T4 (`implementer`) - Refresco seguro en DropDown y concurrencia.
- **Wave 5 (Ronda 323 / Fase 1610)**: T5 (`implementer`) - Ciclo de vida post-ejecución, `lblAgModo` y retorno de `FormVinculaciones`.
- **Wave 6 (Ronda 324 / Fase 1611)**: T6 (`build-verifier`) - Verificación integral Full (0 regresiones).
- **Wave 7 (Ronda 325 / Fase 1612)**: T7 (`implementer`) - Documentación técnica del cambio.
