# Combo "Corrida" en Vinculaciones (MODO AG): filtros ocultos y corridas anteriores

**Fecha:** 2026-11-09 · **Plan:** `.agents/plans/2026-11-09-corridas-ag-formvinculaciones-plan.md` (R-007) · **Fases:** 1488–1495

---

## 1. Síntoma reportado

En `FormVinculaciones` (MODO AG) el estado mostraba *"hay sesiones en staging pero 0 coinciden con los filtros"*, el combo de corridas aparecía superpuesto sobre los botones y no era posible consultar corridas anteriores.

## 2. Diagnóstico (datos reales del vault `VALLE`, 2026-11-09)

| Dato | Valor |
|---|---|
| `config` | `id_periodo = 2 (2026-02)`, `subunidad = VALLE` |
| `ag_corrida` #46 | período 3 · BOLIVAR · estado **4 (Error)** · **0 vinc / 0 ses** |
| `ag_corrida` #44 | período 3 · VALLE · estado 2 · **23 vinc / 244 ses** |
| `ag_corrida` #37 | período 3 · VALLE · **0 / 0** |
| `ag_corrida` #1 | período **2** · VALLE · 18 / 145 |
| Huérfana | `ag_vinculaciones.id_corrida = 45` (BOLIVAR) con 3 sesiones **sin cabecera** en `ag_corrida` |

Prueba del defecto sobre `vw_AgSesionesDetalle`:

| Consulta | Filas |
|---|---|
| `IdCorrida = 44` | **244** |
| `IdCorrida = 44 AND IdPeriodo = 2` ← `config` | **0** |
| `IdCorrida = 44 AND IdPeriodo = 3` | 244 |
| `IdCorrida = 46` ← default del combo (`MAX(id)`) | **0** |

### Causas raíz

1. **Filtros ocultos**: `CargarCatalogosAsync` preseleccionaba `config.id_periodo` y `config.subunidad` en combos invisibles (`Location(-100,-100)`, `Visible=false`) y `EjecutarConsultaAsync` los aplicaba como filtros reales.
2. **Selección por `MAX(id)`**: el combo abría en #46 (otra territorial, programa y estado Error) en lugar de la corrida cargada en el Módulo R (#44).
3. **Lista incompleta**: se construía sólo desde `ag_corrida`, mostrando corridas vacías (#46, #37) y ocultando el staging huérfano (#45).
4. **Layout**: el combo se creaba en runtime en `pnlFilters.Width - 310`, tapando *Viáticos / Filtrar / Adicionar*.
5. **Mensaje no accionable**: los contadores usaban SQL crudo con otro criterio de estado y sin los filtros de la UI.

## 3. Solución implementada

| Fase | Cambio |
|---|---|
| 1488 | `CorridaStagingItem`, `AgStagingResumen` en `Domain/Models/Genetica`; `AgCorridaSelector`, `AgFiltrosPolicy` en `Application/Services` |
| 1489 | `GetCorridasStagingAsync` (`ag_vinculaciones` ∪ `ag_corrida`, con conteos) y `GetResumenStagingAsync` (consulta A sin filtros / B con filtros) en `VinculacionQueryRepository` |
| 1490 | `BuildWhere` único compartido por el listado y los contadores; delegación y población de `Etiqueta` en `VinculacionListService` (Application → Domain) |
| 1491 | En MODO AG no se aplican período/territorial de `config`; comparación `UPPER(TRIM(...))` para subunidad |
| 1492 | Selector en el **Designer** (fila 3), `idCorridaInicial` desde el Módulo R, refresco en `DropDown` sin `async void`, una sola consulta inicial |
| 1493 | Mensaje accionable + botón **«Ajustar filtros»**; guardia para «Todas las corridas» |
| 1494 | `pnlFilters` 80→104 px, controles en fila 3 (`12,78` / `75,75` w505 / `585,74` w170 / `765,74` w140); etiqueta **MODO AG (staging)** alineada al extremo derecho de la fila 3; Producción conserva 80 px y oculta los controles de corrida |
| 1495 | 23 pruebas nuevas (selector, política, integración con fixture real, delegación) |

### Etiqueta del combo

```
#44 · 2027-01 · APT-53 · VALLE · 244 ses
#46 · 2027-01 · MAP-429 · BOLIVAR · (sin sesiones) · Error
```

## 4. Resultado verificado sobre el vault real

- SQL de corridas ejecutado sobre la copia descifrada (read-only): devuelve **#46(0), #45(3), #44(244), #37(0), #1(145)** — la huérfana #45 ahora es visible.
- Resumen de #44: `SesionesCorrida=244`, `VincCorrida=23`, `PeriodosCorrida="2027-01"`, `SubUnidadesCorrida="VALLE"`; con filtro período 2 → **0**, con período 3 → **244**.

## 5. Cómo cargar corridas anteriores

| Corrida | Estado |
|---|---|
| #44 (2027-01, 244 ses) | Seleccionable directamente; al abrir desde el Módulo R queda preseleccionada |
| #1 (2026-02, 145 ses) | Seleccionable (ya no la filtra `config`) |
| #45 (BOLIVAR, 3 ses) | Ahora visible pese a no tener cabecera en `ag_corrida` |
| #37, #46 | Visibles como `(sin sesiones)` |
| Anteriores del mismo (período, programa) | **No recuperables**: la purga de Fase 502 las elimina al iniciar una nueva corrida |

## 6. Validación

```
scripts/verify.ps1 -Mode Build                → 0 errores
Fast ~Vinculacion                             → 37 superadas / 0 con error
Fast ~AgCorridaSelectorTests|~AgFiltrosPolicy → 14 superadas / 0 con error
scripts/verify.ps1 -Mode Full                 → ver _evidence/
```

## 7. Eliminación de corridas (Fase 1516)

El botón **«Eliminar corrida»** (solo MODO AG, junto a «Ajustar filtros») borra el staging de la corrida seleccionada:

- Confirmación previa con período, programa, territorial y conteos; aviso extra si coincide con `_idCorridaInicial` (la cargada en el Módulo R).
- No actúa sobre el ítem «Todas las corridas» (sentinel `IdCorrida = 0`).
- `EliminarCorridaStagingAsync` es transaccional y borra en orden: `ag_vinculacion_gen` → `ag_sesiones` → `ag_salvedades_conflictos` → `ag_vinculaciones` → `ag_violacion` → `ag_bitacora` → `ag_promocion` → `ag_gen` → `ag_individuo` → `ag_corrida`.
- Las tablas opcionales se omiten si el vault no las tiene (consulta a `sqlite_master`), por lo que funciona en vaults legados.
- Rechaza ids inválidos y corridas inexistentes (`"La corrida #N no existe en el staging."`).
- Tras eliminar: recarga el combo y relanza la consulta; el estado informa los conteos borrados.
- Las tablas `ag_*` **no** están en `SyncInitializer.CrdtTables`, así que el borrado no genera tombstones de producción.

## 8. Layout final de la fila 3

```
Corrida: [12,78]  cmbCorrida [75,75 · 505]  «Ajustar filtros» [585,74 · 170]  «Eliminar corrida» [765,74 · 140]  MODO AG (staging) (rojo, borde derecho)
```

`pnlFilters.MinimumSize = (0, 104)`; en Producción la altura vuelve a 80 px, los tres controles de corrida se ocultan y la etiqueta regresa a `(derecha, 10)`.

