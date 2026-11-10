# Módulo R: selector de corridas y reparación de la promoción a producción

**Fecha:** 2026-11-09 · **Plan:** `.agents/plans/2026-11-09-modulo-r-combo-corridas-plan.md` (R-011) · **Fases:** 1518–1521

---

## 1. Cambios de interfaz

| Antes | Ahora |
|---|---|
| `btnCargarUltimaOptimizacion` — "Cargar última optimización" `(535,10) 195×32` | **`cmbCargarCorridas`** `(535,10) 195×32`, `DropDownWidth = 430` |
| Sin acción de aprobación visible | **`btnAprobarCorrida`** — "✅ Aprobar corrida" en `pnlAgModo` `(510,4) 150×28` |
| `lblAgModo` en `(520,10)` AutoSize | `lblAgModo` en `(670,10)`, `AutoSize = false`, `Width = 276`, `AutoEllipsis` |

**Etiqueta del combo:** `#44 · 2027-01 · APT-53 · VALLE · 244 ses` · sin staging: `… · (sin sesiones) · Error`.
Se lista con `GetCorridasStagingAsync(período, programa, subunidad activa)` (mismo criterio que el botón anterior) y se preselecciona con `AgCorridaSelector.ResolverSeleccionInicial(items, _idCorridaCargada)`. Se refresca al abrir el desplegable y al cambiar período/programa. Sin corridas ⇒ `— sin corridas para el período/programa —` sin disparar carga.

## 2. Diagnóstico de "Promover a Producción" (estaba roto)

| # | Defecto | Evidencia |
|---|---|---|
| **G1** | `AprobarCorridaYAulaAsync()` existía pero **no se invocaba** desde ninguna parte → la corrida quedaba en `Completada (2)` y el gate `Estado != Aprobada` siempre rechazaba | `grep AprobarCorridaYAulaAsync` → solo definición. Vault real: #44, #1, #37 en estado 2; #46 en 4; ninguna en 5 |
| **G2** | El servicio exigía aula válida **sin** tolerar `id_aula = 0` (el formulario sí lo toleraba) | `GetAulaByIdAsync(0)` → `Result.Failure("Aula no encontrada.")`; vault real: #44 y #1 tienen `id_aula = 0` |
| **G3** | `ind.Id` se usaba sin validar | corridas sin `id_individuo_optimo` → *"No se encontró el individuo 0"* |
| **G4** | El botón se deshabilitaba antes de las compuertas y no se restauraba al rechazar | `Enabled = false` + `return` sin re-habilitar |

## 3. Correcciones

1. **G1:** `btnAprobarCorrida` → confirmación → `AprobarCorridaYAulaAsync()` (corrida a `Aprobada`, aulas referenciadas a `Aprobado`). Se habilita solo cuando hay corrida cargada e individuo con `Id > 0`.
2. **G2:** `AgPromotorVinculaciones` omite el gate de aula cuando `IdAula <= 0`, dejando `LoggerHelper.Warning` (paridad con el formulario).
3. **G3:** validación explícita de `ind.Id > 0` con mensaje accionable antes de cualquier escritura.
4. **G4:** el botón se deshabilita al iniciar la promoción y se restaura en `finally` según `HardViolations == 0`.

## 4. Flujo resultante

```
Selector de corridas → cargar (#44, 244 sesiones)
        ↓
✅ Aprobar corrida  → ag_corrida.estado = 5 · ag_aula.estado = 4
        ↓
🚀 Promover a Producción → valida HARD (0) → valida estado (5) → omite aula si id_aula = 0
        ↓
AgPromotorVinculaciones (transacción): vinculaciones/sesiones de producción + ag_promocion
```

## 5. Verificación

| Gate | Resultado |
|---|---|
| `-Mode Build` | 0 errores |
| `-Mode Fast ~AgPromotorVinculaciones` | 5/5 (incluye `WhenCorridaHasNoAula_SkipsAulaGate` y `WhenCorridaHasAulaNotApproved_ShouldReturnFailure`) |
| `-Mode Full` | ver `_evidence/<stamp>/summary.md` |

## 6. Fuera de alcance

Retención de corridas, cambios en `FormVinculaciones`, migraciones de esquema.


## 7. Ajuste de layout y robustez del botón Promover (Fases 1522–1523)

**Layout de `pnlHeaderResultados`** (956 px de ancho, verificado sin solapes):

| Control | x | ancho | fin |
|---|---|---|---|
| `lblVacantes` | 355 | 85 | 440 |
| **`lblCargarCorridas`** "Cargar corridas:" | 444 | 95 (real 90) | 534 |
| **`cmbCargarCorridas`** | 545 | 185 | 730 |
| `btnPromoverProduccion` | 740 | 200 | 940 |

El combo pasa de `195×32` a **`185×25`** (con altura 32 el texto se dibujaba desalineado y parecía duplicado).

**Diagnóstico del botón Promover.** Con datos reales: #44 tiene `id_individuo_optimo = 30` y 48 genes, pero `estado = 2 (Completada)` e `id_aula = 0`. La promoción exige estado **Aprobada**, por lo que rechazaba el intento. Cambios:

1. `Enabled` se **recalcula** tras cada carga (`ind != null && HardViolations == 0`), sin heredar el estado del AG.
2. Cada compuerta deja traza en el log: inicio (corrida/individuo/HARD/vacantes), estado real, aula y error del promotor.
3. El mensaje de compuerta indica el estado actual y señala el botón «✅ Aprobar corrida».
4. `AprobarCorridaYAulaAsync` valida el resultado de `UpdateCorridaAsync` (antes se ignoraba) y `AprobarCorridaDesdeModuloRAsync` **verifica el estado final**, habilitando la promoción solo si quedó en `Aprobada`.

Flujo esperado sobre #44: selector → cargar (244 sesiones) → **✅ Aprobar corrida** (`ag_corrida.estado = 5`, aulas → 4) → **🚀 Promover** (gate de aula omitido por `id_aula = 0`).


## 8. Violaciones HARD: contador persistido vs detalle materializado (Fases 1524–1525)

**Síntoma:** el log mostraba `Corrida #44 cargada: individuo #30, HARD=148, promover=False`, mientras la cabecera indicaba `Violaciones: 0 (HARD 0, SOFT 0)`.

**Causa (datos reales):**

| Tabla | Valor |
|---|---|
| `ag_individuo` #30 (corrida #44) | `hard_violations = 148`, `soft_violations = 0` |
| `ag_violacion` (individuo 30) | **0 filas** (las 138 filas existentes son del individuo #1) |

- La cabecera contaba sólo la **lista** `ind.Violaciones` (0) → mostraba HARD 0.
- El gate de promoción usa el **escalar** `ind.HardViolations` (148) → bloqueaba.
- Origen del detalle vacío: `AgValidadorReglas` fijaba `HardViolations = faltantes` y hacía `return` **sin registrar** las violaciones cuando el cromosoma no tenía genes.

**Correcciones:**

1. **Cabecera honesta:** `HARD = max(hard_violations, detalle)` con sufijo `· detalle no disponible` y `LoggerHelper.Warning` cuando el escalar supera el detalle.
2. **Gate explicado:** el mensaje indica el número de violaciones HARD y advierte que, si el panel aparece vacío, el detalle no se persistió.
3. **Validador:** al no haber genes se registra una violación HARD por asignatura faltante (antes sólo el contador), de modo que `ag_violacion` queda poblado en nuevas corridas.

**Consecuencia operativa:** la corrida #44 **no es promovible** (148 violaciones HARD, compuerta de diseño). Debe corregirse la programación y volver a ejecutar el AG para obtener HARD 0.
