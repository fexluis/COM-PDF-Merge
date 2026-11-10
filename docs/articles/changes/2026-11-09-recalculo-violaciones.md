# Recálculo manual de violaciones y resaltado en la vista por fechas

**Fecha:** 2026-11-09 · **Plan:** fix de R-014 (no requiere plan swarm) · **Fases:** 1526–1529 · **Ronda:** 292

---

## 1. Problema

Una corrida generada con una compilación anterior podía quedar con `ag_individuo.hard_violations > 0` y
`ag_violacion` vacía. El encabezado mostraba `HARD 0`, la compuerta de promoción bloqueaba con 95/148 y el panel
«Violaciones de Reglas» aparecía en blanco: no había forma de saber **qué** regla se incumplía sin volver a ejecutar
el AG completo.

## 2. Recálculo manual (Módulo R)

`IAgRecalculoViolacionesService` / `AgRecalculoViolacionesService`:

```
ag_corrida(idCorrida) ──► id_individuo_optimo
ag_individuo(id) ──────► genes
contexto del AG ───────► AgValidadorReglas.Evaluar()
                        └─► DELETE + INSERT ag_violacion (transacción)
                            UPDATE ag_individuo (hard/soft/estado)
                            UPDATE ag_corrida   (hard/soft)
```

- **Idempotente:** el `DELETE` previo y `Violaciones.Clear()` del validador garantizan el mismo resultado en cada ejecución.
- **No reejecuta el AG:** reutiliza los genes ya materializados y el contexto del formulario (`_contextoActual`); si es `null`,
  lo reconstruye con `ConstruirContextoAsync()`.
- **Atómico:** una transacción sobre `ag_violacion`, `ag_individuo` y `ag_corrida`.
- **Trazable:** `LoggerHelper.Info("[AG-Recalculo] …")` con HARD, SOFT, genes evaluados y desglose por regla.

### Contratos nuevos

| Elemento | Ubicación |
|---|---|
| `AgRecalculoResultado` | `Academic/Domain/Models/Genetica/` |
| `IAgRecalculoViolacionesService` | `Academic/Application/Services/Genetica/` |
| `AgRecalculoViolacionesService` | `Academic/Application/Services/Genetica/` |
| `IAgScheduleRepository.ReemplazarViolacionesAsync` | `Academic/Domain/Interfaces/` |

### Interfaz

`pnlAgModo` pasa de `956×38` a `956×66`. Fila 2:

| Control | Posición | Texto |
|---|---|---|
| `btnRecalcularViolaciones` | `(10,34) 200×28` | «🔁 Recalcular violaciones» |
| `lblAgModo` | `(220,40) 726×15` | indicador de corrida / staging |

Se habilita solo cuando hay corrida cargada con individuo persistido (`Id > 0`).

## 3. Visualización en `FormVerVinculaciones`

`CargarViolacionesAsync()` lee el detalle de la corrida mostrada:

```sql
SELECT s.id AS IdSesion, v.regla, v.es_dura, v.mensaje
FROM ag_violacion v
INNER JOIN ag_sesiones s ON s.idgen = v.id_gen AND s.idindividuo = v.id_individuo
INNER JOIN ag_vinculaciones av ON av.id = s.idvinculacion
WHERE av.id_corrida = @IdCorrida
```

y lo indexa en `_violacionesPorSesion`. En `ConstruirArbolPorFechas` las violaciones **HARD** tienen prioridad sobre
las salvedades:

| Estado | Fondo | Texto | Marca |
|---|---|---|---|
| Violación HARD | `#FFCDD2` | `Firebrick` | `⛔` + reglas + tooltip con el mensaje |
| Conflicto de salvedad | `MistyRose` | `DarkRed` | `⚠️` |
| Revisar viabilidad | `LightYellow` | `DarkGoldenrod` | `⚡` |
| Salvedad vigente | `Honeydew` | `DarkGreen` | `✓` |
| Plataforma | `LightCyan` | `DarkBlue` | `🌐` |

El `ToolTipText` de la sesión concatena el detalle de las violaciones con el detalle del cruce de salvedad, y la barra
inferior antepone `⛔ Violaciones HARD: {n} en {m} sesiones`.

## 4. Criterios de aceptación verificados

1. Recálculo sobre una corrida antigua deja `ag_violacion` poblada y `hard_violations` coherente.
2. Segunda ejecución produce el mismo conteo y desglose (idempotencia).
3. La vista por fechas muestra `⛔ Rxx` por sesión, tooltip con el mensaje y contador inferior.
4. Una salvedad vigente **no** oculta una violación HARD.
5. `scripts/verify.ps1 -Mode Full` en verde.

## 5. Rollback

El recálculo solo escribe `ag_violacion`, `ag_individuo.hard_violations/soft_violations/estado` y
`ag_corrida.hard_violations/soft_violations`. No toca tablas de producción; revertir el commit restaura el
comportamiento anterior sin migración de datos.