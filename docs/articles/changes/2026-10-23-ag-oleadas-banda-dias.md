# Oleadas por bandas de días (v3) — Fases 1224–1235

## Problema
El planificador usaba carriles de 2 ranuras con rotación `i % carriles` y oleadas secuenciales:
- el carril limitaba los días simultáneos de una asignatura (ventanas de 3 días dejaban 1 día sin usar);
- `olas = Min(olas, W)` forzaba el excedente a la última oleada sin chequear saturación (Fase 1239/1243);
- la siembra de `dias_patron` de un solo día contradecía el plan (Fase 1244).

## Solución
`AgBandasDias.Particiones` genera las 2^(n-1) particiones contiguas disjuntas de la ventana.
`AgBandasDias.Resolver` elige la asignación asignatura→banda con objetivo lexicográfico
(infactibles, patrones violados, makespan, desperdicio, -nº bandas), exhaustivo si `k^n ≤ 4096`, si no BFD.
`AgEscaleraVentanas` sugiere `VIE-SAB → JUE-SAB → 4,5,6,7` (máx. 2 pasos, idempotente) y
`AgEscaladorVentanaAula` decide (dry-run) y aplica en `ag_aula`/`ag_arbol_nodo`.

## Resultados
| Caso | Demanda | Partición | Semanas | Ocupación |
|------|---------|-----------|---------|-----------|
| 1 | 60 sesiones | 2 bandas | 20 | 100 % |
| 2 | 64 sesiones | 2 bandas | 22 | 92,8 % (+`I5_HOLGURA_BAJA`) |

## Contrato
`AgPlanOleada.Carril` = índice de banda; `Oleada` = 0 (bandas paralelas).
`AgPlanOleadasResultado` expone `Particion`, `VentanaSugerida`, `PatronesRenormalizados`, `OcupacionPorcentaje`.

## Migración / rollback
Sin cambios de esquema. Rollback: `USAR_BANDAS_DIAS = false` y revertir `AgPlanificadorOleadas` a la rama de carriles.
