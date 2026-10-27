# Handoff — ag-penalizaciones (F2)

**Plan**: `.agents/plans/2026-10-27-ag-penalizaciones-plan.md`
**Estado**: BLOQUEADO hasta R1 (workspace real) y R3 (baseline medido).

1. **Workspace**: abrir una copia del **árbol real** (fuera de OneDrive). No usar `C:\dev\ConversorPDF`: es un clon de HEAD pre-Academic (98 `.cs` en XLL, 15 en Tests).
2. **Preflight obligatorio** (debe devolver `0` antes de cualquier ola):
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ruta> -Plan .agents/plans/2026-10-27-ag-penalizaciones-plan.md
   ```
   Si devuelve `1` (workspace incompleto), `2` (prerequisitos/baseline) o `3` (OneDrive): **detener**.
3. **Baseline**: recalibrar (`verify.ps1 -Mode Full` + `dotnet test --list-tests`) y actualizar `tests/baseline/test_baseline.txt` con el `TOTAL` real. El `TOTAL: 51` actual corresponde al clon obsoleto y no es válido.
4. **Ejecutar**: `/swarm-parallel-task .agents/plans/2026-10-27-ag-penalizaciones-plan.md`
5. **Gate**: `plan-reviewer` antes de la ola 1. Un solo `build-verifier` a la vez (builds concurrentes fallan con `DNA-577853131`).
6. **Aceptación**: T1..T6 en `done` con evidencia RED→GREEN enlazada, `Full` exit 0, 0 regresiones, suite no reducida, y `files edited/created` ⊆ `files_lock`.
7. **Prohibido**: tocar `.agents/rules/**`, este handoff, `_evidence/**` en commits, `git push`, y cambiar `APP_VERSION`/`Version`.

**Salida esperada**: plan actualizado con ledger (status/evidence/log por tarea) + resumen de ejecución + rutas `_evidence/<timestamp>/summary.md`.
