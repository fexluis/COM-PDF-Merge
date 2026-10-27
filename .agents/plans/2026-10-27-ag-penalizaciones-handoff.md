# Handoff — ag-penalizaciones (F2)

**Plan**: `.agents/plans/2026-10-27-ag-penalizaciones-plan.md`
**Estado**: LISTO — R1+R2+R3 completados el 2026-10-27.

1. **Workspace**: `C:\dev\ConversorPDF` (clon fiel de `6ae4995`: 452 `.cs` en XLL, 138 en Tests, slice `Academic`).
2. **Preflight obligatorio** (debe devolver `0` antes de cualquier ola):
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ruta> -Plan .agents/plans/2026-10-27-ag-penalizaciones-plan.md
   ```
   Si devuelve `1` (workspace incompleto), `2` (prerequisitos/baseline) o `3` (OneDrive): **detener**.
3. **Baseline**: `TOTAL: 939` medido el 2026-10-27 sobre `6ae4995` (939 passed, 0 failed, ~58 s). No recalibrar salvo que cambie el árbol.
4. **Ejecutar**: `/swarm-parallel-task .agents/plans/2026-10-27-ag-penalizaciones-plan.md`
5. **Gate**: `plan-reviewer` antes de la ola 1. Un solo `build-verifier` a la vez (builds concurrentes fallan con `DNA-577853131`).
6. **Aceptación**: T1..T6 en `done` con evidencia RED→GREEN enlazada, `Full` exit 0, 0 regresiones, suite no reducida, y `files edited/created` ⊆ `files_lock`.
7. **Prohibido**: tocar `.agents/rules/**`, este handoff, `_evidence/**` en commits, `git push`, y cambiar `APP_VERSION`/`Version`.

**Salida esperada**: plan actualizado con ledger (status/evidence/log por tarea) + resumen de ejecución + rutas `_evidence/<timestamp>/summary.md`.
