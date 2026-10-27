# Handoff — usuarios duplicados + subunidad no persiste

**Plan**: `.agents/plans/2026-10-27-usuarios-duplicados-subunidad-plan.md`
**Estado**: `draft` — requiere aprobación humana (pasar a `approved`).

1. **Workspace**: `C:\dev\ConversorPDF` (clon fiel de `d32115e`, 452/138 `.cs`, `Academic` presente).
2. **Preflight obligatorio** (debe devolver `0` antes de cualquier ola):
   ```powershell
   powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-10-27-usuarios-duplicados-subunidad-plan.md
   ```
3. **Ejecutar**: `/swarm-parallel-task .agents/plans/2026-10-27-usuarios-duplicados-subunidad-plan.md`
4. **Gate**: `plan-reviewer` antes de la ola 1. Un solo `build-verifier` a la vez (`DNA-577853131` si se solapan).
5. **Aceptación** (además de `Full` en verde):
   - dos filas del mismo `nombreusuario` ⇒ el grid muestra **una**;
   - cambiar subunidad ⇒ guardar ⇒ **cerrar Excel** ⇒ reabrir ⇒ el valor nuevo persiste.
6. **Prohibido**: `git push`; tocar el vault sin backup (T7 lo exige); cambiar `APP_VERSION`/`Version`; debilitar tests; escribir SQL nuevo en code-behind.
7. **Nota de alcance**: T6 (mover usuario entre vaults) es **condicional** a lo que demuestre T2; si no se demuestra, se cierra con `Ruling`. T7 reescribe IDs en una tabla CRDT: backup + `PRAGMA integrity_check` obligatorios.

**Salida esperada**: plan con ledger (status/evidence/log por tarea), commits `T<n>:` locales, `_evidence/<timestamp>/summary.md` y `docs/articles/changes/2026-10-27-usuarios-duplicados-subunidad.md`.
