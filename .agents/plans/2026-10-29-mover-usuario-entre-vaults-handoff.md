# Handoff: Mover usuario entre vaults al cambiar de subunidad (MUV)

**Generado**: 2026-10-29 · **Autor**: Cline/Swarm-Planner · **Plan**: `.agents/plans/2026-10-29-mover-usuario-entre-vaults-plan.md`

1. **Workspace a abrir**: el árbol real (456 `.cs` XLL / 143 `.cs` Tests). Hoy está dentro de OneDrive: si se ejecuta ahí, `preflight.ps1` exige `-AllowOneDrive` y `verify.ps1` emite aviso de locks. **Recomendado**: copia fuera de OneDrive (`C:\dev\ConversorPDF`) para la Fase B.
2. **Rama y estado pre-flight**: `refactor/academic-modularization`, `git status --porcelain` **vacío** (HEAD de partida `7381836`). Comando: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ws> -Plan .agents/plans/2026-10-29-mover-usuario-entre-vaults-plan.md -AllowOneDrive` ⇒ exit `0`.
3. **Comando de ejecución**: `/swarm-parallel-task .agents/plans/2026-10-29-mover-usuario-entre-vaults-plan.md` · fallback sin subagentes: `executing-plans-local` sobre el mismo plan.
4. **Gate obligatorio**: invocación inicial de `plan-reviewer`. Si devuelve `BLOCKED` o cualquier hallazgo `Critical` ⇒ **abortar** y reportar (no se toca código).
5. **Criterios de aceptación globales**: RED→GREEN por tarea; `verify.ps1 -Mode Full` con `TOTAL >= 952` y 0 regresiones; 0 archivos fuera de `files_lock`; **ningún `DELETE FROM usuarios`** (R-MUV-01); `config.subunidad` solo para el usuario de Windows (R-MUV-02).
6. **Prohibiciones**: `git push`/merge/publish · tocar `SyncContext.cs`, `SubunidadResolver.cs`, `UsuarioVaultProvisioner.cs`, migraciones o `RibbonController.cs` (fuera de `files_lock`) · C# 12 · SQL nuevo en code-behind · debilitar tests · cerrar una tarea como `completed` sin evidencia (lección de T6 del plan 2026-10-27).
7. **Entregables esperados**: plan con ledger completo (`status`, `log`, `evidence`, `files edited/created` por tarea), commits locales atómicos por tarea (`T<n>: …`), directorio `_evidence/<timestamp>/`, `walkthrough.md` y `docs/articles/changes/2026-10-29-mover-usuario-entre-vaults.md`.

---

## Notas de arranque para el orquestador

- **Ola 1 es `integrator`** (L0: `.csproj`). Sin ella, T2 no compila: el RED necesita `IUsuarioVaultMover`, los 2 DTOs y los esqueletos.
- **Ola 3 admite exactamente 2 escritores** (`max_parallel_writers: 2`): T3 (`Infrastructure/Persistence`) y T4 (`Application/Services`). No añadir una tercera.
- **`FormUsuarios.cs` se toca una sola vez** (T5). Si T5 se rechaza en revisión, el fix va en un `T-fix`, no reabriendo T5.
- **AC3 (arranque) es manual**: requiere Excel. Si no se puede ejecutar, se registra como deuda declarada en el ledger, no como verde.
- **Rulings vigentes que el orquestador no puede contradecir sin emitir uno nuevo**: `R-20261029-MUV-01` (sin `DELETE`), `-02` (`config` es del dispositivo), `-03` (no reutilizar el provisioner), `-04` (seam de test por delegados).
