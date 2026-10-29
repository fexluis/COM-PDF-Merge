---
name: swarm-verify
description: >
  [EXPLICIT INVOCATION ONLY] Cierra un plan swarm (Fase C): re-verifica en disco, audita
  leases e invariantes de arquitectura y sella el plan como closed. Nunca reabre el plan.
metadata:
  invocation: explicit-only
---

# Swarm Verify (Fase C — Cline cierra)

Tercera fase del protocolo A → B → C (ver [`.agents/plans/README.md`](../../plans/README.md) §2). La ejecuta **Cline** una vez que `swarm-parallel-task` (o `executing-plans-local`) terminó la Fase B.

> [!CRITICAL]
> **No aceptas el ledger**: lo re-verificas en disco. Eres el único escritor del plan en esta fase.
> **Nunca ejecutes `git push`**. Si hay hallazgos **no reabres** el plan: creas un plan `T-fix`.

Fuentes de verdad:

- Contrato del plan: [`plan-contract.md`](../swarm-planner/references/plan-contract.md)
- Runbook, plantillas y bloque de cierre: [`plans/README.md`](../../plans/README.md) §5
- Leases y concurrencia: [`parallel-work.md`](../../rules/parallel-work.md)

---

## 1. Entrada: ¿es realmente Fase C?

| Requisito | Comprobación |
|---|---|
| Plan en `.agents/plans/<YYYY-MM-DD>-<slug>-plan.md` | `Test-Path` |
| Cabecera `**Estado**: executing` | primera línea del plan |
| Todos los `T<n>` en `status: done` | `Get-Content <plan> \| Select-String 'status:'` |
| Bloque `## Execution envelope` presente | cabecera del plan |
| `base` = SHA anterior a T1 | lo aporta el usuario |

Si el plan está en `draft` (Fase A), `approved` (Fase B pendiente) o ya `closed`, **detente y reporta**: no hay nada que sellar.

## 2. C0 — Precondiciones

```powershell
git status --porcelain                                                                          # debe estar VACIO
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ws> -Plan <plan>      # exit 0
```

Un working tree sucio invalida el ledger (no se puede atribuir cada cambio a una tarea). En OneDrive `preflight.ps1` devuelve `3`: usar `-AllowOneDrive` sólo con aviso explícito en el acta de cierre.

## 3. C1 — Verificación independiente

```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```

| exit | Significado | Veredicto |
|---|---|---|
| `0` | OK y `TOTAL >= baseline` | continuar |
| `1` | Build falló | `FIX_PLAN` |
| `2` | Tests fallaron / infraestructura | `FIX_PLAN` |
| `3` | Regresión vs baseline, `SUITE_SHRUNK` o `NO_BASELINE` | `FIX_PLAN` (bloqueante) |
| `4` | Tooling o workspace inválido | `BLOCKED` |

Evidencia: `_evidence/<stamp>/summary.md` — se enlaza por ruta y **no se commitea**.

## 4. C2 — Auditoría de leases

```powershell
foreach ($c in (git rev-list "$base..HEAD")) {
  git log -1 --pretty=format:'%h %s' $c; git show --name-only --pretty=format: $c; ''
}
```

Regla: para cada commit `T<n>: ...`, los archivos tocados deben ser subconjunto de `files_lock(T<n>)` ∪ `{.agents/plans/<plan>.md}`. Cualquier archivo adicional ⇒ `LEASE-VIOLATION` ⇒ `FIX_PLAN`.

## 5. C3 — Invariantes de arquitectura

```powershell
Select-String -Path ConversorPDF.XLL\Academic\Application\**\*.cs -Pattern 'Excel\.|Microsoft\.Office|DataGridView'
Select-String -Path ConversorPDF.XLL\Academic\**\*.cs -Pattern 'new \[\]|DateTime\.Now|\.Result\b|\.Wait\(\)'
```

Esperado: **0 resultados**. Interop en capas limpias, o `DateTime.Now` / `.Result` / `.Wait()` en el slice ⇒ `FIX_PLAN`.

## 6. C4 — Publicación y documentación

```powershell
git log --oneline origin/refactor/<slug>..HEAD      # los commits DEBEN seguir sin publicar
```

Artefactos exigidos por `development-process` (Finalization):

- `docs/articles/changes/<YYYY-MM-DD>-<slug>.md` (Purpose · Changes · How to use · Test coverage)
- `walkthrough.md`

## 7. C5 — Sellado

1. En la cabecera, `**Estado**: executing` → `closed` (único cambio sobre el ledger de Fase B).
2. Añadir el bloque `## Cierre (Fase C)` con la plantilla literal de [`plans/README.md`](../../plans/README.md) §5.3: `verify Full`, `HEAD`, evidencia, leases, arquitectura, publicación, docs y desviaciones.
3. Registrar el veredicto y la lista de commits auditados.

Ciclo de vida del estado: `draft` → `approved` → `executing` → **`closed`**.

## 8. C6 — Commit local (sin push)

```powershell
git add -- .agents/plans/<plan>.md walkthrough.md docs/articles/changes/<fecha>-<slug>.md
git diff --cached --name-only        # auditar: solo esos archivos
git commit -m "docs(fase-c): cierre <slug> — verify Full exit 0, 0 regresiones"
```

## 9. Contrato de salida

Responde con:

- **Veredicto**: `SELLADO` | `FIX_PLAN` | `BLOCKED`
- **Evidencia**: exit code de `verify.ps1`, `TOTAL` medido vs baseline, `HEAD`, archivos commiteados
- **Hallazgos**: numerados, con la `T<n>` afectada y el riesgo de rotura
- **Si procede**: ruta del `<YYYY-MM-DD>-<slug>-fix-plan.md` creado

## 10. Prohibiciones

- **No reabrir** el plan cuando hay hallazgos: dos escritores sobre el ledger. Se crea `<YYYY-MM-DD>-<slug>-fix-plan.md` (precedente: `2026-10-27-usuarios-duplicados-subunidad-plan.md` §12).
- **No push, merge ni publish.**
- **No aceptar el ledger** sin re-ejecutar `verify.ps1 -Mode Full` en disco.
- **No commitear** `_evidence/**` ni archivos fuera del acta de cierre.
- **No sellar** con exit `1`, `2`, `3` ni `4`.

