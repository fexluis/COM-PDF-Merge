# Registro de Numeración — Ronda / Wave / Fase / Task

> [!IMPORTANT]
> **Append-only.** Solo se añaden filas; nunca se editan filas históricas ni se bajan contadores.
> Escritores autorizados: `swarm-planner` (Fase A → reserva) y `swarm-verify` (Fase C → cierre). Ver [`parallel-work.md`](../rules/parallel-work.md) §4.4.

Este archivo es la **fuente única de la numeración global** de revisiones (`.clinerules/workflows/Ronda.md`) y de olas/tareas (planes swarm). Sin él, cada plan reinicia `Wave 1` / `T1` y se pierde el hilo del historial de cambios.

---

## 1. Equivalencia de nomenclatura

| Artefacto de revisión (`Ronda.md`, `implementation_plan.md`, `walkthrough.md`) | Plan swarm (`.agents/plans/<fecha>-<slug>-plan.md`) | Nivel |
|---|---|---|
| `## Ronda {n}: {título}` | `Wave {n}` (parallel execution group) | Grupo de ejecución paralela |
| `## Fase {n}: {título}` | `Task T{n}` (`### T{n}: {Nombre}`) | Tarea atómica |
| `| Fase | Severidad | Archivo | Líneas |` (tabla resumen) | Tabla de hallazgos de `audit-tasks.ps1` | Reporte |

`Ronda` y `Fase` son **globales y monótonas** (no se reutilizan números). `Wave` y `T{n}` son **locales del plan**: la fila de asignación fija el mapeo (`Wave k` ⇔ `Ronda r0+k-1`, `T{n}` ⇔ `Fase f0+n-1`).

---

## 2. Contadores (fuente de verdad)

| Contador | Último asignado | Siguiente | Actualizado |
|---|---|---|---|
| `ronda` | 258 | 259 | 2026-10-29 |
| `fase` | 1419 | 1420 | 2026-10-29 |

Sembrado el 2026-10-29 leyendo `implementation_plan.md` y `walkthrough.md` (`Ronda {n}` → máx. 258; `Fase {n}` → máx. 1419). Los artefactos anteriores a esa fecha no son auditables retroactivamente: no se reasignan.

---

## 3. Asignaciones

| # | Fecha | Plan | Waves (Ronda) | Tasks (Fase) | Estado | verify | HEAD |
|---|---|---|---|---|---|---|---|
| R-001 | 2026-10-29 | `2026-10-29-mover-usuario-entre-vaults-plan.md` | `Wave 1..5` → `Ronda 259–263` | `T1..T6` → `Fase 1420–1425` | reserved | pendiente | `f83b20b` |

Estados: `reserved` (Fase A) · `executing` (Fase B) · `closed` (Fase C, sellado) · `abandoned` (plan descartado: los números **no** se reciclan).

> **R-001**: plan creado antes de este contrato (sin bloque `numbering:` ni campo `fase`). Al sellarlo, la Fase C hace el *retrofit* una sola vez: añade `numbering:` al envelope y `- **fase**: 1420..1425` a `T1..T6`. El retrofit está permitido sólo durante el cierre, nunca durante la Fase B.

Los planes anteriores al registro (`2026-10-27-*`) quedan como `pre-registro`: sus `Wave`/`T{n}` nunca tuvieron número global y no se les asigna uno retroactivo.

---

## 4. Reglas de asignación

1. **Reserva en Fase A**: `swarm-planner` toma `ronda+1` y `fase+1`, reserva un rango contiguo igual al número de olas y de tareas previstos, escribe el bloque `numbering:` del envelope y añade la fila `reserved`.
2. **Idempotencia**: un plan re-planificado conserva sus números; solo se amplía el rango si aparecen tareas nuevas (nunca se compacta ni se reordena).
3. **Plan de fix**: `<fecha>-<slug>-fix-plan.md` recibe un **rango nuevo** y referencia el `# ` de la fila original (precedente: `2026-10-27-usuarios-duplicados-subunidad-plan.md` §12).
4. **Olas reales ≠ olas previstas**: si la Fase B serializa más olas por colisión de leases (`LEASE-SERIALIZED`), el rango de `Ronda` se ajusta al alza en el cierre y se anota la desviación; nunca se reduce.
5. **Sellado en Fase C**: `swarm-verify` cambia `reserved|executing` → `closed`, escribe `verify`, `HEAD` y el rango real de olas, y actualiza los contadores.
6. **Prohibido**: reutilizar un número, inventar contadores a mano o numerar hallazgos de una revisión sin fila previa en este registro.

---

## 5. Cómo continuar la numeración (comandos)

```powershell
# Leer contadores
Get-Content .agents\plans\REGISTRY.md | Select-String '^\| `(ronda|fase)`'

# Contador siguiente en PowerShell
$reg = Get-Content .agents\plans\REGISTRY.md -Raw
$nextRonda = [int]([regex]::Match($reg, '\|\s*`ronda`\s*\|\s*(\d+)\s*\|').Groups[1].Value) + 1
$nextFase  = [int]([regex]::Match($reg, '\|\s*`fase`\s*\|\s*(\d+)\s*\|').Groups[1].Value) + 1
"$nextRonda / $nextFase"
```

Bloque que `swarm-planner` inserta en el envelope (contrato: [`plan-contract.md`](../skills/swarm-planner/references/plan-contract.md) §2.1):

```markdown
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-002
  - ronda_base: 264          # Wave 1 = Ronda 264
  - fase_base: 1426          # T1 = Fase 1426
  - waves: 4                 # Ronda 264–267
  - tasks: 6                 # Fase 1426–1431
```

Y cada tarea declara su número global (campo adicional de `plan-contract.md` §3.12):

```markdown
- **fase**: 1426
```

---

## 6. Trazabilidad con la revisión por rondas

- Un hallazgo de `scripts/audit-tasks.ps1` o de una revisión `Ronda.md` se reporta como `Fase {n}` usando el número **global** del registro, con la `T{n}` de origen entre paréntesis: `Fase 1427 (T3)`.
- El plan de fix que corrige esos hallazgos consume el rango siguiente (`Ronda`/`Fase`) y cita las fases que corrige.
- `walkthrough.md` e `implementation_plan.md` siguen usando `Ronda`/`Fase`; su contador se sincroniza con §2 al cerrar cada ciclo.
