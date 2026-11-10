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
| `ronda` | 292 | 293 | 2026-11-09 |
| `fase` | 1529 | 1530 | 2026-11-09 |

Sembrado el 2026-10-29 leyendo `implementation_plan.md` y `walkthrough.md` (`Ronda {n}` → máx. 258; `Fase {n}` → máx. 1419). Los artefactos anteriores a esa fecha no son auditables retroactivamente: no se reasignan.

---

## 3. Asignaciones

| # | Fecha | Plan | Waves (Ronda) | Tasks (Fase) | Estado | verify | HEAD |
|---|---|---|---|---|---|---|---|
| R-001 | 2026-10-29 | `2026-10-29-mover-usuario-entre-vaults-plan.md` | `Wave 1..5` → `Ronda 259–263` | `T1..T6` → `Fase 1420–1425` | reserved | pendiente | `f83b20b` |
| R-002 | 2026-10-30 | `2026-10-30-gestion-documental-rango-yy-plan.md` | `Wave 1..2` → `Ronda 264–265` | `T1..T11` → `Fase 1426–1436` | closed | pendiente | `3bb7b18` |
| R-003 | 2026-10-31 | `2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md` | `Wave 1..3` → `Ronda 266–268` | `T1..T6` → `Fase 1437–1442` | executing | pendiente (gate Full) | `3bb7b18` |
| R-004 | 2026-10-31 | `(mantenimiento sin plan: bump 1.6.3.83)` | `Wave 272` → `Ronda 272` | `T1` → `Fase 1461` | executing | pendiente (gate conjunto) | pendiente |
| R-005 | 2026-11-01 | `2026-11-01-verificacion-solicitudes-plan.md` | `Wave 1..3` → `Ronda 273–275` | `T1..T6` → `Fase 1462–1467` | abandoned | — | pendiente |
| R-006 | 2026-10-31 | `2026-10-31-duplicados-y-auditoria-informe-plan.md` | `Wave 1..3` → `Ronda 279–281` | `T1..T5` → `Fase 1483–1487` | closed | 0 (Full, 999 tests) | sin commit |
| R-007 | 2026-11-09 | `2026-11-09-corridas-ag-formvinculaciones-plan.md` | `Wave 1..3` → `Ronda 282–284` | `T1..T8` → `Fase 1488–1495` · `T9..T10` → `Fase 1516–1517` | executing | PASS (Full 1064) | pendiente |
| R-008 | 2026-11-09 | `(revisión manual Ronda 285 — correcciones R-007)` | `Wave —` | `Fase 1496–1503` | closed | — | pendiente |
| R-009 | 2026-11-09 | `(revisión manual Ronda 286 — Fases 1504–1509)` | `Wave —` | `Fase 1504–1509` | closed | — | pendiente |
| R-010 | 2026-11-09 | `(revisión manual Ronda 287 — Fases 1510–1514)` | `Wave —` | `Fase 1510–1514` | closed | — | pendiente |
| R-012 | 2026-11-09 | `2026-11-09-modulo-r-combo-corridas-plan.md` | `Wave 1` → `Ronda 289` | `T11..T14` → `Fase 1518–1521` | closed | PASS (Full 1070) | pendiente |
| R-013 | 2026-11-09 | `(fix de R-012: layout del selector y robustez de la promoción)` | `Wave —` | `Fase 1522–1523` | closed | PASS (Full 1070) | pendiente |
| R-014 | 2026-11-09 | `(fix de R-012: coherencia de violaciones HARD)` | `Wave —` | `Fase 1524–1525` | closed | PASS (Full 1071) | pendiente |
| R-011 | 2026-11-09 | `(revisión manual Ronda 288 — Fases 1515–1516)` | `Wave —` | `Fase 1515–1516` | closed | — | pendiente |
| R-015 | 2026-11-09 | `(fix de R-014: recálculo manual de violaciones y resaltado en vista por fechas)` | `Wave —` | `Fase 1526–1529` | closed | PASS (Full) | pendiente |

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
