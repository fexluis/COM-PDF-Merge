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
| `ronda` | 391 | 392 | 2026-11-14 |
| `fase` | 1805 | 1806 | 2026-11-14 |

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
| R-016 | 2026-11-10 | `(revisión manual Ronda 293 — Fases 1530–1536)` | `Wave —` | `Fase 1530–1536` | closed | — | pendiente |
| R-017 | 2026-11-10 | `(revisión manual Ronda 294 — Fases 1537–1540)` | `Wave —` | `Fase 1537–1540` | closed | — | pendiente |
| R-018 | 2026-11-10 | `(revisión manual Ronda 295 — Fase 1541)` | `Wave —` | `Fase 1541` | closed | — | pendiente |
| R-019 | 2026-11-10 | `(revisión manual Ronda 296 — Fases 1542–1544)` | `Wave —` | `Fase 1542–1544` | closed | — | pendiente |
| R-020 | 2026-11-10 | `(revisión manual Ronda 297 — Fases 1545–1548)` | `Wave —` | `Fase 1545–1548` | closed | — | pendiente |
| R-021 | 2026-11-10 | `(revisión manual Ronda 298 — Fases 1549–1550)` | `Wave —` | `Fase 1549–1550` | closed | — | pendiente |
| R-022 | 2026-11-10 | `(revisión manual Ronda 299 — Fase 1551)` | `Wave —` | `Fase 1551` | closed | — | pendiente |
| R-023 | 2026-11-10 | `(revisión manual Ronda 300 — Fases 1552–1553)` | `Wave —` | `Fase 1552–1553` | closed | — | pendiente |
| R-024 | 2026-11-11 | `(revisión manual Ronda 302 — Fases 1554–1555)` | `Wave 301` | `Fase 1554–1555` | closed | — | pendiente |
| R-025 | 2026-11-11 | `(revisión manual Ronda 303 — Fase 1556)` | `Wave 302` | `Fase 1556` | closed | — | pendiente |
| R-026 | 2026-11-11 | `(revisión manual Ronda 304 — Fases 1557–1558)` | `Wave 304` | `Fase 1557–1558` | closed | — | pendiente |
| R-027 | 2026-11-11 | `(revisión manual Ronda 305 — Fase 1559)` | `Wave 305` | `Fase 1559` | closed | — | pendiente |
| R-028 | 2026-11-11 | `(revisión manual Ronda 306 — Fase 1562)` | `Wave 306` | `Fase 1562` | closed | — | pendiente |
| R-029 | 2026-11-11 | `(revisión manual Ronda 307 — Fases 1563–1567: corrida huérfana)` | `Wave 307` | `Fase 1563–1567` | executing | PASS (Full 1102) | e53d2f5 |
| R-030 | 2026-11-11 | `(revisión manual Ronda 308 — Fases 1568–1574: validador oleadas)` | `Wave 308` | `Fase 1568–1574` | closed | PASS (Fast 1123 + #48 HARD 0) | HEAD |
| R-031 | 2026-11-11 | `(revisión manual Ronda 309 — Fases 1575–1578: contexto homónimo Módulo R)` | `Wave 309` | `Fase 1575–1578` | closed | PASS (Fast 1124 + recálculo #48 125→0) | HEAD |
| R-032 | 2026-11-11 | `(revisión manual Ronda 310 — Fase 1579: sede en el log del validador de ventana)` | `Wave —` | `Fase 1579` | closed | PASS (Fast 21/21 AgValidadorVentanaTests) | HEAD |
| R-033 | 2026-11-11 | `(revisión manual Ronda 311 — Fase 1580: alcance completo al cargar perfil de árbol)` | `Wave —` | `Fase 1580` | closed | PASS (Fast 27 tests alcance+validador) | HEAD |
| R-034 | 2026-11-11 | `(revisión manual Ronda 312 — Fase 1581: partición de carriles consciente del patrón)` | `Wave —` | `Fase 1581` | closed | PASS (Fast 1156, 49/49 oleadas) | HEAD |
| R-035 | 2026-11-11 | `(revisión manual Ronda 313 — Fases 1582–1589: fidelidad de patrón y robustecimiento)` | `Wave —` | `Fase 1582–1589` | closed | PASS (Fast/Full) | HEAD |
| R-036 | 2026-11-11 | `(revisión manual Ronda 314 — Fases 1590–1595: separación de casos, diagnóstico y gate)` | `Wave —` | `Fase 1590–1595` | closed | PASS (Fast 1157, gate y fixtures calibrados) | HEAD |
| R-037 | 2026-11-11 | `(revisión manual Ronda 315 — Fases 1596–1598: índices/bloques gate, pool docente y festivo)` | `Wave —` | `Fase 1596–1598` | closed | PASS (Fast/Full) | HEAD |
| R-038 | 2026-11-11 | `(revisión manual Ronda 316 — Fase 1599: API canónica RegistrarDocenteHabilitado en Gate G-1 y calibración Caso 7.6)` | `Wave —` | `Fase 1599` | closed | PASS (Fast 1158, 44/44 focalizados) | HEAD |
| R-039 | 2026-11-11 | `(revisión manual Ronda 317 — Fases 1600–1602: reparto compatible con patrón y diagnóstico verificable de carriles)` | `Wave —` | `Fase 1600–1602` | closed | PASS (Fast/Full) | HEAD |
| R-040 | 2026-11-11 | `(revisión manual Ronda 318 — Fases 1603–1605: corrección de tipos en pruebas, calibración de fixtures y cobertura festivo)` | `Wave —` | `Fase 1603–1605` | closed | PASS (Full 1166) | HEAD |
| R-041 | 2026-11-12 | `2026-11-12-sincronizacion-corridas-modulo-r-plan.md` | `Wave 1..7` → `Ronda 319–325` | `T1..T7` → `Fase 1606–1612` | closed | PASS (Full 1172) | c15738e |
| R-042 | 2026-11-12 | `(revisión manual Ronda 326 — Fases 1613–1620: desincronización commits, popup desplegable y reentradas)` | `Wave —` | `Fase 1613–1620` | closed | PASS (Full 1172) | HEAD |
| R-043 | 2026-11-12 | `(correcciones Ronda 327 — Fases 1621–1626: refresco diferido, coalescencia de cargas y marcador fuera de filtro)` | `Wave —` | `Fase 1621–1626` | closed | pendiente (gate Full) | HEAD |
| R-044 | 2026-11-12 | `(correcciones Ronda 328 — Fases 1627–1629: refresco anticipado selector y desacoplamiento modal)` | `Wave —` | `Fase 1627–1629` | closed | pendiente (gate Full) | HEAD |
| R-045 | 2026-11-12 | `(correcciones Ronda 329 — Fases 1630–1632: reapertura segura de popup, deduplicación de consultas y robustez de handle)` | `Wave —` | `Fase 1630–1632` | closed | pendiente (gate Full) | HEAD |
| R-046 | 2026-11-12 | `(correcciones Ronda 330 — Fases 1633–1635: desacoplamiento de reapertura, flag cierre interno y precisión de timestamp)` | `Wave —` | `Fase 1633–1635` | closed | pendiente (gate Full) | HEAD |
| R-047 | 2026-11-12 | `(correcciones Ronda 331 — Fases 1636–1637: reapertura con validación de foco y guard contra cascadas espurias)` | `Wave —` | `Fase 1636–1637` | executing | pendiente (gate Full) | pendiente |
| R-048 | 2026-11-12 | `(revisión manual Ronda 332 — Fases 1651–1660: ejecutabilidad del plan 4→8 h)` | `Wave —` | `Fase 1651–1660` | closed | manual | `931f529` |
| R-049 | 2026-11-12 | `(revisión manual Ronda 333 — Fases 1661–1666: contrato de migración y procedencia de patrón)` | `Wave —` | `Fase 1661–1666` | closed | manual | `931f529` |
| R-050 | 2026-11-12 | `(revisión manual Ronda 334 — Fases 1667–1671: fidelidad estricta y migración no aditiva)` | `Wave —` | `Fase 1667–1671` | closed | manual | `931f529` |
| R-051 | 2026-11-12 | `(revisión manual Ronda 335 — Fases 1672–1676: Sellar destruye el plan explícito)` | `Wave —` | `Fase 1672–1676` | closed | manual | `931f529` |
| R-052 | 2026-11-12 | `(revisión manual Ronda 336 — Fases 1677–1681: guarda quirúrgica y parser único)` | `Wave —` | `Fase 1677–1681` | closed | manual | `931f529` |
| R-053 | 2026-11-12 | `(revisión manual Ronda 337 — Fases 1682–1686: Clonar y tolerancia de déficit)` | `Wave —` | `Fase 1682–1686` | closed | manual | `931f529` |
| R-054 | 2026-11-12 | (revisión documental Ronda 338: dependencias, plan explícito y cobertura) | Wave — | Fase 1687–1691 | closed | revisión documental; sin build/tests | c0d986d |
| R-055 | 2026-11-12 | (revisión documental Ronda 339: sincronización efectiva del plan) | Wave — | Fase 1692 | closed | revisión documental; sin build/tests | c0d986d |
| R-056 | 2026-11-12 | (corrección documental Ronda 340: rutas T6/T10 y numeración) | Wave — | Fase 1693–1694 | closed | inspección de archivos; sin build/tests | c0d986d |
| R-057 | 2026-11-12 | `(revisión estática Ronda 344 — Fases 1701–1706: consolidación escalera ventanas)` | `Wave —` | `Fase 1701–1706` | closed | revisión documental; sin build/tests | `e523b51` |
| R-058 | 2026-11-12 | `2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom-plan.md` | `Wave 1..6` → `Ronda 345–350` | `T1..T8` → `Fase 1707–1714` | closed | PASS (Full 1223) | `1a97cfc` |
| R-059 | 2026-11-12 | `(revisión Ronda 351 — Fases 1715–1719: correcciones antes de release/smoke)` | `Wave —` | `Fase 1715–1719` | closed | revisión y correcciones; sin build/tests | `1a97cfc` |
| R-060 | 2026-11-12 | `(revisión Ronda 352 — Fases 1720–1723: robustez SQL, recarga checked y exclusión mutua pre-vuelo)` | `Wave —` | `Fase 1720–1723` | closed | revisión y correcciones; sin build/tests | `1a97cfc` |
| R-061 | 2026-11-12 | `(revisión Ronda 353 — Fases 1724–1727: tipos en tests, Result<int> en sincronización y anti-reentrada)` | `Wave —` | `Fase 1724–1727` | closed | revisión y correcciones; sin build/tests | `1a97cfc` |
| R-062 | 2026-11-12 | `(revisión Ronda 354 — Fases 1728–1730: revisión concurrente, bloqueo diagnóstico y cierre confirmado)` | `Wave —` | `Fase 1728–1730` | closed | revisión y correcciones; sin build/tests | `1a97cfc` |
| R-063 | 2026-11-12 | `(revisión Ronda 355 — Fase 1731: aislamiento de subunidad en dimensiones pendientes)` | `Wave —` | `Fase 1731` | closed | revisión y correcciones; sin build/tests | `1a97cfc` |
| R-064 | 2026-11-12 | `(revisión Ronda 356 — Fase 1732: corrección de columnas en fixtures ag_aula de tests)` | `Wave —` | `Fase 1732` | closed | revisión y correcciones; sin build/tests | `1a97cfc` |
| R-065 | 2026-11-12 | `(re-gate post-cierre Ronda 357 — Fase 1733: correcciones R-059..R-064)` | `Wave —` | `Fase 1733` | closed | PASS (Full 1236) | `27b7c71` |
| R-066 | 2026-11-12 | `2026-11-12-escalado-automatico-tres-casos-plan.md` | `Wave 1..8` → `Ronda 358–365` | `T1..T8` → `Fase 1734–1741` | closed | PASS (Full 1253) | `c34bec4` |
| R-067 | 2026-11-12 | `(revisión Ronda 366 — Fases 1742–1746: criterios mejora, persistencia patrones)` | `Wave —` | `Fase 1742–1746` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-068 | 2026-11-12 | `(revisión Ronda 367 — Fases 1747–1752: carriles reales, semanas requeridas)` | `Wave —` | `Fase 1747–1752` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-069 | 2026-11-12 | `(revisión Ronda 368 — Fases 1753–1758: bloque huérfano, bounds carril, fecha fin ventana)` | `Wave —` | `Fase 1753–1758` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-070 | 2026-11-12 | `(revisión Ronda 369 — Fases 1759–1762: granularidad slot/franja, contención reset)` | `Wave —` | `Fase 1759–1762` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-071 | 2026-11-12 | `(revisión Ronda 370 — Fases 1763–1767: exposición métricas, I5_DIA_SATURADO, retiro bandas)` | `Wave —` | `Fase 1763–1767` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-072 | 2026-11-12 | `(revisión Ronda 371 — Fases 1768–1771: supresión oleada contradictoria, déficit global ordenado)` | `Wave —` | `Fase 1768–1771` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-073 | 2026-11-12 | `(revisión Ronda 372 — Fases 1772–1774: reducción materias fuera de ventana, mejora real)` | `Wave —` | `Fase 1772–1774` | closed | PASS (Full 1259) | `9265db3` |
| R-074 | 2026-11-12 | `(revisión Ronda 373 — Fases 1775–1779: cota horizonte fecha, autoridad planificador, idempotencia)` | `Wave —` | `Fase 1775–1779` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-075 | 2026-11-12 | `(revisión Ronda 374 — Fases 1780–1783: métricas destino, confirmadas idempotencia, huella)` | `Wave —` | `Fase 1780–1783` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-076 | 2026-11-12 | `(revisión Ronda 375 — Fases 1784–1787: rollback vacío, demanda pendiente real, huella/idempotencia)` | `Wave —` | `Fase 1784–1787` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-077 | 2026-11-14 | `2026-11-14-reclasificacion-solo-pendientes-plan.md` | `Wave 1..8` → `Ronda 376–383` | `T1..T10` → `Fase 1788–1797` | closed | PASS (Full 1356) | `93137d3` |
| R-078 | 2026-11-14 | `2026-11-14-corrida92-completar-huecos-plan.md` | `Wave 1..8` → `Ronda 384–391` | `T1..T8` → `Fase 1798–1805` | closed | PASS (Full 1387) | pendiente |


Estados: `reserved` (Fase A) · `executing` (Fase B) · `closed` (Fase C, sellado) · `abandoned` (plan descartado: los números **no** se reciclan).

> **R-001**: plan creado antes de este contrato (sin bloque `numbering:` ni campo `fase`). Al sellarlo, la Fase C hace el *retrofit* una sola vez: añade `numbering:` al envelope y `- **fase**: 1420..1425` a `T1..T6`. El retrofit está permitido sólo durante el cierre, nunca durante la Fase B.

> **R-048..R-053**: revisiones manuales de las Rondas 332–337 (Fases 1651–1686) sobre el plan de ranuras dinámicas. No ejecutaron código: solo reporte y correcciones documentales. Se registran para fijar la numeración consumida. `R-047` sigue `executing` y su commit es precondición de cualquier reserva nueva.
> **Plan pendiente de reserva**: `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` (`T1..T10`, 8 olas). Sin fila en este registro hasta que `preflight` devuelva 0; su `ronda_base`/`fase_base` se calculan entonces a partir de §2.

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
