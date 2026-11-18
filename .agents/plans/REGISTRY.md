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
| `ronda` | 443 | 444 | 2026-11-18 |
| `fase` | 1970 | 1971 | 2026-11-18 |

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
| R-078 | 2026-11-14 | `2026-11-14-corrida92-completar-huecos-plan.md` | `Wave 1..8` → `Ronda 384–391` | `T1..T8` → `Fase 1798–1805` | closed | PASS (Full 1387) | `2587279` |
| R-079 | 2026-11-15 | `2026-11-15-libro8-corrida94-plan.md` | `Wave 1..6` → `Ronda 392–397` | `T1..T9` → `Fase 1806–1814` | closed | PASS (Full 1446) | `f3e95be` |
| R-080 | 2026-11-16 | `2026-11-15-planificador-limite-ventana-design.md` | `Wave —` → `Ronda 398` | `Fase 1815–1819` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `5676f5b` |
| R-081 | 2026-11-16 | `(revisión Ronda 399 — Fases 1820–1824: gate verify.ps1, citas vivas, cierres sin evidencia, registro)` | `Wave —` → `Ronda 399` | `Fase 1820–1824` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `5676f5b` |
| R-082 | 2026-11-16 | `(revisión Rondas 400–401 — Fases 1825–1828: gate verify.ps1 (ACL DENY DC), sha ancestro, fila faltante)` | `Wave —` → `Ronda 400–401` | `Fase 1825–1828` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `5676f5b` |
| R-083 | 2026-11-16 | `(revisión Rondas 402–403 — Fases 1829–1832: causa ACL, desviación de gate, SHA canónico)` | `Wave —` → `Ronda 402–403` | `Fase 1829–1832` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `5676f5b` |
| R-084 | 2026-11-16 | `(revisión Ronda 404 — Fases 1833–1835: HEAD canónico, fila R-083, desviación de gate con TRX)` | `Wave —` → `Ronda 404` | `Fase 1833–1835` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `5676f5b` |
| R-085 | 2026-11-16 | `(revisión Ronda 405 — Fases 1836–1838: respaldo blob, gate TRX 1584 tests, preflight adaptado)` | `Wave —` → `Ronda 405` | `Fase 1836–1838` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `5676f5b` |
| R-086 | 2026-11-16 | `(revisión Ronda 406 — Fases 1839–1841: commit de contenido, doc reconciliado, respaldo byte-a-byte)` | `Wave —` → `Ronda 406` | `Fase 1839–1841` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `534931d` |
| R-087 | 2026-11-16 | `(revisión Ronda 407 — Fases 1842–1843: SHA propio R-086, política .gitattributes EOL)` | `Wave —` → `Ronda 407` | `Fase 1842–1843` | closed | `BLOCKED-ACL (verify.ps1); evidencia: _evidence/20261116-gate-alt/gate-alternativo.trx (1584 pruebas, 0 fallas)` | `d97bc8a` |
| R-088 | 2026-11-17 | `(implementación general: aula aprobada sin origen productivo ⇒ demanda pendiente de programar)` | `Wave —` → `Ronda 408` | `Fase 1844` | closed | `BLOCKED-ACL (verify.ps1); gate alternativo por lotes: _evidence/20261117-aprobada-sin-origen (1629 pruebas, 0 fallas)` | `0596274` |
| R-089 | 2026-11-17 | `(empaquetado: Newtonsoft.Json declarado en .dna/MSI + contrato check-addin-packaging)` | `Wave —` → `Ronda 409` | `Fase 1845` | closed | `BLOCKED-ACL (verify.ps1); gate alternativo por lotes: 1629 pruebas, 0 fallas; EMPACADO_OK (bin 20 / dna 16 / wix 21)` | `2da0788` |
| R-090 | 2026-11-17 | `(revisión Ronda 410/411 — Fases 1846–1849: factibilidad, escalador decisión, fail-closed snapshot)` | `Wave —` → `Ronda 410–411` | `Fase 1846–1849` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-091 | 2026-11-17 | `(revisión Ronda 412 — Fases 1850–1853: flush alcance semáforo, snapshot divergente, trazabilidad y tests)` | `Wave —` → `Ronda 412` | `Fase 1850–1853` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-092 | 2026-11-17 | `(revisión Ronda 413 — Fases 1854–1857: extensión de período, fallo cerrado ventana, paridad demanda escalador)` | `Wave —` → `Ronda 413` | `Fase 1854–1857` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-093 | 2026-11-17 | `(revisión Ronda 414 — Fases 1858–1866: corrección tipado Result, asigsAula, Marcar combo, contabilidad demanda y no pérdida)` | `Wave —` → `Ronda 414` | `Fase 1858–1866` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-094 | 2026-11-17 | `(revisión Ronda 415 — Fases 1867–1869: filtro catálogo en simulación, DtpVentana tipado y escritura fuera de UI)` | `Wave —` → `Ronda 415` | `Fase 1867–1869` | closed | PASS (Full 1641, Release v1.6.4.21) | `60ea604` |
| R-095 | 2026-11-17 | `(revisión Ronda 416 — Fases 1870–1873: vacantes consentidas en certificado, propagación contexto y no throw post-staging)` | `Wave —` → `Ronda 416` | `Fase 1870–1873` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-096 | 2026-11-17 | `(revisión Ronda 417 — Fases 1874–1878: rehidratación consentimiento en compuerta, IdTercero <= 0 en cobertura, asserts estrictos y purga staging)` | `Wave —` → `Ronda 417` | `Fase 1874–1878` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-097 | 2026-11-17 | `(revisión Ronda 418 — Fases 1879–1883: corrección compilación tests, purga staging no certificable y columna tipada vacantes_consentidas)` | `Wave —` → `Ronda 418` | `Fase 1879–1883` | closed | revisión y correcciones; sin build/tests | pendiente |
| R-098 | 2026-11-17 | `(revisión Ronda 419 — Fases 1884–1886: using AgCorrida, migración V84 vacantes_consentidas, persistencia repo y hot path tipado)` | `Wave —` → `Ronda 419` | `Fase 1884–1886` | closed | revisión y correcciones (1884-1886); sin build/tests | pendiente |
| R-099 | 2026-11-17 | `(revisión Ronda 420 — Fases 1887–1888: serialización segura vacantes_consentidas, helper introspección unificado y test compatibilidad)` | `Wave —` → `Ronda 420` | `Fase 1887–1888` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-100 | 2026-11-17 | `(revisión Ronda 421 — Fases 1889–1891: tri-estado ExistenciaSql, TablaExiste unificado y saneamiento de identificador)` | `Wave —` → `Ronda 421` | `Fase 1889–1891` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-101 | 2026-11-17 | `(revisión Ronda 422 — Fases 1892–1894: propagación Indeterminada en V84, degradación segura en repositorio y exclusión prefijo sqlite_)` | `Wave —` → `Ronda 422` | `Fase 1892–1894` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-102 | 2026-11-17 | `(revisión Ronda 423 — Fases 1892–1894: verificación en v1.6.4.22, tri-estado y exclusión sqlite_)` | `Wave —` → `Ronda 423` | `Fase —` | closed | sin fases nuevas obligatorias; procede plan R424 | `v1.6.4.22` |
| R-103 | 2026-11-17 | `(implementación Ronda 424 — Fases 1895–1898: AG fuera de UI, logging pre-purga, detalle vacante y corrección ruta vacantes)` | `Wave —` → `Ronda 424` | `Fase 1895–1898` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-104 | 2026-11-17 | `(revisión Ronda 425 — Fases 1899–1903: vacantes docentes en R5, cierre diferido, snapshot contexto y tests)` | `Wave —` → `Ronda 425` | `Fase 1899–1903` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-105 | 2026-11-17 | `(revisión Ronda 426 — Fases 1904–1908: motivo vacío a R5, ciclo de vida cierre, guarda Invoke, snapshot total y bloqueo mutaciones)` | `Wave —` → `Ronda 426` | `Fase 1904–1908` | closed | sin escisión (bajo umbral); sin build/tests | `29cdd0f` |
| R-106 | 2026-11-17 | `(revisión manual Ronda 427 — Fases 1909–1919: diagnóstico forense vacantes, tipado enum regla y log)` | `Wave —` → `Ronda 427` | `Fase 1909–1919` | closed | retirada 1909; 1914/1915/1918 consolidadas en R428 | `29cdd0f` |
| R-107 | 2026-11-17 | `(revisión e implementación Ronda 428 — Fases 1920–1927: no bloqueo por vacantes en corrida/staging/GA, botón promoción con justificación, y diagnóstico legible)` | `Wave —` → `Ronda 428` | `Fase 1920–1927` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-108 | 2026-11-17 | `(revisión e implementación Ronda 429 — Fases 1928–1930: calidad vs promovible, limpieza código muerto y etiqueta precisa)` | `Wave —` → `Ronda 429` | `Fase 1928–1930` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-109 | 2026-11-17 | `(consolidado e implementación Ronda 430 — Fases 1931–1935: alcance completo corrida, cuadre independiente, identidad fila-vinculación, visibilidad homónimos y prueba última fecha)` | `Wave —` → `Ronda 430` | `Fase 1931–1935` | closed | PASS (Full 1648) | pendiente |
| R-110 | 2026-11-17 | `(revisión e implementación Ronda 431 — Fases 1936–1938: semántica banner cuadre, fallback clave IdVinculacion <= 0 y detección sesiones huérfanas)` | `Wave —` → `Ronda 431` | `Fase 1936–1938` | closed | PASS (8 tests SemestresExcelExporterTests) | pendiente |
| R-111 | 2026-11-17 | `(revisión e implementación Ronda 432 — Fases 1939–1941: discriminantes sede/subsede/periodo en clave fallback, precisión texto huérfanas y cuadre en módulo legado)` | `Wave —` → `Ronda 432` | `Fase 1939–1941` | closed | sin escisión (bajo umbral); sin build/tests | pendiente |
| R-112 | 2026-11-17 | `2026-11-17-rendimiento-app-corrida105-plan.md` | `Wave 1..6` → `Ronda 433–438` | `T1..T10` → `Fase 1942–1951` | closed | PASS (Gate Full 1714, AUDIT_PASS) | 2cd3f42 |
| R-113 | 2026-11-18 | `(consolidación Ronda 439 — Fases 1952–1961: reparación cruce, clave completa, telemetría y puertas P1/P2/recortes)` | `Wave —` → `Ronda 439` | `Fase 1952–1961` | executing | pendiente (abierta; sin build/tests) | pendiente |
| R-114 | 2026-11-18 | `(trazabilidad R439 y correcciones motor AG: duplicados índices, orden sellado cruce, bypass validador y coherencia mutación)` | `Wave —` → `Ronda 440` | `Fase 1962–1966` | executing | F1962 closed (documental); F1963-1966 pendiente | pendiente |
| R-115 | 2026-11-18 | `(aptitud heredada en cruce, sinceramiento R440 y neutralización estado copiado en HeredarOX)` | `Wave —` → `Ronda 441` | `Fase 1967–1969` | executing | F1967/F1969 pendientes; F1968 reabierta | pendiente |
| R-116 | 2026-11-18 | `(ajuste F1970 — escenario parametrizado por relaciones y prueba de rutas costosas)` | `Wave 442` → `Ronda 442` | `T1` → `Fase 1970` · `Fase 1968` (cerrada) | reserved | F1968 closed; F1970 reserved (sin build/tests) | pendiente |
| R-117 | 2026-11-18 | `(revisión Wave 442 — F1968 cerrada, F1970 reservada con addendas A1–A4; no procede build release)` | `Wave —` → `Ronda 443` | `Fase — (sin fases nuevas)` | executing | sin escisión (bajo umbral); sin build/tests | pendiente |




Estados: `reserved` (Fase A) · `executing` (Fase B) · `closed` (Fase C, sellado) · `abandoned` (plan descartado: los números **no** se reciclan).

> **R-001**: plan creado antes de este contrato (sin bloque `numbering:` ni campo `fase`). Al sellarlo, la Fase C hace el *retrofit* una sola vez: añade `numbering:` al envelope y `- **fase**: 1420..1425` a `T1..T6`. El retrofit está permitido sólo durante el cierre, nunca durante la Fase B.

> **R-048..R-053**: revisiones manuales de las Rondas 332–337 (Fases 1651–1686) sobre el plan de ranuras dinámicas. No ejecutaron código: solo reporte y correcciones documentales. Se registran para fijar la numeración consumida. `R-047` sigue `executing` y su commit es precondición de cualquier reserva nueva.
> **Plan pendiente de reserva**: `.agents/plans/2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md` (`T1..T10`, 8 olas). Sin fila en este registro hasta que `preflight` devuelva 0; su `ronda_base`/`fase_base` se calculan entonces a partir de §2.

> `HEAD` registra el **commit de contenido** de la ola (el que introduce el TRX, la fila y los artefactos),
> no el `HEAD` móvil de la rama. Requisito: `git merge-base --is-ancestor <sha> HEAD` = 0 y
> `git show --name-only <sha>` contiene los archivos de la ola. Cada fila nueva cita **su** commit de contenido;
> nunca se hereda el de la fila anterior. Prohibido emitir un commit extra sólo para actualizar este campo.
> Desviación autorizada: `verify.ps1` irrecuperable por ACL heredado `Todos:(I)(CI)(DENY)(DC)` (borrado fijado en índice). Respaldo del blob en `_evidence/20261116-gate-alt/verify.ps1.bak`.
> Hasta `icacls scripts /remove:d Todos`, el gate es `dotnet build` + `dotnet test` con TRX hacia `_evidence/20261116-gate-alt/gate-alternativo.trx` (1584 pruebas, 0 fallas); `preflight.ps1` adaptado para tolerar la desviación bajo advertencia WARN.

> **History rotation**: `implementation_plan.md` y `walkthrough.md` se escinden por rangos de fase según `.agents/rules/history-rotation.md` (herramienta `.agents/tools/split-history.ps1`; trozos en `docs/plans/history/`). Los archivos vivos conservan el índice y las últimas rondas.

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
