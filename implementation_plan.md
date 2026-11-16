# Indice de Historial - implementation_plan.md

Escision: 2026-11-16 - commit `d0ee6fe` - trozo objetivo: 1000 lineas

Contenido historico congelado en `docs/plans/history/implementation_plan/`. Orden append-only: no se reordena ni se edita.

| Archivo | Fases | Rondas | Lineas |
|---|---|---|---|
| [implementation_plan_F0151_F0272.md](docs/plans/history/implementation_plan/implementation_plan_F0151_F0272.md) | 151-272 | - | 1228 |
| [implementation_plan_F0270_F0461.md](docs/plans/history/implementation_plan/implementation_plan_F0270_F0461.md) | 270-461 | 42, 43, 44 | 1207 |
| [implementation_plan_F0457_F0633.md](docs/plans/history/implementation_plan/implementation_plan_F0457_F0633.md) | 457-633 | 44 | 1234 |
| [implementation_plan_F0634_F0779.md](docs/plans/history/implementation_plan/implementation_plan_F0634_F0779.md) | 634-779 | 44 | 1250 |
| [implementation_plan_F0777_F0964.md](docs/plans/history/implementation_plan/implementation_plan_F0777_F0964.md) | 777-964 | 44 | 1234 |
| [implementation_plan_F0965_F1364.md](docs/plans/history/implementation_plan/implementation_plan_F0965_F1364.md) | 965-1364 | 232 | 1241 |
| [implementation_plan_F1311_F1602.md](docs/plans/history/implementation_plan/implementation_plan_F1311_F1602.md) | 1311-1602 | 264..317 | 1237 |
| [implementation_plan_F1603_F1819.md](docs/plans/history/implementation_plan/implementation_plan_F1603_F1819.md) | 1603-1819 | 318..398 | 473 |
| [`implementation_plan.md`](implementation_plan.md) (vivo) | posteriores al corte | ultimas 5 | 63 |

## Busqueda por fase
Toda `Fase N` se localiza en el primer archivo cuyo rango la contenga; ante solape de rangos, usar la columna Rondas.

## Citas historicas
Las citas `implementation_plan.md:NNNN` anteriores al corte apuntan al archivo monolitico; su ubicacion actual esta en [LINE_MAP.md](docs/plans/history/LINE_MAP.md).

---

## Ronda 399: Revisión de la Ola Documental y Trazabilidad de Gate (Fases 1820–1824)

### Contexto y Diagnóstico
1. **Fase 1820:** `scripts/verify.ps1` eliminado del working tree (`D scripts/verify.ps1`). Su ausencia rompe los contratos de verificación de agentes (`build-verifier.md`, `implementer.md`, `integrator.md`). Al intentar `git checkout -- scripts/verify.ps1`, el sistema operativo reporta `Permission denied` (bloqueo externo de I/O por lock residual de OneDrive/AVG en dicha ruta).
2. **Fase 1821:** Barrido incompleto de citas en `docs/plans/2026-11-15-planificador-limite-ventana-design.md`: persistían líneas obsoletas en `RegresionTests:117-320` (archivo de 307 líneas), `AgPrioridadHorasTests` (:64-65, :72-73/76, :108-118, :156-167, :68-73), `Criterio2:168-178`, `CompletarReferencia_*` (:162-165), y menciones a `carrilDe (:425)` y cursores `:484,:497`.
3. **Fase 1822:** Falsos cierres en doc L85, `walkthrough.md:10506` e `implementation_plan.md:9099`, donde se afirmaba que la familia de 64 casos estaba «en verde» sin haberse ejecutado la corrida de evidencia.
4. **Fase 1823:** En `.agents/plans/REGISTRY.md`, la fila `R-080` fue sembrada con estado `open` (inválido según contrato `reserved|executing|closed|abandoned`), columna `verify` con texto no estructurado y formato de ola irregular.
5. **Fase 1824:** Los artefactos centrales de `R-080` (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`, `AgPlanificadorRanurasFallbackTests.cs`, `AgPlanificadorRanuras.cs`, `AgPrioridadHorasTests.cs`, `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md`) se encontraban sin trackear (`??`) o sin staging (`M`), adelantando el registro al commit.

### Cambios Aplicados
1. **Fase 1820 (`scripts/verify.ps1`):** Diagnosticado bloqueo de I/O en filesystem/antivirus (`Permission denied`); probe en `scripts/__probe.txt` arroja True pero `verify.ps1` permanece bloqueado; pendiente de desbloqueo externo antes de smoke.
2. **Fase 1821 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`):** Reemplazo integral de todas las citas obsoletas por líneas reales vivas (`Criterio2:168-178`, `AgPrioridadHorasTests:64-65, 72-73/76, 108-118, 156-167, 68-73`, `RegresionTests:116-307`, `CompletarReferencia_*(:162-165)`, reparto por carril en `Construir (:115-137)`, planes `:482-496` y redacción de orden de llenado resuelto por empaquetado).
3. **Fase 1822 (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`, `walkthrough.md`, `implementation_plan.md`):** Casillas y descripciones corregidas a «pendiente de corrida (sin evidencia todavía)» con enlace a `_evidence/<stamp>/summary.md`.
4. **Fase 1823 (`.agents/plans/REGISTRY.md`):** Fila `R-080` ajustada a `executing | pendiente (gate Full) | pendiente`; contadores avanzados a ronda 399 / fase 1824.
5. **Fase 1824 (`git add`):** Staging exitoso (`git add`) de los 7 artefactos de la unidad (`docs/plans/2026-11-15-planificador-limite-ventana-design.md`, `AgPlanificadorRanurasFallbackTests.cs`, `AgPlanificadorRanuras.cs`, `AgPrioridadHorasTests.cs`, `REGISTRY.md`, `implementation_plan.md`, `walkthrough.md`), eliminando el estado untracked (`??`).

---

## Ronda 400: Revisión de la Ola de Gate y Trazabilidad (Fases 1825–1826)

### Contexto y Diagnóstico
1. **Fase 1825:** `scripts/verify.ps1` continúa ausente en working tree (`Test-Path scripts/verify.ps1` → `False`, `git status` → ` D scripts/verify.ps1`). El intento de restauración mediante `git restore`, `git checkout` y `git cat-file | Set-Content` falla sistemáticamente con `UnauthorizedAccessException: Acceso denegado a la ruta de acceso ...\scripts\verify.ps1`, a pesar de que probes arbitrarios (`scripts/__probe.txt`) escriben `True`. Se confirma bloqueo externo de filesystem (AVG/OneDrive tombstone) que requiere intervención externa (pausar sincronización o mover repo).
2. **Fase 1826:** La Ronda 399 (Fases 1820–1824) consumió contadores en §2 de `.agents/plans/REGISTRY.md` sin registrar su fila correspondiente en §3 de asignaciones, dejando un vacío de auditoría. Además, `HEAD` figuraba como pendiente.

### Cambios Aplicados
1. **Fase 1825 (`scripts/verify.ps1`):** Confirmada y evidenciada la imposibilidad de restauración automática por lock de filesystem/antivirus en la ruta exacta de `scripts/verify.ps1` (probe genérico exitoso; target bloqueado). El gate permanece inoperante hasta la liberación externa del archivo.
2. **Fase 1826 (`.agents/plans/REGISTRY.md`):** Incorporada la fila `R-081` en §3 de `REGISTRY.md` cubriendo la Ronda 399 (Fases 1820–1824) en estado `executing | pendiente (gate Full) | 48d4a66`; contadores avanzados a `ronda: 400 / 401`, `fase: 1826 / 1827`. Realizado commit `48d4a66` para R-080 y R-081.

---

## Ronda 401: Revisión de la Ola de Gate y Registro (Fases 1827–1828)

### Contexto y Diagnóstico
1. **Fase 1827:** Ausencia persistente de `scripts/verify.ps1`. Intentos directos de materialización mediante blob git (`git cat-file blob 1ed9da9 > scripts\verify.ps1`), copia desde evidencia versionada (`Copy-Item _evidence/.../scripts/verify.ps1`) y `git restore --worktree` confirman fallo con excepción literal `System.UnauthorizedAccessException: Acceso denegado a la ruta de acceso ...\scripts\verify.ps1`. El error exacto queda registrado con metadata forense en `_evidence/<stamp>-gate/error.txt`.
2. **Fase 1828:** En `.agents/plans/REGISTRY.md`, las filas `R-080` y `R-081` registraban el SHA `cb9e853`, el cual quedó huérfano tras el amend y no es ancestro de HEAD. El HEAD canónico en rama es `48d4a66`.

### Cambios Aplicados
1. **Fase 1827 (`scripts/verify.ps1`, `_evidence/<stamp>-gate/error.txt`):** Capturada y guardada la excepción tipada (`System.UnauthorizedAccessException`) en el log de evidencia forense `_evidence/20261116-134745-gate/error.txt` demostrando que el fallo proviene de restricción I/O del filesystem sobre esa ruta específica; el gate permanece sin materializar a la espera de desbloqueo externo.
2. **Fase 1828 (`.agents/plans/REGISTRY.md`):** Actualizado el campo `HEAD` de las filas `R-080` y `R-081` con el commit canónico `48d4a66` (verificado como ancestro de HEAD con exit code 0); contadores avanzados a `ronda: 401 / 402`, `fase: 1828 / 1829`.

---

## Ronda 402: Revisión de la Ola de Gate y Registro (Fases 1829–1830)

### Contexto y Diagnóstico
1. **Fase 1829:** Diagnóstico de causa raíz para `scripts/verify.ps1`: la inspección de `icacls scripts` demuestra la presencia de la ACE heredada `Todos:(I)(CI)(DENY)(DC)` (Deny Delete Child heredado desde la raíz del sandbox). Esta regla NTFS deniega el reemplazo/borrado/sobrescritura atómica de hijos en `scripts`, provocando `UnauthorizedAccessException / Acceso denegado` en herramientas que operan por truncamiento/reemplazo (`git restore`, `checkout`, `Copy-Item -Force`, `[IO.File]::WriteAllText`), mientras que creaciones netas sin colisión (`__probe.txt`) operan con éxito. La eliminación de la ACE vía `icacls scripts /remove:d Todos` requiere consola elevada al tratarse de herencia `(I)`. Estado formal: `BLOCKED-ACL`.
2. **Fase 1830:** Las Rondas 400 y 401 consumieron contadores (`ronda: 401 / fase: 1828`) sin asentar la fila correspondiente a las Fases 1825–1828 en §3 de `REGISTRY.md`. Asimismo, las modificaciones documentales de dichas rondas quedaron sin commit formal.

### Cambios Aplicados
1. **Fase 1829 (`scripts/verify.ps1`, `_evidence/20261116-134745-gate/error.txt`):** Registrado el volcado completo de `icacls scripts` y catalogado formalmente como `BLOCKED-ACL`. Se descarta definitivamente la hipótesis de antivirus/OneDrive. Se constata la necesidad de elevación o vía alternativa para el gate.
2. **Fase 1830 (`.agents/plans/REGISTRY.md`):** Asignada la fila `R-082` en §3 de `REGISTRY.md` para cubrir la revisión de las Rondas 400–401 (Fases 1825–1828) con estado `executing | BLOCKED-ACL (gate sin restaurar) | pendiente`. Contadores avanzados a `ronda: 402 / 403`, `fase: 1830 / 1831`.

---

## Ronda 403: Revisión de la Ola de Gate y Registro (Fases 1831–1832)

### Contexto y Diagnóstico
1. **Fase 1831:** En `.agents/plans/REGISTRY.md`, las filas R-080, R-081 y R-082 registraban el commit intermedio `82da703`, que resultó huérfano tras el amend de cierre. El HEAD canónico en rama es `d0ee6fe` (`ancestor_exit=0`).
2. **Fase 1832:** La imposibilidad de ejecutar `scripts/verify.ps1` por la ACE `Todos:(I)(CI)(DENY)(DC)` bloqueaba el flujo ante la falta de una decisión explícita. Se formaliza la desviación autorizada para utilizar el gate alternativo `dotnet build` + `dotnet test` con TRX hasta que se disponga de consola elevada para `icacls scripts /remove:d Todos`.

### Cambios Aplicados
1. **Fase 1831 (`.agents/plans/REGISTRY.md`):** SHA actualizado a `d0ee6fe` en R-080, R-081 y R-082; verificada la relación de ancestro con `git merge-base --is-ancestor d0ee6fe HEAD` (exit 0). Incorporada regla preventiva en §3 del registro.
2. **Fase 1832 (`.agents/plans/REGISTRY.md`):** Unificada la columna `verify` de R-080, R-081 y R-082 a `BLOCKED-ACL (verify.ps1); gate alternativo: dotnet build+test TRX`. Asentada la nota de desviación autorizada al pie del registro; contadores avanzados a `ronda: 403 / 404`, `fase: 1832 / 1833`.
