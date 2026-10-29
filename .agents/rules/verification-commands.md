# Contrato de Verificación y Comandos Estándar

> [!IMPORTANT]
> **Trigger**: `always_on`. Todo agente, orquestador o revisor debe utilizar estrictamente estos comandos para certificar compilación y pruebas.

---

## 1. Comandos Congelados

Las tres formas oficiales de verificar el código en el repositorio son:

### 1.0 Preflight de Workspace (obligatorio antes de planificar o ejecutar)
Valida que el workspace sea una copia **real** del árbol (tamaño, slice, baseline coherente):
```powershell
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <ruta> [-Plan <plan.md>] [-AllowOneDrive]
```
Exit codes: `0` OK | `1` workspace incompleto (no es copia del árbol real) | `2` faltan prerequisitos | `3` workspace en OneDrive.

> Un `git clone` de un repositorio con working tree sucio **no** es una copia válida: solo trae HEAD.

### 1.1 Verificación de Compilación (Build Gate)
Compila ambos proyectos (`ConversorPDF.XLL` y `ConversorPDF.Tests`) en modo `Debug|x64`:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
```

### 1.2 Verificación Rápida de Ciclo TDD (Fast Gate)
Ejecuta únicamente las pruebas relacionadas con la clase bajo prueba (SUT):
```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~<ClaseBajoPrueba>"
```

### 1.3 Verificación Completa de la Suite (Full Gate)
Ejecuta la totalidad de la suite del workspace, genera el archivo TRX y contrasta con el baseline (regresiones **y** tamaño de suite):
```powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
```

### 1.4 Auditoría de Código por Tarea (Audit Gate)
Complementa a `verify.ps1`: no compila, solo analiza el código de las tareas de un plan con detectores deterministas de fallos de ejecución (bloqueos STA, Interop en capas limpias, SQL destructivo en tablas CRDT, migraciones duplicadas, tests debilitados…):
```powershell
powershell -ExecutionPolicy Bypass -File scripts/audit-tasks.ps1 -Plan .agents\plans\<plan>.md [-Tasks T1,T2]
```
Exit codes: `0` sin hallazgos bloqueantes (puede haber `Warning`) | `3` ≥1 hallazgo `Compilation`/`Runtime`/`Logic` | `4` entrada inválida.

Evidencia: `_evidence/<timestamp>-audit/audit.md` con la tabla `| Fase | Task | Severidad | Archivo | Líneas |`. Es obligatorio en la Fase C (`swarm-verify`) y recomendado en la revisión por rondas (`.clinerules/workflows/Ronda.md`).

---

## 2. Contrato de Evidencia RED -> GREEN

1. **Evidencia de Fase RED**:
   - El test nuevo o actualizado DEBE compilar.
   - Debe fallar **exactamente por la razón esperada** (por ejemplo, `NotImplementedException` o aserción de negocio insatisfecha).
   - Se debe registrar en el ledger o reporte el nombre completo de la prueba y el mensaje de error exacto extraído de `summary.md` / `test.trx`.
2. **Evidencia de Fase GREEN**:
   - El mismo test debe pasar a estado `Passed`.
   - La suite completa (`-Mode Full`) debe ejecutarse y terminar con exit code `0` (0 regresiones vs baseline).
   - La ruta del directorio `_evidence/<timestamp>` debe vincularse en el ledger.

---

## 3. Prácticas Estrictamente Prohibidas
- ❌ Aplicar `--filter` que excluya intencionalmente pruebas existentes para disimular roturas.
- ❌ Modificar aserciones existentes o debilitar condiciones de prueba para lograr un falso verde.
- ❌ Aumentar timeouts arbitrarios para disimular problemas de concurrencia o locks COM.
- ❌ Declarar completada una tarea sin adjuntar la ruta de evidencia generada por `verify.ps1`.

---

## 4. Reglas de Integridad del Gate

- `Full` devuelve `3` si la suite ejecutada es menor que `TOTAL` del baseline (`SUITE_SHRUNK`): hacer desaparecer tests no es "verde".
- `Full` exige baseline; sin `tests/baseline/test_baseline.txt` devuelve `3` (`NO_BASELINE`).
- `Full` devuelve `2` si el TRX no trae resultados o es anterior a la ejecución (`INFRA_ERROR`).
- El gate es **serial**: nunca dos `verify.ps1` concurrentes en el mismo workspace (Excel-DNA bloquea `bin/*.xll`, error `DNA-577853131`).
- `summary.md` registra `HEAD`, archivos con cambios y veredicto (`PASS|RED|REGRESSION|SUITE_SHRUNK|NO_BASELINE|INFRA_ERROR`).
- La evidencia vive en `_evidence/` (ignorada por Git) y se enlaza por ruta en el ledger.
