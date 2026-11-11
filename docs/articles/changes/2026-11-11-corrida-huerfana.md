# Corrida huérfana en el Módulo R: diagnóstico y reparación (Fases 1563–1566)

## 1. Síntoma

En el Módulo R la corrida **#48** (2027-01 · APT-52 · VALLE · 180 sesiones) se mostraba con:

- `Restricciones Hard: —` y el panel **«Violaciones de Reglas» vacío**.
- Botones «Aprobar corrida» y «Promover a producción» deshabilitados («promover no funciona»).
- Etiqueta del selector con el sufijo **`sin resultado AG`**.
- Log repetido:

```
[AG] Corrida #48: staging referencia individuo(s) [33] sin fila en ag_individuo.
[AG] La corrida #48 tiene staging pero no cabecera en ag_corrida ni individuo óptimo persistido.
```

## 2. Causa raíz (evidencia)

| Evidencia | Valor |
|---|---|
| `log_20261109.txt` | `FinCorridaAG IdCorrida=48 | Completada | Hard=95 | Vinculaciones=15 | Sesiones=180` |
| `log_20261109.txt` | `Corrida #48 aprobada (estado Aprobada). Promoción habilitada.` |
| `log_20261110.txt` | `[Purga] vinc 33->33 | ses 325->325 | gen 70->70 | corridas 2->1 | idPeriodo=3 idPrograma=1 corridasPeriodo=48` |
| Vault (`ag_vinculaciones`) | corrida 48 → 15 filas con **`idperiodo=3`, `idprograma=2`** |
| Vault (`ag_gen`) | 30 genes de `id_individuo=33`, **`idprograma=2`**, aulas 12/30/31 |
| Vault (`ag_corrida` / `ag_individuo`) | **sin** fila 48 / **sin** fila 33 |
| `programas` | `id=1 → APT-53`, `id=2 → APT-52` (**mismo nombre** «ADMINISTRACION PÚBLICA TERRITORIAL») |

`AgFabricaIndividuos` estampa cada gen con el **programa de la asignatura** (`asig.IdPrograma = 2`),
mientras `AgCorrida.IdPrograma` se toma del **combo del formulario** (`1`).

`AgScheduleRepository.LimpiarStagingAsync(idPeriodo, idPrograma)` borraba:

- `ag_corrida`, `ag_individuo`, `ag_violacion`, `ag_bitacora`, `ag_promocion` → por el par **de `ag_corrida`** `(3,1)`;
- `ag_vinculaciones`, `ag_sesiones`, `ag_gen` → por el par **de cada fila** `(3,2)`.

Con programas distintos, la cabecera y el individuo desaparecen y el staging sobrevive **huérfano**.
Sin `ag_individuo` no hay `ag_violacion` que leer (el detalle se borró con el individuo) y sin
`ag_corrida` la compuerta `Estado == Aprobada` de la promoción es inalcanzable.

## 3. Cambios aplicados

### Fase 1563 — Purga anclada a `id_corrida`
`AgScheduleRepository.LimpiarStagingAsync(int, int)`:

- `idsDel = ag_corrida(idperiodo,idprograma) ∪ ag_vinculaciones(idperiodo,idprograma)`.
- Todos los borrados (`ag_vinculacion_gen`, `ag_sesiones`, `ag_salvedades_conflictos`, `ag_vinculaciones`,
  `ag_violacion`, `ag_bitacora`, `ag_promocion`, `ag_gen`, `ag_individuo`, `ag_corrida`) usan
  `id_corrida IN @idsDel`.
- Limpieza adicional de genes/violaciones de individuos **referenciados por el staging** aunque su fila en
  `ag_individuo` ya no exista (huérfanos de purgas anteriores).
- Invariante: si tras la purga quedan `ag_vinculaciones`/`ag_sesiones` de esas corridas → `LoggerHelper.Error`.
- Guardas `TablaExiste` para vaults legados sin tablas opcionales.

### Fase 1564 — Fallar rápido en lugar de corromper en silencio
- `UpdateCorridaAsync` verifica filas afectadas: `0` filas ⇒ `Result.Failure("La corrida #N ya no existe en ag_corrida…")`.
- `EjecutarOptimizacionAsync` aborta antes de materializar staging si la cabecera no se pudo consolidar.
- Tras `SaveIndividuoAsync` se registra el número real de filas de `ag_violacion` frente al objeto en memoria.

### Fase 1565 — Observabilidad en el Módulo R
- `CargarModuloResultadosAsync` ya no descarta `corridaRes.Error`; distingue **cabecera ausente (huérfana)**,
  **cabecera sin individuo** y **lectura fallida**.
- En estado huérfano habilita **«Recalcular violaciones»** (vía de reparación) y explica el estado en `lblAgModo`.
- `PromoverAProduccionAsync` dejó de retornar en silencio: informa la causa y la acción recomendada.

### Fase 1566 — Reparación sin re-ejecutar el AG
Nuevos miembros:

- `IAgScheduleRepository.GetGenesByIndividuoAsync`, `DetectarCorridaHuerfanaAsync`,
  `ReinsertarCabeceraHuerfanaAsync`.
- `AgCorridaHuerfana`, `AgReparacionCorridaResultado` (Domain) y `AgReparacionCorridaService` (Application).

`AgReparacionCorridaService.RepararAsync`:

1. Deriva período/programa/subunidad/individuo **del staging**.
2. Carga los genes de `ag_gen` y **valida cobertura del contexto antes de escribir**.
3. Evalúa el individuo con `AgEvaluadorAptitud` (mismo evaluador del AG) → Fitness, HARD/SOFT y detalle.
4. Reinserta `ag_individuo` y `ag_corrida` con **ids explícitos** (33 y 48) para respetar el staging.
5. Persiste el detalle con `ReemplazarViolacionesAsync` (registra `[AG-Recalculo]`).

Se invoca desde el botón existente **«Recalcular violaciones»** cuando la corrida es huérfana:
no se añadieron controles nuevos a la interfaz.

### Fase 1567 — El nombre de la regla llegaba como «0»
`ag_violacion.regla` está declarada **INTEGER** en el vault, pero el código persiste el **nombre** de la regla
(`R1_SuperposicionDocente`, …). Sin `CAST`, System.Data.SQLite aplica la afinidad de la columna y devuelve `0`,
por lo que el panel mostraba «regla 0» y la vista por fechas perdía el identificador.

- `AgScheduleRepository.GetViolacionesByIndividuoAsync`: `CAST(regla AS TEXT) AS Regla`.
- `FormVerVinculaciones.CargarViolacionesAsync`: `CAST(v.regla AS TEXT) AS Regla`.
- Prueba de regresión con la columna declarada `INTEGER` (`ReemplazarViolaciones_ConColumnaReglaInteger_…`).

## 4. Procedimiento para recuperar la corrida #48

1. Abrir el AG → Módulo R → seleccionar `#48`.
2. En el árbol, seleccionar **período 2027-01 + programa APT-52** y las cohortes de las aulas 12/30/31.
3. Pulsar **«Recalcular violaciones»** → confirmar «Reparar corrida huérfana».
4. El panel muestra el desglose real (`ag_violacion`) y la cabecera queda en estado `Completada`.
5. Pulsar **«Aprobar corrida»**.

> **Compuerta de seguridad:** la promoción exige `HARD = 0` (`AgPromotorVinculaciones:85-88`).
> La #48 tiene `HARD = 95`, por lo que seguirá bloqueada por diseño hasta corregir la programación y
> volver a ejecutar el AG. La reparación devuelve la **visibilidad y el diagnóstico**, no relaja la compuerta.

## 5. Pruebas

- `ConversorPDF.Tests/Integration/Academic/AgScheduleRepositoryHuerfanasTests.cs`
  - `AgScheduleRepositoryPurgaTests`: purga del caso real (cabecera `(3,1)` + staging `(3,2)`) deja 0 filas en
    todas las tablas; purga de staging sin cabecera; `UpdateCorridaAsync` sobre id inexistente falla.
  - `AgScheduleRepositoryReparacionTests`: detección, genes huérfanos, reinserción con ids 48/33 e idempotencia.
- `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgReparacionCorridaServiceTests.cs`:
  aborta sin escribir si el contexto no cubre las aulas; éxito persiste violaciones una sola vez.

## 6. Archivos afectados

- `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgScheduleRepository.cs`
- `ConversorPDF.XLL/Academic/Domain/Interfaces/IAgScheduleRepository.cs`
- `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgCorridaHuerfana.cs` (nuevo)
- `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgReparacionCorridaResultado.cs` (nuevo)
- `ConversorPDF.XLL/Academic/Application/Services/Genetica/IAgReparacionCorridaService.cs` (nuevo)
- `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgReparacionCorridaService.cs` (nuevo)
- `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- `ConversorPDF.XLL/ConversorPDF.XLL.csproj`

