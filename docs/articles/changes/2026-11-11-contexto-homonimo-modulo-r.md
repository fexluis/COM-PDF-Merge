# Contexto homónimo en el Módulo R: recálculo y detalle staging (Fases 1575–1578)

## 1. Síntoma

Con el validador de oleadas ya desplegado (Ronda 308), la interfaz seguía mostrando `125 HARD` y el panel con
101 filas de `R14/R2/R3/R4`, y además fallaban dos acciones:

- **«Recalcular violaciones»** → *“El contexto del AG no corresponde a la corrida #48: contexto=3/1, corrida=3/2.
  Seleccione el período y programa de la corrida y vuelva a intentarlo.”*
- **«Detalle (AG)»** → *“No se encontraron sesiones registradas en la vista de detalle staging (vw_AgSesionesDetalle).”*

## 2. Causa raíz

`ConstruirContextoAsync` fuerza el programa de la corrida (`idProgramaForzado = 2`) y **después** aplicaba la
resolución de homónimos (Fases 446/477): al existir varios programas con el **mismo nombre**
(`APT-53 id=1` y `APT-52 id=2`), sustituía el programa solicitado por el que tuviera más asignaturas del semestre
(1). El contexto quedaba en `3/1`.

Consecuencias:

1. `AgRecalculoViolacionesService` (Fase 1532) **aborta** por el desajuste de programa ⇒ el validador nuevo nunca
   se ejecutaba y el panel seguía leyendo las 101 filas persistidas el 11/11 a las 14:18 (producidas por el
   validador **antiguo**). No había regresión: era dato obsoleto.
2. `AbrirDetalleSesionesModoAgAsync` filtraba `vw_AgSesionesDetalle` con el programa del contexto (1) mientras el
   staging de la corrida es programa (2) ⇒ 0 sesiones.
3. `AgReparacionCorridaService` **no** validaba el programa del contexto, por lo que la reparación inicial persistió
   violaciones calculadas con el pensum equivocado.

## 3. Cambios aplicados

| Fase | Archivo | Cambio |
|---|---|---|
| **1575** | `AgFormGeneradorHorarios.ConstruirContextoAsync` | Con `idProgramaForzado > 0` se **omite** la resolución de homónimos y `ProgramasAlcance` se fija al programa de la corrida. Traza `ProgramaHomonimoOmitido`. |
| **1576** | `AgFormGeneradorHorarios.AbrirDetalleSesionesModoAgAsync` | Si hay corrida cargada, período y programa se leen de `ag_corrida`, no del árbol/contexto. |
| **1577** | `AgReparacionCorridaService.RepararAsync` | Rechaza el contexto cuyo período/programa no coincida con la corrida, **antes de escribir nada** (mensaje accionable citando los ids). |
| **1578** | Tests + verificación | `RepararAsync_ConContextoDeOtroPrograma_FallaSinEscribir` y arnés de recálculo real. |

## 4. Verificación sobre el vault real (copia, solo lectura)

| Momento | `ag_corrida#48.hard` | `ag_individuo#33.hard` | filas `ag_violacion` |
|---|---|---|---|
| Antes | 125 | 125 | 101 |
| Después de `RecalcularAsync(48, contexto 3/2)` | **0** | **0** | **0** |

- Genes evaluados: **30** · desglose: *“sin violaciones”*.
- El contexto usado fue `3/2` (el de la corrida), demostrando que F1575 elimina la sustitución por el homónimo.
- Suite completa: **1123 tests** en verde (más el nuevo caso de F1578).

## 5. Lección

La Ronda 308 validó el **validador**, pero no la cadena **UI → contexto → recálculo**. Con programas homónimos
(mismo nombre, distinto `id`) el contexto debe **obedecer siempre a la corrida** cuando se recalcula o repara, y
cualquier servicio que evalúe genes debe rechazar contextos de otro período/programa.
