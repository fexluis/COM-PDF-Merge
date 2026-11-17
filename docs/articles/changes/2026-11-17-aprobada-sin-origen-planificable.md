# Aula aprobada sin origen productivo: demanda obligatoria, nunca bloqueo

## Propósito

Eliminar el callejón sin salida que dejaba un aula en estado `Aprobado` cuando **no existe ninguna fila
productiva** para su cohorte (período + sede + programa + semestre + grupo + modalidad) en la ventana.

Antes: el aula aprobada obligaba a leer producción, la lectura no encontraba nada, la validación la marcaba
`PLAN_NO_VERIFICABLE` y la corrida quedaba `Bloqueado` con `REVISION_RECHAZADA`. El aula no se podía
conservar (no había origen) ni programar (`PuedeGenerarAula` devolvía `false`).

Ahora: esa situación se reconoce como **demanda pendiente de programar**. El aula se planifica con sus
oleadas igual que una `Disponible`; no se elimina, no se inactiva y no se inventa origen.

## Regla general (no depende de ningún aula concreta)

1. Toda aula con relaciones `ag_aula_asignatura.requerida = 1` y horas tiene demanda obligatoria.
2. El estado del aula **sólo** decide si se conserva un plan existente; nunca si se planifica.
3. `Aprobado` + ninguna fila productiva reconciliada ⇒ `PendienteDeProgramar`.
4. `Aprobado` + fila productiva (aunque incompleta, ambigua o fuera de ventana) ⇒ se mantiene el régimen
   estricto anterior (`PlanNoVerificable` / `RequiereRatificacion`). No se relaja un origen parcial.
5. `Aprobado` + origen completo ⇒ exige ratificación T4 antes de conservar.

## Cambios

| Archivo | Cambio |
|---|---|
| `Academic/Domain/Models/Genetica/AgValidacionPlanAprobado.cs` | Nuevo estado `PendienteDeProgramar`; `BloqueaPreparacion` sólo considera `PlanNoVerificable` y `RequiereRatificacion`. |
| `Academic/Domain/Models/Genetica/AgContexto.cs` | Nuevo `TieneOrigenConservable(idAula)`; `PuedeGenerarAula` admite genes para una aprobada sin origen. |
| `Academic/Application/Services/Genetica/AgValidadorPlanAprobado.cs` | Clasificación por aula: globales → sin origen (`PendienteDeProgramar`) → cobertura/ratificación. |
| `Academic/Application/Services/Genetica/AgProyectorSolucionCobertura.cs` | `Conservar` excluye las aulas pendientes: no exige decisión publicada para lo que no existe. |
| `Academic/Application/Services/Genetica/AgGeneradorHorarios.cs` | `ValidarAprobadas` sólo exige conservación para aprobadas con origen. |
| `Academic/UI/Forms/AgFormGeneradorHorarios.cs` | Separa `conservables` de `pendientesProgramacion`, registra `[PLAN CANÓNICO] … estado=PendienteDeProgramar origen=AUSENTE acción=planificar` y no bloquea la preparación por ellas. |

## Cómo verificarlo

1. Perfil con un aula aprobada cuyo período no tiene filas en `vinculaciones`/`sesiones`.
2. Ejecutar la corrida: el log muestra el aula como `PendienteDeProgramar` y **no** aparece
   `REVISION_RECHAZADA`/`estado=Bloqueado` por ese motivo.
3. `ag_sesiones` recibe las oleadas del aula dentro de la ventana semestral.
4. Si la demanda no cabe en el carril, se mantiene el flujo existente de escalado con confirmación humana
   (`AgEscaleraVentanas.Siguiente`, p. ej. `MAR-VIE → LUN-VIE`); no se recorta la cobertura.
5. Tras aprobar y promover, el aula vuelve a `Aprobado` con origen productivo real y el régimen estricto
   aplica sin cambios.

## Cobertura de pruebas

`ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAprobadaSinOrigenTests.cs`
(10 casos: `TieneOrigenConservable`, `PuedeGenerarAula` en sus cuatro rutas, clasificación del validador
sin origen y con origen parcial, `Conservar` sin origen y gate `ValidarAprobadas` abierto/cerrado).

`AgValidadorPlanAprobadoTests` se ajusta al requisito nuevo: un manifiesto sin ninguna vinculación es
demanda pendiente (antes `PlanNoVerificable`), mientras que una vinculación existente sin sesiones sigue
siendo `PlanNoVerificable`.

## Límites

- No autoriza conservar, no sustituye T4 y no modifica `ag_aula` por sí mismo.
- La corrida sigue exigiendo factibilidad: si el carril no alcanza, se requiere el escalón confirmado.
- Los festivos `PRESENCIAL` mal asignados a otro período siguen afectando sólo viáticos y la alerta
  `Presenciales: 0`; no bloquean la planificación.
