# Recuperación de planificación por ranuras — 2026-11-14

Estado: implementación reconstruida e integrada; sin compilación ni ejecución de pruebas. Sin nueva reserva de ronda/fases. No es cierre de las fases 1788–1793.

## Qué ocurrió

Al iniciar esta recuperación faltaban AgPlanificadorRanuras.cs, AgSesionPlanificada.cs y AgHuellaPlanVentana.cs, sus inclusiones en el csproj y los contratos de sesiones del plan. No aparecen en el historial Git consultado. HEAD: f6e8e35. Eran cambios sin confirmar de esta conversación. Git no identifica cuándo ni quién retiró esos cambios del directorio de trabajo. No se atribuye la retirada a una persona/modelo ni a OneDrive.

Los respaldos existentes bajo docs/plans/_backup-escalado-1788 no contenían los tres archivos nuevos. Se reconstruyeron contra los contratos actuales. No se hizo reset, checkout, limpieza, publicación ni instalación.

## Cambios recuperados

- AgSesionPlanificada: asignatura, fecha y ranura inmutables. AgPlanOleada transporta una colección de sesiones de solo lectura.
- AgPlanificadorRanuras: calendario real sin festivos; búsqueda de bandas con presupuesto de 12000 estados y cancelación; resultado parcial con pendientes y búsqueda agotada separados de factibilidad.
- AgPlanificadorOleadas: materializa la referencia y la compara con el candidato. Busca compactación incluso si la referencia ya es factible en una ventana ampliada. Conserva referencia si no hay mejora completa; mantiene su contrato público de Failure/planes completos. La evaluación interna admite parciales fechados.
- Fábrica: consume las mismas fechas por asignatura y ranura. Ignora como restricción los textos de patrones AUTOMATICO. FIJO/LEGADO continúan restringiendo días.
- AgHuellaPlanVentana: SHA-256 canónico por contenido de ventana, aula, programas, asignaturas, relaciones, bloques y calendario. Se integra en la clave de caché de la fábrica.
- Evaluador: pendientes y ocio derivados de sesiones realmente ubicadas. Fecha proyectada diagnóstica con horizonte de dos años; ninguna modificación de las fechas recibidas.
- Escalador/validador/modal: consumen esa evaluación; el modal usa pendientes posteriores reales y denomina la fecha proyectada.
- Calendario: recorridos implicados protegidos al alcanzar DateTime.MaxValue.
- csproj: cada nuevo archivo incluido exactamente una vez; versionado preexistente preservado.

## Alcance y límites

Actualización posterior a la recuperación: se implementó la instantánea SQL completa, la comprobación del contexto vigente y perfil, la reaplicación idempotente y el refuerzo transaccional de propuestas. Detalle en docs/articles/changes/2026-11-14-huellas-escalado.md. La validación por ejecución sigue pendiente; no se declara cierre verificado de la fase 1791.

El planificador recibe demanda, patrones, bloques y calendario; no recibe disponibilidad global de docentes ni todas las restricciones de aulas físicas. Un plan local fechado no certifica la factibilidad global de la corrida. La banda única puede utilizar más días por materia que los carriles canónicos. Deben verificarse reglas académicas adicionales y regresiones de genes/coherencia antes de release.

El límite utilizado es ctx.Ventana.FechaFin. Si se sobrescribió antes el plazo original en BD, estos archivos no recuperan por sí solos ese valor histórico. La fecha de contexto no se amplía ni reduce automáticamente.

## Pruebas preparadas

Archivo: ConversorPDF.Tests/Academic/Application/Services/Genetica/AgPlanificadorRanurasTests.cs.

- Demandas nominales de CALI III (256 h/64 sesiones) y IV (272 h/68 sesiones), con calendario sintético y sin restricciones docentes.
- Compactación aun con ventana ampliada a julio y referencia factible.
- Fábrica conserva exactamente las fechas del plan y las horas totales.
- Tres escalones de capacidad 4W a 5W; dos evaluaciones separadas VIE-DOM → JUE-DOM → MIE-DOM para demanda 6W.
- Patrones fijos viernes preservados; pendientes explícitos.
- Unidades 4/8 h y unicidad de franjas de sábado.
- Cambio de contenido con mismos conteos invalida huella/caché.
- Presupuesto agotado y cancelación.
- Última fecha representable del calendario.

No representan una repetición validada de la corrida #77. No se modificó Libro4.xlsx ni BD operativa.

## Verificación y siguiente paso

Revisión estática de inclusiones y git diff --check sobre archivos de recuperación. No se ejecutó proyecto, compilador, pruebas, Excel ni instalador. El usuario solicitó commit y push; consultar el historial Git para la referencia de entrega. Existe copia previa y copia recuperada fuera del proyecto en la carpeta de trabajo de esta conversación.

Tras autorización: Fast dirigido a AgPlanificadorRanurasTests y regresiones de fábrica/oleadas/escalador/validador; Full Debug/x64; Full Release/x64; smoke sobre copia de BD con datos de la corrida #77. Corregir cualquier regresión de contrato sin debilitar aserciones para aprobar. Release no acreditado.
