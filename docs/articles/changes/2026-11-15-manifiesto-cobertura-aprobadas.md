# T2 — Manifiesto, integración UI y validación canónica (parcial, no verificado)

## Estado y límites

Soporte parcial de **S1, S2, S6 y S7**. **S8** respetada: no se ejecutaron pruebas, compilación, SQL, aplicación ni acceso a bases reales. No se declara resuelto I5 ni completada T2. Se preservaron las unidades anteriores y los cambios ajenos; esta unidad no implementa T3/T4/T5/T6.

**Toda aula aprobada sigue bloqueada en preparación.** Un origen reconciliado queda `REQUIERE_RATIFICACION`: faltan la revisión y ratificación auditadas de T4. Tener horas suficientes no concede conservación ni permiso de reemplazo. Origen desconocido, extras, duplicados, sesiones inválidas o ausencia de plan producen `PLAN_NO_VERIFICABLE`, con detalle para reconciliar. No se convierte el estado del aula a Disponible ni se ocultan aulas. No se ofrece todavía la propuesta MAR-VIE → LUN-VIE para aprobadas: corresponde a T3, sobre un origen ratificable y con la protección durable T4.

## Evidencia de esquema y composición

Lectura exclusivamente de archivos del repositorio:

- `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Sql/AcademicSchema.sql:91–113,147–188`: territorial/subunidad y dimensiones de vinculaciones; sesiones con IDs canónicos, fecha, horario descriptivo, horas y docente. No hay `idaula`, estado ni revisión de aprobación productiva.
- `MigrationV14_UpdateSesionesSchema.cs:42–53` confirma horario descriptivo. El lector existente reutiliza `HorarioSesionParser.TryParse`; no se fabrican intervalos ante horarios ilegibles.
- `MigrationV49_AddAgHorariosTables.cs:235–249`: `ag_promocion` enlaza staging/producción, pero no acredita por sí solo una revisión de negocio ni ratificación. No se amplió el SQL del lector ni se infirió aprobación de esos enlaces.
- `AgScheduleRepository.cs:21–26` y el constructor de `AgFormGeneradorHorarios` establecen `Func<IDbConnection>` como convención. La UI crea `AgPlanAprobadoRepository` con esa fábrica; admite inyectar `IAgPlanAprobadoRepository` por constructor. No fue necesario editar una fábrica/registro externo.

Las rutas abreviadas de migraciones/repositorios anteriores están bajo `ConversorPDF.XLL/Academic/Infrastructure/Persistence/`.

## Contratos preservados de la unidad previa

- `AgConstructorManifiestoCobertura.Construir` recibe IDs explícitos, territorial/subunidad y ventana. Reutiliza `AgEvaluadorDemandaVentana.HorasObjetivo`, conserva demanda original y efectiva separadas y exige todas las relaciones `Requerida == 1`. No deriva demanda de genes/staging ni completa selección mediante primera aula.
- La clave incluye territorial, subunidad, período, sede, programa, aula, asignatura, semestre, grupo y modalidad; el manifiesto es inmutable, versionado y sellado por contenido.
- El lector conserva todas las sesiones del alcance, incluidas las exteriores a la ventana y vinculaciones vacías. Consultas separadas dentro de una transacción de lectura evitan multiplicar sesiones por joins. No hay escrituras ni corte temporal ni primera página.
- Payload completo restringido a subunidad seleccionada. Particiones ajenas conocidas se excluyen; incertidumbres de origen y cruces de partición usan metadatos mínimos. Huérfanas seleccionadas/desconocidas generan diagnóstico explícito y evidencia en la huella sin atribuirles aula arbitrariamente. No se cambiaron estas consultas ni sus presupuestos.
- Huella SHA-256 determinista sobre manifiesto, catálogo y filas del alcance; contenido desconocido ajeno sólo aporta los metadatos mínimos ya definidos. No sustituye revisión de negocio, autorización ni control transaccional de concurrencia.

## Integración de esta unidad

`PrepararCoberturaAntesDeCambiosAsync` es común a diagnóstico y optimización:

1. Resuelve cada hoja confirmada contra una única aula por dimensiones y `AulaRef` exacta. Rechaza selecciones vacías, duplicadas, faltantes, ambiguas o con aulas adicionales. Los fallbacks históricos de homónimos/cohorte no autorizan un manifiesto distinto del seleccionado.
2. Resuelve una única territorial por la subunidad del contexto, mediante consulta de sólo lectura basada en el esquema del repositorio. No lee credenciales/configuración personal ni usa una territorial inventada.
3. Construye el manifiesto completo con el constructor existente. La ventana es la confirmada: rechaza fin posterior al **30/05/2027**, no extiende una ventana menor ni cambia fechas para hacerla viable. Estados no programables y relaciones/identidades inválidas bloquean explícitamente.
4. Sólo si hay aprobadas carga el snapshot canónico. Fallo técnico y cancelación mantienen categoría propia (`TipoPreparacionPlan`) y el código del lector; no se convierten en colección vacía ni en déficit. Un snapshot de otro manifiesto se rechaza por comparación exacta de huella.
5. Valida y muestra todas las aulas/relaciones afectadas en el registro; establece estado terminal antes del modal. Informa expresamente que falta T4 y que no autoriza conservación/replanificación. No se crea corrida para este bloqueo.
6. El gate precede al flush pendiente de dimensiones y al escalado mutable. Para Disponible mantiene el escalado/generación heredados, sin exigir lectura canónica. Si se sincronizan dimensiones se recarga y prepara nuevamente; también se prepara el contexto vigente antes de aplicar escalado y el recargado después de aplicarlo. No se pierde el manifiesto antes de crear corrida.

El pendiente de dimensiones debe pertenecer a un aula Disponible del alcance; no se usa ese flush para cambiar una aprobada o un aula no seleccionada. La protección no es una transacción de autorización: cambios concurrentes entre lectura y escritura siguen siendo riesgo abierto.

## Validación y semántica de horas

`AgValidadorPlanAprobado.Validar` no muta el snapshot ni el contexto:

- Todas las relaciones requeridas aparecen, incluso sin vinculación/sesiones. Cero sesiones no constituye un plan completo.
- Conserva diagnósticos del lector, rechaza claves desconocidas/extras/ambiguas, vinculaciones múltiples y sesiones canónicas duplicadas. No suma duplicados como cobertura.
- Comprueba IDs, vínculo sesión/relación, subunidad, docente válido/no centinela, fecha/intervalo interpretables, horas finitas positivas y duración ocupada positiva dentro de un día. Las horas lectivas no pueden exceder el intervalo ocupado; no se inventa conversión de hora académica. No certifica aquí patrones semanales ni disponibilidad contractual docente.
- Usa la ventana **del manifiesto** y los festivos capturados del contexto, no una ventana ampliada ni el reloj actual. Fechas fuera de plazo/festivos no cuentan como horas válidas.
- Compara intervalos reales por fecha para aula y docente entre sesiones reconciliadas del snapshot. Extremos contiguos no se solapan; una sesión no se cruza consigo misma. Sesiones conflictivas no suman cobertura válida.
- Resultado tipado por aula y relación: esperado, horas válidas sin ratificar, preservadas, déficit, exceso, horas fuera de ventana y diagnósticos. Déficit/exceso se calculan por relación, sin compensarlos entre asignaturas. **HorasPreservadas es siempre cero** hasta T4; horas válidas son evidencia diagnóstica, no conservación autorizada. Un déficit con origen reconciliado sigue esperando ratificación, no habilita automáticamente propuesta ni generación.

**Límite de conflictos:** sólo se comparan sesiones expuestas/reconciliadas por el snapshot. Otras cohortes productivas no seleccionadas no vienen con identidad exacta en este DTO; `OcupacionExterna` agrupa franjas sin ID canónico y contiene las propias sesiones productivas, por lo que no permite certificar cruces externos exactos evitando autocruces. T5 debe integrar ocupaciones inmutables con identidad y calendario completo; no se afirma resuelto ese control aquí.

## Reproducibilidad y verificaciones pendientes

Cuando puede crearse corrida (actualmente sólo alcance sin aprobadas), `ParametrosJson` incorpora `manifiestoCobertura`, `planCanonico` y `validacionPlan`, además del snapshot previo, festivos, configuración e intento. `planCanonico` es null explícito si no hubo lectura requerida. Las aprobadas bloqueadas conservan evidencia en contexto/registro, **no auditoría durable de ratificación ni snapshot persistido en una corrida ficticia**. T4 sigue siendo propietario de esa persistencia.

Fallo de captura de huella, serialización inicial o incorporación serializada del plan fechado falla claramente; se retiró el JSON de fallback que permitía continuar sin snapshot completo. No se probó la serialización/reconstrucción de ida y vuelta ni la persistencia real. El consumidor histórico existente no se migró para reconstruir estos nuevos contratos.

### Única corrección estructural autorizada de esta unidad

La revisión independiente detectó que el catch exterior reemplazaba `SNAPSHOT_SERIALIZACION_ERROR` por `FALLO_EJECUCION` antes de emitir FIN. `CerrarEstadoVisualOptimizacion` ahora retorna si ya existe un terminal: conserva fase, motivo y presentación específicos, también frente a cancelación tardía o fallo de presentación tras completar. Sin terminal previo mantiene los respaldos Error/`FALLO_EJECUCION` y Cancelado/`CANCELACION`. El inicio existente limpia terminal, fase y motivo antes del primer await, por lo que no se conserva el terminal de un intento anterior. No se modificó el cierre persistente de corrida ni su protección `corridaFinalizada`.

Se añadieron dos regresiones NO ejecutadas: causa de snapshot/terminal completado/cancelado frente al helper genérico y contrato de fuente para los catch, FIN y reinicio de intento. La regresión genérica anterior ahora inicia una fase activa antes de cada cierre y comprueba sus códigos de error/cancelación. Esta es la única corrección consumida; no certifica el recorrido completo de serialización, logging ni persistencia y no amplía T2 ni resuelve I5. La comprobación sigue limitada a lectura/diff y `git diff --check`, sin pruebas, build, SQL, aplicación ni revisión nativa.

Se añadieron dos includes puntuales al proyecto XLL no-SDK para modelo y validador; el proyecto de pruebas SDK recoge los fixtures nuevos sin ampliar includes globales.

Pruebas escritas, **no ejecutadas**: validador público (ratificación, vacío/ausencia, límites, duración, festivos, duplicados, identidad, docente, déficit/exceso, cruces de aula/docente, contigüidad, incertidumbres/extras y regresión Disponible); contratos de fuente UI para ambas rutas, gate previo a escrituras y captura sin fallback; composición pública con lector inyectado sin abrir conexión. Se mantienen las pruebas anteriores de manifiesto/lector y preservación de histórico. No hay evidencia RED/GREEN, compilabilidad, regresión ejecutada ni aprobación funcional. La prohibición de S8 es la excepción explícita a ejecutar test-first.

Comprobación autorizada: lectura estructural y `git diff --check`; este último sólo examina archivos rastreados, no valida compilación ni prueba archivos nuevos. No se crearon scratch/cachés.

Riesgos abiertos (ítems 1–4): correspondencia histórica/productiva sin `idaula`, cobertura y ratificación pendientes; permisos y auditoría T4; contratos/serialización compartidos y protección Disponible no ejecutados; cambios concurrentes entre carga y escritura/promoción. Continúan pendientes simulación T3, autorización/ratificación durable T4, conservación/ocupaciones/cobertura integral T5 y reemplazo/exportación T6. Se requiere verificación independiente del padre; no hay entrega completa ni autorización de release.
