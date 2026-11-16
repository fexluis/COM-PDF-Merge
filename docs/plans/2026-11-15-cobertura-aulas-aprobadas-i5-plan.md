# Cobertura de aulas aprobadas e I5 — plan de implementación

Fecha: 2026-11-15. Estado: PROPUESTO, ajustado al flujo oferta → confirmación → reevaluación → bloqueo si corresponde. No aplicado al código. No compilar ni ejecutar pruebas sin aprobación.

## Evidencia y objetivo

- Corrida #95: I5_AULA_APROBADA_CON_DEFICIT, CALI I-B PRESENCIAL, aula 14, estado Aprobado; rechazo del generador en 31 ms.
- Aula 14: 6 relaciones, 256 horas, 64 sesiones de 4 h. Hasta 30/05/2027, con festivos del snapshot, bandas martes-miércoles: 30 fechas frente a 32 sesiones; jueves-viernes: 31 frente a 32. Déficit: 3 sesiones. Aula 13 (I-A) tiene igualmente 256 horas.
- #96: ventana manual hasta 30/06/2027; finalizó en 143 segundos. Staging: 2064 sesiones, fecha máxima 30/05/2027; sin sesiones de aulas 13 y 14. Libro9: 136 filas, fecha máxima 20/05/2027; CALI II-A a VII-B, sin I-A ni I-B.
- Ausencia en staging/Excel NO demuestra ausencia en producción. Investigar sesiones aprobadas canónicas antes de decidir conservación o replanificación.
- Objetivo: toda relación requerida del alcance seleccionado queda cubierta por sesiones aprobadas conservadas o nuevas, dentro de 05/02/2027–30/05/2027. Ampliar calendario no constituye aceptación.

## Flujo obligatorio antes del bloqueo definitivo

1. Cargar alcance y plan aprobado real. Si está completo dentro de la ventana, conservarlo e incluirlo en cobertura.
2. Si existe déficit verificable, evaluar el siguiente escalón incluso para aulas Aprobado, sin modificarlas y sin exigir autorización previa para simular. No devolver I5 todavía.
3. Ofrecer una propuesta concreta por aula: ventana actual/siguiente, días añadidos, sesiones fechadas, déficit antes/después, patrones que requieren redistribución y fecha final fija 30/05/2027.
4. La confirmación autoriza ese único paso y, para Aprobado, su replanificación candidata. Redistribución LEGADO se consiente explícitamente por asignatura dentro del mismo formulario. No cambiar estado/productivo todavía.
5. Verificar huellas y permisos, guardar de forma atómica la configuración confirmada y su autorización, reconstruir contexto/plan y reevaluar cobertura. Una confirmación repetida no agrega otro paso.
6. Cobertura completa: continuar generación, validación y promoción. Cobertura insuficiente: I5 terminal con detalle; si hay otro escalón, informar que se ofrecerá en la siguiente corrida. No encadenarlo dentro de la corrida actual.
7. Revisión rechazada: I5 terminal para el aula aprobada, sin cambios. Horario aprobado no verificable tras lectura exitosa: I5 terminal con motivo PLAN_NO_VERIFICABLE; puede ofrecer revisión de datos, pero no simular autorización sobre un origen desconocido.
8. Error técnico de lectura, huella obsoleta o cancelación conserva su código específico; no atribuirlo a déficit. No silenciar el aula ni ampliar fechas.

| Actual | Siguiente | Límite |
|---|---|---|
| LUN-JUE | LUN-VIE; añadir viernes | LUN-VIE terminal |
| MAR-VIE | LUN-VIE; añadir lunes | LUN-VIE terminal |
| VIE-DOM | JUE-DOM; añadir jueves | Primer paso |
| JUE-DOM | MIE-DOM; añadir miércoles | Segundo paso desde VIE-DOM; MIE-DOM terminal |
| VIE-SAB | JUE-SAB | Primer paso histórico |
| JUE-SAB | ISO 4,5,6,7 | Segundo paso histórico; terminal |

Un paso por aula y corrida/confirmación, máximo dos acumulados desde VIE-DOM y desde VIE-SAB. No confundir ISO 4,5,6,7 histórico con el estado canónico JUE-DOM: conservar origen de escalera, paso, jornadas y semántica de ranuras (incluido sábado M/T). La autorización protege la escritura; no impide calcular ni presentar propuestas.

## Fase 1. Unificar alcance y demanda esperada

Causa: AgGeneradorHorarios.ValidarAprobadas evalúa aulas Aprobado; AgFabricaIndividuos.CrearAleatorio, línea 73, sólo genera para Disponible. La compuerta compara genes existentes contra staging; no detecta relaciones omitidas sin genes.

Acción:
1. Crear AgManifiestoCobertura e ítems por clave estable: territorial, período, sede, programa, aula, asignatura, modalidad/cohorte. No usar sólo asignatura o semestre como identidad.
2. Construir manifiesto desde selección confirmada + relaciones aula-asignatura requeridas. No derivarlo de genes, staging o filas exportadas. Usar AgEvaluadorDemandaVentana.HorasObjetivo para demanda efectiva y conservar demanda original para huella.
3. Durante preparación cada ítem puede quedar PendienteRevision; antes de ejecutar declara destino definitivo: ConservarAprobada o Replanificar. Aulas no programables seleccionadas producen bloqueo explícito, nunca exclusión silenciosa. Carga global de otras cohortes aporta ocupaciones; no agrega demanda seleccionada.
4. Guardar manifiesto, alcance, ventana, decisiones y versión de esquema en ParametrosJson. Sin fallback a la primera aula si faltan aulas seleccionadas. Desactivar fallback de fábrica a aula nula cuando existe un manifiesto explícito.

Archivos: AgContexto.cs; construcción de contexto/snapshot en AgFormGeneradorHorarios.cs; DTO nuevo de manifiesto; AgFabricaIndividuos.cs.

Aceptación: CALI I-A e I-B figuran como conservadas o replanificadas, cada una con 6 asignaturas y 256 h; ninguna desaparece al crear genes.

## Fase 2. Recuperar y verificar sesiones aprobadas reales

Causa: ValidarAprobadas simula de nuevo la demanda de aulas aprobadas, sin validar el horario aprobado existente. La interfaz IAgScheduleRepository ofrece lecturas de staging, insuficientes para certificar producción.

Acción:
1. Crear IAgPlanAprobadoRepository en Domain y repositorio de Infrastructure. Consultar vinculaciones/sesiones canónicas aprobadas mediante su vínculo/auditoría de promoción. Verificar esquema y estados reales antes de escribir SQL; no presumir que estado_ag del staging acredita aprobación productiva.
2. Consultar por identidad completa del manifiesto. No limitar consulta a mayo: también recuperar sesiones fuera de plazo y reportarlas. Resolver paginación completa, duplicados, huérfanos y relaciones sin correspondencia única.
3. Crear DTO inmutable de sesión conservada: identidad de relación, IDs canónicos, fecha, hora inicio/fin, horas lectivas, docente, estado y revisión de origen.
4. Validar por relación cobertura horaria exacta, fechas dentro de ventana, duración, festivos y solapamientos. No sumar dos veces misma sesión por joins.
5. Lectura exitosa sin sesiones verificables: I5 con motivo PLAN_NO_VERIFICABLE, sin omitir el aula. Horas insuficientes o fuera del plazo: devolver diagnóstico tipado RequierePropuesta con detalle por asignatura, esperado, conservado, faltante, fechas y origen; ofrecer escalón antes de I5. Lectura fallida: error técnico específico, nunca colección vacía interpretada como éxito.
6. Si existe plan aprobado completo dentro del plazo, conservarlo aunque una simulación desde cero encuentre déficit. No inventar ni modificar fechas de sesiones aprobadas.

Contrato NUEVO propuesto:

```csharp
public interface IAgPlanAprobadoRepository
{
    Task<Result<IReadOnlyCollection<AgSesionAprobadaDto>>> ObtenerAsync(
        AgAlcanceCobertura alcance, CancellationToken ct);
}
```

Dependencia: fase 1. Añadir documentación XML y registro de inyección; incluir archivos nuevos explícitamente en csproj si corresponde.

## Fase 3. Ofrecer escalado y confirmar antes de I5

Causa: EsAulaElegible exige Disponible y ValidarAprobadas devuelve I5 antes de una propuesta para Aprobado. Confunde permiso de escritura con permiso de simulación. Ampliar hasta junio hace pasar el cálculo previo sin crear sesiones para esas aulas.

Acción:
1. Reemplazar retorno anticipado de ValidarAprobadas por diagnóstico tipado que derive a propuesta de escalado para déficit verificable. Invocar esta preparación tanto desde diagnóstico como desde optimización; el generador recibe sólo contexto preparado y autorización vigente. Mostrar TODAS las aulas afectadas. Ofrecer el siguiente escalón sin revisión previa separada; confirmar la propuesta constituye autorización específica para replanificar. Conservar plan completo verificable y permitir cancelar.
2. Registrar autorización persistente y auditada: usuario, motivo, aula, versión/huella del plan aprobado, ventana, alcance exacto y fecha. No autorizar un aula distinta por compartir asignatura. Un cambio de configuración invalida la autorización.
3. Mantener horario aprobado anterior hasta promoción transaccional del reemplazo. No hacer UPDATE masivo a Disponible ni borrar sesiones para eludir I5.
4. AgEscaladorRanuraAula: separar PuedeEvaluarPropuesta (Disponible o Aprobado con identidad/configuración verificables) de PuedeAplicarPropuesta (permiso específico vigente para Aprobado y confirmación válida). Sustituir el filtro EsAulaElegible en todos los recorridos de simulación, incluida búsqueda de reclasificación y revalidación; mantener autorización estricta en cada escritura. No cambiar el helper global a permitir escrituras sobre cualquier Aprobado. Los demás estados quedan bloqueados.
5. AgFabricaIndividuos consume exclusivamente aulas marcadas Replanificar con permiso vigente; las conservadas no reciben genes mutables. Mutación/cruce/coherencia no pueden introducir aulas fuera de ese conjunto.
6. Aplicar tabla de escaleras del flujo obligatorio. Para CALI I-B ofrecer MAR-VIE → LUN-VIE con 30/05/2027 fijo antes de I5. Simular sin persistir también la redistribución LEGADO necesaria; diferenciar resultado sin/con redistribución. Consentimiento explícito por asignatura en el mismo formulario. Mostrar sólo pendientes y cambio real; lunes no está autorizado por ampliar la etiqueta si el patrón sigue limitado a 2,3 o 4,5. FIJO no se convierte por consentimiento genérico de LEGADO.
7. Evaluador y planificador deben acreditar sesiones fechadas completas, sin cruces con sesiones conservadas. Las 3 sesiones deficitarias de I-B son demanda de capacidad, no una garantía de disponibilidad docente.
8. Rechazo explícito → I5 con motivo REVISION_RECHAZADA. Tras único paso confirmado, cobertura insuficiente → I5 con motivo COBERTURA_INSUFICIENTE y déficit residual. Si no existe siguiente escalón, usar este último motivo e indicar terminal. No mostrar I5 previo a la oferta aplicable ni encadenar segundo paso en la misma corrida. Si el paso intermedio mejora capacidad pero sigue incompleto, permitir confirmarlo con advertencia explícita de que no habilita generación/promoción; persistir paso confirmado, detener por I5 y ofrecer el siguiente en otra corrida. No ampliar fecha final automáticamente.
9. Registrar origen de escalera, índice de paso e identificador de intento/propuesta. Guardar ventana, patrones consentidos y autorización en una sola transacción con revisión de origen. En reintentos recuperar la operación aplicada; no tratar cada recálculo como otra confirmación. La reevaluación posterior no vuelve a abrir el formulario durante esa corrida.

Dependencias: fases 1–2. Archivos: formulario, escalador, fábrica, autorización/repositorio nuevos, snapshot y huellas.

## Fase 4. Incorporar aprobadas como ocupaciones fijas

Causa: horarios conservados fuera del individuo pueden no intervenir en conflictos; incluirlos como genes ordinarios permitiría mutarlos o duplicarlos.

Acción:
1. AgContexto incorpora sesiones conservadas inmutables. Extender AgCalendarioConflictos para combinar sesiones nuevas y conservadas, identificadas por origen e ID canónico.
2. Comparar intervalos reales de fecha/hora para aula y docente. OcupacionExterna actual basada sólo en franja no sustituye ese calendario exacto. Mantener duración lectiva separada de intervalo ocupado.
3. Integrar en AgAsignadorTercero, reglas de aptitud y AgOperadorMutacion. Evitar autocruces al comparar una sesión consigo misma.
4. Al replanificar un aula autorizada, sus sesiones sustituidas no se cuentan como conservadas en la solución candidata. Seguir bloqueando otras ocupaciones y no retirar las sesiones anteriores de producción antes de confirmar reemplazo.
5. Cachear calendario fijo por huella y proyecciones por individuo/revisión. Invalidar al mutar docentes, fechas, bloques o aulas. No cachear sólo por IdAula.

Dependencias: fases 2–3.

## Fase 5. Cobertura integral, resultados y exportación

Causa: igualdad genes ↔ staging permite éxito aunque no se hayan generado genes para parte del alcance.

Acción:
1. Mantener AgCompuertaPromocion.ValidarCobertura para consistencia exacta de genes y staging nuevos.
2. Añadir AgValidadorCoberturaSeleccion: contrastar manifiesto esperado con unión disjunta de sesiones conservadas verificadas y nuevas materializadas. Validar cada relación; rechazar déficit, exceso, relación inesperada y duplicación. VACANTE sigue siendo falta de docente; no elimina demanda ni equivale a sesión faltante autorizada.
3. Añadir control de fecha final y conflictos al conjunto completo. HARD=0 del individuo no sustituye este control.
4. Si todo el alcance está aprobado y completo, permitir resultado Sólo conservación, sin población artificial ni exigencia de staging no vacío. Ajustar compuerta de orquestación para esta modalidad; no debilitar la validación del staging generado.
5. Guardar manifiesto y referencias/snapshot de conservadas vinculados a la corrida, para reproducción histórica sin depender de cambios posteriores en producción.
6. Resultados/Excel muestran sesiones nuevas y conservadas con origen visible, sin insertarlas dos veces en staging ni producción. Resumen incluye aulas seleccionadas, conservadas, replanificadas, bloqueadas, horas esperadas/cubiertas y vacantes. Corregir aviso UI de primera combinación si el alcance real es múltiple.
7. Promoción: no tocar conservadas. Sustituir sólo sesiones del alcance autorizado; conservar auditoría y vínculos de revisión. Revalidar horas, permisos, ocupaciones y huellas inmediatamente antes de confirmar, dentro de transacción y con control de versión.

Esquema lógico NUEVO; no es código existente:

```csharp
foreach (var item in manifiesto.Relaciones)
{
    var sesiones = solucion.SesionesUnicasDe(item.Clave);
    if (Math.Abs(sesiones.Sum(s => s.Horas) - item.HorasEsperadas) > 0.001)
        return Result.Failure($"COBERTURA_INCOMPLETA: {item.Clave}");
    if (sesiones.Any(s => s.Fecha < inicio || s.Fecha > limite))
        return Result.Failure($"SESION_FUERA_DE_VENTANA: {item.Clave}");
}
```

El servicio también debe comprobar IDs duplicados, relaciones inesperadas, origen, exactitud temporal y permisos; sumar horas por sí solo no certifica cobertura.

Dependencias: fases 1–4. Archivos: compuerta, materializador, promotor, resultados, consultas/exportador AG y DTO del resultado.

## Fase 6. Validar antes de modificar; conservar historial

Causa: AgFormGeneradorHorarios llama LimpiarStagingAsync(periodo, programa) antes del generador. #95 purgó 184 vinculaciones y 2064 sesiones de #94 y después abortó.

Acción:
1. Eliminar purga global del inicio de la optimización. Validar alcance y sesiones aprobadas, ofrecer escalado, recoger confirmación, guardar paso autorizado y reevaluar antes de crear una corrida ejecutable. Crear identidad de intento previa para limitar pasos aunque no se llegue a insertar ag_corrida.
2. Crear corrida nueva con identidad propia. Guardar genes/staging sólo bajo su IdCorrida. Mantener histórico anterior y limpiar únicamente resultados parciales de la nueva corrida cuando corresponda; no limpiar aprobadas ni referencias necesarias para reproducción.
3. Sustitución productiva sólo en promoción transaccional después de compuerta y verificación de versiones. La limpieza histórica será una acción separada, explícita, que respete referencias.
4. Si una corrida ya fue creada y falla, persistir estado terminal mediante UpdateCorridaAsync: Error para bloqueo/error y Cancelada para cancelación; guardar motivo y tipo estructurado en diagnóstico. No añadir valores enum sin migración/compatibilidad.
5. Registrar intentos bloqueados antes de crear corrida en auditoría/diagnóstico. No simular finalización correcta ni dejar EnProgreso.

Dependencias: fases 1–5.

## Fase 7. Corregir estado de pantalla y respuesta de Excel

Causa A: rama optResult.IsFailure (formulario, alrededor de 2504) muestra MessageBox y hace return sin cambiar lblStatusAG. finally restaura botones, no estado. Por eso I5 dejó Optimizando.

Acción A:
1. Estados UI explícitos: Preparando, Validando, RequiereRevision, Optimizando, Materializando, Completado, Bloqueado, Cancelado, Error. Transiciones centralizadas; separar estado de corrida persistida de estado visual.
2. Antes de confirmar: Estado: Propuesta de escalado pendiente de confirmación; mostrar actual → siguiente y déficit. I5 sólo después de revisión rechazada, cobertura todavía insuficiente o plan no verificable: Estado: Bloqueado — aula aprobada requiere revisión, con motivo y siguiente acción. Durante simulación indicar Evaluando escalado; tras aceptación indicar Reevaluando cobertura. No tratar bloqueo esperado como excepción inesperada.
3. Establecer estado terminal ANTES del modal y de cada return. finally garantiza estado terminal de respaldo, restaura controles y desactiva Cancelar; no debe sobrescribir un estado terminal válido.
4. Usar resultado de preparación tipado con código/motivo/aulas, no analizar mensajes traducidos para decidir flujo. Mientras se migra, los errores genéricos siguen teniendo salida terminal Error.
5. Progreso lleva identificador de ejecución y sólo actualiza la pantalla si esa ejecución continúa activa en fase Optimizando. Desactivar recepción antes de terminar o bloquear, para que callbacks pendientes no reviertan el estado.

Corrección inmediata del retorno actual (insertar antes de mostrar el modal; conservar traza y log existentes):

```csharp
if (optResult.IsFailure)
{
    lblStatusAG.Text = "Estado: Optimización detenida. Revise el diagnóstico.";
    lblStatusAG.ForeColor = Color.DarkOrange;
    progressBarAG.Value = 0;
    btnPromoverProduccion.Enabled = false;
    // Conservar aquí trazaOpt.Fin, traza.Fallo, traza.Fin y Log.
    MessageBox.Show(this, optResult.Error, "Optimización detenida",
        MessageBoxButtons.OK, MessageBoxIcon.Warning);
    return;
}
```

Este parche visual no resuelve cobertura ni clasifica errores. La solución final tipada diferencia I5 (Bloqueado: aula aprobada requiere revisión), cancelación y error técnico. Aplicar el control de callbacks indicado arriba.

Causa B: UI espera directamente al generador; ValidarFactibilidadAsync usa Task.FromResult y no desplaza el trabajo CPU. Inicialización/evolución ocupan el hilo UI.

Acción B:
1. Capturar parámetros, contexto aislado y CancellationToken antes de lanzar cálculo; no leer controles dentro del trabajador.
2. Ejecutar el generador fuera de UI mediante Task.Run, usando Progress<T> creado en el hilo UI. Verificar que el contexto y dependencias del generador no usan Excel/COM ni controles. Sustituir traza que toque pantalla por adaptador seguro.
3. Propagar CancellationToken a evaluación de aprobadas, planificadores, población, validación de conflictos y bucles internos largos. No basta comprobarlo entre generaciones.
4. Rechazar ejecuciones simultáneas; no permitir editar selección durante cálculo. Esperar que termine antes de liberar contexto/dependencias. No usar DoEvents, .Wait() o .Result para resolver la falta de respuesta.

Patrón de llamada propuesto (sustituye llamada directa; contexto aislado se debe implementar):

```csharp
var token = _ctsEjecucion.Token;
var contextoTrabajo = contextoAislado; // Sin controles ni objetos COM.
var optResult = await Task.Run(() =>
    _generadorHorarios.EjecutarOptimizacionAsync(
        contextoTrabajo, corrida.PoblacionTamano, corrida.GeneracionesMax,
        corrida.TasaCruce, corrida.TasaMutacion, corrida.Elitismo,
        progressReporter, token), token);
```

Todo fallo esperado debe dejar estado terminal aunque no lance excepción. Persistir el resultado terminal sin reutilizar un token ya cancelado; emplear token de cierre acotado y registrar fallo de persistencia.

## Fase 8. Huellas y diagnóstico

Causa: conservar aprobadas introduce datos externos al plan generado que pueden cambiar entre evaluación y guardado.

Acción:
1. AgHuellaPlanVentana incluye manifiesto, política por aula, autorizaciones/revisiones y sesiones conservadas: IDs, relación, docente, fecha, intervalo, horas y estado. Orden determinista y representación invariante. Mantener horas originales y efectivas con semánticas distintas.
2. Recalcular/releer revisiones en confirmación y promoción; usar comparación exacta de huella completa, nunca StartsWith. Cambio concurrente aborta sin modificar producción. Operación ya aplicada se verifica con su identidad y efecto persistido para idempotencia.
3. DEBUG por ejecución/aula: INICIO/FIN, fase, duración, demanda, sesiones conservadas/nuevas, déficit, ventana, origen de cobertura y motivo de bloqueo. Registrar por qué cada aula fue conservada o replanificada.
4. Medir carga de aprobadas, validación, primera población, aptitud por generación, materialización y exportación. Incluir eventos terminales de fallo/cancelación y estado de UI; no volcar datos personales innecesarios.

Dependencias: fases 1–7.

## Fase 9. Verificación antes de release

Pruebas unitarias en memoria; integración separada con SQLite temporal y exportación. Ejecutar sólo tras aprobación.

1. Aprobada con plan canónico completo dentro de mayo: conserva fechas/docentes; no genera genes nuevos; cobertura incluye aula.
2. Aprobada sin plan verificable tras lectura exitosa: I5/PLAN_NO_VERIFICABLE antes de purgar/crear staging, sin éxito silencioso. Error de lectura conserva código técnico y no se etiqueta I5.
3. Aprobada con déficit real y escalón disponible: comprobar que se muestra propuesta ANTES de cualquier I5. Aceptación con cobertura completa continúa; rechazo o cobertura residual produce I5 con motivo y estado UI Bloqueado, botones restaurados. Validar retorno Failure, excepción y cancelación con sus códigos propios.
4. Alcance mixto: disponibles + aprobadas. Eliminar una relación completa y comprobar que falla cobertura aunque genes y staging coincidan entre sí.
5. Todo aprobado: éxito de sólo conservación, cero staging nuevo y sin duplicación de producción.
6. CALI I-A/I-B: 6 relaciones y 256 h por aula. Replanificación autorizada + LEGADO autorizado a AUTO, lunes disponibles, 64 sesiones/aula hasta 30/05/2027. Incluir festivos y conflicto docente que impida una solución para verificar rechazo honesto.
7. Rechazar revisión o cambio de patrón mantiene aprobación y sesiones anteriores, termina con estado correcto.
8. Nueva corrida fallida no borra #94/#96 ni sus sesiones. Promoción parcial fallida revierte transacción.
9. Cambio de sesión aprobada/autorización/fecha después de propuesta invalida huella; doble confirmación no duplica sesiones.
10. Sesión conservada cruza docente/aula de nueva: detectado por intervalos exactos. Replanificación autorizada no se autocruza con sesiones que reemplaza.
11. Vista/Excel incluyen CALI I-A/I-B con origen; reconciliar cada relación contra manifiesto y total de horas, no sólo max(fecha).
12. Prueba de respuesta UI: trabajo deliberadamente lento, progreso visible, cancelar responde, callback tardío no cambia Bloqueado/Cancelado; sin acceso a controles desde trabajador.
13. Regresión en Disponible y Aprobado: LUN-JUE → LUN-VIE añade viernes; MAR-VIE → LUN-VIE añade lunes; terminales sin nueva oferta si cabe. Desde VIE-DOM: primera corrida ofrece sólo JUE-DOM; si queda déficit, I5 tras confirmación; siguiente corrida ofrece sólo MIE-DOM, terminal. MIE-DOM insuficiente no inventa otro escalón. VIE-SAB → JUE-SAB → ISO 4,5,6,7 conserva dos pasos históricos. Reintento/doble clic no consume paso extra; cancelar no modifica configuración ni permisos.
14. Huella modificada tras oferta rechaza aplicación sin persistencia parcial. Patrones LEGADO no consentidos permanecen intactos; sólo consentir días no habilita lunes fuera del patrón. Verificar que cada filtro de evaluación permite propuesta de aprobadas y cada escritura exige su autorización.
15. Tras autorización: pruebas focalizadas, regresión requerida y después build Release/smoke. No certificar cierre con compilación únicamente.

## Orden y cierre

Orden: 1 → 2 → 3 → 4 → 5 → 6 → 7 → 8 → 9. Añadir pruebas con cada cambio y mantener compilabilidad de contratos/dependencias por paso. Entregar lista real de archivos modificados y evidencia por fase.

Cierre: propuesta del siguiente escalón ofrecida antes de I5 para déficit verificable; un paso por confirmación/corrida y terminales respetados; I5 reservado a revisión rechazada, cobertura insuficiente tras el flujo o plan aprobado no verificable; ninguna relación seleccionada omitida; aprobadas preservadas o reemplazadas con autorización auditada; cobertura exacta y conflictos validados; límite 30/05/2027 cumplido para todo el conjunto; UI nunca queda Optimizando tras rechazo; histórico anterior intacto; huellas verificadas al guardar/promover. No release hasta cumplirlo.
