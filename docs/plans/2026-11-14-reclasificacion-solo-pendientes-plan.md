# Reclasificación solo de materias pendientes dentro del plazo

Estado: plan de corrección para otro modelo. Sin implementación ni ejecución en esta revisión.
Base inspeccionada: HEAD 3c2d1b4; volver a comprobar estado del repositorio antes de editar.
T1–T7 son identificadores locales de tareas; no reservan numeración global de rondas/fases.

## Objetivo e invariantes

Evaluar primero el plan fechado dentro del plazo aprobado. Ofrecer reclasificación exclusivamente de materias LEGADO con sesiones pendientes en ese plan. Materias completas no aparecen, aunque su aula tenga otras materias pendientes. FIJO se conserva; AUTOMATICO ya es redistribuible.

- Plazo del escenario: 05/02/2027–30/05/2027. No extenderlo para hacer factible la propuesta.
- Respetar las cohortes seleccionadas y sus IDs de programa, sede, semestre, grupo y modalidad. No fusionar programas por nombres homónimos.
- Usar la misma elegibilidad en evaluación, lista, autorización y guardado.
- Conservar huellas, guardas SQL, transacción, rollback e idempotencia.
- No compilar, ejecutar pruebas/AG, modificar datos operativos, hacer commit o push durante esta implementación sin autorización posterior específica.

## Diagnóstico verificado

1. AgFormGeneradorHorarios, alrededor de líneas 1634–1659, obtiene candidatos directamente desde IdentificarPatronesRedistribuibles(contexto). La lista se construye sin usar los pendientes de la evaluación que se ejecuta después.
2. IdentificarPatronesRedistribuibles(contexto, aula) filtra por LEGADO y patrón no vacío. No comprueba demanda efectiva, sesiones pendientes ni estado del aula.
3. EvaluarInterno valida autorizaciones exigiendo Estado=Disponible. La lista permite aulas que luego ese mismo flujo rechaza.
4. Caso concreto leído en la base temporal operativa: aula 12, referencia 3|1|2|X|A|MIXTA, estado 4=Aprobado. Sus materias aparecen en la captura. Si se selecciona cualquiera, no encontrará aula Disponible y producirá el rechazo. No cambiar su estado para evitar la validación.
5. Log 14/11/2026: versión 1.6.4.13, DEBUG apagado; fallos a las 17:55:09 en pre-vuelo y 17:55:16 en optimización, con «Reclasificación ajena o configuración modificada». Falta identificación del registro rechazado; no se demuestra cuál fue seleccionado ni un cambio concurrente real.
6. La excepción esperable de validación se propaga como ExcepcionNoControlada. Impide completar el flujo antes de iniciar la optimización.

## T1 — Unificar elegibilidad y fijar casos de regresión

Archivos: AgEscaladorRanuraAula.cs; suites AgRedistribucionCorreccionesTests y AgRedistribucionRanurasAceptacionTests.

- Centralizar el predicado de elegibilidad: aula seleccionada, referencia válida y Estado=Disponible.
- Aplicarlo a evaluación, generación de candidatos y validación de selección. Verificarlo otra vez al guardar con contexto vigente.
- No ofrecer Aprobado, Inactiva, Mantenimiento o Reservada. Mantener sus registros intactos.
- Escribir primero pruebas: aula elegible; aula 12 Aprobado; dos aulas con misma asignatura y distinto estado; dos programas homónimos seleccionados.

Aceptación: todo candidato generado cumple la política que posteriormente lo valida. Los cambios posteriores de estado se rechazan de forma controlada.

## T2 — Generar candidatos desde el resultado fechado

Depende de T1. Archivos: AgEvaluadorDemandaVentana.cs; AgEscaladorRanuraAula.cs; AgReclasificacionPatron.cs si necesita métricas adicionales.

- Evaluar demanda real con bloques, duración y festivos, dentro del plazo aprobado.
- Reutilizar IdAsignaturasSaturadas y las sesiones del resultado. Calcular pendientes por materia, no mediante capacidad nominal ni fechas de corridas antiguas.
- Candidata = aula elegible + relación incluida en la demanda efectiva + LEGADO con días restringidos + sesiones pendientes positivas.
- Si CabeEnPlazo=true, devolver cero candidatas.
- Si existe error estructural, devolver diagnóstico; no abrir un formulario con todos los LEGADO como alternativa.
- Si BusquedaAgotada=true, indicar «sesiones no ubicadas en búsqueda acotada»; no declarar imposibilidad demostrada ni ampliar automáticamente el plazo.
- Mantener la búsqueda de alternativas para AUTOMATICO y la escalera de ventanas aunque no existan candidatas LEGADO.
- No incluir materias completas como posibles liberadoras de capacidad. Si las pendientes seleccionadas no bastan, informar el déficit restante.
- Eliminar o restringir la sobrecarga pública que enumera candidatos sin evaluación para impedir que la UI vuelva a usarla.

Lógica orientativa, adaptada al contrato definitivo:

```csharp
if (!string.IsNullOrEmpty(evaluacion.Error) || evaluacion.CabeEnPlazo)
{
    return new List<AgReclasificacionPatron>();
}
var pendientes = new HashSet<int>(evaluacion.IdAsignaturasSaturadas);
// Intersectar con relaciones de demanda efectiva y LEGADO; nunca devolver todas las relaciones del aula.
```

Aceptación: dentro de un aula mixta, solo aparecen las materias con déficit del plan evaluado. Cada fila informa sesiones/horas pendientes, días actuales y fecha límite.

## T3 — Cambiar orden y contenido de la UI

Depende de T2. Archivos: AgFormGeneradorHorarios.cs; FormAutorizarReclasificacion.cs.

Flujo obligatorio:

```text
Capturar contexto y plazo
  → evaluar plan fechado
  → obtener únicamente candidatas pendientes de ese resultado
  → si existen, mostrar selección individual desmarcada
  → simular selección en copia y evaluar misma ventana
  → evaluar siguiente escalón solo si siguen pendientes y aporta mejora
  → mostrar propuesta definitiva
  → confirmar y comprobar huellas
  → guardar ese mismo lote y recargar
```

- Usar un resultado de evaluación compartido; no construir una lista independiente antes de evaluar.
- Sin candidatas: no mostrar FormAutorizarReclasificacion. Continuar solo con escalados válidos si los hay; de lo contrario, SinCambios o diagnóstico de restricciones.
- Cancelar/cerrar detiene el flujo sin escribir.
- Selección vacía no reclasifica. No impedir una propuesta nominal de escalado válida.
- Etiquetas legibles: sede, programa identificable, semestre/grupo, materia, patrón y déficit. Mantener la clave técnica en el objeto, no como único identificador visible.
- Ajustar texto cuando hay reclasificación sin ampliar días; no afirmar que todas las aulas tienen sobrecarga global.

Aceptación: ningún formulario para aulas cuyo plan está completo; ninguna materia completa en la lista de un aula saturada.

## T4 — Convertir rechazos esperados en resultados controlados

Depende de T1–T3. Archivos: AgEscaladorRanuraAula.cs; UI; AgDiagnosticoTraza si corresponde.

- Conservar validación completa: IdAula, AulaRef, Subunidad, IdAsignatura, origen anterior, días anteriores y pertenencia a candidatas evaluadas.
- Validar selección mediante Result o resultado tipado. No usar ArgumentException como salida normal de una selección desactualizada.
- Diferenciar motivos: aula no elegible, identidad ajena, relación ausente, patrón cambiado, duplicado, selección fuera del diagnóstico o huella cambiada.
- UI muestra causa concreta y permite recalcular; no convertir un rechazo en éxito ni descartar silenciosamente una autorización inválida.
- Mantener comprobación de cambios entre evaluación, selección y guardado. No sustituir el contexto sellado por uno nuevo sin volver a presentar la propuesta.
- Registrar operación/traza, IDs de aula/asignatura, estado y motivo en WARN aun con DEBUG apagado. Resumen INFO de aulas evaluadas, materias completas, pendientes y candidatas. Detalle de búsqueda sigue en DEBUG.

Aceptación: reproducir selección de aula 12 mediante llamada inválida devuelve fallo controlado, cero escrituras y motivo «aula no elegible»; no ExcepcionNoControlada.

## T5 — Preservar escalera, planes y persistencia

Depende de T2–T4. Archivos: AgEscaladorRanuraAula.cs; AgHuellaPlanVentana.cs; AgHuellaEscaladoPersistido.cs; fábrica y materializador solo si cambia el contrato.

- LUN-JUE/MAR-VIE → LUN-VIE. VIE-DOM → JUE-DOM → MIE-DOM. Un paso por propuesta; MIE-DOM terminal; conservar tope de VIE-SAB.
- Proponer ampliación solo si reduce pendientes conservando cobertura previa. Más ranuras nominales no prueban mejora.
- Reclasificación puede resolver el problema dentro de la misma ventana, incluso siendo terminal.
- Actualizar exclusivamente relaciones autorizadas; no limpiar LEGADO/FIJO no seleccionados ni ampliar fechas.
- Huellas deben cubrir identidad, clasificación, patrón, plazo y alcance. Validar antes de filtrar acciones y antes de escribir.
- Segunda aplicación sin nuevas escrituras. Recalcular contexto/caché después del commit.
- FechasGen y sesiones materializadas deben reflejar el plan confirmado; un plan de fechas factible no acredita por sí solo disponibilidad docente global.

## T6 — Completar pruebas sin ejecutarlas

Depende de T1–T5.

1. Todas las materias caben: cero candidatas, sin diálogo.
2. Un aula pendiente y otra completa: solo pendientes de la primera.
3. Una materia completa y otra pendiente en el mismo aula: excluir la completa.
4. LEGADO pendiente frente a FIJO pendiente: solo LEGADO es seleccionable; FIJO se informa como restricción.
5. AUTOMATICO pendiente: evaluar escalera sin diálogo de reclasificación.
6. Aula Aprobado/Reservada/Mantenimiento/Inactiva: nunca candidata.
7. Misma materia en dos aulas/programas: autorizar una no afecta a la otra.
8. Cambio de estado, patrón, plazo o alcance durante el diálogo: rechazo controlado y rollback completo.
9. Selección ajena a candidatas, duplicada o manipulada: rechazo sin escrituras.
10. Cancelar, cerrar y aceptar sin seleccionar: comportamiento previsto, sin reclasificación.
11. Reclasificación resuelve en LUN-VIE/MIE-DOM: misma ventana, cero pendientes.
12. Escalera con capacidad real 4→5 y 4→5→6, persistencia e idempotencia.
13. Fechas y horas del plan confirmado iguales en staging; sin fechas posteriores al 30/05/2027 en el escenario factible.
14. Búsqueda agotada/error estructural: diagnóstico preciso; no lista indiscriminada.

Actualizar pruebas anteriores que exigían candidatos de LEGADO completos: ese comportamiento contradice la petición actual. Conservar las pruebas de aislamiento y huellas.

## T7 — Entrega y validación posterior

- Registrar archivos modificados, causas y pruebas añadidas; registrar nuevos archivos en csproj si el proyecto XLL los requiere.
- Revisión estática de contratos, llamadas, filtros y diferencias. No presentar esto como compilación aprobada.
- Tras autorización independiente: verificaciones Fast focalizadas, regresión requerida y smoke test con DEBUG desde pre-vuelo.
- En captura real comprobar ausencia de aula 12 y materias completas en el formulario, diagnóstico de las pendientes, confirmación exacta y ausencia de excepciones esperadas no controladas.
- No recomendar release hasta validar compilación, pruebas y resultado funcional. No publicar automáticamente por existir un workflow local de release.
