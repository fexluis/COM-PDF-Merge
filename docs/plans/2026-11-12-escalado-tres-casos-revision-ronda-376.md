> Estado vigente 2026-11-14: el código descrito previamente había desaparecido del directorio de trabajo. Se reconstruyeron los tres archivos de ranuras y su integración. Las afirmaciones anteriores de implementación completa, especialmente fase 1791/persistencia, no describen el estado vigente. Ver [recuperación actual](2026-11-14-recuperacion-planificador-ranuras.md). Sin compilación/pruebas; sin nueva reserva en REGISTRY.

## Estado de implementación — 2026-11-13

Ronda 376 propuesta; fases 1788–1793. Sin reserva nueva en REGISTRY.

| Fase | Estado | Cambio |
|---|---|---|
| 1788 | Código editado; sin ejecutar | Evaluador, validador y fábrica usan AgPlanificadorOleadas; misma normalización de horas. |
| 1789 | Código editado; sin ejecutar | Sesiones con asignatura/fecha/ranura; pendientes derivados de sesiones; búsqueda acotada con estado agotado. |
| 1790 | Código editado; sin ejecutar | FechaFinProyectada nullable; horizonte diagnóstico de dos años; plazo original intacto; cancelación y límites del calendario. |
| 1791 | Código editado; sin ejecutar | Huellas por contenido de contexto y BD, comparación exacta, reaplicación antes de validar estado anterior, postcondiciones y commit del lote. |
| 1792 | Pruebas escritas; no ejecutadas | Tres rutas, fechas de fábrica, 4/8 h, calendario, caché, contenido modificado, cancelación, rollback y reaplicación incompleta. |
| 1793 | Pendiente de aprobación | Fast, Full Debug/x64, Full Release/x64 y smoke sobre copia de BD. |

La implementación conserva planes de referencia factibles. Si fallan, prueba calendarios de bandas alternativas, incluida banda única; consume ocurrencias concretas, sin reservar semanas completas vacías. La fábrica conserva esas mismas fechas. Este fallback puede cambiar la partición y debe revisarse con las regresiones de carriles; no declara optimalidad global.

La UI bloquea edición del perfil durante evaluación/aplicación, evalúa fuera del hilo visual, recarga contexto antes de escribir y después del commit. Rechaza instantáneas obsoletas. Las decisiones manuales sin sello ya no autorizan escritura.

Comprobación estática: git diff --check correcto; tres archivos nuevos incluidos exactamente una vez en el csproj.

No se compiló, ejecutó proyecto/pruebas, abrió Excel ni modificó BD operativa. Ninguna prueba nueva se declara aprobada. La viabilidad del Excel original al 30/05/2027 requiere su calendario y restricciones reales; no queda acreditada por capacidad nominal.

El diagnóstico y fragmentos siguientes documentan la revisión previa; el estado vigente es esta tabla.

---

# Plan revisado: cierre operativo del escalado automático

Ronda de revisión propuesta: 376. Fases propuestas: 1788–1793.
Estado previo: plan de revisión; implementación actual detallada arriba.
Fuente: adjunto del usuario «Plan para cumplir los casos obligatorios — Ronda 373 / Fases 1775–1780», contrastado con código actual.

## Numeración y estado real

REGISTRY consultado: ronda 375, fase 1787; siguiente 376/1788. R-074, R-075 y R-076 ya están closed. No reutilizar sus números ni reabrirlos. R-077 es solo candidato; confirmar registro al reservar.

Este documento usa una ronda de revisión y seis fases secuenciales. No declara seis waves equivalentes a una sola ronda. Si se convierte a plan swarm de seis waves, reservar seis rondas desde el contador vigente y actualizar el envelope completo. Ninguna reserva efectuada aquí.

El adjunto está obsoleto en numeración, HEAD, workspace limpio y porcentajes. No asumir HEAD f6e8e35 ni árbol limpio. Leer estado y evidencia reales antes de ejecutar. El Full histórico de R-073 no valida las revisiones posteriores.

## Acuerdos y correcciones al adjunto

Se conserva: rutas de escalera, etiqueta versus ISO terminal, un paso por corrida, máximo dos desde VIE-DOM y VIE-SAB, fechas originales inmutables, FIJO/LEGADO preservados, 4 h/8 h y builds serializados.

Ya implementado parcialmente: fecha nullable y recorrido acotado; exclusión de patrones AUTOMATICO en fábrica/validador/evaluador; planRes antes del bloqueo heurístico del validador; guard Desde; contadores Confirmadas/Modificadas; catch transaccional devuelve contadores vacíos. Verificar regresiones; no rehacer por los porcentajes del adjunto.

Pendiente crítico: el escalador sigue evaluando heurísticas, no planes efectivos. El nuevo contador de pendientes sigue saliendo de una asignación estimada a slots, sin sesiones fechadas. La huella compara solo prefijo ventana/jornada. Se comprueba antes de idempotencia, por lo que reaplicar con huella original falla al encontrar ventana destino. La verificación final revisa nodo, no todos los efectos autorizados.

No aceptar como falso positivo el problema completo porque el planificador aislado pase un ejemplo: el fallo puede residir en el escalador o su integración. No aplazar huella a otra etapa. No afirmar ausencia de cambios L0: si se añade servicio/contrato en proyecto no SDK, revisar csproj y sus reglas vigentes.

## Ejecución permitida

El usuario autorizó corregir e implementar pendientes. Compilar, ejecutar código/pruebas, reservar registro o publicar requieren autorización posterior. Gates y smoke requieren aprobación posterior del usuario. No crear subagentes por este documento.

## Fase 1788 — Una autoridad de planificación

Causa: AgEscaladorVentanaAula.Evaluar consume AgEvaluadorDemandaVentana; AgValidadorVentana consume también AgPlanificadorOleadas. Dos algoritmos pueden discrepar.

Archivos: AgEvaluadorDemandaVentana, AgPlanificadorOleadas, AgEscaladorVentanaAula, AgValidadorVentana, AgFabricaIndividuos y pruebas. Añadir servicio/contrato interno solo si evita duplicación; registrar archivos en csproj cuando corresponda.

Implementar evaluación compartida, pura, sobre instantánea de entrada: materias válidas, fallback de horas según catálogo, restricciones/origen, unidad, bloques, fechas y festivos. Reutilizar el constructor de planes efectivos en escalador, validador y fábrica. Mantener identidad día/franja/horas. La heurística puede ordenar candidatos, nunca certificar imposibilidad.

Contrato propuesto — miembros nuevos por implementar, no APIs existentes:
~~~csharp
var antes = evaluadorPlan.Evaluar(entrada, aula.DiasVentana);
if (antes.EsFactible) continue;

string paso = AgEscaleraVentanas.Siguiente(aula.DiasVentana);
if (string.IsNullOrEmpty(paso)) continue;

var despues = evaluadorPlan.Evaluar(entrada, paso);
if (!despues.MejoraRespectoA(antes)) continue;
~~~

Intentar redistribución AUTOMATICO en ventana actual antes de ampliarla. Si basta, no ofrecer escalado. FIJO/LEGADO conservan restricciones; ningún origen ambiguo se convierte silenciosamente. No resembrar patrones por módulo al regenerar aulas.

Criterio: estado actual factible por plan efectivo implica cero propuestas de ampliación. Probar escalador, no solo planificador, con 15 semanas y demandas [12,12,8,8,12,12], y fijo [2,3] en partición alternativa.

Dependencia: ninguna. Salida: misma entrada produce misma factibilidad en tres consumidores.

## Fase 1789 — Plan parcial verificable y pendientes reales

Causa: DemandaPendiente ahora contabiliza fracciones que caben en slots estimados; todavía no mide sesiones reales. AgPlanificadorOleadas puede fallar anticipadamente por capacidad global, impidiendo comparar mejoras parciales entre 4W,5W,6W.

Archivos: servicio/resultado compartidos, planificador y modal.

Añadir resultado de evaluación parcial sin cambiar inadvertidamente el contrato de optimización existente. Cuando demanda exceda capacidad global, devolver métricas parciales válidas o un resultado explícitamente indeterminado; no presentar un Failure genérico como cero sesiones pendientes. Si algoritmo greedy falla, búsqueda determinista acotada con estado INDETERMINADO separado de imposibilidad demostrada.

Cada sesión contabilizada debe tener asignatura, fecha, ranura y duración. Evitar duplicados/solapamientos; respetar calendario y reglas existentes. Inicializar todas las materias, incluidas las no planificadas.

Código de agregación una vez disponible el plan parcial real:
~~~csharp
var ubicadas = demandas.ToDictionary(d => d.IdAsignatura, _ => 0);
foreach (var sesion in sesionesProgramadas)
{
    ubicadas[sesion.IdAsignatura]++;
}
int pendientes = demandas.Sum(d =>
    Math.Max(0, d.SesionesRequeridas - ubicadas[d.IdAsignatura]));
~~~

Rechazar candidatos que pierdan materias ya planificadas o aumenten pendientes. Ordenar mejora por pendientes, materias infactibles y fecha final real. Añadir días sin mejorar estas métricas no basta. En VIE-DOM con demanda 6W permitir JUE-DOM parcial útil, sin saltar a MIE-DOM en la misma corrida.

Modal usa métricas posteriores, separa actuales/destino y dice «sin pendientes según el plan evaluado» solo con evaluación completa. Eliminar cualquier cálculo paralelo demanda-capacidad en advertencias. No anunciar necesariamente otro escalado si el siguiente es terminal o no útil.

Dependencia: 1788. Salida: pendientes iguales a demanda menos sesiones efectivamente ubicadas.

## Fase 1790 — Plazo original y proyección diagnóstica distintos

Causa del adjunto: propone buscar fuera del plazo, pero traduce null como «no cabe dentro del plazo». Agotar un horizonte no demuestra imposibilidad. Un único día con máxima carga tampoco acredita una fecha mínima global bajo festivos asimétricos.

Mantener fecha límite original para factibilidad y programación. Una proyección fuera del semestre es diagnóstico opcional, nunca ampliación de fechas. Nombre FechaFinProyectada; no FechaFinMinima salvo prueba de optimalidad. Derivar fechas del calendario completo de ranuras, no solo del día con más carga.

Contrato recomendado:
~~~csharp
public bool CabeEnPlazo { get; set; }
public DateTime? FechaFinProyectada { get; set; }
public bool HorizonteAgotado { get; set; }
~~~

null no equivale automáticamente a no cabe: mostrar «sin proyección disponible» o «horizonte de búsqueda agotado», según causa. CabeEnPlazo deriva del plan real dentro de ini..fin.

Elegir una sola política de horizonte acotado y cancelable; el adjunto mezcla fórmula dinámica y dos años. Usar configuración explícita, validar límites de DateTime y unidades; no inferir límite mediante semanas heurísticas. Calendario completamente festivo debe terminar. Evitar AddDays después de alcanzar el último día permitido.

Para expectativa 17/10/2027: fijar ini=05/02/2027, 36 sesiones solo domingo y exactamente 28/03/2027 como domingo excluido en el fixture; especificar cualquier otra exclusión. No atribuir esa fecha al Excel sin cargar el calendario real de la corrida.

Dependencia: 1789. Salida: fechas originales intactas y proyección honesta.

## Fase 1791 — Huellas completas, idempotencia y commit

Causas: huella incompleta y comparación StartsWith; verificación anterior al estado final rompe reaplicación; faltan comprobaciones finales de patrones. MarkDirty/log dentro del catch de transacción puede convertir un error posterior a commit en supuesto rollback.

Archivos: Decision, AgEscaladorVentanaAula y pruebas SQLite temporal.

Guardar instantánea canónica versionada: identidad aula/subunidad/perfil, fechas, jornada/unidad, demanda por materia, patrón/origen, bloques y calendario efectivo. Serializar ordenado y comparar hash exacto; sin prefijos. Componentes no persistidos deben contrastarse con su fuente vigente antes de aplicar. No confiar en agregados de demanda/capacidad.

La decisión contiene huella anterior y especificación/huella posterior esperada. Una reaplicación verifica estado posterior completo antes de comparar entrada anterior:
~~~csharp
if (EstadoFinalCompleto(conn, trans, decision))
{
    provisional.Confirmadas++;
    continue;
}
if (!string.Equals(decision.HuellaAntes,
    CalcularHuellaActual(conn, trans, decision), StringComparison.Ordinal))
{
    throw new InvalidOperationException("Configuración modificada.");
}
// Guard Desde; actualizar aula, nodo y cambios autorizados de patrones.
// Verificar postcondiciones completas antes de confirmar esta aula.
~~~

Estado final completo incluye componentes no modificados de entrada y efectos autorizados exactos; no basta aula en destino ni revisar solo nodo. Si aún no aplicado, faltan datos requeridos o hay huella vacía: fallo, sin omitir validación. No migrar FIJO/LEGADO automáticamente. Reaplicación correcta: confirmada sin nuevas escrituras.

Contadores provisionales solo se publican tras Commit. Excepción previa revierte lote completo y devuelve fallo explícito; excepción posterior a commit se informa como problema de notificación/persistencia secundaria, sin afirmar rollback. Separar resultado de aplicación de número de filas: éxito, confirmadas y modificadas. UI no muta memoria antes de ese éxito.

Recargar desde BD las relaciones/orígenes y claves afectadas; invalidar planes cacheados por contenido. No simular un reset de memoria que BD no hizo. Verificar existencia/unicidad del nodo del perfil cuando el contrato lo exige. Mantener aislamiento por subunidad y perfil; no prometer aislamiento de tablas compartidas entre perfiles.

Dependencia: 1790. Salida: decisión obsoleta rechazada; reaplicación completa exitosa; error en nodo/patrón revierte todo.

## Fase 1792 — Matriz de aceptación completa

Dependencia: 1791. Diseñar y escribir pruebas; ejecutarlas solo con aprobación. No sustituir demanda 5W por una materia fija en el día que falta.

Fixtures base: W semanas completas, catálogo completo, sin festivos; horas y demanda concretas compatibles con partición/cursores. Separar casos de semanas parciales, festivos, bloque ausente, restricciones docentes y unidades 8 h.

1. LUN-JUE: 4W→5W nominal, propone LUN-VIE, plan real de 5W sesiones distintas dentro del plazo, viernes usado, persistencia y segunda corrida sin propuesta.
2. MAR-VIE: mismo contrato, lunes usado.
3. VIE-DOM con 5W: propone JUE-DOM, jueves usado, cero pendientes si fixture factible, sin segunda propuesta.
4. VIE-DOM con 6W: primer paso JUE-DOM con pendientes explícitos; recargar de BD; siguiente MIE-DOM; miércoles usado; tercera evaluación vacía.
5. Fijo viernes/lunes/jueves: escalón útil solo si resuelve o mejora ese caso; restricción preservada. Son pruebas adicionales, no reemplazan las anteriores.
6. Tres materias fijas domingo: añadir jueves/miércoles no reduce pendientes; ninguna propuesta falsa. Proyección opcional separada del plazo.
7. AUTOMATICO con antiguo texto domingo: se redistribuye sin convertirlo en fijo; FIJO/LEGADO no cambian. Preparación legacy exige clasificación explícita cuando sea necesaria.
8. 4 h: rótulo MIE-DOM=3,4,5,6M,6T,7. 8 h: 3,4,5,6,7. Dos franjas sábado separadas; cero doble conteo.
9. Terminales LUN-VIE/MIE-DOM; VIE-SAB máximo dos pasos; ISO 4,5,6,7 terminal distinto de JUE-DOM.
10. Mismo número de materias/bloques pero contenido distinto invalida huella/cache. Cambio de horas manteniendo demanda total también invalida.
11. Reaplicación real con huella y postestado completo: éxito sin escrituras. Aula en destino con nodo/patrón incompleto: fallo. Error SQL tras actualizar aula: rollback y cero confirmaciones publicadas.
12. Todo festivo, horizonte agotado y cancelación: terminación acotada; ningún resultado indeterminado rotulado factible.
13. Mismo pre-vuelo desde diagnóstico y optimización; aceptar/rechazar, guardado/recarga fallidos, sincronización pendiente y doble clic.
14. Fixture anonimizado del Excel: demanda y cuellos de botella reales; recuperar calendario/docentes antes de prometer cierre al 30/05/2027. Capacidad nominal sola no prueba viabilidad.

Contar sesiones fechadas y horas, no SesionesRequeridas declaradas. Comprobar reglas y ausencia de solapamientos. TDD compilable: si se añaden contratos, crearlos antes de pruebas que los invoquen; RED por aserción, no por errores de tipo/esquema.

## Fase 1793 — Evidencia, Release y smoke posteriores a aprobación

Dependencia: 1792. No ejecutar durante esta revisión.

Confirmar branch, workspace, cambios ajenos, instrucciones locales, baseline y evidencia de HEAD real. No asumir copia fuera de OneDrive ni crearla automáticamente. Aplicar preflight vigente y documentar opción OneDrive si corresponde.

Tras autorización: Fast dirigido, Full Debug/x64, después Full Release/x64 con Configuration Release explícita. Ejecutar serializado. No exigir crecimiento artificial de suite; cero regresiones y ninguna reducción no explicada. Resolver discrepancias de baseline con TRX/evidencia real; nunca bajar para aprobar.

Luego smoke sobre copia de BD: tres rutas completas, legado/fijo, persistencia tras reinicio, rechazo, fallo y segunda corrida. No cerrar por Full Debug solamente. Registrar limitaciones externas.

Actualizar plan/handoff y evidencia con commit exacto verificado. Registro por procedimiento autorizado, append-only, sin reutilizar R-074 ni bajar contadores. Distinguir implementación terminada, pruebas aprobadas y smoke pendiente. No publicar ni modificar instalador/versionado sin alcance autorizado.

## Dependencias

1788 → 1789 → 1790 → 1791 → 1792 → 1793. Escritura secuencial recomendada por archivos compartidos. No declarar tareas completadas por existir un fragmento parcial de código.

## Veredicto de esta revisión

Implementación editada y pruebas preparadas. No recomendar release aprobado todavía: ejecutar verificación dirigida, Full Debug/x64 y Full Release/x64 tras autorización; smoke posterior sobre copia de BD. Registro sin nueva reserva.
