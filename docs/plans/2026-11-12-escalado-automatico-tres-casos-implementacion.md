# Plan de implementación: escalado automático de tres ventanas

Fecha: 2026-11-12. Entrega para otro modelo. Estado: propuesto; no ejecutado.

## Objetivo y límites

Ofrecer y aplicar ampliación de días cuando la programación no cabe en las fechas originales, considerando demanda por carril y restricciones de asignatura. Aprovechar los días añadidos redistribuyendo únicamente patrones automáticos. No ampliar FechaInicio/FechaFin automáticamente.

Casos obligatorios:
- LUN-JUE → LUN-VIE; añadir viernes.
- MAR-VIE → LUN-VIE; añadir lunes.
- VIE-DOM → JUE-DOM → MIE-DOM; añadir jueves y después miércoles; un paso por confirmación/corrida, máximo dos desde VIE-DOM.
- LUN-VIE y MIE-DOM terminales. Mantener VIE-SAB → JUE-SAB → ISO 4,5,6,7 terminal: dos pasos históricos.
- Conservar semántica existente: etiqueta JUE-DOM permite otro paso; ISO 4,5,6,7 permanece terminal. No convertir indistintamente etiqueta e ISO.
- 6M/6T son rótulos de salida. Nunca persistirlos dentro de dias_ventana ni pasarlos al parser de días.
- No compilar, ejecutar pruebas, instalar, publicar ni hacer commits durante la preparación. Ejecutar verificaciones únicamente tras aprobación del usuario.

## Diagnóstico que debe resolver

AgEscaladorVentanaAula.Evaluar descarta casos cuya demanda total cabe en capacidad total. No detecta saturación por patrones/carriles. AplicarAsync actualiza ventana, pero no trata los patrones previos.

AgFormGeneradorHorarios.GenerarEncabezadosAulasAsync siembra bandas por número de días y alterna asignaturas por módulo; ignora demanda y capacidad desigual. VIE-DOM produce días [5,6] y [7], equivalentes a tres y una ranuras. LUN-VIE produce bandas de dos, dos y una ranuras. Distribuir igual número de asignaturas no balancea estas capacidades.

AgPlanificadorOleadas conserva los días sembrados; AgParticionCarriles deja de buscar cuando ya no hay infactibles y solo acepta alternativas con menos infactibles. Con una fecha final ampliada puede aceptar resultados largos y poco ocupados.

La prueba anterior de 5W usa demandas 2W,2W,W y patrones ya balanceados. No cubre el flujo real de generación ni la demanda del Excel.

## Contrato funcional

1. Completar sincronizaciones pendientes y congelar selección/perfil durante evaluación y aplicación. Mantener guardas existentes de revisión de pendientes, subunidad, semáforo, cierre y reentrada.
2. Construir una instantánea con fechas originales, bloques, festivos, demanda por asignatura y restricciones. No modificar BD ni contexto para simular candidatos.
3. Evaluar programación actual. Intentar redistribución de patrones automáticos dentro de la misma ventana. Si basta, ofrecer reorganización sin añadir días; aplicar solo con aceptación cuando cambie configuración persistida.
4. Si sigue sin caber, evaluar únicamente AgEscaleraVentanas.Siguiente. No saltar pasos sin ganancia real.
5. Proponer el paso si añade ranuras reales y mejora la programación: menos sesiones sin ubicar o menos asignaturas infactibles; con igual factibilidad, menor fecha de terminación. Capacidad nominal mayor, por sí sola, no basta.
6. Un paso puede mejorar sin resolver toda la demanda: mostrar explícitamente demanda pendiente. Ejemplo 6W en VIE-DOM: JUE-DOM todavía puede quedar corto; siguiente corrida ofrece MIE-DOM. Nunca presentar solución parcial como factible.
7. Si restricciones fijas impiden aprovechar el día añadido, explicar el bloqueo; no prometer que escalar lo resuelve.
8. Rechazar propuesta: ninguna escritura. Aceptar: transacción, recarga desde BD y revalidación. Segunda corrida sobre un estado factible: ninguna propuesta.

## Fase 1736 — Contrato de aceptación y reproducción

Numeración propuesta de este documento, no reservada. Antes de registrar ejecución, contrastar REGISTRY y asignar rango libre sin reutilizar números. No reabrir planes cerrados.

Archivos: pruebas de AgEscaleraVentanas, AgPlanificadorOleadas, AgEscaladorVentanaAula y persistencia.

Crear fixtures de W semanas completas, calendario sin festivos y catálogo completo; separar pruebas con festivos, bloques ausentes, semanas parciales y restricciones fijas. Usar 4 h y 8 h. No usar Excel real, WinForms ni BD del usuario para unitarios.

Reproducir seis materias con sesiones [12,12,8,8,12,12]: total 64. En LUN-VIE, siembra antigua alternada carga 20 sesiones al viernes; demostrar que otro reparto cabe en el plazo del fixture. En VIE-DOM, siembra antigua por días puede cargar 32/36 sesiones al domingo; demostrar cuello de botella y una redistribución por ranuras factible en un calendario que tenga capacidad real suficiente.

No certificar RED si falla por compilación o esquema SQL. Agregar contratos compilables antes de tests que los invoquen. Registrar RED específico; prohibido debilitar restricciones para hacer verde una prueba.

Salida: fixtures y expectativas concretas, incluidos los casos de la matriz final.

## Fase 1737 — Origen y representación de patrones

Depende de 1736. Archivos: AgAulaAsignatura, persistencia/consultas/mappers de esa entidad, flujo de siembra, migración nueva según mecanismo real del repositorio.

Añadir origen explícito de patrón: AUTOMATICO, FIJO, LEGADO. Migración conserva dias_patron existente y lo marca LEGADO; no inferir origen por coincidir con una banda.

Para nuevas relaciones generadas por la aplicación: AUTOMATICO. No persistir una banda rotatoria como restricción fija. Un patrón vacío automático significa que lo asigna el planificador; no significa ausencia de clases.

FIJO mantiene la restricción declarada. LEGADO se trata conservadoramente como fijo hasta clasificación expresa. Para perfiles existentes, ofrecer conversión explícita y acotada de patrones seleccionados a AUTOMATICO, con vista previa por asignatura; sin conversión masiva silenciosa. Esto es necesario para corregir las aulas antiguas del Excel.

Usar AgRanura (día, franja, horas) en el plan generado. No reducir una asignación [5,6M] a la lista de días [5,6] y después regenerar ambos sábados. Patrones automáticos calculados en memoria; persistir su origen y las configuraciones necesarias para recomputarlos determinísticamente. Restricciones manuales por días siguen limitando días, sin inventar restricciones de franja.

Auditar todos los escritores de dias_patron: generación, edición, regeneración, importación y sincronización. Regenerar aulas no debe sobreescribir patrones FIJO/LEGADO.

Salida: ida y vuelta de origen y datos; legacy intacto; sábado mañana/tarde conservado en planificación.

## Fase 1738 — Planificación por ranuras y demanda

Depende de 1737. Archivos principales: AgPlanificadorOleadas, AgParticionCarriles, AgBandasDias y pruebas. AgCarriles/AgRanuras solo si hace falta adaptar consumidores; no cambiar reglas horarias.

Reemplazar siembra modular por asignación determinista de materias automáticas según sesiones requeridas y capacidad efectiva disponible. Carriles canónicos como referencia, nunca cuotas iguales por número de materias. Permitir asignar distinto número de materias a cada carril.

Evaluar capacidad y fechas usando el planificador temporal real; no solo suma de cargas: los redondeos de semanas, festivos y cursores también pueden impedir acomodación. Si greedy no encuentra solución, búsqueda acotada determinista con poda y memoización. Distinguir búsqueda agotada de infactibilidad demostrada; no ampliar calendario por un simple fallo heurístico.

Mantener ranuras disjuntas, unicidad de sesión del aula y conservación exacta de horas. Respetar restricciones fijas y reglas existentes R4/R5/R8/R10/R11 según su definición real. No eliminar controles para acomodar ejemplos.

Comparar candidatos con orden explícito: sesiones sin ubicar, asignaturas infactibles, última fecha programada, huecos internos elegibles, clave estable. Un candidato no puede perder asignaturas ya planificadas ni ampliar días fijos. Definir hueco sobre ranuras elegibles descontando festivos/bloques ausentes; no confundir capacidad nominal con disponible.

Corregir selector para considerar mejora temporal con igual número de infactibles y para trabajar con automáticos sin patrón sembrado. No relajar fidelidad de FIJO/LEGADO.

Salida: asignación reproducible, fechas dentro del fixture y aprovechamiento de viernes/lunes/jueves/miércoles añadidos cuando sean necesarios y elegibles.

## Fase 1739 — Evaluador compartido y decisión de escalado

Depende de 1738. Archivos: nuevo servicio interno puro en Academic/Application/Services/Genetica; AgEscaladorVentanaAula; AgValidadorVentana; AgFabricaIndividuos.

Centralizar selección de materias, filtro por programa/semestre, horas efectivas, unidad 4/8 h, patrones y métricas. Unificar fallback de HorasObjetivo con catálogo: escalador y planificador no deben medir demanda distinta.

Validador, escalador y fábrica deben consumir el mismo evaluador y resultado. Evitar tres implementaciones de factibilidad. Eliminar la salida anticipada basada exclusivamente en demanda <= capacidad total.

La decisión incluye identidad de aula y subunidad, perfil, ventana anterior/destino, fechas inmutables, cambios de origen autorizados, métricas antes/después, sesiones pendientes y huella de entrada. Distinguir REORGANIZACION de ESCALADO. No mantener Decision.Aplica basado únicamente en CapacidadDespues > CapacidadAntes.

Mantener etiquetas y parser existentes si sus contratos ya están verdes. Escalera.Siguiente decide ruta; el evaluador decide utilidad real del paso. Terminales no vuelven a proponerse.

Salida: propuestas justificadas por el mismo resultado que ejecutará la planificación, incluso cuando el total cabe pero hay saturación de carril.

## Fase 1740 — Aplicación transaccional e invalidación

Depende de 1739. Archivos: AgEscaladorVentanaAula y repositorios afectados; tests SQLite temporal.

Aplicar ventana en ag_aula y nodo del perfil activo y las conversiones/cambios autorizados de patrones en una misma transacción. Antes de escribir, comprobar huella y estado anterior de todas las filas; si cambió demanda, bloques, fechas, patrones o perfil desde la propuesta, abortar y recalcular. No usar únicamente conteo agregado de filas como prueba de éxito.

No tocar nodos de otros perfiles ni aulas de otra subunidad. Documentar que ag_aula/ag_aula_asignatura son datos compartidos si así lo confirma el esquema: no prometer aislamiento total de perfiles sin modelo que lo garantice. Si no hay perfil activo, no dejar claves visibles antiguas: sincronizar por identidad desde ag_aula o exigir perfil antes de ofrecer la acción.

Retornar Result con resultados por aula; rollback integral ante fallo/cancelación. Idempotencia por estado final. MarkDirty solo después de commit. No mutar memoria antes de confirmación de persistencia.

Recargar aulas, relaciones, orígenes y claves; reconstruir índices. Invalidar cache de planes afectada. Corregir huella de cache para incluir contenido de demanda, patrones/origen, bloques, festivos y fechas: contar filas no identifica cambios de contenido.

Salida: transacción, rollback, no interferencia, recarga real e idempotencia comprobados.

## Fase 1741 — Integración UI y flujo de optimización

Depende de 1740. Archivos: AgFormGeneradorHorarios y coordinación de diagnóstico/optimización.

Unificar el pre-vuelo usado por Diagnosticar y Optimizar. No permitir que Optimizar omita el flujo de escalado confirmado o utilice decisiones antiguas. No aplicar cambios sin confirmación.

Modal: aula, ruta inmediata, días añadidos, fechas originales sin cambios, horas/sesiones ubicadas antes/después, pendientes restantes y cambios de patrones automáticos. Si necesita conversión LEGADO, separarla claramente y listar materias afectadas. No mostrar 5W como garantía sin calendario y restricciones del caso.

Mantener semáforo, guardas de reentrada, revisión de pendientes y subunidad capturada. Bloquear todas las acciones que puedan cambiar contexto durante evaluación/aplicación. Recargar solo nodos seleccionados con validación previa en claves temporales. Error de guardado/recarga: abortar planificación y conservar cambios pendientes.

Siguiente corrida desde JUE-DOM puede ofrecer MIE-DOM si aún es insuficiente. Si JUE-DOM basta, no ofrecer segundo paso. Ofrecer ampliación de fechas solo como acción separada después de explicar insuficiencia residual; nunca usarla para ocultar fragmentación de patrones automáticos.

Salida: mismos resultados desde ambos botones; cancelación sin escrituras; segunda corrida estable.

## Fase 1742 — Pruebas integrales y validación del caso real

Depende de 1741. No ejecutar hasta aprobación.

Matriz obligatoria:
| Caso | Resultado |
|---|---|
| LUN-JUE, demanda 5W | Ofrece LUN-VIE; 4W→5W nominal; acepta; programa 5W sesiones distintas; viernes usado; recarga e idénticas fechas; segunda corrida sin propuesta |
| MAR-VIE, demanda 5W | Mismo contrato; lunes usado |
| VIE-DOM, demanda 5W | Ofrece JUE-DOM; 4W→5W; jueves usado; segunda corrida sin propuesta si todo cabe |
| VIE-DOM, demanda 6W | Primera corrida solo JUE-DOM y pendiente explícito; recarga; segunda MIE-DOM; miércoles usado; tercera sin propuesta |
| MIE-DOM/LUN-VIE | Terminales; insuficiencia residual informada sin inventar otro escalón |
| VIE-SAB | Dos pasos históricos intactos; ISO 4,5,6,7 terminal |
| Paso inmediato sin ganancia efectiva | No saltar al siguiente; causa explícita |
| Demanda global cabe pero carril automático saturado | Reorganiza dentro de ventana si es viable; si no, evalúa siguiente paso |
| Restricción fija domingo | No convertir en jueves/sábado ni prometer resolverla por capacidad global |
| Sábado 4 h | Dos ranuras disjuntas, sin duplicar ni perder mañana/tarde |
| Sábado 8 h | Una ranura; capacidades VIE-DOM=3, JUE-DOM=4, MIE-DOM=5 por semana completa |
| Rotular MIE-DOM | 4 h: 3,4,5,6M,6T,7; 8 h: 3,4,5,6,7 |
| Rechazo/fallo SQL/huella obsoleta | Sin mutación parcial ni planificación optimista |
| Mismo conteo, distinto contenido | Cache invalidada y plan recalculado |

Contar sesiones materializadas y horas reales, no solo SesionesRequeridas. Verificar ausencia de solapamientos por aula/docente donde corresponda y cumplimiento de reglas.

Referencia real: C:/Users/USUARIO/Downloads/Libro3.xlsx; Semestres. Periodo 2027-01, límite original confirmado 30/05/2027; exportación termina 17/10/2027. 136 materias, 5952 h, 800 h posteriores al límite. El Águila III-A: filas 22,25,28,31,34,37; tres materias de 48 h dominicales requieren 36 domingos y solo hay 17 nominales en ventana original.

Convertir esos datos a fixture mínimo anonimizado. Recuperar por separado inicio real, bloques, festivos, restricciones y disponibilidad docente de la corrida. El Excel no contiene esos datos completos ni prueba estados persistidos. No afirmar que los 23 grupos caben antes del 30/05 solo con capacidad nominal. Resultado esperado: cabe con restricciones reales o informe específico de imposibilidad; nunca alargar silenciosamente hasta octubre.

## Fase 1743 — Gates y entrega

Depende de 1742. Aprobación del usuario necesaria antes de ejecutar.

- Leer instrucciones y preflight vigentes. No asumir workspace limpio ni hacer commits de cambios ajenos. Coordinar con otros planes que toquen los mismos archivos.
- Verificar baseline: archivo consultado declara 1223; REGISTRY R-065 declara Full 1236. Resolver contra evidencia ejecutada antes de actualizar cifras. Nunca bajar baseline para aprobar.
- Tras autorización: Fast por componente y Full Debug/x64; luego Full Release/x64 con -Configuration Release explícito. Builds/tests serializados. Confirmar comandos y filtros existentes antes de usarlos.
- Revisar alta de archivos/migración en csproj y compatibilidad .NET Framework 4.8/C# 7.3; no incrementar versión ni modificar instalador salvo alcance autorizado.
- Smoke sobre copia de BD y perfil: tres rutas, aceptación/rechazo, persistencia tras reinicio, patrones legacy/fijos, 4/8 h, error de guardado y segunda corrida. Sin modificar BD del usuario durante pruebas.
- Guardar evidencia por caso: entrada, huella, propuesta, aceptación, datos antes/después, fechas, ocupación efectiva, sesiones, pendientes y commit verificado.
- Entregar resumen, archivos modificados, migración y tratamiento de legacy, pruebas ejecutadas y límites. No declarar terminado por parser/escalera verdes si falta flujo completo o smoke.

## Dependencias y exclusiones

1736 → 1737 → 1738 → 1739 → 1740 → 1741 → 1742 → 1743. Ejecución secuencial propuesta. Este documento no ordena crear subagentes, reservar números ni iniciar implementación.

Fuera de alcance: publicación, instalador, ampliación automática de calendario, relajación de restricciones explícitas y corrección masiva de datos de producción.
