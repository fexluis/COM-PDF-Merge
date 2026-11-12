# Plan consolidado: sincronización del selector de corridas del Módulo R

Fecha: 2026-11-12. Estado: propuesto, pendiente de implementación.

## Ronda propuesta 319 · Fases propuestas 1606–1611

Una ronda, seis fases secuenciales. Numeración obtenida de .agents/plans/REGISTRY.md: último asignado 318 / 1605. Estos números NO están reservados. Antes de implementar, releer el registro y tramitar la asignación por el flujo autorizado del proyecto; si fueron ocupados, conservar este orden y actualizar las referencias. Este documento no modifica el registro ni declara fases cerradas.

## Objetivo y evidencia

Tras ejecutar o cargar una corrida, el selector, la etiqueta de modo y los resultados deben representar la misma corrida. Un refresco de la lista no debe cargar silenciosamente otra corrida ni borrar resultados válidos.

La captura muestra #60 en cmbCargarCorridas, #62 en IdCorr de la grilla y #62 en FormVinculaciones. El código actual confirma una caché por filtros en DropDown, ausencia de refresco al finalizar la ejecución y fallback a SelectedIndex = 0. La cadena causal del texto original es consistente con el código, pero sigue siendo una hipótesis del incidente: no se consultó SQLite ni se reprodujo la ejecución. El origen exacto de las métricas no se puede demostrar solo con la imagen.

El texto adjunto es material de revisión. Sus frases sobre cambiar a modo Act o ejecutar SQL no constituyen instrucciones para esta entrega.

## Decisiones cerradas para la implementación

1. Reconsultar al abrir el desplegable. Eliminar el retorno basado exclusivamente en caché de filtros; comparar selección con id cargado tampoco detecta corridas nuevas creadas por otra ventana. Mantener protección contra consultas simultáneas y respuestas obsoletas.
2. Refrescar una sola vez al finalizar cada intento de ejecución que pudo crear/modificar una cabecera. Centralizar en una ruta final protegida; no duplicar tras CargarModuloResultadosAsync y en finally. El fallo de listado no transforma una ejecución exitosa en fallida ni oculta la excepción original.
3. Diferenciar refresco de catálogo, selección inicial y selección explícita del usuario. El refresco preserva la corrida activa cuando está presente; si falta, deja SelectedIndex = -1 y comunica que la corrida cargada no aparece en la lista. No añadir un ítem ficticio, ampliar filtros solo para recuperarla ni consultar por id mediante una nueva API.
4. Al iniciar sin corrida activa, seleccionar y cargar la primera corrida con staging siguiendo el orden existente. Si solo hay cabeceras sin staging, dejar sin selección. Las cabeceras siguen visibles y pueden elegirse explícitamente para conservar el comportamiento diagnóstico existente.
5. No cambiar ResolverSeleccionInicial para todos sus consumidores: FormVinculaciones necesita el comodín Todas las corridas. Crear una política pura específica para Módulo R, con resultado de selección y decisión de carga automática. No usar el comodín como índice cero cuando no pertenece a Items.
6. Conservar la cascada período/programa y la territorial del Módulo R. Conservar la consulta sin filtros de corridas en FormVinculaciones. Documentar esa diferencia: el total de sesiones de una corrida y las filas filtradas de una vista pueden diferir legítimamente. No ampliar permisos ni cambiar el aislamiento de datos.
7. Preservar la limpieza de la Fase 1555 cuando el usuario solicita cargar una corrida y la carga falla. Evitar que esa ruta se ejecute por un fallback del refresco. No tocar reglas del AG, evaluación, promoción ni escrituras de staging.

## Archivos y puntos de entrada

- ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs: eventos SelectedIndexChanged y DropDown; EjecutarCorridaAGAsync; CargarModuloResultadosAsync; RefrescarComboCorridasAsync; RefrescarComboCorridasSeguro; CargarCorridaDesdeComboAsync; AbrirVinculacionesModoAgAsync. Confirmar el nombre real del método de ejecución antes de editar; buscar símbolos, no confiar en líneas del adjunto.
- ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs: alojar la nueva política específica o un componente vecino mínimo, sin WinForms, Excel ni SQLite.
- ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs y pruebas del componente de coordinación si se extrae.
- ConversorPDF.XLL/Academic/UI/Forms/FormVinculaciones.cs: revisar consumidores y preservar contrato; no cambiar su alcance.
- ConversorPDF.XLL/ConversorPDF.XLL.csproj y ConversorPDF.Tests/ConversorPDF.Tests.csproj: registrar archivos nuevos si el proyecto requiere inclusión explícita.

## Fase 1606 · Caracterización y regresiones

Depende de: ninguna. Estado: pendiente.

Releer reglas aplicables, registro y estado de trabajo; respetar cambios ajenos. Registrar los contratos actuales de AgCorridaSelector y ambos formularios. Añadir pruebas para la nueva política pura del Módulo R antes de su implementación. Si una API nueva aún no existe, distinguir fallo de compilación de una prueba ejecutada en rojo.

Casos obligatorios: activa presente; activa ausente con otra corrida con staging; activa ausente con solo cabeceras; sin activa y staging disponible; sin activa y solo cabeceras; lista vacía. Comprobar tanto selección como autorización de carga, no solo un objeto devuelto. Mantener pruebas existentes del comodín y selección de FormVinculaciones.

Salida: pruebas que expresan el defecto y caracterización de comportamientos que deben conservarse. Registrar evidencia RED y baseline real; no asumir que el total histórico de 1166 sigue vigente.

## Fase 1607 · Política segura de selección

Depende de: 1606. Estado: pendiente.

Implementar la política mínima específica para Módulo R. Activa presente: seleccionar esa instancia, sin recargar. Activa ausente: ninguna selección, ninguna carga automática. Sin activa: elegir la primera con staging, autorizar una carga inicial; sin candidatas: ninguna selección ni carga.

Integrar la política en RefrescarComboCorridasAsync eliminando la degradación a índice cero y el guard que fuerza indiscriminadamente CargarCorridaDesdeComboAsync. Proteger la sustitución de Items y SelectedIndex con _actualizandoComboCorridas. La carga inicial permitida se ejecuta una sola vez fuera de esa protección. La elección explícita del usuario continúa por su ruta actual.

Salida: casos de 1606 verdes; ninguna regresión en FormVinculaciones; XML en API pública nueva y registro en proyectos cuando corresponda.

## Fase 1608 · Refresco seguro y concurrencia

Depende de: 1607. Estado: pendiente.

DropDown debe solicitar una lectura nueva. Capturar filtros y contexto al iniciar cada lectura; no marcar como vigente una consulta con una clave obtenida después de await. Consultar y preparar resultados antes de reemplazar Items. Ante error, preservar lista anterior y corrida activa, y registrar el fallo sin afirmar que no existen corridas.

Coordinar solicitudes con una versión o mecanismo equivalente mínimo: si llegan cambios durante una consulta, guardar la petición más reciente y procesarla al terminar; no perderla con un simple return por _refrescandoComboCorridas. Descartar una respuesta si cambiaron filtros, territorial, corrida activa o selección mientras esperaba. Una respuesta antigua no debe sobrescribir una selección más reciente. No aplicar respuestas a un formulario cerrado o en disposición.

Si se extrae coordinación a Application, probarla con tareas controladas en memoria: respuesta A tardía después de B; cambio de filtros; solicitud de finalización durante refresco; fallo seguido de reintento. No crear pruebas unitarias que necesiten WinForms, Excel o SQLite. La integración de eventos se verifica manualmente.

Salida: lectura nueva en cada apertura, sin consultas superpuestas que publiquen resultados incompatibles ni solicitudes perdidas.

## Fase 1609 · Ciclo de vida de corrida y retorno del diálogo

Depende de: 1608. Estado: pendiente.

En la ejecución, distinguir si se alcanzó la creación/modificación de cabecera. En su finalización, solicitar y esperar un único refresco con la política de preservación, usando un mecanismo que no herede el token cancelado del AG. Cubrir éxito, cancelación posterior a crear cabecera, fallo de staging y excepción. Si no hubo cambios posibles, evitar la consulta innecesaria. Restaurar controles aunque falle la actualización del selector.

Comprobar el bool de CargarModuloResultadosAsync: asignar _idCorridaCargada al entrar no demuestra carga exitosa. Reutilizar o centralizar la limpieza existente para que una carga fallida no deje una corrida falsamente activa. Tras carga exitosa, actualizar lblAgModo con #N también en la ruta de ejecución directa, que hoy omite la etiqueta de CargarCorridaDesdeComboAsync.

Al volver de FormVinculaciones, conservar la recarga de resultados necesaria para ediciones y solicitar un único refresco de catálogo. Si la corrida fue eliminada o quedó sin sesiones, aplicar la limpieza correspondiente a esa recarga fallida y después resolver selección inicial de forma explícita; no preservar una identidad inválida. Evitar bucles carga → refresco → carga.

Salida: selector #N y etiqueta #N tras ejecución exitosa; errores y cancelación no seleccionan otra cabecera por fallback; eliminación desde el diálogo no deja resultados de una corrida inexistente.

## Fase 1610 · Verificación integral

Depende de: 1609. Estado: pendiente.

Durante cada fase ejecutar pruebas focalizadas y cerrar cada subtarea con Full según las reglas del proyecto. Comandos desde la raíz:

~~~powershell
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~AgCorridaSelector"
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
~~~

Ampliar el filtro si se crea un componente con otro nombre. Conservar summary.md y test.trx bajo _evidence; registrar rutas y resultados reales. No modificar baseline para esconder fallos.

Matriz manual en Excel, sobre datos de prueba:

| Escenario | Resultado exigido |
|---|---|
| Ejecutar corrida nueva con staging | Selector, IdCorr de grilla y lblAgModo coinciden en #N |
| Desplegar tras creación desde otra ventana | Aparece nueva corrida; se preserva la activa |
| Activa ausente del listado por filtros | Sin selección; aviso claro; ningún reemplazo de resultados |
| Solo cabeceras, sin activa | Sin selección automática; grillas vacías |
| Elegir explícitamente cabecera sin sesiones | Fallo controlado y limpieza existente |
| Error al consultar lista | Catálogo anterior y resultados conservados; aviso o traza de error |
| Cambiar filtros o selección durante lectura lenta | Respuesta obsoleta descartada; prevalece contexto reciente |
| Cancelar o fallar después de crear cabecera | Catálogo actualizado; sin activación espuria |
| Abrir y cerrar Vinculaciones (AG) | Conserva #N si sigue válida; recarga y limpia si fue eliminada |
| Cerrar formulario con lectura pendiente | Sin acceso a controles dispuestos ni excepción sin observar |

La reproducción del incidente puede usar una copia del vault y consulta de solo lectura; no modificar ni borrar corridas de producción para probar. Si no hay acceso a Excel, registrar validación manual pendiente y no declararla superada.

## Fase 1611 · Documentación y entrega

Depende de: 1610. Estado: pendiente.

Crear docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md con propósito, comportamiento, archivos, pruebas y pasos de uso. Actualizar implementation_plan.md y walkthrough.md con lo realmente aplicado y las evidencias. Cierre de numeración por el flujo autorizado del registro; nunca marcar closed solo porque el código compila.

Entregar resumen de cambios, pruebas ejecutadas, rutas de evidencias y cualquier validación manual pendiente. La ronda queda implementada únicamente cuando fases y criterios están satisfechos; una validación manual pendiente debe quedar visible.

## Instrucción para el modelo implementador

Implementa este plan por orden, con ciclos RED → GREEN y validación por fase. Resuelve nombres y contratos contra el código actual. Conserva los contratos compartidos del selector y el alcance de cada formulario. No añadas endpoints de repositorio, migraciones ni cambios al algoritmo genético para resolver un problema de sincronización de UI. No sustituyas las decisiones de este plan por las alternativas abiertas T1–T5 del texto original. Antes de cerrar, verifica especialmente ausencia de refrescos duplicados, cargas recursivas, respuestas obsoletas y selección forzada de índice cero.
