# Correcciones de reclasificación y fechas

## Purpose

Cerrar errores de autorización, ocupación de ranuras y pruebas detectados al revisar la entrega de redistribución. Sin modificar datos del usuario ni ejecutar el AG.

## Changes

- AgFormGeneradorHorarios: Cancelar o cerrar la selección detiene el flujo; no continúa con escalados nominales.
- AgReclasificacionPatron y AgEscaladorRanuraAula: identidad completa obligatoria; rechazo de duplicados, relaciones ajenas, origen/patrón anterior desactualizado y transiciones no previstas. Solo LEGADO puede pasar a AUTOMATICO (patrón vacío) o FIJO (mismos días). Selección copiada y normalizada antes de simular.
- Huellas de memoria/SQL y postcondiciones: coincidencia conjunta de aula, referencia, subunidad y asignatura autorizada. Guardado con origen y días anteriores como condiciones.
- AplicarDetalladoAsync: comprobar huellas antes de filtrar decisiones no aplicables. Una acción alterada para desactivarla también se rechaza; candidatos intactos no seleccionados se omiten.
- AgCoherenciaGrupo: reemplazar etiquetas de franja por comparación de horas de inicio/fin. Consultar fechas actuales de las otras asignaturas, sin eliminar ocupaciones compartidas por error. Buscar huecos desde el inicio, conservar fechas válidas, acotar al plazo y mantener horas demandadas si no alcanza la capacidad.
- AgMaterializadorSesiones: reparación de cada grupo consulta ocupación de todos los genes del individuo.
- Pruebas: eliminado bloque duplicado de aserciones fuera de un método en AgRedistribucionRanurasIntegracionTests. Corregido día/rango del fixture dominical y asignatura del fixture de aislamiento, sin reducir sus aserciones.

## How to use

Seleccionar materias, simular, revisar la propuesta definitiva y confirmar. Cancelar detiene el flujo. Patrones FIJO y LEGADO no seleccionados se conservan.

Tras aprobación, verificar con el flujo de pruebas del repositorio y un pre-vuelo DEBUG hasta 30/05/2027. Revisar pendientes y colisiones globales antes de considerar release. Esta reparación consulta ocupación del individuo recibido; no sustituye validadores de disponibilidad docente ni de otras corridas.

## Test coverage

AgRedistribucionCorreccionesTests agrega casos de transición válida/inválida, identidad incorrecta, duplicados, aislamiento entre aulas, copia de autorización, huella alterada al desactivar una acción, huecos anteriores, intervalos solapados/adyacentes, otra aula y falta de capacidad con demanda conservada.

AgRedistribucionRanurasAceptacionTests refuerza aislamiento mediante evaluación real de la selección.

Solo revisión estática y comprobación de diferencias. Compilación, pruebas y AG no ejecutados por restricción del usuario. No se declara aprobación de release.
