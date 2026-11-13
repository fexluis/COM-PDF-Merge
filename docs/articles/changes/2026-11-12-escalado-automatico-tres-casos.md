# Escalado Automático de Ventanas en Tres Casos: LUN-JUE, MAR-VIE y VIE-DOM (Fases 1734–1741)

**Fecha**: 2026-11-12  
**Módulo**: Académico / Algoritmo Genético  
**Versión base**: 1.6.4.09  
**Autores**: Antigravity / Swarm Parallel Engine  

---

## 1. Motivación y Problema Resuelto

En escenarios de optimización de horarios académicos (.NET Framework 4.8 / Excel-DNA), existían configuraciones de ventana semestral donde la demanda global de sesiones cabía aritméticamente en el número total de ranuras de la ventana, pero colapsaba por asimetría de carriles en días específicos.

El caso más crítico ocurría en la ventana `VIE-DOM` ("5,6,7"):
- En una ventana típica de 17 semanas con unidad de 4 horas:
  - Viernes (1 ranura) + Sábado (2 ranuras: mañana y tarde) = 3 ranuras por fin de semana (51 ranuras en 17 semanas).
  - Domingo (1 ranura: mañana) = 17 ranuras en 17 semanas.
  - Capacidad global total = 68 ranuras.
- Si una cohorte demandaba 4 asignaturas de 16 sesiones (64 sesiones en total), la capacidad global (68) parecía suficiente.
- Sin embargo, la siembra ingenua en dos bandas (Banda 0: [5,6]; Banda 1: [7]) asignaba la mitad de las materias al Domingo (32 sesiones), superando la capacidad del Domingo (17) por 15 sesiones.
- Como `AgEscaladorVentanaAula.Evaluar` descartaba el caso si `demandaTotal <= capacidadTotal`, el escalador no detectaba la saturación y el motor genético se veía forzado a dejar sesiones sin ubicar o a desbordar el calendario hasta meses posteriores fuera de la ventana semestral.

---

## 2. Solución Arquitectónica Implementada

### A. Origen Explícito de Patrones (`AgOrigenPatron` y Migración V80)
Se incorporó el concepto de procedencia del patrón de días en `ag_aula_asignatura`:
- `AUTOMATICO`: Patrones asignados dinámicamente por el planificador según la demanda.
- `FIJO`: Restricciones inmutables expresadas por el usuario.
- `LEGADO`: Registros preexistentes en la base de datos (tratados conservadoramente como fijos).

Se introdujo `MigrationV80_AddOrigenPatronAgAulaAsignatura.cs` para añadir la columna `origen_patron TEXT DEFAULT 'LEGADO'` sin alterar la integridad de datos preexistentes.

### B. Evaluador de Demanda y Saturación de Carriles (`AgEvaluadorDemandaVentana`)
Se desacopló y enriqueció la lógica de evaluación en `AgEvaluadorDemandaVentana.cs`:
1. Partición de la ventana en bandas según la cardinalidad de días.
2. Cálculo exacto de ranuras reales considerando días festivos y bloques activos de la franja.
3. Evaluación de factibilidad entera mediante backtracking (`EsParticionableEnBandas`): si las materias con patrones fijos o automáticos no pueden distribuirse sin saturar alguna de las bandas, se marca `HaySaturacion = true` y `SaturacionCarriles = true`.

### C. Escalado Estricto de Un Paso y Detección de Demanda Pendiente
Siguiendo el contrato funcional:
- Solo se evalúa el escalón inmediato determinado por `AgEscaleraVentanas.Siguiente`.
- Si el escalón mejora la capacidad o resuelve la saturación, se formula la propuesta de escalado.
- Si el escalón no es suficiente para cubrir toda la demanda (por ejemplo, 6 materias = 96 sesiones en `VIE-DOM` escalando a `JUE-DOM` con capacidad 85), la decisión registra explícitamente la **demanda pendiente** (11 sesiones) y advierte al usuario que se trata de una solución parcial que requerirá un segundo paso.
- En la siguiente corrida, el aula en `JUE-DOM` con 96 sesiones se propone para `MIE-DOM` (capacidad 102), alcanzando un estado 100% factible. Una tercera corrida sobre `MIE-DOM` resulta en 0 propuestas (estabilidad de segunda corrida).

### D. Invalidación Transaccional de Patrones Automáticos
Al ejecutarse `AgEscaladorVentanaAula.AplicarAsync`:
- Se actualiza `dias_ventana = d.Hacia` en `ag_aula`.
- Se actualiza `dias_semana = d.Hacia` en los nodos de `ag_arbol_nodo` del perfil activo.
- Se resetea `dias_patron = ''` exclusivamente para los registros con `origen_patron = 'AUTOMATICO'`, permitiendo que el planificador redistribuya libremente sobre la nueva ventana ampliada.

### E. Unificación en Interfaz de Usuario (`AgFormGeneradorHorarios`)
Se integró el mismo flujo de diagnóstico y diálogo modal interactivo tanto en `EjecutarDiagnosticoPreVueloAsync` como en `IniciarOptimizacionAsync`:
- Presenta el aula, la ruta (`Desde → Hacia`), los días añadidos, la capacidad anterior y posterior, la demanda requerida y la demanda pendiente si aplica.
- Exclusión mutua garantizada mediante `SemaphoreSlim` (`_semaforoSyncDimensiones`) y bloqueo temporal de controles de edición durante la evaluación y recarga.
- Recarga segura de las dimensiones del perfil en el árbol (`RecargarDimensionesPerfilActivoTrasEscaladoAsync`).

---

## 3. Matriz de Validación y Cobertura de Pruebas

Se implementaron y certificaron suites dedicadas en `ConversorPDF.Tests`:
- `AgEscaladoTresCasosTests.cs`: Reproducción de cuellos de botella y contrato inicial (3 tests).
- `AgOrigenPatronTests.cs`: Comportamiento de enum, migración V80 y asignación (3 tests).
- `AgParticionCarrilesTests.cs`: Alternativas de partición y resiembra sin degradación (14 tests).
- `AgEvaluadorDemandaVentanaTests.cs`: Detección de saturación global y por carril en domingo (4 tests).
- `AgEscaladorVentanaAulaPersistenciaTests.cs`: Persistencia transaccional e invalidación selectiva de patrones (13 tests).
- `AgEscaladoTresCasosMatrizTests.cs`: Validación exhaustiva de los tres casos, multi-paso 6W con pendiente, terminalidad de `MIE-DOM` y `LUN-VIE`, y diferenciación sábado 4h (68 ranuras) vs 8h (51 ranuras) (6 tests).

**Veredicto Final**: 1259 tests ejecutados, 1259 superados, 0 fallos, 0 regresiones.

