# Partición de Carriles Consciente del Patrón Sembrado (`AgParticionCarriles`)

> **Fecha:** 2026-11-11  
> **Ronda:** 312 | **Fase:** 1581 | **Registro:** `R-034`  
> **Componente:** `Academic/Application/Services/Genetica`

---

## 1. Problema y Diagnóstico Crítico

En `AgPlanificadorOleadas`, la partición de ranuras en carriles se realizaba rígidamente utilizando la partición canónica fija de `AgCarriles` (bloques de 2 ranuras contiguas: p. ej., `[4,5]|[6]|[7]`).
Cuando las asignaturas tenían patrones sembrados (`dias_patron`, p. ej. `"6,7"` en ventana `"4,5,6,7"`), el patrón no quedaba contenido en ningún carril canónico individual. Esto provocaba que la asignatura fuera rechazada de la pasada 1 y, en la pasada 3, fallara por la restricción `R8_DiaFueraDePatron`, marcándola como infactible aun cuando la capacidad del aula y el número de semanas hábiles eran suficientes.

---

## 2. Solución Arquitectónica (Opción C)

1. **Evaluador Único Reutilizable:** Se extrajo la lógica de asignación y oleadas de `AgPlanificadorOleadas` en un método privado `EvaluarParticion(ContextoEvaluacion ctx, string diasVentana, IReadOnlyList<IReadOnlyList<AgRanura>> bandas)`.
2. **Reordenamiento Defensivo:** Se evaluó `R8_DiaFueraDePatron` antes de invocar la construcción de `AgPatronSemanal`, evitando trabajo computacional innecesario y preservando idénticos mensajes y orden de `DetalleInfactibles`.
3. **Componente Selector `AgParticionCarriles`:**
   - Enumera particiones contiguas de días provenientes de `AgBandasDias.Particiones(ventana)` acotadas entre $k_{min}$ y $k_{actual}$.
   - Deduplica por firma de ranuras (`Día:Franja:Horas`).
   - Aplica los 4 filtros de aceptación conservadora estricta frente a la referencia canónica:
     $$\text{NoAlojables}(c) < \text{NoAlojables}(ref)$$
     $$\text{Alojables}(ref) \subseteq \text{Alojables}(c)$$
     $$\text{Planes}(ref) \subseteq \text{Planes}(c)$$
     $$|\text{Infactibles}(c)| < |\text{Infactibles}(ref)|$$
   - Desempate determinista lexicográfico: `Infactibles` ascendente, `PatronesNoAlojables` ascendente, `Makespan` ascendente, `Bandas` descendente, `MaxRanurasPorBanda` ascendente, `Clave` ordinal.
4. **Trazabilidad:** Se registra un log informativo mediante `LoggerHelper.Info` únicamente cuando `seleccion.Cambiada == true`.
5. **Gate G-1 (Evaluación de propagación a genes):** Se comprobó que las bandas seleccionadas agrupan consistentemente los días sembrados requeridos (`plan.RanurasPatron`), documentando el impacto en los días permitidos.

---

## 3. Pruebas y Evidencia

- **Caracterización:** `AgPlanificadorOleadasCaracterizacionTests` (7/7 PASS) validó que los casos canónicos sin patrones, patrones ya alojables, R8 cruzado y fallos de capacidad se mantienen 100% idénticos.
- **Selector Sintético:** `AgParticionCarrilesTests` (12/12 PASS) verificó cada condición de guarda, rechazo y desempate lexicográfico.
- **Integración:** `AgParticionCarrilesIntegracionTests` (5/5 PASS) comprobó el caso homogéneo real (`"6,7"`), el caso mixto (`"4,5"` y `"6,7"`), el desempate por no perder asignaturas previas y el manejo de festivos.
- **Gate G-1:** `AgGateG1ParticionCarrilesTests` (1/1 PASS).
- **Suite Completa:** `verify.ps1 -Mode Fast` ejecutó **1156/1156 tests PASS** sin regresiones.

