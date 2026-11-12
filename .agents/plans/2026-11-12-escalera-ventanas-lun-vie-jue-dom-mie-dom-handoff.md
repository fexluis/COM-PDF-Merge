# Handoff: Escalado Automático de Ventanas LUN-VIE / JUE-DOM / MIE-DOM (R-058)

**Fecha**: 2026-11-12  
**Plan asociado**: `.agents/plans/2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom-plan.md`  
**Registro**: Fila `R-058` en `.agents/plans/REGISTRY.md` (Ronda 345–350, Fases 1707–1714)  
**Estado inicial**: `reserved` (Fase A completada)

---

## 1. Contexto y Problema

El planificador académico y el escalador de ventanas (`AgEscaladorVentanaAula`) actualmente operan con una escalera restringida (`LUN-JUE`, `MAR-VIE`, `VIE-SAB`, `JUE-SAB`, `"4,5,6,7"`).
Se requiere incorporar de forma segura y canónica las ventanas extendidas:
- `LUN-VIE` ("1,2,3,4,5") como escalón de `LUN-JUE` y `MAR-VIE`.
- `JUE-DOM` ("4,5,6,7") y `MIE-DOM` ("3,4,5,6,7") como escalones sucesivos de `VIE-DOM`.

### Hallazgos de la Revisión Estática Integrados (Ronda 344 / Fases 1701–1706)
1. **Fase 1701 (Grafo y Olas)**: Se ajusta el orden a 6 olas y 8 tareas para respetar causalidad estricta y leases disjuntos (≤ 2 escritores concurrentes).
2. **Fase 1702 (Compilación RED)**: Introducción de stub previo en T1 para `Rotular` evitando errores de compilación CS en la suite de tests (T2/T3).
3. **Fase 1703 (Evaluación estricta de `Siguiente`)**: `AgEscaladorVentanaAula.Evaluar` evalúa exclusivamente `AgEscaleraVentanas.Siguiente()`, impidiendo saltarse escalones inmediatos sin ganancia.
4. **Fase 1704 (Capacidad Nominal vs Efectiva)**: Definición formal de fixtures (`4W -> 5W`) con semanas completas `W`, sin festivos espurios y catálogo completo.
5. **Fase 1705 (Persistencia y Recarga)**: Pruebas sobre SQLite temporal verificando mutación aislada sobre `idArbol` y ausencia de efectos colaterales en otros árboles.
6. **Fase 1706 (UI Segura y Leyenda no Clicable)**: Recarga específica de DTOs de nivel 6 y renderizado de tooltip/leyenda fuera de la zona de hit-test del selector de dimensiones.

---

## 2. Olas de Ejecución

- **Wave 1 (Ronda 345 / Fases 1707–1708)**:
  - T1: Stub de `Rotular` en `AgEscaleraVentanas.cs` (habilita compilación RED).
  - T2: RED del parser en `AgDiasVentanaTests.cs` (falla por aserción).
- **Wave 2 (Ronda 346 / Fases 1709–1710)**:
  - T3: RED de dominio, integración y persistencia (múltiples fixtures falsables).
  - T4: Implementación parser canónico en `AgDiasVentana.cs` (GREEN para T2).
- **Wave 3 (Ronda 347 / Fase 1711)**:
  - T5: Implementación de la escalera ampliada y `Rotular` real en `AgEscaleraVentanas.cs`.
- **Wave 4 (Ronda 348 / Fase 1712)**:
  - T6: Implementación en `AgEscaladorVentanaAula.cs` evaluando solo `Siguiente`.
- **Wave 5 (Ronda 349 / Fase 1713)**:
  - T7: UI en `AgFormGeneradorHorarios.cs` (modal genérico, recarga acotada y tooltip seguro).
- **Wave 6 (Ronda 350 / Fase 1714)**:
  - T8: Documentación de cambio y cierre formal del plan (`Mode Full` con 1172 tests mínimo).

---

## 3. Instrucciones para la Fase B (`swarm-parallel-task`)

1. Validar preflight con `-AllowOneDrive` en el workspace actual.
2. Invocar `plan-reviewer` para auditar `.agents/plans/2026-11-12-escalera-ventanas-lun-vie-jue-dom-mie-dom-plan.md`.
3. Iniciar Wave 1 (Ronda 345) ejecutando T1 y T2 en paralelo.

