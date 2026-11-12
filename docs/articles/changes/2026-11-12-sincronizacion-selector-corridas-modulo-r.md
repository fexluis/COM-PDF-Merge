# Sincronización del Selector de Corridas del Módulo R (Fases 1606–1612)

## 1. Contexto y Problema

Durante las pruebas operativas del Módulo R (Generador y Optimizador de Horarios por Algoritmo Genético), se observó un desfasaje crítico entre la interfaz de selección y los resultados presentados en pantalla: tras ejecutar o cargar una optimización, el desplegable `cmbCargarCorridas` mostraba una corrida anterior (por ejemplo, `#60`), mientras que la etiqueta `lblAgModo` y las grillas de horarios generados (`dgvHorariosGenerados`) y violaciones (`dgvViolaciones`) reflejaban los datos de la corrida activa recién calculada (por ejemplo, `#62`).

Asimismo, al desplegar manualmente el combo para buscar la corrida en curso, la lista no se actualizaba con las corridas recién creadas en la base de datos de staging. En otros casos, al interactuar con el combo o disparar refrescos automáticos, una corrida activa era reemplazada de inmediato por la primera opción de la lista (`SelectedIndex = 0`), sobreescribiendo los resultados evaluados en memoria y obligando al usuario a reiniciar el proceso.

### Causas Raíz Identificadas

El análisis exhaustivo del flujo de datos reveló tres causas fundamentales:

1. **Bloqueo del evento `DropDown` por caché de filtros**:
   En la Fase 1544 se introdujo una guarda de rendimiento en el evento `DropDown`:
   ```csharp
   if (cmbCargarCorridas.Items.Count > 0 && ClaveFiltroCorridas() == _claveFiltroCorridas) return;
   ```
   Esta guarda asumía erróneamente que la lista de corridas solo cambia si el usuario altera el período o programa en los controles superiores. Cuando el AG finaliza una nueva optimización dentro del mismo período y programa, la clave de filtro no varía, por lo que el despliegue del ComboBox descartaba silenciosamente la reconsulta a la base de datos SQLite y mantenía en memoria el catálogo obsoleto.

2. **Ausencia de sincronización garantizada post-ejecución del AG**:
   El refresco del combo al término de `EjecutarOptimizacionAsync` no estaba centralizado en un bloque `finally`. Si el proceso era cancelado por el usuario (`CancellationToken`), o si ocurría un fallo posterior a la inserción de la cabecera en `ag_corrida` (por ejemplo, durante la evaluación o persistencia de staging), el flujo abortaba sin refrescar el catálogo de corridas. Además, cuando la optimización culminaba exitosamente por ruta directa, la etiqueta `lblAgModo` no se actualizaba con el identificador `#N` asignado, dejando un estado visual inconsistente.

3. **Degradación forzada a `SelectedIndex = 0` por reutilización de `ResolverSeleccionInicial`**:
   El método `AgCorridaSelector.ResolverSeleccionInicial` fue concebido para `FormVinculaciones`, el cual inserta un ítem comodín (`-- Todas las corridas --`) en el índice 0. En el Módulo R no existe dicho comodín; cada ítem es una corrida física real. Al invocar `ResolverSeleccionInicial`:
   - Si la corrida activa no se encontraba en el catálogo filtrado, el método devolvía el comodín con ID 0.
   - En `AgFormGeneradorHorarios.RefrescarComboCorridasAsync`, el código ejecutaba una degradación forzada:
     ```csharp
     int idx = items.FindIndex(c => c.IdCorrida == elegida.IdCorrida);
     cmbCargarCorridas.SelectedIndex = idx >= 0 ? idx : 0;
     await CargarCorridaDesdeComboAsync();
     ```
   - Al degradar a `0`, se seleccionaba forzosamente la primera corrida del catálogo y se disparaba la recarga destructiva del staging, pisando los resultados de la corrida activa.

---

## 2. Solución Implementada

Para resolver integralmente el desfasaje sin introducir regresiones en `FormVinculaciones`, se implementó una arquitectura en capas basada en una política de dominio pura, desacoplamiento de tokens, protección de concurrencia y gestión defensiva del ciclo de vida de la corrida.

### 2.1. Política Pura `AgCorridaSelector.ResolverSeleccionModuloR` (Fases 1606–1608)
Se creó el DTO inmutable `AgSeleccionModuloR` y se implementó la función estática pura `ResolverSeleccionModuloR` en `AgCorridaSelector`:

```csharp
public sealed class AgSeleccionModuloR
{
    public CorridaStagingItem ItemSeleccionado { get; }
    public int IndiceSeleccionado { get; }
    public bool RequiereCargaInicial { get; }
    public bool CorridaActivaNoEncontrada { get; }
}
```

La política resuelve de manera determinista y sin efectos secundarios:
- **Activa presente en la lista**: Devuelve el ítem de la corrida activa con su índice real (`>= 0`) y `RequiereCargaInicial = false`. La selección se sincroniza en el combo sin disparar recargas innecesarias en las grillas.
- **Activa ausente en la lista**: Si hay una corrida activa cargada pero los filtros actuales (período/programa/subunidad) no la contienen, devuelve `ItemSeleccionado = null`, `IndiceSeleccionado = -1`, `RequiereCargaInicial = false` y `CorridaActivaNoEncontrada = true`.
- **Sin activa y con staging disponible**: Al abrir el formulario o tras un reset, si existen corridas con sesiones en staging, selecciona la primera (`IndiceSeleccionado >= 0`) y marca `RequiereCargaInicial = true` para autorizar la carga inicial automática exactamente una vez.
- **Sin activa y solo cabeceras (sin staging)**: Si solo existen corridas sin sesiones persistidas, devuelve `IndiceSeleccionado = -1` y `RequiereCargaInicial = false`, evitando estados inconsistentes.
- **Lista vacía o nula**: Devuelve `IndiceSeleccionado = -1` y `RequiereCargaInicial = false`.

El contrato previo `ResolverSeleccionInicial` y el comodín `-- Todas las corridas --` se mantuvieron completamente intactos para `FormVinculaciones`.

### 2.2. Preservación Estricta de la Corrida Activa (Fase 1608)
En `AgFormGeneradorHorarios.RefrescarComboCorridasAsync`:
- Se eliminó el fallback destructivo `idx >= 0 ? idx : 0`.
- Si `seleccion.CorridaActivaNoEncontrada == true`, se fija `cmbCargarCorridas.SelectedIndex = -1` y se actualiza `lblAgModo` con el mensaje informativo:
  ```
  Modo AG · corrida cargada #N no visible con los filtros actuales
  ```
  Las grillas de horarios, violaciones y vacantes permanecen intactas en pantalla, garantizando que un cambio en los filtros del árbol o del formulario no destruya la corrida en revisión.

### 2.3. Refresco Fresco en `DropDown` y Control de Concurrencia (Fase 1609)
- **Eliminación de la caché bloqueante**: El evento `DropDown` ahora solicita siempre una consulta fresca a la base de datos:
  ```csharp
  cmbCargarCorridas.DropDown += (s, e) => RefrescarComboCorridasSeguro();
  ```
- **Coalescencia de consultas concurrentes**: Se implementó un bucle de drenaje con la bandera `_solicitudPendienteRefrescoCombo`. Si entra una nueva petición de refresco mientras otra operación asíncrona está en vuelo (por ejemplo, ráfagas de eventos de UI o apertura rápida de controles), la segunda no se descarta; se procesa de forma secuencial con los parámetros más recientes.
- **Detección de respuestas obsoletas**: Se captura `claveFiltrosAlIniciar = ClaveFiltroCorridas()` antes de los `await`. Si los filtros cambian durante la lectura en SQLite o si el formulario entra en disposición (`IsDisposed || Disposing`), la respuesta se descarta y se reencola una nueva solicitud limpia.
- **Resiliencia ante fallos**: Si la lectura en SQLite falla, se emite una advertencia en el log y se preserva el catálogo previo sin limpiar los controles ni deseleccionar la corrida activa.

### 2.4. Sincronización Post-Ejecución y Ciclo de Vida en `finally` (Fase 1610)
- **Bloque `finally` desacoplado**: En `EjecutarOptimizacionAsync`, se registra `idCorridaCreada = idCorrida` en cuanto se inserta la cabecera en `ag_corrida`. En el bloque `finally`, si `idCorridaCreada > 0`, se invoca `RefrescarComboCorridasAsync()` de manera protegida contra excepciones y con un token independiente (`CancellationToken.None`), garantizando que la corrida aparezca en el combo aun si el AG se cancela durante las generaciones evolutivas.
- **Actualización de `lblAgModo` en ruta directa**: Tras materializar las sesiones y el individuo óptimo en `CargarModuloResultadosAsync`, se asigna `_idCorridaCargada = idCorrida` y se actualiza la etiqueta en la ruta de ejecución directa:
  ```csharp
  _idCorridaCargada = idCorrida;
  lblAgModo.Text = $"Modo AG · corrida #{idCorrida} · {AgModoDatos.Ag.VistaSesionesDetalle}";
  ```
- **Limpieza ante fallo de staging**: Si `CargarModuloResultadosAsync` no logra cargar las sesiones o la corrida queda vacía, retorna `false` y limpia todas las variables de estado (`_idCorridaCargada = 0`, métricas a cero y grillas nulas), evitando que el formulario conserve un identificador huérfano.
- **Manejo defensivo en `AbrirVinculacionesModoAgAsync`**: Tras cerrar el diálogo modal de vinculaciones, se verifica el resultado de recargar la corrida activa. Si la corrida fue eliminada o purgada dentro del diálogo, se restablece `_idCorridaCargada = 0` y se solicita un refresco limpio del combo para que el selector aplique la política de selección inicial.

---

## 3. Archivos Modificados y Creados

| Archivo | Operación | Fase / Tarea | Descripción |
|---|---|---|---|
| `ConversorPDF.XLL/Academic/Application/DTOs/AgSeleccionModuloR.cs` | **Creado** | Fase 1606 (T1) | DTO inmutable que encapsula el resultado de la política de selección del Módulo R. |
| `ConversorPDF.XLL/ConversorPDF.XLL.csproj` | **Modificado** | Fase 1606 (T1) | Registro del nuevo archivo `AgSeleccionModuloR.cs` en la compilación del proyecto XLL. |
| `ConversorPDF.Tests/Academic/Application/Services/AgCorridaSelectorTests.cs` | **Modificado** | Fase 1607 (T2) | 6 pruebas unitarias RED/GREEN para caracterizar todos los escenarios de `ResolverSeleccionModuloR`. |
| `ConversorPDF.XLL/Academic/Application/Services/AgCorridaSelector.cs` | **Modificado** | Fase 1608 (T3) | Implementación de `ResolverSeleccionModuloR` sin efectos colaterales en `FormVinculaciones`. |
| `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs` | **Modificado** | Fases 1608–1610 (T3–T5) | Enlace seguro en `RefrescarComboCorridasAsync`, eliminación de fallback `idx >= 0 ? idx : 0`, refresco fresco en `DropDown`, coalescencia de refrescos, bloque `finally` desacoplado, sincronización de `lblAgModo` y manejo defensivo en `AbrirVinculacionesModoAgAsync`. |
| `docs/articles/changes/2026-11-12-sincronizacion-selector-corridas-modulo-r.md` | **Creado** | Fase 1612 (T7) | Artículo de documentación técnica del cambio y cierre de la ronda R-041. |

---

## 4. Pruebas y Evidencias

La solución fue desarrollada bajo la metodología TDD estricta (RED -> GREEN -> REFACTOR) y validada en múltiples niveles de la suite de pruebas automatizadas:

### 4.1. Pruebas Unitarias de `AgCorridaSelector` (Fases 1607–1608)
Se ejecutaron las 24 pruebas del componente selector (`AgCorridaSelectorTests`), cubriendo los 6 casos de negocio obligatorios:
1. `ResolverSeleccionModuloR_CuandoActivaPresente_DebeSeleccionarActivaYSinCargaInicial`: Activa presente se selecciona con su índice sin forzar recarga.
2. `ResolverSeleccionModuloR_CuandoActivaAusenteConStagingDisponible_NoDebeSeleccionarNiCargarYMarcaNoEncontrada`: Activa ausente pero con corridas con staging devuelve índice `-1` y bandera `CorridaActivaNoEncontrada = true`.
3. `ResolverSeleccionModuloR_CuandoActivaAusenteConSoloCabeceras_NoDebeSeleccionarNiCargarYMarcaNoEncontrada`: Activa ausente con solo cabeceras devuelve índice `-1` y bandera `CorridaActivaNoEncontrada = true`.
4. `ResolverSeleccionModuloR_CuandoSinActivaYConStagingDisponible_DebeSeleccionarPrimeraConStagingYRequiereCargaInicial`: Sin activa previa, selecciona la primera corrida con staging y habilita carga inicial.
5. `ResolverSeleccionModuloR_CuandoSinActivaYSoloCabeceras_NoDebeSeleccionarNiCargar`: Sin activa previa y sin staging, no selecciona ni carga.
6. `ResolverSeleccionModuloR_CuandoListaVaciaONula_NoDebeSeleccionarNiCargar`: Casos borde de colecciones vacías o nulas devuelven estado neutro.
7. 18 pruebas preexistentes de formateo de etiquetas, cascada tolerante y selección inicial con comodín para `FormVinculaciones` (100% PASS).

**Nota de robustez y concurrencia (Ronda 326 / Fases 1615–1618)**:
- El refresco en UI se realiza al cerrar el desplegable (`DropDownClosed`) o con el popup cerrado, evitando mutar `Items` mientras `DroppedDown == true`.
- `CargarCorridaDesdeComboAsync` cuenta con exclusión mutua mediante `_cargandoCorridaDesdeCombo` para prevenir colisiones asíncronas concurrentes.
- La carga inicial automática está estrictamente acotada a la apertura del formulario (`permitirCargaInicial: true` en `OnShown`), sin dispararse tras errores de staging ni en refrescos manuales.
- Cuando la corrida activa queda excluida por los filtros de período/programa, se reinyecta en el catálogo como ítem fuera de filtro, asegurando que el selector nunca quede vacío y siempre represente la corrida `#N` activa.

**Resultado**: 24/24 pruebas unitarias aprobadas (0 fallos).

### 4.2. Pruebas de Regresión del Módulo AG (Fase 1610)
Verificación de todos los servicios, algoritmos genéticos y controladores del módulo académico (`FullyQualifiedName~Ag`):
- **Resultado**: 617/617 pruebas aprobadas (0 fallos).

### 4.3. Verificación Integral Full (Fase 1611)
Ejecución completa de la suite de pruebas del proyecto en modo `Full` (`scripts/verify.ps1 -Mode Full`):
- **Total de pruebas ejecutadas**: 1172
- **Pruebas superadas (Passed)**: 1172
- **Pruebas fallidas (Failed)**: 0
- **Pruebas omitidas (Skipped)**: 0
- **Regresiones vs Baseline (952 pruebas)**: 0
- **Veredicto**: `PASS`
- **Evidencia registrada**: `_evidence/20261112-074457/summary.md` y `_evidence/20261112-074457/test.trx`.