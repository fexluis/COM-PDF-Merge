# Renombrado de Modalidades de Tutoría: Presencial / Plataforma / Mixta

**Fecha:** 2026-10-21  
**Módulo:** Académico (`ConversorPDF.XLL.Academic`)  
**Ronda:** 148 (Fases 783 a 791)

## Propósito
Renombrar las etiquetas visibles en la interfaz de usuario y los valores canónicos persistidos de **Tipo Tutoría** en toda la solución:
- La etiqueta y valor **`Virtual`** pasa a ser **`Plataforma`**.
- La etiqueta y valor **`Semipresencial`** pasa a ser **`Mixta`** (sustituyendo integralmente a *Semipresencial* y absorbiendo variantes legadas como *A Distancia* / *Distancia*).
- La modalidad **`Presencial`** se mantiene intacta sin cambios.

El cambio unifica la terminología institucional en la UI, formularios de gestión académica (`FormSesionesAcademicas`, `FormVerVinculaciones`), árbol de configuración del Generador de Horarios (`AgFormGeneradorHorarios`), vistas, mappers, servicios de cálculo de viáticos y migraciones de base de datos SQLite (V66).

## Cambios
1. **Catálogo Canónico (`ModalidadTutoriaCatalog.cs`, Fase 783):**
   - Nuevo catálogo central en `Academic.Domain.Models` para evitar dependencias circulares entre capas.
   - Constantes canónicas `Presencial = "Presencial"`, `Plataforma = "Plataforma"`, `Mixta = "Mixta"`.
   - Propiedades `Opciones` (`string[]`), `Pares` (`IReadOnlyList<(string Codigo, string Etiqueta)>`).
   - Métodos utilitarios: `Normalizar(modalidad)` con orden de evaluación estricto (Plataforma/Virtual → Mixta/Semi/Distancia → Presencial → Trim/Fallback), `AEtiqueta(codigo)`, `EsPresencial(modalidad)`, `EsPlataforma(modalidad)`, `EsMixta(modalidad)`.

2. **Nueva Migración V66 Aditiva y Tolerante (`MigrationV66_RenombrarModalidadTutoria.cs`, Fase 784):**
   - `Version = 66`, `Name = "RenombrarModalidadTutoria"`, `Aditiva = true`.
   - Registrada en `SqliteHelper.BuildMigrator()` posterior a V65.
   - Actualización set-based e idempotente de columnas de modalidad en 4 tablas:
     - `vinculaciones.modalidadVinculacion`
     - `ag_vinculaciones.modalidadVinculacion`
     - `ag_aula.modalidadVinculacion`
     - `ag_arbol_nodo.modalidad`
   - Actualización de identificadores derivados en `ag_aula`:
     - Reemplazo seguro de tokens en `aula_ref`, `codigo`, `nombre`: `|VIRTUAL` → `|PLATAFORMA`, `|SEMIPRESENCIAL` → `|MIXTA`, `|A DISTANCIA` → `|MIXTA`.
   - Protección contra colisión en `uq_ag_arbol_nodo(idarbol, aula_ref)`:
     - Deduplicación explícita previa eliminando filas legadas duplicadas en el mismo perfil antes de ejecutar el `UPDATE`.
   - Preservación explícita fuera de alcance: `TipoAula` (`VIRTUAL`/`ESTANDAR`), `EsVirtual`/`es_virtual`, `programas.modalidad` (`PREGRADO`/`POSGRADO`), festivos `PRESENCIAL`.

3. **Dominio y DTOs (`AgAula.cs`, `AgVinculacion.cs`, `AgContextoSeleccionDto.cs`, Fase 785):**
   - Valores por defecto inicializados con `ModalidadTutoriaCatalog.Presencial`.
   - Normalización canónica en constructores de `AgContextoSeleccionDto`.
   - Propiedad `ModalidadEfectiva => EsVirtual > 0 ? ModalidadTutoriaCatalog.Plataforma : ModalidadTutoriaCatalog.Normalizar(Modalidad)`.

4. **Factorías y Mappers (Fase 786):**
   - `AgAulaFactory.ConstruirAulaRef`: normalización previa a `ToUpperInvariant()`.
   - `AgAulaFactory.CrearDesdeDto`: asignación `modVinculacion = esVirtual == 1 ? Plataforma : mod`.
   - `AgArbolMapper`: normalización bidireccional entre nodos del árbol y DTOs de contexto.
   - `AgExpansorSesiones`, `AgAulaToAgVinculacionMapper`, `AgVinculacionToVinculacionMapper`, `AgGenToAgVinculacionMapper`, `AgEncabezadoVinculacionFactory`: mapeo y normalización con fallback canónico.

5. **Servicios de Aplicación (Fase 787):**
   - `CalculadoraViaticaSesion`: evaluación con `ModalidadTutoriaCatalog.EsPlataforma`, `EsPresencial`, `EsMixta`. Semántica idéntica (Plataforma nunca viática; Presencial viática con desplazamiento; Mixta viática exclusivamente dentro de rangos presenciales).
   - `SesionConflictAnalyzer`: omisión de sesiones de modalidad Plataforma mediante `ModalidadTutoriaCatalog.EsPlataforma(s.Modalidad)`.

6. **Interfaz de Usuario (Fases 788 y 789):**
   - `FormSesionesAcademicas.Designer.cs`: controles renombrados `cmbTipoJornada` → `cmbTipoTutoria`, `lblTipoJornada` → `lblTipoTutoria`.
   - `FormSesionesAcademicas.cs`: carga de opciones `Presencial`, `Plataforma`, `Mixta`; persistencia y recuperación de estado; mensaje de advertencia actualizado explícitamente: *"Con la modalidad Mixta, las sesiones se generarán sin viáticas asignadas..."*.
   - `FormVerVinculaciones.cs`: uso de `ModalidadTutoriaCatalog.EsPlataforma`, distintivo visual `" | 🌐 Plataforma"`, exclusión de conflictos para Plataforma.
   - `AgFormGeneradorHorarios.cs`: nivel 5 (modalidad) poblado mediante pares `(par.Codigo, par.Etiqueta)` (`key6 = new AgNodoArbolKey(6, 0, par.Codigo, par.Codigo)`, `new TreeNode(par.Etiqueta)`); mensajes de error de aulas faltantes con etiqueta canónica `[Plataforma]` / `[Mixta]`; compatibilidad y normalización en carga de perfiles guardados.

7. **Pruebas y Documentación (Fases 790 y 791):**
   - Creación de `ModalidadTutoriaCatalogTests.cs` y `MigrationV66Tests.cs`.
   - Actualización de suites existentes: `CalculadoraViaticaSesionTests`, `AgAulaFactoryTests`, `AgArbolServiceTests`, `AgArbolMapperTests`, `AgAulaToAgVinculacionMapperTests`, `AgTreeViewCheckServiceTests`.
   - Actualización de especificación en `AG-2.txt` (fila `cmbTipoTutoria`).

## Cómo usar
1. Al iniciar la aplicación, `SchemaMigrator` ejecuta automáticamente la migración V66 actualizando los datos existentes.
2. En **Sesiones Académicas**:
   - El desplegable **Tipo Tutoría** presenta tres opciones: **Presencial**, **Plataforma** y **Mixta** (por defecto *Presencial*).
   - Al seleccionar *Plataforma*, las sesiones se generan sin viáticos ni restricciones de franja física.
   - Al seleccionar *Mixta*, las sesiones generan viáticos únicamente si caen dentro de un rango presencial configurado.
   - Vinculaciones históricas marcadas como *Virtual* se presentan automáticamente como *Plataforma*; vinculaciones marcadas como *Semipresencial* o *A Distancia* se presentan como *Mixta*.
3. En el **Generador de Horarios (AG)**:
   - Los nodos de nivel de modalidad en el árbol muestran: `Presencial`, `Plataforma`, `Mixta`.
   - Al generar aulas, se generan claves canónicas con sufijos `|PLATAFORMA` y `|MIXTA`.
   - Si faltan aulas por configurar, el diálogo reportará claramente los nombres canónicos entre corchetes (ej. `... [Plataforma]`, `... [Mixta]`).

## Cobertura de pruebas (Fase 790)
- `ModalidadTutoriaCatalogTests.cs`:
  - `Opciones_ShouldContainCanonicalValuesInCorrectOrder`
  - `Pares_ShouldMatchCodeAndLabel`
  - `Normalizar_WhenGivenVariedInputs_ShouldMapCorrectly`
  - `AEtiqueta_WhenGivenInputs_ShouldReturnExpectedDisplayLabel`
  - `Predicates_ShouldIdentifyModalitiesAccurately`
- `MigrationV66Tests.cs`:
  - `MigrationV66_Metadata_ShouldBeCorrect`
  - `Up_WhenTablesMissing_ShouldNotThrowAndLogInfo`
  - `Up_WhenLegacyDataPresent_ShouldNormalizeToPlataformaAndMixta`
  - `Up_WhenCollisionInArbolNodoUniqueIndex_ShouldDeduplicateAndSucceed`
  - `Up_WhenExecutedMultipleTimes_ShouldBeIdempotent`
- Suites de regresión y compatibilidad actualizadas:
  - `CalculadoraViaticaSesionTests.cs` (casos canónicos y compatibilidad legada)
  - `AgAulaFactoryTests.cs` (generación y formateo con Plataforma y Mixta)
  - `AgArbolServiceTests.cs` (mapeo y carga de perfiles)
  - `AgArbolMapperTests.cs` (mapeo bidireccional DTO/Nodo)
  - `AgAulaToAgVinculacionMapperTests.cs` (mapeo aula a vinculación)
  - `AgTreeViewCheckServiceTests.cs` (expansión cartesiana de combinaciones)
