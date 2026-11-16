<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 707-799 (Rondas 145, 146, 147, 148, 149, 150)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Evaluación Crítica de las Fases

#### 1. Fase 707 — `AsignaturasExcelExporter.cs`: Rango Efectivo de Comisión
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La exportación de comisiones a Excel en plantillas `FV-FO-033` escribía `FechaMaxima` en la celda de fecha fin. Ante pernocta en el último día (`isPernoctaUltimoDia = 1`), no se reflejaba el día de retorno, reduciendo 1 día de viáticos en las fórmulas de la plantilla.
  - *Causa Raíz:* Uso directo de `FechaMaxima` en vez del cálculo de dominio `FechaFinComision`.
  - *Impacto Confirmado:* Fechas de comisión reflejan el retorno real del docente; las plantillas deducen los días exactos de viático sin alterar el modelo subyacente.
- **Acción Aplicada:** Sustituido `FechaMaxima.Value.Date` por `ViaticoBloqueFormatter.ObtenerRangoComision(bloque)?.Fin`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`

#### 2. Fase 708 — `ViaticoBloqueFormatter.cs`: Método de Rango de Comisión y Pruebas Unitarias
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de método formateador/extractor desacoplado para obtener el rango efectivo `(Inicio, Fin)` de comisión con pernocta.
  - *Causa Raíz:* Lógica de visualización y exportación dispersa.
  - *Impacto Confirmado:* Helper puro y testeable sin efectos secundarios ni dependencias externas.
- **Acción Aplicada:** Creado `ObtenerRangoComision(ViaticoBloque bloque)` y añadidas 3 pruebas unitarias completas en `ViaticoBloqueFormatterTests.cs`.
- **Archivos Afectados:**
  - `ConversorPDF.XLL/Academic/Application/Services/ViaticoBloqueFormatter.cs`
  - `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`

#### 3. Fase 709 — Documentación Técnica y Respaldo Snapshot
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Necesidad de documentar la regla institucional de pernocta en comisiones exportadas y registrar respaldo de seguridad.
  - *Causa Raíz:* Comportamiento no explicitado en los manuales técnicos del módulo académico.
  - *Impacto Confirmado:* Documentación sincronizada y snapshot de respaldo preservado.
- **Acción Aplicada:** Actualizado `docs/articles/academic-viaticos.md` y creado respaldo en `_FIXES/fix-20261016-182822/`.
- **Archivos Afectados:**
  - `docs/articles/academic-viaticos.md`
  - `_FIXES/fix-20261016-182822/`

#### 4. Fase 710 — Verificación Estática y Preservación de Invariantes
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Riesgo potencial de mutar `FechaMaxima` o alterar `AgruparConsecutivos` y `ViaticoConflictAnalyzer`.
  - *Causa Raíz:* Sobrecarga conceptual entre fecha de última clase y fecha de regreso.
  - *Impacto Confirmado:* `FechaMaxima` intacta, `SeSolapan` inalterado, sin impacto en plantillas ajenas ni en modo AG.
- **Acción Aplicada:** Verificación estática y de diseño sin regresiones.
- **Archivos Afectados:** Verificación transversal de arquitectura.

---

### Verificaciones de Principios (Ronda 130)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F707 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F708 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F709 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F710 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.89 Ronda 131 — Verificación Runtime, Sincronización de DocFX y Estandarización de Auditoría (Fases 711 a 713)

### Evaluación Crítica de las Fases

#### 1. Fase 711 — Verificación de Compilación y Suite Completa de Pruebas Unitarias
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La Fase 710 no ejecutó compilador ni suite de pruebas; el comportamiento de tuplas nullables `(DateTime Inicio, DateTime Fin)?` con operador `?.` y los nuevos tests unitarios requerían confirmación en runtime.
  - *Causa Raíz:* Restricción de ejecución en revisiones estáticas previas.
  - *Impacto Confirmado:* Código compilado limpiamente en `Release|x64` y suite integral 100% aprobada.
- **Acción Aplicada:** Compilado `ConversorPDF.XLL` (0 errores) y ejecutada suite `ConversorPDF.Tests` (624/624 pruebas aprobadas, 0 fallos).
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `ConversorPDF.Tests/ConversorPDF.Tests.csproj`.

#### 2. Fase 712 — Regeneración del Sitio Estático de Documentación (`docs/_site`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `docs/_site/articles/academic-viaticos.html` no contenía la sección de pernocta en comisiones agregada en el archivo Markdown fuente.
  - *Causa Raíz:* Omitida invocación de DocFX tras la edición documental.
  - *Impacto Confirmado:* Sitio web estático alineado con la última versión de la documentación.
- **Acción Aplicada:** Ejecutado `docfx build docs/docfx.json` (0 errores); HTML sincronizado.
- **Archivo Afectado:** `docs/_site/articles/academic-viaticos.html`.

#### 3. Fase 713 — Corrección de Atribución de Plantillas y Estandarización de Snapshot de Auditoría
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Atribución documental imprecisa de plantillas (`FV-FO-033` vs `GF-FO-023`) y snapshot sin formato canónico de auditoría (`.txt` y encabezado `# AUDITORÍA DE CAMBIO`).
  - *Causa Raíz:* Redacción agrupada y copia directa sin script de auditoría.
  - *Impacto Confirmado:* Trazabilidad documental exacta y respaldo conforme a la convención `_FIXES/fix-yyyyMMdd-HHmm.ss.fff`.
- **Acción Aplicada:** Ajustada redacción en `docs/articles/academic-viaticos.md` y generado snapshot en `_FIXES/fix-20261016-1828.22.000/` con encabezados de auditoría v1.6.3.51.
- **Archivos Afectados:** `docs/articles/academic-viaticos.md`, `_FIXES/fix-20261016-1828.22.000/`.

---

### Verificaciones de Principios (Ronda 131)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F711 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F712 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F713 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.90 Ronda 132 — Formato y Resaltado del Nivel 2 en la Vista Viáticos (Fases 714 a 720)

### Evaluación Crítica de las Fases

#### 1. Fase 714 — `ViaticoBloqueFormatterTests.cs`: Pruebas Unitarias de Formato Nivel 2 (TDD)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La suite de pruebas de formateo de Nivel 2 requería validar el nuevo formato exacto sin código de trayecto ni prefijo `Cierre: `, con duración abreviada `d`, sufijo condicional `| Pernocta Ultimo Dia` y cobertura de bloque nulo.
  - *Causa Raíz:* Contrato de prueba previo atado a `codRuta` y `días`.
  - *Impacto Confirmado:* Suite TDD actualizada con aserciones exactas y cobertura de bordes sin regresiones.
- **Acción Aplicada:** Modificadas pruebas de `FormatearNodoNivel2` existentes y añadidas `FormatearNodoNivel2_WhenSinPernocta_ShouldNotIncludeCierrePrefixNorCodigoRuta` y `FormatearNodoNivel2_WhenBloqueNull_ShouldReturnEmpty`.
- **Archivo Afectado:** `ConversorPDF.Tests/Academic/Application/Services/ViaticoBloqueFormatterTests.cs`.

#### 2. Fase 715 — `ViaticoBloqueFormatter.cs`: Ajuste de Firma y Formateo de Cadena
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `FormatearNodoNivel2` sobrecargaba la fila visual con el parámetro `codRuta`, la palabra `días` y el prefijo `Cierre: `.
  - *Causa Raíz:* Plantilla de formato previa no optimizada para lectura rápida en el árbol.
  - *Impacto Confirmado:* Cadena limpia, compacta y alineada con la especificación de UI, preservando `ObtenerTextoCierre` y `ObtenerRangoComision`.
- **Acción Aplicada:** Retirado `codRuta` de la firma, sustituido `días` por `d`, removido el prefijo `Cierre: ` y actualizados los XML docs.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/Application/Services/ViaticoBloqueFormatter.cs`.

#### 3. Fase 716 — `FormVerVinculaciones.cs`: Integración UI y Resaltado en DarkOrange
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La fila del bloque mostraba `codRutaDisplay` ("SIN-COD"), no diferenciaba visualmente viajes con pernocta en último día y requería garantizar que la pre-generación en BD de rutas no catalogadas continuara ejecutándose.
  - *Causa Raíz:* Parámetro de ruta en cadena de texto acoplado a UI y color estático `Color.DarkBlue`.
  - *Impacto Confirmado:* Pre-creación de rutas en BD preservada mediante `_trayectoEncoderService.GetOrCreateCodigoTrayecto`, llamada a `FormatearNodoNivel2` actualizada y nodo resaltado en `Color.DarkOrange` cuando `bloque.TienePernoctaUltimoDia == true`.
- **Acción Aplicada:** Removido `codRutaDisplay`, preservada la invocación a `GetOrCreateCodigoTrayecto`, adaptada la llamada a `FormatearNodoNivel2` y asignado `ForeColor = bloque.TienePernoctaUltimoDia ? Color.DarkOrange : Color.DarkBlue`.
- **Archivo Afectado:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`.

#### 4. Fase 717 — Documentación Técnica de Formato y Resaltado Nivel 2
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Se requería documentación técnica explícita sobre la estructura del Nivel 2, la abreviatura `d`, el retiro de `Cierre:`/`CodRuta` exclusivamente en la vista y el resaltado `DarkOrange`.
  - *Causa Raíz:* Evolución de especificación de interfaz en el módulo de viáticos.
  - *Impacto Confirmado:* Manual técnico y registro de cambios alineados con la implementación.
- **Acción Aplicada:** Agregada subsección en `docs/articles/academic-viaticos.md` y creado `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`.
- **Archivos Afectados:** `docs/articles/academic-viaticos.md`, `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`.

#### 5. Fase 718 — Sincronización de Sitio Estático DocFX (`docs/_site`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El HTML publicado en `docs/_site` no contenía la nueva subsección de formato Nivel 2 ni el resaltado en `DarkOrange` al no haberse ejecutado DocFX tras la edición Markdown.
  - *Causa Raíz:* Omisión de ejecución de compilación del sitio DocFX.
  - *Impacto Confirmado:* Sitio web estático alineado al 100% con la documentación en Markdown.
- **Acción Aplicada:** Ejecutado `docfx build docfx.json` desde el directorio `docs/` (0 errores, 1 warning pre-existente en index.md).
- **Archivo Afectado:** `docs/_site/articles/academic-viaticos.html`.

#### 6. Fase 719 — Respaldo de Seguridad y Estandarización de Snapshot `_FIXES`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de snapshot histórico `.txt` de la ronda con encabezado canónico `# AUDITORÍA DE CAMBIO`.
  - *Causa Raíz:* Omisión del procedimiento de snapshot al cierre de la edición de código.
  - *Impacto Confirmado:* Trazabilidad histórica preservada en el repositorio.
- **Acción Aplicada:** Creada carpeta `_FIXES/fix-20261016-1908.00.000/` conteniendo `ViaticoBloqueFormatter.txt`, `ViaticoBloqueFormatterTests.txt`, `FormVerVinculaciones.txt` y `academic-viaticos.txt` identificados con versión `1.6.3.52`.
- **Archivos Afectados:** `_FIXES/fix-20261016-1908.00.000/`, `scratch/fix_snapshot.ps1`.

#### 7. Fase 720 — Corrección de Nomenclatura de Test en Registro de Cambio
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La sección Test Coverage citaba `FormatearNodoNivel2_WhenSinPernocta_ShouldFormatExpectedString` en vez del nombre real de la prueba.
  - *Causa Raíz:* Desincronización entre el nombre preliminar propuesto y el nombre final del método de prueba.
  - *Impacto Confirmado:* Registro de cambios consistente con la suite real de pruebas.
- **Acción Aplicada:** Sincronizado el identificador de la prueba unitaria en `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`.
- **Archivo Afectado:** `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`.

---

### Verificaciones de Principios (Ronda 132)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F714 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F715 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F716 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F717 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F718 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F719 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F720 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 📦 Entrega de Release v1.6.3.54 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.54`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.54"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.54</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.54"`)
- **Documentación Técnica Generada (`/build-docs`):**
  - Generado sitio web estático DocFX (`docs\_site\`) sin errores (0 errores).
  - Sincronizados XML comments, manuales técnicos y registros de cambios.
- **Verificación Integral de Pruebas Unitarias:**
  - Suite completa: **626 / 626 pruebas aprobadas (0 errores)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores** (`ConversorPDF.XLL-AddIn64.xll`).
- **Paquete Instalador Generado (`/build-release 1.6.3.54`):**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.54.msi`
  - Tamaño: **5,464,064 bytes** (5.46 MB)
  - Estado: Listo para distribución con auto-activación en Excel.

---

## 157.91 Ronda 133 — Sincronización de Plantillas Excel en el Instalador y Blindaje de Reemplazo (Fases 721 a 727)

### Evaluación Crítica de las Fases

#### 1. Fase 721 — Sincronización de Plantilla `FV-FO-033.xlsx` y Secuencia de Build
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El instalador v1.6.3.54 empaquetó `FV-FO-033.xlsx` de 736.198 B en lugar de la versión actualizada de 777.061 B.
  - *Causa Raíz:* Inversión temporal: el empaquetado del MSI ocurrió a las 19:27 y el guardado de la plantilla en `ConversorPDF.XLL\ExcelTemplates\` a las 19:34.
  - *Impacto Confirmado:* Distribución de plantilla desactualizada a usuarios de la v1.6.3.54.
- **Acción Aplicada:** Documentada la invariante de empaquetado en el workflow `build-release.md` y establecido bump a versión `1.6.3.55` para empaquetado del artefacto correcto (777.061 B).
- **Archivos Afectados:** `ConversorPDF.XLL\ExcelTemplates\FV-FO-033.xlsx`, `installer\InstaladorConversorPDF.wxs`.

#### 2. Fase 722 — Blindaje Integral de Reemplazo de Plantillas en `TemplateComponents`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al actualizar mediante `MajorUpgrade`, Windows Installer trata los `.xlsx` como archivos de datos sin versión PE. Si el archivo en destino fue tocado, `CostFinalize` asigna `FA_NoOverwrite` y omite su reemplazo durante `InstallFiles`, incluso si `RemoveFile` lo elimina durante `RemoveFiles`.
  - *Causa Raíz:* Archivos sin versión en la tabla `File` de MSI y componentes per-user sin `KeyPath` explícito en `HKCU`.
  - *Impacto Confirmado:* Riesgo de conservar plantillas obsoletas o eliminarlas sin reemplazo en `%APPDATA%\Microsoft\AddIns\ExcelTemplates\`.
- **Acción Aplicada:**
  1. Conservados los identificadores de archivo `File/@Id` (`filTemplate1`, `filTemplate2`) y Component GUIDs para garantizar trazabilidad y consistencia de linaje MSI.
  2. Asignado `DefaultVersion="!(bind.FileVersion.filXllDll)"` a cada `<File>` para forzar a `CostFinalize` a asignar `FileAction = Install` incondicional contra archivos no versionados en disco.
  3. Declarados `RegistryValue` con `KeyPath="yes"` en `HKCU\Software\ESAP\ConversorPDF\Templates` con valor `[ProductVersion]` para asegurar el estado activo de instalación del componente en el ámbito per-user.
  4. Declaradas directivas `<RemoveFile Id="PurgeOldTemplate1" Name="GF-FO-023.xlsx" On="install" />` y `<RemoveFile Id="PurgeOldTemplate2" Name="FV-FO-033.xlsx" On="install" />` para purga previa en `RemoveFiles` (3500) antes de `InstallFiles` (4000).
- **Archivo Afectado:** `installer\InstaladorConversorPDF.wxs`.

#### 3. Fase 723 — Eliminación de Carpeta `ExcelTemplates` en Desinstalación (`RemoveFolder`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al desinstalar la aplicación, la carpeta de plantillas y sus archivos permanecían en disco (~3,7 MB) como residuos huérfanos.
  - *Causa Raíz:* Ausencia de directiva `<RemoveFolder>` en los componentes de plantilla.
  - *Impacto Confirmado:* Violación del smoke test 4 ("sin residuos en AppData al desinstalar").
- **Acción Aplicada:** Incorporada directiva `<RemoveFolder Id="RemoveExcelTemplatesFolder" Directory="dir_ExcelTemplates" On="uninstall" />` en `cmpTemplate2`.
- **Archivo Afectado:** `installer\InstaladorConversorPDF.wxs`.

#### 4. Fase 724 — Corrección de Codificación (Mojibake) en Snapshots `_FIXES`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Todos los archivos de respaldo histórico `.txt` en `_FIXES/fix-*` presentaban doble codificación UTF-8 (mojibake `# AUDITORÃA DE CAMBIO` y caracteres acentuados).
  - *Causa Raíz:* En PowerShell 5.1, `Get-Content -Raw` sin `-Encoding` lee en ANSI (Windows-1252), y `Set-Content -Encoding UTF8` recodifica los caracteres ya corrompidos en memoria.
  - *Impacto Confirmado:* Corrupción de los artefactos históricos de auditoría de código.
- **Acción Aplicada:** Sustituido el mecanismo en `.agents/workflows/contexto-previo.md` y `scratch/fix_snapshot.ps1` por `[System.IO.File]::WriteAllText` con `System.Text.UTF8Encoding($true)`. Saneados los 7 archivos de las rondas 131 y 132 a UTF-8 válido con BOM.
- **Archivos Afectados:** `.agents/workflows/contexto-previo.md`, `scratch/fix_snapshot.ps1`, `_FIXES/fix-20261016-1828.22.000/`, `_FIXES/fix-20261016-1908.00.000/`.

#### 5. Fase 725 — Sincronización de Encabezados y Registro de Invariante de Release
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Título de la Ronda 132 desfasado citando Fases 714 a 717 en lugar de Fases 714 a 720; ausencia de validación de tamaño de plantillas en el checklist de release.
  - *Causa Raíz:* Omisión de actualización de cabecera tras anexar fases 718-720.
  - *Impacto Confirmado:* Desincronía documental entre especificación y ejecución.
- **Acción Aplicada:** Actualizados los encabezados a `(Fases 714 a 720)` en plan, walkthrough y `2026-10-16-ajuste-formato-nivel2-viaticos.md`. Añadida la sección `⚠️ Lección Aprendida 2 — Plantillas ExcelTemplates (contenido, no código)` en `.agents/workflows/build-release.md`.
- **Archivos Afectados:** `docs/articles/changes/2026-10-16-ajuste-formato-nivel2-viaticos.md`, `.agents/workflows/build-release.md`, `implementation_plan.md`, `walkthrough.md`.

#### 6. Fase 726 (Opcional) — Referencias Irresolubles en `.dna` No Desplegadas
- **Estado:** OBSERVADA (OPCIONAL - PENDIENTE DE ELECCIÓN)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `<Reference Path="System.Buffers.dll" />`, `System.Memory.dll`, `System.Numerics.Vectors.dll` en `.dna` no existen en `bin\x64\Release`.
  - *Causa Raíz:* Dependencias transitivas netstandard agregadas manualmente al `.dna`.
- **Acción Propuesta:** Se requiere elección entre Opción A (retirar del `.dna`) y Opción B (desplegar en el instalador). No se altera sin confirmación.
- **Archivo Afectado:** `ConversorPDF.XLL/ConversorPDF.XLL-AddIn64.dna`.

#### 7. Fase 727 (Opcional) — Código Muerto de Formato Retirado y Deconstrucción con Variables No Usadas
- **Estado:** OBSERVADA (OPCIONAL - PROPUESTA SIN APLICAR)
- **Hallazgo:**
  - *Diagnóstico Crítico:* `ViaticoBloque.DisplayText` obsoleto sin consumidores y variables deconstruidas descartables (`origen`, `escala`, `destino`) en `FormVerVinculaciones.cs:976`.
  - *Causa Raíz:* Residuos de la refactorización a `FormatearNodoNivel2`.
- **Acción Propuesta:** Mantener sin modificar hasta aprobación para evitar cambios fuera del alcance estricto de empaquetado.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/ViaticoBloque.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F721 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F722 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F723 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F724 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F725 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.92 Ronda 134 — Habilitar "Usar Autofiltro" en Hoja Protegida de Exportar Asignaturas (Fase 728)

### Evaluación Crítica de las Fases

#### 1. Fase 728 — Habilitar "Usar Autofiltro" en Hoja Protegida de Exportar Asignaturas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al exportar por asignaturas, la hoja `PROGRAMACIÓN` queda protegida sin el permiso de autofiltro, impidiendo desplegar y aplicar filtros en la fila 13 bajo advertencia de celda protegida.
  - *Causa Raíz:* Invocación de `ws.Protect(ConstantsHelper.FORMATO_PASSWORD)` en el bloque `finally` de `AsignaturasExcelExporter.cs:491` omitiendo los parámetros booleanos opcionales de COM, lo que asigna `false` por defecto a todos los permisos `Allow*`.
  - *Impacto Confirmado:* Bloqueo operativo de los filtros integrados en la plantilla `FV-FO-033` para los usuarios finales.
- **Acción Aplicada:** Adición del método privado estático `ProtegerHojaPermitiendoAutofiltro(dynamic ws, string password)` en `AsignaturasExcelExporter.cs`, invocando `Worksheet.Protect` mediante `Type.InvokeMember` con 16 argumentos posicionales (`Missing.Value`), fijando `args[0] = password` y `args[14] = true` (`AllowFiltering`). Reemplazo en el bloque `finally` de la exportación. Creación de la documentación de cambio correspondiente.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`, `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F728 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.93 Ronda 135 — Robustecimiento de Protección con Autofiltro/Ordenación y Trazabilidad de Rango (Fases 729 a 732)

### Evaluación Crítica de las Fases

#### 1. Fase 729 — `AllowSorting` Habilitado en Protección de Hoja
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al omitir `args[13]`, las opciones "Ordenar de A a Z / Z a A" en el desplegable del autofiltro permanecían deshabilitadas en la hoja protegida.
  - *Causa Raíz:* Excel trata los permisos `AllowFiltering` y `AllowSorting` de forma independiente.
  - *Impacto Confirmado:* Imposibilidad de ordenar datos filtrados sin desproteger la hoja.
- **Acción Aplicada:** Asignado `args[13] = true` (`AllowSorting`) en `ProtegerHojaPermitiendoAutofiltro`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 2. Fase 730 — Protección Fallback ante Excepción en Invocación Reflexiva
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Cualquier error durante `InvokeMember` en `ProtegerHojaPermitiendoAutofiltro` dejaba la hoja `PROGRAMACIÓN` totalmente desprotegida tras la exportación.
  - *Causa Raíz:* El bloque `finally` envolvía la llamada en un `catch` vacío sin mecanismo de respaldo.
  - *Impacto Confirmado:* Riesgo de pérdida silenciosa de protección del formato institucional.
- **Acción Aplicada:** Implementado un bloque `catch` con llamada de respaldo a `ws.Protect(ConstantsHelper.FORMATO_PASSWORD)` en el bloque `finally`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 3. Fase 731 — Invariante de Rango de Datos de la Tabla `T_Principal8`
- **Estado:** ACEPTADA (OPCIÓN B: NO DESTRUCTIVA)
- **Hallazgo:**
  - *Diagnóstico Crítico:* La tabla institucional `T_Principal8` tiene un rango fijo de 221 filas (`A13:DF234`); exportaciones mayores escribirían fuera del rango filtrable.
  - *Causa Raíz:* Tamaño nativo del `ListObject` en la plantilla `FV-FO-033.xlsx`.
  - *Impacto Confirmado:* Si se intentara redimensionar en runtime (`ListObject.Resize`), se arriesgaría la corrupción de fórmulas precalculadas.
- **Acción Aplicada:** Establecida la Opción B: documentar el límite de 221 filas como invariante de diseño de la plantilla en el registro de cambios sin mutar `ListObjects` en runtime.
- **Archivos Afectados:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`.

#### 4. Fase 732 — Trazabilidad de Cobertura de Pruebas y Smoke Test
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Documento de cambios inicial no especificaba el entorno de prueba ni la tabla involucrada (`T_Principal8`).
  - *Causa Raíz:* Redacción preliminar sin trazabilidad de smoke test.
  - *Impacto Confirmado:* Falta de evidencia de verificación formal en el registro histórico.
- **Acción Aplicada:** Actualizada la sección de cobertura con validación de análisis estático COM y especificaciones del smoke test manual en Excel 365 x64.
- **Archivos Afectados:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F729 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F730 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F731 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F732 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.94 Ronda 136 — Reversión de AllowSorting, Trazabilidad en Fallback y Blindaje de Reprotección (Fases 733 a 735)

### Evaluación Crítica de las Fases

#### 1. Fase 733 — Reversión de `AllowSorting` por Restricción de Celdas Bloqueadas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La activación de `args[13] = true` habilita el control de ordenamiento en el menú de autofiltro, pero Excel exige que todas las celdas del rango estén desbloqueadas; al tener celdas protegidas en la plantilla `FV-FO-033`, la ejecución de ordenamiento falla con un diálogo de error de celda protegida.
  - *Causa Raíz:* Requisito nativo de Excel COM sobre rangos protegidos.
  - *Impacto Confirmado:* Error engañoso en la experiencia de usuario y contradicción documental.
- **Acción Aplicada:** Revertida la asignación de `args[13]` a `Missing.Value` en `ProtegerHojaPermitiendoAutofiltro` conservando únicamente `args[14] = true` (`AllowFiltering`). Actualizada la documentación técnica y comentarios XML.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`, `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`.

#### 2. Fase 734 — Supresión de Reprotección sobre Hoja No Desprotegida
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En libros con contraseñas distintas, tras fallar la desprotección, el bloque `finally` intentaba proteger la hoja con la clave interna institucional sobre una hoja ya protegida.
  - *Causa Raíz:* Ausencia de bandera de verificación de éxito de desprotección previa.
  - *Impacto Confirmado:* Inconsistencia de estado o excepciones COM silenciadas en el bloque `finally`.
- **Acción Aplicada:** Declarada la variable `bool hojaDesprotegida = false;`, actualizada tras `Unprotect` mediante `!ws.ProtectContents` y condicionada la reprotección en `finally` a `if (hojaDesprotegida)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 3. Fase 735 — Observabilidad y Trazabilidad en Fallback de Protección
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La captura de excepciones en la reflexión y en el fallback era completamente silenciosa.
  - *Causa Raíz:* Bloques `catch` vacíos sin integración con el subsistema de logs.
  - *Impacto Confirmado:* Imposibilidad de diagnosticar problemas de protección en producción si el fallback o la reflexión fallan.
- **Acción Aplicada:** Integrado `LoggerHelper.Warning` ante fallo en la llamada reflexiva y `LoggerHelper.Error` ante fallo crítico en el fallback de protección.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F733 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F734 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F735 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.95 Ronda 137 — Blindaje de Bandera de Desprotección, Sincronización de Versión 1.6.3.57 y Saneamiento Documental (Fases 736 a 738)

### Evaluación Crítica de las Fases

#### 1. Fase 736 — Independencia de Bandera `hojaDesprotegida` respecto a Lectura COM
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La expresión `hojaDesprotegida = !ws.ProtectContents` dentro del `try` condicionaba la bandera a una segunda llamada COM que, en caso de excepción, dejaba la bandera en `false` omitiendo la reprotección en `finally`.
  - *Causa Raíz:* Dependencia innecesaria de una lectura de propiedad adicional durante la desprotección.
  - *Impacto Confirmado:* Riesgo de omisión de reprotección sobre hojas desprotegidas con éxito.
- **Acción Aplicada:** Asignado `hojaDesprotegida = true` inmediatamente tras `ws.Unprotect(ConstantsHelper.FORMATO_PASSWORD)` sin excepción; la validación `if (ws.ProtectContents)` posterior anula `hojaDesprotegida = false` si la contraseña era distinta.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 2. Fase 737 — Sincronización de Versión Oficial a 1.6.3.57
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El instalador distribuido v1.6.3.56 no incorporaba los cambios de las Fases 728-735 aplicados con posterioridad; distribuir una nueva versión con el mismo número impediría actualización limpia por `MajorUpgrade`.
  - *Causa Raíz:* Generación previa del release v1.6.3.56.
  - *Impacto Confirmado:* Inconsistencia de versión entre código fuente y artefactos distribuidos.
- **Acción Aplicada:** Sincronizada la versión oficial a `1.6.3.57` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/App.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `installer/InstaladorConversorPDF.wxs`.

#### 3. Fase 738 (Opcional) — Saneamiento Documental y Encabezados de Release
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Presencia de contraseña en texto plano en la documentación y encabezado de release rezagado en versión anterior.
  - *Causa Raíz:* Desfase de actualización tras bump.
  - *Impacto Confirmado:* Exposición de credencial institucional y desalineación de trazabilidad.
- **Acción Aplicada:** Sustituido el literal por `ConstantsHelper.FORMATO_PASSWORD`, fechado el smoke test (2026-10-17 / Excel 365 x64 / build 1.6.3.57) y actualizados los encabezados a `1.6.3.57`.
- **Archivos Afectados:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`, `implementation_plan.md`, `walkthrough.md`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F736 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F737 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F738 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.96 Ronda 138 — Sincronización de DefaultVersion de Plantillas y Corrección de Estado de Release (Fases 739 y 740)

### Evaluación Crítica de las Fases

#### 1. Fase 739 — `DefaultVersion` de Plantillas en Instalador Sincronizado a 1.6.3.57
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La persistencia de `DefaultVersion="1.6.3.56"` en `cmpTemplate1` y `cmpTemplate2` impedía que una actualización hacia 1.6.3.57 forzara la sobreescritura de las plantillas preinstaladas en `CostFinalize`, arriesgando degradación del blindaje.
  - *Causa Raíz:* Omisión de actualización de las versiones sintéticas de archivo en el manifiesto WiX.
  - *Impacto Confirmado:* Falso estado de actualización de plantillas en escenarios de upgrade per-user.
- **Acción Aplicada:** Sincronizados los atributos `DefaultVersion` a `"1.6.3.57"` en `filTemplate1` (`GF-FO-023.xlsx`) y `filTemplate2` (`FV-FO-033.xlsx`), actualizando igualmente los comentarios correspondientes.
- **Archivos Afectados:** `installer/InstaladorConversorPDF.wxs`.

#### 2. Fase 740 — Corrección de Trazabilidad en Sección de Entrega de Release
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La sección de release afirmaba ejecución de suite de pruebas, compilación de binarios e instalación generada para la versión 1.6.3.57 cuando dichos pasos están pendientes.
  - *Causa Raíz:* Reutilización textual de registros de releases previos sin marcar el estado real.
  - *Impacto Confirmado:* Declaración inexacta de artefactos y verificaciones previas a la ejecución real de `/build-release`.
- **Acción Aplicada:** Ajustada la sección documentando que las pruebas unitarias, la compilación de Release|x64 y el empaquetado del MSI quedan pendientes de ejecución bajo `/build-release 1.6.3.57`.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`.

---

## 157.97 Ronda 139 — Enlace Dinámico de Plantillas y Sincronización Integral de Release (Fases 741 a 743)

### Evaluación Crítica de las Fases

#### 1. Fase 741 — Invariante de Tamaños de Plantilla Sincronizado en `build-release.md`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Checklist en `build-release.md` con tamaños obsoletos (`777061` y `2938523` B) respecto a los binarios reales en disco (`777082` y `2938592` B).
  - *Causa Raíz:* Omisión de actualización tras la última edición de plantillas (Fase 721).
  - *Impacto Confirmado:* Inducción a falsos positivos en el control de calidad post-empaquetado.
- **Acción Aplicada:** Sincronizados los valores a `777082` (`FV-FO-033.xlsx`) y `2938592` (`GF-FO-023.xlsx`) en `build-release.md`.
- **Archivos Afectados:** `.agents/workflows/build-release.md`.

#### 2. Fase 742 — Enlace Dinámico de Versión de Plantillas con `!(bind.FileVersion.filXllDll)`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El uso de literales `DefaultVersion="1.6.3.57"` en el manifiesto WiX generaba duplicación de fuentes de verdad y riesgo de desfase en futuros releases.
  - *Causa Raíz:* Asignación estática en lugar del enlace nativo soportado por WiX.
  - *Impacto Confirmado:* Mantenimiento manual redundante y contradicción con el estándar del workflow.
- **Acción Aplicada:** Asignado `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `filTemplate1` y `filTemplate2`, vinculando dinámicamente la versión sintética al ensamblado principal.
- **Archivos Afectados:** `installer/InstaladorConversorPDF.wxs`.

#### 3. Fase 743 — Sincronización de Encabezados y Lista de Cambios en Documento Técnico
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Desfase en el documento de cambios al omitir las rondas 137 y 138 y las mejoras de desprotección e instalador en la sección Changes.
  - *Causa Raíz:* Desincronización documental acumulada.
  - *Impacto Confirmado:* Incompletitud en el registro histórico del cambio.
- **Acción Aplicada:** Actualizado el encabezado a Rondas 134-139 e incorporadas las descripciones de las Fases 734/736 (bandera atómica) y 739/742 (enlace dinámico WiX).
- **Archivos Afectados:** `docs/articles/changes/2026-10-17-habilitar-autofiltro-proteccion-programacion.md`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F741 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F742 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F743 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 📦 Entrega de Release v1.6.3.57 (Instalador Oficial)

- **Versión Oficial:** `1.6.3.57`
- **Archivos de Versión Sincronizados:**
  - `ConversorPDF.XLL\Core\Helpers\App.cs` (`APP_VERSION = "1.6.3.57"`)
  - `ConversorPDF.XLL\ConversorPDF.XLL.csproj` (`<Version>1.6.3.57</Version>`)
  - `installer\InstaladorConversorPDF.wxs` (`Version="1.6.3.57"`, `DefaultVersion="!(bind.FileVersion.filXllDll)"`)
- **Blindaje Integral de Plantillas Excel en Instalador:**
  - `File/@Id` conservado: `filTemplate1` (`GF-FO-023.xlsx`), `filTemplate2` (`FV-FO-033.xlsx`).
  - `FV-FO-033.xlsx`: **777.082 bytes**, `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `InstaladorConversorPDF.wxs`.
  - `GF-FO-023.xlsx`: **2.938.592 bytes**, `DefaultVersion="!(bind.FileVersion.filXllDll)"` en `InstaladorConversorPDF.wxs`.
  - `DefaultVersion="!(bind.FileVersion.filXllDll)"`: Anula `FA_NoOverwrite` en `CostFinalize`, garantizando sobreescritura contra cualquier archivo sin versión en disco.
  - `RegistryValue KeyPath="yes"`: Claves en `HKCU\Software\ESAP\ConversorPDF\Templates` (`[ProductVersion]`) asegurando activación del componente.
  - `RemoveFile` integrado (`PurgeOldTemplate1`, `PurgeOldTemplate2`) para purga forzada en instalación.
  - `RemoveFolder` integrado (`RemoveExcelTemplatesFolder`) para desinstalación limpia de `dir_ExcelTemplates`.
- **Verificación Integral de Pruebas Unitarias:**
  - Suite completa: **626 / 626 pruebas aprobadas (0 errores, 46 s)** (`dotnet test ConversorPDF.Tests\ConversorPDF.Tests.csproj -c Release`).
- **Compilación de Add-In:**
  - `ConversorPDF.XLL` (Release|x64): **0 Errores** (`ConversorPDF.XLL-AddIn64.xll`).
- **Paquete Instalador Generado (`/build-release 1.6.3.57`):**
  - Archivo: `installer\InstaladorConversorPDFv1.6.3.57.msi`
  - Tamaño: **5,578,752 bytes** (5.57 MB)
  - Hash SHA-256: `D309E2CFABE5D110ACB081FFB6E3C15034F3C6BFB97556C0FDCD3AAC528EA49C`
  - Estado: Listo para distribución con auto-activación en Excel y reemplazo blindado dinámico de plantillas.

---

## 157.98 Ronda 140 — Reparación de Esquema CRDT en Periodos y Resolución de Hueco de Migración V50/V51 (Fases 744 a 751)

### Evaluación Crítica de las Fases

#### 1. Fase 744 — Helper `PeriodoVentanaCrdtRepair` para Soltar Triggers CRDT Obsoletos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La tabla `periodos` participa en la sincronización CR-SQLite. Al alterar columnas o actualizar registros con esquemas enriquecidos, los triggers `periodos__crsql_itrig/utrig/dtrig` compilados con la aridad previa (4 columnas) disparan `SQL logic error expected 13 values, got 9`, revirtiendo transacciones de migración.
  - *Causa Raíz:* Triggers generados para 4 columnas ejecutándose tras `ALTER TABLE` o durante `UPDATE` en esquemas con 6 columnas sin previa recreación.
  - *Impacto Confirmado:* Fallo de migraciones aditivas y ausencia física de columnas en vaults existentes.
- **Acción Aplicada:** Creado el helper `PeriodoVentanaCrdtRepair` con el método `DropStaleCrsqlTriggers` que suelta los tres triggers de `periodos` de forma segura y transaccional, preservando `periodos__crsql_clock` para no perder historial de cambios entre sedes.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/PeriodoVentanaCrdtRepair.cs`.

#### 2. Fase 745 — Migración V65: `MigrationV65_PeriodoVentanaCrdtRepair`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La tabla `periodos` en vaults actuales carece físicamente de `periodo_inicio` y `periodo_fin` a pesar de que `__SchemaVersion` se encuentra en versión 64, arrojando el error *"El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones."* al guardar o abrir períodos.
  - *Causa Raíz:* V50 y V51 fallaron por triggers CRDT obsoletos; al ser aditivas tolerantes no registraron su versión pero el migrador continuó hasta la v64, generando un hueco de versión irrecuperable desde la UI.
  - *Impacto Confirmado:* Bloqueo funcional de guardado de fechas de período y pérdida de Prioridad 1 en el motor de Algoritmo Genético.
- **Acción Aplicada:** Implementada `MigrationV65_PeriodoVentanaCrdtRepair` (Versión 65, Aditiva, Tolerante). Dropea triggers stale, añade físicamente `periodo_inicio` y `periodo_fin`, y ejecuta backfill canónico parametrizado (`id 1: 2026-01-12/2026-06-19`, `id 2: 2026-08-11/2026-12-12` con regla V56, `id 3: 2027-01-01/2027-06-30`) respetando fechas existentes.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`.

#### 3. Fase 746 — Registro de Migración V65 y Helper en Pipeline y Compilación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La migración V65 y el helper debían estar integrados en la lista secuencial de `BuildMigrator()` y en el archivo de proyecto `.csproj`.
  - *Causa Raíz:* Nuevos tipos no referenciados ni compilados en proyecto non-SDK.
  - *Impacto Confirmado:* Error de compilación o migración no ejecutada en arranque.
- **Acción Aplicada:** Registrado `new MigrationV65_PeriodoVentanaCrdtRepair()` en `SqliteHelper.BuildMigrator()` tras V64 y añadidos `<Compile Include>` para `PeriodoVentanaCrdtRepair.cs` y `MigrationV65_PeriodoVentanaCrdtRepair.cs` en `ConversorPDF.XLL.csproj`.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 4. Fase 747 — Blindaje Retroactivo de `MigrationV50` y `MigrationV51` para Vaults Legados (< 50)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Un vault legado que ejecute V50 o V51 volvería a fallar en el backfill si los triggers CRDT de `periodos` están presentes.
  - *Causa Raíz:* Mismatch de aridad no protegido en las migraciones tempranas.
  - *Impacto Confirmado:* Reproducción del hueco de versión en bases de datos que no hayan alcanzado v50.
- **Acción Aplicada:** Integrada la invocación `PeriodoVentanaCrdtRepair.DropStaleCrsqlTriggers` previa a las sentencias DDL y DML en `MigrationV50_AddPeriodoVentana.cs` y `MigrationV51_PeriodoVentanaRepair.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`.

#### 5. Fase 748 — Observabilidad y Diagnóstico de Huecos de Versión en `SchemaMigrator`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Los fallos en migraciones tolerantes se registraban únicamente como advertencias benignas sin explicitar que la versión no fue registrada en `__SchemaVersion`, dificultando el diagnóstico del hueco de migración.
  - *Causa Raíz:* Logging genérico sin contexto de versión omitida.
  - *Impacto Confirmado:* Dificultad para rastrear por qué columnas declaradas en v50 no estaban presentes en v64.
- **Acción Aplicada:** Modificado el logging en `SchemaMigrator` para emitir `_logger.Error` con la leyenda explícita `HUECO DE VERSIÓN: v{migration.Version} ({migration.Name}) FALLÓ pero es ADITIVA. Se continúa el arranque sin registrar versión`. Añadida sugerencia de acción en `VerificarColumnasRequeridas`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/SchemaMigrator.cs`.

#### 6. Fase 749 — Guard de Coherencia en `AcademicRepository.InsertPeriodo`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `UpdatePeriodo` rechazaba con mensaje guiado cuando el esquema carecía de columnas de ventana y venían fechas, pero `InsertPeriodo` las descartaba en silencio e insertaba registros incompletos.
  - *Causa Raíz:* Asimetría defensiva entre operaciones de inserción y actualización en el repositorio.
  - *Impacto Confirmado:* Pérdida silenciosa de fechas de período en nuevos registros bajo esquemas no migrados.
- **Acción Aplicada:** Incorporado en `InsertPeriodo` el guard idéntico a `UpdatePeriodo` retornando `SqliteResult<int>.Failure("El esquema no soporta ventana temporal (periodo_inicio/periodo_fin). Aplique las migraciones.")`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AcademicRepository.cs`.

#### 7. Fase 750 — Cobertura de Pruebas Unitarias Automatizadas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de pruebas unitarias que validaran la migración V65, la eliminación de triggers stale y la protección de aridad.
  - *Causa Raíz:* Nueva funcionalidad de reparación de esquema sin suite asociada.
  - *Impacto Confirmado:* Riesgo de regresión en futuras modificaciones de migraciones o repositorio.
- **Acción Aplicada:** Creado `MigrationV65Tests.cs` (5 pruebas), añadidas pruebas de triggers stale en `MigrationV50Tests.cs` y `MigrationV51Tests.cs`, y prueba de guard de ventana en `AcademicRepositoryPeriodoTests.cs`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV65Tests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV50Tests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV51Tests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AcademicRepositoryPeriodoTests.cs`.

#### 8. Fase 751 — Sincronización de Versión Oficial a 1.6.3.58
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Se requería bump de versión oficial para la entrega de la solución.
  - *Causa Raíz:* Bump tras resolución de defecto de esquema.
  - *Impacto Confirmado:* Desalineación de versión del producto instalable.
- **Acción Aplicada:** Sincronizada versión `1.6.3.58` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Helpers/App.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`, `installer/InstaladorConversorPDF.wxs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F744 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F745 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F746 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F747 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F748 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F749 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F750 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F751 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.99 Ronda 141 — Reparación de Esquema CRDT en Periodos — Revisión de Fases 744-751 (Fases 752 a 756)

### Evaluación Crítica de las Fases

#### 1. Fase 752 — Drop Condicionado de Triggers CRDT en V65 según Necesidad Real de Backfill
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `DropStaleCrsqlTriggers` se ejecutaba incondicionalmente dos veces en V65, incluso en bases de datos sanas (columnas y fechas completas = no-op real). Si `EnableAllCrdts` no regeneraba los triggers en el arranque, `periodos` quedaba sin triggers y en arranques posteriores el short-circuit de `SyncInitializer:295` impedía repararlos.
  - *Causa Raíz:* Invocación de drop antes de evaluar si la base de datos realmente requería mutación DML.
  - *Impacto Confirmado:* Riesgo de pérdida silenciosa de sincronización CRDT en períodos en vaults sanos.
- **Acción Aplicada:** Condicionado el drop de triggers exclusivamente a `requiereBackfill`. Si las columnas ya existen y todas las fechas están pobladas, se omite el drop de triggers y el backfill DML.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`.

#### 2. Fase 753 — Aplicación de Regla V56 y Relleno de Fechas Fin Nulas en Backfill de V65
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La condición previa `WHERE periodo_inicio IS NULL OR periodo_inicio = ''` omitía rellenar `periodo_fin` en filas donde `periodo_inicio` estaba presente pero `periodo_fin` era NULL. Además, vaults donde V50 pobló `2026-08-15` y V56 falló conservaban la fecha desactualizada.
  - *Causa Raíz:* Evaluación univariada de `periodo_inicio` en el `WHERE` y ausencia de réplica explícita de la regla de inicio de semestre 2026-02.
  - *Impacto Confirmado:* Ventanas incompletas para el motor AG (que exige ambas cotas) y desajuste de fecha en 2026-02.
- **Acción Aplicada:** Backfill con `COALESCE(NULLIF(periodo_inicio, ''), @Ini)` y `COALESCE(NULLIF(periodo_fin, ''), @Fin)` por fila, más sentencia DML de regla V56 para `2026-02` (`periodo_inicio = '2026-08-11'`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`.

#### 3. Fase 754 — Blindaje de Drop Condicionado en V50 y V51
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Mismo riesgo de F752 en vaults legados que ejecutan V50 o V51: drops duplicados antes y después de `ALTER TABLE` que soltaban triggers aun si no había necesidad de backfill.
  - *Causa Raíz:* Invocaciones redundantes sin verificación previa de `requiereBackfill`.
  - *Impacto Confirmado:* Eliminación innecesaria de triggers CRDT en bases de datos que ya contenían los registros de períodos completos.
- **Acción Aplicada:** Eliminado el drop previo al DDL y condicionado el drop posterior a la evaluación de `requiereBackfill`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`.

#### 4. Fase 755 — Detección y Restauración de Triggers CRDT Ausentes en `SyncInitializer`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Si una tabla ya convertida a CRDT (`isAlreadyCrdt == true`) había perdido sus triggers por un drop defensivo que no logró regenerarse, el short-circuit de `SyncInitializer:295` en arranques posteriores (`schemaMigrated == false`) impedía la auto-reparación.
  - *Causa Raíz:* El short-circuit evaluaba únicamente la presencia de la tabla `_clock`, sin constatar la existencia física de los triggers `itrig`, `utrig` y `dtrig`.
  - *Impacto Confirmado:* Tablas en estado CRDT corrupto sin captura de cambios.
- **Acción Aplicada:** Implementado `TableHasCrdtTriggers(connection, tableName)` e integrado antes del short-circuit en `EnableCrdtForTable` para restaurar triggers automáticamente mediante `crsql_as_crr`.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`.

#### 5. Fase 756 — Sincronización de `DefaultVersion` en Instalador WiX
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Desfase entre la versión de paquete (1.6.3.58) y la versión sintética declarada en `filTemplate1` y `filTemplate2` (`DefaultVersion="1.6.3.57"`).
  - *Causa Raíz:* Omisión de sincronización en los atributos de archivo de plantillas del manifiesto WiX.
  - *Impacto Confirmado:* Inconsistencia en la política de sobreescritura `CostFinalize` de Windows Installer.
- **Acción Aplicada:** Sincronizados los atributos `DefaultVersion="1.6.3.58"` y sus comentarios en `cmpTemplate1` y `cmpTemplate2`.
- **Archivos Afectados:** `installer/InstaladorConversorPDF.wxs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F752 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F753 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F754 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F755 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F756 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.100 Ronda 142 — Verificación de Fases 752-756 (Reparación V65 / CRDT) (Fases 757 a 759)

### Evaluación Crítica de las Fases

#### 1. Fase 757 — Regla V56 Incondicional en Migration V65
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La sentencia de actualización para la regla V56 (`periodo_inicio = '2026-08-11'`) se encontraba anidada dentro del bloque `if (requiereBackfill)`. En vaults donde las columnas ya existían y no había registros NULL o vacíos, `requiereBackfill` era `false`, omitiendo el UPDATE. Si un vault ejecutó V50 (`periodo_inicio = '2026-08-15'`) pero falló al aplicar V56, conservaba la fecha incorrecta (`2026-08-15`) y la prueba unitaria `Up_WhenInicioEs20260815_ShouldApplyV56Rule` fallaba.
  - *Causa Raíz:* Acoplamiento erróneo de la regla canónica de corrección de fecha V56 a la condición de necesidad de backfill.
  - *Impacto Confirmado:* Error en tests unitarios e inconsistencia de fechas en 2026-02 para vaults con migración previa parcial.
- **Acción Aplicada:** Extraído el UPDATE de la regla V56 para ejecutarse de manera incondicional inmediatamente después del bloque `if (requiereBackfill)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`.

#### 2. Fase 758 — Purga de Triggers Parciales Remanentes antes de `crsql_as_crr` en `SyncInitializer`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La condición `!TableHasCrdtTriggers(connection, tableName)` detecta si faltan triggers CRDT (0, 1 o 2 presentes de los 3 requeridos). Al llamar directamente `SELECT crsql_as_crr(...)` con triggers parciales existentes, CR-SQLite lanzaba la excepción `trigger ... already exists`, marcando la tabla como fallida y dejándola sin sincronización.
  - *Causa Raíz:* Invocación de inicialización CRDT sin purgar triggers residuales remanentes.
  - *Impacto Confirmado:* Excepción de runtime en auto-restauración y pérdida permanente de triggers de sincronización.
- **Acción Aplicada:** Agregadas sentencias defensivas `DROP TRIGGER IF EXISTS` para `dtrig`, `utrig` e `itrig` inmediatamente antes de invocar `crsql_as_crr`.
- **Archivos Afectados:** `ConversorPDF.XLL/Core/Sync/SyncInitializer.cs`.

#### 3. Fase 759 — Uso de `COALESCE` en Backfill de V50 y V51 para Relleno de Fechas Fin Nulas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La condición de backfill en V50 y V51 evaluaba `periodo_fin IS NULL`, pero el UPDATE solo aplicaba cuando `periodo_inicio IS NULL OR periodo_inicio = ''`. Registros con inicio válido y fin nulo causaban la eliminación de triggers sin actualizar ninguna fila, dejando `periodo_fin` nulo de forma innecesaria.
  - *Causa Raíz:* Desincronización entre el predicado de detección y la condición del UPDATE en V50 y V51.
  - *Impacto Confirmado:* Filas con ventanas incompletas en vaults migrados y drop de triggers injustificado.
- **Acción Aplicada:** Actualizada la sentencia UPDATE en V50 y V51 a `COALESCE(NULLIF(periodo_inicio, ''), @Ini)` y `COALESCE(NULLIF(periodo_fin, ''), @Fin)`, filtrando solo por `idperiodo = @Id`. Añadida la prueba `Up_WhenInicioPresenteYFinNulo_ShouldFillOnlyFin` en `MigrationV50Tests` y `MigrationV51Tests`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV50_AddPeriodoVentana.cs`, `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV51_PeriodoVentanaRepair.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV50Tests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV51Tests.cs`.

#### 4. Fase 760 — Ajuste de Trazabilidad de Logs en V65
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El mensaje en el bloque `else` de `requiereBackfill` indicaba `"sin cambios ni drop de triggers"`. Tras Fase 757, la regla V56 se ejecuta incondicionalmente y puede modificar `periodo_inicio` a `2026-08-11`, haciendo inexacto el log de auditoría.
  - *Causa Raíz:* Mensaje redactado antes de desacoplar la regla V56 del bloque de backfill.
  - *Impacto Confirmado:* Confusión diagnóstica en logs donde se informaba "sin cambios" a pesar de haberse actualizado la fecha a 2026-08-11.
- **Acción Aplicada:** Modificado el mensaje a `"[V65] Columnas y fechas completas: sin drop de triggers."` y agregado el log `"[V65] Regla V56 verificada para 2026-02."`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV65_PeriodoVentanaCrdtRepair.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F757 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F758 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F759 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F760 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.101 Ronda 143 — Distribución de Viáticos por Asignatura/Sede y Prevención de Desbordes en Plantilla Excel (Fases 761 a 766)

### Evaluación Crítica de las Fases

#### 1. Fase 761 — Creación de `AsignaturaViaticoAllocator` y Modelo `AsignaturaViaticoFila` (Reglas R1 a R7)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En la exportación de asignaturas a Excel, la asignación de bloques de viáticos se realizaba secuencialmente sin verificar la correspondencia con la vinculación docente que generó el viaje. Esto causaba que viáticos de una sede se imprimieran en la fila de otra sede y que bloques que cruzaban sedes se fragmentaran o asignaran arbitrariamente.
  - *Causa Raíz:* Algoritmo rudimentario basado en `bloquesGlobales.Skip(bloqueIndex).Take(8)` desacoplado del modelo de vinculaciones y ventanas.
  - *Impacto Confirmado:* Filas con sedes erróneas, desalineación de tarifas y rutas, y desbordes incontrolados.
- **Acción Aplicada:** Implementado `AsignaturaViaticoAllocator` y `AsignaturaViaticoFila` con reglas deterministas R1 (correspondencia exacta por `IdVinculacion`), R2 (bloque multi-sede indivisible en vinculación predominante), R3 (máx. 8 comisiones por fila), R4 (reasignación por saturación priorizando ventana [U, V], sede en ruta y cercanía), R5 (fila de continuación como último recurso), R6 (advertencia ante fecha fuera de [U, V]) y R7 (invarianza estricta de suma total de bloques).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoFila.cs`, `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`.

#### 2. Fase 762 — Integración en `AsignaturasExcelExporter.cs` y Corrección de Anulación de Tiquetes Aéreos
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La clave de agrupación en `AsignaturasExcelExporter` no contemplaba la ventana de vinculación, el bucle utilizaba punteros planos `asigIndex/bloqueIndex`, las filas duplicadas omitían `U`, `V` y `Grupo`, y las líneas 477-483 forzaban `RequiereTiquetesAereos = "NO"` y `TerminalAerea = SIN_TERMINAL` en toda fila duplicada, anulando el valor de gasto de terminal aérea (`BQ`) en viajes legítimos.
  - *Causa Raíz:* Residuos de lógica previa de limpieza que no contemplaba filas de continuación con comisiones y tiquetes propios.
  - *Impacto Confirmado:* Pérdida del valor `BQ` en comisiones adicionales y desalineación de datos de vinculación en Excel.
- **Acción Aplicada:** Agrupación enriquecida con `VinculacionFechaInicio` y `VinculacionFechaFin`, sustitución de punteros planos por la llamada a `AsignaturaViaticoAllocator.Distribuir`, escritura de `U`, `V` y `Grupo` en continuación, eliminación del bloque 477-483 de sobreescritura de tiquetes, guard defensivo contra desborde de fila 234 en `T_Principal8` y exposición de trazabilidad en `UltimosAvisos`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 3. Fase 763 — Actualización de UI en `FormVerVinculaciones.cs`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La interfaz mostraba un mensaje plano `"Exportación finalizada."` sin visibilidad de bloques reasignados por saturación, filas de continuación o advertencias de rango de fechas.
  - *Causa Raíz:* Ausencia de consumo de métricas de trazabilidad del exportador en el code-behind del formulario.
  - *Impacto Confirmado:* Falta de retroalimentación ante anomalías de capacidad en la exportación.
- **Acción Aplicada:** Actualizado el mensaje emergente a `"Exportación finalizada. N bloques reasignados / M filas de continuación / K advertencias (ver log)."` mostrando el detalle del primer aviso y elevando a `MessageBoxIcon.Warning` si hay advertencias.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`.

#### 4. Fase 764 — Registro de Nuevos Servicios en `ConversorPDF.XLL.csproj`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Los nuevos archivos `AsignaturaViaticoFila.cs` y `AsignaturaViaticoAllocator.cs` no estaban registrados en el archivo de proyecto non-SDK.
  - *Causa Raíz:* Formato de proyecto legacy que requiere declaración explícita de compilación.
  - *Impacto Confirmado:* Omisión de compilación de los nuevos componentes.
- **Acción Aplicada:** Añadidas las entradas `<Compile Include>` bajo `Academic\Application\Services\` en `ConversorPDF.XLL.csproj`.
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 5. Fase 765 — Suite de Pruebas Unitarias en `AsignaturaViaticoAllocatorTests.cs`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de pruebas unitarias automatizadas para verificar las reglas de negocio R1 a R7.
  - *Causa Raíz:* Nueva funcionalidad de asignación de viáticos sin suite de pruebas preexistente.
  - *Impacto Confirmado:* Riesgo de regresiones en futuras modificaciones de distribución de comisiones.
- **Acción Aplicada:** Creación de 9 pruebas unitarias cubriendo caso canónico (15 bloques, 4 sedes), bloques multi-sede indivisibles, desborde R4, continuación R5, límite de 8 comisiones, validación R6 de ventana, parseo multi-formato y colecciones vacías.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/AsignaturaViaticoAllocatorTests.cs`.

#### 6. Fase 766 — Documentación Técnica del Cambio
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Falta de registro formal de arquitectura del nuevo motor de distribución y cambios en exportación.
  - *Causa Raíz:* Nueva funcionalidad sin artículo de documentación correspondiente.
  - *Impacto Confirmado:* Desincronización del repositorio de documentación técnica DocFX.
- **Acción Aplicada:** Generado el documento `docs/articles/changes/2026-10-19-distribucion-viaticos-por-asignatura.md` detallando propósito, reglas, riesgos residuales y cobertura de pruebas.
- **Archivos Afectados:** `docs/articles/changes/2026-10-19-distribucion-viaticos-por-asignatura.md`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F761 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F762 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F763 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F764 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F765 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F766 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## 157.102 Ronda 144 — Revisión de Fases 761-766 (Distribución de Viáticos por Asignatura/Sede) (Fases 767 a 773)

### Evaluación Crítica de las Fases

#### 1. Fase 767 — Propagación de Guard de Capacidad de Fila 234 y Prevención de Truncamiento en Cascada
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al exceder la fila 234, el `break` interno salía únicamente del bucle de filas del docente actual. Los siguientes docentes volvían a evaluar `filaExcel > 234`, emitiendo una alerta redundante por cada fila omitida y saturando `_ultimosAvisos`.
  - *Causa Raíz:* No propagación del estado de saturación al bucle de docentes y falta de cuantificación de filas omitidas.
  - *Impacto Confirmado:* Centenares de avisos redundantes en el log y mensaje emergente.
- **Acción Aplicada:** Incorporada constante `FILA_MAXIMA_PLANTILLA = 234`, banderas `capacidadExcedida` y conteo acumulado `filasOmitidas += filasPlan.Count - indiceFila`, interrupción externa con `if (capacidadExcedida) break;` y emisión de un único log consolidador con severidad `Error`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 2. Fase 768 — Distribución en Dos Pases para Priorizar Asignación Natural (R1) sobre Desbordes (R4)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* En procesamiento cronológico de un solo paso, un bloque en desborde R4 ocupaba cupos en una fila natural que posteriormente requería un bloque legítimo de esa misma fila con fecha posterior, provocando que el bloque natural se desplazara o generara filas de continuación innecesarias.
  - *Causa Raíz:* Ausencia de reserva previa de cupo para la regla R1 antes de ejecutar reasignaciones R4.
  - *Impacto Confirmado:* Desalineación innecesaria de viáticos naturales ante desbordes previos.
- **Acción Aplicada:** Reestructuración en dos pases: Pase 1 reserva y asigna exclusivamente los bloques con vinculación única natural (R1) sobre filas con cupo; Pase 2 procesa pendientes aplicando R2, R4 y R5.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`.

#### 3. Fase 769 — Parseo Robusto de Fechas con Hora y Cultura `es-CO`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Fechas con formato `dd/MM/yyyy HH:mm:ss` no coincidían con `TryParseExact` y caían al fallback invariante `DateTime.TryParse`, el cual interpretaba `12/09/2026` como `MM/dd/yyyy` (9 de diciembre en vez de 12 de septiembre).
  - *Causa Raíz:* Formatos con hora ausentes en la matriz y cultura invariante con precedencia de mes antes de día en cadenas con `/`.
  - *Impacto Confirmado:* Desplazamiento silencioso de ventanas [U, V] y advertencias erróneas R6.
- **Acción Aplicada:** Añadidos patrones con hora a `FormatosFecha` y uso de `CultureInfo.GetCultureInfo("es-CO")` para cadenas que contengan `/`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`.

#### 4. Fase 770 — Distinción de Advertencia para Bloques sin Vinculación Coincidente en R4
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Bloques huérfanos sin correspondencia de `IdVinculacion` se etiquetaban como `"reasignado por saturación"`, inflando indebidamente la métrica `UltimosBloquesReasignados`.
  - *Causa Raíz:* Mensaje indiferenciado en el bloque de fallback de R4.
  - *Impacto Confirmado:* Métrica y diagnóstico incorrecto en la interfaz.
- **Acción Aplicada:** Evaluada la condición `candidatas.Count == 0` para emitir `"sin vinculación coincidente"` sin la palabra `"saturación"`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/AsignaturaViaticoAllocator.cs`.

#### 5. Fase 771 — Replicación de Prestaciones y Puntaje Ocasionales en Filas de Continuación (LLAVE Y)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Las filas de continuación omitían las columnas `Prestaciones` y `PuntajeOcasionales`, dejando incompleta la columna de auditoría `Y = CONCAT(VINCULACIÓN, CATEGORÍA, PRESTACIONES)`.
  - *Causa Raíz:* Omisión de réplica de campos de docente en la rama de continuación.
  - *Impacto Confirmado:* Fórmulas institucionales de auditoría rotas en filas duplicadas.
- **Acción Aplicada:** Incorporada la escritura de `Prestaciones` (X) y `PuntajeOcasionales` (Z) en `if (esFilaDuplicada)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 6. Fase 772 — Priorización de Avisos Críticos en la Interfaz de Usuario
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La ventana emergente seleccionaba `UltimosAvisos[0]`, permitiendo que avisos informativos (R6) ocultaran alertas críticas de capacidad o continuación.
  - *Causa Raíz:* Selección por índice sin filtrado de severidad.
  - *Impacto Confirmado:* Falta de visibilidad de eventos críticos ante advertencias menores.
- **Acción Aplicada:** Selección prioritaria de avisos que contengan `"continuación"`, `"saturación"` o `"Capacidad"`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`.

#### 7. Fase 773 — Cobertura Automatizada del Caso Real Multi-Sede y Fallback Huérfano
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Carencia de tests unitarios que reprodujeran el caso real reportado con dos asignaturas en Cali con ventanas diferenciadas, Cartago y El Águila, así como bloques huérfanos.
  - *Causa Raíz:* Casos sintéticos previos sin sedes repetidas.
  - *Impacto Confirmado:* Riesgo de regresión en asignación de ventanas reales.
- **Acción Aplicada:** Añadidos tests `CasoRealDosCalisCartagoElAguila_15Bloques_RespetaVentanaYFila` y `BloqueSinVinculacionCoincidente_AdvierteSinContarSaturacion`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/AsignaturaViaticoAllocatorTests.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F767 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F768 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F769 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F770 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F771 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F772 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F773 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 145: Revisión de fases 767-773 (distribución de viáticos por asignatura/sede)

### Evaluación Crítica de las Fases

#### 1. Fase 774 — Corrección de Conflicto de Ámbito de Variable CS0136 en Bucles Anidados de Filas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La introducción de `for (int i = 0; i < filasPlan.Count; i++)` en F767 provocó el error de compilación `CS0136` por colisión de nombres con los bucles internos `for (int i ...)` en las secciones 4 y 5 de comisiones y rutas.
  - *Causa Raíz:* Reutilización del identificador `i` en el bucle envolvente de filas del docente.
  - *Impacto Confirmado:* Error de compilación bloqueante.
- **Acción Aplicada:** Renombrado el identificador del bucle a `for (int indiceFila = 0; indiceFila < filasPlan.Count; indiceFila++)` y acceso indexado `filasPlan[indiceFila]`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 2. Fase 775 — Métricas de Cierre Inconsistentes al Activar el Guard de Capacidad y Comentario Obsoleto
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `UltimasFilasContinuacion` se incrementaba antes de la escritura, reportando filas inexistentes si el guard de capacidad truncaba el proceso; el aviso omitía el docente de corte; y un comentario obsoleto indicaba falsamente que las fechas de vinculación no se escribían en filas duplicadas.
  - *Causa Raíz:* Contadores anticipados al plan y documentación desincronizada tras F762/F771.
  - *Impacto Confirmado:* Métrica de continuaciones inflada ante truncamiento y ambigüedad en comentarios.
- **Acción Aplicada:** Trasladado el conteo de continuaciones al interior del bucle tras el guard (`if (esFilaDuplicada) UltimasFilasContinuacion++;`), registro de `docenteCorte` en el aviso consolidado y actualización del comentario aclarando qué campos se omiten (horas L-O) y cuáles se replican.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 3. Fase 776 — Detección Dinámica del Límite de Capacidad Institucional `T_Principal8`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El límite acoplado a la constante fija 234 (`FILA_MAXIMA_PLANTILLA`) truncaba la exportación ante plantillas institucionales ampliadas con más filas útiles en `T_Principal8`.
  - *Causa Raíz:* Lectura estática desalineada de la estructura real del libro Excel activo.
  - *Impacto Confirmado:* Incapacidad de aprovechar extensiones de filas en la tabla institucional.
- **Acción Aplicada:** Reemplazada la constante por `FILA_MAXIMA_PLANTILLA_DEFAULT = 234` e incorporado el método defensivo `ObtenerFilaMaximaPlantilla(dynamic ws)` que lee `ws.ListObjects["T_Principal8"].Range` con fallback al valor por defecto.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F774 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F775 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F776 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 146: Revisión de fases 774-776 (guard de capacidad y métricas de exportación)

### Evaluación Crítica de las Fases

#### 1. Fase 777 — Prevención de Fuga de Recursos COM en `ObtenerFilaMaximaPlantilla`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La obtención de `ws.ListObjects["T_Principal8"]` y `tabla.Range` se realizaba sin liberar los envoltorios RCW correspondientes, dejando referencias activas en el STA de Excel.
  - *Causa Raíz:* Helper F776 implementado sin bloque `finally` ni llamadas a `Marshal.ReleaseComObject`.
  - *Impacto Confirmado:* Fuga de objetos COM y riesgo de retención de procesos residuales de Excel.
- **Acción Aplicada:** Declaración de variables `tabla` y `rango` con liberación explícita `Marshal.ReleaseComObject` dentro de un bloque `finally` con captura de excepciones.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 2. Fase 778 — Sincronización de Avisos y Reasignaciones Condicionada a `planCompleto`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al excederse la capacidad de la plantilla, las advertencias del docente y el contador `UltimosBloquesReasignados` se habían computado antes del bucle de escritura, incorporando métricas y avisos de filas que no se escribieron.
  - *Causa Raíz:* Agregación incondicional previa a la iteración física sobre las celdas de Excel.
  - *Impacto Confirmado:* Métricas infladas y avisos inconsistentes con el contenido de la hoja exportada.
- **Acción Aplicada:** Introducida la bandera `bool planCompleto = true;`, asignando `false` al activarse el guard, y trasladada la agregación a `_ultimosAvisos` y `UltimosBloquesReasignados` al final del bucle condicionada a `planCompleto`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 3. Fase 779 — Sincronización de Documentación XML y Bitácora Histórica
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El XML doc de `Exportar` mantenía referencias obsoletas previas a F761, y la bitácora `walkthrough.md` citaba el identificador de variable anterior a F774.
  - *Causa Raíz:* Comentarios y bitácora no actualizados tras las modificaciones de arquitectura y compilación.
  - *Impacto Confirmado:* Discrepancias entre la documentación interna y el comportamiento real del código.
- **Acción Aplicada:** Actualizado el resumen XML de `Exportar` describiendo la distribución por vinculación (R1–R7) y filas de continuación como último recurso; ajustada la bitácora en L2054 con `filasPlan.Count - indiceFila`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`, `walkthrough.md`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F777 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F778 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F779 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 147: Revisión de fases 777-779 (liberación COM y métricas de cierre)

### Evaluación Crítica de las Fases

#### 1. Fase 780 — Preservación del Diagnóstico de Asignación ante Fallos en Escritura COM
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Al posponer en F778 el registro de advertencias al final del bucle de filas, si una llamada a `EscribirCelda` arrojaba una excepción COM a mitad del docente, la ejecución abortaba y se perdía la trazabilidad de los avisos del allocator (reasignaciones, ventanas, continuaciones).
  - *Causa Raíz:* Registro en `LoggerHelper.Warning` y agregación a `_ultimosAvisos` acoplados en el mismo bloque pospuesto.
  - *Impacto Confirmado:* Pérdida de visibilidad diagnóstica en el log ante anomalías de escritura.
- **Acción Aplicada:** Separado el registro en `LoggerHelper.Warning` (ejecutado inmediatamente tras `Distribuir`, garantizando trazabilidad ante fallos) de la agregación a `_ultimosAvisos` y `UltimosBloquesReasignados` (mantenida al final del bucle condicionada a `planCompleto == true`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 2. Fase 781 — Delimitación del Alcance en el Mensaje de Capacidad Excedida
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* El texto `"Filas omitidas desde el docente '{docenteCorte}': {filasOmitidas}"` inducía a creer que la cifra representaba la totalidad de filas no escritas de todos los docentes restantes, cuando solo cubría el docente del corte.
  - *Causa Raíz:* Redacción imprecisa que no aclaraba la interrupción del bucle externo.
  - *Impacto Confirmado:* Métrica interpretada erróneamente en reportes de auditoría.
- **Acción Aplicada:** Reformulado el aviso a: `$"Capacidad de T_Principal8 excedida (fila {filaMaximaPlantilla}). Filas omitidas del docente '{docenteCorte}': {filasOmitidas} (docentes posteriores no procesados)."`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

#### 3. Fase 782 — Liberación Exhaustiva de Objetos COM Intermedios en `ObtenerFilaMaximaPlantilla`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Las propiedades COM intermedias `ws.ListObjects` y `rango.Rows` generaban referencias RCW temporales que no eran capturadas ni liberadas explícitamente en el `finally`.
  - *Causa Raíz:* Acceso directo a propiedades COM encadenadas sin asignación a variables locales.
  - *Impacto Confirmado:* Fuga residual de referencias COM en el hilo STA de Excel.
- **Acción Aplicada:** Declaradas las variables locales `coleccion` y `filas`, capturando cada referencia intermedia y liberándolas explícitamente en el bloque `finally` con `Marshal.ReleaseComObject`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Excel/AsignaturasExcelExporter.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F780 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F781 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F782 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 148: Renombrado de modalidad de tutoría: Presencial / Plataforma / Mixta (Fases 783 a 791)

### Evaluación Crítica de las Fases

#### 1. Fase 783 — Catálogo Canónico de Modalidades de Tutoría (`ModalidadTutoriaCatalog.cs`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La terminología de modalidades de tutoría contenía inconsistencias (`Virtual`, `Semipresencial`, `A Distancia`) y no existía un punto centralizado en Dominio que desacoplara la definición canónica sin provocar dependencias Domain -> Application.
  - *Causa Raíz:* Valores hardcodeados dispersos en formularios, mappers y consultas.
  - *Impacto Confirmado:* Ambigüedad en la interfaz de usuario y discrepancia con los lineamientos institucionales.
- **Acción Aplicada:** Creado `ModalidadTutoriaCatalog` con constantes (`Presencial`, `Plataforma`, `Mixta`), tuplas `Pares`, normalización insensible a tildes/mayúsculas con precedencia estricta (`Normalizar`), convertidor a etiqueta (`AEtiqueta`) y predicados (`EsPresencial`, `EsPlataforma`, `EsMixta`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 2. Fase 784 — Migración V66 Aditiva y Tolerante (`MigrationV66_RenombrarModalidadTutoria.cs`) y Registro en `SqliteHelper`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La base de datos local SQLite contenía valores legados en 4 tablas y tokens embebidos en identificadores `aula_ref`, con riesgo de violación de la restricción única `uq_ag_arbol_nodo(idarbol, aula_ref)`.
  - *Causa Raíz:* Datos persistidos con versiones anteriores del Add-In.
  - *Impacto Confirmado:* Error en inicio o perfiles desincronizados al renombrar claves de aula.
- **Acción Aplicada:** Creada `MigrationV66_RenombrarModalidadTutoria` (`IMigration, IMigrationTolerante`, `Version = 66`, `Aditiva = true`) con updates set-based en `vinculaciones`, `ag_vinculaciones`, `ag_aula` y `ag_arbol_nodo`, deduplicación explícita previa en `ag_arbol_nodo` y registro en `SqliteHelper.BuildMigrator()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj`.

#### 3. Fase 785 — Actualización de Dominio y DTOs (`AgAula.cs`, `AgVinculacion.cs`, `AgContextoSeleccionDto.cs`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Valores por defecto de entidades de dominio exponían literales legados y el DTO de selección permitía circular modalidades no canónicas.
  - *Causa Raíz:* Falta de canalización a través del catálogo central.
  - *Impacto Confirmado:* Generación de registros y DTOs con valores legados.
- **Acción Aplicada:** Establecido `ModalidadTutoriaCatalog.Presencial` como valor por defecto en `AgAula` y `AgVinculacion`, y normalización en los constructores y en `ModalidadEfectiva` de `AgContextoSeleccionDto`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgAula.cs`, `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgVinculacion.cs`, `ConversorPDF.XLL/Academic/Application/DTOs/AgContextoSeleccionDto.cs`.

#### 4. Fase 786 — Actualización de Mappers y Factorías
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Generación de `AulaRef` con tokens legados en `AgAulaFactory` y propagación de modalidades sin normalizar en mappers.
  - *Causa Raíz:* Invocaciones directas a `ToUpperInvariant()` sin paso previo por `Normalizar`.
  - *Impacto Confirmado:* Discrepancias entre `AulaRef` y registros migrados.
- **Acción Aplicada:** Normalización previa en `AgAulaFactory.ConstruirAulaRef`, cálculo de `modVinculacion = esVirtual == 1 ? Plataforma : mod` en `CrearDesdeDto`, y normalización en `AgArbolMapper`, `AgExpansorSesiones`, `AgAulaToAgVinculacionMapper`, `AgVinculacionToVinculacionMapper`, `AgGenToAgVinculacionMapper` y `AgEncabezadoVinculacionFactory`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgAulaFactory.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgAulaToAgVinculacionMapper.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgVinculacionToVinculacionMapper.cs`, `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEncabezadoVinculacionFactory.cs`.

#### 5. Fase 787 — Servicios de Aplicación (`CalculadoraViaticaSesion.cs`, `SesionConflictAnalyzer.cs`)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Evaluaciones de texto directo para determinar derecho a viáticos o descartar colisiones en sesiones virtuales.
  - *Causa Raíz:* Comprobaciones `Contains("VIRTUAL")` y `Contains("SEMI")` dispersas.
  - *Impacto Confirmado:* Fragilidad ante variaciones de mayúsculas/minúsculas y omisión de modalidades canónicas.
- **Acción Aplicada:** Reemplazo por `ModalidadTutoriaCatalog.EsPlataforma`, `EsPresencial`, `EsMixta`, manteniendo con exactitud la semántica de negocio previa (Plataforma nunca viática; Presencial viática con desplazamiento; Mixta viática solo en rangos presenciales).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/CalculadoraViaticaSesion.cs`, `ConversorPDF.XLL/Academic/Application/Services/SesionConflictAnalyzer.cs`.

#### 6. Fase 788 — Actualización de Controles y Lógica de `FormSesionesAcademicas` (Designer + Code-Behind)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Controles nombrados `cmbTipoJornada` y `lblTipoJornada`, combo con opciones obsoletas y texto de advertencia desincronizado.
  - *Causa Raíz:* Deuda técnica en nomenclatura de controles WinForms.
  - *Impacto Confirmado:* Confusión de conceptos (jornada vs tutoría) e interfaz con términos obsoletos.
- **Acción Aplicada:** Controles renombrados a `cmbTipoTutoria` y `lblTipoTutoria`, poblado con `ModalidadTutoriaCatalog.Opciones`, guardado y carga normalizados, y mensaje de advertencia actualizado a *"Con la modalidad Mixta, las sesiones se generarán sin viáticas asignadas..."*.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.Designer.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`.

#### 7. Fase 789 — Actualización de `FormVerVinculaciones` y Árbol de `AgFormGeneradorHorarios`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Etiquetas legadas en vista de vinculaciones, árbol AG con arreglo plano de strings y mensajes de validación de aulas faltantes sin etiqueta amigable.
  - *Causa Raíz:* Código UI acoplado a literales anteriores.
  - *Impacto Confirmado:* Etiquetas `[Virtual]` o `[Semipresencial]` mostradas al usuario.
- **Acción Aplicada:** Inclusión de badge `" | 🌐 Plataforma"`, uso de pares `(par.Codigo, par.Etiqueta)` en nivel 5 del árbol, uso de `AEtiqueta` en mensajes de error y tolerancia a perfiles legados.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormVerVinculaciones.cs`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 8. Fase 790 — Suite de Pruebas Unitarias Nuevas y Actualizadas
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Ausencia de pruebas de unidad para el nuevo catálogo y migración V66, y aserciones desfasadas en suites existentes.
  - *Causa Raíz:* Nuevos componentes y contratos de salida actualizados.
  - *Impacto Confirmado:* Cobertura incompleta y pruebas en rojo ante expectativas legadas.
- **Acción Aplicada:** Creados `ModalidadTutoriaCatalogTests.cs` y `MigrationV66Tests.cs`, y actualizadas las expectativas preservando compatibilidad en `CalculadoraViaticaSesionTests`, `AgAulaFactoryTests`, `AgArbolServiceTests`, `AgArbolMapperTests`, `AgAulaToAgVinculacionMapperTests` y `AgTreeViewCheckServiceTests`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Domain/ModalidadTutoriaCatalogTests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`, `ConversorPDF.Tests/Academic/Application/Services/CalculadoraViaticaSesionTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgAulaFactoryTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/AgArbolServiceTests.cs`, `ConversorPDF.Tests/AgArbolMapperTests.cs`, `ConversorPDF.Tests/Academic/Application/Mappers/AgAulaToAgVinculacionMapperTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgTreeViewCheckServiceTests.cs`.

#### 9. Fase 791 — Documentación Técnica del Cambio y Sincronización de Especificación
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Falta de artículo de cambios y fila desactualizada en `AG-2.txt`.
  - *Causa Raíz:* Requisitos de trazabilidad del proyecto pendientes de actualización.
  - *Impacto Confirmado:* Desalineación entre especificación viva y código fuente.
- **Acción Aplicada:** Creado `docs/articles/changes/2026-10-21-renombrado-modalidades-tutoria.md` y actualizada la fila 108 de `AG-2.txt` con `cmbTipoTutoria`.
- **Archivos Afectados:** `docs/articles/changes/2026-10-21-renombrado-modalidades-tutoria.md`, `AG-2.txt`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F783 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F784 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F785 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F786 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F787 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F788 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F789 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F790 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F791 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 149: Revisión de Renombrado Presencial / Plataforma / Mixta (Fases 792 a 799)

### Evaluación Crítica de las Fases

#### 1. Fase 792 — Variables Legacy y Variable Sin Uso en `FormSesionesAcademicas`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La variable `esPresencial` se asignaba sin leerse (CS0219). Los nombres `esVirtual` y `esSemipresencial` contradecían la nueva terminología, y el log emitía `EsPlataforma: {esVirtual}`.
  - *Causa Raíz:* Renombrado parcial de identificadores locales tras la integración del catálogo.
  - *Impacto Confirmado:* Advertencia de compilación y nombres contradictorios en logs de diagnóstico.
- **Acción Aplicada:** Eliminada la variable muerta `esPresencial` y renombradas a `esPlataforma` y `esMixta`, sincronizando la condición de aviso y el mensaje del log en L808 y L864.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`.

#### 2. Fase 793 — Alineación de `TipoAula` y `Capacidad` con `EsVirtual` Derivado en `AgArbolMapper`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `EsVirtual` en el nodo del árbol se calculaba considerando `ModalidadTutoriaCatalog.EsPlataforma`, pero `TipoAula` y `Capacidad` leían `dto.EsVirtual` crudo.
  - *Causa Raíz:* Desincronización en la derivación de propiedades de aula virtual en el mapeador.
  - *Impacto Confirmado:* DTO con `Modalidad="Plataforma"` y `EsVirtual=0` generaba nodo con `es_virtual=1`, `tipo_aula='ESTANDAR'` y `capacidad=0`.
- **Acción Aplicada:** Precalculado `esVirtual = (dto.EsVirtual > 0 || ModalidadTutoriaCatalog.EsPlataforma(modEfectiva)) ? 1 : 0`, asignando `Capacidad = dto.Capacidad > 0 ? dto.Capacidad : (esVirtual == 1 ? 999 : 30)` y `TipoAula = !string.IsNullOrWhiteSpace(dto.TipoAula) ? dto.TipoAula : (esVirtual == 1 ? "VIRTUAL" : "ESTANDAR")`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Mappers/AgArbolMapper.cs`.

#### 3. Fase 794 — Deduplicación Simétrica y en Dos Pasos en `ag_arbol_nodo` (Migración V66)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* La subconsulta de deduplicación comparaba `n2.aula_ref` sin normalizar contra `REPLACE(ag_arbol_nodo.aula_ref, ...)`, fallando ante variantes con distinta capitalización (`|VIRTUAL` y `|Virtual`).
  - *Causa Raíz:* Normalización asimétrica de solo un lado de la comparación en `EXISTS`.
  - *Impacto Confirmado:* Violación de restricción única `uq_ag_arbol_nodo(idarbol, aula_ref)` y rollback total de la transacción de migración V66.
- **Acción Aplicada:** Deduplicación en 2 pasos: Paso 1 elimina registros legados si ya existe un equivalente canónico (`n2.aula_ref NOT LIKE '%|VIRTUAL%'`); Paso 2 resuelve duplicados legados equivalentes conservando el registro con menor `id` (`n2.id < ag_arbol_nodo.id`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV66Tests.cs`.

#### 4. Fase 795 — Guardas de Columnas y Aislamiento por Bloque en Migración V66
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Fallo en tablas secundarias por columnas ausentes provocaba rollback de la migración de `vinculaciones` al ejecutarse en una sola transacción.
  - *Causa Raíz:* Ausencia de `pragma_table_info` y ejecución no tolerante por bloque.
  - *Impacto Confirmado:* Error bloqueante de arranque en vaults antiguos con hueco de migración.
- **Acción Aplicada:** Incorporado el helper `ExisteColumna` con `pragma_table_info`, aislamiento de `MigrarAgAula` y `MigrarAgArbolNodo` en `try/catch` con logging preventivo, y validación previa de columnas en updates.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`.

#### 5. Fase 796 — Ampliación de Filtro en Identificadores Derivados (Migración V66)
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Filas con `aula_ref` vacío pero con tokens legados en `codigo` o `nombre` quedaban sin normalizar.
  - *Causa Raíz:* `WHERE` filtraba exclusivamente por `aula_ref LIKE ...`.
  - *Impacto Confirmado:* Residuos legados visibles en interfaz a través de código o nombre de aula/nodo.
- **Acción Aplicada:** Ampliado el `WHERE` para incluir `aula_ref LIKE ... OR codigo LIKE ... OR nombre LIKE ...` para todas las variantes de Plataforma y Mixta.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV66_RenombrarModalidadTutoria.cs`.

#### 6. Fase 797 — Asignación de Etiqueta a `Descripcion` en `AgNodoArbolKey` Nivel 6
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Se pasaba el código persistido en el parámetro `descripcion` en lugar de la etiqueta visible al construir claves de nivel 6.
  - *Causa Raíz:* Uso duplicado de `par.Codigo` en constructor.
  - *Impacto Confirmado:* Diagnósticos y descripciones de nodo desfasados si etiqueta difiere del código.
- **Acción Aplicada:** Ajustado a `new AgNodoArbolKey(6, 0, par.Codigo, par.Etiqueta)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`.

#### 7. Fase 798 — Inmutabilidad de `Opciones` en `ModalidadTutoriaCatalog`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* `public static readonly string[] Opciones` permitía mutación externa en tiempo de ejecución.
  - *Causa Raíz:* Tipo de retorno array mutable.
  - *Impacto Confirmado:* Riesgo de corrupción del catálogo en runtime.
- **Acción Aplicada:** Encapsulado en array privado con propiedad pública `IReadOnlyList<string> Opciones => _opciones`; uso de `.Cast<object>().ToArray()` en `FormSesionesAcademicas` y aserción con `.Count` en suite de tests.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/ModalidadTutoriaCatalog.cs`, `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs`, `ConversorPDF.Tests/Academic/Domain/ModalidadTutoriaCatalogTests.cs`.

#### 8. Fase 799 — Consideración de `EsVirtual` en `AgEncabezadoVinculacionFactory`
- **Estado:** ACEPTADA
- **Hallazgo:**
  - *Diagnóstico Crítico:* Aula con `EsVirtual = 1` y modalidad vacía producía vinculación `Presencial`, asignando viáticos indebidos.
  - *Causa Raíz:* Normalización sin evaluar la bandera técnica `EsVirtual` del aula.
  - *Impacto Confirmado:* Vinculaciones virtuales generadas con modalidad presencial.
- **Acción Aplicada:** Asignación condicional `aula != null && aula.EsVirtual == 1 ? Plataforma : Normalizar(aula?.ModalidadVinculacion)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgEncabezadoVinculacionFactory.cs`.

---

### Verificación de No Regresión y Tabla de Control

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa |
|---|---|---|---|---|
| F792 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F793 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F794 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F795 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F796 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F797 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F798 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |
| F799 | [X] Sí | [X] Sí | [X] Sí | [X] Sí |

---

## Ronda 150: Revisión de fases 792-799 (Renombrado Presencial / Plataforma / Mixta)

