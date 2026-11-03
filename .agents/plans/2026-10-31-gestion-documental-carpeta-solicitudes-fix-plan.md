# Plan consolidado: Gestión Documental de Viáticos — renombrado sobre `Solicitudes` y desambiguación de comisiones
**Generado**: 2026-10-31 · **Autor**: Cline · **Estado**: approved (handoff a otro modelo) · **Ejecutado**: no

> **Documento autocontenido.** El modelo ejecutor NO debe re-planificar, re-numerar ni re-investigar: todo lo necesario está aquí. Si algún extracto de código no coincide con el archivo real, **detenerse y reportar** en lugar de improvisar.

## Execution envelope
- branch: refactor/academic-modularization (HEAD `3bb7b18`)
- workspace: `c:\Users\USUARIO\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`
- preflight: `powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace <workspace> -Plan .agents/plans/2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md`
- workspace_guard:
  - min_xll_cs: 300
  - min_tests_cs: 100
  - required_slice: Academic
- baseline: `tests/baseline/test_baseline.txt` → `TOTAL: 952` (medido 2026-10-27). **Re-medir** antes de la primera ola; el árbol ya suma pruebas nuevas, así que el `TOTAL` real debe ser ≥ 952.
- verify: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- build_gate: serializado (un único build/tests a la vez; Excel-DNA bloquea `bin/`)
- max_parallel_writers: 2
- agents: plan-reviewer, tdd-red-writer, implementer, integrator, task-reviewer, build-verifier
- hot_paths_L0:
  - ConversorPDF.XLL/ConversorPDF.XLL.csproj
  - ConversorPDF.XLL/Core/Helpers/App.cs
  - ConversorPDF.XLL/**/Migrations/MigrationV*.cs
  - ConversorPDF.XLL/**/Persistence/SchemaMigrator.cs
  - ConversorPDF.XLL/Ribbon/RibbonController.cs
  - Composition Root / IoC Container
  - docs/toc.yml
- stop_conditions:
  - Acción destructiva no reversible sin respaldo
  - Violación de seguridad o datos
  - Intento de push, merge o publicación a remoto
  - Plan en estado irrecuperable (más de 3 rondas de fix fallidas)
  - Renombrar archivos del usuario fuera de una copia de prueba sin aprobación explícita
- ledger_owner: swarm-parallel-task
- no_push: true
- numbering:
  - registry: .agents/plans/REGISTRY.md
  - registry_row: R-003
  - ronda_base: 266          # Wave 1 = Ronda 266
  - fase_base: 1437          # T1 = Fase 1437
  - waves: 3                 # Ronda 266–268
  - tasks: 6                 # Fase 1437–1442

---

## 0. Estado de partida (leer antes de tocar código)

| Elemento | Valor |
|---|---|
| Rama | `refactor/academic-modularization`, HEAD `3bb7b18` |
| Versión en código | `App.APP_VERSION = "1.6.3.83"` (bump ejecutado en Ronda 272 / Fase 1461) |
| Cambios sin commitear | Los de la Ronda 264–265 (R-002): `ViaticoNumeroComision.cs` (nuevo), tests nuevos, interfaces/servicio/procesador con rango `(inicio, fin)`, `FormGestionDocumental` con `TryLeerRango`, docs e `installer` |
| Ya existe y **NO** se rehace | `Core/Helpers/ViaticoNumeroComision.cs` (`HayFiltro`, `TryParse`, `TryParseLimite`, `EnRango`); filtro por rango en `GetSolicitudes`; `TryLeerRango` en el formulario; `InternalsVisibleTo("ConversorPDF.Tests")` |
| Convención de comisión | SIIF `[Consecutivo][YY]`: `126` = No.1/2026, `41326` = No.413/2026, `12625` = No.126/2025 |
| Convención de nombre destino | `[Comisión]-[Cédula]-[Nombre18].pdf`; `Nombre18` = `App.CapitalizarPalabras` + quitar espacios + truncar a 18 |
| Suite de pruebas | xUnit 2.7 + FluentAssertions 6.12 + Moq 4.20 (`ConversorPDF.Tests`, net48, `LangVersion latest`) |
| Log de producción | `%LOCALAPPDATA%\ConversorPDF\log\log_YYYYMMDD.txt`; en pruebas se redirige a `%TEMP%\ConversorPDF.Tests\log` (`TestLogBootstrap.cs`) |

---

## 1. Diagnóstico consolidado (evidencia; no repetir la investigación)

Fuentes verificables: `%LOCALAPPDATA%\ConversorPDF\log\log_20261031.txt`, `scratch\propuesta_renombres.csv` (plan de 256 pendientes), `scratch\diag_consultas.ps1`, `scratch\diag_plan_renombres.ps1`, `scratch\diag_ambiguas.ps1` y la sonda `scratch\pdfprobe\PdfProbe.exe` (PdfPig 0.1.9, misma librería del producto). Todo `scratch/` está en `.gitignore`.

| ID | Defecto | Evidencia |
|---|---|---|
| **D1** | `RenombrarSolicitudes` escanea la carpeta del libro (raíz), no `Solicitudes`, y sin recursión. | `ViaticosService.cs:107-108` pasa `GetOneDriveFolderPath()`; `ViaticosGestionProcessor.cs:62` usa `Directory.GetFiles(folder, "*.pdf")`. Raíz: 15 PDF / **0** renombrables. `Solicitudes`: 419 PDF / **231** renombrables. |
| **D2** | El mensaje `No se encontraron comisiones para el rango indicado` sale con `n == 0` aunque hubo 416 comisiones en rango. | Log `log_20261031.txt`: `[CONSULTAS] filas=416 | enRango=416 | rango=126-41326`. `FormGestionDocumental.cs:82-90`. |
| **D3** | Comisiones `3326`, `21626`, `23826` duplicadas **con filas idénticas** (misma cédula y nombre) ⇒ `matches.Count == 2` fuerza OCR innecesario. | Filas 34/35, 218/219, 241/242 de `CONSULTAS`. |
| **D4** | `FindCC` fallback devuelve el **rubro presupuestal** (`0505023`) en lugar de la cédula. | Sonda PdfPig 0.1.9 sobre los 2 PDF ambiguos: patrón 1 → `52371279` / `66955819` (correctos); patrón 2 → `0505023`. |
| **D5** | Los PDF de `Solicitudes` no llevan metadatos de la app (`Keywords` vacío, productor Excel) ⇒ Capa 1 no aplica. | Sonda PdfPig. |
| **D6** | `Estandarizar` renombra **todos** los archivos del directorio, no solo `*.pdf`. | `ViaticosGestionProcessor.cs:178-186`. |
| **D7** | Sin traza por archivo: solo se registra el resumen de `CONSULTAS`. | El bucle de renombrado no emite logs. |
| **D8** | 25 PDF no tienen comisión en `CONSULTAS`: 23 con comisión `41426`–`43626` (la hoja termina en `41326`) y 2 con prefijo de 6 dígitos (`202426`, `202626`). | Cruce nombre↔hoja (`scratch\propuesta_renombres.csv`). |

### Dataset de referencia (para validar; no se codifica)

| Grupo | Cantidad |
|---|---|
| PDF en `Solicitudes` | 419 (163 ya conformes + 256 pendientes) |
| Pendientes renombrables | **231** = 229 comisión única + 2 duplicadas idénticas |
| Pendientes sin comisión en `CONSULTAS` | 25 (23 fuera del rango de la hoja + 2 con prefijo malformado) |
| PDF en raíz | 15 (1 `YA_CORRECTO`, 3 `SIN_COMISION`, 11 `PREFIJO_INVALIDO`) |

---

## 2. Mapa de Rondas y Fases

| Wave | Ronda | Task | Fase | Alcance | Depends on |
|---|---|---|---|---|---|
| 1 | **266** | T1 | **1437** | Procesador: carpetas, `Estandarizar` solo PDF, modelo de resultado, enumeración y contadores | — |
| 1 | **266** | T2 | **1438** | Interfaces + `ViaticosService` + `FormGestionDocumental`: resultado detallado y mensajes fieles | T1 |
| 2 | **267** | T3 | **1439** | `RenombrarFila`: dedupe `(Num,CC)`, política N0/N1/N2/N3, motivos, `GetSolicitudes` con `Trim` | T1 |
| 2 | **267** | T4 | **1440** | `FindCC` endurecido (validación contra cédulas candidatas) | T3 |
| 3 | **268** | T5 | **1441** | 7 pruebas de regresión (carpeta, dedupe, OCR, colisión, prefijo, error) | T1–T4 |
| 3 | **268** | T6 | **1442** | Docs, changelog, cierre de `R-003` y (opcional) bump a 1.6.3.83 | T5 |

Reglas de numeración (`.agents/plans/REGISTRY.md`): `Wave k` ⇔ `Ronda 265+k`; `T n` ⇔ `Fase 1436+n`. La fila `R-003` ya existe en estado `reserved`: **no** crear otra fila ni renumerar. El cierre (Fase C) cambia `reserved → closed` y actualiza contadores. La tabla de "Can start when" está en §13.

---

### 2.1 Contratos y convenciones que el implementador NO cambia

| Contrato | Valor |
|---|---|
| `RenombrarSolicitudes(string folder, int inicio, int fin)` | Se conserva (`int` = renombrados, `-1` = error de infraestructura) |
| `RenombrarSolicitudesConDetalle(...)` | Nuevo; devuelve `ResultadoRenombrado` y nunca `null` |
| `UnificarSolicitudes` | **Fuera de alcance** (ya usa `folder\Solicitudes` en `pSoli`) |
| `GetSolicitudes(inicio, fin)` | Se conserva `internal` (usado por pruebas vía `InternalsVisibleTo`) |
| `SInfo` | Se conserva `internal class SInfo { public string Num, CC, Nom; }` |
| Mensajes al usuario | Solo desde `FormGestionDocumental`; la capa Core **no** referencia `MessageBox` (R-002 lo eliminó) |
| Nombres de archivo | `[Comisión]-[Cédula]-[Nombre18].pdf`; no se inventan sufijos nuevos |


---

## 3. Fase 1437 (T1) — Procesador: carpetas correctas, modelo de resultado y contadores

- **id**: T1 · **fase**: 1437 · **depends_on**: [] · **status**: pending · **log**: -
- **location**: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
  - `ConversorPDF.XLL/Core/Models/ResultadoRenombrado.cs` (nuevo)
  - `ConversorPDF.XLL/Core/Models/MotivoRenombrado.cs` (nuevo)
  - `ConversorPDF.XLL/ConversorPDF.XLL.csproj` (registrar los 2 archivos nuevos en `<Compile Include>`; el proyecto es de estilo antiguo)
- **description**: enumerar raíz + `Solicitudes`, limitar `Estandarizar` a `*.pdf`, devolver un resultado detallado con contadores y motivos, y emitir log por archivo + resumen.
- **validation**: `RenombrarSolicitudesConDetalle` sobre la carpeta real devuelve `PdfsExaminados = 419 + 15 = 434` y `ComisionesEnRango = 416`.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"`
- **evidence**: `_evidence/<timestamp>/summary.md`

### 3.1 Nuevo `Core/Models/ResultadoRenombrado.cs`

```csharp
using System;
using System.Collections.Generic;

namespace ConversorPDF.XLL.Core.Models
{
    /// <summary>
    /// Resultado detallado del renombrado de solicitudes de viáticos.
    /// Permite distinguir "sin comisiones en el rango" de "comisiones sin PDF coincidentes".
    /// </summary>
    public class ResultadoRenombrado
    {
        /// <summary> Carpeta raíz recibida (carpeta del libro). </summary>
        public string CarpetaRaiz { get; set; }

        /// <summary> Carpetas realmente escaneadas (raíz y, si existe, Solicitudes). </summary>
        public List<string> CarpetasProcesadas { get; set; }

        /// <summary> Mensaje de error de infraestructura; vacío si la ejecución fue normal. </summary>
        public string Error { get; set; }

        /// <summary> True cuando la ejecución terminó por error de infraestructura. </summary>
        public bool HayError { get { return !string.IsNullOrEmpty(Error); } }

        /// <summary> Comisiones de CONSULTAS dentro del rango solicitado. </summary>
        public int ComisionesEnRango { get; set; }

        /// <summary> PDF examinados en todas las carpetas procesadas. </summary>
        public int PdfsExaminados { get; set; }

        public int Renombrados { get; set; }
        public int YaCorrectos { get; set; }
        public int Colisiones { get; set; }
        public int Ambiguos { get; set; }
        public int SinComision { get; set; }
        public int PrefijoInvalido { get; set; }
        public int ErroresIo { get; set; }

        /// <summary> Nombre original del archivo → motivo (ver <see cref="MotivoRenombrado"/>). </summary>
        public Dictionary<string, string> Detalle { get; set; }

        public ResultadoRenombrado()
        {
            CarpetasProcesadas = new List<string>();
            Detalle = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        }
    }
}
```

### 3.2 Nuevo `Core/Models/MotivoRenombrado.cs`

```csharp
namespace ConversorPDF.XLL.Core.Models
{
    /// <summary> Motivos posibles del resultado del renombrado de un PDF. </summary>
    public static class MotivoRenombrado
    {
        public const string Renombrado = "RENOMBRADO";
        public const string YaCorrecto = "YA_CORRECTO";
        public const string Colision = "COLISION";
        public const string Ambigua = "AMBIGUA";
        public const string SinComision = "SIN_COMISION_EN_CONSULTAS";
        public const string PrefijoInvalido = "PREFIJO_INVALIDO";
        public const string ErrorIo = "ERROR_IO";
    }
}
```

### 3.3 Reemplazo de `RenombrarSolicitudes` (líneas 42–74 actuales)

**Antes** (referencia, no editar a mano este bloque): `RenombrarSolicitudes` hace `Estandarizar(folder)`, recorre `Directory.GetFiles(folder, "*.pdf")` y cuenta con `if (RenombrarFila(file, solicitudes, folder)) count++;`.

**Después** (reemplazar el método completo por estos dos miembros):

```csharp
        /// <inheritdoc/>
        public int RenombrarSolicitudes(string folder, int inicio = 0, int fin = 0)
        {
            ResultadoRenombrado resultado = RenombrarSolicitudesConDetalle(folder, inicio, fin);
            return resultado.HayError ? -1 : resultado.Renombrados;
        }

        /// <inheritdoc/>
        public ResultadoRenombrado RenombrarSolicitudesConDetalle(string folder, int inicio = 0, int fin = 0)
        {
            var r = new ResultadoRenombrado { CarpetaRaiz = folder };
            try
            {
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                {
                    r.Error = $"La carpeta '{folder}' no existe.";
                    LoggerHelper.Error("RenombrarSolicitudes: " + r.Error);
                    return r;
                }

                // Comisiones de CONSULTAS filtradas por el rango solicitado
                var solicitudes = GetSolicitudes(inicio, fin);
                r.ComisionesEnRango = solicitudes.Count;
                if (solicitudes.Count == 0)
                {
                    LoggerHelper.Warning($"RenombrarSolicitudes: no se encontraron comisiones en CONSULTAS para el rango {inicio}-{fin}.");
                    return r;
                }

                // Raíz del libro + subcarpeta estándar de solicitudes
                foreach (string carpeta in CarpetasDeSolicitudes(folder))
                {
                    r.CarpetasProcesadas.Add(carpeta);
                    Estandarizar(carpeta);

                    foreach (string file in Directory.GetFiles(carpeta, "*.pdf"))
                    {
                        r.PdfsExaminados++;
                        string motivo = RenombrarFila(file, solicitudes, carpeta);
                        r.Detalle[Path.GetFileName(file)] = motivo;
                        Contabilizar(r, motivo);
                        LoggerHelper.Debug($"RenombrarSolicitudes: '{Path.GetFileName(file)}' -> {motivo}");
                    }
                }

                LoggerHelper.Resumen("RENOMBRAR",
                    ("carpetas", r.CarpetasProcesadas.Count),
                    ("comisionesEnRango", r.ComisionesEnRango),
                    ("pdfs", r.PdfsExaminados),
                    ("renombrados", r.Renombrados),
                    ("yaCorrectos", r.YaCorrectos),
                    ("colisiones", r.Colisiones),
                    ("ambiguos", r.Ambiguos),
                    ("sinComision", r.SinComision),
                    ("prefijoInvalido", r.PrefijoInvalido),
                    ("erroresIo", r.ErroresIo));

                return r;
            }
            catch (Exception ex)
            {
                LoggerHelper.Error("RenombrarSolicitudes", ex);
                r.Error = ex.Message;
                return r;
            }
        }
```

### 3.4 Helpers nuevos (agregar junto a `Estandarizar`)

```csharp
        /// <summary>
        /// Enumera la carpeta raíz y, si existe, la subcarpeta estándar de solicitudes.
        /// Evita duplicar la raíz cuando ya es la carpeta de solicitudes.
        /// </summary>
        private static IEnumerable<string> CarpetasDeSolicitudes(string folder)
        {
            yield return folder;

            string pSoli = Path.Combine(folder, App.FolderNames.SOLICITUDES);
            if (!string.Equals(pSoli, folder, StringComparison.OrdinalIgnoreCase) && Directory.Exists(pSoli))
            {
                yield return pSoli;
            }
        }

        /// <summary> Incrementa el contador correspondiente al motivo devuelto por <see cref="RenombrarFila"/>. </summary>
        private static void Contabilizar(ResultadoRenombrado r, string motivo)
        {
            switch (motivo)
            {
                case MotivoRenombrado.Renombrado: r.Renombrados++; break;
                case MotivoRenombrado.YaCorrecto: r.YaCorrectos++; break;
                case MotivoRenombrado.Colision: r.Colisiones++; break;
                case MotivoRenombrado.Ambigua: r.Ambiguos++; break;
                case MotivoRenombrado.SinComision: r.SinComision++; break;
                case MotivoRenombrado.PrefijoInvalido: r.PrefijoInvalido++; break;
                default: r.ErroresIo++; break;
            }
        }
```

### 3.5 `Estandarizar`: solo PDF (D6)

Cambiar **una línea** en `Estandarizar` (línea 180 actual):

```csharp
            // antes:  foreach (string file in Directory.GetFiles(f))
            foreach (string file in Directory.GetFiles(f, "*.pdf"))
```

### 3.6 `RenombrarFila` versión T1 (compila esta ola; T3 la reemplaza)

`RenombrarFila` debe pasar de `bool` a `string` **en esta misma fase** para que el bucle compile. Conserva la estrategia actual (prefijo + metadatos/OCR) y solo añade el motivo:

```csharp
        /// <summary>
        /// Intenta identificar y renombrar un PDF. Conserva la estrategia vigente
        /// (prefijo + metadatos/OCR) y devuelve el motivo del resultado.
        /// </summary>
        private string RenombrarFila(string file, List<SInfo> sols, string carpeta)
        {
            string oldN = Path.GetFileNameWithoutExtension(file);
            string ext = Path.GetExtension(file);

            string num = oldN.Split('-')[0];
            var matches = sols.Where(s => s.Num == num).ToList();
            if (matches.Count == 0) return MotivoRenombrado.SinComision;

            SInfo m = null;
            try
            {
                var meta = _metadataService.LeerMetadatos(file);
                if (meta != null && meta.MetaDocumento == TipoMetadata.VIA_SOLICITUD && !string.IsNullOrEmpty(meta.PalabrasClave))
                {
                    string metaCC = meta.PalabrasClave.Split(',')[0].Trim();
                    m = matches.FirstOrDefault(s => s.CC == metaCC);
                }
            }
            catch { /* metadatos ilegibles: se continúa con OCR */ }

            if (m == null)
            {
                m = matches.Count == 1 ? matches[0] : matches.FirstOrDefault(s => s.CC == FindCC(file));
            }

            if (m == null) return MotivoRenombrado.Ambigua;

            string newN = $"{m.Num}-{m.CC}-{m.Nom}{ext}";
            if (string.Equals(oldN + ext, newN, StringComparison.OrdinalIgnoreCase)) return MotivoRenombrado.YaCorrecto;

            string p = Path.Combine(carpeta, newN);
            if (File.Exists(p)) return MotivoRenombrado.Colision;

            try
            {
                File.Move(file, p);
                return MotivoRenombrado.Renombrado;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"RenombrarSolicitudes: no se pudo renombrar '{Path.GetFileName(file)}': {ex.Message}");
                return MotivoRenombrado.ErrorIo;
            }
        }
```

**Checklist de cierre de la Fase 1437**
1. `using System.Collections.Generic;`, `using System.IO;`, `using System.Linq;`, `using System.Text.RegularExpressions;` ya existen en el archivo.
2. `using ConversorPDF.XLL.Core.Models;` ya existe (se usa `SInfo`/`TipoMetadata`).
3. Build: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`.
4. Tests vigentes en verde (los 3 de `ViaticosGestionProcessorTests` deben seguir pasando).
5. Commit atómico: `Fase 1437 (T1): renombrar solicitudes sobre raíz y subcarpeta Solicitudes`.

---

## 4. Fase 1438 (T2) — Interfaces, servicio y mensajes del formulario

- **id**: T2 · **fase**: 1438 · **depends_on**: [T1] · **status**: pending · **log**: -
- **location**: `ConversorPDF.XLL/FormProgramador/FormGestionDocumental.cs`
- **files_lock**:
  - `ConversorPDF.XLL/Core/Interfaces/IViaticosGestionProcessor.cs`
  - `ConversorPDF.XLL/Core/Interfaces/IViaticosService.cs`
  - `ConversorPDF.XLL/Core/Services/ViaticosService.cs`
  - `ConversorPDF.XLL/FormProgramador/FormGestionDocumental.cs`
- **description**: exponer el resultado detallado y reemplazar el mensaje que confunde "sin comisiones" con "0 renombrados".
- **validation**: con rango `126-41326` y 416 comisiones, la UI nunca muestra "No se encontraron comisiones…"; muestra el desglose.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Build`
- **evidence**: `_evidence/<timestamp>/summary.md`

### 4.1 `IViaticosGestionProcessor` — añadir miembro

```csharp
        /// <summary>
        /// Renombra las solicitudes devolviendo el desglose completo (carpetas, motivos y contadores).
        /// </summary>
        /// <param name="folder">Carpeta raíz del libro (se procesa también la subcarpeta Solicitudes).</param>
        /// <param name="inicio">Número de viático inicial del rango (0 para sin filtro).</param>
        /// <param name="fin">Número de viático final del rango (0 para sin filtro).</param>
        /// <returns>Resultado con contadores y motivo por archivo; nunca null.</returns>
        ResultadoRenombrado RenombrarSolicitudesConDetalle(string folder, int inicio, int fin);
```

### 4.2 `IViaticosService` — añadir miembro

```csharp
        /// <summary> Renombra solicitudes devolviendo el desglose completo del proceso. </summary>
        ResultadoRenombrado RenombrarSolicitudesConDetalle(int inicio, int fin);
```

### 4.3 `ViaticosService` — implementación (debajo de `RenombrarSolicitudes`)

```csharp
        /// <inheritdoc/>
        public ResultadoRenombrado RenombrarSolicitudesConDetalle(int inicio, int fin)
            => _gestionProcessor.RenombrarSolicitudesConDetalle(_excelRepository.GetOneDriveFolderPath(), inicio, fin);
```

### 4.4 `FormGestionDocumental.BtnRenombrar_Click` — reemplazo completo

```csharp
        private void BtnRenombrar_Click(object sender, EventArgs e)
        {
            if (!TryLeerRango(out int inicio, out int fin))
            {
                return;
            }

            var r = _viaticosService.RenombrarSolicitudesConDetalle(inicio, fin);

            string mensaje;
            MessageBoxIcon icono;

            if (r.HayError)
            {
                mensaje = $"Ocurrió un error al renombrar solicitudes: {r.Error}";
                icono = MessageBoxIcon.Error;
            }
            else if (r.ComisionesEnRango == 0)
            {
                mensaje = $"No hay comisiones en CONSULTAS para el rango {inicio}-{fin}.";
                icono = MessageBoxIcon.Warning;
            }
            else if (r.Renombrados > 0)
            {
                mensaje = $"Finalizado. {r.Renombrados} archivo(s) renombrado(s).\r\n" +
                          $"Comisiones en rango: {r.ComisionesEnRango} · PDF examinados: {r.PdfsExaminados}\r\n" +
                          $"Pendientes: sin comisión {r.SinComision} · ambiguos {r.Ambiguos} · " +
                          $"colisiones {r.Colisiones} · ya correctos {r.YaCorrectos} · prefijo inválido {r.PrefijoInvalido}";
                icono = MessageBoxIcon.Information;
            }
            else
            {
                mensaje = $"Se encontraron {r.ComisionesEnRango} comisiones en el rango, pero no se renombró ningún archivo " +
                          $"(PDF examinados: {r.PdfsExaminados}).\r\n" +
                          $"Verifique que los PDF estén en la carpeta 'Solicitudes' y que su nombre inicie con el número de comisión.";
                icono = MessageBoxIcon.Warning;
            }

            MessageBox.Show(mensaje, App.APP_NAME, MessageBoxButtons.OK, icono);
        }
```

`BtnUnificar_Click` no se toca.

**Checklist de cierre de la Fase 1438**: build OK; `UnitTest1.cs` (usa `Mock<IViaticosService>`) sigue compilando; ningún otro implementador de las interfaces (verificado: solo `ViaticosService` y `ViaticosGestionProcessor`). Commit: `Fase 1438 (T2): resultado detallado y mensajes fieles en Gestión Documental`.

---

## 5. Fase 1439 (T3) — Dedupe `(Num, CC)` y política N0/N1/N2/N3

- **id**: T3 · **fase**: 1439 · **depends_on**: [T1] · **status**: pending · **log**: -
- **location**: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- **files_lock**: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- **description**: sustituir la coincidencia exacta por prefijo por un pipeline explícito: prefijo numérico robusto → dedupe `(Num, CC)` → metadatos → OCR → motivo.
- **validation**: `21626` y `23826` renombran **sin** invocar `IPdfToTexto`; una comisión con dos cédulas distintas sí lo invoca.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"`
- **evidence**: `_evidence/<timestamp>/summary.md`

### 5.1 `GetSolicitudes`: normalizar `Num` y `CC` (líneas 140–152 actuales)

```csharp
                string num = (data[i, ConstantsHelper.CONSULTAS_NUMERO_COMISION]?.ToString() ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(num)) continue;

                total++;
                if (!ViaticoNumeroComision.EnRango(num, inicio, fin)) continue;

                list.Add(new SInfo {
                    Num = num,
                    CC  = (data[i, ConstantsHelper.CONSULTAS_CEDULA]?.ToString() ?? string.Empty).Trim(),
                    Nom = Clean(data[i, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO]?.ToString() ?? string.Empty)
                });
```

### 5.2 Reemplazo de `RenombrarFila` (versión final)

```csharp
        /// <summary>
        /// Intenta identificar y renombrar un PDF aplicando la política de desambiguación:
        /// N0 candidato único por (Num, Cédula) → N1 metadatos → N2 OCR → N3 sin renombrar.
        /// </summary>
        /// <returns>Motivo del resultado (ver <see cref="MotivoRenombrado"/>).</returns>
        private string RenombrarFila(string file, List<SInfo> sols, string carpeta)
        {
            string oldN = Path.GetFileNameWithoutExtension(file);
            string ext = Path.GetExtension(file);

            // El nombre debe iniciar con el número de comisión (admite '-' o '_' como separador)
            Match prefijo = Regex.Match(oldN, @"^(\d+)");
            if (!prefijo.Success) return MotivoRenombrado.PrefijoInvalido;
            string num = prefijo.Groups[1].Value;

            // N0: un solo candidato por (Num, Cédula); las filas duplicadas idénticas colapsan aquí
            var candidatos = sols.Where(s => s.Num == num)
                                 .GroupBy(s => NormalizarDocumento(s.CC))
                                 .Select(g => g.First())
                                 .ToList();
            if (candidatos.Count == 0) return MotivoRenombrado.SinComision;

            SInfo elegido = candidatos.Count == 1 ? candidatos[0] : null;

            // N1: metadatos del propio PDF
            if (elegido == null) elegido = IdentificarPorMetadatos(file, candidatos);

            // N2: OCR de la primera página validado contra las cédulas candidatas
            if (elegido == null) elegido = IdentificarPorOcr(file, candidatos);

            // N3: nada se renombra a ciegas
            if (elegido == null) return MotivoRenombrado.Ambigua;

            string nuevoNombre = $"{elegido.Num}-{elegido.CC}-{elegido.Nom}{ext}";
            if (string.Equals(oldN + ext, nuevoNombre, StringComparison.OrdinalIgnoreCase))
                return MotivoRenombrado.YaCorrecto;

            string destino = Path.Combine(carpeta, nuevoNombre);
            if (File.Exists(destino)) return MotivoRenombrado.Colision;

            try
            {
                File.Move(file, destino);
                return MotivoRenombrado.Renombrado;
            }
            catch (Exception ex)
            {
                LoggerHelper.Warning($"RenombrarSolicitudes: no se pudo renombrar '{Path.GetFileName(file)}': {ex.Message}");
                return MotivoRenombrado.ErrorIo;
            }
        }
```

### 5.3 Helpers N1/N2 (agregar debajo de `RenombrarFila`)

```csharp
        /// <summary> N1: identifica al funcionario por los metadatos del PDF (cédula en PalabrasClave). </summary>
        private SInfo IdentificarPorMetadatos(string file, List<SInfo> candidatos)
        {
            try
            {
                var meta = _metadataService.LeerMetadatos(file);
                if (meta == null || meta.MetaDocumento != TipoMetadata.VIA_SOLICITUD || string.IsNullOrWhiteSpace(meta.PalabrasClave))
                    return null;

                string doc = NormalizarDocumento(meta.PalabrasClave.Split(',')[0]);
                return candidatos.FirstOrDefault(s => NormalizarDocumento(s.CC) == doc);
            }
            catch (Exception ex)
            {
                LoggerHelper.Debug($"RenombrarSolicitudes: metadatos ilegibles en '{Path.GetFileName(file)}': {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// N2: identifica al funcionario por la cédula leída del PDF.
        /// En esta fase usa <c>FindCC(string)</c>; la Fase 1440 lo sustituye por la versión validada.
        /// </summary>
        private SInfo IdentificarPorOcr(string file, List<SInfo> candidatos)
        {
            string doc = NormalizarDocumento(FindCC(file));
            if (string.IsNullOrEmpty(doc)) return null;
            return candidatos.FirstOrDefault(s => NormalizarDocumento(s.CC) == doc);
        }

        /// <summary> Quita separadores de miles y espacios de un número de documento. </summary>
        internal static string NormalizarDocumento(string valor)
        {
            return Regex.Replace(valor ?? string.Empty, @"[\s\.,]", string.Empty);
        }
```

**Checklist de cierre de la Fase 1439**
1. `SInfo` sigue siendo `internal class SInfo { public string Num, CC, Nom; }` (sin cambios).
2. `FindCC` conserva la firma `string FindCC(string f)` en esta fase.
3. Los 3 tests existentes pasan. Commit: `Fase 1439 (T3): dedupe (Num,CC) y motivos de renombrado`.

---

## 6. Fase 1440 (T4) — `FindCC` validado contra cédulas candidatas

- **id**: T4 · **fase**: 1440 · **depends_on**: [T3] · **status**: pending · **log**: -
- **location**: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- **files_lock**: `ConversorPDF.XLL/Core/Processors/ViaticosGestionProcessor.cs`
- **description**: eliminar el falso positivo del rubro (`0505023`): devolver solo una cédula que pertenezca al conjunto candidato.
- **validation**: con los 2 PDF ambiguos, `FindCC` devuelve `52371279` / `66955819`; con un texto que solo contiene `0505023` devuelve cadena vacía.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"`
- **evidence**: `_evidence/<timestamp>/summary.md`

### 6.1 Reemplazar `FindCC` (líneas 243–254 actuales)

```csharp
        /// <summary>
        /// Extrae de la primera página una cédula que pertenezca al conjunto de cédulas candidatas.
        /// Patrón 1: cédula seguida de correo (formato GF-FO-023). Patrón 2: cualquier secuencia
        /// numérica que coincida exactamente con una cédula candidata (evita tomar el rubro presupuestal).
        /// </summary>
        /// <param name="f">Ruta del PDF.</param>
        /// <param name="cedulasCandidatas">Cédulas normalizadas de las filas de CONSULTAS para esa comisión.</param>
        /// <returns>Cédula normalizada encontrada, o cadena vacía.</returns>
        private string FindCC(string f, ICollection<string> cedulasCandidatas)
        {
            if (cedulasCandidatas == null || cedulasCandidatas.Count == 0) return string.Empty;

            string t = _pdfToTexto.ExtraerTextoPagina(f, 1);
            if (string.IsNullOrWhiteSpace(t)) return string.Empty;

            foreach (Match m in Regex.Matches(t, @"(\d[\d\.,\s]{5,})\s+[a-zA-Z0-9._%+-]+@"))
            {
                string candidata = NormalizarDocumento(m.Groups[1].Value);
                if (cedulasCandidatas.Contains(candidata)) return candidata;
            }

            var halladas = Regex.Matches(t, @"\d[\d\.\s]{5,}\d")
                                .Cast<Match>()
                                .Select(m => NormalizarDocumento(m.Value))
                                .Where(c => cedulasCandidatas.Contains(c))
                                .Distinct()
                                .ToList();

            return halladas.Count == 1 ? halladas[0] : string.Empty;
        }
```

### 6.2 Ajustar `IdentificarPorOcr` (único punto de llamada)

```csharp
        private SInfo IdentificarPorOcr(string file, List<SInfo> candidatos)
        {
            var cedulas = new HashSet<string>(candidatos.Select(s => NormalizarDocumento(s.CC)));
            string doc = FindCC(file, cedulas);
            if (string.IsNullOrEmpty(doc)) return null;
            return candidatos.FirstOrDefault(s => NormalizarDocumento(s.CC) == doc);
        }
```

Verificar que **no exista** otra llamada a `FindCC(` con un solo argumento: `Select-String -Pattern 'FindCC\\('`.

**Checklist de cierre de la Fase 1440**: build OK; `_pdfToTexto.ExtraerTextoPagina` se invoca como máximo una vez por archivo ambiguo. Commit: `Fase 1440 (T4): FindCC validado contra cédulas candidatas`.

---

## 7. Fase 1441 (T5) — Pruebas de regresión

- **id**: T5 · **fase**: 1441 · **depends_on**: [T1, T2, T3, T4] · **status**: pending · **log**: -
- **location**: `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`
- **files_lock**: `ConversorPDF.Tests/Core/Processors/ViaticosGestionProcessorTests.cs`
- **description**: 7 pruebas nuevas sobre el fixture existente (`_sut` con `Mock<IExcelRepository>`, `Mock<IViaticosRepository>`, `Mock<IPdfToTexto>`, `Mock<IPdfMetadataService>`).
- **validation**: 10 pruebas en la clase (3 existentes + 7 nuevas) en verde; `TOTAL` de la suite ≥ baseline 952.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**: `_evidence/<timestamp>/test.trx`

### 7.1 Usings a añadir al archivo de pruebas

```csharp
using System.IO;
using ConversorPDF.XLL.Core.Models;
```

### 7.2 Helper de datos y de carpetas (agregar dentro de la clase)

```csharp
        private static object[,] CrearConsultas(params (string Num, string Cc, string Nom)[] filas)
        {
            var data = Array.CreateInstance(typeof(object), new int[] { filas.Length + 1, 35 }, new int[] { 1, 1 }) as object[,];
            for (int i = 0; i < filas.Length; i++)
            {
                data[i + 2, ConstantsHelper.CONSULTAS_NUMERO_COMISION] = filas[i].Num;
                data[i + 2, ConstantsHelper.CONSULTAS_CEDULA] = filas[i].Cc;
                data[i + 2, ConstantsHelper.CONSULTAS_NOMBRE_COMPLETO] = filas[i].Nom;
            }
            return data;
        }

        private static string CrearCarpetaTemporal()
        {
            string root = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(root);
            return root;
        }
```

### 7.3 Pruebas

```csharp
        [Fact]
        public void RenombrarSolicitudes_WhenPdfEnSubcarpetaSolicitudes_ShouldRenombrar()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "17526-Obert-Alejandro-OKEY.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("17526", "79633812", "OBERT ALEJANDRO ORTIZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.HayError.Should().BeFalse();
                r.PdfsExaminados.Should().Be(1);
                r.Renombrados.Should().Be(1);
                File.Exists(Path.Combine(pSoli, "17526-79633812-ObertAlejandroOrti.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenComisionDuplicadaConMismaCedula_ShouldRenombrarSinOcr()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "21626-VIATICOS -YENY RAYO.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(
                              ("21626", "52371279", "YENY PATRICIA RAYO PRADA"),
                              ("21626", "52371279", "YENY PATRICIA RAYO PRADA")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(1);
                _mockPdfToTexto.Verify(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()), Times.Never());
                File.Exists(Path.Combine(pSoli, "21626-52371279-YenyPatriciaRayoPr.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenComisionDuplicadaConCedulasDistintas_ShouldUsarOcr()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "14526-Solicitud.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(
                              ("14526", "12345678", "JUAN PEREZ"),
                              ("14526", "87654321", "ANA GOMEZ")));
                _mockPdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                               .Returns("C.C. 12.345.678  juan.perez@esap.edu.co  Rubro C-0505-1000-4-20306D-0505023-02");

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(1);
                _mockPdfToTexto.Verify(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()), Times.AtLeastOnce());
                File.Exists(Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf")).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }
```

```csharp
        [Fact]
        public void RenombrarSolicitudes_WhenOcrDevuelveNumeroAjeno_ShouldMarcarAmbiguo()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                string origen = Path.Combine(pSoli, "14526-Solicitud.pdf");
                File.WriteAllText(origen, "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(
                              ("14526", "12345678", "JUAN PEREZ"),
                              ("14526", "87654321", "ANA GOMEZ")));
                _mockPdfToTexto.Setup(x => x.ExtraerTextoPagina(It.IsAny<string>(), It.IsAny<int>()))
                               .Returns("Rubro C-0505-1000-4-20306D-0505023-02");

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Renombrados.Should().Be(0);
                r.Ambiguos.Should().Be(1);
                File.Exists(origen).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenDestinoYaExiste_ShouldMarcarColisionYNoSobrescribir()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                string origen = Path.Combine(pSoli, "14526-juan.pdf");
                string destino = Path.Combine(pSoli, "14526-12345678-JuanPerez.pdf");
                File.WriteAllText(origen, "origen");
                File.WriteAllText(destino, "destino");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("14526", "12345678", "JUAN PEREZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.Colisiones.Should().Be(1);
                r.Renombrados.Should().Be(0);
                File.ReadAllText(destino).Should().Be("destino");
                File.Exists(origen).Should().BeTrue();
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenNombreSinPrefijoNumerico_ShouldMarcarPrefijoInvalido()
        {
            string root = CrearCarpetaTemporal();
            try
            {
                string pSoli = Path.Combine(root, App.FolderNames.SOLICITUDES);
                Directory.CreateDirectory(pSoli);
                File.WriteAllText(Path.Combine(pSoli, "ADOLFO-LEON-GUERRERO.pdf"), "x");
                _mockExcel.Setup(x => x.GetWorksheetData(ConstantsHelper.SHEET_CONSULTAS))
                          .Returns(CrearConsultas(("14526", "12345678", "JUAN PEREZ")));

                var r = _sut.RenombrarSolicitudesConDetalle(root, 0, 0);

                r.PrefijoInvalido.Should().Be(1);
                r.Renombrados.Should().Be(0);
            }
            finally { Directory.Delete(root, true); }
        }

        [Fact]
        public void RenombrarSolicitudes_WhenCarpetaNoExiste_ShouldRetornarError()
        {
            string inexistente = Path.Combine(Path.GetTempPath(), "ConversorPDF.Tests", Guid.NewGuid().ToString("N"));

            var r = _sut.RenombrarSolicitudesConDetalle(inexistente, 0, 0);

            r.HayError.Should().BeTrue();
            _sut.RenombrarSolicitudes(inexistente, 0, 0).Should().Be(-1);
        }
```

**Checklist de cierre de la Fase 1441**
1. Invariante que deben cumplir todas las pruebas: `Renombrados + YaCorrectos + Colisiones + Ambiguos + SinComision + PrefijoInvalido + ErroresIo == PdfsExaminados`.
2. `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full` → exit 0.
3. Commit: `Fase 1441 (T5): pruebas de regresión de renombrado (7 casos)`.

---

## 8. Fase 1442 (T6) — Documentación, cierre y versión

- **id**: T6 · **fase**: 1442 · **depends_on**: [T5] · **status**: pending · **log**: -
- **location**: `docs/articles/gestion-documental.md`
- **files_lock**:
  - `docs/articles/gestion-documental.md`
  - `docs/articles/changes/2026-10-31-gestion-documental-carpeta-y-desambiguacion.md` (nuevo)
  - `docs/api/ConversorPDF.XLL.Core.Interfaces.IViaticosGestionProcessor.yml`
  - `.agents/plans/REGISTRY.md`
  - `ConversorPDF.XLL/Core/Helpers/App.cs` **(solo con el bump de versión; hot_path L0 ⇒ serializado)**
  - `installer/InstaladorConversorPDF.wxs` **(solo con el bump)**
- **description**: documentar carpeta + política N0–N3 + motivos; changelog; sincronizar el YAML de API; sellar `R-003`. Opcional: `APP_VERSION = "1.6.3.83"` y versión en el `.wxs`.
- **validation**: `gestion-documental.md` describe la carpeta y la política; `R-003` en `closed` con contadores en 268 / 1442.
- **verification**: `powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full`
- **evidence**: `_evidence/<timestamp>/summary.md`

### 8.1 Contenido mínimo a añadir en `gestion-documental.md`

1. El renombrado escanea **la carpeta del libro y la subcarpeta `Solicitudes`** (no recursivo más allá de esas dos).
2. Política **N0 → N1 → N2 → N3** (tabla del §9 de este plan).
3. Tabla de motivos: `RENOMBRADO`, `YA_CORRECTO`, `COLISION`, `AMBIGUA`, `SIN_COMISION_EN_CONSULTAS`, `PREFIJO_INVALIDO`, `ERROR_IO`.
4. Sección "Datos que debe tener `CONSULTAS`": una fila por comisión (las filas idénticas se deduplican; dos personas en la misma comisión exigen metadatos/OCR).
5. Operación: el resumen queda en el log como `[RENOMBRAR] carpetas=… | comisionesEnRango=… | pdfs=… | renombrados=…`.

### 8.2 Changelog y cierre

- Crear `docs/articles/changes/2026-10-31-gestion-documental-carpeta-y-desambiguacion.md` con la misma estructura que `2026-10-30-gestion-documental-rango-yy.md` (diagnóstico D1–D8, cambios por fase 1437–1442, pruebas, criterios).
- `docs/api/...IViaticosGestionProcessor.yml`: añadir el `uid` del nuevo miembro `RenombrarSolicitudesConDetalle(System.String,System.Int32,System.Int32)` (o regenerar con DocFX si el pipeline lo hace).
- `REGISTRY.md`: fila `R-003` → `closed`, `verify` con el exit code del gate, `HEAD` con el commit del cierre; actualizar §2 (`ronda` 268, `fase` 1442).
- Bump de versión: `App.APP_VERSION = "1.6.3.83"` ejecutado en Ronda 272 (Fase 1461). No compilar MSI hasta cierre con gate verde de R-003 y R-005.

---

## 9. Política de desambiguación (texto de referencia para docs y revisión)

| Nivel | Regla implementada | Cuándo actúa | Resultado |
|---|---|---|---|
| **N0** | `sols.Where(Num==prefijo).GroupBy(NormalizarDocumento(CC))`; si queda 1 candidato | Siempre, primero | Renombra sin OCR (cubre `21626`, `23826`, que tienen filas idénticas) |
| **N1** | `metaDocumento == VIA_SOLICITUD` y la cédula es el primer término de `PalabrasClave` | PDF marcados por la app | Renombra sin OCR |
| **N2** | `FindCC` devuelve solo una cédula **perteneciente al conjunto candidato** | PDF sin metadatos y comisión con >1 persona | Renombra si hay coincidencia única |
| **N3** | Sin coincidencia ⇒ no se toca el archivo; se registra el motivo | N0–N2 no concluyen | `AMBIGUA` en log y en `Detalle` |

Motivo por archivo (clave = nombre original) disponible en `ResultadoRenombrado.Detalle`, lo que permite auditar sin revisar el log.

---

## 10. Gate de aceptación (medición real, tras T5)

Ejecución **en seco sobre una copia** antes del dato real:

```powershell
$one = (Get-ChildItem $env:USERPROFILE -Directory | Where-Object { $_.Name -like 'OneDrive - Escuela Superior*' } | Select-Object -First 1).FullName
$src = Join-Path $one '__2026\__VIATICOS\SOPORTE VIATICOS 2026'
$sandbox = Join-Path $env:TEMP ('sandbox_viaticos_' + (Get-Date -Format 'yyyyMMdd_HHmmss'))
New-Item -ItemType Directory -Path $sandbox | Out-Null
Copy-Item -Path (Join-Path $src '*') -Destination $sandbox -Recurse -Force
# Luego, con Excel abierto en el libro real, pulsar RENOMBRAR SOLICITUDES usando el rango 126-41326
# y contrastar el mensaje con la tabla siguiente.
```

| Métrica | Estimado de referencia (muestra preliminar) |
|---|---|
| `ComisionesEnRango` | 416 |
| `PdfsExaminados` | 434 (419 en `Solicitudes` + 15 en raíz) |
| `Renombrados` | ~231 (sujeto a validación real sobre universo total) |
| `YaCorrectos` | 1 |
| `Colisiones` | 0 |
| `Ambiguos` | **0** |
| `SinComision` | 28 (25 en `Solicitudes` + 3 en raíz) |
| `PrefijoInvalido` | 11 |
| `ErroresIo` | 0 |

**Invariante obligatoria rectora (criterio principal de aceptación)** (si el dataset cambia, la igualdad debe cumplirse y se valida en `Detalle` que cambios de archivos con patrón preexistente sean legítimos):
`Renombrados + YaCorrectos + Colisiones + Ambiguos + SinComision + PrefijoInvalido + ErroresIo == PdfsExaminados`.

Verificación del log: debe existir una línea `[RENOMBRAR] ... renombrados=231 ...` y una línea `Debug` por archivo.

---

## 11. Riesgos, rulings y rollback

| Riesgo | Mitigación definida |
|---|---|
| Renombrar archivos reales del usuario en la primera prueba | Ejecutar primero sobre `$sandbox` (copia); el dato real solo tras revisar el desglose |
| Pérdida de datos por colisión | `RenombrarFila` **nunca** sobrescribe: si el destino existe devuelve `COLISION` y deja ambos archivos |
| Escanear dos veces la misma carpeta | `CarpetasDeSolicitudes` compara rutas con `OrdinalIgnoreCase` y solo añade `Solicitudes` si existe y difiere |
| `Estandarizar` tocando archivos que no son PDF | Restringido a `*.pdf` (D6) |
| OCR caro en comisiones duplicadas | N0 deduplica antes de cualquier lectura; N1 es instantáneo |
| Tests escribiendo en el log de producción | `TestLogBootstrap.cs` ya redirige a `%TEMP%\ConversorPDF.Tests\log` |
| Bump de versión colisionando con `bin/` | Va al final, serializado, con `verify.ps1 -Mode Full` |

**Rollback**: los cambios son de código (6 archivos) + 2 modelos nuevos; revertir con `git checkout -- <archivos>` y eliminar los 2 archivos nuevos. Los renombrados de archivos **no** son reversibles automáticamente: por eso la validación se hace primero en `$sandbox`.

---

## 12. Anexo A — Tareas de datos que el código NO resuelve

| Grupo | Cantidad | Acción |
|---|---|---|
| PDF con comisión `41426`–`43626` | 23 | Actualizar la hoja `CONSULTAS` (hoy termina en `41326`) y volver a ejecutar. Sin fila en `CONSULTAS` no hay nombre destino posible. |
| PDF con prefijo de 6 dígitos (`202426-…`, `202626-…`) | 2 | Verificar en SIIF el número real y renombrar a `<comisión>-<cédula>-<nombre>.pdf`. No se adivina automáticamente. |
| PDF de la raíz con nombre libre (`ADOLFO-…`, `[2]-CARLOS-…`, `Javier-zorrilla-…`) | 11 | Permanecen intactos y se reportan como `PREFIJO_INVALIDO`; requieren identificación manual si se quieren estandarizar. |

---

## 13. Resumen de ejecución para el modelo implementador

```powershell
# 0) Contexto y gate base
git status --short
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace . -Plan .agents/plans/2026-10-31-gestion-documental-carpeta-solicitudes-fix-plan.md
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full

# Wave 1 — Ronda 266
#   T1 (Fase 1437): modelos + RenombrarSolicitudesConDetalle + CarpetasDeSolicitudes + Contabilizar + Estandarizar(*.pdf) + RenombrarFila->string
#   T2 (Fase 1438): interfaces + ViaticosService + FormGestionDocumental
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"

# Wave 2 — Ronda 267 (serializada: mismo archivo)
#   T3 (Fase 1439): dedupe (Num,CC) + IdentificarPorMetadatos/Ocr + NormalizarDocumento + GetSolicitudes.Trim
#   T4 (Fase 1440): FindCC(f, cedulasCandidatas)
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Fast -Filter "FullyQualifiedName~ViaticosGestionProcessorTests"

# Wave 3 — Ronda 268
#   T5 (Fase 1441): 7 pruebas
powershell -ExecutionPolicy Bypass -File scripts/verify.ps1 -Mode Full
#   T6 (Fase 1442): docs + changelog + REGISTRY cerrado (+ bump opcional)
```

**Orden estricto**: 1437 → 1438 → 1439 → 1440 → 1441 → 1442. Aunque `files_lock` de T1 y T2 es disjunto, T2 usa `RenombrarSolicitudesConDetalle`, por lo que **T2 arranca cuando T1 compile**; T3 y T4 comparten archivo y van serializadas.

### Grupos de ejecución paralela
| Wave | Ronda | Tasks | Can start when |
|---|---|---|---|
| 1 | 266 | T1 (L0) | Inmediato |
| 1 | 266 | T2 | T1 compila (`verify.ps1 -Mode Build` = 0) |
| 2 | 267 | T3 (serializada) | T1 y T2 completadas |
| 2 | 267 | T4 | T3 compila |
| 3 | 268 | T5, T6 | T4 completada (T6 requiere T5) |


**Reglas para el implementador**
1. No re-numerar ni reasignar; `R-003` ya existe.
2. Un commit por Fase, con el mensaje indicado en cada checklist.
3. No `push`, no `merge`, no publicación remota.
4. Si un extracto no coincide con el archivo real: detenerse y reportar (posible cambio previo del árbol).
5. Toda evidencia va a `_evidence/<timestamp>/` (ignorado por git).
6. Los tests no deben tocar OneDrive ni `%LOCALAPPDATA%`.
7. Al cerrar: `REGISTRY.md` §2 con `ronda=268`, `fase=1442` y `R-003` en `closed`.

**Definition of Done**
- [ ] 6 Fases ejecutadas y selladas en este documento (`status: done`).
- [ ] `verify.ps1 -Mode Full` → exit 0 con `TOTAL` ≥ baseline medido.
- [ ] Renombrado real con `Renombrados = 231`, `Ambiguos = 0`, `Colisiones = 0`, `ErroresIo = 0`.
- [ ] Mensaje de UI con desglose (nunca "no se encontraron comisiones" cuando `ComisionesEnRango > 0`).
- [ ] `gestion-documental.md` + changelog + `REGISTRY.md` actualizados.











