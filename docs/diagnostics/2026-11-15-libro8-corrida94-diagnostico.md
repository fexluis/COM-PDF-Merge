# Libro8 y corrida #94: diagnóstico

Revisión de datos y código, 2026-11-15. Sin compilar, ejecutar AG/pruebas ni modificar Excel o base de datos.

Actualización: Causa, Acción y código propuesto. Los bloques siguientes son una propuesta para implementar; no son cambios aplicados ni código compilado. No reservan rondas/fases ni reabren R-078.

## Evidencia

- Excel: `C:/Users/USUARIO/Downloads/Libro8.xlsx`, hoja `Semestres`, A1:DM138. `Hoja1` vacía.
- Log: `C:/Users/USUARIO/AppData/Local/ConversorPDF/log/log_20261115.txt`, especialmente 13:12:17–13:13:39.
- Base temporal de la misma sesión: `C:/Users/USUARIO/AppData/Local/Temp/esap_acad_tmp_AcademicData_VALLEdb_2b841e3f.db`. Consultas SQLite con `mode=ro` y `PRAGMA query_only=ON`.
- Corrida 94, individuo óptimo 44, subunidad VALLE, versión registrada 1.6.4.15.

## 1. Recuperación de ranuras comprobada en CALI III/IV

| Cohorte | Materia final | Horas | Fechas efectivas | Filas Excel |
|---|---|---:|---|---|
| CALI III-A | Teorías de las organizaciones públicas | 48 | 26/04–13/05/2027 | 36 |
| CALI III-B | Teorías de las organizaciones públicas | 48 | 26/04–13/05/2027 | 44 |
| CALI IV-A | Presupuesto público | 48 | 26/04–19/05/2027 | 60 |
| CALI IV-B | Presupuesto público | 48 | 26/04–19/05/2027 | 68 |

Las cuatro usan varios días de lunes a viernes y huecos reales. III tiene 12 sesiones por materia final, incluidas fechas de abril y mayo; IV también tiene 12 y evita los miércoles/jueves todavía ocupados por Planeación. El patrón del fallo anterior —última materia confinada al viernes y extensión hasta julio— no aparece en estas cuatro cohortes.

Excel: 136 filas, 5952 horas; horas de cada fila coinciden con la suma de sus celdas diarias. Última fecha exportada: 20/05/2027. No contar columna de total como otra fecha.

La corrida completa incluye dos programas: 184 vinculaciones, 2064 sesiones, 8256 horas; rango de sesiones 05/02–30/05/2027. No confundir el alcance filtrado del Excel con el total de staging.

## 2. No se ejecutó con el límite original

El snapshot persistido de la corrida 94 registra ventana **05/02–30/06/2027**, origen Manual, 146 días. El log registra ajustes sucesivos de 115, 117 y 146 días antes de iniciar esta corrida.

Aunque las sesiones resultantes terminan a más tardar el 30/05, esto no certifica una ejecución íntegra aceptada con ese límite como entrada.

El intento de las 13:12:23 se abortó por:

- CALI I-A, aula 13: demanda 64, capacidad MAR-VIE 61.
- CALI I-B, aula 14: demanda 64, capacidad MAR-VIE 61.

Ambas aulas figuran en la base con `Estado=4 (Aprobado)` y `DiasVentana=MAR-VIE`. `AgEscaladorRanuraAula.EsAulaElegible` solo admite Disponible; el validador, en cambio, incluyó esas aulas y bloqueó la corrida.

Con los festivos del snapshot y el límite original, cálculo independiente de capacidad: MAR-VIE=61 y LUN-VIE=75 sesiones de 4 h. Por capacidad local, el siguiente escalón proporciona espacio suficiente para 64 sesiones; falta simular las demás restricciones.

**Causa:** validación incluye aulas aprobadas, pero escalado solo admite Disponible. El usuario recibe una sugerencia de escalado que ese flujo no puede aplicar al aula bloqueada y termina ampliando fechas.

**Acción:** mantener protegida la aprobación; devolver bloqueo específico para aulas aprobadas con déficit. Su reapertura requiere una transición explícita autorizada mediante el flujo de aprobación, seguida del escalado normal. No convertir Aprobado→Disponible automáticamente ni excluir aulas silenciosamente. Código en A1. Repetir con 30/05 como entrada, sin codificar esa fecha dentro del algoritmo.

## 3. HARD calculados sobre rangos amplios: falsos positivos

`AgValidadorReglas.cs`, líneas 272–376: R1/R2/R3/R4 agrupan por bloque/día/franja y comparan `FechaInicioEstimada`–`FechaFinEstimada`; no intersectan `FechasGen`.

Esto no representa correctamente una materia que aprovecha fechas intercaladas entre otras materias.

Ejemplo comprobado, CALI III-A, aula 36:

- Estadística I: viernes 05/02–30/04, bloque 30.
- Teorías: rango general 26/04–13/05, pero su único viernes efectivo es **07/05**.
- Reglas persistidas: R2 «aula 36 asignada a 2 cursos en bloque 30» y R4 «dos asignaturas día 5, TARDE».
- Sesiones: Estadística termina el viernes 30/04; Teorías usa viernes 07/05. No hay cruce de viernes.

Contraste independiente de las 2064 sesiones de staging, por fecha e intervalo horario: **0 pares con solapamiento dentro de una misma aula**. Esto demuestra que no se deben interpretar todas las alarmas de aula como cruces reales de sesiones.

**Causa:** los rangos generales son envolventes de sesiones discontinuas. Además, agrupar por ID de bloque no detecta bloques diferentes cuyos intervalos reales se superponen. R7 interna ni siquiera compara fechas: agrupa docente/bloque en todo el semestre.

**Acción:** proyectar sesiones con el expansor existente y comparar fecha e intervalo real. Sustituir las comprobaciones temporales de R1/R2/R3/R4/R7; conservar habilitación, cobertura y demás reglas. Cuando no haya FechasGen, el expansor genera las fechas por la misma ruta usada al materializar. Si no puede proyectar, informar error estructural; no asumir agenda libre. Código en A2.

## 4. También existen conflictos docentes reales

Contraste independiente entre vinculaciones distintas, excluyendo sesiones vacantes: **12 pares de sesiones con mismo docente, fecha e intervalo**. No equivalen a 12 registros de reglas: son pares de sesiones concretas.

- Docente 193, Carlos Antonio Espinosa Dorado, Seminario Electivo II: aulas 50 y 56, ocho sábados 08:00–12:00 entre 06/02 y 03/04, excluido 27/03.
- Docente 9, Miguel Ángel Cerón Hurtado: aulas 52 y 15, cuatro jueves 08:00–12:00, 15/04, 22/04, 29/04 y 06/05; Gerencia Pública Integral frente a Pensamientos y teorías de la administración pública.

**Causa:** `AgAsignadorTercero.DocenteDisponible` y `BuscarAlterno` comparan ID de bloque sin fechas; la ocupación externa se consulta solo en `FechaInicioEstimada`. Esto combina rechazos ficticios con omisiones de choques posteriores. Una comprobación de un solo gen tampoco garantiza disponibilidad en toda la asignatura de varios días.

**Acción:** probar candidatos para todos los genes de la misma asignatura/aula, contra cada fecha interna y externa. Cambiar docente de todo el grupo de forma coherente, respetando genes Fijado. Integrar la misma comprobación en siembra, reparación, cruce y mutación; validar de nuevo al finalizar. Si no existe candidato, mantener déficit explícito, sin ocultarlo mediante una asignación incompatible. Código en A3. Una clase compartida requiere autorización/modelo explícito; repetir asignatura no acredita esa excepción.

## 5. Vacantes y estado de éxito

- Staging: **440 sesiones vacantes, 1760 horas**, motivo persistido `SIN_DISPONIBILIDAD`; 1624 sesiones no vacantes, 6496 horas.
- Excel filtrado: 37 filas VACANTE, 1648 horas. Es otro alcance.
- Log de construcción: 29/102 asignaturas sin pool. Es un conteo por asignatura del contexto, no por sesión; no es directamente comparable con 440.
- Corrida: `vacantes=0` en `ag_corrida`, pese a 440 sesiones vacantes en staging.
- Se registra «Completada», «APROBADO» y «Optimización exitosa» con `HardViolations=197`.

**Causa:** UI usa `stagingOk` para anunciar éxito, independientemente de HARD/vacantes/cobertura. No asigna `corrida.Vacantes` y `UpdateCorridaAsync` tampoco incluye esa columna. `AgIndividuo.Vacantes` cuenta genes, mientras que el grid cuenta sesiones.

**Acción:** conservar la unidad actual de `AgCorrida.Vacantes` como genes vacantes, persistirla y mostrar separadamente sesiones/horas vacantes. Completada describe ejecución técnica; no equivale a Aprobada. Promoción exige cero HARD, cero vacantes y cobertura completa, comprobados también en el servicio de promoción. Código en A4. Staging parcial puede conservarse para diagnóstico, claramente identificado.

## 6. Detalle de reglas y advertencia de resultados

Registros persistidos del individuo 44:

| Regla | Registros |
|---|---:|
| R2 Superposición aula | 51 HARD |
| R3 Superposición grupo | 51 HARD |
| R4 Dos asignaturas mismo día | 35 HARD |
| R7 Docente en otra cohorte | 9 HARD |
| R1 Superposición docente | 4 HARD |
| S2 Desplazamiento entre sedes | 58 SOFT |
| R5 Habilitación docente | 1 SOFT |

Total: 209 registros, 150 HARD y 59 SOFT. `HardViolations=197` no implica necesariamente 197 filas: el código suma `cluster.Count-1` por algunos registros. `soft=6225` es una penalización, no cantidad de registros.

Por tanto, la advertencia del log que compara 197/6225 con 209 filas no prueba por sí sola pérdida de detalle.

**Causa del aviso incorrecto:** `CargarModuloResultadosAsync` suma conteo HARD y puntuación SOFT y compara ese resultado con cantidad de filas. Son unidades diferentes.

**Acción:** mostrar conteos de registros y puntuaciones en campos distintos. Comparar memoria y persistencia con la misma unidad; un fallo de lectura debe mostrarse como fallo de lectura. Código en A5.

**Causa del visor vacío, confirmada:** las **209 violaciones tienen `id_gen IS NULL`**. `FormVerVinculaciones.CargarViolacionesAsync` hace `INNER JOIN ag_sesiones ON s.idgen=v.id_gen`, por lo que obtiene cero filas aunque existan los mensajes. La consulta equivalente devuelve cero en esta copia.

**Acción:** cargar primero todas las violaciones del individuo óptimo de la corrida, sin exigir sesión enlazada. Mostrar las globales aparte; para nuevas violaciones guardar genes implicados y fecha concreta. Resolver IDs después de insertar genes, antes de insertar violaciones y dentro de la misma transacción. Código en A6. No asignar una violación antigua sin fecha a todas las sesiones del gen.

## 7. Orden recomendado

1. Aclarar alcance de aulas aprobadas y escalado MAR-VIE→LUN-VIE para CALI I-A/B, conservando 30/05 como entrada.
2. Unificar reglas y disponibilidad sobre sesiones fechadas; regresión específica de Estadística/Teorías sin cruce real.
3. Resolver los cruces docentes efectivos y revisar vacantes con disponibilidad recalculada.
4. Corregir estado de éxito, conteo de vacantes y advertencias de detalle.
5. Tras autorización: pruebas pertinentes y corrida funcional con snapshot reproducible y límite original. DEBUG estuvo desactivado en #94; activar diagnóstico para comprobar alternativas y presupuesto si se reproduce.

**Conclusión:** recuperación de ranuras comprobada en CALI III/IV; corrida global todavía no certificable. El siguiente trabajo debe corregir alcance, validación temporal y asignación docente, sin volver a confinar automáticamente las materias a carriles de dos ranuras.

## A1. Bloqueo específico de aula aprobada con déficit

En `AgGeneradorHorarios`, añadir este método y llamarlo antes del pre-vuelo. Aplicar la misma validación en diagnóstico UI para mostrar la causa antes de proponer escalado. Usar la ventana efectiva recibida, nunca una constante de mayo.

```csharp
private static Result ValidarAprobadas(AgContexto contexto)
{
    if (contexto.Ventana == null)
        return Result.Failure("I1_VENTANA_NO_DEFINIDA");

    foreach (var aula in contexto.Aulas.Where(a => a.Estado == AgEstadoAula.Aprobado))
    {
        var programa = contexto.Programas.FirstOrDefault(p => p.Id == aula.IdPrograma);
        double unidad = AgUnidadesSemanales.HorasSesionAula(aula,
            AgNivelAcademicoCatalogo.Desde(programa?.Modalidad));
        var evaluacion = AgEvaluadorDemandaVentana.Evaluar(
            aula.DiasVentana, aula.Jornada, unidad, contexto, aula,
            contexto.Ventana.FechaInicio, contexto.Ventana.FechaFin);
        if (!string.IsNullOrWhiteSpace(evaluacion.Error))
            return Result.Failure(evaluacion.Error);
        if (!evaluacion.CabeEnPlazo)
            return Result.Failure(
                $"I5_AULA_APROBADA_CON_DEFICIT: {aula.AulaRef}; " +
                "requiere revisión explícita de aprobación para replanificar. " +
                "El escalador no modifica aulas aprobadas.");
    }
    return Result.Success();
}

// En EjecutarOptimizacionAsync, antes de ValidarFactibilidadAsync:
var aprobadas = ValidarAprobadas(contexto);
if (aprobadas.IsFailure)
    return Result.Failure<AgResultadoCorrida>(aprobadas.Error);
```

No agregar `Aprobado` a `EsAulaElegible` como atajo. Tras reapertura autorizada, conservar huellas, simulación, confirmación y transacción del escalador existente. Una ampliación manual debe seguir identificada como Manual, sin presentarla como cumplimiento del plazo original.

## A2. Autoridad común de ocupación fechada

Archivo nuevo propuesto: `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCalendarioConflictos.cs`. Añadir su `Compile Include` al proyecto si mantiene listado explícito. Usings: `System`, `System.Collections.Generic`, `System.Linq`, `System.Globalization`, `CSharpFunctionalExtensions` y los namespaces Academic de Models/Genetica.

```csharp
internal static class AgCalendarioConflictos
{
    internal sealed class Ocupacion
    {
        internal AgGen Gen;
        internal DateTime Fecha;
        internal TimeSpan Inicio, Fin;
        internal string Franja;
    }

    internal static Result<List<Ocupacion>> Crear(
        IEnumerable<AgGen> genes, AgContexto contexto)
    {
        var resultado = new List<Ocupacion>();
        var expansor = new AgExpansorSesiones();
        foreach (var gen in genes)
        {
            if (gen.HorasGen <= 0.0) continue; // Misma semántica que el expansor.
            if (!contexto.BloquesPorId.TryGetValue(gen.IdBloqueHorario, out var bloque)
                || bloque.Activo != 1
                || !TimeSpan.TryParse(bloque.HoraInicio, CultureInfo.InvariantCulture, out var ini)
                || !TimeSpan.TryParse(bloque.HoraFin, CultureInfo.InvariantCulture, out var fin)
                || ini >= fin)
                return Result.Failure<List<Ocupacion>>(
                    $"I6_BLOQUE_INVALIDO: aula={gen.IdAula}, asig={gen.IdAsignatura}");

            var sesiones = expansor.Expandir(gen, 0, contexto);
            if (sesiones.IsFailure)
                return Result.Failure<List<Ocupacion>>(sesiones.Error);
            foreach (var sesion in sesiones.Value)
            {
                if (!sesion.FechaSesion.HasValue)
                    return Result.Failure<List<Ocupacion>>("I5_SESION_SIN_FECHA");
                resultado.Add(new Ocupacion {
                    Gen = gen, Fecha = sesion.FechaSesion.Value.Date,
                    Inicio = ini, Fin = fin, Franja = bloque.Franja ?? string.Empty
                });
            }
        }
        return Result.Success(resultado);
    }

    internal static bool SeCruzan(Ocupacion a, Ocupacion b) =>
        a.Fecha == b.Fecha && a.Inicio < b.Fin && b.Inicio < a.Fin;

    internal static string Cohorte(AgGen g)
    {
        string semestre = AgFormatoAsignatura.SemestreRomano(g.Semestre);
        if (string.IsNullOrWhiteSpace(semestre) || semestre == "SIN ASIGNAR")
            semestre = (g.Semestre ?? string.Empty).Trim();
        return $"{g.IdSede}|{g.IdPrograma}|{semestre}|{(g.Grupo ?? string.Empty).Trim()}";
    }

    internal static bool DocenteReal(AgGen g, AgContexto c) =>
        g.EsVacante == 0 && g.IdTercero > 0 && g.IdTercero != c.IdTerceroVacante;
}
```

En `AgValidadorReglas.Evaluar`, reemplazar bloques temporales anteriores, no agregarlos en paralelo. El siguiente bloque incluye R1/R2/R3/R7 interna y evidencia por par/fecha. Conservar el resto de las reglas y su cierre actual de contadores/estado.

```csharp
var proyeccion = AgCalendarioConflictos.Crear(genes, contexto);
if (proyeccion.IsFailure)
{
    individuo.Violaciones.Add(new AgViolacion(
        AgReglaViolacion.R14_BloqueHorarioNoAsignado, true,
        "Calendario no verificable: " + proyeccion.Error));
    individuo.HardViolations = hardCount + 1;
    individuo.SoftViolations = softScore;
    individuo.Estado = AgEstadoIndividuo.Infactible;
    return; // Fallo estructural: nunca interpretar como ausencia de conflictos.
}
var ocupaciones = proyeccion.Value;
var serializador = new System.Web.Script.Serialization.JavaScriptSerializer();

void Registrar(AgReglaViolacion regla,
    IEnumerable<AgCalendarioConflictos.Ocupacion> implicadas, string mensaje)
{
    var lista = implicadas.ToList();
    var indices = lista.Select(o => genes.IndexOf(o.Gen)).Distinct().ToArray();
    string detalle = serializador.Serialize(new {
        indicesGenes = indices,
        fecha = lista[0].Fecha.ToString("yyyy-MM-dd"),
        intervalos = lista.Select(o => new {
            o.Gen.IdAula, o.Gen.IdAsignatura, o.Gen.IdTercero,
            inicio = o.Inicio.ToString(), fin = o.Fin.ToString()
        }).ToArray()
    });
    individuo.Violaciones.Add(new AgViolacion(regla, true, mensaje,
        detalleJson: detalle)); // IDs definitivos se resuelven en A6.
    hardCount++;
}

foreach (var dia in ocupaciones.GroupBy(o => o.Fecha))
{
    var lista = dia.ToList();
    for (int i = 0; i < lista.Count; i++)
    for (int j = i + 1; j < lista.Count; j++)
    {
        var a = lista[i]; var b = lista[j];
        if (!AgCalendarioConflictos.SeCruzan(a, b)) continue;
        var par = new[] { a, b };
        string mensaje = $"Cruce real {dia.Key:yyyy-MM-dd}: " +
            $"asignaturas {a.Gen.IdAsignatura}/{b.Gen.IdAsignatura}.";
        bool docente = AgCalendarioConflictos.DocenteReal(a.Gen, contexto)
            && AgCalendarioConflictos.DocenteReal(b.Gen, contexto)
            && a.Gen.IdTercero == b.Gen.IdTercero;
        bool mismaCohorte = AgCalendarioConflictos.Cohorte(a.Gen)
            == AgCalendarioConflictos.Cohorte(b.Gen);
        if (docente)
        {
            Registrar(AgReglaViolacion.R1_SuperposicionDocente, par, mensaje);
            if (!mismaCohorte)
                Registrar(AgReglaViolacion.R7_TerceroOtroGrupoFechaHora, par, mensaje);
        }
        if (a.Gen.IdAula > 0 && a.Gen.IdAula == b.Gen.IdAula)
            Registrar(AgReglaViolacion.R2_SuperposicionAula, par, mensaje);
        if (mismaCohorte)
            Registrar(AgReglaViolacion.R3_SuperposicionGrupo, par, mensaje);
    }
}

// R4: unicidad por fecha concreta/franja y tope diario, no por semestre entero.
foreach (var dia in ocupaciones.Where(o => o.Gen.EsVacante == 0)
    .GroupBy(o => new { Cohorte = AgCalendarioConflictos.Cohorte(o.Gen), o.Fecha }))
{
    foreach (var franja in dia.GroupBy(o => o.Franja.Trim().ToUpperInvariant()))
        if (franja.Count() > 1)
            Registrar(AgReglaViolacion.R4_DosAsignaturasMismoDia, franja,
                $"Más de una sesión en la franja {franja.Key} el {dia.Key.Fecha:yyyy-MM-dd}.");
    int tope = dia.Key.Fecha.DayOfWeek == DayOfWeek.Saturday ? 2 : 1;
    if (dia.Select(o => o.Franja.Trim().ToUpperInvariant()).Distinct().Count() > tope)
        Registrar(AgReglaViolacion.R4_DosAsignaturasMismoDia, dia,
            $"Tope diario excedido el {dia.Key.Fecha:yyyy-MM-dd}.");
}

// Sustituye R7 externa que solo revisaba FechaInicioEstimada.
foreach (var o in ocupaciones.Where(o => AgCalendarioConflictos.DocenteReal(o.Gen, contexto)))
    if (contexto.TerceroOcupado(o.Gen.IdTercero, o.Fecha,
        new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(o.Inicio, o.Fin)))
        Registrar(AgReglaViolacion.R7_TerceroOtroGrupoFechaHora, new[] { o },
            $"Ocupación externa del docente {o.Gen.IdTercero} el {o.Fecha:yyyy-MM-dd}.");
```

Invariantes de integración:

- Normalizar cohorte con el mismo fallback de semestre del método `Cohorte` actual; no perder grupos cuyo semestre sea vacío/no canónico al trasladarlo al helper.
- Docentes vacantes no consumen recurso docente, pero sus sesiones siguen consumiendo aula y cohorte.
- No convertir todos los HARD del caso en cero: los cruces docentes reales deben seguir detectándose.
- La nueva unidad de estas reglas es par/fecha o grupo/fecha; ajustar pruebas y etiquetas a esa definición, no exigir conservar 197 como cifra histórica.
- Construir una proyección por evaluación y agrupar por fecha/recurso. La comparación por día mostrada es correcta pero debe indexarse por recurso para grandes poblaciones; medir rendimiento tras autorización. Cachear solo entradas temporales completas, nunca por cantidad de fechas.
- Al recalcular una corrida guardada, recuperar `FechasGen` desde sus sesiones por `(IdIndividuo, IdGen)`. `ag_gen` no conserva esa lista: no reconstruir con el generador legacy las fechas de un plan intercalado ya materializado. Usar el fallback únicamente para genes nuevos que realmente carezcan de plan fechado.

## A3. Disponibilidad para el grupo completo

En `AgAsignadorTercero`, añadir el helper siguiente y reutilizarlo en `DocenteDisponible`, `BuscarAlterno` y reparaciones. Requiere A2. No altera fechas ni restricciones para hacer caber un docente.

```csharp
private static bool MismoGrupo(AgGen a, AgGen b) =>
    a.IdAula == b.IdAula && a.IdAsignatura == b.IdAsignatura
    && a.IdPrograma == b.IdPrograma && a.IdSede == b.IdSede
    && string.Equals(a.Grupo ?? "", b.Grupo ?? "", StringComparison.Ordinal);

private static bool DisponibleGrupo(int docente, AgGen gen,
    AgIndividuo individuo, AgContexto contexto)
{
    if (docente <= 0 || docente == contexto.IdTerceroVacante) return false;
    var grupo = individuo.Genes.Where(g => MismoGrupo(g, gen)).ToList();
    if (grupo.Count == 0) grupo.Add(gen);
    if (grupo.Any(g => g.Estado == AgEstadoGen.Fijado && g.IdTercero != docente))
        return false;
    if (grupo.Any(g => !contexto.EsDocenteHabilitado(
        docente, g.IdPrograma, g.IdAsignatura, g.Semestre))) return false;

    var propias = AgCalendarioConflictos.Crear(grupo, contexto);
    var ajenas = AgCalendarioConflictos.Crear(individuo.Genes.Where(g =>
        !MismoGrupo(g, gen) && AgCalendarioConflictos.DocenteReal(g, contexto)
        && g.IdTercero == docente), contexto);
    if (propias.IsFailure || ajenas.IsFailure) return false;
    foreach (var o in propias.Value)
    {
        if (ajenas.Value.Any(a => AgCalendarioConflictos.SeCruzan(o, a))) return false;
        if (contexto.TerceroOcupado(docente, o.Fecha,
            new ConversorPDF.XLL.Academic.Domain.Models.FranjaHoraria(o.Inicio, o.Fin)))
            return false;
    }
    return true;
}

public bool DocenteDisponible(AgGen gen, AgIndividuo individuo, AgContexto contexto)
{
    if (gen == null || individuo == null || contexto == null) return false;
    return DisponibleGrupo(gen.IdTercero, gen, individuo, contexto);
}

public int BuscarAlterno(AgGen gen, AgContexto contexto, AgIndividuo individuo)
{
    if (gen == null || contexto == null || individuo == null
        || !contexto.PoolPorAsignatura.TryGetValue(gen.IdAsignatura, out var pool))
        return 0;
    return pool.Where(p => p.Activo == 1 && p.IdTercero != gen.IdTercero)
        .OrderBy(p => p.Prioridad).ThenBy(p => p.VecesAsignada).ThenBy(p => p.IdTercero)
        .Select(p => p.IdTercero)
        .FirstOrDefault(id => DisponibleGrupo(id, gen, individuo, contexto));
}

// Tras encontrar candidato, aplicar a todo el grupo, no solo al gen consultado:
// este bloque se inserta en el llamador que dispone de candidato/gen/individuo.
if (candidato > 0 && DisponibleGrupo(candidato, gen, individuo, contexto))
    foreach (var miembro in individuo.Genes.Where(g => MismoGrupo(g, gen)))
    {
        miembro.IdTercero = candidato;
        miembro.EsVacante = 0;
        miembro.MotivoVacante = string.Empty;
    }
```

El bloque final pertenece a la misma clase o a un método de aplicación que exponga esa operación; no llamar helpers privados desde otra clase. No cambiar genes Fijado salvo mantener su mismo docente. Para siembra, formar primero el grupo y después buscar candidato; `AsignarDocente` aislado no basta. Si no se encuentra alterno, no convertir automáticamente una asignación fija en VACANTE; conservar bloqueo y explicar la causa.

No concluir que las 440 sesiones vacantes son todas evitables: primero corregir falsos conflictos y después medir déficit real del pool.

## A4. Métricas, estado y promoción

Dos cambios necesarios para que `vacantes` deje de quedar en cero:

```csharp
// UI, antes de UpdateCorridaAsync:
corrida.Vacantes = _ultimoResultado.MejorIndividuo?.Vacantes ?? 0; // GENES
```

En `AgScheduleRepository.UpdateCorridaAsync`, conservar SQL/parámetros restantes e incorporar:

```sql
-- Entre los campos del SET existente:
vacantes = @Vacantes,
```

```csharp
// En el objeto param existente:
corrida.Vacantes,
```

Después de materializar, leer staging y tratar fallos de lectura como errores, no como cero sesiones. Ejemplo del cierre UI, sustituyendo el éxito basado exclusivamente en `stagingOk`:

```csharp
var sesiones = await _scheduleRepository.GetSesionesStagingAsync(
    idCorrida, _ctsEjecucion.Token);
if (sesiones.IsFailure)
    throw new InvalidOperationException("No se pudo verificar staging: " + sesiones.Error);

var mejor = _ultimoResultado.MejorIndividuo;
if (mejor == null) throw new InvalidOperationException("Corrida sin individuo final.");
int sesionesVacantes = sesiones.Value.Count(s => s.EsVacante == 1);
double horasVacantes = sesiones.Value.Where(s => s.EsVacante == 1).Sum(s => s.Horas);
var calendario = AgCalendarioConflictos.Crear(mejor.Genes, _contextoActual);
if (calendario.IsFailure) throw new InvalidOperationException(calendario.Error);

// Comparar multiconjunto de gen/fecha/horas: detecta pérdida, duplicación o cambio de sesión.
string Clave(long gen, DateTime fecha, double horas) =>
    gen + "|" + fecha.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)
    + "|" + horas.ToString("R", CultureInfo.InvariantCulture);
var previstas = calendario.Value.Select(o => Clave(o.Gen.Id, o.Fecha,
    _contextoActual.BloquesPorId[o.Gen.IdBloqueHorario].DuracionHoras))
    .OrderBy(k => k, StringComparer.Ordinal);
bool fechasValidas = sesiones.Value.All(s => s.FechaSesion.HasValue);
var efectivas = sesiones.Value.Where(s => s.FechaSesion.HasValue)
    .Select(s => Clave(s.IdGen, s.FechaSesion.Value, s.Horas))
    .OrderBy(k => k, StringComparer.Ordinal);
bool cobertura = fechasValidas && previstas.SequenceEqual(efectivas);
bool factible = stagingOk && cobertura && mejor.HardViolations == 0
    && mejor.Vacantes == 0 && sesionesVacantes == 0;

corrida.Vacantes = mejor.Vacantes;
corrida.Estado = stagingOk ? AgEstadoCorrida.Completada : AgEstadoCorrida.Error;
var guardado = await _scheduleRepository.UpdateCorridaAsync(corrida, _ctsEjecucion.Token);
if (guardado.IsFailure) throw new InvalidOperationException(guardado.Error);
btnPromoverProduccion.Enabled = factible;
lblStatusAG.Text = factible ? "Estado: Horario factible."
    : $"Estado: Resultado parcial. HARD={mejor.HardViolations}; " +
      $"sesiones vacantes={sesionesVacantes}; horas vacantes={horasVacantes:0.#}.";
lblStatusAG.ForeColor = factible ? Color.DarkGreen : Color.DarkOrange;
traza.Resultado(factible, lblStatusAG.Text);
traza.Fin(stagingOk, "Ejecución técnica finalizada; " + lblStatusAG.Text);
```

Usar duración lectiva del bloque como autoridad y validar su compatibilidad con el intervalo; catálogos de 8 h con pausas pueden tener intervalo mayor que horas lectivas. No computar pausas como horas docentes. Para conflictos, conservar el intervalo de ocupación autorizado del bloque.

El gate de cobertura también exige que la demanda de todas las relaciones requeridas esté cubierta: la igualdad contra genes existentes no detecta una materia omitida. Conservar validación de demanda y rechazar cualquier grupo en cuarentena. Verificar además intervalo/bloque exactos, como en Criterio11. La proyección usada para aprobar debe ser la del mismo snapshot previo a materializar, sin cambios silenciosos de fechas por coherencia.

Repetir el gate en carga de resultados y en el servicio de promoción. No basta deshabilitar un botón ni añadir un enum `Parcial` inexistente sin su migración/contrato. `Completada` puede seguir siendo estado técnico; `Aprobada` requiere proceso de aprobación separado.

## A5. Conteos y puntuaciones separados

En `CargarModuloResultadosAsync`, reemplazar el bloque que suma `SoftViolations` a la cantidad de filas:

```csharp
var lectura = await _scheduleRepository.GetViolacionesByIndividuoAsync(
    ind.Id, CancellationToken.None);
if (lectura.IsFailure)
{
    lblHardViolaciones.Text = "No se pudo cargar el detalle de violaciones.";
    LoggerHelper.Warning(lectura.Error);
}
else
{
    var detalle = lectura.Value;
    int registrosHard = detalle.Count(v => v.EsDura);
    int registrosSoft = detalle.Count(v => !v.EsDura);
    lblHardViolaciones.Text =
        $"Registros HARD={registrosHard}, SOFT={registrosSoft}; " +
        $"infracciones HARD={ind.HardViolations}; penalización SOFT={ind.SoftViolations:0.##}.";
    lblHardViolaciones.ForeColor = ind.HardViolations == 0 ? Color.DarkGreen : Color.DarkRed;
    if (ind.Violaciones != null && detalle.Count != ind.Violaciones.Count)
        LoggerHelper.Warning("Cantidad de registros persistidos distinta de la evaluada.");
}
```

La comparación memoria/DB sirve inmediatamente después de guardar o cuando la lista en memoria se cargó completamente. Al rehidratar, no declarar pérdida porque una lista no se haya cargado todavía. La verificación existente posterior a `SaveIndividuoAsync` compara ambas cardinalidades correctamente: conservarla.

## A6. Violaciones globales y enlace a sesiones

### Consulta del visor: conservar violaciones sin gen

Reemplazar la consulta que exige `INNER JOIN ag_sesiones` por una consulta principal de violaciones del individuo óptimo. Incorporar usings/DTO según la vista existente; esta consulta devuelve registros, no filas multiplicadas por sesión.

```sql
SELECT v.id AS Id, v.id_individuo AS IdIndividuo, v.id_gen AS IdGen,
       CAST(v.regla AS TEXT) AS Regla, v.es_dura AS EsDura,
       v.mensaje AS Mensaje, v.detalle_json AS DetalleJson,
       c.id AS IdCorrida
FROM ag_corrida c
JOIN ag_individuo i ON i.id = c.id_individuo_optimo AND i.id_corrida = c.id
JOIN ag_violacion v ON v.id_individuo = i.id
WHERE c.id IN @IdCorridas AND c.subunidad = @Subunidad
ORDER BY v.es_dura DESC, v.regla, v.id;
```

Usar la subunidad de la corrida cargada y rechazar mezcla incompatible de contextos. Mostrar en una lista general todos los registros. Indexar por sesión solo los que tengan genes y fecha verificables en detalle. Si hay registros generales pero cero enlaces a sesiones, mostrar «violaciones generales sin sesión asociada», no «sin detalle». No renombrar `ag_sesiones` a una tabla inventada de staging.

### Resolver IDs en la transacción de guardado

En `AgScheduleRepository.SaveIndividuoAsync`, después de asignar `g.Id` a todos los genes y antes de insertar violaciones, resolver el contrato `indicesGenes` de A2. Este fragmento va dentro de la transacción existente; no ejecutar un segundo guardado del individuo para obtener IDs.

```csharp
var serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
foreach (var violacion in individuo.Violaciones)
{
    var detalle = serializer.DeserializeObject(violacion.DetalleJson ?? "{}")
        as Dictionary<string, object>;
    if (detalle == null || !detalle.ContainsKey("indicesGenes")) continue;
    var indices = detalle["indicesGenes"] as object[];
    if (indices == null || indices.Length == 0)
        throw new InvalidOperationException("Detalle de genes inválido.");
    var ids = indices.Select(x => Convert.ToInt32(x)).Select(indice =>
    {
        if (indice < 0 || indice >= individuo.Genes.Count || individuo.Genes[indice].Id <= 0)
            throw new InvalidOperationException("Índice de gen no resoluble al guardar.");
        return individuo.Genes[indice].Id;
    }).Distinct().ToArray();
    detalle["idGenes"] = ids;
    violacion.IdGen = ids[0];
    violacion.DetalleJson = serializer.Serialize(detalle);
}
```

Asegurar que el orden de `individuo.Genes` no cambia entre evaluación y guardado. Si materialización/coherencia cambia el calendario final, revalidar y reemplazar violaciones mediante `ReemplazarViolacionesAsync` para el mismo individuo; no mostrar un fitness de un calendario anterior como resultado del calendario nuevo. A largo plazo, usar una clave temporal estable por gen si se permiten reordenamientos.

Los registros históricos sin `indicesGenes` se conservan como globales; no inventar su asociación. El nuevo detalle permite enlazar ambos genes por `idGenes` y filtrar por `fecha`, evitando marcar todas las sesiones de un gen por un cruce de un solo día.

## Validación pendiente de la implementación propuesta

Preparar pruebas para: aprobadas con déficit y sin mutación; MAR-VIE 61→75 sin ampliar fechas; Estadística/Teorías sin choque; bloques distintos con intervalos superpuestos; conflicto en una fecha posterior al inicio; dos franjas de sábado; duración 8 h; grupo fijado; docente externo ocupado; genes sin fechas precalculadas; expansión inválida; materialización parcial; métricas de vacantes; 209 registros globales visibles sin `id_gen`; detalle por par después de asignar IDs.

Tras aprobación: ejecutar pruebas pertinentes, Full Debug/Release y corrida funcional en copia con límite original. Verificar igualdad de sesiones antes/después del guardado y cero aprobaciones falsas. No dar por resueltas vacantes ni afirmar que todos los HARD eran falsos. No ejecutar estos pasos como parte de esta actualización documental.
