# Corrida 92: completar huecos y unificar evaluación/generación

Estado: propuesta de implementación. Los bloques siguientes NO se aplicaron al código de producción ni se compilaron o ejecutaron.
T1–T7 son tareas locales, no números globales de fases. Implementar secuencialmente con el árbol de trabajo verificado.

## Objetivo

Dentro del plazo aprobado, conservar las restricciones no reclasificadas y completar materias automáticas con sesiones fechadas libres, incluso combinando días de diferentes bandas. Evaluador, validador y fábrica deben elegir el mismo plan completo para entradas iguales.

No ampliar fechas, reclasificar otras materias ni prometer disponibilidad docente a partir de capacidad del aula. No introducir precedencias curriculares que el modelo actual no declara: OrdenCurricular determina prioridad, no una obligación global de terminar una materia antes de iniciar otra.

## Diagnóstico y alcance

- AgPlanificadorOleadas.ConstruirInterno fuerza referencia si permitirParcial=false y existe cualquier patrón restringido. Por ello un aula con cinco LEGADO y una AUTOMATICO no consume el candidato compacto.
- La selección actual también exige una mejora de siete días para generación, mientras evaluación utiliza otra comparación. Eliminar ambos criterios divergentes.
- AgPlanificadorRanuras usa un cursor por banda: deja atrás huecos y obliga a cada materia a una banda de la partición. Liberar el patrón no garantiza combinar huecos entre bandas.
- La búsqueda contiene limiteMayo=30/05/2027. Retirarlo; ningún algoritmo productivo debe depender de ese período concreto.
- #92 se ejecutó después de ampliar manualmente el plazo. El libro tiene 116 horas posteriores al 30/05, aunque la reclasificación sí se persistió.

## T1 — Fijar regresiones antes del cambio

Archivo de pruebas: AgPlanificadorRanurasTests o una nueva suite pura equivalente.

Fixtures CALI, 05/02/2027–30/05/2027, 4 h TARDE, LUN-VIE:

| Caso | Horas por materia, orden de IDs | Patrones de primeras cinco | Última |
|---|---|---|---|
| III | 32,48,48,48,32,48 | 1,2 / 3,4 / 5 / 1,2 / 3,4 | AUTOMATICO, sin restricción |
| IV | 32,48,48,48,48,48 | mismos | AUTOMATICO, sin restricción |

Añadir CARTAGO X-A: 32,32,64,112 h; restricciones de las primeras tres materias 4,5 / 6 / 4,5; última AUTOMATICO; ventana JUE-DOM, 4 h, jueves/viernes noche, sábado mañana/tarde y domingo mañana. Exigir que las 28 sesiones de Opción de Grado aprovechen huecos compatibles y no se reserven exclusivamente a sábado.

Festivos: incluir 22–28/03 y 10/05; reproducir también los restantes configurados del período. Usar bloques reales/canónicos de 13:00–17:00.

- III: 64 sesiones; última materia 12 sesiones. Conservando las cinco anteriores del ejemplo, puede terminar 13/05.
- IV: 68 sesiones; última materia 12 sesiones. Conservando las cinco anteriores del ejemplo, puede terminar 19/05.
- Las fechas 13/05 y 19/05 son regresiones del fixture descrito; no límites ni garantías universales.
- Comparar Construir y ConstruirEvaluacion: mismas sesiones completas para iguales entradas.
- Probar año distinto, ausencia de festivos, unidad de 8 h y sábado M/T.

Ejemplo de prueba para CALI, en suite del proyecto existente (añadir los using de producción, LINQ y xUnit correspondientes):

```csharp
[Theory]
[InlineData(false, 13)]
[InlineData(true, 19)]
public void Mixta_ConservaRestriccionesYCompletaAntesDelLimite(bool cuarto, int diaFin)
{
    var inicio = new DateTime(2027, 2, 5);
    var limite = new DateTime(2027, 5, 30);
    int[] horas = cuarto ? new[] { 32, 48, 48, 48, 48, 48 }
        : new[] { 32, 48, 48, 48, 32, 48 };
    var demandas = horas.Select((h, i) =>
        (IdAsignatura: i + 1, Horas: (double)h, OrdenCurricular: 1)).ToList();
    var patrones = new Dictionary<int, string>
    {
        { 1, "1,2" }, { 2, "3,4" }, { 3, "5" }, { 4, "1,2" }, { 5, "3,4" }
    };
    Func<DateTime, bool> festivo = d =>
        d.Date >= new DateTime(2027, 3, 22) && d.Date <= new DateTime(2027, 3, 28)
        || d.Date == new DateTime(2027, 5, 10);
    var bloques = Enumerable.Range(1, 5).Select(d => new AgBloqueHorario
    {
        Id = d, DiaSemanaNumero = d, Franja = "TARDE", HoraInicio = "13:00",
        HoraFin = "17:00", DuracionHoras = 4, Activo = 1
    }).ToList();
    var sut = new AgPlanificadorOleadas();
    var generado = sut.Construir(demandas, "LUN-VIE", inicio, limite, 4,
        festivo, bloques, "TARDE", patrones);
    var evaluado = sut.ConstruirEvaluacion(demandas, "LUN-VIE", inicio, limite, 4,
        festivo, bloques, "TARDE", patrones);
    Assert.True(generado.IsSuccess);
    Assert.True(evaluado.IsSuccess);
    Assert.Empty(generado.Value.Infactibles);
    var sesiones = generado.Value.Planes.SelectMany(p => p.Sesiones).ToList();
    Assert.Equal(horas.Sum() / 4, sesiones.Count);
    Assert.All(sesiones, s => Assert.True(s.Fecha <= limite));
    var ultima = generado.Value.Planes.Single(p => p.IdAsignatura == 6);
    Assert.Equal(12, ultima.Sesiones.Count);
    Assert.Equal(new DateTime(2027, 5, diaFin), ultima.Sesiones.Max(s => s.Fecha));
    Assert.Equal(sesiones.Count, sesiones.GroupBy(s =>
        new { s.Fecha, s.Ranura.Dia, s.Ranura.Franja }).Count());
    var fechasGeneradas = sesiones.Select(s => $"{s.IdAsignatura}|{s.Fecha:yyyyMMdd}|{s.Ranura}")
        .OrderBy(s => s).ToArray();
    var fechasEvaluadas = evaluado.Value.Planes.SelectMany(p => p.Sesiones)
        .Select(s => $"{s.IdAsignatura}|{s.Fecha:yyyyMMdd}|{s.Ranura}").OrderBy(s => s).ToArray();
    Assert.Equal(fechasGeneradas, fechasEvaluadas);
}
```

Esta prueba comprueba la ruta pública. Añadir otra que invoque CompletarReferencia con una referencia explícita y compare una a una las sesiones restringidas antes/después: no sustituir esa comprobación por mera igualdad de cantidades.

## T2 — Añadir validador común de planes fechados

Archivo: AgPlanificadorRanuras.cs. Añadir `using CSharpFunctionalExtensions;`.
Insertar los métodos siguientes dentro de la clase existente. Reutilizan Calendario, AgPlanOleada y AgPlanOleadasResultado actuales.

El validador comprueba el contrato de ranuras canónicas por aula. Las comprobaciones globales de docentes y los intervalos reales del catálogo siguen siendo obligatorias antes de promover.

```csharp
internal static Result ValidarFechado(
    AgPlanOleadasResultado plan,
    IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> demandas,
    IReadOnlyList<AgRanura> ranuras, DateTime ini, DateTime fin, double unidad,
    Func<DateTime, bool> festivo, IReadOnlyDictionary<int, string> patrones,
    CancellationToken ct)
{
    if (plan == null) { return Result.Failure("I5_PLAN_NULO"); }
    var requeridas = demandas.ToDictionary(a => a.IdAsignatura,
        a => (int)Math.Ceiling(a.Horas / unidad));
    var calendario = Calendario(ranuras, ini.Date, fin.Date, festivo, ct);
    var ocupadas = new HashSet<int>();
    var cantidades = new Dictionary<int, int>();
    foreach (var p in plan.Planes)
    {
        ct.ThrowIfCancellationRequested();
        int requerida;
        if (!requeridas.TryGetValue(p.IdAsignatura, out requerida)
            || cantidades.ContainsKey(p.IdAsignatura)
            || p.SesionesRequeridas != requerida || p.Sesiones.Count > requerida)
        {
            return Result.Failure("I5_PLAN_DEMANDA_INCONSISTENTE");
        }
        cantidades[p.IdAsignatura] = p.Sesiones.Count;
        string texto;
        var dias = patrones != null && patrones.TryGetValue(p.IdAsignatura, out texto)
            ? AgDiasVentana.Dias(texto) : new List<int>();
        if (patrones != null && patrones.TryGetValue(p.IdAsignatura, out texto)
            && !string.IsNullOrWhiteSpace(texto) && dias.Count == 0)
        {
            return Result.Failure("I1_PATRON_INVALIDO");
        }
        foreach (var s in p.Sesiones)
        {
            ct.ThrowIfCancellationRequested();
            if (s.IdAsignatura != p.IdAsignatura || s.Ranura == null
                || s.Fecha != s.Fecha.Date
                || Math.Abs(s.Ranura.Horas - unidad) > 0.001
                || (dias.Count > 0 && !dias.Contains(s.Ranura.Dia))
                || s.Fecha < p.FechaInicio.Date || s.Fecha > p.FechaFin.Date
                || !p.RanurasPatron.Any(r => r.Equals(s.Ranura)))
            {
                return Result.Failure("I5_PLAN_SESION_INVALIDA");
            }
            int indice = calendario.FindIndex(c => c.Fecha == s.Fecha
                && c.Ranura.Equals(s.Ranura));
            if (indice < 0 || !ocupadas.Add(indice))
            {
                return Result.Failure("I5_PLAN_FECHA_O_COLISION");
            }
        }
    }
    var pendientes = new HashSet<int>(requeridas.Where(kv =>
        !cantidades.ContainsKey(kv.Key) || cantidades[kv.Key] < kv.Value)
        .Select(kv => kv.Key));
    return pendientes.SetEquals(plan.Infactibles)
        ? Result.Success() : Result.Failure("I5_PLAN_PENDIENTES_INCONSISTENTES");
}
```

Precondición de estos métodos internos: entradas ya validadas por ConstruirInterno (IDs únicos, demanda positiva finita, unidad válida, ventana y ranuras válidas). Si se exponen a otros llamadores, incorporar esas mismas validaciones de entrada.

## T3 — Completar referencia mediante ocupación por sesión

En AgPlanificadorRanuras.cs, añadir este método. Conserva exactamente las sesiones de materias restringidas completas de la referencia. Después intenta ubicar restringidas pendientes respetando sus días y finalmente las libres. No cambia OrigenPatron ni DiasPatron en configuración.

El resultado es un candidato determinista, no una prueba de optimalidad o imposibilidad. Si queda parcial, conservar la alternativa de búsqueda acotada y reportar pendientes.

```csharp
internal static Result<AgPlanOleadasResultado> CompletarReferencia(
    IReadOnlyList<(int IdAsignatura, double Horas, int OrdenCurricular)> demandas,
    IReadOnlyList<AgRanura> ranuras, DateTime ini, DateTime fin, double unidad,
    Func<DateTime, bool> festivo, IReadOnlyDictionary<int, string> patrones,
    AgPlanOleadasResultado referencia, CancellationToken ct)
{
    var orden = demandas.OrderBy(a => Math.Max(0, a.OrdenCurricular))
        .ThenBy(a => a.IdAsignatura).ToList();
    var requeridas = orden.ToDictionary(a => a.IdAsignatura,
        a => (int)Math.Ceiling(a.Horas / unidad));
    var restricciones = orden.ToDictionary(a => a.IdAsignatura, a =>
    {
        string texto;
        return new HashSet<int>(patrones != null
            && patrones.TryGetValue(a.IdAsignatura, out texto)
            ? AgDiasVentana.Dias(texto) : new List<int>());
    });
    var calendario = Calendario(ranuras, ini.Date, fin.Date, festivo, ct);
    var ocupadas = new bool[calendario.Count];
    var porId = new Dictionary<int, List<AgSesionPlanificada>>();

    if (referencia != null)
    {
        var validacion = ValidarFechado(referencia, demandas, ranuras,
            ini, fin, unidad, festivo, patrones, ct);
        if (validacion.IsFailure)
        {
            return Result.Failure<AgPlanOleadasResultado>(validacion.Error);
        }
        foreach (var p in referencia.Planes.Where(p =>
            restricciones[p.IdAsignatura].Count > 0
            && !referencia.Infactibles.Contains(p.IdAsignatura)))
        {
            ct.ThrowIfCancellationRequested();
            var sesiones = p.Sesiones.OrderBy(s => s.Fecha)
                .ThenBy(s => s.Ranura.Franja, StringComparer.Ordinal).ToList();
            foreach (var s in sesiones)
            {
                int indice = calendario.FindIndex(c => c.Fecha == s.Fecha
                    && c.Ranura.Equals(s.Ranura));
                ocupadas[indice] = true; // Ya validado: existe y no se solapa.
            }
            porId[p.IdAsignatura] = sesiones;
        }
    }

    var restantes = orden.Where(a => !porId.ContainsKey(a.IdAsignatura))
        .OrderBy(a => restricciones[a.IdAsignatura].Count == 0 ? 1 : 0)
        .ThenBy(a => Math.Max(0, a.OrdenCurricular))
        .ThenBy(a => a.IdAsignatura);
    foreach (var a in restantes)
    {
        ct.ThrowIfCancellationRequested();
        var sesiones = new List<AgSesionPlanificada>();
        var dias = restricciones[a.IdAsignatura];
        for (int i = 0; i < calendario.Count && sesiones.Count < requeridas[a.IdAsignatura]; i++)
        {
            ct.ThrowIfCancellationRequested();
            var hueco = calendario[i];
            if (ocupadas[i] || (dias.Count > 0 && !dias.Contains(hueco.Ranura.Dia)))
            {
                continue;
            }
            ocupadas[i] = true;
            sesiones.Add(new AgSesionPlanificada(a.IdAsignatura, hueco.Fecha, hueco.Ranura));
        }
        porId[a.IdAsignatura] = sesiones;
    }

    var planes = new List<AgPlanOleada>();
    foreach (var a in orden)
    {
        var sesiones = porId[a.IdAsignatura];
        if (sesiones.Count == 0) { continue; }
        var usadas = ranuras.Where(r => sesiones.Any(s => s.Ranura.Equals(r))).ToList();
        DateTime primera = sesiones.First().Fecha;
        DateTime ultima = sesiones.Last().Fecha;
        int capacidad = calendario.Count(c => c.Fecha >= primera && c.Fecha <= ultima
            && usadas.Any(r => r.Equals(c.Ranura)));
        // Una banda global; Carril es metadato, no una reserva exclusiva de fechas.
        planes.Add(new AgPlanOleada(a.IdAsignatura, 0, planes.Count,
            usadas.Select(r => r.Dia).Distinct().ToList(), primera, ultima,
            (ultima - primera).Days / 7 + 1, requeridas[a.IdAsignatura], capacidad,
            usadas, sesiones));
    }
    var pendientes = orden.Where(a => porId[a.IdAsignatura].Count < requeridas[a.IdAsignatura])
        .Select(a => a.IdAsignatura).ToList();
    var resultado = new AgPlanOleadasResultado(planes, pendientes,
        pendientes.Select(id => "Asig " + id + ": sesiones no ubicadas; el completado no acredita imposibilidad.").ToList(),
        new List<IReadOnlyList<int>> { ranuras.Select(r => r.Dia).Distinct().ToList() },
        ocupacionPorcentaje: calendario.Count == 0 ? 0.0
            : 100.0 * planes.Sum(p => p.Sesiones.Count) / calendario.Count);
    var comprobacion = ValidarFechado(resultado, demandas, ranuras,
        ini, fin, unidad, festivo, patrones, ct);
    return comprobacion.IsSuccess ? Result.Success(resultado)
        : Result.Failure<AgPlanOleadasResultado>(comprobacion.Error);
}
```

Cambios complementarios en la búsqueda antigua:

- Eliminar `DateTime limiteMayo = new DateTime(2027, 5, 30);`.
- Sustituir su condición de parada por condiciones derivadas de entradas, por ejemplo plan completo y `mejorFin <= fin.Date`, conservando presupuesto y cancelación.
- No aumentar el presupuesto para compensar el problema de representación por bandas.
- El completado nuevo no usa cursor por banda: cada posición del calendario se ocupa individualmente.

## T4 — Una selección común de planes completos

Archivo: AgPlanificadorOleadas.cs, ConstruirInterno.

Conservar validaciones previas y construcción de referencia. Conservar rechazo de errores estructurales no I5. Sustituir desde el comentario «También se buscan mejoras…» hasta el último `return referencia` del método por el bloque siguiente.

No mantener el retorno automático `generacion_con_patrones`, el umbral de siete días ni comparadores diferentes de planes completos.

```csharp
var alternativas = new List<AgPlanOleadasResultado>();
AgPlanOleadasResultado referenciaValida = null;
if (referencia.IsSuccess && AgPlanificadorRanuras.ValidarFechado(referencia.Value,
    asignaturas, ranuras, fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, ct).IsSuccess)
{
    referenciaValida = referencia.Value;
    alternativas.Add(referenciaValida);
}
var completado = AgPlanificadorRanuras.CompletarReferencia(asignaturas, ranuras,
    fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, referenciaValida, ct);
if (completado.IsSuccess) { alternativas.Add(completado.Value); }

// Primero un pase barato sobre huecos. Búsqueda acotada solo cuando no lo resuelve.
if (completado.IsFailure || completado.Value.Infactibles.Count > 0)
{
    var buscado = AgPlanificadorRanuras.Construir(asignaturas, ranuras,
        fechaInicio.Date, fechaFin.Date, unidadSesion, esFestivo, patrones,
        ct: ct, diagnostico: diagnostico);
    if (AgPlanificadorRanuras.ValidarFechado(buscado, asignaturas, ranuras,
        fechaInicio, fechaFin, unidadSesion, esFestivo, patrones, ct).IsSuccess)
    {
        alternativas.Add(buscado);
    }
}

var requeridas = asignaturas.ToDictionary(a => a.IdAsignatura,
    a => (int)Math.Ceiling(a.Horas / unidadSesion));
var completasPrevias = referenciaValida == null ? new List<int>()
    : referenciaValida.Planes.Where(p => !referenciaValida.Infactibles.Contains(p.IdAsignatura))
        .Select(p => p.IdAsignatura).ToList();
var elegibles = alternativas.Where(p => completasPrevias.All(id =>
    p.Planes.Any(x => x.IdAsignatura == id && x.Sesiones.Count == requeridas[id]))).ToList();
var completos = elegibles.Where(p => p.Infactibles.Count == 0).ToList();
if (completos.Count > 0) { elegibles = completos; }
else if (!permitirParcial)
{
    // Conservar contrato de generación: factibles previas y cuarentena de infactibles.
    return referenciaValida != null ? Result.Success(referenciaValida)
        : Result.Failure<AgPlanOleadasResultado>("I5_SIN_PLAN_COMPLETO: quedan sesiones sin ubicar.");
}
var elegido = elegibles.OrderByDescending(p => p.Planes.Sum(x => x.Sesiones.Count))
    .ThenBy(p => p.Infactibles.Count)
    .ThenBy(p => p.Planes.SelectMany(x => x.Sesiones).Select(s => s.Fecha)
        .DefaultIfEmpty(DateTime.MaxValue).Max()).FirstOrDefault();
if (elegido == null)
{
    return Result.Failure<AgPlanOleadasResultado>("I5_SIN_PLAN_FECHADO_VALIDO");
}
LoggerHelper.Debug($"[AG-RANURAS-TMP] ELEGIR {diagnostico} pendientes={elegido.Infactibles.Count} " +
    $"sesiones={elegido.Planes.Sum(p => p.Sesiones.Count)} parcial={permitirParcial}");
return Result.Success(elegido);
```

La diferencia permitida entre evaluación y generación se limita a qué hacer si NO existe plan completo: evaluación puede informar ubicaciones parciales; generación mantiene su contrato de aislamiento de infactibles. Si existe plan completo, ambos deben seleccionar las mismas fechas con desempate determinista.

No presentar resultados parciales como imposibilidad demostrada. Conservar la información de agotamiento de búsqueda y los motivos del intento acotado en el diagnóstico agregado, incluso si se elige otra alternativa parcial.

## T5 — Consumir el plan sin volver a repartir fechas

Archivos: AgFabricaIndividuos, AgCoherenciaGrupo, AgExpansorSesiones, AgMaterializadorSesiones.

- Mantener FechasGen exactas por ranura y HorasGen=cantidad*duración.
- En una materia de 12 sesiones combinadas entre varios días, no redistribuir posteriormente 12/N sesiones por cada gen: los conteos reales del plan mandan.
- No cambiar el bloque del gen por otro horario distinto del plan fechado. En la rama que ya dispone de ranuras del plan, el criterio de coincidencia debe ser exacto:

```csharp
Func<AgBloqueHorario, bool> coincideRanura = b =>
    b.Activo == 1 && b.DiaSemanaNumero == ranura.Dia
    && Math.Abs(b.DuracionHoras - ranura.Horas) < 0.001
    && string.Equals(b.Franja, ranura.Franja, StringComparison.OrdinalIgnoreCase);
```

- Si no existe bloque real para una sesión planificada, devolver diagnóstico de configuración. No afirmar que se materializó el plan ni sustituirlo silenciosamente por horario legacy.
- Si hay genes temporalmente fijados por el usuario fuera de la referencia generada, incorporarlos como reservas explícitas antes de este completado. No moverlos para lograr el fixture. La propuesta no agrega esa representación por inferencia.
- Mantener los validadores globales de docentes y colisiones; #92 tiene siete infracciones duras independientes de la reducción de ventana.

## T6 — Persistir contexto reproducible y registrar decisiones

- Guardar en parametros_json el rango efectivo, perfil/alcance, huella de configuración, versión y parámetros del planificador. #92 conserva `{}` y no acredita por sí solo el plazo utilizado.
- Registrar referencia/completado/búsqueda, sesiones por materia, pendientes, último día y motivo de elección. Resumen INFO y detalle DEBUG, sin nombres de docentes.
- Activar DEBUG explícitamente para la captura funcional; no modificar preferencias globales por defecto.
- Mantener plazo aprobado separado de proyección; las pruebas de 2027 no pueden convertirse en constantes del algoritmo.

## T7 — Aceptación y entrega

Pruebas a escribir antes de aplicar el cambio:

1. CALI III y IV de T1: cero sesiones posteriores al 30/05; 256/272 h exactas; fechas de las cinco materias restringidas conservadas por el completado.
2. Generación y evaluación devuelven el mismo plan completo, incluyendo mezcla LEGADO/AUTOMATICO.
3. La última materia utiliza varios días/bandas cuando sea necesario; no tiene una banda exclusiva artificial.
4. Dos materias libres comparten calendario sin colisiones; reuso de huecos anteriores a otras materias.
5. LEGADO/FIJO no autorizados mantienen días permitidos; no reclasificaciones adicionales.
6. Referencia completa + completado con final anterior por menos de siete días: elegir el anterior.
7. Referencia parcial: no perder materias previamente completas.
8. Fechas festivas, fuera de ventana, repetidas, ranura incompatible, metadatos/pendientes inconsistentes: rechazar candidato.
9. Semana incompleta, sábado M/T, unidad 8 h, cancelación y búsqueda agotada sin falso certificado de imposibilidad.
10. Período distinto de 2027: resultados dependen de entradas, no de limiteMayo.
11. Fábrica → FechasGen → coherencia → staging conserva fechas y horas. Usar mock solo para persistencia, no para sustituir planificador/expansor.
12. Escalera 4→5 y 4→5→6, tope e idempotencia: sin regresiones.

Después de autorización: compilación, pruebas focalizadas, regresión completa y smoke test con 30/05/2027 sin ampliación manual. Si hay restricciones globales adicionales, informar las pendientes; no garantizar factibilidad docente a partir del fixture del aula.

No considerar el plan cerrado hasta tener evidencia del mismo binario/configuración y resultado funcional. No hacer commit/push ni modificar datos del usuario como parte de este documento.
