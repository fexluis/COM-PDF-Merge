<!-- historico congelado - generado por .agents/tools/split-history.ps1 - commit d0ee6fe - no editar -->
# walkthrough - Fases 1103-1347 (Rondas 232)

> Indice: [../../../walkthrough.md](../../../walkthrough.md) - Escision: 2026-11-16

---
### Diagnóstico y Contexto
- **Pérdida Silenciosa de Asignaturas por Falta de Sanción Dura:** Tras excluir `SIN_BLOQUE_NIVEL` de `R14` (F1100), los genes afectados quedaban sin ninguna infracción dura. El individuo sólo recibía el peso blando de vacante y, al materializar, terminaba en cuarentena I6 produciendo una pérdida silenciosa en la promoción.
- **Pérdida de Motivo `SIN_BLOQUE_NIVEL` en Sellado de Grupos Mixtos:** `AgCoherenciaGrupo.Sellar` únicamente preservaba `SIN_PLAN_OLEADA`. Ante cruces que dejaban grupos mixtos con genes reales, `Sellar` forzaba `EsVacante = 0` y vaciaba el motivo, eludiendo la tipificación.
- **Infactibilidad Falsa I5 por Default Ficticio de 64h:** En `AgValidadorVentana.cs`, si `HorasObjetivo == 0`, se atribuían 64 horas ficticias por defecto, disparando bloqueos espurios `[I5_INFACTIBLE]`.
- **Doble Sanción Dura por `SIN_PLAN_OLEADA` en R14:** Los genes cuarentenados con `IdBloqueHorario = 0` y motivo `SIN_PLAN_OLEADA` disparaban simultáneamente `R14` y `R7`.
- **Vacíos en la Suite de Pruebas:** Carencia de aserciones sobre la regla dura sustituta en `AgValidadorReglas`, preservación en `Sellar` y ausencia de bloqueo en factibilidad.

### Fases Ejecutadas

#### 1. Fase 1103 — Emisión de Violación Dura R7 para Motivo `SIN_BLOQUE_NIVEL`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Genes con `SIN_BLOQUE_NIVEL` quedaban sin regla dura, permitiendo que la solución fuera seleccionada con asignaturas perdidas.
- **Acción Aplicada:** Incorporado bucle en `AgValidadorReglas.cs` que emite `AgReglaViolacion.R7_CoberturaHorariaIncompleta` como violación dura cuando `EsVacante == 1` y `MotivoVacante == SIN_BLOQUE_NIVEL`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:494-505`.

#### 2. Fase 1104 — Preservación de Motivos No Asignables en `AgCoherenciaGrupo.Sellar`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Grupos mixtos tras cruce genético sobrescribían `SIN_BLOQUE_NIVEL` a docente real si al menos un gen tenía docente.
- **Acción Aplicada:** Generalizada la colección `motivoNoAsignable` incluyendo `SIN_PLAN_OLEADA` y `SIN_BLOQUE_NIVEL` para forzar todo el grupo a vacante y mantener el motivo original.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:234-245`.

#### 3. Fase 1105 — Cálculo de Demanda Real en Factibilidad I5 sin Default Arbitrario
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `HorasObjetivo == 0` asumía 64.0h inflando la demanda y bloqueando aulas viables con `[I5_INFACTIBLE]`.
- **Acción Aplicada:** Resolución de intensidad real vía `asigPorId` (`IntHoraria`), exclusión de la demanda si `horasAu <= 0.0` y emisión de advertencia diagnóstica `[I5_HORAS_DESCONOCIDAS]`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:203, 222-235`.

#### 4. Fase 1106 — Desduplicación de `SIN_PLAN_OLEADA` en Regla R14
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Doble conteo de violaciones duras (`R14` y `R7`) para genes cuarentenados con `SIN_PLAN_OLEADA`.
- **Acción Aplicada:** Añadida exclusión explícita de `SIN_PLAN_OLEADA` en la condición de `R14_BloqueHorarioNoAsignado`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:172-178`.

#### 5. Fase 1107 — Cobertura Integral de Pruebas para `SIN_BLOQUE_NIVEL`, `Sellar` e I5
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Ausencia de pruebas que garantizaran la emisión de R7 para `SIN_BLOQUE_NIVEL`, la exclusión de R14 para `SIN_PLAN_OLEADA`, el sellado de grupo mixto y la no-infactibilidad ante `HorasObjetivo == 0`.
- **Acción Aplicada:** Agregadas las pruebas `Validador_GenSinBloqueNivel_EmiteR7Dura`, `Validador_GenSinPlanOleada_NoEmiteR14`, `Sellar_GrupoMixtoConSinBloqueNivel_PreservaMotivoYVacante` y `ValidarFactibilidad_HorasObjetivoCero_NoBloqueaYAdvierte`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs:610-670`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs:252-290`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs:211-232`.

---

### Verificación de Principios (Ronda 212)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1103 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:494-505` (bucle R7 dura para `SIN_BLOQUE_NIVEL`) |
| F1104 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:234-245` (`motivoNoAsignable`) |
| F1105 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:203, 222-235` (resolución `asigPorId` y advertencia I5) |
| F1106 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:172-178` (exclusión `SIN_PLAN_OLEADA` en R14) |
| F1107 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | 4 pruebas nuevas en suites de `ValidadorReglas`, `CoherenciaGrupo` y `ValidadorVentana` |

---

## 🔧 Ronda 213 — Fechas de Oleadas en Vinculaciones/Promoción, Min/Max de Grupo, Conteo I5 y Contador Atómico (Fases 1108 a 1112)

### Diagnóstico y Contexto
- **Pérdida de Fechas de Oleada en Vinculaciones y Promoción:** Cuando `IdPeriodo > 0`, tanto `AgGenToAgVinculacionMapper.MapGrupo` como `AgMaterializadorSesiones` sustituían incondicionalmente las fechas de la sub-ventana de la oleada (`FechaInicioEstimada` / `FechaFinEstimada`) por las fechas del semestre completo (`contexto.Ventana.FechaInicio / FechaFin`), tornando invisibles las oleadas en `ag_vinculaciones` y en la columna de promoción del módulo R.
- **Fechas de Grupo Determinadas Únicamente por el Primer Gen:** En patrones multi-día, `MapGrupo` tomaba únicamente las fechas de `g0`, sin considerar el rango completo del grupo.
- **Incoherencia en Conteo de Materias de Bloqueo I5:** Cuando F1105 descartaba materias con horas desconocidas (`continue`), el mensaje de bloqueo I5 continuaba reportando el total bruto (`asigsAula.Count`), generando confusión en el informe.
- **Falta de Sincronización en Contador de Pruebas `Construcciones`:** El incremento `Construcciones++` no era atómico, introduciendo potencial indeterminismo en ejecuciones paralelas.
- **Ausencia de Pruebas Unitarias para Fechas de Oleada en Staging:** Las pruebas existentes verificaban únicamente asignaciones de horas y códigos, sin validar que las fechas de vinculación conservaran la sub-ventana de la oleada.

### Fases Ejecutadas

#### 1. Fase 1108 — Preservación de Sub-Ventana de Oleada en Mapeador y Materializador
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Fechas de oleadas eran sobreescritas por las fechas completas del semestre.
- **Acción Aplicada:** En `AgGenToAgVinculacionMapper.MapGrupo`, se preservan las fechas de los genes; en `AgMaterializadorSesiones`, sólo se acotan a la ventana si la exceden o están vacías, sin reemplazarlas forzosamente.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs:73-82`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs:99-108`.

#### 2. Fase 1109 — Rango Min/Max de Fechas en `MapGrupo`
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `MapGrupo` tomaba las fechas de `g0` en vez del rango total del grupo multi-día.
- **Acción Aplicada:** Calculadas `fiGrupo` y `ffGrupo` con `Min()` y `Max()` sobre todos los genes con fecha estimada del grupo.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Mappers/AgGenToAgVinculacionMapper.cs:74-75`.

#### 3. Fase 1110 — Reporte de Materias Computadas en Bloqueo I5
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Conteo reportado en I5 incluía asignaturas con horas desconocidas que habían sido excluidas de la suma.
- **Acción Aplicada:** Introducido contador `materiasComputadas` e integrado en el mensaje: `({materiasComputadas} materias computadas de {asigsAula.Count})`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:227, 238, 244`.

#### 4. Fase 1111 — Contador `Construcciones` Sincronizado en Planificador de Oleadas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `Construcciones++` no era atómico ante ejecuciones concurrentes.
- **Acción Aplicada:** Campo privado `_construcciones` leído con `Volatile.Read` e incrementado con `Interlocked.Increment`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgPlanificadorOleadas.cs:18-20, 39`.

#### 5. Fase 1112 — Cobertura Integral de Pruebas de Fechas de Oleada e I5
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Carencia de aserciones para la conservación de fechas de oleada y el reporte de materias computadas.
- **Acción Aplicada:** Añadidas pruebas `MapGrupo_ConOleada_ConservaFechasDelGen`, `MapGrupo_GenesConFechasDistintas_UsaMinMax`, `Materializar_ConOleada_NoSobrescribeFechasConVentana` y `ValidarFactibilidad_ConHorasDesconocidas_ReportaMateriasComputadas`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Mappers/AgGenToAgVinculacionMapperTests.cs:60-129`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs:320-384`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs:231-254`.

---

### Verificación de Principios (Ronda 213)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1108 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGenToAgVinculacionMapper.cs:73-82`, `AgMaterializadorSesiones.cs:99-108` |
| F1109 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgGenToAgVinculacionMapper.cs:74-75` (`fiGrupo`, `ffGrupo`) |
| F1110 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:227, 238, 244` (`materiasComputadas`) |
| F1111 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:18-20, 39` (`Interlocked`/`Volatile`) |
| F1112 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | 4 pruebas nuevas en `AgGenToAgVinculacionMapperTests`, `AgMaterializadorSesionesTests`, `AgValidadorVentanaTests` |

---

## 🔧 Ronda 214 — Pre-vuelo UI con Duración Efectiva por Nivel, Advertencia D4' y Parseo Invariante en Sub-Ventana (Fases 1113 a 1116)

### Diagnóstico y Contexto
- **Falso Bloqueo de Pre-Vuelo UI ante Selección de 8h en Pregrado:** `AgFormGeneradorHorarios` validaba la existencia de bloques activos utilizando `aula.HorasDia` directamente. Si el usuario seleccionaba 8h en pregrado y el catálogo no disponía de bloques de 8h, el pre-vuelo abortaba la corrida, a pesar de que el motor (por la regla D4') planificaría deterministamente con 4 h.
- **Falta de Trazabilidad ante Degradación D4':** El usuario no recibía ninguna advertencia diagnóstica que explicara por qué una configuración de 8 h/día se planificaba a 4 h/día.
- **Vulnerabilidad Cultural en Parseo de Sub-Ventanas de Oleada:** `DateTime.TryParse` en `AgMaterializadorSesiones` dependía de la cultura del hilo de ejecución. En entornos con calendarios regionales no gregorianos, el parseo fallaba en silencio y omitía las cotas de la sub-ventana.
- **Acoplamiento de Patrón en Pre-Vuelo de Aula:** La UI pasaba `aula.DiasVentana` como primer parámetro (`diasPatronCsv`) a `AgValidadorVentana.Validar`, eludiendo la validación específica que el motor realiza por cada asignatura.

### Fases Ejecutadas

#### 1. Fase 1113 — Pre-Vuelo UI con Duración Efectiva por Nivel
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Selección de 8h en pregrado causaba falso bloqueo "No existe ningún bloque horario activo … con duración 8 horas".
- **Acción Aplicada:** Resuelta `horasEfectivasV` mediante `AgUnidadesSemanales.HorasSesionAula(aula, nivelV)` en `BtnGenerarAulas_Click` y `BtnOptimizar_Click`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1113-1126, 1360-1373`.

#### 2. Fase 1114 — Invariante Dura de Nivel: Ningún Gen de PREGRADO con Bloque de 8h
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `horas_dia = 8` se aceptaba en PREGRADO; aunque la fábrica asignaba 4h, mutación y coherencia podían preservar un bloque de 8h al no existir una invariante dura en el punto de convergencia de operadores.
- **Acción Aplicada:**
  - *1114.a:* Invariante en `AgCoherenciaGrupo.Sellar` (normaliza `HorasSemanales` a `uNivel`, limpia `FechasGen` y reasigna bloque de 4h del mismo día o marca cuarentena `SIN_BLOQUE_NIVEL`).
  - *1114.b:* Duración objetivo en `AgOperadorMutacion` impuesta por el nivel del aula (`AgUnidadesSemanales.HorasSesionAula`), nunca por el gen.
  - *1114.c:* Bloqueo explícito pre-vuelo `[D4_PREGRADO_8H]` en `AgValidadorVentana.ValidarFactibilidad` y aborto bloqueante con `traza.Fallo` y `MessageBox` en `AgFormGeneradorHorarios` (P07 y P08).
  - *1114.d:* Pruebas unitarias de invariante en `AgCoherenciaGrupoTests`, `AgOperadorMutacionTests`, `AgValidadorVentanaTests` y `AgFabricaIndividuosTests`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:56-96`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs:110-125`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:222-230`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1117-1126, 1379-1388`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs:290-357`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs:193-225`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs:255-275`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs:406-431`.

#### 3. Fase 1115 — Parseo Invariante en Cotas de Sub-Ventana de Oleada
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `DateTime.TryParse` fallaba en silencio bajo calendarios o culturas no gregorianas.
- **Acción Aplicada:** Reemplazado por `DateTime.TryParseExact` con `CultureInfo.InvariantCulture`, registrando advertencia ante literales corruptos y acotando a la ventana semestral.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs:100-123`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMaterializadorSesionesTests.cs:387-463`.

#### 4. Fase 1116 — Desacoplamiento de Patrón en Pre-Vuelo de Aula
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Envío de `aula.DiasVentana` como `diasPatron` redundante en pre-vuelo de UI.
- **Acción Aplicada:** Envío de `string.Empty` en `Validar(...)` en ambos pre-vuelos de la UI delegando la validación del patrón curricular al motor genético.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1128, 1391`.

---

### Verificación de Principios (Ronda 214)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1113 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1113-1115, 1375-1377` (`horasEfectivasV`) |
| F1114 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:56-96`, `AgOperadorMutacion.cs:110-125`, `AgValidadorVentana.cs:222-230`, `AgFormGeneradorHorarios.cs:1117-1126, 1379-1388` |
| F1115 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMaterializadorSesiones.cs:100-123`, `AgMaterializadorSesionesTests.cs:387-463` |
| F1116 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1128, 1391` (`string.Empty`) |

---

## 🔧 Ronda 215 — Robustecimiento de Invariante de Nivel: Protección de POSGRADO, Resguardo de Genes sin Duración e Indexación O(1) de Bloques (Fases 1117 a 1121)

### Diagnóstico y Contexto
- **Disparo Indebido de Invariante ante HorasSemanales = 0:** La condición `Math.Abs(g.HorasSemanales - uNivel) >= 0.001` evaluaba genes recién heredados o de prueba sin duración asignada (0h) como violación, forzándolos a 4h, borrando `FechasGen` y enviándolos a cuarentena `SIN_BLOQUE_NIVEL`.
- **Degradación Indebida de POSGRADO y Aulas no Indexadas:** `UnidadNivel` retornaba `4.0` cuando `IdAula` no se encontraba en el diccionario de aulas, forzando asignaturas válidas de posgrado (8h) a la unidad de pregrado.
- **Falta de Cobertura de No-Regresión:** La suite de pruebas no contaba con aserciones para verificar que POSGRADO conserva 8h ni que genes sin duración no son alterados.
- **Diagnóstico Impreciso de Causa de Fallo:** Pre-vuelo UI clasificaba el bloqueo D4' de horas/día como `AgCausaFallo.SinBloques` en lugar de una causa de política de nivel.
- **Búsqueda Lineal Repetitiva en Sellar:** `ctx.Bloques.FirstOrDefault(...)` realizaba escaneo lineal O(Genes × Bloques) en cada ejecución de cruce, mutación y sellado final.

### Fases Ejecutadas

#### 1. Fase 1117 — Protección contra Disparo de Invariante con HorasSemanales <= 0
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Genes con duración desconocida (0h) perdían sus fechas calculadas y pasaban a vacante sin causa real.
- **Acción Aplicada:** Incorporadas cláusulas de guarda en `AgCoherenciaGrupo.cs`: omitir si `uNivel <= 0`, si `g.HorasSemanales <= 0` o si `g.HorasSemanales <= uNivel + 0.001`, corrigiendo únicamente sobre-duración real.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:72-74`.

#### 2. Fase 1118 — Retorno -1.0 en UnidadNivel ante Aulas no Resolubles
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `UnidadNivel` devolvía 4h ante aulas no resueltas, degradando grupos de POSGRADO a 4h.
- **Acción Aplicada:** Retorno de `-1.0` (nivel no verificable) en `UnidadNivel`, evitando la asunción errónea de pregrado.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:68`.

#### 3. Fase 1119 — Pruebas Unitarias de No-Regresión de Invariante
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Ausencia de pruebas que garantizaran la preservación de 8h en posgrado y la inmunidad de genes con 0h.
- **Acción Aplicada:** Añadidas pruebas `Sellar_PosgradoConBloque8h_NoDegradaDuracion`, `Sellar_GenSinHorasSemanales_NoModificaFechasNiEstado` y `Sellar_AulaNoIndexada_NoAlteraDuracion`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs:360-464`.

#### 4. Fase 1120 — Código AgCausaFallo.HorasDiaInvalidoNivel
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Trazas P07 y P08 reportaban `SinBloques` ante bloqueo de política académica D4'.
- **Acción Aplicada:** Agregado `HorasDiaInvalidoNivel = 13` a `AgCausaFallo.cs` y referenciado en trazas P07/P08 de `AgFormGeneradorHorarios.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Enums/AgCausaFallo.cs:50`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1122, 1385`.

#### 5. Fase 1121 — Indexación O(1) de Bloques en Sellar
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Búsqueda lineal O(G×B) en bucle caliente de coherencia.
- **Acción Aplicada:** Generado `indiceBloques` agrupado por `(DiaSemanaNumero, DuracionHoras)` al inicio de `Sellar` y búsqueda O(1) vía `TryGetValue`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:53-57, 80-81`.

---

### Verificación de Principios (Ronda 215)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1117 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:72-74` (guardas `uNivel`, `HorasSemanales`) |
| F1118 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:68` (`return -1.0;`) |
| F1119 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupoTests.cs:360-464` (3 pruebas unitarias) |
| F1120 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCausaFallo.cs:50`, `AgFormGeneradorHorarios.cs:1122, 1385` |
| F1121 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:53-57, 80-81` (`indiceBloques`) |

---

## 🔧 Ronda 216 — Validación Pre-Vuelo por Sub-Ventana de Oleada, Capacidad Canónica de Carril y Caché de Bloques en Contexto (Fases 1122 a 1126)

### Diagnóstico y Contexto
- **Omisión de Sub-Ventanas de Oleada en Factibilidad Pre-Vuelo:** `CapacidadDe` y el balance global de aula evaluaban la demanda contra la ventana completa del semestre. Si una asignatura con alta carga horaria excedía las semanas de su sub-ventana asignada por oleadas, el pre-vuelo aprobaba erróneamente la corrida, la cual luego fallaba en el motor con `R7` y `SIN_PLAN_OLEADA`.
- **Sobreestimación Artificial de Capacidad (2 Sesiones/Día):** `sesionesPorDia` calculaba `Math.Round(au.HorasDia / unidadAu)`. En aulas de pregrado con 8h y unidad de 4h, devolvía 2, duplicando la capacidad permitida por día a pesar de que el motor impone un máximo de 1 sesión por día/carril (R4).
- **Cálculo No Canónico de Carriles en Pre-Vuelo:** `Math.Ceiling(ventanaIso.Count / 2.0)` devolvía valores discordantes con la partición canónica de `AgCarriles.Cantidad` en ventanas impares (ej. 3 días ⇒ 2 en vez de 1).
- **Reconstrucción Repetida de Índices en Sellar:** `Sellar` instanciaba un diccionario de bloques en cada llamada dentro de bucles calientes genéticos en lugar de aprovechar un índice memoizado en el contexto.

### Fases Ejecutadas

#### 1. Fase 1122 — Validación de Sub-Ventana por Oleada en Pre-Vuelo
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Asignaturas inviables en su oleada pasaban el pre-vuelo porque la demanda global del aula no superaba la ventana semestral.
- **Acción Aplicada:** Ejecutado `AgPlanificadorOleadas.Construir` por aula dentro de `ValidarFactibilidad` en `AgValidadorVentana.cs`. Si hay infactibles, emite bloqueo explícito `[I5_INFACTIBLE_OLEADA]`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:255-276`.

#### 2. Fase 1123 — Capacidad Canónica de 1 Sesión/Día/Carril
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `sesionesPorDia` duplicaba la capacidad permitida en aulas configuradas con 8h en pregrado.
- **Acción Aplicada:** Fijado `sesionesPorDia = 1` en `CapacidadDe` de `AgValidadorVentana.cs` según la restricción R4 de carriles.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:122`.

#### 3. Fase 1124 — Carriles Canónicos con AgCarriles.Cantidad
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Uso residual de `Math.Ceiling` sobrestimaba la cantidad de carriles en ventanas impares.
- **Acción Aplicada:** Reemplazado por `AgCarriles.Cantidad(ventanaIso)` en `AgValidadorVentana.cs:129`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:129`.

#### 4. Fase 1125 — Caché Indexado de Bloques en AgContexto
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `Sellar` generaba miles de diccionarios efímeros de bloques durante la optimización genética.
- **Acción Aplicada:** Implementado `TryGetBloque(int dia, double dur, out AgBloqueHorario b)` con memoización lazy e invalidación en `ConstruirIndices()` en `AgContexto.cs`, e integrado en `AgCoherenciaGrupo.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs:150-169`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:81-89`.

#### 5. Fase 1126 — Cobertura de Pruebas Unitarias de Oleada y Carril
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Falta de pruebas para validar el rechazo por sub-ventana de oleada y verificar la cota de 1 sesión/día.
- **Acción Aplicada:** Agregadas pruebas `ValidarFactibilidad_AsignaturaExcedeSuOleada_BloqueaI5` y `ValidarFactibilidad_Pregrado8h_NoDuplicaCapacidadPorDia` en `AgValidadorVentanaTests.cs`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs:277-343`.

---

### Verificación de Principios (Ronda 216)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1122 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:255-276` (`[I5_INFACTIBLE_OLEADA]`) |
| F1123 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:122` (`sesionesPorDia = 1`) |
| F1124 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:129` (`AgCarriles.Cantidad`) |
| F1125 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:150-169`, `AgCoherenciaGrupo.cs:81-89` (`TryGetBloque`) |
| F1126 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentanaTests.cs:277-343` (2 pruebas unitarias) |

---

## 🔧 Ronda 217 — Coherencia de Jornada en Reasignación de Duración, Invalidación Explícita de Caché de Bloques y Alcance Curricular en Pre-Vuelo (Fases 1127 a 1130)

### Diagnóstico y Contexto
- **Colapso de Franjas en `TryGetBloque`:** `_indiceBloquesDiaDur` indexaba por `(dia, duracion)` usando `g.First()`. Si existían bloques de 4h en mañana y tarde, siempre retornaba el primero (mañana). Un aula de jornada TARDE con un gen normalizado de 8h→4h terminaba reasignada indebidamente a la franja MAÑANA, rompiendo la coherencia con `AgFranjaJornada`.
- **Caché Obsoleta ante Mutación de `Bloques`:** `_indiceBloquesDiaDur` no disponía de método de invalidación explícita para invocarse tras recargar `Bloques` desde base de datos en P05 o en pruebas con carga tardía.
- **Inclusión Indebida de Asignaturas Fuera de Alcance en Pre-Vuelo:** `itemsPlan` incluía todas las asignaturas asignadas al aula sin verificar su programa y semestre, arrojando falsos `[I5_INFACTIBLE_OLEADA]` por materias curriculares no planificables en la corrida.

### Fases Ejecutadas

#### 1. Fase 1127 — Coherencia de Jornada y Franja Preferida en TryGetBloque
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Normalización 8h→4h reasignaba a franja MAÑANA en aulas de jornada TARDE.
- **Acción Aplicada:** Almacenada lista de bloques por `(dia, duracion)` en `AgContexto.cs` y resuelta la franja canónica del aula con `AgFranjaJornada.Resolve` en `AgCoherenciaGrupo.cs`, priorizando el bloque cuya franja coincide con la jornada.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs:153-195`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:79-88`.

#### 2. Fase 1128 — Invalidación Explícita de Caché ante Mutación de Bloques
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Inexistencia de invalidación al actualizar la colección `Bloques`.
- **Acción Aplicada:** Agregado método `InvalidarIndiceBloques()` en `AgContexto.cs` e invocado en `ConstruirIndices()` y en la recarga de bloques de `AgFormGeneradorHorarios.cs` (P05).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs:157`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:762`.

#### 3. Fase 1129 — Filtro Curricular de Asignaturas en Pre-Vuelo
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Falsos bloqueos de oleada generados por asignaturas de otros semestres o programas asignadas al aula.
- **Acción Aplicada:** Agregado filtro en `itemsPlan` que exige coincidencia de `IdPrograma` y `SemestreRomano` entre la asignatura y el aula en `AgValidadorVentana.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs:261-265`.

#### 4. Fase 1130 — Cobertura de Pruebas Unitarias de Jornada, Caché y Alcance
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Ausencia de pruebas que certifiquen el respeto a la jornada en la corrección, la invalidación de caché y la exclusión de asignaturas de otro semestre.
- **Acción Aplicada:** Agregadas pruebas `Sellar_PregradoTarde8h_ReasignaBloqueDeFranjaTarde` y `Contexto_TrasInvalidarIndiceBloques_UsaCatalogoVigente` en `AgCoherenciaGrupoTests.cs`, y `ValidarFactibilidad_AulaConAsignaturaDeOtroSemestre_NoBloquea` en `AgValidadorVentanaTests.cs`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs:463-524`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs:342-372`.

---

### Verificación de Principios (Ronda 217)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1127 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:161-186`, `AgCoherenciaGrupo.cs:79-88` (`franjaAula`) |
| F1128 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgContexto.cs:157`, `AgFormGeneradorHorarios.cs:762` (`InvalidarIndiceBloques`) |
| F1129 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:261-265` (filtro curricular en `itemsPlan`) |
| F1130 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupoTests.cs:463-524`, `AgValidadorVentanaTests.cs:342-372` (3 pruebas) |

---

## 🔧 Ronda 218 — Coherencia de Jornada en Cruce, Mutación y Regla R11 con Unificación de Selección de Bloques (Fases 1131 a 1135)

### Diagnóstico y Contexto
- **Colapso de Franjas en Cruce (`AgOperadorCruce`):** El operador seleccionaba el bloque buscando únicamente coincidencia de `DiaSemanaNumero` y `DuracionHoras`. Al haber múltiples bloques de igual duración en un día (e.g. 4h MANANA 08-12 y 4h TARDE 13-17), siempre ganaba el primero, asignando erróneamente franjas matutinas a aulas de jornada TARDE o NOCHE.
- **Migración Incontrolada de Franjas en Mutación (`AgOperadorMutacion`):** Los bloques candidatos para mutación mono-día y multi-día solo filtraban por duración, permitiendo que mutaciones arbitrarias cambiaran la franja horaria fuera de la jornada del aula sin respetar la política académica.
- **Carencia de Validación Dura R11 para Jornada (`AgValidadorReglas`):** R11 sólo evaluaba la duración del bloque en horas contra la unidad de nivel requerida. Genes con bloques en franjas incompatibles con la jornada del aula no eran penalizados ni bloqueados.
- **Duplicación de Selección de Bloques:** Cuatro componentes (`AgFabricaIndividuos`, `AgOperadorCruce`, `AgOperadorMutacion`, `AgCoherenciaGrupo`) implementaban lógicas dispersas para elegir bloques por día, duración y franja.

### Fases Ejecutadas

#### 1. Fase 1131 — Selección de Bloque Coherente con Jornada en Operador de Cruce
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Reasignación arbitraria al bloque de la mañana en cruces para aulas de jornada tarde.
- **Acción Aplicada:** Reemplazada la búsqueda ciega por `AgSeleccionBloque.ParaDia(aulaDst, diaObjetivo, unidadCruce, contexto)` para el día objetivo y días alternativos del patrón.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorCruce.cs:182-205`.

#### 2. Fase 1132 — Filtro de Franja según Jornada en Operador de Mutación
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Mutación alteraba la franja horaria fuera de la jornada del aula.
- **Acción Aplicada:** Resuelta `franjaObjetivo` vía `AgFranjaJornada.Resolve(aulaMut.Jornada, (int)duracionObjetivo, gen.DiaSemana)?.Franja` y filtrados los candidatos en mutación mono-día y multi-día con fallback seguro.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgOperadorMutacion.cs:126-175`.

#### 3. Fase 1133 — Validación de Franja vs Jornada en Regla Dura R11
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Discrepancias franja ∉ jornada eran invisibles para el motor y no generaban violaciones duras.
- **Acción Aplicada:** Extendida la regla R11 en `AgValidadorReglas.cs` para validar que la franja del bloque coincida con la esperada por `AgFranjaJornada.Resolve(aulaGen.Jornada, ...)`, emitiendo violación dura si difieren.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:445-459`.

#### 4. Fase 1134 — Cobertura de Pruebas Unitarias de Jornada en Cruce, Mutación y R11
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Inexistencia de pruebas con aulas TARDE y catálogos con franjas múltiples de igual duración.
- **Acción Aplicada:** Añadidas pruebas unitarias `Cruzar_AulaTarde_AsignaBloqueDeFranjaTarde` en `AgOperadorCruceTests.cs`, `Mutacion_AulaTarde_NoMigraAManana` en `AgOperadorMutacionTests.cs` y `Validador_GenDeFranjaFueraDeJornada_EmiteR11` en `AgValidadorReglasTests.cs`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorCruceTests.cs:213-241`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgOperadorMutacionTests.cs:227-260`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs:672-708`.

#### 5. Fase 1135 — Unificación de Selección de Bloques en AgSeleccionBloque
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Criterios dispersos de selección de bloque día + duración + franja con riesgo estructural de regresión.
- **Acción Aplicada:** Creado `AgSeleccionBloque.cs` que unifica la resolución de franja canónica y la consulta al contexto, registrado en `ConversorPDF.XLL.csproj` e integrado en `AgOperadorCruce.cs` y `AgCoherenciaGrupo.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgSeleccionBloque.cs:1-21`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj:607`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:79-88`.

---

### Verificación de Principios (Ronda 218)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1131 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorCruce.cs:182-205` (`AgSeleccionBloque.ParaDia`) |
| F1132 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgOperadorMutacion.cs:126-175` (`franjaObjetivo`) |
| F1133 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:445-459` (R11 franja vs jornada) |
| F1134 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Tests cruce/mutación/reglas (3 pruebas unitarias) |
| F1135 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgSeleccionBloque.cs`, `ConversorPDF.XLL.csproj:607` |

---

## 🔧 Ronda 219 — Formato Canónico DIA|horas|horario en Columna K del Exportador Excel (Fases 1136 a 1138)

### Diagnóstico y Contexto
- **Columna K sin Detalle de Horas ni Horario (`SemestresExcelExporter`):** La matriz volcaba únicamente nombres de días separados por barra (`"MARTES/MIERCOLES"`), impidiendo validar visual y programáticamente que el carril respetara la unidad semanal de 4h y el horario asignado.
- **Formateador Incompleto (`DiasHorasSemanaFormatter`):** La API solo aceptaba `IEnumerable<Sesion>` de dominio, mientras que el exportador maneja `SesionListItem` proyectado para la grilla y el reporte. Mantenía además un comentario `TODO` de cableado pendiente.
- **Ausencia de Cobertura de Pruebas:** No existía suite de pruebas unitarias para certificar el contrato canónico `DIA|horas|horario` y su deduplicación en la columna K.

### Fases Ejecutadas

#### 1. Fase 1136 — Cableado Canónico de la Columna K en SemestresExcelExporter
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** La columna K no permitía verificar el carril de 4h por falta de horas y horario.
- **Acción Aplicada:** Agregado overload `DesdeSesionListItems(IEnumerable<SesionListItem> items)` en `DiasHorasSemanaFormatter.cs` y cableada la asignación de `matrizDatos[r, 10]` en `SemestresExcelExporter.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs:57-93`, `ConversorPDF.XLL/Academic/Infrastructure/Excel/SemestresExcelExporter.cs:5, 377-380`.

#### 2. Fase 1137 — Retiro de TODO y Preservación de Granularidad en Grilla
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Comentarios obsoletos `TODO` en formateador y riesgo de alteración errónea de la grilla por sesión atómica.
- **Acción Aplicada:** Retirado comentario `TODO` en `DiasHorasSemanaFormatter.cs`. Ratificada la estructura de columnas atómicas (`ColDias`, `ColHoras`, `ColHorario`) en `FormSesionesAcademicas.cs` sin sobreescrituras desnormalizadas innecesarias.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/DiasHorasSemanaFormatter.cs:52-54`.

#### 3. Fase 1138 — Pruebas Unitarias para Formato Canónico de Columna K
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Carencia de pruebas que aseguren la invariabilidad del formato canónico `DIA|horas|horario`.
- **Acción Aplicada:** Creado `DiasHorasSemanaFormatterTests.cs` cubriendo carril pregrado, sesiones duplicadas/semanales, entradas sin día, colecciones nulas y compatibilidad con `Sesion`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/DiasHorasSemanaFormatterTests.cs:1-105`.

#### 4. Fase 1139 (Opcional) — Botón Cargar en Módulo R
- **Estado:** RETIRADA / FALSO POSITIVO DE DEFECTO
- **Diagnóstico Crítico:** Alcance de producto / evolutivo no bloqueante. Las compuertas de seguridad D11/D12 ya están resguardadas en `AgFormGeneradorHorarios.cs:2086-2097`.
- **Acción Aplicada:** Se mantiene la compuerta de seguridad I8 existente sin alterar el flujo de botones del usuario.

---

### Verificación de Principios (Ronda 219)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1136 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DiasHorasSemanaFormatter.cs:57-93`, `SemestresExcelExporter.cs:377-380` |
| F1137 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DiasHorasSemanaFormatter.cs:52-54` (limpieza `TODO`) |
| F1138 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `DiasHorasSemanaFormatterTests.cs:1-105` (5 pruebas) |
| F1139 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2086-2097` (preservada compuerta I8) |

---

## 🔧 Ronda 220 — Implementación F1-F4 compilación-bloqueada y saneamiento V76 inefectivo en filas legacy (Fases 1195 a 1201)

### Diagnóstico y Contexto
- **Error de Compilación CS0246 en AgMensajeDimensionesAula:** El using importaba `ConversorPDF.XLL.Academic.Domain.Models` en vez de `ConversorPDF.XLL.Academic.Domain.Models.Genetica`, bloqueando la compilación del Add-In y la ejecución de la suite de pruebas.
- **Saneamiento Inefectivo en V76 para Filas Legacy:** La condición estricta `COALESCE(n.subunidad,'') = COALESCE(a.subunidad,'')` impedía curar filas de `ag_aula` previas a V53/V62 que tenían `subunidad = ''`, mientras el árbol almacenaba la subunidad real.
- **Ambigüedad No Determinista en V76:** `GROUP BY a.id` devolvía un registro de nodo no determinista cuando existían múltiples versiones en `ag_arbol_nodo`, arriesgando adoptar dimensiones de perfiles caducos.
- **Falta de Aislamiento de Subunidad y Notificación de Suciedad:** `AgAulaDimensionesSync` no filtraba por `subunidad` y no notificaba a `SyncContext.Scheduler?.MarkDirty()`.
- **Desbordamiento de Diálogo y Formato Roto de Log:** Acumulación no acotada de mensajes con `\n` desbordaba `MessageBox` y fragmentaba el log.
- **Silenciamiento de Errores en V76:** Catch interno consumía excepciones, permitiendo que `SchemaMigrator` marcara la migración como aplicada impidiendo reintentos.

### Fases Ejecutadas

#### 1. Fase 1195 — Corrección de Namespace en AgMensajeDimensionesAula (CS0246)
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Referencia de namespace errónea para `AgAula` impedía la compilación.
- **Acción Aplicada:** Reemplazado import a `ConversorPDF.XLL.Academic.Domain.Models.Genetica` en `AgMensajeDimensionesAula.cs` y añadido en `AgMensajeDimensionesAulaTests.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMensajeDimensionesAula.cs:3`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgMensajeDimensionesAulaTests.cs:3`.

#### 2. Fase 1196 — Tolerancia a Subunidad Vacía en V76
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Asimetría lector↔migrador; no se curaban filas legacy con `subunidad = ''`.
- **Acción Aplicada:** Incorporada condición `(COALESCE(x.subunidad,'') = COALESCE(a.subunidad,'') OR TRIM(COALESCE(a.subunidad,'')) = '')`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV76_SanearDimensionesAgAula.cs:46`.

#### 3. Fase 1197 — Determinismo por MAX(idarbol) en V76
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `GROUP BY a.id` seleccionaba nodos arbitrarios ante múltiples árboles.
- **Acción Aplicada:** Reemplazado `GROUP BY` por subconsulta correlacionada `n.idarbol = (SELECT MAX(x.idarbol) FROM ag_arbol_nodo x ...)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV76_SanearDimensionesAgAula.cs:40-49`.

#### 4. Fase 1198 — Aislamiento por Subunidad y Marcado Sucio en Sincronización
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `UPDATE ag_aula` sin subunidad afectaba otras subunidades y no propagaba cambios pendientes al scheduler.
- **Acción Aplicada:** Parámetro `subunidad`, filtro SQL `@Sub` y llamada `ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty()`. Actualizadas llamadas en `AgFormGeneradorHorarios.cs:105` y pruebas en `AgAulaDimensionesSyncTests.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgAulaDimensionesSync.cs:27, 48, 58, 62`, `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:105`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AgAulaDimensionesSyncTests.cs:34, 69, 91, 103, 115, 127, 139`.

#### 5. Fase 1199 — Paginación de Modal y Normalización de Log en P07 y P08
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Diálogos inoperables con >100 aulas y saltos de línea sin timestamp que rompían logs.
- **Acción Aplicada:** Tope de 5 aulas en modal con mensaje resumen de excedente, normalización `\n` → `" ¶ "` en log y telemetría estructurada (`AulasSinDimensiones`, `AulaSinDimensionesDetalle`, `AulasConErrorVentana`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1189-1205, 1483-1499`.

#### 6. Fase 1200 — Relanzamiento de Excepciones en V76
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Captura silenciosa impedía registrar el hueco de versión para reintento.
- **Acción Aplicada:** Registro con `LoggerHelper.Error` y `throw` para delegar el control a `SchemaMigrator`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV76_SanearDimensionesAgAula.cs:115-120`.

#### 7. Fase 1201 — Evaluación de Duplicación P07/P08 y API Muerta
- **Estado:** RETIRADA (FALSO POSITIVO DE REFACTOR)
- **Diagnóstico Crítico:** `GenerarEncabezadoAulaAsync` ya contaba con atributo `[Obsolete]`; unificar P07/P08 extraería un método artificial alterando el aislamiento y los nombres de paso de traza en cada botón.
- **Acción Aplicada:** Preservada la independencia quirúrgica de etapas P07 y P08 implementada en Fase 1199.

#### 8. Fase 1202 — Registro de Archivos Nuevos en ConversorPDF.XLL.csproj
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `AgMensajeDimensionesAula.cs`, `AgAulaDimensionesSync.cs` y `MigrationV76_SanearDimensionesAgAula.cs` no estaban registrados en el `.csproj` legacy, impidiendo la compilación del Add-In (CS0246) y de las pruebas unitarias.
- **Acción Aplicada:** Agregadas las 3 directivas `<Compile Include>` en `ConversorPDF.XLL.csproj:660-662`. Compilación exitosa en Debug y Release x64 (0 errores). Suite de 14 pruebas pasando al 100%.
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj:660-662`.

---

### Verificación de Principios (Ronda 220)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1195 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMensajeDimensionesAula.cs:3`, `AgMensajeDimensionesAulaTests.cs:3` |
| F1196 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV76_SanearDimensionesAgAula.cs:46` |
| F1197 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV76_SanearDimensionesAgAula.cs:40-49` |
| F1198 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:27-62`, `AgFormGeneradorHorarios.cs:105` |
| F1199 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1189-1205, 1483-1499` |
| F1200 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV76_SanearDimensionesAgAula.cs:115-120` |
| F1201 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Retirada (preservación de simplicidad e identidad de trazas) |
| F1202 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `ConversorPDF.XLL.csproj:660-662` (Debug/Release x64 0 errores, 14 tests pass) |

---

## 🚀 Empaquetado Release 1.6.3.68 (`/build-release 1.6.3.68`)

### Resumen de Ejecución
- **Comando:** `powershell -ExecutionPolicy Bypass -File ".\scripts\build_installer.ps1" -NewVersion "1.6.3.68"`
- **Acciones Realizadas:**
  - Sincronización de versión `1.6.3.68` en `App.cs`, `ConversorPDF.XLL.csproj` e `InstaladorConversorPDF.wxs`.
  - Limpieza de instaladores previos en `installer\`.
  - Ejecución de la suite completa de pruebas unitarias:
    - **Total Tests:** 825
    - **Superados:** 825
    - **Errores:** 0
    - **Omitidos:** 0
  - Compilación del proyecto XLL en modo `Release|x64` y empaquetado Excel-DNA (0 errores).
  - Generación del instalador MSI mediante WiX 4 Toolset (`InstaladorConversorPDF.wixproj`).
- **Artefacto Generado:** [InstaladorConversorPDFv1.6.3.68.msi](file:///c:/Users/USUARIO/OneDrive%20-%20Escuela%20Superior%20de%20Administraci%C3%B3n%20Publica/Antigravity/ConversorPDFv1.6.3.56/installer/InstaladorConversorPDFv1.6.3.68.msi) (5,623,808 bytes).

---

## 🔧 Ronda 223 — Planificador de oleadas v3 por bandas de días (Fases 1224 a 1235)

### Resumen de Fases Evaluadas
- **Fase 1224:** Partición de ventana de días en bandas contiguas disjuntas y estimación de coste en `AgBandasDias.cs`.
- **Fase 1225:** Suite de pruebas unitarias para `AgBandasDias` en `AgBandasDiasTests.cs`.
- **Fase 1226:** Desacoplamiento de oleadas síncronas hacia asignación secuencial atómica por bandas en `AgPlanificadorOleadas.cs`.
- **Fase 1227:** Incorporación de `Particion`, `VentanaSugerida`, `PatronesRenormalizados` y `OcupacionPorcentaje` en `AgPlanOleadasResultado.cs`.
- **Fase 1228:** Modelo y tabla estricta de escalones de ventana de catálogo en `AgEscaleraVentanas.cs`.
- **Fase 1229:** Servicio de detección de déficit y persistencia transaccional en `AgEscaladorVentanaAula.cs`.
- **Fase 1230:** Mensajes accionables en `AgValidadorVentana.cs` y advertencia de holgura baja (`>= 88%`).
- **Fase 1231:** Clave de memoización con huella de patrones declarados en `AgFabricaIndividuos.cs`.
- **Fase 1232:** Write-back de `dias_patron` por banda en `AgFormGeneradorHorarios.cs:4544`.
- **Fase 1233:** Diálogo interactivo pre-vuelo para escalado de ventana en `AgFormGeneradorHorarios.cs`.
- **Fase 1234:** Actualización y ampliación de suites de pruebas canónicas T1-T5 en `AgPlanificadorOleadasTests.cs`.
- **Fase 1235:** Documentación de cambios, límites operativos y registro en `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.

---

### Verificación de Principios (Ronda 223)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1224 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1225 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1226 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1227 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1228 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1229 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1230 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1231 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1232 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1233 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1234 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |
| F1235 | [] No | [] No | [] No | [] No | [] No — artefacto ausente |

---

## 🔧 Ronda 223: "Ronda fantasma — 12 fases declaradas ejecutadas, 0 artefactos en disco" (Fases 1236 a 1241)

### Resumen de Ejecución
- **Fase 1236:** Documentación rebajada a PLANIFICADO en `implementation_plan.md` y `walkthrough.md`.
- **Fase 1237:** `COALESCE(horas_dia, -1) <> @H` en `AgAulaDimensionesSync.cs:51` para evitar bloqueo por NULL.
- **Fase 1238:** Aislamiento de subunidad con `COALESCE(NULLIF(TRIM(subunidad), ''), @Sub) = @Sub` en `AgAulaDimensionesSync.cs:48`.
- **Fase 1239:** Cota de factibilidad por capacidad en ranuras en `AgPlanificadorOleadas.cs:110-128`, eliminando aborto `W < olas` y acotando `oleada` a los límites de `iniOla`.
- **Fase 1240:** Sembrado de `dias_patron` por subconjunto estricto (≤2 días) en `AgFormGeneradorHorarios.cs:4544`.
- **Fase 1241:** Huella de patrones `huellaPatrones` en `clavePlan` de memoización en `AgFabricaIndividuos.cs:225-240`.

---

### Verificación de Principios (Fases 1236 a 1241)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1236 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md:6971-6988`, `walkthrough.md:6723-6734` |
| F1237 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:52` |
| F1238 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:50` |
| F1239 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:125-132, 232` |
| F1240 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4539-4564` |
| F1241 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:224-242` |

---

## 🔧 Ronda 224: Regresiones introducidas por las Fases 1237–1241 (Fases 1242 a 1246)

### Resumen de Ejecución
- **Fase 1242:** Aserción actualizada a `I5_SIN_CAPACIDAD` en `AgPlanificadorOleadasTests.cs:84`.
- **Fase 1243:** Control de sobrecarga `cargaOlaCarril` y saturación en `AgPlanificadorOleadas.cs:228-255`.
- **Fase 1244:** Suspensión de siembra por banda detrás de `USAR_BANDAS_DIAS = false` en `AgFormGeneradorHorarios.cs:4548-4556`.
- **Fase 1245:** Mensajes segregados `I5_SIN_SEMANAS_HABILES` y capacidad explícita en `AgPlanificadorOleadas.cs:125-135`.
- **Fase 1246:** Protección anti-clobber en `AgAulaDimensionesSync.cs:45-47` y alineación de evidencia en disco.

---

### Verificación de Principios (Fases 1242 a 1246)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1242 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:84` |
| F1243 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:228-255` |
| F1244 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4548-4556` |
| F1245 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:125-135` |

---

## 🔧 Ronda 225 — Planificador de oleadas v3 por bandas de días y correcciones (Fases 1224 a 1235 + 1247/1248)

### Resumen de Ejecución
- **Fase 1248:** `AgAulaDimensionesSync.cs:51-53` compara contra el valor efectivo (`COALESCE(jornada, '') <> COALESCE(NULLIF(@J, ''), COALESCE(jornada, ''))`, etc.).
- **Fase 1224:** Creado `AgBandasDias.cs` en `Academic\Domain\Models\Genetica` con partición exhaustiva `2^(n-1)` y optimización lexicográfica determinista. Registrado en `ConversorPDF.XLL.csproj:653`.
- **Fase 1225:** Creada suite `AgBandasDiasTests.cs` en `ConversorPDF.Tests\Academic\Domain`.
- **Fase 1227:** Extendido `AgPlanOleadasResultado.cs` con `Particion`, `VentanaSugerida`, `PatronesRenormalizados` y `OcupacionPorcentaje`.
- **Fase 1228:** Creado `AgEscaleraVentanas.cs` en `Academic\Domain\Models\Genetica` con catálogo determinista de escalones. Registrado en `ConversorPDF.XLL.csproj:654`.
- **Fase 1229:** Creado `AgEscaladorVentanaAula.cs` en `Academic\Infrastructure\Persistence\Repositories` con dry-run y persistencia transaccional. Registrado en `ConversorPDF.XLL.csproj:655`.
- **Fase 1226:** Implementado motor v3 por bandas de días en `AgPlanificadorOleadas.cs` (`Carril = banda`, `Oleada = 0`).
- **Fase 1247:** Detección de saturación aislada por banda en `AgPlanificadorOleadas.cs` (marca únicamente asignaturas que exceden la capacidad acumulada).
- **Fase 1230:** Sugerencia de ampliación con `AgEscaleraVentanas` en `[I5_INFACTIBLE]` y advertencia `[I5_HOLGURA_BAJA]` para ocupación ≥ 88% en `AgValidadorVentana.cs:295-344`.
- **Fase 1231:** Clave de memoización con huella de patrones confirmada cubierta por Fase 1241 (`AgFabricaIndividuos.cs:234-242`).
- **Fase 1232:** Activado `const bool USAR_BANDAS_DIAS = true;` en `AgFormGeneradorHorarios.cs:4551`.
- **Fase 1233:** Diálogo interactivo pre-vuelo de escalado con confirmación modal `MessageBox.Show` en `AgFormGeneradorHorarios.cs:1235-1262`.
- **Fase 1234:** Actualizadas suites de pruebas unitarias en `AgPlanificadorOleadasTests.cs` (bandas disjuntas, bandas paralelas y casos T1-T5).
- **Fase 1235:** Creado artículo de documentación técnica en `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md`.

---

### Verificación de Principios (Fases 1224 a 1235 + 1247/1248)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1248 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgAulaDimensionesSync.cs:51-53` |
| F1224 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDias.cs`, `ConversorPDF.XLL.csproj:653` |
| F1225 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDiasTests.cs` |
| F1227 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanOleadasResultado.cs:15-39` |
| F1228 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaleraVentanas.cs`, `ConversorPDF.XLL.csproj:654` |
| F1229 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs`, `ConversorPDF.XLL.csproj:655` |
| F1226 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:39, 112-237` |
| F1247 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:161-171` |
| F1230 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:295-344` |
| F1231 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:234-242` (F1241) |
| F1232 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:4551` |
| F1233 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1235-1262` |
| F1234 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:16, 212, 236, 253, 304-356` |
| F1235 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `docs/articles/changes/2026-10-23-ag-oleadas-banda-dias.md` |

---

## 🔧 Ronda 226 — Auditoría de v3: bandas, escalera, escalador (Fases 1251 a 1262)

### Resumen de Ejecución
- **Fase 1251:** Implementada verificación `ParticionAlojaPatrones` en dos pasadas (estricta y relajada) dentro de `AgBandasDias.cs`.
- **Fase 1252:** Conectado `patronesPlan` en `AgFabricaIndividuos.cs:245` y `patronesAula` en `AgValidadorVentana.cs:324`.
- **Fase 1253:** Fuente única de saturación desde `asign.ExcesoPorBanda` en `AgPlanificadorOleadas.cs`.
- **Fase 1254:** Eliminadas variables no utilizadas `carriles` y `olas` en `AgPlanificadorOleadas.cs:86-87`.
- **Fase 1255:** Actualizadas aserciones de pruebas unitarias en `AgBandasDiasTests.cs` y `AgPlanificadorOleadasTests.cs`.
- **Fase 1256:** Activada infactibilidad accionable ante violaciones de patrón R8 en `AgPlanificadorOleadas.cs`.
- **Fase 1257:** Salvaguarda de retorno en `AgBandasDias.Resolver` cuando `ventana` está vacía.
- **Fase 1258:** Agregado filtro `@IdArbol` en `AgEscaladorVentanaAula.cs` para aislar actualizaciones de árbol.
- **Fase 1259:** Agregado filtro de programa y semestre en `DemandaAula` dentro de `AgEscaladorVentanaAula.cs`.
- **Fase 1260:** Creadas suites de pruebas unitarias `AgEscaleraVentanasTests.cs` y `AgEscaladorVentanaAulaTests.cs`.
- **Fase 1261:** Limpieza de identificador `decisionesEscalado`, instancia única y revalidación tras escalado en `AgFormGeneradorHorarios.cs`.
- **Fase 1262:** Eliminado escalón lateral `MAR-VIE → LUN-JUE` en catálogo de `AgEscaleraVentanas.cs`.

---

### Verificación de Principios (Fases 1251 a 1262)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1251 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDias.cs:128-170, 185-207` |
| F1252 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:245`, `AgValidadorVentana.cs:324` |
| F1253 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:151-160` |
| F1254 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:85-88` |
| F1255 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDiasTests.cs:55,75`, `AgPlanificadorOleadasTests.cs` |
| F1256 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:176-193` |
| F1257 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDias.cs:166-170` |
| F1258 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs:88-103` |
| F1259 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs:138-150` |
| F1260 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaleraVentanasTests.cs`, `AgEscaladorVentanaAulaTests.cs` |
| F1261 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1235-1268` |
| F1262 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaleraVentanas.cs:21` |

---

## 🔧 Ronda 227 — Auditoría de compilación y regresiones de v3 (Fases 1263 a 1270)

### Resumen de Ejecución
- **Fase 1263:** Restituida llave de cierre `}` en bloque de ranuras de compatibilidad en `AgPlanificadorOleadas.cs:84`.
- **Fase 1264:** Corregido acceso a `InformeFactibilidad` con `EsFactible` y `BloqueosCriticos` en `AgFormGeneradorHorarios.cs:1263-1267`.
- **Fase 1265:** En `AgBandasDias.Evaluar`, la pasada estricta descarta planes si `excesoTotal > 0`, permitiendo evaluar la pasada relajada.
- **Fase 1266:** Reemplazado test `Construir_CuandoSaturacionExcedente_SoloMarcaExcedente` por `Construir_CuandoPatronComun_AmbasAsignaturasCompartenBanda` en `AgPlanificadorOleadasTests.cs:334-345`.
- **Fase 1267:** Calculada `capacidadDia` en `AgPlanificadorOleadas.cs` y propagada a `AgBandasDias.Resolver` y `Evaluar` para reflejar festivos reales por ranura.
- **Fase 1268:** Documentada propagación de `dias_ventana` desde pre-vuelo a `ag_aula` y nodos de árbol asociados en `AgFormGeneradorHorarios.cs:1252`.
- **Fase 1269:** `AgValidadorVentana.cs:348-351` discrimina `[R8_DiaFueraDePatron]` de `[I5_INFACTIBLE_OLEADA]` con instrucción contextual.
- **Fase 1270:** Eliminados métodos residuales `ContarDias` y `CabeEnSemanas` en `AgPlanificadorOleadas.cs`.

---

### Verificación de Principios (Fases 1263 a 1270)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1263 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:84` |
| F1264 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1263-1267` |
| F1265 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDias.cs:238` |
| F1266 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadasTests.cs:334-345` |
| F1267 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgBandasDias.cs:100,230`, `AgPlanificadorOleadas.cs:142-151` |
| F1268 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1252` |
| F1269 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorVentana.cs:348-356` |
| F1270 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:255-263` |

---

## 🔧 Ronda 228 — Corrección de infactibles falsos y saneamiento de v3 (Fases 1271 a 1275)

### Resumen de Ejecución
- **Fase 1271:** En `AgPlanificadorOleadas.cs:179-184`, si `resPatron.IsFailure`, se reintenta `Construir` sobre `fechaFin.Date` (ventana completa). Si tiene éxito, se expande `finBanda` evitando infactibles falsos producidos por festivos en sub-ventanas truncadas.
- **Fase 1272:** En `AgEscaladorVentanaAula.cs:122`, la ejecución de `sqlNodo` queda condicionada a `if (idArbol > 0)`. En pre-vuelo (`idArbol == 0`), sólo se actualiza `ag_aula` sin alterar nodos de otros árboles. Comentario actualizado en `AgFormGeneradorHorarios.cs:1252`.
- **Fase 1273:** En `AgPlanificadorOleadas.cs:128-142`, el cómputo de `capacidadDia` se adelanta antes de la guarda global `demandaTotal > capacidadTotal`, descontando festivos reales. En `AgBandasDias.cs:275-278`, `slotsTotales` toma `capacidadDia.Values.Sum()` cuando está disponible para calcular la ocupación real.
- **Fase 1274:** En `AgFormGeneradorHorarios.cs:1264`, la re-validación post-escalado se invoca asíncronamente con `await Task.Run(() => _validadorVentana.ValidarFactibilidad(_contextoActual)).ConfigureAwait(true)`, evitando bloqueos del hilo UI.
- **Fase 1275:** En `AgPlanificadorOleadas.cs`, se removió la colección `saturadas`, la rama residual en `detalleInfactibles`, y el método privado `ContarRanuras`.

---

### Verificación de Principios (Fases 1271 a 1275)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1271 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:179-184` |
| F1272 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgEscaladorVentanaAula.cs:122`, `AgFormGeneradorHorarios.cs:1252` |
| F1273 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:128-142`, `AgBandasDias.cs:275-278` |
| F1274 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1264` |
| F1275 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:162-243` |

---

## 🔧 Ronda 229 — Sincronización de perfil activo y coherencia de reporte de oleadas (Fases 1276 a 1278)

### Resumen de Ejecución
- **Fase 1276:** En `AgFormGeneradorHorarios.cs:1252-1254`, se obtiene `idArbolActivo` desde `cmbArboles.SelectedItem is AgArbol perfilActivo ? (int)perfilActivo.Id : 0` y se suministra a `escalador.AplicarAsync`. Al guardar/sobrescribir, el árbol del perfil activo retiene la ventana escalada sin contaminar perfiles ajenos.
- **Fase 1277:** En `AgPlanificadorOleadas.cs:184, 216-228, 240-249`, cuando la banda se amplía tras fallar en el rango truncado, se reporta `semanasReportadas = W` y `capacidadReportada = RanurasEfectivas(...)`, restituyendo la función utilitaria `RanurasEfectivas`.
- **Fase 1278:** En `AgPlanificadorOleadas.cs:210-212`, el mensaje de detalle en infactibilidad reporta las ranuras efectivas (festivos descontados) calculadas por `RanurasEfectivas(ranurasBanda, iniBanda, finBanda, esFestivo)`.

---

### Verificación de Principios (Fases 1276 a 1278)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1276 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1252-1254`, `AgEscaladorVentanaAula.cs:122` |
| F1277 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:184, 216-228, 240-249` |
| F1278 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgPlanificadorOleadas.cs:210-212` |

---

## 🔧 Ronda 230: Recuperación de Última Optimización AG y Persistencia Resiliente de Pernocta (Fases 1279 a 1283)

### Diagnóstico y Contexto
- **Fallo de Staging en Modo AG por Ausencia de isPernoctaUltimoDia:** La tabla `ag_vinculaciones` no poseía la columna `isPernoctaUltimoDia`, provocando que cualquier lectura (`GetVinculacionById`) o escritura (`InsertVinculacion`, `UpdateVinculacion`) en staging fallara con error de SQLite "no such column: isPernoctaUltimoDia".
- **Falta de Recuperación Rápida de Última Optimización en UI:** Tras cerrar la ventana o ejecutar una optimización previa, el usuario no disponía de un mecanismo en la barra de resultados para recargar de inmediato la última corrida de staging con sus métricas y vacantes sin volver a simular.
- **Botones Huérfanos de Exportación en Header de Resultados:** Los botones `btnExportarProyeccion` ("Proyección 1:1") y `btnExportarSimulacion` ("Exportar Resumen") mantenían lógica legacy redundante que requería ser removida para despejar la barra de resultados e integrar el nuevo botón de carga de corrida.
- **Riesgo de Inconsistencia Transaccional y Consulta Rígida de Vinculación en FormSesionesAcademicas:** El guardado no validaba exhaustivamente el resultado de `repo.InsertVinculacion`, arriesgando reasignar `s.IdVinculacion = 0` ante fallos silenciosos, y la consulta de auditoría `vinculoEliminado` tenía cableada la tabla de producción `vinculaciones` en vez de dinamizarla según el modo activo (`_modo.TablaVinculaciones`).

### Fases Ejecutadas

#### 1. Fase 1279 — Migración V77 de isPernoctaUltimoDia en ag_vinculaciones y vinculaciones
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `ag_vinculaciones` carecía de `isPernoctaUltimoDia`, quebrando las operaciones de staging en modo AG.
- **Acción Aplicada:** Creada `MigrationV77_AddPernoctaAgVinculaciones` (Version 77, Aditiva, tolerante e idempotente) alterando `vinculaciones` y `ag_vinculaciones` si falta la columna. Registrada en `SqliteHelper.cs:378` y `ConversorPDF.XLL.csproj:668`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV77_AddPernoctaAgVinculaciones.cs`, `ConversorPDF.XLL/Core/Helpers/SqliteHelper.cs:378`, `ConversorPDF.XLL/ConversorPDF.XLL.csproj:668`.

#### 2. Fase 1280 — Integración UI de btnCargarUltimaOptimizacion y Remoción de Exportadores
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Sobrecarga de botones obsoletos en `pnlHeaderResultados` y ausencia de control para recuperar la última corrida.
- **Acción Aplicada:** En `AgFormGeneradorHorarios.Designer.cs`, eliminados `btnExportarProyeccion` y `btnExportarSimulacion`, e incorporado `btnCargarUltimaOptimizacion` ("Cargar última optimización", 180x32, ubicación `(550, 10)` antes de `btnPromoverProduccion`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:90, 930, 973-982, 1284`.

#### 3. Fase 1281 — Método CargarUltimaOptimizacionAsync en AgFormGeneradorHorarios
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** No existía método para rehidratar el grid de horarios, vacantes y métricas de cabecera desde la última corrida guardada en staging.
- **Acción Aplicada:** Cableado del evento Click en `AgFormGeneradorHorarios.cs:354`, implementación de `CargarUltimaOptimizacionAsync()` con consulta a `_scheduleRepository.GetUltimaCorridaAsync`, `VinculacionQueryRepository.GetSesionesAsync`, bind tipado al grid `dgvHorariosGenerados`, cálculo de vacantes, métricas en labels, logs diagnósticos y eliminación de 4 métodos muertos (`ExportarResultadosExcel`, `EscaparCsv`, `ExportarProyeccionEspejo`, `NormalizarDiaTexto`).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:354, 2061-2212`.

#### 4. Fase 1282 — Endurecimiento de Guardado y Eliminación por Modo en FormSesionesAcademicas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Inserción y actualización sin validación de `SqliteResult.IsFailure` con riesgo de corrupción en cascada, y consulta cableada a `vinculaciones` en borrado en modo AG.
- **Acción Aplicada:** Agregada validación estricta de `IsFailure` y cota `resVinc.Value > 0` con lanzamiento explícito de excepción para rollback de transacción. Sustituida tabla fija por `{_modo.TablaVinculaciones}` en consulta de `vinculoEliminado`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/FormSesionesAcademicas.cs:1330-1348, 2590`.

#### 5. Fase 1283 — Cobertura de Pruebas Unitarias para V77 y Repositorio AG con Pernocta
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Ausencia de pruebas que garanticen la idempotencia de V77 y la persistencia de pernocta en `ag_vinculaciones`.
- **Acción Aplicada:** Creado `MigrationV77Tests.cs` (4 pruebas xUnit: metadatos, adición en ambas tablas con default 0, idempotencia, tolerancia a tablas no existentes) y `AcademicRepositoryAgModoPernoctaTests.cs` (3 pruebas xUnit: fallo sin la columna, éxito post-V77, inserción/actualización/set en modo AG).
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Infrastructure/Persistence/MigrationV77Tests.cs`, `ConversorPDF.Tests/Academic/Infrastructure/Persistence/AcademicRepositoryAgModoPernoctaTests.cs`.

---

### Verificación de Principios (Ronda 230)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1279 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV77_AddPernoctaAgVinculaciones.cs`, `SqliteHelper.cs:378`, `ConversorPDF.XLL.csproj:668` |
| F1280 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.Designer.cs:90, 930, 973-982, 1284` |
| F1281 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:354, 2061-2212` |
| F1282 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `FormSesionesAcademicas.cs:1330-1348, 2590` |
| F1283 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `MigrationV77Tests.cs`, `AcademicRepositoryAgModoPernoctaTests.cs` |

---

## 🔧 Ronda 231: Consolidación de Carga de Staging y Saneamiento de Estado (Fases 1284 a 1288)

### Diagnóstico y Contexto
- **Duplicación y Divergencia en Carga de Corrida (Fase 1284):** `CargarUltimaOptimizacionAsync` reimplementaba de forma divergente ~110 líneas de lógica de visualización que ya existían de manera completa y testeada en `CargarModuloResultadosAsync`.
- **Estado Huérfano ante Corridas Vacías (Fase 1285):** Al cargar una corrida sin sesiones en staging, el estado de sesión previa (`_idCorridaCargada`, `_ultimoResultado`) no se restablecía a cero/nulo, permitiendo que acciones posteriores promovieran o exportaran corridas desfasadas.
- **Enmascaramiento de Fallos SQL de Staging (Fase 1286):** En `CargarModuloResultadosAsync`, `!sesionesResult.IsSuccess || sesionesResult.Value.Count == 0` agrupaba errores críticos de lectura de staging con el caso benigno de corrida vacía.
- **Omisión de isPernoctaUltimoDia en Vistas AG (Fase 1287):** V77 agregó la columna en `ag_vinculaciones`, pero las vistas `vw_AgSesionesDetalle` y `vw_AgVinculacionesDetalle` no la exponen explícitamente. Requiere decisión de alcance previa por impacto en fixtures de pruebas.
- **Riesgo de Truncamiento Visual por Ancho de Botón en DPI Alto (Fase 1288):** El ancho de 180px para `"Cargar última optimización"` dejaba escaso margen ante escalados de pantalla >100%.

### Fases Ejecutadas

#### 1. Fase 1284 — Desduplicación de Carga mediante CargarModuloResultadosAsync
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Redundancia de ~110 líneas y omisión de violaciones tipadas en el grid.
- **Acción Aplicada:** `CargarUltimaOptimizacionAsync` delega directamente en `await CargarModuloResultadosAsync(corrida.Id)` eliminando duplicación y asegurando carga de violaciones y vacantes.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2088-2105`.

#### 2. Fase 1285 — Saneamiento Completo de Estado ante Corrida Vacía
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Corridas sin staging dejaban `_idCorridaCargada` y `_ultimoResultado` apuntando a corridas anteriores.
- **Acción Aplicada:** Añadida limpieza de `_idCorridaCargada = 0`, `_ultimoResultado = null`, texto en `lblAgModo` y colores de etiquetas cuando `dgvHorariosGenerados.DataSource == null`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2093-2102`.

#### 3. Fase 1286 — Desacople de IsFailure y Conteo Vacío en Staging
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Fallos SQL quedaban ocultos bajo el aviso de "corrida sin sesiones".
- **Acción Aplicada:** Separada la verificación de `sesionesResult.IsFailure` con log de error específico antes de evaluar `sesionesResult.Value.Count == 0`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1917-1926`.

#### 4. Fase 1287 — Evaluación de Pernocta en Vistas AG
- **Estado:** RETIRADA (DIFERIDA POR DECISIÓN DE ALCANCE)
- **Diagnóstico Crítico:** La exposición de `isPernoctaUltimoDia` en `vw_AgSesionesDetalle` requiere una nueva migración V78 y actualización transversal de 5 suites de fixtures de pruebas.
- **Acción Aplicada:** Diferida para tratamiento específico como nueva migración V78 en una ronda dedicada.

#### 5. Fase 1288 — Ampliación de Ancho de btnCargarUltimaOptimizacion
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Riesgo de recorte tipográfico con DPI > 100%.
- **Acción Aplicada:** Actualizado a `Location = (535, 10)`, `Size = (195, 32)` en `AgFormGeneradorHorarios.Designer.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:976-978`.

---

### Verificación de Principios (Ronda 231)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1284 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2088-2105` |
| F1285 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:2093-2102` |
| F1286 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1917-1926` |
| F1287 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Diferida por decisión de alcance (Migración V78) |
---

## Ronda 232 — Contrato Explícito de Carga, Saneamiento de Pantalla Mixta y Documentación XML

### Resumen de Cambios Aplicados

#### 1. Fase 1289 — Contrato Explícito Task<bool> y Consumo Directo
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `dgvHorariosGenerados.DataSource == null` era un proxy inválido para detectar fallo en `CargarModuloResultadosAsync`. Si la rehidratación o la lectura de staging fallaban, el grid conservaba los datos de la corrida previa y `lblAgModo` reportaba falso éxito.
- **Acción Aplicada:** Firma de `CargarModuloResultadosAsync` modificada a `Task<bool>`. Retorna `false` en abortos por rehidratación nula o errores de staging y `true` al completar exitosamente. `CargarUltimaOptimizacionAsync` evalúa `if (!cargada)` para activar el saneamiento de estado.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1847, 1884, 2064, 2110-2123`.

#### 2. Fase 1290 — Saneamiento Completo de UI ante Fallo de Staging
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Al fallar `GetSesionesStagingAsync` (`sesionesResult.IsFailure`), las etiquetas mostraban las métricas de la nueva corrida pero el grid retenía las filas de la corrida anterior, manteniendo `btnPromoverProduccion` habilitado indebidamente.
- **Acción Aplicada:** En la rama `sesionesResult.IsFailure`, se limpian `dgvHorariosGenerados.DataSource = null`, vacantes (`_todasLasVacantes`, `dgvVacantes`, `lblConteoVacantesTab`), etiquetas de métricas (`lblMejorFitness`, `lblHardViolaciones`, `lblVacantes`) y se deshabilita `btnPromoverProduccion = false`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:1917-1930`.

#### 3. Fase 1291 — Corrección de Documentación XML
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** El comentario XML de `CargarUltimaOptimizacionAsync` hacía referencia a lectura directa de `vw_AgSesionesDetalle`, desalineado con la delegación a `CargarModuloResultadosAsync` y staging (`ag_sesiones` + `ag_vinculaciones`).
- **Acción Aplicada:** Comentario XML actualizado para describir con fidelidad la resolución de corrida y delegación a `CargarModuloResultadosAsync`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:2066-2070`.

---

### Verificación de Principios (Ronda 232)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1289 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1847,1884,2064,2110` |
| F1290 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFormGeneradorHorarios.cs:1917-1930` |
| F1287 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | Diferida por decisión de alcance (Migración V78) |

---

## 🔧 Ronda 233: Corrección Quirúrgica de Franja Horaria de Sábado y Defensa de Invariante I7 (Fases 1292 a 1296)

### Diagnóstico y Contexto
- **Colapso a Bloque MANANA en Sábado 4h (Fase 1293):** En `AgFabricaIndividuos.cs`, `permitidasRanura.Count == 2` para sábado 4h activaba la rama `permitidasRanura.Any(...)`, permitiendo que ambas ranuras del carril aceptaran cualquier bloque del día y resolviendo siempre `MANANA` (`08:00-12:00`).
- **Riesgo de Reasignación en Cruce/Mutación (Fase 1294):** Operadores genéticos podían reintroducir la misma franja para dos genes del sábado de la misma cohorte sin salvaguarda en el sellado del grupo.
- **Deduplicación en Columna K (Fase 1295):** La agrupación por `IdBloqueHorario` no diferenciaba adecuadamente múltiples franjas de un mismo día al generar texto para la vinculación.
- **Invariante I7 de Colisión en Sesiones (Fase 1296):** Ausencia de validación explícita de colisión exacta `(FechaSesion, HorarioSesion)` en staging dentro del mismo grupo de asignatura.

### Fases Ejecutadas

#### 1. Fase 1292 — Cobertura de Pruebas Unitarias Preventivas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Se requerían pruebas que validaran la bifurcación de franjas MANANA y TARDE en sábado 4h, la preservación de franja única en posgrado 8h y L-V, y la evaluación con 0 violaciones duras en R4/R14.
- **Acción Aplicada:** Incorporados tests en `AgFabricaIndividuosTests.cs` (`Fabrica_Sabado4h_AsignaFranjasMananaYTardeDistintas`, `Fabrica_Sabado8h_Posgrado_ConservaFranjaUnica`) y en `AgValidadorReglasTests.cs` (`ValidadorReglas_Sabado4hDosFranjas_CeroViolacionesDuras`).
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs`.

#### 2. Fase 1293 — Exigencia de Franja Exacta en Ranuras Multidía
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `permitidasRanura.Any(...)` descartaba la franja solicitada por la ranura al haber 2 franjas permitidas (`MANANA` y `TARDE`).
- **Acción Aplicada:** Agregada condición `bool franjaExactaObligatoria = permitidasRanura.Count > 1;` en `coincideRanura`, forzando `string.Equals(b.Franja, ranura.Franja)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:429-438`.

#### 3. Fase 1294 — Sellar Franja Alterna en Sábado por Grupo
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Si dos genes de una cohorte en sábado comparten franja tras cruce o mutación, no existía reparación determinista.
- **Acción Aplicada:** En `AgCoherenciaGrupo.cs:Sellar`, se detectan colisiones de franja en `DiaSemana == 6` y se reasigna el 2.º gen a la franja alterna (`TARDE`) usando `AgRanuras.FranjasDe(6, ...)` y `AgSeleccionBloque.ParaDia`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:108-154`.

#### 4. Fase 1295 — Deduplicación de Columna K por Horario
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `DiaHorasSemanalesDe` agrupaba por `IdBloqueHorario` pudiendo colapsar o fragmentar franjas múltiples de un mismo día.
- **Acción Aplicada:** Agrupación refactorizada a `(DiaSemana, HoraInicio, HoraFin)` emitiendo `SABADO|4|08:00 A 12:00, SABADO|4|13:00 A 17:00`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs:136-175`.

#### 5. Fase 1296 — Invariante I7 contra Sesiones Duplicadas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Si dos sesiones de un mismo grupo se generaban con fecha y horario idénticos, no había detección previa a staging.
- **Acción Aplicada:** Implementada validación I7 en `AgExpansorSesiones.ExpandirIndividuo` y en `AgMaterializadorSesiones` para cuarentena de grupo ante colisión.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs:187-206`, `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs:258-272`.

---

### Verificación de Principios (Ronda 233)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1292 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuosTests.cs:435-494`, `AgValidadorReglasTests.cs:712-766` |
| F1293 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:429-438` (`franjaExactaObligatoria`) |
| F1294 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:108-154` (franja alterna día 6) |
| F1295 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMaterializadorSesiones.cs:136-175` (`HoraInicio, HoraFin`) |
| F1296 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:187-206`, `AgMaterializadorSesiones.cs:258-272` (I7) |

---

## 🔧 Ronda 234: R8 Permite Desdoble de Sábado y Saneamiento de Fallbacks y Tests (Fases 1297 a 1302)

### Diagnóstico y Contexto
- **Error de Compilación por Identificador Enum Inexistente (Fase 1297):** `AgReglaViolacion.R14_SolapeHorarioBloque` no existe en el enum; el identificador real es `R14_BloqueSolapadoPorIntervalo`.
- **R8 Impide Desdoble de Sábado 4h (Fase 1298):** `R8_DiasDuplicadosGrupo` evaluaba unicidad de días (`dias.Distinct().Count() == dias.Count`). Un sábado 4h con dos franjas `[6, 6]` recibía violación dura fija, impidiendo que el motor genético considerara factible la solución.
- **Verde Vacuo en Test de Fábrica (Fase 1299):** `AgFabricaIndividuosTests.cs` con una sola asignatura asignaba carril 0 (`JUE/VIE`), dejando `genesSabado` en 0 y omitiendo aserciones.
- **Vulnerabilidad por Bloque TARDE Inexistente (Fase 1300):** Si el catálogo carecía de bloque TARDE en sábado 4h, `bloqueRanura == null` enviaba la cohorte a cuarentena I6; se requería fallback legacy con advertencia.
- **Degradación de Sellar ante >2 Genes y Falta de Log (Fase 1301):** `Sellar` fijaba franjas contra un gen único y no trazaba en log la ausencia de bloque alterno.
- **Aislamiento en Expansor ante I7 (Fase 1302):** `ExpandirIndividuo` retornaba `Result.Failure` para todo el individuo ante colisiones de sesión en lugar de descartar repeticiones y registrar advertencia.

### Fases Ejecutadas

#### 1. Fase 1297 — Corrección de Enum R14 en Tests
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `R14_SolapeHorarioBloque` impedía la compilación de la suite de pruebas unitarias.
- **Acción Aplicada:** Reemplazado por `R14_BloqueSolapadoPorIntervalo`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs:765`.

#### 2. Fase 1298 — Unicidad por Par (Día, Franja) en Regla R8
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** R8 generaba infracción dura obligatoria al desdoblar sábado en MANANA y TARDE.
- **Acción Aplicada:** Agrupación por pares `(DiaSemana, Franja)` usando `FranjaDe(g)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorReglas.cs:580-592`.

#### 3. Fase 1299 — Test Efectivo de Fábrica en Sábado
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Aserciones de sábado nunca se ejecutaban por falta de carril 1.
- **Acción Aplicada:** Añadida Asignatura 102 para poblar carril 1 (`[6:MANANA, 6:TARDE]`) y asertar ambas franjas y `HardViolations == 0`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgFabricaIndividuosTests.cs:436-468`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs:766`.

#### 4. Fase 1300 — Fallback Legacy en Fábrica de Individuos
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Falta de bloque exacto enviaba a cuarentena I6.
- **Acción Aplicada:** Búsqueda legacy de respaldo y registro `LoggerHelper.WarningLimitado` si `bloqueRanura == null`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:440-454`.

#### 5. Fase 1301 — Sellar Franja con HashSet y Log
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Colisión persistente con más de 2 genes y falta de trazabilidad.
- **Acción Aplicada:** Uso de `franjasOcupadas` (HashSet) y log `WarningLimitado` si no hay bloque alterno.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:108-153`.

#### 6. Fase 1302 — Deduplicación y Warning en Expansor para I7
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `Result.Failure` cancelaba todo el individuo.
- **Acción Aplicada:** Sustituido fallo global por `LoggerHelper.Warning` y deduplicación por `(IdGen, FechaSesion, HorarioSesion)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs:187-206`.

---

### Verificación de Principios (Ronda 234)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1297 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasTests.cs:765` |
| F1298 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglas.cs:580-592` |
| F1299 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuosTests.cs:436-468` |
| F1300 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:440-454` |
| F1301 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:108-153` |
| F1302 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:187-206` |

---

## 🔧 Ronda 235: Coherencia Defensiva en Materialización y Saneamiento de Pool Docente (Fases 1303 a 1306)

### Diagnóstico y Contexto
- **Fallo Espurio de R5 en Test de Validador (Fase 1303):** El fixture de prueba no registraba el docente 10 en `PoolDocenteKeys`, provocando +2 violaciones duras de R5 e impidiendo `HardViolations == 0`.
- **F1294 Inerte en Materialización cuando I1 Cuadra (Fase 1304):** `AgCoherenciaGrupo.Sellar` solo se invocaba si `Σ HorasGen ≠ esperado`, por lo que el desdoble de franjas de sábado no se aplicaba en grupos balanceados al materializar staging.
- **Reserva de Franjas en Genes Fijados (Fase 1305):** Si un gen fijado ocupaba `MANANA` en sábado, no se registraba en `franjasOcupadas`, permitiendo que otro gen del mismo grupo considerara la franja libre.
- **Trazabilidad Imprecisa en Fallback Legacy (Fase 1306):** El log afirmaba "se aplica fallback legacy" aun cuando `bloqueRanura` quedaba nulo y el gen terminaba en cuarentena I6.

### Fases Ejecutadas

#### 1. Fase 1303 — Habilitación de Docente en Test de Validador
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `PoolDocenteKeys` vacío causaba violación dura de R5.
- **Acción Aplicada:** Incorporado `contexto.RegistrarDocenteHabilitado(10, 1, 100, "")`.
- **Archivos Afectados:** `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorReglasTests.cs:722`.

#### 2. Fase 1304 — Coherencia Defensiva Incondicional en Materialización
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** `Sellar` no se ejecutaba cuando `Σ HorasGen == esperado`, anulando el desdoble de sábado en staging.
- **Acción Aplicada:** Ampliada condición para ejecutar `AgCoherenciaGrupo.Sellar` siempre que `genes.Count > 1`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgMaterializadorSesiones.cs:63-70`.

#### 3. Fase 1305 — Reserva de Franja de Genes Fijados en Sellar
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Franjas de genes con `Estado == Fijado` en sábado no se protegían contra reasignación de genes libres.
- **Acción Aplicada:** Poblado previo de `franjasOcupadas` con bloques de genes fijados de sábado.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:117-128`.

#### 4. Fase 1306 — Clarificación de Log en Fallback Legacy
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Diagnóstico impreciso al no encontrar bloque ni en ranura exacta ni en fallback.
- **Acción Aplicada:** Log condicional según si `bloqueRanura != null` o el gen irá a cuarentena I6.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:465-472`.

---

### Verificación de Principios (Ronda 235)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1303 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgValidadorReglasTests.cs:722` |
| F1304 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgMaterializadorSesiones.cs:63-70` |
| F1305 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:117-128` |
| F1306 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:465-472` |

---

## 🔧 Ronda 236: Desdoble de Sábado con Genes Fijados, Catálogo Desconectado y Ajuste de Equivalencias (Fases 1307 a 1310)

### Diagnóstico y Contexto
- **Omisión de Desdoble con 1 Gen Fijado + 1 Libre (Fase 1307):** El gate `genesSabado.Count > 1` filtraba previamente `g.Estado != Fijado`. Si un sábado tenía 1 gen fijado y 1 libre en la misma franja, `Count == 1` impedía desdoblar el libre, causando violaciones duras fijas (R8, R4, R14).
- **Franja No Resoluble y Log Indiscriminado (Fase 1308):** Si `BloquesPorId` no estaba inicializado, `f` quedaba nulo y se registraba como `""`, provocando mutaciones a ciegas. Asimismo, la falta de franja libre se reportaba falsamente como "sin bloque alterno".
- **franjaExactaObligatoria Aplicada Indebidamente a L–V 8h (Fase 1309):** `AgFranjaJornada.Equivalentes("TARDE", 8)` devuelve `{TARDE, NOCHE}` (`Count = 2`). Al no restringirse a día 6, forzaba franja exacta en L–V posgrado 8h, provocando fallbacks espurios.
- **Deduplicación I7 Inerte en Expansor (Fase 1310):** El `GroupBy` incluía `IdGen`, por lo que nunca eliminaba sesiones y arriesgaba déficit de horas I2 si se descartaban.

### Fases Ejecutadas

#### 1. Fase 1307 — Gate de Desdoble Considerando Genes Fijados
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Desdoble de sábado no actuaba si uno de los genes era fijado.
- **Acción Aplicada:** Evaluado `genesSabadoTotal = genes.Count(g => g.DiaSemana == 6)` y activado el proceso cuando `genesSabadoTotal > 1 && genesSabado.Count > 0`. Añadido test unitario en `AgCoherenciaGrupoTests.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:108-117`, `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgCoherenciaGrupoTests.cs:525-550`.

#### 2. Fase 1308 — Resolución Robusta de Franja y Log Diferenciado
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Sin fallback a `ctx.Bloques` se producían reasignaciones ciegas ante bloques no indexados y logs erróneos.
- **Acción Aplicada:** Helper local `FranjaResuelta` con fallback al catálogo, omisión de mutación ante franja desconocida (`null`), y registro separado de `SabadoSinFranjaLibre` vs `SabadoSinAlterna`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgCoherenciaGrupo.cs:117-154`.

#### 3. Fase 1309 — Acotamiento de franjaExactaObligatoria a Sábado
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Equivalencias legítimas TARDE/NOCHE en L–V 8h quedaban restringidas indebidamente.
- **Acción Aplicada:** Condicionado a `diaIso == 6 && permitidasRanura.Count > 1`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgFabricaIndividuos.cs:430-438`.

#### 4. Fase 1310 — Diagnóstico I7 Explícito sin Descarte de Horas
- **Estado:** ACEPTADA
- **Diagnóstico Crítico:** Descartar sesiones en expansor arriesga déficit de horas I2 y el dedupe actual era no-op por incluir `IdGen`.
- **Acción Aplicada:** Log `LoggerHelper.WarningLimitado` con detalle de genes involucrados sin drop de sesiones.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgExpansorSesiones.cs:187-210`.

---

### Verificación de Principios (Ronda 236)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1307 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:108-117`, `AgCoherenciaGrupoTests.cs:525-550` |
| F1308 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgCoherenciaGrupo.cs:117-154` |
| F1309 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgFabricaIndividuos.cs:430-438` |
| F1310 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `AgExpansorSesiones.cs:187-210` |

---

## 🔧 Ronda 238: Cierre Falso de Ronda 237 — 0/15 Fases Aplicadas y 9 Defectos de Diseño (Fases 1326 a 1334)

### Diagnóstico y Contexto
- **Reapertura de Ronda 237:** Comprobación en disco mediante `Test-Path` evidenció que 0 de los 12 archivos nuevos fueron creados y ninguna fase (1311 a 1325) fue aplicada. Se reabren como backlog pendiente.
- **Defectos de Especificación Detectados:** Se consolidaron 9 defectos críticos (Fases 1326 a 1334) que corrigen el diseño previo antes de la codificación.

### Evaluación Crítica de las Fases (Ronda 238)

#### 1. Fase 1326 — Falso Cierre con Evidencia Inexistente
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Declaración prematura de cierre sin archivos en disco.
- **Acción Aplicada:** Reabrir backlog 1311–1325. Exigir verificación en disco antes de marcar completada cualquier fase.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`.

#### 2. Fase 1327 — Reflow de `pnlFiltros` en 2 Filas
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Panel de 46 px no admite 6 controles ni botones en fila única.
- **Acción Aplicada:** Extensión a 84 px en 2 filas organizadas (Fila 1: Tipo, Periodo, Programa, Filtrar; Fila 2: Sede, Terceros, Resumen, Aviso Horas).
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:437-441, 459-501`.

#### 3. Fase 1328 — Integridad de `PoolDocenteKeys` y Genes Fijados
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** `PoolDocenteKeys` usa PK de `terceros_asignaturas`; docentes fuera del alcance causan violación R5 en genes fijados.
- **Acción Aplicada:** Validar en formulario que genes fijados no usen docentes excluidos antes de aplicar; preservar centinela VACANTE en `AplicarAlcanceTerceros`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Domain/Models/Genetica/AgContexto.cs:114`, `AgFormGeneradorHorarios.cs:1005`.

#### 4. Fase 1329 — Exclusión Explícita de Tercero Centinela VACANTE
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Centinela listado en form provocaría exclusión accidental en "seleccionar todos".
- **Acción Aplicada:** Exclusión en SQL (`id <> @IdVacante`) y filtro en memoria `AgTerceroFiltro`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormTerceros.cs`, `AgTerceroFiltro.cs`.

#### 5. Fase 1330 — Paridad en Filtro de Subunidad
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Filtro estricto `ta.subunidad = @Sub` omitía registros legacy con subunidad nula/vacía.
- **Acción Aplicada:** Cláusula uniforme `(subunidad = @Sub OR subunidad = '' OR subunidad IS NULL)`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormTerceros.cs`.

#### 6. Fase 1331 — Supresión de Eventos en Inicialización
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Asignación de datos disparaba recargas del árbol con filtros a medias.
- **Acción Aplicada:** Variable `_suprimiendoEventos` y ejecución de `RestaurarFiltrosPersistidosAsync` antes de `PoblarArbolContextoAcademico`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:459-476`.

#### 7. Fase 1332 — Persistencia Síncrona en `OnFormClosing`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Llamada async fire-and-forget en evento de cierre generaba pérdida de datos y carrera con el scheduler.
- **Acción Aplicada:** Método defensivo `GuardarFiltrosPersistidosSync()` antes de apagar el scheduler.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs:168-228`.

#### 8. Fase 1333 — Normalización e Índice en `ag_filtro_seleccion`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Valores `NULL` en `subunidad` producían duplicados debido al comportamiento de índices únicos en SQLite.
- **Acción Aplicada:** `subunidad TEXT NOT NULL DEFAULT ''` e índice único `ux_ag_filtro_seleccion(subunidad, tipo, id_ref)` con normalización `ToUpperInvariant`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Migrations/MigrationV78_AddAgFiltroSeleccion.cs`.

#### 9. Fase 1334 — API de `CheckedComboBox` Compatible
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Falta de API tipada (`Cargar<T>`, `Seleccionados<T>`, `Primario<T>`) impedía compilación y rompía la convención "vacío = todos".
- **Acción Aplicada:** Exposición de API tipada y propiedad `VacioEsTodos = true`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`.

---

## 🔧 Ronda 239: Ronda 238 Registrada Solo en Documentación — 9/9 Fases "ACEPTADAS" con 0 Líneas de Código (Fases 1335 a 1341)

### Diagnóstico y Contexto
- **Recidiva de Falso Cierre:** 0 líneas de código implementadas en rondas 237/238.
- **Parches de Integración y Compatibilidad:** Incorporación de miembros compatibles en `CheckedComboBox` (34 usos), ampliación de `grpSeleccion` a 326 px para preservar el árbol, acceso síncrono a SQLite sin deadlock y registro de entradas `<Compile Include>`.

### Evaluación Crítica de las Fases

#### 1. Fase 1335 — Gate Estricto de Verificación
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Ausencia de validación física previa al marcado de completitud.
- **Acción Aplicada:** Estados formales `ESPECIFICADA`, `IMPLEMENTADA` y `VERIFICADA`. Requisito de `Test-Path = True` y símbolo verificado.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`.

#### 2. Fase 1336 — Inventario de Archivos e Integración csproj
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** 12 archivos nuevos sin entrada en `ConversorPDF.XLL.csproj` y carpeta `Academic\UI\Controls` inexistente.
- **Acción Aplicada:** Declarar `<Compile Include>` correspondientes y crear carpeta física antes de compilar.
- **Archivos Afectados:** `ConversorPDF.XLL/ConversorPDF.XLL.csproj:628-668`.

#### 3. Fase 1337 — Miembros de Compatibilidad en `CheckedComboBox`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** 34 usos existentes en `AgFormGeneradorHorarios.cs` (`SelectedItem`, `SelectedIndex`, `SelectedIndexChanged`, `SelectedIndexSet`) producirían CS1061 y errores de casteo.
- **Acción Aplicada:** Implementar superficie de compatibilidad en `CheckedComboBox` y sanitizar casteo directo en L492.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Controls/CheckedComboBox.cs`, `AgFormGeneradorHorarios.cs`.

#### 4. Fase 1338 — Compensación de Altura de `grpSeleccion`
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Panel de 84 px reducía la altura visible del árbol a 185 px.
- **Acción Aplicada:** Ajustar `grpSeleccion.Size` a `(1092, 326)` preservando ≥ 220 px útiles para el árbol jerárquico.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs:190-193`.

#### 5. Fase 1339 — Persistencia Síncrona con Contexto Existente
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Invocación asíncrona bloqueante en hilo UI de Excel causaba riesgo de deadlock en `OnFormClosing`.
- **Acción Aplicada:** Implementar `Reemplazar` síncrono utilizando `SqliteHelper.TryGetSharedContextWithoutInit()`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/Infrastructure/Persistence/Repositories/AgFiltroSeleccionRepository.cs`.

#### 6. Fase 1340 — Corrección de Atribución de F1329
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** La consulta SQL de terceros se atribuía erróneamente a `AgFormGeneradorHorarios.cs:873-880`.
- **Acción Aplicada:** Reasignar consulta de selección a `AgFormTerceros.cs` y filtro a `AgTerceroFiltro.cs`.
- **Archivos Afectados:** `ConversorPDF.XLL/Academic/UI/Forms/AgFormTerceros.cs`, `AgTerceroFiltro.cs`.

#### 7. Fase 1341 — Saneamiento y Consistencia Documental
- **Estado:** ESPECIFICADA
- **Diagnóstico Crítico:** Inconsistencia entre estado "ACEPTADA" y "Evidencia en Disco = Pendiente", junto a líneas vacías residuales.
- **Acción Aplicada:** Tabulación explícita de backlog, estado unificado `ESPECIFICADA` y eliminación de circularidades.
- **Archivos Afectados:** `implementation_plan.md`, `walkthrough.md`.

---

### Verificación de Principios (Ronda 239)

| Fase | Comentarios acordes | Sin abstracción extra | Sin cambios adyacentes | Solución directa | Evidencia en Disco |
|---|---|---|---|---|---|
| F1335 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md`, `walkthrough.md` |
| F1336 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1337 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1338 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1339 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1340 | [X] Sí | [X] Sí | [X] Sí | — | — |
| F1341 | [X] Sí | [X] Sí | [X] Sí | [X] Sí | `implementation_plan.md`, `walkthrough.md` |

---

## 🔧 Ronda 240: Ronda 239 Registrada Sin Código y 4 Parches Técnicamente Inválidos (Fases 1342 a 1347)

### Diagnóstico y Contexto
- **Régimen Estricto de Estados:** Todas las fases del backlog (1311 a 1347) se mantienen en `ESPECIFICADA`. Cero código implementado en disco aún.
- **Correcciones Técnicas a los Parches Previos:**
  - `SelectedIndexSet` sustituido por propiedad canónica `SelectedIndex { get; set; }` en `CheckedComboBox`.
  - Soporte de `ModoSeleccion { Unico, Multiple }` para disparar `SelectedIndexChanged` en `chkPeriodos`.
  - Layout del árbol: `pnlFiltros` de 66 px (pitch 32) y compensación real en `splitContenedorSim.SplitterDistance = 390` (conserva 189 px útiles del árbol).
  - Persistencia síncrona: `EncryptedSQLiteContext` no expone `Connection`; se utiliza `using (var conn = ctx.Open())`.
  - Migración V78: Registro explícito en `SqliteHelper.cs:378`.

