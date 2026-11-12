# Log del Validador de Ventana con Nombre de Sede (Fase 1579)

## 1. Propósito
Los bloqueos y advertencias del pre-vuelo (`AgValidadorVentana`) identificaban el aula únicamente por
`AulaRef` (`idPeriodo|idSede|idPrograma|semestre|grupo|modalidad`). En corridas multi-sede el usuario veía
mensajes como:

```
[I5_INFACTIBLE] Aula '3|13|1|VI|A|MIXTA': demanda total de 64 sesiones ...
[R8_DiaFueraDePatron] Aula '3|13|1|I|A|PRESENCIAL': asignaturas [1,2,3,4,5,6] ...
```

y debía decodificar el segundo segmento del `AulaRef` para saber a qué sede pertenecía el aula.
Ahora cada mensaje incluye el nombre legible de la sede resuelto desde el catálogo `Sedes`.

## 2. Cambios
- `ConversorPDF.XLL/Academic/Application/Services/Genetica/AgValidadorVentana.cs`
  - Nuevo helper privado `EtiquetaSede(AgContexto, AgAula)`:
    - Resuelve la sede por `contexto.SedesPorId` y, si no hay índice, por `contexto.Sedes`.
    - Prioridad de nombre: `Sede.Nombre` (`Ciudad (Subsede)`) → `Sede.Subunidad` → `Aula.Subunidad` → `Contexto.Subunidad`.
    - Fallback determinista: `Id {n}` cuando existe `IdSede`, o `N/D` sin datos. Nunca devuelve cadena vacía.
  - Se antepone la etiqueta a los mensajes del informe asociados a un aula:
    - `[I5_INFACTIBLE]`, `[I5_INFACTIBLE_OLEADA]`, `[R8_DiaFueraDePatron]`, `[I5_HOLGURA_BAJA]`,
      `[I5_HORAS_DESCONOCIDAS]`, `[D4_PREGRADO_8H]`, `[I1_VENTANA_VACIA]`.
    - Bloqueo "Asignatura … sesiones > capacidad real …" y bloqueo D14 "Asignatura … / Aula …".
- `.agents/plans/REGISTRY.md`: registro `R-032` (Ronda 310 / Fase 1579) y contadores.

### Formato resultante
```
[I5_INFACTIBLE] Sede 'CALI (SUBSEDE NORTE)' Aula '3|13|1|VI|A|MIXTA': demanda total de 64 sesiones ...
```

## 3. Cómo verificarlo
1. Ejecutar el generador de horarios (Módulo AG) con una selección que incluya aulas de más de una sede.
2. Forzar un aula con demanda mayor a la capacidad del semestre (por ejemplo, 6 asignaturas de 64 h en una
   ventana de 4 ranuras/semana).
3. En el log del informe, cada bloqueo/advertencia de aula debe mostrar `Sede '<nombre>'` entre el código y el `AulaRef`.
4. Si la sede no está en catálogo, se muestra `Sede 'Id 13'` o `Sede 'N/D'` sin romper el mensaje.

## 4. Cobertura de pruebas
- `ConversorPDF.Tests/Academic/Application/Services/Genetica/AgValidadorVentanaTests.cs`
  - Nuevo `ValidarFactibilidad_BloqueoI5_IncluyeNombreDeSede`: contexto con `Sede { Id = 13, Ciudad = "BOGOTA", Subsede = "SUBSEDE NORTE" }`
    y un aula que excede capacidad; verifica que el bloqueo contiene `[I5_INFACTIBLE]`, `Sede 'BOGOTA (SUBSEDE NORTE)'`
    y `Aula '1|13|1|I|A|PRESENCIAL'`.
  - Los 21 tests de la clase `AgValidadorVentanaTests` siguen pasando; no se modificaron aserciones existentes.
