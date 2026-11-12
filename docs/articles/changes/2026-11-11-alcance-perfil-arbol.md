# Alcance Completo al Cargar un Perfil de Árbol (Fase 1580)

## 1. ¿Por qué se omitían combinaciones si las sedes están activas?
El aviso *«Se omitieron N combinación(es) no encontradas en el período/sedes activas»* no consultaba
el campo `sedes.activo`. Provenía de `AgFormGeneradorHorarios.AsegurarSubarbol`, que recorre cada
combinación guardada (`ag_arbol_nodo`) y la cuenta como huérfana cuando **no encuentra el nodo** de
período, sede o programa dentro del árbol visible:

```csharp
TreeNode nodoSede = nodoPer.Nodes.Cast<TreeNode>()
    .FirstOrDefault(n => n.Tag is AgNodoArbolKey k && k.Nivel == 2 && k.Id == dto.IdSede);

if (nodoSede == null) { huerfanos++; continue; }
```

El árbol no contiene todas las sedes activas: `PoblarArbolContextoAcademico` lo construye con el
**filtro vigente** de los combos (`FiltrarPeriodos` / `FiltrarSedes` / `FiltrarProgramas` sobre la
selección de `cmbPeriodos`, `cmbSedes`, `cmbProgramas` y el `cmbTipoPrograma`). Además, la carga del
perfil solo sincronizaba el **primer DTO**:

```csharp
// Fase 1208 (comportamiento anterior)
var primerDto = dtos.FirstOrDefault();
cmbSedes.Marcar(new[] { sedeEncontrada.Id });
cmbProgramas.Marcar(new[] { progEncontrado.Id });
```

Resultado: un perfil de 35 combinaciones cargado con el filtro en 1 sede + 1 programa restauraba solo
las 12 combinaciones presentes en el árbol (35 − 23) y el resto se reportaba como «omitido», aunque sus
sedes y programas estuvieran activos.

## 2. Cambios
### Nuevo resolutor de alcance (Application)
`ConversorPDF.XLL/Academic/Application/Services/AgPerfilAlcanceResolver.cs`
- `AgPerfilAlcanceResolver.Resolver(dtos, periodos, sedes, programas)` devuelve:
  - `Periodos`, `Sedes`, `Programas`: ids distintos del perfil **presentes en catálogo**.
  - `TipoPrograma`: `PREGRADO`, `POSGRADO` o `TODOS` (mixto) para que ningún programa quede oculto por el filtro de tipo.
  - `CombosFueraDeCatalogo`: combinaciones cuya tupla período/sede/programa no existe (sede inactiva, programa eliminado, etc.).

### Carga del perfil (UI)
`ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- `CargarPerfilArbolSeleccionadoAsync`: sustituye la sincronización del primer DTO (Fase 1208) por el alcance completo del perfil.
  1. Marca todos los períodos y sedes del perfil presentes en catálogo.
  2. Ajusta `cmbTipoPrograma` al tipo mínimo que cubre esos programas.
  3. `CargarProgramasSegunTipo(tipo, programasDelPerfil)` recarga y marca los programas del perfil.
  4. `PoblarArbolContextoAcademico()` reconstruye el árbol con ese alcance **antes** de restaurar las hojas,
     limpia `_filtroArbolDesactualizado` y persiste los filtros (`GuardarFiltrosAsync(forzar: true)`).
- `AsegurarSubarbol(..., out List<string> omitidas)`: registra el `AulaRef` de cada combinación omitida.
- Mensajes corregidos: ahora dicen *«su período, sede o programa no existe en el catálogo activo»*,
  incluyen hasta 5 `AulaRef` de ejemplo y el listado completo va al log.

## 3. Cómo verificarlo
1. Guardar un perfil con combinaciones de varias sedes/programas (p. ej. 35).
2. Reducir el filtro a 1 sede + 1 programa y pulsar **Cargar**.
3. El árbol se repuebla con las sedes, programas y tipo necesarios y se marcan todas las combinaciones restaurables.
4. El diálogo solo aparece si alguna combinación apunta a una sede/programa/período inexistente o inactivo,
   mostrando su `AulaRef` para diagnóstico.

## 4. Cobertura de pruebas
`ConversorPDF.Tests/Academic/Application/Services/AgPerfilAlcanceResolverTests.cs` (6 tests):
- Alcance con varias sedes activas → todas incluidas.
- Sede/programa inexistente → contabilizada en `CombosFueraDeCatalogo` y excluida del alcance.
- Programas mixtos → `TODOS`; solo posgrado → `POSGRADO`; solo pregrado → `PREGRADO`.
- Perfil vacío → alcance vacío sin excepciones.
