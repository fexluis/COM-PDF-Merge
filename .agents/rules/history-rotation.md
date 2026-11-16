# Regla: rotación del historial (`implementation_plan.md` / `walkthrough.md`)

`Ronda` y `Fase` son globales y su fuente única es `.agents/plans/REGISTRY.md`. Esta regla regula **dónde** vive ese historial, no su numeración.

## Umbrales de escisión (se dispara con cualquiera)

- Archivo vivo > **1500 líneas** o > **150 KB**.
- Más de **40 rondas** desde la última escisión registrada en el índice.
- Algún bloque `## Ronda` supera **1200 líneas**.

## Procedimiento (solo en modo Act)

1. `powershell -ExecutionPolicy Bypass -File .agents/tools/split-history.ps1 -DryRun`
   Revisar la tabla de cortes. Si algún trozo supera `Target * 1.3`, repetir con `-Target` menor.
2. `powershell -ExecutionPolicy Bypass -File .agents/tools/split-history.ps1 -Apply`
   Escribe `docs/plans/history/<base>/<base>_F<min>_F<max>.md`, reescribe el archivo vivo con el índice y regenera `docs/plans/history/LINE_MAP.md`.
3. Validaciones obligatorias (el script aborta si falla la primera):
   - concatenación trozos + vivo == original (comparación exacta);
   - cada `## Ronda` aparece exactamente una vez entre trozos + vivo;
   - toda fila del índice apunta a un archivo existente;
   - `git status` solo muestra trozos nuevos + los dos archivos vivos reescritos + `LINE_MAP.md`.
4. Un único commit: `chore(history): escisión <base> F<min>_F<max> (Ronda <r>)`.
5. Registrar en `.agents/plans/REGISTRY.md` §3, columna `verify` de la fila vigente: `escisión histórica F<a>_F<b>`.

## Invariantes

- **Append-only**: el contenido congelado no se edita, reordena ni borra; lo nuevo se añade al archivo vivo.
- La numeración `Ronda`/`Fase` de `REGISTRY.md` no se toca ni se recicla.
- No se corta dentro de un bloque de código (```` ``` ````/`~~~`), tabla o sección: todo trozo empieza en encabezado `##`-`####`.
- Preservar codificación (UTF-8, BOM si existía) y fin de línea original; el script lo detecta por archivo.
- Las citas históricas `<archivo>.md:NNNN` no se reescriben: se resuelven con `docs/plans/history/LINE_MAP.md`.
- Los trozos son material de consulta: nunca son destino de escritura de una ronda nueva.

## Prohibiciones

- Ejecutar la escisión a mitad de una ronda con fases abiertas sin cerrar el registro.
- Mover contenido entre trozos ya congelados para «cuadrar» rangos de fase.
- Publicar un trozo sin su fila en el índice del archivo vivo.
