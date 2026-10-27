# Workflow: Revisar

## Trigger
Use this workflow when the user asks to review code, detect errors, or validate a recent implementation.

## Objective
Detect errors and potential runtime failures in the requested C# code, explain the root cause of each issue, and propose concrete corrections with code snippets.

## Steps
1. **Identify the scope**
Revisa el actual `implementation_plan.md`, `walkthrough.md` y `task.md` y detecta cualquier error o problema que pueda generar fallos. Task Completed: En caso que la Revisión esté completada y no se encuentren errores se debe responder "Sin errores adicionales detectados".

2. **Read the relevant files**
   - La ruta completa del `implementation_plan.md` se proporciona en el chat.
   - Deduce las rutas del `walkthrough.md` y `task.md` reemplazando el nombre del archivo: están en el mismo directorio.
     - Ejemplo base: `C:\Users\USUARIO\.gemini\antigravity-ide\brain\02b22785-...\implementation_plan.md`
     - → `C:\Users\USUARIO\.gemini\antigravity-ide\brain\02b22785-...\walkthrough.md`
     - → `C:\Users\USUARIO\.gemini\antigravity-ide\brain\02b22785-...\task.md`
   - Usa `list_files` sobre el directorio contenedor para confirmar los archivos presentes.
   - Usa `read_file` sobre `implementation_plan.md` para obtener el plan de implementación (fases, objetivos, estado actual).
   - Usa `read_file` sobre `walkthrough.md` para obtener el registro de ejecución (lo implementado, problemas encontrados, decisiones tomadas).
   - Usa `read_file` sobre `task.md` para obtener la lista de tareas definidas y su estado (pendiente, en progreso, completada).
   - Revisa cada tarea en `task.md` contra el código fuente para verificar que esté correctamente implementada y sin errores.
   - Usa `search_files` para localizar referencias en el código fuente, nombres de variables antiguos o rutas de código relacionadas.

3. **Generate Output Format & update implementation_plan.md**

   - Aplica el siguiente Output Format para documentar los hallazgos de cada error detectado:

   Responde con la mínima cantidad de tokens:
   - Sin relleno, sin cortesías
   - Sin lenguaje emocional
   - Frases cortas y directas
   - Prefiere fragmentos en vez de oraciones completas
   - Quita palabras innecesarias (artículos, relleno, repetición)
   - Enfócate solo en resolver el problema
   - Entrega solo información útil

## Fase {n}: {título} (Error/problem title)

**Ubicación:** {archivo}, líneas {n}-{m} (File and line reference)

**Problema:** {descripción} 

**Causa:** {raíz} (Cause explanation)

**Corrección:**
{código} (Corrected code snippet)

---
**Severity: Compilation / Runtime / Logic / Warning**

Al final, tabla resumen:

| Fase | Severidad | Archivo | Líneas |
|------|-----------|---------|--------|

Y conclusión breve.

Restricción: No compiles ni ejecutes el código.



