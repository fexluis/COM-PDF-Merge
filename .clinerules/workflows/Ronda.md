# Workflow: Ronda

## Trigger
Use this workflow when the user asks to review code, detect errors, or validate a recent implementation.

## Objective
Detect errors and potential runtime failures in the requested C#/SQL code, explain the root cause of each issue, and propose concrete corrections with code snippets.

## Steps
1. **Identify the scope**
Buscar la última ronda y última fase numerada.
Ronda #: Enumera la ronda para iniciar la revisión de las fases de la ronda anterior, continúa con la última numeración:

## Ronda {n}: {tittle} (Error/problem title)

Contexto de la ronda anterior:
-Fases aplicadas y verificadas: {lista}
-Fases retiradas o falso positivo: {lista, si aplica}
-Estado de corrección pendiente: {fases incompletas / fases abiertas}

2. **Detect runtime/logic errors and Propose concrete corrections**
Revisa exhaustivamente el código de las fases {n-?}...{n} aplicadas en la ronda anterior. Confirma que cada corrección se aplicó como fue especificada (líneas, lógica, validaciones) y no introdujo regresiones.
Si detectas errores nuevos o incompletos (compilación, runtime, lógica, UX/warning), enuméralos siguiendo el formato de fase, continuando la numeración exacta desde la última fase conocida.

Detecta cualquier error o problema que pueda generar fallos en su ejecución
   - For each error, explain the cause and provide the corrected C# or SQL snippet.
   - Keep changes minimal and non-destructive.

## Fase {n}: {tittle} (Error/problem title)

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


4. **Save AI tokens and reduce costs**
Responde con la mínima cantidad de tokens:
- Sin relleno, sin cortesías
- Sin lenguaje emocional
- Frases cortas y directas
- Prefiere fragmentos en vez de oraciones completas
- Quita palabras innecesarias (artículos, relleno, repetición)
- Enfócate solo en resolver el problema
- Entrega solo información útil

5. **Verification Plan**
Verifica que las fases no introducen regresiones.
Conclusión con decisión final Ronda {n}:
a. Si se requieren correcciones / Nuevas Fases
b. Sin Fases nuevas obligatorias / No se requieren más correcciones / Fase opcional / Recomendar pasar a build release + smoke test.
 
6. **Restricciones**
- Requiere modo Act para ejecutar build/tests/smoke.
- No compilar, No editar archivos. Esperar aprobación. 


