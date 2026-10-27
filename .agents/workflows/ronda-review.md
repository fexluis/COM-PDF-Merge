---
description: Revisa el último plan de implementación y walkthrough de la conversación activa, detecta errores antes de ejecutar
---

# Workflow: `/ronda-review`

Revisa el `implementation_plan.md` y el `walkthrough.md` de la conversación activa.
Buscar la última ronda y últimas fases numeradas:
- Ronda {n}: {tittle} (Error/problem title)
-- Fase {n}: {tittle} (Error/problem title)

Responde con la mínima cantidad de tokens:

- Sin relleno, sin cortesías
- Sin lenguaje emocional
- Frases cortas y directas
- Prefiere fragmentos en vez de oraciones completas
- Quita palabras innecesarias (artículos, relleno, repetición)
- Enfócate solo en resolver el problema
- Entrega solo información útil

---

## Principios operativos

1. **Pensar antes de codificar** — declarar ambigüedad explícitamente, no asumir.
2. **Simplicidad primero** — no abstraer lo que no se pidió, no "mejorar" lo adyacente.
3. **Cambios quirúrgicos** — cada línea cambiada se rastrea al problema declarado.

---

## Paso 1 — Localizar las fuentes de revisión

Leer los siguientes artefactos de la conversación activa:

- `implementation_plan.md` — qué se propuso hacer y por qué
- `walkthrough.md` — qué se ejecutó realmente y qué archivos se modificaron

Si alguno no existe en la conversación activa, crear el respectivo archivo .md. No asumir. Mantener el hilo y la conversación.

---

## Paso 2 — Leer cada archivo fuente involucrado

Construir una lista explícita:
```
Archivos involucrados según el plan/walkthrough:
- [ruta] → [cambio propuesto/aplicado]
- ...
```

Para cada archivo en la lista, leer el contenido **actual en disco**.
No trabajar desde memoria ni desde el plan. El archivo real es la fuente de verdad.

Si un archivo no existe en la ruta indicada, reportarlo como 🔴 CRÍTICO antes de continuar.

---

## Paso 3 — Declarar ambigüedades

Si hay más de una interpretación posible para un cambio propuesto:
- Listar las interpretaciones numeradas.
- Preguntar al usuario antes de continuar.
- No elegir en silencio.

Si no hay ambigüedades, declararlo explícitamente y continuar.

---

## Paso 4 — Revisión por principios (orden fijo)

### 4.1 Correctitud

- Evalúa críticamente la nueva Ronda {n}: {tittle} (Error/problem title) y las nuevas fases {n}: {tittle} (Error/problem title) que se van aplicar y por cada una de ellas confirma: 
a. ACEPTADA: Si estás de acuerdo con los análisis y se implementó la fase.
b. RETIRADA: Si detectas un Falso Positivo en alguna de las fases y la rechazas explicando las razones.

Validados los hallazgo, respondes:

# Ronda {n}: {tittle} (Error/problem title) 
---  
# Evaluación Crítica Fase {n}: {tittle} (Error/problem title):  
1. Estado: ACEPTADA / RETIRADA  
2. Hallazgo: Diagnóstico Crítico & Causa Raíz & Impacto Confirmado  
3. Acción Aplicada: Código / Solución implementada  
4. Archivo Afectado: Nombre archivo afectado  
---

Para cada cambio verificar:

- ¿Hay referencias a columnas, tablas o propiedades que ya no existen (renombradas o eliminadas en migraciones previas)?
- ¿El orden de operaciones importa? (ej: asignar a campo estático antes de completar todas las lecturas del `GridReader`)
- ¿Hay efectos secundarios si ocurre una excepción parcial? (estado inconsistente en caché, transacción no revertida, etc.)
- ¿El plan describe un cambio en un archivo pero el walkthrough indica que se modificó otro? (desincronía plan/ejecución)

### 4.2 Simplicidad

Aplica los cambios al código siguiendo las instrucciones de cada fase {n} ACEPTADA.
Actualiza en cada archivo toda la documentación interna y comentarios del código para que reflejen las modificaciones realizadas (incluye: descripción de funciones, parámetros, lógica alterada y razón del cambio).

Comentarios obligatorios: en cada archivo modificado, actualiza toda la documentación interna para reflejar el cambio:
- Encabezados/descripción de funciones alteradas.
- Parámetros nuevos o cuyo significado cambió.
- Lógica: Qué hace ahora, por qué.
- Elimina comentarios obsoletos que contradigan el nuevo comportamiento.

No destructivo: no renombres, no reorganices, no "mejores" código adyacente que la fase no tocó.

---

## Paso 5 — Reportar hallazgos Ronda {n}

Actualizar y extender la información de los archivos:

- `implementation_plan.md`
- `walkthrough.md`

Detalla y consolida todos los cambios de la Ronda {n} aplicados al código fuente y reflejando los hallazgos y el soporte de realizado en cada fase {n} afectada.

Para cada fase {n}, responde explícitamente:

[ ] ¿Los comentarios están acordes con los cambios implementados? (doc refleja el comportamiento real, no el anterior)
[ ] ¿Se introdujo abstracción que el plan no solicitó? (clases/métodos/parámetros extra → revertir)
[ ] ¿Se modificó código adyacente fuera del alcance declarado? (→ revertir, reportar)
[ ] ¿Existe una solución más directa con menos líneas que resuelve el mismo problema? (si sí, proponerla como observación sin aplicarla; esperar aprobación)Formato de respuesta obligatorio — tabla Markdown.

Al final responde:

```
Ronda {n} finalizada. Artefactos sincronizados: 
- implementation_plan.md 
- walkthrough.md
```

Si no se encuentran problemas, declararlo explícitamente:

```
Sin hallazgos Ronda {n}. El plan y el walkthrough son consistentes con el código actual.
```
---

## Paso 6 — Restricción
No compilar. No ejecutar el código.


---