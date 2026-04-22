---
description: Para uso en sesiones de análisis, refactoring y resolución de errores
---

# Workflow: Sistema de Documentación de Contexto IA
> Versión 1.0 — Para uso en sesiones de análisis, refactoring y resolución de errores

---

## Visión General

```
INICIO DE SESIÓN
      │
      ▼
┌─────────────────┐
│  Encabezado +   │
│  Plan de ref.   │
└────────┬────────┘
         │
         ▼
┌─────────────────┐        ┌──────────────────────┐
│  Ejecutar tarea │──────▶ │  ¿Breakpoint limpio? │
│  del plan       │        └──────────┬───────────┘
└─────────────────┘                   │
         ▲               NO           │ SÍ
         │◀────────────────┘          ▼
         │                  ┌─────────────────────┐
         │                  │  Documentar entrada  │
         │                  │  en analisis_ia.md   │
         │                  └──────────┬──────────┘
         │                             │
         │◀────────────────────────────┘
         │
         │  ¿Tokens < 800?
         ├──────────────────▶ BP-TOKEN → Interrupción controlada
         │
         │  ¿Tarea completa?
         └──────────────────▶ BP-FINAL → Resumen + Recomendaciones
```

---

## Fase 1: Inicialización de Sesión

### Paso 1.1 — Crear encabezado en `analisis_ia.md`

Al iniciar cualquier sesión de trabajo, el archivo debe abrirse o crearse con esta estructura:

```markdown
# Análisis IA — Sesión YYYY-MM-DD

## Encabezado de Sesión
- Inicio: YYYY-MM-DD HH:MM
- Tarea principal: [descripción de una línea]
- Plan de referencia: [nombre del archivo o plan base, ej: plan_estabilizacion_xll_v3.md]
- Estado global: EN PROCESO

## Índice de Breakpoints
| # | Tipo | Descripción breve | Estado |
|---|------|-------------------|--------|
| — | —    | (se completa durante la sesión) | — |

## Entradas Detalladas
(se completa durante la sesión)

## Resumen Final
(se completa al cerrar la sesión)
```

### Paso 1.2 — Regla de inicio obligatoria

Antes de ejecutar cualquier acción, la IA debe declarar en su primera respuesta:

```
📋 Sesión iniciada.
Plan de referencia: [archivo]
Primer breakpoint esperado: [BP-TIPO] al completar [tarea inicial]
```

---

## Fase 2: Definición de Breakpoints

Un **breakpoint** es el único momento en que se actualiza `analisis_ia.md`.
No se documenta cada micro-paso de razonamiento interno — solo decisiones,
hallazgos y cambios de estado significativos.

### Tabla de Breakpoints

| Código | Nombre | Cuándo se activa |
|--------|--------|-----------------|
| **BP-FASE** | Fase completada | Al terminar una fase completa del plan |
| **BP-DECISIÓN** | Decisión tomada | Cuando se evalúan 2+ alternativas y se elige una |
| **BP-ERROR** | Error encontrado | Al detectar un error inesperado, antes de corregirlo |
| **BP-BLOQUEO** | Esperando input | Cuando se necesita respuesta del usuario para continuar |
| **BP-TOKEN** | Límite de contexto | Si se estiman menos de 800 tokens disponibles |
| **BP-FINAL** | Sesión cerrada | Al completar la tarea o cerrar la sesión voluntariamente |

### Regla de activación

```
¿Terminé una fase completa?         → BP-FASE
¿Elegí entre 2+ opciones?           → BP-DECISIÓN
¿Encontré algo inesperado?          → BP-ERROR
¿Necesito input del usuario?        → BP-BLOQUEO
¿Tokens restantes < 800?            → BP-TOKEN  ← PRIORIDAD MÁXIMA
¿Terminé todo?                      → BP-FINAL
```

---

## Fase 3: Formato de Entrada en Breakpoints

Cada vez que se activa un breakpoint, se agrega esta entrada al archivo:

```markdown
---
## [BP-TIPO] Paso #N — YYYY-MM-DD HH:MM

**Tipo:** Fase | Decisión | Error | Bloqueo | Interrupción | Final
**Estado:** COMPLETADO | FALLIDO | EN ESPERA DE INPUT | INTERRUMPIDO

### Contexto
[Qué tarea se estaba ejecutando y en qué punto del plan general se encuentra]

### Hallazgos
[Datos concretos: archivos encontrados, errores identificados, resultados
de comandos, namespaces localizados, valores exactos obtenidos]

### Decisión tomada
[Qué se eligió hacer y por qué. Si se descartó una alternativa, explicar
el motivo en una línea]

### Estado del sistema al cierre del paso
- Errores resueltos: N de Total
- Archivos modificados: [lista o "ninguno"]
- Variables clave: [valores de estado que importan para continuar]
- Próxima acción: [descripción exacta del siguiente paso]

### Punto de reanudación
> "Reanudar desde analisis_ia.md Paso #N.
>  Ejecutar: [comando o acción literal para retomar]"
```

---

## Fase 4: Protocolo de Interrupción Controlada (BP-TOKEN)

Este protocolo tiene **prioridad sobre cualquier otra acción** cuando el
contexto disponible es insuficiente para completar la siguiente fase.

### Activación

```
SI tokens_estimados < 800 O contexto_cercano_al_límite:
    1. Detener en el próximo punto limpio (no en medio de una operación)
    2. Escribir entrada BP-TOKEN en analisis_ia.md
    3. Enviar mensaje de interrupción al usuario
    4. NO continuar con la siguiente tarea
```

### Entrada BP-TOKEN en el archivo

```markdown
---
## [BP-TOKEN] Paso #N — YYYY-MM-DD HH:MM

**Tipo:** Interrupción
**Estado:** INTERRUMPIDO — Límite de contexto detectado

### Contexto
[Tarea en ejecución al momento de la interrupción]

### Estado del sistema al momento de interrupción
- Última acción completada: [descripción]
- Acción en curso (no completada): [descripción o "ninguna"]
- Errores resueltos hasta aquí: N de Total
- Archivos modificados en esta sesión: [lista]

### Punto de reanudación
> "Reanudar desde analisis_ia.md Paso #N.
>  Próxima acción: [instrucción literal]"
```

### Mensaje al usuario (texto exacto)

```
⚠️ INTERRUPCIÓN CONTROLADA

El contexto disponible es insuficiente para completar la siguiente
fase sin pérdida de información.

Último estado guardado: [descripción de 1 línea]
Archivo actualizado: analisis_ia.md — Paso #N
Errores resueltos: N de Total

Para reanudar, inicia una nueva sesión y escribe:
"Reanudar desde analisis_ia.md — Paso #N. [Próxima acción]"
```

---

## Fase 5: Cierre de Sesión (BP-FINAL)

Al completar la tarea o al cerrar la sesión voluntariamente:

### Entrada BP-FINAL en el archivo

```markdown
---
## [BP-FINAL] Paso #N — YYYY-MM-DD HH:MM

**Tipo:** Final
**Estado:** COMPLETADO | PARCIAL | FALLIDO

### Resumen de la sesión
- Tarea principal: [descripción]
- Resultado: COMPLETADO / PARCIAL / FALLIDO
- Duración estimada: [HH:MM a HH:MM]

### Métricas
- Breakpoints documentados: N
- Errores encontrados: N
- Errores resueltos: N
- Archivos modificados: [lista]
- Decisiones tomadas: N

### Pendientes para próxima sesión
1. [tarea pendiente 1]
2. [tarea pendiente 2]

### Recomendaciones
[Observaciones técnicas o de proceso para el siguiente ciclo de trabajo]

### Índice final actualizado
| # | Tipo | Descripción | Estado |
|---|------|-------------|--------|
| 1 | BP-FASE | ... | COMPLETADO |
| 2 | BP-ERROR | ... | COMPLETADO |
| N | BP-FINAL | Cierre de sesión | COMPLETADO |
```

### Confirmación al usuario

Al finalizar cada respuesta que incluya un breakpoint documentado:

```
analisis_ia.md actualizado — Paso #N [BP-TIPO]: [descripción de 1 línea].
Próximo breakpoint esperado: [BP-TIPO] al completar [tarea].
```

---

## Fase 6: Protocolo de Reanudación

Cuando el usuario inicia una nueva sesión tras una interrupción:

### Instrucción del usuario
```
"Reanudar desde analisis_ia.md — Paso #N. [Próxima acción]"
```

### Acciones de la IA al reanudar

```
1. Leer analisis_ia.md desde el Paso #N indicado
2. Reconstruir el estado del sistema desde los hallazgos documentados
3. Confirmar el estado reconstruido al usuario antes de ejecutar
4. Continuar desde el "Punto de reanudación" del último breakpoint
```

### Confirmación de reanudación (texto al usuario)

```
🔄 REANUDACIÓN DESDE PUNTO #N

Estado reconstruido:
- Tarea principal: [descripción]
- Último paso completado: [BP-TIPO #N — descripción]
- Errores resueltos: N de Total
- Próxima action: [instrucción del punto de reanudación]

¿Confirmas que este estado es correcto? (responde "sí" para continuar
o indica correcciones)
```

---

## Referencia Rápida

```
CUÁNDO DOCUMENTAR        → Solo en breakpoints (BP-FASE, BP-DECISIÓN,
                            BP-ERROR, BP-BLOQUEO, BP-TOKEN, BP-FINAL)

CUÁNDO NO DOCUMENTAR     → Durante razonamiento interno, búsquedas
                            intermedias, micro-pasos sin decisión

PRIORIDAD MÁXIMA         → BP-TOKEN siempre tiene precedencia sobre
                            completar la tarea en curso

REGLA DE ORO             → Un breakpoint = una entrada completa en el
                            archivo. Nunca entrada parcial.

PUNTO DE REANUDACIÓN     → Cada entrada DEBE terminar with instrucción
                            literal para retomar, sin ambigüedad
```

---

*Workflow v1.0 — Sistema de Documentación de Contexto IA*
