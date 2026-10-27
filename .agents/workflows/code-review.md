---
description: Revisa el último plan de implementación y walkthrough de la conversación activa, detecta errores antes de ejecutar
---

# Workflow: `/code-review`

Revisa el `implementation_plan.md` y el `walkthrough.md` de la conversación activa.
Detecta errores. Reporta. Espera aprobación antes de corregir.

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

Si alguno no existe en la conversación activa, declararlo al usuario y preguntar cómo proceder. No asumir.

Construir una lista explícita:
```
Archivos involucrados según el plan/walkthrough:
- [ruta] → [cambio propuesto/aplicado]
- ...
```

---

## Paso 2 — Leer cada archivo fuente involucrado

Para cada archivo listado en el Paso 1, leer el contenido **actual en disco**.
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

Para cada cambio verificar:

- ¿El tipo C# receptor coincide con lo que Dapper/SQLite retorna? (`bool` vs `int`, `decimal` vs `double`, etc.)
- ¿Los nombres de columnas SQL coinciden con las propiedades del modelo destino? (considerar `MatchNamesWithUnderscores`)
- ¿Hay referencias a columnas, tablas o propiedades que ya no existen (renombradas o eliminadas en migraciones previas)?
- ¿El orden de operaciones importa? (ej: asignar a campo estático antes de completar todas las lecturas del `GridReader`)
- ¿Hay efectos secundarios si ocurre una excepción parcial? (estado inconsistente en caché, transacción no revertida, etc.)
- ¿El plan describe un cambio en un archivo pero el walkthrough indica que se modificó otro? (desincronía plan/ejecución)

### 4.2 Simplicidad

Para cada cambio verificar:

- ¿Se introdujo abstracción que el plan no solicitó?
- ¿Se modificó código adyacente fuera del alcance declarado?
- ¿Existe una solución más directa con menos líneas que resuelve el mismo problema?

### 4.3 Alcance quirúrgico

Para cada archivo modificado verificar:

- ¿Cada línea cambiada se rastrea directamente al problema descrito en el plan?
- ¿Se refactorizó algo que no estaba roto?
- ¿Se modificó estilo, nomenclatura o comentarios fuera del alcance?

---

## Paso 5 — Reportar hallazgos

Formato de respuesta obligatorio — tabla Markdown:

| # | Archivo | Línea | Problema | Severidad |
|---|---------|-------|----------|-----------|
| 1 | `ruta/archivo.cs` | 48 | descripción concreta | 🔴 CRÍTICO |

**Leyenda de severidades:**

| Ícono | Nivel | Criterio |
|-------|-------|----------|
| 🔴 | **CRÍTICO** | Fallo en runtime garantizado (type mismatch, columna inexistente, NullRef, excepción no manejada) |
| 🟡 | **RIESGO** | Fallo bajo condiciones específicas (estado parcial en excepción, race condition, datos corruptos silenciosos) |
| 🔵 | **ALCANCE** | Cambio fuera del problema declarado (refactor no solicitado, código adyacente modificado) |
| ⚪ | **INFO** | Observación sin acción requerida — mencionar, no tocar |

Si no se encuentran problemas, declararlo explícitamente:
```
Sin hallazgos. El plan y el walkthrough son consistentes con el código actual.
```

---

## Paso 6 — Proponer correcciones (solo si hay 🔴 o 🟡)

Para cada hallazgo 🔴 o 🟡:

- Mostrar **solo el bloque exacto a reemplazar**, no el archivo completo.
- No introducir cambios adicionales dentro del bloque corregido.
- Si la corrección requiere modificar más archivos de los que el plan original tocaba, declararlo y pedir aprobación antes de continuar.

**No aplicar ninguna corrección hasta que el usuario apruebe el reporte.**

---

## Paso 7 — Esperar aprobación

Terminar la respuesta con:

```
¿Apruebas las correcciones propuestas? Responde "sí" para aplicar, o indica ajustes.
```

Si el usuario aprueba, aplicar **solo** los cambios del reporte aprobado.
Si el usuario pide ajustes, volver al Paso 5 con la nueva información.

---

## Notas de uso

- Si el usuario invoca `/code-review` sin `implementation_plan.md` en la conversación activa, preguntar qué revisar. No asumir.
- Override opcional: `/code-review @[ruta-archivo]` usa el archivo indicado en lugar del artefacto de la conversación.
- Código muerto no relacionado encontrado durante la revisión → reportar como ⚪ INFO. No eliminar.