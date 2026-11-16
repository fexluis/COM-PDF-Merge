---
description: Revisa el último plan de implementación y walkthrough de la conversación activa, detecta errores antes de ejecutar
---

# Workflow: `/ronda-review`

Revisa el `implementation_plan.md` y el `walkthrough.md` de la conversación activa.
Buscar la última ronda y últimas fases numeradas:
- Ronda {n}: {tittle} (Error/problem title)
-- Fase {n}: {tittle} (Error/problem title)

> Los dos artefactos están **escindidos por rangos de fase**. El archivo vivo contiene el índice y las últimas rondas; el historial congelado vive en `docs/plans/history/`. Resolver siempre la ubicación antes de leer (Paso 0).

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

## Paso 0 — Resolver la ubicación en el historial escindido

Regla: `.agents/rules/history-rotation.md`. Herramienta: `.agents/tools/split-history.ps1`.

1. La última `Ronda {n}` / `Fase {n}` se lee **del archivo vivo** (`implementation_plan.md` / `walkthrough.md`); su cola contiene las últimas 5 rondas.
2. Para una ronda o fase anterior al corte:
   - leer la tabla `| Archivo | Fases | Rondas | Lineas |` del archivo vivo;
   - abrir **solo** el trozo cuyo rango de fases contenga el número; si hay solape de rangos, desempatar con la columna `Rondas`;
   - no leer trozos completos ni concatenar historial.
3. Las citas `<archivo>.md:NNNN` anteriores al corte se resuelven con `docs/plans/history/LINE_MAP.md`; no reinterpretar el número contra el archivo vivo.
4. Los trozos son de solo lectura. Toda ronda nueva se escribe al final del archivo vivo, nunca en un trozo.

```
Ubicación resuelta:
- Ronda 402 → implementation_plan.md (vivo, cola)
- Fase 1336 → docs/plans/history/implementation_plan/implementation_plan_F1311_F1602.md
- Cita walkthrough.md:5604 → walkthrough_F0990_F1107.md:644
```

Si el índice no contiene el rango buscado, reportarlo como 🔴 CRÍTICO (historial incompleto) antes de continuar.

---

## Paso 1 — Localizar las fuentes de revisión

Leer los siguientes artefactos de la conversación activa:

- `implementation_plan.md` — índice + rondas posteriores al último corte (qué se propuso hacer y por qué)
- `walkthrough.md` — índice + rondas posteriores al último corte (qué se ejecutó realmente y qué archivos se modificaron)
- `docs/plans/history/<base>/*.md` — solo los trozos resueltos en el Paso 0

Si alguno de los archivos vivos no existe en la conversación activa, crear el respectivo archivo .md. No asumir. Mantener el hilo y la conversación. No recrear historial: si falta un trozo, detenerse y reportar.


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

Las referencias del plan/walkthrough a historial deben citarse con la ruta del trozo resuelto:

```
- ConversorPDF.XLL/.../AgPlanificadorRanuras.cs → cambio aplicado (Fase 1826)
- docs/plans/history/implementation_plan/implementation_plan_F1603_F1819.md:120 → justificación
- implementation_plan.md:45 → ronda vigente (cola del archivo vivo)
```

Si un archivo no existe en la ruta indicada, reportarlo como 🔴 CRÍTICO antes de continuar. Un trozo ausente en `docs/plans/history/` es pérdida de historial: detenerse.

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

Actualizar y extender la información de los archivos **vivos**:

- `implementation_plan.md`
- `walkthrough.md`

Detalla y consolida todos los cambios de la Ronda {n} aplicados al código fuente y reflejando los hallazgos y el soporte de realizado en cada fase {n} afectada. La ronda se anexa al final del archivo vivo; nunca se edita un trozo de `docs/plans/history/`.

Para cada fase {n}, responde explícitamente:

```
[ ] ¿Los comentarios están acordes con los cambios implementados? (doc refleja el comportamiento real, no el anterior)
[ ] ¿Se introdujo abstracción que el plan no solicitó? (clases/métodos/parámetros extra → revertir)
[ ] ¿Se modificó código adyacente fuera del alcance declarado? (→ revertir, reportar)
[ ] ¿Existe una solución más directa con menos líneas que resuelve el mismo problema? (si sí, proponerla como observación sin aplicarla; esperar aprobación)
```

Formato de respuesta obligatorio — tabla Markdown.

### 5.1 Escisión incremental (antes de cerrar)

Verificar umbrales de `.agents/rules/history-rotation.md` (>1500 líneas, >150 KB, >40 rondas desde el último corte, bloque `## Ronda` >1200 líneas):

```powershell
# 1) consulta de cortes: no escribe
powershell -ExecutionPolicy Bypass -File .agents/tools/split-history.ps1
# 2) materializar escisión incremental sobre el archivo ya indexado
powershell -ExecutionPolicy Bypass -File .agents/tools/split-history.ps1 -Apply
# por archivo, si solo uno supera el umbral
powershell -ExecutionPolicy Bypass -File .agents/tools/split-history.ps1 -Base implementation_plan -Apply
```

Comportamiento esperado en modo incremental:

- conserva las filas previas del índice y **añade** una fila por trozo nuevo;
- crea `<base>_F<min>_F<max>.md` con las rondas más antiguas de la cola;
- deja en el vivo la última ronda (o las últimas `-LiveRondas`);
- `LINE_MAP.md` conserva las citas previas y agrega las nuevas.

Verificar tras aplicar:

```
[ ] filas previas del índice intactas
[ ] trozo nuevo abre en encabezado `##`-`####` (sin corte dentro de fence/tabla)
[ ] `## Ronda` de la cola == rondas del trozo nuevo + rondas del vivo (sin duplicados ni pérdidas)
[ ] LINE_MAP.md con las citas previas conservadas
[ ] git status: trozo nuevo + archivo vivo + LINE_MAP.md
```

Si la herramienta reporta `sin contenido nuevo que escindir`, no forzar: el archivo vivo ya está por debajo del umbral.

Registrar el resultado en la fila vigente de `.agents/plans/REGISTRY.md` (§3, columna `verify`): `escisión <base> F<a>_F<b>` o `sin escisión (bajo umbral)`.

---

Al final responde:

```
Ronda {n} finalizada. Artefactos sincronizados:
- implementation_plan.md (índice + cola)
- walkthrough.md (índice + cola)
- escisión: <base>_F<a>_F<b>.md | sin escisión (bajo umbral)
```

Si no se encuentran problemas, declararlo explícitamente:

```
Sin hallazgos Ronda {n}. El plan y el walkthrough son consistentes con el código actual.
Escisión: sin cambios (bajo umbral) | <base>_F<a>_F<b>.md
```
---

## Paso 6 — Restricción
No compilar. No ejecutar el código.


---