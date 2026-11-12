# Escalado Automático de Ventanas: LUN-VIE, JUE-DOM y MIE-DOM (Fases 1707–1714)

## 1. Contexto y Objetivos

Durante la optimización y validación pre-vuelo de horarios en el Módulo R (`AgFormGeneradorHorarios`), se identificaron escenarios donde la demanda semanal o acumulada de sesiones de un aula excede la capacidad nominal de su ventana de días configurada. En las versiones previas, la escalera de ventanas estaba restringida a saltos básicos de fin de semana (`VIE-SAB → JUE-SAB`), dejando desatendidas las aulas con ventanas semanales (`LUN-JUE`, `MAR-VIE`) o intensivas (`VIE-DOM`).

El objetivo de esta intervención (Ronda 345–350, Fases 1707–1714, Plan `R-058`) fue extender la escalera canónica de ventanas para incorporar de manera determinista y predecible:
1. `LUN-VIE` como escalón natural para ventanas semanales (`LUN-JUE` y `MAR-VIE`).
2. `JUE-DOM` y `MIE-DOM` como escalones progresivos para ventanas de fin de semana extendido (`VIE-DOM`).
3. Parser único no envolvente de rangos arbitrarios en `AgDiasVentana.cs`.
4. Rotulación dinámica según unidad horaria (4 h u 8 h) distinguiendo capacidad nominal y efectiva.
5. Sincronización transparente de dimensiones en UI (`AgFormGeneradorHorarios.cs`) evitando clics espurios y recargando el árbol sin rebarajar combos.

---

## 2. Capacidad Nominal vs. Capacidad Efectiva

### 2.1. Capacidad Nominal por Semana (4 h / 8 h)

La capacidad nominal representa el número máximo teórico de ranuras disponibles por semana según la ventana de días y la duración de la sesión (4 horas para pregrado, 8 horas para posgrado):

| Ventana | Días ISO | Ranuras Semanales (4 h) | Ranuras Semanales (8 h) | Descripción de Días |
|---|---|---|---|---|
| `LUN-JUE` | `1,2,3,4` | 4 | 4 | Lunes a Jueves |
| `MAR-VIE` | `2,3,4,5` | 4 | 4 | Martes a Viernes |
| `LUN-VIE` | `1,2,3,4,5` | 5 | 5 | Lunes a Viernes |
| `VIE-SAB` | `5,6` | 3 (V=1, S=2) | 2 (V=1, S=1) | Viernes tarde + Sábado mañana/tarde |
| `JUE-SAB` | `4,5,6` | 4 (J=1, V=1, S=2) | 3 (J=1, V=1, S=1) | Jueves tarde + Viernes tarde + Sábado |
| `VIE-DOM` | `5,6,7` | 4 (V=1, S=2, D=1) | 3 (V=1, S=1, D=1) | Fin de semana extendido con Domingo |
| `JUE-DOM` | `4,5,6,7` | 5 (J=1, V=1, S=2, D=1) | 4 (J=1, V=1, S=1, D=1) | Jueves a Domingo |
| `MIE-DOM` | `3,4,5,6,7` | 6 (M=1, J=1, V=1, S=2, D=1) | 5 (M=1, J=1, V=1, S=1, D=1) | Miércoles a Domingo |

### 2.2. Capacidad Efectiva del Período ($W$ Semanas)

Para un período lectivo con $W$ semanas completas:
$$\text{Capacidad Efectiva} = \text{Capacidad Nominal Semanal} \times W - \text{Festivos Afectados} - \text{Ranuras sin Bloque en Catálogo}$$

En los fixtures de prueba y la evaluación de escalado (`AgEscaladorVentanaAula`):
- Una ventana `LUN-JUE` o `VIE-DOM` de $W$ semanas tiene capacidad nominal de $4W$ ranuras.
- Al escalar a `LUN-VIE` o `JUE-DOM`, la capacidad nominal pasa a $5W$ ranuras.
- Al escalar de `JUE-DOM` a `MIE-DOM`, pasa de $5W$ a $6W$ ranuras.

---

## 3. Rutas de Escalera y Reglas de Negocio

### 3.1. Grafo de Escalones

```text
[Semanal]
LUN-JUE ─────────┐
                 ├───> LUN-VIE (Terminal semanal, sin escalón siguiente)
MAR-VIE ─────────┘

[Fin de Semana Tradicional]
VIE-SAB ───> JUE-SAB ───> JUE-DOM (Alias ISO 4,5,6,7) ───> MIE-DOM (Terminal)

[Fin de Semana Extendido]
VIE-DOM ─────────────────> JUE-DOM ───> MIE-DOM (Terminal intensivo)
```

### 3.2. Reglas Cardinales Implementadas

- **R1 — Sábado como Rótulo y Tokens Rechazados:**
  Los sufijos de franja `6M` y `6T` son tokens exclusivos de salida y rotulación (`AgEscaleraVentanas.Rotular`). El parser de entrada (`AgDiasVentana.Dias()`, `AgDiasVentana.AIso()`) los rechaza estrictamente devolviendo 0 días o lista vacía, garantizando que el sábado no se compute doblemente por error.
- **R2 — Normalización y Claves de Escalones:**
  `AgEscaleraVentanas.Normalizar()` valida primero la coincidencia exacta de catálogo. Si no coincide, recurre a `AgDiasVentana.AIso()`. Para preservar compatibilidad histórica con la BD, `4,5,6,7` se normaliza canónicamente a `JUE-DOM`.
- **R3 — Aislamiento Arquitectural y Cero Cambios L0:**
  No se modificaron esquemas SQLite (`MigrationV*.cs`), entidades de dominio compartido, ni `ConversorPDF.XLL.csproj`.
- **R4 — Escalado Estricto de Un Paso:**
  `AgEscaladorVentanaAula.Evaluar()` invoca exclusivamente `AgEscaleraVentanas.Siguiente()` sobre la ventana actual del aula. Si el siguiente paso no aporta una capacidad estrictamente mayor a la demanda requerida o no existe, no se emite ninguna sugerencia (no se admiten saltos dobles en una sola corrida).

---

## 4. Códigos de Diagnóstico y Planificación

- `I5_INFECTIBLE`: Emitido por el motor genético o el planificador de oleadas cuando las sesiones asignadas a una ventana superan la capacidad disponible o las restricciones de concurrencia de carriles.
- `I5_CAPACIDAD_CARRIL_PATRON`: Alerta generada cuando, a pesar de existir capacidad global nominal suficiente, el reparto de patrones de tutoria fragmenta los carriles e impide encajar las sesiones requeridas.
- `R8`: Regla de validación de suficiencia de ventana en pre-vuelo (`AgValidadorVentana`), la cual dispara el diálogo de escalado asistido en `AgFormGeneradorHorarios`.

---

## 5. UI y Sincronización en el Árbol Académico

1. **Modal Dinámico en Pre-Vuelo:**
   El mensaje de confirmación en `AgFormGeneradorHorarios` ya no contiene cadenas fijas de ventanas. En su lugar, itera las decisiones de escalado y construye la leyenda con `AgEscaleraVentanas.Rotular(d.Hacia, unidad)`.
2. **Recarga Estrecha de Dimensiones:**
   Tras aplicar las decisiones en base de datos (`AplicarAsync`), se invoca `RecargarDimensionesPerfilActivoTrasEscaladoAsync()`, el cual consulta los DTOs actualizados de `_arbolService.CargarPerfilAsync()` y actualiza las claves `AgNodoArbolKey` de los nodos de nivel 6 (modalidades) en `treeContextoAcademico` sin rebarajar combos ni destruir la selección en pantalla.
3. **Guarda Vertical y Tooltips Seguros:**
   Para evitar que clics espurios en el área de radios de días cambien involuntariamente la ventana escalada:
   - Se activó `treeContextoAcademico.ShowNodeToolTips = true`.
   - Se implementó la guarda `int dy = e.Y - e.Node.Bounds.Y; if (dy < 0 || dy > e.Node.Bounds.Height) return;` en `NodeMouseClick`.
   - En nodos cuya ventana actual sea extendida (`LUN-VIE`, `JUE-DOM`, `MIE-DOM`), se asigna `node.ToolTipText = $"Ventana actual: {dv} ({AgDiasVentana.AIso(dv)})"`, manteniéndolos visualmente informativos sin colisionar con los radios estándar.

---

## 6. Coordinación con Planes Posteriores

Este plan sienta las bases para el plan de soporte de ranuras dinámicas 4 a 8 horas (`2026-11-12-ranuras-dinamicas-4-a-8-horas-plan.md`), asegurando que cualquier manipulación de ventanas y bloques de decisión se realice sobre la semántica unificada de `AgDiasVentana` y `AgEscaleraVentanas`.
