# Plan: Reorganización Visual Módulo S (AgFormGeneradorHorarios)

**Fecha**: 2026-10-26  
**Objetivo**: Optimizar la distribución del espacio en la pestaña "Configuración & Ejecución (Módulo S)" de `AgFormGeneradorHorarios`:
1. Ampliar significativamente la altura visible del árbol de 6 niveles (`treeContextoAcademico`).
2. Reducir la altura del cuadro de texto del log (`txtLogAG`) a la mitad (~45-50px).
3. Mover la etiqueta de origen (`lblOrigenVentana`) junto al título del GroupBox de Ventana Temporal (`grpVentana`).
4. Reorganizar la barra inferior en una sola fila compacta: barra de progreso reducida (`progressBarAG`), etiqueta de estado (`lblStatusAG`) a su lado, y botones de acción alineados horizontalmente.

---

## 1. Análisis de Estado Actual vs. Propuesto

| Componente / Contenedor | Estado Actual | Estado Propuesto | Beneficio / Razón |
| :--- | :--- | :--- | :--- |
| **`splitContenedorSim.SplitterDistance`** | `461` (Panel2 = 148px) | `525` (Panel2 = ~84px) | Aumenta el alto de Panel1 en **+64px**. |
| **`txtLogAG`** | Alto ~90px (~10 líneas) | Alto ~45px (~5 líneas, scrollable) | Reducción del log a la mitad solicitada. |
| **`pnlConfiguracionInferior`** | Alto `82px` | Alto `56px` | Ahorra **+26px** adicionales para el árbol. |
| **`treeContextoAcademico`** | Alto `270px` | Alto **~360px** | **+33% de área visible** para selección de tuplas S. |
| **`lblOrigenVentana`** | Fila inferior de `grpVentana` (Y=52) | Cabecera superior junto al título (Y=0 o Y=1) | Elimina fila vacía en `grpVentana`, visualización integrada. |
| **`pnlBotonesSim`** | Alto `58px` (2 filas verticales) | Alto `38px` (1 sola fila horizontal) | Barra y estado en una fila sin desperdicio vertical. |
| **`progressBarAG`** | Ancho `370px`, apilado arriba | Ancho `150px`, Y=9, al lado de estado | Tamaño más compacto y balanceado. |
| **`lblStatusAG`** | Y=26, debajo de progressBar | Y=9, al lado derecho de progressBar | Lectura natural en una sola fila continua. |

---

## 2. Grafo de Dependencias

```
T1 (Reorganización Barra Inferior: Progreso + Estado en 1 fila)
  │
  ├───► T2 (Reubicación de lblOrigenVentana y compactación de grpVentana)
  │       │
  │       └───► T3 (Ajuste de SplitterDistance, txtLogAG y expansión del árbol)
  │               │
  └───────────────┴───► T4 (Validación de compilación, runtime y prueba visual)
```

---

## 3. Tareas Detalladas

### T1: Reorganización de Barra de Progreso y Estado en una sola fila
- **id**: `T1`
- **depends_on**: `[]`
- **location**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **description**:
  - Cambiar `pnlBotonesSim.Size` a `(1092, 38)`.
  - En `pnlProgreso`: fijar `Size = (430, 38)`.
  - Configurar `progressBarAG`: quitar `Dock = Top`, asignar `Location = new Point(5, 9)`, `Size = new Size(130, 20)`.
  - Configurar `lblStatusAG`: quitar `Dock = Top`, asignar `Location = new Point(142, 9)`, `Size = new Size(280, 20)`, `TextAlign = ContentAlignment.MiddleLeft`.
  - Ajustar `flowAcciones` y `chkModoDebug` para alinearse verticalmente a la altura de 38px (`Padding = (0, 4, 8, 4)`).
- **validation**: La barra de progreso y la etiqueta "Estado: En espera" quedan una al lado de la otra en una única fila horizontal sin desbordamiento.
- **status**: Completed
- **log**: Reorganizados pnlProgreso, progressBarAG, lblStatusAG, flowAcciones y botones en una sola fila limpia de 38px.
- **files edited/created**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`

---

### T2: Reubicación de `lblOrigenVentana` a la cabecera y compactación de `grpVentana`
- **id**: `T2`
- **depends_on**: `[T1]`
- **location**:
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`
- **description**:
  - En `AgFormGeneradorHorarios.Designer.cs`:
    - Cambiar `lblOrigenVentana.Location` a `new Point(206, 0)` (inmediatamente a la derecha del texto `"2. Ventana Temporal Semestral (α)"`) y llamar `lblOrigenVentana.BringToFront()`.
    - Reducir `grpVentana.Size` de `(490, 82)` a `(490, 56)`.
    - Reducir `grpParametrosAG.Size` de `(596, 82)` a `(596, 56)`.
    - Reducir `pnlSeparadorConfig.Size` de `(6, 82)` a `(6, 56)`.
    - Reducir `pnlConfiguracionInferior.Size` de `(1092, 82)` a `(1092, 56)`.
  - En `AgFormGeneradorHorarios.cs`:
    - Ajustar el formato en `ActualizarVentanaSemestralAsync` para que el texto empiece con `"· Origen: ..."` o `": Origen: ..."` de forma limpia y legible junto al título del GroupBox.
- **validation**: El texto de origen aparece junto al título de la ventana temporal; la fila inferior de `grpVentana` queda eliminada liberando 26px de espacio vertical.
- **status**: Completed
- **log**: Reubicada lblOrigenVentana a la cabecera del GroupBox, compactado pnlConfiguracionInferior a 56px y agregado prefijo con punto centrado en las trazas y estados.
- **files edited/created**: 
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
  - `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.cs`

---

### T3: Ajuste de SplitterDistance, reducción de `txtLogAG` y expansión del árbol
- **id**: `T3`
- **depends_on**: `[T2]`
- **location**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`
- **description**:
  - Modificar `splitContenedorSim.SplitterDistance` a `525` (anteriormente `461`).
  - El Panel2 de `splitContenedorSim` pasa a tener ~84px de alto total (38px de barra de botones + 46px de `txtLogAG`).
  - `txtLogAG` pasa de 90px a ~46px (exactamente la mitad).
  - Como `pnlConfiguracionInferior` se redujo en 26px y `SplitterDistance` creció en 64px, `grpSeleccion` y `treeContextoAcademico` ganan automáticamente ~90px adicionales de alto vertical (`Size.Height` de `treeContextoAcademico` pasa de 270px a ~360px).
- **validation**: El árbol ocupa la mayor parte de la pantalla visible mostrando más semestres/grupos simultáneamente sin necesidad de scroll excesivo, y el log se reduce a 4-5 líneas útiles.
- **status**: Completed
- **log**: SplitterDistance ajustado a 525, txtLogAG reducido a 46px de alto, otorgando +90px de espacio vertical adicional para treeContextoAcademico.
- **files edited/created**: `ConversorPDF.XLL/Academic/UI/Forms/AgFormGeneradorHorarios.Designer.cs`

---

### T4: Verificación de compilación (.NET 4.8) y revisión estética
- **id**: `T4`
- **depends_on**: `[T1, T2, T3]`
- **location**: Solución `ConversorPDF.sln`
- **description**:
  - Ejecutar compilación con MSBuild (`Release|x64` o `Debug|x64`) asegurando cero errores y cero advertencias nuevas.
  - Verificar que no se rompan eventos, anclajes (`Dock`, `Anchor`), ni cadenas de formato en runtime.
- **validation**: Compilación exitosa `0 errores` y verificación de consistencia visual.
- **status**: Completed
- **log**: Compilación Release|x64 completada con 0 errores. Suite de 939 pruebas unitarias ejecutada y superada al 100% sin regresiones (0 fallas, 939 superados).
- **files edited/created**: Solución completa validada. 

---

## 4. Grupos de Ejecución en Olas (Parallel Execution Waves)

| Ola (Wave) | Tareas | Condición de Inicio |
| :---: | :---: | :--- |
| **Wave 1** | `T1` | Inmediato (Reorganizar barra inferior progreso + estado) |
| **Wave 2** | `T2` | Tras completar `T1` (Reubicar etiqueta origen y compactar GroupBox) |
| **Wave 3** | `T3` | Tras completar `T2` (Ajustar Splitter, reducir log a la mitad y expandir árbol) |
| **Wave 4** | `T4` | Tras completar `T3` (Compilación y verificación integral) |

---

## 5. Riesgos y Mitigaciones
1. **Solapamiento en títulos de GroupBox en diferentes resoluciones / DPI**:
   - *Riesgo*: Si el texto del título `"2. Ventana Temporal Semestral (α)"` cambia de fuente según el DPI, `lblOrigenVentana` podría solaparse si está fijo en coordenadas absolutas.
   - *Mitigación*: En el `Load` o en el evento de redibujado/actualización, se calcula `lblOrigenVentana.Left = grpVentana.Padding.Left + TextRenderer.MeasureText(grpVentana.Text, grpVentana.Font).Width + 8;` o se ubica con margen seguro garantizado.
2. **Auto-scroll de `txtLogAG`**:
   - *Riesgo*: Al reducirse a la mitad (~46px), el log podría ocultar trazas rápidas.
   - *Mitigación*: Se preserva `ScrollBars = ScrollBars.Vertical` y el comportamiento existente de append/scroll al final para que el último evento siempre sea visible.
