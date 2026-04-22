---
description: Build, generate, and preview the technical documentation using DocFX
---

# Build & Preview — Documentación Técnica (DocFX)

Este workflow cubre los pasos necesarios para generar y previsualizar la documentación técnica autogenerada (API) y los artículos conceptuales. El proyecto utiliza **DocFX** para compilar los comentarios XML del código fuente y los archivos Markdown (`docs/*.md`) en un sitio web estático compatible con GitHub Pages.

---

## 1. Generar la Documentación

El script de compilación automatiza todo el proceso: instala DocFX si es necesario, compila el código C# para generar el XML actualizado, construye el sitio usando `docfx build`, y genera el archivo `.nojekyll` para despliegue.

// turbo-all
```powershell
powershell -ExecutionPolicy Bypass -File ".\scripts\build_docs.ps1"
```

## 2. Previsualizar el sitio web generado

Una vez generado, puedes levantar un servidor HTTP ligero incluido con DocFX para asegurar que todo se vea correctamente antes de subir los cambios.

```powershell
docfx serve docs/_site --port 8080
```

*Después de esto abre tu navegador en: [http://localhost:8080/ConversorPDF/](http://localhost:8080/ConversorPDF/)*

---

## 3. Secciones Obligatorias para el Proceso de Documentación

Para asegurar la calidad y mantenibilidad, **todo cambio en la lógica de negocio o arquitectura** debe completar estas secciones obligatorias:

### 3.1 Documentación de API (Triple-slash)
- **Comentarios XML detallados**: Añadir bloques `///` en el código fuente para que DocFX los incluya en la referencia técnica.
- **Listado de Operaciones**: Documentar explícitamente las nuevas operaciones, parámetros y tipos de retorno.

### 3.2 Guías Conceptuales (Arquitectura y Modelos)
- **Modelado de Datos**: Los artículos en `docs/articles/` deben reflejar la arquitectura basada en modelos (ej: `ObligacionExtract`).
- **Lógica de Procesos**: Actualizar los artículos conceptuales para que el código fuente y la documentación hablen el mismo lenguaje (arquitectura y funcionalidad).

### 3.3 Guía del Desarrollador (Ejemplos y Casos de Uso)
- **Sección de Ejemplos**: Añadir siempre un bloque de código en `developer-guide.md` que muestre cómo consumir la nueva funcionalidad o el servicio refactorizado.

---

## 4. Checklist de Validación (Mandatorio)

Antes de finalizar la tarea, verifica:
- [ ] ¿Toda nueva clase/método público tiene comentarios XML?
- [ ] ¿Se actualizó `docs/articles/[area].md` con la lógica funcional del nuevo proceso?
- [ ] ¿Hay un ejemplo reproducible en `developer-guide.md`?
- [ ] ¿El build de DocFX (`build_docs.ps1`) finaliza sin errores (`0 errors`)?

---

## 5. Mejores Prácticas

- **Imágenes**: Si necesitas agregar capturas de pantalla, ponlas en `docs/images/` y referéncialas usando rutas relativas en los `.md`.
- **Previsualización**: Siempre verifica el sitio localmente con `docfx serve` antes de confirmar los cambios.
