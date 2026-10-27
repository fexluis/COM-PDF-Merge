# Directorio de Context Packs (`.agents/context/`)

Este directorio almacena los paquetes de contexto acotados por vertical slice o componente funcional para alimentar a `swarm-planner` y a los subagentes ejecutores.

---

## 1. Principio del Presupuesto de Contexto
- **Límite estricto**: Cada context pack debe contener **≤ 400 líneas** de texto curado.
- **Objetivo**: Evitar inflar la ventana de contexto de los subagentes con los 452 archivos del repositorio legacy, proveyendo únicamente el mapa mental y contratos relevantes del slice objetivo.

---

## 2. Estructura Oficial de un Context Pack (`_template.md`)

Todo context pack debe estructurarse en 6 secciones fijas:

```markdown
# Context Pack: Academic - <SliceName>

## 1. Responsabilidad del Slice y Límites de Dominio
- Descripción concisa de la capacidad de negocio.
- Entidades del Clean World (`Domain/Models`).
- Servicios del Application layer (`Application/Services`).

## 2. Contratos e Interfaces Clave
- Interfaces expuestas para interacción (`Domain/Interfaces`, `Application/Interfaces`).
- DTOs principales.

## 3. Almacenamiento y Esquema SQLite
- Tablas relevantes en SQLite.
- Triggers, índices o restricciones conocidas.
- Migraciones `MigrationV*.cs` asociadas.

## 4. UI y Puntos de Entrada
- Form WinForms asociado (`UI/Forms`).
- Entrada desde `RibbonController.cs`.

## 5. Gotchas y Lecciones Aprendidas
- Problemas de concurrencia, locking de SQLite o STA threads.
- Formatos especiales de fecha, parsing o validaciones.

## 6. Archivos Críticos y Referencias en Disco
- Lista de rutas exactas con enlaces relativos.
```

---

## 3. Generación Automatizada
Se puede generar un borrador preliminar de context pack ejecutando:
```powershell
powershell -ExecutionPolicy Bypass -File scripts/context-pack.ps1 -Slice <NombreSlice>
```
