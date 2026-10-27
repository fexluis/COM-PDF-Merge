# Context Pack: Academic - {{SLICE_NAME}}

**Fecha de generación**: {{DATE}}  
**Presupuesto de contexto**: ≤ 400 líneas

---

## 1. Responsabilidad del Slice y Límites de Dominio
- **Propósito**: Descripción concisa de la capacidad de negocio administrada por este slice.
- **Entidades de Dominio** (`ConversorPDF.XLL/Academic/Domain/Models`):
  - Modelos puros de C# sin dependencias externas.
- **Servicios de Aplicación** (`ConversorPDF.XLL/Academic/Application/Services`):
  - Orquestación y lógica de negocio aislada.

---

## 2. Contratos e Interfaces Clave
- **Interfaces de Dominio** (`Academic/Domain/Interfaces`):
  - Repositorios y contratos limpios.
- **Interfaces de Aplicación** (`Academic/Application/Interfaces`):
  - Servicios de coordinación y DTOs de transporte.

---

## 3. Almacenamiento y Esquema SQLite
- **Tablas principales**:
  - Definición de tablas en SQLite y campos esenciales.
- **Migraciones asociadas**:
  - `MigrationV*.cs` que afectan a este slice.

---

## 4. UI y Puntos de Entrada
- **Formularios WinForms** (`ConversorPDF.XLL/Academic/UI/Forms`):
  - Único formulario público del slice.
- **Acceso Ribbon**:
  - Botón o controlador en `ConversorPDF.XLL/Ribbon/RibbonController.cs`.

---

## 5. Gotchas y Lecciones Aprendidas
- Consideraciones especiales de concurrencia o locks en SQLite.
- Reglas de mapeo o transformaciones de fecha/hora.
- Restricciones de Single-Threaded Apartment (STA) para controles WinForms.

---

## 6. Archivos Críticos y Referencias en Disco
- Lista de rutas relativas de los archivos principales del slice para consulta rápida de los subagentes.
