# Renombrar escalador de ranuras del aula

## Purpose

`AgEscaladorVentanaAula` pasa a `AgEscaladorRanuraAula`: propone ampliar las ranuras semanales disponibles del aula mediante la escalera de días. No amplía la fecha límite del calendario.

## Changes

- Renombrados clase, constructor y archivo de Infrastructure a `AgEscaladorRanuraAula.cs`.
- Actualizadas referencias de UI y registro Compile del proyecto XLL.
- Actualizadas referencias de pruebas; clases y archivos propios pasan a `AgEscaladorRanuraAulaTests` y `AgEscaladorRanuraAulaPersistenciaTests`.
- Actualizada documentación del diagnóstico transitorio. Marcadores `AG-RANURAS-TMP` conservados.
- Planes y registros históricos conservan el nombre vigente cuando se escribieron; corresponde al servicio renombrado.

## How to use

Usar `AgEscaladorRanuraAula` en código y filtros de pruebas futuros. Métodos, decisiones, persistencia y comportamiento conservados.

## Test coverage

Pruebas existentes actualizadas al nuevo símbolo; sin cambios de aserciones.
Verificación estática de referencias y rutas del proyecto. Sin compilación ni ejecución de pruebas, conforme a la restricción del usuario.
