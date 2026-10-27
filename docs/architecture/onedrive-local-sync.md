# Regla de Sincronización: Copia Local y OneDrive

## Contexto y justificación

El árbol principal vive en OneDrive (`...\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.6.3.72`). Compilar (MSBuild/Roslyn/Excel-DNA) y ejecutar tests ahí provoca locks y latencia, y hace inviables los `git worktree`.

**Lección de la auditoría 2026-10-27:** `git clone . C:\dev\ConversorPDF` NO reproduce el árbol cuando el working tree tiene cambios sin commitear. El clon quedó con 98 `.cs` en XLL y 13 en Tests (árbol real: 452/138) y sin el slice `Academic`; toda calibración derivada fue inválida.

## Cómo crear la copia local (dos opciones válidas)

### Opción A — Commit primero, clon después (recomendada a medio plazo)
```powershell
# En el repo de OneDrive: revisar y commitear el árbol real (sin bin/obj/_evidence)
git status --short
git add -A
git commit -m "chore: snapshot del arbol real antes de worktrees"
# Solo ahora el clon es fiel
git clone . C:\dev\ConversorPDF
```

### Opción B — Snapshot sin tocar el historial (inmediata)
```powershell
robocopy "<repo OneDrive>" "C:\dev\ConversorPDF" /MIR /XD .git bin obj TestResults _evidence .vs
# El snapshot no tendrá .git: sirve para compilar/testear, no para commits por tarea
```

## Verificación obligatoria

```powershell
powershell -ExecutionPolicy Bypass -File scripts/preflight.ps1 -Workspace C:\dev\ConversorPDF
```
Debe devolver `0`. Si devuelve `1`, la copia es un snapshot incompleto: **no** se planifica ni se ejecuta swarm sobre ella.

## Reglas

1. **Nunca `git push`.** El swarm y la copia local solo generan commits locales; publicar es una decisión humana explícita (stop condition del envelope).
2. **No sincronizar binarios**: `bin/`, `obj/`, `TestResults/` y `_evidence/` están en `.gitignore`.
3. **Un solo sentido**: la copia local recibe; los cambios vuelven al repo de OneDrive por commits revisados, no por copia manual de archivos aislados.
4. **Un solo workspace activo por sesión**: dos árboles con dos `bin/` distintos producen evidencia no comparable.
