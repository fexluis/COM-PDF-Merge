import os
import re

repo_path = r"c:\Users\usuario\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.2.59.21\ConversorPDF.XLL\Academic\Infrastructure\Persistence\Repositories\AcademicRepository.cs"

with open(repo_path, 'r', encoding='utf-8') as f:
    content = f.read()

# Replace all MarkDirty calls
content = content.replace("ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();", "NotifyMarkDirty();")

# Inject EnsureCrSqliteEnabled() at the beginning of the body of Insert/Update/Delete/ExecuteRawSql/SeedCatalogosTransaccional
# Regex finds: public SqliteResult... MethodName(...) {
# And adds EnsureCrSqliteEnabled(); right after {
pattern = r"(public\s+SqliteResult(?:<[^>]+>)?\s+(?:Insert|Update|Delete|ExecuteRawSql|SeedCatalogosTransaccional)\w*\s*\([^)]*\)\s*\{)"
replacement = r"\1\n            EnsureCrSqliteEnabled();"

content = re.sub(pattern, replacement, content)

with open(repo_path, 'w', encoding='utf-8') as f:
    f.write(content)

print("AcademicRepository updated.")

salvedad_path = r"c:\Users\usuario\OneDrive - Escuela Superior de Administración Publica\Antigravity\ConversorPDFv1.2.59.21\ConversorPDF.XLL\Academic\Infrastructure\Persistence\Repositories\SalvedadConflictoRepository.cs"

with open(salvedad_path, 'r', encoding='utf-8') as f:
    content2 = f.read()

content2 = content2.replace("ConversorPDF.XLL.Core.Sync.SyncContext.Scheduler?.MarkDirty();", "NotifyMarkDirty();")

pattern2 = r"(public\s+async\s+Task<Result>\s+(?:SaveAsync|DeleteAsync)\w*\s*\([^)]*\)\s*\{)"
replacement2 = r"\1\n            EnsureCrSqliteEnabled();"

content2 = re.sub(pattern2, replacement2, content2)

with open(salvedad_path, 'w', encoding='utf-8') as f:
    f.write(content2)

print("SalvedadConflictoRepository updated.")
