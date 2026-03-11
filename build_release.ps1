$ErrorActionPreference = 'Stop'
$env:Configuration = 'Release'
$env:Platform = 'x64'
& .\build_project.ps1