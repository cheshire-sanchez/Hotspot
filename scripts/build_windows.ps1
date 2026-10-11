$ErrorActionPreference = 'Stop'

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
$BuildDir = Join-Path $RepoRoot 'build/windows'
$HotSpotExe = Join-Path $BuildDir 'hotspot.exe'
$GccCommand = Get-Command gcc -ErrorAction SilentlyContinue
if (-not $GccCommand) {
    throw 'MinGW-w64 GCC was not found on PATH. Install MinGW-w64 and reopen PowerShell.'
}

New-Item -ItemType Directory -Force -Path $BuildDir | Out-Null
$Sources = @(
    'hotspot.c', 'microchannel.c', 'materials.c', 'temperature.c', 'RCutil.c',
    'package.c', 'temperature_block.c', 'temperature_grid.c', 'flp.c',
    'flp_desc.c', 'npe.c', 'shape.c', 'util.c', 'wire.c'
) | ForEach-Object { Join-Path $RepoRoot $_ }

Write-Host "Building native Windows HotSpot: $HotSpotExe"
& $GccCommand.Source -O3 -DVERBOSE=1 -DMATHACCEL=0 -DSUPERLU=0 `
    -o $HotSpotExe @Sources -lm
if ($LASTEXITCODE -ne 0) {
    throw "GCC failed with exit code $LASTEXITCODE"
}
Write-Host 'Build succeeded.'
