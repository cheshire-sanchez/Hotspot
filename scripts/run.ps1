param(
    [string]$Design = 'HBFuse-4',
    [string]$Excel = 'W4_power_4.xlsx'
)

$ErrorActionPreference = 'Stop'
$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot '..')).Path
if (-not [System.IO.Path]::IsPathRooted($Design)) {
    $Design = Join-Path $RepoRoot $Design
}
$Design = (Resolve-Path $Design).Path
$DesignName = Split-Path $Design -Leaf
$SourceXlsx = Join-Path (Join-Path $RepoRoot 'psource') (Split-Path $Excel -Leaf)
$Lcf = Join-Path $Design 'example.lcf'
$Config = Join-Path $Design 'example.config'
$Materials = Join-Path $Design 'example.materials'
$HotSpotExe = Join-Path $RepoRoot 'build/windows/hotspot.exe'
$PtraceDir = Join-Path $RepoRoot 'ptrace'
$OutputBase = Join-Path (Join-Path $RepoRoot 'outputs') $DesignName
$SamplingInterval = '0.001'
$DecodeCount = '129'

foreach ($RequiredPath in @($SourceXlsx, $Lcf, $Config, $Materials)) {
    if (-not (Test-Path -LiteralPath $RequiredPath -PathType Leaf)) {
        throw "Required input not found: $RequiredPath"
    }
}

$PythonCommand = Get-Command python -ErrorAction SilentlyContinue
if (-not $PythonCommand) {
    throw 'Python was not found on PATH.'
}
& $PythonCommand.Source -c 'import openpyxl, numpy, matplotlib'
if ($LASTEXITCODE -ne 0) {
    throw 'Python needs openpyxl, numpy, and matplotlib. Install with: python -m pip install openpyxl numpy matplotlib'
}

if (-not (Test-Path -LiteralPath $HotSpotExe -PathType Leaf)) {
    & (Join-Path $PSScriptRoot 'build_windows.ps1')
}
if (-not (Test-Path -LiteralPath $HotSpotExe -PathType Leaf)) {
    throw "Native HotSpot executable was not created: $HotSpotExe"
}

New-Item -ItemType Directory -Force -Path $PtraceDir, $OutputBase | Out-Null
$PrefillTrace = Join-Path $PtraceDir "$DesignName`_prefill.ptrace"
$DecodeTrace = Join-Path $PtraceDir "$DesignName`_decode.ptrace"
$TimePrefix = Join-Path $PtraceDir "$DesignName.time.tsv"
$UnitsFile = Join-Path $PtraceDir "$DesignName.units.tsv"

Write-Host "Generating phase traces from $SourceXlsx"
& $PythonCommand.Source (Join-Path $PSScriptRoot 'ptrace_gen.py') `
    --input-xlsx $SourceXlsx --lcf $Lcf --output-dir $PtraceDir `
    --output-prefill $PrefillTrace --output-decode $DecodeTrace `
    --output-time $TimePrefix --output-units $UnitsFile `
    --sampling-intvl $SamplingInterval --decode-count $DecodeCount
if ($LASTEXITCODE -ne 0) { throw "ptrace generation failed: $LASTEXITCODE" }

$ConfigText = Get-Content -LiteralPath $Config -Raw
if ($ConfigText -notmatch '(?m)^\s*-grid_rows\s+(\d+)') { throw 'Could not read -grid_rows from config.' }
$GridRows = [int]$Matches[1]
if ($ConfigText -notmatch '(?m)^\s*-grid_cols\s+(\d+)') { throw 'Could not read -grid_cols from config.' }
$GridCols = [int]$Matches[1]

function Invoke-HotSpotStage {
    param(
        [string]$Stage,
        [string]$TracePath,
        [string]$StageTimePath
    )

    $StageDir = Join-Path $OutputBase $Stage
    New-Item -ItemType Directory -Force -Path $StageDir | Out-Null
    Copy-Item -LiteralPath $StageTimePath, $UnitsFile -Destination $StageDir -Force
    $GridTransient = Join-Path $StageDir 'example.grid.ttrace'
    $GridSteady = Join-Path $StageDir 'example.grid.steady'
    $Steady = Join-Path $StageDir 'example.steady'
    $Transient = Join-Path $StageDir 'example.ttrace'
    Remove-Item -LiteralPath $GridTransient, $GridSteady, $Steady, $Transient -Force -ErrorAction SilentlyContinue
    New-Item -ItemType File -Path $GridTransient | Out-Null

    Write-Host "Running $Stage HotSpot simulation"
    Push-Location $Design
    try {
        & $HotSpotExe -c $Config -p $TracePath -grid_layer_file $Lcf `
            -materials_file $Materials -model_type grid -detailed_3D on `
            -sampling_intvl $SamplingInterval -init_file '(null)' `
            -steady_file $Steady -grid_steady_file $GridSteady `
            -o $Transient -grid_transient_file $GridTransient
        if ($LASTEXITCODE -ne 0) { throw "HotSpot failed for $Stage (exit $LASTEXITCODE)" }
    }
    finally {
        Pop-Location
    }

    $LayerHeaders = Select-String -LiteralPath $GridSteady -Pattern '^Layer\s+\d+:'
    if (-not $LayerHeaders) { throw "No layer data found in $GridSteady" }
    & $PythonCommand.Source (Join-Path $PSScriptRoot 'split_grid_steady.py') `
        $GridSteady $LayerHeaders.Count $GridRows $GridCols
    if ($LASTEXITCODE -ne 0) { throw "Grid split failed for $Stage" }

    $UnitRows = Import-Csv -LiteralPath $UnitsFile -Delimiter "`t"
    $PowerLayers = $UnitRows | Group-Object layer_number | ForEach-Object { $_.Group[0] }
    foreach ($Layer in $PowerLayers) {
        $FlpPath = $Layer.flp_reference
        if (-not [System.IO.Path]::IsPathRooted($FlpPath)) {
            $FlpPath = Join-Path $Design $FlpPath
        }
        $LayerGrid = Join-Path $StageDir "example_layer$($Layer.layer_number).grid.steady"
        $MapBase = "layer$($Layer.layer_number)_$([System.IO.Path]::GetFileNameWithoutExtension($FlpPath))"
        & $PythonCommand.Source (Join-Path $PSScriptRoot 'grid_thermal_map.py') `
            $FlpPath $LayerGrid $GridRows $GridCols (Join-Path $StageDir "$MapBase.png")
        if ($LASTEXITCODE -ne 0) { throw "Thermal-map generation failed for layer $($Layer.layer_number)" }
    }
}

Invoke-HotSpotStage 'prefill' $PrefillTrace ($TimePrefix.Replace('.tsv', '_prefill.tsv'))
Invoke-HotSpotStage 'decode' $DecodeTrace ($TimePrefix.Replace('.tsv', '_decode.tsv'))
Write-Host "Completed. Results: $OutputBase\prefill and $OutputBase\decode"
