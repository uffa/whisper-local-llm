param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    [string]$Model = "small.en"
)

$ErrorActionPreference = "Stop"

# Resolve the project root from this script's location so the whole folder
# is portable (this file lives at <root>/scripts/transcribe.ps1).
$whisperRoot = Split-Path -Parent $PSScriptRoot
$exe = Join-Path $whisperRoot "whisper-cli.exe"
$modelPath = Join-Path $whisperRoot "models\ggml-$Model.bin"

if (!(Test-Path $exe)) {
    throw "whisper-cli.exe not found at $exe"
}

if (!(Test-Path $modelPath)) {
    throw "Model not found at $modelPath. Download a ggml model from https://huggingface.co/ggerganov/whisper.cpp/tree/main and place it in the models/ folder."
}

if (!(Test-Path $InputFile)) {
    throw "Input file not found: $InputFile"
}

# Auto-detect the newest installed CUDA version so we don't need to update
# this script every time the toolkit is upgraded. Falls through silently if
# CUDA isn't installed — whisper-cli will run on CPU.
$cudaRoot = "C:\Program Files\NVIDIA GPU Computing Toolkit\CUDA"
if (Test-Path $cudaRoot) {
    $latestCuda = Get-ChildItem $cudaRoot -Directory -Filter "v*" -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match '^v(\d+)\.(\d+)$' } |
        Sort-Object { [version]($_.Name -replace '^v') } -Descending |
        Select-Object -First 1
    if ($latestCuda) {
        $cudaBin = Join-Path $latestCuda.FullName "bin"
        if (Test-Path $cudaBin) {
            $env:Path = "$cudaBin;$env:Path"
        }
    }
}

Push-Location $whisperRoot
try {
    & $exe -m $modelPath -f $InputFile -otxt
}
finally {
    Pop-Location
}
