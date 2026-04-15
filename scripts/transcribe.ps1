param(
    [Parameter(Mandatory=$true)]
    [string]$InputFile,

    # Model filename without the "ggml-" prefix or ".bin" extension.
    # Examples: "small.en", "medium.en", "large-v3-turbo", "large-v3".
    [string]$Model = "large-v3-turbo",

    # Initial prompt fed to whisper.cpp to bias the decoder toward specific
    # vocabulary. Keep it short (a sentence or two plus a list of terms).
    # whisper.cpp's prompt is capped by the model's context window, so don't
    # stuff too much here — 200 characters is a safe ceiling.
    [string]$Prompt = "Technical dictation notes. Common terms: Claude, ChatGPT, Anthropic, whisper.cpp, AutoHotkey, ffmpeg, PowerShell, GitHub, CUDA, NVIDIA, JavaScript, TypeScript, Python, npm.",

    # Beam search width. Higher = more accurate, slower. 5 is whisper.cpp's
    # default; 8 is the maximum whisper.cpp allows and gives a noticeable
    # accuracy bump on the larger models at negligible cost on a modern GPU.
    [int]$BeamSize = 8
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
    & $exe `
        -m $modelPath `
        -f $InputFile `
        -otxt `
        --beam-size $BeamSize `
        --prompt $Prompt
}
finally {
    Pop-Location
}
