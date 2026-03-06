param(
  [string]$InputPptx = "",
  [string]$LayoutName = "",
  [string]$OutputDir = "",
  [int]$SnapshotWidthPx = 1280,
  [switch]$SkipSnapshot,
  [switch]$AllowOOXML
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$RepoRoot = (Resolve-Path (Join-Path $PSScriptRoot ".."))
$ScriptPath = Join-Path $RepoRoot "scripts\windows_com_smoke.py"

if (-not $OutputDir) {
  $OutputDir = Join-Path $RepoRoot "artifacts\com-smoke"
}

$env:PYTHONPATH = Join-Path $RepoRoot "python"

$args = @(
  $ScriptPath,
  "--output-dir", $OutputDir,
  "--snapshot-width-px", "$SnapshotWidthPx"
)

if ($InputPptx) {
  $args += @("--input-pptx", (Resolve-Path $InputPptx).Path)
}
if ($LayoutName) {
  $args += @("--layout-name", $LayoutName)
}
if ($SkipSnapshot) {
  $args += "--skip-snapshot"
}
if ($AllowOOXML) {
  $args += "--allow-ooxml"
}

Write-Host "Running Windows COM smoke runner..."
Write-Host "Repo root: $RepoRoot"
Write-Host "Output dir: $OutputDir"

python @args
if ($LASTEXITCODE -ne 0) {
  throw "COM smoke runner failed with exit code $LASTEXITCODE"
}

Write-Host "COM smoke runner completed successfully."
