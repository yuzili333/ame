param(
  [ValidateSet('gateway', 'studio', 'both')]
  [string]$Mode = 'both'
)
$ErrorActionPreference = 'Stop'
Set-Location (Split-Path -Parent $PSScriptRoot)
if (-not (Test-Path '.env')) { throw 'Copy .env.example to .env and set deployment values first.' }
if ($Mode -eq 'gateway' -or $Mode -eq 'both') {
  Start-Process -FilePath 'npm.cmd' -ArgumentList 'run', 'gateway' -WorkingDirectory (Get-Location).Path -NoNewWindow
}
if ($Mode -eq 'studio' -or $Mode -eq 'both') {
  npm run dev
}
