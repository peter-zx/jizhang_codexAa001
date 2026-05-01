param(
    [string]$HostAddress = "127.0.0.1",
    [int]$Port = 8000
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
Set-Location $Root
if (!(Test-Path ".env")) {
    & "$PSScriptRoot\use-env.ps1" -Environment development
}
if (!(Test-Path ".venv\Scripts\python.exe")) {
    throw "Virtual environment is missing. Create it with: py -m venv .venv"
}
& ".venv\Scripts\python.exe" -m uvicorn app.main:app --host $HostAddress --port $Port
