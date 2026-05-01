param(
    [ValidateSet("development", "production")]
    [string]$Environment = "development"
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $PSScriptRoot
$EnvFile = Join-Path $Root "env\$Environment.env"
if (!(Test-Path $EnvFile)) {
    throw "Environment file not found: $EnvFile"
}
Copy-Item $EnvFile (Join-Path $Root ".env") -Force
Write-Host "Activated environment: $Environment"
