param([switch]$CheckInstallation,[switch]$DiagnoseCredential)
$ErrorActionPreference='Stop'
# Compatibilidade com a grafia usada pelo usuario. Mantem um unico verificador.
& (Join-Path $PSScriptRoot 'Test-LinearConnection.ps1') @PSBoundParameters
