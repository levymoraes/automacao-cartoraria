param([switch]$CheckInstallation,[switch]$DiagnoseCredential)
$ErrorActionPreference='Stop'
foreach($filename in @('SecretStorage.ps1','check-linear-connection.cjs')){
 if(-not (Test-Path -LiteralPath (Join-Path $PSScriptRoot $filename))){throw 'LINEAR_VERIFIER_INSTALLATION_INCOMPLETE'}
}
$nodeCommand=Get-Command node -ErrorAction SilentlyContinue
if(-not $nodeCommand){throw 'NODE_22_OR_NEWER_REQUIRED'}
if($CheckInstallation){
 [pscustomobject]@{status='LINEAR_VERIFIER_INSTALLED';script='Test-LinearConnection.ps1';alias='Test-Linear-Connection.ps1';credentialRead=$false;networkRequests=0;openaiCalls=0} | ConvertTo-Json
 return
}
. (Join-Path $PSScriptRoot 'SecretStorage.ps1')
$projectRoot=Split-Path -Parent $PSScriptRoot
$keyPath=Join-Path $projectRoot 'storage\secrets\linear-api-key.txt'
Assert-SecretAcl -Path $keyPath -UserSid (Get-CurrentSecretSid)
# O Node le a chave somente para enviar o header; nunca a devolve ao console.
$arguments=@((Join-Path $PSScriptRoot 'check-linear-connection.cjs'))
if($DiagnoseCredential){$arguments+='--diagnose-credential'}
& $nodeCommand.Source @arguments
if($LASTEXITCODE -ne 0){throw 'LINEAR_CONNECTION_CHECK_FAILED: consulte o codigo seguro acima. -DiagnoseCredential verifica somente codificacao/formato, sem imprimir a chave ou acessar a rede.'}
