# Compativel com Windows PowerShell 5.1 e PowerShell 7. Nenhuma funcao imprime segredos.
Set-StrictMode -Version Latest
function Get-CurrentSecretSid { return [Security.Principal.WindowsIdentity]::GetCurrent().User }
function ConvertTo-ValidatedSecret {
 param([Security.SecureString]$Secret)
 $pointer=[IntPtr]::Zero;$plain=$null;$normalized=$null
 try{
  $pointer=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
  $plain=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer)
  $normalized=[regex]::Replace($plain,'^[\s\u200B\u200E\u200F\u2060\uFEFF]+|[\s\u200B\u200E\u200F\u2060\uFEFF]+$','')
  if($normalized.Length -eq 0 -or $normalized.Length -gt 4096 -or $normalized -cnotmatch '^[A-Za-z0-9_-]+$'){throw 'SECRET_INVALID_CHARACTERS: copie somente o valor da chave, sem rotulos, aspas ou caracteres extras. Nenhum arquivo foi substituido.'}
  return ConvertTo-SecureString -String $normalized -AsPlainText -Force
 }finally{
  if($pointer -ne [IntPtr]::Zero){[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)}
  $plain=$null;$normalized=$null
 }
}
function Read-SecretFromClipboard {
 param([scriptblock]$ClipboardReader={Get-Clipboard -Raw -ErrorAction Stop})
 $plain=$null;$temporary=$null
 try{
  try{$plain=& $ClipboardReader}catch{throw 'SECRET_CLIPBOARD_UNAVAILABLE'}
  if($plain -isnot [string] -or [string]::IsNullOrEmpty($plain)){throw 'SECRET_CLIPBOARD_EMPTY'}
  if($plain -match '(?im)(Set-LocalSecret\.ps1|Test-Linear-?Connection\.ps1|^\s*Set-Location\b|^\s*PS\s+[A-Z]:\\|^\s*\.\\scripts\\)'){
   throw 'SECRET_CLIPBOARD_CONTAINS_COMMAND: a area de transferencia contem um comando, nao a chave. O comando copiado substituiu a chave. Copie o valor da chave depois de iniciar o configurador.'
  }
  $temporary=ConvertTo-SecureString -String $plain -AsPlainText -Force
  return ConvertTo-ValidatedSecret -Secret $temporary
 }finally{$plain=$null;if($temporary){$temporary.Dispose()}}
}
function Read-ConfirmedSecretFromClipboard {
 param(
  [scriptblock]$ConfirmReady={
   Write-Host 'AGORA copie somente o valor da chave no Linear. Nao copie outro comando depois disso.'
   Write-Host 'Volte a esta janela e pressione somente Enter. Nao precisa colar a chave no console.'
   $confirmation=Read-Host 'Pressione Enter quando a chave estiver copiada' -AsSecureString
   if($confirmation){$confirmation.Dispose()}
  },
  [scriptblock]$ClipboardReader={Get-Clipboard -Raw -ErrorAction Stop}
 )
 # Captura somente depois de o comando ja ter sido colado/executado e de nova copia.
 & $ConfirmReady | Out-Null
 return Read-SecretFromClipboard -ClipboardReader $ClipboardReader
}
function Set-SecretSecurity {
 param([string]$Path,[Security.AccessControl.FileSystemSecurity]$Acl,[switch]$Directory)
 # Set-Acl pode solicitar SACL/SeSecurityPrivilege. Persistir somente secoes modificadas.
 $item=if($Directory){[IO.DirectoryInfo]::new($Path)}else{[IO.FileInfo]::new($Path)}
 if($PSVersionTable.PSEdition -eq 'Core'){[IO.FileSystemAclExtensions]::SetAccessControl($item,$Acl)}else{$item.SetAccessControl($Acl)}
}
function New-SecretAcl {
 param([Security.Principal.SecurityIdentifier]$UserSid,[switch]$Directory)
 $acl=if($Directory){[Security.AccessControl.DirectorySecurity]::new()}else{[Security.AccessControl.FileSecurity]::new()}
 $acl.SetAccessRuleProtection($true,$false)
 $inherit=if($Directory){[Security.AccessControl.InheritanceFlags]'ContainerInherit,ObjectInherit'}else{[Security.AccessControl.InheritanceFlags]::None}
 $system=[Security.Principal.SecurityIdentifier]::new('S-1-5-18')
 $systemRights=if($Directory){[Security.AccessControl.FileSystemRights]::ReadAndExecute}else{[Security.AccessControl.FileSystemRights]::Read}
 $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($UserSid,[Security.AccessControl.FileSystemRights]::FullControl,$inherit,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow))
 $acl.AddAccessRule([Security.AccessControl.FileSystemAccessRule]::new($system,$systemRights,$inherit,[Security.AccessControl.PropagationFlags]::None,[Security.AccessControl.AccessControlType]::Allow))
 return $acl
}
function Assert-SecretAcl {
 param([string]$Path,[Security.Principal.SecurityIdentifier]$UserSid,[switch]$AllowOwnerTransfer)
 $item=Get-Item -LiteralPath $Path -Force -ErrorAction Stop
 if(($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'SECRET_REPARSE_POINT_FORBIDDEN'}
 $acl=Get-Acl -LiteralPath $Path -ErrorAction Stop
 if(-not $acl.AreAccessRulesProtected){throw 'SECRET_INHERITANCE_NOT_PROTECTED'}
 if(-not $AllowOwnerTransfer -and $acl.GetOwner([Security.Principal.SecurityIdentifier]).Value -ne $UserSid.Value){throw 'SECRET_OWNER_MISMATCH'}
 $rules=@($acl.GetAccessRules($true,$true,[Security.Principal.SecurityIdentifier]))
 if($rules.Count -ne 2){throw 'SECRET_UNEXPECTED_ACCESS_RULES'}
 $foundUser=$false;$foundSystem=$false
 foreach($rule in $rules){
  if($rule.IsInherited -or $rule.AccessControlType -ne [Security.AccessControl.AccessControlType]::Allow){throw 'SECRET_UNSAFE_ACCESS_RULE'}
  if($rule.IdentityReference.Value -eq $UserSid.Value){
   if($rule.FileSystemRights -ne [Security.AccessControl.FileSystemRights]::FullControl){throw 'SECRET_USER_RIGHTS_INVALID'};$foundUser=$true
  }elseif($rule.IdentityReference.Value -eq 'S-1-5-18'){
   $allowed=[Security.AccessControl.FileSystemRights]'ReadAndExecute,Synchronize'
   if(($rule.FileSystemRights -band (-bnot [int]$allowed)) -ne 0){throw 'SECRET_SYSTEM_RIGHTS_EXCESSIVE'};$foundSystem=$true
  }else{throw 'SECRET_UNAUTHORIZED_PRINCIPAL'}
 }
 if(-not $foundUser -or -not $foundSystem){throw 'SECRET_REQUIRED_PRINCIPAL_MISSING'}
}
function Protect-ExistingSecret {
 param([string]$Path,[Security.Principal.SecurityIdentifier]$UserSid=(Get-CurrentSecretSid),[switch]$OwnerTransfer)
 $item=Get-Item -LiteralPath $Path -Force -ErrorAction Stop
 if(($item.Attributes -band [IO.FileAttributes]::ReparsePoint) -ne 0){throw 'SECRET_REPARSE_POINT_FORBIDDEN'}
 $acl=New-SecretAcl -UserSid $UserSid -Directory:$item.PSIsContainer
 try{Set-SecretSecurity -Path $Path -Acl $acl -Directory:$item.PSIsContainer}catch{throw 'SECRET_OWNER_REPAIR_REQUIRED'}
 if(-not $OwnerTransfer){
  $updated=Get-Acl -LiteralPath $Path;$updated.SetOwner($UserSid)
  try{Set-SecretSecurity -Path $Path -Acl $updated -Directory:$item.PSIsContainer}catch{throw 'SECRET_OWNER_REPAIR_REQUIRED'}
 }
 Assert-SecretAcl -Path $Path -UserSid $UserSid -AllowOwnerTransfer:$OwnerTransfer
}
function Initialize-SecretDirectory {
 param([string]$Path,[Security.Principal.SecurityIdentifier]$UserSid=(Get-CurrentSecretSid))
 if(Test-Path -LiteralPath $Path){Protect-ExistingSecret -Path $Path -UserSid $UserSid;return}
 $acl=New-SecretAcl -UserSid $UserSid -Directory;$acl.SetOwner($UserSid)
 if($PSVersionTable.PSEdition -eq 'Core'){[IO.FileSystemAclExtensions]::Create([IO.DirectoryInfo]::new($Path),$acl)}else{[void][IO.Directory]::CreateDirectory($Path,$acl)}
 Assert-SecretAcl -Path $Path -UserSid $UserSid
}
function Write-RestrictedSecret {
 param([string]$Path,[Security.SecureString]$Secret,[scriptblock]$BeforeWrite={})
 $userSid=Get-CurrentSecretSid
 Initialize-SecretDirectory -Path (Split-Path -Parent $Path) -UserSid $userSid
 if(Test-Path -LiteralPath $Path){Protect-ExistingSecret -Path $Path -UserSid $userSid}
 $temporary=Join-Path (Split-Path -Parent $Path) ('.secret-'+[Guid]::NewGuid().ToString('N')+'.tmp')
 $acl=New-SecretAcl -UserSid $userSid;$acl.SetOwner($userSid)
 $stream=$null;$pointer=[IntPtr]::Zero;$bytes=$null
 try{
  if($PSVersionTable.PSEdition -eq 'Core'){$stream=[IO.FileSystemAclExtensions]::Create([IO.FileInfo]::new($temporary),[IO.FileMode]::CreateNew,[Security.AccessControl.FileSystemRights]::Write,[IO.FileShare]::None,4096,[IO.FileOptions]::WriteThrough,$acl)}
  else{$stream=[IO.FileStream]::new($temporary,[IO.FileMode]::CreateNew,[Security.AccessControl.FileSystemRights]::Write,[IO.FileShare]::None,4096,[IO.FileOptions]::WriteThrough,$acl)}
  Assert-SecretAcl -Path $temporary -UserSid $userSid
  & $BeforeWrite
  $pointer=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($Secret)
  $bytes=[Text.Encoding]::UTF8.GetBytes([Runtime.InteropServices.Marshal]::PtrToStringBSTR($pointer))
  if($bytes.Length -eq 0){throw 'SECRET_EMPTY'}
  $stream.Write($bytes,0,$bytes.Length);$stream.Flush($true);$stream.Dispose();$stream=$null
  Assert-SecretAcl -Path $temporary -UserSid $userSid
  if(Test-Path -LiteralPath $Path){[IO.File]::Replace($temporary,$Path,[System.Management.Automation.Language.NullString]::Value)}else{[IO.File]::Move($temporary,$Path)}
  Assert-SecretAcl -Path $Path -UserSid $userSid
 }finally{
  if($stream){$stream.Dispose()};if($bytes){[Array]::Clear($bytes,0,$bytes.Length)}
  if($pointer -ne [IntPtr]::Zero){[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($pointer)}
  if(Test-Path -LiteralPath $temporary){Remove-Item -LiteralPath $temporary -Force -ErrorAction Stop}
 }
}
function Initialize-SecretPlaceholder {
 param([string]$Path)
 Initialize-SecretDirectory -Path (Split-Path -Parent $Path)
 if(Test-Path -LiteralPath $Path){Protect-ExistingSecret -Path $Path;return}
 $acl=New-SecretAcl -UserSid (Get-CurrentSecretSid);$acl.SetOwner((Get-CurrentSecretSid))
 if($PSVersionTable.PSEdition -eq 'Core'){$stream=[IO.FileSystemAclExtensions]::Create([IO.FileInfo]::new($Path),[IO.FileMode]::CreateNew,[Security.AccessControl.FileSystemRights]::Write,[IO.FileShare]::None,4096,[IO.FileOptions]::WriteThrough,$acl)}
 else{$stream=[IO.FileStream]::new($Path,[IO.FileMode]::CreateNew,[Security.AccessControl.FileSystemRights]::Write,[IO.FileShare]::None,4096,[IO.FileOptions]::WriteThrough,$acl)}
 $stream.Dispose();Assert-SecretAcl -Path $Path -UserSid (Get-CurrentSecretSid)
}
