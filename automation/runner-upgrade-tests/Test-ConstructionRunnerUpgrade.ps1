$ErrorActionPreference='Stop'
$source=Join-Path $PSScriptRoot '../runner-upgrade/Apply-ConstructionRunnerUpgrade.ps1'
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile($source,[ref]$tokens,[ref]$errors)
if($errors.Count){throw ($errors | Out-String)}
foreach($path in @(
  $source,
  (Join-Path $PSScriptRoot '../runner-upgrade/Stage-ConstructionRunnerUpgrade.ps1'),
  (Join-Path $PSScriptRoot '../runner-upgrade-offline/Install-ConstructionRunnerUpgrade.ps1')
)) {
  $parseTokens=$null;$parseErrors=$null
  [void][Management.Automation.Language.Parser]::ParseFile($path,[ref]$parseTokens,[ref]$parseErrors)
  if($parseErrors.Count){throw ($parseErrors | Out-String)}
}
foreach($name in @('Get-ConfigFingerprint','Assert-ConfigFingerprint')) {
  $fn=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq $name},$true)
  if($null -eq $fn){throw 'Required helper missing'}
  Invoke-Expression $fn.Extent.Text
}
function Get-CimInstance {param($ClassName) return @()}
function Get-ListenerVersion($path) {
  return (Get-Content -Raw -LiteralPath (Join-Path (Split-Path $path -Parent) 'marker.txt')).Trim()
}
$transaction=@($ast.EndBlock.Statements | Where-Object {$_ -is [Management.Automation.Language.TryStatementAst]})[-1]
if($null -eq $transaction){throw 'Actual install transaction not found'}
$transactionScript=[scriptblock]::Create($transaction.Extent.Text)
function Test-InstallCase([int]$failureIndex) {
  $sandbox=Join-Path ([IO.Path]::GetTempPath()) ('runner-upgrade-test-'+[guid]::NewGuid().ToString('N'))
  $RunnerRoot=Join-Path $sandbox 'installed'
  $StageRoot=Join-Path $sandbox 'stage'
  $newRoot=Join-Path $sandbox 'new'
  New-Item -ItemType Directory -Force -Path $RunnerRoot,$StageRoot,$newRoot | Out-Null
  try {
    foreach($dir in @('bin','externals')) {
      New-Item -ItemType Directory -Force -Path (Join-Path $RunnerRoot $dir),(Join-Path $newRoot $dir) | Out-Null
      Set-Content -LiteralPath (Join-Path $RunnerRoot "$dir/marker.txt") -Value '2.337.0'
      Set-Content -LiteralPath (Join-Path $newRoot "$dir/marker.txt") -Value '2.338.0'
    }
    foreach($name in @('.runner','.credentials','.env')) {
      Set-Content -LiteralPath (Join-Path $RunnerRoot $name) -Value "protected-$name"
    }
    New-Item -ItemType Directory -Force -Path (Join-Path $RunnerRoot '_work') | Out-Null
    Set-Content -LiteralPath (Join-Path $RunnerRoot '_work/document.docx') -Value 'protected-work-file'
    Set-Content -LiteralPath (Join-Path $RunnerRoot 'custom-start.cmd') -Value 'protected-custom-start'
    Set-Content -LiteralPath (Join-Path $RunnerRoot 'run.cmd') -Value 'old-start'
    Set-Content -LiteralPath (Join-Path $newRoot 'run.cmd') -Value 'new-start'
    $items=@('bin','externals','run.cmd' | ForEach-Object {Get-Item -LiteralPath (Join-Path $newRoot $_)})
    if($failureIndex -ge 0) {
      $items[$failureIndex]=[pscustomobject]@{Name=$items[$failureIndex].Name;FullName=(Join-Path $sandbox 'missing-source')}
    }
    $ExpectedVersion='2.338.0'
    $oldVersion='2.337.0'
    $lock=$null;$backup=$null
    $moved=[System.Collections.Generic.List[string]]::new()
    $written=[System.Collections.Generic.List[string]]::new()
    $baseline=Get-ConfigFingerprint
    $caught=$false
    try {& $transactionScript} catch {$caught=$true}
    if($failureIndex -lt 0 -and $caught){throw 'Successful installation case unexpectedly failed'}
    if($failureIndex -ge 0 -and -not $caught){throw 'Injected copying failure was not detected'}
    $expected=if($failureIndex -lt 0){'2.338.0'}else{'2.337.0'}
    foreach($dir in @('bin','externals')) {
      if((Get-Content -Raw (Join-Path $RunnerRoot "$dir/marker.txt")).Trim() -ne $expected){throw 'Binary copy or rollback incorrect'}
    }
    Assert-ConfigFingerprint $baseline
    if((Get-Content -Raw (Join-Path $RunnerRoot '_work/document.docx')).Trim() -ne 'protected-work-file'){throw 'Work file changed'}
    if((Get-Content -Raw (Join-Path $RunnerRoot 'custom-start.cmd')).Trim() -ne 'protected-custom-start'){throw 'Custom startup changed'}
    $expectedStart=if($failureIndex -lt 0){'new-start'}else{'old-start'}
    if((Get-Content -Raw (Join-Path $RunnerRoot 'run.cmd')).Trim() -ne $expectedStart){throw 'Startup copy or rollback incorrect'}
    Write-Host "PASS_INSTALL_OR_ROLLBACK_CASE=$failureIndex"
  } finally {
    Remove-Item -LiteralPath $sandbox -Recurse -Force
  }
}
Test-InstallCase -1
Test-InstallCase 0
Test-InstallCase 1
Write-Host 'PASS: PowerShell syntax; actual install transaction; rollback on first and second copy failures; protected configuration and work files'
