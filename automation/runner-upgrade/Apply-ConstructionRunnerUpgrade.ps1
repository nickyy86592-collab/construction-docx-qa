param([switch]$ValidateOnly)
$ErrorActionPreference='Stop'
$RunnerRoot='C:\actions-runner'
$StageRoot=$PSScriptRoot
$ExpectedVersion='2.338.0'
$ExpectedHash='f48e0750a21812bca5f82de5f7f5aeae71abee647fab5a582f1742d07eba455f'
$ExpectedRepo='https://github.com/nickyy86592-collab/construction-docx-qa'
function Get-ConfigFingerprint {
  $result=@{}
  Get-ChildItem -LiteralPath $RunnerRoot -File -Force |
    Where-Object {$_.Name.StartsWith('.')} |
    ForEach-Object {$result[$_.Name]=(Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash}
  return $result
}
function Assert-ConfigFingerprint($baseline) {
  $current=Get-ConfigFingerprint
  if($baseline.Count -ne $current.Count){throw 'Registration/configuration file set changed'}
  foreach($name in $baseline.Keys){if($baseline[$name] -ne $current[$name]){throw "Protected configuration changed: $name"}}
}
function Get-ListenerVersion($path) {
  $lines=@(& $path --version)
  if($LASTEXITCODE -ne 0){throw 'Runner version check failed'}
  $version=@($lines | Where-Object {$_ -match '^\d+\.\d+\.\d+$'})
  if($version.Count -ne 1){throw 'Cannot verify exact runner version'}
  return [string]$version[0]
}
$registration=Get-Content -Raw -LiteralPath (Join-Path $RunnerRoot '.runner') | ConvertFrom-Json
if($registration.gitHubUrl.TrimEnd('/') -ne $ExpectedRepo){throw 'Refusing to modify a runner outside the construction repository'}
if($registration.agentName -ne 'DESKTOP-0472OML-word'){throw 'Unexpected construction runner name'}
if(Test-Path -LiteralPath (Join-Path $RunnerRoot '.service')){throw 'This updater supports the interactive construction runner only'}
$package=Join-Path $StageRoot 'runner.zip'
if((Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToLowerInvariant() -ne $ExpectedHash){throw 'Official package checksum mismatch'}
$unpacked=Join-Path $StageRoot 'unpacked'
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive=[IO.Compression.ZipFile]::OpenRead($package)
try {$rootNames=@($archive.Entries | ForEach-Object {($_.FullName -split '[/\\]')[0]} | Where-Object {$_} | Sort-Object -Unique)}
finally {$archive.Dispose()}
$items=@($rootNames | ForEach-Object {Get-Item -LiteralPath (Join-Path $unpacked $_) -Force})
foreach($item in $items) {
  if($item.Name.StartsWith('.') -or $item.Name.StartsWith('_')){throw 'Package attempts to replace protected data'}
  if($item.PSIsContainer -and $item.Name -notin @('bin','externals')){throw "Unexpected package directory: $($item.Name)"}
}
$sourceListener=Join-Path $unpacked 'bin\Runner.Listener.exe'
if((Get-ListenerVersion $sourceListener) -ne $ExpectedVersion){throw 'Staged runner version mismatch'}
if(-not (Test-Path (Join-Path $unpacked 'run.cmd')) -or -not (Test-Path (Join-Path $unpacked 'externals'))){throw 'Staged package incomplete'}
if($ValidateOnly){Write-Host 'VALIDATED: exact official package, version, and construction registration';exit 0}
$lock=$null
$backup=$null
$moved=[System.Collections.Generic.List[string]]::new()
$written=[System.Collections.Generic.List[string]]::new()
$baseline=Get-ConfigFingerprint
$oldVersion=Get-ListenerVersion (Join-Path $RunnerRoot 'bin\Runner.Listener.exe')
if($oldVersion -eq $ExpectedVersion){Write-Host "Already upgraded to $ExpectedVersion. Start the original construction runner shortcut.";exit 0}
if([version]$oldVersion -gt [version]$ExpectedVersion){throw 'Refusing to downgrade a newer runner'}
try {
  $lock=[IO.File]::Open((Join-Path $StageRoot 'upgrade.lock'),[IO.FileMode]::OpenOrCreate,[IO.FileAccess]::ReadWrite,[IO.FileShare]::None)
  $owned=@(Get-CimInstance Win32_Process | Where-Object {
    $_.Name -match '^Runner\.(Listener|Worker|Service)\.exe$' -and
    $_.ExecutablePath -and
    $_.ExecutablePath.StartsWith($RunnerRoot+'\',[StringComparison]::OrdinalIgnoreCase)
  })
  if($owned.Count -gt 0){throw 'Close the CONSTRUCTION runner window first, then double-click this updater again. Do not close the beautification runner.'}
  $backup=Join-Path $StageRoot ('backup-'+(Get-Date -Format 'yyyyMMdd-HHmmss'))
  New-Item -ItemType Directory -Path $backup -ErrorAction Stop | Out-Null
  foreach($item in $items) {
    $destination=Join-Path $RunnerRoot $item.Name
    if(Test-Path -LiteralPath $destination) {
      Move-Item -LiteralPath $destination -Destination (Join-Path $backup $item.Name) -ErrorAction Stop
      $moved.Add($item.Name)
    }
    $written.Add($item.Name)
    Copy-Item -LiteralPath $item.FullName -Destination $destination -Recurse -Force -ErrorAction Stop
  }
  $installed=Get-ListenerVersion (Join-Path $RunnerRoot 'bin\Runner.Listener.exe')
  if($installed -ne $ExpectedVersion){throw 'Installed version verification failed'}
  Assert-ConfigFingerprint $baseline
  @{status='SUCCESS';previous_version=$oldVersion;installed_version=$installed;registration_preserved=$true;backup=$backup;completed_at=(Get-Date).ToString('o')} |
    ConvertTo-Json | Set-Content -LiteralPath (Join-Path $StageRoot 'upgrade-result.json') -Encoding UTF8
  Write-Host ''
  Write-Host "SUCCESS: construction runner upgraded from $oldVersion to $installed." -ForegroundColor Green
  Write-Host 'Registration and work files are preserved. No re-registration is required.'
  Write-Host 'Now start your ORIGINAL construction runner shortcut.'
} catch {
  $failure=$_
  if($null -ne $backup) {
    foreach($name in $written) {
      $destination=Join-Path $RunnerRoot $name
      if(Test-Path -LiteralPath $destination){Remove-Item -LiteralPath $destination -Recurse -Force}
    }
    foreach($name in $moved) {
      Move-Item -LiteralPath (Join-Path $backup $name) -Destination (Join-Path $RunnerRoot $name) -ErrorAction Stop
    }
    Assert-ConfigFingerprint $baseline
    Write-Host 'Original runner binaries restored after unsuccessful upgrade.'
  }
  throw $failure
} finally {
  if($null -ne $lock){$lock.Dispose()}
}
