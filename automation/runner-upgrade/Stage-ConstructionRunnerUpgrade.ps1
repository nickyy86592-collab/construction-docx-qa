$ErrorActionPreference='Stop'
if($env:GITHUB_REPOSITORY -ne 'nickyy86592-collab/construction-docx-qa'){throw 'Repository mismatch'}
if($env:RUNNER_NAME -ne 'DESKTOP-0472OML-word'){throw 'Runner mismatch'}
$RunnerRoot='C:\actions-runner'
$registration=Get-Content -Raw -LiteralPath (Join-Path $RunnerRoot '.runner') | ConvertFrom-Json
if($registration.gitHubUrl.TrimEnd('/') -ne 'https://github.com/nickyy86592-collab/construction-docx-qa'){throw 'Registration mismatch'}
$stageRoot='C:\actions-runner-upgrade\2.338.0'
New-Item -ItemType Directory -Force -Path $stageRoot | Out-Null
$bundleRoot=$PSScriptRoot
$package=Join-Path $bundleRoot 'runner.zip'
$hash=(Get-FileHash -LiteralPath $package -Algorithm SHA256).Hash.ToLowerInvariant()
if($hash -ne 'f48e0750a21812bca5f82de5f7f5aeae71abee647fab5a582f1742d07eba455f'){throw 'Official package checksum mismatch'}
Copy-Item -LiteralPath $package -Destination (Join-Path $stageRoot 'runner.zip') -Force
Copy-Item -LiteralPath (Join-Path $bundleRoot 'Apply-ConstructionRunnerUpgrade.ps1') -Destination $stageRoot -Force
$scriptPath=Join-Path $stageRoot 'Apply-ConstructionRunnerUpgrade.ps1'
$tokens=$null;$errors=$null
[void][Management.Automation.Language.Parser]::ParseFile($scriptPath,[ref]$tokens,[ref]$errors)
if($errors.Count -gt 0){throw ($errors | Out-String)}
$unpacked=Join-Path $stageRoot 'unpacked'
if(Test-Path -LiteralPath $unpacked){Remove-Item -LiteralPath $unpacked -Recurse -Force}
New-Item -ItemType Directory -Path $unpacked | Out-Null
Add-Type -AssemblyName System.IO.Compression.FileSystem
$archive=[IO.Compression.ZipFile]::OpenRead((Join-Path $stageRoot 'runner.zip'))
try {
  $prefix=[IO.Path]::GetFullPath($unpacked)+'\'
  foreach($entry in $archive.Entries) {
    $destination=[IO.Path]::GetFullPath((Join-Path $unpacked $entry.FullName))
    if(-not $destination.StartsWith($prefix,[StringComparison]::OrdinalIgnoreCase)){throw 'Invalid archive path'}
  }
} finally {$archive.Dispose()}
[IO.Compression.ZipFile]::ExtractToDirectory((Join-Path $stageRoot 'runner.zip'),$unpacked)
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File $scriptPath -ValidateOnly
if($LASTEXITCODE -ne 0){throw 'Offline updater validation failed'}
$desktop=[Environment]::GetFolderPath('DesktopDirectory')
if(-not (Test-Path -LiteralPath $desktop)){throw 'Desktop directory unavailable'}
$launcherPath=Join-Path $desktop 'Upgrade-ConstructionRunner.cmd'
$launcher=(@('@echo off','cd /d C:\actions-runner-upgrade\2.338.0','powershell.exe -NoProfile -ExecutionPolicy Bypass -File Apply-ConstructionRunnerUpgrade.ps1','echo.','pause') -join [Environment]::NewLine)+[Environment]::NewLine
[IO.File]::WriteAllText($launcherPath,$launcher,[Text.Encoding]::ASCII)
Write-Host "UPGRADE_READY_VERSION=2.338.0"
Write-Host "UPGRADE_PACKAGE_SHA256=$hash"
Write-Host "UPGRADE_STAGE=$stageRoot"
Write-Host "UPGRADE_LAUNCHER=$launcherPath"
Write-Host 'CURRENT_RUNNING_INSTALLATION_UNCHANGED'
Write-Host 'ACTION_REQUIRED: close only construction runner window, run desktop Upgrade-ConstructionRunner.cmd, then reopen original construction shortcut'
